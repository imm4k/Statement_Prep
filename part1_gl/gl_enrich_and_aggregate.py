from __future__ import annotations

import pandas as pd

from common.sqlite_utils import clear_table, connect, table_exists


def ensure_agg_schema(conn, gl_agg_table: str) -> None:
    conn.execute(
        f"""
        CREATE TABLE IF NOT EXISTS {gl_agg_table} (
            month_start TEXT,
            investor TEXT,
            owner TEXT,
            property_name TEXT,
            property TEXT,
            acquired TEXT,
            categorization TEXT,
            gl_mapping_type TEXT,
            value REAL,
            cash_categorization TEXT,
            cash_value REAL,
            cash_type_mapping TEXT,
            timeframe TEXT
        );
        """
    )
    conn.commit()

def reset_agg_table(conn, gl_agg_table: str) -> None:
    conn.execute(f"DROP TABLE IF EXISTS {gl_agg_table};")
    ensure_agg_schema(conn, gl_agg_table)
    conn.commit()


def apply_mappings_inplace(
    db_path: str,
    gl_raw_table: str,
    investor_table_df: pd.DataFrame,
    gl_mapping_df: pd.DataFrame,
) -> None:
    conn = connect(db_path)
    try:
        conn.execute("DROP TABLE IF EXISTS _investor_map;")
        conn.execute("DROP TABLE IF EXISTS _gl_map;")

        conn.execute(
            """
            CREATE TABLE _investor_map (
                investor TEXT,
                property_name TEXT,
                owner TEXT,
                acquired TEXT,
                investor_type TEXT
            );
            """
        )
        conn.execute(
            """
            CREATE TABLE _gl_map (
                gl_account TEXT,
                categorization TEXT,
                gl_mapping_type TEXT,
                cash_categorization TEXT,
                cash_type_mapping TEXT
            );
            """
        )

        inv_rows = []
        for row in investor_table_df.itertuples(index=False, name=None):
            investor, _property, property_name, owner, acquired, inv_type = row
            inv_rows.append(
                (
                    str(investor).strip(),
                    str(property_name).strip(),
                    str(owner).strip() if str(owner).strip() else None,
                    str(acquired).strip() if str(acquired).strip() else None,
                    str(inv_type).strip() if str(inv_type).strip() else None,
                )
            )

        gl_rows = []
        for row in gl_mapping_df.itertuples(index=False, name=None):
            gl_account, categorization, gl_mapping_type, cash_categorization, cash_type_mapping = row
            gl_rows.append(
                (
                    str(gl_account).strip(),
                    str(categorization).strip() if str(categorization).strip() else None,
                    str(gl_mapping_type).strip() if str(gl_mapping_type).strip() else None,
                    str(cash_categorization).strip() if str(cash_categorization).strip() else None,
                    str(cash_type_mapping).strip() if str(cash_type_mapping).strip() else None,
                )
            )

        conn.executemany(
            "INSERT INTO _investor_map (investor, property_name, owner, acquired, investor_type) VALUES (?, ?, ?, ?, ?);",
            inv_rows,
        )

        conn.executemany(
            "INSERT INTO _gl_map (gl_account, categorization, gl_mapping_type, cash_categorization, cash_type_mapping) VALUES (?, ?, ?, ?, ?);",
            gl_rows,
        )

        conn.execute(
            f"""
            UPDATE {gl_raw_table}
            SET
                investor = (
                    SELECT m.investor
                    FROM _investor_map m
                    WHERE m.property_name = {gl_raw_table}.property_name
                    LIMIT 1
                ),
                owner = (
                    SELECT m.owner
                    FROM _investor_map m
                    WHERE m.property_name = {gl_raw_table}.property_name
                    LIMIT 1
                ),
                acquired = (
                    SELECT m.acquired
                    FROM _investor_map m
                    WHERE m.property_name = {gl_raw_table}.property_name
                    LIMIT 1
                ),
                investor_type = (
                    SELECT m.investor_type
                    FROM _investor_map m
                    WHERE m.property_name = {gl_raw_table}.property_name
                    LIMIT 1
                ),
                categorization = (
                    SELECT g.categorization
                    FROM _gl_map g
                    WHERE g.gl_account = {gl_raw_table}.gl_account
                    LIMIT 1
                ),
                gl_mapping_type = (
                    SELECT g.gl_mapping_type
                    FROM _gl_map g
                    WHERE g.gl_account = {gl_raw_table}.gl_account
                    LIMIT 1
                ),
                cash_categorization = (
                    SELECT g.cash_categorization
                    FROM _gl_map g
                    WHERE g.gl_account = {gl_raw_table}.gl_account
                    LIMIT 1
                ),
                cash_type_mapping = (
                    SELECT g.cash_type_mapping
                    FROM _gl_map g
                    WHERE g.gl_account = {gl_raw_table}.gl_account
                    LIMIT 1
                );
            """
        )

        conn.execute("DROP TABLE IF EXISTS _investor_map;")
        conn.execute("DROP TABLE IF EXISTS _gl_map;")
        conn.commit()
    finally:
        conn.close()


def build_aggregate_table(
    db_path: str,
    gl_raw_table: str,
    gl_agg_table: str,
    investor_table_df: pd.DataFrame,
    statement_thru_date: str,
) -> None:
    conn = connect(db_path)
    try:
        reset_agg_table(conn, gl_agg_table)

        conn.execute(
            f"""
            INSERT INTO {gl_agg_table} (
                month_start, investor, owner, property_name, acquired,
                categorization, gl_mapping_type, value,
                cash_categorization, cash_value, cash_type_mapping
            )
            SELECT
                month_start,
                investor,
                owner,
                property_name,
                MIN(acquired) AS acquired,
                categorization,
                gl_mapping_type,
                COALESCE(SUM(debit), 0.0) - COALESCE(SUM(credit), 0.0) AS value,
                cash_categorization,
                COALESCE(SUM(debit), 0.0) - COALESCE(SUM(credit), 0.0) AS cash_value,
                MIN(cash_type_mapping) AS cash_type_mapping
            FROM {gl_raw_table}
            WHERE COALESCE(cash_categorization, '') <> 'Mortgage'
            GROUP BY
                month_start, investor, owner, property_name, categorization, gl_mapping_type, cash_categorization

            UNION ALL

            SELECT
                month_start,
                investor,
                owner,
                property_name,
                MIN(acquired) AS acquired,
                categorization,
                gl_mapping_type,
                COALESCE(SUM(debit), 0.0) AS value,
                'Mortgage Payment' AS cash_categorization,
                COALESCE(SUM(debit), 0.0) AS cash_value,
                'Outflow' AS cash_type_mapping
            FROM {gl_raw_table}
            WHERE cash_categorization = 'Mortgage'
            GROUP BY
                month_start, investor, owner, property_name, categorization, gl_mapping_type

            UNION ALL

            SELECT
                month_start,
                investor,
                owner,
                property_name,
                MIN(acquired) AS acquired,
                categorization,
                gl_mapping_type,
                0.0 - COALESCE(SUM(credit), 0.0) AS value,
                'Mortgage Loan' AS cash_categorization,
                COALESCE(SUM(credit), 0.0) AS cash_value,
                'Inflow' AS cash_type_mapping
            FROM {gl_raw_table}
            WHERE cash_categorization = 'Mortgage'
            GROUP BY
                month_start, investor, owner, property_name, categorization, gl_mapping_type
            ;
            """
        )

        # Append property from Investor Table (Property Name -> Property)
        conn.execute("DROP TABLE IF EXISTS _prop_map;")
        conn.execute(
            """
            CREATE TABLE _prop_map (
                property_name TEXT,
                property TEXT
            );
            """
        )

        prop_rows = []
        for row in investor_table_df[["Property Name", "Property"]].itertuples(index=False, name=None):
            property_name, prop = row
            prop_rows.append((str(property_name).strip(), str(prop).strip() if str(prop).strip() else None))

        conn.executemany(
            "INSERT INTO _prop_map (property_name, property) VALUES (?, ?);",
            prop_rows,
        )

        conn.execute(
            f"""
            UPDATE {gl_agg_table}
            SET property = (
                SELECT p.property
                FROM _prop_map p
                WHERE p.property_name = {gl_agg_table}.property_name
                LIMIT 1
            );
            """
        )
        conn.execute("DROP TABLE IF EXISTS _prop_map;")

        # Append timeframe [T1]..[T13] based on Statement Thru Date month
        t1_start = pd.to_datetime(statement_thru_date, errors="raise").to_period("M").to_timestamp().strftime("%Y-%m-%d")

        conn.execute(
            f"""
            UPDATE {gl_agg_table}
            SET timeframe = (
                CASE
                    -- Future months
                    WHEN (
                        (CAST(strftime('%Y', month_start) AS INTEGER) * 12 + CAST(strftime('%m', month_start) AS INTEGER))
                        > (CAST(strftime('%Y', ?) AS INTEGER) * 12 + CAST(strftime('%m', ?) AS INTEGER))
                    )
                    THEN 'N/A'

                    -- T1 to T12
                    WHEN (
                        (CAST(strftime('%Y', ?) AS INTEGER) * 12 + CAST(strftime('%m', ?) AS INTEGER))
                        - (CAST(strftime('%Y', month_start) AS INTEGER) * 12 + CAST(strftime('%m', month_start) AS INTEGER))
                    ) BETWEEN 0 AND 11
                    THEN '[T' || (
                        (
                            (CAST(strftime('%Y', ?) AS INTEGER) * 12 + CAST(strftime('%m', ?) AS INTEGER))
                            - (CAST(strftime('%Y', month_start) AS INTEGER) * 12 + CAST(strftime('%m', month_start) AS INTEGER))
                        ) + 1
                    ) || ']'

                    -- Older than T12
                    ELSE '[T13]'
                END
            );
            """,
            (t1_start, t1_start, t1_start, t1_start, t1_start, t1_start),
        )

        conn.commit()

    finally:
        conn.close()
