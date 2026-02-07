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
            categorization TEXT,
            value REAL
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
                gl_mapping_type TEXT
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
            gl_account, categorization, gl_mapping_type = row
            gl_rows.append(
                (
                    str(gl_account).strip(),
                    str(categorization).strip() if str(categorization).strip() else None,
                    str(gl_mapping_type).strip() if str(gl_mapping_type).strip() else None,
                )
            )

        conn.executemany(
            "INSERT INTO _investor_map (investor, property_name, owner, acquired, investor_type) VALUES (?, ?, ?, ?, ?);",
            inv_rows,
        )

        conn.executemany(
            "INSERT INTO _gl_map (gl_account, categorization, gl_mapping_type) VALUES (?, ?, ?);",
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
) -> None:
    conn = connect(db_path)
    try:
        reset_agg_table(conn, gl_agg_table)

        conn.execute(
            f"""
            INSERT INTO {gl_agg_table} (month_start, investor, owner, property_name, categorization, value)
            SELECT
                month_start,
                investor,
                owner,
                property_name,
                categorization,
                COALESCE(SUM(debit), 0.0) - COALESCE(SUM(credit), 0.0) AS value
            FROM {gl_raw_table}
            GROUP BY
                month_start,
                investor,
                owner,
                property_name,
                categorization
            ORDER BY
                month_start ASC,
                investor ASC,
                owner ASC,
                property_name ASC,
                categorization ASC;
            """
        )
        conn.commit()
    finally:
        conn.close()
