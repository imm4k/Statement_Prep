from __future__ import annotations

import argparse

from part1_gl.main import run_part1


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--parts",
        type=str,
        required=True,
        help="Comma separated list of parts to run. Example: 1 or 1,2,3",
    )
    args = parser.parse_args()

    parts = [p.strip() for p in args.parts.split(",") if p.strip()]
    if not parts:
        raise ValueError("No parts provided.")

    for p in parts:
        if p == "1":
            run_part1()
        elif p in {"2", "3", "4"}:
            raise NotImplementedError(f"Part {p} not implemented yet.")
        else:
            raise ValueError(f"Unknown part: {p}")


if __name__ == "__main__":
    main()
