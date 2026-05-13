from __future__ import annotations

import argparse
import json

from .core.processor import process_zip


def main() -> None:
    parser = argparse.ArgumentParser(description="Process talent ninebox Excel zip.")
    parser.add_argument("zip_path")
    parser.add_argument("output_dir")
    args = parser.parse_args()

    result = process_zip(args.zip_path, args.output_dir)
    print(json.dumps(result.summary.as_dict(), ensure_ascii=False, indent=2))
    print(result.output_file)


if __name__ == "__main__":
    main()
