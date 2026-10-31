#!/usr/bin/env bash
# Usage: ./run.sh input_folder output.csv
set -euo pipefail
if [ "$#" -ne 2 ]; then
  echo "Usage: ./run.sh input_folder output.csv" >&2
  exit 2
fi
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
exec python3 "$HERE/src/extract.py" "$1" "$2"