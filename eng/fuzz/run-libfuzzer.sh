#!/usr/bin/env bash
set -euo pipefail

if [[ $# -lt 1 || $# -gt 2 ]]; then
  echo "usage: $0 CORPUS_DIRECTORY [MAX_TOTAL_TIME_SECONDS]" >&2
  exit 2
fi

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
repo_root="$(cd -- "$script_dir/../.." && pwd)"
corpus_dir="$(cd -- "$1" && pwd)"
max_total_time="${2:-120}"

if ! [[ "$max_total_time" =~ ^[1-9][0-9]*$ ]]; then
  echo "MAX_TOTAL_TIME_SECONDS must be a positive integer" >&2
  exit 2
fi

output_root="$repo_root/artifacts/fuzz"
driver="$output_root/driver/libfuzzer-dotnet"
tool_dir="$output_root/tools"
publish_root="$output_root/publish"
crash_dir="$output_root/crashes"
project="$repo_root/fuzz/Cellwright.Fuzz/Cellwright.Fuzz.csproj"

driver_url="https://github.com/Metalnem/libfuzzer-dotnet/releases/download/v2025.05.02.0904/libfuzzer-dotnet-ubuntu"
driver_sha256="c2c2a90d94c409a4af339a0d4f244e0442c5a5d249be0f1252ba07871f285958"

mkdir -p "$(dirname -- "$driver")" "$tool_dir" "$publish_root" "$crash_dir"
publish_dir="$(mktemp -d "$publish_root/run.XXXXXX")"
driver_temp=""
cleanup() {
  rm -rf -- "$publish_dir"
  if [[ -n "$driver_temp" ]]; then
    rm -f -- "$driver_temp"
  fi
}
trap cleanup EXIT

if [[ ! -f "$driver" ]] || ! echo "$driver_sha256  $driver" | sha256sum --check --status; then
  driver_temp="$(mktemp "$(dirname -- "$driver")/.download.XXXXXX")"
  curl --fail --location --silent --show-error "$driver_url" --output "$driver_temp"
  echo "$driver_sha256  $driver_temp" | sha256sum --check --status
  mv -f -- "$driver_temp" "$driver"
  driver_temp=""
fi
echo "$driver_sha256  $driver" | sha256sum --check --status
chmod +x "$driver"

if [[ ! -x "$tool_dir/sharpfuzz" ]]; then
  dotnet tool install SharpFuzz.CommandLine --version 2.3.0 --tool-path "$tool_dir"
fi

dotnet restore "$project" --locked-mode --verbosity minimal
dotnet publish "$project" --configuration Release --no-restore --output "$publish_dir" --verbosity minimal

for assembly in Cellwright.dll Cellwright.Formulas.dll Cellwright.Xls.dll; do
  "$tool_dir/sharpfuzz" "$publish_dir/$assembly"
done

"$driver" \
  -artifact_prefix="$crash_dir/" \
  -dict="$repo_root/fuzz/excel.dict" \
  -max_len=1048576 \
  -max_total_time="$max_total_time" \
  -rss_limit_mb=1024 \
  -timeout=10 \
  --target_path=dotnet \
  --target_arg="$publish_dir/Cellwright.Fuzz.dll" \
  "$corpus_dir"
