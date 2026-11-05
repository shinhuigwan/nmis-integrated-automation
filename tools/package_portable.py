"""Create an update-compatible release ZIP without any local user data."""
import argparse
import hashlib
import json
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from app_updater import PAYLOAD_NAMES, ROOT_NAME
from app_version import APP_VERSION


def package_bundle(bundle, output):
    bundle, output = Path(bundle), Path(output)
    for name in ("NMIS-Automation.exe", "_internal", "apply_update.ps1", "installers/ChromeSetup.exe"):
        if not (bundle / name).exists():
            raise ValueError(f"Required release file missing: {name}")
    (bundle / "BUILD_INFO.json").write_text(json.dumps({"version": APP_VERSION, "platform": "Windows x64", "python": sys.version.split()[0]}, indent=2), encoding="utf-8")
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for name in sorted(PAYLOAD_NAMES):
            path = bundle / name
            if path.is_dir():
                for file in sorted(path.rglob("*")):
                    if file.is_file():
                        archive.write(file, (Path(ROOT_NAME) / file.relative_to(bundle)).as_posix())
            elif path.is_file():
                archive.write(path, (Path(ROOT_NAME) / name).as_posix())
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    output.with_suffix(output.suffix + ".sha256").write_text(f"{digest}  {output.name}\n", encoding="ascii")
    return digest


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("bundle")
    parser.add_argument("output")
    args = parser.parse_args()
    print(package_bundle(args.bundle, args.output))
