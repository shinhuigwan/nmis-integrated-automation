"""Include licenses from the installed build dependencies in the release."""
import importlib.metadata
import sys
from pathlib import Path

sections = []
python_license = Path(sys.base_prefix) / "LICENSE.txt"
if python_license.is_file():
    sections.append("Python runtime\n" + python_license.read_text(encoding="utf-8", errors="replace"))
for distribution in sorted(importlib.metadata.distributions(), key=lambda d: d.metadata["Name"].lower()):
    name = distribution.metadata["Name"]
    sections.append(f"\n{'=' * 72}\n{name} {distribution.version}\nLicense: {distribution.metadata.get('License-Expression', distribution.metadata.get('License', 'See notices below'))}\n")
    for file in distribution.files or []:
        if file.name.lower().startswith(("license", "copying", "notice", "copyright")) and file.suffix.lower() in {"", ".txt", ".md"}:
            path = distribution.locate_file(file)
            if path.is_file():
                sections.append(str(file) + "\n" + path.read_text(encoding="utf-8", errors="replace"))
Path(sys.argv[1]).write_text("\n\n".join(sections), encoding="utf-8")
