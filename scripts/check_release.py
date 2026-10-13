"""Fail CI when documented installation pins disagree with package metadata."""
from pathlib import Path
import re
import tomllib


def main():
    root = Path(__file__).resolve().parents[1]
    version = tomllib.loads((root / "pyproject.toml").read_text())["project"]["version"]
    for name in ("README.md", "opencode.example.json", "OPERATIONS.md"):
        pins = re.findall(r"@v(\d+\.\d+\.\d+)", (root / name).read_text())
        if not pins or any(pin != version for pin in pins):
            raise ValueError(f"{name}: installation pins must match v{version}")
    print(f"Installation pins match v{version}")


if __name__ == "__main__":
    main()
