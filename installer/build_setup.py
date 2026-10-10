"""Build PaperMaker_Setup.exe with the bucket's download address built in.

    python installer/build_setup.py https://pub-xxxxxxxx.r2.dev/

Output: installer/dist/PaperMaker_Setup.exe  (one small file to share).
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent


def main() -> None:
    if len(sys.argv) != 2 or not sys.argv[1].startswith("http"):
        sys.exit(__doc__)
    url = sys.argv[1].rstrip("/") + "/"
    (HERE / "_setup_config.py").write_text(f"BASE_URL = {url!r}\n", "utf-8")
    icon = HERE.parent / "assets" / "paper_maker.ico"
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--onefile", "--windowed",
                    "--name", "PaperMaker_Setup", "--hidden-import", "_setup_config",
                    "--icon", str(icon), "--add-data", f"{icon};assets",
                    "--paths", str(HERE), "--distpath", str(HERE / "dist"),
                    "--workpath", str(HERE / "build"), "--specpath", str(HERE / "build"),
                    str(HERE / "setup_app.py")], check=True)
    exe = HERE / "dist" / "PaperMaker_Setup.exe"
    print(f"\n{exe}  ({exe.stat().st_size / 2**20:.1f} MB)  ->  {url}")


if __name__ == "__main__":
    main()
