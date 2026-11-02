from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def run_command(command: list[str], description: str):
    print(f"\n{'=' * 72}")
    print(f"▶ {description}")
    print(f"{'=' * 72}")
    print("命令:", " ".join(command))
    completed = subprocess.run(command)
    if completed.returncode != 0:
        raise SystemExit(completed.returncode)


def build_executable():
    project_dir = Path(__file__).parent
    dist_dir = project_dir / "dist"
    build_dir = project_dir / "build"

    for path in (dist_dir, build_dir):
        if path.exists():
            shutil.rmtree(path)

    run_command(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--onefile",
            "--windowed",
            "--name=PDF监考名单标记工具",
            f"--distpath={dist_dir}",
            f"--workpath={build_dir}",
            f"--specpath={project_dir}",
            f"--add-data={project_dir / 'data'};data",
            "--collect-data=rapidocr_onnxruntime",
            "--collect-submodules=rapidocr_onnxruntime",
            str(project_dir / "main.py"),
        ],
        "PyInstaller 打包",
    )

    exe_path = dist_dir / "PDF监考名单标记工具.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"打包完成后找不到输出档：{exe_path}")

    size_mb = exe_path.stat().st_size / (1024 * 1024)
    print(f"\n✅ 打包完成：{exe_path}")
    print(f"档案大小：{size_mb:.2f} MB")


if __name__ == "__main__":
    build_executable()
