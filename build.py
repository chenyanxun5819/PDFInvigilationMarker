from __future__ import annotations

import shutil
import subprocess
import sys
from datetime import datetime
from importlib import import_module
from pathlib import Path

from app.version import APP_NAME, APP_VERSION

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


def ensure_required_modules() -> None:
    required_modules = ("pymupdf", "openpyxl", "rapidocr_onnxruntime")
    missing_modules: list[str] = []
    for module_name in required_modules:
        try:
            import_module(module_name)
        except ModuleNotFoundError:
            missing_modules.append(module_name)

    if missing_modules:
        missing_text = ", ".join(missing_modules)
        raise ModuleNotFoundError(
            f"打包环境缺少必要模组：{missing_text}。请先安装 requirements.txt 后再执行 build.py。"
        )


def resolve_output_name(dist_dir: Path, base_name: str) -> str:
    exe_path = dist_dir / f"{base_name}.exe"
    if not exe_path.exists():
        return base_name

    try:
        exe_path.unlink()
        return base_name
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base_name}_{timestamp}"


def build_executable():
    ensure_required_modules()

    project_dir = Path(__file__).parent
    dist_dir = project_dir / "dist"
    build_dir = project_dir / "build"
    output_name = resolve_output_name(dist_dir, f"{APP_NAME}_v{APP_VERSION}")

    if build_dir.exists():
        shutil.rmtree(build_dir)
    dist_dir.mkdir(parents=True, exist_ok=True)

    run_command(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--onefile",
            "--windowed",
            f"--name={output_name}",
            f"--distpath={dist_dir}",
            f"--workpath={build_dir}",
            f"--specpath={project_dir}",
            f"--add-data={project_dir / 'data'};data",
            "--collect-data=rapidocr_onnxruntime",
            "--collect-submodules=rapidocr_onnxruntime",
            "--hidden-import=app.template_generator",
            "--hidden-import=openpyxl",
            "--hidden-import=pymupdf",
            str(project_dir / "main.py"),
        ],
        "PyInstaller 打包",
    )

    exe_path = dist_dir / f"{output_name}.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"打包完成后找不到输出档：{exe_path}")

    size_mb = exe_path.stat().st_size / (1024 * 1024)
    print(f"\n✅ 打包完成：{exe_path}")
    print(f"档案大小：{size_mb:.2f} MB")


if __name__ == "__main__":
    build_executable()
