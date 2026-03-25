"""
Build script for Micro Automation desktop app.

Usage: python build_desktop.py [--skip-frontend] [--skip-installer]

Steps:
1. Build Next.js static export -> web/out/
2. Copy web/out/ -> static_web/
3. Run PyInstaller -> dist/MicroAutomation/
4. Run Inno Setup -> dist/MicroAutomation_Setup_vX.Y.Z.exe
"""

import argparse
import os
import shutil
import subprocess
import sys

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))


def run(cmd: list[str], cwd: str | None = None, env: dict | None = None):
    """Run a command and exit on failure."""
    print(f"\n>>> {' '.join(cmd)}")
    merged_env = {**os.environ, **(env or {})}
    result = subprocess.run(cmd, cwd=cwd, env=merged_env)
    if result.returncode != 0:
        print(f"ERROR: Command failed with exit code {result.returncode}")
        sys.exit(1)


def build_frontend():
    """Build Next.js static export.

    Static export is incompatible with API routes and middleware (they use
    force-dynamic / server-only auth). Since desktop mode doesn't need auth,
    we temporarily move these files out of the way during the build.
    """
    print("\n=== Step 1: Building Next.js static export ===")
    web_dir = os.path.join(PROJECT_ROOT, "web")
    static_dir = os.path.join(PROJECT_ROOT, "static_web")
    src_dir = os.path.join(web_dir, "src")

    # Files incompatible with static export (server-only auth)
    files_to_skip = [
        os.path.join(src_dir, "middleware.ts"),
        os.path.join(src_dir, "app", "api"),  # entire api directory
    ]
    # Backup location OUTSIDE web/src so Next.js doesn't scan it
    backup_dir = os.path.join(PROJECT_ROOT, "_static_build_backup")

    # Install deps if needed
    if not os.path.isdir(os.path.join(web_dir, "node_modules")):
        run(["npm", "ci"], cwd=web_dir)

    # Clean .next cache to avoid stale references
    next_cache = os.path.join(web_dir, ".next")
    if os.path.isdir(next_cache):
        shutil.rmtree(next_cache)
        print("  Cleaned .next cache")

    # Temporarily move incompatible files outside the web directory
    os.makedirs(backup_dir, exist_ok=True)
    moved: list[tuple[str, str]] = []
    for path in files_to_skip:
        if os.path.exists(path):
            basename = os.path.basename(path)
            backup = os.path.join(backup_dir, basename)
            print(f"  Moving aside: {os.path.relpath(path, PROJECT_ROOT)} -> _static_build_backup/{basename}")
            shutil.move(path, backup)
            moved.append((path, backup))

    try:
        # Build with static export
        run(
            ["npm", "run", "build"],
            cwd=web_dir,
            env={
                "BUILD_STATIC": "true",
                "NEXT_PUBLIC_LOCAL_MODE": "true",
                "NEXT_PUBLIC_API_URL": "",
            },
        )
    finally:
        # Always restore moved files, even if build fails
        for original, backup in moved:
            if os.path.exists(backup):
                shutil.move(backup, original)
                print(f"  Restored: {os.path.relpath(original, PROJECT_ROOT)}")
        # Clean up backup dir
        if os.path.isdir(backup_dir):
            shutil.rmtree(backup_dir, ignore_errors=True)

    # Copy output to static_web/
    out_dir = os.path.join(web_dir, "out")
    if not os.path.isdir(out_dir):
        print(f"ERROR: Expected static export at {out_dir} but it doesn't exist")
        sys.exit(1)

    if os.path.isdir(static_dir):
        shutil.rmtree(static_dir)
    shutil.copytree(out_dir, static_dir)
    print(f"Copied {out_dir} -> {static_dir}")


def build_pyinstaller():
    """Run PyInstaller to create the app bundle."""
    print("\n=== Step 2: Running PyInstaller ===")
    run(["pyinstaller", "--clean", "--noconfirm", "desktop.spec"], cwd=PROJECT_ROOT)


def build_installer():
    """Run Inno Setup to create the Windows installer."""
    print("\n=== Step 3: Running Inno Setup ===")
    # iscc is the Inno Setup command-line compiler
    run(["iscc", "installer.iss"], cwd=PROJECT_ROOT)


def main():
    parser = argparse.ArgumentParser(description="Build Micro Automation desktop app")
    parser.add_argument("--skip-frontend", action="store_true", help="Skip Next.js build")
    parser.add_argument("--skip-installer", action="store_true", help="Skip Inno Setup")
    args = parser.parse_args()

    print("=" * 60)
    print("  Micro Automation — Desktop Build")
    print("=" * 60)

    if not args.skip_frontend:
        build_frontend()
    else:
        print("\nSkipping frontend build (--skip-frontend)")

    build_pyinstaller()

    if not args.skip_installer:
        build_installer()
    else:
        print("\nSkipping installer (--skip-installer)")

    print("\n" + "=" * 60)
    print("  BUILD COMPLETE")
    print("=" * 60)
    print(f"\nInstaller: dist/MicroAutomation_Setup_v*.exe")
    print(f"Portable:  dist/MicroAutomation/MicroAutomation.exe")


if __name__ == "__main__":
    main()
