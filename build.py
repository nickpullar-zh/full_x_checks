import os
import shutil
import subprocess
import re

# ==========================================
# Version — update this for each release
# ==========================================
VERSION = "v0.2.2"

# ==========================================
# Build configuration
# ==========================================
BUILDS = [
    {
        "name":          f"X-Checks_Debug_{VERSION}",
        "debug_mode": True,
        "add_test_data": True,
    },
    {
        "name":       f"X-Checks_{VERSION}",
        "debug_mode": False,
        "add_test_data": False,
    },
]

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
MAIN_PY      = os.path.join(PROJECT_ROOT, "main.py")
TEMP_MAIN    = os.path.join(PROJECT_ROOT, "_main_build_temp.py")

def force_remove(func, path, _):
    import stat
    os.chmod(path, stat.S_IWRITE)
    func(path)

def build(config: dict):
    print(f"\n  Building: {config['name']}  (DEBUG_MODE={config['debug_mode']})")

    # --- Read main.py ---
    with open(MAIN_PY, "r", encoding="utf-8") as f:
        source = f.read()

    # --- Patch DEBUG_MODE line ---
    source = re.sub(
        r"DEBUG_MODE\s*=\s*(True|False)",
        f"DEBUG_MODE = {config['debug_mode']}",
        source
    )

    # --- Write patched temp file ---
    with open(TEMP_MAIN, "w", encoding="utf-8") as f:
        f.write(source)

    # --- Pre-clean build folder before PyInstaller runs ---
    build_folder = os.path.join(PROJECT_ROOT, "build", config["name"])
    if os.path.exists(build_folder):
        print(f"  Pre-cleaning: {build_folder}")
        try:
            shutil.rmtree(build_folder)
        except PermissionError:
            shutil.rmtree(build_folder, onexc=force_remove)

    # --- Build PyInstaller command (without --clean) ---
    cmd = [
        "pyinstaller",
        "--onefile",
        "--name", config["name"],
        "--add-data", "templates;templates",
    ]

    if config["add_test_data"]:
        cmd += ["--add-data", "test_data;test_data"]

    cmd.append(TEMP_MAIN)

    print(f"  [BUILD] Command: {' '.join(cmd)}")

    # --- Run PyInstaller ---
    subprocess.run(cmd, check=True)

    # --- Clean up temp file ---
    os.remove(TEMP_MAIN)

    print(f"  Built successfully: dist\\{config['name']}.exe")


if __name__ == "__main__":
    for build_config in BUILDS:
        build(build_config)

    # --- Clean up build artifacts ---
    print("\n  Cleaning up build folders...")
    for build_config in BUILDS:
        build_folder = os.path.join(PROJECT_ROOT, "build", build_config["name"])
        spec_file    = os.path.join(PROJECT_ROOT, f"{build_config['name']}.spec")

        if os.path.exists(build_folder):
            try:
                shutil.rmtree(build_folder)
                print(f"  Removed: {build_folder}")
            except PermissionError:
                # OneDrive or antivirus may be locking files — force remove
                import stat
                def force_remove(func, path, _):
                    os.chmod(path, stat.S_IWRITE)
                    func(path)
                shutil.rmtree(build_folder, onexc=force_remove)
                print(f"  Removed (forced): {build_folder}")

        if os.path.exists(spec_file):
            try:
                os.remove(spec_file)
                print(f"  Removed: {spec_file}")
            except PermissionError:
                print(f"  WARNING: Could not remove {spec_file} — delete manually if needed")

    print(f"\n  All builds complete. Executables are in the dist\\ folder.")