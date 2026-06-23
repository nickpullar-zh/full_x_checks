import os
import sys
import shutil
import subprocess
import re

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from version import __version__

VERSION = f"v{__version__}"

# ==========================================
# Build configuration
# Pass one or more keys as CLI args to build only that subset, e.g.:
#   python build.py             → all builds (default)
#   python build.py debug       → all debug builds (any strategy)
#   python build.py prod        → production only
#
# Strategy branches add their own debug entries (e.g. {key:"xc",
# debug_task:"X-Checks", add_test_data:True}); the infra `main` branch
# ships only the production build.
# ==========================================
BUILDS = [
    {
        "key":         "prod",
        "name":        f"X-Checks_{VERSION}",
        "debug_mode":  False,
        "add_test_data": False,
    },
    {
        "key":          "conditions_debug",
        "name":         f"X-Checks_Debug_Conditions_{VERSION}",
        "debug_mode":   True,
        "debug_task":   "Conditions",
        "add_test_data": True,
    },
    {
        "key":          "full_run_debug",
        "name":         f"X-Checks_Debug_Full_Run_{VERSION}",
        "debug_mode":   True,
        "debug_task":   "Full Run",
        "add_test_data": True,
    },
]

PROJECT_ROOT    = os.path.dirname(os.path.abspath(__file__))
MAIN_PY         = os.path.join(PROJECT_ROOT, "main.py")
SPLASH_TEMPLATE = os.path.join(PROJECT_ROOT, "templates", "splash_template.png")
SPLASH_OUT      = os.path.join(PROJECT_ROOT, "templates", "splash.png")
FONTS_DIR       = os.path.join(PROJECT_ROOT, "templates", "fonts")


def generate_splash(version: str):
    """
    Generates splash.png using Zurich brand fonts and colours.
      Background: Dark Blue #23366F → Zurich Blue #2167AE gradient
      Headline:   "X-Check " in Ogg-Regular (brand emphasis) + "Application" in ZurichSans-Light
      Version:    ZurichSans-Regular
      Loading:    ZurichSans-Light
      All text:   Pure White (only accessible colour on dark blue backgrounds per brand guidelines)
    """
    from PIL import Image, ImageDraw, ImageFont

    W, H = 480, 200
    img  = Image.new("RGB", (W, H))
    draw = ImageDraw.Draw(img)

    # Background gradient: Dark Blue #23366F → Zurich Blue #2167AE
    top_rgb    = (35,  54, 111)
    bottom_rgb = (33, 103, 174)
    for y in range(H):
        t = y / (H - 1)
        draw.line([(0, y), (W - 1, y)], fill=tuple(
            int(top_rgb[i] + (bottom_rgb[i] - top_rgb[i]) * t) for i in range(3)
        ))

    def _font(name, size):
        try:
            return ImageFont.truetype(os.path.join(FONTS_DIR, name), size)
        except OSError:
            return ImageFont.load_default()

    ogg_font     = _font("Ogg-Regular.ttf",        28)
    light_font   = _font("ZurichSans-Light.ttf",   28)
    version_font = _font("ZurichSans-Regular.ttf", 17)
    loading_font = _font("ZurichSans-Light.ttf",   13)

    # Mixed headline: "X-Check " in Ogg (brand emphasis) + "Application" in ZurichSans Light
    # anchor="ls" pins (x, y) to the left-baseline of each segment — true baseline alignment
    t1, t2   = "X-Check ", "Application"
    w1 = draw.textbbox((0, 0), t1, font=ogg_font,   anchor="ls")[2]
    w2 = draw.textbbox((0, 0), t2, font=light_font, anchor="ls")[2]
    x_start  = (W - (w1 + w2)) // 2
    baseline = 80
    draw.text((x_start,        baseline), t1, fill=(255, 255, 255), font=ogg_font,   anchor="ls")
    draw.text((x_start + w1,   baseline), t2, fill=(255, 255, 255), font=light_font, anchor="ls")

    def draw_centered(text, font, y_top):
        bbox = draw.textbbox((0, 0), text, font=font)
        x = (W - (bbox[2] - bbox[0])) // 2
        draw.text((x, y_top - bbox[1]), text, fill=(255, 255, 255), font=font)

    draw_centered(version,      version_font, 105)
    draw_centered("Loading...", loading_font, 155)

    img.save(SPLASH_OUT)
    print(f"  Splash generated: {version}")

def force_remove(func, path, _):
    import stat
    os.chmod(path, stat.S_IWRITE)
    func(path)

def build(config: dict):
    print(f"\n  Building: {config['name']}  (DEBUG_MODE={config['debug_mode']})")

    # --- Generate versioned splash screen ---
    generate_splash(VERSION)

    # --- Read main.py ---
    with open(MAIN_PY, "r", encoding="utf-8") as f:
        source = f.read()

    # --- Patch DEBUG_MODE line ---
    source = re.sub(
        r"DEBUG_MODE\s*=\s*(True|False)",
        f"DEBUG_MODE = {config['debug_mode']}",
        source
    )

    # --- Patch DEBUG_TASK line (debug builds only) ---
    if "debug_task" in config:
        source = re.sub(
            r'DEBUG_TASK\s*=\s*"[^"]*"',
            f'DEBUG_TASK = "{config["debug_task"]}"',
            source
        )

    # --- Write patched temp file (name is unique per build to allow parallel runs) ---
    temp_main = os.path.join(PROJECT_ROOT, f"_main_build_{config['key']}.py")
    with open(temp_main, "w", encoding="utf-8") as f:
        f.write(source)

    # --- Pre-clean build folder before PyInstaller runs ---
    build_folder = os.path.join(PROJECT_ROOT, "build", config["name"])
    if os.path.exists(build_folder):
        print(f"  Pre-cleaning: {build_folder}")
        try:
            shutil.rmtree(build_folder)
        except PermissionError:
            shutil.rmtree(build_folder, onexc=force_remove)

    # --- Generate spec file (enables always_on_top=False on Splash) ---
    datas_entries = [
        f"(r'{os.path.join(PROJECT_ROOT, 'templates')}', 'templates')",
        f"(r'{os.path.join(PROJECT_ROOT, 'version.py')}', '.')",
    ]
    if config["add_test_data"]:
        datas_entries.append(
            f"(r'{os.path.join(PROJECT_ROOT, 'test_data')}', 'test_data')"
        )
    datas_str = "[\n" + ",\n".join(f"        {e}" for e in datas_entries) + "\n    ]"

    # Explicit hidden imports: the lazy `_lazy(module, cls)` factory in
    # task_registry.py defers strategy imports to runtime, so PyInstaller's
    # static analyser cannot see them. List them here so they get bundled.
    # Strategy submodules are loaded lazily from task_registry._lazy(), so
    # PyInstaller's static analyser cannot see them. Each strategy branch
    # extends this list with its own modules. The infra `main` branch ships
    # an empty list because no strategies are registered.
    hidden_imports: list[str] = [
        # Sensitivity labelling via Excel COM (BaseStrategy hook).
        "strategies.sensitivity",
        "win32com.client",
        "win32com",
        "pythoncom",
        "pywintypes",
        # Conditions strategy submodules (loaded lazily via task_registry._lazy).
        "strategies.conditions",
        "strategies.conditions.conditions",
        "strategies.conditions.extract",
        "strategies.conditions.fip",
        "strategies.conditions.compare",
        # Full Run strategy submodules.
        "strategies.full_run",
        "strategies.full_run.full_run",
    ]
    hidden_str = "[" + ", ".join(repr(m) for m in hidden_imports) + "]"

    spec_content = f"""# -*- mode: python ; coding: utf-8 -*-
a = Analysis(
    [r'{temp_main}'],
    pathex=[r'{PROJECT_ROOT}'],
    binaries=[],
    datas={datas_str},
    hiddenimports={hidden_str},
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
splash = Splash(
    r'{SPLASH_OUT}',
    binaries=a.binaries,
    datas=a.datas,
    text_pos=None,
    text_size=12,
    minify_script=True,
    always_on_top=False,
)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    splash,
    splash.binaries,
    [],
    name='{config["name"]}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
"""
    spec_path = os.path.join(PROJECT_ROOT, f"{config['name']}.spec")
    with open(spec_path, "w", encoding="utf-8") as f:
        f.write(spec_content)

    cmd = ["pyinstaller", spec_path]
    print(f"  [BUILD] Command: {' '.join(cmd)}")

    # --- Run PyInstaller ---
    subprocess.run(cmd, check=True)

    # --- Clean up temp file ---
    os.remove(temp_main)

    print(f"  Built successfully: dist\\{config['name']}.exe")


if __name__ == "__main__":
    requested_keys = sys.argv[1:]
    valid_keys = {b["key"] for b in BUILDS} | {"debug"}

    if requested_keys:
        unknown = [k for k in requested_keys if k not in valid_keys]
        if unknown:
            print(f"  Unknown key(s): {unknown}")
            print(f"  Available keys: {sorted(valid_keys)}")
            sys.exit(1)
        selected = []
        for b in BUILDS:
            if "debug" in requested_keys and b["debug_mode"]:
                selected.append(b)
            elif b["key"] in requested_keys:
                selected.append(b)
    else:
        selected = BUILDS

    for build_config in selected:
        build(build_config)

    # --- Clean up build artifacts ---
    print("\n  Cleaning up build folders...")
    for build_config in selected:
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