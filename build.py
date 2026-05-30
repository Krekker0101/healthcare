#!/usr/bin/env python3
"""
Healthcare Application Builder.

Builds launcher, x86/x64 hosts, shared module, and optional installer.
The script keeps the existing csc.exe workflow, but centralizes compiler
settings so release flags and references stay consistent across artifacts.
"""

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path

APP_NAME = "HealthcareSanatoriumInterface"
MODULE_NAME = "HealthcareModules.dll"
DATABASE_NAME = "HealthcareSanatoriumSystem.accdb"
DOCS = ("UserGuide.txt", "UserGuide.rtf", "UserGuide.pdf")
PAYLOAD_FILES = (
    f"{APP_NAME}.exe",
    f"{APP_NAME}.x86.exe",
    f"{APP_NAME}.x64.exe",
    MODULE_NAME,
    DATABASE_NAME,
) + DOCS
COMMON_REFERENCES = (
    "System.dll",
    "System.Core.dll",
    "System.Data.dll",
    "System.Drawing.dll",
    "System.Windows.Forms.dll",
)


def find_csc():
    candidates = [
        r"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
        r"C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path

    for tool in ("csc", "mcs", "mono-csc"):
        path = shutil.which(tool)
        if path:
            return path
    return None


def run(cmd):
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        output = (result.stderr or result.stdout or "Compiler failed").strip()
        raise RuntimeError(output)
    return result


def copy_if_exists(src, dst):
    src = Path(src)
    if src.exists():
        Path(dst).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
        return True
    return False


def compiler_args(csc, target, output, sources, icon=None, platform=None, references=COMMON_REFERENCES, resources=()):
    args = [
        csc,
        "/nologo",
        f"/target:{target}",
        "/optimize+",
        "/codepage:65001",
        "/warn:4",
        "/nowarn:1591",
        "/out:" + str(output),
    ]
    if platform:
        args.append(f"/platform:{platform}")
    if icon:
        args.append("/win32icon:" + str(icon))
    args.extend("/reference:" + reference for reference in references)
    args.extend(resources)
    args.extend(str(source) for source in sources)
    return args


def compile_artifact(csc, label, target, output, sources, icon=None, platform=None, resources=(), references=COMMON_REFERENCES):
    print(label)
    run(compiler_args(csc, target, output, sources, icon=icon, platform=platform, resources=resources, references=references))


def prepare_assets(root, build):
    icon_src = root / "assest" / "app_icon.ico"
    db_src = root / "assest" / "db.accdb"
    db_fixed = build / "HealthcareSanatoriumSystem_FIXED.accdb"

    if not icon_src.exists():
        raise FileNotFoundError("app_icon.ico not found")

    shutil.copy2(icon_src, build / "app_icon.ico")
    db_target = build / DATABASE_NAME
    if db_src.exists():
        shutil.copy2(db_src, db_target)
    elif db_fixed.exists():
        shutil.copy2(db_fixed, db_target)
    else:
        raise FileNotFoundError("Access database not found")

    for doc in DOCS:
        copy_if_exists(root / "docs" / doc, build / doc)

    return icon_src


def copy_payload(build, payload):
    copied = []
    for file_name in PAYLOAD_FILES:
        if copy_if_exists(build / file_name, payload / file_name):
            copied.append(file_name)
    return copied


def parse_args(argv):
    parser = argparse.ArgumentParser(description="Build Healthcare Sanatorium Windows Forms package.")
    parser.add_argument("--clean", action="store_true", help="remove build/dist output before compiling")
    parser.add_argument("--no-installer", action="store_true", help="skip optional installer compilation")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv or sys.argv[1:])
    root = Path(__file__).resolve().parent
    os.chdir(root)

    print("\n" + "=" * 54)
    print("Healthcare Application Builder v3.1")
    print("=" * 54 + "\n")

    csc = find_csc()
    if not csc:
        print("ERROR: C# compiler not found. Install .NET Framework 4.x")
        return 1

    build = root / "build"
    dist = root / "dist"
    payload = dist / "payload"

    if args.clean:
        shutil.rmtree(build, ignore_errors=True)
        shutil.rmtree(dist, ignore_errors=True)

    build.mkdir(parents=True, exist_ok=True)
    payload.mkdir(parents=True, exist_ok=True)

    icon_src = prepare_assets(root, build)

    compile_artifact(
        csc,
        "[1/5] Compiling shared module...",
        "library",
        build / MODULE_NAME,
        [root / "src" / "HealthcareInterface.cs"],
    )
    compile_artifact(
        csc,
        "[2/5] Compiling x86 host...",
        "winexe",
        build / f"{APP_NAME}.x86.exe",
        [root / "src" / "HostProgram.cs"],
        icon=icon_src,
        platform="x86",
    )
    compile_artifact(
        csc,
        "[3/5] Compiling x64 host...",
        "winexe",
        build / f"{APP_NAME}.x64.exe",
        [root / "src" / "HostProgram.cs"],
        icon=icon_src,
        platform="x64",
    )
    compile_artifact(
        csc,
        "[4/5] Compiling launcher...",
        "winexe",
        build / f"{APP_NAME}.exe",
        [root / "src" / "Launcher.cs"],
        icon=icon_src,
    )

    print("[5/5] Copying docs and packaging installer...")
    copied_payload = copy_payload(build, payload)
    setup_resources = [f"/resource:{payload / name},{name}" for name in copied_payload]

    setup_host = root / "installer" / "SetupHost.cs"
    installer_built = False
    if not args.no_installer and setup_host.exists():
        compile_artifact(
            csc,
            "      Compiling installer...",
            "winexe",
            dist / "Setup.exe",
            [setup_host],
            icon=icon_src,
            resources=setup_resources,
            references=("System.dll", "System.Core.dll", "System.Drawing.dll", "System.Windows.Forms.dll"),
        )
        installer_built = True
    elif args.no_installer:
        print("INFO: installer build skipped by --no-installer.")
    else:
        print("INFO: installer/SetupHost.cs not found; installer build skipped.")

    shutil.rmtree(payload, ignore_errors=True)

    print("\n" + "=" * 54)
    print("BUILD SUCCESSFUL!")
    print("=" * 54)
    print(f"Launcher:   {build / (APP_NAME + '.exe')}")
    print(f"x86 host:   {build / (APP_NAME + '.x86.exe')}")
    print(f"x64 host:   {build / (APP_NAME + '.x64.exe')}")
    if installer_built:
        print(f"Installer:  {dist / 'Setup.exe'}")
    else:
        print("Installer:  skipped")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print("ERROR:", exc)
        sys.exit(1)
