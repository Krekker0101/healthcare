
#!/usr/bin/env python3
"""
Healthcare Application Builder
Builds launcher, x86/x64 hosts, shared module, and installer.
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path


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
        raise RuntimeError(result.stderr or result.stdout or "Compiler failed")
    return result


def copy_if_exists(src, dst):
    if os.path.exists(src):
        shutil.copy2(src, dst)
        return True
    return False


def main():
    root = Path(__file__).resolve().parent
    os.chdir(root)

    print("\n" + "=" * 54)
    print("Healthcare Application Builder v3.0")
    print("=" * 54 + "\n")

    csc = find_csc()
    if not csc:
        print("ERROR: C# compiler not found. Install .NET Framework 4.x")
        return 1

    build = root / "build"
    dist = root / "dist"
    payload = dist / "payload"

    build.mkdir(parents=True, exist_ok=True)
    payload.mkdir(parents=True, exist_ok=True)

    icon_src = root / "assest" / "app_icon.ico"
    db_src = root / "assest" / "db.accdb"
    db_fixed = build / "HealthcareSanatoriumSystem_FIXED.accdb"

    if not icon_src.exists():
        print("ERROR: app_icon.ico not found")
        return 1

    shutil.copy2(icon_src, build / "app_icon.ico")
    db_target = build / "HealthcareSanatoriumSystem.accdb"
    if db_src.exists():
        shutil.copy2(db_src, db_target)
    elif db_fixed.exists():
        shutil.copy2(db_fixed, db_target)
    else:
        print("ERROR: Access database not found")
        return 1

    print("[1/5] Compiling shared module...")
    run([
        csc, "/nologo", "/target:library", "/optimize+", "/codepage:65001",
        "/out:" + str(build / "HealthcareModules.dll"),
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Data.dll",
        "/reference:System.Drawing.dll",
        "/reference:System.Windows.Forms.dll",
        str(root / "src" / "HealthcareInterface.cs"),
    ])

    print("[2/5] Compiling x86 host...")
    run([
        csc, "/nologo", "/platform:x86", "/target:winexe", "/optimize+", "/codepage:65001",
        "/out:" + str(build / "HealthcareSanatoriumInterface.x86.exe"),
        "/win32icon:" + str(icon_src),
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Data.dll",
        "/reference:System.Drawing.dll",
        "/reference:System.Windows.Forms.dll",
        str(root / "src" / "HostProgram.cs"),
    ])

    print("[3/5] Compiling x64 host...")
    run([
        csc, "/nologo", "/platform:x64", "/target:winexe", "/optimize+", "/codepage:65001",
        "/out:" + str(build / "HealthcareSanatoriumInterface.x64.exe"),
        "/win32icon:" + str(icon_src),
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Data.dll",
        "/reference:System.Drawing.dll",
        "/reference:System.Windows.Forms.dll",
        str(root / "src" / "HostProgram.cs"),
    ])

    print("[4/5] Compiling launcher...")
    run([
        csc, "/nologo", "/target:winexe", "/optimize+", "/codepage:65001",
        "/out:" + str(build / "HealthcareSanatoriumInterface.exe"),
        "/win32icon:" + str(icon_src),
        "/reference:System.dll",
        "/reference:System.Core.dll",
        "/reference:System.Data.dll",
        "/reference:System.Drawing.dll",
        "/reference:System.Windows.Forms.dll",
        str(root / "src" / "Launcher.cs"),
    ])

    print("[5/5] Copying docs and packaging installer...")
    for doc in ("UserGuide.txt", "UserGuide.rtf", "UserGuide.pdf"):
        src = root / "docs" / doc
        if src.exists():
            shutil.copy2(src, build / doc)

    for file_name in (
        "HealthcareSanatoriumInterface.exe",
        "HealthcareSanatoriumInterface.x86.exe",
        "HealthcareSanatoriumInterface.x64.exe",
        "HealthcareModules.dll",
        "HealthcareSanatoriumSystem.accdb",
        "UserGuide.txt",
        "UserGuide.rtf",
        "UserGuide.pdf",
    ):
        src = build / file_name
        if src.exists():
            shutil.copy2(src, payload / file_name)

    setup_resources = [f"/resource:{payload / f},{f}" for f in (
        "HealthcareSanatoriumInterface.exe",
        "HealthcareSanatoriumInterface.x86.exe",
        "HealthcareSanatoriumInterface.x64.exe",
        "HealthcareModules.dll",
        "HealthcareSanatoriumSystem.accdb",
        "UserGuide.txt",
        "UserGuide.rtf",
        "UserGuide.pdf",
    ) if (payload / f).exists()]

    setup_host = root / "installer" / "SetupHost.cs"
    installer_built = False
    if setup_host.exists():
        run([
            csc, "/nologo", "/target:winexe", "/optimize+", "/codepage:65001",
            "/out:" + str(dist / "Setup.exe"),
            "/win32icon:" + str(icon_src),
            "/reference:System.dll",
            "/reference:System.Core.dll",
            "/reference:System.Drawing.dll",
            "/reference:System.Windows.Forms.dll",
        ] + setup_resources + [str(setup_host)])
        installer_built = True
    else:
        print("WARNING: installer/SetupHost.cs not found; installer build skipped.")

    shutil.rmtree(payload, ignore_errors=True)

    print("\n" + "=" * 54)
    print("BUILD SUCCESSFUL!")
    print("=" * 54)
    print(f"Launcher:   {build / 'HealthcareSanatoriumInterface.exe'}")
    print(f"x86 host:   {build / 'HealthcareSanatoriumInterface.x86.exe'}")
    print(f"x64 host:   {build / 'HealthcareSanatoriumInterface.x64.exe'}")
    if installer_built:
        print(f"Installer:   {dist / 'Setup.exe'}")
    else:
        print("Installer:   skipped (installer/SetupHost.cs not found)")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        print("ERROR:", exc)
        sys.exit(1)
