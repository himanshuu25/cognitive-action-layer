"""
pack_installer.py
Automated Python installer compiled into install.exe.
Bundles excel_mcp.exe inside and auto-configures Claude Desktop.
Supports:
  - Standard Claude Desktop installer (APPDATA/Claude/)
  - Windows Store Claude app (LOCALAPPDATA/Packages/Claude_*/LocalCache/Roaming/Claude/)
Dynamically discovers the Claude Packages folder so it works on any machine.
"""

import os
import sys
import shutil
import json
import glob
from pathlib import Path


# ── Helpers ────────────────────────────────────────────────────────────────────

def update_claude_json(path: Path, exe_path: Path) -> bool:
    """
    Safely reads, merges, and writes the Claude Desktop JSON config.
    Returns True on success, False on failure.
    """
    try:
        path.parent.mkdir(parents=True, exist_ok=True)

        config = {}
        if path.exists():
            try:
                with open(path, "r", encoding="utf-8") as f:
                    config = json.load(f)
            except Exception:
                config = {}

        if "mcpServers" not in config or not isinstance(config["mcpServers"], dict):
            config["mcpServers"] = {}

        config["mcpServers"]["excel-mcp"] = {
            "command": str(exe_path),
            "args": []
        }

        with open(path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)

        print(f"    [OK] Configured: {path}")
        return True

    except Exception as e:
        print(f"    [!!] Could not write to {path}: {e}")
        return False


def find_claude_config_paths() -> list:
    """
    Returns all Claude Desktop config file paths found on this machine.
    Searches both standard and Windows Store (AppX) install locations.
    """
    paths = []
    appdata = os.environ.get("APPDATA", "")
    localappdata = os.environ.get("LOCALAPPDATA", "")

    # 1. Standard Desktop Installer
    if appdata:
        paths.append(Path(appdata) / "Claude" / "claude_desktop_config.json")

    # 2. Windows Store AppX — dynamic discovery (package ID varies per machine)
    if localappdata:
        packages_dir = Path(localappdata) / "Packages"
        if packages_dir.exists():
            # Search for any folder starting with "Claude_"
            matches = list(packages_dir.glob("Claude_*"))
            for match in matches:
                candidate = match / "LocalCache" / "Roaming" / "Claude" / "claude_desktop_config.json"
                paths.append(candidate)

    return paths


# ── Main ────────────────────────────────────────────────────────────────────────

def main():
    print("=" * 56)
    print("  Excel Automation MCP Server — Installer")
    print("=" * 56)
    print()

    # ── Step 1: Locate bundled excel_mcp.exe ──────────────────
    try:
        base_path = Path(sys._MEIPASS)   # PyInstaller temp dir
    except AttributeError:
        base_path = Path(__file__).parent

    bundled_exe = base_path / "excel_mcp.exe"
    if not bundled_exe.exists():
        # Developer fallback
        for fallback in [Path("dist/excel_mcp.exe"), Path("excel_mcp.exe")]:
            if fallback.exists():
                bundled_exe = fallback
                break

    if not bundled_exe.exists():
        print("[ERROR] excel_mcp.exe was not found inside the installer.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # ── Step 2: Copy to permanent install location ─────────────
    target_dir = Path(os.environ["USERPROFILE"]) / "AppData" / "Local" / "Programs" / "ExcelMCP"
    target_exe = target_dir / "excel_mcp.exe"

    print(f"[1/3] Installing excel_mcp.exe ...")
    print(f"      Destination: {target_exe}")
    target_dir.mkdir(parents=True, exist_ok=True)

    try:
        shutil.copy2(bundled_exe, target_exe)
        print("      Successfully installed.")
    except Exception as e:
        print(f"\n[ERROR] Failed to copy excel_mcp.exe: {e}")
        input("\nPress Enter to exit...")
        sys.exit(1)

    print()

    # ── Step 3: Discover all Claude config paths ───────────────
    print("[2/3] Searching for Claude Desktop installations ...")
    config_paths = find_claude_config_paths()

    if not config_paths:
        print("      [!!] No Claude Desktop installation found.")
        print("      Please install Claude Desktop and try again.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    for p in config_paths:
        if p.exists():
            print(f"      Found existing config: {p}")
        else:
            print(f"      Will create new config: {p}")

    print()

    # ── Step 4: Update every config found ─────────────────────
    print("[3/3] Configuring Claude Desktop ...")
    success_count = 0
    for config_path in config_paths:
        if update_claude_json(config_path, target_exe):
            success_count += 1

    print()

    if success_count == 0:
        print("[ERROR] Could not update any Claude Desktop configuration.")
        print("        Please check folder permissions and try again.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    # ── Done ───────────────────────────────────────────────────
    print("=" * 56)
    print("  Installation Completed Successfully!")
    print("=" * 56)
    print()
    print("  Next step: Fully quit and restart Claude Desktop.")
    print("  Then you will see 'excel-mcp' in your MCP servers.")
    print()
    input("Press Enter to close this window...")


if __name__ == "__main__":
    main()
