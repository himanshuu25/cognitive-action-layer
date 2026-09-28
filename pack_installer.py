import os
import sys
import shutil
import json
import time
from pathlib import Path
import customtkinter as ctk
import psutil
import threading

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class InstallerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("CAL Smart Installer")
        self.geometry("500x350")
        self.resizable(False, False)

        # Title
        self.title_label = ctk.CTkLabel(self, text="Cognitive Action Layer Setup", font=ctk.CTkFont(size=20, weight="bold"))
        self.title_label.pack(pady=(20, 5))

        self.subtitle_label = ctk.CTkLabel(self, text="Select your AI Assistant to configure the Excel MCP Server", text_color="gray")
        self.subtitle_label.pack(pady=(0, 20))

        # Buttons Frame
        self.buttons_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.buttons_frame.pack(pady=10)

        self.btn_antigravity = ctk.CTkButton(
            self.buttons_frame, 
            text="Install for Antigravity", 
            width=200, 
            height=40,
            command=lambda: self.start_install("antigravity")
        )
        self.btn_antigravity.pack(side="left", padx=10)

        self.btn_claude = ctk.CTkButton(
            self.buttons_frame, 
            text="Install for Claude Desktop", 
            width=200, 
            height=40,
            command=lambda: self.start_install("claude")
        )
        self.btn_claude.pack(side="left", padx=10)

        # Status textbox
        self.status_box = ctk.CTkTextbox(self, width=460, height=120, state="disabled")
        self.status_box.pack(pady=20)

    def log(self, message):
        self.status_box.configure(state="normal")
        self.status_box.insert("end", message + "\n")
        self.status_box.see("end")
        self.status_box.configure(state="disabled")
        self.update()

    def check_and_kill_process(self, process_names):
        """Checks if app is running and attempts to kill it."""
        found = False
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'] and proc.info['name'].lower() in [name.lower() for name in process_names]:
                    if not found:
                        self.log(f"[*] Found running process: {proc.info['name']}")
                        self.log("[*] Attempting to close it to avoid file locks...")
                        found = True
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        
        if found:
            time.sleep(2) # Give it time to close completely
            self.log("[+] Process closed.")
        return True

    def install_bundled_exe(self) -> Path:
        """Copies the bundled excel_mcp.exe to the permanent location."""
        try:
            base_path = Path(sys._MEIPASS)
        except AttributeError:
            base_path = Path(__file__).parent

        bundled_exe = base_path / "excel_mcp.exe"
        if not bundled_exe.exists():
            for fallback in [Path("dist/excel_mcp.exe"), Path("excel_mcp.exe")]:
                if fallback.exists():
                    bundled_exe = fallback
                    break

        if not bundled_exe.exists():
            raise FileNotFoundError("excel_mcp.exe was not found inside the installer.")

        target_dir = Path(os.environ.get("USERPROFILE", "")) / "AppData" / "Local" / "Programs" / "ExcelMCP"
        target_exe = target_dir / "excel_mcp.exe"

        self.log(f"[*] Copying server to: {target_dir}")
        target_dir.mkdir(parents=True, exist_ok=True)
        shutil.copy2(bundled_exe, target_exe)
        return target_exe

    def configure_json(self, config_path: Path, exe_path: Path):
        """Reads, updates, and saves the MCP config JSON safely."""
        config_path.parent.mkdir(parents=True, exist_ok=True)

        config = {}
        if config_path.exists():
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
            except Exception as e:
                self.log(f"[!] Warning: existing config was invalid JSON, creating new. ({e})")
                config = {}

        if "mcpServers" not in config or not isinstance(config["mcpServers"], dict):
            config["mcpServers"] = {}

        config["mcpServers"]["excel-mcp"] = {
            "command": str(exe_path),
            "args": []
        }

        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        self.log(f"[+] Successfully updated: {config_path.name}")

    def start_install(self, target):
        # Disable buttons during install
        self.btn_antigravity.configure(state="disabled")
        self.btn_claude.configure(state="disabled")
        self.status_box.configure(state="normal")
        self.status_box.delete("0.0", "end")
        self.status_box.configure(state="disabled")
        
        # Run install in background thread to keep UI responsive
        threading.Thread(target=self._run_install, args=(target,), daemon=True).start()

    def _run_install(self, target):
        try:
            self.log(f"--- Starting Installation for {target.capitalize()} ---")
            
            if target == "claude":
                self.check_and_kill_process(["claude.exe"])
                config_paths = self.get_claude_paths()
            else:
                self.check_and_kill_process(["antigravity.exe", "agy.exe"])
                config_paths = self.get_antigravity_paths()

            if not config_paths:
                self.log(f"[X] Error: Could not find configuration folders for {target.capitalize()}. Is it installed?")
                return

            # Install EXE
            self.log("[*] Extracting MCP server executable...")
            target_exe = self.install_bundled_exe()
            self.log("[+] Executable installed successfully.")

            # Update Configs
            self.log("[*] Registering server with client...")
            success = 0
            for path in config_paths:
                try:
                    self.configure_json(path, target_exe)
                    success += 1
                except Exception as e:
                    self.log(f"[X] Failed to update {path.name}: {str(e)}")

            if success > 0:
                self.log("\n✅ ALL DONE! The MCP server is now installed.")
                self.log("Please restart your AI Assistant to see the changes.")
            else:
                self.log("\n❌ Installation failed. Please check permissions.")

        except Exception as e:
            self.log(f"\n[X] Critical Error: {str(e)}")
        finally:
            self.btn_antigravity.configure(state="normal")
            self.btn_claude.configure(state="normal")

    def get_claude_paths(self):
        paths = []
        appdata = os.environ.get("APPDATA", "")
        localappdata = os.environ.get("LOCALAPPDATA", "")
        if appdata:
            paths.append(Path(appdata) / "Claude" / "claude_desktop_config.json")
        if localappdata:
            packages_dir = Path(localappdata) / "Packages"
            if packages_dir.exists():
                for match in packages_dir.glob("Claude_*"):
                    paths.append(match / "LocalCache" / "Roaming" / "Claude" / "claude_desktop_config.json")
        return paths

    def get_antigravity_paths(self):
        # Global Antigravity Config
        home = Path.home()
        paths = [home / ".gemini" / "config" / "mcp_config.json"]
        return paths

if __name__ == "__main__":
    app = InstallerApp()
    app.mainloop()
