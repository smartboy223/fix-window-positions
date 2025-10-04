# 🪟 Fix-WindowPositions

A simple but powerful PowerShell + Batch tool to bring back **off-screen or lost windows** into your desktop view.  
Perfect for multi-monitor setups where apps sometimes open **outside the visible area** after disconnecting a monitor, changing resolution, or using remote desktop.

---

## 🚀 Why You Need This Tool
- 🖥️ **Multi-monitor setups** often cause windows to get stuck outside the screen.  
- 🎮 **Gamers & streamers** switching resolutions may find apps “missing.”  
- 💼 **Remote workers** who switch between laptops and external monitors lose track of windows.  
- 🔧 This script automatically **detects & repositions windows** back to the visible area.

---

## 📂 Files in Repo
- `Fix-WindowPositions.ps1` → The main PowerShell script.  
- `run.bat` → A quick way to run the script without typing commands.  

---

## ⚡ How to Use

### 🔹 Option 1: Run with PowerShell
1. Open **PowerShell** as Administrator.  
2. Run:
   .\Fix-WindowPositions.ps1
3. To test without applying changes (dry run):
   .\Fix-WindowPositions.ps1 -DryRun

---

### 🔹 Option 2: Run with Batch File (Recommended 🚀)
1. Just **double-click** `run.bat`.  
2. It will automatically launch the PowerShell script with proper permissions.  
3. Press any key when done.

---

## 📝 Notes
- Safe to use on Windows 10/11.  
- Doesn’t modify registry or system files—just moves misplaced windows.  
- Great helper for productivity when windows "disappear."

---

## 📜 License
MIT License – free to use and share.
