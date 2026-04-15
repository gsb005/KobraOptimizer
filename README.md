# 🐍 KobraOptimizer  
### A modular PowerShell + WPF system optimization tool with a neon UI

KobraOptimizer is a Windows optimization utility built with **PowerShell**, **WPF**, and a fully modular architecture.  
It provides a clean neon‑themed interface and a collection of system‑tuning modules designed to improve performance, reduce clutter, and streamline Windows behavior.

---

## ✨ Features

- ⚡ **Modular architecture** — each optimization lives in its own PowerShell module  
- 🖥️ **Neon WPF UI** — modern, glowing interface  
- 🚀 **System cleanup tools** — temp files, logs, caches, browser junk  
- 🌐 **Network tweaks** — DNS, latency, connectivity improvements  
- 🧹 **Startup manager** — disable unnecessary startup apps  
- 🛠️ **OEM bloat removal** — remove manufacturer‑installed junk  
- 🌍 **Browser cleanup** — Chrome, Edge, Firefox  
- 📦 **Portable build** — no installation required  
- 🔧 **EXE builder script** — package your tool into a standalone executable  

---

## 📁 Project Structure

KobraOptimizer/
│
├── Assets/
├── Logs/
├── Modules/
│   ├── Kobra_Browsers.psm1
│   ├── Kobra_Cleanup.psm1
│   ├── Kobra_Network.psm1
│   ├── Kobra_OEM.psm1
│   └── Kobra_Startup.psm1
│
├── Main.ps1
├── Kobra_UI.xaml
├── Build_KobraExe.ps1
├── Launch_Kobra.cmd
└── Release_Notes_v1_5_1.md

Code

---

## 🛠️ How to Run

1. Download the latest release (coming soon)  
2. Extract the ZIP
3. Open the folder:KobraOptimizer_*release_number
4. Run:Launch_Kobra.cmd (As Administrator)

Launch_Kobra.cmd

Code

or directly:

powershell.exe -ExecutionPolicy Bypass -File .\Main.ps1

Code

---

## ❤️ Support Development

If you enjoy this tool and want to support future updates:

👉 **https://ko-fi.com/kobraoptimizer**

Your support helps keep the project alive.

---

## 📜 License

This project is licensed under the **MIT License**.  
See the `LICENSE` file for details.

---

## 📣 Credits

Created by **KobraOptimizer**  
Built with PowerShell, WPF, and a lot of neon.
