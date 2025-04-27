# Oblivion Remastered Simple Mod Manager (Python)

A basic Python script to manage ESP and Pak mods for Oblivion Remastered, primarily handling installations from archive files.

## Features

*   **Install/Uninstall ESP Mods:** Handles `.esp` files, automatically updating `plugins.txt`.
*   **Install/Uninstall Pak Mods:** Handles `.pak`/`.ucas`/`.utoc` file sets, installing them to the correct `Paks/~mods` folder.
*   **Archive Support:** Processes `.zip`, `.rar`, and `.7z` mod archives.
*   **Load Order Management:** Allows interactive reordering of custom ESP plugins in `plugins.txt`.
*   **Drag & Drop:** Supports dragging archive files onto the script/shortcut for batch installation.

## Requirements

*   **Python 3.x:** Download from [python.org](https://www.python.org/)
*   **Windows OS:** Tested on Windows 10/11.
*   **Python Packages:**
    *   `rarfile`: Run `pip install rarfile`
    *   `py7zr`: Run `pip install py7zr`
*   **`unrar` Executable (for .rar support):**
    *   Download the command-line `unrar.exe` from [RARLAB (Other downloads section)](https://www.rarlab.com/rar_add.htm).
    *   Place `unrar.exe` somewhere in your system's PATH (e.g., `C:\Windows`) or in the same directory as the Python script.

## Installation & Setup

1.  **Download:** Get the Python script.
2.  **Install Dependencies:**
    ```bash
    pip install rarfile py7zr
    ```
3.  **Install `unrar`:** Download `unrar.exe` from RARLAB and ensure it's accessible via your PATH or alongside the script.
4.  **!!! Configure Game Path !!!**
    *   Open the downloaded `.py` script in a text editor (like Notepad++).
    *   Find the line `GAME_BASE_PATH = r"C:\..."` near the top.
    *   **Change the path** within the quotes (`r"..."`) to match your **exact** installation directory for Oblivion Remastered (e.g., where the game's main executable is located).
    *   Save the script.

## Usage

1.  **Run Interactively:**
    *   Double-click the `.py` file or run `python your_script_name.py` in the terminal.
    *   Follow the on-screen menu:
        *   `(I)nstall`: Prompts for an archive path to install a mod.
        *   `Uninstall (E)SP`: Lists installed custom ESPs for removal.
        *   `Uninstall (P)ak`: Lists installed Pak mods for removal.
        *   `Change (L)oad Order`: Interactively reorder custom ESPs.
        *   `(Q)uit`: Exits the script.

2.  **Install via Drag & Drop (Recommended for Installation):**
    *   Create a shortcut to the `.py` script (Right-click > Send to > Desktop (create shortcut)).
    *   Drag one or more mod archive files (`.zip`, `.rar`, `.7z`) onto the **shortcut icon**.
    *   The script will attempt to install each dropped archive sequentially.

3.  **Install via Command Line:**
    *   Open Command Prompt/PowerShell in the script's directory.
    *   Run: `python your_script_name.py "C:\path\to\mod1.zip" "D:\mods\mod2.rar"`
    *   The script will process each listed archive.

## Important Notes

*   **Administrator Privileges:** The script requires admin rights to write to the game's installation directory (usually `Program Files`). It will try to prompt for elevation (UAC). If this fails, you must run the script or shortcut explicitly as an Administrator.
*   **`unrar.exe`:** Support for `.rar` files **strictly requires** the external `unrar.exe` to be installed and accessible.
*   **Game Path Configuration:** The script **will not work** if `GAME_BASE_PATH` is incorrect. Double-check it!
*   **Simplicity:** This is a basic tool, not a replacement for advanced mod managers like Mod Organizer 2 or Vortex.
*   **Backups:** Always back up your game's `Data` folder and `plugins.txt` before extensive modding.
