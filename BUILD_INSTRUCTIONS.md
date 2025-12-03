# Building Rash Manager v2 as Windows or macOS Application

This guide explains how to convert the Python program to standalone desktop binaries using PyInstaller.

---

## Windows (.exe) Build

## Prerequisites

1. **Windows PC** (required - PyInstaller builds executables for the platform it runs on)
2. **Python 3.8 or higher** installed on Windows
3. **Internet connection** (to download dependencies)

## Step-by-Step Instructions

### Step 1: Prepare Your Windows Machine

1. Install Python 3.8+ from [python.org](https://www.python.org/downloads/)
   - ⚠️ **Important**: Check "Add Python to PATH" during installation
2. Verify installation by opening Command Prompt and typing:
   ```cmd
   python --version
   ```
   You should see something like `Python 3.11.x`

### Step 2: Copy Project Files

1. Copy the entire project folder to your Windows machine
2. Open Command Prompt (cmd) or PowerShell
3. Navigate to the project folder:
   ```cmd
   cd path\to\Rash_mainnnnn
   ```

### Step 3: Install Dependencies

1. Upgrade pip:
   ```cmd
   python -m pip install --upgrade pip
   ```

2. Install all required packages:
   ```cmd
   pip install -r requirements.txt
   ```

   This will install:
   - pandas, numpy, scipy
   - opencv-python, Pillow
   - pyzbar, qrcode
   - openpyxl, pdf2image
   - pyinstaller

### Step 4: Install Poppler (for PDF processing)

`pdf2image` requires Poppler binaries to process PDF files:

1. Download Poppler for Windows:
   - Visit: https://github.com/oschwartz10612/poppler-windows/releases
   - Download the latest `Release-X.XX.X-X.zip` file
   - Extract it (e.g., to `C:\poppler`)

2. **Option A: Add to PATH** (recommended)
   - Add `C:\poppler\Library\bin` to your Windows PATH environment variable
   - Restart Command Prompt

   **Option B: Copy to exe folder** (easier for distribution)
   - After building, copy the contents of `poppler\Library\bin` folder to the same folder as your `.exe`

### Step 5: Build the Executable

**Easy Method (using batch script):**
```cmd
build_exe.bat
```

**Manual Method:**
```cmd
pyinstaller build_exe.spec --clean --noconfirm
```

The build process may take 5-15 minutes. Wait for it to complete.

### Step 6: Locate Your Executable

After successful build, your executable will be at:
```
dist\RashManager.exe
```

### Step 7: Prepare Distribution Folder

Create a folder for distribution with these files:

```
RashManager/
├── RashManager.exe
├── titul_bubble_koordinatalar_2480x3508.xlsx  (required)
├── Titul.pdf                                  (optional, if available)
└── poppler/                                   (if not in PATH)
    └── bin/
        ├── pdftoppm.exe
        ├── pdfinfo.exe
        └── ... (other poppler DLLs)
```

**Important Notes:**
- The `.exe` file is standalone but requires the Excel coordinate file
- PDF processing requires Poppler binaries in the same folder or in PATH
- The program will create `output/` and `test_keys/` folders automatically
- On first run, it will ask for the keyword "Mirkomil12"

## Troubleshooting

### Problem: "python is not recognized"
**Solution**: Python is not in PATH. Reinstall Python and check "Add Python to PATH"

### Problem: "ModuleNotFoundError" during build
**Solution**: Install missing module with `pip install <module-name>`

### Problem: PDF processing doesn't work
**Solution**: 
1. Ensure Poppler is installed (see Step 4)
2. Copy poppler binaries to the same folder as the `.exe`
3. Or add poppler to your PATH environment variable

### Problem: Executable is very large (200-500 MB)
**Solution**: This is normal. PyInstaller bundles Python and all libraries. You can use `--onefile` in the spec file, but it will be slower to start.

### Problem: Antivirus flags the exe as suspicious
**Solution**: PyInstaller executables are sometimes flagged. Add an exception or use code signing (advanced).

### Problem: GUI doesn't appear / crashes on startup
**Solution**: 
- Check if all data files are in the same folder as the exe
- Try running from Command Prompt to see error messages
- Modify `build_exe.spec` and set `console=True` to see debug output

---

# macOS (.app) Build

## Prerequisites

1. **Mac running macOS 12+** (builds must be created on macOS)
2. **Python 3.8 or higher** (`python3 --version` should work in Terminal)
3. **Xcode Command Line Tools** (for Homebrew) — install via `xcode-select --install`
4. **Homebrew** (optional but recommended) — https://brew.sh

## Step-by-Step Instructions

### Step 1: Prepare your Mac

1. Install Python 3.8+ (via [python.org](https://www.python.org/downloads/macos/) or Homebrew)
2. Confirm Python works:
   ```bash
   python3 --version
   ```

### Step 2: Copy Project Files

```bash
cd /path/to/Rash_mainnnnn
```

### Step 3: Install Dependencies

```bash
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt
```

### Step 4: Install Poppler for macOS

```bash
brew install poppler
```

Poppler binaries will live in `/opt/homebrew/bin` (Apple Silicon) or `/usr/local/bin`. `pdf2image` locates them automatically when they are on `PATH`.

### Step 5: Build the App Bundle

**Easy Method (using shell script):**

```bash
chmod +x build_macos.sh
./build_macos.sh
```

**Manual Method:**

```bash
python3 -m PyInstaller build_exe.spec --clean --noconfirm
```

### Step 6: Locate Your App

```
dist/RashManager.app
```

Inside the bundle, the executable is at `dist/RashManager.app/Contents/MacOS/RashManager`.

### Step 7: Prepare Distribution Folder

```
RashManager-mac/
├── RashManager.app
├── titul_bubble_koordinatalar_2480x3508.xlsx  (required)
├── Titul.pdf                                  (optional)
└── README.txt / instructions for users
```

**Optional:** If target Macs do not have Poppler installed, bundle the Poppler binaries or instruct users to run `brew install poppler`.

### Dealing with Gatekeeper

- Unsigned apps may trigger “App is from an unidentified developer.” Users can right-click the app, choose **Open**, then confirm.
- For distributing widely, consider codesigning and notarizing (advanced, not required for internal use).

### Troubleshooting (macOS)

- **“python3: command not found”** → Install Python 3 and ensure `/usr/local/bin` or `/opt/homebrew/bin` is on PATH.
- **`ModuleNotFoundError`** → Re-run `pip install -r requirements.txt`.
- **PDF pages not detected** → Confirm `poppler` is installed (`brew install poppler`).
- **App is quarantined** → Run `xattr -dr com.apple.quarantine RashManager.app`.
- **App crashes silently** → Rebuild with console enabled by setting `console=True` in `build_exe.spec` to view logs.

---

## Advanced Options

### Create a Single-File Executable

Edit `build_exe.spec` and add `--onefile` to the EXE options, or modify the spec file.

### Add an Icon

1. Create or download an `.ico` file
2. Edit `build_exe.spec` and set:
   ```python
   icon='path/to/icon.ico'
   ```

### Reduce File Size

You can exclude unused modules by editing `build_exe.spec`:
```python
excludes=['matplotlib', 'IPython', 'jupyter', ...]
```

## Distribution

To distribute the application:
1. Zip the entire `RashManager/` folder (including the `.exe` and required files)
2. Users can extract and run `RashManager.exe` directly
3. No Python installation required on target machines!

## Testing

Before distributing:
1. Test the `.exe` on a Windows machine without Python installed
2. Verify all features work (image checking, PDF processing, etc.)
3. Check that the keyword prompt appears on first run
4. Ensure output files are created correctly

## Support

If you encounter issues:
- Check the PyInstaller documentation: https://pyinstaller.org/
- Review error messages in the build output
- Ensure all dependencies are correctly installed

