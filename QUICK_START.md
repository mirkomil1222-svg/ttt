# Quick Start - Building Desktop Apps

## Windows (.exe)

1. **Copy project folder to Windows PC**

2. **Open Command Prompt** in the project folder

3. **Run the build script:**
   ```cmd
   build_exe.bat
   ```

4. **Wait 5-15 minutes** for build to complete

5. **Find your executable:**
   - Location: `dist\RashManager.exe`
   - Size: ~200-500 MB (this is normal)

6. **Before using:**
   - Copy `titul_bubble_koordinatalar_2480x3508.xlsx` to the same folder as the exe
   - Install Poppler for PDF processing (see BUILD_INSTRUCTIONS.md)

## macOS (.app)

1. **Copy project folder to your Mac**

2. **Open Terminal** in the project directory

3. **Run the build script:**
   ```bash
   chmod +x build_macos.sh
   ./build_macos.sh
   ```

4. **Wait 5-15 minutes** for build to complete

5. **Find your app bundle:**
   - Location: `dist/RashManager.app`

6. **Before distributing:**
   - Copy `titul_bubble_koordinatalar_2480x3508.xlsx` next to the app
   - Ensure Poppler is installed on the target Mac (`brew install poppler`)

## Important Notes (All Platforms)

⚠️ **You MUST build on Windows** - PyInstaller creates executables for the platform it runs on.

⚠️ **Poppler is required** for PDF processing. Download from:
   https://github.com/oschwartz10612/poppler-windows/releases

⚠️ **First run** will ask for keyword: `Mirkomil12`

## File Structure After Build

```
dist/
├── RashManager.exe      ← Windows build (if built on Windows)
└── RashManager.app      ← macOS build (if built on macOS)

RashManager/  (folder for distribution)
├── RashManager.exe or RashManager.app
├── titul_bubble_koordinatalar_2480x3508.xlsx  ← Required
├── Titul.pdf  ← Optional
└── poppler/  ← If not in PATH
    └── bin/
        └── *.dll, *.exe
```

## Troubleshooting

**Build fails?**
- Make sure Python 3.8+ is installed
- Run: `pip install -r requirements.txt` first

**App/EXE doesn't work?**
- Check all files are in the same folder (or inside `.app/Contents/Resources`)
- Ensure Poppler binaries are accessible (`brew install poppler` on macOS)
- See BUILD_INSTRUCTIONS.md for detailed help

For full instructions, see **BUILD_INSTRUCTIONS.md**

