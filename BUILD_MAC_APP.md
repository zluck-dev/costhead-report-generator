## Build macOS .app (GIN Mapper)

Prerequisites:

- macOS with Xcode Command Line Tools: `xcode-select --install`
- Python 3.12 (already vendored in `exfile/` venv)
- PyInstaller installed in venv

### 1) Activate virtual environment

```bash
source exfile/bin/activate
```

### 2) Install/Update PyInstaller (if needed)

```bash
python -m pip install --upgrade pip pyinstaller
```

### 3) Create .icns app icon (one-time)

If `assets/app_icon.icns` does not exist, generate from `assets/app_icon.png`:

```bash
mkdir -p assets/AppIcon.iconset
sips -z 16 16   assets/app_icon.png --out assets/AppIcon.iconset/icon_16x16.png
sips -z 32 32   assets/app_icon.png --out assets/AppIcon.iconset/icon_16x16@2x.png
sips -z 32 32   assets/app_icon.png --out assets/AppIcon.iconset/icon_32x32.png
sips -z 64 64   assets/app_icon.png --out assets/AppIcon.iconset/icon_32x32@2x.png
sips -z 128 128 assets/app_icon.png --out assets/AppIcon.iconset/icon_128x128.png
sips -z 256 256 assets/app_icon.png --out assets/AppIcon.iconset/icon_128x128@2x.png
sips -z 256 256 assets/app_icon.png --out assets/AppIcon.iconset/icon_256x256.png
sips -z 512 512 assets/app_icon.png --out assets/AppIcon.iconset/icon_256x256@2x.png
sips -z 512 512 assets/app_icon.png --out assets/AppIcon.iconset/icon_512x512.png
sips -z 1024 1024 assets/app_icon.png --out assets/AppIcon.iconset/icon_512x512@2x.png
iconutil -c icns assets/AppIcon.iconset -o assets/app_icon.icns
```

### 4) Build the app

```bash
pyinstaller \
  --noconfirm \
  --clean \
  --name "GIN Mapper" \
  --windowed \
  --icon "assets/app_icon.icns" \
  --add-data "assets:assets" \
  main.py
```

Outputs:

- App bundle: `dist/GIN Mapper.app`
- Distribution folder: `dist/`

### 5) Run locally

```bash
open "dist/GIN Mapper.app"
```

### 6) Distribute

- Zip the app for sharing:

```bash
cd dist
zip -r "GIN Mapper.zip" "GIN Mapper.app"
```

### 7) (Optional) Sign and notarize

To avoid Gatekeeper warnings, sign and notarize with an Apple Developer ID.
High-level steps:

1. Create a Developer ID Application certificate in Keychain.
2. Codesign the app:

```bash
codesign --deep --force --options runtime --sign "Developer ID Application: YOUR NAME (TEAMID)" "dist/GIN Mapper.app"
```

3. Notarize using `xcrun notarytool` or `altool`.
4. Staple the ticket:

```bash
xcrun stapler staple "dist/GIN Mapper.app"
```

Notes:

- If you change assets or code, re-run step 4.
- If Tk windows don’t show or fonts look off, ensure Tcl/Tk is available on the target macOS.

### Set App Name and App Icon
- App name comes from the `--name` flag (affects the .app bundle name and process name).
- App icon is set with the `--icon` flag and should be an `.icns` file on macOS.

Example:
```bash
pyinstaller \
  --name "GIN Mapper" \
  --windowed \
  --icon assets/app_icon.icns \
  --add-data "assets:assets" \
  main.py
```

Notes:
- macOS requires `.icns` for best results. See the icon generation steps above.
- On Windows use `.ico`, and on Linux `.png` is usually fine.

### PyInstaller options (quick reference)
- `--onefile`: Pack everything into a single executable. On macOS GUI apps, app bundles are often preferred.
- `--windowed` / `--noconsole`: Hide console window for GUI apps (Tkinter, Qt, etc.).
- `--console`: Show console (default for CLI apps).
- `--name NAME`: Sets app name and product name.
- `--icon PATH`: Sets the app icon (`.icns` on macOS, `.ico` on Windows).
- `--add-data SRC:DEST`: Bundle extra files/folders. On macOS/Linux use `:`; on Windows use `;`.
- `--hidden-import MOD`: Force include a module PyInstaller doesn’t auto-detect.
- `--collect-submodules PKG`: Include all submodules of a package.
- `--collect-data PKG`: Include package data files.
- `--exclude-module MOD`: Exclude a module from build.
- `--clean`: Clean PyInstaller cache and remove temporary files.
- `--noconfirm`: Overwrite output folder without asking.
- `--osx-bundle-identifier com.example.app`: Set bundle identifier in Info.plist.
- `--codesign-identity "Developer ID Application: NAME (TEAMID)"`: Codesign the app during build (macOS).
- `--version-file file.rc` (Windows): Embed version info/resources.

### Customizing the .spec file (optional)
PyInstaller writes a spec file on first run (e.g., `GIN Mapper.spec`). You can:
- Add/modify `datas` to include more assets.
- Set `bundle_identifier` (Info.plist) via `osx.BUNDLE` options.
- Adjust Info.plist entries (e.g., `CFBundleName`, `CFBundleIconFile`).

Typical flow:
1) Generate spec: run PyInstaller once.
2) Edit `GIN Mapper.spec` (tune datas, identifiers, entitlements, etc.).
3) Build from spec:
```bash
pyinstaller "GIN Mapper.spec"
```

Example snippet inside a spec (illustrative):
```python
app = BUNDLE(
    appname='GIN Mapper.app',
    name='GIN Mapper',
    icon='assets/app_icon.icns',
    bundle_identifier='com.yourcompany.ginmapper',
    info_plist={
        'CFBundleName': 'GIN Mapper',
        'CFBundleDisplayName': 'GIN Mapper',
    },
    # ...
)
```

Tips:
- If you change icons or Info.plist via the spec, re-run `pyinstaller "GIN Mapper.spec"`.
- Always test the produced `.app` on a clean macOS machine or a separate user to ensure dependencies are bundled.

## Build for Intel Macs (x86_64)

You can produce an Intel-only app on an Apple Silicon Mac using Rosetta, or build directly on an Intel Mac.

### Option A: Build x86_64 on Apple Silicon (via Rosetta)
1) Install Rosetta (once):
```bash
softwareupdate --install-rosetta --agree-to-license
```
2) Create an Intel venv and install dependencies as x86_64:
```bash
/usr/bin/arch -x86_64 /Library/Frameworks/Python.framework/Versions/3.12/bin/python3 -m venv exfile_x86
source exfile_x86/bin/activate
/usr/bin/arch -x86_64 python -m pip install --upgrade pip
/usr/bin/arch -x86_64 python -m pip install -r requirements.txt pyinstaller --no-cache-dir --force-reinstall
```
3) Clean previous build artifacts (important when switching arch):
```bash
rm -rf build dist *.spec
```
4) Build the Intel app:
```bash
/usr/bin/arch -x86_64 pyinstaller \
  --noconfirm --clean \
  --name "GIN Mapper" \
  --windowed \
  --icon assets/app_icon.icns \
  --add-data "assets:assets" \
  --target-arch x86_64 \
  main.py
```
5) Verify the binary architecture:
```bash
lipo -archs "dist/GIN Mapper.app/Contents/MacOS/GIN Mapper"
# should print: x86_64
```

### Option B: Build on an Intel Mac
1) Create venv and install deps normally (no Rosetta needed):
```bash
python3 -m venv exfile
source exfile/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller
```
2) Build the app:
```bash
pyinstaller \
  --noconfirm --clean \
  --name "GIN Mapper" \
  --windowed \
  --icon assets/app_icon.icns \
  --add-data "assets:assets" \
  main.py
```
3) Verify:
```bash
lipo -archs "dist/GIN Mapper.app/Contents/MacOS/GIN Mapper"
# should print: x86_64
```

Notes:
- When distributing, consider renaming builds to clarify arch, e.g., `GIN Mapper (Intel).app` and `GIN Mapper (Apple Silicon).app`.
- If building universal2 fails due to single-arch native wheels (numpy/pandas/lxml), prefer shipping two separate builds (one per arch).
