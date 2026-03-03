# ProPresenter 7 → PowerPoint Converter

A macOS desktop app that converts **ProPresenter 7 `.pro` files** into **Microsoft PowerPoint `.pptx` files** — white text on a black background, ready to present anywhere.

## Features

- **Single File** — pick one `.pro` file and convert it instantly
- **Watch Folder** — monitor your ProPresenter library folder; new files are auto-converted as they appear. Supports multiple libraries (recursive). "Convert Existing" batch-converts everything already in the folder
- **Merge PPTX** — combine multiple `.pptx` files into one, with drag-to-reorder
- **Deduplication** — skips files already converted (unless the source file has changed)
- Output is always a **flat folder** — no subfolders mirrored from the library structure

## Requirements (to run from source)

- macOS (Apple Silicon / arm64)
- [Homebrew](https://brew.sh)
- Python 3.13 + Tk 9.0

```bash
brew install python@3.13 python-tk@3.13
```

## Setup

```bash
git clone https://github.com/YOUR_USERNAME/pro-to-pptx.git
cd pro-to-pptx

# Create virtual environment
/opt/homebrew/bin/python3.13 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Compile ProPresenter 7 protobuf bindings
git clone https://github.com/greyshirtguy/ProPresenter7-Proto.git /tmp/PP7Proto
python -m grpc_tools.protoc \
    -I /tmp/PP7Proto/Proto \
    --python_out=proto_pb2 \
    /tmp/PP7Proto/Proto/*.proto
```

## Running

```bash
./run.sh
```

Or directly:
```bash
venv/bin/python3.13 pro_to_pptx.py
```

## Distributing (macOS .app + DMG)

A pre-built DMG for Apple Silicon is available on the [Releases](../../releases) page.

To build it yourself after setup:

```bash
# 1. Copy dependencies into the bundle
cp -r proto_pb2 venv pro_to_pptx.py "ProPresenter Converter.app/Contents/Resources/"

# 2. Compile the launcher
clang -arch arm64 \
  -o "ProPresenter Converter.app/Contents/MacOS/ProPresenter Converter" \
  launcher.c

# 3. Sign and package
xattr -cr "ProPresenter Converter.app"
codesign --force --deep --sign - "ProPresenter Converter.app"
hdiutil create \
  -volname "ProPresenter Converter" \
  -srcfolder dmg_staging \
  -format UDZO \
  "ProPresenter Converter.dmg"
```

## Notes

- ProPresenter 7 `.pro` files use Google Protocol Buffers (binary format). This tool uses the community schema from [greyshirtguy/ProPresenter7-Proto](https://github.com/greyshirtguy/ProPresenter7-Proto).
- Only **text/lyrics** are extracted — backgrounds, images, and media are not included by design.
- The app requires Python 3.13 + Tk 9.0 due to a crash in Tk 8.5 on macOS 15 Tahoe.
- Unsigned app — on first launch, right-click → Open to bypass Gatekeeper.

## Dependencies

| Package | Purpose |
|---|---|
| `protobuf` | Parse ProPresenter 7 binary format |
| `python-pptx` | Create PowerPoint files |
| `striprtf` | Extract plain text from RTF data |
| `watchdog` | Monitor folder for new files |
| `grpcio-tools` | Compile `.proto` schema files |
