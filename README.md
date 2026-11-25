# PANG & PPT Studio — Professional PDF, PowerPoint & Document Studio

**PANG & PPT Studio** is a Windows desktop application engineered for graphic designers, agencies, and presentation professionals working with Adobe Illustrator exports, vector PDFs, and PowerPoint presentations.

---

## Key Capabilities

1. **Adobe Illustrator PDF Optimizer & Compressor**
   - Strips redundant Illustrator XML metadata (`/Metadata`, `/PieceInfo`, `/AcroForm`).
   - Intelligent stream deflation and object garbage collection.
   - Smart image downsampling (72, 96, 120, 150, 200, 300 DPI) with PNG alpha preservation.
   - Built-in presets: *Balanced Studio (150 DPI)*, *Maximum Quality (300 DPI)*, *Mobile-Friendly (120 DPI)*, *Smallest File Size (96 DPI)*, and *Print-Ready (300 DPI)*.

2. **Rendering & Compatibility Repair**
   - Automatically detects problematic opacity masks, blend modes (`Multiply`, `Screen`, `Overlay`, etc.), and complex transparency groups.
   - Selective high-resolution raster flattening (300 DPI) for mobile devices (iOS / Android / WhatsApp PDF preview compatibility) while maintaining vector sharpness.

3. **PDF to Editable PowerPoint Converter**
   - Converts PDF artboards into native `.pptx` slides preserving exact aspect ratio and dimensions.
   - Reconstructs text as editable text boxes with typography positioning.
   - Extracts vector shapes and raster illustrations as independent objects.
   - High-fidelity mode + Fully editable mode.

4. **PowerPoint to PDF Converter**
   - Converts `.pptx` slides into optimized, publication-ready PDFs.

5. **Apple Pro Dark Studio Design System**
   - Minimal charcoal `#090909` palette with `#8B7CFF` accent.
   - Live side-by-side Before/After comparison frame.
   - Deep analysis inspector displaying embedded fonts, image streams, vector path count, and artboard dimensions.

---

## Quick Start

### 1. Launch the Desktop Application
Double-click `Run_PANG_Studio.bat` or run:
```bash
python desktop_app.py
```

### 2. Command-Line Processing Engine
The core processing engine can also be invoked directly from the CLI or shell scripts:

```bash
# Deep inspection of an Illustrator PDF
python engine.py inspect --input document.pdf

# Optimize PDF with custom settings
python engine.py optimize --input document.pdf --output document_optimized.pdf

# Repair transparency and blend mode issues for mobile viewing
python engine.py repair --input document.pdf --output document_repaired.pdf

# Convert PDF to editable PowerPoint presentation
python engine.py pdf2ppt --input document.pdf --output presentation.pptx

# Convert PowerPoint to PDF
python engine.py ppt2pdf --input presentation.pptx --output presentation.pdf
```
