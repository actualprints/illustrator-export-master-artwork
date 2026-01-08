# Export Master Artwork - Illustrator ExtendScript

Export each artwork layer as a separate high-resolution PNG from Adobe Illustrator, with automatic cut contour hiding, ZIP creation, and PDF backup.

## Features

- Exports each layer as a separate PNG at 1400 DPI
- Automatically hides cut contour lines (CutContour spot colors)
- Creates artboard at correct bleed size for each card thickness
- Auto-generates ZIP file of all exports
- Saves PDF backup of the master artwork
- Organized output folder structure

## Installation

1. Download `export_master_artwork.jsx`
2. Save it anywhere on your computer (e.g., `Documents/Scripts/`)

## Usage

1. Open your master artwork file in Illustrator
2. Ensure each artwork is on a separate layer named by card thickness (35, 55, 75, etc.)
3. Run the script: **File > Scripts > Other Script...**
4. Enter the proof number when prompted (e.g., `6843470155` or `6843470155-1`)
5. Files are exported to: `C:\Users\MasterBrader\Dropbox\STICKERs\ActualPrints\{proofnumber}\`

## Output Structure

```
ActualPrints/
└── 6843470155/
    ├── 6843470155_exports/
    │   ├── 79x90.png
    │   ├── 79x91.png
    │   ├── 79x94.png
    │   ├── 79x95.png
    │   ├── 79x98.png
    │   └── 79x101.png
    ├── 6843470155_exports.zip
    └── 6843470155-MasterArtwork.pdf
```

## Size Mapping

| Layer Name | Output Size | Bleed Dimensions |
|------------|-------------|------------------|
| 35 | 79x90.png | 30mm x 33.75mm |
| 55 | 79x91.png | 30mm x 34.25mm |
| 75 | 79x94.png | 30mm x 35mm |
| 100 | 79x95.png | 30mm x 35.5mm |
| 130 | 79x98.png | 30mm x 36.5mm |
| 180 | 79x101.png | 30mm x 37.5mm |

## Cut Contour Detection

The script automatically hides elements with these patterns in their name or spot color:
- CutContour, Cut Contour
- Dieline, Die Line, Diecut
- Kiss Cut, Thru-Cut, Through Cut
- Perf, Score, Contour

## Configuration

Edit the `CONFIG` object at the top of the script to customize:
- `exportPPI`: Export resolution (default: 1400)
- `sizeMapping`: Card thickness to size mapping
- `cutContourPatterns`: Patterns to identify cut lines

## License

MIT License
