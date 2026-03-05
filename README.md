# Nedev.PptxToPdf

A high-performance .NET library for converting PPTX (PowerPoint) files to PDF 鈥?**with zero third-party dependencies**. Also ships with a ready-to-use command-line tool.

## Feature Completeness

### Core Pipeline
| Area | Status | Notes |
|------|--------|-------|
| PPTX parsing 鈫?Slide rendering 鈫?PDF output | 鉁?Complete | End-to-end conversion chain |
| Library API | 鉁?Complete | Simple `Convert()` method with file path or stream |
| CLI tool | 鉁?Complete | `Nedev.PptxToPdf.Cli` with `--parallel` flag |

### PPTX Parsing
| Feature | Status | Notes |
|---------|--------|-------|
| Slide master / layout inheritance | 鉁?Complete | Color maps, text styles, default formatting |
| Theme parsing (colors, fonts, effects, format scheme) | 鉁?Complete | Full scheme color resolution |
| Slide transitions & timing | 鉁?Parsed | Data model captured; not rendered (N/A for PDF) |
| Animations | 鉁?Parsed | Data model captured; not rendered (N/A for PDF) |
| Speaker notes | 鉁?Parsed | Not rendered to PDF output |
| Comments / comment authors | 鉁?Parsed | Not rendered to PDF output |
| Document properties | 鉁?Complete | Core, extended & custom properties |
| Hyperlinks | 鉁?Parsed | Internal / external link resolution |

### Shape Rendering
| Feature | Status | Notes |
|---------|--------|-------|
| Basic shapes (rect, ellipse, triangle, diamond, 鈥? | 鉁?Complete | |
| Polygons & stars | 鉁?Complete | |
| Arrows (right, left, up, down) | 鉁?Complete | |
| Decorative shapes (heart, cloud) | 鉁?Complete | |
| Custom geometry / freeform paths | 鉁?Complete | Via `PathRenderer` with B茅zier support |
| Shape rotation & transform | 鉁?Complete | |
| Connectors | 鉁?Complete | Straight, elbow, curved |
| Group shapes (nested) | 鉁?Complete | Recursive child parsing |

### Fill & Outline
| Feature | Status | Notes |
|---------|--------|-------|
| Solid fill | 鉁?Complete | |
| Gradient fill (linear, radial, rectangular, path) | 鉁?Complete | Multi-stop stitching supported |
| Pattern fill | 鈿狅笍 Approximate | Simplified rendering |
| Picture fill | 鈿狅笍 Simplified | Basic implementation; complex crops may differ |
| Shape outline (width, color, dash styles) | 鉁?Complete | |

### Text Rendering
| Feature | Status | Notes |
|---------|--------|-------|
| Paragraph alignment & spacing | 鉁?Complete | |
| Word wrap | 鉁?Complete | Estimated glyph metrics |
| Font, size, color, bold, italic, underline, strikethrough | 鉁?Complete | |
| Superscript / subscript | 鉁?Complete | Baseline offset |
| Bullet styles (char, auto-number) | 鉁?Complete | |
| Text body properties (margins, vertical alignment, auto-fit) | 鉁?Complete | |
| CJK (Chinese/Japanese/Korean) detection | 鉁?Complete | Font fallback for CJK text |

### Font Handling
| Feature | Status | Notes |
|---------|--------|-------|
| 14 PDF standard fonts | 鉁?Complete | |
| System font embedding (TrueType 鈫?Type0/CIDFont) | 鉁?Complete | CMap + ToUnicode for Unicode text |
| Font metrics & width tables | 鉁?Complete | |

### Image Support
| Feature | Status | Notes |
|---------|--------|-------|
| JPEG | 鉁?Native | Passed through directly |
| PNG | 鈿狅笍 Simplified | Decoded to raw pixels 鈫?re-encoded |
| GIF / BMP / TIFF | 鈿狅笍 Simplified | Pixel extraction with placeholder fallback |
| Image effects (shadow, glow, reflection, bevel, soft edges, 3-D rotation) | 鈿狅笍 Approximate | |

### Table Rendering
| Feature | Status | Notes |
|---------|--------|-------|
| Row / column structure | 鉁?Complete | |
| Merged cells (horizontal & vertical) | 鉁?Complete | |
| Cell borders (per-edge style) | 鉁?Complete | |
| Cell text with paragraph formatting | 鉁?Complete | |
| Table styles (header row, banding) | 鈿狅笍 Approximate | |

### Charts
| Feature | Status | Notes |
|---------|--------|-------|
| Bar / Column / Line / Pie / Area charts | 鈿狅笍 Simplified | Rendered from parsed data; visual approximation |
| Scatter / Radar / Doughnut / Bubble / Stock | 鈿狅笍 Parsed | Data model only; rendering placeholder |
| Axes, gridlines, legend | 鈿狅笍 Simplified | Basic rendering |

### SmartArt
| Feature | Status | Notes |
|---------|--------|-------|
| List (vertical / horizontal) | 鈿狅笍 Simplified | |
| Process / Cycle / Matrix / Pyramid | 鈿狅笍 Simplified | |
| Hierarchy / Org chart | 鈿狅笍 Simplified | |
| Relationship / Target | 鈿狅笍 Simplified | |

### Background
| Feature | Status | Notes |
|---------|--------|-------|
| Solid color background | 鉁?Complete | |
| Gradient / image background | 鉂?Not yet | Only solid fill parsed |

### Performance
| Feature | Status | Notes |
|---------|--------|-------|
| Slide-level parallel processing | 鉁?Complete | `--parallel` flag / API parameter |
| Serial lock on PDF write phase | 鈿狅笍 By design | Prevents interleaved output |

### Testing
| Area | Status |
|------|--------|
| Automated test suite | 鉂?Not included | 

## Installation

### Build from Source

```bash
git clone <repository-url>
cd PptxToPdf
dotnet build src/Nedev.PptxToPdf/Nedev.PptxToPdf.csproj
```

## Usage

### As a Library

```csharp
using Nedev.PptxToPdf;

var converter = new PptxToPdfConverter();

// Basic conversion
converter.Convert("input.pptx", "output.pdf");

// Parallel mode (faster for large decks)
converter.Convert("input.pptx", "output.pdf", parallel: true);

// Stream-based conversion
using var input = File.OpenRead("input.pptx");
using var output = File.Create("output.pdf");
converter.Convert(input, output);
```

### Command-Line Tool

```bash
# Basic
Nedev.PptxToPdf.Cli input.pptx output.pdf

# With parallel processing
Nedev.PptxToPdf.Cli input.pptx output.pdf --parallel

# Help
Nedev.PptxToPdf.Cli --help
```

## Project Structure

```
src/
鈹溾攢鈹€ Nedev.PptxToPdf/                    # Core library
鈹?  鈹溾攢鈹€ PptxToPdfConverter.cs      # Public API entry point
鈹?  鈹溾攢鈹€ Pptx/                      # PPTX / OOXML parsing
鈹?  鈹?  鈹溾攢鈹€ PptxDocument.cs        # ZIP archive reader & part loader
鈹?  鈹?  鈹溾攢鈹€ Presentation.cs        # Presentation-level properties
鈹?  鈹?  鈹溾攢鈹€ Slide.cs               # Slide & connector parsing
鈹?  鈹?  鈹溾攢鈹€ SlideMaster.cs         # Master, layout, color map, text styles
鈹?  鈹?  鈹溾攢鈹€ Theme.cs               # Theme colors, fonts, effects
鈹?  鈹?  鈹溾攢鈹€ Shape.cs               # AutoShape geometry, fill, text
鈹?  鈹?  鈹溾攢鈹€ GroupShape.cs          # Group shape tree
鈹?  鈹?  鈹溾攢鈹€ Picture.cs             # Embedded image references
鈹?  鈹?  鈹溾攢鈹€ Table.cs               # Table, row, cell, borders
鈹?  鈹?  鈹溾攢鈹€ Chart.cs               # Chart data & series
鈹?  鈹?  鈹溾攢鈹€ SmartArt.cs            # SmartArt diagrams
鈹?  鈹?  鈹溾攢鈹€ Background.cs          # Slide background
鈹?  鈹?  鈹溾攢鈹€ Hyperlink.cs           # Hyperlink resolution
鈹?  鈹?  鈹溾攢鈹€ Animation.cs           # Animation data model
鈹?  鈹?  鈹溾攢鈹€ Notes.cs               # Speaker notes & comments
鈹?  鈹?  鈹斺攢鈹€ DocumentProperties.cs  # Core / extended metadata
鈹?  鈹溾攢鈹€ Pdf/                       # PDF generation
鈹?  鈹?  鈹溾攢鈹€ PdfDocument.cs         # PDF object tree & serialization
鈹?  鈹?  鈹溾攢鈹€ PdfObjects.cs          # Low-level PDF object types
鈹?  鈹?  鈹溾攢鈹€ PdfRenderer.cs         # Slide 鈫?PDF content stream
鈹?  鈹?  鈹溾攢鈹€ FontManager.cs         # Standard font mapping & metrics
鈹?  鈹?  鈹溾攢鈹€ EmbeddedFontManager.cs # System font embedding (Type0)
鈹?  鈹?  鈹溾攢鈹€ FontEmbedder.cs        # TrueType font file reader
鈹?  鈹?  鈹溾攢鈹€ GradientRenderer.cs    # Gradient shading patterns
鈹?  鈹?  鈹溾攢鈹€ PathRenderer.cs        # Custom geometry 鈫?PDF paths
鈹?  鈹?  鈹斺攢鈹€ ImageEffectsRenderer.cs# Shadow, glow, reflection, etc.
鈹?  鈹溾攢鈹€ Image/                     # Image processing
鈹?  鈹?  鈹溾攢鈹€ ImageConverter.cs      # Format conversion (PNG/GIF/BMP/TIFF 鈫?JPEG)
鈹?  鈹?  鈹斺攢鈹€ ImageDecoder.cs        # Raw pixel extraction
鈹?  鈹斺攢鈹€ Models/                    # Shared data models
鈹?      鈹溾攢鈹€ Color.cs               # Color representation & conversion
鈹?      鈹溾攢鈹€ Enums.cs               # Shared enumerations
鈹?      鈹溾攢鈹€ Fill.cs, Outline.cs    # Fill & outline models
鈹?      鈹溾攢鈹€ Paragraph.cs           # Paragraph & run models
鈹?      鈹溾攢鈹€ Rect.cs, Geometry.cs   # Geometry primitives
鈹?      鈹溾攢鈹€ ShapeTypeMapping.cs    # OOXML preset 鈫?internal shape type
鈹?      鈹溾攢鈹€ TextProperties.cs      # Text body / paragraph properties
鈹?      鈹溾攢鈹€ Transform2D.cs         # 2-D transform
鈹?      鈹斺攢鈹€ GradientStop.cs        # Gradient stop model
鈹斺攢鈹€ Nedev.PptxToPdf.Cli/               # Command-line interface
    鈹斺攢鈹€ Program.cs
```

## Requirements

- **.NET 10** (or later)
- No third-party NuGet packages required

## License

This project is licensed under the MIT License.
