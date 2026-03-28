# Nedev.FileConverters.PptxToPdf

A high-performance .NET library for converting PPTX (PowerPoint) files to PDF **with zero third-party dependencies**. Also ships with a ready-to-use command-line tool.

## Features

- **Zero Dependencies**: No external NuGet packages required for core functionality
- **Core Integration**: Supports `Nedev.FileConverters.Core` via explicit registration plus `IFileConverter` adapter
- **Multi-Target**: Supports both .NET 8.0 and .NET Standard 2.1
- **Parallel Processing**: Optional slide-level parallelization for faster conversion
- **Stream-Based**: Support for both file path and stream-based conversion
- **CJK Support**: Proper handling of Chinese, Japanese, and Korean characters

## Feature Completeness

### Core Pipeline
| Area | Status | Notes |
|------|--------|-------|
| PPTX parsing | ✅ Complete | Full OOXML format support |
| Slide rendering | ✅ Complete | End-to-end conversion chain |
| PDF output | ✅ Complete | Standard PDF generation |
| Library API | ✅ Complete | Simple `Convert()` method with file path or stream |
| CLI tool | ✅ Complete | `Nedev.FileConverters.PptxToPdf.Cli` with `--parallel` and `--core` registration modes |

### PPTX Parsing
| Feature | Status | Notes |
|---------|--------|-------|
| Slide master / layout inheritance | ✅ Complete | Color maps, text styles, default formatting |
| Theme parsing (colors, fonts, effects, format scheme) | ✅ Complete | Full scheme color resolution |
| Slide transitions & timing | ✅ Parsed | Data model captured; not rendered (N/A for PDF) |
| Animations | ✅ Parsed | Data model captured; not rendered (N/A for PDF) |
| Speaker notes | ✅ Parsed | Not rendered to PDF output |
| Comments / comment authors | ✅ Parsed | Not rendered to PDF output |
| Document properties | ✅ Complete | Core, extended & custom properties |
| Hyperlinks | ✅ Parsed | Internal / external link resolution |

### Shape Rendering
| Feature | Status | Notes |
|---------|--------|-------|
| Basic shapes (rect, ellipse, triangle, diamond, etc.) | ✅ Complete | |
| Polygons & stars | ✅ Complete | |
| Arrows (right, left, up, down) | ✅ Complete | |
| Decorative shapes (heart, cloud) | ✅ Complete | |
| Custom geometry / freeform paths | ✅ Complete | Via `PathRenderer` with Bézier support |
| Shape rotation & transform | ✅ Complete | |
| Connectors | ✅ Complete | Straight, elbow, curved |
| Group shapes (nested) | ✅ Complete | Recursive child parsing |

### Fill & Outline
| Feature | Status | Notes |
|---------|--------|-------|
| Solid fill | ✅ Complete | |
| Gradient fill (linear, radial, rectangular, path) | ✅ Complete | Multi-stop stitching supported |
| Pattern fill | ⚠️ Approximate | Simplified rendering |
| Picture fill | ⚠️ Simplified | Basic implementation; complex crops may differ |
| Shape outline (width, color, dash styles) | ✅ Complete | |

### Text Rendering
| Feature | Status | Notes |
|---------|--------|-------|
| Paragraph alignment & spacing | ✅ Complete | |
| Word wrap | ✅ Complete | Estimated glyph metrics |
| Font, size, color, bold, italic, underline, strikethrough | ✅ Complete | |
| Superscript / subscript | ✅ Complete | Baseline offset |
| Bullet styles (char, auto-number) | ✅ Complete | |
| Text body properties (margins, vertical alignment, auto-fit) | ✅ Complete | |
| CJK (Chinese/Japanese/Korean) detection | ✅ Complete | Font fallback for CJK text |

### Font Handling
| Feature | Status | Notes |
|---------|--------|-------|
| 14 PDF standard fonts | ✅ Complete | |
| System font embedding (TrueType & Type0/CIDFont) | ✅ Complete | CMap + ToUnicode for Unicode text |
| Font metrics & width tables | ✅ Complete | |

### Image Support
| Feature | Status | Notes |
|---------|--------|-------|
| JPEG | ✅ Native | Passed through directly |
| PNG | ⚠️ Simplified | Decoded to raw pixels & re-encoded |
| BMP | ✅ Complete | Lossless pixel extraction with alpha-mask support |
| GIF / TIFF | ⚠️ Unsupported | Unsupported images are skipped instead of embedding placeholder pixels |
| Image effects (shadow, glow, reflection, bevel, soft edges, 3-D rotation) | ⚠️ Approximate | |

### Table Rendering
| Feature | Status | Notes |
|---------|--------|-------|
| Row / column structure | ✅ Complete | |
| Merged cells (horizontal & vertical) | ✅ Complete | |
| Cell borders (per-edge style) | ✅ Complete | |
| Cell text with paragraph formatting | ✅ Complete | |
| Table styles (header row, banding) | ⚠️ Approximate | |

### Charts
| Feature | Status | Notes |
|---------|--------|-------|
| Bar / Column / Line / Pie / Area charts | ⚠️ Simplified | Rendered from parsed data; visual approximation |
| Scatter / Radar / Doughnut / Bubble / Stock | ⚠️ Parsed | Data model only; rendering placeholder |
| Axes, gridlines, legend | ⚠️ Simplified | Basic rendering |

### SmartArt
| Feature | Status | Notes |
|---------|--------|-------|
| List (vertical / horizontal) | ⚠️ Simplified | |
| Process / Cycle / Matrix / Pyramid | ⚠️ Simplified | |
| Hierarchy / Org chart | ⚠️ Simplified | |
| Relationship / Target | ⚠️ Simplified | |

### Background
| Feature | Status | Notes |
|---------|--------|-------|
| Solid color background | ✅ Complete | |
| Gradient / image background | ⚠️ Partial | Solid and gradient backgrounds render; image backgrounds remain unimplemented |

### Performance
| Feature | Status | Notes |
|---------|--------|-------|
| Slide-level parallel processing | ✅ Complete | `--parallel` flag / API parameter |
| Serial lock on PDF write phase | ⚠️ By design | Prevents interleaved output |

## Installation

### NuGet Package

```powershell
Install-Package Nedev.FileConverters.PptxToPdf
```

### Build from Source

```bash
git clone <repository-url>
cd FileConverters.PptxToPdf
dotnet build src/Nedev.FileConverters.PptxToPdf/Nedev.FileConverters.PptxToPdf.csproj
```

## Usage

### Using Nedev.FileConverters.Core

```csharp
using Nedev.FileConverters;
using Nedev.FileConverters.PptxToPdf;

PptxToPdfCoreRegistration.EnsureRegistered();

using var inputStream = File.OpenRead("input.pptx");
using var convertedStream = Converter.Convert(inputStream, "pptx", "pdf");
using var outputStream = File.Create("output.pdf");
convertedStream.CopyTo(outputStream);
```

### Using The Adapter Directly

```csharp
using Nedev.FileConverters.PptxToPdf;

using var inputStream = File.OpenRead("input.pptx");
var fileConverter = new PptxToPdfFileConverter();
using var convertedStream = fileConverter.Convert(inputStream);
using var outputStream = File.Create("output.pdf");
convertedStream.CopyTo(outputStream);
```

### As a Library

```csharp
using Nedev.FileConverters.PptxToPdf;

var converter = new PptxToPdfConverter();

// Basic conversion
converter.Convert("input.pptx", "output.pdf");

// Parallel mode (faster for large decks)
converter.Convert("input.pptx", "output.pdf", useParallelProcessing: true);

// Stream-based conversion
using var input = File.OpenRead("input.pptx");
using var output = File.Create("output.pdf");
converter.Convert(input, output);
```

### Command-Line Tool

```bash
# Basic
Nedev.FileConverters.PptxToPdf.Cli input.pptx output.pdf

# With parallel processing
Nedev.FileConverters.PptxToPdf.Cli input.pptx output.pdf --parallel

# Using the Core package entry point
Nedev.FileConverters.PptxToPdf.Cli input.pptx output.pdf --core

# Help
Nedev.FileConverters.PptxToPdf.Cli --help
```

## Project Structure

```
src/
├── Nedev.FileConverters.PptxToPdf/                    # Core library
│   ├── PptxToPdfConverter.cs                           # Public API entry point
│   ├── PptxToPdfFileConverter.cs                       # IFileConverter implementation
│   ├── Pptx/                                         # PPTX / OOXML parsing
│   │   ├── PptxDocument.cs                             # ZIP archive reader & part loader
│   │   ├── Presentation.cs                             # Presentation-level properties
│   │   ├── Slide.cs                                    # Slide & connector parsing
│   │   ├── SlideMaster.cs                               # Master, layout, color map, text styles
│   │   ├── Theme.cs                                    # Theme colors, fonts, effects
│   │   ├── Shape.cs                                    # AutoShape geometry, fill, text
│   │   ├── GroupShape.cs                               # Group shape tree
│   │   ├── Picture.cs                                  # Embedded image references
│   │   ├── Table.cs                                    # Table, row, cell, borders
│   │   ├── Chart.cs                                    # Chart data & series
│   │   ├── SmartArt.cs                                 # SmartArt diagrams
│   │   ├── Background.cs                                # Slide background
│   │   ├── Hyperlink.cs                                 # Hyperlink resolution
│   │   ├── Animation.cs                                 # Animation data model
│   │   ├── Notes.cs                                    # Speaker notes & comments
│   │   └── DocumentProperties.cs                        # Core / extended metadata
│   ├── Pdf/                                           # PDF generation
│   │   ├── PdfDocument.cs                               # PDF object tree & serialization
│   │   ├── PdfObjects.cs                                # Low-level PDF object types
│   │   ├── PdfRenderer.cs                               # Slide → PDF content stream
│   │   ├── FontManager.cs                               # Standard font mapping & metrics
│   │   ├── EmbeddedFontManager.cs                         # System font embedding (Type0)
│   │   ├── FontEmbedder.cs                              # TrueType font file reader
│   │   ├── GradientRenderer.cs                           # Gradient shading patterns
│   │   ├── PathRenderer.cs                              # Custom geometry → PDF paths
│   │   └── ImageEffectsRenderer.cs                      # Shadow, glow, reflection, etc.
│   ├── Image/                                         # Image processing
│   │   ├── ImageConverter.cs                           # Format conversion (PNG/GIF/BMP/TIFF → JPEG)
│   │   └── ImageDecoder.cs                             # Raw pixel extraction
│   └── Models/                                        # Shared data models
│       ├── Color.cs                                    # Color representation & conversion
│       ├── Enums.cs                                    # Shared enumerations
│       ├── Fill.cs, Outline.cs                          # Fill & outline models
│       ├── Paragraph.cs                                  # Paragraph & run models
│       ├── Rect.cs, Geometry.cs                          # Geometry primitives
│       ├── ShapeTypeMapping.cs                           # OOXML preset → internal shape type
│       ├── TextProperties.cs                             # Text body / paragraph properties
│       ├── Transform2D.cs                                # 2-D transform
│       ├── GradientStop.cs                              # Gradient stop model
│       └── ShapeEffects.cs                              # Shadow, glow, reflection effects
└── Nedev.FileConverters.PptxToPdf.Cli/               # Command-line interface
    └── Program.cs
```

## Integration with Nedev.FileConverters.Core

This package integrates seamlessly with `Nedev.FileConverters.Core`:

- **Automatic Discovery**: The `PptxToPdfFileConverter` class is decorated with `[FileConverter("pptx", "pdf")]` attribute
- **Unified API**: Use `Nedev.FileConverters.Converter.Convert()` for consistent interface across all converters
- **DI Support**: Can be registered in dependency injection containers
- **Extensibility**: Easy to add more converters following the same pattern

## Requirements

- **.NET 8.0** or **.NET Standard 2.1** compatible runtime
- **Nedev.FileConverters.Core** 0.1.0 or later
- No third-party NuGet packages required for core functionality

## License

This project is licensed under the MIT License.
