# PPTX Media Optimizer

Optimize PowerPoint files by compressing images, transcoding videos, and removing unused media.

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/stemanfredi/pptx-media-optimizer/blob/main/pptx_media_optimizer.ipynb)

## Features

- **Image Optimization**: Compress JPEG/PNG with quality control, resize large images
- **Video Transcoding**: Convert to H.264 or H.265 (HEVC), GPU accelerated with NVENC
- **Audio Optimization**: Transcode to AAC
- **Unused Media Removal**: Detect and remove unreferenced files
- **Orphan Template Detection**: Find media in unused masters/layouts
- **Slide Selection**: Optimize specific slides (e.g., "1-10", "1,3,5")
- **Speaker Notes Support**: Detect media in speaker notes
- **Analysis-First Workflow**: See optimization opportunities before applying

## Quick Start

1. Click "Open in Colab" above (or open `pptx_media_optimizer.ipynb` in Google Colab)
2. Run **Cell 1 (Setup)** - installs dependencies, detects GPU
3. Run **Cell 2 (Upload & Analyze)** - upload your PPTX, see media analysis
4. Configure options in **Cell 3**, then run to optimize and download

## Options

### Slide Selection
- `all` - Process all slides (default)
- `1,3,5` - Specific slides
- `1-10` - Range of slides
- `2-5,8,10-12` - Combined ranges

### Video Codec
| Codec | Compatibility | File Size |
|-------|---------------|-----------|
| H.264 | All PowerPoint versions (2013+) | Larger |
| H.265 | PowerPoint 2019+, Windows 10+ | Smaller |

### Quality Settings
- **CRF** (video): 18-22 high quality, 23-28 balanced, 29-36 smaller files
- **JPEG quality**: 30-95 (default 65)
- **PNG quality**: 40-95 (default 70, uses pngquant)
- **Max video height**: 720, 1080, 1440, 2160

## Requirements

- Google Colab (recommended) or local Python 3.8+
- FFmpeg (auto-installed)
- Pillow (auto-installed)
- pngquant (auto-installed)

## How It Works

1. Extracts PPTX as ZIP archive
2. Parses XML relationships to find media references
3. Detects unused media, orphan templates
4. Compresses images with Pillow/pngquant
5. Transcodes video/audio with FFmpeg (GPU or CPU)
6. Updates XML references when file extensions change
7. Repackages as valid PPTX

## Documentation

See [CLAUDE.md](CLAUDE.md) for full specs, architecture, and task tracking.
