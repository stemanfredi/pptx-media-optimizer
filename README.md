# PPTX Media Optimizer

Optimize PowerPoint files by compressing images, transcoding videos, and removing unused media.

## Features

- **Image Optimization**: Compress JPEG/PNG with quality control, resize large images
- **Video Transcoding**: Convert to H.264/AAC (PowerPoint-compatible), GPU accelerated
- **Audio Optimization**: Transcode to AAC
- **Unused Media Removal**: Detect and remove unreferenced files
- **Analysis-First Workflow**: See optimization opportunities before applying

## Quick Start

1. Open `pptx_media_optimizer.ipynb` in Google Colab
2. Run cells 1-2 to install dependencies
3. Upload your PPTX file
4. Run analysis to see optimization opportunities
5. Configure which optimizations to apply
6. Execute and download the optimized file

[![Open In Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/stemanfredi/pptx-media-optimizer/blob/main/pptx_media_optimizer.ipynb)

## Requirements

- Google Colab (recommended) or local Python 3.8+
- FFmpeg
- python-pptx
- Pillow

## Codec Compatibility

Output videos use **H.264/AAC in MP4** for maximum PowerPoint compatibility:
- PowerPoint 2013+ (Windows, Mac, Web)
- PowerPoint mobile apps
- No codec installation required

## Documentation

See [CLAUDE.md](CLAUDE.md) for full documentation including specs, architecture, and task tracking.
