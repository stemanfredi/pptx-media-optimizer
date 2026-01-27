# PPTX Media Optimizer

Optimize PowerPoint files by compressing images, transcoding videos, and removing unused media.

## Principles

- **DRY** - Don't Repeat Yourself
- **KISS** - Keep It Simple
- **YAGNI** - You Aren't Gonna Need It
- **Specs First** - Update spec before code
- **Less is more** - Best code is no code

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Google Colab Notebook                    │
├─────────────────────────────────────────────────────────────┤
│  Cell 1: Setup     │ Install deps, detect GPU, setup FFmpeg │
├─────────────────────────────────────────────────────────────┤
│  Cell 2: Core      │ PPTX extract/repackage utilities       │
├─────────────────────────────────────────────────────────────┤
│  Cell 3: Analysis  │ Scan media, detect unused, report      │
├─────────────────────────────────────────────────────────────┤
│  Cell 4: Image     │ Pillow-based resize/compress           │
├─────────────────────────────────────────────────────────────┤
│  Cell 5: Video     │ FFmpeg transcode (NVENC/libx264)       │
├─────────────────────────────────────────────────────────────┤
│  Cell 6: Cleanup   │ Remove unreferenced media files        │
├─────────────────────────────────────────────────────────────┤
│  Cell 7: Main      │ UI forms, orchestration, output        │
└─────────────────────────────────────────────────────────────┘
```

## Structure

```
pptx-media-optimizer/
├── pptx_media_optimizer.ipynb   # Main Colab notebook
├── CLAUDE.md                    # This file
└── README.md                    # Usage instructions
```

## Commits

Format: `<type>(<scope>): <description>`

| Type | Use |
|------|-----|
| feat | New feature |
| fix | Bug fix |
| docs | Documentation |
| refactor | Code restructuring |
| test | Adding tests |
| chore | Maintenance |

Rules: subject ≤50 chars, lowercase, no period, imperative mood

---

## Specs

### Environment
- **Platform**: Google Colab (Linux)
- **Dependencies**: `python-pptx`, `Pillow`, FFmpeg
- **GPU**: NVIDIA T4/A100 (NVENC), fallback to CPU (libx264)

### PPTX Handling
- Treat `.pptx` as ZIP archive
- Extract to temp directory, process, repackage
- **Never overwrite source file** - always output to new file
- Preserve XML structure and relationships

### Image Optimization
| Format | Action | Notes |
|--------|--------|-------|
| JPEG | Compress | Quality slider 10-100 |
| PNG | Compress | Preserve RGBA transparency |
| EMF/WMF | Skip | Vector formats, not processable |
| SVG | Skip | Vector format |
| GIF | Skip | Animation preservation complex |

- Downscale if width > max_width (default: 1920px)
- Preserve original file extensions (critical for XML refs)
- Skip already-small images (< threshold)

### Video Optimization
| Input | Output | Codec |
|-------|--------|-------|
| Any video | MP4 | H.264 (AVC) + AAC |

- **H.264/AAC in MP4**: Maximum PowerPoint compatibility (2013+)
- GPU: `h264_nvenc` with `-preset fast`
- CPU fallback: `libx264` with `-preset medium -crf 23`
- Preserve aspect ratio
- Audio: AAC 128kbps (or copy if already AAC)

### Audio Optimization
- Same FFmpeg pipeline as video
- Output: AAC in M4A container
- Bitrate: 128kbps default

### Unused Media Detection
Parse relationship files to find referenced media:
- `ppt/slides/_rels/*.xml.rels`
- `ppt/slideLayouts/_rels/*.xml.rels`
- `ppt/slideMasters/_rels/*.xml.rels`
- `ppt/notesSlides/_rels/*.xml.rels` (if exists)

Compare against `ppt/media/` contents. Unreferenced files are candidates for deletion.

### Workflow
1. **Upload**: File upload widget (simple, no Drive mount)
2. **Extract**: Unzip PPTX to temp workspace
3. **Analyze**: Scan all media, detect unused, calculate potential savings
4. **Present**: Show analysis with per-item toggle options
5. **Execute**: Process only selected optimizations
6. **Repackage**: Create new valid PPTX
7. **Report**: Size comparison (original vs optimized)
8. **Download**: Provide download link

### UI Parameters (Colab Forms)
```python
input_file       # File upload widget
image_quality    # Slider 10-100, default 85
max_image_width  # Integer, default 1920
video_bitrate    # String, default "2M"
enable_gpu       # Boolean, default True (auto-detect)
```

### Error Handling
- Validate ZIP structure before processing
- Skip corrupt/unreadable media files
- Log warnings, don't fail entire process
- Preserve original on any critical error

---

## Tasks

### Current

[No active tasks]

### Backlog

- [ ] Banner crop feature for letterboxed videos (v2)
- [ ] Batch processing multiple PPTX files
- [ ] Custom FFmpeg parameters

### Done

- [x] Initial spec design
- [x] Architecture decision (analysis-first workflow)
- [x] Research PPTX video codec compatibility (H.264/AAC in MP4)
- [x] Research Colab GPU/FFmpeg setup (NVENC detection, libx264 fallback)
- [x] Implement notebook with analysis-first workflow
