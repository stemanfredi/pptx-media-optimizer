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
│  Cell 1: Setup         │ Install deps, detect GPU, define  │
│                        │ all functions and classes          │
├─────────────────────────────────────────────────────────────┤
│  Cell 2: Upload &      │ File upload, extract PPTX, scan   │
│          Analyze       │ media, detect unused/orphan       │
├─────────────────────────────────────────────────────────────┤
│  Cell 3: Optimize &    │ Apply optimizations, repackage,   │
│          Download      │ download result                   │
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
- **Dependencies**: `Pillow`, `pngquant`, FFmpeg
- **GPU**: NVIDIA T4/A100 (NVENC H.264/H.265), fallback to CPU (libx264/libx265)

### PPTX Handling
- Treat `.pptx` as ZIP archive
- Extract to temp directory, process, repackage
- **Never overwrite source file** - always output to new file
- Preserve XML structure and relationships
- Update `[Content_Types].xml` with correct MIME types when changing extensions

### Image Optimization
| Format | Action | Notes |
|--------|--------|-------|
| JPEG | Compress | Quality slider 30-95, default 65 |
| PNG | Compress | pngquant quality 40-95, default 70 |
| EMF/WMF | Skip | Vector formats, not processable |
| SVG | Skip | Vector format |
| GIF | Skip | Animation preservation complex |

- Downscale if width > max_width (default: 1600px)
- pngquant for PNG compression (better quality than Pillow quantize)
- Skip already-small images (< threshold)

### Video Optimization
| Input | Output | Codec |
|-------|--------|-------|
| Any video | MP4 | H.264 or H.265 + AAC |

- **H.264/AAC in MP4**: Maximum PowerPoint compatibility (2013+)
- **H.265/AAC in MP4**: Smaller files, PowerPoint 2019+/Windows 10+ only
- GPU: `h264_nvenc`/`hevc_nvenc` with `-preset fast`
- CPU fallback: `libx264`/`libx265` with `-preset fast`
- CRF-based quality control (default: 26)
- Max video height scaling (720/1080/1440/2160, default: 1080)
- Preserve aspect ratio
- Audio: AAC 96kbps
- Update XML references when extension changes (.wmv → .mp4)

### Audio Optimization
- Same FFmpeg pipeline as video
- Output: AAC in M4A container
- Bitrate: 96kbps
- Update XML references when extension changes (.wav → .m4a)

### Unused Media Detection
Parse relationship files to find referenced media:
- `ppt/slides/_rels/*.xml.rels`
- `ppt/slideLayouts/_rels/*.xml.rels`
- `ppt/slideMasters/_rels/*.xml.rels`
- `ppt/notesSlides/_rels/*.xml.rels`

Compare against `ppt/media/` contents. Unreferenced files are candidates for deletion.

### Orphan Master/Layout Detection
- Trace Slide → Layout → Master chain to find active templates
- Detect layouts/masters not used by any slide
- Identify media only referenced by orphan templates

### Workflow
1. **Setup**: Install dependencies, detect GPU
2. **Upload & Analyze**: Upload file, extract, scan media, show report
3. **Optimize & Download**: Apply selected optimizations, repackage, download

### UI Parameters (Colab Forms)
```python
# Slide Selection
slides = "all"           # "all", "1,3,5", "1-10", "2-5,8,10-12"

# Image Settings
jpeg_quality = 65        # Slider 30-95
png_quality = 70         # Slider 40-95
max_image_width = 1600   # Integer

# Video Settings
video_codec = "h264"     # "h264" or "h265"
video_crf = 26           # Slider 18-36 (lower = better quality)
max_video_height = 1080  # 720, 1080, 1440, 2160

# What to optimize
remove_unused = True
remove_orphan_media = True
optimize_images = True
optimize_videos = True
optimize_audio = True
include_templates = True
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
- [x] Add pngquant for high-quality PNG compression
- [x] Streamline to 3-cell workflow (Setup, Analyze, Optimize)
- [x] Add slide selection feature (range parsing)
- [x] Add orphan master/layout detection
- [x] Add H.265 codec option with compatibility warning
- [x] Add CRF-based quality control
- [x] Add max video height scaling
- [x] Fix Content-Type updates when changing extensions
- [x] Add notesSlides media reference parsing
