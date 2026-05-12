# Video Editor Agent

Complete video editing automation that removes silences, adds jump cuts, applies transitions, and intelligently names output files based on transcription.

## Quick Start

```bash
# Basic editing - removes silences, auto-names based on content
python execution/video_editor_pipeline.py input.mp4

# Full processing with audio enhancement and transitions
python execution/video_editor_pipeline.py input.mp4 --enhance-audio --add-transitions

# Custom output location
python execution/video_editor_pipeline.py input.mp4 --output edited_video.mp4
```

---

## What It Does

1. **Extracts audio** from the input video
2. **Transcribes** using OpenAI Whisper to understand content
3. **Detects speech** using Silero VAD (Voice Activity Detection)
4. **Removes silences** by cutting non-speech segments
5. **Adds jump cuts** between speech segments
6. **Optionally adds transitions** (crossfades) between cuts
7. **Enhances audio** (optional) - EQ, compression, loudness normalization
8. **Generates smart filename** based on video content
9. **Saves transcript** alongside the edited video

---

## Execution Scripts

| Script | Purpose |
|--------|---------|
| `video_editor_pipeline.py` | Main orchestration script - full pipeline |
| `jump_cut_vad_singlepass.py` | Standalone silence removal (VAD-based) |
| `insert_3d_transition.py` | Add 3D swivel transitions |

---

## CLI Arguments

### video_editor_pipeline.py

| Argument | Default | Description |
|----------|---------|-------------|
| `input` | required | Input video file path |
| `--output`, `-o` | auto | Output video file (auto-generated from content) |
| `--output-dir` | `.tmp` | Output directory |
| `--min-silence` | 0.5 | Minimum silence gap to cut (seconds) |
| `--min-speech` | 0.25 | Minimum speech duration to keep (seconds) |
| `--padding` | 100 | Padding around speech in milliseconds |
| `--merge-gap` | 0.3 | Merge segments closer than this (seconds) |
| `--keep-start` | true | Preserve intro (start from 0:00) |
| `--no-keep-start` | - | Allow cutting silence at the beginning |
| `--enhance-audio` | false | Apply audio enhancement chain |
| `--add-transitions` | false | Add crossfade transitions between cuts |
| `--whisper-model` | base | Whisper model size (tiny/base/small/medium/large) |
| `--no-transcribe` | false | Skip transcription, use generic filename |
| `--save-transcript` | true | Save transcript files alongside video |
| `--no-save-transcript` | - | Don't save transcript files |

---

## Output Files

For an input `recording.mp4`, the pipeline generates:

```
.tmp/
├── content_based_name.mp4           # Edited video
├── content_based_name_transcript.json  # Full transcript with timestamps
└── content_based_name_transcript.txt   # Plain text transcript
```

### Smart Filename Generation

The pipeline analyzes the transcript to create meaningful filenames:

| Video Content | Generated Filename |
|--------------|-------------------|
| "Today we're going to learn about machine learning..." | `machine_learning.mp4` |
| "Welcome to my tutorial on Python programming" | `tutorial_python_programming.mp4` |
| "Let me show you how to build a website" | `build_website.mp4` |

---

## Workflow Examples

### Basic Editing (Silence Removal)

```bash
python execution/video_editor_pipeline.py recording.mp4
```

Output: `.tmp/[smart_name].mp4` with silences removed

### Professional Quality

```bash
python execution/video_editor_pipeline.py recording.mp4 \
    --enhance-audio \
    --add-transitions \
    --whisper-model medium
```

- Better audio (EQ, compression, loudness normalization)
- Smooth crossfade transitions between cuts
- More accurate transcription for naming

### Aggressive Cutting (Fast-paced content)

```bash
python execution/video_editor_pipeline.py recording.mp4 \
    --min-silence 0.3 \
    --padding 50 \
    --merge-gap 0.2
```

### Relaxed Cutting (Preserve natural pauses)

```bash
python execution/video_editor_pipeline.py recording.mp4 \
    --min-silence 1.0 \
    --padding 200 \
    --merge-gap 0.5
```

---

## Audio Enhancement

When `--enhance-audio` is enabled:

```
highpass=f=80            # Remove rumble below 80Hz
lowpass=f=12000          # Remove harsh highs above 12kHz
equalizer (200Hz, -1dB)  # Reduce muddiness
equalizer (3kHz, +2dB)   # Boost presence/clarity
acompressor              # Gentle compression (3:1 ratio)
loudnorm=I=-16           # YouTube loudness standard (-16 LUFS)
```

---

## Dependencies

### System Requirements

```bash
# macOS
brew install ffmpeg

# Windows
winget install ffmpeg

# Linux
sudo apt install ffmpeg
```

### Python Dependencies

```bash
pip install torch           # For Silero VAD
pip install openai-whisper  # For transcription
pip install numpy           # General processing
```

---

## Performance

### Hardware Encoding

The pipeline automatically uses hardware encoding when available:

| Platform | Encoder | Speed |
|----------|---------|-------|
| macOS (Apple Silicon) | hevc_videotoolbox | 5-10x faster |
| Windows (NVIDIA) | hevc_nvenc | 5-10x faster |
| Fallback | libx265 | Baseline |

### Processing Time (Approximate)

| Video Length | Processing Time |
|--------------|-----------------|
| 5 min | ~30 seconds |
| 30 min | ~2-3 minutes |
| 60 min | ~5-8 minutes |

---

## Troubleshooting

### "No speech detected"
- Check that audio track exists: `ffprobe -i input.mp4`
- Try lowering `--min-speech` to 0.1
- Ensure audio isn't too quiet

### Cuts feel too aggressive
- Increase `--padding` to 150-200
- Increase `--min-silence` to 0.8-1.0

### Filename not descriptive
- Use `--whisper-model medium` for better transcription
- Check if video has clear speech

### Transcription slow
- Use `--whisper-model tiny` for faster (less accurate) results
- Use `--no-transcribe` to skip entirely

### FFmpeg errors
- Ensure FFmpeg is installed: `ffmpeg -version`
- Check input file isn't corrupted
- Try re-encoding input: `ffmpeg -i input.mp4 -c copy fixed.mp4`

---

## Integration with Other Scripts

### With 3D Transitions

```bash
# First: edit and remove silences
python execution/video_editor_pipeline.py raw.mp4 --output edited.mp4

# Then: add 3D swivel teaser at 3 second mark
python execution/insert_3d_transition.py edited.mp4 final.mp4
```

### With "Cut Cut" Restart Detection

For restart phrase detection, use the standalone script:

```bash
python execution/jump_cut_vad_singlepass.py input.mp4 output.mp4 --detect-restarts
```

---

## Example Output

```
🎬 Video Editor Pipeline
==================================================
   Input: recording.mp4

📏 Video duration: 300.0s (5.0 min)
🎵 Extracting audio...
📝 Transcribing audio with Whisper (base model)...
   ✓ Transcription complete (2847 chars)
   📛 Generated filename: tutorial_building_rest_api
   📁 Output: .tmp/tutorial_building_rest_api.mp4

🎯 Running Silero VAD...
   Found 42 speech segments
   Sample segments:
     1. 0.00s - 12.34s (12.34s)
     2. 13.21s - 28.45s (15.24s)
     3. 29.87s - 45.12s (15.25s)
     ... and 39 more
📎 After merging: 35 segments
🔲 After padding: 35 segments
📌 Extended first segment to start at 0:00
📊 Expected output: 245.3s (4.1 min)

⚡ Processing 35 segments...
🚀 Hardware encoding enabled (hevc_videotoolbox)
   ✓ Processed in 18.2s
   📄 Saved transcript: .tmp/tutorial_building_rest_api_transcript.json
   📄 Saved transcript: .tmp/tutorial_building_rest_api_transcript.txt

==================================================
✅ VIDEO EDITING COMPLETE!
==================================================
   📁 Output: .tmp/tutorial_building_rest_api.mp4
   ⏱️  Original: 300.0s (5.0 min)
   ⏱️  Edited: 245.3s (4.1 min)
   ✂️  Removed: 54.7s (18.2%)
   ⚡ Processing time: 45.3s
   📝 Transcript saved alongside video
```

---

## Best Practices

1. **Always preview** - Watch a portion of the output before processing long videos
2. **Start conservative** - Begin with default settings, then adjust
3. **Use hardware encoding** - Much faster on supported systems
4. **Save transcripts** - Useful for SEO, captions, and content repurposing
5. **Batch process** - Edit multiple videos with consistent settings

---

## Agent Instructions

When a user asks to edit a video:

1. **Confirm the input file exists** - Check the path is valid
2. **Ask about preferences** if not specified:
   - Do they want aggressive or relaxed cutting?
   - Should audio be enhanced?
   - Do they want transitions?
3. **Run the pipeline** with appropriate settings
4. **Report results** - Show filename, duration reduction, and file location
5. **Offer follow-up** - Suggest 3D transitions or other enhancements
