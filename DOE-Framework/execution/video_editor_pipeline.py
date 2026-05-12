#!/usr/bin/env python3
"""
Video Editor Pipeline

Complete video editing pipeline that:
1. Removes silences using VAD (Voice Activity Detection)
2. Adds jump cuts between speech segments
3. Optionally adds transitions
4. Transcribes the video and generates smart output filename
5. Applies audio enhancement

Usage:
    python execution/video_editor_pipeline.py input.mp4
    python execution/video_editor_pipeline.py input.mp4 --output-dir .tmp/edited
    python execution/video_editor_pipeline.py input.mp4 --enhance-audio --add-transitions
"""

import subprocess
import tempfile
import os
import argparse
import time
import re
import json
from pathlib import Path
from datetime import datetime

# Video encoding settings
HARDWARE_ENCODER = "hevc_videotoolbox"
SOFTWARE_ENCODER = "libx265"
HARDWARE_BITRATE = "17M"
SOFTWARE_CRF = "18"
TARGET_FPS = 30

_hardware_encoder_available = None


def check_hardware_encoder_available() -> bool:
    """Check if hardware encoder is available."""
    try:
        result = subprocess.run(
            ["ffmpeg", "-hide_banner", "-encoders"],
            capture_output=True, text=True, timeout=5
        )
        return "hevc_videotoolbox" in result.stdout
    except Exception:
        return False


def get_encoder_args() -> list[str]:
    """Get encoder arguments based on hardware availability."""
    global _hardware_encoder_available
    if _hardware_encoder_available is None:
        _hardware_encoder_available = check_hardware_encoder_available()
        if _hardware_encoder_available:
            print(f"🚀 Hardware encoding enabled (hevc_videotoolbox)")
        else:
            print(f"💻 Using software encoding (libx265)")

    if _hardware_encoder_available:
        return ["-c:v", HARDWARE_ENCODER, "-b:v", HARDWARE_BITRATE, "-r", str(TARGET_FPS), "-tag:v", "hvc1"]
    else:
        return ["-c:v", SOFTWARE_ENCODER, "-preset", "fast", "-crf", SOFTWARE_CRF, "-r", str(TARGET_FPS), "-tag:v", "hvc1"]


def extract_audio(input_path: str, output_path: str, sample_rate: int = 16000):
    """Extract audio from video as WAV."""
    cmd = [
        "ffmpeg", "-y", "-i", input_path,
        "-vn", "-ar", str(sample_rate), "-ac", "1",
        "-acodec", "pcm_s16le",
        "-loglevel", "error", output_path
    ]
    subprocess.run(cmd, capture_output=True, check=True)


def get_duration(input_path: str) -> float:
    """Get video duration in seconds."""
    cmd = [
        "ffprobe", "-v", "error", "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1", input_path
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    return float(result.stdout.strip())


def transcribe_video(audio_path: str, model_name: str = "base") -> dict:
    """
    Transcribe audio using OpenAI Whisper.
    
    Returns:
        dict with 'text' (full transcript), 'segments' (timestamped segments),
        and 'language' (detected language)
    """
    print(f"📝 Transcribing audio with Whisper ({model_name} model)...")
    
    try:
        import whisper
    except ImportError:
        print("   ⚠️  Whisper not installed. Installing...")
        subprocess.run(["pip", "install", "openai-whisper"], check=True)
        import whisper
    
    model = whisper.load_model(model_name)
    result = model.transcribe(audio_path, verbose=False)
    
    print(f"   ✓ Transcription complete ({len(result['text'])} chars)")
    return result


def generate_smart_filename(transcript: str, max_words: int = 6) -> str:
    """
    Generate a smart filename based on video transcript.
    
    Uses the first meaningful sentence or topic from the transcript
    to create a descriptive filename.
    """
    if not transcript or len(transcript.strip()) < 10:
        return f"edited_video_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    # Clean up the transcript
    text = transcript.strip()
    
    # Try to get the first sentence
    sentences = re.split(r'[.!?]', text)
    first_sentence = sentences[0].strip() if sentences else text[:100]
    
    # Get key words (remove common words)
    stop_words = {
        'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for',
        'of', 'with', 'by', 'from', 'as', 'is', 'was', 'are', 'were', 'been',
        'be', 'have', 'has', 'had', 'do', 'does', 'did', 'will', 'would',
        'could', 'should', 'may', 'might', 'must', 'shall', 'can', 'this',
        'that', 'these', 'those', 'i', 'you', 'he', 'she', 'it', 'we', 'they',
        'what', 'which', 'who', 'when', 'where', 'why', 'how', 'all', 'each',
        'every', 'both', 'few', 'more', 'most', 'other', 'some', 'such', 'no',
        'nor', 'not', 'only', 'own', 'same', 'so', 'than', 'too', 'very', 'just',
        'um', 'uh', 'like', 'you know', 'basically', 'actually', 'really',
        'gonna', 'wanna', 'gotta', 'okay', 'ok', 'hey', 'hi', 'hello', 'well',
        'right', 'yeah', 'yes', 'no', 'maybe', 'think', 'know', 'see', 'look',
        'come', 'go', 'get', 'make', 'take', 'say', 'said', 'says', 'let', "let's",
        "i'm", "i'll", "i've", "it's", "that's", "there's", "here's", "what's",
        "don't", "doesn't", "didn't", "won't", "wouldn't", "couldn't", "shouldn't",
        "can't", "haven't", "hasn't", "hadn't", "isn't", "aren't", "wasn't", "weren't"
    }
    
    # Extract meaningful words
    words = re.findall(r'\b[a-zA-Z]{3,}\b', first_sentence.lower())
    meaningful_words = [w for w in words if w not in stop_words][:max_words]
    
    if not meaningful_words:
        # Fallback: just use first few words
        words = re.findall(r'\b[a-zA-Z]+\b', first_sentence.lower())[:max_words]
        meaningful_words = words
    
    if not meaningful_words:
        return f"edited_video_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    
    # Create filename
    filename = "_".join(meaningful_words)
    # Clean up: only alphanumeric and underscores
    filename = re.sub(r'[^a-z0-9_]', '', filename)
    # Limit length
    filename = filename[:50]
    
    return filename


def get_speech_timestamps_silero(audio_path: str, min_speech_duration: float = 0.25, 
                                  min_silence_duration: float = 0.5) -> list:
    """Use Silero VAD to detect speech segments."""
    print(f"🎯 Running Silero VAD...")
    
    import torch
    model, utils = torch.hub.load(
        repo_or_dir='snakers4/silero-vad',
        model='silero_vad',
        force_reload=False,
        trust_repo=True
    )

    (get_speech_timestamps, _, read_audio, _, _) = utils

    SAMPLE_RATE = 16000
    wav = read_audio(audio_path, sampling_rate=SAMPLE_RATE)

    speech_timestamps = get_speech_timestamps(
        wav,
        model,
        sampling_rate=SAMPLE_RATE,
        threshold=0.5,
        min_speech_duration_ms=int(min_speech_duration * 1000),
        min_silence_duration_ms=int(min_silence_duration * 1000),
        speech_pad_ms=100,
    )

    segments = []
    for ts in speech_timestamps:
        start_sec = ts['start'] / SAMPLE_RATE
        end_sec = ts['end'] / SAMPLE_RATE
        segments.append((start_sec, end_sec))

    print(f"   Found {len(segments)} speech segments")
    return segments


def merge_close_segments(segments: list, max_gap: float) -> list:
    """Merge segments that are very close together."""
    if not segments:
        return []

    merged = [segments[0]]
    for start, end in segments[1:]:
        prev_start, prev_end = merged[-1]
        if start - prev_end <= max_gap:
            merged[-1] = (prev_start, end)
        else:
            merged.append((start, end))

    return merged


def add_padding(segments: list, padding_s: float, duration: float) -> list:
    """Add padding around segments and merge overlaps."""
    if not segments:
        return []

    padded = []
    for start, end in segments:
        new_start = max(0, start - padding_s)
        new_end = min(duration, end + padding_s)
        padded.append((new_start, new_end))

    merged = [padded[0]]
    for start, end in padded[1:]:
        prev_start, prev_end = merged[-1]
        if start <= prev_end:
            merged[-1] = (prev_start, max(prev_end, end))
        else:
            merged.append((start, end))

    return merged


def build_trim_concat_filter(segments: list, add_transitions: bool = False) -> str:
    """Build FFmpeg filter for trim+concat with optional transitions."""
    n = len(segments)
    filter_parts = []

    if add_transitions and n > 1:
        # With crossfade transitions between segments
        transition_duration = 0.15  # 150ms crossfade
        
        # Video trim filters
        for i, (start, end) in enumerate(segments):
            filter_parts.append(
                f"[0:v]trim=start={start:.6f}:end={end:.6f},setpts=PTS-STARTPTS[v{i}]"
            )
        
        # Audio trim filters
        for i, (start, end) in enumerate(segments):
            filter_parts.append(
                f"[0:a]atrim=start={start:.6f}:end={end:.6f},asetpts=PTS-STARTPTS[a{i}]"
            )
        
        # Chain crossfade for video
        current_v = "[v0]"
        for i in range(1, n):
            next_v = f"[v{i}]"
            out_v = f"[xv{i}]" if i < n-1 else "[outv]"
            filter_parts.append(
                f"{current_v}{next_v}xfade=transition=fade:duration={transition_duration}:offset={i * 0.5}{out_v}"
            )
            current_v = out_v
        
        # Chain crossfade for audio
        current_a = "[a0]"
        for i in range(1, n):
            next_a = f"[a{i}]"
            out_a = f"[xa{i}]" if i < n-1 else "[outa]"
            filter_parts.append(
                f"{current_a}{next_a}acrossfade=d={transition_duration}{out_a}"
            )
            current_a = out_a
    else:
        # Simple concat without transitions
        for i, (start, end) in enumerate(segments):
            filter_parts.append(
                f"[0:v]trim=start={start:.6f}:end={end:.6f},setpts=PTS-STARTPTS[v{i}]"
            )
        
        for i, (start, end) in enumerate(segments):
            filter_parts.append(
                f"[0:a]atrim=start={start:.6f}:end={end:.6f},asetpts=PTS-STARTPTS[a{i}]"
            )
        
        concat_inputs = "".join(f"[v{i}][a{i}]" for i in range(n))
        filter_parts.append(f"{concat_inputs}concat=n={n}:v=1:a=1[outv][outa]")

    return ";".join(filter_parts)


def build_audio_enhancement_filter() -> str:
    """Build FFmpeg audio enhancement filter chain."""
    return (
        "highpass=f=80,"
        "lowpass=f=12000,"
        "equalizer=f=200:t=q:w=1:g=-1,"
        "equalizer=f=3000:t=q:w=1:g=2,"
        "acompressor=threshold=-20dB:ratio=3:attack=5:release=50,"
        "loudnorm=I=-16:TP=-1.5:LRA=11"
    )


def process_video(
    input_path: str,
    output_path: str,
    segments: list,
    enhance_audio: bool = False,
    add_transitions: bool = False
) -> bool:
    """Process video with FFmpeg using trim+concat."""
    print(f"⚡ Processing {len(segments)} segments...")
    start_time = time.time()

    encoder_args = get_encoder_args()
    filter_complex = build_trim_concat_filter(segments, add_transitions)
    
    # Add audio enhancement if requested
    if enhance_audio:
        audio_filter = build_audio_enhancement_filter()
        filter_complex = filter_complex.replace("[outa]", f"[outa_pre];[outa_pre]{audio_filter}[outa]")
        print("   🎧 Audio enhancement enabled")

    # Write filter to temp file (handles long filter strings)
    with tempfile.NamedTemporaryFile(mode='w', suffix='.txt', delete=False) as f:
        f.write(filter_complex)
        filter_script_path = f.name

    try:
        cmd = [
            "ffmpeg", "-y",
            "-i", input_path,
            "-filter_complex_script", filter_script_path,
            "-map", "[outv]", "-map", "[outa]",
        ]
        cmd.extend(encoder_args)
        cmd.extend([
            "-c:a", "aac", "-b:a", "192k",
            "-movflags", "+faststart",
            "-loglevel", "error",
            output_path
        ])

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode != 0:
            print(f"   ❌ FFmpeg error: {result.stderr[:1000]}")
            return False

        elapsed = time.time() - start_time
        print(f"   ✓ Processed in {elapsed:.1f}s")
        return True
    finally:
        if os.path.exists(filter_script_path):
            os.remove(filter_script_path)


def save_transcript(transcript_data: dict, output_path: str):
    """Save transcript to JSON and TXT files."""
    base_path = Path(output_path).with_suffix('')
    
    # Save full transcript as JSON
    json_path = str(base_path) + "_transcript.json"
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(transcript_data, f, indent=2, ensure_ascii=False)
    print(f"   📄 Saved transcript: {json_path}")
    
    # Save plain text transcript
    txt_path = str(base_path) + "_transcript.txt"
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(transcript_data['text'])
    print(f"   📄 Saved text: {txt_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Complete video editing pipeline - removes silences, adds jump cuts, smart naming",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python video_editor_pipeline.py input.mp4
    python video_editor_pipeline.py input.mp4 --output-dir ./edited
    python video_editor_pipeline.py input.mp4 --enhance-audio --add-transitions
    python video_editor_pipeline.py input.mp4 --min-silence 0.8 --padding 150
        """
    )
    
    parser.add_argument("input", help="Input video file")
    parser.add_argument("--output", "-o", help="Output video file (auto-generated if not specified)")
    parser.add_argument("--output-dir", default=".tmp", help="Output directory (default: .tmp)")
    parser.add_argument("--min-silence", type=float, default=0.5,
                        help="Minimum silence duration to cut (default: 0.5s)")
    parser.add_argument("--min-speech", type=float, default=0.25,
                        help="Minimum speech duration to keep (default: 0.25s)")
    parser.add_argument("--padding", type=int, default=100,
                        help="Padding around speech in ms (default: 100)")
    parser.add_argument("--merge-gap", type=float, default=0.3,
                        help="Merge segments closer than this (default: 0.3s)")
    parser.add_argument("--keep-start", action="store_true", default=True,
                        help="Always start from 0:00 (default: True)")
    parser.add_argument("--no-keep-start", action="store_false", dest="keep_start")
    parser.add_argument("--enhance-audio", action="store_true",
                        help="Apply audio enhancement (EQ, compression, loudness)")
    parser.add_argument("--add-transitions", action="store_true",
                        help="Add crossfade transitions between jump cuts")
    parser.add_argument("--whisper-model", default="base",
                        choices=["tiny", "base", "small", "medium", "large"],
                        help="Whisper model for transcription (default: base)")
    parser.add_argument("--no-transcribe", action="store_true",
                        help="Skip transcription (use generic filename)")
    parser.add_argument("--save-transcript", action="store_true", default=True,
                        help="Save transcript to file (default: True)")
    parser.add_argument("--no-save-transcript", action="store_false", dest="save_transcript")

    args = parser.parse_args()
    
    input_path = args.input
    if not os.path.exists(input_path):
        print(f"❌ Input file not found: {input_path}")
        return 1

    print(f"🎬 Video Editor Pipeline")
    print(f"=" * 50)
    print(f"   Input: {input_path}")
    print()

    overall_start = time.time()
    
    # Get video duration
    duration = get_duration(input_path)
    print(f"📏 Video duration: {duration:.1f}s ({duration/60:.1f} min)")

    # Create output directory
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Extract audio
    print(f"🎵 Extracting audio...")
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
        audio_path = tmp.name

    try:
        extract_audio(input_path, audio_path)

        # Transcribe for smart naming
        transcript_data = None
        smart_name = None
        
        if not args.no_transcribe:
            transcript_data = transcribe_video(audio_path, args.whisper_model)
            smart_name = generate_smart_filename(transcript_data['text'])
            print(f"   📛 Generated filename: {smart_name}")
        
        # Determine output path
        if args.output:
            output_path = args.output
        else:
            if smart_name:
                output_path = str(output_dir / f"{smart_name}.mp4")
            else:
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                output_path = str(output_dir / f"edited_{timestamp}.mp4")
        
        print(f"   📁 Output: {output_path}")
        print()

        # Get speech segments using VAD
        speech_segments = get_speech_timestamps_silero(
            audio_path,
            min_speech_duration=args.min_speech,
            min_silence_duration=args.min_silence
        )

        if not speech_segments:
            print("⚠️  No speech detected in video!")
            return 1

        # Show first few segments
        print("   Sample segments:")
        for i, (start, end) in enumerate(speech_segments[:3]):
            print(f"     {i+1}. {start:.2f}s - {end:.2f}s ({end-start:.2f}s)")
        if len(speech_segments) > 3:
            print(f"     ... and {len(speech_segments) - 3} more")

    finally:
        if os.path.exists(audio_path):
            os.remove(audio_path)

    # Merge close segments
    speech_segments = merge_close_segments(speech_segments, args.merge_gap)
    print(f"📎 After merging: {len(speech_segments)} segments")

    # Add padding
    padding_s = args.padding / 1000
    speech_segments = add_padding(speech_segments, padding_s, duration)
    print(f"🔲 After padding: {len(speech_segments)} segments")

    # Keep start if requested
    if args.keep_start and speech_segments and speech_segments[0][0] > 0:
        first_start, first_end = speech_segments[0]
        speech_segments[0] = (0.0, first_end)
        print(f"📌 Extended first segment to start at 0:00")

    # Calculate expected output
    total_speech = sum(end - start for start, end in speech_segments)
    print(f"📊 Expected output: {total_speech:.1f}s ({total_speech/60:.1f} min)")
    print()

    # Process video
    success = process_video(
        input_path,
        output_path,
        speech_segments,
        enhance_audio=args.enhance_audio,
        add_transitions=args.add_transitions
    )

    if not success:
        print("❌ Video processing failed!")
        return 1

    # Save transcript if requested
    if transcript_data and args.save_transcript:
        save_transcript(transcript_data, output_path)

    # Final stats
    new_duration = get_duration(output_path)
    removed = duration - new_duration
    overall_time = time.time() - overall_start

    print()
    print(f"{'=' * 50}")
    print(f"✅ VIDEO EDITING COMPLETE!")
    print(f"{'=' * 50}")
    print(f"   📁 Output: {output_path}")
    print(f"   ⏱️  Original: {duration:.1f}s ({duration/60:.1f} min)")
    print(f"   ⏱️  Edited: {new_duration:.1f}s ({new_duration/60:.1f} min)")
    print(f"   ✂️  Removed: {removed:.1f}s ({100*removed/duration:.1f}%)")
    print(f"   ⚡ Processing time: {overall_time:.1f}s")
    
    if transcript_data:
        print(f"   📝 Transcript saved alongside video")

    return 0


if __name__ == "__main__":
    exit(main())
