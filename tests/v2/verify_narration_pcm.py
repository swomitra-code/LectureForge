"""Read-only exact prepared-silence verification for the Milestone B test root."""
import json
import subprocess
import sys
import wave
from pathlib import Path


def verify(raw, prepared):
    speech = subprocess.check_output([
        "ffmpeg", "-v", "error", "-i", str(raw), "-ar", "44100", "-ac", "1",
        "-f", "s16le", "-",
    ])
    with wave.open(str(prepared), "rb") as audio:
        assert (audio.getframerate(), audio.getnchannels(), audio.getsampwidth()) == (44100, 1, 2)
        pcm = audio.readframes(audio.getnframes())
    assert pcm == speech + bytes(264600 * 2), "Speech or exact six-second pause mismatch"


root = Path(sys.argv[1])
live = json.loads((root / "live-narration-results.json").read_text(encoding="utf-8-sig"))
project = root / "Projects" / live["project_id"]
for take in live["takes"]:
    verify(project / take["attempts"][-1]["raw_path"], project / take["asset"]["path"])
    print(f"PASS: live Take {take['number']} preserves decoded speech + exactly 6.000 s silence")
for file in (root / "Tests").glob("*/pause.json"):
    pause = json.loads(file.read_text(encoding="utf-8-sig"))
    verify(pause["raw"], pause["prepared"])
    print("PASS: copied fixture preparation preserves decoded speech + exactly 6.000 s silence")
