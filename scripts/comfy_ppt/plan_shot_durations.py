# -*- coding: utf-8 -*-
"""从 slides.json 生成逐镜时长规划，给 huiben-01-creater 做视频排期。

用法：
    python plan_shot_durations.py slides.json --out shot-duration-plan.json
    python plan_shot_durations.py slides.json --out shot-duration-plan.json --check-plan

规划不是配音实测时长，而是生成视频前的字数预算；TTS 合成后必须回填
measuredAudioSec / requiredAtempo，并重新检查语速门禁。
"""
import argparse
import json
import math
import re
from pathlib import Path


DEFAULT_CPS = 3.5
MAX_ATEMPO = 1.15
MIN_SHOT_SEC = 3.0
MAX_SHOT_SEC = 12.0
FPS = 24
FRAME_STEP = 17
FRAME_OFFSET = 5
GAP_SEC = 0.15
PAD_SEC = 0.25


def count_speech_chars(text):
    return len(re.findall(r"[\u3400-\u9fffA-Za-z0-9]", "".join(text)))


def snap_frames(seconds):
    frames = max(1, round(seconds * FPS))
    aligned = ((frames - FRAME_OFFSET + FRAME_STEP - 1) // FRAME_STEP) * FRAME_STEP + FRAME_OFFSET
    return aligned


def snap_frames_within(seconds, max_seconds):
    max_frames = int(max_seconds * FPS)
    max_aligned = ((max_frames - FRAME_OFFSET) // FRAME_STEP) * FRAME_STEP + FRAME_OFFSET
    wanted = snap_frames(seconds)
    return min(wanted, max_aligned)


def make_plan(config, cps, max_atempo, min_sec, max_sec, gap, pad):
    pages = []
    for slide in config["slides"]:
        text = [slide["title"]] if slide.get("kind") == "cover" else slide.get("text", [])
        text = [line for line in text if isinstance(line, str)]
        speech_chars = count_speech_chars(text)
        speaker_switches = sum(
            1
            for line in text
            if "：" in line or "“" in line or "”" in line
        )
        punctuation_pauses = len(re.findall(r"[，。！？；、：]", "".join(text))) * 0.08
        estimated_tts = speech_chars / cps + punctuation_pauses + speaker_switches * 0.18
        slot_seconds = estimated_tts + gap + pad
        slot_seconds = min(max(slot_seconds, min_sec), max_sec)
        frames = snap_frames_within(slot_seconds, max_sec)
        capped = frames / FPS < slot_seconds - 0.01
        pages.append(
            {
                "id": slide["id"],
                "kind": slide.get("kind", "content"),
                "text": text,
                "speechChars": speech_chars,
                "lineCount": len(text),
                "speakerSwitchCount": speaker_switches,
                "estimatedTtsSec": round(estimated_tts, 3),
                "plannedSlotSec": round(frames / FPS, 3),
                "plannedFrames": frames,
                "h3FrameRule": "17n+5",
                "maxNaturalAtempo": max_atempo,
                "maxSpeechSec": round(frames / FPS * max_atempo, 3),
                "needsSplit": capped,
                "status": "needs-split" if capped else "planned",
            }
        )
    return {
        "schema": "huiben-shot-duration-plan/v1",
        "source": "slides.json",
        "assumptions": {
            "narrationCharsPerSec": cps,
            "maxNaturalAtempo": max_atempo,
            "minShotSec": min_sec,
            "maxShotSec": max_sec,
            "fps": FPS,
            "frameRule": "17n+5",
            "gapSec": gap,
            "padSec": pad,
            "note": "estimatedTtsSec uses 3.5 Chinese chars/sec (picture-book narration); measuredAudioSec must be filled after synthesis.",
        },
        "totals": {
            "pageCount": len(pages),
            "speechChars": sum(page["speechChars"] for page in pages),
            "plannedSec": round(sum(page["plannedSlotSec"] for page in pages), 3),
            "minPageChars": min(page["speechChars"] for page in pages),
            "maxPageChars": max(page["speechChars"] for page in pages),
            "minPlannedSec": min(page["plannedSlotSec"] for page in pages),
            "maxPlannedSec": max(page["plannedSlotSec"] for page in pages),
            "needsSplitPages": [page["id"] for page in pages if page["needsSplit"]],
        },
        "pages": pages,
    }


def check_plan(plan):
    max_atempo = plan["assumptions"]["maxNaturalAtempo"]
    max_sec = plan["assumptions"]["maxShotSec"]
    issues = []
    seen = set()
    for page in plan["pages"]:
        shot_id = page["id"]
        if shot_id in seen:
            issues.append(f"duplicate shot id: {shot_id}")
        seen.add(shot_id)
        frames = page["plannedFrames"]
        if (frames - FRAME_OFFSET) % FRAME_STEP != 0:
            issues.append(f"{shot_id}: {frames} frames violates 17n+5")
        slot = page["plannedSlotSec"]
        if slot > max_sec + 1e-6:
            issues.append(f"{shot_id}: planned slot {slot}s exceeds max {max_sec}s")
        if page.get("needsSplit"):
            issues.append(
                f"{shot_id}: text budget exceeds one {max_sec}s shot; split the page "
                "or obtain user approval for a longer shot"
            )
        measured = page.get("measuredAudioSec")
        if measured is None:
            continue
        available_speech = slot - plan["assumptions"]["gapSec"] - plan["assumptions"]["padSec"]
        required_atempo = measured / max(available_speech, 0.01)
        page["requiredAtempo"] = round(required_atempo, 3)
        page["measuredAudioSec"] = round(float(measured), 3)
        page["status"] = "measured"
        if required_atempo > max_atempo + 1e-6:
            issues.append(
                f"{shot_id}: requires atempo {required_atempo:.2f} > {max_atempo}; "
                "lengthen the shot or split the page"
            )
        if page.get("needsSplit") and page.get("splitConfirmed") is not True and page.get("extendedShotConfirmed") is not True:
            issues.append(
                f"{shot_id}: unresolved needsSplit; split the page or set "
                "extendedShotConfirmed=true after user approval"
            )
    plan["check"] = {
        "passed": not issues,
        "issues": issues,
        "maxObservedAtempo": max(
            (page.get("requiredAtempo", 0) for page in plan["pages"]), default=0
        ),
    }
    return plan


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("slides")
    parser.add_argument("--out", required=True)
    parser.add_argument("--cps", type=float, default=DEFAULT_CPS)
    parser.add_argument("--max-atempo", type=float, default=MAX_ATEMPO)
    parser.add_argument("--min-sec", type=float, default=MIN_SHOT_SEC)
    parser.add_argument("--max-sec", type=float, default=MAX_SHOT_SEC)
    parser.add_argument("--gap", type=float, default=GAP_SEC)
    parser.add_argument("--pad", type=float, default=PAD_SEC)
    parser.add_argument("--check-plan", action="store_true")
    args = parser.parse_args()

    slides_path = Path(args.slides).resolve()
    output_path = Path(args.out).resolve()
    config = json.loads(slides_path.read_text(encoding="utf-8"))
    plan = make_plan(
        config,
        args.cps,
        args.max_atempo,
        args.min_sec,
        args.max_sec,
        args.gap,
        args.pad,
    )
    plan = check_plan(plan)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(plan, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        f"saved {output_path}: {plan['totals']['pageCount']} pages, "
        f"{plan['totals']['plannedSec']}s, checkPassed={plan['check']['passed']}"
    )
    if not plan["check"]["passed"]:
        for issue in plan["check"]["issues"]:
            print("ISSUE", issue)
        raise SystemExit(2)


if __name__ == "__main__":
    main()
