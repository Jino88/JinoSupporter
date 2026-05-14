"""DRM 우회 변환 CLI — 외부(C# 등)에서 subprocess 로 호출.

사용법:
    python excel_drm_cli.py --input <원본.xlsx> --output <대상.xlsx> [--mode clipboard]

결과:
    stdout 마지막 줄에 JSON 한 줄.
        성공: {"status":"ok","input":"...","output":"...","elapsed":12.3}
        실패: {"status":"fail","input":"...","error":"...","elapsed":0.5}
    진행 로그는 stderr 로.

Exit code: 0=성공, 1=변환 실패, 2=사용법 오류.
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

from excel_drm_clean import drm_clean_resave


def _emit(payload: dict) -> None:
    sys.stdout.write(json.dumps(payload, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def main() -> int:
    p = argparse.ArgumentParser(description="DRM 우회 xlsx 변환 (단일 파일)")
    p.add_argument("--input", "-i", required=True, help="원본 xlsx 경로")
    p.add_argument("--output", "-o", required=True, help="저장할 결과 xlsx 경로")
    p.add_argument(
        "--mode", "-m",
        default="clipboard",
        choices=["fast", "balanced", "precise", "clipboard"],
        help="변환 모드 (기본: clipboard — 모든 서식 + 빠름)",
    )
    args = p.parse_args()

    src = Path(args.input)
    dst = Path(args.output)
    t0 = time.perf_counter()

    if not src.is_file():
        _emit({
            "status": "fail",
            "input": str(src),
            "error": f"입력 파일 없음: {src}",
            "elapsed": 0.0,
        })
        return 2

    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
    except Exception as exc:
        _emit({
            "status": "fail",
            "input": str(src),
            "error": f"출력 폴더 생성 실패: {exc}",
            "elapsed": round(time.perf_counter() - t0, 3),
        })
        return 1

    def _log(msg: str) -> None:
        sys.stderr.write(str(msg) + "\n")
        sys.stderr.flush()

    def _noop(*_a, **_k) -> None:
        pass

    # drm_clean_resave 는 <output_root>/<src.stem>_clean.xlsx 로 저장.
    # 사용자가 원하는 정확한 출력 경로로 옮기기 위해 임시 폴더로 받는다.
    with tempfile.TemporaryDirectory(prefix="drm_clean_") as tmp:
        try:
            succeeded, failed = drm_clean_resave(
                src_paths=[src],
                output_root=tmp,
                log=_log,
                on_progress=_noop,
                format_mode=args.mode,
            )
        except Exception as exc:
            _emit({
                "status": "fail",
                "input": str(src),
                "error": f"변환 예외: {exc}",
                "elapsed": round(time.perf_counter() - t0, 3),
            })
            return 1

        if failed:
            _, err = failed[0]
            _emit({
                "status": "fail",
                "input": str(src),
                "error": str(err),
                "elapsed": round(time.perf_counter() - t0, 3),
            })
            return 1

        if not succeeded:
            _emit({
                "status": "fail",
                "input": str(src),
                "error": "결과 없음 (skip 처리됨)",
                "elapsed": round(time.perf_counter() - t0, 3),
            })
            return 1

        _, produced = succeeded[0]
        produced_p = Path(produced)
        if not produced_p.is_file():
            _emit({
                "status": "fail",
                "input": str(src),
                "error": f"생성 파일 없음: {produced_p}",
                "elapsed": round(time.perf_counter() - t0, 3),
            })
            return 1

        try:
            if dst.exists():
                dst.unlink()
            shutil.move(str(produced_p), str(dst))
        except Exception as exc:
            _emit({
                "status": "fail",
                "input": str(src),
                "error": f"결과 이동 실패: {exc}",
                "elapsed": round(time.perf_counter() - t0, 3),
            })
            return 1

    _emit({
        "status": "ok",
        "input": str(src),
        "output": str(dst),
        "elapsed": round(time.perf_counter() - t0, 3),
    })
    return 0


if __name__ == "__main__":
    sys.exit(main())
