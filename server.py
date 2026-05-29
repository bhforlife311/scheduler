"""역량평가 시간표 — 로컬 웹 서버.

Usage:
    python server.py

브라우저가 자동으로 열려서 input.html 폼을 보여줍니다.
폼 입력 → "시간표 생성" 버튼 → xlsx 자동 다운로드.

종료: 터미널에서 Ctrl+C.
"""

from __future__ import annotations

import io
import sys
import tempfile
import threading
import time
import webbrowser
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory

# 같은 폴더의 scheduler.py 재사용
sys.path.insert(0, str(Path(__file__).parent))
from scheduler import (  # noqa: E402
    Config, Room, Technique, WaitingRoom,
    _validate_config, best_schedule, filter_by_load,
    generate_sg_candidates, parse_time_str, write_xlsx,
)

# Windows 콘솔 한글 깨짐 방지
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", line_buffering=True)


ROOT = Path(__file__).parent
app = Flask(__name__, static_folder=str(ROOT))


# ───────────────── 라우팅 ─────────────────


@app.route("/")
def index():
    return send_from_directory(str(ROOT), "input.html")


@app.route("/visualizer.html")
def visualizer():
    return send_from_directory(str(ROOT), "visualizer.html")


@app.route("/health")
def health():
    """input.html이 서버 동작 여부를 감지하는 데 사용."""
    return jsonify({"ok": True})


@app.route("/generate", methods=["POST"])
def generate():
    """JSON config 받아서 시간표 xlsx 생성 후 응답."""
    try:
        raw = request.get_json(force=True)
        if not raw:
            return jsonify({"error": "JSON 본문이 비어있음"}), 400
        cfg = _cfg_from_dict(raw)
    except Exception as e:
        return jsonify({"error": f"입력 파싱 실패: {type(e).__name__}: {e}"}), 400

    try:
        sg_all = generate_sg_candidates(cfg)
        sg_filtered = filter_by_load(sg_all)
        if not sg_filtered:
            return jsonify({
                "error": "(s, G) 후보를 못 찾았어요. 대기실/숙지실 수, 그룹 크기 제약을 확인해 주세요."
            }), 400

        best = best_schedule(cfg, sg_filtered)
        if best is None:
            return jsonify({
                "error": "모든 후보에서 스케줄링 실패. 평가실/숙지실/위원 자원이 부족하거나, 점심 윈도우가 좁아서일 수 있어요."
            }), 400

        # xlsx 를 임시파일에 쓰고 메모리에 적재
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = Path(tmp.name)
        try:
            write_xlsx(best, cfg, tmp_path)
            data = tmp_path.read_bytes()
        finally:
            tmp_path.unlink(missing_ok=True)

        safe_title = (cfg.title or "schedule")
        for ch in '\\/:*?"<>|':
            safe_title = safe_title.replace(ch, "_")
        safe_title = safe_title.replace(" ", "_")

        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"{safe_title}_schedule.xlsx",
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"스케줄링 오류: {type(e).__name__}: {e}"}), 500


# ───────────────── 헬퍼 ─────────────────


def _cfg_from_dict(raw: dict) -> Config:
    techniques = []
    for t in raw["techniques"]:
        # 하위호환: duration_min 만 있는 옛 포맷도 받음
        if "duration_min" in t and "eval_duration_min" not in t:
            prep_d = 0
            eval_d = int(t["duration_min"])
        else:
            prep_d = int(t.get("prep_duration_min", 0))
            eval_d = int(t.get("eval_duration_min", 0))
        techniques.append(Technique(
            name=t["name"],
            assessors=int(t["assessors"]),
            candidates=int(t["candidates"]),
            prep_duration_min=prep_d,
            eval_duration_min=eval_d,
        ))

    def mk_room(r: dict) -> Room:
        return Room(
            name=r["name"],
            supported=tuple(r["supported"]),
            capacity=int(r["capacity"]),
            zone=r["zone"],
        )

    rooms = [mk_room(r) for r in raw["rooms"]]
    prep_rooms = [mk_room(r) for r in raw.get("prep_rooms", [])]
    waiting_rooms = [WaitingRoom(name=w["name"], zone=w["zone"])
                     for w in raw.get("waiting_rooms", [])]
    lw = raw.get("lunch_window", ["11:00", "14:00"])

    cfg = Config(
        title=raw.get("title", "역량평가"),
        total_candidates=int(raw["total_candidates"]),
        techniques=techniques,
        start_time_min=parse_time_str(raw.get("start_time", "09:00")),
        transition_min=int(raw.get("transition_min", 10)),
        prep_to_eval_transition_min=int(raw.get("prep_to_eval_transition_min", 5)),
        lunch_min=int(raw.get("lunch_min", 60)),
        lunch_window=(parse_time_str(lw[0]), parse_time_str(lw[1])),
        assessor_mode=raw.get("assessor_mode", "universal"),
        assessor_count_by_technique=raw.get("assessor_count_by_technique", {}),
        rooms=rooms,
        prep_rooms=prep_rooms,
        waiting_rooms=waiting_rooms,
    )
    _validate_config(cfg)
    return cfg


def _open_browser_later(url: str, delay: float = 1.2) -> None:
    def _go():
        time.sleep(delay)
        try:
            webbrowser.open(url)
        except Exception:
            pass
    threading.Thread(target=_go, daemon=True).start()


# ───────────────── 엔트리 포인트 ─────────────────


def main() -> int:
    host = "127.0.0.1"
    port = 5050
    url = f"http://{host}:{port}/"
    print("=" * 56)
    print(" 역량평가 시간표 서버")
    print(f"   주소: {url}")
    print("   브라우저가 곧 자동으로 열립니다...")
    print("   종료: 이 터미널에서 Ctrl+C")
    print("=" * 56)
    _open_browser_later(url)
    # debug=False 권장 (디버그 리로더는 브라우저 중복 오픈 일으킴)
    app.run(host=host, port=port, debug=False)
    return 0


if __name__ == "__main__":
    sys.exit(main())
