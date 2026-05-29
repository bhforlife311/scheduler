"""역량평가 시간표 자동 생성 도구.

Usage:
    python scheduler.py sample_config.json -o schedule.xlsx

설계 개요
---------
B안 (위원수 = 그룹수, 그룹 엇갈리게 진행) 기반.

흐름:
  1) (s, G) 후보 탐색  — 그룹 크기/그룹 수 가능 조합
  2) 위원 부하 8~9세션 1차 필터
  3) 각 후보에 대해 그리디 + 백트래킹 스케줄링 시뮬레이션
  4) 점수 함수: (zone 충돌=0) → |A/G - 8.5| → 총 종료시각+평균대기 순
  5) visualizer.html 호환 xlsx 출력
"""

from __future__ import annotations

import argparse
import copy
import io
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path

# Windows 콘솔 한글 깨짐 방지
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", line_buffering=True)


SLOT_MIN = 5  # 5분 슬롯 단위 (visualizer.html과 동일)


# ───────────────────────── 데이터 모델 ─────────────────────────


@dataclass(frozen=True)
class Technique:
    name: str
    assessors: int                # 한 세션의 위원 수 (a_k)
    candidates: int               # 한 세션의 대상자 수 (c_k)
    prep_duration_min: int        # 숙지 시간 (그룹 전체가 숙지실에서, 위원 없음)
    eval_duration_min: int        # 평가 시간 (그룹원 분산, 위원과 평가)

    @property
    def load_per_candidate(self) -> float:
        """대상자 1명이 이 기법을 받을 때 발생시키는 위원 세션 부하 = a/c."""
        return self.assessors / self.candidates

    @property
    def total_duration_min(self) -> int:
        return self.prep_duration_min + self.eval_duration_min


@dataclass(frozen=True)
class Room:
    name: str
    supported: tuple[str, ...]   # 가능한 기법 이름들
    capacity: int
    zone: str

    def supports(self, technique_name: str) -> bool:
        return technique_name in self.supported


@dataclass(frozen=True)
class WaitingRoom:
    name: str
    zone: str


@dataclass
class Config:
    title: str
    total_candidates: int
    techniques: list[Technique]
    start_time_min: int               # 9:00 → 540
    transition_min: int               # 기법 → 기법 이동 (평가 끝 → 다음 기법 숙지 시작 사이)
    prep_to_eval_transition_min: int  # 숙지 → 평가 이동 (같은 기법 안에서)
    lunch_min: int
    lunch_window: tuple[int, int]     # (start_min, end_min)
    assessor_mode: str                # 'universal' | 'specialized'
    assessor_count_by_technique: dict[str, int]
    rooms: list[Room]                 # 평가실
    prep_rooms: list[Room]            # 숙지실
    waiting_rooms: list[WaitingRoom]

    @property
    def has_group_discussion(self) -> bool:
        return any(t.candidates >= 3 for t in self.techniques)

    @property
    def has_multi_assessor_solo(self) -> bool:
        """1:2 역할수행처럼 (a>=2, c=1) 기법이 있는지."""
        return any(t.assessors >= 2 and t.candidates == 1 for t in self.techniques)

    @property
    def total_assessor_sessions(self) -> int:
        """모든 대상자가 모든 기법 받을 때 발생하는 총 위원 세션 수."""
        per_candidate = sum(t.load_per_candidate for t in self.techniques)
        # int 변환: 통상 정수
        return round(self.total_candidates * per_candidate)


# ───────────────────────── Config 로드 ─────────────────────────


def parse_time_str(s: str) -> int:
    """'HH:MM' → 분."""
    h, m = s.split(":")
    return int(h) * 60 + int(m)


def format_time_min(mins: int) -> str:
    return f"{mins // 60:02d}:{mins % 60:02d}"


def load_config(path: str | Path) -> Config:
    with open(path, "r", encoding="utf-8") as f:
        raw = json.load(f)

    techniques = []
    for t in raw["techniques"]:
        # 하위호환: 옛 포맷은 duration_min 하나만 있었음 → eval로 매핑하고 prep=0
        if "duration_min" in t and "eval_duration_min" not in t:
            prep_d = 0
            eval_d = int(t["duration_min"])
        else:
            prep_d = int(t.get("prep_duration_min", 0))
            eval_d = int(t.get("eval_duration_min", 0))
        techniques.append(
            Technique(
                name=t["name"],
                assessors=int(t["assessors"]),
                candidates=int(t["candidates"]),
                prep_duration_min=prep_d,
                eval_duration_min=eval_d,
            )
        )
    rooms = [
        Room(
            name=r["name"],
            supported=tuple(r["supported"]),
            capacity=int(r["capacity"]),
            zone=r["zone"],
        )
        for r in raw["rooms"]
    ]
    prep_rooms = [
        Room(
            name=r["name"],
            supported=tuple(r["supported"]),
            capacity=int(r["capacity"]),
            zone=r["zone"],
        )
        for r in raw.get("prep_rooms", [])
    ]
    waiting_rooms = [
        WaitingRoom(name=w["name"], zone=w["zone"])
        for w in raw.get("waiting_rooms", [])
    ]
    lunch_window_raw = raw.get("lunch_window", ["11:30", "13:30"])
    lunch_window = (
        parse_time_str(lunch_window_raw[0]),
        parse_time_str(lunch_window_raw[1]),
    )

    cfg = Config(
        title=raw.get("title", "역량평가"),
        total_candidates=int(raw["total_candidates"]),
        techniques=techniques,
        start_time_min=parse_time_str(raw.get("start_time", "09:00")),
        transition_min=int(raw.get("transition_min", 10)),
        prep_to_eval_transition_min=int(raw.get("prep_to_eval_transition_min", 5)),
        lunch_min=int(raw.get("lunch_min", 60)),
        lunch_window=lunch_window,
        assessor_mode=raw.get("assessor_mode", "universal"),
        assessor_count_by_technique=raw.get("assessor_count_by_technique", {}),
        rooms=rooms,
        prep_rooms=prep_rooms,
        waiting_rooms=waiting_rooms,
    )

    _validate_config(cfg)
    return cfg


def _validate_config(cfg: Config) -> None:
    errs = []
    if cfg.total_candidates <= 0:
        errs.append("total_candidates 는 양수여야 함")
    if not cfg.techniques:
        errs.append("techniques 가 비어있음")
    for t in cfg.techniques:
        if t.assessors <= 0 or t.candidates <= 0:
            errs.append(f"기법 '{t.name}' 의 위원/대상자 수가 비정상")
        if t.prep_duration_min < 0 or t.eval_duration_min < 0:
            errs.append(f"기법 '{t.name}' 의 소요시간이 음수")
        if t.total_duration_min <= 0:
            errs.append(f"기법 '{t.name}' 의 총 소요시간이 0 이하 (숙지+평가 ≥ 5분)")
    # 숙지 시간이 있는데 숙지실이 없으면 경고
    needs_prep = any(t.prep_duration_min > 0 for t in cfg.techniques)
    if needs_prep and not cfg.prep_rooms:
        errs.append("숙지 시간이 있는 기법이 있는데 prep_rooms 가 비어있음")
    if cfg.assessor_mode not in ("universal", "specialized"):
        errs.append(f"assessor_mode 는 'universal' 또는 'specialized'")
    if cfg.assessor_mode == "specialized":
        for t in cfg.techniques:
            if t.name not in cfg.assessor_count_by_technique:
                errs.append(f"specialized 모드: '{t.name}' 위원 수 누락")
    if errs:
        raise ValueError("Config 오류:\n  - " + "\n  - ".join(errs))


# ───────────────────── (s, G) 후보 탐색 ─────────────────────


@dataclass(frozen=True)
class SgCandidate:
    s: int                   # 그룹 크기
    G: int                   # 그룹 수
    num_assessors: int       # 위원 수 (universal: G, specialized: sum of per-tech)
    load_per_assessor: float # A / num_assessors
    empty_slots: int         # G*s - N (올림권 여유)


def _required_distinct_assessors_per_candidate(cfg: Config) -> int:
    """대상자 1명이 모든 기법에서 만나야 하는 서로 다른 위원 수.

    각 기법 k에서 대상자 1명은 a_k 명의 위원에게 동시에 평가받음 (모든 기법에서).
    재매칭 금지(대상자 단위) → 이들이 모두 서로 다른 위원이어야 함.
    """
    return sum(t.assessors for t in cfg.techniques)


def _compute_num_assessors(cfg: Config, s: int, G: int) -> int:
    """필요 위원 수 자동 계산.

    하한:
      - 대상자당 필요 서로 다른 위원 수 (재매칭 금지 제약)
      - 한 그룹이 어떤 기법을 동시 진행할 때 필요한 위원 수: (s/c_k) × a_k
    부하:
      - A / num ≈ 8.5 가 되도록
    → max(하한, 부하 기준)
    """
    if cfg.assessor_mode == "specialized":
        return sum(cfg.assessor_count_by_technique.values())

    # 대상자당 필요 위원 수
    per_cand_min = _required_distinct_assessors_per_candidate(cfg)
    # 한 그룹 동시 진행 시 위원 수요 (최대치)
    group_concurrent_min = max(
        (s // t.candidates) * t.assessors for t in cfg.techniques
    )
    hard_min = max(per_cand_min, group_concurrent_min)

    A = cfg.total_assessor_sessions
    target_for_load = round(A / 8.5)

    return max(hard_min, target_for_load)


def generate_sg_candidates(cfg: Config) -> list[SgCandidate]:
    """(s, G) 후보를 제약 안에서 생성."""
    N = cfg.total_candidates
    has_gd = cfg.has_group_discussion

    candidates: list[SgCandidate] = []
    s_values: list[int]
    # 그룹 토론 등의 기법이 있어도 마지막 그룹은 잔여 인원으로 남을 수 있음.
    # _try_schedule_session 에서 마지막 청크가 candidates 보다 작을 때도 처리함.
    s_values = list(range(1, N + 1))

    import math

    for s in s_values:
        G_ceil = math.ceil(N / s)
        # 올림권 여유 적용: 마지막 그룹이 s보다 작을 수 있음 (집단토론도 마찬가지).
        # 마지막 그룹 < tech.candidates 인 경우는 _try_schedule_session 에서 처리
        # (예: 1명만 남으면 집단토론 1세션을 1명 + 3위원으로 진행).
        G_values = [G_ceil]

        for G in G_values:
            if G < 1:
                continue
            num_assessors = _compute_num_assessors(cfg, s, G)

            # 대기실 수 ≥ G 이어야 함
            if len(cfg.waiting_rooms) < G:
                continue

            A = cfg.total_assessor_sessions
            load = A / num_assessors if num_assessors else float("inf")
            empty = G * s - N

            candidates.append(
                SgCandidate(
                    s=s, G=G, num_assessors=num_assessors,
                    load_per_assessor=load, empty_slots=empty,
                )
            )

    return candidates


def filter_by_load(candidates: list[SgCandidate]) -> list[SgCandidate]:
    """위원당 8~9세션 우선 필터. 없으면 가장 가까운 후보들 반환."""
    in_range = [c for c in candidates if 8.0 <= c.load_per_assessor <= 9.0]
    if in_range:
        return in_range
    # 가까운 순으로 상위 5개
    return sorted(candidates, key=lambda c: abs(c.load_per_assessor - 8.5))[:5]


# ───────────────────── 스케줄러 코어 ─────────────────────


@dataclass
class Group:
    idx: int
    name: str                       # "A조", "B조" ...
    candidate_ids: list[int]        # 전역 대상자 ID (0..N-1)
    waiting_room: str
    waiting_zone: str
    remaining_techniques: set[str] = field(default_factory=set)
    busy_until: int = 0             # 분 단위, 현재 세션 종료 + transition 후
    lunch_done: bool = False


@dataclass
class Session:
    technique: str
    phase: str                        # 'prep' (숙지) | 'eval' (평가)
    group_idx: int
    group_name: str
    candidate_ids: list[int]
    assessor_ids: list[int]           # 숙지 단계에는 빈 리스트
    room: str
    start_min: int
    end_min: int


@dataclass
class LunchBlock:
    group_idx: int
    group_name: str
    start_min: int
    end_min: int


@dataclass
class State:
    cfg: Config
    s: int
    G: int
    num_assessors: int
    groups: list[Group]
    sessions: list[Session] = field(default_factory=list)
    lunches: list[LunchBlock] = field(default_factory=list)
    # 위원 i가 평가한 대상자 id 집합
    assessor_history: list[set[int]] = field(default_factory=list)
    # 자원 점유 인터벌: (start, end, owner) — 시간 겹침 검사용
    assessor_busy: list[list[tuple[int, int]]] = field(default_factory=list)
    room_busy: dict[str, list[tuple[int, int]]] = field(default_factory=dict)
    prep_room_busy: dict[str, list[tuple[int, int]]] = field(default_factory=dict)
    # zone i 동안 group_idx 가 점유 — (start, end, gidx)
    zone_busy: dict[str, list[tuple[int, int, int]]] = field(default_factory=dict)
    # 점심은 동시 1팀만 — 단일 인터벌 리스트
    lunch_busy: list[tuple[int, int]] = field(default_factory=list)


def _overlaps(a_start: int, a_end: int, b_start: int, b_end: int) -> bool:
    return a_start < b_end and b_start < a_end


def init_state(cfg: Config, s: int, G: int, num_assessors: int) -> State:
    """초기 상태 생성: 그룹 구성, 대기실 배정, 자원 초기화."""
    N = cfg.total_candidates
    groups: list[Group] = []
    # 그룹 멤버 배정: 0..N-1 을 s씩 묶음, 마지막 그룹은 잔여(<s 가능)
    for gi in range(G):
        members = list(range(gi * s, min((gi + 1) * s, N)))
        wroom = cfg.waiting_rooms[gi]
        groups.append(
            Group(
                idx=gi,
                name=f"{chr(ord('A') + gi)}조",
                candidate_ids=members,
                waiting_room=wroom.name,
                waiting_zone=wroom.zone,
                remaining_techniques={t.name for t in cfg.techniques},
            )
        )
    # zone_busy 키는 평가실 + 숙지실의 모든 zone
    all_zones = {r.zone for r in cfg.rooms} | {r.zone for r in cfg.prep_rooms}
    state = State(
        cfg=cfg, s=s, G=G, num_assessors=num_assessors, groups=groups,
        assessor_history=[set() for _ in range(num_assessors)],
        assessor_busy=[[] for _ in range(num_assessors)],
        room_busy={r.name: [] for r in cfg.rooms},
        prep_room_busy={r.name: [] for r in cfg.prep_rooms},
        zone_busy={z: [] for z in all_zones},
    )
    return state


def _assessors_for_technique(cfg: Config, state: State, technique: str) -> list[int]:
    """이 기법을 담당 가능한 위원 ID 리스트."""
    if cfg.assessor_mode == "universal":
        return list(range(state.num_assessors))
    # specialized: 각 기법별 풀이 따로. ID 범위로 분할.
    pools: dict[str, list[int]] = {}
    cursor = 0
    for t in cfg.techniques:
        cnt = cfg.assessor_count_by_technique.get(t.name, 0)
        pools[t.name] = list(range(cursor, cursor + cnt))
        cursor += cnt
    return pools.get(technique, [])


def _find_free_rooms(
    cfg: Config, state: State, technique: str, count: int, start: int, end: int
) -> list[str] | None:
    """기법 가능 + 시간대 비어있는 평가실을 count개 찾음."""
    found: list[str] = []
    for room in cfg.rooms:
        if not room.supports(technique):
            continue
        # 시간 겹침 확인
        if any(_overlaps(start, end, bs, be) for bs, be in state.room_busy[room.name]):
            continue
        found.append(room.name)
        if len(found) == count:
            return found
    return None


def _find_free_prep_room(
    cfg: Config, state: State, technique: str, capacity_needed: int,
    start: int, end: int,
) -> Room | None:
    """그룹 전체(capacity_needed명)가 들어갈 숙지실 1개를 찾음."""
    for room in cfg.prep_rooms:
        if not room.supports(technique):
            continue
        if room.capacity < capacity_needed:
            continue
        if any(_overlaps(start, end, bs, be) for bs, be in state.prep_room_busy[room.name]):
            continue
        return room
    return None


def _free_assessors_in_window(
    cfg: Config, state: State, technique: str, start: int, end: int,
) -> list[int]:
    """기법 풀 중 시간대만 비어있는 위원 ID 리스트 (재매칭 검사 안 함)."""
    pool = _assessors_for_technique(cfg, state, technique)
    return [
        aid for aid in pool
        if not any(_overlaps(start, end, bs, be) for bs, be in state.assessor_busy[aid])
    ]


def _match_assessors_for_session(
    state: State, available: list[int], cand_chunk: list[int], a_count: int,
) -> list[int] | None:
    """한 세션(대상자 chunk)에 대해 a_count명의 위원 매칭."""
    valid = [
        aid for aid in available
        if not any(cid in state.assessor_history[aid] for cid in cand_chunk)
    ]
    if len(valid) < a_count:
        return None
    return valid[:a_count]


def _match_chunks_with_backtrack(
    state: State, chunks: list[list[int]], free_pool: list[int], a_per_session: int,
) -> list[list[int]] | None:
    """각 chunk(병렬 세션)에 a_per_session명 위원 배정.

    제약:
      - chunk 내 모든 대상자에 대해 클린 (history에 없음)
      - chunk 간 위원 겹침 없음
    백트래킹으로 정확히 풀이.
    """
    from itertools import combinations

    options_per_chunk: list[list[int]] = []
    for chunk in chunks:
        opts = [
            aid for aid in free_pool
            if not any(cid in state.assessor_history[aid] for cid in chunk)
        ]
        options_per_chunk.append(opts)

    # MRV: 옵션 적은 chunk 부터 (남은 변수 중 가장 제약 강한 것)
    order = sorted(range(len(chunks)), key=lambda i: len(options_per_chunk[i]))
    assignment: dict[int, list[int]] = {}
    used: set[int] = set()

    def backtrack(pos: int) -> bool:
        if pos == len(order):
            return True
        i = order[pos]
        opts = [a for a in options_per_chunk[i] if a not in used]
        if len(opts) < a_per_session:
            return False
        for combo in combinations(opts, a_per_session):
            for a in combo:
                used.add(a)
            assignment[i] = list(combo)
            if backtrack(pos + 1):
                return True
            for a in combo:
                used.discard(a)
            assignment.pop(i, None)
        return False

    if backtrack(0):
        return [assignment[i] for i in range(len(chunks))]
    return None


def _zone_conflict(
    state: State, zone: str, start: int, end: int, group_idx: int
) -> bool:
    """zone 에 [start, end) 시간대에 다른 그룹이 점유 중이면 True."""
    for bs, be, gidx in state.zone_busy.get(zone, []):
        if gidx == group_idx:
            continue
        if _overlaps(start, end, bs, be):
            return True
    return False


def _try_schedule_session(
    state: State, group: Group, technique_name: str, start_min: int
) -> bool:
    """그룹 g가 시각 start_min 에 technique 을 시작할 수 있으면 예약하고 True 반환.

    한 기법 = 2단계(숙지 → 평가)로 진행. 둘 다 예약 성공해야 함 (원자적).
    숙지 시간이 0이면 숙지 단계 생략, 평가 시간이 0이면 평가 단계 생략.
    """
    cfg = state.cfg
    tech = next(t for t in cfg.techniques if t.name == technique_name)

    s = state.s

    # === 시간 계획 ===
    prep_start = start_min
    prep_end = prep_start + tech.prep_duration_min
    if tech.prep_duration_min > 0 and tech.eval_duration_min > 0:
        eval_start = prep_end + cfg.prep_to_eval_transition_min
    else:
        eval_start = prep_end  # 둘 중 하나만 있으면 이동시간 없음
    eval_end = eval_start + tech.eval_duration_min

    # === 숙지 단계: 방·zone 확인 ===
    prep_room_obj: Room | None = None
    if tech.prep_duration_min > 0:
        members_count = len(group.candidate_ids)
        if members_count == 0:
            # 그룹원 0 — 의미 없음
            group.remaining_techniques.discard(technique_name)
            return True
        prep_room_obj = _find_free_prep_room(
            cfg, state, technique_name, members_count, prep_start, prep_end
        )
        if prep_room_obj is None:
            return False
        if _zone_conflict(state, prep_room_obj.zone, prep_start, prep_end, group.idx):
            return False

    # === 평가 단계: 방·zone·위원 확인 ===
    eval_rooms: list[str] = []
    cand_chunks: list[list[int]] = []
    chunk_assessors: list[list[int]] = []

    if tech.eval_duration_min > 0:
        members = group.candidate_ids[:]
        # 실제 그룹 인원 기준으로 c명씩 청크. 마지막 청크는 c보다 작을 수 있음
        # (예: 그룹 4명 + 집단토론 c=3 → [3명, 1명]. 1명 세션도 위원 3명 배정됨).
        for i in range(0, len(members), tech.candidates):
            chunk = members[i : i + tech.candidates]
            if chunk:
                cand_chunks.append(chunk)
        if not cand_chunks:
            group.remaining_techniques.discard(technique_name)
            return True

        rooms_needed = len(cand_chunks)
        rooms_found = _find_free_rooms(
            cfg, state, technique_name, rooms_needed, eval_start, eval_end
        )
        if rooms_found is None:
            return False
        eval_rooms = rooms_found

        # zone 확인 — 평가실 zone들이 다른 그룹과 충돌 없어야
        selected_rooms = [r for r in cfg.rooms if r.name in eval_rooms]
        for room in selected_rooms:
            if _zone_conflict(state, room.zone, eval_start, eval_end, group.idx):
                return False

        # 위원 매칭
        free_pool = _free_assessors_in_window(
            cfg, state, technique_name, eval_start, eval_end
        )
        matched = _match_chunks_with_backtrack(
            state, cand_chunks, free_pool, tech.assessors
        )
        if matched is None:
            return False
        chunk_assessors = matched

    # === 모든 검증 통과 — 세션 등록 (commit) ===
    if prep_room_obj is not None:
        sess = Session(
            technique=technique_name,
            phase="prep",
            group_idx=group.idx,
            group_name=group.name,
            candidate_ids=group.candidate_ids[:],
            assessor_ids=[],
            room=prep_room_obj.name,
            start_min=prep_start,
            end_min=prep_end,
        )
        state.sessions.append(sess)
        state.prep_room_busy[prep_room_obj.name].append((prep_start, prep_end))
        state.zone_busy.setdefault(prep_room_obj.zone, []).append(
            (prep_start, prep_end, group.idx)
        )

    if tech.eval_duration_min > 0:
        for i, room_name in enumerate(eval_rooms):
            sess_cands = cand_chunks[i]
            sess_assessors = chunk_assessors[i]
            sess = Session(
                technique=technique_name,
                phase="eval",
                group_idx=group.idx,
                group_name=group.name,
                candidate_ids=sess_cands,
                assessor_ids=sess_assessors,
                room=room_name,
                start_min=eval_start,
                end_min=eval_end,
            )
            state.sessions.append(sess)
            state.room_busy[room_name].append((eval_start, eval_end))
            room = next(r for r in cfg.rooms if r.name == room_name)
            state.zone_busy.setdefault(room.zone, []).append(
                (eval_start, eval_end, group.idx)
            )
            for aid in sess_assessors:
                state.assessor_busy[aid].append((eval_start, eval_end))
                for cid in sess_cands:
                    state.assessor_history[aid].add(cid)

    group.remaining_techniques.discard(technique_name)
    group.busy_until = eval_end + cfg.transition_min
    return True


def _try_schedule_lunch(state: State, group: Group, start_min: int) -> bool:
    """점심을 [start, start+lunch_min) 에 예약.

    각 그룹은 자기 대기실에서 점심을 먹는 모델 → 그룹 간 점심 동시 진행 OK.
    (그룹별 대기실은 따로 있다는 사용자 답변 반영)
    """
    cfg = state.cfg
    end_min = start_min + cfg.lunch_min
    lw_s, lw_e = cfg.lunch_window
    if start_min < lw_s or end_min > lw_e:
        return False
    state.lunches.append(
        LunchBlock(group_idx=group.idx, group_name=group.name,
                   start_min=start_min, end_min=end_min)
    )
    group.lunch_done = True
    group.busy_until = end_min + cfg.transition_min
    return True


def schedule_one(cfg: Config, sg: SgCandidate, debug: bool = False) -> State | None:
    """(s, G) 후보 1개에 대해 시간표 생성 시도. 실패 시 None."""
    state = init_state(cfg, sg.s, sg.G, sg.num_assessors)
    current = cfg.start_time_min
    safety_limit = current + 24 * 60  # 24시간 안에 못 끝내면 실패
    no_progress_count = 0
    last_event_time = current

    while True:
        # 종료 조건: 모든 그룹이 모든 기법 완료 + 점심 완료
        if all(not g.remaining_techniques and g.lunch_done for g in state.groups):
            return state
        if current > safety_limit:
            if debug:
                print(f"    [DBG] 안전 한도 초과 @ {format_time_min(current)}")
                for g in state.groups:
                    print(f"      {g.name}: 남은={g.remaining_techniques}, lunch={g.lunch_done}, busy_until={format_time_min(g.busy_until)}")
            return None

        # idle 그룹: busy_until <= current
        idle = [g for g in state.groups if g.busy_until <= current
                and (g.remaining_techniques or not g.lunch_done)]
        if not idle:
            # 모두 busy — 다음 종료 시각으로 점프
            next_t = min(
                (g.busy_until for g in state.groups
                 if g.busy_until > current
                 and (g.remaining_techniques or not g.lunch_done)),
                default=current + SLOT_MIN,
            )
            current = max(next_t, current + SLOT_MIN)
            continue

        # 점심 안 먹은 그룹 우선, 그 다음 남은 기법 많은 순
        idle.sort(key=lambda g: (g.lunch_done, -len(g.remaining_techniques), g.idx))
        progress_this_round = False

        for g in idle:
            # 점심 우선 검토 (lunch window 안이고 아직 안 먹었으면)
            lw_s, lw_e = cfg.lunch_window
            took_action = False
            if (not g.lunch_done and lw_s <= current
                and current + cfg.lunch_min <= lw_e):
                # 점심 가능하면 일단 점심
                if _try_schedule_lunch(state, g, current):
                    took_action = True
                    progress_this_round = True
            if not took_action:
                # 남은 기법 중 가능한 거 시도
                # 휴리스틱: a_k 큰 것부터 (집단토론 먼저 → 마지막에 위원 풀 막히는 것 방지)
                # 동률이면 긴 기법부터
                tech_by_name = {t.name: t for t in cfg.techniques}
                techs_sorted = sorted(
                    g.remaining_techniques,
                    key=lambda tn: (-tech_by_name[tn].assessors, -tech_by_name[tn].total_duration_min),
                )
                for tn in techs_sorted:
                    if _try_schedule_session(state, g, tn, current):
                        progress_this_round = True
                        break

        if progress_this_round:
            no_progress_count = 0
            last_event_time = current
        else:
            no_progress_count += 1
            # 모든 그룹이 busy면 시간 점프해도 됨
            if all(g.busy_until > current for g in state.groups
                   if g.remaining_techniques or not g.lunch_done):
                next_t = min(
                    g.busy_until for g in state.groups
                    if g.remaining_techniques or not g.lunch_done
                )
                current = next_t
                continue
            if no_progress_count > 240:  # 20시간(슬롯) 진전 없으면 포기
                if debug:
                    print(f"    [DBG] 240슬롯 무진전 @ {format_time_min(current)}")
                    for g in state.groups:
                        print(f"      {g.name}: 남은={g.remaining_techniques}, lunch={g.lunch_done}, busy_until={format_time_min(g.busy_until)}")
                return None

        current += SLOT_MIN


# ───────────────────── 점수 / 최적 선택 ─────────────────────


def schedule_score(state: State) -> tuple:
    """사전순 비교용 점수 (낮을수록 좋음).

    (zone 충돌 수, |A/G - 8.5|, 총 종료시각, 평균 대기시간)
    zone 충돌은 항상 0 (스케줄러가 보장) → 사실상 (load 편차, 종료시각, 평균대기) 비교.
    """
    cfg = state.cfg
    # 총 종료시각 = 마지막 세션/점심 종료
    end_t = max(
        max((s.end_min for s in state.sessions), default=0),
        max((l.end_min for l in state.lunches), default=0),
    )
    # 평균 대기시간 = 그룹별 (마지막 활동 - 첫 활동 - 활동 합)
    total_wait = 0.0
    for g in state.groups:
        g_sessions = [s for s in state.sessions if s.group_idx == g.idx]
        g_lunches = [l for l in state.lunches if l.group_idx == g.idx]
        blocks = sorted(
            [(s.start_min, s.end_min) for s in g_sessions]
            + [(l.start_min, l.end_min) for l in g_lunches]
        )
        if not blocks:
            continue
        active = sum(e - s for s, e in blocks)
        span = blocks[-1][1] - blocks[0][0]
        total_wait += max(0, span - active)
    avg_wait = total_wait / max(1, len(state.groups))

    load = len(state.sessions) * 0  # zone 충돌 가정 0
    # |A/G - 8.5|
    A = state.cfg.total_assessor_sessions
    load_dev = abs(A / state.num_assessors - 8.5)
    return (load, load_dev, end_t, avg_wait)


def best_schedule(cfg: Config, candidates: list[SgCandidate]) -> State | None:
    best: State | None = None
    best_score = None
    for sg in candidates:
        st = schedule_one(cfg, sg, debug=False)
        if st is None:
            print(f"  ✗ s={sg.s} G={sg.G}: 스케줄 실패")
            continue
        sc = schedule_score(st)
        end_t = max(
            max((s.end_min for s in st.sessions), default=0),
            max((l.end_min for l in st.lunches), default=0),
        )
        print(f"  ✓ s={sg.s} G={sg.G}: 종료 {format_time_min(end_t)}, 세션 {len(st.sessions)}개")
        if best is None or sc < best_score:
            best = st
            best_score = sc
    return best


# ───────────────────── xlsx 출력 ─────────────────────


def write_xlsx(state: State, cfg: Config, output_path: Path) -> None:
    """visualizer.html 호환 xlsx 생성.

    시트 1: 조별 시간표 (시간 | 평가위원 | A조 | B조 | ... | 담당 | 행동요령)
    시트 2: 위원별 시간표 (위원당 어디서 누구를 평가하는지)
    시트 3: 대상자별 일정 (개인별 동선)
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
    from openpyxl.utils import get_column_letter
    from collections import defaultdict

    wb = Workbook()
    ws = wb.active
    ws.title = "조별 시간표"

    # 시간 범위
    start_t = cfg.start_time_min
    end_t = max(
        max((s.end_min for s in state.sessions), default=start_t),
        max((l.end_min for l in state.lunches), default=start_t),
    )
    end_t = ((end_t + 29) // 30) * 30
    time_slots = list(range(start_t, end_t, SLOT_MIN))

    # 컬럼 구조
    HEADER_ROW = 2
    COL_TIME = 1
    COL_ASSESSOR = 2
    COL_GROUPS_START = 3
    num_groups = len(state.groups)
    COL_STAFF = COL_GROUPS_START + num_groups
    COL_NOTE = COL_STAFF + 1

    # 제목
    title_cell = ws.cell(row=1, column=COL_TIME, value=cfg.title)
    title_cell.font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=COL_TIME, end_row=1, end_column=COL_NOTE)
    title_cell.alignment = Alignment(horizontal="center")

    # 헤더
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    header_font = Font(bold=True)
    headers = {COL_TIME: "시간", COL_ASSESSOR: "평가위원"}
    for i, g in enumerate(state.groups):
        headers[COL_GROUPS_START + i] = g.name
    headers[COL_STAFF] = "담당"
    headers[COL_NOTE] = "행동요령"
    for c, label in headers.items():
        cell = ws.cell(row=HEADER_ROW, column=c, value=label)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # 시간 칸
    for r, t in enumerate(time_slots):
        cell = ws.cell(row=HEADER_ROW + 1 + r, column=COL_TIME, value=format_time_min(t))
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(size=10)

    # 세션 → (group_idx, technique, phase, start, end) 로 묶어 한 셀로 표현
    grouped: dict[tuple[int, str, str, int, int], list[Session]] = defaultdict(list)
    for sess in state.sessions:
        grouped[(sess.group_idx, sess.technique, sess.phase, sess.start_min, sess.end_min)].append(sess)

    # 활동별 배경색 (visualizer.html 의 ACTIVITY_COLORS 참고)
    activity_color = {
        "서류함": "C8E6C9",
        "구두발표": "CE93D8",
        "1대1역할수행": "90CAF9",
        "1:1역할수행": "90CAF9",
        "1대2역할수행": "64B5F6",
        "1:2역할수행": "64B5F6",
        "집단토론": "FFCC80",
        "역할수행": "90CAF9",
        "점심": "FFE0B2",
    }
    def pick_color(name: str) -> str:
        for k, v in activity_color.items():
            if k in name:
                return v
        return "F5F5F5"

    def slot_idx(t: int) -> int:
        return (t - start_t) // SLOT_MIN

    def fill_cell(col: int, start_min: int, end_min: int, text: str, color: str):
        first = HEADER_ROW + 1 + slot_idx(start_min)
        last = HEADER_ROW + 1 + slot_idx(end_min) - 1
        cell = ws.cell(row=first, column=col, value=text)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(size=9)
        if first < last:
            ws.merge_cells(start_row=first, start_column=col, end_row=last, end_column=col)

    for (gidx, tech_name, phase, smin, emin), sess_list in grouped.items():
        rooms_in = [s.room for s in sess_list]
        rooms_text = ", ".join(sorted(set(rooms_in)))
        duration = emin - smin
        col = COL_GROUPS_START + gidx
        label = "(숙지)" if phase == "prep" else "(평가)"
        text = f"{tech_name} {label}\n{rooms_text} ({duration})"
        # 숙지는 더 옅은 색, 평가는 진한 색
        color = pick_color(tech_name)
        if phase == "prep":
            color = "E8EAF6"   # 옅은 회색 톤
        fill_cell(col, smin, emin, text, color)

    for l in state.lunches:
        col = COL_GROUPS_START + l.group_idx
        duration = l.end_min - l.start_min
        fill_cell(col, l.start_min, l.end_min, f"점심 ({duration})", pick_color("점심"))

    # 컬럼 너비
    ws.column_dimensions[get_column_letter(COL_TIME)].width = 10
    ws.column_dimensions[get_column_letter(COL_ASSESSOR)].width = 14
    for i in range(num_groups):
        ws.column_dimensions[get_column_letter(COL_GROUPS_START + i)].width = 22
    ws.column_dimensions[get_column_letter(COL_STAFF)].width = 14
    ws.column_dimensions[get_column_letter(COL_NOTE)].width = 30
    # 행 높이
    for r in range(HEADER_ROW + 1, HEADER_ROW + 1 + len(time_slots)):
        ws.row_dimensions[r].height = 18

    # 시트 2: 위원별 시간표
    ws2 = wb.create_sheet("위원별 시간표")
    ws2.cell(row=1, column=1, value=f"{cfg.title} — 위원별 일정").font = Font(bold=True, size=14)
    ws2.cell(row=HEADER_ROW, column=1, value="시간").fill = header_fill
    ws2.cell(row=HEADER_ROW, column=1).font = header_font
    for aid in range(state.num_assessors):
        cell = ws2.cell(row=HEADER_ROW, column=2 + aid, value=f"위원 {aid + 1}")
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    for r, t in enumerate(time_slots):
        ws2.cell(row=HEADER_ROW + 1 + r, column=1, value=format_time_min(t)).alignment = Alignment(horizontal="center")

    # 위원별 세션 (숙지 단계는 위원 없으니 자동 제외)
    for sess in state.sessions:
        if sess.phase != "eval":
            continue
        for aid in sess.assessor_ids:
            col = 2 + aid
            first = HEADER_ROW + 1 + slot_idx(sess.start_min)
            last = HEADER_ROW + 1 + slot_idx(sess.end_min) - 1
            text = f"{sess.technique}\n{sess.group_name} @ {sess.room}"
            cell = ws2.cell(row=first, column=col, value=text)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = PatternFill("solid", fgColor=pick_color(sess.technique))
            cell.font = Font(size=9)
            if first < last:
                ws2.merge_cells(start_row=first, start_column=col, end_row=last, end_column=col)

    ws2.column_dimensions["A"].width = 10
    for aid in range(state.num_assessors):
        ws2.column_dimensions[get_column_letter(2 + aid)].width = 20

    # 시트 3: 대상자별 일정
    ws3 = wb.create_sheet("대상자별 일정")
    ws3.cell(row=1, column=1, value=f"{cfg.title} — 대상자별 일정").font = Font(bold=True, size=14)
    ws3.cell(row=HEADER_ROW, column=1, value="시간").fill = header_fill
    ws3.cell(row=HEADER_ROW, column=1).font = header_font
    N = cfg.total_candidates
    for cid in range(N):
        g = next(g for g in state.groups if cid in g.candidate_ids)
        cell = ws3.cell(row=HEADER_ROW, column=2 + cid, value=f"{g.name}-{cid + 1}")
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
    for r, t in enumerate(time_slots):
        ws3.cell(row=HEADER_ROW + 1 + r, column=1, value=format_time_min(t)).alignment = Alignment(horizontal="center")

    for sess in state.sessions:
        for cid in sess.candidate_ids:
            col = 2 + cid
            first = HEADER_ROW + 1 + slot_idx(sess.start_min)
            last = HEADER_ROW + 1 + slot_idx(sess.end_min) - 1
            label = "(숙지)" if sess.phase == "prep" else "(평가)"
            text = f"{sess.technique} {label}\n{sess.room}"
            color = pick_color(sess.technique)
            if sess.phase == "prep":
                color = "E8EAF6"
            cell = ws3.cell(row=first, column=col, value=text)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = PatternFill("solid", fgColor=color)
            cell.font = Font(size=9)
            if first < last:
                ws3.merge_cells(start_row=first, start_column=col, end_row=last, end_column=col)
    for l in state.lunches:
        g = state.groups[l.group_idx]
        for cid in g.candidate_ids:
            col = 2 + cid
            first = HEADER_ROW + 1 + slot_idx(l.start_min)
            last = HEADER_ROW + 1 + slot_idx(l.end_min) - 1
            cell = ws3.cell(row=first, column=col, value=f"점심 ({l.end_min - l.start_min})")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.fill = PatternFill("solid", fgColor=pick_color("점심"))
            cell.font = Font(size=9)
            if first < last:
                ws3.merge_cells(start_row=first, start_column=col, end_row=last, end_column=col)

    ws3.column_dimensions["A"].width = 10
    for cid in range(N):
        ws3.column_dimensions[get_column_letter(2 + cid)].width = 18

    wb.save(output_path)


# ───────────────────── 진단/요약 출력 ─────────────────────


def print_sg_diagnostic(cfg: Config, candidates: list[SgCandidate]) -> None:
    print(f"\n=== Config: {cfg.title} ===")
    print(f"  대상자 {cfg.total_candidates}명, 기법 {len(cfg.techniques)}개")
    print(f"  총 위원 세션 수 A = {cfg.total_assessor_sessions}")
    print(f"  집단토론 포함: {cfg.has_group_discussion}, 1:2 포함: {cfg.has_multi_assessor_solo}")
    print(f"  대기실 {len(cfg.waiting_rooms)}개, 평가실 {len(cfg.rooms)}개")
    print(f"\n=== (s, G) 후보 {len(candidates)}개 ===")
    for c in candidates:
        marker = " ★" if 8.0 <= c.load_per_assessor <= 9.0 else ""
        print(
            f"  s={c.s:2d} G={c.G:2d} "
            f"위원={c.num_assessors:2d} "
            f"위원당={c.load_per_assessor:5.2f}세션 "
            f"빈자리={c.empty_slots}"
            f"{marker}"
        )


# ───────────────────── main ─────────────────────


def main() -> int:
    ap = argparse.ArgumentParser(description="역량평가 시간표 자동 생성")
    ap.add_argument("config", type=Path, help="JSON config 파일")
    ap.add_argument("-o", "--output", type=Path, default=Path("schedule.xlsx"))
    ap.add_argument("--diagnostic", action="store_true", help="(s, G) 후보만 출력하고 종료")
    args = ap.parse_args()

    cfg = load_config(args.config)
    sg_all = generate_sg_candidates(cfg)
    sg_filtered = filter_by_load(sg_all)

    print_sg_diagnostic(cfg, sg_all)
    print(f"\n=== 1차 필터 통과 {len(sg_filtered)}개 ===")
    for c in sg_filtered:
        print(f"  s={c.s} G={c.G} 위원당={c.load_per_assessor:.2f}")

    if args.diagnostic:
        return 0

    if not sg_filtered:
        print("\n[ERR] 가능한 (s, G) 후보가 없습니다.")
        return 1

    print("\n=== 스케줄링 시도 ===")
    best = best_schedule(cfg, sg_filtered)
    if best is None:
        print("\n[ERR] 모든 후보에서 스케줄링 실패")
        return 1

    print(f"\n=== 최적 시간표: s={best.s}, G={best.G}, 위원 {best.num_assessors}명 ===")
    print(f"  총 세션 {len(best.sessions)}개, 점심 {len(best.lunches)}개")
    end_t = max(
        max((s.end_min for s in best.sessions), default=0),
        max((l.end_min for l in best.lunches), default=0),
    )
    print(f"  종료 시각: {format_time_min(end_t)}")

    # 그룹별 요약
    print("\n=== 그룹별 일정 ===")
    for g in best.groups:
        items: list[tuple[int, int, str]] = []
        for s in best.sessions:
            if s.group_idx == g.idx:
                ph = "숙지" if s.phase == "prep" else "평가"
                items.append((s.start_min, s.end_min, f"{s.technique}({ph})@{s.room}"))
        for l in best.lunches:
            if l.group_idx == g.idx:
                items.append((l.start_min, l.end_min, "점심"))
        items.sort()
        print(f"  {g.name} (대기실: {g.waiting_room}):")
        for s_t, e_t, label in items:
            print(f"    {format_time_min(s_t)}-{format_time_min(e_t)}  {label}")

    write_xlsx(best, cfg, args.output)
    print(f"\n[OK] xlsx 저장: {args.output}")
    print(f"     visualizer.html 에서 이 파일을 드롭하여 동선 시각화 확인 가능.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
