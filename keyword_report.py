"""
키워드 → AI 보고서 초안 → HWPX 출력
사용법: python keyword_report.py 키워드1 키워드2 ...
출력: 보고서_YYYYMMDD_HHMMSS.hwpx (데스크탑)

구조:
  □ 주제 (3~4개)
    ○ 핵심 내용 (최대 3개)
       - 세부 사항 (최대 3개)
    [표] 필요 시 삽입
"""
import sys, json, zipfile, html, re, os, subprocess, shutil, glob
from datetime import datetime

TEMPLATE = os.path.join(os.path.dirname(__file__), "보고서 초안 양식.hwpx")
CONTEXT_FILE = os.path.join(os.path.dirname(__file__), "부서정보.md")
MAX_CHARS = 42
_tbl_id = [2000000000]


def load_context() -> str:
    """부서정보.md 로드. 주석·빈 항목 제거 후 반환."""
    if not os.path.exists(CONTEXT_FILE):
        return ""
    lines = []
    with open(CONTEXT_FILE, "r", encoding="utf-8") as f:
        for line in f:
            stripped = line.rstrip()
            if stripped.startswith("#"):
                continue
            if stripped in ("- ", "-", ""):
                continue
            lines.append(stripped)
    content = "\n".join(lines).strip()
    return content


# ── 유틸 ──────────────────────────────────────────────────

def esc(t):
    return html.escape(str(t), quote=True)

def cut(text, limit=MAX_CHARS):
    return text[:limit] if len(text) > limit else text


# ── HWPX XML 빌더 ─────────────────────────────────────────

LINESEG_BODY = ('  <hp:linesegarray>'
                '<hp:lineseg textpos="0" vertpos="0" vertsize="1400" textheight="1400"'
                ' baseline="1190" spacing="772" horzpos="0" horzsize="51024" flags="393216"/>'
                '</hp:linesegarray>')

LINESEG_EMPTY = ('  <hp:linesegarray>'
                 '<hp:lineseg textpos="0" vertpos="0" vertsize="1400" textheight="1400"'
                 ' baseline="1190" spacing="700" horzpos="0" horzsize="51024" flags="393216"/>'
                 '</hp:linesegarray>')


def p(paraPr, charPr, text=""):
    """단락 XML 한 줄 생성."""
    inner = f'<hp:t>{esc(text)}</hp:t>' if text else ''
    return (f'<hp:p id="0" paraPrIDRef="{paraPr}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="{charPr}">{inner}</hp:run>'
            f'{LINESEG_BODY}'
            f'</hp:p>')


def empty_para():
    return (f'<hp:p id="0" paraPrIDRef="34" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="30"/>'
            f'{LINESEG_EMPTY}'
            f'</hp:p>')


def make_table(headers: list[str], rows: list[list[str]], indent: int = 2) -> str:
    """본문 삽입용 HWPX 테이블 XML 생성 (새 양식 스타일).
    indent: 왼쪽 들여쓰기 칸 수 (□하위=2, ○하위=3, -하위=4)
    """
    _tbl_id[0] += 1
    tid = _tbl_id[0]
    col_count = len(headers)
    h_offset = indent * 700          # 1칸 ≈ 700 HWP unit (14pt 반각)
    total_w = 50460 - h_offset
    col_w = total_w // col_count
    col_ws = [col_w] * col_count
    col_ws[-1] = total_w - col_w * (col_count - 1)
    row_h = 1765
    all_rows = len(rows) + 1  # header + data rows

    # borderFillIDRef 매핑: row_idx × col_position(first/mid/last)
    # row 0 (header): first=13, mid=15, last=17
    # row 1:          first=12, mid=14, last=16
    # row 2+ (data):  first=11, mid=9,  last=10
    def get_bfill(row_idx, col_idx):
        last = col_idx == col_count - 1
        first = col_idx == 0
        if row_idx == 0:
            return "13" if first else ("17" if last else "15")
        elif row_idx == 1:
            return "12" if first else ("16" if last else "14")
        else:
            return "11" if first else ("10" if last else "9")

    def cell(text, col, row_idx):
        bfill = get_bfill(row_idx, col)
        w = col_ws[col]
        is_header = row_idx == 0
        return (
            f'<hp:tc name="" header="{1 if is_header else 0}" hasMargin="0" protect="0"'
            f' editable="0" dirty="0" borderFillIDRef="{bfill}">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER"'
            f' linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0"'
            f' hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="2147483648" paraPrIDRef="2" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="5"><hp:t>{esc(cut(text, 20))}</hp:t></hp:run>'
            f'<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1200" textheight="1200"'
            f' baseline="1020" spacing="360" horzpos="0" horzsize="{w - 1020}" flags="393216"/></hp:linesegarray>'
            f'</hp:p>'
            f'</hp:subList>'
            f'<hp:cellAddr colAddr="{col}" rowAddr="{row_idx}"/>'
            f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="{w}" height="{row_h}"/>'
            f'<hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
            f'</hp:tc>'
        )

    rows_xml = []
    rows_xml.append('<hp:tr>' + ''.join(cell(h, c, 0) for c, h in enumerate(headers)) + '</hp:tr>')
    for r, row in enumerate(rows, 1):
        cells = [cell(row[c] if c < len(row) else '', c, r) for c in range(col_count)]
        rows_xml.append('<hp:tr>' + ''.join(cells) + '</hp:tr>')

    tbl_h = row_h * all_rows
    tbl_xml = (
        f'<hp:tbl id="{tid}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM"'
        f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL"'
        f' repeatHeader="1" rowCnt="{all_rows}" colCnt="{col_count}" cellSpacing="0"'
        f' borderFillIDRef="9" noAdjust="0">'
        f'<hp:sz width="{total_w}" widthRelTo="ABSOLUTE" height="{tbl_h}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0"'
        f' holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" horzAlign="LEFT"'
        f' vertOffset="0" horzOffset="{h_offset}"/>'
        f'<hp:outMargin left="283" right="283" top="283" bottom="283"/>'
        f'<hp:inMargin left="510" right="510" top="141" bottom="141"/>'
        + ''.join(rows_xml) +
        f'</hp:tbl>'
    )

    return (
        f'<hp:p id="0" paraPrIDRef="33" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="30">'
        f'{tbl_xml}'
        f'<hp:t/>'
        f'</hp:run>'
        f'<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1400" textheight="1400"'
        f' baseline="1190" spacing="700" horzpos="0" horzsize="51024" flags="393216"/></hp:linesegarray>'
        f'</hp:p>'
    )


def make_section(source: str, topic: str, points: list[dict], table: dict | None = None) -> str:
    """□ 블록 전체 XML 생성 (○ 최대 3개, - 최대 3개, 선택적 표)."""
    parts = []
    # □ 헤더 줄
    parts.append(
        f'<hp:p id="0" paraPrIDRef="33" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="33"><hp:t> □ {esc(source)}, {esc(cut(topic, 30))}</hp:t></hp:run>'
        f'<hp:run charPrIDRef="34"/>'
        f'{LINESEG_BODY}'
        f'</hp:p>'
    )
    # ○ 및 - 줄 (○=13pt:charPr22, -=12pt:charPr5)
    for pt in points[:3]:
        parts.append(p("35", "22", f"  ○ {cut(pt.get('summary', ''))}"))
        for d in pt.get("details", [])[:3]:
            parts.append(p("36", "5", f"   - {cut(d)}"))
        # ○ 아래 표 (들여쓰기 3칸)
        pt_table = pt.get("table")
        if pt_table and pt_table.get("headers") and pt_table.get("rows"):
            parts.append(make_table(pt_table["headers"], pt_table["rows"], indent=3))
    # □ 아래 표 (들여쓰기 2칸)
    if table and table.get("headers") and table.get("rows"):
        parts.append(make_table(table["headers"], table["rows"], indent=2))
    return "\n".join(parts)


# ── Claude CLI 호출 ────────────────────────────────────────

PROMPT_TMPL = """\
키워드: {keywords}
{context_block}
위 키워드를 주제로 한 보고서 초안을 JSON으로 작성해주세요.
부서 컨텍스트가 있다면 실제 정보(구성원명, 시설명, 업무 등)를 반영해주세요.

구조 규칙:
- sections: 3~4개 (□ 주제)
- 각 section.points: 1~3개 (○ 항목)
- 각 point.details: 1~3개 (- 세부)
- table은 section 레벨(□ 아래) 또는 point 레벨(○ 아래) 중 적절한 위치에 배치 (불필요하면 null)
- 문체: 명사형 종결. ~실시, ~예정, ~완료, ~관리, ~추진 등으로 끝낼 것
- ~임, ~함, ~됨, ~있음 등 서술형 어미 사용 절대 금지
- 예시: "기자재 정리 예정", "현장 점검 실시", "예산 편성 완료"
- title: 40자 이내
- source: 10자 이내
- topic: 30자 이내
- summary/detail/셀: 42자 이내 (반드시 한 줄)
- 표 헤더·셀: 20자 이내

JSON만 응답 (설명 없이):
{{
  "title": "보고서 제목",
  "sections": [
    {{
      "source": "기관명",
      "topic": "주제",
      "points": [
        {{
          "summary": "핵심 내용",
          "details": ["세부1", "세부2"],
          "table": {{"headers": ["구분", "내용"], "rows": [["항목1", "값1"]]}}
        }}
      ],
      "table": {{
        "headers": ["구분", "내용"],
        "rows": [["항목1", "값1"], ["항목2", "값2"]]
      }}
    }},
    {{
      "source": "기관명",
      "topic": "주제",
      "points": [{{"summary": "...", "details": ["..."]}}],
      "table": null
    }}
  ]
}}"""


def _normalize_phrase(text: str, default_suffix: str = "정리") -> str:
    text = re.sub(r"\s+", " ", (text or "").strip())
    if not text:
        return default_suffix
    if any(text.endswith(s) for s in ["실시", "예정", "완료", "관리", "추진", "검토", "작성", "공유"]):
        return cut(text)
    return cut(f"{text} {default_suffix}")


def _fallback_report(keywords: list[str]) -> dict:
    raw = [k.strip() for k in keywords if str(k).strip()]
    title = cut(raw[0] if raw else "보고서 초안")
    body = raw[1] if len(raw) > 1 else ""
    parts = [p.strip(" -•\n\t") for p in re.split(r"[\n,]+", body) if p.strip(" -•\n\t")]
    if not parts:
        parts = [title, "회의 목적", "참여자 및 역할"]

    detail_pool = parts[1:] or ["세부사항 정리", "일정 검토"]
    points = []
    for idx, part in enumerate(parts[:3]):
        details = []
        if idx == 0 and body:
            details.append(_normalize_phrase(body, "정리"))
        for extra in detail_pool[idx:idx+2]:
            norm = _normalize_phrase(extra, "정리")
            if norm not in details:
                details.append(norm)
        points.append({
            "summary": _normalize_phrase(part, "정리"),
            "details": details[:3],
            "table": None,
        })

    return {
        "title": title,
        "sections": [
            {
                "source": "시설운영팀",
                "topic": cut(title),
                "points": points[:3],
                "table": None,
            },
            {
                "source": "업무개요",
                "topic": "추진 방향",
                "points": [
                    {
                        "summary": "회의 목적 정리",
                        "details": [
                            _normalize_phrase(parts[0] if parts else title, "정리"),
                            "검토사항 공유",
                        ],
                        "table": None,
                    },
                    {
                        "summary": "참여자 역할 정리",
                        "details": [
                            _normalize_phrase(parts[1] if len(parts) > 1 else "참여자 역할", "정리"),
                            "후속 일정 검토",
                        ],
                        "table": None,
                    },
                ],
                "table": None,
            },
        ],
    }


def _resolve_claude_command() -> list[str] | None:
    _which = shutil.which("claude")
    if _which:
        return [_which]

    _candidates = [
        "/mnt/c/Users/진광진광/AppData/Local/AnthropicClaude/claude.exe",
        os.path.join(os.path.expanduser("~"), "AppData", "Local", "AnthropicClaude", "claude.exe"),
    ]
    _candidates.extend(glob.glob("/mnt/c/Users/*/AppData/Local/AnthropicClaude/claude.exe"))

    for _cmd in _candidates:
        if _cmd and os.path.exists(_cmd):
            return [_cmd]
    return None


def generate_report(keywords: list[str]) -> dict:
    context = load_context()
    context_block = f"\n[부서 컨텍스트]\n{context}\n" if context else ""
    prompt = PROMPT_TMPL.format(keywords=", ".join(keywords), context_block=context_block)
    env = os.environ.copy()
    env.pop("CLAUDECODE", None)

    _claude_cmd = _resolve_claude_command()
    if not _claude_cmd:
        return _fallback_report(keywords)

    result = subprocess.run(
        _claude_cmd + ["-p", prompt],
        capture_output=True, text=True, encoding="utf-8", errors="replace", env=env
    )
    if result.returncode != 0:
        raise RuntimeError(f"claude CLI 오류:\n{result.stderr}")
    text = result.stdout.strip()
    m = re.search(r'\{.*\}', text, re.DOTALL)
    return json.loads(m.group() if m else text)


# ── HWPX 빌드 ─────────────────────────────────────────────

def build_hwpx(report: dict, output_path: str):
    today = datetime.now().strftime('%y.%m.%d')

    with zipfile.ZipFile(TEMPLATE, 'r') as z:
        section_xml = z.read('Contents/section0.xml').decode('utf-8')

    section_xml = section_xml.replace('타이틀 입력', esc(report['title']))
    section_xml = section_xml.replace('{오늘날짜YY.MM.DD}', today)

    body_tag = '<hp:p id="0" paraPrIDRef="33"'
    body_start = section_xml.find(body_tag)
    header_part = section_xml[:body_start]

    parts = [empty_para()]
    for sec in report['sections']:
        parts.append(make_section(
            sec['source'],
            sec['topic'],
            sec.get('points', []),
            sec.get('table')
        ))
        parts.append(empty_para())

    new_section = header_part + "\n".join(parts) + "\n</hs:sec>\n"

    tmp = output_path + '.tmp'
    with zipfile.ZipFile(TEMPLATE, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                raw = zin.read(info.filename)
                if info.filename == 'Contents/section0.xml':
                    raw = new_section.encode('utf-8')
                compress = zipfile.ZIP_STORED if info.filename == 'mimetype' else zipfile.ZIP_DEFLATED
                zout.writestr(info, raw, compress_type=compress)
    os.replace(tmp, output_path)


# ── 설계서 AI 검토 ────────────────────────────────────────

DESIGN_REVIEW_PROMPT = """\
다음 KIST 시설공사 설계서 검토 기준을 참고하여 아래 문서를 검토해주세요.

[검토 기준 및 규정]
{criteria}

[검토 대상]
- 파일명: {filename}
- 추정 문서 유형: {doc_type}
- 내용 (최대 8000자):
{text}

아래 JSON 형식으로만 응답하세요 (설명 없이):
{{
  "doc_type_confirmed": "내역서/도면/시방서/견적서/복합/기타 중 하나",
  "completeness_score": 0~100 사이 정수,
  "grade": "양호/보통/미흡 중 하나",
  "summary": "30자 이내 종합의견",
  "checks": [
    {{"항목": "항목명", "결과": "충족", "의견": ""}},
    {{"항목": "항목명", "결과": "미흡", "의견": "구체적인 보완 내용"}}
  ],
  "risks": [
    {{"항목": "위험 항목명", "근거": "발견된 근거", "의견": "조치 필요 내용"}}
  ],
  "missing_companion_docs": ["도면", "시방서"]
}}"""


def review_design_doc(filename: str, doc_type: str, text: str, criteria_texts: list) -> dict:
    # 기준 MD 문서
    criteria = "\n\n---\n\n".join(criteria_texts)[:4000] if criteria_texts else ""

    # KIST 원규·지침 RAG 검색 (DB 구축된 경우에만)
    rag_context = ""
    try:
        from kist_rag import is_db_ready, search_for_review, format_for_prompt
        if is_db_ready():
            rag_items = search_for_review(doc_type, text, n_results=8)
            rag_context = format_for_prompt(rag_items, max_chars=4000)
    except Exception:
        pass

    # 기준 조합
    if rag_context:
        full_criteria = criteria + "\n\n[KIST 원규·지침 관련 조항 (자동 검색)]\n" + rag_context
    else:
        full_criteria = criteria or "기준 문서 없음"

    prompt = DESIGN_REVIEW_PROMPT.format(
        criteria=full_criteria[:8000],
        filename=filename,
        doc_type=doc_type,
        text=text[:5000],
    )
    env = os.environ.copy()
    env.pop("CLAUDECODE", None)
    result = subprocess.run(
        ["claude", "-p", prompt],
        capture_output=True, text=True, encoding="utf-8", errors="replace", env=env,
        timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(f"AI 검토 오류:\n{result.stderr}")
    raw = result.stdout.strip()
    m = re.search(r'\{.*\}', raw, re.DOTALL)
    return json.loads(m.group() if m else raw)


# ── 진입점 ────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("사용법: python keyword_report.py 키워드1 키워드2 ...")
        sys.exit(1)

    keywords = sys.argv[1:]
    print(f"키워드: {', '.join(keywords)}")
    print("AI 보고서 생성 중...", flush=True)

    report = generate_report(keywords)
    print(f"제목: {report['title']}")
    print(f"섹션: {len(report['sections'])}개")

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    output = os.path.join(os.path.dirname(__file__), f"보고서_{ts}.hwpx")
    build_hwpx(report, output)
    print(f"완료: {output}")
