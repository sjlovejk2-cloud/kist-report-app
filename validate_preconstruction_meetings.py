import json
import re
import sys
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

try:
    import yaml
except ModuleNotFoundError as exc:
    raise SystemExit(
        "PyYAML이 설치되어 있지 않습니다. PowerShell에서 'py -m pip install pyyaml' 또는 'python -m pip install pyyaml' 실행 후 다시 시도하세요."
    ) from exc

API_BASE = "https://api.gov-dooray.com"
DEFAULT_RULES_PATH = Path("preconstruction_validation_rules.yaml")
DEFAULT_RULES_CONFIG = {
    "version": 1,
    "project_scope": {
        "name": "실험실 공사 자동점검",
        "description": "실제 템플릿 기준 최소 규칙",
        "primary_project_id": "3713261754210573091",
    },
    "field_sources": {
        "title": {"source": "post.subject"},
        "stage": {"source": "milestone.name"},
        "status": {"source": "workflow.name"},
        "tags": {"source": "tags[].name"},
        "body_text": {"source": "body.content"},
        "attachments": {"source": "files[]"},
        "attachment_names": {"source": "files[].name"},
        "comment_count": {"source": "logs.count"},
    },
    "parsing": {
        "body_sections": [
            {"key": "construction_name", "label_patterns": ["공사명"], "required_in_mvp": True},
            {"key": "k_drive", "label_patterns": ["K-DRIVE"], "required_in_mvp": False},
            {"key": "approval_doc", "label_patterns": ["시행품의"], "required_in_mvp": False},
            {"key": "account", "label_patterns": ["사용계정"], "required_in_mvp": False},
            {"key": "work_building", "label_patterns": ["1. 건축", "1.  건축"], "required_in_mvp": False},
            {"key": "work_mechanical", "label_patterns": ["2. 설비"], "required_in_mvp": False},
            {"key": "work_electrical", "label_patterns": ["3. 전기"], "required_in_mvp": False},
            {"key": "contract_date", "label_patterns": ["계약일"], "required_in_mvp": False},
            {"key": "start_date", "label_patterns": ["착공일"], "required_in_mvp": False},
            {"key": "completion_date", "label_patterns": ["준공일"], "required_in_mvp": False},
            {"key": "notes", "label_patterns": ["참고사항"], "required_in_mvp": False},
        ]
    },
    "rule_defaults": {"severity_order": ["info", "low", "medium", "high", "critical"]},
    "rules": [
        {
            "id": "required_trade_tag",
            "enabled": True,
            "severity": "high",
            "title": "공종 태그 누락",
            "when": {
                "all": [
                    {
                        "field": "tags",
                        "op": "not_contains_any",
                        "value": ["공종: 건축", "공종: 기계", "공종: 기타", "공종: 석면", "공종: 소방", "공종: 전기", "공종: 조경", "공종: 토목", "공종: 통신"],
                    }
                ]
            },
            "action": {"result": "보완 필요", "message": "공종(필수) 태그 1개 이상이 필요합니다."},
        },
        {
            "id": "missing_construction_name",
            "enabled": True,
            "severity": "high",
            "title": "공사명 누락",
            "when": {"all": [{"field": "parsed.construction_name", "op": "is_empty"}]},
            "action": {"result": "보완 필요", "message": "본문에 공사명을 입력해야 합니다."},
        },
        {
            "id": "missing_construction_content_all_empty",
            "enabled": True,
            "severity": "high",
            "title": "공사내용 누락",
            "when": {
                "all": [
                    {"field": "parsed.work_building", "op": "is_empty"},
                    {"field": "parsed.work_mechanical", "op": "is_empty"},
                    {"field": "parsed.work_electrical", "op": "is_empty"},
                ]
            },
            "action": {"result": "보완 필요", "message": "공사내용(건축/설비/전기) 중 최소 1개는 입력해야 합니다."},
        },
        {
            "id": "no_attachment_at_all",
            "enabled": True,
            "severity": "medium",
            "title": "첨부파일 없음",
            "when": {"all": [{"field": "attachments", "op": "count_eq", "value": 0}]},
            "action": {"result": "보완 권고", "message": "관련 근거자료 또는 현장사진 첨부가 권장됩니다."},
        },
        {
            "id": "approval_stage_without_approval_doc",
            "enabled": True,
            "severity": "high",
            "title": "시행품의 단계 자료 누락",
            "when": {
                "all": [
                    {"field": "stage", "op": "eq", "value": "시행품의"},
                    {"field": "parsed.approval_doc", "op": "is_empty"},
                ]
            },
            "action": {"result": "보완 필요", "message": "시행품의 단계에서는 시행품의 항목 입력이 필요합니다."},
        },
        {
            "id": "contract_stage_without_contract_date",
            "enabled": True,
            "severity": "high",
            "title": "계약체결의뢰 단계 계약일 누락",
            "when": {
                "all": [
                    {"field": "stage", "op": "eq", "value": "계약체결의뢰"},
                    {"field": "parsed.contract_date", "op": "is_empty"},
                ]
            },
            "action": {"result": "보완 필요", "message": "계약체결의뢰 단계에서는 계약일 입력이 필요합니다."},
        },
        {
            "id": "start_stage_without_start_date",
            "enabled": True,
            "severity": "high",
            "title": "착공 단계 착공일 누락",
            "when": {
                "all": [
                    {"field": "stage", "op": "eq", "value": "착공"},
                    {"field": "parsed.start_date", "op": "is_empty"},
                ]
            },
            "action": {"result": "보완 필요", "message": "착공 단계에서는 착공일 입력이 필요합니다."},
        },
        {
            "id": "completion_stage_without_completion_date",
            "enabled": True,
            "severity": "high",
            "title": "준공 단계 준공일 누락",
            "when": {
                "all": [
                    {"field": "stage", "op": "eq", "value": "준공"},
                    {"field": "parsed.completion_date", "op": "is_empty"},
                ]
            },
            "action": {"result": "보완 필요", "message": "준공 단계에서는 준공일 입력이 필요합니다."},
        },
    ],
    "readiness_score": {
        "method": "weighted_checklist",
        "max_score": 100,
        "weights": {"trade_tag": 25, "construction_name": 25, "work_content": 30, "attachment_present": 20},
    },
}


def api_get_json(token: str, url: str):
    req = urllib.request.Request(url, headers={"Authorization": f"dooray-api {token}"})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.load(resp)


def fetch_post_detail(token: str, post_id: str):
    return api_get_json(token, f"{API_BASE}/project/v1/posts/{post_id}")["result"]


def fetch_post_files(token: str, project_id: str, post_id: str):
    payload = api_get_json(token, f"{API_BASE}/project/v1/projects/{project_id}/posts/{post_id}/files")
    return payload.get("result", [])


def fetch_post_logs(token: str, project_id: str, post_id: str):
    payload = api_get_json(token, f"{API_BASE}/project/v1/projects/{project_id}/posts/{post_id}/logs")
    return {
        "count": int(payload.get("totalCount", 0) or 0),
        "items": payload.get("result", []),
    }


def fetch_project_tags(token: str, project_id: str):
    payload = api_get_json(token, f"{API_BASE}/project/v1/projects/{project_id}/tags")
    items = payload.get("result", []) or []
    return {str(item.get("id")): item.get("name", "") for item in items if item.get("id")}


def normalize_user_entry(entry: dict):
    if not isinstance(entry, dict):
        return {"type": None, "id": None, "name": None, "email": None, "workflow_name": None}
    item_type = entry.get("type")
    workflow_name = (entry.get("workflow") or {}).get("name")
    if item_type == "member":
        member = entry.get("member") or {}
        return {
            "type": item_type,
            "id": member.get("organizationMemberId") or member.get("organizationmemberid"),
            "name": member.get("name"),
            "email": member.get("emailAddress"),
            "workflow_name": workflow_name,
        }
    if item_type == "emailUser":
        email_user = entry.get("emailUser") or entry.get("emailuser") or {}
        return {
            "type": item_type,
            "id": None,
            "name": email_user.get("name"),
            "email": email_user.get("emailAddress"),
            "workflow_name": workflow_name,
        }
    if item_type == "group":
        group = entry.get("group") or {}
        return {
            "type": item_type,
            "id": group.get("projectMemberGroupId"),
            "name": None,
            "email": None,
            "workflow_name": workflow_name,
        }
    return {"type": item_type, "id": None, "name": None, "email": None, "workflow_name": workflow_name}


def normalize_attachment(file_item: dict):
    creator = file_item.get("creator") or {}
    return {
        "id": file_item.get("id"),
        "name": file_item.get("name", ""),
        "size": file_item.get("size", 0),
        "mime_type": file_item.get("mimeType", ""),
        "creator_type": creator.get("type"),
    }


def split_label_line(line: str):
    if ":" in line:
        return line.split(":", 1)
    if "：" in line:
        return line.split("：", 1)
    return None, None


def parse_labeled_sections(body_text: str, section_rules: list[dict]):
    result = {rule["key"]: "" for rule in section_rules}
    if not body_text:
        return result

    normalized_lines = [raw.replace("\xa0", " ").rstrip() for raw in body_text.splitlines()]

    def normalize_candidate(text: str):
        text = text.strip()
        text = re.sub(r"^#+\s*", "", text)
        text = re.sub(r"^\d+\)\s*", "", text)
        return text.strip()

    def clean_block_text(lines: list[str]):
        cleaned = []
        for raw_line in lines:
            line = normalize_candidate(raw_line)
            if not line:
                continue
            line = re.sub(r"^[-*]\s*", "", line).strip()
            if line in {"-", "*"}:
                continue
            cleaned.append(line)
        return "\n".join(cleaned).strip()

    pattern_pairs = []
    for rule in section_rules:
        for pattern in rule.get("label_patterns", []):
            pattern_pairs.append((pattern, rule["key"]))

    def match_rule(candidate: str):
        for pattern, key in pattern_pairs:
            escaped = re.escape(pattern)
            if candidate == pattern or re.match(rf"^{escaped}\s*[:：]", candidate):
                return key, pattern
        return None, None

    current_key = None
    current_lines = []

    def flush_current():
        nonlocal current_key, current_lines
        if current_key and not result[current_key]:
            result[current_key] = clean_block_text(current_lines)
        current_key = None
        current_lines = []

    for raw in normalized_lines:
        raw_stripped = raw.strip()
        candidate = normalize_candidate(raw)
        if not candidate:
            continue

        matched_key, matched_pattern = match_rule(candidate)
        if matched_key:
            flush_current()
            current_key = matched_key
            remainder = candidate[len(matched_pattern):].lstrip()
            if remainder.startswith(":") or remainder.startswith("："):
                remainder = remainder[1:].strip()
            if remainder:
                current_lines.append(remainder)
            continue

        if re.match(r"^#+\s*\d+[.)]\s*", raw_stripped):
            flush_current()
            continue

        if raw_stripped.startswith("#"):
            flush_current()
            continue

        if re.match(r"^[^\n]+[:：]\s*$", candidate):
            flush_current()
            continue

        if current_key:
            current_lines.append(candidate)

    flush_current()
    return result


def normalize_post(list_row: dict, detail: dict, files: list[dict], logs: dict, parsing_rules: list[dict], project_tag_map: dict | None = None):
    tags_raw = detail.get("tags") or list_row.get("tags") or []
    users = detail.get("users") or list_row.get("users") or {}
    body_text = ((detail.get("body") or {}).get("content") or "").strip()
    attachments = [normalize_attachment(x) for x in files]
    project_tag_map = project_tag_map or {}
    tag_names = []
    tag_ids = []
    for tag in tags_raw:
        tag_id = tag.get("id")
        tag_name = tag.get("name") or project_tag_map.get(str(tag_id), "")
        if tag_name:
            tag_names.append(tag_name)
        if tag_id:
            tag_ids.append(tag_id)
    normalized = {
        "post_id": detail.get("id") or list_row.get("id"),
        "project_id": ((detail.get("project") or {}).get("id") or (list_row.get("project") or {}).get("id")),
        "project_code": ((detail.get("project") or {}).get("code") or (list_row.get("project") or {}).get("code")),
        "task_number": detail.get("taskNumber") or list_row.get("taskNumber"),
        "number": detail.get("number") or list_row.get("number"),
        "subject": detail.get("subject") or list_row.get("subject") or "",
        "status": ((detail.get("workflow") or {}).get("name") or (list_row.get("workflow") or {}).get("name") or ""),
        "status_class": detail.get("workflowClass") or list_row.get("workflowClass") or "",
        "status_id": ((detail.get("workflow") or {}).get("id") or (list_row.get("workflow") or {}).get("id")),
        "stage": ((detail.get("milestone") or {}).get("name") or (list_row.get("milestone") or {}).get("name") or ""),
        "stage_id": ((detail.get("milestone") or {}).get("id") or (list_row.get("milestone") or {}).get("id")),
        "priority": detail.get("priority") or list_row.get("priority") or "",
        "created_at": detail.get("createdAt") or list_row.get("createdAt") or "",
        "updated_at": detail.get("updatedAt") or list_row.get("updatedAt") or "",
        "due_date": detail.get("dueDate") or list_row.get("dueDate") or "",
        "due_date_flag": detail.get("dueDateFlag") if detail.get("dueDateFlag") is not None else list_row.get("dueDateFlag"),
        "closed": bool(detail.get("closed") if detail.get("closed") is not None else list_row.get("closed")),
        "parent_post_id": ((detail.get("parent") or {}).get("id") or (list_row.get("parent") or {}).get("id")),
        "tags": tag_names,
        "tag_ids": tag_ids,
        "assignees": [normalize_user_entry(x) for x in (users.get("to") or [])],
        "watchers": [normalize_user_entry(x) for x in (users.get("cc") or [])],
        "body_text": body_text,
        "attachments": attachments,
        "attachment_names": [x.get("name", "") for x in attachments if x.get("name")],
        "comment_count": int(logs.get("count", 0)),
        "parsed": parse_labeled_sections(body_text, parsing_rules),
        "findings": [],
        "readiness_score": 0,
    }
    return normalized


def get_field_value(data: dict, field_name: str):
    parts = field_name.split(".")
    current = data
    for part in parts:
        if isinstance(current, dict):
            current = current.get(part)
        else:
            return None
    return current


def op_is_empty(value, expected=None):
    return value is None or value == "" or value == []


def op_eq(value, expected):
    return value == expected


def op_not_contains(value, expected):
    return expected not in (value or [])


def op_not_contains_any(value, expected):
    values = value or []
    return all(item not in values for item in expected)


def op_not_prefix_exists(value, expected):
    values = value or []
    return not any(str(item).startswith(expected) for item in values)


def op_count_eq(value, expected):
    return len(value or []) == expected


def op_none_match_regex(value, expected):
    pattern = re.compile(expected)
    return not any(pattern.search(str(item)) for item in (value or []))


OPS = {
    "is_empty": op_is_empty,
    "eq": op_eq,
    "not_contains": op_not_contains,
    "not_contains_any": op_not_contains_any,
    "not_prefix_exists": op_not_prefix_exists,
    "count_eq": op_count_eq,
    "none_match_regex": op_none_match_regex,
}


def evaluate_condition(node, data):
    if "all" in node:
        return all(evaluate_condition(child, data) for child in node["all"])
    if "any" in node:
        return any(evaluate_condition(child, data) for child in node["any"])
    field_name = node["field"]
    op_name = node["op"]
    value = get_field_value(data, field_name)
    return OPS[op_name](value, node.get("value"))


def calculate_readiness_score(normalized_post: dict, readiness_cfg: dict):
    parsed = normalized_post.get("parsed", {})
    tags = normalized_post.get("tags", [])
    score = 0
    weights = readiness_cfg.get("weights", {})
    trade_values = {
        "공종: 건축",
        "공종: 기계",
        "공종: 기타",
        "공종: 석면",
        "공종: 소방",
        "공종: 전기",
        "공종: 조경",
        "공종: 토목",
        "공종: 통신",
    }
    if any(t in trade_values for t in tags):
        score += weights.get("trade_tag", 0)
    if parsed.get("construction_name"):
        score += weights.get("construction_name", 0)
    if parsed.get("work_building") or parsed.get("work_mechanical") or parsed.get("work_electrical"):
        score += weights.get("work_content", 0)
    if normalized_post.get("attachments"):
        score += weights.get("attachment_present", 0)
    if parsed.get("approval_doc"):
        score += weights.get("approval_doc", 0)
    if parsed.get("notes"):
        score += weights.get("reference_note", 0)
    return min(score, readiness_cfg.get("max_score", 100))


def validate_post(normalized_post: dict, rules_cfg: dict):
    findings = []
    for rule in rules_cfg.get("rules", []):
        if not rule.get("enabled", True):
            continue
        if evaluate_condition(rule.get("when", {}), normalized_post):
            findings.append(
                {
                    "id": rule["id"],
                    "severity": rule["severity"],
                    "title": rule["title"],
                    "result": rule["action"]["result"],
                    "message": rule["action"]["message"],
                }
            )
    normalized_post["findings"] = findings
    normalized_post["readiness_score"] = calculate_readiness_score(normalized_post, rules_cfg.get("readiness_score", {}))
    return normalized_post


def render_report(posts: list[dict]):
    lines = []
    lines.append("[Dooray 사전검토·예가산정 자동점검 리포트]")
    lines.append(f"생성시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"ascii_total_items={len(posts)}")
    critical = sum(1 for p in posts if any(f["severity"] == "critical" for f in p["findings"]))
    high = sum(1 for p in posts if any(f["severity"] == "high" for f in p["findings"]))
    lines.append(f"ascii_critical_items={critical}")
    lines.append(f"ascii_high_items={high}")
    lines.append("")
    lines.append("[요약]")
    lines.append(f"- 점검 업무 수: {len(posts)}건")
    lines.append(f"- Critical 포함 업무: {critical}건")
    lines.append(f"- High 포함 업무: {high}건")
    lines.append("")
    lines.append("[업무별 결과]")
    if not posts:
        lines.append("- 없음")
        return "\n".join(lines)
    for post in posts:
        findings = post.get("findings", [])
        if findings:
            finding_text = "; ".join(f"[{f['severity']}] {f['title']}" for f in findings)
        else:
            finding_text = "이상 없음"
        lines.append(
            f"- {post.get('task_number') or post.get('post_id')} | 상태={post.get('status')} | 단계={post.get('stage')} | 점수={post.get('readiness_score')} | {post.get('subject')} | {finding_text}"
        )
    return "\n".join(lines)


def load_rules(rules_path: Path):
    try:
        if rules_path.exists():
            return yaml.safe_load(rules_path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"warning: rules file load failed ({exc}). built-in default rules will be used.")
    return DEFAULT_RULES_CONFIG


def main():
    if len(sys.argv) < 3:
        print("Usage: python validate_preconstruction_meetings.py <DOORAY_TOKEN> <list_json_path> [rules_yaml_path] [output_txt_path]")
        sys.exit(1)

    token = sys.argv[1]
    list_path = Path(sys.argv[2])
    rules_path = Path(sys.argv[3]) if len(sys.argv) >= 4 else DEFAULT_RULES_PATH
    output_path = Path(sys.argv[4]) if len(sys.argv) >= 5 else Path("preconstruction_validation_report.txt")

    rules_cfg = load_rules(rules_path)
    parsing_rules = rules_cfg.get("parsing", {}).get("body_sections", [])

    list_payload = json.loads(list_path.read_text(encoding="utf-8"))
    rows = list_payload.get("result", [])

    validated_posts = []
    project_tag_maps = {}
    for row in rows:
        project_id = ((row.get("project") or {}).get("id") or rules_cfg.get("project_scope", {}).get("primary_project_id"))
        post_id = row["id"]
        if project_id not in project_tag_maps:
            project_tag_maps[project_id] = fetch_project_tags(token, project_id)
        detail = fetch_post_detail(token, post_id)
        files = fetch_post_files(token, project_id, post_id)
        logs = fetch_post_logs(token, project_id, post_id)
        normalized = normalize_post(row, detail, files, logs, parsing_rules, project_tag_maps.get(project_id))
        validated = validate_post(normalized, rules_cfg)
        validated_posts.append(validated)

    report_text = render_report(validated_posts)
    output_path.write_text(report_text, encoding="utf-8-sig")
    json_output_path = output_path.with_suffix(".json")
    json_output_path.write_text(json.dumps(validated_posts, ensure_ascii=False, indent=2), encoding="utf-8-sig")

    print(f"saved_report={output_path}")
    print(f"saved_json={json_output_path}")
    print(f"validated_count={len(validated_posts)}")


if __name__ == "__main__":
    main()
