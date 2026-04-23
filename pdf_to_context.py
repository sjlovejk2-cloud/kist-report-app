"""
PDF → 부서정보.md 자동 작성
사용법: python pdf_to_context.py 파일.pdf [파일2.pdf ...]
여러 PDF를 한 번에 넣을 수 있습니다.
"""
import sys, os, re, subprocess
import pdfplumber

CONTEXT_FILE = os.path.join(os.path.dirname(__file__), "부서정보.md")
MAX_CHARS_PER_PDF = 6000  # PDF당 최대 추출 글자수


def extract_text(pdf_path: str) -> str:
    """PDF에서 텍스트 추출."""
    texts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texts.append(t)
    full = "\n".join(texts)
    return full[:MAX_CHARS_PER_PDF]


def ask_claude(prompt: str) -> str:
    env = os.environ.copy()
    env.pop("CLAUDECODE", None)
    result = subprocess.run(
        ["claude", "-p", prompt],
        capture_output=True, text=True, encoding="utf-8", errors="replace", env=env
    )
    if result.returncode != 0:
        raise RuntimeError(result.stderr)
    return result.stdout.strip()


def build_context(pdf_paths: list[str]) -> str:
    """PDF 텍스트를 Claude로 분석해 부서정보 마크다운 생성."""
    all_text = ""
    for path in pdf_paths:
        name = os.path.basename(path)
        print(f"  읽는 중: {name}")
        text = extract_text(path)
        all_text += f"\n\n[파일: {name}]\n{text}"

    prompt = f"""아래는 부서 관련 문서에서 추출한 텍스트입니다.
이 내용을 분석해서 부서정보 마크다운 파일을 작성해주세요.

정보가 없는 항목은 비워두고, 있는 정보만 최대한 구체적으로 채워주세요.
텍스트에서 확인되는 실제 이름, 수치, 시설명 등을 그대로 사용하세요.

--- 추출 텍스트 ---
{all_text}
---

아래 형식으로만 응답하세요 (설명 없이):

## 부서 기본 정보
- 부서명:
- 소속 기관:
- 위치:
- 구성원 수:

## 구성원
-

## 주요 업무
-
-
-

## 현재 추진 과제 / 현안
-
-

## 보유 시설 및 관할 범위
-

## 예산 / 인력 현황
-

## 기타 특이사항
-
"""
    return ask_claude(prompt)


def save_context(content: str):
    header = ("# 부서 컨텍스트 파일 (PDF에서 자동 생성)\n"
              "# 내용을 직접 수정하거나 추가할 수 있습니다.\n\n")
    with open(CONTEXT_FILE, "w", encoding="utf-8") as f:
        f.write(header + content)
    print(f"\n저장 완료: {CONTEXT_FILE}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법: python pdf_to_context.py 파일.pdf [파일2.pdf ...]")
        sys.exit(1)

    pdf_paths = sys.argv[1:]
    missing = [p for p in pdf_paths if not os.path.exists(p)]
    if missing:
        print(f"파일을 찾을 수 없습니다: {', '.join(missing)}")
        sys.exit(1)

    print(f"PDF {len(pdf_paths)}개 분석 중...")
    context = build_context(pdf_paths)

    print("\n--- 생성된 부서정보 ---")
    print(context)

    save = input("\n부서정보.md에 저장할까요? [Y/n]: ").strip().lower()
    if save in ("", "y", "yes"):
        save_context(context)
        print("다음 보고서부터 자동 반영됩니다.")
    else:
        print("저장 취소됨.")
