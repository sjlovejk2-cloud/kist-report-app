"""
KIST 원규·지침 HWP → 벡터DB 인덱싱 (1회 실행)
실행: python build_index.py
소요시간: 약 10~20분 (202개 파일)
"""
import os, re, time, sys

REGULATIONS_DIR = r"F:\스마트규정시스템"
DB_PATH = os.path.join(os.path.dirname(__file__), "kist_db")


# ── HWP 텍스트 추출 (한글 COM 자동화) ─────────────────────────

def extract_hwp_text(filepath: str) -> str:
    try:
        import win32com.client
    except ImportError:
        print("  [오류] pywin32 미설치 — pip install pywin32")
        return ""

    hwp = None
    try:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.Open(os.path.abspath(filepath), "HWP", "forceopen:true")
        text = hwp.GetTextFile("TEXT", "") or ""
        return text
    except Exception as e:
        print(f"  [오류] {e}")
        return ""
    finally:
        if hwp:
            try:
                hwp.Quit()
            except Exception:
                pass


# ── 청크 분할 (조 단위) ───────────────────────────────────────

ARTICLE_RE = re.compile(r'(제\s*\d+\s*조(?:의\s*\d+)?(?:\s*\([^)]{1,30}\))?)')

def split_chunks(text: str, source: str, max_chars: int = 800) -> list[dict]:
    parts = ARTICLE_RE.split(text)
    chunks = []
    current_article = ""
    buffer = ""

    def flush():
        body = (current_article + "\n" + buffer).strip()
        if len(body) > 40:
            chunks.append({
                "text": body[:max_chars],
                "source": source,
                "article": current_article or "전문",
            })

    for part in parts:
        if ARTICLE_RE.match(part):
            flush()
            current_article = part.strip()
            buffer = ""
        else:
            buffer += part

    flush()

    # 조 구분 없는 문서는 고정 크기로 분할
    if not chunks:
        text = text.strip()
        for i in range(0, len(text), max_chars):
            chunk = text[i:i + max_chars].strip()
            if len(chunk) > 40:
                chunks.append({
                    "text": chunk,
                    "source": source,
                    "article": f"단락{i // max_chars + 1}",
                })

    return chunks


# ── 인덱싱 메인 ───────────────────────────────────────────────

def main():
    try:
        import chromadb
        from chromadb.utils.embedding_functions import DefaultEmbeddingFunction
    except ImportError:
        print("[오류] chromadb 미설치 — pip install chromadb")
        sys.exit(1)

    # 파일 목록 수집
    all_files = []
    for category in ["원규정보", "지침정보"]:
        folder = os.path.join(REGULATIONS_DIR, category)
        if not os.path.isdir(folder):
            print(f"[경고] 폴더 없음: {folder}")
            continue
        for fname in sorted(os.listdir(folder)):
            if fname.lower().endswith(".hwp"):
                all_files.append((category, os.path.join(folder, fname), fname))

    print(f"총 {len(all_files)}개 HWP 파일 인덱싱 시작")
    print(f"DB 저장 경로: {DB_PATH}\n")

    # ChromaDB 초기화
    client = chromadb.PersistentClient(path=DB_PATH)
    try:
        client.delete_collection("kist_regulations")
        print("기존 DB 삭제 완료\n")
    except Exception:
        pass

    collection = client.create_collection(
        name="kist_regulations",
        embedding_function=DefaultEmbeddingFunction(),
        metadata={"hnsw:space": "cosine"},
    )

    total_chunks = 0
    errors = []

    for idx, (category, filepath, fname) in enumerate(all_files, 1):
        short_name = fname[:45] + ("..." if len(fname) > 45 else "")
        print(f"[{idx:>3}/{len(all_files)}] {short_name}", end=" ", flush=True)

        text = extract_hwp_text(filepath)
        if not text.strip():
            print("→ 텍스트 없음, 건너뜀")
            errors.append(fname)
            continue

        source_label = f"[{category}] {fname.replace('.hwp', '')}"
        chunks = split_chunks(text, source_label)

        if not chunks:
            print("→ 청크 없음, 건너뜀")
            continue

        ids = [f"doc{idx}_chunk{i}" for i in range(len(chunks))]
        texts = [c["text"] for c in chunks]
        metas = [{"source": c["source"], "article": c["article"], "category": category} for c in chunks]

        try:
            collection.add(documents=texts, metadatas=metas, ids=ids)
            total_chunks += len(chunks)
            print(f"→ {len(chunks)}개 청크")
        except Exception as e:
            print(f"→ DB 추가 오류: {e}")
            errors.append(fname)

        time.sleep(0.05)  # COM 안정화

    print(f"\n{'='*55}")
    print(f"완료!  총 {total_chunks}개 청크 인덱싱")
    print(f"오류 파일 {len(errors)}개: {errors[:5]}" if errors else "오류 없음")
    print(f"DB 위치: {DB_PATH}")
    print("이제 app.py 에서 KIST 원규·지침 검색을 사용할 수 있습니다.")


if __name__ == "__main__":
    main()
