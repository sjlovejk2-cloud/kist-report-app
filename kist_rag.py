"""
KIST 원규·지침 벡터DB 검색 모듈
build_index.py 실행 후 사용 가능
"""
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "kist_db")
_collection = None


def _get_collection():
    global _collection
    if _collection is None:
        import chromadb
        from chromadb.utils.embedding_functions import DefaultEmbeddingFunction
        client = chromadb.PersistentClient(path=DB_PATH)
        _collection = client.get_collection(
            name="kist_regulations",
            embedding_function=DefaultEmbeddingFunction(),
        )
    return _collection


def is_db_ready() -> bool:
    """벡터DB가 구축되어 있는지 확인"""
    try:
        return _get_collection().count() > 0
    except Exception:
        return False


def db_stats() -> dict:
    """DB 통계 반환"""
    try:
        col = _get_collection()
        count = col.count()
        return {"total_chunks": count, "ready": count > 0}
    except Exception:
        return {"total_chunks": 0, "ready": False}


def search(query: str, n_results: int = 6, category: str = None) -> list[dict]:
    """순수 벡터 검색 (내부용)"""
    col = _get_collection()
    where = {"category": category} if category else None
    kwargs = dict(query_texts=[query], n_results=n_results)
    if where:
        kwargs["where"] = where
    results = col.query(**kwargs)
    items = []
    for i, doc in enumerate(results["documents"][0]):
        meta = results["metadatas"][0][i]
        items.append({
            "text": doc,
            "source": meta.get("source", ""),
            "article": meta.get("article", ""),
            "category": meta.get("category", ""),
        })
    return items


def search_hybrid(query: str, n_results: int = 6, category: str = None) -> list[dict]:
    """
    하이브리드 검색: 키워드 직접 포함 필터 우선 → 부족하면 벡터 검색 보충
    한국어 법령 텍스트에서 정확한 조문을 찾는 데 최적화
    """
    col = _get_collection()
    where = {"category": category} if category else None

    # 검색어에서 2자 이상 키워드 추출
    # + 붙여쓰기 대응: 4자 이상 단어는 3-gram으로 분해해 띄어쓰기 차이를 극복
    _raw_kws = [w for w in query.split() if len(w) >= 2]
    keywords: list[str] = []
    for _kw in _raw_kws:
        keywords.append(_kw)
        if len(_kw) >= 4:
            # 3글자 슬라이딩 윈도우: "연봉제운영지침" → "연봉제","봉제운","제운영","운영지","영지침"
            for _i in range(len(_kw) - 2):
                _ng = _kw[_i:_i + 3]
                if _ng not in keywords:
                    keywords.append(_ng)

    seen: set = set()
    results: list[dict] = []

    def _add(docs, metas):
        for doc, meta in zip(docs, metas):
            key = meta.get("source", "") + meta.get("article", "")
            if key not in seen:
                seen.add(key)
                results.append({
                    "text": doc,
                    "source": meta.get("source", ""),
                    "article": meta.get("article", ""),
                    "category": meta.get("category", ""),
                })

    # 1단계: 각 키워드가 실제로 포함된 청크를 벡터 유사도로 랭킹
    for kw in keywords:
        if len(results) >= n_results * 2:
            break
        try:
            kwargs: dict = dict(
                query_texts=[query],
                n_results=min(n_results * 3, 30),
                where_document={"$contains": kw},
            )
            if where:
                kwargs["where"] = where
            r = col.query(**kwargs)
            _add(r["documents"][0], r["metadatas"][0])
        except Exception:
            pass

    # 2단계: 키워드 결과 부족 시 순수 벡터 검색으로 보충
    if len(results) < n_results:
        try:
            kwargs = dict(query_texts=[query], n_results=n_results * 2)
            if where:
                kwargs["where"] = where
            r = col.query(**kwargs)
            _add(r["documents"][0], r["metadatas"][0])
        except Exception:
            pass

    return results[:n_results]


def search_for_review(doc_type: str, content_snippet: str, n_results: int = 8) -> list[dict]:
    """설계서 검토용 관련 조항 검색"""
    type_keywords = {
        "내역서": "원가계산 간접공사비 산출내역 계약금액",
        "시방서": "시방서 품질기준 자재규격 시공기준",
        "도면": "도면 설계도서 설계기준",
        "견적서": "견적 수의계약 비교견적 계약방법",
    }
    extra = type_keywords.get(doc_type, "설계서 시설공사")
    query = f"설계서 {doc_type} 검토 {extra} {content_snippet[:150]}"
    return search_hybrid(query, n_results=n_results)


def format_for_prompt(items: list[dict], max_chars: int = 4000) -> str:
    """검색 결과를 Claude 프롬프트용 텍스트로 변환"""
    lines = []
    total = 0
    for item in items:
        block = f"[출처: {item['source']} / {item['article']}]\n{item['text']}\n"
        if total + len(block) > max_chars:
            break
        lines.append(block)
        total += len(block)
    return "\n".join(lines)


def highlight(text: str, keywords: list[str]) -> str:
    """검색어 키워드 하이라이트 (HTML)"""
    import re
    for kw in keywords:
        if not kw:
            continue
        text = re.sub(
            f"({re.escape(kw)})",
            r'<mark style="background:#fff176;padding:0 2px;border-radius:2px;">\1</mark>',
            text,
        )
    return text
