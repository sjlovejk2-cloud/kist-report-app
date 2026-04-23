import pdfplumber, io, sys, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

FILES = {
    "계약업무요령": r"F:\스마트규정시스템\계약관련 정리\계약업무요령(2024년도 11월 개정).pdf",
    "회계규정": r"F:\스마트규정시스템\계약관련 정리\회계규정(2021년도 4월 개정).pdf",
}

KEYWORDS = ['수의계약','고시금액','추정가격','추정금액','경쟁입찰','제한경쟁','지명경쟁',
            '천만원','억원','만원','2천만','5천만','1억','2억','3억','5억',
            '천만 원','억 원','제26조','제24조','제23조','제27조','제28조',
            '소액수의','소액 수의','수의계약기준']

for label, fp in FILES.items():
    if not os.path.exists(fp):
        print(f"[없음] {fp}")
        continue
    print(f"\n{'='*60}")
    print(f"  {label}  ({fp})")
    print(f"{'='*60}")
    with pdfplumber.open(fp) as pdf:
        print(f"  총 {len(pdf.pages)}페이지\n")
        for i, pg in enumerate(pdf.pages, 1):
            t = pg.extract_text() or ''
            if any(kw in t for kw in KEYWORDS):
                print(f"--- p{i} ---")
                print(t[:4000])
                print()
