# -*- coding: utf-8 -*-
"""
효주기업 견적서 대조·추출 도구
─────────────────────────────────────────────
O:\ 견적서 폴더(회사이름\...\년\월\일)를 훑어서
  1) 폴더의 회사·날짜  vs  PDF 안의 날짜  일치 여부 확인
  2) 자재비 / 가공비 / 합계 자동 추출
  3) 결과를 CSV(엑셀) + JSON(앱 가져오기용) 으로 저장
(선택) 도면 폴더도 주면 GEO/DXF 를 회사·날짜로 견적과 짝지어 줍니다.

■ 준비 (최초 1회): 명령프롬프트에서
      pip install pdfplumber
■ 실행: 아래 '설정'만 본인 경로로 고치고 저장한 뒤
      python 효주_견적대조.py
"""

# ═══════════════ 설정 (여기만 고치세요) ═══════════════
견적서_폴더 = r"O:\김유진\효주기업"          # 견적서 PDF들이 들어있는 최상위 폴더
도면_폴더   = r""                              # (선택) GEO/DXF 폴더. 비우면 도면대조 생략
결과_폴더   = r"O:\김유진\효주기업\_견적대조결과"   # 결과 CSV/JSON 저장 위치
# ══════════════════════════════════════════════════════

import os, re, csv, json, sys

try:
    import pdfplumber
except ImportError:
    print("[안내] pdfplumber 가 설치되어 있지 않습니다.")
    print("      명령프롬프트에 아래를 입력해 설치 후 다시 실행하세요:")
    print("      pip install pdfplumber")
    sys.exit(1)

FAM_PAT = [('SUS','SUS'),('STS','SUS'),('EGI','EGI'),('SGCC','GI'),('GI','GI'),
           ('SK5','SK5'),('SPHC','SS'),('SPCC','CR'),('AL','AL'),('CR','CR'),
           ('PO','PO'),('HOT','HOT'),('SS','SS')]

def won2int(s):
    if s is None: return None
    d = re.sub(r'[^0-9]', '', str(s))
    return int(d) if d else None

def parse_material(spec):
    s = (spec or '').upper().replace(' ', '')
    fam = next((f for k, f in FAM_PAT if k in s), None)
    m = re.search(r'(\d+(?:\.\d+)?)\s*T', s)
    return fam, (float(m.group(1)) if m else None)

def folder_meta(path):
    """경로 세그먼트에서 (회사명, 날짜YYYY-MM-DD) 추출. 초성 한글자 폴더는 건너뜀."""
    segs = re.split(r'[\\/]+', path)
    for i, s in enumerate(segs):
        if re.fullmatch(r'20\d{2}', s) and i + 2 < len(segs) \
           and re.fullmatch(r'\d{1,2}', segs[i+1]) and re.fullmatch(r'\d{1,2}', segs[i+2]):
            date = f"{segs[i]}-{int(segs[i+1]):02d}-{int(segs[i+2]):02d}"
            j = i - 1
            while j > 0 and len(segs[j]) <= 1:  # 초성폴더(ㄷ 등) 스킵
                j -= 1
            return (segs[j] if j >= 0 else None), date
    m = re.search(r'(\d{2})(\d{2})(\d{2})', segs[-1])  # 파일명 6자리 날짜 대비
    if m:
        return None, f"20{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return None, None

def parse_quote_pdf(path):
    """견적서 PDF → 날짜, 라인아이템, 자재/가공/합계."""
    out = {'날짜': None, 'items': [], '자재비': None, '가공비': None, '합계': None}
    with pdfplumber.open(path) as pdf:
        txt = '\n'.join((p.extract_text() or '') for p in pdf.pages)
        tables = [t for p in pdf.pages for t in (p.extract_tables() or [])]
    m = re.search(r'DATE\s*(\d{4}-\d{2}-\d{2})', txt)
    if m: out['날짜'] = m.group(1)
    m = re.search(r'견적금액\s*([\d,]+)', txt) or re.search(r'금\s*액\s*합\s*계[^0-9]*([\d,]+)', txt)
    if m: out['합계'] = won2int(m.group(1))
    m = re.search(r'②[^0-9]*([\d,]+)', txt)
    if m: out['가공비'] = won2int(m.group(1))
    for t in tables:
        for row in t:
            if len(row) >= 6 and str(row[0]).strip().isdigit():
                spec = (row[1] or '').strip()
                nums = [won2int(c) for c in row if won2int(c) and won2int(c) > 100]
                if spec and len(nums) >= 3:
                    out['items'].append({'규격': spec, '자재': nums[0], '가공': nums[1], '금액': nums[2]})
    if out['items']:
        if out['자재비'] is None:
            out['자재비'] = sum(i['자재'] for i in out['items'] if i['자재'])
        if out['가공비'] is None:
            out['가공비'] = sum(i['가공'] for i in out['items'] if i['가공'])
        if out['합계'] is None:
            out['합계'] = sum(i['금액'] for i in out['items'] if i['금액'])
    return out

def geo_source_path(path):
    """GEO 파일 안의 원본 DXF 경로(CP949)를 디코딩해 회사·날짜 추출용으로 반환."""
    try:
        for line in open(path, 'rb').read().split(b'\n'):
            if line.startswith(b'SOURCE_FILE_NAME@'):
                p = line[len(b'SOURCE_FILE_NAME@'):].strip()
                try: return p.decode('cp949')
                except Exception: return p.decode('latin-1', 'replace')
    except Exception:
        pass
    return None

def main():
    os.makedirs(결과_폴더, exist_ok=True)
    rows, records = [], []
    n_ok = n_datemis = n_fail = 0

    print(f"[1/2] 견적서 폴더 스캔: {견적서_폴더}")
    for root, _, files in os.walk(견적서_폴더):
        if os.path.abspath(root).startswith(os.path.abspath(결과_폴더)):
            continue
        for fn in files:
            if not fn.lower().endswith('.pdf'):
                continue
            full = os.path.join(root, fn)
            company, fdate = folder_meta(full)
            try:
                q = parse_quote_pdf(full)
            except Exception as e:
                n_fail += 1
                rows.append([full, company, fdate, '', '', '', '', '', '', '', f'추출실패: {e}'])
                continue
            spec = q['items'][0]['규격'] if q['items'] else ''
            fam, thk = parse_material(spec)
            pdate = q['날짜']
            if q['합계'] is None:
                status = '금액못읽음'; n_fail += 1
            elif fdate and pdate and fdate != pdate:
                status = '⚠날짜불일치'; n_datemis += 1
            else:
                status = '정상'; n_ok += 1
            rows.append([full, company, fdate, pdate,
                         '일치' if (fdate == pdate and fdate) else '',
                         spec, fam or '', thk or '', q['자재비'], q['가공비'], q['합계'], status])
            if q['합계'] is not None:
                records.append({
                    'v': company, 'd': pdate or fdate, 'm': spec or (f"{fam or ''} {thk or ''}").strip(),
                    'f': fam, 't': thk, 'jc': q['자재비'], 'gc': q['가공비'], 'h': q['합계'],
                    'src': 0, '_job': True, '_file': os.path.relpath(full, 견적서_폴더),
                    '_items': q['items'],
                })

    # (선택) 도면 대조
    draw_rows = []
    if 도면_폴더 and os.path.isdir(도면_폴더):
        print(f"[+] 도면 폴더 스캔: {도면_폴더}")
        idx = {}
        for r in records:
            idx.setdefault((r['v'], r['d']), []).append(r)
        for root, _, files in os.walk(도면_폴더):
            for fn in files:
                low = fn.lower()
                if not (low.endswith('.geo') or low.endswith('.dxf')):
                    continue
                full = os.path.join(root, fn)
                src = geo_source_path(full) if low.endswith('.geo') else full
                dc, dd = folder_meta(src or full)
                hit = idx.get((dc, dd))
                draw_rows.append([fn, dc, dd, '견적있음' if hit else '견적없음',
                                  (hit[0]['h'] if hit else '')])

    # ── 저장 ──
    csv_path = os.path.join(결과_폴더, '효주_견적대조표.csv')
    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerow(['파일경로', '폴더회사', '폴더날짜', 'PDF날짜', '날짜일치',
                    '품명', '재질', '두께', '자재비', '가공비', '합계', '상태'])
        w.writerows(rows)
    json_path = os.path.join(결과_폴더, '효주_견적이력.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(records, f, ensure_ascii=False, indent=1)

    if draw_rows:
        with open(os.path.join(결과_폴더, '도면_견적_매칭.csv'), 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f); w.writerow(['도면파일', '회사', '날짜', '견적매칭', '견적합계']); w.writerows(draw_rows)

    print(f"\n[2/2] 완료")
    print(f"  정상 {n_ok} · ⚠날짜불일치 {n_datemis} · 실패 {n_fail}  (총 {len(rows)}건)")
    print(f"  → 대조표(엑셀): {csv_path}")
    print(f"  → 이력 JSON  : {json_path}")
    if draw_rows:
        print(f"  → 도면매칭   : {len(draw_rows)}건")
    print("\n※ '⚠날짜불일치' 행은 폴더 날짜와 견적서 날짜가 달라 확인이 필요한 건입니다.")

if __name__ == '__main__':
    main()
