# 결정 기록(my-wiki/decisions/): d-001 d-002 d-003 d-004 d-005
#   d-001 처리상태 «종결»→«완료», «대기»→«보류»
#   d-002 처리시간 음수는 지우지 않고 _표시 칸에 이유
#   d-003 만족도 6~10점은 지우지 않고 _표시 칸에 이유 (평균 만족도에서만 제외)
#   d-004 칸 이름은 원본 그대로 (_표시 는 맨 끝에 덧붙임)
#   d-005 고객번호 앞뒤 공백 제거
# 그 밖(테스트 접수·중복 제거, 표기 통일, 7월 이후 전화 상담시간 초→분)은
# 원본 «작성기준» 시트에 적힌 내용과 앞서 하던 정리다.
# 원본은 읽기만 한다. 쓰는 파일: 정리본/상담원장_정리.csv, 정리본/지운줄.csv, 정리본/변경요약.md
import sys
from pathlib import Path
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')
BASE = Path(__file__).resolve().parent.parent
SRC = BASE.parent / 'messy' / '상담원장_2025.xlsx'
OUT = BASE / '정리본'
OUT.mkdir(exist_ok=True)

raw = pd.read_excel(SRC, sheet_name='원장', dtype=str)
cols = list(raw.columns)                      # d-004: 원본 칸 이름 그대로
n_raw = len(raw)

# ── 지우는 줄 ────────────────────────────────────────────
is_test = raw['상담원ID'] == 'AG999'
body = raw[~is_test]
order = body.sort_values(['접수번호', '최종수정일'], kind='stable')
keep_idx = order.drop_duplicates('접수번호', keep='last').index
is_dup = ~raw.index.isin(keep_idx) & ~is_test
gone = raw[is_test | is_dup].copy()
gone.insert(0, '지운이유', ['테스트 접수(상담원ID AG999)' if t else '중복(같은 접수번호, 최종수정일이 더 늦은 줄을 남김)'
                          for t in is_test[is_test | is_dup]])
gone.to_csv(OUT / '지운줄.csv', index=False, encoding='utf-8-sig')

v = raw.loc[keep_idx].copy()
r = raw.loc[v.index]                           # 비교용 원본

# ── 표기 통일 ────────────────────────────────────────────
v['접수일'] = pd.to_datetime(v['접수일'].str.replace(r'[/.]', '-', regex=True)).dt.strftime('%Y-%m-%d')
CH = {'전화': '전화', '유선': '전화', '전화상담': '전화', 'TEL': '전화',
      '챗봇': '챗봇', 'chatbot': '챗봇', 'Chatbot': '챗봇', '챗 봇': '챗봇',
      '앱': '앱', '모바일앱': '앱', 'app': '앱', 'APP': '앱',
      '이메일': '이메일', 'e-mail': '이메일', 'E-mail': '이메일', '메일': '이메일'}
v['접수채널'] = v['접수채널'].str.strip().map(CH)
v['상담원팀'] = v['상담원팀'].str.replace(' ', '')
v['월정액'] = pd.to_numeric(v['월정액'].str.replace(r'[,원]', '', regex=True))
v['고객번호'] = v['고객번호'].str.strip()                                   # d-005

sec = (v['접수채널'] == '전화') & (v['접수일'] >= '2025-07-01')
t = v['상담시간'].astype(float)
v['상담시간'] = t.where(~sec, t / 60).round(2)                              # 7월 이후 전화: 초→분

v['처리상태'] = v['처리상태'].str.strip().map(
    {'완료': '완료', '처리완료': '완료', '종결': '완료',                      # d-001
     '진행중': '진행중', '처리중': '진행중', '보류': '보류', '대기': '보류'})
v['재문의'] = v['재문의'].str.strip().str.upper().map(
    {'Y': 'Y', '예': 'Y', 'TRUE': 'Y', 'N': 'N', '아니오': 'N', 'FALSE': 'N'})

sat = pd.to_numeric(v['만족도'].str.replace('점', ''), errors='coerce')
sat = sat.where(sat != 0)                      # 0 · - · 미응답 은 응답 없음(작성기준)
v['만족도'] = sat.astype('Int64')

# ── _표시 (맨 끝 칸) ─────────────────────────────────────
neg = pd.to_numeric(v['처리시간']) < 0
bad_sat = v['만족도'] > 5
MSG_NEG = '처리시간 음수: 틀린 값이나 진짜 값은 모름, 값은 그대로 둠(d-002)'
MSG_SAT = '만족도 6~10점: NPS를 잘못 넣은 것으로 보이나 미확인, 평균 만족도에서 제외(d-003)'
mark = pd.Series('', index=v.index)
mark = mark.mask(neg, MSG_NEG)
mark = mark.mask(bad_sat & neg, MSG_NEG + '; ' + MSG_SAT)
mark = mark.mask(bad_sat & ~neg, MSG_SAT)
v['_표시'] = mark

v = v.sort_values(['접수일', '접수시각', '접수번호'])
r = raw.loc[v.index]                           # 정렬 뒤에도 같은 순서로 비교
v[cols + ['_표시']].to_csv(OUT / '상담원장_정리.csv', index=False, encoding='utf-8-sig')

# ── 다섯 숫자 (쓴 파일을 다시 읽어서 센다) ─────────────────
f = pd.read_csv(OUT / '상담원장_정리.csv', encoding='utf-8-sig', dtype={'고객번호': str}, low_memory=False)
re = f['재문의'].dropna()
sv = f.loc[~f['_표시'].fillna('').str.contains('만족도'), '만족도'].dropna()
m = {
    '상담 건수': len(f),
    '평균 상담시간(분)': f['상담시간'].mean(),
    '재문의율(%)': (re == 'Y').mean() * 100,
    '평균 만족도': sv.mean(),
    '완료율(%)': (f['처리상태'] == '완료').mean() * 100,
    '고객 수': f['고객번호'].nunique(),
}


# ── 변경요약.md ─────────────────────────────────────────
def chg(col):
    return int((v[col].astype(str) != r[col].astype(str)).sum())


st = r['처리상태'].str.strip()
n_jong, n_dae = int((st == '종결').sum()), int((st == '대기').sum())
zero_dash = int(r['만족도'].isin(['0', '-', '미응답']).sum())
re_chg = int(((v['재문의'].astype(str) != r['재문의'].astype(str)) & r['재문의'].notna()).sum())
L = []
L.append('# 변경 요약 — 상담원장_2025.xlsx → 정리본/상담원장_정리.csv\n')
L.append('스크립트: 정리/정리_상담원장.py · 결정 기록: d-001 ~ d-005 (my-wiki/decisions/)\n')
L.append('## 줄 수\n')
L.append('| 단계 | 줄 수 |\n|---|---|')
L.append(f'| 원본 | {n_raw:,} |')
L.append(f'| 테스트 접수(AG999) 삭제 | -{int(is_test.sum()):,} |')
L.append(f'| 중복 삭제(같은 접수번호, 최종수정일이 늦은 줄을 남김) | -{int(is_dup.sum()):,} |')
L.append(f'| 정리본 | {len(v):,} |\n')
L.append('삭제한 줄은 정리본/지운줄.csv 에 이유와 함께 남겼다.\n')
L.append('## 바꾼 값 (남은 줄 기준, 몇 줄)\n')
L.append('| 칸 | 무엇을 | 줄 수 | 근거 |\n|---|---|---|---|')
L.append(f'| 상담시간 | **2025-07-01 이후 전화 접수 줄을 초에서 분으로 바꿈 (60으로 나눔, 소수 둘째 자리까지)** | {int(sec.sum()):,} | 작성기준 시트 |')
L.append(f'| 처리상태 | «종결»→«완료» | {n_jong:,} | d-001 |')
L.append(f'| 처리상태 | «대기»→«보류» | {n_dae:,} | d-001 |')
L.append(f'| 처리상태 | 그 밖의 표기 통일(처리완료→완료, 처리중→진행중, 앞뒤 공백) | {chg("처리상태") - n_jong - n_dae:,} | 표기 통일 |')
L.append(f'| 고객번호 | 앞뒤 공백 제거 | {chg("고객번호"):,} | d-005 |')
L.append(f'| 접수일 | 날짜 형식을 YYYY-MM-DD 로 | {chg("접수일"):,} | 표기 통일 |')
L.append(f'| 접수채널 | 전화·챗봇·앱·이메일 4종으로 | {chg("접수채널"):,} | 표기 통일 |')
L.append(f'| 상담원팀 | 이름 속 띄어쓰기 제거 | {chg("상담원팀"):,} | 표기 통일 |')
L.append(f'| 월정액 | «원»·쉼표 제거, 숫자로 | {chg("월정액"):,} | 표기 통일 |')
L.append(f'| 재문의 | Y/N 으로 (예·TRUE→Y, 아니오·FALSE·소문자→N) | {re_chg:,} | 표기 통일 |')
L.append(f'| 만족도 | «5점»→5 처럼 «점» 제거 | {int(r["만족도"].str.endswith("점", na=False).sum()):,} | 표기 통일 |')
L.append(f'| 만족도 | 0 · - · 미응답 → 빈칸(응답 없음) | {zero_dash:,} | 작성기준 시트 |\n')
L.append('## 값은 두고 표시만 한 줄 (_표시 칸)\n')
L.append('| 무엇 | 줄 수 | 근거 |\n|---|---|---|')
L.append(f'| 처리시간 음수 | {int(neg.sum()):,} | d-002 |')
L.append(f'| 만족도 6~10점 (평균 만족도에서만 제외) | {int(bad_sat.sum()):,} | d-003 |\n')
L.append('## 바꾸지 않은 것\n')
L.append('- 칸 이름 24개는 원본 그대로, `_표시` 칸 하나만 맨 끝에 덧붙였다 (d-004)')
L.append('- 재문의·만족도·NPS·비고의 빈칸은 그대로 빈칸\n')
L.append('## 다섯 숫자와 고객 수 (정리본을 다시 읽어서 센 값)\n')
L.append('| 지표 | 값 |\n|---|---|')
L.append(f'| 상담 건수 | {m["상담 건수"]:,} |')
L.append(f'| 평균 상담시간(분) | {m["평균 상담시간(분)"]:.1f} |')
L.append(f'| 재문의율(%) | {m["재문의율(%)"]:.1f} (기록 있는 {len(re):,}건 중) |')
L.append(f'| 평균 만족도 | {m["평균 만족도"]:.2f} (응답 {len(sv):,}건, 표시한 {int(bad_sat.sum())}줄 제외) |')
L.append(f'| 완료율(%) | {m["완료율(%)"]:.1f} |')
L.append(f'| 고객 수 | {m["고객 수"]:,} |')
(OUT / '변경요약.md').write_text('\n'.join(L) + '\n', encoding='utf-8')

for k, x in m.items():
    print(k, round(x, 4) if isinstance(x, float) else x)
