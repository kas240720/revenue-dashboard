# Revenue Overview Dashboard — 설계 문서

> 최초 작성: 2026-03-13 / 마지막 업데이트: 2026-03-16  
> 목적: 이 문서는 Revenue Overview 대시보드 구축을 위해 나눈 대화와 설계 내용을 정리한 것입니다. 나중에 읽어도 전체 맥락과 구조를 빠르게 파악할 수 있도록 작성되었습니다.

---

## AI를 위한 프로젝트 컨텍스트

이 문서만 보고 바로 작업을 이어받을 수 있도록 핵심 맥락을 정리합니다.

- **현재 상태**: `pipeline.py` 구현 및 테스트 완료. DB 적재 확인됨. 다음 단계는 Power BI 연결
- **작업자**: 금융회사 내부 직원. 브로커별 거래 데이터를 Excel로 관리하고 있으며, 이를 통합 시각화하는 것이 목표
- **파일 규모**: Excel 파일 총 4개 (브로커별 1개씩). 대용량 아님, 단순하게 유지할 것
- **우선순위**: 복잡한 설계보다 빠르게 동작하는 결과물. 확장성보다 실용성 우선
- **DB 선택**: SQLite (로컬 파일, 서버 불필요)
- **시각화 도구**: Power BI Desktop (로컬 환경, 클라우드 불필요)
- **Python 환경**: Anaconda (base 환경), `C:\Users\Minji Song\anaconda3\python.exe`

---

## 배경 및 목표

브로커별로 개별 **Trade Blotter (Excel 파일, 총 4개)** 을 보유하고 있으며, 이 데이터를 통합하여 아래 4가지 Revenue 지표를 한눈에 볼 수 있는 대시보드를 구축하는 것이 목표입니다.

### 대시보드 4개 섹션

| # | 섹션 | 설명 |
|---|------|------|
| A | **Daily Commission Revenue** | 날짜별 커미션 수익 |
| B | **Monthly Revenue Trend** | 월별 수익 추세 |
| C | **Revenue by Desk** | 데스크별 수익 |
| D | **Revenue by Broker** | 브로커별 수익 |

---

## 데이터 현황

### 실제 파일 목록 (`data/blotters/`)

| 파일명 | 형식 | 브로커 | 사용 시트 |
|--------|------|--------|-----------|
| `Trade Blotter_2026_031326.xlsm` | .xlsm | **Apex** | `Combined` |
| `Trade Blotter_Cantor_031326.xlsb` | .xlsb | **Cantor** | `Cantor` |
| `Trade Blotter_DriveWealth_031326.xlsb` | .xlsb | **DriveWealth** | `DW` |
| `Trade Blotter_Velocity_031326.xlsb` | .xlsb | **Velocity** | `Velocity` |

> Apex 파일은 파일명에 브로커명 대신 연도(2026)만 표기되어 있음. 나머지 3개는 파일명에 브로커명 포함.

### 실제 컬럼 매핑 (확인 완료)

| 브로커 | 원본 컬럼명 | 표준화 후 | 비고 |
|--------|------------|-----------|------|
| Apex | `Trade Date` | `trade_date` | datetime 형식 |
| Apex | `Desk` | `desk_name` | Institutions / Retail / Retail Offline |
| Apex | `Qty` | `quantity` | |
| Apex | `Principal` | `principal` | |
| Apex | `Commission` | `commission` | |
| Cantor/DW/Velocity | `TD` | `trade_date` | "MM/DD/YYYY" 문자열 → 변환 필요 |
| Cantor/DW/Velocity | `Shs` | `quantity` | |
| Cantor/DW/Velocity | `Principal` | `principal` | |
| Cantor/DW/Velocity | `Comm` | `commission` | |
| Cantor/DW/Velocity | (없음) | `desk_name` | NULL — 해당 파일에 데스크 정보 없음 |

---

## 기술 스택

```
Excel 4개 (.xlsm / .xlsb)
        ↓
Python - pandas + openpyxl + pyxlsb   ← 데이터 수집 및 정제
        ↓
SQLite (trades.db)                     ← 로컬 DB 저장
        ↓
Power BI Desktop                       ← 대시보드 시각화
```

---

## 프로젝트 폴더 구조

```
revenue-dashboard/
├── data/
│   └── blotters/                        # 브로커별 Excel 파일 4개 (gitignore됨)
│       ├── Trade Blotter_2026_031326.xlsm        ← Apex
│       ├── Trade Blotter_Cantor_031326.xlsb      ← Cantor
│       ├── Trade Blotter_DriveWealth_031326.xlsb ← DriveWealth
│       └── Trade Blotter_Velocity_031326.xlsb    ← Velocity
├── pipeline.py                          # 전체 파이프라인 (읽기 + 정제 + DB 저장)
├── db/
│   └── trades.db                        # SQLite DB (gitignore됨)
├── requirements.txt                     # Python 패키지 버전 명세
├── .gitignore                           # data/, db/, *.pbix 제외
└── powerbi/
    └── revenue_dashboard.pbix           # Power BI 대시보드 (미완성)
```

---

## SQL 데이터 모델

단일 플랫 테이블. broker/desk는 별도 차원 테이블 없이 텍스트로 저장.

### `trades` 테이블

| 컬럼 | 타입 | 설명 |
|------|------|------|
| `trade_date` | DATE | 거래일 |
| `broker_name` | TEXT | Apex / Cantor / DriveWealth / Velocity |
| `desk_name` | TEXT | Apex만 존재 (Institutions / Retail / Retail Offline). 나머지는 NULL |
| `quantity` | FLOAT | 수량 |
| `principal` | FLOAT | 명목 거래금액 |
| `commission` | FLOAT | 커미션 금액 |

### 적재 결과 (2026-03-16 기준)

| 브로커 | 행 수 | 총 커미션 |
|--------|-------|-----------|
| Apex | 94,252 | $3,765,045 |
| DriveWealth | 78,247 | $2,067,478 |
| Velocity | 19,562 | $852,330 |
| Cantor | 1,128 | $36,066 |
| **합계** | **193,189** | **$6,720,919** |

---

## Python Pipeline (`pipeline.py`)

파이프라인 전체를 파일 하나에서 관리. 실행 시 DB를 초기화하고 4개 파일을 새로 적재.

### 동작 순서

1. `data/blotters/` 폴더 순회
2. 파일명으로 브로커 판별 → 브로커별 읽기 함수 호출
3. 컬럼명 표준화 + 날짜 변환 + 결측값 제거
4. 4개 DataFrame을 하나로 합치기
5. SQLite `trades` 테이블에 전체 재적재 (`if_exists="replace"`)

### 실행 방법

```bash
"C:\Users\Minji Song\anaconda3\python.exe" pipeline.py
```

### 결과 확인 (Python)

```bash
"C:\Users\Minji Song\anaconda3\python.exe" -c "import sqlite3, pandas as pd; con = sqlite3.connect(r'db/trades.db'); print(pd.read_sql('SELECT broker_name, COUNT(*) as 건수, ROUND(SUM(commission),2) as 총커미션 FROM trades GROUP BY broker_name', con))"
```

---

## Power BI 대시보드 구성

| 섹션 | 차트 유형 | 주요 측정값 | 필터 |
|------|-----------|------------|------|
| **Daily Commission Revenue** | 세로 막대 차트 | `SUM(commission)` by `trade_date` | 브로커, 데스크, 기간 |
| **Monthly Revenue Trend** | 꺾은선 차트 + KPI 카드 | 월별 합계 + 전월 대비 증감률 | 연도, 데스크 |
| **Revenue by Desk** | 가로 막대 차트 | `SUM(commission)` by `desk_name` | 기간 |
| **Revenue by Broker** | 랭킹 막대 차트 | `SUM(commission)` by `broker_name` | 기간, 데스크 |

- **날짜 테이블**: Power BI 자동 날짜 계층 사용 (별도 dim_date 불필요)
- **DB 연결**: Power BI Desktop → Get Data → SQLite → `db/trades.db`
- **데이터 갱신**: `pipeline.py` 실행 후 Power BI Desktop에서 수동 새로고침

> ⚠️ Revenue by Desk는 Apex 데이터만 desk_name 존재. Cantor/DW/Velocity는 NULL이므로 필터 시 주의.

---

## 다음 단계 (To-Do)

- [x] `requirements.txt` 작성
- [x] `pipeline.py` 구현 및 테스트
- [x] DB 적재 결과 확인
- [ ] Power BI에서 SQLite DB 연결
- [ ] 4개 섹션 대시보드 완성

---

## 작업 재개 시 AI에게

이 프로젝트를 이어받는다면 아래 순서로 진행하면 됩니다:

1. 이 문서 전체 읽기
2. `pipeline.py` 실행해서 DB 최신화 (Excel 파일이 업데이트됐을 경우)
3. Power BI Desktop → `db/trades.db` 연결 → 4개 섹션 대시보드 구성

**절대 과설계 금지.** 파일 4개짜리 소규모 프로젝트이므로 단순하고 빠르게 동작하는 것 우선.
