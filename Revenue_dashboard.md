# Revenue Overview Dashboard — 설계 문서

> 작성일: 2026-03-13  
> 목적: 이 문서는 Revenue Overview 대시보드 구축을 위해 나눈 대화와 설계 내용을 정리한 것입니다. 나중에 읽어도 전체 맥락과 구조를 빠르게 파악할 수 있도록 작성되었습니다.

---

## AI를 위한 프로젝트 컨텍스트

이 문서만 보고 바로 작업을 이어받을 수 있도록 핵심 맥락을 정리합니다.

- **현재 상태**: 설계 문서만 존재. `pipeline.py`, `requirements.txt`, Power BI 파일 모두 아직 미구현
- **작업자**: 금융회사 내부 직원. 브로커별 거래 데이터를 Excel로 관리하고 있으며, 이를 통합 시각화하는 것이 목표
- **파일 규모**: Excel 파일 총 4개 (브로커별 1개씩). 대용량 아님, 단순하게 유지할 것
- **우선순위**: 복잡한 설계보다 빠르게 동작하는 결과물. 확장성보다 실용성 우선
- **DB 선택**: SQLite (로컬 파일, 서버 불필요)
- **시각화 도구**: Power BI Desktop (로컬 환경, 클라우드 불필요)

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

- **데이터 형식**: Excel (.xlsx), 총 4개 파일
- **파일 구조**: 브로커마다 별도의 파일 (예: `broker_A.xlsx`, `broker_B.xlsx`, ...)
- **브로커명 추출 방식**: 파일명에서 자동 추출. 예) `broker_A.xlsx` → `broker_name = "broker_A"`
- **주요 필드 (Excel 원본 컬럼명 → 표준화 후)**:

| Excel 원본 컬럼명 (예시) | 표준화 후 컬럼명 | 비고 |
|--------------------------|-----------------|------|
| `Trade Date` / `거래일` | `trade_date` | 날짜 형식 통일 필요 |
| `Commission` / `Commission $` | `commission` | float 변환 |
| `Notional` | `notional` | float 변환 |
| `Desk` / `데스크` | `desk_name` | 텍스트 |
| `수량` / `Quantity` | `quantity` | float 변환 |
| `금액` / `Amount` | `amount` | float 변환 |

> ⚠️ 실제 Excel 파일의 컬럼명은 브로커마다 다를 수 있음. `pipeline.py` 구현 전 실제 파일 컬럼 확인 후 매핑 테이블 업데이트 필요.

---

## 기술 스택

```
Excel 4개 (.xlsx)
        ↓
Python (pandas + openpyxl)   ← 데이터 수집 및 정제
        ↓
SQLite (.db)                 ← 로컬 DB 저장
        ↓
Power BI                     ← 대시보드 시각화
```

---

## 프로젝트 폴더 구조

```
revenue-dashboard/
├── data/
│   └── blotters/              # 브로커별 Excel 파일 4개
│       ├── broker_A.xlsx
│       ├── broker_B.xlsx
│       └── ...
├── pipeline.py                # 전체 파이프라인 (읽기 + 정제 + DB 저장)
├── db/
│   └── trades.db              # SQLite DB 파일
├── requirements.txt           # Python 패키지 버전 명세
└── powerbi/
    └── revenue_dashboard.pbix # Power BI 대시보드
```

---

## SQL 데이터 모델

파일이 4개로 소규모이므로 단일 플랫 테이블로 단순하게 관리합니다.  
broker, desk는 별도 차원 테이블 없이 텍스트로 저장합니다.

### `trades` — 거래 통합 테이블

| 컬럼 | 타입 | 설명 |
|------|------|------|
| `id` | INTEGER PK | 자동 증가 |
| `trade_date` | DATE | 거래일 |
| `broker_name` | TEXT | 파일명에서 자동 추출 |
| `desk_name` | TEXT | 담당 데스크 |
| `commission` | FLOAT | 커미션 금액 |
| `notional` | FLOAT | 명목 거래금액 |
| `quantity` | FLOAT | 수량 |
| `amount` | FLOAT | 거래 금액 |

---

## Python Pipeline 개요 (`pipeline.py`)

파이프라인 전체를 파일 하나에서 관리합니다. 실행 시 DB를 초기화하고 4개 파일을 새로 적재합니다.

### 동작 순서

1. `data/blotters/` 폴더 내 `.xlsx` 파일 전체 읽기
2. 파일명에서 브로커명 자동 추출
3. 컬럼명 표준화 (예: `"Commission $"` → `commission`)
4. 날짜 형식 통일, 결측값 처리
5. SQLite `trades` 테이블에 전체 재적재 (매 실행 시 덮어쓰기)

### 실행 방법

```bash
python pipeline.py
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
- **데이터 갱신**: `pipeline.py` 실행 후 Power BI Desktop에서 수동 새로고침

---

## 다음 단계 (To-Do)

- [ ] `requirements.txt` 작성
- [ ] `pipeline.py` 구현 및 테스트
- [ ] Power BI에서 SQLite DB 연결
- [ ] 4개 섹션 대시보드 완성

---

## 작업 재개 시 AI에게

이 프로젝트를 이어받는다면 아래 순서로 진행하면 됩니다:

1. 이 문서를 전부 읽고 전체 구조 파악
2. `data/blotters/` 안의 실제 Excel 파일 컬럼명 확인 후 위 매핑 테이블 업데이트
3. `requirements.txt` 생성 → `pipeline.py` 작성 → 테스트
4. Power BI 연결 및 대시보드 구성

**절대 과설계 금지.** 파일 4개짜리 소규모 프로젝트이므로 단순하고 빠르게 동작하는 것 우선.
