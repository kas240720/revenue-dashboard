# Revenue Overview Dashboard — 설계 문서

> 작성일: 2026-03-13  
> 목적: 이 문서는 Revenue Overview 대시보드 구축을 위해 나눈 대화와 설계 내용을 정리한 것입니다. 나중에 읽어도 전체 맥락과 구조를 빠르게 파악할 수 있도록 작성되었습니다.

---

## 배경 및 목표

브로커별로 개별 **Trade Blotter (Excel 파일)** 을 보유하고 있으며, 이 데이터를 통합하여 아래 4가지 Revenue 지표를 한눈에 볼 수 있는 대시보드를 구축하는 것이 목표입니다.

### 대시보드 4개 섹션

| # | 섹션 | 설명 |
|---|------|------|
| A | **Daily Commission Revenue** | 날짜별 커미션 수익 |
| B | **Monthly Revenue Trend** | 월별 수익 추세 |
| C | **Revenue by Desk** | 데스크별 수익 |
| D | **Revenue by Broker** | 브로커별 수익 |

---

## 데이터 현황

- **데이터 형식**: Excel (.xlsx)
- **파일 구조**: 브로커마다 별도의 파일 (예: `broker_A.xlsx`, `broker_B.xlsx`, ...)
- **주요 필드**:
  - `Trade Date` — 거래일
  - `Commission` — 커미션 금액
  - `Notional` — 명목 거래금액
  - `Desk` — 담당 데스크
  - `수량` — 거래 수량
  - `금액` — 거래 금액
  - 기타 필드 존재

---

## 기술 스택

```
Excel (브로커별 .xlsx)
        ↓
Python (pandas + openpyxl)   ← 데이터 수집 및 정제
        ↓
SQL Database (SQLite or PostgreSQL)   ← 구조화된 데이터 저장
        ↓
Power BI   ← 대시보드 시각화
```

---

## 전체 아키텍처

```
data/blotters/
├── broker_A.xlsx
├── broker_B.xlsx
└── brokerN.xlsx
        ↓ (ingest.py)
    파일명에서 broker명 추출 후 데이터 읽기
        ↓ (transform.py)
    컬럼명 표준화, 날짜 파싱, 결측값 처리
        ↓ (load.py)
    SQL DB에 upsert
        ↓
    Power BI 연결 → 4개 섹션 대시보드
```

---

## 프로젝트 폴더 구조

```
revenue-dashboard/
├── data/
│   └── blotters/              # 브로커별 Excel 파일 보관
│       ├── broker_A.xlsx
│       ├── broker_B.xlsx
│       └── ...
├── pipeline/
│   ├── ingest.py              # Excel 읽기 + 브로커명 추출
│   ├── transform.py           # 컬럼 표준화, 클렌징
│   └── load.py                # SQL DB 적재 (SQLAlchemy)
├── db/
│   └── trades.db              # SQLite DB 파일
└── powerbi/
    └── revenue_dashboard.pbix # Power BI 대시보드
```

---

## SQL 데이터 모델

### `dim_broker` — 브로커 마스터 테이블
| 컬럼 | 타입 | 설명 |
|------|------|------|
| `broker_id` | INTEGER PK | 브로커 고유 ID |
| `broker_name` | TEXT | 브로커 이름 |
| `desk_id` | INTEGER FK | 소속 데스크 |

### `dim_desk` — 데스크 마스터 테이블
| 컬럼 | 타입 | 설명 |
|------|------|------|
| `desk_id` | INTEGER PK | 데스크 고유 ID |
| `desk_name` | TEXT | 데스크 이름 |

### `fact_trades` — 거래 사실 테이블 (핵심)
| 컬럼 | 타입 | 설명 |
|------|------|------|
| `trade_id` | INTEGER PK | 거래 고유 ID |
| `trade_date` | DATE | 거래일 |
| `broker_id` | INTEGER FK | 브로커 참조 |
| `desk_id` | INTEGER FK | 데스크 참조 |
| `commission` | FLOAT | 커미션 금액 |
| `notional` | FLOAT | 명목 거래금액 |
| `quantity` | FLOAT | 수량 |
| `amount` | FLOAT | 거래 금액 |

---

## Python Pipeline 개요

### 1. `ingest.py`
- `data/blotters/` 폴더 내 모든 `.xlsx` 파일을 순회
- 파일명에서 브로커명 자동 추출
- `pandas.read_excel()`로 데이터 로드

### 2. `transform.py`
- 컬럼명 표준화 매핑 (예: `"Commission $"` → `commission`)
- 날짜 형식 통일 (`trade_date`)
- 결측값 처리 및 데이터 타입 변환

### 3. `load.py`
- `SQLAlchemy`를 사용해 SQLite(또는 PostgreSQL)에 연결
- `fact_trades` 테이블에 데이터 upsert
- 중복 방지 로직 포함

---

## Power BI 대시보드 구성

| 섹션 | 차트 유형 | 주요 측정값 | 필터 |
|------|-----------|------------|------|
| **Daily Commission Revenue** | 세로 막대 차트 | `SUM(commission)` by `trade_date` | 브로커, 데스크, 기간 |
| **Monthly Revenue Trend** | 꺾은선 차트 + KPI 카드 | 월별 합계 + 전월 대비 증감률 | 연도, 데스크 |
| **Revenue by Desk** | 가로 막대 또는 파이 차트 | `SUM(commission)` by `desk_name` | 기간 |
| **Revenue by Broker** | 랭킹 막대 차트 | `SUM(commission)` by `broker_name` | 기간, 데스크 |

---

## 다음 단계 (To-Do)

- [ ] SQL 스키마 DDL 파일 작성 (`schema.sql`)
- [ ] `ingest.py`, `transform.py`, `load.py` 구현
- [ ] 샘플 데이터로 파이프라인 테스트
- [ ] Power BI에서 DB 연결 및 4개 섹션 대시보드 완성
