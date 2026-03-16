import pandas as pd
import sqlite3
import os

# --- 설정 ---
BLOTTERS_DIR = os.path.join(os.path.dirname(__file__), "data", "blotters")
DB_PATH = os.path.join(os.path.dirname(__file__), "db", "trades.db")

# 최종 컬럼 순서
FINAL_COLS = ["trade_date", "broker_name", "desk_name", "quantity", "principal", "commission"]


# --- 브로커별 읽기 함수 ---

def read_apex(path):
    """
    Apex: Combined 시트 사용
    - Desk, Trade Date, Qty, Principal, Commission 컬럼 사용
    - desk_name이 있는 실데이터 행만 필터링
    """
    df = pd.read_excel(path, sheet_name="Combined", engine="openpyxl")
    df = df[["Desk", "Trade Date", "Qty", "Principal", "Commission"]].copy()
    df.columns = ["desk_name", "trade_date", "quantity", "principal", "commission"]

    # 유효한 행만 남기기 (Desk가 알려진 값이고 날짜/커미션이 있는 행)
    valid_desks = ["Institutions", "Retail", "Retail Offline"]
    df = df[df["desk_name"].isin(valid_desks)]
    df = df.dropna(subset=["trade_date", "commission"])

    df["broker_name"] = "Apex"
    df["trade_date"] = pd.to_datetime(df["trade_date"]).dt.date
    return df[FINAL_COLS]


def read_broker_xlsb(path, sheet_name, broker_name):
    """
    Cantor / DriveWealth / Velocity 공통
    - 브로커별 메인 시트 사용: TD, Shs, Principal, Comm
    - desk_name 없음 (해당 파일에 데스크 정보 없음)
    """
    df = pd.read_excel(path, engine="pyxlsb", sheet_name=sheet_name)
    df = df[["TD", "Shs", "Principal", "Comm"]].copy()
    df.columns = ["trade_date", "quantity", "principal", "commission"]

    df = df.dropna(subset=["trade_date", "commission"])

    df["broker_name"] = broker_name
    df["desk_name"] = None
    df["trade_date"] = pd.to_datetime(df["trade_date"], dayfirst=False).dt.date
    return df[FINAL_COLS]


# --- 메인 파이프라인 ---

def run():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)

    frames = []

    for filename in os.listdir(BLOTTERS_DIR):
        path = os.path.join(BLOTTERS_DIR, filename)
        print(f"읽는 중: {filename}")

        try:
            if "2026" in filename and filename.endswith(".xlsm"):
                # Apex (파일명에 브로커명 없고 2026만 있는 파일)
                df = read_apex(path)

            elif "Cantor" in filename:
                df = read_broker_xlsb(path, "Cantor", "Cantor")

            elif "DriveWealth" in filename:
                df = read_broker_xlsb(path, "DW", "DriveWealth")

            elif "Velocity" in filename:
                df = read_broker_xlsb(path, "Velocity", "Velocity")

            else:
                print(f"  → 처리 규칙 없음, 건너뜀")
                continue

            print(f"  → {len(df)}행 읽음")
            frames.append(df)

        except Exception as e:
            print(f"  → 오류: {e}")

    if not frames:
        print("처리된 데이터가 없습니다.")
        return

    all_trades = pd.concat(frames, ignore_index=True)

    # DB 저장 (매 실행 시 전체 교체)
    con = sqlite3.connect(DB_PATH)
    all_trades.to_sql("trades", con, if_exists="replace", index=False)
    con.close()

    print(f"\n완료: 총 {len(all_trades)}행 → {DB_PATH}")
    print(all_trades.groupby("broker_name")["commission"].sum().rename("총 커미션"))


if __name__ == "__main__":
    run()
