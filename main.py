"""指数趋势强度统计工具 - 主脚本"""

import os
import time
from datetime import datetime
from typing import Optional

import akshare as ak
import numpy as np
import pandas as pd
from tabulate import tabulate

from config import INDEX_CONFIG, MA_PERIOD, START_DATE

# 最大重试次数
MAX_RETRIES = 3
RETRY_DELAY = 2  # 秒


def fetch_with_retry(fetch_fn, retries=MAX_RETRIES):
    """带重试的数据获取"""
    for attempt in range(retries):
        try:
            return fetch_fn()
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(RETRY_DELAY)
            else:
                raise e


def fetch_index_data(cfg: dict) -> pd.DataFrame:
    """根据配置获取指数日线数据，返回统一格式 DataFrame[date, close]"""
    source = cfg["source"]
    symbol = cfg["symbol"]
    today = datetime.now().strftime("%Y%m%d")

    if source == "a_share_sina":
        # 新浪A股指数接口，symbol 需要 sh/sz 前缀
        df = fetch_with_retry(lambda: ak.stock_zh_index_daily(symbol=symbol))

    elif source == "a_share_em":
        # 东财A股指数接口，用于新浪不支持的指数
        df = fetch_with_retry(lambda: ak.index_zh_a_hist(
            symbol=symbol, period="daily",
            start_date=START_DATE, end_date=today
        ))
        df = df.rename(columns={"日期": "date", "收盘": "close"})

    elif source == "hk_sina":
        # 新浪港股指数接口
        df = fetch_with_retry(lambda: ak.stock_hk_index_daily_sina(symbol=symbol))

    elif source == "us_etf":
        # 美股ETF接口
        df = fetch_with_retry(lambda: ak.stock_us_daily(symbol=symbol, adjust="qfq"))

    elif source == "global_sina":
        # 新浪全球指数接口
        df = fetch_with_retry(lambda: ak.index_global_hist_sina(symbol=symbol))

    else:
        raise ValueError(f"未知的数据源类型: {source}")

    # 统一列名
    if "date" not in df.columns:
        # 可能列名是中文
        col_map = {}
        for col in df.columns:
            if col in ("日期",):
                col_map[col] = "date"
            elif col in ("收盘",):
                col_map[col] = "close"
        if col_map:
            df = df.rename(columns=col_map)

    df = df[["date", "close"]].copy()
    df["date"] = pd.to_datetime(df["date"])
    df["close"] = df["close"].astype(float)
    df = df.sort_values("date").reset_index(drop=True)

    # 只保留 START_DATE 之后的数据
    start = pd.to_datetime(START_DATE)
    df = df[df["date"] >= start].reset_index(drop=True)

    return df


def detect_cross(df: pd.DataFrame) -> Optional[dict]:
    """
    检测最近一次收盘价穿越MA20的事件。
    返回: {cross_date, critical_value, cross_close, direction}
    """
    valid = df.dropna(subset=["ma20"]).reset_index(drop=True)
    if len(valid) < 2:
        return None

    # 判断每日 close 相对 MA20 的位置: 1=上方, -1=下方
    valid["position"] = np.where(valid["close"] > valid["ma20"], 1, -1)

    # 从最新往前找位置变化点
    for i in range(len(valid) - 1, 0, -1):
        if valid["position"].iloc[i] != valid["position"].iloc[i - 1]:
            return {
                "cross_date": valid["date"].iloc[i],
                "critical_value": round(valid["ma20"].iloc[i], 2),
                "cross_close": round(valid["close"].iloc[i], 2),
                "direction": "上穿" if valid["position"].iloc[i] == 1 else "下穿",
            }

    # 没有穿越：取第一个有效数据
    return {
        "cross_date": valid["date"].iloc[0],
        "critical_value": round(valid["ma20"].iloc[0], 2),
        "cross_close": round(valid["close"].iloc[0], 2),
        "direction": "上穿" if valid["position"].iloc[-1] == 1 else "下穿",
    }


def calc_daily_change(df: pd.DataFrame) -> float:
    """计算当日涨幅%"""
    if len(df) < 2:
        return 0.0
    prev_close = df["close"].iloc[-2]
    curr_close = df["close"].iloc[-1]
    return round((curr_close - prev_close) / prev_close * 100, 2)


def process_index(cfg: dict) -> Optional[dict]:
    """处理单个指数，返回结果字典"""
    try:
        df = fetch_index_data(cfg)
        if len(df) < MA_PERIOD + 1:
            print(f"数据不足({len(df)}条)，跳过")
            return None

        # 计算 MA20
        df["ma20"] = df["close"].rolling(MA_PERIOD).mean()

        # 检测穿越
        cross_info = detect_cross(df)
        if cross_info is None:
            print("无法检测穿越，跳过")
            return None

        # 当日数据
        latest_close = df["close"].iloc[-1]
        latest_date = df["date"].iloc[-1]
        daily_change = calc_daily_change(df)

        # 偏离率 = (收盘价 - 临界值点) / 临界值点
        critical_value = cross_info["critical_value"]
        deviation = round((latest_close - critical_value) / critical_value * 100, 2)

        # 区间涨幅 = (今日收盘 - 穿越当天收盘价) / 穿越当天收盘价
        cross_close = cross_info["cross_close"]
        range_change = round((latest_close - cross_close) / cross_close * 100, 2)

        # 状态
        status = "YES" if deviation > 0 else "NO"

        return {
            "指数名称": cfg["name"],
            "状态": status,
            "当日涨幅%": daily_change,
            "收盘点位": round(latest_close, 2),
            "临界值点": critical_value,
            "偏离率%": deviation,
            "穿越日期": cross_info["cross_date"].strftime("%Y-%m-%d"),
            "区间涨幅%": range_change,
            "趋势方向": cross_info["direction"],
            "数据日期": latest_date.strftime("%Y-%m-%d"),
        }
    except Exception as e:
        print(f"获取失败: {e}")
        return None


def export_markdown(result_df: pd.DataFrame, output_dir: str = "output"):
    """导出结果为 Markdown 表格文件"""
    if result_df.empty:
        print("没有可用数据。")
        return

    os.makedirs(output_dir, exist_ok=True)

    data_date = result_df["数据日期"].iloc[0]
    today_str = datetime.now().strftime("%Y%m%d")
    filepath = os.path.join(output_dir, f"trend_{today_str}.md")

    export_cols = ["排名", "指数名称", "状态", "当日涨幅%", "收盘点位",
                   "临界值点", "偏离率%", "穿越日期", "区间涨幅%"]
    display_df = result_df[export_cols]

    yes_count = (result_df["状态"] == "YES").sum()
    total = len(result_df)

    md_table = tabulate(display_df, headers="keys", tablefmt="github",
                        showindex=False, numalign="right", stralign="center")

    content = (
        f"# 指数趋势强度统计\n\n"
        f"**数据日期**: {data_date} | **MA周期**: {MA_PERIOD} | "
        f"**多头(YES)**: {yes_count}/{total} | **空头(NO)**: {total - yes_count}/{total}\n\n"
        f"{md_table}\n"
    )

    # 终端打印
    print(f"\n{content}")

    # 写入文件
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"  已导出: {filepath}")


def main():
    print("正在获取指数数据...\n")

    results = []
    for i, cfg in enumerate(INDEX_CONFIG):
        print(f"  [{i + 1}/{len(INDEX_CONFIG)}] {cfg['name']}... ", end="", flush=True)
        result = process_index(cfg)
        if result:
            results.append(result)
            print(f"OK  收盘:{result['收盘点位']}  偏离率:{result['偏离率%']}%")
        time.sleep(0.5)  # 防止请求过快

    if not results:
        print("\n所有指数获取失败，请检查网络连接。")
        return

    # 构建 DataFrame，按偏离率降序排名
    result_df = pd.DataFrame(results)
    result_df = result_df.sort_values("偏离率%", ascending=False).reset_index(drop=True)
    result_df.insert(0, "排名", range(1, len(result_df) + 1))

    # 输出
    export_markdown(result_df)


if __name__ == "__main__":
    main()
