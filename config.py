# 指数趋势强度统计工具 - 配置文件

# MA 周期
MA_PERIOD = 20

# 数据获取起始日期（需要足够天数计算MA20，取3个月数据）
START_DATE = "20250101"

# 指数配置列表
# source 类型:
#   a_share_sina  - 新浪A股指数接口 (stock_zh_index_daily)，需要 sh/sz 前缀
#   a_share_em    - 东财A股指数接口 (index_zh_a_hist)，用于新浪不支持的指数
#   hk_sina       - 新浪港股指数接口 (stock_hk_index_daily_sina)
#   us_etf        - 美股ETF接口 (stock_us_daily)
#   global_sina   - 新浪全球指数接口 (index_global_hist_sina)
INDEX_CONFIG = [
    # ========== A股指数（新浪接口） ==========
    {"name": "上证50",   "source": "a_share_sina", "symbol": "sh000016"},
    {"name": "沪深300",  "source": "a_share_sina", "symbol": "sz399300"},
    {"name": "中证500",  "source": "a_share_sina", "symbol": "sh000905"},
    {"name": "中证1000", "source": "a_share_sina", "symbol": "sh000852"},
    {"name": "中证A500", "source": "a_share_sina", "symbol": "sh000510"},
    {"name": "创业板指", "source": "a_share_sina", "symbol": "sz399006"},
    {"name": "科创50",   "source": "a_share_sina", "symbol": "sh000688"},

# ========== 港股指数 ==========
    {"name": "恒生指数", "source": "hk_sina", "symbol": "HSI"},
    {"name": "国企指数", "source": "hk_sina", "symbol": "HSCEI"},
    {"name": "恒生科技", "source": "hk_sina", "symbol": "HSTECH"},

    # ========== 美股ETF ==========
    {"name": "标普500",  "source": "us_etf", "symbol": "SPY"},
    {"name": "纳指100",  "source": "us_etf", "symbol": "QQQ"},

    # ========== 全球指数 ==========
    {"name": "日经225",  "source": "global_sina", "symbol": "日经225指数"},

    # ========== 贵金属（用ETF替代） ==========
    {"name": "黄金(GLD)", "source": "us_etf", "symbol": "GLD"},
    {"name": "白银(SLV)", "source": "us_etf", "symbol": "SLV"},
]
