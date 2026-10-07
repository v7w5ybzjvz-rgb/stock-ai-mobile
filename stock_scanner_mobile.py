import requests
import pandas as pd

print("=== 台股全市場掃描器 ===")

# 取得上市資料
print("正在抓取上市股票...")
twse_url = "https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL"
twse_data = requests.get(twse_url, timeout=20).json()
twse = pd.DataFrame(twse_data)
twse["市場"] = "上市"

# 取得上櫃資料
print("正在抓取上櫃股票...")
tpex_url = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_daily_close_quotes"
tpex_data = requests.get(tpex_url, timeout=20).json()
tpex = pd.DataFrame(tpex_data)
tpex["市場"] = "上櫃"

print("原始上市筆數：", len(twse))
print("原始上櫃筆數：", len(tpex))

# 只保留4位數股票代號
twse_code = twse.columns[1]
tpex_code = tpex.columns[1]

twse = twse[twse[twse_code].astype(str).str.fullmatch(r"[0-9]{4}", na=False)]
tpex = tpex[tpex[tpex_code].astype(str).str.fullmatch(r"[0-9]{4}", na=False)]

print("篩選後上市股票：", len(twse))
print("篩選後上櫃股票：", len(tpex))

# 儲存結果
twse.to_excel("上市普通股.xlsx", index=False)
tpex.to_excel("上櫃普通股.xlsx", index=False)

print("完成！")
print("已建立：上市普通股.xlsx")
print("已建立：上櫃普通股.xlsx")

# ===== 第一版自動選股 =====
print("\n開始自動選股...")

# 合併上市 + 上櫃
all_stocks = pd.concat([twse, tpex], ignore_index=True)

# 找出共同欄位：股票代號、名稱
code_col = all_stocks.columns[1]
name_col = all_stocks.columns[2]

# 顯示股票總數
print("普通股總數：", len(all_stocks))

# 先建立候選名單
candidates = all_stocks.copy()

# 儲存候選名單
candidates.to_excel("選股候選名單.xlsx", index=False)

print("第一版選股完成！")
print("候選股票：", len(candidates))
print("已建立：選股候選名單.xlsx")

# ===== 第二版：成交量篩選 =====
print("\n開始篩選成交量...")

# TradeVolume 轉成數字
candidates["TradeVolume"] = pd.to_numeric(
    candidates["TradeVolume"], errors="coerce"
)

# 保留當日成交量至少 100,000 股
candidates = candidates[candidates["TradeVolume"] >= 100000].copy()

# 成交量由大到小排列
candidates = candidates.sort_values(
    by="TradeVolume",
    ascending=False
)

print("成交量篩選後：", len(candidates), "檔")

# 另存結果
candidates.to_excel("成交量篩選.xlsx", index=False)

print("已建立：成交量篩選.xlsx")
# 第三版：股價篩選
print("開始股價篩選...")

price_col = "ClosingPrice"

candidates[price_col] = pd.to_numeric(
    candidates[price_col], errors="coerce"
)

candidates = candidates[
    (candidates[price_col] >= 10) &
    (candidates[price_col] <= 500)
].copy()

print("股價篩選後：", len(candidates), "檔")

candidates.to_excel("股價篩選.xlsx", index=False)

print("已建立：股價篩選.xlsx")

# 第四關：漲跌幅篩選
print("\n開始篩選漲跌幅...")

# 把漲跌數字轉成數值
candidates["Change"] = pd.to_numeric(
    candidates["Change"], errors="coerce"
)

# 保留今日上漲的股票
candidates = candidates[
    candidates["Change"] > 0
].copy()

print("上漲股票篩選後：", len(candidates), "檔")

# 儲存結果
candidates.to_excel("上漲股票篩選.xlsx", index=False)

print("已建立：上漲股票篩選.xlsx")

# 第五關：強勢漲幅篩選
print("\n開始篩選強勢漲幅...")

# 開盤價轉成數字
candidates["OpeningPrice"] = pd.to_numeric(
    candidates["OpeningPrice"], errors="coerce"
)

# 計算 開盤 -> 收盤 漲幅百分比
candidates["漲幅%"] = (
    (candidates["ClosingPrice"] - candidates["OpeningPrice"])
    / candidates["OpeningPrice"] * 100
)

# 留下漲幅至少 2% 的股票
candidates = candidates[
    candidates["漲幅%"] >= 2
].copy()

# 漲幅由大到小排列
candidates = candidates.sort_values(
    by="漲幅%",
    ascending=False
)

print("強勢股票篩選後：", len(candidates), "檔")

candidates.to_excel("強勢股票篩選.xlsx", index=False)

print("已建立：強勢股票篩選.xlsx")

# 第六關：強勢股排名
print("\n開始強勢股排名...")

# 漲幅排名：越高越好
candidates["漲幅排名"] = candidates["漲幅%"].rank(
    ascending=False,
    method="min"
)

# 成交量排名：越大越好
candidates["成交量排名"] = candidates["TradeVolume"].rank(
    ascending=False,
    method="min"
)

# 綜合分數：數字越小代表排名越前
candidates["綜合排名分數"] = (
    candidates["漲幅排名"] +
    candidates["成交量排名"]
)

# 按綜合排名排列
candidates = candidates.sort_values(
    by="綜合排名分數",
    ascending=True
)

# 只留下前 30 名
candidates = candidates.head(30).copy()

print("最終強勢候選股：", len(candidates), "檔")

# 儲存結果
candidates.to_excel("強勢股TOP30.xlsx", index=False)

print("已建立：強勢股TOP30.xlsx")

# 第七關：TOP 10 綜合評分
print("\n開始計算綜合分數...")

# 漲幅排名：越高分數越高
candidates["漲幅分數"] = candidates["Change"].rank(
    ascending=False,
    method="min"
)

# 成交量排名：越大分數越高
candidates["成交量分數"] = candidates["TradeVolume"].rank(
    ascending=False,
    method="min"
)

# 綜合排名
candidates["綜合分數"] = (
    candidates["漲幅分數"] +
    candidates["成交量分數"]
)

# 分數越小代表排名越前面
candidates = candidates.sort_values(
    by="綜合分數",
    ascending=True
)

# 只留前10名
top10 = candidates.head(10).copy()

print("最終 TOP10：", len(top10), "檔")

top10.to_excel("強勢股TOP10.xlsx", index=False)

print("已建立：強勢股TOP10.xlsx")

# 第八關：建立最終排名
print("\n開始建立最終排名...")

# 依照綜合分數重新排序
top10 = top10.sort_values(
    by="綜合分數",
    ascending=True
).copy()

# 加入第1～10名
top10.insert(0, "最終排名", range(1, len(top10) + 1))

# 重新輸出 Excel
top10.to_excel("強勢股TOP10排名版.xlsx", index=False)

print("完成！")
print("已建立：強勢股TOP10排名版.xlsx")

# 第八關：精選 TOP3
print("\n開始挑選最終 TOP3...")

top3 = candidates.head(3).copy()

# 加上最終排名
top3.insert(0, "最終排名", range(1, len(top3) + 1))

# 儲存 TOP3
top3.to_excel("今日精選TOP3.xlsx", index=False)

print("今日精選 TOP3：", len(top3), "檔")
print(top3[["最終排名", "Code", "Name", "ClosingPrice", "Change"]])

print("已建立：今日精選TOP3.xlsx")

# 第九關：計算收盤強度

top10["ClosingPrice"] = pd.to_numeric(top10["ClosingPrice"], errors="coerce")
top10["HighestPrice"] = pd.to_numeric(top10["HighestPrice"], errors="coerce")
top10["LowestPrice"] = pd.to_numeric(top10["LowestPrice"], errors="coerce")
top10["收盤強度"] = (top10["ClosingPrice"] - top10["LowestPrice"]) / (top10["HighestPrice"] - top10["LowestPrice"]) * 100
top10 = top10.sort_values(by="收盤強度", ascending=False).copy()

top10["最終排名"] = range(1, len(top10) + 1)
top10.to_excel("強勢股TOP10收盤強度.xlsx", index=False)




























    

