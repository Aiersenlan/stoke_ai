import urllib.request
import json

def get_json(url):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    res = urllib.request.urlopen(req)
    return json.loads(res.read())

print("TWSE MI_INDEX:")
d1 = get_json("https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date=20260223&type=ALLBUT0999")
for tb in d1['tables']:
    if tb['title'] and '每日收盤行情' in tb['title']:
        fields = tb['fields']
        print(f"Code: {fields.index('證券代號')}, Name: {fields.index('證券名稱')}, Close: {fields.index('收盤價')}")
        print(f"Vol: {fields.index('成交股數')}, Val: {fields.index('成交金額')}")
        break

print("\nTPEX MI_INDEX:")
d2 = get_json("https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&d=115/02/23")
for i, f in enumerate(d2['tables'][0]['fields']):
    print(f"{i}: {f}")
