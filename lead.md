# Lead Scraper — Runbook

## 什么时候用这个

当你说「帮我 scrap [地点] 的 [类型] lead」，直接按下面步骤跑，不需要重新 coding。

---

## 前置条件

`.env` 里需要有：
```
GOOGLE_API_KEY=你的key
```

依赖已安装（第一次需要）：
```
py -m pip install googlemaps requests beautifulsoup4 openpyxl python-dotenv
```

---

## 每次用的流程

### 第一步：改 query

打开 `scrape_washington_leads.py`，找到这一行（在 `main()` 函数里）：

```python
query = "restaurants in Washington DC"
```

改成用户要的地点和类型，例如：
- `"restaurants in Seattle WA"`
- `"restaurants in New York NY"`
- `"cafes in Austin TX"`
- `"restaurants in Chicago IL"`

### 第二步：改输出文件名

同一个 `main()` 里，改这行：

```python
out = os.path.abspath("leads_output/washington_restaurants.xlsx")
```

改成对应的名字，例如：
- `"leads_output/seattle_restaurants.xlsx"`
- `"leads_output/newyork_restaurants.xlsx"`

### 第三步：改数量（可选，默认 10）

```python
places = search_places(gmaps, query, max_results=10)
```

把 `10` 改成想要的数量（最多约 60，Google Places API 限制）。

### 第四步：跑脚本

```
py scrape_washington_leads.py
```

---

## 输出

Excel 文件存在 `leads_output/` 文件夹，包含以下栏位：

| 栏位 | 说明 |
|------|------|
| Restaurant Name | 店家名字 |
| Website | 官网 |
| Instagram | IG 主页链接 |
| Facebook | FB 主页链接 |
| Email | 公开 email（没有就空白） |

---

## 注意事项

- Email 不一定每家都有，要看店家官网有没有公开。
- Facebook 有时会抓到 `facebook.com/tr`（追踪像素），这是假链接，可以手动删掉。
- Google Places API 每次调用会消耗 quota，留意用量。
- 脚本跑完之后记得 commit + push（CLAUDE.md 规定）。

---

## 快速改动示例

如果用户说「帮我找 10 个 Los Angeles 的餐厅 lead」：

1. `query = "restaurants in Los Angeles CA"`
2. `out = "leads_output/losangeles_restaurants.xlsx"`
3. `py scrape_washington_leads.py`
