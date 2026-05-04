# Pain Point 分析流程手册 (pp.md)

## 任务目标

读取一份包含餐厅/咖啡馆名单的 Excel 文件，根据 ICP.md 的标准自动筛选 leads，并为符合条件的商家生成个人化 pain point 描述（中文），输出新的 Excel 文件。

---

## 直接运行（每次执行只需这一条命令）

```
py analyze_leads.py <输入文件.xlsx> <输出文件.xlsx>
```

**示例：**
```
py analyze_leads.py leads_output/mount_austin_cafe_restaurant.xlsx leads_output/mount_austin_qualified.xlsx
```

不传参数时默认使用上面的路径（用于快速测试）。

---

## 前置条件（首次设置，之后不用重复）

### 1. 安装依赖
```
py -m pip install requests openpyxl python-dotenv
```

### 2. 配置 `.env` 文件
项目根目录下已有 `.env`，内容为：
```
GEMINI_API_KEY=AIzaSy...（已配置好，无需更改）
```

### 3. 输入 Excel 格式要求
必须包含以下列（列名区分大小写）：
| 列名 | 必须 | 说明 |
|------|------|------|
| `Restaurant Name` | 是 | 商家名称 |
| `Instagram` | 建议 | IG 主页完整 URL |
| `Facebook` | 备用 | 没有 IG 时使用 FB |
| `Phone` / `Email` / `Website` | 可选 | 原样保留到输出 |

---

## 脚本工作原理

```
输入 Excel
    ↓
读取每一行 lead
    ↓
提取 IG/FB 链接 → 解析出用户名 (@handle)
    ↓
构造 prompt → 调用 Gemini 2.5 Flash + Google Search grounding
（Gemini 自动搜索该账号的粉丝数、发帖频率、内容质量等）
    ↓
ICP 筛选判断（符合 / 不符合）
    ↓
若符合 → 生成中文 pain point（2-4句，高度个人化）
    ↓
输出 Excel（保留原始列 + 新增 Pain Point 列）
```

---

## 关键配置参数（在 `analyze_leads.py` 顶部调整）

| 参数 | 当前值 | 说明 |
|------|--------|------|
| `RATE_LIMIT_DELAY` | `2` 秒 | 每次 Gemini 调用之间的间隔，100+ leads 可调高到 3-5 秒防止限速 |
| `maxOutputTokens` | `16384` | 每次 API 调用（即每个 lead）的最大 token 数。注意：这是单次调用上限，不是全部 leads 的总量。每个 lead 独立调用一次 Gemini，各自有 16384 token 空间，完全足够处理最长的中文 pain point。 |
| `temperature` | `0.4` | 生成创意度，0.4 平衡稳定与多样性，不建议改动 |

---

## 已知限制与解决方案

### Instagram 无法直接抓取
Instagram 对所有自动访问工具（requests、instaloader、headless browser）均返回 403 或要求登录。

**解决方案：** 脚本使用 Gemini + Google Search grounding——Gemini 通过搜索引擎查找账号公开信息（粉丝数通常出现在搜索结果 snippet 中），无需直接访问 IG。

### Gemini 模型版本
- 可用模型：`gemini-2.5-flash`（v1beta 端点）
- 不可用（已下线）：`gemini-1.5-flash`、`gemini-2.0-flash`
- 端点：`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent`

### Windows 控制台中文乱码
脚本内部已处理（cp1252 编码转换），不影响 Excel 文件输出。Excel 文件中文正常显示。

### JSON 解析失败（极少数情况）
脚本内置重试机制：若带 Google Search 的响应 JSON 解析失败，自动改用不带 search 的 Gemini 再试一次。

---

## 100+ Leads 注意事项

1. **运行时间估算：** 每个 lead 约 5-10 秒（含 API 调用 + rate limit delay），100 个 leads 约 10-17 分钟
2. **Rate limit：** 如遇 HTTP 429 错误，将 `RATE_LIMIT_DELAY` 从 2 调高到 5
3. **Gemini API 免费额度：** 注意 `gemini-2.5-flash` 每分钟请求数限制，大批量可考虑分批运行
4. **断点续传：** 脚本目前不支持断点续传；若中途中断，建议将输入 Excel 分割成多个小文件分批运行

---

## 输出文件说明

输出 Excel 包含原始所有列 + 新增 `Pain Point` 列：
- 只保留**符合 ICP 标准**的 leads（不符合的行被过滤掉）
- Pain Point 为 1-2 句简体中文，纯客观描述，不使用第二人称（不写「您/你/你们」）
- 表头行蓝色加粗，Pain Point 列自动换行

---

## ICP 筛选标准快速参考

**合格条件（满足任意一条即合格）：**
1. 粉丝 < 50,000
2. 发帖约每月一次或更少
3. 超过一个月没有新帖
4. Feed 以手机随拍为主（无后期、构图随意）
5. Feed 全是 Reels 且无食物海报（需粉丝 < 20,000）
6. 有食物海报但设计粗糙（Canva 套模板、字体乱）
7. 偶尔发食物内容但占比很低

**排除条件（满足任意一条即排除）：**
- 粉丝 ≥ 50,000
- Feed 以专业摄影为主
- 完全没有发过食物内容
- Feed 全是 Reels 且粉丝 ≥ 20,000
- 不是 F&B 商家（如培训学校、非餐饮业）
- 超过 6 个月没有发新帖（账号已实质性废弃）

---

## 文件结构

```
Claude Project/
├── analyze_leads.py          ← 主脚本，每次执行这个
├── pp.md                     ← 本流程手册
├── ICP.md                    ← ICP 标准完整版（脚本 prompt 基于此）
├── .env                      ← GEMINI_API_KEY（不提交 git）
└── leads_output/
    ├── <城市>_cafe_restaurant.xlsx     ← 输入：原始名单
    └── <城市>_qualified.xlsx          ← 输出：筛选后带 pain point
```

---

## 历史运行记录

| 日期 | 输入文件 | Leads 数 | 合格数 | 输出文件 |
|------|----------|----------|--------|----------|
| 2026-05-03 | mount_austin_cafe_restaurant.xlsx | 4 | 3 | mount_austin_qualified.xlsx |
| 2026-05-03 | johor_jaya_cafe.xlsx | 6 | 2 | johor_jaya_qualified.xlsx |
| 2026-05-03 | Wyoming list 1.xlsx | 59 | 15 | wyoming_qualified.xlsx |
| 2026-05-04 | Wyoming list 2.xlsx | rows 108–150 (43 rows) | 17 new (84 total) | wyoming2_qualified.xlsx |

---

*此文件由 Claude Code 生成并维护。每次成功运行后更新历史记录。*
