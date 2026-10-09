# templates/ — 用户自定义 PPT 母版

这个目录存放**用户上传的 `.pptx` 母版**。生成脚本会读取母版里的主题（`ppt/theme/theme1.xml`），
把母版的**配色方案和字体**套用到新生成的 PPT 上。

> 与另外两个目录的分工：
> - `assets/`  → 运行时功能资产（如 `round-card-template.html` 问卷引擎）
> - `references/` → 知识性规范文档（`.md`，供 agent 阅读）
> - `templates/`  → **用户的 `.pptx` 母版**（本目录）

---

## 1. 上传模板

把 `.pptx` 文件直接放进本目录即可：

```
templates/
├── README.md
├── registry.json
├── business-blue.pptx      ← 用户上传
└── tech-dark.pptx          ← 用户上传
```

约定：
- 只放 `.pptx`；`~$xxx.pptx` 之类的 Office 临时锁文件会被自动忽略。
- 文件名建议用英文小写加连字符（如 `business-blue.pptx`），便于命令行引用。

## 2. 登记元数据（可选但推荐）

编辑 `registry.json`，以**文件名**为键补充展示信息：

```json
{
  "business-blue.pptx": {
    "name": "商务蓝",
    "style": "深蓝配色、稳重",
    "scene": "商务汇报、年终总结",
    "desc": "公司标准汇报模板"
  }
}
```

字段说明：

| 字段 | 说明 | 缺省行为 |
|------|------|---------|
| `name` | 模板展示名（命令行 `--template` 可直接用这个名字） | 取文件名（去扩展名） |
| `style` | 风格关键词 | 空 |
| `scene` | 适合场景 | 空 |
| `desc` | 备注 | 空 |

> 不登记也能用，只是界面/命令行里显示的是文件名而非中文名。

## 3. 使用

### Node 路线（pptxgenjs）

```bash
# 列出所有可用模板
node scripts/generate.js --list-templates

# 按文件名引用
node scripts/generate.js --title "Q1 销售汇报" --template business-blue.pptx

# 按登记的中文名引用
node scripts/generate.js --title "Q1 销售汇报" --template 商务蓝

# 直接给绝对/相对路径（模板放在别处时）
node scripts/generate.js --title "Q1 销售汇报" --template ~/Downloads/my-deck.pptx
```

`--template` 优先级高于 `--palette`：给了模板就按模板主题配色，忽略 `--palette`。

### Python 路线（python-pptx，真正的母版继承）

四个 Python 脚本参数一致：

```bash
python3 scripts/generate_ppt.py  --list-templates
python3 scripts/generate_ppt.py  --template 商务蓝 -o ~/Desktop/demo.pptx
python3 scripts/generate_full.py --template 商务蓝 --keep-colors
```

| 参数 | 说明 |
|------|------|
| `-m, --template` | 模板名 / 文件名 / 路径 |
| `-o, --output` | 输出 `.pptx` 路径 |
| `-t, --title` | 标题（`generate_ppt.py` 使用） |
| `--keep-colors` | 继承母版/版式/字体，但**保留脚本内置配色** |
| `--list-templates` | 列出可用模板后退出 |

## 4. 实现边界（重要）

两条路线的能力**不同**：

| 能力 | Node（`pptxgenjs`） | Python（`python-pptx`） |
|------|--------------------|------------------------|
| 主题配色（`accent1-6`/`dk1`/`lt1`/`lt2`） | ✅ 提取后重绘 | ✅ 提取后重绘 |
| 主题字体 | ✅ 写入输出主题 | ✅ 继承 + 逐 run 指定 |
| 母版 `slideMaster` | ❌ | ✅ **整体继承** |
| 版式 `slideLayout`（11 个） | ❌ | ✅ **整体继承** |
| 母版自带幻灯片 | — | ✅ 生成前清空，不会串出原稿 |

原因：`pptxgenjs` 无法把已有 `.pptx` 当作「基底文件」再往里追加页面；而 `python-pptx`
的 `Presentation(template_path)` 会真正继承母版与版式。因此 Python 侧的做法是：

```python
prs = Presentation(template_path)   # 继承 slideMaster / slideLayout / theme
strip_slides(prs)                   # 清空母版自带幻灯片，避免污染输出
```

**仍未复用**的部分：脚本依旧用 `pick_blank_layout(prs)`（母版里的空白版式）作底，
再手工绘制内容，所以母版版式里的**占位符位置/背景图/装饰图形**不会被逐像素铺进去——
输出「携带」了母版的版式资源，但页面上不主动使用它们。

若要做到「把内容填进母版占位符」的逐版式复用，需要在脚本里改成
`prs.slides.add_slide(prs.slide_layouts[1])`（Title and Content）后按占位符写入，
目前尚未实现。

## 5. 解析依赖

- Node：主题解析用系统 `unzip`（macOS / Linux 自带），**不需要**额外 npm 依赖。
- Python：依赖 `python-pptx`（`pip install python-pptx`），主题解析用标准库 `zipfile`。

若模板无有效主题或找不到文件，两条路线都会给出明确报错并回退到内置配色。
