# grill-ppt

**先拷问，再产出的 PPT 生成 Skill。** 一个毫不留情的提问者 + 一位专业的 PPT 设计师，二合一。

- **第一阶段 · 穷追式提问**：把这份 PPT 的每个设计决策（主题 / 目的 / 受众 / 页数 / 语言 / 风格 / 模板 / 配色 / 结构 / 数据来源 / 交付形态）映射成一棵**设计树**，按轮次追问，直到前沿为空、双方达成共识。
- **第二阶段 · 生成**：共识达成后再联网检索内容 → 定配色/模板 → 设计结构 → 生成 `.pptx` → 视觉检查。

> 铁律：在用户明确确认「我们已达成共识」之前，绝不开始生成。

支持所有行业（商务、教育、科技、医疗、金融、创意等）与中英文双语。

---

## 安装为 Cursor Skill

把本目录整个放进 Cursor 的 skills 目录，`SKILL.md` 即入口：

```bash
# 个人级（所有项目可用）
git clone <repo-url> ~/.cursor/skills/grill-ppt

# 或项目级（仅当前项目）
git clone <repo-url> .cursor/skills/grill-ppt
```

放好后，用触发词即可唤起：`grill-ppt`、`做PPT`、`生成PPT`、`帮我做个PPT`、`make ppt`、`create presentation` 等。

---

## 目录结构

```
grill-ppt/
├── SKILL.md                        # 技能主入口：角色定义 + 两阶段流程 + 设计规范
├── README.md                       # 本文件
├── assets/
│   └── round-card-template.html    # 单轮问卷引擎（data-driven，复制只改数据）
├── references/                     # 知识性规范文档，供 agent 阅读
│   ├── 配色方案.md
│   ├── MBE插画风格规范.md
│   ├── 复古卡通风格规范.md
│   ├── 孟菲斯设计风格规范.md
│   ├── 9套新增风格规范.md
│   ├── 7套新增风格规范.md
│   └── …                           # 其余单独风格规范
├── scripts/                        # PPT 生成脚本（Node + Python 两条路线）
│   ├── generate.js                 # Node / pptxgenjs 主线
│   ├── template-loader.js          # 模板扫描 / 主题解析（generate.js 依赖）
│   ├── generate_ppt.py             # Python / python-pptx（支持母版继承）
│   ├── generate_full.py
│   ├── generate_grand.py
│   ├── generate_mbe.py
│   └── pptx_template.py            # 母版继承 + --list-templates 公共模块
└── templates/                      # 用户的 .pptx 母版（见 templates/README.md）
    ├── README.md
    └── registry.json               # 模板元数据登记（name/style/scene/desc）
```

---

## 工作流

### 第一阶段 · 穷追式提问

1. 把设计决策映射成一棵**设计树**，按**轮次**推进。
2. 每一轮的**前沿（frontier）**= 所有前置条件已确定、现在就能问的决策。
3. 一轮 = 一份**独立问卷界面**（`show_widget`），先把整轮内容在内部构建完毕，再**一次性原子推送**，绝不边渲染边生成。
   - 呈现前先调 `show_widget_guide`，再调 `show_widget`（`interactive=true`）。
   - 问卷引擎直接复用 `assets/round-card-template.html`，**只改**脚本里标注「只改这里」的 `ROUND` 与 `QUESTIONS`。
   - 每题给出**推荐答案并默认选中**，让用户「提交即等于接受推荐」；同时每题保留「其他」输入框（与选项严格互斥）。
   - 无界面渠道或渲染失败时，回退到纯文字提问（兜底，非默认）。
4. 用户每轮回答都**重塑**这棵树：确定项把前沿向外推进并解锁下游决策，直到前沿为空。

典型第 1 轮前沿：主题与目的、受众、语言、交付形态、页数与时长、风格基调、模板/配色、硬约束与数据来源。

> 事实由 agent 负责查（`web_search` / `task` 派子代理），**决策交给用户**。

### 第二阶段 · 生成 PPT

前置条件：前沿为空且用户确认达成共识。

0. **回放共识** → 输出一段「设计简报」。
1. **AI 内容检索与填充**：识别内容缺口，联网检索市场数据/趋势/案例，每页正文补到 80–150 字。
2. **确定模板或配色**：见下方两条路线。
3. **设计结构**：每页 3–5 个要点、有数据支撑、布局多样。
4. **生成 `.pptx`**：调用 `scripts/` 下脚本。
5. **视觉检查**：查文字溢出、配色统一、布局整齐、内容完整。

---

## 两条生成路线

| 能力 | Node（`generate.js`） | Python（`generate_*.py`） |
|------|----------------------|--------------------------|
| 主题配色 + 对比度保护 | ✅ | ✅ |
| 主题字体 | ✅ | ✅ |
| 母版 `slideMaster` / 版式 `slideLayout` | ❌ | ✅ **整体继承** |
| 母版自带幻灯片 | — | ✅ 生成前清空，不串原稿 |

- **Node**：`pptxgenjs`，只能**提取主题**（无法以已有 `.pptx` 为基底追加页面）。
- **Python**：`Presentation(template_path)` + `strip_slides()`，输出**携带**母版/版式/主题；加 `--keep-colors` 可只继承母版/版式/字体、保留脚本内置配色。

详见 `templates/README.md`。

---

## 使用方法

### 常用命令

| 命令 | 说明 |
|------|------|
| `grill-ppt [主题]` / `做PPT [主题]` | 先进入拷问，达成共识后生成 PPT |
| `用模板[编号] 做PPT [主题]` | 指定模板后拷问并生成 |
| `查看模板` | 列出所有可用模板（`--list-templates`） |
| `修改第[N]页` / `换个配色` / `预览PPT` | 迭代已生成的 PPT |

### 脚本

```bash
# 列出可用模板
node scripts/generate.js --list-templates
python3 scripts/generate_ppt.py --list-templates

# Node：按配色或模板生成
node scripts/generate.js --topic "Q1 销售汇报" --pages 10 --lang zh
node scripts/generate.js --topic "Q1 销售汇报" --template business-blue.pptx --output ~/Desktop/Q1销售汇报.pptx

# Python：套用用户模板（真正的母版继承）
python3 scripts/generate_ppt.py --template 商务蓝 -o ~/Desktop/Q1销售汇报.pptx
python3 scripts/generate_full.py --template 商务蓝 --keep-colors

# 其他风格脚本
python3 scripts/generate_grand.py   # 大气版
python3 scripts/generate_mbe.py     # MBE 插画
```

`generate.js` 参数：`-t/--title/--topic`、`-p/--palette`、`-m/--template`、`-l/--lang`、`-o/--output`、`--list-templates`。
Python 脚本参数统一：`-m/--template`、`-o/--output`、`-t/--title`、`--keep-colors`、`--list-templates`。

`--template` 优先级高于 `--palette`。

---

## 模板库

> 用户上传的 `.pptx` 母版放在 `templates/`，元数据登记在 `templates/registry.json`。
> **以实际文件为准**：新增/删除模板后运行 `--list-templates` 同步 `SKILL.md` 中的模板表。

添加模板：把 `.pptx` 放进 `templates/`，在 `registry.json` 里以**文件名**为键登记 `{ "name": 中文名, "style": 风格, "scene": 适合场景, "desc": 备注 }`。
完整说明见 `templates/README.md`。

---

## 依赖安装

```bash
# Node.js
npm install -g pptxgenjs

# Python（生成 + 母版继承）
pip install python-pptx

# Python（预览用）
pip install "markitdown[pptx]" Pillow

# LibreOffice（PPT 转图片）
brew install libreoffice   # macOS
```

---

## 设计规范速览

- **排版**：页面标题 36–44pt 粗体、章节标题 20–24pt、正文 14–16pt、注释 10–12pt 灰。
- **布局**：左文右图 / 图标行列 / 大数据展示 / 时间轴 / 2×2 矩阵 / 全图背景 / 对比列。
- **禁止**：纯文字页、标题下划线、默认蓝色（非商务）、同一布局重复 >3 次、正文 <14pt、对比度不足。

完整内容见 `SKILL.md`，配色与风格细节见 `references/`。
