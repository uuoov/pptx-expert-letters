# pptx-expert-letters 📄✉️

从 A4 信函式 PPTX 模板 + 会议议程，**批量生成整套专家沟通函**（主席 / 致辞 / 主持 / 讲者 / 讨论，一人一份），含简历简介页、程序审计与渲染验收。

这是一个 [ZCode](https://github.com/) / Claude Code 风格的 **Agent Skill**（也可当普通脚本工具集用）。诞生于医药学术会议会务的真实生产环境：一套 20~40 份的专家函，从模板解剖到全员交付，全流程脚本化。

## 它解决什么问题

会务场景里最常见的活儿：**"照上一场会议的沟通函，按新议程把这场会议的全套函做出来"**。手工做一套 20 份要改几百处，且极易漏改联动位置（换一位主持要牵动 5 个文件）。本技能把这件事变成：

```
模板解剖 → 议程任务矩阵 → 简历收集 → 配置驱动生成 → 程序审计 + 渲染验收 → 变更联动清单
```

## 目录

```
├── SKILL.md                  # Agent 工作流入口（触发词、五步流程、踩坑速查）
├── references/
│   ├── anatomy.md            # A4 信函模板解剖详解 + 版式参数表（行高 EMU 等）
│   ├── pitfalls.md           # 踩坑手册（XML 倒序、信头错图、溢出补偿、V2 实战…）
│   └── v2-station-playbook.md# V2 版式档案 + 站点复刻七步 + 二次审核清单
├── scripts/
│   ├── analyze_template.py   # 模板解剖：dump 文本框/表格/背景/行高/run 结构
│   ├── extract_bios.py       # 简历提取（旧版）：串场PPT批量 / docx·pptx 单份
│   ├── extract_bios_v2.py    # 简历提取（V2 目录式）：二进制转换/画布过滤/扫描件渲染转写
│   ├── build_letters.py      # 生成引擎（V1 旧版式）：config JSON → 整套 PPTX
│   ├── build_v2.py           # 生成引擎（V2 新版式）：站点 JSON → 整套 PPTX
│   ├── audit_v2.py           # V2 审计：校验点由站点配置自动派生
│   └── qa_render.py          # 验收：LibreOffice 转 PDF + PyMuPDF 查空页
└── examples/
    ├── fuzhou_config.json    # V1 真实会议(20份)完整配置样例
    ├── v2_chengdu.json       # V2 真实会议(32份)完整站点配置样例
    └── bios.example.json     # 简历库样例（虚构数据）
```

## 下载与安装

**你不需要会写代码，也不需要手动 clone——把一段话发给你的 Agent（ZCode / Claude Code 均可），剩下的它来做。**

### 第一步：对 Agent 说这段话（【】内换成你的实际信息）

```text
请帮我安装并使用 pptx-expert-letters 这个技能：

1. 把 https://github.com/uuoov/pptx-expert-letters.git 克隆到你的技能目录
   （ZCode 是 ~/.agents/skills/，Claude Code 是 ~/.claude/skills/）；
2. pip install python-pptx pymupdf pillow，并确认本机装有 LibreOffice（没有就提示我安装）；
3. 然后按该技能的 SKILL.md 流程，参照 examples/v2_chengdu.json 的配置写法，
   为我生成会议沟通函：
   - 捐体模板（上一场会议的沟通函）：【文件夹路径】
   - 本场议程：【飞书表格链接 / 截图 / 文字】
   - 专家简历目录：【文件夹路径】
   - 会议信息：【X月X日 星期X 09:00-12:00，XX酒店XX厅】
   - 讨论问题沿用样例里的套用规则，会议信息里如无特殊说明不放议程页；
4. 生成后按技能自带的 audit 和 qa_render 自检，并把这些事项列出来问我：
   需要我拍板的差异（如医院署名冲突）、有简历但不在议程的候补人员、
   简历被截断的人、两场讨论共用问题套是否可以。
```

### 第二步：回答 Agent 的确认问题

Agent 会在生成前把「人×任务矩阵」拿给你确认（谁是主持、每位嘉宾参与哪个环节），
并可能追问：某位专家的医院署名以简历还是议程为准、超长简历是否接受截断等。
逐项回复后它会继续执行到验收完成，最终交付一人一份的 PPTX 沟通函。

### 之后再次使用

技能装好后是一劳永逸的。下一场会议只需对 Agent 说：

> 「用 pptx-expert-letters，议程在【链接/截图】，简历在【路径】，地点【XX酒店XX厅】，
> 照上一场（【路径】）做全套沟通函」

### 手动安装（不用 Agent 时）

```bash
git clone https://github.com/uuoov/pptx-expert-letters.git
pip install python-pptx pymupdf pillow
```

环境要求：Python 3.8+；[LibreOffice](https://www.libreoffice.org/)（渲染验收用，默认路径可被自动识别）；
V2 版式需有一份上一场会议的沟通函作捐体模板（含真实专家信息，不入库，用自己经手的历史会议文件）。
所有脚本也可脱离 Agent 单独运行（见下方快速开始）。

## 快速开始

### V2 版式（2026 CSCO 巡讲系列：页级信头 / 主席双表 / 简介页带照片 / 逐拍任务表）

先用 `scripts/extract_bios_v2.py` 从简历目录提取（自动二进制转换、扫描件渲染转写），
再复制 `examples/v2_chengdu.json` 改成本站配置（议程/人员/地点/日期/讨论问题），
三条命令跑完全套：

```bash
python scripts/extract_bios_v2.py --res 简历目录/ --out bios.json --port portraits/ --conv converted/
python scripts/build_v2.py station.json      # 生成全套
python scripts/audit_v2.py  station.json     # 配置派生自动审计
python scripts/qa_render.py out/             # 渲染验收
```

版式细节、已定稿规范与二次审核清单见 `references/v2-station-playbook.md`。

### V1 旧版式（福州/杭州式）

```bash
# 1. 解剖一份旧会议的沟通函作模板（必做，产出全部版式参数）
python scripts/analyze_template.py 上一场/讨论函.pptx > anatomy.txt

# 2. 从串场PPT/散简历提取专家简历
python scripts/extract_bios.py --deck 串场.pptx --names 周某,李某,王某 --out bios.json
python scripts/extract_bios.py --file 某某简介.docx --name 陈某 --out bios.json --append

# 3. 照 examples/fuzhou_config.json 写本场配置（人×任务矩阵）
python scripts/build_letters.py config.json

# 4. 渲染验收（LibreOffice + PyMuPDF）
python scripts/qa_render.py out/
```

## 内置的实战经验

- **信头背景核对**：模板页级背景可能是别的会议的旧信头（真实事故：40/41 份函信头错图）
- **XML 防倒序**：重建段落 run 必须链式 `addnext`
- **固定高度文本框溢出补偿**：加备注行必须同步删空段，否则"任务环节"压到表格
- **简介页主题防截断**：长话题自动降字号（24pt→18pt）
- **邀请人交叉核对**：讲者函的邀请人必须与该环节主持函互相印证
- **换人联动清单**：换主持牵动 5 处，改配置重跑全量，禁止手补单文件
- **验收三件套**：程序审计 + LibreOffice 渲染 + 人工抽查（PowerPoint COM 常报 E_FAIL，别走那条路）

依赖：`python-pptx`、`PyMuPDF`、LibreOffice（仅验收用）。

## 数据安全

`examples/fuzhou_config.json` 中的会议议程信息来自公开会议日程；**专家简历（bios）已替换为虚构样例**——请勿把真实专家个人信息提交到公开仓库。

## License

MIT
