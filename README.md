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
│   └── pitfalls.md           # 17 条真实踩坑手册（XML 倒序、信头错图、溢出补偿…）
├── scripts/
│   ├── analyze_template.py   # 模板解剖：dump 文本框/表格/背景/行高/run 结构
│   ├── extract_bios.py       # 简历提取：串场PPT批量 / docx·pptx 单份 → bios.json
│   ├── build_letters.py      # 生成引擎：config JSON → 整套 PPTX
│   └── qa_render.py          # 验收：LibreOffice 转 PDF + PyMuPDF 查空页
└── examples/
    ├── fuzhou_config.json    # 真实会议(20份)完整配置样例
    └── bios.example.json     # 简历库样例（虚构数据）
```

## 快速开始

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
