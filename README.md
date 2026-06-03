# baoyan-skill

> **商科保研文书专家** —— 一个面向中国高校经管商科推免（保研、夏令营、预推免）的 Claude Code Skill。专门处理**中文个人陈述（PS）**与**导师套瓷信**的生成、改写与优化。

## 项目背景

国内推免文书最常见的问题不是"写得不好看"，而是写得**像简历流水账、像通用 AI 输出、像所有人都能套的模板**。招生老师和意向导师真正想看到的是：扎实的学术基础、明确的研究方向、可信的潜力，以及与目标项目/导师之间自然、可信的衔接。

本 Skill 把这件事工程化：

- **不**做简历的散文化扩写
- **不**让用户填写僵硬的字段模板
- **不**用空泛抒情和形容词堆砌
- **要**把"学术基础 → 科研训练 → 实践支撑 → 未来规划"这条主线搭起来
- **要**让事实（GPA、数据库、研究方法、论文/项目）自己承担说服力
- **要**做"因专业而异"的写法适配，禁止换学校名就当新文书

## 适用方向

经管商科及其交叉方向：

- 经济学 / 金融 / 数字经济 / 碳经济
- 国际商务 / 工商管理
- 公共管理 / 区域经济
- 教育经济与管理 / 教育学
- 新闻传播 / 国际传播 / 智能传播
- 农林经济管理 / 农业与农村发展

## 能力范围

| 任务类型 | 输入 | 输出 |
| --- | --- | --- |
| 从零生成个人陈述 | 简历 / 个人材料 / 目标专业 | 完整 PS 草稿（带或不带小标题，可指定字数） |
| 改写优化已有 PS | 现有 PS + 目标专业 | 强化学术感、调整结构、压缩/扩写、改专业方向 |
| 从零生成套瓷信 | 简历 + 导师信息（主页/论文） | 正式、克制、学术导向的中文套瓷信 |
| 改写优化套瓷信 | 现有套瓷信 | 强化导师匹配段、压缩字数、提升正式度 |

## 仓库结构

```
baoyan-skill/
├── SKILL.md                      # 主 Skill 协议（流程、规则、输出格式）
├── build_statement.py            # 把结构化 PS 内容导出为符合规范的 .docx
├── references/                   # 写作规则参考
│   ├── structure-guide.md        # 结构搭建指南
│   ├── tone-guide.md             # 语气与措辞指南
│   └── writing-pitfalls.md       # 常见写作陷阱
├── examples/                     # 9 篇脱敏案例（按方向归类）
│   ├── ps-edu-econ-01.md
│   ├── ps-agri-econ-02.md
│   ├── ps-agri-management-03.md
│   ├── ps-rural-modernization-04.md
│   ├── ps-digital-carbon-05.md
│   ├── ps-education-06.md
│   ├── ps-public-management-07.md
│   ├── ps-international-communication-08.md
│   ├── ps-intelligent-communication-09.md
│   └── README.md                 # 案例使用规范
├── scripts/
│   └── convert_to_word.py
└── assets/
```

## 使用方式

### 作为 Claude Code Skill

1. 把整个目录放到 Claude Code 的 skills 目录（或工程内 `.claude/skills/`）
2. 在对话中提及 "保研"、"推免"、"个人陈述"、"PS"、"套瓷信" 等关键词，Claude 会自动加载本 Skill
3. 直接提供你的简历、已有文书或文件夹，无需逐项填表

### 生成 docx 文件

`build_statement.py` 提供了符合常见投递规范的 docx 模板（宋体 / Times New Roman、小四、1.25 倍行距、首行缩进 2 字符、上下左右 1 英寸页边距）：

```bash
pip install python-docx
# 编辑 build_statement.py 中的 SECTIONS / OUTPUT / AUTHOR
python build_statement.py
```

## 设计原则

1. **事实优先于辞藻** —— GPA、数据库、方法、样本量、论文题目，让信息密度本身有说服力
2. **不编造、不夸大** —— 弱事实必须克制表达，避免虚构导师评价、论文状态、研究结论
3. **专业适配是硬要求** —— 同一申请人申请不同专业，不能只换学校名
4. **Skill 管流程，检索补知识** —— 目标专业较新或较交叉时，先检索官方项目信息再写作
5. **案例是参照物，不是模板库** —— `examples/` 仅供结构、语气、专业迁移参考，禁止照抄事实

## 隐私

仓库内不包含任何真实申请人的个人材料。
- `examples/` 中的 9 篇案例均为脱敏后的结构参考样本
- `outputs/` 中的生成稿件为虚构人物、虚构经历的演示样本，不涉及真实申请人
- `assets/input-notes.md` 等本地输入材料已在 `.gitignore` 中排除
- `build_statement.py` 中的 `SECTIONS` 为占位示例，使用时请替换为自己的内容

## License

MIT
