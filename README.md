# Cost Statistics

将多个受测对象的过程指标统计Excel汇总为简短的成本统计Word报告。包含十项指标对比表、平均总Token及耗时图、简短分析，默认封面加一至两页正文。支持同一对象的结果分散在多个文件或子目录，以及不同题量和同题多次运行。

## 自然语言安装

```text
请从 https://github.com/Ma-Ruinan/Cost-Statistics 安装 skills/cost-statistics 技能。
```

## 自然语言调用

```text
请使用 cost-statistics 技能，读取 <Excel所在目录>（例如：E:\evaluation\成本统计结果）及其子目录中的各对象过程指标统计Excel，生成成本统计报告，保存到 <结果目录>（例如：E:\evaluation\报告）。同一对象分散在多份Excel中的结果先合并并去重，再按共同题目比较，保留汇总表、总Token与耗时图及简短分析，适度关注MobileWork的实际表现。自动完成数据汇总、报告生成和排版检查。
```

Excel目录可按执行人员或数据集分设子目录，同一对象可提供多份包含 `_过程指标统计` 的Excel；也可直接给文件路径。对象名从表内标题取得，文件重命名不影响分组。只放入本次比较范围的结果。结果目录用于保存Word报告，未填写时由Agent在当前任务目录创建。对象名称、题目数和执行次数从Excel取得，不需要端侧工作空间或测试时间条件。

不同题目自动合并；同题不同次执行从逐次记录重新求均值。同一文件副本及会话ID、用户消息ID相同的执行只计一次，测试ID重新编号不会造成重复或遗漏。合并记录保留在本地汇总JSON，报告样式不变。

## 脚本入口

```powershell
python -m pip install -r skills/cost-statistics/requirements.txt
python skills/cost-statistics/scripts/aggregate.py --input-dir "<Excel所在目录>" --output "<工作目录>/汇总数据.json"
python skills/cost-statistics/scripts/report.py "<工作目录>/汇总数据.json" --analysis "<工作目录>/分析.json" --output "<结果目录>/成本统计报告.docx"
python skills/cost-statistics/scripts/render_preview.py "<结果目录>/成本统计报告.docx" --output "<工作目录>/预览"
```

Agent根据汇总数据撰写分析JSON，格式为 `[{"heading":"标题","text":"正文"}]`。不带 --analysis 可生成数据驱动的初稿观察。渲染需要本机Office导出能力或LibreOffice，以及Poppler的pdftoppm；字体可通过 --font 指定。完整规则见[SKILL.md](skills/cost-statistics/SKILL.md)。
