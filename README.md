# auto-koubei-collector

汽车之家车型口碑采集 skill。

用于从汽车之家车型口碑页批量采集全部口碑，并从详情页提取 `评价详情`、`最满意`、`最不满意` 后导出为结构化 Excel。

## 功能

- 支持汽车之家车型口碑分页抓取
- 默认按全部口碑列表抓取
- 支持自动探测总页数
- 支持从详情页提取全文正文，不抓评论区
- 支持从详情页拆出 `最满意` / `最不满意`
- 支持 `综合口碑` 优先按详情页 7 个维度分数求平均
- 支持导出单 sheet Excel
- 支持终端进度条、`progress.json` 轮询和飞书 webhook
- 导出结果旁自动生成同名 `.validation.json`

## 目录结构

```text
auto-koubei-collector/
├── SKILL.md
├── SPEC.md
├── CHANGELOG.md
├── README.md
├── references/
│   └── rules.md
├── scripts/
│   └── export_autohome_koubei.py
└── tests/
    └── test_export_autohome_koubei.py
```

## 使用示例

```bash
python3 scripts/export_autohome_koubei.py \
  --series-id 8140 \
  --start-page 1 \
  --auto-detect-pages \
  --output ./ZJ启源A06口碑_全量.xlsx \
  --workdir /Users/xyc/.openclaw/workspace
```

如果要看进度，可再加：

- `--progress`
- `--progress-file /tmp/job.progress.json`
- `--progress-webhook https://...`
- `--feishu-webhook https://...`
- `--feishu-secret <secret>`

## 导出结构

当前输出为单个 `口碑` sheet，每行一条口碑，包含：

- 基础结构字段
- `综合口碑`（优先按详情页 7 个维度分数求平均；分项缺失时回退到列表页值）
- `评价详情`
- `最满意`
- `最不满意`
- `来源链接`
- `抓取页码`

## 测试

单元测试：

```bash
python3 -m unittest tests/test_export_autohome_koubei.py -v
```

语法检查：

```bash
python3 -m py_compile scripts/export_autohome_koubei.py tests/test_export_autohome_koubei.py
```

烟测示例：

```bash
python3 scripts/export_autohome_koubei.py \
  --series-id 8140 \
  --start-page 1 \
  --end-page 2 \
  --output ./ZJ启源A06口碑_1-2页.xlsx \
  --workdir /Users/xyc/.openclaw/workspace
```

## 安装与 Release

Release 会附带打包好的 `.skill` 文件，可直接按你的 OpenClaw / 技能安装流程使用。

## 说明

更多规则请看：

- `SPEC.md`
- `references/rules.md`
