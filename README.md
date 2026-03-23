# auto-koubei-collector

汽车之家车型口碑采集 skill。

用于从汽车之家口碑页批量采集指定车型的“最满意 / 最不满意”评价，并导出为结构化 Excel。

## 功能

- 支持汽车之家车型口碑页抓取
- 支持 `最满意` / `最不满意` 两个维度
- 支持分页批量抓取
- 区分：
  - 车主购车口碑
  - 试驾探店口碑
- 按正式版结构导出 Excel
- 支持文件命名规则：
  - `ZJ+车型名称+最满意or最不满意_页数范围.xlsx`

## 目录结构

```text
auto-koubei-collector/
├── SKILL.md
├── SPEC.md
├── README.md
├── references/
│   └── rules.md
└── scripts/
    └── export_autohome_koubei.py
```

## 安装

### 方式 1：作为本地 skill 使用

把 skill 目录放到 OpenClaw skills 目录中，例如：

```bash
mkdir -p ~/.openclaw/workspace/skills
cp -R auto-koubei-collector ~/.openclaw/workspace/skills/
```

### 方式 2：使用打包后的 `.skill`

如果已有打包文件：

```bash
# 示例：将 .skill 文件放到本地后按你的 OpenClaw/技能管理流程安装
# 当前已打包文件示例：
# auto-koubei-collector.skill
```

## 测试

### 1. 脚本测试

运行导出脚本，抓取指定车型页码范围：

```bash
python3 scripts/export_autohome_koubei.py \
  --series-id 8208 \
  --start-page 1 \
  --end-page 13 \
  --output ./ZJ风云X3L最满意or最不满意_1-13页.xlsx \
  --workdir /Users/xyc/.openclaw/workspace
```

### 2. 已验证车型

- 风云X3L
- 风云T11

### 3. 测试重点

建议重点检查：

- 字段是否完整
- 用户名规范化是否正确
- 购车口碑 / 试驾探店口碑是否正确分流
- 最满意 / 最不满意两个维度是否符合业务预期
- Excel 命名和 sheet 结构是否正确

## 调用方式

### 自然语言唤起

可以直接这样说：

- 帮我抓取汽车之家某车型最满意和最不满意口碑，导出 Excel
- 抓风云X3L口碑页 1-13 页，分 sheet 导出
- 整理汽车之家口碑数据，区分购车口碑和试驾探店口碑
- 用汽车之家车型口碑采集这个 skill，把链接内容做成表格

### 脚本调用

```bash
python3 scripts/export_autohome_koubei.py \
  --series-id <车型ID> \
  --start-page <起始页> \
  --end-page <结束页> \
  --output ./ZJ车型名称最满意or最不满意_页数范围.xlsx \
  --workdir <workspace路径>
```

## 输出规则

### 文件命名

```text
ZJ+车型名称+最满意or最不满意_页数范围.xlsx
```

例如：

- `ZJ风云X3L最满意or最不满意_1-13页.xlsx`
- `ZJ风云T11最满意or最不满意_1-3页.xlsx`

### 默认 sheet 结构

正式版默认拆成 4 个 sheet：

- `最满意_车主购车口碑`
- `最满意_试驾探店口碑`
- `最不满意_车主购车口碑`
- `最不满意_试驾探店口碑`

## 字段

正式版字段：

1. 口碑维度
2. 数据类型
3. 用户名
4. 发表日期
5. 口碑标题
6. 综合口碑
7. 车型
8. 行驶里程
9. 电耗
10. 裸车购买价
11. 参考价格
12. 购买时间
13. 探店时间
14. 购买地点
15. 探店地点
16. 评价详情
17. 来源链接
18. 抓取页码

## 说明

更多规则请看：

- `SPEC.md`
- `references/rules.md`
