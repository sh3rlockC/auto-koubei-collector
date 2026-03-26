# auto-koubei-collector

汽车之家车型口碑采集 skill。

用于从汽车之家口碑页批量采集指定车型的“最满意 / 最不满意”评价，并导出为**按来源链接严格对齐**的结构化 Excel。

## 当前正式输出

- 只导出 **2 个 sheet**：
  - `购车口碑`
  - `试驾口碑`
- 每一行代表 **同一个来源链接** 的同一条口碑
- 同时包含两列：
  - `最满意`
  - `最不满意`
- 仅当同一来源链接下 **最满意 / 最不满意都有内容** 时，才进入正式结果
- 导出后自动生成同名 `.validation.json` 校验报告
- 支持 `--strict-validate`，校验失败时直接返回非零退出码

## 功能

- 支持汽车之家车型口碑页抓取
- 支持 `最满意` / `最不满意` 两个维度
- 支持分页批量抓取
- 按来源链接对齐同一条口碑
- 区分：
  - 车主购车口碑
  - 试驾探店口碑
- 导出为双 sheet 对齐版 Excel
- 导出后自动生成同名 `.validation.json` 校验报告
- 支持文件命名规则：
  - `ZJ+车型名称+口碑对齐版_日期.xlsx`

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

## 测试

### 脚本测试

```bash
python3 scripts/export_autohome_koubei.py \
  --series-id 8208 \
  --start-page 1 \
  --end-page 13 \
  --output ./ZJ风云X3L口碑对齐版_2026-03-26.xlsx \
  --workdir /Users/xyc/.openclaw/workspace
```

### 严格校验

```bash
python3 scripts/export_autohome_koubei.py \
  --series-id 8208 \
  --start-page 1 \
  --end-page 13 \
  --output ./ZJ风云X3L口碑对齐版_2026-03-26.xlsx \
  --workdir /Users/xyc/.openclaw/workspace \
  --strict-validate
```

## 输出字段

1. 数据类型
2. 用户名
3. 发表日期
4. 口碑标题
5. 综合口碑
6. 车型
7. 行驶里程
8. 电耗
9. 裸车购买价
10. 参考价格
11. 购买时间
12. 探店时间
13. 购买地点
14. 探店地点
15. 最满意
16. 最不满意
17. 来源链接
18. 最满意抓取页码
19. 最不满意抓取页码

## 校验逻辑

校验报告会包含：

- 原始抓取条数（最满意 / 最不满意）
- 对齐后正式结果条数
- 两个 sheet 的数量
- 每页解析到的卡片数量
- 未对齐链接清单
- 异常记录清单

## 说明

更多规则请看：

- `SPEC.md`
- `references/rules.md`
