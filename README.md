# Hub i18n 词条提取与对比工具

从 Vue I18n 打包后的 `language-*.js` 中提取**中文 → 英文**词条，与标准 Excel 对照表比对，并导出便于人工校对的 Excel 报告。

适用于 Hub AI 等项目的界面文案版本核对与翻译一致性检查。

---

## 功能

- 解析 `language-*.js` 中的 `Xe`（英文）与 `Ke`（中文）语言包
- 生成路径级中英配对 JSON
- 与 `rule file` 中的标准 xlsx 自动对比（一致 / 不一致 / 仅 JS / 仅标准）
- 按标准表表头格式导出中英对照 Excel，**文件名可带时间戳**
- 通过 `config.json` 配置路径，换版本无需改代码

---

## 目录结构

```
.
├── config.json              # 路径与输出配置
├── extract_compare.py       # 步骤 1：提取 + 对比
├── export_js_to_excel.py    # 步骤 2：导出 Excel
├── i18n_config.py           # 配置模块
├── doc/
│   └── 操作文档.md          # 详细操作说明（中文）
├── new file/                # 输入：待提取的 language-*.js
├── rule file/               # 规范：标准对照 Excel
└── report/                  # 输出：json / 报告 / xlsx
```

| 目录 | 说明 |
|------|------|
| `new file` | 放置每次构建产出的 `language-xxxxx.js` |
| `rule file` | 放置标准翻译对照表（`.xlsx`） |
| `report` | 脚本生成结果，建议纳入 `.gitignore` 或仅提交示例 |

---

## 环境要求

- **Python** 3.10+
- **Node.js**（用于解析 JS 对象字面量）
- Python 包：`openpyxl`

```bash
pip install openpyxl
```

---

## 快速开始

### 1. 准备文件

1. 将 `language-xxxxx.js` 放入 `new file/`
2. 将标准对照表放入 `rule file/`（如 `GrandHub_AI_V1.1_新增中文词条.xlsx`）
3. 编辑 `config.json`：

```json
{
  "paths": {
    "js_dir": "new file",
    "js_file": "language-B1vEaFmS.js",
    "rule_dir": "rule file",
    "standard_xlsx": "GrandHub_AI_V1.1_新增中文词条.xlsx",
    "output_dir": "report"
  },
  "output": {
    "xlsx_add_timestamp": true,
    "timestamp_format": "%Y%m%d_%H%M%S"
  },
  "compare": {
    "enabled": true
  }
}
```

`js_file` 留空 `""` 时，将自动使用 `new file` 下**最新修改**的 `language-*.js`。

### 2. 运行

```bash
# 提取词条并与标准表对比
python extract_compare.py

# 导出带时间戳的 Excel
python export_js_to_excel.py
```

### 3. 查看输出（`report/`）

| 文件 | 说明 |
|------|------|
| `*_zh_en.json` | 全部中英词条（含 i18n 路径） |
| `*_compare_report.txt` | 与标准表的差异报告 |
| `*_中英对照_YYYYMMDD_HHMMSS.xlsx` | 两列中英对照表 |

---

## 更换 JS 版本

1. 新 JS 放入 `new file/`
2. 修改 `config.json` 中的 `js_file`
3. 重新执行上述两条命令

更完整的步骤、可选参数与排错说明见 **[doc/操作文档.md](doc/操作文档.md)**。

---

## 配置说明

| 配置项 | 默认值 | 说明 |
|--------|--------|------|
| `paths.js_dir` | `new file` | JS 输入目录 |
| `paths.js_file` | — | JS 文件名；空则自动选最新 |
| `paths.rule_dir` | `rule file` | 标准表目录 |
| `paths.standard_xlsx` | — | 标准表文件名 |
| `paths.output_dir` | `report` | 输出目录 |
| `output.xlsx_add_timestamp` | `true` | xlsx 是否追加时间戳 |
| `compare.enabled` | `true` | 是否与标准表对比 |

命令行可临时覆盖（详见操作文档）：

```bash
python extract_compare.py language-NewHash.js
python extract_compare.py -c /path/to/config.json
python export_js_to_excel.py language-NewHash.js
```

---

## 技术说明

- 目标 JS 需为 vue-i18n 打包格式，包含 `const Xe=`（英文）与 `Ke={global:`（中文）结构
- 对比以**中文键完全一致**为准；空格/括号差异会导致「仅 JS 有」与「仅标准有」分项，属预期行为
- `vxe` 等外部引用字段在解析时按空对象跳过

---

## 建议的 `.gitignore`

上传 GitHub 时，可按需忽略输入输出与缓存：

```gitignore
__pycache__/
*.pyc
report/*
!report/README.txt
new file/*.js
rule file/*.xlsx
_tmp_i18n_*
```

若希望仓库保留目录结构，可保留各目录下的 `README.txt`。

---

## 文档

- [操作文档（中文）](doc/操作文档.md)

---

## License

内部工具 / 按需自行添加开源协议。
