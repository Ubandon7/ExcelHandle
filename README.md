# ExcelHandle

本项目是一个本地优先的 Excel 文件处理网站。它支持多 Sheet 工作簿预览、嵌套 AND/OR 条件筛选、按优先级分流到多个结果 Sheet，以及对全部 Sheet 的同名表头或固定列位置求和。

## 已实现功能

- 支持 `.xlsx`、`.xls`、`.xlsm`、`.xlsb`、`.xltx`、`.xltm`、`.xlt`、SpreadsheetML `.xml`、`.csv`、`.tsv`。
- 自动识别所有 Sheet、表头、实际数据范围和重复表头。
- 支持同时上传并保留多个工作簿，可在左侧文件列表中切换当前处理文件。
- 支持以前缀开头、以后缀结尾、无前缀、无后缀、包含、不包含、为空、非空、属于、不属于。
- 以前缀/后缀筛选时，会按 `-` 自动识别当前列中的候选值并提供带数量的下拉选择；无前缀/无后缀可选择或输入标识符。
- 支持最多 5 层 AND/OR 条件组，同一表头可重复选择。
- 筛选条件可按结果 Sheet 名称持久化保存，并在筛选页下方形成可复用方案列表；点击方案可继续修改。
- 一个结果组生成一个 Sheet；结果组按顺序匹配，一行仅进入首个命中的结果组。
- 每组默认处理全部 Sheet，也可选择单个 Sheet。
- 全部 Sheet 模式采用联合表头，并增加“来源 Sheet”列；缺失字段留空。
- 未命中数据固定输出到最后一个“未匹配”Sheet。
- 求和支持按规范化表头或固定列位置匹配。
- 求和页会列出全部上传工作簿、已保存筛选结果和当前临时筛选结果；无需先保存筛选结果即可直接求和。
- `$1,234.50`、`12%`、`(100)` 会去除符号后解析；`-100` 按负数计算；含中英文字母的内容不会参与求和并保留在异常明细中。
- 处理结果可在线分页查看，也可下载为 `.xlsx`；筛选完成后先作为临时结果展示，由用户决定是否保存到处理记录。
- 左侧工作簿栏支持长文件名自动换行完整展示，并显示文件大小和 Sheet 数量。

## 一键启动

环境要求：Windows PowerShell、Python 3.11 或更高版本。仓库已包含构建后的网页；如需重新构建前端，还需要 Node.js 20 或更高版本。

```powershell
.\start.ps1
```

启动后打开 [http://127.0.0.1:8000](http://127.0.0.1:8000)。首次启动会自动创建 `.venv` 并安装 Python 依赖。

## 开发模式

后端：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r backend\requirements.txt
.\.venv\Scripts\python.exe -m uvicorn app.main:app --app-dir backend --reload
```

前端：

```powershell
cd frontend
npm install
npm run dev
```

开发服务器地址为 `http://127.0.0.1:5173`，`/api` 会代理到后端的 `8000` 端口。

## 测试与构建

```powershell
.\.venv\Scripts\python.exe -m pytest
cd frontend
npm run build
```

自动化测试覆盖筛选首次命中、嵌套 OR、联合表头、未匹配行、两种求和方式、符号数字解析、HTTP 完整流程和 Excel 导出。示例文件验收还会检查原始数据与参考结果均为 6991 行。

## 项目结构

```text
backend/app/        FastAPI、文件解析、筛选、求和、导出与临时存储
backend/tests/      核心逻辑和 API 集成测试
frontend/src/       React + TypeScript 用户界面
docs/               产品需求、技术设计和网页布局说明
prototype/          早期交互原型与设计预览
样例/               原始数据和参考筛选结果
```

运行数据默认写入 `backend/data`，包括上传文件、临时任务结果和下载文件。只有用户点击“保存处理记录”后，筛选结果才进入可恢复的历史列表，并持续保留到用户手动删除记录或删除所属文件；未保存结果在服务重启时清理。可通过环境变量 `EXCEL_HANDLE_DATA_DIR` 修改位置。上传文件不会执行宏；加密或损坏的工作簿需要先在 Excel 中解密或修复。

## API 文档

服务启动后可访问：

- Swagger UI：`http://127.0.0.1:8000/docs`
- OpenAPI JSON：`http://127.0.0.1:8000/openapi.json`

完整业务规则与接口设计见 [开发需求与技术设计](docs/Excel文件处理网站-开发需求与技术设计.md)，页面结构见 [网页布局与交互设计](docs/网页布局与交互设计.md)。
