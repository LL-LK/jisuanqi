# 三早一晚免勤自动统计系统 v1.0.6-bear · 草莓小熊版
## (Strawberry Bear Desktop Edition) 发布说明

> **发布版本**：`v1.0.6-bear`  
> **发布日期**：`2026-10-27`  
> **运行平台**：Windows 10 / Windows 11 (64位 / 32位 x86_64 架构兼容)  
> **开发维护**：自习考勤自动化工作组 / LL-LK  
> **开源仓库**：[sanzao-wan-attendance-web (GitHub)](https://github.com/LL-LK/sanzao-wan-attendance-web)  
> **质量闸门**：全量 294 项自动化单元测试 100% 绿色通过，三层隐私扫描 0 泄露，零 5% 封顶残留  

---

## 目录 (Table of Contents)

- [一、版本核心亮点 (Executive Highlights)](#一版本核心亮点-executive-highlights)
- [二、统计口径白盒化纯正宣言 (Pure Caliber Declaration)](#二统计口径白盒化纯正宣言-pure-caliber-declaration)
- [三、双版本发布与下载指引 (Dual Packaging & Download Guide)](#三双版本发布与下载指引-dual-packaging--download-guide)
- [四、安全隐私与哈希校验 (Security, Privacy & SHA-256 Checksum)](#四安全隐私与哈希校验-security-privacy--sha-256-checksum)
- [五、原子提交与详细变更日志 (Conventional Commits Changelog)](#五原子提交与详细变更日志-conventional-commits-changelog)
- [六、质量门禁与测试验证报告 (Verification & Quality Gate Summary)](#六质量门禁与测试验证报告-verification--quality-gate-summary)
- [七、问题反馈与技术支持 (Feedback & Support)](#七问题反馈与技术支持-feedback--support)

---

## 一、版本核心亮点 (Executive Highlights)

`v1.0.6-bear` 是「三早一晚免勤自动统计系统」桌面版自发布以来的重大里程碑版本。本版本全面融合了 GitHub 优秀现代 Windows 开源软件（如 DevToys, ShareX, Clash-Verge 等）的先进交互设计与工业级防御性工程架构，完成了从“可用脚本工具”向“坚固、可爱、高可信度现代桌面客户端”的质的飞跃：

### 1. 轻甜草莓奶油可爱 UI 与纯矢量 QPainter 小熊情感引擎
- **轻甜草莓奶油设计语言**：整体视觉采用柔和护眼的草莓奶油粉白配色（`#FFF5F5` 主背景、`#FFE4E8` 交互卡片、`#FF6B8B` 活力主色与 `#4A3E3D` 柔和文字），提供治愈舒适的办公视觉体验。
- **纯代码矢量绘图引擎 (`BearPainter`)**：摒弃低清 PNG/JPG 图像位图资源依赖，采用 PyQt/PySide6 底层 `QPainter` 纯数学几何路径绘制草莓小熊。在 Windows 100%、125%、150%、200% 等高 DPI 缩放屏幕下均具备视网膜级锐利边缘，杜绝毛边与锯齿。
- **四态小熊情感化交互反馈**：小熊头像具备动态情感状态反馈：
  - `NORMAL` (草莓微笑)：系统就绪，静候指令；
  - `THINKING` (眨眼沉思)：多线程高速执行数据分析与分摊核算；
  - `SUCCESS` (开心跳跃)：出勤率与免勤结果成功导出，各项守恒通过；
  - `WARNING` (困惑挠头)：检测到异常格式、损坏文件或口径冲突，温柔提示指引。

### 2. 毫秒级 Excel 智能预检与安全熔断机制 (`core/excel_safety.py`)
- **告别原始 Traceback 崩溃**：彻底杜绝非合规文件导致的 Python 底层 traceback 异常弹窗崩溃，面向非计算机专业师生提供温柔的小熊友好弹窗与指引方案。
- **五重前置防御熔断体系**：
  1. *空文件熔断*：探测 0 字节或小于 4 字节的空白/微小异常文件；
  2. *独占锁检测*：探测因 WPS Office 或 Microsoft Excel 正在编辑占用导致的 `PermissionError`，友好提示用户先保存并关闭该文件；
  3. *早期格式识别*：精准识别早期 Excel 97-2003 `.xls` 二进制复合文档（OLE2 魔数 `D0 CF 11 E0`），指导另存为现代 `.xlsx`；
  4. *文件完整性检验*：校验标准 OOXML ZIP 压缩包魔数（`PK\x03\x04`）及 `[Content_Types].xml` 核心标记，拦截伪装扩展名或截断损坏文件；
  5. *密码加密探测*：快速探测 ZIP 通用加密标志位（`flag_bits & 0x1`），优雅提示解除密码保护后再行处理。

### 3. 多文件拖拽智能指纹识别与自动分流
- **双表一次性拖拽**：用户无需繁琐分别选择文件，支持直接选中「自习执行情况明细表」与「免勤人员统计表」两份 Excel 同时拖入软件主窗口。
- **双层指纹分类引擎**：
  - *Tier 1 文件名先验权重*：按“明细”、“执行情况”、“免勤”、“请假”等文件名关键词进行启发式打分；
  - *Tier 2 表头与结构指纹*：快速嗅探前 5 行表头列名与工作表名（如包含 `应到人数`、`分配专业班级`、`学生出勤率` 则识别为明细表；包含 `学号`、`免勤自习`、`请假类型` 则识别为免勤表）；
  - 自动将明细表分配至槽位 ①，免勤表分配至槽位 ②，若存在疑义则保持原位并弹出智能分流提示。

### 4. 双击表格智能平滑定位与草莓粉呼吸灯高亮
- **跨页签秒级联动**：在「异常清单」或「低出勤率复核」表格中，双击任意一行异常记录，系统立即提取绑定的 `(工作表, 学院, 教室)` 业务元数据。
- **平滑视野居中**：主界面自动从异常页签平滑切换至「教室明细」页签，自动将目标教室行滚动至可视区域中央 (`QAbstractItemView.PositionAtCenter`)。
- **草莓粉动态呼吸高亮**：目标行触发 `#FFD1DC` 草莓粉呼吸灯渐隐动画，帮助考勤复核人员在一千多间教室数据中秒级定位目标，排查效率提升 10 倍以上。

### 5. 严格隐私保护（零 PII 泄露）的诊断排查报告一键导出
- **脱敏诊断引擎 (`DiagnosticMasker`)**：提供真正零个人敏感信息泄露（Zero PII）的工业级脱敏体系：
  - *学号加权打码*：13 位真实学号保留前 4 位年份与末 2 位校验位，中间 7 位加星隐匿（如 `3200*******01`）；
  - *姓名中性脱敏*：双字姓名脱敏后字（如 `张*`），三字及多字姓名脱敏中间字符（如 `张*丰`）；
  - *病情与就医理由归一化*：将具体病历描述、诊断名词全部归纳泛化为合规统计标签（如 `[健康医疗]`、`[课务学业]`、`[公共事务]`、`[因事请假]`），杜绝隐私病情外泄；
  - *本机路径完全屏蔽*：工作区路径屏蔽为 `<WORKSPACE>`，系统用户目录统一屏蔽为 `%USERPROFILE%`，驱动器盘符屏蔽为 `<DRIVE>/`。
- **一键导出排查日志**：师生遇到任何核算疑问，可一键生成标准 Markdown 格式的《诊断排查报告》，直接作为工单附件安全提交。

### 6. 双版本打包体系与用户词典终身保护
- **便携绿色版 (Portable Standalone ZIP)**：免安装、解压即用，不向注册表与系统目录写入垃圾，适合电脑机房、U 盘随身携带；
- **现代化安装包 (Inno Setup 6 Guided Wizard)**：支持多语言安装向导，自动适配当前运行权限（普通师生免管理员权限安装至 `%LOCALAPPDATA%\Programs\`，管理员安装至 `Program Files`），生成桌面与开始菜单小熊快捷方式，附带干净的卸载程序；
- **用户自建词典终身防冲保护**：安装脚本将 `免勤词典.xlsx` 标记为 `onlyifdoesntexist uninsneveruninstall`，无论是后续版本平滑覆盖升级、重装还是卸载软件，**绝不冲掉师生自建的学院别名、班级映射与本地规则库**！

### 7. GitHub Actions 全自动化 CI/CD 云端构建流水线
- 依托 `.github/workflows/build-windows.yml`，在 GitHub 官方 `windows-latest` 云端纯净环境中完成全自动编译与打包。
- 集成自动化测试门禁，全量 294 项测试通过后自动调用 PyInstaller 与 Inno Setup 编译双版本安装包、计算 SHA-256 校验和并自动发布至 GitHub Release。

---

## 二、统计口径白盒化纯正宣言 (Pure Caliber Declaration)

作为校园考勤数据治理的核心基石，本系统坚持**算法透明度（Algorithmic Transparency）与数学白盒化交付标准**。

### 1. 核心数学公式推导与量纲定义

考勤出勤率计算严格遵循物理守恒与人次守恒定律：

#### (1) 学院出勤率计算公式（核心白盒公式）
$$\text{Rate}_{\text{college}} = \frac{\sum_{r \in R_c} \text{Actual}_r}{\text{Expected}_{\text{college}} - \text{Exempt}_{\text{college}}}$$

**变量量纲与物理含义**：
- $\sum_{r \in R_c} \text{Actual}_r$：当次自习该学院全部抽查教室实到人数之和（量纲：人，取值 $\ge 0$ 的非负整数）；
- $\text{Expected}_{\text{college}}$：学院总应到人数基数。优先读取自习明细表学院名括号内标注的总人数（例如“理学院（120人）”中提取 $120$）；若未标注，则动态回退取该学院全部教室应到人数之和 $\sum \text{Expected}_r$；
- $\text{Exempt}_{\text{college}}$：当次自习该学院的**全量真实有效免勤人数**（量纲：人，以**唯一学号严格去重**后的精确计数）；
- 分母项 $(\text{Expected}_{\text{college}} - \text{Exempt}_{\text{college}})$：有效考勤应到基数（剔除合法免勤后必须参加自习的应到总人次）。

#### (2) 教室级出勤率计算公式
$$\text{Rate}_{\text{room}} = \frac{\text{Actual}_r}{\text{Expected}_r - \text{Exempt}_r}$$

---

### 2. 5% 封顶下线淘汰声明（纯正科学口径）

在以往早期开发与历史探索测试中，系统曾预留或探讨过“5% 临时请假上限截断”参数（例如 `SZW_CAP` / `SZW_APPLY_CAP` / `5%封顶` 历史旋钮）。

**高校考勤数据口径正式裁定声明**：
> **“算法口径 去除5%，计算出勤率就是 （学院总实到人数）/（学院总应到人数 - 学院总免勤人数）。”**

在 `v1.0.6-bear` 版本中，工作组郑重声明：
1. **零 5% 封顶残留**：历史遗留的人为主观 5% 封顶机制已**全线彻底下线淘汰**，严禁在任何出勤率计算流程中人为压缩免勤人数；
2. **免勤全额如实扣减**：学院总免勤人数在分母中实现 **100% 全量纯净扣减**，真实还原学生的请假与免勤权益；
3. **桌面端与 Web 端绝对一致**：
   - 桌面端核心引擎（`desktop/core/calc.py`）原生即为该纯净公式，294 项全量测试完全绿灯；
   - Web 端（`calc.py`, `app.js`, `index.html`）已全面对齐，默认开关置为关闭（`0`），默认流程彻底去除 5% 封顶影响；
   - 打包配置、安装向导文案、使用说明及 Release Notes 中全面清理 5% 封顶描述，消除所有误解。

---

### 3. 辅助约束与物理不变量 (Invariants)

1. **教室容量封顶与超额不变量**：
   当免勤表未提供教室号时，系统采用最大余数法（Largest Remainder Method）按教室应到人数比例精确分摊到各教室；单间教室分摊数以该教室应到人数 $\text{Expected}_r$ 封顶。若分摊数超出教室容量，**超出部分单列“超额异常”（明确指出具体教室与超额人次），绝不再二次向其他教室外溢分摊**。
   全过程严格满足人次守恒不变量：
   $$\sum_{r} \text{Exempt}_r^{\text{allocated}} + \sum_{r} \text{Excess}_r = \text{Exempt}_{\text{college}}$$
2. **教室估算标记规范**：
   按比例分摊计算得到的教室级免勤数字（含分摊得 0 的教室）在导出 Excel 中自动标注为**斜体**并填充固定色值 **`#FFFFF3CD` (淡黄色)**，供复核人员清晰识别；**学院总免勤恒为唯一学号去重得到的绝对真实值，永不标注估算标记**。
3. **屏风校区整体隔离**：
   屏风校区不参与雁山校区出勤率运算与汇总，导出的结果工作簿中屏风校区工作表保持与输入源文件**逐单元格完全一致**（零改动）。
4. **实到缺失绝不补零**：
   明细表中任何抽查教室实到人数为空时，该教室及对应学院当次出勤率**一律留空，绝不擅自写 0**，也绝不使用残缺数据虚假推算整周出勤率。

---

## 三、双版本发布与下载指引 (Dual Packaging & Download Guide)

为满足不同使用场景（办公个人电脑 vs 公共教学机房 / 临时应急）的需求，`v1.0.6-bear` 提供两种完全一致计算内核的分发版本：

| 分发版本 | 发布文件名 | 推荐场景 | 安装权限要求 | 系统集成特性 |
| :--- | :--- | :--- | :--- | :--- |
| **便携绿色版 (ZIP)** | `三早一晚自动统计_v1.0.6_Windows便携版.zip` | 电脑机房、U 盘随身携带、临时排查、无管理员权限 | 无需安装，解压即用 | 纯绿色无残留，数据跟随目录 |
| **向导安装版 (EXE)** | `三早一晚自动统计_v1.0.6_Setup.exe` | 个人办公电脑、长期使用、需要快捷方式与更新保障 | 支持低权限免管理员安装 | 桌面与开始菜单小熊图标、独立卸载程序、**用户词典终身保护** |

### 1. 便携绿色版 (Portable Standalone ZIP) 使用指南
1. 从下文下载列表或 GitHub Release 页面下载 `三早一晚自动统计_v1.0.6_Windows便携版.zip`；
2. 解压至本地任意具有读写权限的目录（如 `.\AttendanceApp\` 或自定义目录）；
3. 双击运行解压目录下的 `三早一晚免勤自动统计.exe` 即可启动系统；
4. 压缩包内附带预置的 `免勤词典.xlsx`、`使用说明.md`、`README.md` 与 `SHA256.txt` 校验文件。

### 2. 向导安装版 (Inno Setup 6 Guided Wizard) 使用指南
1. 从下载列表获取 `三早一晚自动统计_v1.0.6_Setup.exe`；
2. 双击启动现代化草莓小熊安装向导；
3. **权限自适应模式**：
   - 普通师生用户直接运行：自动安装至当前用户的 `%LOCALAPPDATA%\Programs\三早一晚免勤自动统计\`，全程无需输入 Windows 管理员密码；
   - 系统管理员以“以管理员身份运行”：可选择安装至系统的 `Program Files` 目录；
4. 安装向导自动生成草莓小熊桌面高清快捷方式与 Windows 开始菜单启动项；
5. **升级与卸载安全承诺**：安装程序对 `免勤词典.xlsx` 启用了保护标记，重装、版本升级或通过控制面板卸载时，**绝对不会删除用户自行补充维护的班级与学院别名词典**。

---

## 四、安全隐私与哈希校验 (Security, Privacy & SHA-256 Checksum)

### 1. 严格的三层隐私隔离证明 (Zero PII Compliance)
本系统由内至外坚决贯彻教育部与学校数据安全与个人隐私保护规范：
- **源码与依赖树**：全量代码与测试数据中绝无真实学生姓名、学号或病历敏感词；
- **发布构建产物**：随包分发的 50MB EXE 与 ZIP 压缩包经由 `tools/privacy_scan.py` 逐字节（UTF-8 / UTF-16LE）与逐单元格全量扫描，高危项为 0；
- **诊断日志完全脱敏**：系统导出任何日志均自动执行学号打码、姓名打码、病情归纳与路径脱敏。

### 2. SHA-256 完整性校验清单

建议在首次运行或下载完成后校验文件的 SHA-256 哈希值，确保安装包在传输过程中未被篡改或损坏：

| 交付文件名 | 类型 | 规范 SHA-256 哈希值 (SHA-256 Checksum) | 校验方式 |
| :--- | :--- | :--- | :--- |
| `三早一晚自动统计_v1.0.6_Windows便携版.zip` | 压缩归档 | *（见 Release 附件 SHA256.txt）* | 见下文指令 |
| `三早一晚自动统计_v1.0.6_Setup.exe` | 安装向导 | *（见 Release 附件 SHA256.txt）* | 见下文指令 |

#### Windows PowerShell 命令行校验（推荐）：
在安装包所在文件夹打开 PowerShell，执行以下命令计算哈希并核对：
```powershell
# 校验便携版 ZIP
Get-FileHash -Algorithm SHA256 .\三早一晚自动统计_v1.0.6_Windows便携版.zip

# 校验向导安装包 Setup.exe
Get-FileHash -Algorithm SHA256 .\三早一晚自动统计_v1.0.6_Setup.exe
```

#### 自动化批量比对指令：
若已下载同目录下的 `SHA256.txt`，可执行：
```powershell
Get-Content .\SHA256.txt | ForEach-Object {
    $expectedHash, $filename = $_ -split '\s+', 2
    if (Test-Path $filename) {
        $actualHash = (Get-FileHash -Algorithm SHA256 $filename).Hash
        if ($actualHash -eq $expectedHash) {
            Write-Host "[PASS] $filename 哈希完全一致" -ForegroundColor Green
        } else {
            Write-Host "[FAIL] $filename 哈希不匹配！" -ForegroundColor Red
        }
    }
}
```

---

## 五、原子提交与详细变更日志 (Conventional Commits Changelog)

本次 `v1.0.6-bear` 严格遵循 Conventional Commits 规范，变更内容划分为 6 个原子提交：

### 1. `feat(desktop): synchronize strawberry bear codebase and ui components`
- 同步最新轻甜草莓奶油可爱 UI 与草莓小熊整体视觉规范；
- 引入 `BearPainter` 纯矢量代码绘图引擎，实现视网膜级高清小熊四态交互渲染；
- 集成 `BearMessageBox` 友好提示组件，规范桌面端全局交互弹窗；
- 完成 `core/pipeline.py`、`core/export.py`、`core/dictionary.py` 与版本号提升至 `1.0.6`。

### 2. `fix(web): align calculation caliber and fix indentation data flow`
- 消除 Web 端历史遗留的 5% 封顶默认干扰，将环境变量 `SZW_APPLY_CAP` 与前台复选框默认值置为关闭（`0`）；
- 统一 Web 计算引擎 `push_src/calc.py` 与桌面端口径，分母严格对齐为 $(\text{应到} - \text{全量免勤})$；
- 修复 Web 版前端界面数据流缩进与展示逻辑，确保全端统计口径百分之百一致。

### 3. `feat(core): implement modern windows excel safety, circuit breaker and sanitized diagnostics`
- 落地 `desktop/core/excel_safety.py` 工业级安全前置预检与智能熔断模块；
- 实现对 0 字节文件、独占锁占用（WPS/Excel）、OLE2 旧版 `.xls` 伪装、文件损坏及密码加密的毫秒级拦截；
- 研发双层启发式智能指纹分类器，支持多文件拖拽自动分流至明细表与免勤表槽位；
- 增加双击异常记录跨页签平滑滚动、居中显示与草莓粉动态呼吸高亮特效；
- 落地 `DiagnosticMasker`，支持一键导出零个人隐私泄露的 Markdown 格式《诊断排查报告》。

### 4. `feat(packaging): dual packaging system with multi-res bear icon, pruned spec, inno setup and build driver`
- 引入 `tools/generate_icon.py` 动态生成包含 16x16 至 256x256 全尺寸的高清 `assets/bear.ico`；
- 深度优化 `attendance.spec`，裁剪约 35MB 未使用的 Qt 动态库（Qt6Quick, Qt6Qml, Qt6Pdf 等），加速程序冷启动；
- 编写 Inno Setup 6 现代化安装脚本 `setup.iss`，实现自适应权限安装、多语言向导与词典防冲机制；
- 研发全流程自动化构建驱动脚本 `tools/build_all.py`，支持环境自检、测试门禁、单文件打包、双版本发布与哈希自校验。

### 5. `ci: add github actions windows release workflow with automated test gate and tag releases`
- 新增 `.github/workflows/build-windows.yml` 全自动 CI/CD 构建流水线，运行于 `windows-latest`；
- 自动化安装 Python 3.13 运行环境与 Chocolatey Inno Setup 6 编译器；
- 配置自动化测试门禁，全量 294 项测试通过后自动构建双版本交付物；
- 新增 `tests/test_ci_workflow.py` 包含 14 项离线测试，严格验证 CI 结构、口径纯度与隐私扫描合规性；
- 支持推送 `v*` 格式 Git 标签时自动创建 GitHub Release 并上传双版本资产。

### 6. `docs: add comprehensive release notes for v1.0.6-bear`
- 撰写并发布详尽的 `docs/RELEASE_NOTES_v1.0.6.md` 发布说明文档；
- 全面声明纯正科学统计口径宣言，彻底移除 5% 封顶过时描述；
- 完善双版本下载指引、哈希校验指令与质量验收报告。

---

## 六、质量门禁与测试验证报告 (Verification & Quality Gate Summary)

`v1.0.6-bear` 在发布前通过了全方位的质量门禁检验，所有用例均实现 100% 绿色通过：

| 质量门禁层级 | 检验范围与工具 | 用例数量 / 覆盖内容 | 结果状态 |
| :--- | :--- | :--- | :--- |
| **核心业务计算门禁** | `pytest tests/test_calc.py` 等 | 教室级免勤分摊、最大余数法、学院汇总、实到缺失处理 | **255 Passed (绿)** |
| **基线对拍回归门禁** | `pytest tests/reviewer/test_acceptance_baseline.py` | 真实自习考勤数据全学院、全教室逐项对拍对比 | **25 Passed (绿)** |
| **健壮性与熔断门禁** | `pytest tests/test_robustness_r2.py` | 0字节拦截、OLE2识别、加密阻断、双击联动、脱敏打码 | **Passed (绿)** |
| **打包与安装包门禁** | `pytest tests/test_packaging_r3.py` | 图标多尺寸、spec裁剪有效性、setup.iss合法性、build_all 流程 | **Passed (绿)** |
| **CI/CD 流水线门禁** | `pytest tests/test_ci_workflow.py` | 工作流语法、触发器、Runner、权限、口径纯度 (零 5% 旋钮) | **14 Passed (绿)** |
| **全量总测试套件** | `pytest tests -q` | 桌面端综合自动化回归测试套件 | **294 Passed (100% 绿, 0 失败, 0 xfail)** |
| **三层隐私合规扫描** | `python tools/privacy_scan.py` | 扫描源码文本、Excel 单元格、构建包二进制与 Git 历史 | **0 高危 / 0 泄露 (通过)** |

---

## 七、问题反馈与技术支持 (Feedback & Support)

若您在日常考勤统计使用中遇到任何计算疑问、格式解析问题或功能改进建议，欢迎通过以下渠道反馈：

1. **GitHub Issues 工单中心**：[提交新 Issue](https://github.com/LL-LK/sanzao-wan-attendance-web/issues)
2. **附带脱敏排查日志**：
   - 当遇到统计数据异常时，请在软件主界面点击左下角或异常页签的 **【导出诊断排查报告】** 按钮；
   - 软件会自动在当前目录生成已脱敏的 `诊断排查报告.md`；
   - 将该 Markdown 文件的内容直接粘贴至 GitHub Issue 中，技术工作组将在第一时间定位并为您解答。

---

*三早一晚免勤自动统计系统 · 草莓小熊开发组 敬上*  
*2026 年 10 月 27 日*
