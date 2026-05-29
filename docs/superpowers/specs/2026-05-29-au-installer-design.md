# AU 安装器设计 (2026-05-29)

## 1. 背景与目标

母项目 `shein-extract` 已打包为 Windows 安装程序 `SheinExtract.exe`,但创建了**两个图标**(主程序图标 + 单独的"配置"图标 `--config`)。

把 `shein-extract-au` 同样打包成 Windows 安装程序,**简化为一个图标**:启动后用两项菜单引导用户选择动作。

## 2. 决策记录(用户已确认)

| 决策 | 选择 |
|---|---|
| 复用母项目的密码门、更新检查、AI 标题(Claude Haiku) | ✅ 三项全要 |
| 菜单 UX | ✅ 控制台输入(沿用母项目"黑窗口 + pause"风格) |
| 首次运行行为 | ✅ 总是出菜单(选 [2] 但无配置时提示先选 [1]) |
| 入口数量 | ✅ 1 个 exe、1 个图标 |

## 3. 命名与标识

| 项 | 值 | 备注 |
|---|---|---|
| 中文显示名 | `SHEIN 上架工具 AU` | |
| ASCII 名 | `SheinExtractAU` | |
| Exe 文件名 | `SheinExtractAU.exe` | |
| 安装目录 | `{localappdata}\SheinExtractAU\` | 用户级,无需管理员 |
| 配置目录 | `%APPDATA%\shein-extract-au\config.env` | `config.py` 已读这里,不变 |
| Chrome 资料目录 | `%USERPROFILE%\shein-cdp-profile-au\` | **与母项目分开**(原 `shein-cdp-profile`)避免冲突 |
| Chrome CDP 端口 | `9223`(母项目用 `9222`) | 避免双开端口冲突 |
| AppId GUID | **新 GUID(实施时生成)** | 与母项目互不卸载 |
| 起始版本 | `0.1.0` | 在 `version.py` |
| GitHub release feed | `MikeLiu93/shein-extract-au` | 更新检查目标 |

## 4. 启动流程

```
SheinExtractAU.exe  (无参数)
├─ UTF-8 控制台初始化
├─ 横幅:  "SHEIN 上架工具 AU  v0.1.0"
├─ auth.gate()                          失败 → 退出码 2
├─ update_check.check_for_update()      异常静默跳过(打印一行)
├─ 进入菜单循环:
│      请选择:
│        [1] 配置目标路径
│        [2] 直接跑
│        [Q] 退出
│      > _
├─ 输入 "1":
│      ├─ setup_wizard.run_wizard()
│      ├─ 保存 → 回菜单(可继续选 [2])
│      └─ 用户取消 → 回菜单
├─ 输入 "2":
│      ├─ 检查 config.env 中 SUBMITTED_DIR 是否就绪
│      │      └─ 否 → 打印 "尚未配置,请先选 [1]" → 回菜单
│      ├─ run_excel.main()
│      └─ 结束(无论成功失败) → 退出菜单(去到 pause)
├─ 输入 "Q"/"q"/空回车 → 退出菜单
└─ pause("按 Enter 关闭窗口...")
```

**CLI 旁路参数**:
- `--config`:跳过菜单,直接进配置向导,完了退出
- `--run`:跳过菜单,直接跑 `run_excel.main()`,完了退出
- 未识别参数:透传给 `run_excel.main()`(用于 `--run --no-price-gate "path.xlsx"` 等场景)

## 5. 组件清单

**全部新建文件**,均放在 `shein-extract-au` 仓库根目录:

| 文件 | 来源 | 改动要点 |
|---|---|---|
| `app_main.py` | 移植母项目 | **重写主流程为菜单驱动**;CLI 由 `--config` 扩展为 `--config` / `--run` |
| `setup_wizard.py` | 移植母项目 | 配置项改为 `SHEIN_SUBMITTED_DIR` / `SHEIN_OUTPUT_DIR` / `SHEIN_INPUT_FILENAME` / `ANTHROPIC_API_KEY`;路径校验/默认值推导逻辑同母项目 |
| `auth.py` | 移植母项目 | **独立密码体系**(独立的 `auth.json`);逻辑同母项目 |
| `auth.json` | 新建 | AU 的密码哈希(Mike 用 `make_password_hash.py` 生成,**不进 git** 或者进 git 但只含 hash) |
| `update_check.py` | 移植母项目 | GitHub 仓库地址改为 `MikeLiu93/shein-extract-au`;其它逻辑不变 |
| `version.py` | 新建 | `VERSION = "0.1.0"`;手动 bump |
| `make_password_hash.py` | 移植母项目 | Mike 用来生成 `auth.json` 的离线工具 |
| `pyinstaller.spec` | 移植母项目 | 主入口 `app_main.py`,输出 `SheinExtractAU.exe`;打包 `config.py`、`run_excel.py`、`shein_scraper.py`、`notify.py`、`setup_wizard.py`、`auth.py`、`update_check.py`、`version.py`、`auth.json` |
| `build.bat` | 移植母项目 | PyInstaller 调用 + 输出路径调整 |
| `installer.iss` | 移植母项目 | **只创建 1 个程序图标**(去掉 `--config` 那个);新 GUID;名字、路径全部改 AU |
| `INSTALL_GUIDE_CN.md` | 移植母项目 | 简化版安装说明 |

**`shein_scraper.py` 的小改动**(已知必需):
- `PERSISTENT_PROFILE_DIR` 改为 `~\shein-cdp-profile-au`
- `CDP_PORT` 改为 `9223`

**保持不变的现有文件**:`run_excel.py` / `config.py` / `notify.py` / `requirements.txt` / `test_variant_merge.py`

## 6. 配置向导细节

`setup_wizard.run_wizard()` 依次提示:

1. **输入表所在目录** `SHEIN_SUBMITTED_DIR`
   - 默认值:已有 config 中的值,否则 `D:\共享云端硬盘\02 希音\澳洲站`
   - 校验:目录存在且可读;否则重提示
2. **输出根目录** `SHEIN_OUTPUT_DIR`
   - 默认值:`{SUBMITTED_DIR}\上架资料-已完成`
   - 校验:目录可写;不存在则提示是否创建
3. **指定输入文件名** `SHEIN_INPUT_FILENAME`(可选)
   - 提示:"留空 = 处理目录下所有 .xlsx;填名字 = 只跑这一个"
   - 默认值:当前 config 中的值(目前是 `希音链接 - LU.xlsx`)
4. **Anthropic API Key** `ANTHROPIC_API_KEY`
   - 提示:"留空 = 不启用 AI 标题"
   - 现存值显示为掩码(前 8 + `...` + 后 4)
   - 输入"keep"/"保留"则不修改现有值

向导写出 UTF-8 编码的 `%APPDATA%\shein-extract-au\config.env`(无引号、KEY=VALUE 格式,与 `config.py._load_env_file` 兼容)。

`is_first_run_complete()` 判断:文件存在且至少包含 `SHEIN_SUBMITTED_DIR=`。

## 7. 安装器骨架(`installer.iss`)

```ini
[Setup]
AppId={{<NEW-GUID-HERE>}
AppName=SHEIN 上架工具 AU
AppVersion={#MyAppVersion}
DefaultDirName={localappdata}\SheinExtractAU
PrivilegesRequired=lowest
OutputBaseFilename=SheinExtractAU-Setup-{#MyAppVersion}

[Tasks]
Name: desktopicon;  Description: "在桌面创建快捷方式";  Flags: unchecked
Name: startmenuicon; Description: "在开始菜单创建快捷方式"

[Files]
Source: dist\SheinExtractAU.exe; DestDir: {app}; Flags: ignoreversion

[Icons]
Name: {group}\SHEIN 上架工具 AU; Filename: {app}\SheinExtractAU.exe; Tasks: startmenuicon
Name: {userdesktop}\SHEIN 上架工具 AU; Filename: {app}\SheinExtractAU.exe; Tasks: desktopicon
Name: {group}\卸载 SHEIN 上架工具 AU; Filename: {uninstallexe}; Tasks: startmenuicon
; 只 1 个程序图标 —— 没有 "配置 SHEIN 上架工具 AU" 那一项

[Run]
Filename: {app}\SheinExtractAU.exe; Description: "立即启动"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; 用户数据保留:%APPDATA%\shein-extract-au\, %USERPROFILE%\shein-cdp-profile-au\
```

## 8. 母项目 vs AU 共存

| 维度 | 母项目 | AU |
|---|---|---|
| 安装目录 | `{localappdata}\SheinExtract` | `{localappdata}\SheinExtractAU` |
| Exe 名 | `SheinExtract.exe` | `SheinExtractAU.exe` |
| AppId | 既有 GUID | 新 GUID |
| 配置目录 | `%APPDATA%\shein-extract\` | `%APPDATA%\shein-extract-au\` |
| Chrome 资料 | `~\shein-cdp-profile` | `~\shein-cdp-profile-au` |
| CDP 端口 | 9222 | 9223 |
| 密码体系 | 自己的 `auth.json` | 自己的 `auth.json` |
| 更新源 | `MikeLiu93/shein-extract` | `MikeLiu93/shein-extract-au` |

两者可在同一台机器上**并存且同时运行**(不同端口、不同 profile)。

## 9. 错误处理

| 情况 | 行为 |
|---|---|
| 密码错(达上限) | 退出码 2 |
| 更新检查失败 | 静默 + 一行提示,不阻塞 |
| 配置缺失且选 [2] | 提示"请先选 [1]" → 回菜单 |
| 配置存在但路径失效 | run_excel 报错 → 捕获 → 打印 → pause |
| 向导中 Ctrl+C | 提示"已取消" → 回菜单 |
| pipeline 内 `RateLimitError` | 已有处理:打印 → 走完 pause |
| 任何未捕获异常 | print traceback → pause |
| 菜单输入无效 | 重提示 |

## 10. 测试策略

**自动测试**(扩展现有的 `test_variant_merge.py` 测试集):
- `test_setup_wizard.py` — monkeypatch stdin,断言 config.env 写出正确;路径校验失败时正确报错并重提示
- `test_app_main_menu.py` — monkeypatch stdin,测试菜单各分支:1→wizard 被调用;2→检查 config 后正确分支;Q→直接退;无效输入→重提示

**手动测试**(Mike):
1. `build.bat` → 出 `dist/SheinExtractAU.exe`,直接双击,验证菜单流程
2. `iscc installer.iss` → 出 `dist/SheinExtractAU-Setup-0.1.0.exe`
3. 在新 Windows 用户/同台机器并存母项目下安装,验证无冲突
4. 走完整个流程:输密码 → 出菜单 → 选 [1] → 配 → 回菜单 → 选 [2] → 跑
5. 卸载,验证用户数据保留(配置 + Chrome 资料)

## 11. 已解决的决策

1. **AU 起始密码**:**`Ace2025`**。Mike 用 `make_password_hash.py` 生成 hash 写进 `auth.json`,该 json **进 git**(只含 hash,不含明文)。明文 `Ace2025` 不进 git
2. **未识别 CLI 参数透传给 `run_excel.main()`**:**YES**。`SheinExtractAU.exe --run --no-price-gate "path.xlsx"` 可用
3. **图标 .ico**:**用 Inno Setup 默认**。后期想换专属图标时,扔个 `.ico` 到仓库根目录再改 `installer.iss` 即可
4. **GitHub release 流程**:**Spec 收纳(见 §14)**,并配 `release.bat` 半自动化

## 12. 工作量估计

- 移植/改写 6 个新 .py:1-2 小时
- 配置 `pyinstaller.spec` + `build.bat`:30 分钟
- 写 `installer.iss` + Inno Setup 验证:30 分钟
- 写自动测试:1 小时
- 手动构建 & 烟雾测试:30 分钟

合计 **~3.5-4.5 小时**。

## 13. 实施顺序(待 Plan 细化)

1. `version.py` / `auth.py` / `auth.json`(用 `Ace2025` 生成 hash) / `make_password_hash.py` 移植
2. `update_check.py` 移植 + 改仓库地址为 `MikeLiu93/shein-extract-au`
3. `setup_wizard.py` 改写 AU 配置项
4. `shein_scraper.py` 小改:`PERSISTENT_PROFILE_DIR` → `shein-cdp-profile-au`、`CDP_PORT` → `9223`
5. `app_main.py` 菜单驱动主程序
6. 自动测试(`test_setup_wizard.py` + `test_app_main_menu.py`) + 跑通现有 `test_variant_merge.py`
7. `pyinstaller.spec` + `build.bat` → 出 `dist\SheinExtractAU.exe`
8. `installer.iss` → 出 `dist\SheinExtractAU-Setup-0.1.0.exe`
9. 烟雾测试 + 修补
10. `INSTALL_GUIDE_CN.md`(简化版安装说明)
11. `release.bat`(半自动化发版脚本,见 §14)
12. v0.1.0 首发:走 §14 流程

详细 Plan 在 Spec 批准后用 writing-plans skill 出。

---

## 14. 发版流程

### 14.1 用户视角的更新机制

`update_check.py` 在每次 exe 启动时(最多每 24 小时一次)请求:

```
GET https://api.github.com/repos/MikeLiu93/shein-extract-au/releases/latest
```

返回的 `tag_name`(如 `v0.1.1`)与本地 `version.py` 中的 `VERSION` 比较;新版存在则提示用户下载该 release 的 setup.exe 资源(`SheinExtractAU-Setup-0.1.1.exe`)。

**关键**:`update_check` 看的是 **GitHub Releases**,**不是 git push 或 commits**。仅推代码不会触发任何提示——必须创建 Release 并把 setup.exe 上传为附件(asset)。

### 14.2 发新版要做的事(每次)

1. **bump `version.py`** —— 改 `VERSION = "0.1.x"`,提交并 push
2. **本地构建** —— `build.bat`(出 exe)+ `iscc installer.iss`(出 setup),产物在 `dist\`
3. **打 tag 并 push** —— `git tag v0.1.x` + `git push origin v0.1.x`
4. **GitHub 上创建 Release** —— 仓库页 → Releases → "Draft a new release",选刚 push 的 tag(`v0.1.x`),填发布说明,**把 `dist\SheinExtractAU-Setup-0.1.x.exe` 拖进 Assets 区**,Publish
5. (已安装的 exe 下次启动检测到 → 提示用户升级)

### 14.3 `release.bat` 半自动化

第 1-3 步可以脚本化。`release.bat` 接受新版号作参数:

```bat
release.bat 0.1.1
```

脚本做的事:
- 检查工作区干净、当前在 main、与 origin 同步
- 改 `version.py` 的 `VERSION = "0.1.1"`,git commit + push
- 跑 `build.bat`
- 跑 `iscc installer.iss`
- 验证 `dist\SheinExtractAU-Setup-0.1.1.exe` 存在
- `git tag v0.1.1` + `git push origin v0.1.1`
- 打开浏览器到 `https://github.com/MikeLiu93/shein-extract-au/releases/new?tag=v0.1.1` 让 Mike 手动填说明 + 拖文件 + 点 Publish

第 4 步(写发布说明、上传附件、Publish)**不自动化**——内容需要人工写,且 GitHub release 的附件上传可以用 `gh release create` 但本 Spec 不强制(可选增强)。

### 14.4 可选增强:全自动 `gh release create`

如装了 GitHub CLI(`gh`),`release.bat` 末尾可加:

```bat
gh release create v%VERSION% dist\SheinExtractAU-Setup-%VERSION%.exe ^
   --title "v%VERSION%" ^
   --notes-file release_notes.md
```

要求 Mike 提前写 `release_notes.md`。这步是 nice-to-have,实施 Plan 可标为 optional。
