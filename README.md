# KK 成长记录跨平台项目

这是跨平台项目：

- 电脑端：Electron 窗口，保存时直接修改 `KK个人成长系统.docx`。
- 手机端：PWA，可在手机浏览器中添加到主屏幕，像 App 一样打开。
- 本地数据库：浏览器 IndexedDB，每条记录有唯一 ID，可导入 JSON 合并，避免按日期覆盖记录。
- 电脑端 Word 主文档：电脑 App 保存时直接修改 `KK个人成长系统.docx`，先自动备份，再写入对应表格并按日期排序。
- 手机转电脑：手机点击“导出记录 JSON”，把 JSON 文件传给电脑；电脑点击“导入记录 JSON”，按唯一记录 ID 合并到 Word，不覆盖已有记录。
- 手机端 Word 导出：手机可导出当前本地记录的 Word 副本，但不会直接修改电脑硬盘上的 Word。
- 云同步：当前未加入。

## 网站构建

项目内置 GitHub Actions：

- `Build KK Growth Record`：手动运行或推送 `v*` 标签时，生成 Windows `.exe` 和手机端网站构建产物。
- `Publish KK Growth Record Mobile Web App`：推送到 `main` 后发布 GitHub Pages，手机打开页面后可添加到主屏幕。

GitHub Pages 的 source 选择 GitHub Actions。网站端数据仍保存在手机本地，JSON 导出/导入用于手机和电脑之间合并记录。

## 运行

```powershell
npm install
npm run dev
```

双击 `启动桌面开发版.bat` 也可以启动电脑端窗口。要生成桌面运行目录，运行：

```powershell
npm run package:dir
```

然后双击 `dist/win-unpacked/KK成长记录.exe`。要生成可直接双击的便携版 Windows 程序，运行：

```powershell
npm run package:win
```

手机端和电脑端分开打包时，使用下面的命令。电脑端会输出到 `desktop-dist`，不会覆盖 `dist` 中的手机端文件：

```powershell
npm run package:dir:separate
npm run package:win:separate
```

重新打包时，请先关闭正在运行的电脑端程序和 Word 主文档，再运行对应命令。

首次打包会下载 Electron Windows 运行时，生成文件在 `dist` 目录中。首次使用在左侧选择 `KK个人成长系统.docx`；之后保存会直接写入该文件。

手机测试时，让手机和电脑连接同一个局域网，然后运行：

```powershell
npm run web
```

在手机浏览器打开电脑显示的局域网地址。生产环境要通过 HTTPS 部署后，才能稳定使用 PWA 的离线能力和“添加到主屏幕”。

## 记录合并

手机和电脑记录先保存为独立记录，导入 JSON 时按唯一 ID 合并；Word 导出时按记录类型和日期重新生成表格，因此 10 月 2 日的手机记录会自动排在 10 月 1 日和 10 月 3 日之间。
