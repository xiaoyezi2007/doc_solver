# doc_solver

常用文件处理脚本集合。默认把源文件放到 `input/`，脚本输出放到 `output/`。

## 脚本

- `app.py`: 图形界面入口，首页以功能卡片组织，每个功能有独立处理页面；功能页统一为左侧文件队列预览、右侧选项和底部开始按钮。
- `ppt2image.py`: 将 PPT/PPTX 每页导出为高清 PNG。
- `ppt_flatten.py`: 将 PPT 扁平化为每页一张背景图的新 PPTX；视频和可识别的 GIF 会覆盖回原位置。
- `ppt_export_media.py`: 从 PPTX 包中导出原始图片、视频、音频等素材，尽量保持原始清晰度。

## 用法

把文件放进 `input/` 后运行：

```powershell
python .\app.py
python .\ppt2image.py
python .\ppt_flatten.py
python .\ppt_export_media.py
```

也可以指定单个文件或目录：

```powershell
python .\ppt2image.py .\input\demo.pptx
python .\ppt_flatten.py .\input\demo.pptx
python .\ppt_export_media.py .\input\demo.pptx
```

## 依赖

需要 Windows + Microsoft PowerPoint，Python 包依赖见 `requirements.txt`。

## 打包 exe

运行：

```powershell
.\build_exe.ps1
```

打包产物在 `dist/DocSolver.exe`。把 `DocSolver.exe`、`input/`、`output/` 放在同一个目录，用户双击 exe 即可使用。

## 前端交互约定

后续新增功能统一按现有功能页框架实现：

- 左侧为文件队列预览区，支持拖入多个文件、预览首页、拖动调整处理顺序。
- 左侧底部放置加号按钮，用于打开文件选择器继续添加文件。
- 右侧为任务选项，每个任务至少包含输出路径选项：与源文件相同或指定路径。
- 右侧底部固定放置醒目的“开始”按钮。
- 每次任务输出到独立目录，命名为第一个文件名加秒级时间戳，避免覆盖已有结果。
