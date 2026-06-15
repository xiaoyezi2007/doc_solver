# doc_solver

常用文件处理脚本集合。默认把源文件放到 `input/`，脚本输出放到 `output/`。

## 脚本

- `ppt2image.py`: 将 PPT/PPTX 每页导出为高清 PNG。
- `ppt_flatten.py`: 将 PPT 扁平化为每页一张背景图的新 PPTX；视频和可识别的 GIF 会覆盖回原位置。
- `ppt_export_media.py`: 从 PPTX 包中导出原始图片、视频、音频等素材，尽量保持原始清晰度。

## 用法

把文件放进 `input/` 后运行：

```powershell
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

需要 Windows + Microsoft PowerPoint，Python 包依赖 `comtypes`。
