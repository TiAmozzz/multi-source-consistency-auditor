# 文件选择核验工具

双击 `file_picker.cmd`，或运行：

```bash
python tools/file_picker.py
```

页面支持分次选择文件。每次点击“选择文件”都会把新文件追加到已选列表，可逐个移除或清空，然后一次核验全部资料。支持 CSV、TSV、XLSX、XLSM、JSON、DOCX、PDF、TXT、Markdown 或 HTML 文件。

输出保存在 `user_data/file-picker-runs/`，所选文件的本机副本保存在 `user_data/file-picker-uploads/`。这些目录已加入 `.gitignore`，不会提交到代码仓库。

原始文件不会被修改、重命名或转换。服务只监听 `127.0.0.1`，不向互联网发送选中资料。
