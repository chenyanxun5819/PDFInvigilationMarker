# PDF 监考名单标记工具

这个工具会：

1. 让使用者自行选择任意监考 PDF。
2. 读取内建设定或匯入的老师名单 Excel。
3. 若 PDF 有文字层，会直接读取 PDF 里的老师代号与座标，不再依赖 OCR。
4. 若同资料夹内有排班 Excel，会从主表自动抓老师代号并去重；**不会直接把「去重」工作表当成输出清单**。
5. 只有在 PDF 没有可用文字层时，才回退用 OCR 找老师代号位置。
6. 依座位区块排序后，为每位老师输出独立标记页。
7. 另存新的 PDF，并可再用 PyInstaller 打包成 `.exe`。

## 资料设定

- 内建设定：`data/default_config.json`
- 使用者设定：`%APPDATA%/.pdf_invigilation_marker/config.json`

若老师名单或座位区块有变动，可直接在程式内按 **「匯入老师名单 Excel」** 重新匯入，或手动修改使用者设定 JSON。

若输入 PDF 同资料夹内存在排班 Excel，程式会自动：

- 从主表逐格扫描老师代号并自动去重
- 以老师名单主档中的姓名 / 代号 / 座号输出左上角标签
- 若 PDF 本身有文字层，则以 PDF 文字座标作为 `*` 标记位置

## 启动 GUI

```powershell
d:\chhsban\chhsban-acadoc\.venv\Scripts\python.exe d:\chhsban\pdf_invigilation_marker\main.py
```

## 命令列处理

```powershell
d:\chhsban\chhsban-acadoc\.venv\Scripts\python.exe d:\chhsban\pdf_invigilation_marker\main.py `
  --input "D:\path\input.pdf" `
  --output "D:\path\output.pdf"
```

## 匯入 Excel 设定

```powershell
d:\chhsban\chhsban-acadoc\.venv\Scripts\python.exe d:\chhsban\pdf_invigilation_marker\main.py `
  --import-excel "D:\path\老师名单.xlsx"
```

## 打包成可执行档

```powershell
d:\chhsban\chhsban-acadoc\.venv\Scripts\python.exe d:\chhsban\pdf_invigilation_marker\build.py
```

输出档会在 `pdf_invigilation_marker/dist/PDF监考名单标记工具.exe`。

> exe 现已改为无主控台视窗模式；双击时不会再跳出黑底 cmd 背景视窗。
