# 大学节讲座抽签

读取 `student.xlsx`，在浏览器显示 SVG 滚轮抽签动画，并自动保存 `result.xlsx`。无需联网使用。

## 分配规则

- 每位同学抽两次，分别对应上半场、下半场，两次大学不同。
- 同一半场内，每所大学只安排一位同学；同一所大学可以在两个半场安排不同同学。
- 自动读取实际名单人数。同场大学数量必须不少于学生人数，否则页面提示名额不足，不能开始抽签。
- 先随机生成满足规则的完整方案，再逐次揭晓，避免最后一位同学无有效选项。每次揭晓前自动保存。

## 在 GitHub 上打包 Windows EXE

1. 把本项目源码提交到 GitHub 仓库的默认分支，保留 `.github/workflows/build-windows.yml` 的目录结构。上传 `lottery.py`、`lottery.html`、两份 requirements 文件、`test_lottery.py` 和 `tests/smoke_packaged.py`。打包不需要上传学生名单。
2. 打开仓库 **Actions → Build Windows EXE → Run workflow**。
3. 等待工作流成功，打开该次运行，在 **Artifacts** 下载 `university-lottery-windows-x64`。
4. 解压，得到 `university-lottery.exe`。它已包含 Python、所需库和网页，教室电脑不需要另外安装 Python。

工作流在 Windows x64 / Python 3.11 下打包，并检查生成的 EXE。保留控制台窗口，以便显示启动错误和关闭程序。

## 在教室使用

将文件放在同一文件夹中，例如：

```text
大学节抽签/
  university-lottery.exe
  student.xlsx
```

`student.xlsx` 使用第一张工作表。推荐 A1 写“姓名”，从 A2 起一行一位；也支持无表头、首行空白、带“学生姓名”表头及空行。重名按名单中的不同记录处理。

双击 EXE，浏览器会自动打开本地页面。点击“抽一次”或按空格键逐次揭晓，也可以“自动抽取”；点击“暂停自动”后完成当前抽取再暂停。可用页面上的全屏按钮进行投影。

每次抽签后，程序将结果写入 EXE 旁的 `result.xlsx`，包含序号、学生姓名、上半场、下半场四列。尚未抽取的单元格为空。“下载结果”可以另存一份。

关闭浏览器不会关闭程序；结束时关闭程序控制台，或按 `Ctrl+C`。下次启动会从结果文件恢复进度。若要全新抽签，先关闭程序，将旧 `result.xlsx` 改名备份后再启动。更换名单后也要备份移走旧结果。

若提示保存失败，请关闭 Excel 中已打开的结果文件，再重试；未成功保存的抽签不会推进进度。若浏览器没有自动打开，将控制台显示的本地地址复制到浏览器。

## 直接运行 Python

需要 Python 3.10 或更新版本，`lottery.py` 和 `lottery.html` 放在一起：

```sh
python -m pip install -r requirements.txt
python lottery.py
```

macOS 上可把命令中的 `python` 改为 `python3`。名单和结果默认位于程序旁边，与从哪个文件夹启动无关。

可选参数：

```sh
python lottery.py --input student.xlsx --output result.xlsx --port 8765 --no-browser
python -m unittest -v test_lottery
```

在 Windows 本地手动打包：

```sh
python -m pip install -r requirements-build.txt
python -m PyInstaller --clean --noconfirm --onefile --console --name university-lottery --add-data "lottery.html;." lottery.py
```

大学列表集中在 `lottery.py` 的 `UNIVERSITIES` 中；修改后重新打包即可。提供的 24 所大学名称均已保留。
