# MDZ Launcher

MDZ Launcher 是一个用于管理 `.mdz` 文件的桌面应用程序。需要自定义 Typora 和 7-Zip路径，允许用户轻松编辑和打包 Markdown 文件。

使用chatgpt辅助编写

## 功能

- 新建和打开 `.mdz` 文件
- 自动监控 `document.md` 的保存操作并自动打包
- 配置 7-Zip 和 Typora 的路径
- 实时日志记录和错误报告

## 安装

### 前提条件

- [Python 3.x](https://www.python.org/downloads/)
- [PyInstaller](https://www.pyinstaller.org/)
- [7-Zip](https://www.7-zip.org/)
- [Typora](https://typora.io/)

### 克隆仓库

```bash
git clone https://github.com/ArchieZhao/mdzlauncher.git
cd mdzlauncher
```



### 安装依赖

```
pip install -r requirements.txt
```



### 打包为可执行文件

```
pyinstaller --onefile --windowed --icon=icon.ico mdzlauncher_V1.3.py
```

打包完成后，生成的 `mdzlauncher.exe` 文件位于 `dist` 目录下。
