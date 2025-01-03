# MDZ Launcher

**MDZ Launcher** 是一款专为管理 `.mdz` 文件设计的桌面应用程序，它与强大的 Markdown 编辑器 **Typora** 和高效的压缩工具 **7-Zip** 协同工作。通过简化 Markdown 文档及其相关资源的编辑与打包流程，MDZ Launcher 旨在为用户提供一个高效、便捷的文档管理解决方案。

使用chatgpt辅助编写。

## **主要功能**

- **新建 `.mdz` 文件**：一键创建包含 `document.md` 和 `document.assets` 目录的全新 `.mdz` 文件，快速开始您的 Markdown 项目。

- **打开和编辑 `.mdz` 文件**：轻松打开已有的 `.mdz` 文件，自动解压到临时目录，并使用 Typora 进行实时编辑。

- **自动打包**：在 Typora 中保存 `document.md` 文件时，MDZ Launcher 会自动监控文件变更并重新打包 `.mdz` 文件，确保您的文档始终处于最新状态。

- **路径配置**：通过友好的图形界面，用户可以轻松配置 7-Zip 和 Typora 的路径，支持相对路径和绝对路径，适应不同的系统环境。

- **实时日志记录**：集成多行日志显示功能，实时记录程序的操作和状态更新，帮助用户追踪和排查问题。

- **文件关联**：将 `.mdz` 文件默认关联到 MDZ Launcher，实现双击文件自动启动并编辑的便捷操作。

## **使用场景**

- **文档管理**：适用于需要组织和管理大量 Markdown 文档及其相关资源的用户，如技术文档编写、项目说明书制作等。

- **内容创作**：为内容创作者提供一个集成的编辑和打包工具，简化创作流程，提高工作效率。

- **协作项目**：在团队协作中，MDZ Launcher 能够帮助团队成员统一管理和编辑 Markdown 文档，确保文档的一致性和最新性。

## **优势与特点**

- **协同工作**：通过与 Typora 和 7-Zip 的协同工作，提供统一的用户体验，避免频繁切换工具。

- **自动化流程**：通过自动监控和打包功能，减少手动操作，降低出错概率，提高工作效率。

- **用户友好**：简洁直观的图形界面，易于上手，无需复杂的配置，即可快速开始使用。

- **可扩展性**：支持自定义路径配置，适应不同的系统环境和用户需求，具有良好的灵活性。

## **安装与使用**

1. **下载并安装**：

   - 前往 [GitHub Releases](https://github.com/archiezhao/mdzlauncher/releases) 页面，下载最新版本的 `mdzlauncher.exe` 安装包。
   - 双击安装包，按照提示完成安装。

2. **安装依赖工具**：

   - 确保您已在系统中安装 **Typora** 和 **7-Zip**。您可以从以下链接下载并安装它们：
     - [Typora](https://typora.io/)
     - [7-Zip](https://www.7-zip.org/)

3. **配置路径**：

   - 启动 MDZ Launcher 后，导航至菜单栏的 **“设置”** > **“配置路径”**。
   - 设置 7-Zip 和 Typora 的可执行文件路径，确保程序能够正确调用这两个工具。

4. **创建或打开 `.mdz` 文件**：

   - 通过 **“文件”** 菜单选择 **“新建 .mdz 文件”**，选择保存位置，开始创建新的 Markdown 项目。
   - 或者选择 **“打开 .mdz 文件”**，加载已有的项目，进行编辑和管理。

5. **编辑与打包**：

   - 使用 Typora 编辑 `document.md` 文件，保存时程序会自动监控并重新打包 `.mdz` 文件。
   - 关闭 Typora 后，程序将完成最终的打包工作，并清理临时目录。

## **贡献与支持**

**MDZ Launcher** 是一个开源项目，欢迎任何形式的贡献！您可以通过以下方式参与：

- **报告问题**：在 [Issues](https://github.com/archiezhao/mdzlauncher/issues) 页面提交您遇到的问题或建议的功能。
- **贡献代码**：Fork 本仓库，创建新的分支，提交 Pull Request，我们会尽快审核您的贡献。
- **讨论与反馈**：加入我们的讨论，分享您的使用体验和改进建议。

## **许可证**

本项目采用 [MIT 许可证](LICENSE)，您可以自由使用、修改和分发该软件，但需保留原作者的版权声明和许可证说明。

---

感谢您选择 **MDZ Launcher**！我们致力于不断优化和改进，为您提供更好的文档管理体验。如果您有任何问题或建议，欢迎随时与我们联系。

### 打包为可执行文件

```
pyinstaller --onefile --windowed --icon=icon.ico mdzlauncher_V1.3.py
```

打包完成后，生成的 `mdzlauncher.exe` 文件位于 `dist` 目录下。

### typora需要设置插入图片复制到文件夹内
![Typora_2024-12-20_16-47-20 677_星期五](https://github.com/user-attachments/assets/dc33013f-3041-4da3-8f73-ba13db84bbcd)


### 使用前需要先配置typora和7z的路径
![mdzlauncher_V1 3_2024-12-20_16-49-39 861_星期五](https://github.com/user-attachments/assets/85fcb358-0766-4a73-9dfe-9cf78af4d9b3)
