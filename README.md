# 软件安装助理 (Software Installer Assistant)

![Mainform](https://github.com/wanglanaaa/Software-Installer-Assistant/blob/main/Mainform.png)

## 项目简介

“软件安装助理”是一个桌面应用程序，旨在帮助用户高效地管理和安装各类软件。它提供了一个直观的用户界面，允许用户对软件进行分类、管理安装方案，并支持智能化的静默安装。

## ✨ 主要功能

*   **软件分类管理**：用户可以创建、删除软件分类，方便地组织和查找软件。
*   **软件信息管理**：支持添加、删除软件条目，并记录软件名称、版本、大小、描述、安装路径和静默安装参数等信息。
*   **智能静默安装**：
    *   程序内置了多种常见 EXE 安装包的静默参数（如 `/S`, `/VERYSILENT`, `/q` 等）。
    *   在安装过程中，程序会智能尝试这些参数，并通过**窗口检测机制**判断静默安装是否成功。
    *   如果成功，程序会询问用户确认，并将有效的静默参数**自动保存**到配置中，实现“一次学习，永久使用”。
    *   **高质量数据持久化**：只有经过验证的、成功的静默安装配置才会被保存到 `software_list.json` 知识库中，确保数据的纯粹性和有效性。
*   **按需加载界面**：程序启动时界面简洁，左侧树形列表默认为空，只显示用户当前会话中手动添加的软件，提高操作效率和专注度。
*   **安装方案管理**：
    *   **保存方案**：用户可以将当前选中的软件列表及其勾选状态保存为一个安装方案，方便日后快速加载。
    *   **加载方案**：通过下拉框选择并自动加载已保存的安装方案，恢复软件的勾选状态。
    *   **删除方案**：管理和删除不再需要的安装方案。
*   **安装路径选择**：用户可以自定义软件的安装路径。
*   **现代简约UI**：采用类似 Microsoft Edge 风格的简约用户界面，提供清晰、专业的视觉体验。

## 🛠️ 技术栈

*   **C#**：主要开发语言。
*   **.NET 9.0**：应用程序框架。
*   **Windows Forms**：用于构建桌面用户界面。
*   **System.Text.Json**：用于方案和软件列表的序列化和反序列化。

## 📂 项目结构

项目采用分层结构，将不同的功能模块分离到独立的项目中，以实现高内聚、低耦合。

*   `SoftwareInstaller.UI`：用户界面层，包含主窗体 `MainForm`。
*   `SoftwareInstaller.Core`：核心业务逻辑层，负责管理软件数据、核心功能（如智能安装）和业务相关的管理（如方案管理）。
*   `SoftwareInstaller.Models`：数据模型层，定义了如 `SoftwareItem` 等核心数据结构。
*   `SoftwareInstaller.Utils`：通用工具层，提供最底层的、与具体业务无关的辅助功能，如文件元数据读取、进程调用、错误处理和窗口检测等。

## 🚀 如何运行

### 方式一：为普通用户 (推荐)

1.  前往本项目的 [**Releases 页面**](https://github.com/wanglanaaa/Software-Installer-Assistant/releases)。
2.  下载最新版本的 `.zip` 压缩包文件（例如 `Software.Installer.Assistant.vX.X.X.zip`）。
3.  将压缩包解压到你希望的任意位置。
4.  进入解压后的文件夹，找到并双击运行 `SoftwareInstaller.UI.exe` 即可启动程序。

### 方式二：为开发者

如果你希望自行编译或对代码进行修改，请遵循以下步骤：

1.  **环境准备**：确保你已安装 [.NET 9.0 SDK](https://dotnet.microsoft.com/download/dotnet/9.0) 或更高版本。
2.  **克隆仓库**：
    ```bash
    git clone https://github.com/wanglanaaa/Software-Installer-Assistant.git
    cd Software-Installer-Assistant
    ```
3.  **运行程序**：
    ```bash
    dotnet run --project SoftwareInstaller.UI
    ```

## 📝 未来计划 (TODO)

- [ ] 优化软件列表的加载机制，支持更丰富的元数据（例如，从文件读取更多信息）。
- [ ] 提供更美观的自定义主题或皮肤功能。
- [ ] 考虑集成更多安装包类型（如 MSI）的静默参数支持。

## 🤝 贡献

欢迎任何形式的贡献！如果您有任何建议或发现 Bug，请随时提交 Issue 或 Pull Request。

## 📄 许可证

[待定]