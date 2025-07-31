# 📝 Word 写作计划助手 (Writing Planner Add-in)

这是一个为 Microsoft Word 设计的加载项（Add-in），旨在帮助写作者、学生和内容创作者更高效地管理他们的写作项目。通过设定目标、追踪进度和可视化分析，您可以更好地掌控写作节奏，告别拖延。

<img width="506" height="525" alt="写作助手界面" src="https://github.com/user-attachments/assets/ee22a7ff-ea31-4274-abd5-a69d801452a3" />

**在线体验和安装**: [https://chunyanghu.github.io/WritingPlannerAdd/](https://chunyanghu.github.io/WritingPlannerAdd/)

---

## ✨ 主要功能

*   **🎯 计划设置**:
    *   为您的写作项目设定一个清晰的名称。
    *   设置总目标字数和最终截止日期。
    *   规划每日需要完成的最低字数。
*   **📈 进度追踪**:
    *   一键更新，实时获取并记录当前文档的总字数。
    *   直观的进度条，显示整体项目完成百分比。
    *   关键数据概览：当前字数、目标字数、剩余天数、今日已写字数。
*   **📊 统计分析**:
    *   **可视化图表**: 以折线图形式展示累计字数和每日新增字数的趋势。
    *   **写作历史**: 查看最近的写作记录，包括日期、总字数和每日增量。
*   **⏰ 提醒功能**:
    *   自定义每日提醒时间。
    *   到点后，如果尚未完成当日目标，插件会发送提醒。

---

## 🚀 快速安装

您可以通过以下简单的步骤在您的 Word 中安装并使用此插件：

1.  **访问项目主页**:
    打开浏览器，访问 [https://chunyanghu.github.io/WritingPlannerAdd/](https://chunyanghu.github.io/WritingPlannerAdd/)

2.  **下载安装文件**:
    在页面上，点击 **“下载安装文件 (manifest.xml)”** 按钮，将 `manifest.xml` 文件保存到您的电脑上。

3.  **在 Word 中上传加载项**:
    *   打开 Microsoft Word (支持 Word 2016+, Word Online, Word for Mac)。
    *   点击顶部菜单栏的 **`插入 (Insert)`**。
    *   点击 **`我的加载项 (My Add-ins)`**。
    *   在弹出的窗口中，选择 **`上传我的加载项 (Upload My Add-in)`**。
    *   在文件选择框中，找到并选择您刚刚下载的那个 `manifest.xml` 文件。

4.  **开始使用**:
    *   安装成功后，在 Word 的 **`开始 (Home)`** 选项卡下，您会看到一个新的分组“写作计划”，点击 **“打开写作计划”** 按钮即可启动插件。

---

## 🛠️ 技术栈

*   **核心**: HTML, CSS, JavaScript (ES6+)
*   **Office API**: [Office.js](https://learn.microsoft.com/zh-cn/javascript/api/office?view=common-js-preview)
*   **UI & 样式**: [Bootstrap](https://getbootstrap.com/)
*   **图表库**: [Chart.js](https://www.chartjs.org/)
*   **开发工具**: Node.js, Webpack, Yeoman Office Add-in generator

---

## 🧑‍💻 本地开发

如果您想对本项目进行修改或贡献，可以按照以下步骤设置本地开发环境：

1.  **克隆仓库**:
    ```bash
    git clone https://github.com/chunyanghu/WritingPlannerAdd.git
    cd WritingPlannerAdd
    ```

2.  **安装依赖**:
    ```bash
    npm install
    ```

3.  **安装并信任开发证书**:
    这是 Office 加载项本地开发所必需的，它会为 `localhost` 创建一个受信任的 HTTPS 证书。
    ```bash
    npx office-addin-dev-certs install
    ```
    *在弹出的安全提示中，请务必选择“是”。*

4.  **启动开发服务器**:
    ```bash
    npm start
    ```
    服务器将在 `https://localhost:3000` 上运行。

5.  **在 Word 中旁加载 (Sideload)**:
    *   在 Word 中，按照“快速安装”的步骤3操作，但这次不是上传文件，而是通过信任共享文件夹的方式。
    *   具体请参考 [Office 加载项官方旁加载文档](https://learn.microsoft.com/zh-cn/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)。

---

## 🤝 贡献

欢迎任何形式的贡献！如果您有任何好的建议或发现了 Bug，请随时提交 [Issue](https://github.com/chunyanghu/WritingPlannerAdd/issues) 或 Pull Request。

## 📄 许可证

本项目采用 [MIT License](LICENSE) 开源。
