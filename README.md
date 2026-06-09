# 👨‍🏫 智能教师作业批改助手

> AI 驱动的多维度作业批改智能体 · 支持 9 大学科 · 完全免费使用

[![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=Streamlit&logoColor=white)](https://streamlit.io)
[![Python](https://img.shields.io/badge/Python-3.10+-blue?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](LICENSE)

---

## ✨ 功能亮点

| 🆓 完全免费 | 📚 9大学科 | 🤖 7款模型 | 📂 3种输入 |
|:---:|:---:|:---:|:---:|
| 集成多款免费 API | 语数英物化生史地政 | DeepSeek/Qwen/GLM等 | 粘贴/单文件/批量 |

- 🎯 **5维度评分**：每学科定制 100 分制评分体系
- 📥 **报告导出**：一键下载 Markdown 批改报告
- 📜 **历史记录**：自动保存，随时回溯
- 🎨 **精美 UI**：现代化设计，流畅交互

## 🚀 快速开始

```bash
# 1. 安装依赖
pip install streamlit openai python-docx PyPDF2

# 2. 启动应用
streamlit run app.py

# 3. 浏览器打开 http://localhost:8501
```

## 🆓 免费 API 获取

**推荐：硅基流动（最简单）**

1. 访问 [siliconflow.cn](https://siliconflow.cn) 注册
2. 控制台 → API 密钥 → 复制
3. 在应用中选择「硅基流动」系列模型，粘贴密钥

> 更多免费方案见 [产品说明书](./产品说明书.md)

## 🔗 在线体验

> **https://teacher-grading-assistant.streamlit.app**

*（部署后替换为实际链接）*

## 📖 文档

- [完整产品说明书](./产品说明书.md) — 功能详解、操作指南、FAQ
- [部署指南](#-部署到-streamlit-cloud) — 一键部署获取公开链接

## 🌐 部署到 Streamlit Cloud

```bash
# 1. 推送到 GitHub
git init && git add . && git commit -m "v2.0"
git remote add origin https://github.com/YOUR_USERNAME/teacher-grading-assistant.git
git push -u origin main

# 2. 访问 https://share.streamlit.io → New app → 选择仓库 → Deploy
# 3. 1分钟后获得公开链接！
```

## 📁 项目结构

```
├── app.py                # 主应用（单文件）
├── requirements.txt      # Python 依赖
├── packages.txt          # 系统依赖
├── .streamlit/config.toml# 主题配置
└── 产品说明书.md          # 完整文档
```

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 License

MIT License

---

> **让 AI 成为每位教师的得力助手，把时间还给教学本身。**
