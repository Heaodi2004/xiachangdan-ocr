# 下场单智能生成系统

基于 OCR 光学字符识别 + 大语言模型的设备下场单自动化生成工具。上传设备铭牌照片，OCR识别文字后由大模型智能提取设备信息，一键生成规范 Excel 下场单。

[![Deploy on Hugging Face](https://huggingface.co/datasets/huggingface/badges/resolve/main/deploy-on-spaces-sm.svg)](https://huggingface.co/spaces)

## ✨ 功能特性

- 📷 **图片自动识别**：支持拖拽上传，自动识别设备铭牌信息
- 🤖 **大模型智能提取**：基于通义千问 qwen3.7-plus 模型，智能提取设备信息
- 🏷️ **多类别管理**：外检设备、带回设备、退检设备三种分类
- 📝 **8个字段自动提取**：证书日期、设备名称、型号、出厂编号、设备编号、生产厂家、备注、负责人
- 📊 **Excel 一键导出**：生成规范格式的下场单 Excel 表格
- 🎨 **颜色标记**：不同类别设备用不同背景色区分
- 🔍 **OCR 原始内容展示**：方便人工核对识别结果
- 🐳 **Docker 支持**：一键部署到 Hugging Face Spaces

## 🛠️ 技术栈

| 类别 | 技术 |
|------|------|
| 后端 | Python + Flask + Gunicorn |
| OCR 引擎 | EasyOCR（基于 PyTorch） |
| 大模型 | 通义千问 qwen3.7-plus（阿里云百炼） |
| 图像处理 | OpenCV + NumPy |
| Excel 生成 | OpenPyXL |
| 前端 | HTML5 + Tailwind CSS + 原生 JavaScript |
| 部署 | Docker / Hugging Face Spaces |

## 🚀 快速开始

### 环境要求

- Python 3.8 ~ 3.12
- 4GB 及以上内存
- 大模型 API 密钥（阿里云百炼）
- 现代浏览器（Chrome / Edge / Firefox）

### 本地运行

```bash
# 1. 克隆项目
git clone <仓库地址>
cd 下场单

# 2. 安装依赖
pip install -r requirements.txt

# 3. 配置大模型API
cp api.csv.example api.csv
# 编辑 api.csv，填入你的阿里云百炼 API 密钥

# 4. 启动应用
python app.py

# 5. 访问 http://127.0.0.1:5000
```

> **首次启动**：EasyOCR 会自动下载模型文件（约 100MB），请确保网络连接正常。
>
> **大模型配置**：将 `api.csv.example` 复制为 `api.csv`，填入你的阿里云百炼 API 配置。如果不配置大模型，系统将仅使用正则表达式提取信息。

### Docker 本地运行

```bash
# 构建镜像
docker build -t xiachangdan .

# 运行容器（配置API）
docker run -p 7860:7860 \
  -e USE_LLM=True \
  -e LLM_MODEL=qwen3.7-plus \
  xiachangdan

# 访问 http://127.0.0.1:7860
```

## ☁️ 部署到 Hugging Face Spaces

本项目已配置好 Docker 部署文件，可一键部署到 Hugging Face Spaces（免费 16GB 内存！）

### 一键部署步骤

1. 将代码推送到 GitHub 仓库
2. 登录 [huggingface.co](https://huggingface.co)
3. 进入 Spaces → Create new Space
4. 填写 Space 名称，选择 **Docker** SDK
5. 选择 "Public" 或 "Private"
6. 连接你的 GitHub 仓库
7. 等待构建完成，即可访问！

详细步骤请参考 [部署指南.md](./部署指南.md)

## 📖 使用说明

### 基本流程

1. 填写公司信息（单位名称、联系人、地址等）
2. 上传设备铭牌图片到对应分类区域
3. 核对设备明细表中的识别结果
4. 点击"导出 Excel 表格"生成下场单

### 支持的图片格式

PNG、JPG、JPEG、GIF、BMP、WebP、TIFF、TIF

### 设备分类

| 分类 | 颜色 | 说明 |
|------|------|------|
| 外检设备 | 🟢 绿色 | 现场检定，不需要带回 |
| 带回设备 | 🔵 蓝色 | 需要带回实验室检定 |
| 退检设备 | 🔴 红色 | 退检不做的设备 |

## 📁 项目结构

```
下场单/
├── app.py                    # 主应用程序
├── Dockerfile                # Docker 部署配置
├── .dockerignore           # Docker 忽略规则
├── requirements.txt          # Python 依赖
├── .gitignore               # Git 忽略规则
├── templates/
│   └── index.html           # 前端页面
├── uploads/                 # 上传的图片（运行时创建）
├── output/                  # 生成的 Excel（运行时创建）
├── 使用教程.md              # 详细使用教程
├── 部署指南.md              # Hugging Face 部署指南
└── README.md                # 本文档
```

## 🔧 API 接口

### 图片上传识别

```
POST /upload
Content-Type: multipart/form-data

参数:
  - file: 图片文件
  - category: 设备类别 (inspected/brought_back/returned)

返回: JSON格式的设备信息
```

### 生成 Excel

```
POST /generate_excel
Content-Type: application/json

参数:
  - company_info: 公司信息对象
  - devices: 设备列表数组

返回: 生成的文件名
```

### 下载 Excel

```
GET /download/<filename>
```

## ⚠️ 注意事项

1. OCR 识别结果仅供参考，请人工核对后再导出
2. 首次启动需要下载 OCR 模型，可能需要几分钟
3. 图片越清晰，识别准确率越高
4. 建议图片中铭牌文字清晰、光线充足
5. Hugging Face Spaces 免费版 48 小时无活动会暂停，点击即可唤醒

## 📝 更新日志

### v1.0.0

- 初始版本发布
- 支持图片 OCR 识别和设备信息提取
- 支持三种设备分类
- 支持 Excel 一键导出
- 多策略图像预处理，提高识别准确率
- 支持 Docker 部署和 Hugging Face Spaces 部署

## 📄 许可证

MIT License

---

**提示**：
- 更详细的使用说明请参考 [使用教程.md](./使用教程.md)
- 部署步骤请参考 [部署指南.md](./部署指南.md)

