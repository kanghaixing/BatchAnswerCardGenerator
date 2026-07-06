# 答题卡批量生成器 - 前后端完整版

这是一个可在本地或服务器运行的答题卡批量生成工具。

你只需要上传 Excel 学生信息，系统会自动生成：

- 合并版 PDF：适合直接批量打印
- 单人版 PDF ZIP：每个学生一份，便于归档或单独发送
- 第一页预览图：用于快速检查版式

当前项目已内置本次配置好的答题卡模板，包含：

- 顶部新增 `考试名称：______`
- `考生姓名` 自动写入 Excel 中的姓名
- `考号` 替换为 `识别号`
- 右上角二维码内容为识别号
- 修复并重绘姓名/识别号区域附近四个小黑色视觉定位点
- 前端页面：白色主面板 + 蓝色辅助色，支持点击上传和拖拽上传

---

## 目录结构

```text
answer_card_batch_web_project/
├── app.py                         # FastAPI 后端入口
├── generate_answer_cards.py        # PDF 生成核心逻辑，也可命令行运行
├── config.py                       # 坐标与样式配置
├── requirements.txt                # Python 依赖
├── run_web.sh                      # macOS / Linux 一键启动脚本
├── run_web.bat                     # Windows 一键启动脚本
├── static/
│   └── index.html                  # 前端 HTML/CSS/JS
├── template/
│   └── answer_card_template.pdf    # 固定答题卡模板
├── input/
│   ├── 学生信息.xlsx                # 示例 Excel
│   ├── 学生信息_示例.xlsx
│   └── 学生信息_84人示例.xlsx
├── output/                         # 命令行模式输出目录
└── jobs/                           # 网页模式临时任务输出目录
```

---

## 1. 安装环境

建议使用 Python 3.10 或更高版本。

进入项目目录：

```bash
cd answer_card_batch_web_project
```

安装依赖：

```bash
pip install -r requirements.txt
```

---

## 2. 启动网页工具

### macOS / Linux

```bash
./run_web.sh
```

或者：

```bash
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

### Windows

双击：

```text
run_web.bat
```

或者在命令行运行：

```bash
python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

启动后打开浏览器访问：

```text
http://127.0.0.1:8000
```

---

## 3. 网页使用方式

1. 打开 `http://127.0.0.1:8000`
2. 上传 Excel 学生信息表
3. 可选填写考试名称
   - 留空：PDF 中显示 `考试名称：______`，方便打印后手写
   - 填写：PDF 中直接印上考试名称
4. 点击 `生成答题卡 PDF`
5. 下载：
   - 合并版 PDF
   - 单人版 ZIP

---

## 4. Excel 格式要求

Excel 第一行必须是表头。

默认支持以下列名：

| 字段 | 可接受列名 |
|---|---|
| 姓名 | 姓名、学生姓名、考生姓名、name |
| 识别号 | 识别号、考号、准考证号、学生编号、id、student_id |

最简单格式：

| 姓名 | 年级 | 识别号 |
|---|---|---|
| 张三 | 高1 | 1S001 |
| 李四 | 高2 | 2S002 |

程序只强制需要 `姓名` 和 `识别号` 两列，其他列可以保留。

---

## 5. 命令行模式仍然可用

除了网页模式，你也可以继续使用命令行批量生成。

默认运行：

```bash
python generate_answer_cards.py
```

指定 Excel 和输出目录：

```bash
python generate_answer_cards.py --excel input/新的学生信息.xlsx --out output_新批次
```

预先印上考试名称：

```bash
python generate_answer_cards.py --exam-name "高三数学阶段测试"
```

---

## 6. 输出文件位置

### 网页模式

每次网页生成会创建一个独立任务目录：

```text
jobs/<任务ID>/output/
├── 批量答题卡_xx人合并版.pdf
├── 批量答题卡_单人版.zip
├── 第一页预览.png
└── 单人版/
    ├── 01_识别号_姓名.pdf
    ├── 02_识别号_姓名.pdf
    └── ...
```

网页上的下载按钮会直接指向对应文件。

### 命令行模式

默认输出到：

```text
output/
```

---

## 7. 部署到自己的服务器

如果只给自己或内部人员使用，可以直接运行：

```bash
python -m uvicorn app:app --host 0.0.0.0 --port 8000
```

然后通过服务器 IP 或域名访问。

如果要长期稳定运行，建议使用：

- Nginx 反向代理
- systemd 或 pm2 管理进程
- 定期清理 `jobs/` 里的历史任务文件

---

## 8. 修改前端样式

前端全部在一个文件中：

```text
static/index.html
```

可以直接修改 HTML、CSS、JS。

当前视觉风格：

- 背景：浅蓝灰
- 主面板：白色
- 按钮和强调色：蓝色
- 布局：左侧上传，右侧流程与下载结果

---

## 9. 修改 PDF 坐标

如果未来模板改版，主要修改：

```text
config.py
```

重点配置项：

- `white_masks`：白色遮罩区域
- `exam_title`：考试名称和横线位置
- `student_text`：姓名和识别号位置
- `qr`：二维码位置和大小
- `locator_squares`：四个小黑色定位点位置和大小

坐标说明：

- x：距离页面左侧
- y：距离页面顶部
- 单位：PDF pt 点

---

## 10. 常见问题

### 生成失败，提示找不到姓名或识别号列

请检查 Excel 第一行表头。默认需要能匹配 `姓名` 和 `识别号`。

### 网页上传后没有反应

请先看终端窗口是否有报错。常见原因是依赖未安装完整，可以重新运行：

```bash
pip install -r requirements.txt
```

### 二维码内容是什么？

二维码内容就是该学生的识别号，例如：

```text
3S001
```

### 是否可以替换模板？

可以。替换：

```text
template/answer_card_template.pdf
```

但如果新版模板位置发生变化，需要同步调整 `config.py` 中的坐标。
