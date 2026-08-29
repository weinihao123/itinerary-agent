# 行程记录与分析智能体

把口语化行程描述解析为结构化记录，支持派生天数计算、讲师层级确认、编辑删除、Excel 导出与多维看板分析。

## 功能

- 口语录入：粘贴一段口语描述，自动抽取运营商、省/市、项目类型、执行区间、出发/返程日期，缺失项标黄追问。
- 派生计算：执行天数、花费时间按含首尾口径自动推算（例如 8月1日到8月3日 = 3天）。
- 讲师层级：同项目逐期加一，文本中明确提到等级时优先识别，保存前可改。
- 编辑删除：记录可改任意字段，删除为软删除（可恢复）。
- Excel 导出：明细 + 汇总两表。
- 看板：按周/月/年切换，维度（省份/地市/类型/运营商）独立汇总，指标可选出行次数/执行天数/花费时间。

## 本地运行

```bash
cd itinerary-agent
python app.py
# 浏览器打开 http://127.0.0.1:8788
```

## 云端部署

代码已适配云端：服务读取 `PORT` 环境变量并绑定 `0.0.0.0`，仓库含 `requirements.txt`、`Procfile`、`Dockerfile`。

### 方式一 Hugging Face Spaces（免费、稳定）

1. 登录 huggingface.co，新建 Space，SDK 选 Docker。
2. 在 Space 设置中连接本 GitHub 仓库，或把仓库内容推到 Space。
3. 端口保持默认 7860，平台自动注入 `PORT`。
4. 部署完成后获得 `https://<用户名>-<space名>.hf.space` 公网地址。

### 方式二 Render（免费 Web 服务，15 分钟无活动会休眠）

1. 登录 render.com，新建 Web Service，连接本 GitHub 仓库。
2. Build Command：`pip install -r requirements.txt`；Start Command：`python app.py`。
3. Render 自动注入 `PORT`，服务监听该端口。
4. 部署完成后获得 `https://<服务名>.onrender.com` 公网地址。

## 数据持久说明

数据存于 `data/trips.csv`。Hugging Face Spaces 的文件系统在 Space 生命周期内持久；Render 免费版磁盘为临时（每次部署重置），如需持久数据建议挂载 Render Disk 或后续接入数据库。
