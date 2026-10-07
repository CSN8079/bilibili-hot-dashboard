# 哔哩哔哩热门视频定时导出与实时看板

程序默认启动后立即抓取一次，之后每 10 分钟抓取哔哩哔哩全站热门视频，并同时更新 Excel 和 HTML 看板。

## 数据不会因为 Excel 占用而丢失

每轮抓取结果会先写入 `pending_updates.jsonl`，然后才尝试更新 Excel。

- Excel 正常：队列中的数据写入 Excel，写入成功后清空队列。
- Excel 被打开或写入失败：该轮数据继续保存在队列中。
- 后续周期：程序自动把尚未写入的全部快照一次补写到 Excel。
- 手动补写：运行 `run.bat --flush-pending`。
- 因此，即使下一轮抓取前一直打开 Excel，历史快照也不会丢失，关闭 Excel 后会自动补写。

## Excel 被占用时的系统提醒

当检测到 Excel 文件被占用时，程序会弹出 Windows 系统提示：

```text
检测到 Excel 文件正在被占用，本次无法写入。
请关闭 Excel 进行更新。
数据已安全保存在待补写队列中，不会丢失。
```

提示频率可以在 `config.ini` 中调整：

```ini
[notifications]
enabled = true
cooldown_seconds = 600
```

## 排名变化、封面与转发量

- Excel 在“排名”后增加“排名变化”列。
- Excel 在“视频名称”前增加“封面链接”列。
- Excel 在“收藏量”后增加“转发量”列。
- 实时看板在视频标题前显示封面缩略图。
- 看板表格在收藏量后显示转发量，并支持按转发量排序。
- 排名上升使用红色 `↑`，排名下降使用绿色 `↓`，符合常用股票颜色逻辑。
- 首次进入榜单的视频显示 `NEW`。

## 分类分布与播放量对比

- 看板提供“视频分类分布”饼图。
- 看板提供“分类播放量对比”横向柱状图。
- 分类筛选由下拉框改为可折叠的点击标签。
- 默认折叠；展开后一次性显示全部选项，不再使用滚动条。
- 支持单独点击多个分类，也支持“全选”和“删除已选”。
- 搜索、分类筛选和排序可以组合使用。
- 饼图使用 3D 投影扇形柱体，包含透视顶面、分段侧壁和厚度阴影；鼠标移入扇区会抬升并显示分类、占比和播放量详情。
- 页面顶部提供“立即刷新数据”按钮，可以通知正在运行的抓取程序马上执行一轮抓取。

## 实时 HTML 看板

程序启动后会同时开启本地看板服务：

```text
http://127.0.0.1:8765/
```

可以：

- 双击 `bilibili_hot_dashboard.url` 打开。
- 双击 `view_dashboard.bat` 打开。
- 浏览器保持打开时，默认每 10 秒自动获取最新数据。
- 直接打开 `bilibili_hot_dashboard.html` 可以查看静态快照。
- 看板支持搜索、分区筛选、按播放/点赞/投币/收藏排序以及明暗主题。

如果默认端口 `8765` 被占用，程序会自动尝试后续端口，并在启动日志中显示实际网址。

## 文件说明

- `bilibili_hot_excel.py`：主程序。
- `snapshot_store.py`：数据队列和原子写入。
- `dashboard_server.py`：HTML 看板生成及本地服务。
- `dashboard_template.html`：看板样式和交互模板。
- `config.ini`：全部常用配置。
- `pending_updates.jsonl`：Excel 写入失败时的待补写数据。
- `bilibili_hot_dashboard.html`：可独立打开的 HTML 看板。
- `dashboard_data.json`：看板实时数据接口的数据文件。
- `bilibili_hot_dashboard.url`：实时看板快捷方式。
- `run.bat`：前台启动抓取和看板服务。
- `view_dashboard.bat`：打开实时看板。

## 运行与停止

双击 `run.bat` 前台运行，命令行窗口会持续显示抓取和补写日志。

停止：

1. 点击运行窗口。
2. 按 `Ctrl + C`。
3. 如果提示 `Terminate batch job (Y/N)?`，输入 `Y` 并回车。

## 常用配置

```ini
[schedule]
interval_minutes = 10

[storage]
pending_file = pending_updates.jsonl

[dashboard]
enabled = true
host = 127.0.0.1
port = 8765
refresh_seconds = 10
```

命令行示例：

```powershell
# 只抓取一次
python .\bilibili_hot_excel.py --once

# 修改间隔和保存数量
python .\bilibili_hot_excel.py --interval-minutes 5 --top-n 50

# 手动补写 Excel
python .\bilibili_hot_excel.py --flush-pending

# 启动但不提供实时看板服务
python .\bilibili_hot_excel.py --no-dashboard
```

公开排行榜通常不需要 Cookie。如确需设置，可在 `config.ini` 的 `[optional_auth]` 中配置，或设置环境变量 `BILIBILI_COOKIE`。