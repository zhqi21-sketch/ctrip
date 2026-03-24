# 携程签证订单助手 — 使用说明

## 📁 文件结构
```
ctrip-scraper/
  start.bat     ← 双击这个启动！
  server.js     ← 后端程序（不需要改）
  index.html    ← 可视化界面（不需要改）
  package.json  ← 配置文件（不需要改）
```

## 🚀 使用步骤

### 第一步：用调试模式打开 Chrome
每次使用前，先关闭所有 Chrome 窗口，然后运行：

**方法A（推荐）**：在命令提示符输入：
```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
```

**方法B**：创建一个 chrome-debug.bat 文件，内容：
```
"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
```
以后双击这个 bat 文件就行。

### 第二步：登录携程后台
在打开的 Chrome 里登录携程，并打开订单列表页面：
```
https://vbooking.ctrip.com/order/visaOrderList?obsolete=1
```

### 第三步：启动抓取工具
双击 `start.bat`，会自动安装依赖并打开操作界面（http://localhost:3333）

### 第四步：开始抓取
点击界面上的「▶ 开始抓取」按钮，程序会：
- 自动翻页抓取所有订单
- 实时显示进度和日志
- 进入每个订单详情页提取出行人信息
- 点击查看子订单状态

### 第五步：导出 Excel
抓取完成后点击「⬇ 导出 Excel」，文件自动下载。

## ⚠ 注意事项
- Chrome 必须用调试模式（--remote-debugging-port=9222）打开才能连接
- 抓取过程中不要手动操作已打开的 Chrome 浏览器
- 每次使用前确认已登录携程后台
