# 屏幕指定区域自动截图 → OCR → Excel → 自动翻页

Win7 x64 可用:自动截取屏幕指定区域 → 用 **Umi-OCR** 识别文字 → 逐行写入 Excel → 点击"下一页"按钮 → 循环 N 次。

全程只需用户:双击 `ScreenOCR.exe` → 输入循环次数 → 框选截图区域 → 点一下"下一页"按钮位置 → 选择 Excel 保存文件 → 点"开始抓取"。

## 一、目录结构

```
screen-ocr-excel/
├── app/
│   └── main.py                # 主程序源码
├── build/
│   └── build_win.bat          # Windows 一键构建脚本(自动合并 Umi-OCR 引擎)
├── .github/workflows/build.yml # GitHub Actions 自动构建(推仓库即可, 无需本机装东西)
├── requirements.txt
└── README.md
```

## 二、构建 exe(任选其一)

> 说明:exe 必须在一台 **Windows** 机器上用 Python 编译(macOS/Linux 无法交叉编译)。
> 如果这台 Win7 机器本身就装了 Python 3.8, 也可以直接在它上面构建。

**方式 A —— 任意 Windows 电脑双击脚本**
1. 安装 Python(构建机建议 Python **3.8**, 只有 3.8 编出的 exe 才能装到 Win7)。
2. 双击 `build\build_win.bat`, 等待完成。
3. 产物:`dist\ScreenOCR\ScreenOCR.exe`。

**方式 B —— GitHub Actions(无需本地装任何东西)**
1. 把本目录推进一个 GitHub 仓库(`.github/workflows/build.yml` 在仓库内)。
2. 仓库 `Actions` 页面 → 运行 `build-exe` 工作流。
3. 完成后在 Actions 运行页右下角 **Artifacts** 下载 `ScreenOCR.zip`, 解压得到 exe。

## 三、部署到 Win7(最终用户操作, 只需一步)

构建产物是 **自包含完整包**:`dist\ScreenOCR\` 文件夹里已经包含 `ScreenOCR.exe` 和 `Umi-OCR` 引擎(构建时自动从官方下载合并)。最终用户只要:

1. 下载 `ScreenOCR.zip` → 解压(得到 `ScreenOCR` 文件夹)。
2. 整个文件夹拷到 Win7 电脑, 双击 `ScreenOCR.exe` 即可。

程序会自动在后台拉起 Umi-OCR 引擎, 无需用户手动下载或拷贝任何东西。

> 若网络/速度受限(构建时下载这个 100MB 引擎较慢), 也可只构建不含引擎的 exe,
> 并把官方 Umi-OCR(Rapid 版)解压后的 `Umi-OCR` 文件夹放到 `ScreenOCR.exe` 旁边, 效果相同。

## 四、使用说明

1. 双击 `ScreenOCR.exe`, 输入**循环次数**(要抓多少页)。
2. 点 **① 选择截图区域**:按住左键拖出一个框, 松开即选定。
3. 点 **② 定位下一页按钮**:鼠标移到"下一页"按钮上单击一下。
4. 点 **③ 选择Excel保存文件**:文件已存在则接着往后面追加, 不存在则新建。
5. 点 **▶ 开始抓取**。

流程自动循环:截图 → OCR → 写入 Excel → 点击下一页 → 截图 → ……
Excel 每页三列:`页码 | 行号 | 识别内容`(每行文字一个单元格, 适合表格 / 字段名+值)。

**紧急停止**:运行中把鼠标甩到屏幕左上角(0,0), 立即中止。

## 五、常见问题

- **提示"未找到 OCR 引擎"**:一般不会发生(引擎已随包自带)。若出现, 确认 `ScreenOCR.exe` 旁边有 `Umi-OCR` 文件夹; 或者手动打开 Umi-OCR 再点开始。
- **若用户自己已手动打开 Umi-OCR, 程序会直接复用, 不再重复启动。**
- **识别乱码/语言不对**:程序会自动使用引擎默认语言(简体中文)。如需英文等, 在 Umi-OCR 界面"文字识别 → 引擎"里切换语言库; 程序会自动跟随, 无需改代码。
- **报"引擎返回错误 code=xxx"**:多为引擎启动初期未就绪或参数不适配, 程序已内置重试与自适应降级, 一般自动恢复。若持续出现, 打开 Umi-OCR 看其状态页, 或确认 1224 端口没有被其他程序占用。
- **页面等太久/太快**:修改 `main.py` 顶部 `PAGE_WAIT`(点下一页后等待秒数)。
- **截图框别盖住"下一页"按钮**, 否则点击会被自己触发的前端干扰。
- **Win7 上双击无反应**:装过 kb/缺少运行库时, 查看 exe 旁边的 `run.log`(如有)。
- **Umi-OCR 关闭后不退出**:Umi-OCR 官方说明, 关闭软件时若 HTTP 连接未断会导致其进程残留, 属正常, 任务管理器结束即可。

## 六、技术栈

- `pyautogui` 截屏(`region=...`)与鼠标点击
- 调 Umi-OCR 本地 HTTP 接口 `POST /api/ocr`(base64 传图)
- `openpyxl` 写 Excel
- `tkinter` 全屏框选/取点 + 主界面
- `PyInstaller 5.x (Python 3.8)` 打包, 保证 Win7 x64 兼容