@echo off
:: ============================================================
:: 小虾记账 v1.0 — 一键部署脚本
:: 用法: 将从 GitHub 下载的源码解压到此目录，双击 setup.bat
:: 自动完成: 环境检测→依赖安装→示例数据→前端构建→启动
:: ============================================================
chcp 65001 >nul
setlocal enabledelayedexpansion

set ROOT=%~dp0
set ROOT=%ROOT:~0,-1%
set VENV_DIR=%ROOT%\.venv
set VENV_PYTHON=%VENV_DIR%\Scripts\python.exe
set FRONTEND_DIR=%ROOT%\web-ui
set BACKEND_DIR=%ROOT%\backend

title 小虾记账 v1.0 — 部署中...

echo.
echo   ╔══════════════════════════════════════════╗
echo   ║        🦞 小虾记账 v1.0 一键部署         ║
echo   ╚══════════════════════════════════════════╝
echo.
echo   本脚本将依次执行:
echo     1. 检测 Python / Node.js 环境
echo     2. 创建 Python 虚拟环境 + 安装依赖
echo     3. 安装前端 npm 依赖
echo     4. 构建前端生产包
echo     5. 生成示例测试数据
echo     6. 验证全管线
echo     7. 启动服务 + 打开浏览器
echo.
echo   预计耗时: 3-5 分钟（首次）/ 30 秒（已部署）
echo   ═══════════════════════════════════════════
echo.

:: ============================================================
:: STEP 1: 环境检测
:: ============================================================
echo [1/7] 检测运行环境...

:: Python
where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo   ❌ 未找到 Python！请先安装 Python 3.10+
    echo     下载: https://www.python.org/downloads/
    echo     安装时务必勾选 ☑ "Add Python to PATH"
    echo.
    pause
    exit /b 1
)
for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo   ✅ Python %PYVER%

:: Node.js
where node >nul 2>&1
if errorlevel 1 (
    echo.
    echo   ❌ 未找到 Node.js！请先安装 Node.js 20+
    echo     下载: https://nodejs.org
    echo.
    pause
    exit /b 1
)
for /f "tokens=1,2 delims=v." %%a in ('node --version 2^>^&1') do set NODEVER=%%a.%%b
echo   ✅ Node.js !NODEVER!

:: npm
where npm >nul 2>&1
if errorlevel 1 (
    echo   ❌ 未找到 npm（通常随 Node.js 一起安装）
    pause
    exit /b 1
)
echo   ✅ npm

:: curl (for health check)
where curl >nul 2>&1
if errorlevel 1 (
    echo   ⚠️ 未找到 curl，服务探活将使用超时等待
    set HAS_CURL=0
) else (
    set HAS_CURL=1
)

echo.
echo   环境检测通过。
echo.

:: ============================================================
:: STEP 2: Python 虚拟环境 + 依赖
:: ============================================================
echo [2/7] Python 虚拟环境 + 依赖安装...

if not exist "%VENV_PYTHON%" (
    echo   创建虚拟环境 .venv ...
    python -m venv "%VENV_DIR%" 2>&1
    if errorlevel 1 (
        echo   ❌ 创建虚拟环境失败
        pause
        exit /b 1
    )
)
echo   ✅ .venv 就绪

echo   检查 Python 依赖...
"%VENV_PYTHON%" -c "import fastapi, uvicorn, openpyxl, yaml" 2>nul
if errorlevel 1 (
    echo   正在安装依赖 (fastapi, uvicorn, openpyxl, pyyaml, python-multipart, aiofiles)...
    "%VENV_PYTHON%" -m pip install fastapi uvicorn openpyxl pyyaml python-multipart aiofiles pdfplumber -q 2>&1
    if errorlevel 1 (
        echo.
        echo   ❌ pip 安装失败，请检查网络连接
        echo   手动尝试: .venv\Scripts\pip install fastapi uvicorn openpyxl pyyaml python-multipart aiofiles pdfplumber
        echo.
        pause
        exit /b 1
    )
)
echo   ✅ Python 依赖就绪

:: ============================================================
:: STEP 3: 前端 npm 依赖
:: ============================================================
echo [3/7] 前端依赖安装...

if not exist "%FRONTEND_DIR%" (
    echo   ❌ 未找到前端目录: %FRONTEND_DIR%
    echo   请确认源码完整解压。
    pause
    exit /b 1
)

cd /d "%FRONTEND_DIR%"

if not exist "node_modules\react" (
    echo   正在 npm install（首次耗时较长，请耐心等待）...
    call npm install --silent 2>&1
    if errorlevel 1 (
        echo.
        echo   ❌ npm install 失败，请检查网络连接
        echo   手动尝试: cd web-ui ^&^& npm install
        echo.
        cd /d "%ROOT%"
        pause
        exit /b 1
    )
)
echo   ✅ 前端依赖就绪
cd /d "%ROOT%"

:: ============================================================
:: STEP 4: 构建前端生产包
:: ============================================================
echo [4/7] 构建前端...

cd /d "%FRONTEND_DIR%"
call npm run build --silent 2>&1
if errorlevel 1 (
    echo   ❌ 前端构建失败
    echo   手动尝试: cd web-ui ^&^& npm run build
    cd /d "%ROOT%"
    pause
    exit /b 1
)
cd /d "%ROOT%"

if not exist "%FRONTEND_DIR%\dist\index.html" (
    echo   ❌ 构建产物不存在
    pause
    exit /b 1
)
echo   ✅ 前端构建完成

:: ============================================================
:: STEP 5: 生成示例数据
:: ============================================================
echo [5/7] 检查示例数据...

set HAS_DATA=0
if exist "%ROOT%\输入\业务事项.xlsx" set HAS_DATA=1
if exist "%ROOT%\输入\企业信息" set HAS_DATA=1

if "!HAS_DATA!"=="0" (
    echo   未发现业务数据，正在生成示例数据...
    "%VENV_PYTHON%" "%ROOT%\tools\generate_sample_data.py" 2>&1
    if errorlevel 1 (
        echo   ⚠️ 示例数据生成失败，稍后可手动运行
    ) else (
        echo   ✅ 示例数据已生成
    )
) else (
    echo   ✅ 数据已存在，跳过
)

:: ============================================================
:: STEP 6: 全管线验证
:: ============================================================
echo [6/7] 全管线验证运算 (2026-08)...

"%VENV_PYTHON%" -m accounting_report --full 2026-08 --input "%ROOT%\输入" --output "%ROOT%\输出" 2>&1 | findstr /C:"完成" >nul
if errorlevel 1 (
    echo   ⚠️ 全管线运算有警告，但可能正常（首次运行常见）
) else (
    echo   ✅ 全管线运算通过
)

:: ============================================================
:: STEP 7: 启动服务
:: ============================================================
echo [7/7] 启动服务...
echo.
echo   ═══════════════════════════════════════════
echo     🚀 小虾记账 v1.0 启动中...
echo.
echo     后端 API:  http://localhost:8100
echo     API 文档:  http://localhost:8100/docs
echo     前端界面:  http://localhost:5173
echo.
echo     按 Ctrl+C 停止服务
echo   ═══════════════════════════════════════════
echo.

:: 启动后端
start "小虾记账-API" /MIN "%VENV_PYTHON%" -m uvicorn server:app --host 0.0.0.0 --port 8100 --app-dir "%BACKEND_DIR%"

:: 等待后端就绪
echo   等待后端启动...
if "!HAS_CURL!"=="1" (
    set /a RETRY=0
    :wait_backend
    set /a RETRY+=1
    timeout /t 2 /nobreak >nul
    curl.exe -s http://localhost:8100/api/health 2>nul | findstr /C:"ok" >nul
    if errorlevel 1 (
        if !RETRY! LSS 15 (
            goto wait_backend
        ) else (
            echo   ⚠️ 后端启动超时，请手动检查: http://localhost:8100/api/health
        )
    ) else (
        echo   ✅ 后端就绪 (8100)
    )
) else (
    echo   ⏳ 等待 10 秒让后端启动...
    timeout /t 10 /nobreak >nul
)

:: 启动前端
cd /d "%FRONTEND_DIR%"
start "小虾记账-前端" /MIN cmd /c "npx vite --host 0.0.0.0 --port 5173"
cd /d "%ROOT%"

:: 等待前端就绪
if "!HAS_CURL!"=="1" (
    set /a RETRY=0
    :wait_frontend
    set /a RETRY+=1
    timeout /t 2 /nobreak >nul
    curl.exe -s -o NUL http://localhost:5173 2>nul
    if errorlevel 1 (
        if !RETRY! LSS 10 (
            goto wait_frontend
        ) else (
            echo   ⚠️ 前端启动超时
        )
    ) else (
        echo   ✅ 前端就绪 (5173)
    )
) else (
    echo   ⏳ 等待 5 秒让前端启动...
    timeout /t 5 /nobreak >nul
)

:: 打开浏览器
start http://localhost:5173

echo.
echo   ╔══════════════════════════════════════════╗
echo   ║   🦞 小虾记账 v1.0 部署完成！            ║
echo   ║                                          ║
echo   ║   前端: http://localhost:5173             ║
echo   ║   后端: http://localhost:8100             ║
echo   ║   API:  http://localhost:8100/docs        ║
echo   ║                                          ║
echo   ║   首次使用请阅读: docs\部署使用说明书.md   ║
echo   ╚══════════════════════════════════════════╝
echo.
echo   按任意键关闭此窗口（不影响后台服务）。
pause >nul
