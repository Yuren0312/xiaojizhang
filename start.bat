:: 小虾记账 v1.0 — 一键启动脚本（含依赖自检）
:: 同时启动后端 API 服务 + 前端开发服务器
@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set ROOT=%~dp0
set VENV_PYTHON=%ROOT%.venv\Scripts\python.exe
set FRONTEND_DIR=%ROOT%web-ui

echo.
echo  ?? 小虾记账 v1.0
echo ===========================================
echo.

:: ── 步骤 1: 检查 Python 虚拟环境 ──
echo [1/5] 检查 Python 虚拟环境...
if not exist "%VENV_PYTHON%" (
    echo    [!] 未找到 .venv，正在创建...
    python -m venv .venv
    if errorlevel 1 (
        echo    [错误] 创建虚拟环境失败，请确认已安装 Python 3.10+
        pause
        exit /b 1
    )
    echo    [?] 虚拟环境已创建
)
echo    [?] .venv 已就绪

:: ── 步骤 2: 安装/更新 Python 依赖 ──
echo [2/5] 检查 Python 依赖...
"%VENV_PYTHON%" -c "import fastapi, uvicorn, openpyxl, yaml" 2>nul
if errorlevel 1 (
    echo    [!] 正在安装依赖 (fastapi, uvicorn, openpyxl, pyyaml, python-multipart)...
    "%VENV_PYTHON%" -m pip install fastapi uvicorn openpyxl pyyaml python-multipart aiofiles pdfplumber -q
    if errorlevel 1 (
        echo    [错误] pip 安装失败，请检查网络连接
        pause
        exit /b 1
    )
    echo    [?] Python 依赖安装完成
) else (
    echo    [?] Python 依赖已就绪
)

:: ── 步骤 3: 检查 Node.js + 前端依赖 ──
echo [3/5] 检查 Node.js 环境...
where npm >nul 2>&1
if errorlevel 1 (
    echo    [错误] 未找到 npm，请先安装 Node.js: https://nodejs.org
    echo    推荐版本: Node.js 20+
    pause
    exit /b 1
)

if not exist "%FRONTEND_DIR%\node_modules\react" (
    echo    [!] 正在安装前端依赖 (npm install)...
    cd /d "%FRONTEND_DIR%"
    call npm install --silent
    if errorlevel 1 (
        echo    [错误] npm install 失败，请检查网络连接
        cd /d "%ROOT%"
        pause
        exit /b 1
    )
    cd /d "%ROOT%"
    echo    [?] 前端依赖安装完成
) else (
    echo    [?] 前端依赖已就绪
)

:: ── 步骤 4: 生成示例数据（若输入目录为空）─
echo [4/5] 检查数据文件...
set HAS_DATA=0
if exist "%ROOT%输入\业务事项.xlsx" set HAS_DATA=1
if exist "%ROOT%输入\企业信息" set HAS_DATA=1
if "!HAS_DATA!" == "0" (
    echo    [!] 未发现业务数据，正在生成示例数据...
    "%VENV_PYTHON%" tools\generate_sample_data.py
    if errorlevel 1 (
        echo    [警告] 示例数据生成失败，可稍后手动运行
    ) else (
        echo    [?] 示例数据已生成在 输入/ 目录
    )
) else (
    echo    [?] 输入数据已存在，跳过示例数据生成
)

:: ── 步骤 5: 启动服务 ──
echo [5/5] 启动服务...
echo.
echo    ? 后端 API: http://localhost:8100
echo    ? 前端界面: http://localhost:5173
echo.
echo    ?? 按 Ctrl+C 停止所有服务
echo ===========================================
echo.

:: 启动后端 (后台)
start "小虾记账-API" /MIN "%VENV_PYTHON%" -m uvicorn server:app --host 0.0.0.0 --port 8100 --app-dir backend

:: 等待后端就绪（最多 20 秒）
echo    ? 等待后端启动...
set /a RETRY=0
:wait_backend
set /a RETRY+=1
timeout /t 2 /nobreak >nul
curl.exe -s http://localhost:8100/api/health 2>nul | findstr /C:"ok" >nul
if errorlevel 1 (
    if !RETRY! LSS 10 (
        goto wait_backend
    ) else (
        echo    [警告] 后端启动超时，请手动检查
    )
) else (
    echo    [?] 后端已就绪
)

:: 启动前端
cd /d "%FRONTEND_DIR%"
start "小虾记账-前端" /MIN cmd /c "npm run dev 2>&1 | findstr /V """
cd /d "%ROOT%"

:: 等待前端就绪
echo    ? 等待前端启动...
set /a RETRY=0
:wait_frontend
set /a RETRY+=1
timeout /t 2 /nobreak >nul
curl.exe -s -o NUL http://localhost:5173 2>nul
if errorlevel 1 (
    if !RETRY! LSS 10 (
        goto wait_frontend
    ) else (
        echo    [警告] 前端启动超时
    )
) else (
    echo    [?] 前端已就绪
)

:: 打开浏览器
echo.
echo    ? 正在打开浏览器...
start http://localhost:5173

echo.
echo  ?? 小虾记账 v1.0 启动完成！
echo.
echo    后端: http://localhost:8100  (API docs: http://localhost:8100/docs)
echo    前端: http://localhost:5173
echo.
echo    按任意键关闭... (不影响后台服务)
pause >nul
