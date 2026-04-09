@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] 安装打包依赖...
python -m pip install -q pyinstaller pandas openpyxl selenium rapidfuzz

echo [2/3] 正在打包（约 1~3 分钟）...
pyinstaller --noconfirm --clean ^
  --onefile ^
  --console ^
  --name fx_ludan ^
  --add-data "std_keywords.txt;." ^
  --add-data "synonyms.json;." ^
  --hidden-import=pandas ^
  --hidden-import=openpyxl ^
  --hidden-import=openpyxl.cell._writer ^
  --hidden-import=rapidfuzz ^
  --collect-all selenium ^
  fx_ludan.py

if errorlevel 1 (
  echo 打包失败，请检查是否已安装 Python 3.10+ 并加入 PATH
  pause
  exit /b 1
)

echo [3/3] 完成。
echo.
echo 生成文件: dist\fx_ludan.exe
echo 请将 dist\fx_ludan.exe 复制到任意文件夹；可选在同目录放置自定义 std_keywords.txt / synonyms.json 覆盖内置词库。
echo 本机仍需安装 Google Chrome；若 chromedriver 不在 PATH，可设置环境变量 FX_CHROMEDRIVER 指向 chromedriver.exe
echo.
pause
