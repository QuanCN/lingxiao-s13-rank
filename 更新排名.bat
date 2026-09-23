@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/3] 正在根据 data.xlsx 生成排名页...
python build.py
if errorlevel 1 (
  echo 失败：请确认已安装 Python，并执行 pip install -r requirements.txt
  pause
  exit /b 1
)
echo [2/3] 提交并推送到 GitHub...
git add data.xlsx index.html
git -c user.email="quancn@users.noreply.github.com" -c user.name="QuanCN" commit -m "Update ranking from data.xlsx"
if errorlevel 1 (
  echo 没有变更可提交（表格可能没改）
) else (
  git push
  if errorlevel 1 (
    echo 推送失败，请检查网络或登录状态
    pause
    exit /b 1
  )
)
echo [3/3] 完成。约1分钟后刷新：
echo https://quancn.github.io/lingxiao-s13-rank/
pause