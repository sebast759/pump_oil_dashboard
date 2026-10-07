@echo off
rem Refresh brent_cache.csv from Investing.com (blocked on GitHub's IPs) and push it.
rem Pushing triggers the GitHub workflow, which rebuilds and deploys the site.
cd /d "%~dp0"
(
  echo === %date% %time% ===
  "C:\Users\sebas\AppData\Local\Programs\Python\Python312\python.exe" generate_oil_dashboard.py --output site/index.html || exit /b 1
  git diff --quiet brent_cache.csv && (echo brent_cache.csv unchanged & exit /b 0)
  git add brent_cache.csv
  git commit -m "Refresh Brent cache" -m "Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
  git pull --rebase
  git push
) >> refresh_brent.log 2>&1
