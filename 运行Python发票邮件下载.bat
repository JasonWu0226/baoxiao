@echo off
cd /d "%~dp0"
if not exist invoice_mail_config.json (
  python invoice_mail_downloader.py --init-config
  echo.
  echo 已创建 invoice_mail_config.json，请先填写邮箱账号、授权码或设置 INVOICE_MAIL_PASSWORD。
  pause
  exit /b
)
python invoice_mail_downloader.py --config invoice_mail_config.json
pause
