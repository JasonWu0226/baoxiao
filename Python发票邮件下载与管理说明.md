# Python 发票邮件下载与管理说明

## 文件

- `invoice_mail_downloader.py`：主程序。
- `invoice_mail_config.example.json`：配置模板。
- `invoice_mail_config.json`：首次运行时自动从模板生成，本地使用配置。
- `requirements_invoice_mail.txt`：PDF 文本读取依赖。
- `运行Python发票邮件下载.bat`：双击运行入口。

## 安装依赖

```powershell
pip install -r requirements_invoice_mail.txt
```

程序只依赖 `pypdf` 读取 PDF 文本；如果不安装，也能下载和哈希查重，但无法从 PDF 提取发票号、金额、开票日期。

## 配置

首次双击 `运行Python发票邮件下载.bat` 会生成 `invoice_mail_config.json`。

建议不要把邮箱授权码直接写进配置文件，可以在 PowerShell 里设置环境变量：

```powershell
$env:INVOICE_MAIL_PASSWORD="你的邮箱授权码"
python invoice_mail_downloader.py --config invoice_mail_config.json
```

也可以把授权码写到配置里的 `email.password`，但不推荐长期保存明文密码。

## 处理逻辑

1. 从 `start_time` 之后扫描邮箱。
2. 对每封邮件记录：
   - 邮件 ID
   - 邮件主题
   - 发件人
   - 接收时间
   - 是否含附件
   - 是否为发票邮件
   - 处理状态
   - 下载状态
   - 异常原因
3. 只有标题或正文命中发票关键词的邮件才下载附件。
4. 附件优先保留 PDF；同一封邮件同时有 PDF 和图片时，只下载 PDF。
5. 文件统一保存为：

```text
发票_YYYYMMDD_发件人_唯一标识.pdf
```

6. 查重依据：
   - 文件 MD5
   - 发票号码
   - 发票金额
   - 开票日期

## 输出

- `规则/发票邮件处理记录.json`：完整机器记录，支持增量处理和重新处理。
- `输出/发票邮件处理记录.csv`：人工查询表。
- `所有发票`：下载后的发票文件。

## 状态

- `正常已下载`
- `非发票邮件`
- `重复发票`
- `异常`

下载状态包含：

- `正常已下载`
- `重复已忽略`
- `异常`

异常原因包含：

- `无附件`
- `无关键词`
- `无 PDF`
- `格式错误`
- `下载失败`
- `格式不支持`
- `无发票号`
- `发票日期早于开始时间`

## 常用命令

正常增量处理：

```powershell
python invoice_mail_downloader.py --config invoice_mail_config.json
```

只重新处理异常邮件：

```powershell
python invoice_mail_downloader.py --config invoice_mail_config.json --reprocess-abnormal
```

重新处理指定邮件：

```powershell
python invoice_mail_downloader.py --config invoice_mail_config.json --reprocess-message INBOX:12345
```

只根据 JSON 重新导出 CSV：

```powershell
python invoice_mail_downloader.py --config invoice_mail_config.json --export-csv
```
