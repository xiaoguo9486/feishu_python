# 项目更新日志 (Changelog)



## [V4] - 2026-06-22 -增加报告生成后自动邮件通知功能+报告直接输出到光伏组目录

更新内容：在现有基础上新增【生成报告后自动邮件通知】的功能。发送邮件复用之前的是【pthoto_edit.py】中的【def send_mail(smtp_server, smtp_port, user, password, to_addrs, subject, body):】函数

​	原来程序中业务处理对应的excel名称是【YH光伏巡检-照片巡检_xg简易自动巡检系统_巡检记录-照片xg (2).xlsx】，现在是【DH光伏巡检-逆变器+并网点_xg简易自动巡检系统_总-昨日.xlsx】。更新程序中excel文件的匹配规则：EXCEL_PATTERN

更新文件：【generate_report.py】



更新内容：报告直接输出到光伏组目录

```
# OUTPUT_FOLDER = "../output/生成的报告/"       # 本机测试环境的路径
OUTPUT_FOLDER = "E:/光伏运维/01流水资料/自动生成的报告/"       # 本机测试环境的路径
```

更新文件：【generate_report.py】



## [V4] - 2026-06-22 - 修复excel文件名问题

更新内容：原来程序中业务处理对应的excel名称是【YH光伏巡检-照片巡检_xg简易自动巡检系统_巡检记录-照片xg (2).xlsx】，现在是【DH光伏巡检-逆变器+并网点_xg简易自动巡检系统_总-昨日.xlsx】。更新程序中excel文件的匹配规则：EXCEL_PATTERN

更新文件：【generate_report.py】、【生成巡检报告V1-20260618.exe】



## [历史] - 2026-06-22以前

$ git log --oneline
d84711d (HEAD -> main, origin/main) V4生成高压低压巡检报告功能+发送邮件功能
676cca8 V3实现excel当行的高压模版报告导出
a1cb182 处理完成后自动发送邮件报告
0fcb47d V2发送邮件-未完成
b2c2715 实现了照片处理功能
3ade933 首次提交

### d84711d (HEAD -> main, origin/main) V4生成高压低压巡检报告功能+发送邮件功能

Date:   Tue Jun 16 13:26:55 2026 +0800

完全实现生成高压低压巡检报告功能，对应的文档有【generate_report.py】、【光伏巡检报告模板-低压.docx】、【光伏巡检报告模板-高压.docx】、【巡检照片整理工具-使用说明书-20260617.docx】、【生成巡检报告.exe】

【generate_report.py】是源程序，

【光伏巡检报告模板-低压.docx】、【光伏巡检报告模板-高压.docx】是程序会去调用的模板

【巡检照片整理工具-使用说明书-20260617.docx】是说明书

【生成巡检报告.exe】是编译后的文件



### 676cca8 V3实现excel单行的高压模版报告导出

Date:   Thu Jun 11 13:28:42 2026 +0800

### a1cb182 处理完成后自动发送邮件报告

Date:   Tue Jun 9 14:46:11 2026 +0800

### 0fcb47d V2发送邮件-未完成

Date:   Fri Jun 5 16:19:00 2026 +0800

### b2c2715 实现了照片处理功能

Date:   Fri Jun 5 16:02:19 2026 +0800

### 3ade933 首次提交

Date:   Fri Jun 5 14:28:11 2026 +0800


