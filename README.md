# fortify_format
该脚本用于格式化fortify导出报告的风险路径。
# 环境要求 #
- python3.x

# 模板导出要求 #
- BIRT report -> 模板选项 -> Developer WorkBook -> I ssue Filter Settings(-> filter -> 'anlysis is exploit') -> Format -> XLS
- 导出的原生xls另存为[Excel工作薄(*.xlsx)]或[Excel 97-2003工作薄(*.xls)],命名123或其它任意名字


# 使用方法
- 放置在报告同目录下
- 打开cmd,进入报告目录
## python环境运行
```python 
python -m pip install -r requires.txt
excel_format5.0.py [xxx.xlsx | xxx.xls , default=123.xlsx]
```
## 无python环境运行
```
excel_format5.0.exe [xxx.xlsx | xxx.xls , default=123.xlsx]
```
- 例如：excel_format5.0.exe 1.xlsx

# 注意事项：
- 建议在cmd下运行，可视化进度
- p3默认安装xlrd2，其不支持xlsx，使用1.2.0版本支持。
- 包中*.xls(x)用于脚本测试

# 版本历史
- excel_format5.1.exe
· 修复路径去重bug
· 若干优化
- excel_format5.0.exe
· 缺陷点去重检查
· 新的输出格式
- excel_format4.0-pack.exe
· 更新支持的模板：Developer WorkBook,不再支持CWE Top25 2019
- excel_format3.0-pack.exe
· 支持输入xlsx和xls类型
· 新的算法
· 新的格式
· 新的功能
- excel_format2.0.exe
· 支持自定义文件名，缺省为123.xls，cmd下执行：excel_format2.0.exe [example].xls
- excel_format1.0.exe
· 默认待格式化文件名为123.xls，直接运行即可



