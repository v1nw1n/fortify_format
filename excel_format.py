import xlrd.book
import xlwt.Workbook,xlwt.Style,xlwt.Worksheet
import os.path
import sys
import time
import pandas as pd
import re

'''
v5.1
1.算法改进（去重检查）,修复去重bug
2.其他优化
'''

#pyinstaller打包(pyinstaller -F excel_format.py -i f.ico)使用
#workdir = os.path.dirname(os.path.realpath(sys.executable))+'\\'
#脚本环境使用
#workdir = os.path.dirname(__file__)+'\\'
workdir = os.getcwd()+'\\'
def process():
  excelfile = '' 
  #自定义输入文件
  if len(sys.argv) == 2:
    excelfile = sys.argv[1]
  else:
    excelfile = '123.xlsx'
  #读取输入文件
  try:
    book = xlrd.open_workbook(workdir+excelfile)
  except xlrd.biffh.XLRDError:
    print('[info]'+"格式不正确，原始导出报告需另存为xlsx(建议)或xls")
    return
  else:
    sheet_Report = book.sheet_by_name('Report')
  #获取行数
  rows = sheet_Report.nrows
  #创建输出文件workbook实例
  outfile = xlwt.Workbook()  
  #添加sheet1
  sheet_1 = outfile.add_sheet(u'sheet_1', cell_overwrite_ok=True)
  #写入表头
  table_head = [u'漏洞名称（总计）',u'序号', u'漏洞路径']
  for i in range(0, len(table_head)):
    sheet_1.write(0, i, table_head[i])
  #转储数据变量：用于构造dataFrame完成聚合
  vuln_data = {
    'path':[],
    'point':[]
  }
  #新sheet行号
  new_row=1
  #漏洞名称
  vuln_name=[]
  #缺陷点计数
  vuln_sum=[]
  #vuln_sum、vuln_name索引对应
  #获取build id,替换\/:*?"<>|
  build_id = sheet_Report.row_values(2)[0]
  if build_id == "":
    build_id = "Fortify_Report"
  re.sub("\\/:\*\?\"<>|", build_id, "_")
  #统计序号
  no=0
  #创建第一列的样式
  XFstyle = xlwt.Style.easyxf("align: vert centre")
  #处理数据
  start = time.perf_counter()
  for r in range(0, rows):
    #打印进度
    num = r*100//2//rows
    if r == rows-1:
      loading = "\r[%3.3s%%]: |%-50s|\n" % (r/rows*100+1, '=' * (num+1))
    else:
      loading = "\r[%3.3s%%]: |%-50s|" % (r/rows*100+1, '=' * (num+1))
    print(loading, end='', flush=True)
    #获取行数据
    r_values = sheet_Report.row_values(r)
    if "Total" == r_values[0] or (r == (rows-1)) :
      #获取漏洞名称
      if len(vuln_name) == 0 or (r != (rows-1) and (len(vuln_name) != len(vuln_sum))):
        vuln_name.append(sheet_Report.row_values(r-1)[0])
      #判断是否完成一类漏洞的遍历
      if no != 0:
        #记录缺陷点数量
        vuln_sum.append(no)
        #重置缺陷点计数器
        no=0
        #聚合数据
        #构造dataframe
        df = pd.DataFrame(vuln_data)
        df = df.groupby("path").apply(lambda x: '、'.join(x.point)).to_frame("point").reset_index()
        #写入数据
        top_row=new_row
        for df_r in range(len(df)):
          sheet_1.write(new_row,0,vuln_name[len(vuln_sum)-1])
          sheet_1.write(new_row,1,str(df_r+1))
          sheet_1.write(new_row,2,df.loc[df_r][0]+' '+df.loc[df_r][1]+'行')
          new_row += 1
          #第一列按类型合并单元格
          sheet_1.write_merge(top_row, new_row-1, 0, 0, vuln_name[len(vuln_sum)-1]+"（"+str(vuln_sum[-1])+"）",XFstyle)
        vuln_data['path'] = []
        vuln_data['point'] = []
      continue
    if str(sheet_Report.row_values(r)[0]) == "Issue Details":
      #缺陷路径格式标准化
      path_point = sheet_Report.row_values(r-1)[0].replace(", line ","*").replace(" ("+vuln_name[-1]+")","").split("*")
      if path_point[1] not in vuln_data['point'] and path_point[0] not in vuln_data['path']:
        #5.0 update:缺陷点去重
        vuln_data['point'].append(path_point[1])
        vuln_data['path'].append(path_point[0])
        no += 1
  end = time.perf_counter()
  print('[info]转换成功，用时：{:.2f}s，报告路径:'.format(end-start))
  #输出xlsx将导致文件损坏
  outfile_name = workdir + build_id+'_'+time.strftime("%Y%m%d%H%M%S",time.localtime())+'.xls'
  print('--'+outfile_name)
  print('[info]安全风险统计:')
  for a in range(len(vuln_sum)):
    print('--'+vuln_name[a]+":"+str(vuln_sum[a]))
  sheet_1.col(0).width = 256 * 60
  sheet_1.col(2).width = 256 * 150
  outfile.save(outfile_name)
process()
