import pandas as pd

# 初期的准备工作

filepath = "19机电4班result.xls"
writer = pd.ExcelWriter(filepath)
df_2 = pd.DataFrame(pd.read_excel('table.xls',sheet_name='Sheet1'))

# 中断控制LED闪烁

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='中断控制LED闪烁'))
df_result1 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result1.sort_values(by='序号',inplace=True)
df_result1.to_excel(writer,sheet_name='中断控制LED闪烁',index=False)

# 花样霓虹灯

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='花样霓虹灯'))
df_result2 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result2.sort_values(by='序号',inplace=True)
df_result2.to_excel(excel_writer=writer,index=False,sheet_name='花样霓虹灯')

# 8个LED闪烁

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='8个LED闪烁'))
df_result3 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result3.sort_values(by='序号',inplace=True)
df_result3.to_excel(excel_writer=writer,index=False,sheet_name='8个LED闪烁')

# LED点阵姓名显示

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='LED点阵姓名显示'))
df_result4 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result4.sort_values(by='序号',inplace=True)
df_result4.to_excel(excel_writer=writer,index=False,sheet_name='LED点阵姓名显示')

# 库函数控制流水灯

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='库函数控制流水灯'))
df_result5 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result5.sort_values(by='序号',inplace=True)
df_result5.to_excel(excel_writer=writer,index=False,sheet_name='库函数控制流水灯')

# 简易计数报警器

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='简易计数报警器'))
df_result6 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result6.sort_values(by='序号',inplace=True)
df_result6.to_excel(excel_writer=writer,index=False,sheet_name='简易计数报警器')

# LED动态显示自己生日

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='LED动态显示自己生日'))
df_result7 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result7.sort_values(by='序号',inplace=True)
df_result7.to_excel(excel_writer=writer,index=False,sheet_name='LED动态显示自己生日')

# LED显示矩阵键盘按键号

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='LED显示矩阵键盘按键号'))
df_result8 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result8.sort_values(by='序号',inplace=True)
df_result8.to_excel(excel_writer=writer,index=False,sheet_name='LED显示矩阵键盘按键号')

# 综合

df_1 = pd.DataFrame(pd.read_excel('19机电4班课内实践成绩(空).xls',sheet_name='综合'))
df_result9 = pd.merge(df_2,df_1,how='left',on=["成绩"])
df_result9.sort_values(by='序号',inplace=True)
df_result9.to_excel(excel_writer=writer,index=False,sheet_name='综合')

writer.save()
writer.close()