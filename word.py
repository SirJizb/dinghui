from win32com.client import Dispatch
from datetime import datetime
import os

zpath = os.getcwd() + '\\'
app = Dispatch('Word.Application')
# app.Visible = True
xlApp = Dispatch("Excel.Application")
# xlApp.Visible = True
xlBook = xlApp.Workbooks.Open(zpath + 'blog_list13.xlsx')

for i in range(5):
    name = xlBook.Worksheets('blog').Cells(i + 2, 1).Value
    line = str(xlBook.Worksheets('blog').Cells(i + 2, 2).Value)
    time = str(xlBook.Worksheets('blog').Cells(i + 2, 3).Value)
    content = str(xlBook.Worksheets('blog').Cells(i + 2, 4).Value)
    print(name, line, time, content)
    doc = app.Documents.Add(zpath + '博客标题.dotx')
    doc.Bookmarks("name").Range.Text = name
    doc.Bookmarks("line").Range.Text = line
    doc.Bookmarks("time").Range.Text = time
    doc.Bookmarks("content").Range.Text = content < br >　　 <br > doc.SaveAs(
        zpath + '1905-2 20194157 韩佳作  ' + time + '.doc')
　　print(zpath + '1905-2 20194157   ' + time + '.doc')
　　print(i)
1
app.Documents.Close()
app.Quit()
xlBook.Close()
xlApp.Quit()
print("运行结束！！")