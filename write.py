import xlwt
from datetime import datetime

class WriteExcel(object):
    def WriteDocEx(self,**lstProd):
        font0 = xlwt.Font()
        font0.name = 'Times New Roman'
        #font0.colour_index = 2
        font0.bold = True

        style0 = xlwt.XFStyle()
        style0.font = font0

        style1 = xlwt.XFStyle()
        style1.num_format_str = 'D-MMM-YY'

        wb = xlwt.Workbook()
        ws = wb.add_sheet('Sheet 1')

        iter=0
        for keys,values in lstProd.items():
            ws.write(iter, 0, keys, style0)
            ws.write(iter, 1, values, style0)
            iter+=1

        wb.save('choiceProd.xls')
