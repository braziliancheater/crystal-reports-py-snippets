# brazilian
import sys
from win32com.client import Dispatch

def get_params(self):
	return self.ParameterFields

app = Dispatch('CrystalRunTime.Application')
print(app)
rep = app.OpenReport("./reports/<file>.rpt")
print(rep)
tbl = rep.Database.Tables.Item(1)
print(tbl)

prop = tbl.ConnectionProperties('Password')
prop.Value = "<password of the database>"
prop = tbl.ConnectionProperties('Database')
prop.Value = '<database name>'

rep.ExportOptions.FormatType = 29 # 1 - rpt | 31 - pdf | 29 - xls | 7 - csv | 30 - xls data only | for more: https://www.tek-tips.com/faqs.cfm?fid=3331
rep.ExportOptions.DestinationType = 1
rep.ExportOptions.DiskFileName = "<file location in disk>"
rep.Export(False)
rep.PrintOut(promptUser=False) # if anything else other than this, it just errors out...

# params
#params = rep.ParameterFields
#print(params)
#p1 = params(1)
#print(p1)
#p2 = params(2)
#print(p2)
#p3 = params(3)
#print(p3)
