import shutil
from time import sleep
from datetime import date
today = date.today()
d1 = today.strftime("%d-%m-%Y")

print(d1)
 
shutil.copyfile(r'\\10.199.199.35\soc team\SOC_Daily Report\2023\Feb\Nexthink feb Report.xlsx', r"\\10.199.199.35\soc team\SOC_Daily Report\2023\Feb\Backup\Nxthink/"+d1+".xlsx")
sleep(2)
shutil.copyfile(r'\\10.199.199.35\soc team\SOC_Daily Report\2023\Feb\BMC February Report.xlsx', r"\\10.199.199.35\soc team\SOC_Daily Report\2023\Feb\Backup\BMC/"+d1+".xlsx")
