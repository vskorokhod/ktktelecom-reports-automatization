import subprocess
from datetime import date, timedelta, datetime


def _report_composition():
    today = datetime.strftime(date.today(), '%d%m%Y')
    day, month, year = today[0:2], today[2:4], today[4:8]
    day_modified = day[1:] if day < '10' else day
    month_modified = month[1:] if month < '10' else month
    
    
    commands = ['"C:\Program Files (x86)\WinSCP\winscp.com" /ini=nul /command "open sftp://root:cnp200@HW@10.10.180.2 -hostkey=""ssh-rsa 2048 b9:17:dd:3b:f9:ac:98:6e:6a:6e:da:0a:f0:d5:ba:5f""" "call cd /opt/HUAWEI/cgp/workshop/omu/share/export/nefile/ne_5/VLROutPut/MSS01SMF_User\ data_{}-{}-{}*" "get ""*.txt"" ""C:\Reports\VLR Badashev\\""" "close" "exit"'.format(year, month_modified, day_modified),
                'if not exist "G:\Share\ТД\Мониторинг\Отчёты\VLR_Бадашеву\{0}.{1}\\" mkdir "G:\Share\ТД\Мониторинг\Отчёты\VLR_Бадашеву\{0}.{1}\\"'.format(month, year),
                'move "C:\Reports\VLR Badashev\*.txt" "G:\Share\ТД\Мониторинг\Отчёты\VLR_Бадашеву\{}.{}\\"'.format(month, year)]
             
    for command in commands:
        subprocess.run(command, shell=True)

        
if __name__ == '__main__':
    _report_composition()