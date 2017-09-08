import subprocess
from datetime import date, timedelta, datetime


def _report_composition():
    today = datetime.strftime(date.today(), '%d.%m.%Y')
    
    commands = ['"C:\Program Files (x86)\WinSCP\winscp.com" /ini=nul /command "open sftp://ossuser:Changeme_123@10.10.180.65 -hostkey=""ssh-rsa 2048 61:5b:c1:1f:ef:64:75:52:53:11:b1:d5:3f:6b:ae:4b""" "get -delete ""/opt/oss/server/var/fileint/pm/NOC_reports/VLR_Sharkov/*.xlsx"" ""Z:\ТД\Мониторинг\Отчёты\Недельный отчёт VLR Шаркову - 9ч. пятница\\""" "close" "exit"',
                'ren "Z:\ТД\Мониторинг\Отчёты\Недельный отчёт VLR Шаркову - 9ч. пятница\VLR Subscribers.xlsx" {}.xlsx'.format(today)]
             
    for command in commands:
        subprocess.run(command, shell=True)

        
if __name__ == '__main__':
    _report_composition()