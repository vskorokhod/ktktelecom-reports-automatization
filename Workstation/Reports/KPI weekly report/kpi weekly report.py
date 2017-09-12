import subprocess
from datetime import date, timedelta, datetime


def _report_composition():
    yesterday = datetime.strftime(date.today() - timedelta(days=1), '%Y.%m.%d')
    week_ago = datetime.strftime(date.today() - timedelta(days=8), '%Y.%m.%d')
    
    commands = ['"C:\Program Files (x86)\WinSCP\winscp.com" /ini=nul /command "open sftp://ossuser:Changeme_123@10.10.180.65 -hostkey=""ssh-rsa 2048 61:5b:c1:1f:ef:64:75:52:53:11:b1:d5:3f:6b:ae:4b""" "get ""/opt/oss/server/var/fileint/pm/NOC_reports/KPI_weekly_report/*.xlsx"" ""C:\Reports\KPI weekly report\\""" "close" "exit"',
                'mkdir "Z:\ТД\Мониторинг\Отчёты\Недельный отчет в разрезе месяца - 9ч. понедельник\{}\\"'.format(yesterday),
                'robocopy "Z:\ТД\Мониторинг\Отчёты\Недельный отчет в разрезе месяца - 9ч. понедельник\{}" "Z:\ТД\Мониторинг\Отчёты\Недельный отчет в разрезе месяца - 9ч. понедельник\{}"'.format(yesterday, week_ago),
                'move "C:\Reports\KPI weekly report\*.xlsx" "Z:\ТД\Мониторинг\Отчёты\Недельный отчет в разрезе месяца - 9ч. понедельник\{}\\"'.format(yesterday)]
             
    for command in commands:
        subprocess.run(command, shell=True)

        
if __name__ == '__main__':
    _report_composition()