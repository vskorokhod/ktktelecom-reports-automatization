import subprocess
from datetime import date, timedelta, datetime


def _report_composition():
    today = datetime.strftime(date.today(), '%d%m%Y')
    yesterday = datetime.strftime(date.today() - timedelta(days=1), '%d%m%Y')
    day, month, year = today[0:2], today[2:4], today[4:8]
    day_yesterday, month_yesterday, year_yesterday = yesterday[0:2], yesterday[2:4], yesterday[4:8]
  
    month_str_dict = {'01': 'января',
                      '02': 'февраля',
                      '03': 'марта', 
                      '04': 'апреля', 
                      '05': 'мая', 
                      '06': 'июня', 
                      '07': 'июля', 
                      '08': 'августа', 
                      '09': 'сентября', 
                      '10': 'октября', 
                      '11': 'ноября', 
                      '12': 'декабря'}
  
    month_str = month_str_dict[month]
    month_str_yesterday = month_str_dict[month_yesterday]
    
    commands = ['"C:\Program Files (x86)\WinSCP\winscp.com" /ini=nul /command "open sftp://ossuser:Changeme_123@10.10.180.65 -hostkey=""ssh-rsa 2048 61:5b:c1:1f:ef:64:75:52:53:11:b1:d5:3f:6b:ae:4b""" "get ""/opt/oss/server/var/fileint/pm/NOC_reports/Daily_report/*.xlsx"" ""C:\Reports\Daily report\\""" "close" "exit"',
                'mkdir "G:\Share\ТД\Мониторинг\Отчёты\Суточный отчет\{0}м{1}\{2} {3} {0}\\"'.format(year, month, day, month_str),
                'robocopy "G:\Share\ТД\Мониторинг\Отчёты\Суточный отчет\{0}м{1}\{2} {3} {0}" "G:\Share\ТД\Мониторинг\Отчёты\Суточный отчет\{4}м{5}\{6} {7} {4}"'.format(year_yesterday, month_yesterday, day_yesterday, month_str_yesterday, year, month, day, month_str),
                'move "C:\Reports\Daily report\*.xlsx" "G:\Share\ТД\Мониторинг\Отчёты\Суточный отчет\{0}м{1}\{2} {3} {0}\\"'.format(year, month, day, month_str)]
             
    for command in commands:
        subprocess.run(command, shell=True)

        
if __name__ == '__main__':
    _report_composition()