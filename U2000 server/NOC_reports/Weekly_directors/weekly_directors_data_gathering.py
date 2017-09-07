import sys
sys.path.append('..')
from collect_u2000_export import collect_data

_filenames = ['1.1 CSSR_15.csv', 
			 '1.2 Call Drop Rate 2G_15.csv', 
			 '1.3 Call Drop rate 3G_15.csv', 
			 '1.4 Voice Trafic 2G 3G_60.csv', 
			 '1.5 Data Traffic 2G GBytes_60.csv',
			 '1.6 Data Traffic 3G GBytes_60.csv',
			 '1.7 Data Traffic 4G GBytes_60.csv',
			 '1.8 2G Roamers Data Traffic_60.csv',
			 '1.9 3G Roamers Data Traffic_60.csv',
			 '2.0 4G Roamers Data Traffic_60.csv']

if __name__ == '__main__':
	for filename in _filenames:
		collect_data(filename, 'Weekly directors report.xlsx')