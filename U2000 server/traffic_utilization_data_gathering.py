import sys
sys.path.append('..')
from collect_u2000_export import collect_data

_filenames = ['Outgoing Traffic Trunk Groups_60.csv', 
			 'Incoming Traffic Trunk Groups_60.csv', 
			 'Outgoing Traffic Trunk Groups_1440.csv', 
			 'Incoming Traffic Trunk Groups_1440.csv', 
			 'Outgoing Traffic Office Directions_60.csv',
			 'Incoming Traffic Office Directions_60.csv',
			 'Outgoing Traffic Office Directions_1440.csv',
			 'Incoming Traffic Office Directions_1440.csv']

if __name__ == '__main__':
	for filename in _filenames:
		collect_data(filename, 'Traffic utilization report.xlsx')
