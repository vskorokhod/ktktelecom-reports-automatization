import sys
sys.path.append('..')
from collect_u2000_export import collect_data

if __name__ == '__main__':
	collect_data('1.0 VLR_Subscribers.csv', 'VLR Subscribers.xlsx')