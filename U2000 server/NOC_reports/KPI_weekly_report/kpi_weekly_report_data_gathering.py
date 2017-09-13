import sys
sys.path.append('..')
from collect_u2000_export import collect_data

_filenames = ['1.0 VLR_Subscribers.csv',
			 '1.1 CSSR.csv',
			 '1.2 PDP SR 2G.csv',
			 '1.3 PDP SR 3G.csv',
			 '1.4 Attach SR 2G_KTK_MTS.csv',
			 '1.5 Attach SR 3G_KTK_MTS.csv',
			 '1.6 Attach SR 4G_KTK_MTS.csv',
			 '1.7 Voice Trafic 2G 3G.csv',
			 '1.8 Call Drop Rate 2G.csv',
			 '1.9 Call Drop rate 3G.csv',
			 '2.0 SMS SR.csv',
			 '2.1 Data Traffic 2G GBytes_UL_DL.csv',
			 '2.2 2G Roamers Data Traffic_UL_DL.csv',
			 '2.3 Data Traffic 3G GBytes_UL_DL.csv',
			 '2.4 3G Roamers Data Traffic_UL_DL.csv',
			 '2.5 Data Traffic 4G GBytes_UL_DL.csv',
			 '2.6 4G Roamers Data Traffic_UL_DL.csv',
			 '2.7 MTCSUCCESS.csv',
			 'Congection Rate_60.csv',
			 '2.9 Trunk Utilization_Trunk group.csv']

if __name__ == '__main__':
	for filename in _filenames:
		collect_data(filename, 'KPI weekly report.xlsx', daily_report=False, kpi_weekly_report=True)