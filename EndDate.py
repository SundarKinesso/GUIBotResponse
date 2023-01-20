import datetime

import pandas as pd
#create a series object
s = pd.Series(['20220418_20220717'])
#Extract the substring
s.str.extract(r'(20220717)',expand=False)
#Convert string to datetime
datestr = s.str.extract(r'(20220717)',expand=False)[0]
dateobj = datetime.datetime.strptime(datestr,'%Y%m%d')
print(dateobj.strftime('%m/%d/%Y'))