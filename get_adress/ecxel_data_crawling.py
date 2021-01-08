import pandas as pd

Location = 'C:\Users\user\OneDrive - 중앙대학교\바탕 화면'
File = 'input.xlsx'

Row = 0
Column = 0

data_pd = pd.read_excel('{}/{}'.format(Location,File),
                        header=None,index_col=None,names=None)
data_np = pd.DataFrame.to_numpy(data_pd)

print(data_pd)
print(data_np[Row][Column])
