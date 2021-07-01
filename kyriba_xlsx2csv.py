import pandas as pd
import sys
import os

path=sys.argv[1]
excelfiles=[]

for fname in os.listdir(path):
    if os.path.isdir(path + '/' + fname):
        pass
    else:
        if fname[-5:].lower()=='.xlsx' and fname[0:2] != '~$':
            excelfiles.append(fname)

print('{0} excel files found:'.format(len(excelfiles)))
print(excelfiles)

csv_count=0
for i,f in enumerate(excelfiles):
    try:
        df=pd.read_excel(path + '/' + f)
        df['Flow'] = df['Flow'].str.replace('+', ' +')
        df['Flow']=df['Flow'].str.replace('-',' -')
        df['Update time']=pd.Series([val.time() for val in df['Update time']])
        df.to_csv(path + '/' + f.replace('.xlsx', '.csv'),index=False,date_format='%m/%d/%Y')
        os.rename(path + '/' + f, path + '/archive/' + f )
        csv_count += 1
    except Exception as e:
        print('conversion of {0} into csv failed  : {1}'.format(f,str(e)))


print('{0} csv generated from excel !'.format(csv_count))