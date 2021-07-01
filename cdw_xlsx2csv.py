import pandas as pd
import sys
import os
import logging

SrcFiles=sys.argv[1]

def get_logger(script,logname,loglevel="DEBUG"):
    # Default the logging level to INFO if no LOGGING_LEVEL param defined
    level = os.environ.get("FRINGE_LOGGING_LEVEL")
    if level == None:
        level = loglevel
    # Create/Get a custom logger
    logger = logging.getLogger(script,)
    logger.setLevel(logging.DEBUG)

    fh = logging.FileHandler(logname)
    # Set the level
    if level:
        fh.setLevel(logging.getLevelName(level))
    else:
        fh.setLevel(logging.DEBUG)

    # create console handler with a higher log level
    ch = logging.StreamHandler()
    # ch.setLevel(logging.ERROR)
    ch.setLevel(logging.DEBUG)

    # Create formatter and add it to handler
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)

    # add the handlers to the logger
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger
#default logger
logger = get_logger('CDW_XLSX2CSV',SrcFiles+'/CDW/Scripts/cdw_xlsx2csv_log.txt')
print = logger.debug

excelfiles=[]

User_SrcFiles=SrcFiles+'/CDW/User_SrcFiles'
for fname in os.listdir(User_SrcFiles):
    if os.path.isdir(User_SrcFiles + '/' + fname):
        pass
    else:
        if fname[-5:].lower()=='.xlsx' and fname[0:2] != '~$':
            excelfiles.append(fname)

print('{0} excel files found:'.format(len(excelfiles)))
print(excelfiles)

export_file = pd.read_csv(SrcFiles+"/TM1/File_Name_Export.txt", sep=',',header=None,encoding = "ISO-8859-1")
rownumbers=[1]*len(excelfiles)
for ix,i in enumerate(excelfiles):
    match=False
    searchstr = i.split('.')[0]  # removing .xlsx
    searchstr = searchstr.split('_')[0:4]  # capturing first 4 values separated by _
    searchstr = '_'.join(searchstr)
    for jx, j in enumerate(export_file[3]):
        if searchstr.lower() == j.lower():
            rownumbers[ix] = export_file.iloc[jx, 4]
            print('File:{0}--->String:{1}--->FileFormat:{2}--->DataHeader:{3}'.format(i, searchstr, j, str(rownumbers[ix])))
            match=True
            break
    if not match:
        print('File:{0}--->String:{1}--->No Match hence DataHeader:1'.format(i, searchstr))
csv_count=0
for i,f in enumerate(excelfiles):
    try:
        df=pd.read_excel(User_SrcFiles + '/' + f,skiprows=rownumbers[i]-1,header=None)
        df.replace(to_replace=[r"\\t|\\n|\\r", "\t|\n|\r"], value=[" ", " "], regex=True, inplace=True)
        df.to_csv(User_SrcFiles + '/' + f.replace('.xlsx', '.csv'),index=False,date_format='%m/%d/%Y')
        csv_count += 1
    except Exception as e:
        print('conversion of {0} into csv failed  : {1}'.format(f,str(e)))

print('{0} csv generated !'.format(csv_count))