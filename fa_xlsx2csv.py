#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
usecols=["Fiscal Period", "Accounted (Ledger Functional) Amount","Business Unit", "Account","Ledger","Entered (Transactional) Amount","Trans Type","Currency","Fx Rate","USD (calc)"]
df=pd.read_excel('M.01.a_AM_FY20_GL tab export.xlsx',usecols=usecols)
df=df.dropna(how='all')


# In[3]:


def convertFloat(val):
    try:
        return float(str(val).replace(' ','').replace(',',''))
    except:
        print(val)
        return None
def convertInt(val):
    try:
        return int(float(str(val).replace(' ','').replace(',','')))
    except:
        print(val)
        return None
df['Entered (Transactional) Amount']=df['Entered (Transactional) Amount'].apply(lambda x: convertFloat(x))
df['USD (calc)']=df['USD (calc)'].apply(lambda x: convertFloat(x))
df['Fiscal Period']=df['Fiscal Period'].fillna(-1).apply(lambda x: convertInt(x))
df['Business Unit']=df['Business Unit'].fillna(-1).apply(lambda x: convertInt(x))
df['Account']=df['Account'].fillna(-1).apply(lambda x: convertInt(x))
# df.to_csv('M01_GL_REPORT.csv',index=False)


# In[7]:


df.dtypes


# In[4]:


df.dtypes


# In[5]:


# df.to_csv(r'\\bank\infa_corpfin_prd\sus9_r12\SrcFiles\EDW\M01_GL_REPORT.csv',index=False)
df.to_csv('M01_GL_REPORT_NEW.csv',index=False)

