import pandas as pd
import os
import time

def CallType(dtf):
    
    dtf = dtf.dropna()
    dtf['DateTime'] = dtf['DateTime'].apply(lambda x:x.date())
    dtf['ASA'] = dtf['ASA'].apply(lambda t:(t.second+t.minute*60+t.hour*3600)) 
    dtf['TalkTime'] = dtf['TalkTime'].apply(lambda t:(t.second+t.minute*60+t.hour*3600)) 
    dtf['Answered <60'] = dtf['SL']*dtf['Answered']

    df1 = dtf.copy()
    cols = ['Call Type', 'DateTime', 'Offered', 'Answered', 'Answered <60', 'Aban', 'TalkTime', 'ASA', 'Calls Error', 'Flow Out']
    df1 = df1[cols]

    df2 = df1.groupby(['Call Type', 'DateTime']).sum().reset_index()
    df2.insert(5, '% Answered <60', df2['Answered <60']/df2['Answered'])

    df2['Average TalkTime'] = df2.apply(lambda _: '', axis=1)
    df2['Average ASA'] = df2.apply(lambda _: '', axis=1)
    df2['True Calls Offered'] = df2.apply(lambda _: '', axis=1)
    df2['% Call Abandoned'] = df2.apply(lambda _: '', axis=1)

    df2['True Calls Offered'] = df2['Offered'] - (df2['Calls Error'] + df2['Flow Out'])
    df2['Average TalkTime'] = df2['TalkTime']/df2['Answered']
    df2['Average ASA'] = df2['ASA']/df2['Answered']
    df2['% Call Abandoned'] = df2['Aban']/df2['True Calls Offered']

    lst = ['TalkTime','ASA']
    for var in lst:
        i=0
        for x in df2['Average '+var]:
            a = time.strftime("%H:%M:%S", time.gmtime(x))
            df2['Average '+var][i] = a
            i+=1

    df2['Offered'] = df2['True Calls Offered']
    df2 = df2.drop(['True Calls Offered','TalkTime','ASA', 'Calls Error', 'Flow Out'], axis=1)
    df2 = df2.rename(columns = {"Call Type": "Service Desk", 
                            "DateTime":"Date", 
                            "Offered": "Call Offered",
                            "Answered": "Call Answered",
                            "% Answered <60":"% Answered <60", 
                            "Aban": "Call Abandoned",
                            "Average TalkTime": "Average Talk Time",
                            "Average ASA": "Average Speed to Answer"}) 
    return df2.sort_values(by=['Date'])

def FuncSAP(dtf):
    
    dtf = dtf = dtf[3]
    dtf = dtf.drop(columns=0)
    dtf = dtf[-1:]
    dtf = dtf.rename(columns = {"0": "Call Offered",
                            "1": "Call Answered",
                            "2":"% Answered <60", 
                            "3": "Call Abandoned",
                            "4": "Average Talk Time",
                            "5": "Average Speed to Answer"}) 
    
    return dtf

def bool_CallType_DATA(df): #If Call type.xls has not data from daily stats returns true
    dfx = df[0]
    if list(dfx.columns)==list(dfx.iloc[1]):
        return True
    
