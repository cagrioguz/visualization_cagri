# -*- coding: utf-8 -*-
"""
Created on Thu Jul  1 20:41:35 2021

@author: osman
"""
def subject_decoding_results_to_excel(df_svm, filepath):
    import pandas as pd
    import openpyxl
    
    df_svm_sig = df_svm.loc[lambda df: df_svm.p_value < .05, :]
    df_svm_sig = df_svm_sig.sort_values(by=['lead'])

    df_svm_sig_piv = df_svm_sig.pivot_table(values='accuracy', index=['subject', 'lead'],
                                            columns='classification_type',aggfunc='first')

    df_svm_sig_piv = df_svm_sig_piv.fillna("ns")
    
    """adding column for ac specificity""" 
    df_svm_sig_piv = df_svm_sig_piv.append(
        {"AC_Specificity": []}, ignore_index=True)
    
    lenght_piv = len(df_svm_sig_piv)
    
    """Naming AC Specificities""" 
    for i in range(lenght_piv-1):
        if (df_svm_sig_piv.iloc[i, 1] != 'ns' and df_svm_sig_piv.iloc[i, 2] != 'ns' and df_svm_sig_piv.iloc[i, 0] != 'ns'):
            df_svm_sig_piv.AC_Specificity.iloc[i] = 'ALL'
        elif (df_svm_sig_piv.iloc[i, 0] != 'ns' and df_svm_sig_piv.iloc[i, 2] != 'ns'):
            df_svm_sig_piv.AC_Specificity.iloc[i] = 'MAN'
        elif (df_svm_sig_piv.iloc[i, 1] != 'ns' and df_svm_sig_piv.iloc[i, 2] != 'ns'):
            df_svm_sig_piv.AC_Specificity.iloc[i] = 'SD'
        elif (df_svm_sig_piv.iloc[i, 1] != 'ns' and df_svm_sig_piv.iloc[i, 0] != 'ns'):
            df_svm_sig_piv.AC_Specificity.iloc[i] = 'IP'
        else:
            df_svm_sig_piv.AC_Specificity.iloc[i] = 'NONE'

    

    book = openpyxl.load_workbook(filepath)
    
    """Dataframes for each subject (3er sheet) to excel"""
    with pd.ExcelWriter(filepath, engine='openpyxl', mode='a') as writer:
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        df_svm.to_excel(
            writer, sheet_name=df_svm.subject.iloc[0] + ' all results', index=False)
        df_svm_sig.to_excel(
            writer, sheet_name=df_svm.subject.iloc[0] + ' sig. results', index=False)
        df_svm_sig_piv.to_excel(
            writer, sheet_name=df_svm.subject.iloc[0] + ' sig - ac grouped') 
    

    allresult_sheets = ['AllSubjects','All Significant Results','All Sig Ac']
    
    """Bendeki Türkçe Excel olduğu için excel'i oluşturduğumuzda kendiliğinden çıkan Sayfa1 sheet'ini yok ettim""" 
    """Yok edilmezse sondaki for loop'ta range'in gözden geçirilmesi gerekiyor"""
    keylist=list(writer.sheets.keys())
    if 'Sayfa1' in keylist:
        std = book['Sayfa1']
        book.remove(std)
        book.save(filepath)
    
    """All result sheetlerinin farklı subjectler için functionı çağırdığımızda yenilenmesini ve sonda olmasını istiyoruz
    O yüzden yok edip tekrar oluşturdum""" 
    for i in range(len(allresult_sheets)):
        if allresult_sheets[i] in keylist:
            std = book[allresult_sheets[i]]
            book.remove(std)
            book.save(filepath)
        book.create_sheet(allresult_sheets[i])
        book.save(filepath)
    
    keylist=list(writer.sheets.keys())
    
    """Bütün subjectler için all subjectin spesifik sheet'ine karşılık gelen subject sheetlerindeki
    veriler, all subject için belirtilen sheete alt alta eklenmiş oluyor""" 
    with pd.ExcelWriter(filepath, engine='openpyxl', mode='a') as writer:
        writer.book = book
        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
        for i in range(0,len(keylist)-3,3):
            df = pd.read_excel(filepath,sheet_name = keylist[i])
            target_sheet = pd.read_excel(filepath,sheet_name = 'AllSubjects')
            df.to_excel(writer,sheet_name = 'AllSubjects',index=False,header=True,startrow=len(target_sheet)+1)
            book.save(filepath)
        for i in range(1,len(keylist)-3,3):
            df = pd.read_excel(filepath,sheet_name = keylist[i])
            target_sheet = pd.read_excel(filepath,sheet_name = 'All Significant Results')
            df.to_excel(writer,sheet_name = 'All Significant Results',index=False,header=True,startrow=len(target_sheet)+1)
            book.save(filepath)
        for i in range(2,len(keylist)-3,3):
            df = pd.read_excel(filepath,sheet_name = keylist[i])
            target_sheet = pd.read_excel(filepath,sheet_name = 'All Sig Ac')
            df.to_excel(writer,sheet_name = 'All Sig Ac',index=False,header=True,startrow=len(target_sheet)+1)
            book.save(filepath)
            
import pandas as pd
import datetime 
farci_df = pd.read_pickle('FarciM_decoding_results.pkl')
guigli_df = pd.read_pickle('GuigliF_decoding_results.pkl')
folder_path = r'C:\Users\osman\Desktop\Intracerebral'
date = datetime.datetime.today().strftime('%d-%m')
filepath = folder_path + '\\' + date + '.xlsx'
subject_decoding_results_to_excel(farci_df, filepath)
subject_decoding_results_to_excel(guigli_df, filepath) 
