import pandas as pd
import pickle
import xlrd
import tia.bbg.datamgr as dm
from tia.bbg import LocalTerminal
import numpy as np
desired_width = 320
pd.set_option('display.width', desired_width)
mgr = dm.BbgDataManager()
addstring = ' EQUITY'
# CANADA
xl_workbook = xlrd.open_workbook('Underlying\Data\originalTickers.xlsx')
xl_sheet = xl_workbook.sheet_by_name('CAD2')
cad_tickers = xl_sheet.row_values(0)
# Convert to strings
cad_tickers = [str(x) for x in cad_tickers]
# Add " EQUITY" to list of strings via list comprehension
cad_tickers = [x + addstring for x in cad_tickers]
sids_cad = mgr[cad_tickers]
cad_df = sids_cad.get_historical('TOT_RETURN_INDEX_GROSS_DVDS', '11/1/2014', '08/1/2016')
pickle_out = open('caddb2', 'wb')
pickle.dump(cad_df, pickle_out)
pickle_out.close()
# JAPAN
xl_workbook = xlrd.open_workbook('Underlying\Data\originalTickers.xlsx')
xl_sheet = xl_workbook.sheet_by_name('JPY2')
jpy_tickers = xl_sheet.row_values(0)
jpy_tickers = [str(x) for x in jpy_tickers]
jpy_tickers = [x + addstring for x in jpy_tickers]
sids_jpy = mgr[jpy_tickers]
jpy_df = sids_jpy.get_historical('TOT_RETURN_INDEX_GROSS_DVDS', '11/1/2014', '08/1/2016')
pickle_out = open('jpydb2', 'wb')
pickle.dump(jpy_df, pickle_out)
pickle_out.close()
# AUSTRALIA
xl_workbook = xlrd.open_workbook('Underlying\Data\originalTickers.xlsx')
xl_sheet = xl_workbook.sheet_by_name('AUD2')
aud_tickers = xl_sheet.row_values(0)
aud_tickers = [str(x) for x in aud_tickers]
aud_tickers = [x + addstring for x in aud_tickers]
sids_aud = mgr[aud_tickers]
aud_df = sids_aud.get_historical('TOT_RETURN_INDEX_GROSS_DVDS', '11/1/2014', '08/1/2016')
pickle_out = open('auddb2', 'wb')
pickle.dump(aud_df, pickle_out)
pickle_out.close()
# USA
xl_workbook = xlrd.open_workbook('Underlying\Data\originalTickers.xlsx')
xl_sheet = xl_workbook.sheet_by_name('USA2')
usd_tickers = xl_sheet.row_values(0)
usd_tickers = [str(x) for x in usd_tickers]
usd_tickers = [x + addstring for x in usd_tickers]
sids_usd = mgr[usd_tickers]
usd_df = sids_usd.get_historical('TOT_RETURN_INDEX_GROSS_DVDS', '11/1/2014', '08/1/2016')
pickle_out = open('usddb2', 'wb')
pickle.dump(usd_df, pickle_out)
pickle_out.close()
# EUR
xl_workbook = xlrd.open_workbook('Underlying\Data\originalTickers.xlsx')
xl_sheet = xl_workbook.sheet_by_name('EUR2')
eur_tickers = xl_sheet.row_values(0)
eur_tickers = [str(x) for x in eur_tickers]
eur_tickers = [x + addstring for x in eur_tickers]
sids_eur = mgr[eur_tickers]
eur_df = sids_eur.get_historical('TOT_RETURN_INDEX_GROSS_DVDS', '11/1/2014', '08/1/2016',currency="EUR")
# eur_df = sids_eur.get_historical('TOT_RETURN_INDEX_GROSS_DVDS', '11/1/2014', '08/1/2016',
#                                  non_trading_day_fill_option="ALL_CALENDAR_DAYS",
#                                  non_trading_day_fill_method="PREVIOUS_VALUE",
#                                  currency="EUR")
pickle_out = open('Underlying\eurdb2', 'wb')
pickle.dump(eur_df, pickle_out)
pickle_out.close()

# GBP
xl_workbook = xlrd.open_workbook('Underlying\Data\originalTickers.xlsx')
xl_sheet = xl_workbook.sheet_by_name('GBP2')
gbp_tickers = xl_sheet.row_values(0)
gbp_tickers = [str(x) for x in gbp_tickers]
gbp_tickers = [x + addstring for x in gbp_tickers]
sids_gbp = mgr[gbp_tickers]
gbp_df = sids_gbp.get_historical('TOT_RETURN_INDEX_GROSS_DVDS', '11/1/2014', '08/1/2016')
pickle_out = open('gbpdb2', 'wb')
pickle.dump(gbp_df, pickle_out)
pickle_out.close()

# Import Data from Pickle and Calculate Returns
# CANADA
cad_data = pd.read_pickle('Underlying/caddb2').sort_index(axis=1)
# Forward Fill to account for public holidays
cad_data = cad_data.asfreq("D",method='ffill')
cad_data = cad_data.fillna(method='ffill')
cad_return = cad_data.pct_change()
# JAPAN
jpy_data = pd.read_pickle('Underlying/jpydb2').sort_index(axis=1)
jpy_data = jpy_data.asfreq("D",method='ffill')
jpy_data = jpy_data.fillna(method='ffill')
jpy_return = jpy_data.pct_change()
# AUSTRALIA
aud_data = pd.read_pickle('Underlying/auddb2').sort_index(axis=1)
aud_data = aud_data.asfreq("D",method='ffill')
aud_data = aud_data.fillna(method='ffill')
aud_return = aud_data.pct_change()
# USA
usd_data = pd.read_pickle('Underlying/usddb2').sort_index(axis=1)
usd_data = usd_data.asfreq("D",method='ffill')
usd_data = usd_data.fillna(method='ffill')
usd_return = usd_data.pct_change()
# EUROPE
eur_data = pd.read_pickle('Underlying/eurdb2').sort_index(axis=1)
eur_data = eur_data.asfreq("D",method='ffill')
eur_data = eur_data.fillna(method='ffill')
eur_return = eur_data.pct_change()
# UK
gbp_data = pd.read_pickle('Underlying/gbpdb2').sort_index(axis=1)
gbp_data = gbp_data.asfreq("D",method='ffill')
gbp_data = gbp_data.fillna(method='ffill')
gbp_return = gbp_data.pct_change()

######################################### LONG ONLY ############################################
######################################### RAW BETA #############################################
# CAD
cad_factor_wt = pd.read_excel('Underlying\Data\originalTickers.xlsx',sheetname='CADWt', index_col=0)
cad_factor_pivot = cad_factor_wt.pivot_table(cad_factor_wt, index=cad_factor_wt.index, columns='UnifiedTicker')
new_index = cad_data.index
cad_factor_pivot = cad_factor_pivot.reindex(new_index, method='ffill')
cad_factor_rawbeta = cad_factor_pivot['Raw_Beta'].sort_index(axis=1).shift(1)
cad_factor_rawbeta[cad_factor_rawbeta < 0] = 0
cad_factor_rawbeta_wt = cad_factor_rawbeta.div(cad_factor_rawbeta.sum(axis=1),axis=0)
cad_factor_rawbeta_wtret = (np.multiply(cad_return, cad_factor_rawbeta)).sum(axis=1)
cad_rawbeta = (1.0+cad_factor_rawbeta_wtret).cumprod(axis=0)
# JPY
jpy_factor_wt = pd.read_excel('Underlying\Data\originalTickers.xlsx',sheetname='JPYWt', index_col=0)
jpy_factor_pivot = jpy_factor_wt.pivot_table(jpy_factor_wt, index=jpy_factor_wt.index, columns='UnifiedTicker')
new_index = jpy_data.index
jpy_factor_pivot = jpy_factor_pivot.reindex(new_index, method='ffill')
jpy_factor_rawbeta = jpy_factor_pivot['Raw_Beta'].sort_index(axis=1).shift(1)
jpy_factor_rawbeta[jpy_factor_rawbeta < 0] = 0
jpy_factor_rawbeta_wt = jpy_factor_rawbeta.div(jpy_factor_rawbeta.sum(axis=1),axis=0)
jpy_factor_rawbeta_wtret = (np.multiply(jpy_return, jpy_factor_rawbeta)).sum(axis=1)
jpy_rawbeta = (1.0+jpy_factor_rawbeta_wtret).cumprod(axis=0)
# USD
usd_factor_wt = pd.read_excel('Underlying\Data\originalTickers.xlsx',sheetname='USAWt', index_col=0)
usd_factor_pivot = usd_factor_wt.pivot_table(usd_factor_wt,index=usd_factor_wt.index, columns='UnifiedTicker')
new_index = usd_data.index
usd_factor_pivot = usd_factor_pivot.reindex(new_index, method='ffill')
usd_factor_rawbeta = usd_factor_pivot['Raw_Beta'].sort_index(axis=1).shift(1)
usd_factor_rawbeta[usd_factor_rawbeta < 0] = 0
usd_factor_rawbeta_wt = usd_factor_rawbeta.div(usd_factor_rawbeta.sum(axis=1),axis=0)
usd_factor_rawbeta_wtret = (np.multiply(usd_return, usd_factor_rawbeta)).sum(axis=1)
usd_rawbeta = (1.0+usd_factor_rawbeta_wtret).cumprod(axis=0)
# GBP
gbp_factor_wt = pd.read_excel('Underlying\Data\originalTickers.xlsx',sheetname='GBPWt', index_col=0)
gbp_factor_pivot = gbp_factor_wt.pivot_table(gbp_factor_wt, index=gbp_factor_wt.index, columns='UnifiedTicker')
new_index = gbp_data.index
gbp_factor_pivot = gbp_factor_pivot.reindex(new_index, method='ffill')
gbp_factor_rawbeta = gbp_factor_pivot['Raw_Beta'].sort_index(axis=1).shift(1)
gbp_factor_rawbeta[gbp_factor_rawbeta < 0] = 0
gbp_factor_rawbeta_wt = gbp_factor_rawbeta.div(gbp_factor_rawbeta.sum(axis=1),axis=0)
gbp_factor_rawbeta_wtret = (np.multiply(gbp_return, gbp_factor_rawbeta)).sum(axis=1)
gbp_rawbeta = (1.0+gbp_factor_rawbeta_wtret).cumprod(axis=0)
# EUR
eur_factor_wt = pd.read_excel('Underlying\Data\originalTickers.xlsx',sheetname='EURWt', index_col=0)
eur_factor_pivot = eur_factor_wt.pivot_table(eur_factor_wt, index=eur_factor_wt.index, columns='UnifiedTicker')
new_index = eur_data.index
eur_factor_pivot = eur_factor_pivot.reindex(new_index, method='ffill')
eur_factor_rawbeta = eur_factor_pivot['Raw_Beta'].sort_index(axis=1).shift(1)
eur_factor_rawbeta[eur_factor_rawbeta < 0] = 0
eur_factor_rawbeta_wt = eur_factor_rawbeta.div(eur_factor_rawbeta.sum(axis=1),axis=0)
eur_factor_rawbeta_wtret = (np.multiply(eur_return, eur_factor_rawbeta)).sum(axis=1)
eur_rawbeta = (1.0+eur_factor_rawbeta_wtret).cumprod(axis=0)
# AUD
aud_factor_wt = pd.read_excel('Underlying\Data\originalTickers.xlsx',sheetname='AUDWt', index_col=0)
aud_factor_pivot = aud_factor_wt.pivot_table(aud_factor_wt, index=aud_factor_wt.index, columns='UnifiedTicker')
new_index = aud_data.index
aud_factor_pivot = aud_factor_pivot.reindex(new_index, method='ffill')
aud_factor_rawbeta = aud_factor_pivot['Raw_Beta'].sort_index(axis=1).shift(1)
aud_factor_rawbeta[aud_factor_rawbeta < 0] = 0
aud_factor_rawbeta_wt = aud_factor_rawbeta.div(aud_factor_rawbeta.sum(axis=1),axis=0)
aud_factor_rawbeta_wtret = (np.multiply(aud_return, aud_factor_rawbeta)).sum(axis=1)
aud_rawbeta = (1.0+aud_factor_rawbeta_wtret).cumprod(axis=0)

# Compile Daily and Month End Series into 2 Dataframes
rawbeta_df = pd.concat([jpy_rawbeta, cad_rawbeta, usd_rawbeta, gbp_rawbeta, eur_rawbeta, aud_rawbeta], axis=1)
rawbeta_df.columns = ['JPY','CAD','USD','GBP','EUR','AUD']
rawbeta_df = rawbeta_df.asfreq("D",method='ffill')
rawbeta_df = rawbeta_df.fillna(method='ffill')
rawbeta_df['AC EW Mean Beta'] = rawbeta_df.mean(axis=1)
rawbeta_monthly_df = rawbeta_df.resample('M').last()
######################################### RAW VALUE #############################################
# CAD
# Shift everything forward by t+1
cad_factor_rawvalue = cad_factor_pivot['Raw_Value'].sort_index(axis=1).shift(1)
cad_factor_rawvalue[cad_factor_rawvalue < 0] = 0
cad_factor_rawvalue_wt = cad_factor_rawvalue.div(cad_factor_rawvalue.sum(axis=1),axis=0)
cad_factor_rawvalue_wtret = (np.multiply(cad_return, cad_factor_rawvalue_wt)).sum(axis=1)
cad_rawvalue = (1.0+cad_factor_rawvalue_wtret).cumprod(axis=0)
# JPY
jpy_factor_rawvalue = jpy_factor_pivot['Raw_Value'].sort_index(axis=1).shift(1)
jpy_factor_rawvalue[jpy_factor_rawvalue < 0] = 0
jpy_factor_rawvalue_wt = jpy_factor_rawvalue.div(jpy_factor_rawvalue.sum(axis=1),axis=0)
jpy_factor_rawvalue_wtret = (np.multiply(jpy_return, jpy_factor_rawvalue_wt)).sum(axis=1)
jpy_rawvalue = (1.0+jpy_factor_rawvalue_wtret).cumprod(axis=0)
# USD
usd_factor_rawvalue = usd_factor_pivot['Raw_Value'].sort_index(axis=1).shift(1)
usd_factor_rawvalue[usd_factor_rawvalue < 0] = 0
usd_factor_rawvalue_wt = usd_factor_rawvalue.div(usd_factor_rawvalue.sum(axis=1),axis=0)
usd_factor_rawvalue_wtret = (np.multiply(usd_return, usd_factor_rawvalue_wt)).sum(axis=1)
usd_rawvalue = (1.0+usd_factor_rawvalue_wtret).cumprod(axis=0)
# EUR
eur_factor_rawvalue = eur_factor_pivot['Raw_Value'].sort_index(axis=1).shift(1)
eur_factor_rawvalue[eur_factor_rawvalue < 0] = 0
eur_factor_rawvalue_wt = eur_factor_rawvalue.div(eur_factor_rawvalue.sum(axis=1),axis=0)
eur_factor_rawvalue_wtret = (np.multiply(eur_return, eur_factor_rawvalue_wt)).sum(axis=1)
eur_rawvalue = (1.0+eur_factor_rawvalue_wtret).cumprod(axis=0)
# GBP
gbp_factor_rawvalue = gbp_factor_pivot['Raw_Value'].sort_index(axis=1).shift(1)
gbp_factor_rawvalue[gbp_factor_rawvalue < 0] = 0
gbp_factor_rawvalue_wt = gbp_factor_rawvalue.div(gbp_factor_rawvalue.sum(axis=1),axis=0)
gbp_factor_rawvalue_wtret = (np.multiply(gbp_return, gbp_factor_rawvalue_wt)).sum(axis=1)
gbp_rawvalue = (1.0+gbp_factor_rawvalue_wtret).cumprod(axis=0)
# AUD
aud_factor_rawvalue = aud_factor_pivot['Raw_Value'].sort_index(axis=1).shift(1)
aud_factor_rawvalue[aud_factor_rawvalue < 0] = 0
aud_factor_rawvalue_wt = aud_factor_rawvalue.div(aud_factor_rawvalue.sum(axis=1),axis=0)
aud_factor_rawvalue_wtret = (np.multiply(aud_return, aud_factor_rawvalue_wt)).sum(axis=1)
aud_rawvalue = (1.0+aud_factor_rawvalue_wtret).cumprod(axis=0)
# Compile Daily and Month End Series into 2 Dataframes
rawvalue_df = pd.concat([jpy_rawvalue, cad_rawvalue, usd_rawvalue, gbp_rawvalue, eur_rawvalue, aud_rawvalue], axis=1)
rawvalue_df.columns = ['JPY','CAD','USD','GBP','EUR','AUD']
rawvalue_df = rawvalue_df.asfreq("D",method='ffill')
rawvalue_df = rawvalue_df.fillna(method='ffill')
rawvalue_df['AC EW Mean Value'] = rawvalue_df.mean(axis=1)
rawvalue_monthly_df = rawvalue_df.resample('M').last()

######################################### MOMENTUM #############################################
# CAD
cad_factor_rawsmom = cad_factor_pivot['Raw_SMOM'].sort_index(axis=1).shift(1)
cad_factor_rawsmom = cad_factor_rawsmom[cad_factor_rawsmom.index < '2015-08-12']
# If scaled momentum was positive, then invest in Raw MOM
cad_mom_table = cad_factor_pivot['Scaled_SMOM'].sort_index(axis=1).shift(1)
cad_mom_table = cad_mom_table[cad_mom_table.index < '2015-08-12']
cad_mom_table[cad_mom_table < 0] = 0
cad_mom_table[cad_mom_table > 0] = 1
cad_factor_rawsmom = np.multiply(cad_mom_table, cad_factor_rawsmom)
cad_factor_rawrmom = cad_factor_pivot['Raw_RMOM'].sort_index(axis=1)
cad_factor_rawrmom = cad_factor_rawrmom[cad_factor_rawrmom.index >= '2015-08-12']
cad_factor_mom = pd.concat([cad_factor_rawsmom, cad_factor_rawrmom],axis=0).sort_index(axis=1)
cad_factor_mom[cad_factor_mom < 0] = 0
cad_factor_mom_wt = cad_factor_mom
cad_factor_mom_wt.head()
cad_factor_mom_wt[cad_factor_mom.sum(axis=1)>0] = cad_factor_mom.div(cad_factor_mom.sum(axis=1),axis=0)
cad_mom_score = np.multiply(cad_factor_mom_wt, cad_return)
cad_mom_score.head()
cad_mom_sum = cad_mom_score.sum(axis=1, skipna='True')
cad_mom_sum.head()
cad_mom = (1.0+cad_mom_sum).cumprod(axis=0)
cad_mom.head()
# JPY
jpy_factor_rawsmom = jpy_factor_pivot['Raw_SMOM'].sort_index(axis=1).shift(1)
jpy_factor_rawsmom = jpy_factor_rawsmom[jpy_factor_rawsmom.index < '2015-08-12']
jpy_mom_table = jpy_factor_pivot['Scaled_SMOM'].sort_index(axis=1).shift(1)
jpy_mom_table = jpy_mom_table[jpy_mom_table.index < '2015-08-12']
jpy_mom_table[jpy_mom_table < 0] = 0
jpy_mom_table[jpy_mom_table > 0] = 1
jpy_factor_rawsmom = np.multiply(jpy_mom_table, jpy_factor_rawsmom)
jpy_factor_rawrmom = jpy_factor_pivot['Raw_RMOM'].sort_index(axis=1)
jpy_factor_rawrmom = jpy_factor_rawrmom[jpy_factor_rawrmom.index >= '2015-08-12']
jpy_factor_mom = pd.concat([jpy_factor_rawsmom, jpy_factor_rawrmom],axis=0).sort_index(axis=1)
jpy_factor_mom[jpy_factor_mom < 0] = 0
jpy_factor_mom_wt = jpy_factor_mom
jpy_factor_mom_wt[jpy_factor_mom.sum(axis=1)>0] = jpy_factor_mom.div(jpy_factor_mom.sum(axis=1),axis=0)
jpy_mom_score = np.multiply(jpy_factor_mom_wt, jpy_return)
jpy_mom_sum = jpy_mom_score.sum(axis=1, skipna=True)
jpy_mom = (1.0+jpy_mom_sum).cumprod(axis=0)
# USD
usd_factor_rawsmom = usd_factor_pivot['Raw_SMOM'].sort_index(axis=1).shift(1)
usd_factor_rawsmom = usd_factor_rawsmom[usd_factor_rawsmom.index < '2015-08-12']
usd_mom_table = usd_factor_pivot['Scaled_SMOM'].sort_index(axis=1).shift(1)
usd_mom_table = usd_mom_table[usd_mom_table.index < '2015-08-12']
usd_mom_table[usd_mom_table < 0] = 0
usd_mom_table[usd_mom_table > 0] = 1
usd_factor_rawsmom = np.multiply(usd_mom_table, usd_factor_rawsmom)
usd_factor_rawrmom = usd_factor_pivot['Raw_RMOM'].sort_index(axis=1)
usd_factor_rawrmom = usd_factor_rawrmom[usd_factor_rawrmom.index >= '2015-08-12']
usd_factor_mom = pd.concat([usd_factor_rawsmom, usd_factor_rawrmom],axis=0).sort_index(axis=1)
usd_factor_mom[usd_factor_mom < 0] = 0
usd_factor_mom_wt = usd_factor_mom
usd_factor_mom_wt[usd_factor_mom.sum(axis=1)>0] = usd_factor_mom.div(usd_factor_mom.sum(axis=1),axis=0)
usd_mom_score = np.multiply(usd_factor_mom_wt, usd_return)
usd_mom_sum = usd_mom_score.sum(axis=1)
usd_mom = (1.0+usd_mom_sum).cumprod(axis=0)
# EUR
eur_factor_rawsmom = eur_factor_pivot['Raw_SMOM'].sort_index(axis=1).shift(1)
eur_factor_rawsmom = eur_factor_rawsmom[eur_factor_rawsmom.index < '2015-08-12']
eur_mom_table = eur_factor_pivot['Scaled_SMOM'].sort_index(axis=1).shift(1)
eur_mom_table = eur_mom_table[eur_mom_table.index < '2015-08-12']
eur_mom_table[eur_mom_table < 0] = 0
eur_mom_table[eur_mom_table > 0] = 1
eur_factor_rawsmom = np.multiply(eur_mom_table, eur_factor_rawsmom)
eur_factor_rawrmom = eur_factor_pivot['Raw_RMOM'].sort_index(axis=1)
eur_factor_rawrmom = eur_factor_rawrmom[eur_factor_rawrmom.index >= '2015-08-12']
eur_factor_mom = pd.concat([eur_factor_rawsmom, eur_factor_rawrmom],axis=0).sort_index(axis=1)
eur_factor_mom[eur_factor_mom < 0] = 0
eur_factor_mom_wt = eur_factor_mom
eur_factor_mom_wt[eur_factor_mom.sum(axis=1)>0] = eur_factor_mom.div(eur_factor_mom.sum(axis=1),axis=0)
eur_mom_score = np.multiply(eur_factor_mom_wt, eur_return)
eur_mom_sum = eur_mom_score.sum(axis=1)
eur_mom = (1.0+eur_mom_sum).cumprod(axis=0)
# GBP
gbp_factor_rawsmom = gbp_factor_pivot['Raw_SMOM'].sort_index(axis=1).shift(1)
gbp_factor_rawsmom = gbp_factor_rawsmom[gbp_factor_rawsmom.index < '2015-08-12']
gbp_mom_table = gbp_factor_pivot['Scaled_SMOM'].sort_index(axis=1).shift(1)
gbp_mom_table = gbp_mom_table[gbp_mom_table.index < '2015-08-12']
gbp_mom_table[gbp_mom_table < 0] = 0
gbp_mom_table[gbp_mom_table > 0] = 1
gbp_factor_rawsmom = np.multiply(gbp_mom_table, gbp_factor_rawsmom)
gbp_factor_rawrmom = gbp_factor_pivot['Raw_RMOM'].sort_index(axis=1)
gbp_factor_rawrmom = gbp_factor_rawrmom[gbp_factor_rawrmom.index >= '2015-08-12']
gbp_factor_mom = pd.concat([gbp_factor_rawsmom, gbp_factor_rawrmom],axis=0).sort_index(axis=1)
gbp_factor_mom[gbp_factor_mom < 0] = 0
gbp_factor_mom_wt = gbp_factor_mom
gbp_factor_mom_wt[gbp_factor_mom.sum(axis=1)>0] = gbp_factor_mom.div(gbp_factor_mom.sum(axis=1),axis=0)
gbp_mom_score = np.multiply(gbp_factor_mom_wt, gbp_return)
gbp_mom_sum = gbp_mom_score.sum(axis=1)
gbp_mom = (1.0+gbp_mom_sum).cumprod(axis=0)
# AUD
aud_factor_rawsmom = aud_factor_pivot['Raw_SMOM'].sort_index(axis=1).shift(1)
aud_factor_rawsmom = aud_factor_rawsmom[aud_factor_rawsmom.index < '2015-08-12']
aud_mom_table = aud_factor_pivot['Scaled_SMOM'].sort_index(axis=1).shift(1)
aud_mom_table = aud_mom_table[aud_mom_table.index < '2015-08-12']
aud_mom_table[aud_mom_table < 0] = 0
aud_mom_table[aud_mom_table > 0] = 1
aud_factor_rawsmom = np.multiply(aud_mom_table, aud_factor_rawsmom)
aud_factor_rawrmom = aud_factor_pivot['Raw_RMOM'].sort_index(axis=1)
aud_factor_rawrmom = aud_factor_rawrmom[aud_factor_rawrmom.index >= '2015-08-12']
aud_factor_mom = pd.concat([aud_factor_rawsmom, aud_factor_rawrmom],axis=0).sort_index(axis=1)
aud_factor_mom[aud_factor_mom < 0] = 0
aud_factor_mom_wt = aud_factor_mom
aud_factor_mom_wt[aud_factor_mom.sum(axis=1)>0] = aud_factor_mom.div(aud_factor_mom.sum(axis=1),axis=0)
aud_mom_score = np.multiply(aud_factor_mom_wt, aud_return)
aud_mom_sum = aud_mom_score.sum(axis=1)
aud_mom = (1.0+aud_mom_sum).cumprod(axis=0)
# Compile Daily and Month End Series into 2 Dataframes
mom_df = pd.concat([jpy_mom, cad_mom, usd_mom, gbp_mom, eur_mom, aud_mom], axis=1)
mom_df.columns = ['JPY','CAD','USD','GBP','EUR','AUD']
mom_df = mom_df.asfreq("D",method='ffill')
mom_df = mom_df.fillna(method='ffill')
mom_df['AC EW Mean'] = mom_df.mean(axis=1)
mom_df.head()
mom_monthly_df = mom_df.resample('M').last()

#######################################MONTHLY RETURNS###########################################
ff_df_beta = rawbeta_monthly_df
ff_df_beta = ff_df_beta.drop(['AUD','AC EW Mean Beta'], axis=1)

ff_df_value = rawvalue_monthly_df
ff_df_value = ff_df_value.drop(['AUD','AC EW Mean Value'], axis=1)

ff_df_mom = mom_monthly_df
ff_df_mom = ff_df_mom.drop(['AUD','AC EW Mean'], axis=1)
#######################################PRINT OUTPUT###########################################
# GBP
writer = pd.ExcelWriter('NewReporting\GBP\gbp_return.xlsx')
gbp_return.to_excel(writer,'GBPRET')
writer.save()
# RawBeta
writer = pd.ExcelWriter('NewReporting\GBP\gbp_factor_rawbeta_wt.xlsx')
gbp_factor_rawbeta_wt.to_excel(writer,'betaweights')
writer.save()
# RawValue
writer = pd.ExcelWriter('NewReporting\GBP\gbp_factor_rawvalue_wt.xlsx')
gbp_factor_rawvalue_wt.to_excel(writer,'valueweights')
writer.save()
# Momentum
writer = pd.ExcelWriter('NewReporting\GBP\gbp_factor_mom_wt.xlsx')
gbp_factor_mom_wt.to_excel(writer,'momweights')
writer.save()
# CAD
writer = pd.ExcelWriter('NewReporting\CAD\cad_return.xlsx')
cad_return.to_excel(writer,'cadRET')
writer.save()
# RawBeta
writer = pd.ExcelWriter('NewReporting\CAD\cad_factor_rawbeta_wt.xlsx')
cad_factor_rawbeta_wt.to_excel(writer,'betaweights')
writer.save()
# RawValue
writer = pd.ExcelWriter('NewReporting\CAD\cad_factor_rawvalue_wt.xlsx')
cad_factor_rawvalue_wt.to_excel(writer,'valueweights')
writer.save()
# Momentum
writer = pd.ExcelWriter('NewReporting\CAD\cad_factor_mom_wt.xlsx')
cad_factor_mom_wt.to_excel(writer,'momweights')
writer.save()
# JPY
writer = pd.ExcelWriter('NewReporting\JPY\jpy_return.xlsx')
jpy_return.to_excel(writer,'jpyRET')
writer.save()
# RawBeta
writer = pd.ExcelWriter('NewReporting\JPY\jpy_factor_rawbeta_wt.xlsx')
jpy_factor_rawbeta_wt.to_excel(writer,'betaweights')
writer.save()
# RawValue
writer = pd.ExcelWriter('NewReporting\JPY\jpy_factor_rawvalue_wt.xlsx')
jpy_factor_rawvalue_wt.to_excel(writer,'valueweights')
writer.save()
# Momentum
writer = pd.ExcelWriter('NewReporting\JPY\jpy_factor_mom_wt.xlsx')
jpy_factor_mom_wt.to_excel(writer,'momweights')
writer.save()
# USD
writer = pd.ExcelWriter('NewReporting\USD\usd_return.xlsx')
usd_return.to_excel(writer,'usdRET')
writer.save()
# RawBeta
writer = pd.ExcelWriter('NewReporting\USD\usd_factor_rawbeta_wt.xlsx')
usd_factor_rawbeta_wt.to_excel(writer,'betaweights')
writer.save()
# RawValue
writer = pd.ExcelWriter('NewReporting\USD\usd_factor_rawvalue_wt.xlsx')
usd_factor_rawvalue_wt.to_excel(writer,'valueweights')
writer.save()
# Momentum
writer = pd.ExcelWriter('NewReporting\USD\usd_factor_mom_wt.xlsx')
usd_factor_mom_wt.to_excel(writer,'momweights')
writer.save()
# EUR
writer = pd.ExcelWriter('NewReporting\EUR\eur_return.xlsx')
eur_return.to_excel(writer,'eurRET')
writer.save()
# RawBeta
writer = pd.ExcelWriter('NewReporting\EUR\eur_factor_rawbeta_wt.xlsx')
eur_factor_rawbeta_wt.to_excel(writer,'betaweights')
writer.save()
# RawValue
writer = pd.ExcelWriter('NewReporting\EUR\eur_factor_rawvalue_wt.xlsx')
eur_factor_rawvalue_wt.to_excel(writer,'valueweights')
writer.save()
# Momentum
writer = pd.ExcelWriter('NewReporting\EUR\eur_factor_mom_wt.xlsx')
eur_factor_mom_wt.to_excel(writer,'momweights')
writer.save()




writer = pd.ExcelWriter('ACValueDaily.xlsx')
rawvalue_df.to_excel(writer, 'Value')
writer.save()

writer = pd.ExcelWriter('ACMomDaily.xlsx')
mom_df.to_excel(writer, 'RawData')
writer.save()

writer = pd.ExcelWriter('ACLowBetaDaily.xlsx')
rawbeta_df.to_excel(writer, 'Beta')
writer.save()

writer = pd.ExcelWriter('ACValueMonthly.xlsx')
ff_df_value.to_excel(writer, 'Value')
writer.save()

writer = pd.ExcelWriter('ACLowBetaMonthly.xlsx')
ff_df_beta.to_excel(writer, 'Beta')
writer.save()

writer = pd.ExcelWriter('ACMOMMonthly.xlsx')
ff_df_mom.to_excel(writer, 'MOM')
writer.save()


writer = pd.ExcelWriter('gbp_return.xlsx')
gbp_return.to_excel(writer,'RET')
writer.save()

writer = pd.ExcelWriter('gbp_factor_rawbeta_wt.xlsx')
gbp_factor_rawbeta_wt.to_excel(writer,'weights')
writer.save()

writer = pd.ExcelWriter('jpyrawbeta.xlsx')
# Series
jpy_rawbeta.to_frame().to_excel(writer,'returns')
writer.save()

##################################### UNSCALED MOMENTUM ################################################

# jpy_factor_rawsmom = jpy_factor_pivot['Raw_SMOM'].sort_index(axis=1)
# jpy_factor_rawsmom = jpy_factor_rawsmom[jpy_factor_rawsmom.index < '2015-08-11']
# jpy_mom_table = jpy_factor_pivot['Scaled_SMOM'].sort_index(axis=1)
# jpy_mom_table = jpy_mom_table[jpy_mom_table.index < '2015-08-11']
# jpy_mom_table[jpy_mom_table < 0] = 0
# jpy_mom_table[jpy_mom_table > 0] = 1
# jpy_factor_rawsmom = np.multiply(jpy_mom_table, jpy_factor_rawsmom)
# jpy_factor_rawrmom = jpy_factor_pivot['Raw_RMOM'].sort_index(axis=1)
# jpy_factor_rawrmom = jpy_factor_rawrmom[jpy_factor_rawrmom.index >= '2015-08-11']
# jpy_factor_mom = pd.concat([jpy_factor_rawsmom, jpy_factor_rawrmom],axis=0).sort_index(axis=1)
# jpy_factor_mom[jpy_factor_mom < 0] = 0
# jpy_mom_score = np.multiply(jpy_factor_mom, jpy_return)
# jpy_mom_sum = jpy_mom_score.sum(axis=1)
# jpy_mom = (1.0+jpy_mom_sum).cumprod(axis=0)

# CAD
# cad_factor_rawsmom = cad_factor_pivot['Raw_SMOM'].sort_index(axis=1)
# cad_factor_rawsmom = cad_factor_rawsmom[cad_factor_rawsmom.index < '2015-08-11']
# # If scaled momentum was negative, then invest in Raw MOM
# cad_mom_table = cad_factor_pivot['Scaled_SMOM'].sort_index(axis=1)
# cad_mom_table = cad_mom_table[cad_mom_table.index < '2015-08-11']
# cad_mom_table[cad_mom_table < 0] = 0
# cad_mom_table[cad_mom_table > 0] = 1
# cad_factor_rawsmom = np.multiply(cad_mom_table, cad_factor_rawsmom)
# cad_factor_rawrmom = cad_factor_pivot['Raw_RMOM'].sort_index(axis=1)
# cad_factor_rawrmom = cad_factor_rawrmom[cad_factor_rawrmom.index >= '2015-08-11']
# cad_factor_mom = pd.concat([cad_factor_rawsmom, cad_factor_rawrmom],axis=0).sort_index(axis=1)
# cad_factor_mom[cad_factor_mom < 0] = 0
# cad_factor_mom_wt = cad_factor_mom
# cad_factor_mom_wt[cad_factor_mom.sum(axis=1)>0] = cad_factor_mom.div(cad_factor_mom.sum(axis=1),axis=0)
# cad_mom_score = np.multiply(cad_factor_mom_wt, cad_return)
# cad_mom_sum = cad_mom_score.sum(axis=1)
# cad_mom = (1.0+cad_mom_sum).cumprod(axis=0)