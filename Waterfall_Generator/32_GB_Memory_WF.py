'''
Created on May 30, 2017

@author: v52k
'''
import pandas as pd
BP_File = 'C:/Users/v52k/OneDrive/Python/Waterfall_Generator/Lynx_Jupiter_ZaP_BP_10_30_17.xlsx'
pkl_file = 'C:/Users/v52k/OneDrive/Python/Waterfall_Generator/Lynx_SKU_BP.pkl'
Waterfall_File = 'C:/Users/v52k/OneDrive/Python/Waterfall_Generator/Memory_Waterfall.xlsx'
Where_Used='C:/Users/v52k/OneDrive/Python/Waterfall_Generator/Pro4_EnO/Lynx_SKU.xlsx'

lead_time = 10 #weeks
cancel_win = 4 #weeks
Unwanted_Cols = ['SKU #','Country','Type','CPU','CPU (2)','SSD','DRAM','Pallet QTY','Total in SKU BP',' ','SKU Description']

#######################################################################################################
##File Operations
# BP_df = pd.read_excel(BP_File, date_format='%Y%m%d')
#    
#    
# BP_df_fCol= (BP_df['Vertex-Fresh Build BP'].apply(lambda x: pd.Series(str(x).split(' ')))).rename(columns={0:'Snapshot_Date'})
#    
# BP_df['Vertex-Fresh Build BP']=BP_df_fCol['Snapshot_Date']
#    
# BP_df.to_pickle(pkl_file)
     
#      
#######################################################################################################


BP_df= pd.read_pickle(pkl_file)
 
BP_df.iloc[:,0]=pd.to_datetime(BP_df.iloc[:,0]) #Convert first column to Date 
first_col = BP_df.iloc[0,0]

 
BP_df = BP_df.rename(columns={'Vertex-Fresh Build BP':'Snapshot_Date'})

#Get Where Used Data
BOM_df=pd.read_excel(Where_Used)

Bom_df=BOM_df['Item Number']

Subset_df= BP_df #Saving the BP in Subset DF
# Subset_df=Subset_df.set_index('SKU #')
# Subset_df=(Subset_df.loc[BOM_df['Item Number']])


for col_name in Subset_df.columns :
    if col_name in Unwanted_Cols :
        Subset_df=Subset_df.drop(col_name,1)
#         print(col_name)

# Subset_df=Subset_df.astype(str).convert_objects(convert_numeric=True)
# Subset_df=Subset_df.apply(pd.to_numeric, axis = 0, errors='coerce')
Date_df=Subset_df['Snapshot_Date']
Subset_df=Subset_df.drop('Snapshot_Date',1).apply(pd.to_numeric, axis = 0, errors='coerce')
# Subset_df=Subset_df.insert(1,'Snapshot_Date',0)
Subset_df=pd.concat([Date_df, Subset_df], axis = 1, join = 'inner')
print(Subset_df)

sum_df = Subset_df.groupby(['Snapshot_Date']).sum() #Sum all rows with the same key

sum_df.columns = pd.to_datetime(sum_df.columns) #Format all column headers to date

  
# print(sum_df.head())
  
# sum_df=sum_df.resample('W-THU',axis=1).sum() #Group from days to weeks
  
col_index = sum_df.columns.searchsorted(first_col)
  
  
 
sum_df= pd.concat([(sum_df.iloc[:,0:col_index]).sum(axis=1),sum_df.iloc[:,col_index:]], axis = 1, join='inner')
sum_df['Total Demand']=sum_df.sum(axis=1)
sum_df=sum_df.rename(columns={0:'Previous Weeks'})
# 
# new_col = len(sum_df.columns)
# 
# for row in range(0,len(sum_df)):
#     sum_df.loc[sum_df.index[row],new_col]=sum_df.columns[row+1]
#     sum_df.loc[sum_df.index[row],new_col+1]=sum_df.columns[row+1+lead_time]
# sum_df=sum_df.rename(columns={new_col:'Start Date'})
# sum_df=sum_df.rename(columns={new_col+1:'End Date'})
# 
# 
# new_col = len(sum_df.columns)
# 
# 
# for row in range(0,len(sum_df)):
#     sum_df.loc[sum_df.index[row],new_col]=(sum_df.iloc[row,(1+row):(1+row+lead_time)]).sum()
# sum_df=sum_df.rename(columns={new_col:'Demand at LT'})
# 
# new_col = len(sum_df.columns)
# 
# for row in range(0,len(sum_df)):
#     sum_df.loc[sum_df.index[row],new_col]=(sum_df.iloc[row,(1+row):(1+row+cancel_win)]).sum()
# sum_df=sum_df.rename(columns={new_col:'Non Cancellable Demand'})  
# 
#    
# sum_df['Cancellable Demand']=sum_df['Demand at LT']-sum_df['Non Cancellable Demand']

sum_df.to_excel(Waterfall_File)
