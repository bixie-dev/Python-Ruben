import  numpy as  np
import pandas as  pd
import csv
import os
import datetime

#source settings
routesource=['c:/sources'] #Here is the route of the source files and can be csv, xls or xlsx. Multiple files can be entered.
columnstochecksourceinventory='R'  #The column where it will take the value for the inventory from “routesource” files
columnstochecksourcesku='O'   #Format to enter is KK,AA,This is a value to search in source file. If is written “whole” it will search in whole columns.
#if columnstochecksourcesku=='whole':
sheetnamesource='book2'  #The sheet name to work for source files.
rowstartsource=0  #the row where it will start the information and skip the headers for source files.


#destination file settings
routefiletocheck=['c:/destinations']   #the file or files where we will search for the elements route of the files to check
columnstocheckfile ='O'  # Where to check the value in destination files. If is written “whole” it will search in whole ( SKU search ) .    
columnstocheckdestinationinventory='R'  # Where to check the value in destination files.( Stock Value)
sheetnamedest='book2'  #The sheet name to work for destination files.
rowstartdest=0 #the row where it will start the information and skip the headers for destination files.

stockquantitytocheck=50   #Value to enter to check minimum stock to search in “columnstocheckfile”


#safe files settings
safefiles=['safe.xlsx']
safefilecolumns='B,D'
safefilesenable=False

filtersourcetoinventory=True  #If says True then it will not consider the values with lower value entered in
filterdestinationinventory=True  #If says True it will remove all the rows in destination files with “stockquantitytocheck” lower than the value entered.

coldestinationmode='fixed'  #this will be a variable with how the destination mode columns will be arrive to“columnstocheckfile”.

if coldestinationmode=='fixed':
    decimalsdest=2

percetnadddestcol=True   #f “percent” is true then it will increase or decrease a percent from “percentstatus” to the “columnstocheckfile”

if percetnadddestcol==True:
    percent=5
    percentstatus=["increase", "decrease"]



removerows=True


def readFolder(folder):
    dir_list = os.listdir(folder)
    return dir_list



def read_file(filename,filetype,sheetname,start):
    dataframe=None

    name=str(filename+filetype).strip()
    if os.path.exists(name):
    
        if filetype=='.xlsx' or  filetype=='.xls':
            dataframe=pd.read_excel(name,sheet_name = sheetname,skiprows = start) 
            initcolumns=dataframe.columns
            number_of_columns=len(dataframe.columns)
            dataframe.columns =[ excel_style(i) for i in range(1,number_of_columns+1)]
        
        elif filetype=='.csv':
            dataframe=pd.read_csv(name, sep=',', lineterminator='\r',sheet_name=sheetname ,skiprows = start)    # reading source file with its sheet
            number_of_columns=len(dataframe.columns)
            initcolumns=dataframe.columns
            dataframe.columns =[ excel_style(i) for i in range(1,number_of_columns+1)] 
            
        elif filetype=='.txt':
            dataframe=pd.read_csv(name, sep='\t', lineterminator='\r',skiprows = start)   # reading source file with its sheet
            number_of_columns=len(dataframe.columns)
            initcolumns=dataframe.columns
            dataframe.columns =[ excel_style(i) for i in range(1,number_of_columns+1)]
            dataframe=dataframe.replace(r'\n','', regex=True) 
       
    else:
        print(filename +" does not exist ")
    
  
    return (dataframe,initcolumns)
    
    

def save_to_report(data,source,des,removed_or_not,fileextension):
    output_path = 'Reports' 
    if not os.path.exists(output_path):
        os.makedirs(output_path)

    if fileextension.lower()=='.xlsx':
        data.to_excel(output_path+"/"+source+des+removed_or_not+".xlsx", index = False) 
        
    elif fileextension.lower()=='.csv':
        data.to_csv(output_path+"/"+source+des+removed_or_not+".csv", index = False)
            
    else:
        data.to_csv(output_path+"/"+source+des+removed_or_not+".txt", sep="\t", index = False)


def  save_to_output(data,filename,fileextension):
    output_path = 'Output' 
    if not os.path.exists(output_path):
        os.makedirs(output_path)
       
    if fileextension.lower()=='.xlsx':
        data.to_excel(output_path+"/"+filename+".xlsx", index = False) 
        
    elif fileextension.lower()=='.csv':
        data.to_csv(output_path+"/"+filename+".csv", index = False)
            
    else:
        data.to_csv(output_path+"/"+filename+".txt", sep="\t", index = False)
    
            
 

    
    
def excel_style(col):
    LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    """ Convert given row and column number to an Excel-style cell name. """
    result = []
    while col:
        col, rem = divmod(col-1, 26)
        result[:0] = LETTERS[rem]
    return ''.join(result)
    
    
    
def search(A,B,sheetnameA,startA,sheetnameB,startB,sku1,sku2,inv1,inv2):
    for k in range(0,len(A)):
        files=readFolder(A[k])

        for i in range(0,len(files)):
            filenames, file_extensions = os.path.splitext(A[k]+'/'+files[i])
            dfsource,initcolumns=read_file(filenames,file_extensions,sheetnameA,startA)
    
            print("processing "+filenames+file_extensions+" . . . with ")
            
            for j in range(0,len(B)):
                filesd=readFolder(B[j])

                for ik in range(0,len(filesd)):
                    filenamed, file_extensiond = os.path.splitext(B[ik]+'/'+filesd[i])
                    dfdest,initcolumns=read_file(filenamed,file_extensiond,sheetnameB,startB)
        

                    print("  processing "+filenamed+file_extensiond+"")
                    
                    

                    
                    if dfsource is not None and dfdest is not None:
                        items_inv_source=[]
                        items_sku_source=[]
                        
                        items_inv_destination=[]
                        items_sku_destination=[]
                
                        if  sku1.lower()=='whole':
                            items_sku_source=dfsource.columns
                        else:
                            items_sku_source=sku1.split(',')
                        
                        if  sku2.lower()=='whole':
                            items_sku_destination=dfdest.columns
                        else:
                            items_sku_destination=sku2.split(',')
                            
                         
                        if inv2.lower()=='whole':
                            items_inv_destination=dfdest.columns
                        else:
                            items_inv_destination=inv2.split(',')
                            
                            
                        if inv1.lower()=='whole':
                            items_inv_source=dfsource.columns
                        else:
                            items_inv_source=inv1.split(',')
                            
                           
                        
                      
                                
                                
                        filtered=dfdest
                        dataframe_found=pd.DataFrame()  
                            
                        for itemrow_frame1 in items_sku_source:
                           
                            for item2_frame2 in items_sku_destination:
                               
                                dataframe_found=pd.concat([dataframe_found,filtered[filtered[item2_frame2].isin(dfsource[itemrow_frame1].values)==True]])
                                filtered=filtered[filtered[item2_frame2].isin(dfsource[itemrow_frame1].values)==True]
                                
                               
                                
                        found_values=dataframe_found.drop_duplicates()    
                        not_found_values=pd.concat([dfdest,dataframe_found]).drop_duplicates(keep=False)
                      
                        
                        
                        filtered_found=found_values
                        final_found_values=pd.DataFrame()
                        
                        
                        print(filtered_found)
                        for item2_frame2 in items_inv_destination:
                            filtered_found=found_values[found_values[item2_frame2]<stockquantitytocheck]
                            final_found_values=pd.concat([final_found_values,filtered_found])
                            
                        
                        removed_values=pd.concat([not_found_values,final_found_values.drop_duplicates()])
                     
                        
                        not_removed_values=pd.concat([dfdest,removed_values]).drop_duplicates(keep=False)
                        
                        
                      
                        if len(removed_values)>0:
                            removed="removed"
                        else:
                            removed="not removed"
                            
                        
                        removed_values.columns=initcolumns
                        not_removed_values.columns=initcolumns
                        
                        
                        save_to_output(not_removed_values,files[i]+filesd[j]+removed+datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),file_extensiond)
                   
                        save_to_report(removed_values,files[i]+filesd[j]+removed+datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),"",removed,file_extensiond)
                    
                        print("Procssing "+filenames+file_extensions+"with"+filenamed+file_extensiond+" is complete")
                    
                        
    
def update_inventory(A,B,sheetnameA,startA,sheetnameB,startB,sku1,sku2,inv1,inv2):
    fillzeroes=True
    for k in range(0,len(A)):
        files=readFolder(A[k])

        for i in range(0,len(files)):
            filenames, file_extensions = os.path.splitext(A[k]+'/'+files[i])
            dfsource,initcolumns=read_file(filenames,file_extensions,sheetnameA,startA)
    
            print("processing "+filenames+file_extensions+" . . . with ")
            
            for j in range(0,len(B)):
                filesd=readFolder(B[j])

                for ik in range(0,len(filesd)):
                    filenamed, file_extensiond = os.path.splitext(B[ik]+'/'+filesd[i])
                    dfdest,initcolumns=read_file(filenamed,file_extensiond,sheetnameB,startB)
        

                    print("  processing "+filenamed+file_extensiond+"")
     

            
                    if dfsource is not None and dfdest is not None:
                        items_inv_source=[]
                        items_sku_source=[]
                        
                        items_inv_destination=[]
                        items_sku_destination=[]
                
                        if  sku1.lower()=='whole':
                            items_sku_source=dfsource.columns
                        else:
                            items_sku_source=sku1.split(',')
                        
                        if  sku2.lower()=='whole':
                            items_sku_destination=dfdest.columns
                        else:
                            items_sku_destination=sku2.split(',')
                            
                         
                        if inv2.lower()=='whole':
                            items_inv_destination=dfdest.columns
                        else:
                            items_inv_destination=inv2.split(',')
                            
                            
                        if inv1.lower()=='whole':
                            items_inv_source=dfsource.columns
                        else:
                            items_inv_source=inv1.split(',')
                            
                                
                        filtered=dfdest
                        dataframe_found=pd.DataFrame()  
                            
                        for itemrow_frame1 in items_sku_source:
                           
                            for item2_frame2 in items_sku_destination:
                               
                                dataframe_found=pd.concat([dataframe_found,filtered[filtered[item2_frame2].isin(dfsource[itemrow_frame1].values)==True]])
                                filtered=filtered[filtered[item2_frame2].isin(dfsource[itemrow_frame1].values)==True]
                                
                               
                                
                        found_values=dataframe_found.drop_duplicates()    
                        not_found_values=pd.concat([dfdest,dataframe_found]).drop_duplicates(keep=False)
                      
                        
                        
                        filtered_found=found_values
                        final_found_values=pd.DataFrame()
                        
                        
                     
                        for item2_frame2 in items_inv_destination:
                            filtered_found=found_values[found_values[item2_frame2]==0]
                            final_found_values=pd.concat([final_found_values,filtered_found])
                            
                        
                        removed_values=pd.concat([not_found_values,final_found_values.drop_duplicates()])
                        
                      
                        
                        
                        not_removed_values=pd.concat([dfdest,removed_values]).drop_duplicates(keep=False)
                        
          
                                                       
                        for kk,ll in zip(items_sku_destination,items_inv_destination):
                            v=not_removed_values[kk]
                            for index in v:
                                for ii,jj in zip(items_sku_source,items_inv_source):
                                    x=dfsource[dfsource[ii]==index][jj]
                               
                                    y=not_removed_values[not_removed_values[kk]==index][ll]
                                   
                                    not_removed_values.loc[not_removed_values[kk]==index,ll]=float(x)
                                    
                      
                            
                  
                        
                        
                        
                        
                        
                        
                        
                        
                        if fillzeroes:
                            for k in items_inv_destination:
                                removed_values[k] = 0
                   
                            not_removed_values=pd.concat([not_removed_values,removed_values])
                        
                        
                   
           
                      
                      
                        if len(removed_values)>0:
                            removed="removed"
                        else:
                            removed="not removed"
                            
                       
                            
             
                            
                        removed_values.columns=initcolumns
                        
                        not_removed_values.columns=initcolumns
                        
                        
                       
                        
                        
                       
                        
                        
                        save_to_output(not_removed_values,files[i]+filesd[j]+removed+datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),file_extensiond)
                   
                        save_to_report(removed_values,files[i]+filesd[j]+removed+datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),"",removed,file_extensiond)
                    
                        print("Procssing "+filenames+file_extensions+"with"+filenamed+file_extensiond+" is complete")
                        
                      
                      
     
                            
            
    
    
    
    
    
'''
Script Starts Here
'''


howtoprocess='cleandest'

#howtoprocess='cleansourc'

#howtoprocess='full - search from source to destination files'

#howtoprocess='full - search from destination to source files'

#howtoprocess='updateinventory source files to destination files'

#howtoprocess='updateinventory destination files to source files'




if howtoprocess=='cleandest':

    for k in range(0,len(routefiletocheck)):
        files=readFolder(routefiletocheck[k])

        for i in range(0,len(files)):
            filename, file_extension = os.path.splitext(routefiletocheck[k]+'/'+files[i])
            df,initcolumns=read_file(filename,file_extension,sheetnamedest,rowstartdest)
            
            print("processing "+filename+file_extension+" . . .",df)
            
            if df is not None:
            
                if columnstocheckdestinationinventory.lower()=='whole':
                    items=data_frame_destination.columns
                else:
                    items=columnstocheckdestinationinventory.split(',')
            
            
                data_frame_valid_values=df
                
                for ik in items:
                    data_frame_valid_values=data_frame_valid_values[data_frame_valid_values[ik]>stockquantitytocheck]
                
                removed_values=pd.concat([df,data_frame_valid_values]).drop_duplicates(keep=False)
           
               
                if len(removed_values)>0:
                    removed="removed"
                else:
                    removed="not removed"
                    
                    
                removed_values.columns=initcolumns
                data_frame_valid_values.columns=initcolumns
            
            
                save_to_output(data_frame_valid_values,files[i]+datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),file_extension)
               
                save_to_report(removed_values,files[i]+datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),"",removed,file_extension)
                
                print("Procssing "+filename+file_extension+" is complete &  saved to "+files[i])
           
    
   



if howtoprocess=='cleansourc':

    for k in range(0,len(routesource)):
        files=readFolder(routesource[k])

        for i in range(0,len(files)):
            filename, file_extension = os.path.splitext(routesource[k]+'/'+files[i])
            df,initcolumns=read_file(filename,file_extension,sheetnamesource,rowstartsource)
            
        
            print("processing "+filename+file_extension+" . . .")
            
            if df is not None:
            
                if columnstochecksourceinventory.lower()=='whole':
                    items=data_frame_destination.columns
                else:
                    items=columnstochecksourceinventory.split(',')
            
            
                data_frame_valid_values=df
                for ik in items:
                    data_frame_valid_values=data_frame_valid_values[data_frame_valid_values[ik]>stockquantitytocheck]
                
                removed_values=pd.concat([df,data_frame_valid_values]).drop_duplicates(keep=False)
           
               
                if len(removed_values)>0:
                    removed="removed"
                else:
                    removed="not removed"
                    
                removed_values.columns=initcolumns
                data_frame_valid_values.columns=initcolumns
            
            
                save_to_output(data_frame_valid_values,files[i]+datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),file_extension)
               
                save_to_report(removed_values,files[i]+datetime.datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),"",removed,file_extension)
                
                print("Procssing "+filename+file_extension+" is complete")



if howtoprocess=='full - search from source to destination files':

    search(routesource,routefiletocheck,sheetnamesource,rowstartsource,sheetnamedest,rowstartdest,columnstochecksourcesku,columnstocheckfile,columnstochecksourceinventory,columnstocheckdestinationinventory)



if howtoprocess=='full - search from destination to source files':

    search(routefiletocheck,routesource,sheetnamedest,rowstartdest,sheetnamesource,rowstartsource,columnstocheckfile,columnstochecksourcesku,columnstocheckdestinationinventory,columnstochecksourceinventory)
        
      


if howtoprocess=='updateinventory source files to destination files':

     update_inventory(routesource,routefiletocheck,sheetnamesource,rowstartsource,sheetnamedest,rowstartdest,columnstochecksourcesku,columnstocheckfile,columnstochecksourceinventory,columnstocheckdestinationinventory)
    


if  howtoprocess=='updateinventory destination files to source files':
    update_inventory(routefiletocheck,routesource,sheetnamedest,rowstartdest,sheetnamesource,rowstartsource,columnstocheckfile,columnstochecksourcesku,columnstocheckdestinationinventory,columnstochecksourceinventory)
    

 




'''

Script Ends Here

'''