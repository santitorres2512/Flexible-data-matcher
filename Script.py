import pandas as pd
from fuzzywuzzy import fuzz
import numpy as np



def match(Col1,Col2): ## Function will receive 2 columns to compare
    tosalesforce=[]
    nottosalesforcematch=[]  ## dynamic list creation
    nottosalesforce=[]
    nottosalesforceratio=[]
    #i=i+1
    for n in Col1:
        print("Company Name: ",n)   ##Grab each company name
        result=[(fuzz.partial_ratio(n, n2),n2)     
                for n2 in Col2 if fuzz.ratio(n, n2)>87  ##And apply fuzz method to find only matches with more than 87% similarity                     ##################################### RATIO PERCENTAGE OF TOLERANCE###############################
               ]
        
        
        if len(result):    ## if any matches
            result.sort()  
            print("Approximate Company Name found, choosing best match...")
            print('Matches: {}'.format(result))
            print("Best Match : {}".format(result[-1][1]))
            print("")
            nottosalesforcematch.append(result[-1][1])    ##append the best result as the best match
            nottosalesforce.append(n)
            nottosalesforceratio.append(result[-1][0])
            #print(Col2.ID.loc[result[-1][1]])
            #not2salesforceIDs.append()
            
        else:
            print("Company Name not found, adding to the 'To SalesForce List'")
            print("")
            tosalesforce.append(n)
    return [tosalesforce,nottosalesforce,nottosalesforcematch,nottosalesforceratio]  #returns list with all results

##let's open the CSV

with open("DataBase_File.csv", encoding="utf8", errors='ignore') as DataBase:
    DataDF = pd.DataFrame(DataBase, columns = ["Company_Name"])

#DataBase = pd.read_csv("List2_DNC.csv",)

with open("Input_File.csv", encoding="utf8", errors='ignore') as InputData:
    InputDF = pd.DataFrame(InputData, columns = ["Company Name"])

#InputData = pd.read_excel("List1_DNC.xlsx").decode('unicode_escape')

#print(horizontal_stack3)
#InputDF = pd.DataFrame(InputData, columns = ["Company_Name"])

InputList = InputDF.values.tolist()
DataList = DataDF.values.tolist()

## and let's start removing some common words from Companies, for better matches
## the desired result is just the company name, like coca cola without any LLC. CO. 

df1=pd.DataFrame(InputList)
df2=pd.DataFrame(DataList)


InputContacts = len(df1)
DataBaseContacts = len(df2)

df1.columns=['Name']
df2.columns=['Name']


df2['Original_Name_Masterlist'] = df2['Name']

#df1O=pd.DataFrame(InputList)
#df2O=pd.DataFrame(DataList)



#df2O.columns=['Name']

#df2['ID'] = df2.groupby('Name',sort=False).ngroup()
df2 = df2.drop_duplicates(subset=['Name'], keep='first') ##removes duplicates
#writer = pd.ExcelWriter('Test.xlsx', engine='xlsxwriter')
#horizontal4.to_excel(writer, sheet_name='Email_Match', index=False)


#df2.to_excel(writer, sheet_name='Statistics Table', index=False)

#nottosalesforceDF.to_excel(writer, sheet_name='Not for SalesForce', index=False)

#tosalesforcelist.to_excel(writer, sheet_name='For SalesForce', index=False)


#writer.save()




df1['Name'].replace(',','', regex=True, inplace=True)                          #############################KEY COMMON WORDS TO EXCLUDE ON MATCHING STAGE###################################
df2['Name'].replace(',','', regex=True, inplace=True)

df1['Name'].replace('\.','', regex=True, inplace=True)
df2['Name'].replace('\.','', regex=True, inplace=True)

df1['Name'].replace('Inc','', regex=True, inplace=True)
df2['Name'].replace('Inc','', regex=True, inplace=True)

df1['Name'].replace('Company','', regex=True, inplace=True)
df2['Name'].replace('Company','', regex=True, inplace=True)

df1['Name'].replace('Corporation','', regex=True, inplace=True)
df2['Name'].replace('Corporation','', regex=True, inplace=True)

df1['Name'].replace('Foods','', regex=True, inplace=True)
df2['Name'].replace('Foods','', regex=True, inplace=True)

df1['Name'].replace('Products','', regex=True, inplace=True)
df2['Name'].replace('Products','', regex=True, inplace=True)

df1['Name'].replace('LLC','', regex=True, inplace=True)
df2['Name'].replace('LLC','', regex=True, inplace=True)

df1['Name'].replace('Dairy','', regex=True, inplace=True)
df2['Name'].replace('Dairy','', regex=True, inplace=True)

df1['Name'].replace('international','', regex=True, inplace=True)
df2['Name'].replace('International','', regex=True, inplace=True)

df1['Name'].replace('Brands','', regex=True, inplace=True)
df2['Name'].replace('Brands','', regex=True, inplace=True)

df1['Name'].replace('North America','', regex=True, inplace=True)
df2['Name'].replace('North America','', regex=True, inplace=True)

df1['Name'].replace('Beverage','', regex=True, inplace=True)
df2['Name'].replace('Beverage','', regex=True, inplace=True)

df1['Name'].replace('Bakery','', regex=True, inplace=True)
df2['Name'].replace('Bakery','', regex=True, inplace=True)

df1['Name'].replace('Sugar','', regex=True, inplace=True)
df2['Name'].replace('Sugar','', regex=True, inplace=True)

df1['Name'].replace('Chocolate','', regex=True, inplace=True)
df2['Name'].replace('Chocolate','', regex=True, inplace=True)

df1['Name'].replace('Food','', regex=True, inplace=True)
df2['Name'].replace('Food','', regex=True, inplace=True)

df1['Name'].replace('Cheese','', regex=True, inplace=True)
df2['Name'].replace('Cheese','', regex=True, inplace=True)

df1['Name'].replace('Seafood','', regex=True, inplace=True)
df2['Name'].replace('Seafood','', regex=True, inplace=True)

df1['Name'].replace('Ice Cream','', regex=True, inplace=True)
df2['Name'].replace('Ice Cream','', regex=True, inplace=True)

df1['Name'].replace('Factory','', regex=True, inplace=True)
df2['Name'].replace('Factory','', regex=True, inplace=True)

df1['Name'].replace('Creamery','', regex=True, inplace=True)
df2['Name'].replace('Creamery','', regex=True, inplace=True)

df1['Name'].replace('Grocers','', regex=True, inplace=True)
df2['Name'].replace('Grocers','', regex=True, inplace=True)

df1['Name'].replace('Vegetables','', regex=True, inplace=True)
df2['Name'].replace('Vegetables','', regex=True, inplace=True)

df1['Name'].replace('Market','', regex=True, inplace=True)
df2['Name'].replace('Market','', regex=True, inplace=True)

df1['Name'].replace('Producers','', regex=True, inplace=True)
df2['Name'].replace('Producers','', regex=True, inplace=True)

df1['Name'].replace('Candy','', regex=True, inplace=True)
df2['Name'].replace('Candy','', regex=True, inplace=True)

df1['Name'].replace('Fisheries','', regex=True, inplace=True)
df2['Name'].replace('Fisheries','', regex=True, inplace=True)

df1['Name'].replace('Bakeries','', regex=True, inplace=True)
df2['Name'].replace('Bakeries','', regex=True, inplace=True)

df2['Match']=df2['Name']


#df1.replace('\.','', regex=True, inplace=True)
#df2.replace('\.','', regex=True, inplace=True)


#df2.drop_duplicates(subset ="Name", keep ='first', inplace = True)


#df2O.columns=["Company_Name"]

#df2 = pd.concat([df2, df2O], axis=1)


#df2['ID'] = df2.groupby('Name',sort=False).ngroup()



df2Unique = df2

UniqueDataBaseContacts = len(df2Unique)


lista = match(df1.Name,df2.Name)
##creates the new columns for the matches found and also the similarity ratio

tosalesforcelist = pd.DataFrame(lista[0], columns = ["Name"])
nottosalesforcelist = pd.DataFrame(lista[1], columns = ["Name"])
nottosalesforcematchlist = pd.DataFrame(lista[2], columns = ["Match"])
nottosalesforceratiolist = pd.DataFrame(lista[3], columns = ["Ratio"])


nottosalesforceDF = pd.concat([nottosalesforcelist, nottosalesforcematchlist], axis=1)
nottosalesforceDF = pd.concat([nottosalesforceDF, nottosalesforceratiolist], axis=1)
#nottosalesforceDF = pd.concat([nottosalesforceDF, df2.Company_Name], axis=1)

#nottosalesforceDF['ID'] = nottosalesforceDF.apply(lambda x: (np.where(x['Match'] == df2['Name'])),axis=1)
#nottosalesforceDF['ID'] = nottosalesforceDF['ID'].astype(str)
#nottosalesforceDF['ID'] = nottosalesforceDF['ID'].str.extract('(\d+)', expand=False)

#LastID = df2['ID'].max()


ListLength = len(tosalesforcelist["Name"])
#tosalesforcelistIDs = pd.DataFrame('New_IDs': range(LastID, ListLength + LastID + 1))
#sample_data = pd.DataFrame({'B' : range(1, N + 1)})
#tosalesforcelistIDs = pd.DataFrame({        
#        'ID' : range(LastID+1, ListLength + LastID+1)}
#     )

#.concat([tosalesforcelist, tosalesforcelistIDs], axis=1)
#tosalesforcelist = pd.concat([tosalesforcelist, tosalesforcelistIDs], axis=1)


## merges the information making sure all is on is place

nottosalesforceDF = nottosalesforceDF.merge(df1, on='Name',how = 'left')
nottosalesforceDF = nottosalesforceDF.merge(df2, on='Match',how = 'left')

tosalesforcelist = tosalesforcelist.merge(df1, on='Name',how = 'left')
#del tosalesforcelist['Original_Name_y']
#print("To SalesForce list -->",tosalesforcelist)
#nottosalesforceDF['Company_Name']=df2O
#nottosalesforceDF['Original_Name'] = nottosalesforceDF.apply(lambda x: (np.where(x['Assigned_ID'] == df2['ID'])),axis=1)
print(nottosalesforceDF)

print(tosalesforcelist)

#LastIDtosalesforce = tosalesforcelist['ID'].max()
#print(match(df1.Name,df2.Name))

amount2salesforce = len(tosalesforcelist)
amountnot2salesforce = len(nottosalesforceDF)

table = [['Input DNC Amount', InputContacts], ['Initial DataBase DNC Amount ', DataBaseContacts],  ['Amount DNC for SalesForce', amount2salesforce], ['Amount DNC not for SalesForce', amountnot2salesforce], ['Last ID Assigned on DB'], ['Last ID assigned to new DNCs']] 

finaldf3 = pd.DataFrame(table, columns = ['Category','Value'])


#df1['Original_Name'].replace('"','', regex=True, inplace=True)


del nottosalesforceDF['Name_y']
#del nottosalesforceDF['ID_y']

writer = pd.ExcelWriter('Output_File.xlsx', engine='xlsxwriter')
#horizontal4.to_excel(writer, sheet_name='Email_Match', index=False)

##Finally writes the CSV with all the results

finaldf3.to_excel(writer, sheet_name='Statistics Table', index=False)

nottosalesforceDF.to_excel(writer, sheet_name='Matched Records', index=False)

tosalesforcelist.to_excel(writer, sheet_name='Independent Records', index=False)


writer.save()


#print("Last ID",LastID)
#def cross_fuzz(df):
#    ct = pd.crosstab(df['Name'], df['Name'])
#    ct = ct.apply(lambda col: [fuzz.ratio(col.name, x) for x in col.index])
#    return ct

#resultado = df1.groupby('Name').apply(cross_fuzz)



#print('input names:', combined_list)
print("*******************************")
#print("in :", InputList)

