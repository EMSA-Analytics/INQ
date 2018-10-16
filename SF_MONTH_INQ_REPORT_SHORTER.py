
# coding: utf-8

# In[1]:


#This is the functional Version. Use this for creating the report
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
get_ipython().run_line_magic('matplotlib', 'inline')


# In[2]:


#disable warnings
import warnings; warnings.simplefilter('ignore')


# The source files for this script are SalesForce reports Interactions.Contacts.csv and Interactions.Opportunity.Term.csv. The source files have to be stored in a NON-GIT folder so that they will never be available to git hub. Whenever the source files are updated into box the old source files should be placed in a depreciated state in a different folder. This script therefore will always pull the most recent Salesforce report results.
# 

# In[3]:


#INT is the contact information for INTERACTIONS
INT=pd.read_csv('C:/Users/mjc341/Desktop/UMAN 1507 Monthly INQ summary Report/Interactions.Contacts.2.csv',skipfooter=5,encoding='latin-1',engine ='python')
#INT=pd.read_csv('C:/Users/sxh706/Desktop/Interactions.by.Month/2018/August/Interactions.Contacts1.csv',skipfooter=5,encoding='latin-1',engine ='python')


# Please Use a new dataframe name if you are going to break any code
# 
# 

# In[4]:


#CON is the CONTACT TERM information
CON=pd.read_csv('C:/Users/mjc341/Desktop/UMAN 1507 Monthly INQ summary Report/Interactions.Opportunity.Term.2.csv',skipfooter=5,encoding='latin-1',engine ='python')
#CON=pd.read_csv('C:/Users/sxh706/Desktop/Interactions.by.Month/2018/August/Interactions.Opportunity.Term.csv',skipfooter=5,encoding='latin-1',engine ='python')


# In[5]:


len(INT)


# In[6]:


len(CON)


# In[7]:


#INT.info()


# In[8]:


INT.columns


# In[9]:


CON.columns


# In[10]:


#CON.info()


# In[11]:


#remove columns with identifing info
CON = CON.drop(columns =['Full Name'])


# In[12]:


#remove columns with identifing info
INT = INT.drop(columns= ['Contact: First Name', 'Contact: Middle Name',
       'Contact: Last Name'])


# In[13]:


#changed the name of columns using a dictionary
INT = INT.rename(columns ={'Contact: Contact ID (18-digit)':'Contact_ID','Interaction: Interaction Name':'Int_Name',
       'Admit Type':'Admit_Type','Contact: EMPLID': 'EMPLID','Contact: Ethnic Group': 'Ethnic_Group', 'Contact: Ethnicity':'Ethnicity',
       'Contact: Gender':'Gender', 'Term':'Term', 'Lead Source':'Lead_Source', 'Interaction: Created Date':'Create_Date',
       'Contact: Mailing Street':'Mailing_Street', 'Contact: Mailing City':'Mailing_City',
       'Contact: Mailing State/Province':'Mailing_State', 'Contact: Mailing Zip/Postal Code':'Mailing_Postal',
       'Contact: Mailing Country':'Mailing_Country', 'Contact: Other City':'Other_City',
       'Contact: Other State/Province':'Other_State', 'Contact: Other Zip/Postal Code':'Other_Postal',
       'Contact: Other Country':'Mailing_Country', 'Contact: Market Segment Code':'Market_Segment_Code',
       'Contact: Market Segment':'Market_Segment'},inplace =False)


# In[14]:


#changed the name of columns
CON = CON.rename(columns={  'Contact ID (18-digit)':'Contact_ID', 'Opportunity Name':'OppName', 
       'Term: Term Name':'CONTerm', 'Opportunity Record Type':'Record_Type', 'Inquiry':'Inquiry', 'Inquiry Date':'Inquiry_Date',
       'Opportunity ID (18-digit)':'Opp_ID', 'Empl ID':'EMPL', 'Application Number':'App_Number'})


# In[15]:


#Sort INT by Contact_ID and Interaction Names
INT = INT.sort_values(by = ['Contact_ID','Int_Name'])


# In[16]:


#Sort CON by Contact_ID and Opportunity Name
CON = CON.sort_values(by = ['Contact_ID','Inquiry_Date'])


# In[17]:


#Dropping the duplicates to then merge later
CON=CON.drop_duplicates(['Contact_ID'], keep='last')


# In[18]:


#Inner Join
INNER_INT = pd.merge(left = INT, right = CON, left_on= 'Contact_ID', right_on = 'Contact_ID')


# In[19]:


len(INNER_INT)


# INT has been MERGED with CON on Contact_ID this will allow us to find one TERM for each contact ID

# In[20]:


INNER_INT.head(1)


# INNER_INT now has a nested tuple multiple index as a result of the merge. Have to create a unique key and replace the index

# In[138]:


INNER_INT.Contact_ID.duplicated().sum()


# In[139]:


r= len(INNER_INT)
r


# In[140]:


N = list(range(0,r))
INNER_INT['NEW'] = N
INNER_INT['NEW'] = INNER_INT['NEW'].astype(str)
INNER_INT['NEW'].dtype


# In[24]:


INNER_INT['NEW_INDEX'] = INNER_INT['Contact_ID']+INNER_INT['NEW']
INNER_INT.NEW_INDEX.duplicated().sum()


# In[25]:


INNER_INT.set_index(['NEW_INDEX'], inplace=True)
INNER_INT.shape


# Create an IPEDS Ethnicity column by combining ethnic group and ethnicity and using a dictionary to assign IPEDS values

# In[26]:


INNER_INT['ETH'] = np.where(pd.isnull(INNER_INT['Ethnic_Group']), INNER_INT['Ethnicity'], INNER_INT['Ethnic_Group'])


# In[27]:


IPEDSETH_DICT = {'Caucasian':'White', 'WHITE':'White','Not specified':'Unknown','Hispanic':'Hispanic/Latino'
                 ,'Black/African American':'Black/African American'
                 ,'Asian/Asian American/Pacific Islander':'Asian','ASIAN':'Asian','BLACK':'Black/African American'
                 ,'NSPEC':'Unknown','Multi-Ethnic':'Two or More Races'
                 ,'Other':'Unknown','SOAMER':'Hispanic/Latino','PUERTOR':'Hispanic/Latino'
                 ,'CENTAMER':'Hispanic/Latino','MEXAMER':'Hispanic/Latino'
                 ,'CUBAN':'Hispanic/Latino', 'HISPA':'Hispanic/Latino'
                 ,'American Indian/Alaskan Native':'American Indian/Alaska Native','NHISP':'Unknown'
                 ,'PACIF':'Native Hawaiian/Other Pacific Islander','AMIND':'American Indian/Alaska Native'
                 ,'SPANISH':'Hispanic/Latino'}


# In[28]:


INNER_INT['IPEDS'] = INNER_INT['ETH'].map(IPEDSETH_DICT)


# In[29]:


INNER_INT['IPEDS'].value_counts(dropna=False)


# In[30]:


INNER_INT['IPEDS2'] = np.where(pd.isnull(INNER_INT['IPEDS']), "Unknown", INNER_INT['IPEDS'])


# In[31]:


INNER_INT['IPEDS2'].value_counts(dropna=False)


# Create the columns and values for States, Clean the state columns, create state dictionary and combine the State columns

# In[32]:


import re


# In[33]:


#clean the market segment column, take the 1st 2 letters of the string only
INNER_INT['STATE'] = INNER_INT['Market_Segment'].str.extract('([A-Z]{2})',expand=True)


# In[34]:


#check the strings are cleaned
INNER_INT['STATE'].value_counts(dropna=False)


# In[35]:


#if the EMPL ID is null pass the STATE value to the new COL otherwise give me the value from "Other_State"
INNER_INT['STATE2'] = np.where(pd.isnull(INNER_INT['EMPLID']), INNER_INT['STATE'], INNER_INT['Other_State'])


# In[36]:


#create a reference list to compare the possible state values
import csv
with open('C:/Users/mjc341/Desktop/UMAN 1507 Monthly INQ summary Report/states.csv','r') as f:
    reader= csv.reader(f)
    states = list(reader)
flat_states =[item for sublist in states for item in sublist]
flat_states


# In[37]:


INNER_INT['STATE2&MAILING'] = np.where(pd.isnull(INNER_INT['Mailing_State']), INNER_INT['STATE2'], INNER_INT['Mailing_State'])


# In[38]:


#STATE2 is the cleaned MKTSEGEMNT + OTHER_STATE where EMPLID is null
#STATE3 DEFINES STATE2 as existing in the flat_state file or being NOT in USA
INNER_INT['STATE3'] = np.where(INNER_INT['STATE2'].isin(flat_states),INNER_INT['STATE2'], "Not USA")


# In[39]:


INNER_INT['STATE4'] = np.where(INNER_INT['STATE2&MAILING'].isin(flat_states),INNER_INT['STATE2&MAILING'], "Not USA")


# In[40]:


INNER_INT['STATE4'].value_counts(dropna=False)


# In[41]:


state_dict= {'ALABAMA':'AL','Alabama':'AL','AL':'AL','ALASKA':'AK','Alaska':'AK','AK':'AK','ARIZONA':'AZ','Arizona':'AZ','AZ':'AZ','ARKANSAS':'AR','Arkansas':'AR','AR':'AR','CALIFORNIA':'CA'
,'California':'CA','CA':'CA','COLORADO':'CO','Colorado': 'CO','CO':'CO','CONNECTICUT':'CT'
,'Connecticut':'CT','CT':'CT','DELAWARE':'DE','Delaware':'DE','DE':'DE','FLORIDA':'FL','Florida':'FL'
,'FL':'FL','GEORGIA':'GA','Georgia':'GA','GA':'GA','HAWAII':'HI','Hawaii':'HI','HI':'HI','IDAHO':'ID'
, 'Idaho':'ID','ID':'ID','ILLINOIS':'IL','Illinois':'IL','IL':'IL','INDIANA':'IN','Indiana':'IN','IN':'IN'
,'IOWA':'IA','Iowa':'IA','IA':'IA','KANSAS':'KS','Kansas':'KS','KS':'KS','KENTUCKY':'KY','Kentucky':'KY'
,'KY':'KY','LOUISIANA':'LA','Louisiana':'LA','LA':'LA','MAINE':'ME','Maine':'ME','ME':'ME'
,'MARYLAND':'MD','Maryland':'MD','MD':'MD','MASSACHUSETTS':'MA','Massachusetts':'MA','MA':'MA'
,'MICHIGAN':'MI','Michigan':'MI','MI':'MI','MINNESOTA':'MN','Minnesota':'MN','MN':'MN','MISSISSIPPI':'MS'
,'Mississippi':'MS','MS':'MS','MISSOURI':'MO','Missouri':'MO','MO':'MO','MONTANA':'MT','Montana':'MT'
,'MT':'MT','NEBRASKA':'NE','Nebraska':'NE','NE':'NE','NEVADA':'NV','Nevada':'NV','NV':'NV'
,'NEW HAMPSHIRE':'NH','New Hampshire':'NH','NH':'NH','NEW JERSEY':'NJ','New Jersey':'NJ','NJ':'NJ'
,'NEW MEXICO':'NM','New Mexico':'NM','NM':'NM','NEW YORK':'NY','New York':'NY', 'NY':'NY'
,'NORTH CAROLINA':'NC','North Carolina':'NC','NC':'NC','NORTH DAKOTA':'ND','North Dakota':'ND','ND':'ND'
,'OHIO':'OH','Ohio':'OH','OH':'OH','OKLAHOMA':'OK','Oklahoma':'OK','OK':'OK','OREGON':'OR','Oregon':'OR'
,'OR':'OR','PENNSYLVANIA':'PA','Pennsylvania':'PA','PA':'PA','RHODE ISLAND':'RI'
,'Rhode Island':'RI','RI':'RI','SOUTH CAROLINA':'SC','South Carolina':'SC','SC':'SC','SOUTH DAKOTA':'SD'
,'South Dakota':'SD','SD':'SD','TENNESSEE':'TN','Tennessee':'TN','TN':'TN','TEXAS':'TX','Texas':'TX'
,'TX':'TX','UTAH':'UT','Utah':'UT','UT':'UT','VERMONT':'VT','Vermont':'VT','VT':'VT','VIRGINIA':'VA'
,'Virginia':'VA','VA':'VA','WASHINGTON':'WA','Washington':'WA','WA':'WA','WEST VIRGINIA':'WV'
,'West Virginia':'WV','WV':'WV','WISCONSIN':'WI','Wisconsin':'WI','WI':'WI','WYOMING':'WY','Wyoming':'WY'
,'WY':'WY','GUAM':'GU','Guam':'GU','GU':'GU','DISTRICT OF COLUMBIA':'DC','District of Columbia':'DC'
,'DC':'DC','FEDERATED STATES OF MICRONESIA':'FM','Federated States of Micronesia':'FM','FM':'FM'
,'MARSHALL ISLANDS':'MH','Marshall Islands':'MH','MH':'MH','NORTHERN MARIANA ISLANDS':'MP'
,'Northern Mariana Islands':'MP','MP':'MP', 'PALAU':'PW','Palau':'PW','PW':'PW','PUERTO RICO':'PR'
,'Puerto Rico':'PR','PR':'PR','VIRGIN ISLANDS':'VI','Virgin Islands':'VI','VI':'VI'}


# In[42]:


INNER_INT['STATE5'] = INNER_INT['STATE4'].map(state_dict)


# In[43]:


INNER_INT['STATE5'].value_counts(dropna=False)


# In[44]:


INNER_INT['STATE5'].value_counts(dropna=False).sum()


# In[45]:


INNER_INT.columns


# Return to previous code string and create all of the date variables

# In[46]:


# Create a datetime variable from Create_Date this operation takes up to 3 min
#INT['Create_Date']= pd.to_datetime(INT.Create_Date)
INNER_INT['Create_Date']= pd.to_datetime(INNER_INT.Create_Date)


# In[47]:


#INT.info()


# In[48]:


#find earliest create date
#INT['Create_Date'].min()
INNER_INT['Create_Date'].min()


# Here we will insert the lambda function to extract the YEAR and MONTH from the new datetime type so we can filter on YEAR and groupby MONTH also add M col that can be dictionaried to Names of MONTHS instead of numbers. Question can we format the month names at the conversion?

# In[49]:


#INT.loc[:,'YEAR'] =INT['Create_Date'].apply(lambda x: "%d" % (x.year))
INNER_INT.loc[:,'YEAR'] =INNER_INT['Create_Date'].apply(lambda x: "%d" % (x.year))


# In[50]:


#INT.loc[:,'MONTH'] = INT['Create_Date'].apply(lambda x: "%d" % (x.month))
INNER_INT.loc[:,'MONTH'] = INNER_INT['Create_Date'].apply(lambda x: "%d" % (x.month))


# In[51]:


INNER_INT.head(1)


# In[52]:


# filter on term for 2017,2018,2019,2020
#TERM2017 = INT.loc[(INT.Term == 'Fall 2017')]
#TERM2018 = INT.loc[(INT.Term == 'Fall 2018')]
#TERM2019 = INT.loc[(INT.Term == 'Fall 2019')]
#TERM2020 = INT.loc[(INT.Term == 'Fall 2020')]
TERM2017 = INNER_INT.loc[(INNER_INT.CONTerm == 'Fall 2017')]
TERM2018 = INNER_INT.loc[(INNER_INT.CONTerm == 'Fall 2018')]
TERM2019 = INNER_INT.loc[(INNER_INT.CONTerm == 'Fall 2019')]
TERM2020 = INNER_INT.loc[(INNER_INT.CONTerm == 'Fall 2020')]


# In[53]:


len(TERM2017)


# In[54]:


len(TERM2018)


# In[55]:


len(TERM2019)


# In[56]:


len(TERM2020)


# In[57]:


#filter on YEAR for  2016,2017,2018,2019
#INQ2017 df is inquiries made in 2016 for 2017 TERM
INQ2017 = TERM2017[TERM2017['YEAR'] == '2016']
INQ2018 = TERM2018[TERM2018['YEAR'] == '2017']
INQ2019 = TERM2019[TERM2019['YEAR'] == '2018']
INQ2020 = TERM2020[TERM2020['YEAR'] == '2018']
INQ2017year2 = TERM2017[TERM2017['YEAR'] == '2017']
INQ2018year2 = TERM2018[TERM2018['YEAR'] == '2017']
INQ2019year2 = TERM2019[TERM2019['YEAR'] == '2018']


# In[58]:


len(INQ2017)


# In[59]:


len(INQ2018)


# In[60]:


len(INQ2019)


# In[61]:


len(INQ2020)


# In[62]:


#show head of INQ2017 same for 2017,2018,2019
INQ2017.head(1)


# In[63]:


INQ2017.columns


# In[64]:


#confirm that filtered data is a datafram
type(INQ2017)


# In[65]:


#pd.show_versions()


# In[66]:


#get count by month of each term
INQ2017.Create_Date.dt.month.value_counts().sort_index()


# In[67]:


INQ2018.Create_Date.dt.month.value_counts().sort_index()


# In[68]:


INQ2019.Create_Date.dt.month.value_counts().sort_index()


# In[69]:


#this creates a list for 2018 dates but will eventually have to be changed to 2019 dates
INQ2020.Create_Date.dt.month.value_counts().sort_index()


# In[70]:


#create a timestamp to get values from the FAll 2018 term that are prior to Jan 1 2017
#also created a timestamp for february and august values of the subsequent year of each opp term
#a current timestamp was also created as well as the Y-o-Y time stamps (12-15).
tsprior = pd.Timestamp('1/1/2016')
tsprior2 = pd.Timestamp('1/1/2017')
tsprior3 = pd.Timestamp('1/1/2018')
tsprior4 = pd.Timestamp('1/1/2019')
tsprior5 = pd.Timestamp('2/1/2017')
tsprior6 = pd.Timestamp('2/1/2018')
tsprior7 = pd.Timestamp('2/1/2019')
tsprior8 = pd.Timestamp('8/1/2017')
tsprior9 = pd.Timestamp('8/1/2018')
tsprior10 = pd.Timestamp('9/1/2019')
tsprior11 = pd.Timestamp('9/1/2018')
tsprior12 = pd.Timestamp('9/1/2017')
tsprior13 = pd.Timestamp('9/1/2016')


# In[71]:


tsprior


# In[72]:


B4JAN16 = TERM2017.loc[TERM2017.Create_Date < tsprior,:]
B4JAN17 = TERM2018.loc[TERM2018.Create_Date < tsprior2,:]
B4JAN18 = TERM2019.loc[TERM2019.Create_Date < tsprior3,:]
B4JAN19 = TERM2020.loc[TERM2020.Create_Date < tsprior3,:]
B4FEB17 = TERM2017.loc[TERM2017.Create_Date < tsprior5,:]
B4FEB18 = TERM2018.loc[TERM2018.Create_Date < tsprior6,:]
B4FEB19 = TERM2019.loc[TERM2019.Create_Date < tsprior7,:]
B4FEB20 = TERM2020.loc[TERM2020.Create_Date < tsprior7,:]
B4AUG17 = TERM2017.loc[TERM2017.Create_Date < tsprior8,:]
B4AUG18 = TERM2018.loc[TERM2018.Create_Date < tsprior9,:]
TOTAL2017 = TERM2017.loc[TERM2017.Create_Date < tsprior10,:]
TOTAL2018 = TERM2018.loc[TERM2018.Create_Date < tsprior10,:]
TOTAL2019 = TERM2019.loc[TERM2019.Create_Date < tsprior10,:]
TOTAL2020 = TERM2020.loc[TERM2020.Create_Date < tsprior10,:]
YoY2017 = TERM2017.loc[TERM2017.Create_Date < tsprior13,:]
YoY2018 = TERM2018.loc[TERM2018.Create_Date < tsprior12,:]
YoY2019 = TERM2019.loc[TERM2019.Create_Date < tsprior11,:]


# In[73]:


len(B4JAN16)


# In[74]:


len(B4JAN17)


# In[75]:


len(B4JAN18)


# In[76]:


len(B4JAN19)


# In[77]:


#test the counts
B4JAN16['Contact_ID'].value_counts().sum()
#B4JAN17['Contact_ID'].value_counts().sum()
#B4JAN18['Contact_ID'].value_counts().sum()
#B4JAN19['Contact_ID'].value_counts().sum()
#B4FEB17['Contact_ID'].value_counts().sum()
#B4FEB18['Contact_ID'].value_counts().sum()
#B4FEB19['Contact_ID'].value_counts().sum()
#B4FEB20['Contact_ID'].value_counts().sum()
#B4AUG17['Contact_ID'].value_counts().sum()
#B4AUG18['Contact_ID'].value_counts().sum()
#TOTAL2017['Contact_ID'].value_counts().sum()
#TOTAL2018['Contact_ID'].value_counts().sum()
#TOTAL2019['Contact_ID'].value_counts().sum()
#TOTAL2020['Contact_ID'].value_counts().sum()


# In[78]:


#B4JAN19['Contact_ID'].value_counts().sum()


# In[79]:


R17 = INQ2017.Create_Date.dt.month.value_counts().sort_index()
R18 = INQ2018.Create_Date.dt.month.value_counts().sort_index()
R19 = INQ2019.Create_Date.dt.month.value_counts().sort_index()
R20 = INQ2020.Create_Date.dt.month.value_counts().sort_index()


# In[80]:


R17df = pd.DataFrame([R17]).T
R18df = pd.DataFrame([R18]).T
R19df = pd.DataFrame([R19]).T
R20df = pd.DataFrame([R20]).T


# In[81]:


#R16df.head()


# In[82]:


a =R17df.rename({'Create_Date': 'Count-2016'}, axis = 'columns')
#R16df.rename({'Create_Date': 'Count-2016'}, axis = 'columns')


# In[83]:


b = R18df.rename({'Create_Date': 'Count-2017'}, axis = 'columns')
#R17df.rename({'Create_Date': 'Count-2017'}, axis = 'columns')


# In[84]:


#R18df.rename({'Create_Date': 'Count-2018'}, axis = 'columns')
c = R19df.rename({'Create_Date': 'Count-2018'}, axis = 'columns')


# In[85]:


d = R20df.rename({'Create_Date': 'Count-2018'}, axis = 'columns')


# In[86]:


j = pd.concat([a, b], axis=1, join_axes=[a.index])
    


# In[87]:


#j.head(1)


# In[88]:


k = pd.concat([j, c], axis=1, join_axes=[a.index])


# In[89]:


k2 = pd.concat([k, d], axis=1, join_axes=[a.index])


# In[90]:


k2.head()


# In[91]:


#k.head()


# In[92]:


type(k2)


# In[93]:


k2.dtypes


# In[94]:


k =  k2.replace(np.nan,0)


# In[95]:


k['Count-2018'] = k['Count-2018'].astype(int)
k


# In[96]:


#k.dtypes


# In[97]:


k.index


# In[98]:


# create  a dictionary to map the month values
dict = {1:'JAN',2:'FEB',3:'MAR',4:'APR',5:'MAY',6:'JUN',7:'JUL',8:'AUG',9:'SEP',10:'OCT',11:'NOV',12:'DEC'}


# In[99]:


#map the dictionary to the new col on index
k ['MONTH'] = k.index.map(dict)


# In[100]:


k


# In[101]:


#change the name of the df and rearrange the cols
INQbyMONTH = k[['MONTH','Count-2016','Count-2017','Count-2018']]


# In[102]:


#this is the result
INQbyMONTH


# In[103]:


#it is a df
type(INQbyMONTH)


# In[104]:


#before January values
a = B4JAN16['Contact_ID'].value_counts().sum()
b = B4JAN17['Contact_ID'].value_counts().sum()
c = B4JAN18['Contact_ID'].value_counts().sum()
d = B4JAN19['Contact_ID'].value_counts().sum()


# In[105]:


PRIOR = pd.DataFrame({'Prior to Jan 2016': a , 'Prior to Jan 2017': b , 'Prior to Jan 2018': c, 'Prior to Jan 2018.2': d},index = [0])


# In[106]:


PRIOR


# In[107]:


#before February values
e = B4FEB17['Contact_ID'].value_counts().sum()
f = B4FEB18['Contact_ID'].value_counts().sum()


# In[108]:


PRIOR2 = pd.DataFrame({'Prior to Feb 2017': e , 'Prior to Feb 2018': f },index = [0])


# In[109]:


PRIOR2


# In[110]:


#before August values
g = B4AUG17['Contact_ID'].value_counts().sum()
h = B4AUG18['Contact_ID'].value_counts().sum()


# In[111]:


PRIOR3 = pd.DataFrame({'Prior to Aug 2017': g , 'Prior to Aug 2018': h },index = [0])


# In[112]:


PRIOR3


# In[113]:


#current values
i = TOTAL2017['Contact_ID'].value_counts().sum()
j = TOTAL2018['Contact_ID'].value_counts().sum()
k = TOTAL2019['Contact_ID'].value_counts().sum()
l = TOTAL2020['Contact_ID'].value_counts().sum()


# In[114]:


TOTAL4 = pd.DataFrame({'Total 2017': i , 'Total 2018': j, 'Total 2019': k, 'Total 2020': l},index = [0])


# In[115]:


TOTAL4


# In[116]:


#create a df for the y-o-y comparison of 2017, 2018, and 2019. #redundant???
#tspriorFEB17 = pd.Timestamp('2/1/2017')
#tspriorFEB18 = pd.Timestamp ('2/1/2018')
#tspriorFEB19 = pd.Timestamp('2/1/2019')
#tspriorAUG17 = pd.Timestamp('8/1/2017')
#tspriorAUG18 = pd.Timestamp ('8/1/2018')
#tspriorAUG19 = pd.Timestamp('8/1/2019')
#tsJUN312017 = pd.Timestamp('6/30/2017')
#tsJUN312018 = pd.Timestamp('6/30/2018')
#tsJUN312019 = pd.Timestamp('6/30/2019')


# In[117]:


#LeadSource Counts for AdmitTerm 2017 that occured in 2016
LEAD17 = YoY2017.pivot_table(index=['Lead_Source'],values =['Contact_ID'],aggfunc =len)
LEAD17.rename(index= str, columns={'Contact_ID':'2017'},inplace =True)
LEAD17


# In[118]:


#LeadSource Counts for AdmitTerm 2018 that occured in 2017
LEAD18 = YoY2018.pivot_table(index=['Lead_Source'],values =['Contact_ID'],aggfunc =len)
LEAD18.rename(index= str, columns={'Contact_ID':'2018'},inplace =True)
LEAD18


# In[119]:


#LeadSource Counts for AdmitTerm 2019 that occured in 2018
LEAD19 = YoY2019.pivot_table(index=['Lead_Source'],values =['Contact_ID'],aggfunc =len)
LEAD19.rename(index= str, columns={'Contact_ID':'2019'},inplace =True)
LEAD19


# In[120]:


LEAD1 = pd.concat([LEAD17, LEAD18], axis=1, join_axes=[LEAD18.index])
LEAD1


# In[121]:


Lead2 = pd.concat([LEAD1, LEAD19], axis=1, join_axes=[LEAD18.index])
Lead2


# In[122]:


LEAD = Lead2.fillna(0)
LEAD['2018'] = LEAD['2018'].astype(int)
LEAD['2019'] = LEAD['2019'].astype(int)
LEAD


# In[123]:


#Count all Inq for the 2017 Admit term that occured before August of 2017  #Redundant???
#B4AUG17 = TERM2017.loc[TERM2017.Create_Date <= tspriorAUG17,:]
#len(B4AUG17)


# In[124]:


#Count all Inq for the 2018 Admit term that occured before August of 2018 #Redundant???
#B4AUG18 = TERM2018.loc[TERM2018.Create_Date <= tspriorAUG18,:]
#len(B4AUG18)


# In[125]:


# create the path where the results willbe sent
path_final = r'C:/Users/mjc341/Desktop/EMSA-GIT/COLAB'
#path_final = r'C:/Users/sxh706/Desktop'


# In[126]:


#path_final


# In[127]:


#Ethnicity Counts for AdmitTerm 2017 that occured in 2016
ETH17 = YoY2017.pivot_table(index=['IPEDS2'],values =['Contact_ID'],aggfunc =len)
ETH17.rename(index= str, columns={'Contact_ID':'2017'},inplace =True)
ETH17


# In[128]:


#ETH Counts for AdmitTerm 2018 that occured in 2017
ETH18 = YoY2018.pivot_table(index=['IPEDS2'],values =['Contact_ID'],aggfunc =len)
ETH18.rename(index= str, columns={'Contact_ID':'2018'},inplace =True)
ETH18


# In[129]:


#Ethnicity Counts for AdmitTerm 2019 that occured in 2018
ETH19 = YoY2019.pivot_table(index=['IPEDS2'],values =['Contact_ID'],aggfunc =len)
ETH19.rename(index= str, columns={'Contact_ID':'2019'},inplace =True)
ETH19


# In[130]:


ETH1 = pd.concat([ETH17, ETH18], axis=1, join_axes=[ETH18.index])
ETH2 = pd.concat([ETH1, ETH19], axis=1, join_axes=[ETH18.index])
ETH2


# In[131]:


ETH2.info()


# In[132]:


#State Counts for AdmitTerm 2017 that occured in 2016
STATE17 = YoY2017.pivot_table(index=['STATE5'],values =['Contact_ID'],aggfunc =len)
STATE17.rename(index= str, columns={'Contact_ID':'2017'},inplace =True)
STATE18 = YoY2018.pivot_table(index=['STATE5'],values =['Contact_ID'],aggfunc =len)
STATE18.rename(index= str, columns={'Contact_ID':'2018'},inplace =True)
STATE19 = YoY2019.pivot_table(index=['STATE5'],values =['Contact_ID'],aggfunc =len)
STATE19.rename(index= str, columns={'Contact_ID':'2019'},inplace =True)


# In[133]:


ST1 = pd.concat([STATE17, STATE18], axis=1, join_axes=[STATE18.index])
ST2 = pd.concat([ST1, STATE19], axis=1, join_axes=[STATE18.index])
ST2


# In[134]:


writer = pd.ExcelWriter(path_final + '/' + 'SalesForceByMonthSept' + '.xlsx')
INQbyMONTH.to_excel(writer,'Results')
PRIOR.to_excel(writer, 'PRIOR')
PRIOR2.to_excel(writer, 'PRIOR2')
PRIOR3.to_excel(writer, 'PRIOR3')
TOTAL4.to_excel(writer, 'TOTAL4')
LEAD.to_excel(writer,'LEAD')
ETH2.to_excel(writer,'ETH2')
ST2.to_excel(writer,'ST2')

workbook = writer.book
worksheet = writer.sheets['Results'] 
worksheet2 = writer.sheets ['PRIOR']
worksheet3 = writer.sheets ['PRIOR2']
worksheet4 = writer.sheets ['PRIOR3']
worksheet5 = writer.sheets ['TOTAL4']
worksheet6 = writer.sheets ['LEAD']
worksheet7 = writer.sheets ['ETH2']
worksheet8 = writer.sheets['ST2']

format1 = workbook.add_format({'num_format': '#,##', 'bold': True, 'border': '1'})
format2 = workbook.add_format({'num_format': '##,###', 'border': '1','bg_color':'#FFC7CE'})
format3 = workbook.add_format({'border': '1','bg_color':'#DAF7A6'})
format4 = workbook.add_format({'border': '1','bg_color':'#DAF7A6'})
#added bkground color so we can show different years

# Note: It isn't possible to format any cells that already have a format such
# as the index or headers or any cells that contain dates or datetimes.

worksheet.set_column('A:A', 10)
worksheet.set_column('B:E', 12)
worksheet.set_row(1, None, format2)
worksheet.set_row(2, None, format2)
worksheet.set_row(3, None, format2)
worksheet.set_row(4, None, format2)
worksheet.set_row(5, None, format2)
worksheet.set_row(6, None, format2)
worksheet.set_row(7, None, format2)
worksheet.set_row(8, None, format2)
worksheet.set_row(9, None, format2)
worksheet.set_row(10, None, format2)
worksheet.set_row(11, None, format2)
worksheet.set_row(12, None, format2)

worksheet2.set_column('A:A', 10)
worksheet2.set_column('B:E', 18)
worksheet2.set_row(1, None, format3)
worksheet2.set_row(2, None, format4)

worksheet3.set_column('A:A', 10)
worksheet3.set_column('B:C', 18)
worksheet3.set_row(1, None, format3)
worksheet3.set_row(2, None, format4)

worksheet4.set_column('A:A', 10)
worksheet4.set_column('B:C', 18)
worksheet4.set_row(1, None, format3)
worksheet4.set_row(2, None, format4)

worksheet5.set_column('A:A', 10)
worksheet5.set_column('B:E', 18)
worksheet5.set_row(1, None, format3)
worksheet5.set_row(2, None, format4)

worksheet6.set_column('A:A', 10)
worksheet6.set_column('B:E', 18)
worksheet6.set_row(1, None, format3)
worksheet6.set_row(2, None, format4)
worksheet6.set_row(3, None, format4)
worksheet6.set_row(4, None, format4)
worksheet6.set_row(5, None, format4)
worksheet6.set_row(6, None, format4)
worksheet6.set_row(7, None, format4)
worksheet6.set_row(8, None, format4)
worksheet6.set_row(9, None, format4)
worksheet6.set_row(10, None, format4)
worksheet6.set_row(11, None, format4)
worksheet6.set_row(12, None, format4)
worksheet6.set_row(13, None, format4)
worksheet6.set_row(14, None, format4)
worksheet6.set_row(15, None, format4)
worksheet6.set_row(16, None, format4)
worksheet6.set_row(17, None, format4)
worksheet6.set_row(18, None, format4)
worksheet6.set_row(19, None, format4)
worksheet6.set_row(20, None, format4)

worksheet7.set_column('A:A', 10)
worksheet7.set_column('B:E', 18)
worksheet7.set_row(1, None, format3)
worksheet7.set_row(2, None, format4)
worksheet7.set_row(3, None, format4)
worksheet7.set_row(4, None, format4)
worksheet7.set_row(5, None, format4)
worksheet7.set_row(6, None, format4)
worksheet7.set_row(7, None, format4)
worksheet7.set_row(8, None, format4)
worksheet7.set_row(9, None, format4)

worksheet8.set_column('A:A', 10)
worksheet8.set_column('B:E', 18)
worksheet8.set_row(1, None, format3)
worksheet8.set_row(2, None, format4)
worksheet8.set_row(3, None, format4)
worksheet8.set_row(4, None, format4)
worksheet8.set_row(5, None, format4)
worksheet8.set_row(6, None, format4)
worksheet8.set_row(7, None, format4)
worksheet8.set_row(8, None, format4)
worksheet8.set_row(9, None, format4)
worksheet8.set_row(10, None, format4)
worksheet8.set_row(11, None, format4)
worksheet8.set_row(12, None, format4)
worksheet8.set_row(13, None, format4)
worksheet8.set_row(14, None, format4)
worksheet8.set_row(15, None, format4)
worksheet8.set_row(16, None, format4)
worksheet8.set_row(17, None, format4)
worksheet8.set_row(18, None, format4)
worksheet8.set_row(19, None, format4)
worksheet8.set_row(20, None, format4)
worksheet8.set_row(21, None, format4)
worksheet8.set_row(22, None, format4)
worksheet8.set_row(23, None, format4)
worksheet8.set_row(24, None, format4)
worksheet8.set_row(25, None, format4)
worksheet8.set_row(26, None, format4)
worksheet8.set_row(27, None, format4)
worksheet8.set_row(28, None, format4)
worksheet8.set_row(29, None, format4)
worksheet8.set_row(30, None, format4)
worksheet8.set_row(31, None, format4)
worksheet8.set_row(32, None, format4)
worksheet8.set_row(33, None, format4)
worksheet8.set_row(34, None, format4)
worksheet8.set_row(35, None, format4)
worksheet8.set_row(36, None, format4)
worksheet8.set_row(37, None, format4)
worksheet8.set_row(38, None, format4)
worksheet8.set_row(39, None, format4)
worksheet8.set_row(40, None, format4)
worksheet8.set_row(41, None, format4)
worksheet8.set_row(42, None, format4)
worksheet8.set_row(43, None, format4)
worksheet8.set_row(44, None, format4)
worksheet8.set_row(45, None, format4)
worksheet8.set_row(46, None, format4)
worksheet8.set_row(47, None, format4)
worksheet8.set_row(48, None, format4)
worksheet8.set_row(49, None, format4)
worksheet8.set_row(50, None, format4)
worksheet8.set_row(51, None, format4)
worksheet8.set_row(52, None, format4)
worksheet8.set_row(53, None, format4)
worksheet8.set_row(54, None, format4)
worksheet8.set_row(55, None, format4)
worksheet8.set_row(56, None, format4)
worksheet8.set_row(57, None, format4)

#dont write until you check the code then remove the # below
writer.save()

