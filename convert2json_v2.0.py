from genericpath import exists
from operator import index
from re import T
from turtle import left
import pandas as pd # pip3 install pandas
import numpy as np # pip3 install numpy
import ipaddress
import json
import openpyxl # pip3 install openpyxl
import os, datetime, time, sys
import glob # pip3 install glob
#import pysftp as sftp # pip3 install pysftp
import warnings
import yaml # pip3 install pyyaml

warnings.filterwarnings("ignore", category=UserWarning)

def progressbar(it, prefix="", size=60, out=sys.stdout): # Python3.3+
    count = len(it)
    def show(j):
        x = int(size*j/count)
        print("{}[{}{}] {}/{}".format(prefix, u"■"*x, "."*(size-x), j, count), 
                end='\r', file=out, flush=True)
    show(0)
    for i, item in enumerate(it):
        yield item
        show(i+1)
    print("\n", flush=True, file=out)

'''
## version 1.4 Update ##
Need to delete row 21-29 - DONE
Need to rename header name RAN_Vendor to Vendor - DONE on WFM
Need to inject and modify all value on CSG_IP_System, CSG_LAG_Desc, CSG_Port1, CSG_Port2, CSG_Port3, CSG_Lag_ID to dummy null - DONE
New_Port_Type, New_Port_Setup_CSG to be added on grup_C - DONE
Convert value VLAN-ADD to VA and VLAN-ADD+NEW-PORT to VA_NP - DONE
'''
userName = 'lukman' # <== plz update your name accordingly
datestring = datetime.datetime.now().strftime("%d-%m-%Y")
# cur_Path = os.getcwd() # using os.getcwd() sometime didn't give exact path of subdirectory depending upon OS. Use below alternative for exact path
cur_Path = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
excel_File = "%s\RAN_SOR_TEMPLATE_v2.0.xlsx" % cur_Path
result_Path = os.path.join(cur_Path, datestring) 
isExist = os.path.exists(result_Path)

if not isExist:
    os.mkdir(result_Path)

# print (f"The name is {excel_File}")

grup_A = pd.read_excel('%s' % excel_File, usecols=["FON", "Network", "Subnet", "VLAN", "Traffic_Type", "RAN_Vendor", "Region", "MW_Network", "MW_Subnet","FTTM_RAN_Port_1", "FTTM_RAN_Port_2", "FTTM_MW_Port","FTTM_MW_Port_2","FTTM_VLAN_REF"])
grup_A = grup_A.dropna(how = 'all') #or use this# grup_A = grup_A[grup_A["FON"].notna()]
grup_A.rename(columns = {'FTTM_MW_Port':'FTTM_MW_RAN_Port'}, inplace = True)
grup_A.rename(columns = {'FTTM_MW_Port_2':'FTTM_MW_RAN_Port_2'}, inplace = True)
grup_A.drop(grup_A.loc[19:26].index, inplace=True)
grup_A["IpGw"] = [str(ipaddress.ip_address(i) + 1) for i in grup_A["Network"]]
grup_A["IpNextHopMW"] = [str(ipaddress.ip_address(i) + 4) for i in grup_A["Network"]]
# print (grup_A)
if grup_A["FTTM_VLAN_REF"].isnull().values.any() == True :
    grup_A[["FON", "VLAN"]] = grup_A[["FON", "VLAN"]].astype(np.int64)
else:
    grup_A[["FON", "VLAN","FTTM_VLAN_REF"]] = grup_A[["FON", "VLAN","FTTM_VLAN_REF"]].astype(np.int64)

# print(grup_A)
grup_A["FTTM_MW_RAN_Port"].fillna(grup_A["FTTM_MW_RAN_Port"].values[0], inplace=True)


#Grup B will be use for FTTM scenario later
grup_B = pd.read_excel('%s' % excel_File, usecols=["FTTM_RAN_Port_1", "FTTM_RAN_Port_2", "FTTM_MW_Port","FTTM_MW_Port_2"])
grup_B = grup_B.dropna(how = 'all')
grup_B.rename(columns = {'FTTM_MW_Port':'FTTM_MW_RAN_Port'}, inplace = True)
grup_B.rename(columns = {'FTTM_MW_Port_2':'FTTM_MW_RAN_Port_2'}, inplace = True)

# grup_B["FTTM_MW_RAN_Port"].fillna("ASEM", inplace=True)
# grup_B["FTTM_MW_RAN_Port"] = grup_B["FTTM_MW_RAN_Port"].str.replace("ASEM", grup_B["FTTM_MW_RAN_Port"].values[0])
a = grup_B["FTTM_RAN_Port_1"]
b = grup_B["FTTM_RAN_Port_2"]
c = grup_B["FTTM_MW_RAN_Port"]
d = grup_B["FTTM_MW_RAN_Port_2"]
# grup_D = a.append(b).append(c).reset_index(drop=True).sort_values().drop_duplicates().reset_index(drop=True).dropna(how = 'all') 



grup_C = pd.read_excel('%s' % excel_File, usecols=["Req_Type", "SOR_CRQ", "New_Port_Type_CSG", "New_Port_Setup_CSG"])
grup_C.rename(columns = {'New_Port_Type_CSG':'New_Port_Type'}, inplace = True)
grup_C[["CSG_Port1","CSG_Port2","CSG_Port3","CSG_Lag_ID","CSG_LAG_Desc","CSG_IP_System"]] = np.NaN
grup_C = grup_C.dropna(how = 'all')
grup_C["Req_Type"] = grup_C["Req_Type"].str.replace('VLAN-ADD', 'VA').str.replace('NEW-PORT', 'NP').str.replace('+', '_', regex=False)
# print (grup_C)

'''
FutureWarning: The series.append method is deprecated and will be removed from pandas in a future version. Use pandas.concat instead.
'''

if grup_C["Req_Type"].values[0] == "VA" :
    grup_D = grup_A["FTTM_VLAN_REF"].drop_duplicates() 
else:
    grup_D = pd.concat((a, b, c, d)).reset_index(drop=True).sort_values().drop_duplicates().reset_index(drop=True).dropna(how = 'all')


grup_A.to_json('input1.json', orient='records')
#grup_B.to_json('input2.json', orient='records')
grup_C.to_json('input3.json', orient='records')
grup_D.to_json('input4.json', orient='records')

fon = grup_A.loc[0,'FON']
sor = grup_C.loc[0,'SOR_CRQ']

baba = '''
{
    "GroupA": [],
    "GroupB": [],
    "GroupC": [],
    "GroupD": []
}
'''

file_json = '%s/%s_%s_%s.json' % (result_Path,userName,fon,sor)
file_yaml = '%s/%s_%s_%s.yaml' % (result_Path,userName,fon,sor)

# Create initial file and reset content of file_json
with open(file_json, 'w') as f:
    f.write(baba)

# Function to add to JSON
def write_json(new_data, filename=file_json):
    with open(filename, 'r+') as file:
        file_data = json.load(file)
        file_data["GroupA"].append(new_data)
        file.seek(0)
        json.dump(file_data, file, indent=4)

# def write_json2(new_data, filename=file_json):
#     with open(filename, 'r+') as file:
#         file_data = json.load(file)
#         file_data["GroupB"].append(new_data)
#         file.seek(0)
#         json.dump(file_data, file, indent=4)

def write_json3(new_data, filename=file_json):
    with open(filename, 'r+') as file:
        file_data = json.load(file)
        file_data["GroupC"].append(new_data)
        file.seek(0)
        json.dump(file_data, file, indent=4)

def write_json4(new_data, filename=file_json):
    with open(filename, 'r+') as file:
        file_data = json.load(file)
        file_data["GroupD"].append(new_data)
        file.seek(0)
        json.dump(file_data, file, indent=4)
        
with open('input1.json', '+r') as f:
    data = json.load(f)
    for i in data:
        write_json(i)

# with open('input2.json', '+r') as f:
#     data = json.load(f)
#     for i in data:
#         write_json2(i)
with open('input3.json', '+r') as f:
    data = json.load(f)
    for i in data:
        write_json3(i)

with open('input4.json', '+r') as f:
    data = json.load(f)
    for i in data:
        write_json4(i)
        
with open(file_json, 'r+') as f:
    lines = f.readlines()
    lines.insert(0, '{"result":\n')
    f.seek(0)
    f.writelines(lines)

with open(file_json, 'a') as f:
    f.write('\n}')

for file in glob.glob("input[1-4].json"):
    os.remove(file) 

with open(file_json, 'r') as file:
    configuration = json.load(file)

with open(file_yaml, 'w') as yaml_file:
    yaml.dump(configuration, yaml_file)

print ("")

for i in progressbar(range(100), "Computing: ", 40):
    time.sleep(0.001) 

result_Name = 'JSON and YAML File created successfully! \n "%s" \n "%s"' % (file_json,file_yaml)
print ("- " * ((len(file_json) // 2)+2))
print (result_Name)
print ("- " * ((len(file_json) // 2)+2))