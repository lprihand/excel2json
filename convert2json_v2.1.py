import ipaddress
import json
import os, datetime, time, sys
import warnings
import pandas as pd # pip3 install pandas
import numpy as np # pip3 install numpy
import openpyxl # pip3 install openpyxl
import glob # pip3 install glob
import yaml #pip3 install pyyaml
#import pysftp as sftp # pip3 install pysftp

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

userName = 'lukman' # <== plz update your name accordingly
datestring = datetime.datetime.now().strftime("%d-%m-%Y")
# cur_Path = os.getcwd() # using os.getcwd() sometime didn't give exact path of subdirectory depending upon OS. Use below alternative for exact path
cur_Path = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
excel_File = f"{cur_Path}\RAN_SOR_TEMPLATE_v2.1_455.xlsx"
result_Path = os.path.join(cur_Path, datestring) 
isExist = os.path.exists(result_Path)

if not isExist:
    os.mkdir(result_Path)

grup_A = pd.read_excel(f'{excel_File}', usecols=["FON", "Network", "Subnet", "VLAN", "Traffic_Type", "RAN_Vendor", "Region", "MW_Network", "MW_Subnet",
                                                "FTTM_RAN_Port_1", "FTTM_RAN_Port_2", "FTTM_MW_Port_1","FTTM_MW_Port_2","FTTM_VLAN_REF"])

grup_A = grup_A.dropna(subset=['FON']) # or use this #grup_A = grup_A[grup_A['FON'].notna()]
# grup_A = grup_A.dropna(how = 'all')

grup_A.rename(columns = {'FTTM_MW_Port_1':'FTTM_MW_RAN_Port_1', 'FTTM_MW_Port_2':'FTTM_MW_RAN_Port_2'}, inplace = True)
grup_A.drop(grup_A.loc[19:26].index, inplace=True)
#print (grup_A)
grup_A["IpGw"] = [str(ipaddress.ip_address(i) + 1) for i in grup_A["Network"]]
grup_A["IpNextHopMW"] = [str(ipaddress.ip_address(i) + 4) for i in grup_A["Network"]]

if grup_A["FTTM_VLAN_REF"].isnull().values.any() == True :
    grup_A[["FON", "VLAN"]] = grup_A[["FON", "VLAN"]].astype(np.int64)
else:
    grup_A[["FON", "VLAN","FTTM_VLAN_REF"]] = grup_A[["FON", "VLAN","FTTM_VLAN_REF"]].astype(np.int64)

# print(grup_A)
grup_A["FTTM_MW_RAN_Port_1"].fillna(grup_A["FTTM_MW_RAN_Port_1"].values[0], inplace=True)
grup_A["FTTM_MW_RAN_Port_2"].fillna(grup_A["FTTM_MW_RAN_Port_2"].values[0], inplace=True)

print(grup_A)

grup_C = pd.read_excel(f'{excel_File}', usecols=["Req_Type", "SOR_CRQ", "New_Port_Type_CSG", "New_Port_Setup_CSG"])
grup_C.rename(columns = {'New_Port_Type_CSG':'New_Port_Type'}, inplace = True)
grup_C[["CSG_Port1","CSG_Port2","CSG_Port3","CSG_Lag_ID","CSG_LAG_Desc","CSG_IP_System"]] = np.NaN
grup_C = grup_C.dropna(how = 'all')
grup_C["Req_Type"] = grup_C["Req_Type"].str.replace('VLAN-ADD', 'VA').str.replace('NEW-PORT', 'NP').str.replace('+', '_', regex=False)
# print (grup_C)

# grup_D = a.append(b).append(c).reset_index(drop=True).sort_values().drop_duplicates().reset_index(drop=True).dropna(how = 'all') 
'''
FutureWarning: The series.append method is deprecated and will be removed from pandas in a future version. Use pandas.concat instead.
'''

if grup_C["Req_Type"].values[0] == "VA" :
    grup_D = pd.DataFrame()
    grup_E = grup_A["FTTM_VLAN_REF"].drop_duplicates() 
else:
    grup_D = pd.concat((grup_A["FTTM_RAN_Port_1"], grup_A["FTTM_RAN_Port_2"], grup_A["FTTM_MW_RAN_Port_1"], grup_A["FTTM_MW_RAN_Port_2"]
                        )).reset_index(drop=True).sort_values().drop_duplicates().dropna(how = 'all')
    grup_E = pd.DataFrame()
    #grup_D = pd.concat((a, b, c, d)).reset_index(drop=True).sort_values().drop_duplicates().reset_index(drop=True).dropna(how = 'all')

grup_A.to_json('inputA.json', orient='records')
grup_C.to_json('inputC.json', orient='records')
grup_D.to_json('inputD.json', orient='records')
grup_E.to_json('inputE.json', orient='records')

baba = '''
{
    "result": {
        "GroupA": [],
        "GroupC": [],
        "GroupD": [],
        "GroupE": []
    }
}
'''

fon = grup_A.loc[0,'FON']
sor = grup_C.loc[0,'SOR_CRQ']
grouping = ["A", "C", "D", "E"]
file_json = f'{result_Path}/{userName}_{fon}_{sor}.json'
file_yaml = f'{result_Path}/{userName}_{fon}_{sor}.yaml'

# Create initial file and reset content of file_json
with open(file_json, 'w') as f:
    f.write(baba)

# Function to add to JSON

def write_json(new_data, filename=file_json):
    with open(filename, 'r+') as file:
        file_data = json.load(file)
        file_data["result"][f"Group{i}"].append(new_data)
        file.seek(0)
        json.dump(file_data, file, indent=4)

for i in grouping:
    with open(f'input{i}.json', '+r') as f:
        data = json.load(f)
        for j in data:
            write_json(j)

for i in grouping:
    for file in glob.glob(f"input{grouping}.json"):
        os.remove(file) 

with open(file_json, 'r') as file:
    configuration = json.load(file)

with open(file_yaml, 'w') as yaml_file:
    yaml.dump(configuration, yaml_file)

print ("")

for i in progressbar(range(100), "Computing: ", 50):
    time.sleep(0.00001) 

result_Name = f'JSON and YAML File created successfully! \n "{file_json}" \n "{file_yaml}"'
print ("- " * ((len(file_json) // 2)+2))
print (result_Name)
print ("- " * ((len(file_json) // 2)+2))
