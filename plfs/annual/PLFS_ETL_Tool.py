import pandas as pd
import configparser as cp
import psycopg2
from openpyxl.utils import column_index_from_string
from openpyxl.utils import get_column_letter
import time
import math

start_time = time.perf_counter()

config = cp.RawConfigParser()
config.read('plfs_2024.properties')

file_path = config.get('master_properties', 'file_path')
file_name = config.get('master_properties', 'file_name')

sheet_name = config.get('master_properties', 'sheet_name')

dsn_string = "host=10.24.89.9 port=5432 dbname=plfs_new_db user=postgres password=root456"

connection = psycopg2.connect(
    dsn=dsn_string,
    keepalives=1,
    keepalives_idle=30,
    keepalives_interval=10,
    keepalives_count=5
)

cursor = connection.cursor()

indicator_list = ["LFPR (Labour Force Participation Rate, in per cent)","WPR (Worker Population Ratio, in per cent)","UR (Unemployment Rate, in per cent)","Percentage distribution of workers","Percentage of regular wage/ salaried employees with employment condition as","Average wage/salary earnings (Rs. 0.00) during the preceding calendar month from regular wage/salaried employment","Average wage earnings (Rs. 0.00) per day from casual labour work other than public works","Average gross earnings (Rs. 0.00) during last 30 days from self-employment"]
status_list = ["PS+SS","CWS"]
industry_list = ["all","industry: (05-99)","industry: (014, 016, 017 , 02-99)","01-03 (agriculture)","05-09 (mining & quarrying)","10-33 (manufacturing)","35-39 (electricity and water supply)","41-43 (construction)","05-43 (secondary)","45-47 (trade)","49-53( transport)","55-56 (accommodation & food services)","58-99 (other services)","45-99 (tertiary)"]
broad_industry_work_list = ["primary sector (Div. 01-03 of NIC 2008)","secondary sector (Div. 05-43 of NIC 2008)","tertiary sector (Div. 45-99 of NIC 2008)","all"]
broad_status_employment_list = ["1.self-employed: own account worker, employer","2.self-employed: helper in household enterprise","3.all self employed","4.regular wage/salary","5.casual labour","all"]
education_list = ["1.not literate","2.literate & upto primary","3.middle","4.secondary","5.higher secondary","6.diploma/ certificate course","7.graduate","8.post graduate & above","9.secondary & above","all"]
industry_section_list = ["All","A_Agriculture, forestry and fishing","B_Mining and quarrying","C_Manufacturing","D_Electricity, gas, steam and air conditioning supply","E_Water supply; sewerage, waste management and remediation activities","F_Construction","G_Wholesale and retail trade; repair of motor vehicles and motorcycles","H_Transportation and storage","I_Accommodation and Food service activities","J_Information and communication","K_Financial and insurance activities","L_Real estate activities","M_Professional, scientific and technical activities","N_Administrative and support service activities","O_Public administration and defence compulsory social security","P_Education","Q_Human health and social work activities","R_Arts, entertainment and recreation","S_Other service activities","T_Activities of hhds as employers, undiff goods services prod actvs of hhds for own use","U_Activities of extraterritorial organizations and bodies"]
nco_division_list = ["all","NCO Division 1: LEGISLATORS, SENIOR OFFICIALS AND MANAGERS","NCO Division 2: PROFESSIONALS","NCO Division 3: TECHNICIANS AND ASSOCIATE PROFESSIONALS","NCO Division 4: CLERKS","NCO Division 5: SERVICE WORKERS AND SHOP & MARKET SALES WORKERS","NCO Division 6: SKILLED AGRICULTURAL AND FISHERY WORKERS","NCO Division 7: CRAFT AND RELATED TRADES WORKERS","NCO Division 8: PLANT AND MACHINE OPERATORS AND ASSEMBLERS","NCO Division 9: ELEMENTARY OCCUPATIONS","NCO Division X: WORKERS NOT CLASSIFIED BY OCCUPATIONS"]
enterprise_type_list = ["1.proprietary and partnership","2.govt./local body/ public sector enterprises","3.autonomous bodies","4.public/ private limited company","5.cooperative societies","6.trust/ other non profit inst","7.employer’s households","8.others","all"]
employee_contract_list = ["1.with no written job contract","2.not eligible for paid leave","3.without any social security benefit","4.not eligible for paid leave, w.o. written job contract and w.o. any SSB","all"]
religion_list = ["all","Hinduism","Islam","Christianity","Sikhism"]
social_group_list = ["all","scheduled tribe","scheduled caste","other backward class","others"]
age_group_list = ["15 years and above","15-29 years","15-59 years","all"]
state_code_list = ["Andhra Pradesh","Arunachal Pradesh","Assam","Bihar","Chhattisgarh","Delhi","Goa","Gujarat","Haryana","Himachal Pradesh","Jammu & Kashmir","Jharkhand","Karnataka","Kerala","Madhya Pradesh","Maharashtra","Manipur","Meghalaya","Mizoram","Nagaland","Odisha","Punjab","Rajasthan","Sikkim","Tamil Nadu","Telangana","Tripura","Uttarakhand","Uttar Pradesh","West Bengal","Andaman & N. Island","Chandigarh","Dadra & Nagar Haveli","Daman & Diu","Lakshadweep","Puducherry","Ladakh","Dadra & Nagar Haveli & Daman & Diu","All India"]

def sql_val(val):
    if val is None or str(val).lower() == 'none' or str(val).lower() == 'nan':
        return "NULL"
    return f"'{val}'" if isinstance(val, str) else str(val)

def insert_data():
    input_df = pd.read_excel(file_path+'//'+file_name,sheet_name,dtype=str)
    print('Dataframe loaded successfully')
    print('Number of rows in the dataframe:',len(input_df))
    count = 0
    
    for row in range(len(input_df)):
    #for row in range(10):
        insert_query_prefix = "insert into plfs_fact("+""
        insert_query_suffix = " values"
        temp = None 
        try:
            year = input_df.iloc[row].loc['Year'].strip()
            
            #print('Indicator name:',input_df.iloc[row].loc['Indicator'])
            temp = input_df.iloc[row].loc['Indicator'].strip()
            indicator_code = indicator_list.index(temp)+1
            #print('indicator_code:',indicator_code)

            temp = input_df.iloc[row].loc['Status'].strip()
            weekly_status_code = None
            
            if temp != None or temp != 'nan':
                weekly_status_code = status_list.index(temp)+1
                
            #print('Industry:',input_df.iloc[row].loc['Industry'])
            
            temp = input_df.iloc[row].loc['Industry']
            industry_code = None
            #print('Industry:',temp)
            
            if not pd.isna(temp):
                industry_code = industry_list.index(temp)+1
            #print('industry_code:',industry_code)
            temp = input_df.iloc[row].loc['BroadIndustryofWork']
            broad_industry_work_code = None
            if not pd.isna(temp):
                broad_industry_work_code = broad_industry_work_list.index(temp)+1

            temp = input_df.iloc[row].loc['BroadStatusInEmployment']
            #print('BroadStatusInEmployment:',temp)
            broad_status_employment_code = None
            if not pd.isna(temp):
                broad_status_employment_code = broad_status_employment_list.index(temp)+1
            #print('broad_status_employment_code:',broad_status_employment_code)
            temp = input_df.iloc[row].loc['Education']
            #print('Education:',temp)
            
            education_code = None
            if not pd.isna(temp):
                education_code = education_list.index(temp)+1
            #print('education_code:',education_code)

            temp = input_df.iloc[row].loc['IndustrySection']
            industry_section_code = None
            #print('IndustrySection:',temp)
            if not pd.isna(temp):
                industry_section_code = industry_section_list.index(temp)+1
            #print('industry_section_code:',industry_section_code)
            temp = input_df.iloc[row].loc['NCO Division']
            nco_division_code = None
            if not pd.isna(temp):
                for i in range(len(nco_division_list)):
                    if temp in nco_division_list[i]:
                        nco_division_code = i+1
                        break

            temp = input_df.iloc[row].loc['EnterpriseType']
            enterprise_type_code = None
            if not pd.isna(temp):
                enterprise_type_code = enterprise_type_list.index(temp)+1
            
            temp = input_df.iloc[row].loc['EmpCond']
            employee_contract_code = None
            if not pd.isna(temp):
                employee_contract_code = employee_contract_list.index(temp)+1
            
            temp = input_df.iloc[row].loc['Religion']
            religion_code = None
            if not pd.isna(temp):
                religion_code = religion_list.index(temp)+1

            social_group_code = None
            temp = input_df.iloc[row].loc['SocialGroup']
            if not pd.isna(temp):
                social_group_code = social_group_list.index(temp)+1

            temp = input_df.iloc[row].loc['age_group']
            age_group_code = None
            if not pd.isna(temp):
                age_group_code = age_group_list.index(temp)+1

            temp = input_df.iloc[row].loc['State/UT']
            state_code = None
            if not pd.isna(temp):
                state_code = state_code_list.index(temp)+1
            if temp == 'All India':
                state_code = 99
            rm = input_df.iloc[row].loc['rural_male'] or ''
            rf = input_df.iloc[row].loc['rural_female'] or ''
            rp = input_df.iloc[row].loc['rural_person'] or ''

            um = input_df.iloc[row].loc['urban_male'] or ''
            uf = input_df.iloc[row].loc['urban_female'] or ''
            up = input_df.iloc[row].loc['urban_person'] or ''

            cm = input_df.iloc[row].loc['rural_urban_male'] or ''
            cf = input_df.iloc[row].loc['rural_urban_female'] or ''
            cp = input_df.iloc[row].loc['rural_urban_person'] or ''

            year = input_df.iloc[row].loc['Year'].strip()
            fact_code = file_name+'_'+sheet_name+'_'+str(row)+'_'
            

            sql_query = """
                INSERT INTO plfs_fact (
                    plfs_fact_code, age_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, sector_code,
                    social_category_code, state_code, weekly_status_code,
                    indicator_value, frequency_code, year_type_code, year, gender_code
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """

            cursor.execute(sql_query, (
                    fact_code+'rm', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 1,
                    social_group_code, state_code, weekly_status_code,
                    rm, 1, 2, year, 1
                ))
            cursor.execute(sql_query, (
                    fact_code+'rf', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 1,
                    social_group_code, state_code, weekly_status_code,
                    rf, 1, 2, year, 2
                ))
            cursor.execute(sql_query, (
                    fact_code+'rp', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 1,
                    social_group_code, state_code, weekly_status_code,
                    rp, 1, 2, year, 3
                ))
            cursor.execute(sql_query, (
                    fact_code+'um', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 2,
                    social_group_code, state_code, weekly_status_code,
                    um, 1, 2, year, 1
                ))
            cursor.execute(sql_query, (
                    fact_code+'uf', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 2,
                    social_group_code, state_code, weekly_status_code,
                    uf, 1, 2, year, 2
                ))
            cursor.execute(sql_query, (
                    fact_code+'up', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 2,
                    social_group_code, state_code, weekly_status_code,
                    up, 1, 2, year, 3
                ))
            cursor.execute(sql_query, (
                    fact_code+'cm', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 3,
                    social_group_code, state_code, weekly_status_code,
                    cm, 1, 2, year, 1
                ))
            cursor.execute(sql_query, (
                    fact_code+'cf', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 3,
                    social_group_code, state_code, weekly_status_code,
                    cf, 1, 2, year, 2
                ))
            cursor.execute(sql_query, (
                    fact_code+'cp', age_group_code, broad_industry_work_code,
                    broad_status_employment_code, education_code,
                    employee_contract_code, enterprise_type_code,
                    indicator_code, industry_section_code,
                    nco_division_code, religion_code, 3,
                    social_group_code, state_code, weekly_status_code,
                    cp, 1, 2, year, 3
                ))



            #print('Insert query:',count)
            
            count = count + 1
            connection.commit()
            #print('Number of rows in the dataframe:',len(input_df))
        except ValueError as e:
            print("Value not found in the list:",temp,"in row:",row,'exception:',e)
    print('Number of Data Points Inserted:',count)
    return count

total_data = insert_data()
print("Total no of data points inserted:",total_data)
cursor.close()

end_time = time.perf_counter()
elapsed_time = end_time - start_time
print(f"Time taken to Execute ETL: {elapsed_time:.4f} seconds")

