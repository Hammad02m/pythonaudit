import os
import sys
import pandas as pd
import numpy as np
import math
import json


# Function to find the risk based on NACE code
def get_risk_by_nace(filepath, nace_code):
    excelpath = resource_path(filepath)
    data = pd.read_excel(excelpath, header=None)
    # Iterate through the rows to find the matching NACE code
    for idx, row in data.iterrows():
        # Check each of the first 4 columns (NACE codes)
        for col in range(4):  # Columns 0, 1, 2, 3
            if pd.notna(row[col]) and str(row[col]) == str(nace_code):
                # Risk is located in column 5 (F in Excel)
                risk = row[5]
                # If the risk is empty, check the row below
                if pd.isna(risk) and idx + 1 < len(data):
                    risk = data.iloc[idx + 1, 5]
                return risk.upper() if pd.notna(risk) else "LOW"
    return "LOW"  # Default to LOW if no matching NACE code is found
  
def extract_ISO9001values(file_path, sheet_name, employeescount):
    excelfile = resource_path(file_path)
    excel_data = pd.ExcelFile(excelfile)
    data = excel_data.parse(sheet_name)
    data.columns = data.iloc[1]  # Set the second row as header
    data = data[2:]  # Remove the first two rows used for headers
    data.reset_index(drop=True, inplace=True)

    # Convert numerical columns to numeric types for calculations
    required_columns = ['Employees', '80%rule', 'visit', 'prep.', 'audit', 'report']
    for col in required_columns:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors='coerce')

    selected_data = data[required_columns]
    #totaleffectiveemployees = employeescount
    matching_rows = selected_data[selected_data['Employees'] <= employeescount]
    if not matching_rows.empty:
        matching_row = matching_rows.iloc[matching_rows['Employees'].idxmax()]

        visit_time = matching_row['visit']
        prep_time = matching_row['prep.']
        audit_time = matching_row['audit']
        report_time = matching_row['report']
        return {"visit": visit_time, "prep": prep_time, "audit": audit_time, "report": report_time, "risk": "N/A"}
    else:
        return None

def extract_ISO45001cert_values(file_path, personnel_count, risk):

    # Read the Excel file
    excelpath = resource_path(file_path)
    data = pd.read_excel(excelpath, header=None)
    
    # Define column mappings based on risk
    risk_columns = {
        "LOW": {"Visit": 5, "Prep": 6, "Audit": 7, "Report": 8},
        "MEDIUM": {"Visit": 13, "Prep": 14, "Audit": 15, "Report": 16},
        "HIGH": {"Visit": 21, "Prep": 22, "Audit": 23, "Report": 24},
    }
    
    # Select the correct column indices for the given risk
    columns = risk_columns[risk]
    
    # Iterate through the rows to find the correct personnel range
    for i in range(4, len(data)):  # Start from row 5 (index 4) to skip the headers
        lower_bound = data.iloc[i, 0]  # Start of personnel range
        upper_bound = data.iloc[i + 1, 0] if i + 1 < len(data) else float('inf')  # End of range or infinity

        # Ensure bounds are converted to numbers for comparison
        try:
            lower_bound = int(lower_bound)
            upper_bound = int(upper_bound)
        except ValueError:
            continue  # Skip rows where bounds aren't numbers

        if lower_bound <= personnel_count < upper_bound:
            # Extract values from the corresponding columns
            visit = data.iloc[i, columns["Visit"]]
            prep = data.iloc[i, columns["Prep"]]
            audit = data.iloc[i, columns["Audit"]]
            report = data.iloc[i, columns["Report"]]

            prep/=8
            report/=8
            return {"visit": visit, "prep": prep, "audit": audit, "report": report, "risk": risk}
    
    # If no matching range is found, return None
    return None

def extract_ISO45001recert_values(file_path, personnel_count, risk):
    # Read the Excel file
    excelpath = resource_path(file_path)
    data = pd.read_excel(excelpath, header=None)
    
    # Define column mappings based on risk
    risk_columns = {
        "LOW": {"Prep": 5, "Audit": 6, "Report": 7},
        "MEDIUM": {"Prep": 12, "Audit": 13, "Report": 14},
        "HIGH": {"Prep": 19, "Audit": 20, "Report": 21},
    }
    
    # Select the correct column indices for the given risk
    columns = risk_columns[risk]
    
    # Iterate through the rows to find the correct personnel range
    for i in range(4, len(data)):  # Start from row 5 (index 4) to skip the headers
        lower_bound = data.iloc[i, 0]  # Start of personnel range
        upper_bound = data.iloc[i + 1, 0] if i + 1 < len(data) else float('inf')  # End of range or infinity

        # Ensure bounds are converted to numbers for comparison
        try:
            lower_bound = int(lower_bound)
            upper_bound = int(upper_bound)
        except ValueError:
            continue  # Skip rows where bounds aren't numbers

        if lower_bound <= personnel_count < upper_bound:
            # Extract values from the corresponding columns
            prep = data.iloc[i, columns["Prep"]]
            audit = data.iloc[i, columns["Audit"]]
            report = data.iloc[i, columns["Report"]]
            prep/=8
            report/=8
            visit=0
            return {"prep": prep, "audit": audit, "report": report, "visit": visit, "risk": risk}
    
    # If no matching range is found, return None
    return None

def extract_ISO14001recert_values(file_path, personnel_count, risk):
        # Read the Excel file
    excelpath = resource_path(file_path)
    data = pd.read_excel(excelpath, header=None)
    
    # Define column mappings based on risk
    risk_columns = {
        "LIMITED": {"Prep": 5, "Audit": 6, "Report": 7},
        "LOW": {"Prep": 12, "Audit": 13, "Report": 14},
        "MEDIUM": {"Prep": 19, "Audit": 20, "Report": 21},
        "HIGH": {"Prep": 26, "Audit": 27, "Report": 28},
    }
    
    # Select the correct column indices for the given risk
    columns = risk_columns[risk]
    
    # Iterate through the rows to find the correct personnel range
    for i in range(4, len(data)):  # Start from row 5 (index 4) to skip the headers
        lower_bound = data.iloc[i, 0]  # Start of personnel range
        upper_bound = data.iloc[i + 1, 0] if i + 1 < len(data) else float('inf')  # End of range or infinity

        # Ensure bounds are converted to numbers for comparison
        try:
            lower_bound = int(lower_bound)
            upper_bound = int(upper_bound)
        except ValueError:
            continue  # Skip rows where bounds aren't numbers

        if lower_bound <= personnel_count < upper_bound:
            # Extract values from the corresponding columns
            prep = data.iloc[i, columns["Prep"]]
            audit = data.iloc[i, columns["Audit"]]
            report = data.iloc[i, columns["Report"]]
            prep/=8
            report/=8
            visit = 0
            return {"visit": visit, "prep": prep, "audit": audit, "report": report, "risk": risk}
    
    # If no matching range is found, return None
    return None

def extract_ISO14001cert_values(file_path, personnel_count, risk):
        # Read the Excel file
    excelpath = resource_path(file_path)
    data = pd.read_excel(excelpath, header=None)
    
    # Define column mappings based on risk
    risk_columns = {
        "LIMITED": {"Visit": 5, "Prep": 6, "Audit": 7, "Report": 8},
        "LOW": {"Visit": 13, "Prep": 14, "Audit": 15, "Report": 16},
        "MEDIUM": {"Visit": 21, "Prep": 22, "Audit": 23, "Report": 24},
        "HIGH": {"Visit": 29, "Prep": 30, "Audit": 31, "Report": 32},
    }
    
    # Select the correct column indices for the given risk
    columns = risk_columns[risk]
    
    # Iterate through the rows to find the correct personnel range
    for i in range(4, len(data)):  # Start from row 5 (index 4) to skip the headers
        lower_bound = data.iloc[i, 0]  # Start of personnel range
        upper_bound = data.iloc[i + 1, 0] if i + 1 < len(data) else float('inf')  # End of range or infinity

        # Ensure bounds are converted to numbers for comparison
        try:
            lower_bound = int(lower_bound)
            upper_bound = int(upper_bound)
        except ValueError:
            continue  # Skip rows where bounds aren't numbers

        if lower_bound <= personnel_count < upper_bound:
            # Extract values from the corresponding columns
            prep = data.iloc[i, columns["Prep"]]
            audit = data.iloc[i, columns["Audit"]]
            report = data.iloc[i, columns["Report"]]
            visit = data.iloc[i, columns["Visit"]]
            prep/=8
            report/=8
            return {"visit": visit, "prep": prep, "audit": audit, "report": report, "risk": risk}
    
    # If no matching range is found, return None
    return None

def readstandarddata(nace_code, standard, renewaud, path, employees, sheet, complexity):
    if(standard=="ISO 9001"):
        auditdata = extract_ISO9001values(path, sheet, employees)
        return auditdata
    if(standard=="ISO 45001"):
        risk = complexity
        if(renewaud==True):
            auditdata = extract_ISO45001recert_values(path,employees, risk)
            return auditdata
        else:
            auditdata = extract_ISO45001cert_values(path, employees, risk)
            return auditdata
    if(standard=="ISO 14001"):
        risk = complexity
        if(renewaud==True):
            auditdata = extract_ISO14001recert_values(path, employees, risk)
            return auditdata
        else:
            auditdata = extract_ISO14001cert_values(path, employees, risk) 
            return auditdata
        
def integrationestimate(x,y):
    reduction_matrix = {
    0: {0:0, 10:0, 20:0, 30:0, 40:0, 50:0, 60:0, 70:0, 80:0, 90:0, 100:0},
    10: {0:0, 10:0, 20:0, 30:0, 40:0, 50:0, 60:0, 70:0, 80:0, 90:0, 100:0},
    20: {0:0, 10:0, 20:0, 30:0, 40:0, 50:0, 60:0, 70:0, 80:0, 90:0, 100:0},
    30: {0:0, 10:0, 20:0, 30:5, 40:5, 50:5, 60:5, 70:5, 80:5, 90:5, 100:5},
    40: {0:0, 10:0, 20:0, 30:5, 40:5, 50:5, 60:5, 70:5, 80:5, 90:5, 100:5},
    50: {0:0, 10:0, 20:0, 30:5, 40:5, 50:10, 60:10, 70:0, 80:10, 90:10, 100:10},
    60: {0:0, 10:0, 20:0, 30:5, 40:5, 50:10, 60:10, 70:0, 80:10, 90:10, 100:10},
    70: {0:0, 10:0, 20:0, 30:5, 40:5, 50:10, 60:10, 70:15, 80:15, 90:15, 100:15},
    80: {0:0, 10:0, 20:0, 30:5, 40:5, 50:10, 60:10, 70:15, 80:15, 90:15, 100:15},
    90: {0:0, 10:0, 20:0, 30:5, 40:5, 50:10, 60:10, 70:15, 80:15, 90:20, 100:20},
    100: {0:0, 10:0, 20:0, 30:5, 40:5, 50:10, 60:10, 70:15, 80:15, 90:20, 100:20}
}


    levels = sorted(reduction_matrix.keys())  # X-axis: Level of Integration %
    audits = sorted(reduction_matrix[levels[0]].keys())  # Y-axis: Ability to Perform Combined Audit %

    # Find the closest lower or equal value for X and Y
    x_floor = max([val for val in levels if val <= x], default=levels[0])
    y_floor = max([val for val in audits if val <= y], default=audits[0])

    # Ensure values don't exceed the highest available values
    x_floor = min(x_floor, max(levels))
    y_floor = min(y_floor, max(audits))
    # Return the predefined reduction value
    reductionpercentage = reduction_matrix[x_floor][y_floor] 
    reductionpercentage /= 100
    return reductionpercentage


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def audittime(auditdata):

#  Received data from React: {
  #companyName: 'Viscount',
  #numberOfSites: '1',
  #renewal: false,
  #surveillanceAudits: '2',
  #isoStandards: [ 'ISO 9001', 'ISO 14001' ],
  #sites: {
    #'0': {
      #management: 3,
      #otherManagers: 6,
      #partTimeEmployees: '',
      #standardsData: [Object],
      #highlySkilledWorkers: 19,
      #unskilledWorkers: 429,
     # naceCode: 25.1,
    #  centralFunctions: 3
   # }
  #},
 # selectedFactors: {},
 # integrationResponses: {
   # 'An integrated documentation set, including work instructions to a good level of development as appropriate': true,
   # 'Management Reviews that consider the overall business strategy and plan': true,
   # 'An integrated approach to internal audits': true,
   # 'An integrated approach to policy and objectives': true,
    #'An integrated approach to improvement mechanisms (corrective and preventive action; measurement and continual improvement)': true,
    #'An integrated approach to system processes': true,
   # 'Integrated management support and responsibilities': true
  #},
  #auditorCount: '2',
  #auditorQualifications: [ 2, '2', '1' ],
  #selectedStandards: [],
 # integrationLevel: 100,
 # auditCombinationAbility: 100
#}
    try:
        integrationlevel = float(auditdata["integrationLevel"])
    except (ValueError, TypeError, KeyError):
        integrationlevel = 0 # Default to 0 if there's an error
    
    try:
        auditcombinationability = float(auditdata["auditCombinationAbility"])
    except (ValueError, TypeError, KeyError):
        auditcombinationability = 0 # Default to 0 if there's an error

    integrationfactor = integrationestimate(integrationlevel, auditcombinationability)
    total_site_times = {}  # Total audit time for each site
    total_stage1_times = {}  # Total Stage 1 time for each site
    total_surv_times = {}  # Total surveillance time for each site
    file_paths = {
        "ISO 9001": {
            "recert": 'recert9001.xlsx',
            "cert": 'cert9001.xlsx',
            "sheet": {"recert": 'Recert 9001', "cert": 'Cert 9001'}
        },
        "ISO 45001": {
            "recert": 'recert45001.xlsx',
            "cert": 'cert45001.xlsx',
            "sheet": {"recert": 'Recert 45001', "cert": 'Cert 45001'}
        },
        "ISO 14001": {
            "recert": 'recert14001.xlsx',
            "cert": 'cert14001.xlsx',
            "sheet": {"recert": 'Recert 14001', "cert": 'Cert 14001'}
        }
    }
    try:
        surv = int(auditdata["surveillanceAudits"])
    except (ValueError, TypeError, KeyError):
        surv = 2  # Default to 2 if there's an error
    try:
        rene = auditdata["renewal"]
    except (ValueError, TypeError, KeyError):
        rene = False
    try:
        noofsites = int(auditdata["numberOfSites"])
    except (ValueError, TypeError, KeyError):
        noofsites = 1  # Default to 1 if there's an error


    results = {standard: [] for standard in auditdata["isoStandards"]}
    totals = {site_id: [] for site_id in auditdata["sites"]}
    for site_id, site in auditdata["sites"].items():
        site_standards = site.get("standardsData", {})

        for standard, standard_data in site_standards.items():
            
            try:
                skilled_percent = float(standard_data.get("skilledPercentage", 100)) / 100
            except (ValueError, TypeError, KeyError):
                skilled_percent = 1.0  # Default to 100% if there's an error

            try:
                unskilled_percent = float(standard_data.get("unskilledPercentage", 5)) / 100
            except (ValueError, TypeError, KeyError):
                unskilled_percent = 0.05  # Default to 5% if there's an error

            try:
                justification = standard_data.get("justification", "")
            except (ValueError, TypeError, KeyError):
                justification = ""  # Default to empty string if there's an error


            try:
                if site_id == 0:
                    centralsitefun = site.get("centralFunctions", 0)
            except (ValueError, TypeError, KeyError):
                centralsitefun = 0  # Default to 0 if there's an error

            try:
                nacecode = site.get("naceCode", 1)
            except (ValueError, TypeError, KeyError):
                nacecode = 1  # Default to empty string if there's an error

            try:
                management = int(site.get("management", 0))
            except (ValueError, TypeError, KeyError):
                management = 0  # Default to 0 if there's an error

            try:
                other_managers = int(site.get("otherManagers", 0))
            except (ValueError, TypeError, KeyError):
                other_managers = 0  # Default to 0 if there's an error

            try:
                highly_skilled = int(site.get("highlySkilledWorkers", 0))
            except (ValueError, TypeError, KeyError):
                highly_skilled = 0  # Default to 0 if there's an error

            try:
                unskilled = int(site.get("unskilledWorkers", 0))
            except (ValueError, TypeError, KeyError):
                unskilled = 0  # Default to 0 if there's an error

            try: 
                partTimeEmployees = int(site.get("partTimeEmployees", 0))
            except (ValueError, TypeError, KeyError):
                partTimeEmployees = 0

            try:
                hoursWorkedPartTime = int(site.get("hoursWorkedPartTime", 0))
            except (ValueError, TypeError, KeyError):
                hoursWorkedPartTime = 0

            try:
                reductionfactor = float(auditdata["totalReductionPercentages"][standard])
            except (ValueError, TypeError, KeyError):
                reductionfactor = 0

            try: 
                increaseDays = float(auditdata["increaseDays"][standard])
            except (ValueError, TypeError, KeyError):
                increaseDays = 0

            try:
                sitefunctions = int(site.get("centralFunctions", 0))
            except (ValueError, TypeError, KeyError):
                sitefunctions = 0  # Default to 0 if there's an error

            if(standard=="ISO 9001"):
                riskk = "N/A"
            if(standard=="ISO 45001"):
                riskk = get_risk_by_nace('risk45001.xlsx', nacecode)
            if(standard=="ISO 14001"):          
                riskk = get_risk_by_nace('risk14001.xlsx', nacecode)
            
            try:
                #formData.sites[selectedSite]?.standardsData?.[standard]?.complexit
                complexity = (standard_data.get("complexity", riskk))
            except (ValueError, TypeError, KeyError):
                complexity = riskk  # Default to LOW if there's an error

            if site_id not in total_site_times:
                    total_site_times[site_id] = 0
                    total_stage1_times[site_id] = 0
                    total_surv_times[site_id] = 0
           # part_time = int(site.get("partTimeEmployees", 0))
           # part_time_hours = int(site.get("hoursWorkedPartTime", 0))
            effective_employees = math.ceil(
                management + other_managers +
                (highly_skilled * skilled_percent) +
                (unskilled * unskilled_percent) + ((partTimeEmployees * hoursWorkedPartTime) / 8)
            )
            file_path = file_paths[standard]["recert"] if rene == True else file_paths[standard]["cert"]
            sheet_name = file_paths[standard]["sheet"]["recert"] if rene == True else file_paths[standard]["sheet"]["cert"]

            audit_data = readstandarddata(nacecode, standard, rene, file_path, effective_employees, sheet_name,complexity)
            visit_time = round(audit_data["visit"], 3)
            prep_time = round(audit_data["prep"],3)
            audit_time = round(audit_data["audit"], 3)
            report_time = round(audit_data["report"], 3)
            stage1plus2 = visit_time + audit_time
            total_time = round(visit_time + prep_time + audit_time + report_time,2)

            addred=0
            whatami = int(str(site_id))
            if(whatami==0):
                central_fun_count = sitefunctions
                addred=0
            if(whatami>0):
                audit_time = audit_time + visit_time
                visit_time=0
                if(sitefunctions>0):
                    addred = ((sitefunctions - central_fun_count)/central_fun_count)*20
                else: 
                    addred = 20
            
            adjusted_audit_time = round(audit_time * (1-((reductionfactor+addred)/100)),2) + increaseDays
            adjusted_total_time = round(adjusted_audit_time + prep_time + report_time+ visit_time,2)
            adjusted_stage1plus2 = round(adjusted_audit_time + visit_time,2)

            if(surv>2):
                survauditdays = 0
                surveil_prep_report = 0
            else:
                if(standard=="ISO 9001"):
                    survauditdays = adjusted_audit_time / 3
                else:
                    survauditdays = adjusted_stage1plus2 / 3

            surveil_prep_report = (prep_time + report_time)

            if(standard=="ISO 9001"):
                if rene == False:
                    surveil_prep_report /= (surv+1)
                else:
                    if(surv==2): 
                        surveil_prep_report /= 2
                    if(surv==3):
                        surveil_prep_report *= (3/8)
                    if(surv==4): 
                        surveil_prep_report *= (3/10)
                    if(surv==5):
                        surveil_prep_report/=4
            else:
                if(rene==True):
                    surveil_prep_report /= (surv+1)
                else:
                    if(surv==2):
                        surveil_prep_report /= 2
                    if(surv==3):
                        surveil_prep_report *= (3/8)
                    if(surv==4):
                        surveil_prep_report *= (3/10)
                    if(surv==5):
                        surveil_prep_report /= 4


            total_site_times[site_id] += adjusted_audit_time
            total_stage1_times[site_id] += visit_time
            total_surv_times[site_id] += survauditdays

            results[standard].append({
                "site_id": site_id,
                "effective_employees": effective_employees,
                'visit_time': visit_time,
                'prep_time': prep_time, 
                'audit_time': audit_time,
                'report_time': report_time,
                'total_time': total_time,
                'stage1plus2': stage1plus2,
                'justification': justification,
                'surveillance': surv,
                'noofsites': noofsites,
                'whatami': whatami,
                'addred': addred,
                'redfac': reductionfactor,
                'central_fun_count': central_fun_count,
                'sitefunctions': sitefunctions,
                'adjusted_audit_time': round(adjusted_audit_time,2),
                'adjusted_total_time': round(adjusted_total_time,2),
                'adjusted_stage1plus2': round(adjusted_stage1plus2,2),
                'survauditdays': round(survauditdays,2),
                'surveil_prep_report': round(surveil_prep_report,2),
                'risk' : riskk,
                'complexity' : complexity,
            })
    
    results["totals"] = {
    "Integrated Times for Site": [
        {
            "site_id": site_id,
            "total_site_time": round(total_site_times[site_id],2),
            "total_stage1_time": round(total_stage1_times[site_id],2),
            "total_surv_time": round(total_surv_times[site_id],2),
            "Integration Reduction Percentage": integrationfactor,
            "Total Integrated Audit Time": round(total_site_times[site_id] * (1 - integrationfactor), 2),
            "Total Integrated Stage 1 Time": round(total_stage1_times[site_id] * (1 - integrationfactor), 2),
            "Total Integrated Surveillance Time": round(total_surv_times[site_id] * (1 - integrationfactor), 2),
        }
        for site_id in total_site_times
    ]
    }


    
    # Output the JSON result (Node.js reads this)
    print(json.dumps(results))

# Read JSON from stdin (sent by Node.js)
if __name__ == "__main__":
    input_data = sys.stdin.read()
    auditdata = json.loads(input_data)
    audittime(auditdata)


