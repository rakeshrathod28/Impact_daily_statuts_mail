import pandas as pd
import os
import win32com.client
from datetime import datetime, timedelta

# ==========================================
# 1. CONFIGURATION
# ==========================================
UNORGANIZED_INPUT = r"data/unorganized_input.txt"
ORGANIZED_FILE = r"data/organized_file.xlsx"
GIVEN_DATA_FILE = r"data/given_data.xlsx"
OUTPUT_DATA1 = r"data/output_data1.xlsx"
MATCH_COLUMN = "ALIAS"
RECORD_COUNTS_FILE = r"data/job_counts.xlsx"

# ==========================================
# 2. STEP 1: ORGANIZE DATA
# ==========================================
def organize_input_data():
    print("Step 1: Organizing raw input data...")
    try:
        if not os.path.exists(UNORGANIZED_INPUT):
            print(f"Error: {UNORGANIZED_INPUT} not found!")
            return None

        df = pd.read_csv(UNORGANIZED_INPUT, skipinitialspace=True)
        df.columns = df.columns.str.strip().str.upper()
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)

        time_cols = ['STARTED', 'COMPLETED']
        for col in time_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], format='%Y.%m.%d %H.%M', errors='coerce').dt.strftime('%I:%M %p ET')

        if 'DURATION' in df.columns:
            td = pd.to_timedelta(df['DURATION'].astype(str), errors='coerce')
            df['DURATION'] = td.apply(lambda x: 
                f"{int(x.total_seconds() // 3600)} Hours and {int((x.total_seconds() % 3600) // 60)} Minutes" 
                if pd.notnull(x) else "N/A"
            )

        if 'STATUS' in df.columns:
            df['STATUS'] = df['STATUS'].apply(lambda x: "Completed" if str(x).strip().upper() == "COMPLETED NORMAL" else "")

        with pd.ExcelWriter(ORGANIZED_FILE, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']
            green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})

            if 'STATUS' in df.columns:
                status_col_idx = df.columns.get_loc('STATUS')
                worksheet.conditional_format(1, status_col_idx, len(df), status_col_idx, {
                    'type':     'cell',
                    'criteria': 'equal to',
                    'value':    '"Completed"',
                    'format':   green_format
                })

        print(f" -> SUCCESS: Cleaned and color-coded data saved to {ORGANIZED_FILE}")
        return df

    except Exception as e:
        print(f"Error in Step 1: {e}")
        return None

# ==========================================
# 3. STEP 2: COMPARE & COMPUTE
# ==========================================
def process_and_match(organized_df):
    print("Step 2: Comparing with given data and filtering...")
    try:
        if not os.path.exists(GIVEN_DATA_FILE):
            print(f"Error: {GIVEN_DATA_FILE} not found!")
            return None

        given_df = pd.read_excel(GIVEN_DATA_FILE)
        given_df.columns = given_df.columns.str.strip().str.upper()

        matched_df = pd.merge(organized_df, given_df, on=MATCH_COLUMN, how='inner')
        available_cols = [MATCH_COLUMN, 'STARTED', 'STATUS', 'COMPLETED', 'DURATION']
        final_cols = [c for c in available_cols if c in matched_df.columns]
        
        final_df = matched_df[final_cols]
        final_df.to_excel(OUTPUT_DATA1, index=False)
        print(f" -> SUCCESS: Extracted {len(final_df)} matches into {OUTPUT_DATA1}")
        return final_df
    except Exception as e:
        print(f"Error in Step 2: {e}")
        return None

# ==========================================
# 4. STEP 3: EXTRACT TIMINGS & STRINGS SAFELY
# ==========================================
def prepare_job_metrics():
    print("Extracting job metrics and formatting values...")
    df = pd.read_excel(ORGANIZED_FILE)
    
    # Standardized list of all trackable process variables
    jobs_to_init = [
        'jobs_formatted', 'jobs_formatted0', 'jobs_formatted1', 'jobs_formatted2', 
        'jobs_formatted3', 'jobs_formatted4', 'jobs_formatted5', 'jobs_formatted_PSTP5092',
        'jobs_formatted_TAPPS1SD', 'jobs_formatted_PSTP2083', 'jobs_formatted_PSTP9050',
        'jobs_formatted_PSTP9051', 'jobs_formatted_PSTP2084', 'jobs_formatted_PSTP9052',
        'jobs_formatted_PSTP208G', 'jobs_formatted_PSTP0030', 'jobs_formatted_PSTP3160',
        'jobs_formatted_PSTP6506', 'jobs_formatted_PSTP0038', 'jobs_formatted_PSTP1000',
        'jobs_formatted_PSTP2035', 'jobs_formatted_PSTP2088', 'jobs_formatted_PSTP3155',
        'jobs_formatted_PSTP9060', 'jobs_formatted_PSTP3156', 'jobs_formatted_PSTP9061',
        'jobs_formatted_PSTP9062', 'jobs_formatted_PSTP5026', 'jobs_formatted_PSTP2010',
        'jobs_formatted_PSTP2115', 'jobs_formatted_PSTP2126', 'jobs_formatted_PSTP2220',
        'jobs_formatted_PSTP2330', 'jobs_formatted_PSTP3090', 'jobs_formatted_PSTP3095',
        'jobs_formatted_PSTP3097', 'jobs_formatted_PSTP3098', 'jobs_formatted_PSTP3099',
        'jobs_formatted_PSTP3169', 'jobs_formatted_PSTP3170', 'jobs_formatted_PSTP509F',
        'jobs_formatted_PSTP4070', 'jobs_formatted_PSTP4170', 'jobs_formatted_PSTP4180',
        'jobs_formatted_PSTP4210', 'jobs_formatted_PSTP3142', 'jobs_formatted_PSTP3145',
        'jobs_formatted_PSTP3149', 'jobs_formatted_PSTP3147', 'jobs_formatted_PSTP5010',
        'jobs_formatted_PSTP9601', 'jobs_formatted_PSTP5050', 'jobs_formatted_PSTP5106',
        'jobs_formatted_PSTP5089', 'jobs_formatted_PSTP5096', 'jobs_formatted_PSTP4095',
        'jobs_formatted_PSTP508E', 'jobs_formatted_PSTP4320', 'jobs_formatted_CICSPPS1',
        'jobs_formatted_PSTP5000', 'jobs_formatted_PSTP5092A', 'jobs_formatted_TAPPS1SDA',
        'jobs_formatted_PSTP2083A', 'jobs_formatted_PSTP9050A', 'jobs_formatted_PSTP9051A',
        'jobs_formatted_PSTP2084A', 'jobs_formatted_PSTP9052A', 'jobs_formatted_PSTP208GA',
        'jobs_formatted_PSTP0030A', 'jobs_formatted_PSTP3160A', 'jobs_formatted_PSTP6506A',
        'jobs_formatted_PSTP0038A', 'jobs_formatted_PSTP1000A', 'jobs_formatted_PSTP2035A',
        'jobs_formatted_PSTP2088A', 'jobs_formatted_PSTP3155A', 'jobs_formatted_PSTP9060A',
        'jobs_formatted_PSTP3156A', 'jobs_formatted_PSTP9061A', 'jobs_formatted_PSTP9062A',
        'jobs_formatted_PSTP5026A', 'jobs_formatted_PSTP2010A', 'jobs_formatted_PSTP2115A',
        'jobs_formatted_PSTP2126A', 'jobs_formatted_PSTP2220A', 'jobs_formatted_PSTP2330A',
        'jobs_formatted_PSTP3090A', 'jobs_formatted_PSTP3095A', 'jobs_formatted_PSTP3097A',
        'jobs_formatted_PSTP3098A', 'jobs_formatted_PSTP3099A', 'jobs_formatted_PSTP3169A',
        'jobs_formatted_PSTP3170A', 'jobs_formatted_PSTP509FA', 'jobs_formatted_PSTP4070A',
        'jobs_formatted_PSTP4170A', 'jobs_formatted_PSTP4180A', 'jobs_formatted_P210A' if 'jobs_formatted_P210A' in locals() else 'jobs_formatted_PSTP4210A',
        'jobs_formatted_PSTP3142A', 'jobs_formatted_PSTP3145A', 'jobs_formatted_PSTP3149A',
        'jobs_formatted_PSTP3147A', 'jobs_formatted_PSTP5010A', 'jobs_formatted_PSTP9601A',
        'jobs_formatted_PSTP5050A', 'jobs_formatted_PSTP5106A', 'jobs_formatted_PSTP5089A',
        'jobs_formatted_PSTP5096A', 'jobs_formatted_PSTP4095A', 'jobs_formatted_PSTP508EA',
        'jobs_formatted_PSTP4320A', 'jobs_formatted_CICSPPS1A', 'jobs_formatted_PSTP5000A'
    ]
    
    context = {var: "" for var in jobs_to_init}
    context['jobs_formatted4'] = "N/A"

    def get_val(alias, col):
        sub = df[df['ALIAS'] == alias]
        return str(sub[col].iloc[0]).strip() if not sub.empty and pd.notnull(sub[col].iloc[0]) else ""

    # Top Email Block Strings
    context['jobs_formatted'] = get_val('PSTP4020', 'COMPLETED')   # Pay process complete
    context['jobs_formatted0'] = get_val('PSTP4020', 'DURATION')    # Pay process duration
    context['jobs_formatted1'] = get_val('PSTP3095', 'COMPLETED')   # CALC Balancing complete
    context['jobs_formatted2'] = get_val('PSTP2330', 'STARTED')     # IMPACT Cycle start
    context['jobs_formatted3'] = get_val('PSTP2200 (Feed) (Exc 1st ,2nd B-Day)', 'COMPLETED')
    context['jobs_formatted5'] = get_val('PSTP3010', 'STARTED')     # CALC start time

    # Calculate Segment Duration
    s_job = df[df['ALIAS'] == 'PSTP3010']
    e_job = df[df['ALIAS'] == 'PSTP3095']
    if not s_job.empty and not e_job.empty:
        try:
            start_t = pd.to_datetime(s_job['STARTED'].iloc[0], format='%I:%M %p ET')
            end_t = pd.to_datetime(e_job['COMPLETED'].iloc[0], format='%I:%M %p ET')
            diff = end_t - start_t
            hours, remainder = divmod(int(diff.total_seconds()), 3600)
            minutes, _ = divmod(remainder, 60)
            context['jobs_formatted4'] = f"{hours} Hours and {minutes} Minutes"
        except Exception as e:
            print(f"Duration Calculation Warning: {e}")

    # Explicitly Map to clean dictionary values
    aliases_map = {
        'PSTP5092': 'jobs_formatted_PSTP5092', 'TAPPS1SD': 'jobs_formatted_TAPPS1SD', 'PSTP2083': 'jobs_formatted_PSTP2083',
        'PSTP9050': 'jobs_formatted_PSTP9050', 'PSTP9051': 'jobs_formatted_PSTP9051', 'PSTP2084': 'jobs_formatted_PSTP2084',
        'PSTP9052': 'jobs_formatted_PSTP9052', 'PSTP208G': 'jobs_formatted_PSTP208G', 'PSTP0030': 'jobs_formatted_PSTP0030',
        'PSTP3160': 'jobs_formatted_PSTP3160', 'PSTP6506': 'jobs_formatted_PSTP6506', 'PSTP0038': 'jobs_formatted_PSTP0038',
        'PSTP1000': 'jobs_formatted_PSTP1000', 'PSTP2035': 'jobs_formatted_PSTP2035', 'PSTP2088': 'jobs_formatted_PSTP2088',
        'PSTP3155': 'jobs_formatted_PSTP3155', 'PSTP9060': 'jobs_formatted_PSTP9060', 'PSTP3156': 'jobs_formatted_PSTP3156',
        'PSTP9061': 'jobs_formatted_PSTP9061', 'PSTP9062': 'jobs_formatted_PSTP9062', 'PSTP5026': 'jobs_formatted_PSTP5026',
        'PSTP2010': 'jobs_formatted_PSTP2010', 'PSTP2115': 'jobs_formatted_PSTP2115', 'PSTP2126': 'jobs_formatted_PSTP2126',
        'PSTP2220': 'jobs_formatted_PSTP2220', 'PSTP2330': 'jobs_formatted_PSTP2330', 'PSTP3090': 'jobs_formatted_PSTP3090',
        'PSTP3095': 'jobs_formatted_PSTP3095', 'PSTP3097': 'jobs_formatted_PSTP3097', 'PSTP3098': 'jobs_formatted_PSTP3098',
        'PSTP3099': 'jobs_formatted_PSTP3099', 'PSTP3169': 'jobs_formatted_PSTP3169', 'PSTP3170': 'jobs_formatted_PSTP3170',
        'PSTP509F': 'jobs_formatted_PSTP509F', 'PSTP4070': 'jobs_formatted_PSTP4070', 'PSTP4170': 'jobs_formatted_PSTP4170',
        'PSTP4180': 'jobs_formatted_PSTP4180', 'PSTP4210': 'jobs_formatted_PSTP4210', 'PSTP3142': 'jobs_formatted_PSTP3142',
        'PSTP3145': 'jobs_formatted_PSTP3145', 'PSTP3149': 'jobs_formatted_PSTP3149', 'PSTP3147': 'jobs_formatted_PSTP3147',
        'PSTP5010': 'jobs_formatted_PSTP5010', 'PSTP9601': 'jobs_formatted_PSTP9601', 'PSTP5050': 'jobs_formatted_PSTP5050',
        'PSTP5106': 'jobs_formatted_PSTP5106', 'PSTP5089': 'jobs_formatted_PSTP5089', 'PSTP5096': 'jobs_formatted_PSTP5096',
        'PSTP4095': 'jobs_formatted_PSTP4095', 'PSTP508E': 'jobs_formatted_PSTP508E', 'PSTP4320': 'jobs_formatted_PSTP4320',
        'CICSPPS1': 'jobs_formatted_CICSPPS1', 'PSTP5000': 'jobs_formatted_PSTP5000'
    }

    for alias, var_name in aliases_map.items():
        context[var_name] = get_val(alias, 'COMPLETED')
        context[var_name + "A"] = get_val(alias, 'STATUS')

    return context

# ==========================================
# 5. STEP 4: EMAIL LOGIC
# ==========================================
def send_outlook_email(df_matched, metrics_context):
    print("Step 3: Preparing email via Outlook...")
    try:
        today = datetime.today().date()
        yesterday_formatted = (today - timedelta(days=1)).strftime('%m/%d/%Y')

        # --- Read record counts from Excel ---
        counts_dict = {}
        if os.path.exists(RECORD_COUNTS_FILE):
            try:
                df_counts = pd.read_excel(RECORD_COUNTS_FILE)
                df_counts.columns = df_counts.columns.str.strip().str.upper()
                
                alias_col = [c for c in df_counts.columns if 'ALIAS' in c or 'JOB' in c or 'NAME' in c][0]
                count_col = [c for c in df_counts.columns if 'COUNT' in c or 'RECORD' in c or 'VOLUME' in c][0]
                
                counts_dict = dict(zip(df_counts[alias_col].astype(str).str.strip(), df_counts[count_col]))
            except Exception as ex:
                print(f"Warning: Could not parse record counts from Excel: {ex}")
        else:
            print(f"Warning: {RECORD_COUNTS_FILE} not found. Defaulting record counts to N/A.")

        # --- Helper function to grab count safely & cleanly handle formatting ---
        def get_record_count(alias):
            raw_val = counts_dict.get(alias, "")
            if pd.isna(raw_val) or raw_val == "" or raw_val is None:
                return ""
            try:
                # Convert float entries (like 112234.0) into pure integers
                return str(int(float(raw_val)))
            except (ValueError, TypeError):
                return str(raw_val)

        # --- Define your jobs, their baseline status, and metadata mappings ---
        upstream_config = [
            {"file": "NVACY.P9990.EXTR.ADMRPTS.PPLUS ", "interface": "Concord Vantage (PSTP2140)", "alias": "PSTP2140", "default_status": "Received"},
            {"file": "NVACY.P9990.EXTR.VTRAIL ", "interface": "Concord Vantage Trails (PSTP2150) ", "alias": "PSTP2150", "default_status": "Received"},
            {"file": "FLLECCN.LIN2JP.CA7VTG.COMRPTS.PPLUS", "interface": "Fort Wayne Vantage (PSTP2200) ", "alias": "PSTP2200", "default_status": "Received"},
            {"file": "FLLECCN.LIN2JP.CA7VTG.ABCTRG ", "interface": "Fort Wayne Vantage Trails (PSTP2135) ", "alias": "PSTP2135", "default_status": "Received"},
            {"file": "NV1LF.PRDAADN.ADMRPTS.PPLUS ", "interface": "Fusion Vantage  (PSTP2210) ", "alias": "PSTP2210", "default_status": "Received"},
            {"file": "NV1LF.PRDAADN.VDR.SUPEXT.REQABC ", "interface": "Fusion Vantage Trails (PSTP2130) ", "alias": "PSTP2130", "default_status": "Received"},
            {"file": "IADM.IFCX9300.PP.TRX ", "interface": "Life-Comm Flex (PSTP2160) ", "alias": "PSTP2160", "default_status": "Received"},
            {"file": "IADM.IACX9300.PP.TRX ", "interface": "Life-Comm  AHL (PSTP2180) ", "alias": "PSTP2180", "default_status": "Received"},
            {"file": "NPSIT.DSN.PSITOP01.AEASALES.FILE ", "interface": "AEA Bonus Sales (PSTP5109) ", "alias": "PSTP5109", "default_status": "NA"},
            {"file": "NPSIT.DSN.PSITOP01.AEAENTY.FILE ", "interface": "AEA Bonus Entity (PSTP5112) ", "alias": "PSTP5112", "default_status": "NA"}
        ]

        # Process lists dynamically based on the empty/missing condition
        file_names = []
        interface_names = []
        file_statuses = []
        record_counts = []

        for item in upstream_config:
            alias = item["alias"]
            formatted_count = get_record_count(alias)
            
            if formatted_count == "":
                final_status = ""
                final_count = ""
            else:
                final_status = item["default_status"]
                final_count = formatted_count

            file_names.append(item["file"])
            interface_names.append(item["interface"])
            file_statuses.append(final_status)
            record_counts.append(final_count)
            
        # 1. Upstream Summary Data Setup
        upstream_data = {
            "File Name": file_names,
            "Interface Name": interface_names,
            "File Status": file_statuses,
            "Record Count": record_counts
        }
        upstream_html = pd.DataFrame(upstream_data).to_html(index=False, border=1, justify='center')

        # 2. Calc Segments HTML Setup
        upstream_data1 = {
            "CALC Segment": ["11 Segments"],
            "Number of Packets": [get_record_count("PSTP3095")],
            "Start Time": [metrics_context['jobs_formatted5']],
            "End Time": [metrics_context['jobs_formatted1']],
            "Duration": [metrics_context['jobs_formatted4']]
        }
        upstream_html1 = pd.DataFrame(upstream_data1).to_html(index=False, border=1, justify='center')

        # 3. Process Execution Matrix Setup
        upstream_data2 = { 
            "Owner": ["DXC", "DXC", "DXC", "Mindtree", "COG", "COG", "DXC", "DXC", "DXC", "Mindtree", "DXC", "COG", "COG", "DXC", "Mindtree", "COG", "COG", "COG", "Mindtree", "Mindtree", "DXC", "DXC", "Mindtree", "DXC", "DXC", "DXC", "DXC", "COG", "DXC", "COG", "DXC", "COG", "COG", "COG", "COG", "DXC", "DXC", "DXC", "DXC", "DXC", "DXC", "DXC", "DXC", "COG", "COG", "COG", "DXC", "DXC"],
            "Process": ["CSA - User extract file (PSTP5092)", "CICS Online down (TAPPS1SD)", "CSA – Full/Delta Policy Producer Extract (PSTP2083/PSTP9050)", "CDW – Full/Delta Policy Producer Extract(PSTP9051/PSTP2084)", "ADW – Full Policy Producer Extract (PSTP9052)", "NBIT – Life writer Extracts (PSTP2081 series A thru G)", "Initialization & Pre cycle image copy job (PSTP0030)", "ADDRESS table update (PSTP3160)", "Purge old PDR records (PSTP6506)", "Director - ACF files (PSTP0038)", "Bridge Jobs (PSTP1000)", "PEP+/Kyriba -ACH files (PSTP2035)", "B2B - Full PPE file with Trailer record (PSTP2088)", "CSA - Full/Delta Producer Demographic Extract (PSTP3155/ PSTP9060)", "CDW - Full/Delta Producer Demographic Extract (PSTP3156/ PSTP9061)", "ADW - Full Producer Demographic Extract (PSTP9062)", "FCRM - Party/Relationship files (PSTP5026)", "Back feed from CDS (PSTP2010)", "ACE/Director(PSTP2115)", "COMM/REV (PSTP2126)", "Fort Wayne - New Business Underwriting (PSTP2220)", "IMPACT Coordinate Cycle input & pre-process (PSTP2330)", "AGMR – Calc job (PSTP3090)", "CALC Balancing (PSTP3095)", "FWV, FV, Concord Vantage - CALC completion (PSTP3097)", "Unload of tables for Lifecomm (PSTP3098)", "FV - PCR & OBA table extract (PSTP3099)", "B2B - non-ACF DTCC extract (PSTP3169)", "BTT/PSOPs- Internal Replacement Report (PSTP3170)", "CSW - AML extract (PSTP5090 series A thru F)", "IMPACT - Pay process job (PSTP4070)", "XPression - Statement files (PSTP4170)", "CDS - Check file (PSTP4180)", "Stockbroker - AV Pay (PSTP4210)", "B2B/ DTCC - ACF files (PSTP3142)", "FWV - ACF files (PSTP3145)", "RPS - ACF files (PSTP3149)", "TRAC Alliance - ACF files (PSTP3147)", "Pay On Demand Processor (PSTP5010)", "All Impact database updates completed (PSTP9601)", "Post cycle backup (PSTP5050)", "Unload of tables for PCMN/AGMR (PSTP5106)", "FWV/FV - ABC Valuation process (PSTP5089/ PSTP5096)", "E2 Ledger file (PSTP4095)", "B2B/DTCC - POS/FAR files (PSTP5080 series A thru E)", "B2B/DTCC - Commission files (PSTP4320)", "CICS Online Up (CICSPPS1)", "Impact cycle balancing report (PSTP5000)"],
            "SLO": ["7:30 PM", "8:00 PM", "8:30 PM", "8:30 PM", "8:30 PM", "9:00 PM", "9:30 PM", "9:45 PM", "9:45 PM", "9:45 PM", "10:00 PM", "10:00 PM", "10:15 PM", "10:30 PM", "10:30 PM", "10:30 PM", "10:30 PM", "1:00 AM", "1:00 AM", "1:00 AM", "1:00 AM", "1:00 AM", "2:30 AM", "3:00 AM", "3:00 AM", "3:00 AM", "3:00 AM", "3:15 AM", "3:30 AM", "4:30 AM", "4:30 AM", "4:45 AM", "5:00 AM", "5:00 AM", "5:00 AM", "5:00 AM", "5:00 AM", "5:00 AM", "5:15 AM", "5:40 AM", "5:40 AM", "5:30 AM", "5:40 AM", "6:15 AM", "6:15 AM", "6:30 AM", "6:45 AM", "7:00 AM"],
            "Process Completion Status" : [
                metrics_context['jobs_formatted_PSTP5092A'], metrics_context['jobs_formatted_TAPPS1SDA']+'Completed', metrics_context['jobs_formatted_PSTP2083A'], metrics_context['jobs_formatted_PSTP9051A'],  
                metrics_context['jobs_formatted_PSTP9052A'], metrics_context['jobs_formatted_PSTP208GA'], metrics_context['jobs_formatted_PSTP0030A'], metrics_context['jobs_formatted_PSTP3160A'], metrics_context['jobs_formatted_PSTP6506A'], metrics_context['jobs_formatted_PSTP0038A'], 
                metrics_context['jobs_formatted_PSTP1000A'], metrics_context['jobs_formatted_PSTP2035A'], metrics_context['jobs_formatted_PSTP2088A'], metrics_context['jobs_formatted_PSTP3155A'],  
                metrics_context['jobs_formatted_PSTP3156A'], metrics_context['jobs_formatted_PSTP9062A'], metrics_context['jobs_formatted_PSTP5026A'], metrics_context['jobs_formatted_PSTP2010A'], metrics_context['jobs_formatted_PSTP2115A'], metrics_context['jobs_formatted_PSTP2126A'], 
                metrics_context['jobs_formatted_PSTP2220A'], metrics_context['jobs_formatted_PSTP2330A'], metrics_context['jobs_formatted_PSTP3090A'], metrics_context['jobs_formatted_PSTP3095A'], metrics_context['jobs_formatted_PSTP3097A'], metrics_context['jobs_formatted_PSTP3098A'], 
                metrics_context['jobs_formatted_PSTP3099A'], metrics_context['jobs_formatted_PSTP3169A'], metrics_context['jobs_formatted_PSTP3170A'], metrics_context['jobs_formatted_PSTP509FA'], metrics_context['jobs_formatted_PSTP4070A'], metrics_context['jobs_formatted_PSTP4170A'], 
                metrics_context['jobs_formatted_PSTP4180A'], metrics_context['jobs_formatted_PSTP4210A'], metrics_context['jobs_formatted_PSTP3142A'], metrics_context['jobs_formatted_PSTP3145A'], metrics_context['jobs_formatted_PSTP3149A'], metrics_context['jobs_formatted_PSTP3147A'], 
                metrics_context['jobs_formatted_PSTP5010A'], metrics_context['jobs_formatted_PSTP9601A'], metrics_context['jobs_formatted_PSTP5050A'], metrics_context['jobs_formatted_PSTP5106A'], metrics_context['jobs_formatted_PSTP5089A'], metrics_context['jobs_formatted_PSTP4095A'], 
                metrics_context['jobs_formatted_PSTP508EA'], metrics_context['jobs_formatted_PSTP4320A'], metrics_context['jobs_formatted_CICSPPS1A'], metrics_context['jobs_formatted_PSTP5000A']
            ],
            "Process Completion Time" : [
                metrics_context['jobs_formatted_PSTP5092'], metrics_context['jobs_formatted_TAPPS1SD']+'09:00 PM ET', metrics_context['jobs_formatted_PSTP2083'], metrics_context['jobs_formatted_PSTP9051'],  
                metrics_context['jobs_formatted_PSTP9052'], metrics_context['jobs_formatted_PSTP208G'], metrics_context['jobs_formatted_PSTP0030'], metrics_context['jobs_formatted_PSTP3160'], metrics_context['jobs_formatted_PSTP6506'], metrics_context['jobs_formatted_PSTP0038'], 
                metrics_context['jobs_formatted_PSTP1000'], metrics_context['jobs_formatted_P0235' if 'jobs_formatted_P0235' in metrics_context else 'jobs_formatted_PSTP2035'], metrics_context['jobs_formatted_PSTP2088'], metrics_context['jobs_formatted_PSTP3155'], 
                metrics_context['jobs_formatted_PSTP3156'], metrics_context['jobs_formatted_PSTP9062'], metrics_context['jobs_formatted_PSTP5026'], metrics_context['jobs_formatted_PSTP2010'], metrics_context['jobs_formatted_PSTP2115'], metrics_context['jobs_formatted_PSTP2126'], 
                metrics_context['jobs_formatted_PSTP2220'], metrics_context['jobs_formatted_PSTP2330'], metrics_context['jobs_formatted_PSTP3090'], metrics_context['jobs_formatted_PSTP3095'], metrics_context['jobs_formatted_PSTP3097'], metrics_context['jobs_formatted_PSTP3098'], 
                metrics_context['jobs_formatted_PSTP3099'], metrics_context['jobs_formatted_PSTP3169'], metrics_context['jobs_formatted_PSTP3170'], metrics_context['jobs_formatted_PSTP509F'], metrics_context['jobs_formatted_PSTP4070'], metrics_context['jobs_formatted_PSTP4170'], 
                metrics_context['jobs_formatted_PSTP4180'], metrics_context['jobs_formatted_PSTP4210'], metrics_context['jobs_formatted_PSTP3142'], metrics_context['jobs_formatted_PSTP3145'], metrics_context['jobs_formatted_PSTP3149'], metrics_context['jobs_formatted_PSTP3147'], 
                metrics_context['jobs_formatted_PSTP5010'], metrics_context['jobs_formatted_PSTP9601'], metrics_context['jobs_formatted_PSTP5050'], metrics_context['jobs_formatted_PSTP5106'], metrics_context['jobs_formatted_PSTP5089'], 
                metrics_context['jobs_formatted_PSTP4095'], metrics_context['jobs_formatted_PSTP508E'], metrics_context['jobs_formatted_PSTP4320'], metrics_context['jobs_formatted_CICSPPS1'], metrics_context['jobs_formatted_PSTP5000']
            ]
        }
        
        df_owner = pd.DataFrame(upstream_data2)
        owner_html = df_owner.to_html(index=False, border=1, justify='center')

        # 4. Outlook Dispatch
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = 'rakesh.rathod@dxc.com'
        mail.CC = 'harshitha.cherukuri@dxc.com; s.sampara@dxc.com; khushboo.i@dxc.com; rakesh.rathod@lfg.com; rakesh.rathod@dxc.com; meghana.km@dxc.com; abhishek.rajput2@dxc.com; smandloi3@dxc.com; sagrawal69@dxc.com; nsharma37@dxc.com;'
        mail.Subject = f'TEST-IMPACT PROD Cycle Status Notification {yesterday_formatted} - Second Notification (secure)'
        
        mail.HTMLBody = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Calibri, sans-serif; font-size: 15px; color: #000000; line-height: 1.4;}}
                table {{ border-collapse: collapse; font-family: Calibri, sans-serif; width: auto; margin-bottom: 20px 0; }}
                th {{ background-color: #663399; color: white; padding: 11px 16px; border: 1.5px solid #003366; text-transform: uppercase; font-size: 13px; font-family: Calibri, sans-serif; text-align: left; white-space: nowrap; }}
                td {{ padding: 09px 16px; text-align: left; border: 1.5px solid #000000;; font-size: 14px; font-family: Calibri, sans-serif; }}
                tr:nth-child(even) {{ background-color: #f2f2f2; }}
                .section-header {{ font-weight: bold; font-size: 17px; color: #333; margin-top: 20px; }}
            </style>
        </head>
        <body>
            <p>IMPACT cycle is in <strong>Progress</strong>. FW Vantage (daily transaction) feed received at <strong>{metrics_context['jobs_formatted3']}</strong>.</p>
            <p>CALC jobs processed <strong>{get_record_count("PSTP3095")}</strong> transactions and ran for <strong>{metrics_context['jobs_formatted4']}</strong> and completed at <strong>{metrics_context['jobs_formatted1']}</strong>.</p>
            <p>Pay process job have completed at <strong>{metrics_context['jobs_formatted']}</strong> and ran for <strong>{metrics_context['jobs_formatted0']}</strong>. </p>
            <p> IMPACT cycle for <strong>{yesterday_formatted}</strong> started at <strong>{metrics_context['jobs_formatted2']}</strong> ,CICS region came down at <strong>09:00 PM ET</strong> </p>
            <p> Note, the following feeds were received after the critical path started and will be included in the next cycle: None.</p>
            
            <u>Upstream Files Status</u><br><br>
            {upstream_html}
            
            <p>CALC Packets Processed <strong>{get_record_count("PSTP3095")}</strong> packets.</p>
            {upstream_html1}<br><br>
            {owner_html}
            
           <p class="section-header">NEXT STATUS: At 06:00 AM ET</p>
        </body>
        </html>
        """
        mail.Send()
        print("SUCCESS: Email sent successfully!")
    except Exception as e:
        print(f"FAILED to send email: {e}")

# ==========================================
# 6. MAIN EXECUTION
# ==========================================
if __name__ == "__main__":
    cleaned_data = organize_input_data()
    
    if cleaned_data is not None:
        final_matched_data = process_and_match(cleaned_data)
        metrics_context = prepare_job_metrics()
        
        if final_matched_data is not None and not final_matched_data.empty:
            send_outlook_email(final_matched_data, metrics_context)
        else:
            print("No matching data found. Email not sent.")