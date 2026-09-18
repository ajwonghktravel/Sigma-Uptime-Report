import argparse
import glob
import os
import re
from datetime import datetime
from pathlib import Path
import dateparser
import numpy as np
import pandas as pd
data_folder = Path.cwd() / "data"
output_dir = Path("reports") / datetime.now().strftime('%Y-%m-%d')
output_dir.mkdir(parents=True, exist_ok=True)
parser = argparse.ArgumentParser(description="Generate quarterly reports for CRM and Charge Data.")
parser.add_argument("--report", choices=["CRM", "ChargeData", "Both"], default="Both", help="Specify which report to generate.")
parser.add_argument("--quarter", type=int, choices=[1, 2, 3, 4], help="Specify the quarter (1-4) for the report.")
args = parser.parse_args()
def filename_handling(report_type = None):
    # Get the current working directory
    cwd = data_folder
    # Find all Excel files in the current directory
    excel_files = list(cwd.glob('*.xlsx'))
    if not excel_files:
        print("No Excel files found in the current directory.")
        return None, None, None
    # Sort files by modification time (newest first)
    excel_files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    for file in excel_files:
        if file.name.startswith('~$'):
            continue  # Skip temporary files
        if re.search(report_type, file.name, re.IGNORECASE):
                latest_file = file
                break
    else:
        print("No matching Excel files found.")
        return None, None, None
    return latest_file, latest_file.stem, datetime.now().strftime('%Y-%m-%d')
downtime = filename_handling(report_type="Downtime")
uptime = filename_handling(report_type="Uptime")
print(f"Latest Downtime Report: {downtime[0]}")
print(f"Latest Uptime Report: {uptime[0]}")
charge_sessions = filename_handling(report_type="Charge Sessions")
print(f"Latest Charge Sessions Report: {charge_sessions[0]}")
def header_handling(file_path, sheet_name, expected_columns):
    real_header_row = None
    for i in range(10):  # Check the first 10 rows
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=i)
        if all(col in df.columns for col in expected_columns):
            real_header_row = i
            break
    if real_header_row is None:
        raise ValueError(f"Could not find the correct header row in the {sheet_name}.")
    return real_header_row
downtime_header = header_handling(
    file_path=downtime[0], 
    sheet_name='Downtime Report', 
    expected_columns=['charger_id', 'downtime_start_time']
)
sigma_downtime_df = pd.read_excel(downtime[0], sheet_name='Downtime Report', header=downtime_header)
print(sigma_downtime_df.head())    
uptime_header = header_handling(
    file_path=uptime[0], 
    sheet_name='Uptime Report', 
    expected_columns=['charger_id', 'uptime_start_time']
)
sigma_uptime_df = pd.read_excel(uptime[0], sheet_name='Uptime Report', header=uptime_header)
print(sigma_uptime_df.head())
charge_sessions_header = header_handling(
    file_path=charge_sessions[0], 
    sheet_name='Charge Sessions Report', 
    expected_columns=['charger_id', 'session_start_time']
)
sigma_chargesessions = pd.read_excel(charge_sessions[0], sheet_name='Charge Sessions Report', header=charge_sessions_header)
print(sigma_chargesessions.head())
if args.quarter:
    quarter = args.quarter
else:
    current_month = datetime.now().month
    if current_month in [1, 2, 3]:
        quarter = 1
    elif current_month in [4, 5, 6]:
        quarter = 2
    elif current_month in [7, 8, 9]:
        quarter = 3
    else:
        quarter = 4
if quarter == 1:
    # Do something
    quarter_start = datetime(datetime.now().year, 1, 1, 0, 0 , 0)
    quarter_end = datetime(datetime.now().year, 3, 31, 11, 59, 59)
elif quarter == 2:
    quarter_start = datetime(datetime.now().year, 4, 1, 0, 0 , 0)
    quarter_end = datetime(datetime.now().year, 6, 30, 11, 59, 59)
elif quarter == 3: 
    quarter_start = datetime(datetime.now().year, 7, 1, 0, 0 , 0)
    quarter_end = datetime(datetime.now().year, 9, 30, 11, 59, 59)
else:
    quarter_start = datetime(datetime.now().year, 10, 1, 0, 0 , 0)
    quarter_end = datetime(datetime.now().year, 12, 31, 11, 59, 59)
def quarterly_report_CRM(quarter=quarter):
    filtered_downtime = sigma_downtime_df[(sigma_downtime_df['calendar_quarter'] == f"Q{quarter}")]
    filtered_downtime['downtime_start_time'] = pd.to_datetime(filtered_downtime['downtime_start_time'], errors='coerce')
    filtered_downtime['downtime_end_time'] = pd.to_datetime(filtered_downtime['downtime_end_time'], errors='coerce')
    filtered_uptime = sigma_uptime_df[(sigma_uptime_df['calendar_quarter'] == f"Q{quarter}")]
    quarter_total_sums = filtered_downtime.pivot_table(index='charger_id', columns='calendar_quarter', values='downtime_seconds', aggfunc='sum')
    quarter_total_sums = quarter_total_sums.rename(columns={f"Q{quarter}": "total_downtime_seconds"})
    quarter_total_sums['System S/N'] = filtered_downtime.groupby('charger_id')['serial_number'].first()
    quarter_total_sums['earliest_start_time'] = (
        filtered_uptime
        .loc[filtered_uptime['uptime_start_time']<=quarter_end]
        .groupby('charger_id')['uptime_start_time']
        .min()
        .clip(lower=quarter_start)
    )   
    #pull in earliest start time for uptime report
    #quarter_total_sums = quarter_total_sums.reset_index()
    quarter_total_sums['total_quarter_time'] = (quarter_end - quarter_total_sums['earliest_start_time']).dt.total_seconds()
    uptime_percentage = 100-((quarter_total_sums['total_downtime_seconds'] / quarter_total_sums['total_quarter_time']) * 100)
    quarter_total_sums['uptime_percentage'] = uptime_percentage
    quarter_total_sums['median_downtime'] = filtered_downtime.groupby('charger_id')['downtime_seconds'].median()
    quarter_total_sums['mean_downtime'] = filtered_downtime.groupby('charger_id')['downtime_seconds'].mean()
    quarter_total_sums['minimum_downtime_duration'] = filtered_downtime.groupby('charger_id')['downtime_seconds'].min()
    quarter_total_sums['maximum_downtime_duration'] = filtered_downtime.groupby('charger_id')['downtime_seconds'].max()
    quarter_total_sums['downtime_causes'] = filtered_downtime.groupby('charger_id')['Downtime Types'].unique().astype(object).apply(lambda x: ", ".join(x))
    less_than_97 = quarter_total_sums[quarter_total_sums['uptime_percentage'] < 97].reset_index()
    sorted_by_reason = filtered_downtime.groupby(['charger_id', 'Downtime Exclusion Types'])['downtime_seconds'].sum()
    quarter_total_sums = quarter_total_sums.merge(sorted_by_reason, on='charger_id', how='left')
    quarter_total_sums = quarter_total_sums.rename(columns={'downtime_seconds': 'excluded_downtime_seconds'}).fillna(0)
    quarter_total_sums['excluded_downtime_reasons'] = (
    filtered_downtime
    .groupby('charger_id')['Downtime Exclusion Types']
    .apply(lambda x: list(x.dropna().unique()) if x.dropna().any() else 'No Exclusions')
)
    quarter_total_sums['Total Uptime percentage'] = 100 - ((quarter_total_sums['total_downtime_seconds'] - quarter_total_sums['excluded_downtime_seconds']) / quarter_total_sums['total_quarter_time']) * 100
    quarter_total_sums.reset_index(inplace=True)
    quarter_total_sums.to_csv(output_dir / f"quarterly_report_CRM_{datetime.now().strftime('%Y%m%d%H%M')}.csv", index=False)
    #print(less_than_97)
    print(quarter_total_sums)
    print(sorted_by_reason)
def sessions_report_chargedata(quarter=quarter):
    filtered_sessions = sigma_chargesessions[(sigma_chargesessions['calendar_quarter'] == f"Q{quarter}")]
    #print(filtered_sessions['charge_session_status'].unique())
    filtered_sessions['session_start_time'] = pd.to_datetime(filtered_sessions['session_start_time'], errors='coerce')
    filtered_sessions['session_end_time'] = pd.to_datetime(filtered_sessions['session_end_time'], errors='coerce')
    filtered_sessions['session_duration_seconds'] = (filtered_sessions['session_end_time'] - filtered_sessions['session_start_time']).dt.total_seconds()
    charger_hardware_errorlist = ['OverCurrentFailure', 'GroundFailure', 'PowerSwitchFailure', 'UnderVoltage']
    error_pattern = '|'.join(charger_hardware_errorlist)
    filtered_sessions['charger_hardware_error'] = filtered_sessions['session_errors'].str.contains(error_pattern, na=False)
    filtered_sessions['other_error'] = filtered_sessions['session_errors'].apply(lambda x: not pd.isna(x) and not any(error in x for error in charger_hardware_errorlist))
    filtered_sessions['charge_session_status'] = filtered_sessions.apply(lambda row: 'ChargerHardwareError' if row['charger_hardware_error'] else ('OtherError' if row['other_error'] else 'Successful'), axis=1)
    chargerid_to_sn = filtered_sessions.groupby('charger_id')['charger_serial_number'].first().to_dict()
    #print(filtered_sessions['charge_session_status'].unique())
    session_summary = filtered_sessions.groupby('charger_id').agg(
        total_sessions=('charge_event_id', 'count'),
        # total failed sessions is either hardware or other)
        total_failed_sessions=('charge_session_status', lambda x: x.isin(['ChargerHardwareError', 'OtherError']).sum()),
        charger_hardware_failures=('session_errors', lambda val: val.isin(charger_hardware_errorlist).sum()),
        other_failures = ('session_errors', lambda val: val.apply(lambda x: x == 'OtherError' or x == ()).sum())
    ).reset_index()
    session_summary['serial_number'] = session_summary['charger_id'].map(chargerid_to_sn)
    chars_to_strip = " ,0123456789"
    session_error_types = filtered_sessions.copy()
    session_error_types['session_errors'] = session_error_types['session_errors'].str.strip(chars_to_strip)
    session_error_types = pd.crosstab(
        index=session_error_types['charger_id'], 
        columns=session_error_types['session_errors']
    ).reset_index()    
    session_error_types['serial_number'] = session_error_types['charger_id'].map(chargerid_to_sn)
    session_error_types.to_csv(output_dir / f"session_error_types_{datetime.now().strftime('%Y%m%d%H%M')}.csv", index=False)
    session_summary.to_csv(output_dir / f"session_summary_{datetime.now().strftime('%Y%m%d%H%M')}.csv", index=False)
    print(session_summary)


if __name__ == "__main__":
    args = parser.parse_args()
    if args.report == "CRM":
        quarterly_report_CRM(quarter=args.quarter)
    elif args.report == "ChargeData":
        sessions_report_chargedata(quarter=args.quarter)
    elif args.report == "Both":
        quarterly_report_CRM(quarter=args.quarter)
        sessions_report_chargedata(quarter=args.quarter)