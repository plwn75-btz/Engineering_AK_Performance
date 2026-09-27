"""
Analytic Evaluation Data Processor: Rebaseline Plan vs. Actual Revision Lifecycle Engine
Integrates:
1. MM-ASK-1A-GEN01-ENG-MDR-0001_B1- Cut off 18-Sep-26.xlsx (5 tabs)
2. ASK Technical Document Log Report - 18 Sep 26.xlsx
3. ASK Vendor Document Log Report - 18 Sep 26.xls

Generates:
- analytics_data.js for index_analytics.html

Contractual SLA & Delay Attribution Methodology:
- Client Review Contractual SLA: <= 10 Working Days (Mon-Fri)
- Contractor Planned Incorporation Allowance: max(0, Plan Cycle WD - 10 WD)
- Contractor Allowable Extended Cycle: Actual Client Review WD + Contractor Planned Incorporation Allowance
- Delay Cause Classification:
  1. IMPROPER_PLAN: Plan Cycle WD <= 10 WD (Contractor allocated <= 0 days for incorporation)
  2. LATE_CLIENT: Actual Cycle > Plan Cycle BUT Actual Cycle <= Allowable Cycle (delay 100% from Client taking > 10 WD)
  3. LATE_CONTRACTOR: Client passed SLA (<= 10 WD), but Contractor took longer than planned incorporation allowance
  4. COMPOUND_DELAY: Client took > 10 WD AND Contractor took > planned incorporation allowance (quantified by party)
  5. MET_PLAN: Actual Cycle <= Plan Cycle
- Macro Justification:
  If (Late Contractor + Compound Delay) / Total Overdue > 50%, confirms systematically improper/over-optimistic Contractor scheduling.
"""

import os
import re
import sys
import glob
import json
from pathlib import Path
from datetime import datetime, timedelta, date
import openpyxl
import pandas as pd

def calc_working_days(d1, d2):
    if d1 is None or d2 is None:
        return None
    if d2 < d1:
        return 0
    cur = d1 + timedelta(days=1)
    work_days = 0
    while cur <= d2:
        if cur.weekday() < 5:  # Monday to Friday (0-4)
            work_days += 1
        cur += timedelta(days=1)
    return work_days

def parse_date_obj(v):
    if v is None:
        return None
    if isinstance(v, (datetime, date)):
        return v if isinstance(v, date) else v.date()
    s = str(v).strip()
    if s.upper() in ['N/A', '-', '', 'NONE', 'NULL']:
        return None
    for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%b-%Y']:
        try:
            return datetime.strptime(s, fmt).date()
        except:
            pass
    return None

def process_analytics(mdr_path=None, tech_log_path=None, vendor_log_path=None):
    # Auto-discover files if not provided
    if not mdr_path:
        files = sorted([f for f in glob.glob("*MDR*.xlsx") if not os.path.basename(f).startswith("~$")], reverse=True)
        if not files:
            print("Error: No MDR Excel file found.")
            return False
        mdr_path = files[0]

    if not tech_log_path:
        files = sorted([f for f in glob.glob("ASK Technical Document Log Report*.xlsx") if not os.path.basename(f).startswith("~$")], reverse=True)
        if not files:
            print("Error: No Technical Document Log Excel file found.")
            return False
        tech_log_path = files[0]

    print(f"[Analytics] Ingesting Technical Log: {tech_log_path}...")
    wb_log = openpyxl.load_workbook(tech_log_path, data_only=True)
    sheet_log = wb_log['Page 1']
    log_vals = sheet_log.values
    for _ in range(9): next(log_vals)
    headers_log = next(log_vals)
    df_log = pd.DataFrame(log_vals, columns=headers_log)
    j_log = df_log[df_log['Submitted By'].astype(str).str.contains('zm167dc@jutal.com', case=False, na=False)].copy()

    doc_log_map = {}
    for _, row in j_log.iterrows():
        d_no = str(row['Document No.'] or '').strip()
        if not d_no: continue
        if d_no not in doc_log_map:
            doc_log_map[d_no] = []
        
        in_str = str(row['Actual Submit In'] or '').strip()
        out_str = str(row['Actual Submit Out'] or '').strip()
        in_dt, out_dt = None, None
        if in_str:
            try: in_dt = datetime.strptime(in_str, '%d-%b-%Y %H:%M:%S')
            except: pass
        if out_str:
            try: out_dt = datetime.strptime(out_str, '%d-%b-%Y %H:%M:%S')
            except: pass
        
        dur = calc_working_days(in_dt.date(), out_dt.date()) if (in_dt and out_dt) else None
        doc_log_map[d_no].append({
            'rev': str(row['Rev.'] or '').strip(),
            'submittal_id': str(row['Submittal ID'] or '').strip(),
            'doc_category': str(row.get('Document Category') or '').strip(),
            'submit_in_raw': in_str,
            'submit_out_raw': out_str,
            'in_dt': in_dt.strftime('%Y-%m-%d') if in_dt else None,
            'out_dt': out_dt.strftime('%Y-%m-%d') if out_dt else None,
            'in_date': in_dt.date() if in_dt else None,
            'out_date': out_dt.date() if out_dt else None,
            'review_days': dur,
            'status_code': str(row['Return Status Code'] or '').strip(),
            'status_name': str(row['Return Status Name'] or '').strip()
        })

    print(f"[Analytics] Indexed {len(doc_log_map)} documents from Technical Log.")

    print(f"[Analytics] Ingesting MDR: {mdr_path}...")
    wb_mdr = openpyxl.load_workbook(mdr_path, data_only=True)
    sheets = ['WP01 TOPSIDE', 'WP01 JACKET', 'WP02 TOPSIDE', 'WP02 JACKET', 'PRO.ENGINEERING_MR TBE']

    transitions = []
    doc_records = []

    for sname in sheets:
        sheet = wb_mdr[sname]
        header_row = 11 if sname == 'PRO.ENGINEERING_MR TBE' else 13
        
        for r in range(header_row + 1, sheet.max_row + 1):
            doc_no_val = sheet.cell(row=r, column=9).value
            if not doc_no_val: continue
            doc_no = str(doc_no_val).strip()
            if not doc_no.startswith('MM-'): continue
            
            title = str(sheet.cell(row=r, column=10).value or '').strip()
            disc = str(sheet.cell(row=r, column=4).value or '').strip()
            area = str(sheet.cell(row=r, column=3).value or '').strip()
            doc_type = str(sheet.cell(row=r, column=5).value or '').strip()
            doc_class = str(sheet.cell(row=r, column=11).value or '').strip()
            
            # Rebaseline Plan Dates
            ifr_reb = parse_date_obj(sheet.cell(row=r, column=23).value)
            ifa_reb = parse_date_obj(sheet.cell(row=r, column=29).value)
            afc_reb = parse_date_obj(sheet.cell(row=r, column=35).value)
            ap_reb = parse_date_obj(sheet.cell(row=r, column=41).value)

            # Original Plan Dates
            ifr_orig = parse_date_obj(sheet.cell(row=r, column=22).value)
            ifa_orig = parse_date_obj(sheet.cell(row=r, column=28).value)
            afc_orig = parse_date_obj(sheet.cell(row=r, column=34).value)
            ap_orig = parse_date_obj(sheet.cell(row=r, column=40).value)

            is_ifa_skipped = (sheet.cell(row=r, column=29).value is not None and str(sheet.cell(row=r, column=29).value).strip().upper() in ['N/A', '-'])
            
            logs = doc_log_map.get(doc_no, [])
            logs.sort(key=lambda x: x['in_dt'] or '9999-99-99')
            
            # Find submittals matching revision families
            ifr_log = next((l for l in logs if l['rev'].upper().startswith('A')), None)
            ifa_log = next((l for l in logs if l['rev'].upper().startswith('B')), None)
            afc_log = next((l for l in logs if l['rev'].upper().startswith('C')), None)
            
            if not ifr_log and len(logs) > 0: ifr_log = logs[0]
            if not ifa_log and not is_ifa_skipped and len(logs) > 1: ifa_log = logs[1]
            if not afc_log and len(logs) > 2: afc_log = logs[2]

            # AP Closeout log
            ap_log = None
            if afc_log and afc_log['out_dt']:
                after_afc = [l for l in logs if l['in_dt'] and l['in_dt'] > afc_log['out_dt']]
                if after_afc:
                    ap_log = after_afc[0]
                elif 'APPR' in (afc_log['status_code'] or '').upper():
                    ap_log = afc_log

            # Determine Document Category
            cat = None
            for l in logs:
                if l.get('doc_category'): cat = l['doc_category']; break
            if not cat:
                if any(x in doc_no for x in ['-DWG-', '-PID-', '-PFD-', '-UFD-', '-LAY-', '-SCH-', '-SLD-', '-ISO-']) or doc_type in ['DWG', 'PID', 'PFD', 'UFD', 'LAY', 'SCH', 'SLD', 'ISO', 'GA']:
                    cat = 'Part C - Engineering Dwg'
                else:
                    cat = 'Part B - Engineering Doc'
            clean_category = 'Part C - Engineering Dwg' if 'Part C' in cat else 'Part B - Engineering Doc'

            # Define potential transitions for this document
            candidate_transitions = []
            
            # 1. IFR -> IFA (if IFA is planned and not skipped)
            if not is_ifa_skipped and ifr_reb and ifa_reb:
                candidate_transitions.append({
                    'transition_type': 'IFR -> IFA',
                    'stage_from': 'IFR',
                    'stage_to': 'IFA',
                    'plan_from': ifr_reb,
                    'plan_to': ifa_reb,
                    'log_from': ifr_log,
                    'log_to': ifa_log
                })
            # 2. IFR -> AFC (Fast-Track Corridor, when IFA is skipped)
            elif is_ifa_skipped and ifr_reb and afc_reb:
                candidate_transitions.append({
                    'transition_type': 'IFR -> AFC (Fast-Track)',
                    'stage_from': 'IFR',
                    'stage_to': 'AFC',
                    'plan_from': ifr_reb,
                    'plan_to': afc_reb,
                    'log_from': ifr_log,
                    'log_to': afc_log
                })

            # 3. IFA -> AFC (for documents with IFA)
            if not is_ifa_skipped and ifa_reb and afc_reb:
                candidate_transitions.append({
                    'transition_type': 'IFA -> AFC',
                    'stage_from': 'IFA',
                    'stage_to': 'AFC',
                    'plan_from': ifa_reb,
                    'plan_to': afc_reb,
                    'log_from': ifa_log,
                    'log_to': afc_log
                })

            # 4. AFC -> AP Closeout
            if afc_reb and ap_reb:
                candidate_transitions.append({
                    'transition_type': 'AFC -> AP',
                    'stage_from': 'AFC',
                    'stage_to': 'AP',
                    'plan_from': afc_reb,
                    'plan_to': ap_reb,
                    'log_from': afc_log,
                    'log_to': ap_log
                })

            # Evaluate each transition
            for ct in candidate_transitions:
                p_from = ct['plan_from']
                p_to = ct['plan_to']
                p_wd = calc_working_days(p_from, p_to)
                p_incorp_allowance = max(0, p_wd - 10) if p_wd is not None else 0

                # Plan Feasibility
                if p_wd is None:
                    p_feasibility = 'UNKNOWN'
                elif p_wd <= 10:
                    p_feasibility = 'DEFECTIVE_IMPROPER'
                elif p_wd <= 15:
                    p_feasibility = 'TIGHT_AGGRESSIVE'
                else:
                    p_feasibility = 'REALISTIC_EFFECTIVE'

                lf = ct['log_from']
                lt = ct['log_to']

                has_actual = bool(lf and lt and lf.get('in_date') and lt.get('in_date'))
                
                act_wd = None
                client_rev_wd = None
                cont_inc_wd = None
                cycle_variance = None
                allowable_cycle = None
                client_sla_status = 'N/A'
                client_over_wd = 0
                cont_over_wd = 0
                cause_code = 0
                cause_key = 'PENDING'
                cause_label = 'Pending Actual Execution'

                if has_actual:
                    dt_in_from = lf['in_date']
                    dt_out_from = lf['out_date']
                    dt_in_to = lt['in_date']

                    act_wd = calc_working_days(dt_in_from, dt_in_to)
                    client_rev_wd = calc_working_days(dt_in_from, dt_out_from) if dt_out_from else None
                    cont_inc_wd = calc_working_days(dt_out_from, dt_in_to) if dt_out_from else None

                    # Client review SLA check (Contractual: <= 10 WD)
                    if client_rev_wd is not None:
                        if client_rev_wd <= 10:
                            client_sla_status = 'PASS'
                            client_over_wd = 0
                        else:
                            client_sla_status = 'OVER'
                            client_over_wd = client_rev_wd - 10

                    # Contractor's allowable extended cycle under User's Rule:
                    # If Client took > 10 WD, Contractor maintains their full planned incorporation allowance!
                    effective_client_wd = client_rev_wd if client_rev_wd is not None else 10
                    allowable_cycle = effective_client_wd + p_incorp_allowance
                    cycle_variance = (act_wd - p_wd) if (act_wd is not None and p_wd is not None) else 0

                    if cycle_variance <= 0:
                        cause_key = 'MET_PLAN'
                        cause_label = 'Met Rebaseline Plan'
                        cause_code = 0
                        client_over_wd = 0
                        cont_over_wd = 0
                    else:
                        # Delayed / Overdue cycle
                        if p_wd is not None and p_wd <= 10:
                            # 1. Improperly planned: Contractor planned <= 10 WD, leaving zero/negative days for incorp
                            cause_key = 'IMPROPER_PLAN'
                            cause_label = '(1) Improperly Planned'
                            cause_code = 1
                            client_over_wd = max(0, (client_rev_wd or 10) - 10)
                            cont_over_wd = max(0, cycle_variance - client_over_wd)
                        elif allowable_cycle is not None and act_wd <= allowable_cycle:
                            # 2. Contractor submitted within allowable extended cycle!
                            # Overdue against original plan is 100% caused by Client taking > 10 WD
                            cause_key = 'LATE_CLIENT'
                            cause_label = '(3) Late Client Review'
                            cause_code = 3
                            client_over_wd = max(0, (client_rev_wd or 10) - 10)
                            cont_over_wd = 0
                        elif client_sla_status == 'PASS':
                            # 3. Client passed 10 WD SLA, so delay is 100% Contractor late incorporation
                            cause_key = 'LATE_CONTRACTOR'
                            cause_label = '(2) Late Contractor Incorp'
                            cause_code = 2
                            client_over_wd = 0
                            cont_over_wd = max(0, (cont_inc_wd or act_wd) - p_incorp_allowance)
                        else:
                            # 4. Compound / Shared Delay: Both Client exceeded 10 WD AND Contractor exceeded allowance
                            cause_key = 'COMPOUND_DELAY'
                            cause_label = '(4) Compound Delay (Client + Cont.)'
                            cause_code = 4
                            client_over_wd = max(0, (client_rev_wd or 10) - 10)
                            cont_over_wd = max(0, (cont_inc_wd or 0) - p_incorp_allowance)

                transitions.append({
                    'id': f"{doc_no}__{ct['transition_type'].replace(' ', '_').replace('->', 'to')}",
                    'doc_no': doc_no,
                    'title': title,
                    'work_package': sname,
                    'discipline': disc,
                    'area': area,
                    'doc_type': doc_type,
                    'doc_class': doc_class,
                    'doc_category': clean_category,
                    'transition_type': ct['transition_type'],
                    'stage_from': ct['stage_from'],
                    'stage_to': ct['stage_to'],
                    
                    'plan_from': p_from.strftime('%Y-%m-%d') if p_from else None,
                    'plan_to': p_to.strftime('%Y-%m-%d') if p_to else None,
                    'plan_wd': p_wd,
                    'plan_incorp_allowance_wd': p_incorp_allowance,
                    'plan_feasibility': p_feasibility,
                    
                    'has_actual': has_actual,
                    'rev_from': lf['rev'] if lf else None,
                    'rev_to': lt['rev'] if lt else None,
                    'actual_in_from': lf['in_dt'] if lf else None,
                    'actual_out_from': lf['out_dt'] if lf else None,
                    'actual_in_to': lt['in_dt'] if lt else None,
                    
                    'actual_cycle_wd': act_wd,
                    'client_rev_wd': client_rev_wd,
                    'client_sla_status': client_sla_status,
                    'cont_inc_wd': cont_inc_wd,
                    'allowable_cycle_wd': allowable_cycle,
                    'cycle_variance_wd': cycle_variance,
                    
                    'delay_cause': cause_key,
                    'delay_cause_code': cause_code,
                    'delay_cause_label': cause_label,
                    'client_delay_days': client_over_wd,
                    'contractor_delay_days': cont_over_wd,
                    
                    'submittal_id_from': lf['submittal_id'] if lf else None,
                    'submittal_id_to': lt['submittal_id'] if lt else None,
                    'status_code_from': lf['status_code'] if lf else None,
                    'status_code_to': lt['status_code'] if lt else None
                })

    # Summary Metrics Calculation
    evaluated = [t for t in transitions if t['has_actual'] and t['actual_cycle_wd'] is not None and t['plan_wd'] is not None]
    total_eval = len(evaluated)
    met_plan_list = [t for t in evaluated if t['delay_cause'] == 'MET_PLAN']
    overdue_list = [t for t in evaluated if t['delay_cause'] != 'MET_PLAN']
    
    improper_list = [t for t in overdue_list if t['delay_cause'] == 'IMPROPER_PLAN']
    late_cont_list = [t for t in overdue_list if t['delay_cause'] == 'LATE_CONTRACTOR']
    late_client_list = [t for t in overdue_list if t['delay_cause'] == 'LATE_CLIENT']
    compound_list = [t for t in overdue_list if t['delay_cause'] == 'COMPOUND_DELAY']

    cont_driven_count = len(late_cont_list) + len(compound_list)
    cont_overrun_pct = round((cont_driven_count / len(overdue_list) * 100), 1) if overdue_list else 0.0
    macro_justification_flag = (cont_overrun_pct > 50.0)

    # Client review SLA compliance across all evaluated transitions
    client_evaluated = [t for t in evaluated if t['client_rev_wd'] is not None]
    client_pass = [t for t in client_evaluated if t['client_sla_status'] == 'PASS']
    client_sla_rate = round(len(client_pass) / len(client_evaluated) * 100, 1) if client_evaluated else 0.0

    avg_plan_wd = round(sum(t['plan_wd'] for t in evaluated) / total_eval, 1) if total_eval else 0.0
    avg_actual_wd = round(sum(t['actual_cycle_wd'] for t in evaluated) / total_eval, 1) if total_eval else 0.0
    avg_client_wd = round(sum(t['client_rev_wd'] for t in client_evaluated) / len(client_evaluated), 1) if client_evaluated else 0.0
    
    cont_inc_evaluated = [t for t in evaluated if t['cont_inc_wd'] is not None]
    avg_cont_inc_wd = round(sum(t['cont_inc_wd'] for t in cont_inc_evaluated) / len(cont_inc_evaluated), 1) if cont_inc_evaluated else 0.0

    # Discipline breakdown
    disc_map = {}
    for t in evaluated:
        d = t['discipline'] or 'Other'
        if d not in disc_map:
            disc_map[d] = {'total': 0, 'met': 0, 'improper': 0, 'late_cont': 0, 'late_client': 0, 'compound': 0}
        disc_map[d]['total'] += 1
        c = t['delay_cause']
        if c == 'MET_PLAN': disc_map[d]['met'] += 1
        elif c == 'IMPROPER_PLAN': disc_map[d]['improper'] += 1
        elif c == 'LATE_CONTRACTOR': disc_map[d]['late_cont'] += 1
        elif c == 'LATE_CLIENT': disc_map[d]['late_client'] += 1
        elif c == 'COMPOUND_DELAY': disc_map[d]['compound'] += 1

    # Transition type breakdown
    trans_map = {}
    for t in evaluated:
        tt = t['transition_type']
        if tt not in trans_map:
            trans_map[tt] = {'total': 0, 'met': 0, 'improper': 0, 'late_cont': 0, 'late_client': 0, 'compound': 0}
        trans_map[tt]['total'] += 1
        c = t['delay_cause']
        if c == 'MET_PLAN': trans_map[tt]['met'] += 1
        elif c == 'IMPROPER_PLAN': trans_map[tt]['improper'] += 1
        elif c == 'LATE_CONTRACTOR': trans_map[tt]['late_cont'] += 1
        elif c == 'LATE_CLIENT': trans_map[tt]['late_client'] += 1
        elif c == 'COMPOUND_DELAY': trans_map[tt]['compound'] += 1

    # Macro Statement of Conclusion
    macro_statement = (
        f"Forensic delay attribution across {total_eval:,} revision transitions reveals that {cont_driven_count:,} out of "
        f"{len(overdue_list):,} delayed transitions ({cont_overrun_pct}%) are driven by Contractor comment incorporation exceeding planned allowances. "
        f"Because this exceeds the 50% threshold, it provides conclusive analytical justification that the Contractor's Rebaseline Plan "
        f"was systematically over-optimistic and improperly planned relative to their actual engineering capacity."
        if macro_justification_flag else
        f"Delay attribution indicates balanced operational friction: Client review accounted for {len(late_client_list):,} delays, "
        f"while Contractor incorporation accounted for {len(late_cont_list):,} delays."
    )

    summary_obj = {
        'total_evaluated_transitions': total_eval,
        'met_plan_count': len(met_plan_list),
        'met_plan_pct': round(len(met_plan_list) / total_eval * 100, 1) if total_eval else 0.0,
        'overdue_count': len(overdue_list),
        'overdue_pct': round(len(overdue_list) / total_eval * 100, 1) if total_eval else 0.0,
        
        'improper_plan_count': len(improper_list),
        'improper_plan_pct_of_overdue': round(len(improper_list) / len(overdue_list) * 100, 1) if overdue_list else 0.0,
        
        'late_contractor_count': len(late_cont_list),
        'late_contractor_pct_of_overdue': round(len(late_cont_list) / len(overdue_list) * 100, 1) if overdue_list else 0.0,
        
        'late_client_count': len(late_client_list),
        'late_client_pct_of_overdue': round(len(late_client_list) / len(overdue_list) * 100, 1) if overdue_list else 0.0,
        
        'compound_delay_count': len(compound_list),
        'compound_delay_pct_of_overdue': round(len(compound_list) / len(overdue_list) * 100, 1) if overdue_list else 0.0,
        
        'contractor_driven_count': cont_driven_count,
        'contractor_overrun_pct': cont_overrun_pct,
        'macro_justification_flag': macro_justification_flag,
        'macro_statement': macro_statement,

        'client_sla_compliance_rate': client_sla_rate,
        'avg_plan_wd': avg_plan_wd,
        'avg_actual_wd': avg_actual_wd,
        'avg_client_rev_wd': avg_client_wd,
        'avg_cont_inc_wd': avg_cont_inc_wd,
        
        'discipline_breakdown': disc_map,
        'transition_breakdown': trans_map
    }

    # Extract cut-off date from filename
    base_name = Path(mdr_path).name
    m = re.search(r'(\d{1,2})[-\s]([A-Za-z]{3})[-\s](\d{2,4})', base_name)
    report_date = f"{m.group(1).zfill(2)}-{m.group(2).capitalize()}-{('20' + m.group(3)) if len(m.group(3)) == 2 else m.group(3)}" if m else "18-Sep-2026"

    out_file = Path('analytics_data.js')
    with open(out_file, 'w', encoding='utf-8') as f:
        f.write(f'window.ANALYTICS_REPORT_DATE = "{report_date}";\n')
        f.write('window.ANALYTICS_SUMMARY = ' + json.dumps(summary_obj, ensure_ascii=False) + ';\n')
        f.write('window.ANALYTICS_DATA = ' + json.dumps(transitions, ensure_ascii=False) + ';\n')

    print(f"[Analytics] Successfully generated {out_file} ({len(transitions)} total transitions, {total_eval} evaluated).")
    print(f"[Analytics] Macro Justification: {cont_overrun_pct}% Contractor overrun (Flag: {macro_justification_flag})")
    return True

if __name__ == '__main__':
    process_analytics()
