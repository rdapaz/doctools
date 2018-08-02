
import psycopg2 
import win32com.client


xlApp = win32com.client.gencache.EnsureDispatch('Excel.Application')
xlApp.Visible = True

wk = xlApp.Workbooks.Open(r'C:\Users\rdapaz\Desktop\Belmont DC Exit - Implementation\Execution\VM Migration\Belmont_to_AWS-Migration.xlsx')
sh = wk.Worksheets('Belmont_VMs')

eof = sh.Range('D65536').End(-4162).Row


conn = psycopg2.connect("dbname='CPM_VMs' user=postgres")
cur = conn.cursor()

sql = '''
    SELECT
        d.vm_migrated,
        d.backup_cfg,
        d.ip_address,
        d.chg_no,
        d.migration_type,
        d.vm_type,
        d.proposed_stt_dttm,
        d.proposed_fin_dttm,
        d.approved_stt_dttm,
        d.approved_fin_dttm,
        d.actual_stt_dttm,
        d.actual_fin_dttm 
    FROM
        "DevTest" d 
    WHERE
        vm = %s
'''

for row in range(9, eof+1):
    vm = sh.Range(f'D{row}').Value
    cur.execute(sql, (vm,))
    rows = cur.fetchall()
    for row_data in rows:
        vm_migrated, backup_cfg, ip_address, chg_no, migration_type, vm_type, proposed_stt_dttm, proposed_fin_dttm, approved_stt_dttm, approved_fin_dttm, actual_stt_dttm, actual_fin_dttm = row_data
        vm_migrated = 'Yes' if vm_migrated else 'No'
        backup_cfg = 'Yes' if backup_cfg else 'No'
        ip_address = ip_address if ip_address  else ''
        chg_no = chg_no if chg_no else ''
        proposed_stt_dttm = f'{proposed_stt_dttm}' if proposed_stt_dttm else ''
        approved_stt_dttm = f'{approved_stt_dttm}' if approved_stt_dttm else ''
        actual_stt_dttm = f'{actual_stt_dttm}' if actual_stt_dttm else ''
        proposed_fin_dttm = f'{proposed_fin_dttm}' if proposed_fin_dttm else ''
        approved_fin_dttm = f'{approved_fin_dttm}' if approved_fin_dttm else ''
        actual_fin_dttm = f'{actual_fin_dttm}' if actual_fin_dttm else ''
        print(vm_migrated, backup_cfg, ip_address, chg_no, migration_type, vm_type, proposed_stt_dttm, proposed_fin_dttm, approved_stt_dttm, approved_fin_dttm, actual_stt_dttm, actual_fin_dttm, sep="|")
        sh.Range(f'T{row}').Value = vm_migrated
        sh.Range(f'U{row}').Value = backup_cfg
        sh.Range(f'V{row}').Value = ip_address
        sh.Range(f'W{row}').Value = chg_no
        sh.Range(f'X{row}').Value = migration_type
        sh.Range(f'Y{row}').Value = vm_type
        sh.Range(f'Z{row}').Value = proposed_stt_dttm
        sh.Range(f'AA{row}').Value = proposed_fin_dttm
        sh.Range(f'AB{row}').Value = approved_stt_dttm
        sh.Range(f'AC{row}').Value = approved_fin_dttm
        sh.Range(f'AD{row}').Value = actual_stt_dttm
        sh.Range(f'AE{row}').Value = actual_fin_dttm