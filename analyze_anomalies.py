# -*- coding: utf-8 -*-
import csv
import sys
import io
import os
from collections import defaultdict

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(SCRIPT_DIR, 'rpt_data.csv')

KALASIN_NETWORKS = {
    '10709': 'เมืองกาฬสินธุ์', '11077': 'นามน', '11078': 'กมลาไสย',
    '11079': 'ร่องคำ', '11080': 'เขาวง', '11081': 'ยางตลาด',
    '11082': 'ห้วยเม็ก', '11083': 'สหัสขันธ์', '11084': 'คำม่วง',
    '11085': 'ท่าคันโท', '11086': 'หนองกุงศรี', '11087': 'สมเด็จ',
    '11088': 'ห้วยผึ้ง', '11449': 'กุฉินารายณ์', '28017': 'นาคู',
    '28789': 'ฆ้องชัย', '28790': 'ดอนจาน', '28791': 'สามชัย',
}

if not os.path.exists(INPUT_FILE):
    print(f"❌ ไม่พบไฟล์ข้อมูลเพื่อตรวจสอบ: {INPUT_FILE}")
    sys.exit(1)

rows = []
with open(INPUT_FILE, encoding='utf-8') as f:
    r = csv.reader(f)
    header = next(r)
    for row in r:
        if len(row) >= 8 and row[1].strip() in KALASIN_NETWORKS:
            rows.append(row)

units = {}
for r in rows:
    net_code = r[1].strip()
    unit_type = r[2].strip()
    unit_code = r[3].strip()
    unit_name = r[4].strip()
    vaccine_type = r[5].strip()
    measure = r[6].strip()
    try: val = int(float(r[7].replace(',', '').strip()))
    except: val = 0
    if unit_code not in units:
        units[unit_code] = {
            'network': net_code, 'district': KALASIN_NETWORKS[net_code],
            'type': unit_type, 'name': unit_name,
            'v11_target_unit': 0, 'v11_target_total': 0, 'v11_vaccinated': 0,
            'v12_target_unit': 0, 'v12_target_total': 0, 'v12_vaccinated': 0,
        }
    is_v11 = 'V11' in vaccine_type
    is_v12 = 'V12' in vaccine_type
    if measure == 'ฉีด':
        if is_v11: units[unit_code]['v11_vaccinated'] = val
        elif is_v12: units[unit_code]['v12_vaccinated'] = val
    elif measure == 'เป้าหมาย(รายหน่วย)':
        if is_v11: units[unit_code]['v11_target_unit'] = val
        elif is_v12: units[unit_code]['v12_target_unit'] = val
    elif measure == 'เป้าหมาย(ทั้งหมด)':
        if is_v11: units[unit_code]['v11_target_total'] = val
        elif is_v12: units[unit_code]['v12_target_total'] = val

# Aggregate networks
networks = defaultdict(lambda: {
    'district': '', 'main_name': '', 'main_code': '', 'sub_count': 0,
    'v11_target': 0, 'v11_vaccinated': 0, 'v12_target': 0, 'v12_vaccinated': 0, 'units': []
})
for code, info in units.items():
    net = info['network']
    networks[net]['district'] = info['district']
    if info['type'] == 'แม่ข่าย':
        networks[net]['main_name'] = info['name']
        networks[net]['main_code'] = code
        if info['v11_target_total'] > 0: networks[net]['v11_target'] = info['v11_target_total']
        if info['v12_target_total'] > 0: networks[net]['v12_target'] = info['v12_target_total']
    else:
        networks[net]['sub_count'] += 1
    networks[net]['v11_vaccinated'] += info['v11_vaccinated']
    networks[net]['v12_vaccinated'] += info['v12_vaccinated']
    networks[net]['units'].append(info | {'code': code})

for net, data in networks.items():
    if data['v11_target'] == 0:
        data['v11_target'] = sum(u['v11_target_unit'] for u in data['units'])
    if data['v12_target'] == 0:
        data['v12_target'] = sum(u['v12_target_unit'] for u in data['units'])

# Aggregate districts
districts = defaultdict(lambda: {
    'v11_target': 0, 'v11_vaccinated': 0, 'v12_target': 0, 'v12_vaccinated': 0
})
for net_code, data in networks.items():
    d = data['district']
    districts[d]['v11_target'] += data['v11_target']
    districts[d]['v11_vaccinated'] += data['v11_vaccinated']
    districts[d]['v12_target'] += data['v12_target']
    districts[d]['v12_vaccinated'] += data['v12_vaccinated']

print('--- 1. อำเภอที่ผลงานเกิน 100% (ตรวจสอบเป้าหมาย) ---')
has_dist = False
for dn, d in districts.items():
    tt = d['v11_target'] + d['v12_target']
    tv = d['v11_vaccinated'] + d['v12_vaccinated']
    pct = (tv / tt) * 100 if tt > 0 else 0
    if pct > 100:
        has_dist = True
        print(f'อำเภอ {dn}: เป้าหมาย {tt:,} | ฉีด {tv:,} | คิดเป็น {pct:.2f}%')
if not has_dist:
    print('ไม่มีอำเภอใดที่ยอดฉีดภาพรวมเกิน 100%')

print('\n--- 2. เครือข่าย (CUP) ที่ผลงานเกิน 100% ---')
has_net = False
for nc, n in networks.items():
    tt = n['v11_target'] + n['v12_target']
    tv = n['v11_vaccinated'] + n['v12_vaccinated']
    pct = (tv / tt) * 100 if tt > 0 else 0
    if pct > 100:
        has_net = True
        print(f'เครือข่าย {n["main_name"]} ({n["district"]}): เป้าหมาย {tt:,} | ฉีด {tv:,} | คิดเป็น {pct:.2f}%')
if not has_net:
    print('ไม่มีเครือข่ายใดที่ยอดฉีดเกิน 100%')

print('\n--- 3. หน่วยบริการที่มีผลงานเกิน 100% ---')
unit_anomalies = []
for code, u in units.items():
    tt = u['v11_target_unit'] + u['v12_target_unit']
    tv = u['v11_vaccinated'] + u['v12_vaccinated']
    pct = (tv / tt) * 100 if tt > 0 else 0
    if pct > 100:
        unit_anomalies.append((u, tt, tv, pct))

unit_anomalies.sort(key=lambda x: x[3], reverse=True)
for u, tt, tv, pct in unit_anomalies:
    print(f'[{u["type"]}] {u["name"]} ({u["district"]}): เป้าหมาย {tt:,} | ฉีด {tv:,} | คิดเป็น {pct:.2f}% (รหัสเครือข่าย {u["network"]})')
if not unit_anomalies:
    print('ไม่มีหน่วยบริการใดที่ยอดฉีดเกิน 100%')

print('\n--- 4. ตรวจสอบเป้าหมายเครือข่ายกุฉินารายณ์ ---')
d_kuchi = districts.get('กุฉินารายณ์', {'v11_target':0,'v12_target':0,'v11_vaccinated':0,'v12_vaccinated':0})
tt_kuchi = d_kuchi['v11_target'] + d_kuchi['v12_target']
tv_kuchi = d_kuchi['v11_vaccinated'] + d_kuchi['v12_vaccinated']
pct_kuchi = (tv_kuchi / tt_kuchi) * 100 if tt_kuchi > 0 else 0
print(f'อำเภอกุฉินารายณ์: เป้าหมายรวม={tt_kuchi:,} | ยอดฉีดรวม={tv_kuchi:,} | ร้อยละ={pct_kuchi:.2f}%')
