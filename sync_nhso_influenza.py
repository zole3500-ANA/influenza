# -*- coding: utf-8 -*-
"""
ระบบดึงข้อมูล NHSO อัตโนมัติและอัพเดท Influenza แดชบอร์ด
============================================================
"""

import sys
import io
import os
import csv
import json
import shutil
import subprocess
from datetime import datetime

# Prevent encoding issues on Windows
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# 18 cup network codes in Kalasin
KALASIN_NETWORKS = {
    '10709', '11077', '11078', '11079', '11080', '11081',
    '11082', '11083', '11084', '11085', '11086', '11087',
    '11088', '11449', '28017', '28789', '28790', '28791',
}

def find_git_exe():
    git_path = shutil.which('git')
    if git_path:
        return git_path
    common_paths = [
        r"C:\Program Files\Git\cmd\git.exe",
        r"C:\Program Files (x86)\Git\cmd\git.exe",
        r"C:\Users\CDCKsn2\AppData\Local\Programs\Git\cmd\git.exe",
    ]
    for p in common_paths:
        if os.path.exists(p):
            return p
    return "git"

GIT_EXE = find_git_exe()

def scrape_nhso_data():
    print("============================================================")
    print("🌐 กำลังดึงข้อมูลจาก NHSO Tableau Dashboard...")
    print("============================================================")
    
    # We will use playwright inside Python
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ ไม่พบ playwright package กรุณาติดตั้งโดยรัน: pip install playwright && playwright install")
        sys.exit(1)
        
    url = "https://medata.nhso.go.th/me/public/dashboard/41/881"
    
    with sync_playwright() as p:
        # Launch browser
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        
        try:
            # Navigate
            page.goto(url, timeout=90000, wait_until="domcontentloaded")
            print("DOM โหลดสำเร็จ กำลังรอ Tableau Viz ทำงาน...")
            
            # Give the page extra time to fully initialize Tableau
            page.wait_for_timeout(10000)
            
            # Combined: wait for workbook readiness + switch sheet + pull data
            # Uses try-catch to handle the workbookImpl race condition
            js_query = """
                async () => {
                    // Phase 1: Wait for Tableau Viz workbook to fully initialize
                    const maxWaitMs = 60000;
                    const startTime = Date.now();
                    
                    while (Date.now() - startTime < maxWaitMs) {
                        try {
                            const viz = document.getElementById("tableau-viz");
                            if (viz && viz.workbook && viz.workbook.publishedSheetsInfo && viz.workbook.publishedSheetsInfo.length > 3) {
                                break; // Ready!
                            }
                        } catch (e) {
                            // workbookImpl not ready yet, keep waiting
                        }
                        await new Promise(r => setTimeout(r, 2000));
                    }
                    
                    // Phase 2: Verify workbook is truly ready
                    const viz = document.getElementById("tableau-viz");
                    if (!viz) return { error: "tableau-viz element not found" };
                    
                    let workbook;
                    try {
                        workbook = viz.workbook;
                        if (!workbook || !workbook.publishedSheetsInfo) {
                            return { error: "Workbook not ready after waiting " + maxWaitMs + "ms" };
                        }
                    } catch (e) {
                        return { error: "Workbook access error: " + e.message };
                    }
                    
                    // Phase 3: Switch to Dashboard รายงานบริการ (index 3)
                    const sheetName = workbook.publishedSheetsInfo[3].name;
                    const activeSheet = await workbook.activateSheetAsync(sheetName);
                    
                    // Wait for the new sheet to fully render
                    await new Promise(r => setTimeout(r, 12000));
                    
                    // Phase 4: Find rpt worksheet and pull data
                    const rptWs = activeSheet.worksheets.find(ws => ws.name === "rpt");
                    if (!rptWs) return { error: "Worksheet rpt not found in active sheet. Available: " + activeSheet.worksheets.map(w => w.name).join(", ") };
                    
                    const dataTable = await rptWs.getSummaryDataAsync({
                        maxRows: 0,
                        ignoreSelection: true
                    });
                    
                    // Phase 5: Filter for Kalasin networks
                    const kalasinNetworks = new Set([
                        '10709', '11077', '11078', '11079', '11080', '11081',
                        '11082', '11083', '11084', '11085', '11086', '11087',
                        '11088', '11449', '28017', '28789', '28790', '28791'
                    ]);
                    
                    const filteredRows = [];
                    for (const row of dataTable.data) {
                        const netCode = row[1].value ? row[1].value.toString().trim() : '';
                        if (kalasinNetworks.has(netCode)) {
                            filteredRows.push(row.map(cell => ({
                                value: cell.value,
                                formattedValue: cell.formattedValue
                            })));
                        }
                    }
                    
                    return {
                        success: true,
                        rowCount: filteredRows.length,
                        rows: filteredRows
                    };
                }
            """
            
            # Retry up to 3 times
            result = None
            last_error = None
            for attempt in range(1, 4):
                print(f"  ▶ ความพยายามครั้งที่ {attempt}/3...")
                try:
                    result = page.evaluate(js_query)
                    if result and result.get("success"):
                        break
                    last_error = result.get("error", "Unknown error") if result else "No result"
                    print(f"  ⚠️ ครั้งที่ {attempt} ไม่สำเร็จ: {last_error}")
                except Exception as e:
                    last_error = str(e)
                    print(f"  ⚠️ ครั้งที่ {attempt} เกิด exception: {last_error}")
                
                if attempt < 3:
                    print("  ⏳ รอ 15 วินาทีแล้วลองใหม่...")
                    # Reload the page for a fresh retry
                    page.goto(url, timeout=90000, wait_until="domcontentloaded")
                    page.wait_for_timeout(15000)
            
            if not result or "error" in result:
                print("❌ เกิดข้อผิดพลาดในการดึงข้อมูลจาก Tableau:", result.get("error", "Unknown error"))
                sys.exit(1)
                
            print(f"✅ ดึงข้อมูลสำเร็จ! พบข้อมูลหน่วยบริการในกาฬสินธุ์ทั้งหมด {result['rowCount']} รายการ")
            
            # Map and write to rpt_data.csv
            csv_path = os.path.join(SCRIPT_DIR, 'rpt_data.csv')
            with open(csv_path, 'w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                # Header
                writer.writerow(['เขต', 'เครือข่าย', 'ประเภท', 'รหัสหน่วยบริการ', 'ชื่อหน่วยบริการ', 'วัคซีน', 'Measure Names', 'Measure Values'])
                
                for row in result['rows']:
                    # Row mappings:
                    # 0: เขต (Zone)
                    # 1: เครือข่าย (Network ID)
                    # 2: ประเภท (Type) -> Main/Child -> แม่ข่าย/ลูกข่าย
                    # 3: รหัสหน่วยบริการ (Service Unit Code)
                    # 4: ชื่อหน่วยบริการ (Service Unit Name)
                    # 5: วัคซีน (Vaccine)
                    # 6: Measure Names -> formattedValue (ฉีด/เป้าหมาย(รายหน่วย)/เป้าหมาย(ทั้งหมด))
                    # 7: Measure Values -> raw value (integer)
                    
                    zone = row[0]['value']
                    net_code = row[1]['value']
                    
                    raw_type = row[2]['value']
                    unit_type = 'แม่ข่าย' if raw_type == 'Main' else 'ลูกข่าย'
                    
                    unit_code = row[3]['value']
                    unit_name = row[4]['value'].strip() if row[4]['value'] else ''
                    vaccine = row[5]['value']
                    
                    measure_name = row[6]['formattedValue']
                    
                    raw_val = row[7]['value']
                    measure_value = 0 if raw_val is None or raw_val == '%null%' else int(float(raw_val))
                    
                    writer.writerow([zone, net_code, unit_type, unit_code, unit_name, vaccine, measure_name, measure_value])
                    
            print(f"💾 เขียนข้อมูลดิบลงไฟล์สำเร็จ: {csv_path}")
            
        except Exception as e:
            print("❌ เกิดข้อผิดพลาดขณะดึงข้อมูล:", e)
            sys.exit(1)
        finally:
            browser.close()

def run_update_and_deploy():
    print("============================================================")
    print("📊 กำลังประมวลผลสรุปข้อมูลวัคซีนไข้หวัดใหญ่...")
    print("============================================================")
    
    # Run update_report.py
    update_script = os.path.join(SCRIPT_DIR, 'update_report.py')
    try:
        subprocess.run([sys.executable, update_script], cwd=SCRIPT_DIR, check=True)
        print("✅ ประมวลผลและสร้าง Fludashboard.html สำเร็จ!")
    except Exception as e:
        print("❌ เกิดข้อผิดพลาดขณะรัน update_report.py:", e)
        sys.exit(1)
        
    # Run analyze_anomalies.py for verification
    anomalies_script = os.path.join(SCRIPT_DIR, 'analyze_anomalies.py')
    try:
        print("\n🔍 กำลังตรวจสอบความถูกต้องของข้อมูล (analyze_anomalies.py)...")
        subprocess.run([sys.executable, anomalies_script], cwd=SCRIPT_DIR, check=True)
        print("✅ ตรวจสอบความถูกต้องเรียบร้อย!")
    except Exception as e:
        print("⚠️ การรันตรวจสอบพบบันทึกเตือนหรือข้อผิดพลาด แต่จะยังดำเนินการจัดส่งต่อ")
        
    # Copy Fludashboard.html to index.html
    flu_html = os.path.join(SCRIPT_DIR, 'Fludashboard.html')
    index_html = os.path.join(SCRIPT_DIR, 'index.html')
    try:
        shutil.copyfile(flu_html, index_html)
        print("📋 คัดลอก Fludashboard.html ไปยัง index.html สำเร็จ!")
    except Exception as e:
        print("❌ คัดลอกไฟล์ล้มเหลว:", e)
        sys.exit(1)
        
    # Git commit and push
    print("\n🚀 กำลังส่งออกข้อมูลขึ้น GitHub...")
    try:
        # Check if it is a git repo
        if not os.path.isdir(os.path.join(SCRIPT_DIR, '.git')):
            print("⚠️ โฟลเดอร์ปัจจุบันไม่ใช่ Git Repository (ข้ามการส่งออก GitHub)")
            return
            
        # git add
        subprocess.run([GIT_EXE, 'add', 'rpt_data.csv', 'Fludashboard.html', 'index.html'], cwd=SCRIPT_DIR, check=True)
        # git commit
        commit_msg = f"Auto-update Influenza Dashboard from NHSO API (ข้อมูล ณ วันที่ {datetime.now().strftime('%d/%m/%Y %H:%M')})"
        subprocess.run([GIT_EXE, 'commit', '-m', commit_msg], cwd=SCRIPT_DIR, capture_output=True)
        # git push
        subprocess.run([GIT_EXE, 'push', 'origin', 'main'], cwd=SCRIPT_DIR, check=True)
        print("🎉 ดันข้อมูลขึ้น GitHub Pages เรียบร้อยแล้ว!")
        print("🔗 https://zole3500-ana.github.io/influenza/index.html")
    except Exception as e:
        print("❌ เกิดข้อผิดพลาดในขั้นตอน Git Push:", e)
        sys.exit(1)

if __name__ == "__main__":
    scrape_nhso_data()
    run_update_and_deploy()
