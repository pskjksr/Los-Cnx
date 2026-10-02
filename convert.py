import pandas as pd
import openpyxl
import datetime
import json

file_name = "CNX LoS.xlsx"
wb = openpyxl.load_workbook(file_name, data_only=True)

all_records = []

def to_total_seconds(val):
    """แปลงค่าเวลา MM:SS ให้ออกมาเป็นจำนวนวินาทีรวม"""
    if val is None or str(val).strip() in ["", "-", "None", "min.", "min"]:
        return None
        
    if isinstance(val, (datetime.time, datetime.datetime)):
        return val.hour * 60 + val.minute + (val.second / 60.0)
        
    if isinstance(val, str):
        v = val.strip().replace("min.", "").replace("min", "").strip()
        if ":" in v:
            parts = v.split(":")
            try:
                if len(parts) == 2:
                    return float(parts[0]) * 60 + float(parts[1])
                elif len(parts) == 3:
                    if float(parts[0]) == 0:
                        return float(parts[1]) * 60 + float(parts[2])
                    else:
                        return float(parts[0]) * 60 + float(parts[1]) + (float(parts[2]) / 60.0)
            except ValueError:
                return None
        try:
            return float(v) * 60
        except ValueError:
            return None
            
    if isinstance(val, (int, float)):
        if 0 < val < 1:
            total_mins = val * 24 * 60
            if total_mins > 60:
                total_mins = total_mins / 60.0
            return total_mins * 60
        return float(val) * 60
        
    return None

def calc_waiting_seconds(start_val, finish_val, waiting_raw):
    """คำนวณเวลารอคอยเป็นวินาที (Finish - Start)"""
    s_sec = to_total_seconds(start_val)
    f_sec = to_total_seconds(finish_val)
    
    if s_sec is not None and f_sec is not None:
        diff = f_sec - s_sec
        if diff >= 0:
            return round(diff, 1)
            
    w_sec = to_total_seconds(waiting_raw)
    if w_sec is not None and w_sec >= 0:
        return round(w_sec, 1)
        
    return 0.0

def fmt_mmss(val):
    """จัดรูปแบบแสดงผลเวลา Start และ Finish ให้เป็น MM:SS"""
    if val is None or str(val).strip() in ["", "-", "None"]:
        return "-"
    if isinstance(val, (datetime.time, datetime.datetime)):
        return "{:02d}:{:02d}".format(val.hour, val.minute)
    v = str(val).strip()
    if ":" in v:
        parts = v.split(":")
        if len(parts) >= 2:
            try:
                m = int(float(parts[0]))
                s = int(float(parts[1]))
                return "{:02d}:{:02d}".format(m, s)
            except ValueError:
                return v
    return v

def seconds_to_mmss(total_seconds):
    """แปลงวินาทีให้แสดงผลเป็น MM:SS (เช่น 56 วิ -> 00:56)"""
    if total_seconds is None or total_seconds <= 0:
        return "00:00"
    m = int(total_seconds // 60)
    s = int(round(total_seconds % 60))
    if s >= 60:
        m += 1
        s = 0
    return "{:02d}:{:02d}".format(m, s)

for sheet in wb.sheetnames:
    ws = wb[sheet]
    sheet_name = sheet.strip()

    parts = sheet_name.split()
    month_name = parts[0] if len(parts) > 0 else sheet_name
    year_val = "2026"
    if len(parts) > 1 and parts[1].isdigit():
        y_num = int(parts[1])
        if len(parts[1]) == 2:
            year_val = f"20{y_num}"
        elif y_num >= 2500:
            year_val = str(y_num - 543)
        else:
            year_val = str(y_num)

    current_flight_type = "International"
    current_direction = "Departure"

    r = 1
    max_row = ws.max_row
    while r <= max_row:
        row_cells = [str(ws.cell(row=r, column=c).value or "").strip().lower() for c in range(1, 6)]
        row_str = " ".join(row_cells)

        if "international" in row_str:
            current_flight_type = "International"
            current_direction = "Departure"
            r += 1
            continue
        elif "domestic" in row_str:
            current_flight_type = "Domestic"
            current_direction = "Departure"
            r += 1
            continue

        if "departure" in row_str:
            current_direction = "Departure"
            r += 1
            continue
        elif "arrival" in row_str:
            current_direction = "Arrival"
            r += 1
            continue

        cell_a = str(ws.cell(row=r, column=1).value or "").strip().replace("\n", " ")
        cell_b = str(ws.cell(row=r, column=2).value or "").strip().lower()

        if "start" in cell_b:
            service_name = cell_a if cell_a else ""
            if not service_name:
                for back_r in range(r, max(1, r-4), -1):
                    val_back = str(ws.cell(row=back_r, column=1).value or "").strip().replace("\n", " ")
                    if val_back and not any(x in val_back.lower() for x in ["departure", "arrival", "international", "domestic", "วัน"]):
                        service_name = val_back
                        break

            s_clean = service_name
            s_lower = service_name.lower()
            if "check" in s_lower:
                s_clean = "Check-In"
            elif "in-line" in s_lower:
                s_clean = "In-Line Screening"
            elif "security" in s_lower:
                s_clean = "Security Screening"
            elif "immigra" in s_lower:
                s_clean = "Immigration"
            elif "baggag" in s_lower:
                s_clean = "Baggage Claim"
            elif "taxi" in s_lower:
                s_clean = "Taxi"

            real_direction = current_direction
            if s_clean in ["Check-In", "In-Line Screening", "Security Screening"]:
                real_direction = "Departure"
            elif s_clean in ["Immigration", "Baggage Claim", "Taxi"]:
                real_direction = "Arrival"

            has_pax = True
            cell_pax_label = str(ws.cell(row=r+3, column=2).value or "").strip().lower()
            if "passenger" not in cell_pax_label:
                has_pax = False

            for day in range(1, 32):
                col_idx = day + 2
                
                head_val = str(ws.cell(row=2, column=col_idx).value or "").strip().lower()
                if "average" in head_val:
                    break

                start_val = ws.cell(row=r, column=col_idx).value
                finish_val = ws.cell(row=r+1, column=col_idx).value
                waiting_raw = ws.cell(row=r+2, column=col_idx).value
                pax_val = ws.cell(row=r+3, column=col_idx).value if has_pax else 0

                wait_seconds = calc_waiting_seconds(start_val, finish_val, waiting_raw)

                pax_int = 0
                if has_pax:
                    if isinstance(pax_val, (int, float)):
                        pax_int = int(pax_val)
                    elif isinstance(pax_val, str) and pax_val.strip().isdigit():
                        pax_int = int(pax_val.strip())

                if wait_seconds > 0 or pax_int > 0 or (start_val and str(start_val).strip() not in ["-", ""]):
                    all_records.append({
                        "ปี": year_val,
                        "เดือน": month_name,
                        "วันที่": day,
                        "ประเภท": current_flight_type,
                        "ขาเข้า/ขาออก": real_direction,
                        "Service": s_clean,
                        "Start": fmt_mmss(start_val),
                        "Finish": fmt_mmss(finish_val),
                        "WaitingSec": wait_seconds,
                        "Waiting": seconds_to_mmss(wait_seconds),
                        "Passenger": pax_int
                    })

            r += (4 if has_pax else 3)
        else:
            r += 1

# บันทึกเป็น Excel
df_clean = pd.DataFrame(all_records)
df_clean.to_excel("CNX_LoS_clean.xlsx", index=False)

# แปลงเป็น JSON สำหรับ HTML
json_data = json.dumps(all_records, ensure_ascii=False)

# เขียน HTML Dashboard
html_template = """<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CNX Airport Service Quality Dashboard</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Prompt', sans-serif; }
  </style>
</head>
<body class="bg-slate-50 text-slate-800 min-h-screen">

  <header class="bg-blue-900 text-white shadow-md">
    <div class="max-w-7xl mx-auto px-4 py-5 flex flex-col md:flex-row md:items-center md:justify-between gap-4">
      <div>
        <h1 class="text-2xl font-bold flex items-center gap-2">
          <span>✈</span> ท่าอากาศยานเชียงใหม่ (CNX) - Level of Service
        </h1>
        <p class="text-blue-200 text-sm mt-1">ระบบติดตามและวิเคราะห์คุณภาพการบริการผู้โดยสาร (LoS Analysis)</p>
      </div>
      <div class="text-xs bg-emerald-800/80 text-emerald-200 px-3 py-1.5 rounded-full inline-flex items-center gap-2 self-start md:self-auto">
        <span class="w-2 h-2 rounded-full bg-emerald-400"></span> ข้อมูลพร้อมใช้งาน (Standalone)
      </div>
    </div>
  </header>

  <main class="max-w-7xl mx-auto px-4 py-6 space-y-6">

    <!-- Filters Section -->
    <section class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
      <div class="flex items-center justify-between pb-3 border-b border-slate-100 mb-4">
        <h2 class="font-semibold text-slate-700 flex items-center gap-2">
          <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 4a1 1 0 011-1h16a1 1 0 011 1v2.586a1 1 0 01-.293.707l-6.414 6.414a1 1 0 00-.293.707V17l-4 4v-6.586a1 1 0 00-.293-.707L3.293 7.293A1 1 0 013 6.586V4z"/></svg>
          ตัวกรองข้อมูล
        </h2>
        <button id="resetBtn" class="text-xs text-blue-600 hover:text-blue-800 font-medium">ล้างตัวกรองทั้งหมด</button>
      </div>

      <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-6 gap-3">
        <div>
          <label class="block text-xs font-medium text-slate-500 mb-1">ปี (Year)</label>
          <select id="yearFilter" class="w-full border border-slate-300 rounded-lg px-2.5 py-2 text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none">
            <option value="ALL">ทุกปี</option>
          </select>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-500 mb-1">เดือน (Month)</label>
          <select id="monthFilter" class="w-full border border-slate-300 rounded-lg px-2.5 py-2 text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none">
            <option value="ALL">ทั้งหมด</option>
          </select>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-500 mb-1">ประเภทเที่ยวบิน</label>
          <select id="flightTypeFilter" class="w-full border border-slate-300 rounded-lg px-2.5 py-2 text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none">
            <option value="ALL">ทั้งหมด</option>
          </select>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-500 mb-1">ขาเข้า/ขาออก</label>
          <select id="directionFilter" class="w-full border border-slate-300 rounded-lg px-2.5 py-2 text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none">
            <option value="ALL">ทั้งหมด</option>
          </select>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-500 mb-1">จุดบริการ (Service)</label>
          <select id="serviceFilter" class="w-full border border-slate-300 rounded-lg px-2.5 py-2 text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none">
            <option value="ALL">ทั้งหมด</option>
          </select>
        </div>
        <div>
          <label class="block text-xs font-medium text-slate-500 mb-1">วันที่</label>
          <select id="dayFilter" class="w-full border border-slate-300 rounded-lg px-2.5 py-2 text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none">
            <option value="ALL">ทุกวัน</option>
          </select>
        </div>
      </div>
    </section>

    <!-- KPI Summary Cards -->
    <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-4">
      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <p class="text-xs font-medium text-slate-500 uppercase">ผู้โดยสารรวม</p>
        <h3 id="kpiTotalPax" class="text-2xl font-bold text-slate-800 mt-1">0</h3>
        <p class="text-xs text-slate-400 mt-1">คน</p>
      </div>

      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <p class="text-xs font-medium text-emerald-600 uppercase font-semibold">ผู้โดยสารเฉลี่ย/วัน</p>
        <h3 id="kpiAvgPax" class="text-2xl font-bold text-emerald-600 mt-1">0</h3>
        <p class="text-xs text-slate-400 mt-1">คน / วัน</p>
      </div>

      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <p class="text-xs font-medium text-blue-600 uppercase font-semibold">เวลารอคอยเฉลี่ย</p>
        <h3 id="kpiAvgWait" class="text-2xl font-bold text-blue-600 mt-1">00:00</h3>
        <p class="text-xs text-slate-400 mt-1">(นาที:วินาที)</p>
      </div>

      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <p class="text-xs font-medium text-teal-600 uppercase font-semibold">เวลารอคอยต่ำสุด</p>
        <h3 id="kpiMinWait" class="text-2xl font-bold text-teal-600 mt-1">00:00</h3>
        <p class="text-xs text-slate-400 mt-1">(นาที:วินาที)</p>
      </div>

      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <p class="text-xs font-medium text-rose-600 uppercase font-semibold">เวลารอคอยสูงสุด</p>
        <h3 id="kpiMaxWait" class="text-2xl font-bold text-rose-600 mt-1">00:00</h3>
        <p class="text-xs text-slate-400 mt-1">(นาที:วินาที)</p>
      </div>

      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4">
        <p class="text-xs font-medium text-slate-500 uppercase">จำนวนรายการ</p>
        <h3 id="kpiTotalRecords" class="text-2xl font-bold text-slate-800 mt-1">0</h3>
        <p class="text-xs text-slate-400 mt-1">เรคคอร์ด</p>
      </div>
    </section>

    <!-- Chart Section -->
    <section class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
      <h2 class="font-semibold text-slate-700 mb-4 flex items-center gap-2">
        <svg class="w-5 h-5 text-blue-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"/></svg>
        แนวโน้มเวลารอคอยเฉลี่ยตามวัน (MM:SS)
      </h2>
      <div class="h-64 sm:h-72">
        <canvas id="trendChart"></canvas>
      </div>
    </section>

    <!-- Table Section -->
    <section class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
      <div class="p-5 border-b border-slate-100 flex items-center justify-between">
        <h2 class="font-semibold text-slate-700">รายละเอียดข้อมูล LoS</h2>
        <span id="tableCounter" class="text-xs text-slate-500">แสดง 0 รายการ</span>
      </div>
      <div class="overflow-x-auto max-h-96">
        <table class="w-full text-sm text-left">
          <thead class="bg-slate-100 text-slate-600 uppercase text-xs sticky top-0">
            <tr>
              <th class="px-4 py-3">ปี</th>
              <th class="px-4 py-3">เดือน</th>
              <th class="px-4 py-3">วันที่</th>
              <th class="px-4 py-3">ประเภท</th>
              <th class="px-4 py-3">ขาเข้า/ออก</th>
              <th class="px-4 py-3">จุดบริการ (Service)</th>
              <th class="px-4 py-3 text-center">Start (MM:SS)</th>
              <th class="px-4 py-3 text-center">Finish (MM:SS)</th>
              <th class="px-4 py-3 text-center text-blue-600 font-bold">Waiting (MM:SS)</th>
              <th class="px-4 py-3 text-right">Passenger (คน)</th>
            </tr>
          </thead>
          <tbody id="dataTableBody" class="divide-y divide-slate-100 text-slate-700"></tbody>
        </table>
      </div>
    </section>

  </main>

  <script>
    const rawData = """ + json_data + """;
    let trendChartInstance = null;

    function formatSecToMMSS(seconds) {
      if (!seconds || seconds <= 0 || isNaN(seconds)) return '00:00';
      const m = Math.floor(seconds / 60);
      const s = Math.round(seconds % 60);
      return String(m).padStart(2, '0') + ':' + String(s).padStart(2, '0');
    }

    function parseMMSSToSec(str) {
      if (!str || str === '-' || str === '00:00') return 0;
      const parts = String(str).trim().split(':');
      if (parts.length === 2) {
        return (parseFloat(parts[0]) || 0) * 60 + (parseFloat(parts[1]) || 0);
      }
      return 0;
    }

    function populateSelect(elemId, items, keepValue = false) {
      const select = document.getElementById(elemId);
      const prevVal = select.value;
      select.innerHTML = '<option value="ALL">ทั้งหมด</option>';
      items.forEach(item => {
        const opt = document.createElement('option');
        opt.value = item;
        opt.textContent = item;
        select.appendChild(opt);
      });
      if (keepValue && items.includes(prevVal)) {
        select.value = prevVal;
      } else {
        select.value = 'ALL';
      }
    }

    // กรองตัวเลือก Service: แยกชัดเจนว่า International ไม่มี Taxi
    function updateServiceOptions() {
      const dirVal = document.getElementById('directionFilter').value;
      const ftVal = document.getElementById('flightTypeFilter').value;

      let allowedList = [];

      if (dirVal === 'Departure') {
        allowedList = ['Check-In', 'In-Line Screening', 'Security Screening'];
      } else if (dirVal === 'Arrival') {
        if (ftVal === 'Domestic') {
          // Domestic Arrival มี Baggage Claim และ Taxi
          allowedList = ['Baggage Claim', 'Taxi'];
        } else if (ftVal === 'International') {
          // International Arrival ไม่มี Taxi
          allowedList = ['Immigration', 'Baggage Claim'];
        } else {
          // Arrival ทั้งหมด
          allowedList = ['Immigration', 'Baggage Claim', 'Taxi'];
        }
      } else {
        // กรณีเลือกขาเข้า/ขาออก เป็น ALL
        if (ftVal === 'Domestic') {
          allowedList = ['Check-In', 'In-Line Screening', 'Security Screening', 'Baggage Claim', 'Taxi'];
        } else if (ftVal === 'International') {
          // International ทั้งหมด ไม่มี Taxi
          allowedList = ['Check-In', 'In-Line Screening', 'Security Screening', 'Immigration', 'Baggage Claim'];
        } else {
          allowedList = ['Check-In', 'In-Line Screening', 'Security Screening', 'Immigration', 'Baggage Claim', 'Taxi'];
        }
      }

      const existingInRaw = new Set(rawData.map(d => d['Service']));
      const finalList = allowedList.filter(s => existingInRaw.has(s));

      populateSelect('serviceFilter', finalList, true);
    }

    function initFilters() {
      const years = [...new Set(rawData.map(d => d['ปี']).filter(Boolean))].sort();
      const months = [...new Set(rawData.map(d => d['เดือน']).filter(Boolean))];
      const flightTypes = ['International', 'Domestic'];
      const directions = ['Departure', 'Arrival'];
      const days = [...new Set(rawData.map(d => d['วันที่']).filter(Boolean))].sort((a,b) => a - b);

      populateSelect('yearFilter', years);
      populateSelect('monthFilter', months);
      populateSelect('flightTypeFilter', flightTypes);
      populateSelect('directionFilter', directions);
      populateSelect('dayFilter', days);

      updateServiceOptions();

      document.getElementById('flightTypeFilter').addEventListener('change', () => {
        updateServiceOptions();
        renderDashboard();
      });

      document.getElementById('directionFilter').addEventListener('change', () => {
        updateServiceOptions();
        renderDashboard();
      });

      document.querySelectorAll('#yearFilter, #monthFilter, #serviceFilter, #dayFilter').forEach(el => {
        el.addEventListener('change', renderDashboard);
      });

      document.getElementById('resetBtn').addEventListener('click', () => {
        document.getElementById('yearFilter').value = 'ALL';
        document.getElementById('monthFilter').value = 'ALL';
        document.getElementById('flightTypeFilter').value = 'ALL';
        document.getElementById('directionFilter').value = 'ALL';
        document.getElementById('dayFilter').value = 'ALL';
        updateServiceOptions();
        renderDashboard();
      });
    }

    function renderDashboard() {
      const yVal = document.getElementById('yearFilter').value;
      const mVal = document.getElementById('monthFilter').value;
      const ftVal = document.getElementById('flightTypeFilter').value;
      const dirVal = document.getElementById('directionFilter').value;
      const sVal = document.getElementById('serviceFilter').value;
      const dVal = document.getElementById('dayFilter').value;

      const filtered = rawData.filter(d => {
        if (yVal !== 'ALL' && String(d['ปี']) !== yVal) return false;
        if (mVal !== 'ALL' && String(d['เดือน']) !== mVal) return false;
        if (ftVal !== 'ALL' && String(d['ประเภท']) !== ftVal) return false;
        if (dirVal !== 'ALL' && String(d['ขาเข้า/ขาออก']) !== dirVal) return false;
        if (sVal !== 'ALL' && String(d['Service']) !== sVal) return false;
        if (dVal !== 'ALL' && String(d['วันที่']) !== dVal) return false;
        return true;
      });

      const tbody = document.getElementById('dataTableBody');
      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="10" class="text-center py-8 text-slate-400">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</td></tr>`;
      } else {
        tbody.innerHTML = filtered.map(row => `
          <tr class="hover:bg-slate-50 transition-colors">
            <td class="px-4 py-2.5 font-medium">${row['ปี'] || '-'}</td>
            <td class="px-4 py-2.5">${row['เดือน'] || '-'}</td>
            <td class="px-4 py-2.5">${row['วันที่'] || '-'}</td>
            <td class="px-4 py-2.5"><span class="px-2 py-0.5 text-xs rounded-full ${row['ประเภท'] === 'International' ? 'bg-purple-100 text-purple-700' : 'bg-emerald-100 text-emerald-700'}">${row['ประเภท'] || '-'}</span></td>
            <td class="px-4 py-2.5"><span class="px-2 py-0.5 text-xs rounded-full ${row['ขาเข้า/ขาออก'] === 'Departure' ? 'bg-blue-100 text-blue-700' : 'bg-amber-100 text-amber-700'}">${row['ขาเข้า/ขาออก'] || '-'}</span></td>
            <td class="px-4 py-2.5 font-medium text-slate-800">${row['Service'] || '-'}</td>
            <td class="px-4 py-2.5 text-center font-mono text-xs text-slate-600">${row['Start'] || '-'}</td>
            <td class="px-4 py-2.5 text-center font-mono text-xs text-slate-600">${row['Finish'] || '-'}</td>
            <td class="px-4 py-2.5 text-center font-mono text-xs font-bold text-blue-600">${row['Waiting'] || '00:00'}</td>
            <td class="px-4 py-2.5 text-right">${Number(row['Passenger'] || 0).toLocaleString()}</td>
          </tr>
        `).join('');
      }

      let totalPax = 0;
      let totalWaitSec = 0;
      let validWaitRowsCount = 0;
      let maxSec = 0;
      let minSec = Infinity;
      const dayPaxMap = {};

      filtered.forEach(d => {
        const pax = parseFloat(d['Passenger']) || 0;
        totalPax += pax;

        const waitSec = parseMMSSToSec(d['Waiting']);
        const dayKey = `${d['ปี']}_${d['เดือน']}_${d['วันที่']}`;
        dayPaxMap[dayKey] = (dayPaxMap[dayKey] || 0) + pax;

        if (waitSec > 0) {
          totalWaitSec += waitSec;
          validWaitRowsCount++;
          if (waitSec > maxSec) maxSec = waitSec;
          if (waitSec < minSec) minSec = waitSec;
        }
      });

      const uniqueDays = Object.keys(dayPaxMap).length;
      const avgPaxPerDay = uniqueDays > 0 ? Math.round(totalPax / uniqueDays) : 0;
      const avgWaitSec = validWaitRowsCount > 0 ? (totalWaitSec / validWaitRowsCount) : 0;
      const displayMin = minSec !== Infinity ? minSec : 0;

      document.getElementById('kpiTotalPax').textContent = Number(totalPax).toLocaleString();
      document.getElementById('kpiAvgPax').textContent = Number(avgPaxPerDay).toLocaleString();
      document.getElementById('kpiAvgWait').textContent = formatSecToMMSS(avgWaitSec);
      document.getElementById('kpiMinWait').textContent = formatSecToMMSS(displayMin);
      document.getElementById('kpiMaxWait').textContent = formatSecToMMSS(maxSec);
      document.getElementById('kpiTotalRecords').textContent = Number(filtered.length).toLocaleString();
      document.getElementById('tableCounter').textContent = `แสดง ${filtered.length} รายการ`;

      renderChart(filtered);
    }

    function renderChart(data) {
      const dayMap = {};
      data.forEach(d => {
        const day = d['วันที่'];
        const waitSec = parseMMSSToSec(d['Waiting']);
        if (day && waitSec > 0) {
          if (!dayMap[day]) dayMap[day] = { sum: 0, count: 0 };
          dayMap[day].sum += waitSec;
          dayMap[day].count += 1;
        }
      });

      const sortedDays = Object.keys(dayMap).sort((a,b) => a - b);
      const avgWaitPerDaySec = sortedDays.map(d => {
        const item = dayMap[d];
        return item.count > 0 ? (item.sum / item.count) : 0;
      });

      const ctx = document.getElementById('trendChart').getContext('2d');
      if (trendChartInstance) trendChartInstance.destroy();

      trendChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: sortedDays.map(d => `วันที่ ${d}`),
          datasets: [{
            label: 'เวลารอคอยเฉลี่ย (วินาที)',
            data: avgWaitPerDaySec,
            backgroundColor: 'rgba(37, 99, 235, 0.7)',
            borderColor: 'rgb(37, 99, 235)',
            borderWidth: 1,
            borderRadius: 4
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { 
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: function(context) {
                  return 'เฉลี่ย: ' + formatSecToMMSS(context.parsed.y) + ' (MM:SS)';
                }
              }
            }
          },
          scales: {
            y: { 
              beginAtZero: true, 
              title: { display: true, text: 'วินาที' },
              ticks: {
                callback: function(val) {
                  return formatSecToMMSS(val);
                }
              }
            }
          }
        }
      });
    }

    window.addEventListener('DOMContentLoaded', () => {
      initFilters();
      renderDashboard();
    });
  </script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_template)

print("="*50)
print(" สำเร็จ 100%! ล็อกให้ Taxi มีเฉพาะใน Domestic เรียบร้อยแล้ว")
print("="*50)