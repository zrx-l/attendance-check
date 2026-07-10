import sys
import io
# 强制行缓冲，确保所有 print 实时输出
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)

import pandas as pd
import win32com.client as win32
from pathlib import Path
from datetime import datetime, timedelta
import re
import argparse

# ================== 颜色常量 ==================
GREEN_RGB = 65280       # RGB(0,255,0)
GRAY_RGB = 8421504      # RGB(128,128,128)
WHITE_RGB = 16777215    # RGB(255,255,255) 白色（离职人员）

# ================== 打卡日报列映射 ==================
CHECKIN_SHEET = 1
CHECKIN_COL_DATE = 0
CHECKIN_COL_ACCOUNT = 2
CHECKIN_COL_TYPE = 6
CHECKIN_COL_TIME = 8
CHECKIN_COL_STATUS = 9

HIDE_JOB_LEVELS = ["总监（业务）", "高级经理（业务）", "经理（业务）", "副经理（业务）"]

def parse_leave_data(file_path):
    """解析休假文件，返回 (emp_id, date) -> (leave_type, total_hours, start, end)"""
    df = pd.read_excel(file_path, header=0, skiprows=[1], dtype=str)
    df.columns = df.columns.str.strip()
    leaves = {}
    for _, row in df.iterrows():
        emp_id = str(row.iloc[4]).strip()
        date_range = row.iloc[8]
        if pd.isna(date_range):
            continue
        try:
            if "至" in date_range:
                start_str, end_str = date_range.split("至")
                start = pd.to_datetime(start_str.strip()).date()
                end = pd.to_datetime(end_str.strip()).date()
            else:
                start = end = pd.to_datetime(date_range.strip()).date()
        except:
            continue
        leave_type = row.iloc[7]
        total_hours = float(row.iloc[9])
        delta = (end - start).days + 1
        for i in range(delta):
            date = start + timedelta(days=i)
            leaves[(emp_id, date)] = (leave_type, total_hours, start, end)
    return leaves

def parse_checkin_data(file_path):
    df = pd.read_excel(file_path, sheet_name=CHECKIN_SHEET, dtype=str)
    checkins = {}
    for _, row in df.iterrows():
        date_str = row.iloc[CHECKIN_COL_DATE]
        try:
            date = pd.to_datetime(date_str.split()[0]).date()
        except:
            continue
        emp_id = str(row.iloc[CHECKIN_COL_ACCOUNT]).strip()
        punch_type = row.iloc[CHECKIN_COL_TYPE]
        punch_time = row.iloc[CHECKIN_COL_TIME]
        status = row.iloc[CHECKIN_COL_STATUS] if CHECKIN_COL_STATUS < len(row) else ""
        key = (emp_id, date)
        if key not in checkins:
            checkins[key] = {"上班": "", "下班": "", "外出": [], "异常": ""}
        if pd.notna(punch_time) and str(punch_time).strip():
            time_str = str(punch_time).strip()
            if time_str == "--":
                time_str = ""
            else:
                match = re.search(r'\d{1,2}:\d{2}', time_str)
                if match:
                    time_str = match.group()
        else:
            time_str = ""
        if punch_type == "上班":
            checkins[key]["上班"] = time_str
        elif punch_type == "下班":
            checkins[key]["下班"] = time_str
        elif "外出" in punch_type:
            if time_str:
                checkins[key]["外出"].append(time_str)
        if status and "缺卡" in status:
            checkins[key]["异常"] = status
    return checkins

def parse_remote_data(file_path):
    """解析远程办公申请文件，返回 {(emp_id, date): (申请类型, 地点)}"""
    try:
        df = pd.read_excel(file_path, header=0, dtype=str)
        remote_dict = {}
        for _, row in df.iterrows():
            emp_id = str(row.iloc[1]).strip()
            date_range = row.iloc[6]
            leave_type = row.iloc[7]
            location = row.iloc[8]
            if pd.isna(date_range) or pd.isna(emp_id):
                continue
            try:
                if " - " in date_range:
                    start_str, end_str = date_range.split(" - ")
                    start = pd.to_datetime(start_str.strip()).date()
                    end = pd.to_datetime(end_str.strip()).date()
                else:
                    start = end = pd.to_datetime(date_range.strip()).date()
            except:
                continue
            delta = (end - start).days + 1
            for i in range(delta):
                date = start + timedelta(days=i)
                remote_dict[(emp_id, date)] = (leave_type, location)
        return remote_dict
    except Exception as e:
        print(f"读取远程办公文件失败: {e}")
        return {}

def get_date_from_cell(cell_value):
    if isinstance(cell_value, datetime):
        return cell_value.date()
    s = str(cell_value).strip()
    if '.' in s:
        s = s.split('.')[0]
    if len(s) == 8 and s.isdigit():
        return datetime.strptime(s, "%Y%m%d").date()
    elif '-' in s:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    return None

def is_weekend(date):
    return date.weekday() >= 5

def generate_cell_text(emp_id, date, leaves, checkins):
    """
    生成基础文本（非休假场景）
    注意：此函数中的 leaves 仅用于判断，实际休假已由主函数处理，此处传空字典即可
    """
    if (emp_id, date) in leaves:
        leave_type, hours = leaves[(emp_id, date)]
        text = f"{leave_type}{hours}h"
        if (emp_id, date) in checkins:
            cinfo = checkins[(emp_id, date)]
            times = []
            if cinfo["上班"]:
                times.append(f"上班{cinfo['上班']}")
            if cinfo["下班"]:
                times.append(f"下班{cinfo['下班']}")
            if cinfo["外出"]:
                times.extend([f"外出{t}" for t in cinfo["外出"]])
            if times:
                text += " " + " ".join(times)
        return text
    if (emp_id, date) in checkins:
        cinfo = checkins[(emp_id, date)]
        has_上班 = bool(cinfo["上班"])
        has_下班 = bool(cinfo["下班"])
        has_外出 = len(cinfo["外出"]) > 0
        if not has_上班 and not has_下班:
            text = "缺卡2次"
            if has_外出:
                text += " " + " ".join([f"外出{t}" for t in cinfo["外出"]])
            return text
        parts = []
        if has_上班:
            parts.append(f"上班{cinfo['上班']}")
        else:
            parts.append("缺卡1次")
        if has_下班:
            parts.append(f"下班{cinfo['下班']}")
        else:
            parts.append("缺卡1次")
        if has_外出:
            parts.extend([f"外出{t}" for t in cinfo["外出"]])
        return " ".join(parts)
    return "缺卡2次"

def process_template_win32com(template_path, leaves, checkins, remote_dict, output_file):
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    wb = None
    try:
        wb = excel.Workbooks.Open(template_path)
        ws = wb.Worksheets("报表区")
        
        try:
            ws.Unprotect()
        except:
            pass
        
        # 识别日期列
        date_row = 4
        first_date_col = 16
        date_cols = {}
        col = first_date_col
        empty_count = 0
        max_empty = 3
        while True:
            cell_value = ws.Cells(date_row, col).Value
            if cell_value is None or str(cell_value).strip() == "":
                empty_count += 1
                if empty_count >= max_empty:
                    break
                col += 1
                continue
            empty_count = 0
            date = get_date_from_cell(cell_value)
            if date:
                date_cols[col] = date
            col += 1
        
        if not date_cols:
            print("错误：未能识别日期列，请检查模板！")
            sys.stdout.flush()
            return
        print(f"识别到的日期列数: {len(date_cols)}")
        sys.stdout.flush()
        
        # 动态识别实际数据行数
        last_row = 6
        for r in range(6, 501):
            val = ws.Cells(r, 2).Value
            if val is not None and str(val).strip() != "":
                last_row = r
            else:
                empty_count = 0
                for i in range(r, min(r+5, 501)):
                    v = ws.Cells(i, 2).Value
                    if v is None or str(v).strip() == "":
                        empty_count += 1
                    else:
                        break
                if empty_count >= 5:
                    break
        max_row = min(last_row + 2, 500)
        print(f"扫描行数: 6 到 {max_row}（动态识别）")
        sys.stdout.flush()
        
        rows_to_hide = []
        modified_count = 0
        SKIP_COLORS = {65280, 8421504, 12632256, 16777215}
        red_cells = []
        
        for row in range(6, max_row + 1):
            emp_cell = ws.Cells(row, 2)
            emp_id = None
            try:
                if emp_cell.MergeCells:
                    emp_id = emp_cell.MergeArea.Cells(1, 1).Value
                else:
                    emp_id = emp_cell.Value
            except:
                emp_id = emp_cell.Value
            
            if not emp_id or str(emp_id).strip() == "":
                continue
            emp_id = str(emp_id).strip()
            
            job_level = ws.Cells(row, 4).Value
            if job_level:
                job_level_str = str(job_level).strip()
                if any(level in job_level_str for level in HIDE_JOB_LEVELS):
                    rows_to_hide.append(row)
                    continue
            
            emp_leave_records = []
            for (eid, date), (leave_type, total_hours, start, end) in leaves.items():
                if eid == emp_id:
                    emp_leave_records.append((leave_type, total_hours, start, end))
            unique_records = []
            seen = set()
            for rec in emp_leave_records:
                key = (rec[1], rec[2], rec[3])
                if key not in seen:
                    seen.add(key)
                    unique_records.append(rec)
            
            leave_assignment = {}
            for leave_type, total_hours, start, end in unique_records:
                workday_count = 0
                for col2, date2 in date_cols.items():
                    if start <= date2 <= end:
                        cell_color = ws.Cells(row, col2).Interior.Color
                        if cell_color is not None:
                            color_val = int(cell_color)
                            if color_val not in SKIP_COLORS:
                                workday_count += 1
                        else:
                            workday_count += 1
                if workday_count > 0:
                    daily_hours = total_hours / workday_count
                    for col2, date2 in date_cols.items():
                        if start <= date2 <= end:
                            cell_color = ws.Cells(row, col2).Interior.Color
                            if cell_color is not None:
                                color_val = int(cell_color)
                                if color_val not in SKIP_COLORS:
                                    leave_assignment[(emp_id, date2)] = (leave_type, daily_hours)
                            else:
                                leave_assignment[(emp_id, date2)] = (leave_type, daily_hours)
            
            for col, date in date_cols.items():
                cell = ws.Cells(row, col)
                color_raw = cell.Interior.Color
                if color_raw is None:
                    continue
                color = int(color_raw)
                if color in SKIP_COLORS:
                    continue
                if color == 255:
                    red_cells.append((row, col))
                
                has_remote = remote_dict and (emp_id, date) in remote_dict
                remote_suffix = ""
                if has_remote:
                    leave_type, location = remote_dict[(emp_id, date)]
                    if pd.isna(leave_type) or str(leave_type).strip() == "":
                        leave_type = "远程工作"
                    remote_suffix = f" {leave_type}（{location}）"
                
                if (emp_id, date) in leave_assignment:
                    leave_type, daily_hours = leave_assignment[(emp_id, date)]
                    text = f"{leave_type}{daily_hours}h"
                    if (emp_id, date) in checkins:
                        cinfo = checkins[(emp_id, date)]
                        times = []
                        if cinfo["上班"]:
                            times.append(f"上班{cinfo['上班']}")
                        if cinfo["下班"]:
                            times.append(f"下班{cinfo['下班']}")
                        if cinfo["外出"]:
                            times.extend([f"外出{t}" for t in cinfo["外出"]])
                        if times:
                            text += " " + " ".join(times)
                    if has_remote:
                        text += remote_suffix
                    cell.Value = text
                    modified_count += 1
                    continue
                
                text = generate_cell_text(emp_id, date, {}, checkins)
                if text is not None:
                    if has_remote:
                        cell.Value = text + remote_suffix
                    else:
                        cell.Value = text
                    modified_count += 1
        
        # 隐藏职级行
        for row in rows_to_hide:
            try:
                ws.Rows(row).Hidden = True
            except:
                pass
        
        # 设置列宽
        ws.Activate()
        for col in date_cols.keys():
            try:
                ws.Columns(col).AutoFit()
                ws.Columns(col).WrapText = True
            except:
                pass
        
        # 创建异常数据区（保留格式）
        ws_new = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws_new.Name = "异常数据区"
        
        red_rows = set()
        red_cols = set()
        for (r, c) in red_cells:
            red_rows.add(r)
            red_cols.add(c)
        
        max_col = max(date_cols.keys()) if date_cols else 46
        source_range = ws.Range(ws.Cells(1, 1), ws.Cells(max_row, max_col))
        source_range.Copy(ws_new.Cells(1, 1))
        
        for row in range(6, max_row + 1):
            for col in date_cols.keys():
                cell = ws_new.Cells(row, col)
                color_raw = cell.Interior.Color
                if color_raw is not None and int(color_raw) == 255:
                    continue
                else:
                    cell.ClearContents()
        
        rows_with_red = red_rows
        rows_to_delete = []
        for row in range(6, max_row + 1):
            if row not in rows_with_red:
                rows_to_delete.append(row)
        for row in sorted(rows_to_delete, reverse=True):
            try:
                ws_new.Rows(row).Delete()
            except:
                pass
        
        cols_with_red = red_cols
        cols_to_delete = []
        for col in date_cols.keys():
            if col not in cols_with_red:
                cols_to_delete.append(col)
        for col in sorted(cols_to_delete, reverse=True):
            try:
                ws_new.Columns(col).Delete()
            except:
                pass
        
        # 保存工作簿
        output_file.parent.mkdir(parents=True, exist_ok=True)
        wb.SaveAs(str(output_file))
        
        print(f"已生成考勤明细: {output_file}")
        print(f"隐藏职级行数: {len(rows_to_hide)}")
        print(f"实际修改单元格数: {modified_count}")
        print(f"异常数据区行数: {len(red_rows)} 行, 列数: {len(red_cols)} 列")
        sys.stdout.flush()
        
    except Exception as e:
        print(f"处理模板时出错: {e}")
        sys.stdout.flush()
        raise
    finally:
        if wb:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        try:
            excel.Quit()
        except:
            pass
        import gc
        gc.collect()

def run_attendance_check(start_date, end_date, department, template_file, checkin_file, leave_file, remote_file=None):
    print("正在读取休假数据...")
    sys.stdout.flush()
    leaves = parse_leave_data(leave_file)
    print(f"  共 {len(leaves)} 条请假记录（按天展开）")
    sys.stdout.flush()
    print("正在读取打卡数据...")
    sys.stdout.flush()
    checkins = parse_checkin_data(checkin_file)
    print(f"  共 {len(checkins)} 个员工-日期打卡记录")
    sys.stdout.flush()
    remote_dict = {}
    if remote_file and Path(remote_file).exists():
        print("正在读取远程办公数据...")
        sys.stdout.flush()
        remote_dict = parse_remote_data(remote_file)
        print(f"  共 {len(remote_dict)} 条远程办公记录")
        sys.stdout.flush()
    else:
        print("未提供远程办公文件，跳过")
        sys.stdout.flush()
    print("正在处理考勤模板...")
    sys.stdout.flush()
    
    base_dir = Path(__file__).parent
    output_dir = base_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    start = datetime.strptime(start_date, "%Y-%m-%d")
    end = datetime.strptime(end_date, "%Y-%m-%d")
    
    start_str = f"{start.year}.{start.month}.{start.day}"
    end_str = f"{end.year}.{end.month}.{end.day}"
    
    filename = f"{department}考勤（{start_str}-{end_str}）.xlsx"
    output_file = output_dir / filename
    
    process_template_win32com(template_file, leaves, checkins, remote_dict, output_file)
    print("处理完成")
    sys.stdout.flush()

def main():
    parser = argparse.ArgumentParser(description="考勤核对工具")
    parser.add_argument("--start", required=True, help="起始日期，如 2026-04-01")
    parser.add_argument("--end", required=True, help="截止日期，如 2026-04-30")
    parser.add_argument("--dept", default="工程造价一部", help="部门名称")
    parser.add_argument("--template", required=True, help="考勤模板文件路径")
    parser.add_argument("--checkin", required=True, help="打卡日报文件路径")
    parser.add_argument("--leave", required=True, help="休假数据文件路径")
    parser.add_argument("--remote", default=None, help="远程办公文件路径（可选）")
    args = parser.parse_args()
    run_attendance_check(args.start, args.end, args.dept, args.template, args.checkin, args.leave, args.remote)

if __name__ == "__main__":
    main()