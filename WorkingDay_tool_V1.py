import streamlit as st
import pandas as pd
from datetime import datetime, timedelta, date
import re
import os
from typing import List, Tuple, Dict, Optional, Set
import tempfile
import io
import zipfile
from io import BytesIO

# ==================== 配置常量 ====================
YEAR = 2026
DATE_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%d/%m/%Y"]
COLUMN_MAPPING = {
    "姓名": ["姓名", "name", "学生姓名"],
    "学生休假开始日期": ["开始日期", "休假开始", "start", "开始"],
    "学生休假结束日期": ["结束日期", "休假结束", "end", "结束"],
}
INVALID_FILENAME_CHARS = r'[\\/*?:"<>|]'

# ==================== 工具函数 ====================
def format_date_list(dates: List[date]) -> str:
    return ', '.join(d.strftime('%Y-%m-%d') for d in dates)

def safe_filename(name: str) -> str:
    return re.sub(INVALID_FILENAME_CHARS, '_', name)

def parse_date_robust(val):
    if pd.isna(val):
        return None
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, date) else val.date()
    if isinstance(val, (int, float)):
        try:
            return pd.to_datetime(val, unit='D', origin='1899-12-30').date()
        except:
            pass
    s = str(val).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    try:
        return pd.to_datetime(s).date()
    except:
        return None

# ==================== HolidayManager ====================
class HolidayManager:
    def __init__(self, year: int = YEAR):
        self.year = year
        self.holidays: Set[date] = set()
        self.makeup_days: Set[date] = set()
        self.names: Dict[date, str] = {}
        self._load_default()

    def _load_default(self):
        default_holidays = {
            date(2026, 1, 1): "元旦", date(2026, 1, 2): "元旦", date(2026, 1, 3): "元旦",
            date(2026, 2, 15): "春节", date(2026, 2, 16): "春节", date(2026, 2, 17): "春节",
            date(2026, 2, 18): "春节", date(2026, 2, 19): "春节", date(2026, 2, 20): "春节",
            date(2026, 2, 21): "春节", date(2026, 2, 22): "春节", date(2026, 2, 23): "春节",
            date(2026, 4, 4): "清明", date(2026, 4, 5): "清明", date(2026, 4, 6): "清明",
            date(2026, 5, 1): "劳动节", date(2026, 5, 2): "劳动节", date(2026, 5, 3): "劳动节",
            date(2026, 5, 4): "劳动节", date(2026, 5, 5): "劳动节",
            date(2026, 6, 19): "端午", date(2026, 6, 20): "端午", date(2026, 6, 21): "端午",
            date(2026, 9, 25): "中秋", date(2026, 9, 26): "中秋", date(2026, 9, 27): "中秋",
            date(2026, 10, 1): "国庆", date(2026, 10, 2): "国庆", date(2026, 10, 3): "国庆",
            date(2026, 10, 4): "国庆", date(2026, 10, 5): "国庆", date(2026, 10, 6): "国庆",
            date(2026, 10, 7): "国庆",
        }
        default_makeup = {
            date(2026, 1, 4): "元旦调休",
            date(2026, 2, 14): "春节调休",
            date(2026, 2, 28): "春节调休",
            date(2026, 5, 9): "劳动节调休",
            date(2026, 9, 20): "国庆调休",
            date(2026, 10, 10): "国庆调休",
        }
        self.holidays = set(default_holidays.keys())
        self.makeup_days = set(default_makeup.keys())
        self.names = {**default_holidays, **default_makeup}

    def is_workday(self, d: date) -> bool:
        if d in self.holidays:
            return False
        if d.weekday() >= 5 and d not in self.makeup_days:
            return False
        return True

    def classify_date(self, d: date) -> Tuple[str, Optional[str]]:
        if d in self.holidays:
            return 'holiday', self.names.get(d, '法定假日')
        if d.weekday() >= 5 and d in self.makeup_days:
            return 'makeup', self.names.get(d, '调休')
        if d.weekday() >= 5:
            return 'weekend', None
        return 'workday', None

    def import_from_file(self, file_path: str) -> Tuple[int, int, int]:
        try:
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path, encoding='utf-8-sig')
            else:
                df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as e:
            raise ValueError(f"读取文件失败: {e}")

        required = ['类型', '日期']
        if not all(col in df.columns for col in required):
            raise ValueError("文件必须包含 '类型' 和 '日期' 列（可包含 '名称' 列）")

        new_holidays = set()
        new_makeup = set()
        new_names = {}
        skipped = 0

        for idx, row in df.iterrows():
            try:
                typ = str(row['类型']).strip()
                date_str = str(row['日期']).strip()
                d = parse_date_robust(date_str)
                if d is None:
                    skipped += 1
                    continue
                if d.year != self.year:
                    skipped += 1
                    continue
                name = ""
                if '名称' in df.columns and pd.notna(row['名称']):
                    name = str(row['名称']).strip()
                if typ == '法定节假日':
                    new_holidays.add(d)
                    if name:
                        new_names[d] = name
                elif typ == '调休日':
                    new_makeup.add(d)
                    if name:
                        new_names[d] = name
                else:
                    skipped += 1
            except Exception:
                skipped += 1

        self.holidays = new_holidays
        self.makeup_days = new_makeup
        self.names = new_names
        return len(new_holidays), len(new_makeup), skipped

    def export_to_file(self, file_path: str):
        holiday_items = [(d, self.names.get(d, '')) for d in sorted(self.holidays)]
        makeup_items = [(d, self.names.get(d, '调休')) for d in sorted(self.makeup_days)]
        all_items = [('法定节假日', d, name) for d, name in holiday_items] + \
                    [('调休日', d, name) for d, name in makeup_items]
        all_items.sort(key=lambda x: x[1])

        rows = []
        for idx, (typ, d, name) in enumerate(all_items, start=1):
            rows.append([idx, typ, name, d.strftime('%Y-%m-%d')])

        df = pd.DataFrame(rows, columns=['序号', '类型', '名称', '日期'])
        try:
            if file_path.lower().endswith('.csv'):
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                df.to_excel(file_path, index=False, engine='openpyxl')
        except Exception as e:
            raise ValueError(f"导出失败: {e}")

    def get_summary(self) -> str:
        return f"法定节假日: {len(self.holidays)} 天, 调休日: {len(self.makeup_days)} 天"

    def to_dataframe(self) -> pd.DataFrame:
        items = []
        for d in self.holidays:
            items.append(('法定节假日', self.names.get(d, ''), d))
        for d in self.makeup_days:
            items.append(('调休日', self.names.get(d, '调休'), d))
        items.sort(key=lambda x: x[2])
        data = []
        for idx, (typ, name, d) in enumerate(items, start=1):
            data.append([idx, typ, name, d.strftime('%Y-%m-%d')])
        return pd.DataFrame(data, columns=['序号', '类型', '名称', '日期'])

# ==================== StatisticsEngine（修改核心逻辑） ====================
class StatisticsEngine:
    def __init__(self, holiday_mgr: HolidayManager):
        self.holiday_mgr = holiday_mgr

    def collect_date_categories_v2(self, start_date: date, end_date: date) -> Dict:
        """
        统计区间内：
        - full_months: 完整自然月数（该月所有天数都在区间内）
        - workdays: 非完整月份中的工作日天数（完整月不计入）
        - 各类日期列表（包含完整月和非完整月的所有日期）
        """
        if start_date > end_date:
            raise ValueError("开始日期不能晚于结束日期")

        full_months = 0
        workday_count = 0
        holidays = []
        weekends = []
        makeups = []

        cur = date(start_date.year, start_date.month, 1)
        while cur <= end_date:
            # 计算当月最后一天
            if cur.month == 12:
                next_month = date(cur.year + 1, 1, 1)
            else:
                next_month = date(cur.year, cur.month + 1, 1)
            last_day = next_month - timedelta(days=1)

            # 判断该月是否完全包含在区间内
            if start_date <= cur and end_date >= last_day:
                # 完整自然月：只计数，不累加工作日
                full_months += 1
                d = cur
                while d <= last_day:
                    cat, _ = self.holiday_mgr.classify_date(d)
                    if cat == 'holiday':
                        holidays.append(d)
                    elif cat == 'weekend':
                        weekends.append(d)
                    elif cat == 'makeup':
                        makeups.append(d)
                    d += timedelta(days=1)
            else:
                # 非完整月：遍历区间覆盖的天数，累加工作日
                month_start = max(start_date, cur)
                month_end = min(end_date, last_day)
                d = month_start
                while d <= month_end:
                    if self.holiday_mgr.is_workday(d):
                        workday_count += 1
                    cat, _ = self.holiday_mgr.classify_date(d)
                    if cat == 'holiday':
                        holidays.append(d)
                    elif cat == 'weekend':
                        weekends.append(d)
                    elif cat == 'makeup':
                        makeups.append(d)
                    d += timedelta(days=1)

            cur = next_month

        return {
            'full_months': full_months,
            'workdays': workday_count,
            'holidays': sorted(holidays),
            'weekends': sorted(weekends),
            'makeups': sorted(makeups),
        }

    def process_file(self, file_path: str) -> Tuple[str, pd.DataFrame, List[str]]:
        errors = []
        try:
            if file_path.lower().endswith(('.xlsx', '.xls')):
                df_raw = pd.read_excel(file_path, engine='openpyxl')
            else:
                df_raw = pd.read_csv(file_path, encoding='utf-8-sig')
        except Exception as e:
            raise ValueError(f"读取文件失败: {e}")

        df_raw.columns = [str(c).strip() for c in df_raw.columns]

        col_map = {}
        for target, possible in COLUMN_MAPPING.items():
            found = None
            for col in df_raw.columns:
                if col in possible:
                    found = col
                    break
            if found is None:
                for col in df_raw.columns:
                    for p in possible:
                        if p in col:
                            found = col
                            break
                    if found:
                        break
            if found is None:
                raise ValueError(f"未找到列: {target} (候选: {possible})")
            col_map[target] = found

        df_sel = df_raw[[col_map['姓名'], col_map['学生休假开始日期'], col_map['学生休假结束日期']]].copy()
        df_sel.columns = ['姓名', '开始日期', '结束日期']

        df_sel['开始日期'] = df_sel['开始日期'].apply(parse_date_robust)
        df_sel['结束日期'] = df_sel['结束日期'].apply(parse_date_robust)

        invalid = df_sel[df_sel['开始日期'].isna() | df_sel['结束日期'].isna()]
        if not invalid.empty:
            errors.append(f"跳过 {len(invalid)} 行因日期无效")
            df_sel = df_sel.dropna(subset=['开始日期', '结束日期'])

        results = []
        for idx, row in df_sel.iterrows():
            try:
                stats = self.collect_date_categories_v2(row['开始日期'], row['结束日期'])
                results.append({
                    '姓名': row['姓名'],
                    '开始日期': row['开始日期'],
                    '结束日期': row['结束日期'],
                    '完整自然月数': stats['full_months'],
                    '工作日天数': stats['workdays'],
                    '调休日日期': format_date_list(stats['makeups']),
                    '节假日日期': format_date_list(stats['holidays']),
                    '周末日期': format_date_list(stats['weekends']),
                })
            except ValueError as e:
                errors.append(f"第{idx+2}行错误: {e}")

        df_result = pd.DataFrame(results)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        return file_name, df_result, errors

# ==================== Streamlit 应用 ====================
def init_session_state():
    if 'holiday_mgr' not in st.session_state:
        st.session_state.holiday_mgr = HolidayManager()
    if 'files' not in st.session_state:
        st.session_state.files = []          # 元素: {'name': str, 'data': bytes}
    if 'results' not in st.session_state:
        st.session_state.results = {}
    if 'all_results_df' not in st.session_state:
        st.session_state.all_results_df = pd.DataFrame()
    if 'processing' not in st.session_state:
        st.session_state.processing = False
    if 'error_messages' not in st.session_state:
        st.session_state.error_messages = []
    if 'uploader_key' not in st.session_state:
        st.session_state.uploader_key = 0

def main():
    st.set_page_config(page_title="休假期间工作日统计工具 (2026)", layout="wide")
    init_session_state()

    st.markdown("""
    <style>
        .main-title {
            color: #1f77b4;
            font-size: 2.5rem;
            font-weight: bold;
            margin-bottom: 1rem;
        }
        .footer {
            text-align: center;
            color: #6c757d;
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid #dee2e6;
        }
        .sidebar-footer {
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid #dee2e6;
            color: #6c757d;
            font-size: 0.9rem;
        }
        .stButton button {
            background-color: #1f77b4;
            color: white;
            border-radius: 5px;
            border: none;
            padding: 0.5rem 1rem;
            font-weight: bold;
        }
        .stButton button:hover {
            background-color: #145a8a;
            color: white;
        }
        .stButton button:disabled {
            background-color: #cccccc;
            color: #666666;
        }
    </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.header("📅 节假日配置")
        st.write(st.session_state.holiday_mgr.get_summary())

        if st.button("预览当前配置"):
            df_config = st.session_state.holiday_mgr.to_dataframe()
            if df_config.empty:
                st.info("当前无配置数据")
            else:
                st.dataframe(df_config, use_container_width=True)

        st.divider()
        st.subheader("导入节假日配置")
        imported_file = st.file_uploader("选择节假日配置文件 (CSV/Excel)", type=['csv', 'xlsx', 'xls'],
                                         key="import_holiday")
        if imported_file is not None:
            suffix = '.csv' if imported_file.name.endswith('.csv') else '.xlsx'
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp.write(imported_file.getvalue())
                tmp_path = tmp.name
            try:
                h, m, skipped = st.session_state.holiday_mgr.import_from_file(tmp_path)
                st.success(f"导入成功！法定 {h} 天，调休 {m} 天" + (f"，跳过 {skipped} 行" if skipped else ""))
                st.rerun()
            except Exception as e:
                st.error(f"导入失败: {e}")
            finally:
                os.unlink(tmp_path)

        st.subheader("导出节假日配置")
        export_format = st.radio("导出格式", ["CSV", "Excel"], horizontal=True)
        if st.button("导出节假日配置"):
            suffix = '.csv' if export_format == "CSV" else '.xlsx'
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                tmp_path = tmp.name
            try:
                st.session_state.holiday_mgr.export_to_file(tmp_path)
                with open(tmp_path, 'rb') as f:
                    data = f.read()
                st.download_button(
                    label="📥 下载节假日配置文件",
                    data=data,
                    file_name=f"holiday_config_{YEAR}{suffix}",
                    mime="text/csv" if suffix == '.csv' else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"导出失败: {e}")
            finally:
                os.unlink(tmp_path)

        st.sidebar.divider()
        st.sidebar.markdown('<div class="sidebar-footer">Author：Alan &nbsp;|&nbsp; Thanks：Jiao Xu</div>', unsafe_allow_html=True)

    st.markdown('<div class="main-title">🐮  休假期间工作日统计工具 (2026) 🐴</div>', unsafe_allow_html=True)

    uploader_key = f"uploader_{st.session_state.uploader_key}"
    uploaded_files = st.file_uploader(
        "添加数据文件 (Excel/CSV)",
        type=['xlsx', 'xls', 'csv'],
        accept_multiple_files=True,
        key=uploader_key
    )

    if uploaded_files:
        for file in uploaded_files:
            file_name = str(file.name)
            if not any(f['name'] == file_name for f in st.session_state.files):
                st.session_state.files.append({
                    'name': file_name,
                    'data': file.getvalue()
                })
        st.session_state.uploader_key += 1
        st.rerun()

    if st.session_state.files:
        st.subheader(f"已添加文件 ({len(st.session_state.files)})")
        cols = st.columns([4, 1, 1])
        cols[0].write("**文件名**")
        cols[1].write("**状态**")
        cols[2].write("**操作**")
        for i, file_dict in enumerate(st.session_state.files):
            cols = st.columns([4, 1, 1])
            cols[0].write(file_dict['name'])
            if file_dict['name'] in st.session_state.results:
                cols[1].write("✅ 已统计")
            else:
                cols[1].write("⏳ 待统计")
            if cols[2].button("🗑️ 删除", key=f"del_{i}"):
                del st.session_state.files[i]
                if file_dict['name'] in st.session_state.results:
                    del st.session_state.results[file_dict['name']]
                    all_dfs = []
                    for name, df in st.session_state.results.items():
                        df_copy = df.copy()
                        df_copy.insert(0, '文件名', name)
                        all_dfs.append(df_copy)
                    st.session_state.all_results_df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
                st.rerun()
        st.divider()
    else:
        st.info("请上传数据文件")

    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        if st.button("🚀 开始统计", disabled=st.session_state.processing or not st.session_state.files):
            st.session_state.processing = True
            st.session_state.results = {}
            st.session_state.all_results_df = pd.DataFrame()
            st.session_state.error_messages = []

            progress_bar = st.progress(0, text="准备统计...")
            total = len(st.session_state.files)
            engine = StatisticsEngine(st.session_state.holiday_mgr)

            for idx, file_dict in enumerate(st.session_state.files):
                progress_bar.progress((idx) / total, text=f"正在处理: {file_dict['name']} ({idx+1}/{total})")
                suffix = '.csv' if file_dict['name'].endswith('.csv') else '.xlsx'
                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(file_dict['data'])
                    tmp_path = tmp.name
                try:
                    file_name, df_result, errors = engine.process_file(tmp_path)
                    if not df_result.empty:
                        base_name = os.path.splitext(file_dict['name'])[0]
                        st.session_state.results[base_name] = df_result
                        df_copy = df_result.copy()
                        df_copy.insert(0, '文件名', base_name)
                        st.session_state.all_results_df = pd.concat(
                            [st.session_state.all_results_df, df_copy], ignore_index=True
                        )
                    if errors:
                        st.session_state.error_messages.extend(errors)
                except Exception as e:
                    st.session_state.error_messages.append(f"{file_dict['name']}: {e}")
                finally:
                    os.unlink(tmp_path)

            progress_bar.progress(1.0, text="统计完成！")
            st.session_state.processing = False
            st.success(f"✅ 统计完成！共处理 {len(st.session_state.files)} 个文件，{len(st.session_state.all_results_df)} 条记录。")
            st.rerun()

    with col2:
        if st.button("📊 导出结果", disabled=st.session_state.all_results_df.empty):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                for sheet_name, df in st.session_state.results.items():
                    clean = re.sub(r'[\\/*?:\[\]]', '_', sheet_name)[:31]
                    if clean in writer.sheets:
                        suffix = 1
                        while f"{clean}_{suffix}" in writer.sheets:
                            suffix += 1
                        clean = f"{clean}_{suffix}"
                    df.to_excel(writer, sheet_name=clean, index=False)
            output.seek(0)
            st.download_button(
                label="📥 下载汇总 Excel",
                data=output,
                file_name="汇总统计结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w') as zf:
                for sheet_name, df in st.session_state.results.items():
                    safe_name = safe_filename(sheet_name)
                    sub_name = f"{safe_name}_Result.xlsx"
                    sub_output = BytesIO()
                    df.to_excel(sub_output, index=False, engine='openpyxl')
                    sub_output.seek(0)
                    zf.writestr(sub_name, sub_output.getvalue())
            zip_buffer.seek(0)
            st.download_button(
                label="📥 下载单独文件 (zip包)",
                data=zip_buffer,
                file_name="各文件结果.zip",
                mime="application/zip"
            )

    if st.session_state.error_messages:
        with st.expander("⚠️ 统计警告/错误"):
            for msg in st.session_state.error_messages:
                st.warning(msg)

    if not st.session_state.all_results_df.empty:
        st.subheader(f"📋 统计结果 ({len(st.session_state.all_results_df)} 条记录)")
        display_df = st.session_state.all_results_df.copy()
        if '开始日期' in display_df.columns:
            display_df['开始日期'] = display_df['开始日期'].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '')
            display_df['结束日期'] = display_df['结束日期'].apply(lambda x: x.strftime('%Y-%m-%d') if pd.notna(x) else '')
        st.dataframe(display_df, use_container_width=True)

   
if __name__ == "__main__":
    main()
