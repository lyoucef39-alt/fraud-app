import streamlit as st
import pandas as pd
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# --- إعدادات الواجهة ---
st.set_page_config(page_title="Rapport A79 - Youcef", layout="wide")
st.markdown("<h1 style='text-align: center; color: #2F5597;'>📊 منصة معالجة تقارير الغش A79</h1>", unsafe_allow_html=True)

# --- التنسيقات ---
months_fr = {1:"Janvier", 2:"Février", 3:"Mars", 4:"Avril", 5:"Mai", 6:"Juin", 
             7:"Juillet", 8:"Août", 9:"Septembre", 10:"Octobre", 11:"Novembre", 12:"Décembre"}
header_fill = PatternFill(start_color="2F5597", fill_type="solid")
sub_fill = PatternFill(start_color="4472C4", fill_type="solid")
total_fill = PatternFill(start_color="D9E1F2", fill_type="solid")
std_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def run_mega_process():
    st.sidebar.header("📁 رفع الملفات")
    v_files = st.sidebar.file_uploader("ملفات Detail_Ventes", accept_multiple_files=True)
    c_files = st.sidebar.file_uploader("ملفات Detail_CA_Energie", accept_multiple_files=True)

    if st.sidebar.button("🚀 معالجة البيانات"):
        if not v_files or not c_files:
            st.error("الرجاء رفع الملفات من الجهة اليسرى أولاً!")
            return

        all_data_list = []
        v_dict = {f.name: f for f in v_files}
        c_dict = {f.name: f for f in c_files}

        for energy in ['BT', 'BP']:
            v_f = next((f for n, f in v_dict.items() if energy in n.upper()), None)
            c_f = next((f for n, f in c_dict.items() if energy in n.upper()), None)
            
            if v_f and c_f:
                df_v = pd.read_excel(v_f, skiprows=5)
                df_ca = pd.read_excel(c_f, skiprows=5)
                df_v.columns = [str(c).strip() for c in df_v.columns]
                df_ca.columns = [str(c).strip() for c in df_ca.columns]

                ev_col = next((c for c in df_v.columns if 'V\u00c8NEMENT' in c.upper() or 'EVENEMENT' in c.upper()), 'Numéro évènement')
                ft_col = 'Numéro Facture'
                dt_col = next((c for c in df_v.columns if 'DATE' in c.upper()), 'Date')
                ttc_col = next((c for c in df_ca.columns if 'TTC' in c.upper()), 'Montant TTC')

                # تنظيف أرقام الفواتير (نحي .0)
                for df in [df_v, df_ca]:
                    if ft_col in df.columns:
                        df[ft_col] = df[ft_col].apply(lambda x: str(x).split('.')[0].strip() if pd.notnull(x) else "")
                    if ev_col in df.columns:
                        df[ev_col] = df[ev_col].astype(str).str.strip()

                df_v_a79 = df_v[df_v[ev_col].str.contains('A79', na=False)].copy()
                cols_ca = [ev_col, ft_col]
                if ttc_col in df_ca.columns: cols_ca.append(ttc_col)
                
                df_m = pd.merge(df_v_a79, df_ca[cols_ca], on=[ev_col, ft_col], how='inner')

                if not df_m.empty:
                    # حل مشكلة KeyError: 'Mois'
                    df_m['Mois'] = pd.to_datetime(df_m[dt_col], errors='coerce').dt.month.fillna(1).astype(int)
                    df_m['Energy_Type'] = 'ÉLECTRICITÉ' if energy == 'BT' else 'GAZ'
                    df_m['Mtt_TTC'] = pd.to_numeric(df_m[ttc_col], errors='coerce').fillna(0)
                    # تصنيف AO/FSM
                    type_c = [c for c in df_m.columns if 'TYPE' in c.upper() or 'CLIENT' in c.upper()]
                    df_m['Cat'] = df_m[type_c[0]].apply(lambda x: 'AO' if 'AO' in str(x).upper() else ('FSM' if 'FSM' in str(x).upper() else 'OTHER')) if type_c else 'OTHER'
                    all_data_list.append(df_m)

        if all_data_list:
            final_df = pd.concat(all_data_list, ignore_index=True)
            st.success("✅ تم استخراج البيانات بنجاح!")
            
            # --- إنشاء ملف Excel في الذاكرة ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # 1. صفحة التفاصيل
                final_df.to_excel(writer, index=False, sheet_name="Détails_Fraude")
                ws2 = writer.sheets["Détails_Fraude"]
                for i in range(1, len(final_df.columns)+1):
                    ws2.cell(1, i).fill = header_fill
                    ws2.cell(1, i).font = Font(color="FFFFFF", bold=True)

                # 2. صفحة الملخص (Synthèse)
                ws1 = writer.book.create_sheet("Synthèse Fraude", 0)
                ws1.sheet_view.showGridLines = False
                
                # منطق بناء جدول الملخص
                agencies = sorted(final_df['Agence'].unique())
                months = sorted(final_df['Mois'].unique())
                
                row = 2
                for m in months:
                    ws1.cell(row, 1, f"MOIS DE : {months_fr.get(m, 'Inconnu').upper()}").font = Font(size=14, bold=True)
                    row += 2
                    # رؤوس الجدول
                    ws1.cell(row, 1, "AGENCE").fill = sub_fill
                    ws1.cell(row, 1).font = Font(color="FFFFFF")
                    for i, cat in enumerate(['AO', 'FSM', 'Total']):
                        c_idx = 2 + (i*2)
                        ws1.merge_cells(start_row=row, start_column=c_idx, end_row=row, end_column=c_idx+1)
                        ws1.cell(row, c_idx, cat).fill = sub_fill
                        ws1.cell(row, c_idx).alignment = Alignment(horizontal='center')
                    row += 1
                    
                    for ag in agencies:
                        ws1.cell(row, 1, ag).border = std_border
                        ag_data = final_df[(final_df['Agence']==ag) & (final_df['Mois']==m)]
                        for i, cat in enumerate(['AO', 'FSM', 'Total']):
                            c_df = ag_data[ag_data['Cat']==cat] if cat != 'Total' else ag_data
                            ws1.cell(row, 2+(i*2), len(c_df)).border = std_border
                            ws1.cell(row, 3+(i*2), c_df['Mtt_TTC'].sum()).border = std_border
                            ws1.cell(row, 3+(i*2)).number_format = '# ##0.00'
                        row += 1
                    row += 2

            st.download_button(
                label="📥 تحميل التقرير النهائي (Excel)",
                data=output.getvalue(),
                file_name="Rapport_A79_Youcef_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.balloons()
        else:
            st.warning("لم يتم العثور على أي بيانات مطابقة لـ A79.")

run_mega_process()
