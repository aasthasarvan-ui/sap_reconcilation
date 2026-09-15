import streamlit as st
import pandas as pd
import io

# Page Configuration
st.set_page_config(page_title="SAP Stock Reconciliation & True Exact Match Auditor", layout="wide")

st.title("📦 SAP Stock Reconciliation & True Exact Match Auditor")
st.markdown("Python-powered reconciliation with **Strict Mathematical Exact Subset Matching** (Zero Approximation).")

# 3 File Uploaders
col1, col2, col3 = st.columns(3)
with col1:
    export_file = st.file_uploader("1. SAP Export File (.xlsx)", type=["xlsx", "xls"])
with col2:
    mb51_file = st.file_uploader("2. MB51 Transactions File (.xlsx)", type=["xlsx", "xls"])
with col3:
    phy_file = st.file_uploader("3. Physical Stock File (.xlsx) [Optional]", type=["xlsx", "xls"])

if export_file and mb51_file:
    try:
        df_export = pd.read_excel(export_file)
        df_mb51 = pd.read_excel(mb51_file)
        
        df_phy = None
        if phy_file:
            df_phy = pd.read_excel(phy_file)

        df_export.columns = df_export.columns.str.strip()
        df_mb51.columns = df_mb51.columns.str.strip()

        # Dynamic mapping
        mat_exp = next((c for c in df_export.columns if 'material' in c.lower()), None)
        op_col = next((c for c in df_export.columns if 'opening' in c.lower()), None)
        rec_exp = next((c for c in df_export.columns if 'receipt' in c.lower()), None)
        iss_exp = next((c for c in df_export.columns if 'issue' in c.lower()), None)
        cls_col = next((c for c in df_export.columns if 'closing' in c.lower()), None)
        plant_col = next((c for c in df_export.columns if 'plant' in c.lower()), None)

        mat_mb51 = next((c for c in df_mb51.columns if 'material' in c.lower() and 'doc' not in c.lower()), None)
        qty_mb51 = next((c for c in df_mb51.columns if 'qty' in c.lower() or 'quantity' in c.lower()), None)
        date_mb51 = next((c for c in df_mb51.columns if 'date' in c.lower() or 'posting' in c.lower()), None)
        doc_mb51 = next((c for c in df_mb51.columns if 'doc' in c.lower() and 'material' in c.lower()), None)
        mov_mb51 = next((c for c in df_mb51.columns if 'movement' in c.lower() or 'mov' in c.lower()), None)

        if mat_exp and op_col and mat_mb51 and qty_mb51:
            df_export['Material'] = df_export[mat_exp].astype(str).str.strip()
            df_export = df_export[~df_export['Material'].str.lower().str.contains('grand total', na=False)]

            phy_map = {}
            if df_phy is not None:
                m_phy = next((c for c in df_phy.columns if 'material' in c.lower() or 'code' in c.lower()), None)
                q_phy = next((c for c in df_phy.columns if 'qty' in c.lower() or 'stock' in c.lower() or 'bag' in c.lower()), None)
                if m_phy and q_phy:
                    for _, r in df_phy.iterrows():
                        phy_map[str(r[m_phy]).strip()] = float(r[q_phy]) if pd.notnull(r[q_phy]) else 0.0

            df_mb51['Clean_Material'] = df_mb51[mat_mb51].astype(str).str.strip()
            df_mb51['Clean_Qty'] = pd.to_numeric(df_mb51[qty_mb51], errors='coerce').fillna(0)

            rec_df = df_mb51[df_mb51['Clean_Qty'] > 0].groupby('Clean_Material')['Clean_Qty'].sum().reset_index()
            rec_df.rename(columns={'Clean_Qty': 'MB51_Raw_Receipts', 'Clean_Material': 'Material'}, inplace=True)

            iss_df = df_mb51[df_mb51['Clean_Qty'] < 0].groupby('Clean_Material')['Clean_Qty'].sum().abs().reset_index()
            iss_df.rename(columns={'Clean_Qty': 'MB51_Raw_Issues', 'Clean_Material': 'Material'}, inplace=True)

            summary_list = []
            for _, row in df_export.iterrows():
                mat = row['Material']
                plant = row[plant_col] if plant_col and plant_col in row else '2100'
                opening = float(row[op_col]) if pd.notnull(row[op_col]) else 0.0
                off_rec = float(row[rec_exp]) if rec_exp and pd.notnull(row[rec_exp]) else 0.0
                off_iss = float(row[iss_exp]) if iss_exp and pd.notnull(row[iss_exp]) else 0.0
                sap_close = float(row[cls_col]) if cls_col and pd.notnull(row[cls_col]) else 0.0

                mb_rec = rec_df[rec_df['Material'] == mat]['MB51_Raw_Receipts'].values
                mb_rec_val = float(mb_rec[0]) if len(mb_rec) > 0 else 0.0

                mb_iss = iss_df[iss_df['Material'] == mat]['MB51_Raw_Issues'].values
                mb_iss_val = float(mb_iss[0]) if len(mb_iss) > 0 else 0.0

                mb_close = opening + mb_rec_val - mb_iss_val
                rec_diff = mb_rec_val - off_rec
                iss_diff = mb_iss_val - off_iss
                phy_val = phy_map.get(mat, sap_close)
                variance = phy_val - sap_close

                summary_list.append({
                    'Plant': plant,
                    'Material Code': mat,
                    'Opening Stock': opening,
                    'Official Receipts (+)': off_rec,
                    'MB51 Raw Receipts (+)': mb_rec_val,
                    'Receipts Diff': rec_diff,
                    'Official Issues (-)': off_iss,
                    'MB51 Raw Issues (-)': mb_iss_val,
                    'Issues Diff': iss_diff,
                    'SAP Official Closing': sap_close,
                    'MB51 Calc Closing': mb_close,
                    'Physical Stock': phy_val,
                    'Variance (Phy vs SAP)': variance,
                    'Status': 'Movement Log Gap Found' if (rec_diff != 0 or iss_diff != 0) else 'Matched'
                })

            summary_df = pd.DataFrame(summary_list)
            
            st.success("Files successfully processed with Python Pandas!")
            st.subheader("📊 Comparative Audit Summary")
            st.dataframe(summary_df, use_container_width=True)

            # Export Button
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='Audit_Summary', index=False)
            st.download_button(
                label="📥 Download Full Comparison Excel Report",
                data=output.getvalue(),
                file_name="SAP_Python_Audit_Summary.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.divider()
            st.subheader("🔍 Official vs Raw Chronological Ledger & Exact Match Inspector")
            
            material_options = [s['Material Code'] for s in summary_list]
            selected_mat = st.selectbox("Select Material Code for Audit Inspection:", material_options)

            if selected_mat:
                mat_summary = next((s for s in summary_list if s['Material Code'] == selected_mat), None)
                
                if mat_summary:
                    # Metric Cards
                    m1, m2, m3, m4, m5 = st.columns(5)
                    m1.metric("Official Receipts", f"+{mat_summary['Official Receipts (+)']:,}", delta=f"Diff: {mat_summary['Receipts Diff']}")
                    m2.metric("MB51 Raw Receipts", f"+{mat_summary['MB51 Raw Receipts (+)']:,}")
                    m3.metric("Official Issues", f"-{mat_summary['Official Issues (-)']:,}", delta=f"Diff: {mat_summary['Issues Diff']}", delta_color="inverse")
                    m4.metric("MB51 Raw Issues", f"-{mat_summary['MB51 Raw Issues (-)']:,}")
                    m5.metric("SAP Official Closing", f"{mat_summary['SAP Official Closing']:,}", delta=f"Phy Var: {mat_summary['Variance (Phy vs SAP)']}")

                view_mode = st.radio(
                    "Select Ledger Inspection Mode:",
                    [
                        "Full Chronological Ledger (Opening + All Transactions)",
                        f"Strict Exact Match: Official Receipts Filter (Target: +{mat_summary['Official Receipts (+)']})",
                        f"Strict Exact Match: Official Issues Filter (Target: -{mat_summary['Official Issues (-)']})"
                    ],
                    horizontal=True
                )

                mat_opening = mat_summary['Opening Stock'] if mat_summary else 0.0
                mat_rows = df_mb51[df_mb51['Clean_Material'] == selected_mat].copy()
                
                if date_mb51 and date_mb51 in mat_rows.columns:
                    mat_rows = mat_rows.sort_values(by=date_mb51, ascending=True)

                def find_exact_subset_sum(rows_df, target_val):
                    """True exact subset sum algorithm using recursive backtracking to guarantee 100% exact match"""
                    items = []
                    for idx, r in rows_df.iterrows():
                        items.append((float(r['Clean_Qty']), idx, r))
                    
                    # Sort descending for optimization
                    items.sort(key=lambda x: abs(x[0]), reverse=True)
                    
                    def solve(index, current_sum, current_combination):
                        if abs(current_sum - target_val) < 1e-5:
                            return current_combination
                        if index >= len(items) or current_sum > target_val + 1e-5:
                            return None
                        
                        # Try including current item
                        q, idx, row = items[index]
                        if current_sum + q <= target_val + 1e-5:
                            res = solve(index + 1, current_sum + q, current_combination + [row])
                            if res is not None:
                                return res
                        
                        # Try excluding current item
                        res = solve(index + 1, current_sum, current_combination)
                        if res is not None:
                            return res
                        
                        return None

                    matched = solve(0, 0.0, [])
                    if matched:
                        return pd.DataFrame(matched)
                    return pd.DataFrame() # Return empty if no 100% exact match exists in raw log

                ledger_data = []

                if "Official Receipts Filter" in view_mode:
                    target_rec = mat_summary['Official Receipts (+)']
                    pos_rows = mat_rows[mat_rows['Clean_Qty'] > 0]
                    matched_receipts = find_exact_subset_sum(pos_rows, target_rec)
                    
                    if not matched_receipts.empty:
                        running_tot = 0.0
                        for _, r in matched_receipts.iterrows():
                            p_date = r[date_mb51] if date_mb51 and pd.notnull(r[date_mb51]) else 'N/A'
                            mov_t = r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else 'N/A'
                            doc_t = r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else 'N/A'
                            q_val = float(r['Clean_Qty'])

                            running_tot += q_val
                            ledger_data.append({
                                'Material Code': selected_mat,
                                'Posting Date': str(p_date),
                                'Movement Type': f"Mov {mov_t}",
                                'Material Document': str(doc_t),
                                'Quantity': q_val,
                                'Running Balance': running_tot
                            })
                        st.success(f"✅ 100% Exact Match Found! Net Total: **{running_tot}** (Target: {target_rec})")
                        ledger_df = pd.DataFrame(ledger_data)
                    else:
                        st.error(f"❌ MB51 raw log me aisi koi exact combination nahi milti jiska sum theek **{target_rec}** ho. Iska matlab official report me kuch rounding ya net-off adjustment shamil hai jo raw log se alag hai.")
                        ledger_df = pd.DataFrame(columns=['Material Code', 'Posting Date', 'Movement Type', 'Material Document', 'Quantity', 'Running Balance'])

                elif "Official Issues Filter" in view_mode:
                    target_iss = mat_summary['Official Issues (-)']
                    neg_rows = mat_rows[mat_rows['Clean_Qty'] < 0].copy()
                    neg_rows['Abs_Qty'] = neg_rows['Clean_Qty'].abs()
                    
                    # Convert to positive for subset sum search
                    neg_rows_pos = neg_rows.rename(columns={'Abs_Qty': 'Clean_Qty'})
                    matched_issues = find_exact_subset_sum(neg_rows_pos, target_iss)
                    
                    if not matched_issues.empty:
                        running_tot = 0.0
                        for _, r in matched_issues.iterrows():
                            p_date = r[date_mb51] if date_mb51 and pd.notnull(r[date_mb51]) else 'N/A'
                            mov_t = r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else 'N/A'
                            doc_t = r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else 'N/A'
                            q_val = -float(r['Clean_Qty']) # Make negative back

                            running_tot += q_val
                            ledger_data.append({
                                'Material Code': selected_mat,
                                'Posting Date': str(p_date),
                                'Movement Type': f"Mov {mov_t}",
                                'Material Document': str(doc_t),
                                'Quantity': q_val,
                                'Running Balance': running_tot
                            })
                        st.success(f"✅ 100% Exact Match Found! Net Total: **{abs(running_tot)}** (Target: {target_iss})")
                        ledger_df = pd.DataFrame(ledger_data)
                    else:
                        st.error(f"❌ MB51 raw log me aisi koi exact combination nahi milti jiska sum theek **{target_iss}** ho.")
                        ledger_df = pd.DataFrame(columns=['Material Code', 'Posting Date', 'Movement Type', 'Material Document', 'Quantity', 'Running Balance'])

                else:
                    running_tot = mat_opening
                    ledger_data.append({
                        'Material Code': selected_mat,
                        'Posting Date': 'Opening Balance',
                        'Movement Type': '-',
                        'Material Document': '-',
                        'Quantity': mat_opening,
                        'Running Balance': running_tot
                    })

                    for _, r in mat_rows.iterrows():
                        p_date = r[date_mb51] if date_mb51 and pd.notnull(r[date_mb51]) else 'N/A'
                        mov_t = r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else 'N/A'
                        doc_t = r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else 'N/A'
                        q_val = float(r['Clean_Qty'])

                        running_tot += q_val
                        ledger_data.append({
                            'Material Code': selected_mat,
                            'Posting Date': str(p_date),
                            'Movement Type': f"Mov {mov_t}",
                            'Material Document': str(doc_t),
                            'Quantity': q_val,
                            'Running Balance': running_tot
                        })
                    ledger_df = pd.DataFrame(ledger_data)

                st.dataframe(ledger_df, use_container_width=True)

                # Download Button
                if not ledger_df.empty:
                    ledger_output = io.BytesIO()
                    with pd.ExcelWriter(ledger_output, engine='openpyxl') as writer:
                        ledger_df.to_excel(writer, sheet_name='Filtered_Ledger', index=False)
                    st.download_button(
                        label=f"📥 Download Ledger for {selected_mat} (.xlsx)",
                        data=ledger_output.getvalue(),
                        file_name=f"{selected_mat}_Exact_Match_Ledger.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        else:
            st.error("Error: Could not automatically detect required columns in your files.")
    except Exception as e:
            st.error(f"An error occurred: {e}")
