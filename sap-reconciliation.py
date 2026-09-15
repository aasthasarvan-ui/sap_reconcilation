import streamlit as st
import pandas as pd
import io

# Page Configuration
st.set_page_config(page_title="SAP Stock & Physical Variance Audit Report", layout="wide")

st.title("📦 SAP Stock vs Physical Audit & Automated Variance Reporter")
st.markdown("Upload your **SAP Stock Export**, **Physical Stock Count**, and **MB51 Transaction Log** to get automated variance reasons and audit remarks.")

# 3 File Uploaders
col1, col2, col3 = st.columns(3)
with col1:
    sap_file = st.file_uploader("1. SAP Stock File (.xlsx)", type=["xlsx", "xls"])
with col2:
    phy_file = st.file_uploader("2. Physical Stock File (.xlsx)", type=["xlsx", "xls"])
with col3:
    mb51_file = st.file_uploader("3. MB51 Transactions File (.xlsx)", type=["xlsx", "xls"])

if sap_file and phy_file and mb51_file:
    try:
        df_sap = pd.read_excel(sap_file)
        df_phy = pd.read_excel(phy_file)
        df_mb51 = pd.read_excel(mb51_file)

        # Clean column names
        df_sap.columns = df_sap.columns.str.strip()
        df_phy.columns = df_phy.columns.str.strip()
        df_mb51.columns = df_mb51.columns.str.strip()

        # Dynamic column detection
        mat_sap = next((c for c in df_sap.columns if 'material' in c.lower() or 'code' in c.lower() or 'fg' in c.lower()), None)
        qty_sap = next((c for c in df_sap.columns if 'bag' in c.lower() or 'qty' in c.lower() or 'stock' in c.lower() or 'closing' in c.lower()), None)

        mat_phy = next((c for c in df_phy.columns if 'material' in c.lower() or 'code' in c.lower() or 'fg' in c.lower()), None)
        qty_phy = next((c for c in df_phy.columns if 'bag' in c.lower() or 'qty' in c.lower() or 'stock' in c.lower()), None)

        mat_mb51 = next((c for c in df_mb51.columns if 'material' in c.lower() and 'doc' not in c.lower()), None)
        qty_mb51 = next((c for c in df_mb51.columns if 'qty' in c.lower() or 'quantity' in c.lower()), None)
        mov_mb51 = next((c for c in df_mb51.columns if 'movement' in c.lower() or 'mov' in c.lower()), None)

        if mat_sap and qty_sap and mat_phy and qty_phy and mat_mb51 and qty_mb51:
            # Standardize Material codes
            df_sap['Material'] = df_sap[mat_sap].astype(str).str.strip()
            df_phy['Material'] = df_phy[mat_phy].astype(str).str.strip()
            df_mb51['Clean_Material'] = df_mb51[mat_mb51].astype(str).str.strip()
            df_mb51['Clean_Qty'] = pd.to_numeric(df_mb51[qty_mb51], errors='coerce').fillna(0)

            # Aggregate SAP and Physical stocks
            sap_grouped = df_sap.groupby('Material')[qty_sap].sum().reset_index()
            sap_grouped.rename(columns={qty_sap: 'SAP_Stock'}, inplace=True)

            phy_grouped = df_phy.groupby('Material')[qty_phy].sum().reset_index()
            phy_grouped.rename(columns={qty_phy: 'Physical_Stock'}, inplace=True)

            # Merge SAP and Physical
            merged = pd.merge(sap_grouped, phy_grouped, on='Material', how='outer').fillna(0)
            merged['Variance (Phy - SAP)'] = merged['Physical_Stock'] - merged['SAP_Stock']

            # Analyze MB51 transaction patterns per material
            audit_report = []
            for _, row in merged.iterrows():
                mat = row['Material']
                sap_val = row['SAP_Stock']
                phy_val = row['Physical_Stock']
                var = row['Variance (Phy - SAP)']

                # Get MB51 movements for this material
                mat_mb = df_mb51[df_mb51['Clean_Material'] == mat]
                total_receipts = mat_mb[mat_mb['Clean_Qty'] > 0]['Clean_Qty'].sum()
                total_issues = mat_mb[mat_mb['Clean_Qty'] < 0]['Clean_Qty'].sum()

                # Automated Audit Remarks Logic
                if var == 0:
                    status = "Matched (SAP Stock is Accurate)"
                    remark = "SAP Stock and Physical Stock match perfectly. No discrepancy found in system audit."
                elif var > 0:
                    status = "Excess Physical Stock"
                    remark = f"Physical stock is {int(abs(var))} bags higher than SAP. Possible causes: Unposted production receipts (Mov 101), physical count overstatement, or unrecorded return entries."
                else:
                    status = "Shortage / Variance"
                    remark = f"Physical stock is {int(abs(var))} bags lower than SAP. Possible causes: Unbilled sales dispatches (Mov 601), physical pilferage/breakage, or delayed system goods issue posting."

                audit_report.append({
                    'Material Code': mat,
                    'SAP System Stock': sap_val,
                    'Physical Count Stock': phy_val,
                    'Variance (Phy vs SAP)': var,
                    'MB51 Total Receipts (+)': total_receipts,
                    'MB51 Total Issues (-)': total_issues,
                    'Audit Status': status,
                    'Valid Audit Remarks & Probable Causes': remark
                })

            report_df = pd.DataFrame(audit_report)

            st.success("Files successfully processed and analyzed!")
            st.subheader("📋 Automated SAP Audit Variance & Remark Report")
            st.dataframe(report_df, use_container_width=True)

            # Export Button for Audit Report
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                report_df.to_excel(writer, sheet_name='Audit_Variance_Remarks', index=False)
            st.download_button(
                label="📥 Download Audit Variance & Remarks Report (.xlsx)",
                data=output.getvalue(),
                file_name="SAP_Audit_Variance_Remarks_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Error: Could not automatically detect required material code or quantity columns across your files.")
    except Exception as e:
        st.error(f"An error occurred during processing: {e}")
else:
    st.info("💡 Please upload all three files (SAP Stock, Physical Stock, and MB51) to generate the automated audit report.")
