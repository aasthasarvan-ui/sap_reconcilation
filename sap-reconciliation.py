import io
import pandas as pd
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="SAP Stock Reconciliation & Master Foolproof Auditor",
    layout="wide",
)

st.title("📦 SAP Stock Reconciliation & Master Foolproof Auditor")
st.markdown(
    "Python-powered official reconciliation featuring **Master Merged SAP,"
    " Physical, & Movement-Wise Breakdown Reports**."
)

# 3 File Uploaders
col1, col2, col3 = st.columns(3)
with col1:
  export_file = st.file_uploader("1. SAP Export File (.xlsx)", type=["xlsx", "xls"])
with col2:
  mb51_file = st.file_uploader(
      "2. MB51 Transactions File (.xlsx)", type=["xlsx", "xls"]
  )
with col3:
  phy_file = st.file_uploader(
      "3. Physical Stock File (.xlsx) [Optional]", type=["xlsx", "xls"]
  )

if export_file and mb51_file:
  try:
    df_export = pd.read_excel(export_file)
    df_mb51 = pd.read_excel(mb51_file)

    df_phy = None
    if phy_file:
      df_phy = pd.read_excel(phy_file)

    df_export.columns = df_export.columns.str.strip()
    df_mb51.columns = df_mb51.columns.str.strip()

    # Dynamic mapping for SAP Export
    mat_exp = next(
        (c for c in df_export.columns if "material" in c.lower()), None
    )
    op_col = next((c for c in df_export.columns if "opening" in c.lower()), None)
    rec_exp = next(
        (c for c in df_export.columns if "receipt" in c.lower()), None
    )
    iss_exp = next((c for c in df_export.columns if "issue" in c.lower()), None)
    cls_col = next(
        (c for c in df_export.columns if "closing" in c.lower()), None
    )
    plant_col = next((c for c in df_export.columns if "plant" in c.lower()), None)

    # Dynamic mapping for MB51
    mat_mb51 = next(
        (
            c
            for c in df_mb51.columns
            if "material" in c.lower() and "doc" not in c.lower()
        ),
        None,
    )
    qty_mb51 = next(
        (
            c
            for c in df_mb51.columns
            if "qty" in c.lower() or "quantity" in c.lower()
        ),
        None,
    )
    date_mb51 = next(
        (
            c
            for c in df_mb51.columns
            if "date" in c.lower() or "posting" in c.lower()
        ),
        None,
    )
    doc_mb51 = next(
        (
            c
            for c in df_mb51.columns
            if "doc" in c.lower() and "material" in c.lower()
        ),
        None,
    )
    mov_mb51 = next(
        (
            c
            for c in df_mb51.columns
            if "movement" in c.lower() or "mov" in c.lower()
        ),
        None,
    )

    if mat_exp and op_col and mat_mb51 and qty_mb51:
      df_export["Material"] = df_export[mat_exp].astype(str).str.strip()
      df_export = df_export[
          ~df_export["Material"].str.lower().str.contains("grand total", na=False)
      ]

      # MB51 Date Range Filtering in Sidebar
      if date_mb51 and date_mb51 in df_mb51.columns:
        df_mb51[date_mb51] = pd.to_datetime(
            df_mb51[date_mb51], errors="coerce"
        )
        min_d = df_mb51[date_mb51].min().date()
        max_d = df_mb51[date_mb51].max().date()

        st.sidebar.subheader("📅 MB51 Transaction Date Filter")
        start_date = st.sidebar.date_input(
            "Start Date", min_d, min_value=min_d, max_value=max_d
        )
        end_date = st.sidebar.date_input(
            "End Date", max_d, min_value=min_d, max_value=max_d
        )

        df_mb51 = df_mb51[
            (df_mb51[date_mb51].dt.date >= start_date)
            & (df_mb51[date_mb51].dt.date <= end_date)
        ]

      phy_map = {}
      if df_phy is not None:
        m_phy = next(
            (
                c
                for c in df_phy.columns
                if "material" in c.lower()
                or "code" in c.lower()
                or "fg_code" in c.lower()
            ),
            None,
        )
        q_phy = next(
            (
                c
                for c in df_phy.columns
                if "qty" in c.lower()
                or "stock" in c.lower()
                or "bag" in c.lower()
            ),
            None,
        )
        if m_phy and q_phy:
          for _, r in df_phy.iterrows():
            phy_map[str(r[m_phy]).strip()] = (
                float(r[q_phy]) if pd.notnull(r[q_phy]) else 0.0
            )

      df_mb51["Clean_Material"] = df_mb51[mat_mb51].astype(str).str.strip()
      df_mb51["Clean_Qty"] = (
          pd.to_numeric(df_mb51[qty_mb51], errors="coerce").fillna(0)
      )

      rec_df = (
          df_mb51[df_mb51["Clean_Qty"] > 0]
          .groupby("Clean_Material")["Clean_Qty"]
          .sum()
          .reset_index()
      )
      rec_df.rename(
          columns={
              "Clean_Qty": "MB51_Raw_Receipts",
              "Clean_Material": "Material",
          },
          inplace=True,
      )

      iss_df = (
          df_mb51[df_mb51["Clean_Qty"] < 0]
          .groupby("Clean_Material")["Clean_Qty"]
          .sum()
          .abs()
          .reset_index()
      )
      iss_df.rename(
          columns={"Clean_Qty": "MB51_Raw_Issues", "Clean_Material": "Material"},
          inplace=True,
      )

      if "audit_notes" not in st.session_state:
        st.session_state.audit_notes = {}

      summary_list = []
      for _, row in df_export.iterrows():
        mat = row["Material"]
        plant = row[plant_col] if plant_col and plant_col in row else "2100"
        opening = float(row[op_col]) if pd.notnull(row[op_col]) else 0.0
        off_rec = (
            float(row[rec_exp])
            if rec_exp and pd.notnull(row[rec_exp])
            else 0.0
        )
        off_iss = (
            float(row[iss_exp])
            if iss_exp and pd.notnull(row[iss_exp])
            else 0.0
        )
        sap_close = (
            float(row[cls_col])
            if cls_col and pd.notnull(row[cls_col])
            else 0.0
        )

        mb_rec = rec_df[rec_df["Material"] == mat][
            "MB51_Raw_Receipts"
        ].values
        mb_rec_val = float(mb_rec[0]) if len(mb_rec) > 0 else 0.0

        mb_iss = iss_df[iss_df["Material"] == mat]["MB51_Raw_Issues"].values
        mb_iss_val = float(mb_iss[0]) if len(mb_iss) > 0 else 0.0

        mb_close = opening + off_rec - off_iss
        rec_diff = mb_rec_val - off_rec
        iss_diff = mb_iss_val - off_iss
        phy_val = phy_map.get(mat, sap_close)
        variance = phy_val - sap_close

        probable_reason = "No Discrepancy Found - Fully Balanced"
        if variance > 0:
          probable_reason = (
              "Physical Excess: Possible unposted production receipts (Mov"
              " 101) or unrecorded floor returns."
          )
        elif variance < 0:
          probable_reason = (
              "SAP Excess: Possible dispatches (Mov 601) delivered physically"
              " but documentation/posting delayed."
          )
        elif rec_diff != 0 or iss_diff != 0:
          probable_reason = (
              "Ledger Mismatch: Difference between official report summary and"
              " MB51 transaction logs."
          )

        default_note = st.session_state.audit_notes.get(mat, probable_reason)

        summary_list.append({
            "Plant": plant,
            "Material Code": mat,
            "Opening Stock": opening,
            "Official Receipts (+)": off_rec,
            "MB51 Raw Receipts (+)": mb_rec_val,
            "Receipts Diff": rec_diff,
            "Official Issues (-)": off_iss,
            "MB51 Raw Issues (-)": mb_iss_val,
            "Issues Diff": iss_diff,
            "SAP Official Closing": sap_close,
            "MB51 Calc Closing": mb_close,
            "Physical Stock": phy_val,
            "Variance (Phy vs SAP)": variance,
            "Status": "Matched" if variance == 0 else "Gap Found",
            "Audit Remarks / Root Cause": default_note,
        })

      summary_df = pd.DataFrame(summary_list)

      st.success("Files successfully processed with Python Pandas!")

      # Metric Cards
      total_items = len(summary_df)
      matched_items = len(summary_df[summary_df["Variance (Phy vs SAP)"] == 0])
      variance_items = total_items - matched_items
      net_var_sum = summary_df["Variance (Phy vs SAP)"].sum()

      m1, m2, m3, m4 = st.columns(4)
      m1.metric(label="Total SKUs", value=total_items)
      m2.metric(label="Fully Matched", value=matched_items)
      m3.metric(label="Variance SKUs", value=variance_items)
      m4.metric(label="Net Physical Variance", value=f"{net_var_sum:+,}")

      st.subheader(
          "📊 Comparative Audit Summary & Editable Root Cause Tagging"
      )
      edited_summary_df = st.data_editor(
          summary_df,
          use_container_width=True,
          num_rows="fixed",
          key="editable_summary",
      )

      for _, r in edited_summary_df.iterrows():
        st.session_state.audit_notes[r["Material Code"]] = r[
            "Audit Remarks / Root Cause"
        ]

      # MASTER MERGED REPORT: Combining Summary + Movement Pivot Table
      st.subheader(
          "🔗 Master Merged Report (SAP Summary + Physical + Movement Breakdown)"
      )
      if mov_mb51 and mov_mb51 in df_mb51.columns:
        movement_pivot = df_mb51.pivot_table(
            index="Clean_Material",
            columns=mov_mb51,
            values="Clean_Qty",
            aggfunc="sum",
            fill_value=0,
        ).reset_index()
        movement_pivot.rename(
            columns={"Clean_Material": "Material Code"}, inplace=True
        )

        # Merge summary with movement pivot
        master_merged_df = pd.merge(
            edited_summary_df, movement_pivot, on="Material Code", how="left"
        ).fillna(0)
        st.dataframe(master_merged_df, use_container_width=True)

        # Excel export for Master Merged Report
        master_output = io.BytesIO()
        with pd.ExcelWriter(master_output, engine="openpyxl") as writer:
          master_merged_df.to_excel(
              writer, sheet_name="Master_Merged_Report", index=False
          )
        st.download_button(
            label=(
                "📥 Download Master Merged Excel Report (Summary + Movements)"
                " (.xlsx)"
            ),
            data=master_output.getvalue(),
            file_name="SAP_Physical_MB51_Master_Merged_Report.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.dataframe(edited_summary_df, use_container_width=True)

      st.subheader("📈 Material Variance Distribution Chart")
      chart_data = edited_summary_df.set_index("Material Code")[
          "Variance (Phy vs SAP)"
      ]
      st.bar_chart(chart_data)

      st.divider()
      st.subheader(
          "🔍 Single Material Color-Coded Ledger & 9-Mode Gap Extractor"
      )

      material_options = edited_summary_df["Material Code"].tolist()
      selected_mat = st.selectbox(
          "Select Material Code for Detailed Audit:", material_options
      )

      if selected_mat:
        mat_summary = edited_summary_df[
            edited_summary_df["Material Code"] == selected_mat
        ].iloc[0]

        c1, c2, c3, c4, c5, c6, c7, c8, c9 = st.columns(9)
        c1.metric("1. Opening", f"{mat_summary['Opening Stock']:,}")
        c2.metric("2. Off. Rec", f"+{mat_summary['Official Receipts (+)']:,}")
        c3.metric("3. MB51 Rec", f"+{mat_summary['MB51 Raw Receipts (+)']:,}")
        c4.metric(
            "4. Rec Diff", f"{mat_summary['Receipts Diff']:,}", delta_color="off"
        )
        c5.metric("5. Off. Iss", f"-{mat_summary['Official Issues (-)']:,}")
        c6.metric("6. MB51 Iss", f"-{mat_summary['MB51 Raw Issues (-)']:,}")
        c7.metric(
            "7. Iss Diff", f"{mat_summary['Issues Diff']:,}", delta_color="off"
        )
        c8.metric("8. SAP Close", f"{mat_summary['SAP Official Closing']:,}")
        c9.metric(
            "9. Phy Var",
            f"{mat_summary['Variance (Phy vs SAP)']}",
            delta_color="off",
        )

        st.info(
            f"💡 **AI / Expert Root Cause Analysis for {selected_mat}:**"
            f" {mat_summary['Audit Remarks / Root Cause']}"
        )

        view_mode = st.radio(
            "Select Audit Inspection Mode:",
            [
                "1. Full Chronological Ledger (Opening + All Transactions + Running Balance)",
                (
                    f"2. 🔍 COMPARE TABLES & EXTRACT MISSING RECEIPT ROWS"
                    f" (Receipts Gap: {mat_summary['Receipts Diff']})"
                ),
                (
                    f"3. Strict Exact Match: Official Receipts Filter (Target:"
                    f" +{mat_summary['Official Receipts (+)']})"
                ),
                (
                    f"4. Complete Log: MB51 Raw Receipts (Target:"
                    f" +{mat_summary['MB51 Raw Receipts (+)']})"
                ),
                (
                    f"5. 🔍 COMPARE TABLES & EXTRACT MISSING ISSUE ROWS (Issues"
                    f" Gap: {mat_summary['Issues Diff']})"
                ),
                (
                    f"6. Strict Exact Match: Official Issues Filter (Target:"
                    f" -{mat_summary['Official Issues (-)']})"
                ),
                (
                    f"7. Complete Log: MB51 Raw Issues (Target:"
                    f" -{mat_summary['MB51 Raw Issues (-)']})"
                ),
                (
                    f"8. SAP Official Closing Verification (Target Closing:"
                    f" {mat_summary['SAP Official Closing']})"
                ),
                (
                    f"9. Physical Stock & Variance Analysis (Physical Qty:"
                    f" {mat_summary['Physical Stock']}, Variance:"
                    f" {mat_summary['Variance (Phy vs SAP)']})"
                ),
            ],
        )

        mat_opening = mat_summary["Opening Stock"]
        mat_rows = df_mb51[df_mb51["Clean_Material"] == selected_mat].copy()

        if date_mb51 and date_mb51 in mat_rows.columns:
          mat_rows = mat_rows.sort_values(by=date_mb51, ascending=True)

        def find_exact_subset_sum(rows_df, target_val):
          items = []
          for idx, r in rows_df.iterrows():
            items.append((float(r["Clean_Qty"]), r))
          items.sort(key=lambda x: abs(x[0]), reverse=True)

          def solve(index, current_sum, current_combination):
            if abs(current_sum - target_val) < 1e-5:
              return current_combination
            if index >= len(items) or current_sum > target_val + 1e-5:
              return None
            q, row = items[index]
            if current_sum + q <= target_val + 1e-5:
              res = solve(
                  index + 1, current_sum + q, current_combination + [row]
              )
              if res is not None:
                return res
            res = solve(index + 1, current_sum, current_combination)
            if res is not None:
              return res
            return None

          matched = solve(0, 0.0, [])
          if matched:
            return pd.DataFrame(matched)
          return pd.DataFrame()

        ledger_data = []

        if "2. 🔍 COMPARE TABLES & EXTRACT MISSING RECEIPT ROWS" in view_mode:
          rec_gap = mat_summary["Receipts Diff"]
          pos_rows = mat_rows[mat_rows["Clean_Qty"] > 0]
          gap_rows = find_exact_subset_sum(pos_rows, abs(rec_gap))
          if gap_rows.empty:
            gap_rows = pos_rows

          for idx, r in gap_rows.reset_index().iterrows():
            p_date = (
                r[date_mb51]
                if date_mb51 and pd.notnull(r[date_mb51])
                else "N/A"
            )
            mov_t = (
                r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else "N/A"
            )
            doc_t = (
                r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else "N/A"
            )
            q_val = float(r["Clean_Qty"])
            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity": f"+{q_val}",
                "Transaction Type": "Receipt",
            })

        elif "3. Strict Exact Match: Official Receipts Filter" in view_mode:
          target_rec = mat_summary["Official Receipts (+)"]
          pos_rows = mat_rows[mat_rows["Clean_Qty"] > 0]
          matched_receipts = find_exact_subset_sum(pos_rows, target_rec)
          if matched_receipts.empty:
            matched_receipts = pos_rows

          for idx, r in matched_receipts.reset_index().iterrows():
            p_date = (
                r[date_mb51]
                if date_mb51 and pd.notnull(r[date_mb51])
                else "N/A"
            )
            mov_t = (
                r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else "N/A"
            )
            doc_t = (
                r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else "N/A"
            )
            q_val = float(r["Clean_Qty"])
            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity": f"+{q_val}",
                "Transaction Type": "Receipt",
            })

        elif "4. Complete Log: MB51 Raw Receipts" in view_mode:
          pos_rows = mat_rows[mat_rows["Clean_Qty"] > 0]
          for idx, r in pos_rows.reset_index().iterrows():
            p_date = (
                r[date_mb51]
                if date_mb51 and pd.notnull(r[date_mb51])
                else "N/A"
            )
            mov_t = (
                r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else "N/A"
            )
            doc_t = (
                r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else "N/A"
            )
            q_val = float(r["Clean_Qty"])
            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity": f"+{q_val}",
                "Transaction Type": "Receipt",
            })

        elif "5. 🔍 COMPARE TABLES & EXTRACT MISSING ISSUE ROWS" in view_mode:
          iss_gap = mat_summary["Issues Diff"]
          neg_rows = mat_rows[mat_rows["Clean_Qty"] < 0].copy()
          neg_rows["Abs_Qty"] = neg_rows["Clean_Qty"].abs()
          items = []
          for idx, r in neg_rows.iterrows():
            items.append((float(r["Abs_Qty"]), r))
          items.sort(key=lambda x: abs(x[0]), reverse=True)

          def solve_gap(index, current_sum, current_combination):
            if abs(current_sum - abs(iss_gap)) < 1e-5:
              return current_combination
            if index >= len(items) or current_sum > abs(iss_gap) + 1e-5:
              return None
            q, row = items[index]
            if current_sum + q <= abs(iss_gap) + 1e-5:
              res = solve_gap(
                  index + 1, current_sum + q, current_combination + [row]
              )
              if res is not None:
                return res
            res = solve_gap(index + 1, current_sum, current_combination)
            if res is not None:
              return res
            return None

          gap_iss_list = solve_gap(0, 0.0, [])
          gap_matched_iss = (
              pd.DataFrame(gap_iss_list) if gap_iss_list else neg_rows
          )

          for idx, r in gap_matched_iss.reset_index().iterrows():
            p_date = (
                r[date_mb51]
                if date_mb51 and pd.notnull(r[date_mb51])
                else "N/A"
            )
            mov_t = (
                r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else "N/A"
            )
            doc_t = (
                r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else "N/A"
            )
            q_val = float(r["Clean_Qty"])
            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity": f"{q_val}",
                "Transaction Type": "Issue",
            })

        elif "6. Strict Exact Match: Official Issues Filter" in view_mode:
          target_iss = mat_summary["Official Issues (-)"]
          neg_rows = mat_rows[mat_rows["Clean_Qty"] < 0].copy()
          items = []
          for idx, r in neg_rows.iterrows():
            items.append((abs(float(r["Clean_Qty"])), r))
          items.sort(key=lambda x: abs(x[0]), reverse=True)

          def solve_iss(index, current_sum, current_combination):
            if abs(current_sum - target_iss) < 1e-5:
              return current_combination
            if index >= len(items) or current_sum > target_iss + 1e-5:
              return None
            q, row = items[index]
            if current_sum + q <= target_iss + 1e-5:
              res = solve_iss(
                  index + 1, current_sum + q, current_combination + [row]
              )
              if res is not None:
                return res
            res = solve_iss(index + 1, current_sum, current_combination)
            if res is not None:
              return res
            return None

          matched_iss_list = solve_iss(0, 0.0, [])
          matched_issues = (
              pd.DataFrame(matched_iss_list) if matched_iss_list else neg_rows
          )

          for idx, r in matched_issues.reset_index().iterrows():
            p_date = (
                r[date_mb51]
                if date_mb51 and pd.notnull(r[date_mb51])
                else "N/A"
            )
            mov_t = (
                r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else "N/A"
            )
            doc_t = (
                r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else "N/A"
            )
            q_val = float(r["Clean_Qty"])
            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity": f"{q_val}",
                "Transaction Type": "Issue",
            })

        elif "7. Complete Log: MB51 Raw Issues" in view_mode:
          neg_rows = mat_rows[mat_rows["Clean_Qty"] < 0]
          for idx, r in neg_rows.reset_index().iterrows():
            p_date = (
                r[date_mb51]
                if date_mb51 and pd.notnull(r[date_mb51])
                else "N/A"
            )
            mov_t = (
                r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else "N/A"
            )
            doc_t = (
                r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else "N/A"
            )
            q_val = float(r["Clean_Qty"])
            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity": f"{q_val}",
                "Transaction Type": "Issue",
            })

        elif "8. SAP Official Closing Verification" in view_mode:
          off_rec = mat_summary["Official Receipts (+)"]
          off_iss = mat_summary["Official Issues (-)"]
          ledger_data.append({
              "Posting Date": "Opening Balance",
              "Movement Type": "-",
              "Material Document": "-",
              "Quantity": mat_opening,
              "Transaction Type": "Opening",
          })
          ledger_data.append({
              "Posting Date": "Official Net Flow",
              "Movement Type": "SUMMARY",
              "Material Document": "-",
              "Quantity": off_rec - off_iss,
              "Transaction Type": "Summary",
          })

        elif "9. Physical Stock & Variance Analysis" in view_mode:
          phy_val = mat_summary["Physical Stock"]
          ledger_data.append({
              "Posting Date": "Physical Count Date",
              "Movement Type": "PHYSICAL",
              "Material Document": "-",
              "Quantity": phy_val,
              "Transaction Type": "Physical",
          })

        else:
          running_tot = mat_opening
          ledger_data.append({
              "Posting Date": "Opening Balance",
              "Movement Type": "-",
              "Material Document": "-",
              "Quantity": mat_opening,
              "Running Balance": running_tot,
              "Transaction Type": "Opening",
          })
          for idx, r in mat_rows.reset_index().iterrows():
            p_date = (
                r[date_mb51]
                if date_mb51 and pd.notnull(r[date_mb51])
                else "N/A"
            )
            mov_t = (
                r[mov_mb51] if mov_mb51 and pd.notnull(r[mov_mb51]) else "N/A"
            )
            doc_t = (
                r[doc_mb51] if doc_mb51 and pd.notnull(r[doc_mb51]) else "N/A"
            )
            q_val = float(r["Clean_Qty"])
            running_tot += q_val
            t_type = "Receipt" if q_val > 0 else "Issue"
            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity": q_val,
                "Running Balance": running_tot,
                "Transaction Type": t_type,
            })

        ledger_df = pd.DataFrame(ledger_data)

        def color_coded_ledger(row):
          t_type = row.get("Transaction Type", "")
          if t_type == "Receipt":
            return [
                "background-color: #d4edda; color: #155724; font-weight:"
                " bold;"
            ] * len(row)
          elif t_type == "Issue":
            return [
                "background-color: #f8d7da; color: #721c24; font-weight:"
                " bold;"
            ] * len(row)
          else:
            return [
                "background-color: #eef2f7; color: #333333; font-weight:"
                " bold;"
            ] * len(row)

        if not ledger_df.empty:
          styled_ledger = ledger_df.style.apply(color_coded_ledger, axis=1)
          st.dataframe(styled_ledger, use_container_width=True)

          ledger_output = io.BytesIO()
          with pd.ExcelWriter(ledger_output, engine="openpyxl") as writer:
            ledger_df.to_excel(
                writer, sheet_name="Master_Color_Coded_Ledger", index=False
            )
          st.download_button(
              label=(
                  f"📥 Download Color-Coded Ledger with Running Total for"
                  f" {selected_mat} (.xlsx)"
              ),
              data=ledger_output.getvalue(),
              file_name=f"{selected_mat}_Master_Ledger.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )
    else:
      st.error(
          "Error: Could not automatically detect required columns in your"
          " files."
      )
  except Exception as e:
    st.error(f"An error occurred: {e}")
