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
    "Python-powered official reconciliation featuring **Exact"
    " Subset-Matching, Table Comparison Gap Extraction, and Color-Coded"
    " Ledgers**."
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
            "Status": (
                "Gap Found" if (rec_diff != 0 or iss_diff != 0) else "Matched"
            ),
        })

      summary_df = pd.DataFrame(summary_list)

      st.success("Files successfully processed with Python Pandas!")
      st.subheader("📊 Comparative Audit Summary")
      st.dataframe(summary_df, use_container_width=True)

      # Export Button for Full Summary
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Audit_Summary", index=False)
      st.download_button(
          label="📥 Download Full Comparison Excel Report",
          data=output.getvalue(),
          file_name="SAP_Python_Audit_Summary.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )

      st.divider()
      st.subheader(
          "🔍 Single Material Color-Coded Ledger & Table Comparison Gap Extractor"
      )

      material_options = [s["Material Code"] for s in summary_list]
      selected_mat = st.selectbox(
          "Select Material Code for Detailed Audit:", material_options
      )

      if selected_mat:
        mat_summary = next(
            (s for s in summary_list if s["Material Code"] == selected_mat),
            None,
        )

        if mat_summary:
          c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
          c1.metric("1. Off. Receipts", f"+{mat_summary['Official Receipts (+)']:,}")
          c2.metric("2. MB51 Receipts", f"+{mat_summary['MB51 Raw Receipts (+)']:,}")
          c3.metric(
              "3. Rec. Diff", f"{mat_summary['Receipts Diff']:,}", delta_color="off"
          )
          c4.metric("4. Off. Issues", f"-{mat_summary['Official Issues (-)']:,}")
          c5.metric("5. MB51 Issues", f"-{mat_summary['MB51 Raw Issues (-)']:,}")
          c6.metric(
              "6. Iss. Diff", f"{mat_summary['Issues Diff']:,}", delta_color="off"
          )
          c7.metric(
              "7. SAP Closing",
              f"{mat_summary['SAP Official Closing']:,}",
              delta=f"PhyVar: {mat_summary['Variance (Phy vs SAP)']}",
          )

        # Inspection Modes including Table Comparison Gap Extractor
        view_mode = st.radio(
            "Select Audit Inspection Mode:",
            [
                "1. Full Chronological Ledger (Opening + All Transactions)",
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

        mat_opening = mat_summary["Opening Stock"] if mat_summary else 0.0
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

        # Mode 2: Compare Tables & Extract Missing Receipt Rows
        if "2. 🔍 COMPARE TABLES & EXTRACT MISSING RECEIPT ROWS" in view_mode:
          rec_gap = mat_summary["Receipts Diff"]
          st.warning(
              "🔍 Comparing Official Receipts vs MB51 Raw Receipts to isolate the"
              f" exact **{rec_gap}** missing units..."
          )

          pos_rows = mat_rows[mat_rows["Clean_Qty"] > 0]
          gap_rows = find_exact_subset_sum(pos_rows, abs(rec_gap))
          if gap_rows.empty:
            gap_rows = pos_rows
            st.info("ℹ️ Showing all raw receipt transactions for comparison.")
          else:
            st.success(
                "✅ Successfully extracted exact missing rows making up the"
                f" **{rec_gap}** receipts gap!"
            )

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
                "Quantity (Gap Contribution)": f"+{q_val}",
                "Metric_Type": "Receipt Gap",
            })

          ledger_data.append({
              "Posting Date": "Total Sum of Missing Rows",
              "Movement Type": "",
              "Material Document": "",
              "Quantity (Gap Contribution)": f"{rec_gap} units",
              "Metric_Type": "Summary",
          })
          ledger_df = pd.DataFrame(ledger_data)

        # Mode 3: Official Receipts Exact Match Filter
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
                "Quantity (Gap Contribution)": f"+{q_val}",
                "Metric_Type": "Official Receipt",
            })
          ledger_df = pd.DataFrame(ledger_data)

        # Mode 4: MB51 Raw Receipts Complete Log
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
                "Quantity (Gap Contribution)": f"+{q_val}",
                "Metric_Type": "Raw Receipt",
            })
          ledger_df = pd.DataFrame(ledger_data)

        # Mode 5: Compare Tables & Extract Missing Issue Rows
        elif "5. 🔍 COMPARE TABLES & EXTRACT MISSING ISSUE ROWS" in view_mode:
          iss_gap = mat_summary["Issues Diff"]
          st.warning(
              "🔍 Comparing Official Issues vs MB51 Raw Issues to isolate the"
              f" exact **{iss_gap}** missing units..."
          )

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
                "Quantity (Gap Contribution)": f"{q_val}",
                "Metric_Type": "Issue Gap",
            })

          ledger_data.append({
              "Posting Date": "Total Sum of Missing Rows",
              "Movement Type": "",
              "Material Document": "",
              "Quantity (Gap Contribution)": f"{iss_gap} units",
              "Metric_Type": "Summary",
          })
          ledger_df = pd.DataFrame(ledger_data)

        # Mode 6: Official Issues Exact Match Filter
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
            q_val = abs(float(r["Clean_Qty"]))

            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity (Gap Contribution)": f"-{q_val}",
                "Metric_Type": "Official Issue",
            })
          ledger_df = pd.DataFrame(ledger_data)

        # Mode 7: MB51 Raw Issues Complete Log
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
            q_val = abs(float(r["Clean_Qty"]))

            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity (Gap Contribution)": f"-{q_val}",
                "Metric_Type": "Raw Issue",
            })
          ledger_df = pd.DataFrame(ledger_data)

        # Mode 8: SAP Official Closing Verification
        elif "8. SAP Official Closing Verification" in view_mode:
          target_cls = mat_summary["SAP Official Closing"]
          off_rec = mat_summary["Official Receipts (+)"]
          off_iss = mat_summary["Official Issues (-)"]
          calc_official_closing = mat_opening + off_rec - off_iss

          ledger_data.append({
              "Posting Date": "Opening Balance",
              "Movement Type": "-",
              "Material Document": "-",
              "Quantity (Gap Contribution)": f"{mat_opening}",
              "Metric_Type": "Opening/Closing",
          })
          ledger_data.append({
              "Posting Date": "Official Net Flow",
              "Movement Type": "SUMMARY",
              "Material Document": "-",
              "Quantity (Gap Contribution)": f"{off_rec - off_iss}",
              "Metric_Type": "Closing",
          })
          ledger_df = pd.DataFrame(ledger_data)

        # Mode 9: Physical Stock & Variance Analysis
        elif "9. Physical Stock & Variance Analysis" in view_mode:
          phy_val = mat_summary["Physical Stock"]
          ledger_data.append({
              "Posting Date": "Physical Stock Audit Summary",
              "Movement Type": "PHYSICAL",
              "Material Document": "-",
              "Quantity (Gap Contribution)": f"{phy_val}",
              "Metric_Type": "Variance",
          })
          ledger_df = pd.DataFrame(ledger_data)

        # Default Mode 1: Full Chronological Ledger
        else:
          running_tot = mat_opening
          ledger_data.append({
              "Posting Date": "Opening Balance",
              "Movement Type": "-",
              "Material Document": "-",
              "Quantity (Gap Contribution)": f"{mat_opening}",
              "Metric_Type": "Opening/Closing",
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
            m_type = "Receipt" if q_val > 0 else "Issue"
            ledger_data.append({
                "Posting Date": str(p_date)[:10],
                "Movement Type": f"Mov {mov_t}",
                "Material Document": str(doc_t),
                "Quantity (Gap Contribution)": f"{q_val}",
                "Metric_Type": m_type,
            })
          ledger_df = pd.DataFrame(ledger_data)

        def highlight_metric_rows(row):
          m_type = row.get("Metric_Type", "")
          if "Summary" in m_type or "Variance" in m_type:
            return [
                "background-color: #fff3cd; color: #856404; font-weight:"
                " bold;"
            ] * len(row)
          elif "Receipt" in m_type or "Gap" in m_type:
            return [
                "background-color: #d4edda; color: #155724; font-weight:"
                " bold;"
            ] * len(row)
          elif "Issue" in m_type:
            return [
                "background-color: #f8d7da; color: #721c24; font-weight:"
                " bold;"
            ] * len(row)
          else:
            return [
                "background-color: #eef2f7; color: #333333; font-weight:"
                " bold;"
            ] * len(row)

        if not ledger_df.empty and "Metric_Type" in ledger_df.columns:
          display_df = ledger_df.drop(columns=["Metric_Type"])
          styled_df = ledger_df.style.apply(highlight_metric_rows, axis=1)
          st.dataframe(styled_df, use_container_width=True)
        else:
          st.dataframe(ledger_df, use_container_width=True)

        if not ledger_df.empty:
          ledger_output = io.BytesIO()
          with pd.ExcelWriter(ledger_output, engine="openpyxl") as writer:
            display_df = (
                ledger_df.drop(columns=["Metric_Type"])
                if "Metric_Type" in ledger_df.columns
                else ledger_df
            )
            display_df.to_excel(
                writer, sheet_name="Master_Audit_Ledger", index=False
            )
          st.download_button(
              label=(
                  f"📥 Download Master Color-Coded Ledger for {selected_mat}"
                  " (.xlsx)"
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
