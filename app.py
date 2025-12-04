import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="K-Electric Distribution Network Tool Form", layout="wide")
st.title("🔧 K-Electric Distribution Network Tool Form")

@st.cache_data
def load_data():
    """
    Load employees, tools and existing requests from absolute paths.
    """

    # ---------- EMPLOYEES ----------
    try:
        employees = pd.read_excel("employees.xlsx")
        employees.columns = [c.strip() for c in employees.columns]
        employees = employees.rename(
            columns={
                "Employee Number": "EmployeeNumber",
                "EmployeeNumber": "EmployeeNumber",
                "Name": "Name",
                "Designation": "Designation",
                "Cluster": "Cluster",
            }
        )
        employees = employees[["EmployeeNumber", "Name", "Designation", "Cluster"]]
    except Exception as e:
        st.error(f"Could not read employees.xlsx: {e}")
        employees = pd.DataFrame(
            columns=["EmployeeNumber", "Name", "Designation", "Cluster"]
        )

    # ---------- TOOLS ----------
    try:
        tools = pd.read_excel("Tool_mapping.xlsx")
        tools.columns = [c.strip() for c in tools.columns]
        tools = tools.rename(
            columns={
                "Designation": "Designation",
                "Tool Name": "ToolName",
                "ToolName": "ToolName",
            }
        )
        tools = tools[["Designation", "ToolName"]]
    except Exception as e:
        st.error(f"Could not read Tool_mapping.xlsx: {e}")
        tools = pd.DataFrame(columns=["Designation", "ToolName"])

    # ---------- REQUESTS ----------
    try:
        requests = pd.read_excel("requests.xlsx")
        requests.columns = [c.strip() for c in requests.columns]
    except Exception:
        requests = pd.DataFrame(
            columns=[
                "EmployeeNumber",
                "Name",
                "Designation",
                "Cluster",
                "AOC",
                "ToolName",
                "Quantity",
                "Date",
                "Status",
            ]
        )

    return employees, tools, requests

employees_df, tools_df, requests_df_initial = load_data()

if "requests_df" not in st.session_state:
    st.session_state["requests_df"] = requests_df_initial.copy()

with st.sidebar:
    st.info(
        "1. Enter Employee Number\n"
        "2. AOC select\n"
        "3. Tick tools & qty\n"
        "4. Submit"
    )

col1, col2 = st.columns([1, 2])

# ========== EMPLOYEE & LOCKED FIELDS ==========
with col1:
    st.subheader("👤 Employee Details")

    emp_num = st.text_input("Employee Number")

    if emp_num:
        # ensure EmployeeNumber is string for matching
        df_emp = employees_df.copy()
        df_emp["EmployeeNumber"] = df_emp["EmployeeNumber"].astype(str)

        row = df_emp[df_emp["EmployeeNumber"] == str(emp_num)]

        if not row.empty:
            rec = row.iloc[0]
            st.success("Employee found")

            # LOCKED display (labels only, user cannot edit)
            c1, c2_, c3 = st.columns(3)
            with c1:
                st.text("Name")
                st.write(f"**{rec['Name']}**")
            with c2_:
                st.text("Designation")
                st.write(f"**{rec['Designation']}**")
            with c3:
                st.text("Cluster")
                st.write(f"**{rec['Cluster']}**")

            st.session_state["emp_data"] = rec.to_dict()
        else:
            st.error("Employee not found")
            if "emp_data" in st.session_state:
                del st.session_state["emp_data"]

with col2:
    st.subheader("🏢 AOC")
    offices = ["Industrial Zone 1", "Industrial Zone 2", "Gizri", "Defence", "Korangi"]
    selected_office = st.selectbox("Select AOC", offices)

# ========== TOOLS BY DESIGNATION ==========
if "emp_data" in st.session_state:
    desig = st.session_state["emp_data"]["Designation"]
    st.subheader(f"🛠️ Tools for {desig}")

    tools_for_desig = tools_df[tools_df["Designation"] == desig]

    if tools_for_desig.empty:
        st.info("No tools configured for this designation.")
    else:
        tool_selections = {}
        cols = st.columns(3)

        for i, (_, r) in enumerate(tools_for_desig.iterrows()):
            col_idx = i % 3
            with cols[col_idx]:
                tname = r["ToolName"]
                checked = st.checkbox(tname, key=f"chk_{tname}")
                if checked:
                    qty = st.number_input(
                        "Qty",
                        min_value=1,
                        value=1,
                        step=1,
                        key=f"qty_{tname}",
                    )
                    tool_selections[tname] = qty

        if st.button("💾 Submit Request", type="primary"):
            if not tool_selections:
                st.warning("Select at least one tool.")
            else:
                emp = st.session_state["emp_data"]
                new_rows = []
                for tool_name, qty in tool_selections.items():
                    new_rows.append(
                        {
                            "EmployeeNumber": emp["EmployeeNumber"],
                            "Name": emp["Name"],
                            "Designation": emp["Designation"],
                            "Cluster": emp["Cluster"],
                            "AOC": selected_office,
                            "ToolName": tool_name,
                            "Quantity": qty,
                            "Date": datetime.now().strftime(
                                "%Y-%m-%d %H:%M:%S"
                            ),
                            "Status": "Submitted",
                        }
                    )
                new_df = pd.DataFrame(new_rows)
                st.session_state["requests_df"] = pd.concat(
                    [st.session_state["requests_df"], new_df],
                    ignore_index=True,
                )

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    st.session_state["requests_df"].to_excel(
                        writer, sheet_name="Requests", index=False
                    )

                st.download_button(
                    "📥 Download Updated requests.xlsx",
                    data=output.getvalue(),
                    file_name="requests.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.success(f"{len(new_rows)} request(s) submitted.")
                st.balloons()
else:
    st.info("Enter a valid Employee Number first.")

# ========== DASHBOARD & RECENT REQUESTS ==========
st.markdown("---")
st.subheader("📊 Requests Dashboard")

if not st.session_state["requests_df"].empty:
    df = st.session_state["requests_df"]

    # --- KPI cards ---
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Total Requests", len(df))
    with c2:
        st.metric("Submitted", int((df["Status"] == "Submitted").sum()))
    with c3:
        st.metric(
            "Technician Requests",
            int((df["Designation"] == "Technician").sum()),
        )
    with c4:
        today = datetime.now().strftime("%Y-%m-%d")
        st.metric(
            "Today's Requests",
            int(df["Date"].astype(str).str.contains(today).sum()),
        )

    st.markdown("### 📄 Recent Requests (Last 50)")
    st.dataframe(
        df.tail(50)[
            [
                "Employee Number",
                "Name",
                "Designation",
                "Tool Name",
                "Quantity",
                "AOC",
                "Date",
                "Status",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )
else:
    st.info("No requests yet.")
