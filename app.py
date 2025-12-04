import streamlit as st
import pandas as pd
import io
from datetime import datetime


st.set_page_config(
    page_title="K-Electric Distribution Network Tool Form",
    layout="wide"
)
st.title("🔧 K-Electric Distribution Network Tool Form")


@st.cache_data
def load_data():
    """
    Load employees, tools and existing requests.
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

# ---------- init password attempts ----------
if "download_attempts" not in st.session_state:
    st.session_state["download_attempts"] = 0

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

    emp_num = st.text_input("Employee Number", key="emp_number_input")

    if emp_num:
        df_emp = employees_df.copy()
        df_emp["EmployeeNumber"] = df_emp["EmployeeNumber"].astype(str)

        row = df_emp[df_emp["EmployeeNumber"] == str(emp_num)]

        if not row.empty:
            rec = row.iloc[0]
            st.success("Employee found")

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

        # ---------- SUBMIT ----------
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

                # Save full history to requests.xlsx
                try:
                    st.session_state["requests_df"].to_excel(
                        "requests.xlsx", sheet_name="Requests", index=False
                    )
                except Exception as e:
                    st.warning(f"Could not write requests.xlsx: {e}")

                                st.success(f"{len(new_rows)} request(s) submitted.")
                st.balloons()

                # -------- CLEAR FORM & REFRESH PAGE --------
                # remove stored employee data
                st.session_state.pop("emp_data", None)

                # clear all tool checkboxes & qty inputs
                for tname in tools_for_desig["ToolName"]:
                    st.session_state.pop(f"chk_{tname}", None)
                    st.session_state.pop(f"qty_{tname}", None)

                # clear employee number input
                st.session_state["emp_number_input"] = ""

                # optional: reset password attempts too
                st.session_state["download_attempts"] = 0

                st.rerun()

else:
    st.info("Enter a valid Employee Number first.")

# ========== ALWAYS-AVAILABLE DOWNLOAD BUTTON ==========
st.markdown("---")
st.subheader("📥 Export Requests")

if not st.session_state["requests_df"].empty:
    # password gate
    pwd = st.text_input("Enter password to download requests.xlsx", type="password")
    if st.button("Unlock download"):
        if pwd == "2313":
            st.session_state["download_attempts"] = 0
            st.success("Password correct. You can download the file below.")
        else:
            st.session_state["download_attempts"] += 1
            attempts_left = 3 - st.session_state["download_attempts"]
            if attempts_left > 0:
                st.error(f"Wrong password. Attempts left: {attempts_left}")
            else:
                st.error("3 wrong attempts. Refreshing the page.")
                st.session_state["download_attempts"] = 0
                st.rerun()

    if st.session_state["download_attempts"] < 3 and pwd == "2313":
        export_output = io.BytesIO()
        with pd.ExcelWriter(export_output, engine="openpyxl") as writer:
            st.session_state["requests_df"].to_excel(
                writer, sheet_name="Requests", index=False
            )

        st.download_button(
            "Download all requests.xlsx",
            data=export_output.getvalue(),
            file_name="requests.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("No requests yet to download.")

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
                "EmployeeNumber",
                "Name",
                "Designation",
                "ToolName",
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
