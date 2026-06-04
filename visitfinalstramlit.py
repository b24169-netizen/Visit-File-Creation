
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Visit File Generator", layout="wide")

st.title("Visit File Generator")

st.write(
    "Upload EMS Analysis, Invoice Multiple Tabs Export, and Employee Export files."
)

# =====================================================
# FILE UPLOADS
# =====================================================

invoice_file = st.file_uploader(
    "Upload Invoice Multiple Tabs Export",
    type=["xlsx"]
)

ems_file = st.file_uploader(
    "Upload EMS Analysis",
    type=["xlsx"]
)

employee_file = st.file_uploader(
    "Upload Employee Export",
    type=["csv"]
)

# =====================================================
# PROCESS BUTTON
# =====================================================

if st.button("Generate Visit File"):

    if not invoice_file or not ems_file or not employee_file:
        st.error("Please upload all three files.")
        st.stop()

    try:

        # =====================================================
        # LOAD SERVICES SHEET
        # =====================================================

        services_df = pd.read_excel(
            invoice_file,
            sheet_name="Services"
        )

        services_df["Visit Date"] = pd.to_datetime(
            services_df["Visit Date"],
            dayfirst=True,
            errors="coerce"
        )

        services_df["VisitDateKey"] = (
            services_df["Visit Date"]
            .dt.date
        )

        services_df["StartTimeKey"] = pd.to_datetime(
            services_df["Visit Start Time"],
            errors="coerce"
        ).dt.strftime("%H:%M")

        services_df["EndTimeKey"] = pd.to_datetime(
            services_df["Visit End Time"],
            errors="coerce"
        ).dt.strftime("%H:%M")

        services_df["CustomerKey"] = (
            services_df["Customer Name"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        services_df["PlannedVisitDate"] = (
            services_df["Visit Date"]
            .dt.strftime("%d-%m-%Y")
        )

        services_df["PlannedEntryTime"] = (
            services_df["StartTimeKey"] + ":00"
        )

        services_df["PlannedExitTime"] = (
            services_df["EndTimeKey"] + ":00"
        )

        # =====================================================
        # LOAD EMS
        # =====================================================

        ems_df = pd.read_excel(
            ems_file,
            header=[2, 3]
        )

        ems_df.columns = [
            "_".join(
                [str(x) for x in col if pd.notna(x)]
            )
            for col in ems_df.columns
        ]

        ems_df.rename(
            columns={
                "Funder_Unnamed: 0_level_1": "Funder",
                "Customer_Unnamed: 1_level_1": "Customer",
                "Date_Unnamed: 3_level_1": "VisitDate",
                "Actual_Start": "ActualStart",
                "Actual_End": "ActualEnd",
                "Actual_Duration": "ActualDuration",
                "Actual_Employee": "EmployeeName",
            },
            inplace=True
        )

        ems_df["VisitDate"] = pd.to_datetime(
            ems_df["VisitDate"],
            dayfirst=True,
            errors="coerce"
        )

        ems_df["VisitDateKey"] = (
            ems_df["VisitDate"]
            .dt.date
        )

        # Match on ACTUAL TIMES

        ems_df["StartTimeKey"] = pd.to_datetime(
            ems_df["ActualStart"],
            errors="coerce"
        ).dt.strftime("%H:%M")

        ems_df["EndTimeKey"] = pd.to_datetime(
            ems_df["ActualEnd"],
            errors="coerce"
        ).dt.strftime("%H:%M")

        ems_df["CustomerKey"] = (
            ems_df["Customer"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        ems_df["EmployeeName"] = (
            ems_df["EmployeeName"]
            .astype(str)
            .str.strip()
        )

        # =====================================================
        # MATCH SERVICES TO EMS
        # =====================================================

        merged_df = services_df.merge(
            ems_df[
                [
                    "CustomerKey",
                    "VisitDateKey",
                    "StartTimeKey",
                    "EndTimeKey",
                    "EmployeeName",
                    "ActualStart",
                    "ActualEnd",
                    "ActualDuration"
                ]
            ],
            on=[
                "CustomerKey",
                "VisitDateKey",
                "StartTimeKey",
                "EndTimeKey"
            ],
            how="inner"
        )

        # =====================================================
        # LOAD EMPLOYEE EXPORT
        # =====================================================

        emp_df = pd.read_csv(employee_file)

        emp_df["EmployeeName"] = (
            emp_df["Last Name"]
            .fillna("")
            .str.strip()
            + ", "
            + emp_df["First Name"]
            .fillna("")
            .str.strip()
        )

        emp_df["MobizioID"] = (
            emp_df["MobizioID"]
            .astype(str)
            .str.strip()
        )

        merged_df = merged_df.merge(
            emp_df[
                [
                    "EmployeeName",
                    "MobizioID"
                ]
            ],
            on="EmployeeName",
            how="left"
        )

        # =====================================================
        # FINAL OUTPUT
        # =====================================================

        output_df = pd.DataFrame()

        output_df["VisitRef"] = merged_df["Service Duty Id"]
        output_df["SSRef"] = merged_df["Customer External Id"]
        output_df["CarerRef"] = merged_df["MobizioID"]
        output_df["EmployeeName"] = merged_df["EmployeeName"]

        output_df["PlannedVisitDate"] = (
            merged_df["PlannedVisitDate"]
        )

        output_df["PlannedEntryTime"] = (
            merged_df["PlannedEntryTime"]
        )

        output_df["PlannedExitTime"] = (
            merged_df["PlannedExitTime"]
        )

        output_df["Customer"] = (
            merged_df["Customer Name"]
        )

        output_df["Funder"] = (
            merged_df["Funder Name"]
        )

        output_df["ActualStart"] = (
            merged_df["ActualStart"]
        )

        output_df["ActualEnd"] = (
            merged_df["ActualEnd"]
        )

        output_df["ActualDuration"] = (
            merged_df["ActualDuration"]
        )

        # =====================================================
        # DISPLAY RESULTS
        # =====================================================

        st.success(
            f"Visit File Created Successfully! Rows: {len(output_df)}"
        )

        st.dataframe(output_df.head(20))

        # =====================================================
        # EXCEL DOWNLOAD
        # =====================================================

        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:
            output_df.to_excel(
                writer,
                index=False,
                sheet_name="Visit File"
            )

        output.seek(0)

        st.download_button(
            label="Download Visit File",
            data=output,
            file_name="Visit_File_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {str(e)}")

