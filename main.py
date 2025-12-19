from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import pandas as pd
import io

app = FastAPI()

# ===============================
# CORS
# ===============================
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===============================
# STEP 2: Progress Tracker
# ===============================
progress_status = {
    "percent": 0,
    "message": "Idle"
}

# ===============================
# Helper: Safe Date Formatter
# ===============================
def format_date_safe(val):
    """
    Convert any date format to DD-MMM-YY (e.g. 14-Oct-25)
    Always return STRING (important for Excel)
    """
    if pd.isna(val) or str(val).strip() in ["", "-", "nan", "None", "NaN"]:
        return ""
    try:
        return pd.to_datetime(val, errors="coerce").strftime("%d-%b-%y")
    except Exception:
        return ""

# ===============================
# Progress API
# ===============================
@app.get("/progress")
def get_progress():
    return progress_status

# ===============================
# Main Processing API
# ===============================
@app.post("/process")
async def process_excel(file: UploadFile = File(...)):
    global progress_status

    # Reset progress
    progress_status["percent"] = 0
    progress_status["message"] = "Starting processing"

    # Read Excel
    content = await file.read()
    df = pd.read_excel(io.BytesIO(content))

    rows = []

    total_leads = df["LEAD ID NUMBER"].nunique(dropna=False)
    processed = 0

    # Group by Lead ID
    for lid, group in df.groupby("LEAD ID NUMBER", dropna=False):
        processed += 1

        percent = int((processed / total_leads) * 100)
        progress_status["percent"] = percent
        progress_status["message"] = f"Processing lead {processed} of {total_leads}"

        group = group.reset_index(drop=True)
        first = group.loc[0]

        # ---- DATE FORMAT (FIXED) ----
        date_rec = format_date_safe(first.get("Date of Rec. (Email OR Whatapp)", ""))
        date_done = format_date_safe(first.get("Date Report Sent/Completed", ""))

        venue = str(first.get("Venue", "")).strip()
        city = str(first.get("City/Loction", "")).strip()
        pi = str(first.get("Proforma Invoice (PI Number)", "")).strip()
        cid = str(first.get("Customer id", "")).strip()

        v_blank = venue in ["", "-", "nan", "None", "NaN"]
        c_blank = city in ["", "-", "nan", "None", "NaN"]

        # COMMENTS logic
        comments = ""
        if v_blank and c_blank:
            comments = "Venue and City not shared by executive"
        elif v_blank:
            comments = "Venue not shared by executive"
        elif c_blank:
            comments = "City not shared by executive"

        # REMARKS logic
        lead_blank = pd.isna(lid) or str(lid).strip() in ["", "-"]
        remarks = ""
        if lead_blank and pi in ["", "-"] and cid in ["", "-"]:
            remarks = "Lead ID not shared"

        # Counts
        shazam_count = (group["Type"] == "Shazam").sum()
        ppl_count = (group["PPL / NON PPL"] == "PPL").sum()

        # Recd From mapping (robust)
        recd_from = ""
        possible_cols = [
            "Data Report Received",
            "Data Report Received from",
            "Data Report Recieved",
            "DATA REPORT RECEIVED",
            "Data report received"
        ]
        for col in possible_cols:
            if col in df.columns:
                recd_from = first.get(col, "")
                break

        if pd.isna(recd_from) or recd_from in ["-", "nan", "None"]:
            recd_from = ""

        # -----------------------------
        # Build Final Row (ALL dates use formatted strings)
        # -----------------------------
        row = {
            "Stream": "PP - WHATAPP",
            "Task Recd. Date": date_rec,
            "Tasks on Hand": "WHATSAPP LABEL CONFIRMATION",
            "LEAD ID NO": "" if lead_blank else lid,
            "PROFOMA INVOICE (PI NO)": pi,
            "CUSTOMER ID": cid,
            "Type": first.get("Type", ""),
            "Event (PP)": "-",
            "Venue (PP)": "" if v_blank else venue,
            "City/Location": "" if c_blank else city,
            "Date of EVENT OR video": "-",
            "Number of Videos": "-",
            "Duration of Video": "-",
            "Shazam": shazam_count,
            "Recd From": recd_from,
            "Contact No": first.get("Contact No", ""),
            "Assigned to for 2nd process": first.get("Processed By", ""),
            "Date Assigned for 2nd process": date_rec,
            "Date Completed by team": date_done,
            "Date Final File uploaded / Sent": date_done,
            "Status": "COMPLETED",
            "Comments": comments,
            "Remarks": remarks,
            "Total Count of Final Files": len(group),
            "PPL Song Count": ppl_count,
        }

        rows.append(row)

    progress_status["percent"] = 100
    progress_status["message"] = "Completed"

    out_df = pd.DataFrame(rows)

    # ===============================
    # STEP 1 FINAL FIX:
    # FORCE DATE COLUMNS AS TEXT
    # ===============================
    date_columns = [
        "Task Recd. Date",
        "Date Assigned for 2nd process",
        "Date Completed by team",
        "Date Final File uploaded / Sent",
    ]

    for col in date_columns:
        if col in out_df.columns:
            out_df[col] = out_df[col].astype(str)

    # Write Excel
    output_stream = io.BytesIO()
    out_df.to_excel(output_stream, index=False)
    output_stream.seek(0)

    return StreamingResponse(
        output_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="Filled_Data.xlsx"'},
    )
