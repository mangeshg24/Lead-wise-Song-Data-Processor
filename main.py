from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import pandas as pd
import io

app = FastAPI()

# Allow all origins (you can later restrict to your frontend domain)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def fix_date(val):
    if pd.isna(val):
        return ""
    try:
        return pd.to_datetime(val).strftime("%d-%m-%Y")
    except Exception:
        return str(val)


@app.post("/process")
async def process_excel(file: UploadFile = File(...)):
    # Read uploaded Excel into DataFrame
    content = await file.read()
    df = pd.read_excel(io.BytesIO(content))

    rows = []

    # Group by LEAD ID NUMBER (include NaNs)
    for lid, group in df.groupby("LEAD ID NUMBER", dropna=False):
        group = group.reset_index(drop=True)
        first = group.loc[0]

        date_rec = fix_date(first.get("Date of Rec. (Email OR Whatapp)", ""))
        date_done = fix_date(first.get("Date Report Sent/Completed", ""))

        venue = str(first.get("Venue", "")).strip()
        city = str(first.get("City/Loction", "")).strip()
        pi = str(first.get("Proforma Invoice (PI Number)", "")).strip()
        cid = str(first.get("Customer id", "")).strip()

        # Blank checks
        v_blank = venue in ["", "-", "nan", "None", "NaN"]
        c_blank = city in ["", "-", "nan", "None", "NaN"]

        # COMMENTS logic (Venue / City)
        comments = ""
        if v_blank and c_blank:
            comments = "Venue and City not shared by executive"
        elif v_blank:
            comments = "Venue not shared by executive"
        elif c_blank:
            comments = "City not shared by executive"

        # REMARKS logic (Lead ID missing)
        lead_blank = pd.isna(lid) or str(lid).strip() in ["", "-"]
        remarks = ""
        if lead_blank and (pi in ["", "-"]) and (cid in ["", "-"]):
            remarks = "Lead ID not shared"

        # Shazam count
        shazam_count = (group["Type"] == "Shazam").sum()

        # PPL song count
        ppl_count = (group["PPL / NON PPL"] == "PPL").sum()

        # Build final row
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
            "Recd From": first.get("Data Report Received from", ""),
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

    out_df = pd.DataFrame(rows)

    # Write to in-memory Excel
    output_stream = io.BytesIO()
    out_df.to_excel(output_stream, index=False)
    output_stream.seek(0)

    # Return as downloadable file
    return StreamingResponse(
        output_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="Filled_Data.xlsx"'
        },
    )
