import streamlit as st
import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
from dotenv import load_dotenv
import os
import json
import tempfile

# === Load UPS Credentials ===
load_dotenv()
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")

# === Get Access Token ===
def get_access_token():
    auth_url = "https://onlinetools.ups.com/security/v1/oauth/token"
    try:
        response = requests.post(
            auth_url,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            data={"grant_type": "client_credentials"},
            auth=HTTPBasicAuth(client_id, client_secret),
            timeout=10
        )
        response.raise_for_status()
        return response.json().get("access_token")
    except Exception as e:
        print("❌ Error getting token:", e)
        return None

# === Get Tracking Info ===
def get_tracking_info(tracking_number, token):
    url = f"https://onlinetools.ups.com/api/track/v1/details/{tracking_number}"
    headers = {
        "transId": "track_batch",
        "transactionSrc": "myApp",
        "Authorization": f"Bearer {token}"
    }
    query = {
        "locale": "en_US",
        "returnSignature": "True",
        "returnMilestones": "True",
        "returnPOD": "True"
    }

    result = {
        "Tracking Number": tracking_number,
        "Current Status": "",
        "Delivery Date": "",
        "Requested Delivery Date": "",
        "Delivery Time": "",
        "Delivered To": "",
        "Delivery Location": "",
        "Delivery Signature Present?": "No",
        "Weight (lbs)": "",
        "Delayed by Inspection?": "No",
        "First Scan Date": "",
        "Pickup Date": "",
        "Delivery Transit Days": "",
        "Purchase Order Number": "",
        "Waybill Number": "",
        "Full History":""
    }

    try:
        response = requests.get(url, headers=headers, params=query, timeout=10)
        response.raise_for_status()
        data = response.json()

        shipment_list = data.get('trackResponse', {}).get('shipment', [])


        if not shipment_list:
            print(f"⚠️ No shipment info for {tracking_number}")
            return result

        shipment = shipment_list[0]

        package_list = shipment.get('package', [])
        if not package_list:
            print(f"⚠️ No package info for {tracking_number}")
            return result

        package = package_list[0]
        # Current Status
        current_status = package.get('currentStatus')
        if isinstance(current_status, dict):
            result["Current Status"] = current_status.get('description', '')

        # Delivery Dates
        delivery_dates = package.get('deliveryDate', [])
        if isinstance(delivery_dates, list):
            for d in delivery_dates:
                if not isinstance(d, dict):
                    continue
                if d.get('type') == "DEL":
                    result["Delivery Date"] = d.get('date', '')
                elif d.get('type','').upper() == "RDD":
                    result["Requested Delivery Date"] = d.get('date', '')
        # Delivery Time
        delivery_time = package.get('deliveryTime')
        if isinstance(delivery_time, dict):
            result["Delivery Time"] = delivery_time.get('endTime', '')

        # Delivered To
        delivery_info = package.get('deliveryInformation')
        if isinstance(delivery_info, dict):
            result["Delivered To"] = delivery_info.get('receivedBy', '')
            # Signature presence
            sig = delivery_info.get('signature')
            if isinstance(sig, dict) and sig.get('image'):
                result["Delivery Signature Present?"] = "Yes"
            result["Delivery Location"] = delivery_info.get('location', '')

        # Weight
        weight = package.get('weight')
        if isinstance(weight, dict):
            w = weight.get('weight')
            unit = weight.get('unitOfMeasurement')
        # If unit is a string, just use it directly
        if w and isinstance(unit, str):
            result["Weight (lbs)"] = f"{w} {unit}"
        # If unit is a dict (just in case), get 'code'
        elif w and isinstance(unit, dict):
            result["Weight (lbs)"] = f"{w} {unit.get('code', '')}"

        activities = package.get('activity', [])

        # Check if delayed for inspection
        inspection_keywords = ["x-ray inspection", "inspection", "customs"]
        delayed = any(
            any(kw.lower() in act.get('status', {}).get('description', '').lower() for kw in inspection_keywords)
                for act in activities if isinstance(act, dict)
            )
        result["Delayed by Inspection?"] = "Yes" if delayed else "No"

        # First Scan Date (oldest activity)
        if activities:
            first_act = activities[-1]
            if isinstance(first_act, dict):
                result["First Scan Date"] = first_act.get('date', '')

        # Latest Status (Optional)
        latest_act = activities[0]
        if isinstance(latest_act, dict):
            latest_status = latest_act.get('status')
            if isinstance(latest_status, dict):
                result["Current Status"] = latest_status.get('description', result["Current Status"])
        # Pickup Date
        pickup_date = shipment.get('pickupDate')
        if isinstance(pickup_date, str):
            result["Pickup Date"] = pickup_date

        # Delivery Transit Days
        from datetime import datetime
        try:
            if result["Delivery Date"] and result["Pickup Date"]:
                fmt = "%Y%m%d"
                d_date = datetime.strptime(result["Delivery Date"], fmt)
                p_date = datetime.strptime(result["Pickup Date"], fmt)
                delta = (d_date - p_date).days
                result["Delivery Transit Days"] = str(delta)
        except Exception:
            pass
        # Purchase Order Number & Waybill Number from references
        package = shipment.get("package", [])[0]  # safely get first package
        references = package.get('referenceNumber', [])
        if isinstance(references, list):
            for ref in references:
                if not isinstance(ref, dict):
                    continue
                desc = ref.get('description', '').lower()
                number = ref.get('number', '').strip()
                if "purchase order" in desc:
                    result["Purchase Order Number"] = number
                elif "waybill" in desc:
                    result["Waybill Number"] = number
        #For Getting the last 5 Status
        activity = package.get('activity',[])
        history_lines = []
        for act in activity[:5]:  # limit to 5 most recent
            date = act.get('date', '')
            time = act.get('time', '')
            desc = act.get('status', {}).get('description', '')
            loc = act.get('location', {})
            loc_text = ", ".join(filter(None, [loc.get('city', ''), loc.get('country', '')]))
            history_lines.append(f"{date} {time} - {desc} ({loc_text})")
        result["Full History"] = "\n".join(history_lines)    
    
    
    except Exception as e:
        print(f"❌ Error extracting data for {tracking_number}: {e}")
        for key in result:
            if key != "Tracking Number":
                result[key] = "❌ Error"

    return result
# === Main Logic ===
def main():
    input_file = r"C:\Users\AI use-case\API Tracking\TrackingNumber.xlsx"  # Update path for the Excel file where Tracking Number are present
    output_file = "tracking_output_full.xlsx"

    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print("❌ Failed to read Excel:", e)
        return

    if "Tracking Number" not in df.columns:
        print("⚠️ Column 'Tracking Number' not found in Excel.")
        return

    token = get_access_token()
    if not token:
        print("❌ Cannot proceed without token.")
        return

    # New expanded output columns based on enriched tracking info
    output_columns = [
        "Tracking Number",
        "Current Status",
        "Delivery Date",
        "Requested Delivery Date",
        "Delivery Time",
        "Delivered To",
        "Delivery Location",
        "Delivery Signature Present?",
        "Weight (lbs)",
        "Delayed by Inspection?",
        "First Scan Date",
        "Pickup Date",
        "Delivery Transit Days",
        "Purchase Order Number",
        "Waybill Number",
        "Full History"
    ]

    # Initialize new columns with empty strings
    for col in output_columns:
        if col not in df.columns:
            df[col] = ""

    # Process each tracking number
    for idx, tn in df["Tracking Number"].items():
        if pd.isna(tn):
            for col in output_columns:
                df.at[idx, col] = "No tracking number"
            continue
        print(f"🔎 Tracking {tn}...")
        details = get_tracking_info(str(tn), token)
        for col in output_columns:
            df.at[idx, col] = details.get(col, "")

    df.to_excel(output_file, index=False)
    print(f"\n✅ Done! Saved detailed tracking to '{output_file}'")
#Run using Streamlit 
st.set_page_config(page_title="UPS API Tracking Tool", layout="centered")
st.title("📦 UPS Tracking Status Checker via API")

uploaded_file = st.file_uploader("📁 Upload Excel file with a 'Tracking Number' column", type=["xlsx"])

if uploaded_file:
    if "processed_df" not in st.session_state:
        df = pd.read_excel(uploaded_file)

        if "Tracking Number" not in df.columns:
            st.error("⚠️ 'Tracking Number' column not found in Excel.")
        else:
            token = get_access_token()
            if not token:
                st.error("❌ Failed to get UPS access token.")
            else:
            # Output columns
                output_columns = [
                    "Tracking Number",
                    "Current Status",
                    "Delivery Date",
                    "Requested Delivery Date",
                    "Delivery Time",
                    "Delivered To",
                    "Delivery Location",
                    "Delivery Signature Present?",
                    "Weight (lbs)",
                    "Delayed by Inspection?",
                    "First Scan Date",
                    "Pickup Date",
                    "Delivery Transit Days",
                    "Purchase Order Number",
                    "Waybill Number",
                    "Full History"
                ]

                for col in output_columns:
                    if col not in df.columns:
                        df[col] = ""

            # Set up UI components
                progress_bar = st.progress(0)
                status_text = st.empty()

                tracking_numbers = df["Tracking Number"]
                total = len(tracking_numbers)

                for idx, (i, tn) in enumerate(tracking_numbers.items()):
                    if pd.isna(tn):
                        for col in output_columns:
                            df.at[i, col] = "No tracking number"
                        continue

                    status_text.text(f"🔄 Processing {idx + 1} of {total}...")
                    result = get_tracking_info(str(tn), token)
                    for col in output_columns:
                        df.at[i, col] = result.get(col, "")
                    progress_bar.progress((idx + 1) / total)

                status_text.text("✅ All tracking numbers processed.")
                st.session_state.processed_df = df
    else:
        df = st.session_state.processed_df  
        # Save to temp file and provide download
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        df.to_excel(tmp.name, index=False)
        tmp_path = tmp.name
        st.success("✅ Tracking completed!")
        with open(tmp_path, "rb") as f:
            st.download_button(
                label="📥 Download Updated Excel",
                data=f,
                file_name="ups_tracking_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

