# Logistics-UPS-shipment-Tracker-API

A Python + Streamlit application that automates bulk UPS shipment tracking using the official UPS REST Tracking API.
The tool enables operations, logistics, and supply-chain teams to upload an Excel file with tracking numbers, retrieve real-time shipment status via authenticated API calls, and download an enriched Excel file with delivery milestones and exception indicators

**Problem This Solves**
In many Logitics Teams:
1. Shipment tracking is done manually on carrier websites
2. status updates are copied into Excel or ERP notes
3. Delivery exceptions (delays, inspections) are discovered late
This project demonstrates how API-driven automation replaces manual tracking with a scalable, repeatable workflow.

**Key Capabilities**

1. OAuth 2.0 Client Credentials authentication
2. Bulk tracking via Excel upload
3. Live shipment status retrieval from UPS
4. Automated enrichment with:
      Current shipment status
      Pickup & delivery dates
      Requested delivery date 
      First scan date
      Inspection / customs delay detection & many more
5. Streamlit UI with progress tracking 
6. Downloadable enriched Excel output (using Pandas Library)

**High-Level Architecture**
 Run Streamlit UI > Excel Upload > UPS OAuth Token Service > UPS Tracking REST API > JSON Normalization > Enriched Excel Output

**Future possibilities**
This can serve as first playbook for other Carrier Integrations using similar strucutres. 
**To Run Locally**
streamlit run (file_name).py
