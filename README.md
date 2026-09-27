# ASK Field Development Project — Document Intelligence Hub

An enterprise executive dashboard suite for tracking document review durations, SLA compliance, transmittal turnaround workflows, and multi-revision lifecycles between **Contractor** and the **Client Engineering Team**.

---

## 🌟 Dashboard Architecture Overview

The system consists of three interconnected, high-performance dashboards:

1. **[Technical Document Log Dashboard (`index.html`)](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Doc_Dashboard/index.html)**
   * **Source Data:** `ASK Technical Document Log Report - <DATE>.xlsx` (~1,822 submissions).
   * **Scope:** 13 Engineering Disciplines (Process, Mechanical, Piping, Structural, Electrical, Instrumentation, Telecom, Safety, HVAC, Naval, Marine, Pipeline, General).
   * **Analytics:** 10 Working Days contractual benchmark compliance, Review Duration Histogram Tiers with discipline drill-down, transmittal timeline drawer, and CSV data export.

2. **[Vendor Document Log Dashboard (`index_vendor.html`)](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Doc_Dashboard/index_vendor.html)**
   * **Source Data:** `ASK Vendor Document Log Report - <DATE>.xls` (~188 submissions across 172 unique documents).
   * **Scope:** 9 Vendor Equipment Packages (Chemical Injection Skid, Plate, Pedestal Cranes, Wellhead Control Panel, Closed Drain Sump Pump, Storage Tank, Open Drain Pump, Diesel Transfer Pump, General).
   * **Analytics:** Review Duration Distribution Tiers (with interactive click-to-view detail modal), package turnaround rankings, multi-transmittal history drawer, and return status monitoring.

3. **[MDR Revision Lifecycle Dashboard (`index_mdr.html`)](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Doc_Dashboard/index_mdr.html)**
   * **Source Data:** `MM-ASK-1A-GEN01-ENG-MDR-0001_B1- Cut off <DATE>.xlsx` (~1,904 deliverables across 5 Work Package tabs).
   * **Scope:** Multi-stage revision workflow tracking:
     * **IFR** (Issue for Review — Rev A1)
     * **IFA** (Issue for Approval — Rev B1, with automatic detection of skipped IFA for Part C Engineering Drawings)
     * **AFC** (Approved for Construction — Rev 0)
     * **AP** (Approved / Close-out)
   * **Analytics:** Contractor comment incorporation turnaround vs Client review durations, TOC constraint bottleneck analysis, RACI accountability matrix, and executive management summaries.

---

## 📖 Definition of "SLA" (Service Level Agreement)

> [!IMPORTANT]
> **SLA = Service Level Agreement**
> In the ASK Project contract, the SLA establishes the procedural standard and maximum turnaround period allowed for the **Client Engineering Team** to review and return submitted deliverables.

* **Contractual Benchmark:** **10 Working Days (WD)**.
* **Meet Agreement (Within SLA):** Review duration **$\le$ 10 Working Days**. The document was evaluated, commented, and returned within the agreed window.
* **Over Agreement (Exceeded SLA):** Review duration **$>$ 10 Working Days**. The document exceeded contractual turnaround, triggering delay tracking.
* **Calculation Formula:**
  $$\text{Working Days} = \text{Count of days between } [\text{Actual Submit In}] \text{ and } [\text{Actual Submit Out}]$$
  * **Start Date:** The first business day following official transmittal receipt (`Actual Submit In`).
  * **Working Days Only:** Monday through Friday are counted. Saturdays and Sundays are excluded.
  * **End Date:** The date the official transmittal is issued back (`Actual Submit Out`).
* **Interactive Modals:** Every dashboard contains an **"ℹ️ What is SLA?"** button and interactive SLA badges explaining these rules in detail with visual examples.

---

## 🔍 Review Duration Distribution Tiers Drill-down

Deliverables are grouped into turnaround tiers:
* **$\le 5$ WD:** Rapid review turnaround.
* **6 – 10 WD:** Contractual compliance sweet spot (contract limit is 10 WD).
* **11 – 15 WD:** Slight delay ($1-5$ days over SLA).
* **16 – 20 WD:** Moderate delay.
* **$> 20$ WD:** Significant bottleneck requiring management escalation.

### How to View Details:
* Simply **click on any tier bar** or click the **`View Details 🔍`** button next to each tier row.
* An interactive **Tier Deliverables Modal** opens showing:
  * Total deliverable count and working-day range.
  * Quick-filter discipline chips (e.g. *All Disciplines*, *Mechanical*, *Structural*).
  * In-tier real-time search box.
  * Complete registry table showing Submittal ID, Document No, Rev, Vendor Package, Dates, Working Days, and Return Status.
  * An **"Apply Filter to Main Table"** button to filter the main registry table by the chosen tier.

---

## ⚡ Automated Weekly Data Update Engine

Contractors receive updated Excel log files on a weekly basis where **only the date in the filename changes** (e.g., changing from `18 Sep 26` to `25 Sep 26`).

This project includes a fully automated updater that handles this seamlessly:

### 1. In-Browser Drag & Drop Upload (No Code Needed!)
1. Click the **`📁 Upload New File`** button in the top navigation of any dashboard.
2. Drag and drop your weekly Excel file (or click **`Browse Files...`**).
3. The system automatically detects the file type (Technical Log, Vendor Log, or MDR Register), runs the Python conversion engine, updates the database, and refreshes the dashboard in real-time.

### 2. Automated Python CLI Batch Updater (`update_all.py`)
To process all files automatically on your local machine:
```bash
python update_all.py
```
* **Auto-Discovery:** Scans the folder using wildcards (`ASK Technical Document Log Report*.xlsx`, `ASK Vendor Document Log Report*.xls*`, and `*MDR*.xlsx`).
* **Latest Date Detection:** Compares dates across matching files and selects the newest cut-off report automatically.
* **Recalculation:** Recalculates all working-day turnaround times excluding weekends.
* **Outputs Generated:** Updates `data.js`, `vendor_data.js`, `mdr_data.js`, and creates `latest_update.json`.

---

## 🚀 Deployment Guide: GitHub & Render.com

The application includes production web server configuration ([`server.py`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Doc_Dashboard/server.py), [`Procfile`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Doc_Dashboard/Procfile), [`requirements.txt`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Doc_Dashboard/requirements.txt), and [`render.yaml`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Doc_Dashboard/render.yaml)).

### Step 1: Push Project to GitHub

1. Open PowerShell in this project directory:
   ```powershell
   git init
   git add .
   git commit -m "Initial commit: ASK Document Intelligence Hub with SLA definitions, tier modals, and weekly auto-updater"
   ```
2. Create a new repository on your GitHub account (e.g., `ASK-Doc-Dashboard`).
3. Link and push your local branch:
   ```powershell
   git remote add origin https://github.com/YOUR_USERNAME/ASK-Doc-Dashboard.git
   git branch -M main
   git push -u origin main
   ```

---

### Step 2: Deploy to Render.com (Free & Fast)

#### Option A: Direct Web Service (Recommended)
1. Sign in to **[Render.com](https://render.com/)**.
2. Click **New +** $\rightarrow$ **Web Service**.
3. Select **Build and deploy from a Git repository** and connect your GitHub repo (`ASK-Doc-Dashboard`).
4. Configure the settings:
   * **Name:** `ask-doc-dashboard` (or your choice).
   * **Language / Environment:** `Python 3`.
   * **Region:** Singapore / Oregon / Frankfurt (choose closest to your team).
   * **Branch:** `main`.
   * **Build Command:**
     ```bash
     pip install -r requirements.txt
     ```
   * **Start Command:**
     ```bash
     gunicorn server:app
     ```
   * **Instance Type:** `Free`.
5. Click **Create Web Service**.
6. Render will automatically build the service and issue your live HTTPS URL (e.g., `https://ask-doc-dashboard.onrender.com`).

#### Option B: Deploy with Render Blueprint (`render.yaml`)
1. Click **New +** $\rightarrow$ **Blueprint**.
2. Connect your repository.
3. Render reads [`render.yaml`](file:///c:/Users/pipes/OneDrive/Documents/Google_AntiGravity/Project/Doc_Dashboard/render.yaml) and automatically configures the web service, build commands, and start commands without manual input.

---

## 🛠️ File Structure

```text
Doc_Dashboard/
├── Procfile                    # Render process entry point (gunicorn server:app)
├── render.yaml                 # Render Infrastructure-as-Code Blueprint
├── requirements.txt            # Python dependencies (flask, gunicorn, pandas, openpyxl, xlrd)
├── server.py                   # Production Flask server with /api/upload endpoint
├── update_all.py               # Master CLI auto-detection & batch update engine
├── update_data.py              # Parser for Technical Document Log Report
├── update_vendor_data.py       # Parser for Vendor Document Log Report
├── update_mdr_data.py          # Parser for MDR Master Document Register (5 tabs)
├── index.html                  # Dashboard 1: Technical Document Log
├── app.js                      # Logic for Dashboard 1
├── data.js                     # Dataset for Dashboard 1
├── index_vendor.html           # Dashboard 2: Vendor Document Log
├── app_vendor.js               # Logic for Dashboard 2 (tier drilldown, package rankings)
├── vendor_data.js              # Dataset for Dashboard 2
├── index_mdr.html              # Dashboard 3: MDR Revision Lifecycle
├── app_mdr.js                  # Logic for Dashboard 3 (IFR->IFA->AFC->AP, TOC bottleneck)
├── mdr_data.js                 # Dataset for Dashboard 3
├── styles.css                  # Core CSS design tokens, modals, tables, and themes
├── styles_mdr.css              # MDR lifecycle & executive analysis styling
├── latest_update.json          # Audit metadata recording latest file sources & cut-off dates
└── .gitignore                  # Ignore rules for python cache and excel lock files
```

---

## 📋 Weekly Maintenance Runbook

When you receive new weekly reports:
1. **On the Web (Render.com):**
   * Simply open your live website, click **`📁 Upload New File`**, and drop the new file in. The server handles extraction, recalculation, and page reload automatically.
2. **Via Git / Local:**
   * Copy the new `.xlsx` / `.xls` files into this directory.
   * Run:
     ```bash
     python update_all.py
     ```
   * Commit and push to GitHub:
     ```bash
     git add .
     git commit -m "Update weekly cut-off data"
     git push origin main
     ```
   * Render will automatically detect the commit and redeploy your live site in ~60 seconds!
