# 🇮🇳 PIB UPSC Daily Intelligence & Headline Automation

An end-to-end automated pipeline that extracts daily press releases from the Press Information Bureau (**`https://www.pib.gov.in/allRel.aspx?reg=48&lang=1`**), filters out ceremonial government fluff, aligns high-yield news to the **UPSC Civil Services Examination (CSE) Syllabus (GS 1, 2, 3, 4)**, and compiles them daily into formatted **Word (`.docx`)** and **Excel (`.xlsx`)** documents.

---

## 🌟 Key Highlights & Engineering Features

- **Reverse-Engineered ASP.NET PostBack Handling**: Overcomes PIB's legacy ASP.NET WebForms architecture by dynamically passing `__VIEWSTATE` and form controls to seamlessly query both today's live feed and historical dates.
- **Fast Reader-Pane Extraction**: Directly requests `PressReleasePage.aspx?PRID={id}`, bypassing heavy client-side JavaScript shells for sub-100ms response times.
- **Two-Tier UPSC Syllabus Classifier**:
  - **Tier 1 (Noise Filter)**: Strips routine ceremonial releases (sports congratulations, festival greetings, condolences, protocol events).
  - **Tier 2 (Taxonomy Mapping)**: Classifies releases into **GS-1**, **GS-2**, **GS-3**, and **GS-4** with topic tags and relevance explanations.
- **Dual Formatted Output**:
  - **Word Brief (`.docx`)**: Styled executive brief with executive metrics, grouped by GS Papers, topic keywords, syllabus alignment reasons, and direct links.
  - **Excel Tracker (`.xlsx`)**: 2 sheets ("UPSC Relevant News" with color-coded GS Papers and auto-wrapping + "All PIB Releases" full audit log).
- **Flexible Execution**:
  1. **Cloud-Native Automation (GitHub Actions)**: 100% automated in the cloud without requiring your PC to be on.
  2. **Local Windows Task Scheduler**: 1-command installer (`python main.py --install-task 21:00`).
  3. **Interactive Web Dashboard**: Streamlit interface (`streamlit run app.py`) to visually explore releases and download on demand.

---

## 📁 Repository Structure

```
PIB Automation/
├── config.py                 # Core constants, timeouts, and UPSC syllabus keyword taxonomy
├── pib_scraper.py            # Resilient PIB scraper with ViewState & retry logic
├── upsc_classifier.py        # 2-Tier noise filtering & UPSC GS 1-4 syllabus alignment
├── doc_builder.py            # Document generators for Word (.docx) & Excel (.xlsx)
├── main.py                   # Central CLI runner & Windows scheduler installer
├── app.py                    # Interactive Streamlit Web Dashboard
├── requirements.txt          # Python dependencies
├── .github/
│   └── workflows/
│       └── daily_pib_upsc.yml # GitHub Actions workflow for 100% free cloud automation
└── output/                   # Directory where generated daily briefs are saved
```

---

## 🚀 Quickstart Guide

### 1. Installation

Ensure Python 3.10+ is installed, then install the required packages:

```bash
pip install -r requirements.txt
```

---

### 2. Running via Command Line

#### Fetch today's news and generate documents:
```bash
python main.py --today
```

#### Fetch any specific past date:
```bash
python main.py --date 2026-10-04
```

#### Run and immediately open the generated Word brief:
```bash
python main.py --today --open
```

All generated files are saved in the `output/` folder:
- `output/PIB_UPSC_Daily_YYYY-MM-DD.docx`
- `output/PIB_UPSC_Daily_YYYY-MM-DD.xlsx`

---

### 3. Launch the Interactive Web Dashboard

To visually browse news, filter by GS papers, and download files with 1 click:

```bash
streamlit run app.py
```

Features:
- Calendar date picker for any past date.
- Real-time classification metrics and GS paper counts.
- Dedicated tabs for **GS-1**, **GS-2**, **GS-3**, and **GS-4**.
- Single-click download buttons for `.docx` and `.xlsx`.

---

### 4. Setting Up Daily Automation at a Fixed Time

#### Option A: 100% Cloud-Native Automation (GitHub Actions) ⭐ *(Recommended)*
1. Push this folder to a GitHub repository (public or private).
2. The workflow file [daily_pib_upsc.yml](file:///.github/workflows/daily_pib_upsc.yml) is pre-configured to run automatically every day at **9:00 PM IST (15:30 UTC)**.
3. It will run in GitHub's cloud, generate the files, commit them to an `output/` archive branch, and keep downloadable artifacts.
4. **Benefit**: Your computer does not need to be turned on.

#### Option B: Local Windows Task Scheduler
To automatically run the script on your local Windows PC every day at a fixed time (e.g. 9:00 PM):

```bash
python main.py --install-task 21:00
```
This registers a daily job named `PIB_UPSC_Daily_Automation` in your Windows Task Scheduler.

#### Option C: Background Python Loop / Daemon
To leave a background terminal process running that triggers at your chosen time:

```bash
python main.py --daemon 21:00
```

---

## 🎯 UPSC Syllabus Alignment Framework

| GS Paper | Focus Areas | PIB Target Themes |
|---|---|---|
| **GS-1** | Art, Culture, History, Geography, Society | Commemorations, freedom struggle icons, archaeological excavations, monuments, monsoons, cyclones, earthquakes, tribal culture. |
| **GS-2** | Governance, Constitution, Polity, Social Justice, IR | Cabinet approvals, Bills, Acts, Supreme Court rulings, welfare schemes (PM-Kisan, Ayushman Bharat), bilateral MoUs, multilateral summits (G20, ASEAN, Quad). |
| **GS-3** | Economy, Science & Tech, Environment, Security | GDP, inflation, RBI, GST, infrastructure, agriculture (MSP, irrigation), ISRO, DRDO, nuclear tech, wildlife (GIB, tiger reserves, IUCN), DRI drug busts, cyber security. |
| **GS-4** | Ethics, Integrity & Aptitude | Vigilance awareness, probity, administrative reforms, citizen charter, transparency. |

---

## ⚖️ License & Disclaimer

Designed for educational and preparation purposes for the UPSC Civil Services Examination. All source press releases are the property of the Press Information Bureau, Government of India.
