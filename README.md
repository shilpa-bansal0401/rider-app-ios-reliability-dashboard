# 📊 Rider iOS Crash & Hang Analysis Dashboard

Automated Sentry-based dashboard to analyze **fatal crashes** and **app hangs** for Rider iOS, with monthly trends, weekly insights, and categorized grouping for faster debugging and prioritization.

---

## 🚀 Overview

This project generates interactive HTML reports using Sentry data to provide:

- 📅 **Monthly comparison** (Feb, Mar, Apr 2026)
- 📈 **Weekly trend analysis** (April breakdown)
- 🧩 **Categorized issue grouping** (Mapbox, Watchdog, QUIC, etc.)
- 👥 **Deduplicated user impact metrics**
- 🔗 **Direct Sentry deep links** for investigation

---

## 🎯 Purpose

- Improve **App Quality Score (AQS)**
- Reduce **Mean Time to Detect (MTTD)**
- Identify **high-impact failure categories**
- Support **data-driven prioritization** for stability improvements

---

## 🏗️ How It Works

The scripts:
1. Query Sentry using:
   is:unresolved level:fatal handled:no

2. Group issues into predefined categories:
   - QUIC Protocol  
   - Data Sync Upload  
   - Watchdog / Fatal Hangs  
   - Mapbox  
   - Location  
   - Core Logic / Concurrency  
   - UIKit / Foundation  

3. Deduplicate issues:
   - Each issue is counted once
   - Assigned based on priority order

4. Generate:
   - Overview dashboard
   - Monthly breakdown tabs
   - Weekly trend tab

---

## 🛠️ Setup

### 1. Clone repo
git clone <repo-url>
cd rider-ios-crash-hang-analysis

### 2. Set Sentry token
export SENTRY_AUTH_TOKEN=your_token_here

### 3. Run report

#### Crash report
python3 generate_crash_report.py

#### Hang report
python3 generate_hang_report.py

---

## 📂 Output

Reports are generated as static HTML:

crash_report/index.html  
hang_report/index.html

---

## 🌐 Sharing the Report

Since Confluence doesn’t support HTML directly:

### Recommended approach
- Deploy the report via GitHub Pages
- Share the GitHub Pages link in Confluence

### Alternative
- Export as PDF
- Add screenshots + tables in Confluence

---

## 📊 Features

- Interactive charts (Chart.js)
- Clickable rows → opens Sentry query
- Delta comparison (week-over-week / month-over-month)
- Clean UI for stakeholder sharing

---

## ⚠️ Notes

- Data is deduplicated by issue
- April data is partial (ongoing month)
- Category accuracy depends on Sentry query quality

---

## ⚙️ Automated Report Generation

Reports are automatically generated and committed to this repo every **Monday and Thursday at 11:25 AM CST**.

---

## 🔮 Future Improvements

- Slack / email alerts
- Cross-platform support
- Unified crash + hang comparison

---

## 🏷️ Tags

ios, sentry, crash-analysis, app-hangs, reliability, dashboard, observability, aqs

---

## 👤 Author
Shilpa Bansal
