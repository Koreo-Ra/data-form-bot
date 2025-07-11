# 🧾 data-form-bot — Autofill Web Forms from Google Sheets

## Description

`data-form-bot` is an automation tool built for transferring applications from **Google Sheets** to a **web form**.  
Originally created to automate the submission of customer requests for *Ukrtelecom's optical internet service*, this tool significantly reduces manual data entry time.

⚠️ This repository contains a **demo version** with all sensitive data removed or replaced. You can adapt it to other projects and websites.

---

## 💡 Features
- Reads data from a Google Spreadsheet
- Parses street, house, flat, region, district, city using a mix of custom rules and Google Maps API
- Validates and normalizes phone numbers
- Fills out a web form using Selenium
- Supports resume mode — remembers the last submitted row
- Designed for scalability and adaptation to other form types

---

## 🛠 Technologies
- Python
- Google Sheets API
- Google Maps Geocoding API
- Selenium
- Regular Expressions
- Chrome WebDriver

---

## 🚀 How to run
1. Clone the repository:
```bash
git clone https://github.com/your-username/data-form-bot.git
cd data-form-bot