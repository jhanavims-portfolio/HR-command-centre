# HR-command-centre
An Excel-based HR Recruitment and Onboarding Analytics project for tracking candidates, interview progress, compliance, offers, joining status, induction, and overall onboarding progress.
# HR Recruitment & Onboarding Analytics

## 📌 Project Overview

This project is an **Excel-based HR Recruitment and Onboarding Analytics system** designed to track candidates throughout the recruitment and onboarding process.

It helps HR teams monitor candidate progress from **application and interview stages to selection, compliance, offer, joining, induction, and onboarding**.

The project demonstrates how Excel formulas and data tracking can be used to manage HR operations efficiently.

---

## 🎯 Project Objectives

* Track candidate recruitment progress.
* Monitor interview status.
* Track selected candidates.
* Identify compliance-related pending activities.
* Monitor offer and joining status.
* Track induction and onboarding progress.
* Automatically generate the overall candidate status.

---

## 📊 Recruitment and Onboarding Process

The project tracks the following stages:

**Candidate Application → Interview → Selection → Compliance → Offer → Joining → Induction → Onboarding**

---

## 🛠️ Tools Used

* Microsoft Excel
* Excel Functions
* VLOOKUP
* INDEX and MATCH
* IF and IFERROR
* AND Function
* Data Validation
* Conditional Logic

---

## 📂 Project Structure

### 1. Interview Tracker

Tracks candidate interview details and recruitment status.

Information includes:

* Candidate Name
* Interview Status
* Selection Status
* Interview Progress

---

### 2. Onboarding Progress

Tracks candidates after selection.

Information includes:

* Candidate Name
* Compliance Status
* Offer Status
* Joining Status
* Induction Status
* Onboarding Progress

---

### 3. Automated Status Tracking

Excel formulas are used to automatically determine the candidate's overall status.

Example statuses include:

* Compliance Pending
* Offer Extended – Awaiting Joining
* Induction Overdue
* Awaiting Interview
* On Track
* Not Onboarded

---

## ⚙️ Key Excel Formulas

### VLOOKUP for Onboarding Status

```excel
=IFERROR(VLOOKUP(A9,Onboarding_Progress!$A$2:$D$9,4,FALSE),"Not Onboarded")
```

This formula searches for a candidate in the **Onboarding Progress** sheet and returns the relevant onboarding information.

If the candidate is not found, it displays:

**Not Onboarded**

---

### Automated Candidate Status

```excel
=IF(E7="Compliance Pending","Compliance Pending",
IF(AND(F7="",D7="Selected"),"Offer Extended - Awaiting Joining",
IF(AND(F7<>"",G7="N"),"Induction Overdue",
IF(D7="Not Yet Interviewed","Awaiting Interview","On Track"))))
```

This formula automatically checks multiple conditions and generates the appropriate candidate status.

---

## 💡 Key Features

* Automated candidate tracking.
* Interview status monitoring.
* Recruitment pipeline management.
* Compliance tracking.
* Offer and joining status tracking.
* Induction monitoring.
* Automated onboarding status.
* Error handling using IFERROR.
* Lookup functions for connecting multiple sheets.

---

## 📈 Business Value

This project helps HR teams:

* Reduce manual tracking.
* Improve recruitment visibility.
* Identify pending compliance activities.
* Monitor candidates awaiting joining.
* Identify induction delays.
* Maintain an organized recruitment and onboarding process.

---

## 🚀 Future Improvements

Future enhancements could include:

* HR Analytics Dashboard
* Power BI Integration
* Recruitment KPI Tracking
* Automated Email Notifications
* Candidate Turnaround Time Analysis
* Recruitment Funnel Analysis
* Attrition Tracking
* Automated Reporting

---

## 👤 Author

M.S JHANAVI

MBA – HR 
KL University

---

## ⭐ Conclusion

This project demonstrates the practical use of **Microsoft Excel for HR Recruitment and Onboarding Operations**. It provides a structured approach to tracking candidates and automating status updates throughout the recruitment and onboarding process.
