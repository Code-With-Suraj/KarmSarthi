# KarmSarthi 🛡️

> **"Har din, har chhutti ka bharosa"**

![KarmSarthi Logo](https://media.licdn.com/dms/image/v2/D5622AQF9EXLdfJsDAg/feedshare-shrink_1280/B56ZtVIUq9HkAs-/0/1766659804554?e=1768435200&v=beta&t=O62do53dbIFIIGlNq4Ov5n6VAPnmkwDfeJX-Ylg0QXw)

## 📋 About The Project

**KarmSarthi** is a comprehensive, professional **Leave Management & Attendance System** built on Google Apps Script. It helps organizations streamline their workforce management by digitizing leave requests, automating attendance tracking, and determining shifts, all while maintaining a user-friendly and responsive interface.

Designed for modern teams, it bridges the gap between Employees, Managers, and HR with distinct, role-based workflows or "Views".

## 🚀 Key Features

### 📅 Advanced Leave Management
*   **Seamless Application:** Employees can apply for casual, sick, emergency, or personal leaves.
*   **Flexible Durations:** Supports **Full Day**, **First Half** (Morning), and **Second Half** (Afternoon) requests.
*   **Real-Time Balances:** Auto-calculation of leave balances and "Insufficient Balance" warnings with salary deduction alerts.
*   **Approval Workflow:** Managers receive email notifications to Approve or Reject requests with comments.

### 📍 Smart Attendance Tracking
*   **GPS-Fenced Check-In:** Enforces location validation (Geofencing) for **Office** mode to ensure employees are on-site.
*   **Multiple Work Modes:**
    *   🏢 **Office:** Requires GPS location within a specific radius.
    *   🏠 **Work From Home (WFH):** Tracks attendance without strict location binding.
    *   🚗 **On-Duty:** For field visits and client sites.
*   **Live Statistics:** View "Present Days", "Attendance Rate", and daily work hours in real-time.
*   **Regularization:** Employees can request regularization for missed punches (e.g., forgot to check in) which managers can approve.

### 🕐 Shift Management
*   **Dynamic Rostering:** Support for multiple shifts (General, Morning, Evening, Night).
*   **Late Marking Logic:** Automatically marks attendance as "Late" if check-in exceeds the shift start time + grace period.

### 👥 HR & Admin Tools
*   **Employee Onboarding:** Digital document collection and verification process.
*   **Document Verification:** HR view to verify uploaded documents (Aadhaar, PAN, etc.) with preview capabilities.
*   **Salary Slips:** Generate and download professional PDF salary slips with detailed breakdown of earnings and deductions.
*   **Team Calendar:** Managers can view the entire team's leave schedule at a glance.

## 🛠️ Technology Stack

*   **Backend:** Google Apps Script (GAS)
*   **Database:** Google Sheets
*   **Frontend:** HTML5, CSS3, JavaScript (Vanilla)
*   **Assets:** SVG Icons, Mobile-Responsive Design

## 📱 User Interface

KarmSarthi features a modern, clean, and mobile-responsive UI that works seamlessly across desktop and mobile devices.
*   **Employee View:** For applying leaves, checking balances, and marking attendance.
*   **Manager View:** For approving requests and viewing team availability.
*   **HR View:** For document verification and employee management.
*   **Attendance View:** Dedicated interface for punch-in/punch-out.

## ⚙️ Setup & Deployment

1.  **Deploy as Web App:** This project is designed to be deployed as a Google Apps Script Web App.
2.  **Permissions:** Execute as "User accessing the web app" to ensure capturing the correct email identity.
3.  **Sheet Configuration:** The system automatically initializes the required Google Sheets (`Employees`, `Leave Requests`, `Attendance Records`, etc.) upon first run.

## 📄 License

Internal Use - Proprietary Software.

---
*Built with ❤️ by the KarmSarthi Team*
