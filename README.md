# event-driven-program-demo
A Windows Forms Application demonstrating Visual Basic .NET + Microsoft Access Database Integration

## Overview
- Database connectivity using **ADO.NET** and **Microsoft Access**
- **CRUD operations** on `Employees` table
- **Payroll calculation** with overtime logic
- **Data validation**, **error handling**, and **clean UI**

---

## Features
| Feature | Implemented |
|-------|-------------|
| Connect to `EmployeeDB.accdb` | Yes |
| Display employees in `DataGridView` | Yes |
| Add new employee with validation | Yes |
| Search & select employee by double-click | Yes |
| Payroll form with real-time calculation | Yes |
| Overtime (1.5x rate if >40 hrs) | Yes |
| Save payroll record to database | Yes |
| Input validation & clear form | Yes |

---

## Database
- File: `EmployeeDB.accdb`
- Tables:
  - `Employees` (EmployeeID, FirstName, LastName, Department)
  - `Payroll` (PayrollID, EmployeeID, HoursWorked, HourlyRate, TotalPay, Overtime)

> **Note:** The `.accdb` file is included. Copy to output directory: **"Copy if newer"**

---

## How to Run
1. Open in **Visual Studio** (Windows Forms App .NET Framework)
2. Build → Run (`frmPayroll` starts)
3. Add employees → double-click to select → calculate pay

---

## Project Structure
/
├── EmployeeDB.accdb
├── frmEmployees.vb
├── frmPayroll.vb
├── Program.vb
└── README.md
