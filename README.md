# 🏋️ GymTrack

GymTrack is a beginner-friendly, terminal-based gym membership management system built using Python.

It helps gym staff manage member records and export the data to an Excel file.

## 🚀 Features

- Add new gym members
- View all gym members
- Update member details
- Delete members
- Choose membership plans
- Track payment status
- Store data using JSON
- Export member data to Excel
- Attractive terminal interface using Rich

## 🛠️ Technologies Used

- Python
- JSON
- Rich
- OpenPyXL
- Git
- GitHub

## 📋 CRUD Operations

| Operation | Function |
|----------|----------|
| Create | Add Member |
| Read | View Members |
| Update | Update Member |
| Delete | Delete Member |

## 📊 Excel Export

The application can export all gym member records to:

`exports/gym_members.xlsx`

## 📁 Project Structure

```text
gymtrack/
│
├── data/
│   └── members.json
│
├── exports/
│   └── gym_members.xlsx
│
├── venv/
│
├── main.py
├── requirements.txt
├── README.md
└── .gitignore