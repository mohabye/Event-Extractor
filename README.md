# Event-Extractor

Overview
The Event Viewer Filter Tool is a PowerShell script that provides a graphical user interface (GUI) for filtering and exporting Windows Event Log entries to a CSV file. 📊 This tool allows users to specify date ranges, event IDs, event levels, sources, keywords, and time zones to retrieve and analyze event log data efficiently. 🚀 It’s designed for system administrators and IT professionals who need to monitor and investigate system, application, security, and setup logs.

Features
Flexible Filtering: Filter events by start and end date/time, event IDs (single or comma-separated), event levels (Information, Warning, Error, Critical, or All), sources (Application, System, Security, Setup, or All), and optional keywords. 🔍
Time Zone Support: Select from a comprehensive list of time zones worldwide or use the local system time zone to adjust event timestamps. 🌍
Multiple Event ID Support: Enter multiple event IDs as a comma-separated list (e.g., 4642, 7001) or retrieve all event IDs. 📋
Export to CSV: Export filtered events to a CSV file with columns for TimeCreated, ID, LevelDisplayName, ProviderName, and Message. 📝
User-Friendly Interface: A GUI built with Windows Forms for easy interaction and configuration. 🖥️
Interactive Feedback: Provides real-time feedback via message boxes for errors, no results, or successful exports, and console output for debugging. ✅
Requirements
Operating System: Windows (compatible with Windows 7 and later). 🪟
PowerShell: PowerShell 5.1 or later (installed by default on most Windows systems). 🛠️
Permissions: Administrative privileges are required to access certain event logs (e.g., Security log). 🔐
Dependencies: The script uses the System.Windows.Forms and System.Drawing assemblies, which are part of the .NET Framework (included in Windows). 📚
Installation
Save the script with a .ps1 extension (e.g., EventViewerFilterTool.ps1). 💾
Ensure PowerShell execution policy allows running scripts:
Open PowerShell as Administrator and run:
powershell

Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
Confirm the change by typing Y and pressing Enter. 👍
Run the script with administrative privileges:
Right-click the script file and select "Run with PowerShell" or open PowerShell as Administrator and navigate to the script’s directory, then run:
powershell

.\EventViewerFilterTool.ps1

Usage
Launch the Tool: Run the script as described in the Installation section. A GUI window titled "Event Viewer Filter Tool" will appear. 🚀
Configure Filters:
Start Date/Time and End Date/Time: Use the date-time pickers to specify the time range for events (format: MM/dd/yyyy HH:mm:ss). 📅
Event IDs: Enter one or more event IDs as a comma-separated list (e.g., 4642, 7001), or check "All Event IDs" to retrieve all events. 🔢
Event Level: Select an event level (Information, Warning, Error, Critical, or All) from the dropdown. ⚠️
Source: Choose the event log source (Application, System, Security, Setup, or All) from the dropdown. 🗂️
Keyword (optional): Enter a keyword to filter events by message content or event ID (numeric keywords match event IDs). 🔍
Time Zone: Select a time zone from the list (sorted alphabetically) or use "Local Time" to apply the system’s local time zone. 🌐
Output File Path: Specify the path for the CSV output file, or use the "Browse" button to select a location and filename. 📂
Export to CSV: Click the "Export to CSV" button to retrieve and export the filtered events. The button will turn light blue during processing and return to light gray afterward. 📤
Review Results: If events are found, they will be exported to the specified CSV file, and a success message will appear. If no events match the criteria, an error message will notify you.







![image](https://github.com/user-attachments/assets/9648f95b-ee3d-462a-a442-ac614878a068)
