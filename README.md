# Windows-Update-Automation-Script-Project

A PowerShell script that automates Windows updates, includes logging, error handling, and can create restore points before updating. Think: sysadmins

Running with Admin. privileges in Microsoft PowerShell, we can check for avaliable windows updates, create a sys. restore point before updating as a safety net, download and install updates, and log all operations. With a cohesive color-coded console output, error handling, and the creation of a detailed summary report to boot.

Example console output with user input in PowerShell:

PS C:\Users\evanj\Downloads\ResumeProjects>  .\WindowsUpdateAutomation.ps1
[2026-01-01 15:14:42] [INFO] ========================================
[2026-01-01 15:14:42] [INFO] Windows Update Automation Script v1.0
[2026-01-01 15:14:42] [INFO] Author: Evan Jania
[2026-01-01 15:14:42] [INFO] ========================================
[2026-01-01 15:14:42] [SUCCESS] Running with Administrator privileges
[2026-01-01 15:14:42] [INFO] Checking for available Windows updates...
[2026-01-01 15:14:42] [INFO] Searching for updates (this may take a few minutes)...
[2026-01-01 15:14:54] [INFO] Found 1 available update(s)
[2026-01-01 15:14:54] [INFO]   - Security Intelligence Update for Microsoft Defender Antivirus - KB2267602 (Version 1.443.454.0) - Current Channel (Broad)
Do you want to proceed with downloading and installing these updates? (Y/N): N
[2026-01-01 15:15:01] [WARNING] Update process cancelled by user
Press Enter to exit:
