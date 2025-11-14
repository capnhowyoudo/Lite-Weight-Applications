Summary
This script is designed to help users convert .reg file entries into PowerShell scripts that can be used to recreate registry keys and values. 
It includes a simple WPF GUI with useful features like syntax checking, registry path testing, and theme switching. 
It’s particularly useful for automating registry edits, converting registry backups, or generating scripts from registry settings for system administration tasks.

1. Load .reg File 📂

Icon: 📂 (Folder/File)
The user selects a .reg file.
The file content is displayed in the txtReg textbox.

2. Convert to PowerShell Script 🔄
Icon: 🔄 (Arrows forming a circle or transformation)
The user clicks Convert, which triggers the Convert-RegToPS function.
The registry entries are converted into PowerShell commands and displayed in the txtPS textbox.

3. Check PowerShell Script Syntax ✅
Icon: ✅ (Check Mark)
The user clicks Check Syntax to validate the PowerShell script.
The script is parsed for errors. If there are no issues, a success message appears. If errors are found, the user gets an error message.

4. Test Registry Path 🔍
Icon: 🔍 (Magnifying Glass)
The user inputs a registry path and clicks Test Path to verify if the path exists in the system's registry.
It checks if the path exists and displays a success or warning message.
5. Save PowerShell Script 💾
Icon: 💾 (Floppy Disk)
The user clicks Save .ps1 to save the converted PowerShell script as a .ps1 file.
You can specify the location and filename for the script file.

6. Toggle Theme 🌙☀️
Icon: 🌙 (Moon) / ☀️ (Sun)
The user can switch between Light and Dark themes to adjust the appearance of the GUI.
Clicking the theme toggle button changes the interface between light and dark modes.
Visual Breakdown with Icons:
Load .reg File 📂: Select and load a .reg file into the GUI.
Convert to PowerShell 🔄: Convert registry keys and values to PowerShell script.
Check Syntax ✅: Validate the syntax of the generated PowerShell script.
Test Registry Path 🔍: Test whether the registry path exists on the system.
Save .ps1 Script 💾: Save the PowerShell script as a .ps1 file.
Toggle Theme 🌙☀️: Switch between dark and light themes in the GUI.
