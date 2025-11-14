This PowerShell script is a GUI-based application that allows users to compile PowerShell (.ps1) scripts into executable (.exe) files using the PS2EXE module. The script creates a user-friendly interface with several features to customize the compilation process. Here's a breakdown of what the script does, with added icons for each key step:

Step-by-Step Overview with Icons
1. GUI Initialization
Icon: ⚙️
The script initializes a Windows Presentation Framework (WPF)-based GUI for the PowerShell-to-EXE compiler, allowing users to easily input the necessary data for the conversion process.

2. Check PS2EXE Module
Icon: 🔍
The script verifies if the PS2EXE module is loaded. If not, it attempts to load it or install it by adjusting the ExecutionPolicy if necessary.


If the module isn't available, it tries to bypass ExecutionPolicy restrictions temporarily, or if that fails, it requests administrator permissions to adjust the ExecutionPolicy and install the module.


Logs are displayed for each step of the process.

3. File Input (Select PowerShell Script, Output EXE, and Icon)
Icon: 📂

Users can select:

Input .PS1: Browse for the PowerShell script to convert.

Output .EXE: Select the desired output location for the EXE file.


Icon (.ICO): Choose an icon to embed into the EXE file.



4. Metadata Configuration
Icon: 📝
The user can input metadata to embed in the executable:


Product name


Description


Company


Version



5. Options Configuration
Icon: ⚙️
Users can configure several options for the EXE:


Hide Console: Hide the PowerShell console when the EXE runs.


Suppress Errors: Prevent error messages from displaying.


No Output: Disable output messages.


Require Admin: Set the EXE to always run as an administrator.


Architecture: Choose either x64 or x86.



6. Show Command
Icon: 💬
Users can view the generated PS2EXE command based on their selections. This command is displayed in a message box for reference.

7. Progress and Logging
Icon: ⏳
A ProgressBar displays the compilation process. A log output box is also available to track the script's status, with real-time updates (e.g., "Compiling...", "Compilation complete!", etc.).

8. Save and Load Presets
Icon: 💾
Users can:


Save Settings: Save their configuration as a JSON file for future use.


Load Settings: Load previously saved settings to easily reuse them in future compilations.



9. Save Log
Icon: 📜
Users can save the log output to a text file for later reference.

10. Compile the EXE
Icon: 🚀
When the Compile EXE button is pressed, the script uses the PS2EXE module to compile the PowerShell script into an executable file with the specified settings.


During compilation, the progress bar updates, and logs are generated.


If successful, the user is notified that the EXE was created.


If there’s an error, a message is shown in the logs.



11. Exit Application
Icon: ❌
The Exit button closes the application.

Summary of Key Features


PS2EXE Module Integration: Converts PowerShell scripts to executables.


File Browsing: Choose input, output, and icon files.


Metadata Embedding: Add details like product name, version, and description.


Customization Options: Toggle console visibility, suppress errors, and set admin privileges.


Architecture Selection: Choose between 64-bit (x64) or 32-bit (x86).


Logs and Progress: View logs and track progress with a progress bar.


Preset Support: Save and load configuration presets for future use.


Execution Policy Fix: Automatically fixes Execution Policy restrictions to load the PS2EXE module.


This script provides a convenient and fully-featured interface for compiling PowerShell scripts into standalone executables, with customization and ease of use in mind.
