This PowerShell script is a GUI tool that allows you to retrieve and display saved Wi-Fi profiles and their passwords (if available) from your Windows machine. It uses a WPF (Windows Presentation Foundation) interface to make the process more user-friendly and displays the information in a console-like text output. It also includes a mechanism to automatically elevate the script to run as an administrator, as the netsh wlan command requires elevated permissions to retrieve passwords.

1. Key Features:

2. Administrator Elevation:
The script checks if it's being run as an administrator. If it's not, it automatically restarts itself with elevated privileges. This is important because retrieving Wi-Fi passwords requires administrator access.

3. GUI-Based Interface:
The GUI has a button (Get Wi-Fi Passwords) to trigger the process of retrieving stored Wi-Fi profiles. It also has a scrollable text box to display the results in a console-style format.

3. Retrieve Wi-Fi Profiles and Passwords:

The script uses the netsh wlan show profiles command to list all saved Wi-Fi profiles on the system.

For each profile, it attempts to extract the associated password (if available) by running the netsh wlan show profile name="ProfileName" key=clear command.

The password is displayed in the text output if it is accessible. If no password is set or it’s hidden, the script will display "[Not Set / Hidden]" for that profile.

4. User Feedback with Icons:
The script uses icons (emojis) for user feedback in the output, such as:

A magnifying glass (🔍) when retrieving profiles.

A checkmark (✅) to indicate that the process is complete.
