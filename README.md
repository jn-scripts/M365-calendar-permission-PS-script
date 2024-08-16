Microsoft 365 Calendar Permission Script

This PowerShell script is designed to streamline the process of managing calendar permissions within Microsoft 365. It allows administrators to assign specific calendar permissions to users in multiple languages, ensuring that the script can be used in a wide range of environments.

Supported Languages
The script supports calendar folder names in the following languages:

English
Spanish
French
German
Italian
Portuguese
Dutch
Chinese (Simplified)
Japanese
Korean
Russian
Arabic
Turkish
Hindi
Bengali
Urdu
Swedish
Polish
Danish
Finnish
Greek
Hebrew
Thai
Vietnamese
Norwegian




This script requires the ExchangeOnlineManagement module. If the module is not installed, the script will prompt the user to install it.
Admin UPN: The script requires the User Principal Name (UPN) of an administrator account on the Microsoft 365 tenant to log in.

The script allows the administrator to apply the same permission settings to additional users if needed.

The script terminates after all permissions have been assigned.

Help Commands: Enter ? at any prompt to receive more detailed instructions.

Error Handling: The script includes basic error handling to ensure smooth execution, including input validation and module installation checks.
Written by jn-scripts.

License
This script is provided as-is without warranty of any kind. Use it at your own risk.
