## How to use
To run the PowerShell scripts here you have to enable remote script execution in PowerShell elevated session, i.e. Run As Administrator:
`Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process`.

To see more about PowerShell execution policies, go to https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-5.1 .

Last tested against ORBBIDental version 1.3.15.

## Read-OrbbiDentDatabase
Reads the `p.dent` file from `C:\OrbbiData` and outputs details about your patients in a CSV file. Intended for migrating your patients.
