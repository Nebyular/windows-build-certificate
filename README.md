# windows-build-certificate
#### Language: C# .NET
#### Application Type: WPF Form
#### Requirements:  .NET v4.6.1 Runtime; Windows 10 recommended 1709 and onwards

This is replacing an existing closed source tool relying on Powershell window building and Powershell queries which currently presented as a cumbersome and slow user experience.

This new tool is a C# .NET program that displays Information about Windows, installed Microsoft software, hardware information, network information, domain information and domain roles. The program uses Windows Management Instrumentation Queries where it can to retireve information, otherwise .NET classes are used to retrieve information if unable with WMI. At the current moment it is just a direct replacement with most relevent features ported over with some additions from Windows 10.

The tool will mainly be used for information gathering for System Administrators to troubleshoot and resolve issues.

#### Note
The coding on this isn't the greatest so there's plenty of ways to do certain methods more efficiently. The tool is not completed so certain features may also be missing/broken. No support is provided for this tool, use at your own risk.
