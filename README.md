This Powershell Script file was tested under Microsoft Windows 11 (10.0.26200.7462 aka 25H2) German edition and Powershell 7.5.4.

First the Powershell Script looks at the entries of Microsoft Device Manager and collects all devices.

Then it opens notepad.exe and lists all drivers incl. Version in a CSV File.

Then it looks in an enhanced way for drivers for for Windows 11, 25H2 at Microsoft Catalog.

Then it checks all entries of the list regarding positive find (trigger: "Updates:") at Microsoft Catalog.

Last it opens one Tab in Chrome for every positive find.

Note: You need to download and update the latest version of the drivers yourself.


Example CSV File:

Line 1: Hersteller,Gerätename,Type,Vendor ID,Device ID,Installierte_Version

Line x: Intel,Intel(R) Management Engine Interface #1,Type=PCI,Vendor ID=8086,Device ID=7AE8,2540.8.7.0


Example Tab in Chrome:
https://www.catalog.update.microsoft.com/Search.aspx?q=windows%2011%20client%2C%20version%2025h2%20Intel%28R%29%20Management%20Engine%20Interface%20%231

Equals in search: windows 11 client, version 25h2 Intel(R) Management Engine Interface #1

Feel free to adapt it to your OS, Version and language.
