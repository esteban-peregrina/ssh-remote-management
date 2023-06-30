# ssh-remote-management
Set of PowerShell scripts (.ps1) that install OpenSSH Server on both 2019 and 2012 R2 Windows Servers, and proceed automatic reboots according to uptime preferences.

Additionnaly, the servers management script send an email to administrator's team converting data into HTML & CSS.

Note that 2012 R2 Windows Servers installation requires an external OpenSSH install folder that should be stored on a shared repository that as to be accessible by servers due to internet installation policies issues.

CAUTION: Those scripts are the result of one month of internship during my first year of french "CPI". I started with a first year of bachelor's degree C and Python level, and not a single shell scripting knowledge, neither in PowerShell, nor in OOP. This is why scripts feels very different between one and an other, as I was constantly learning and improving. Thus, please use those scripts carrefully, especially the install one and the archived one, as they may contain serious security issues and lack of errors management.
