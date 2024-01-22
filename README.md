# Windows 60Hz automatic refresh rate switch on battery
I've set this as a function available in my Powershell 7 (pwsh) using Powershell modules. 
The function is called from a task I made in task scheduler, of which you see the XML.
The function will switch the rate only if one monitor is currently being used (Qres doesn't correctly support multiple monitors, it will apply the refresh rate to every monitor, which is not granted to be good if you got different refresh rate monitors), and on power sets the maximum available refresh rate for any computer you might be using. On battery it sets the refresh to 60Hz.

It requires QRes to be installed: https://www.majorgeeks.com/files/details/qres.html
