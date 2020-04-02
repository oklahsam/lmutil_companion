I made these scripts to help view the status of an Autodesk license server.  
This can also be used for other license servers that run on the "lmutil/lmtools" server software, but this one is set up
specifically for Autodesk.  

There are a few comments in the license.ps1 script that explain different parts of the scripts.


# Edit: 4/2/2020
This script relies on Autodesk continuing to maintain their Feature Code lists, which you can find links to here: 
https://knowledge.autodesk.com/customer-service/network-license-administration/managing-network-licenses/interpreting-your-license-file/feature-codes

As of this edit, 2021 Product Licenses are available, but Autodesk does not have the 2021 Feature Code page up, so the script is broken until they do put it up if you've already added 2021 licenses to your license server. I may go through and modify the code to work without Autodesk's lists and just use the existing Feature Codes, if the 2021 list continues to not be available. For now, I'm just going to wait and see what happens.
