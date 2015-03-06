## Work-Office2013Installer
========================

This program uses a Microsoft Tool to create an offline (USB/DVD) install for Microsoft Office 2013 & 365 "Click-To-Run" versions (Pretty much anything sold in the US that's a digital download or retail key only package).  The Microsoft Tool requires a configuration XML.  My program creates one to the users specification and simplifies the experience.

How to use:
1. Download the Office Deployment Tool & place setup.exe into the Setup folder.
2. Run Office2013Setup.exe with /download command switch (Office2013Setup.exe /download) (I'm working on a better solution atm.).
3. The Setup.exe program will run and download necessary files.  It's a little over 1GB.
4. Once Setup.exe has closed.  Run Office2013Setup.exe.
5. Select the product you have.  Feel free to customize the install to exclude programs you do not want.


Notes:
- You can rerun the download to get an updated copy.
- You can rerun the installer to add or remove programs at any time (Excel, OneDrive, etc).

