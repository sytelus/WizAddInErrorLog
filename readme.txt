What' This:
----------------
This is an add-in (i.e. plug in) for Microsoft Visual Basic 6.0. It does following of things:
1. Adds/removes line numbers from Visual Basic projects so that you can use Erl function to exactly locate where the error occurred. Unlike other add-ins, this one has been thoroughly tested on live complex projects for more then a year and has lots of code to workaround VB IDE object model bugs. It comes with source code, does numbering on whole project group and also does smart numbering (for example it won't number trivial assignment statements in Property Get).
2. With a click of a button, this add-in can also add error handlers in all methods. Again unlike other similar utilities, this one has been thoroughly tested, accepts a directive to turn off error handler and has some decision making code to decide whether an error handler would effect actual logic.
3. Fix for VB's recent menu item bug (MRU). You might have noticed VB's Recent project item menu always gets messed up. This add-in would fix them. You can see the screenshot in WizAddIn.jpg.

Included with this project,
-Sample error handler code for Vector project.
-Developer's manual on how to use add-in and error logging.
-Test program which uses Add-In for error tracing purposes and shows it in nice dialog box. It also produces screenshot and error log in Word document. You will need VSFlexGrid control for this.


How To Install Add-In:
--------------------------------
Execute WizAddInInstall.exe. This will add an entry in to VBAddins.ini and register the DLL.


This Project's Homepage:
------------------------
You can find this project's latest updates, more info, get compiled binaries and request support at http://www.ShitalShah.com. Check for the programmer's tools, geeks only or downloads sections.


Who wrote it?
-----------
This is Shital Shah. Visit my web site at http://www.ShitalShah.com.


Source Code:
--------------------
This program comes with Source Code. It's freeware but you can't redistribute modified source code without author's explicit permission.



Version History:
----------------
Spring 2001	1.0.0.0	Created
Jan 2002		1.0.0.0	Fixed bugs while testing on Vector projects
May 2002		1.1.0.0	Fixed bugs and released to public
Jul 2002		1.2.0.0	Rewrote line number and error handlers code to use same functions


