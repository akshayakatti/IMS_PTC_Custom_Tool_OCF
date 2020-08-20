' ===========================================================================
'
'   Name:   OpenContainingFolder.vbs
'
'   Desc:   Opens file explorer with current working directory of selected sandbox
'           member
'
'   Note:   This script must be configured within the Custom buttons
'           of the MKS Integrity Client Viewsets.  See below for details
'
'       1) Copy this script (OpenContainingFolder.vbs) to <Client Install Dir>\MKS\IntegrityClient2009\viewset
'       2) Go to Viewset -> Customize, select Action tab and Action Group "Custom"
'       3) Edit first free "User Action #"
'
'       Customize User Button
'       ---------------------
'        Name:             Open Containing Folder
'        Program:          wscript.exe
'        Parameters:       //H:WScript //NoLogo "<Client Install Dir>\MKS\IntegrityClient2009\viewset\OpenContainingFolder.vbs"
'        Description:      Launch explorer with current directory
'        Icon File:        OpenContainingFolder.gif
'        Environment File: -
' ===========================================================================

fexplore = "explorer.exe /e,"

' explorer.exe options
'   Option            Function
'   ----------------------------------------------------------------------
'   /n                Opens a new single-pane window for the default selection. This is usually the root of the drive that
'                     Windows is installed on. If the window is already open, a duplicate opens.
'   /e                Opens Windows Explorer in its default view.
'   /root,<object>    Opens a window view of the specified object.
'   /select,<object>  Opens a window view with the specified folder, file, or program selected.
'
'   Examples
'   -----------------------------------------------------------------------
'   Example 1: Explorer /select,C:\TestDir\TestProg.exe
'              Opens a window view with TestProg selected.
'   Example 2: Explorer /e,/root,C:\TestDir\TestProg.exe
'              Opens Explorer with drive C expanded and TestProg selected.
'   Example 3: Explorer /root,\\TestSvr\TestShare
'              Opens a window view of the specified share.
'   Example 4: Explorer /root,\\TestSvr\TestShare,select,TestProg.exe
'              Opens a window view of the specified share with TestProg selected.


'

'-----------------------------------------------------
' Create a windows host scripting object
Set WshShell = WScript.CreateObject("WScript.Shell")

' Create an object to extract environment variables
Set objEnv = WshShell.Environment("Process")

' If any members are selected, then the following variables apply:
' MKSSI_MEMBERxx_SANDBOX=sandbox-side-project/subproject-path at sandbox
member =  objEnv("MKSSI_MEMBER1_SANDBOX")

' If any subprojects are selected, then the following variables apply:
'MKSSI_SUBPROJECTyy_PROJECT=sandbox-side-project/subproject-path
project = objEnv("MKSSI_SUBPROJECT1_SANDBOX")

file = objEnv("MKSSI_FILE")

cmd = ""

if member <> "" then
cmd = fexplore + Left(member,InstrRev(member,"\"))
elseif project <> "" then
cmd = fexplore + Left(project,InstrRev(project,"\"))
elseif file <> "" then
cmd = fexplore + Left(file,InstrRev(file,"\"))
else
cmd = ""
end if

if cmd <> "" then
WSHShell.Run(cmd)
end if
