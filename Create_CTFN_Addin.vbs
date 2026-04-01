' ============================================================
' Double-click this file to create CTFN.xlam automatically.
' It opens Excel, imports the VBA code, saves the add-in,
' and closes. Takes about 5 seconds.
' ============================================================

Dim fso, scriptDir, basFile, xlamFile, addinsDir
Dim xlApp, wb, vbProj

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
basFile = fso.BuildPath(scriptDir, "CTFN.bas")

' Check that CTFN.bas exists next to this script
If Not fso.FileExists(basFile) Then
    MsgBox "Cannot find CTFN.bas in:" & vbCrLf & vbCrLf & scriptDir & vbCrLf & vbCrLf & _
           "Make sure CTFN.bas is in the same folder as this script.", _
           vbExclamation, "CTFN Setup"
    WScript.Quit
End If

' Output path: save both in this folder and in the Excel Add-ins folder
xlamFile = fso.BuildPath(scriptDir, "CTFN.xlam")

' Also figure out the user's Add-ins folder for auto-install
Dim shell
Set shell = CreateObject("WScript.Shell")
addinsDir = shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns"
If Not fso.FolderExists(addinsDir) Then fso.CreateFolder(addinsDir)

On Error Resume Next

' Open Excel
Set xlApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "Could not start Excel. Make sure Microsoft Excel is installed.", _
           vbExclamation, "CTFN Setup"
    WScript.Quit
End If
On Error GoTo 0

xlApp.Visible = False
xlApp.DisplayAlerts = False

' Check that VBA project access is enabled
On Error Resume Next
Set wb = xlApp.Workbooks.Add
Set vbProj = wb.VBProject
If Err.Number <> 0 Then
    wb.Close False
    xlApp.Quit
    MsgBox "Excel is blocking access to VBA projects." & vbCrLf & vbCrLf & _
           "To fix this:" & vbCrLf & _
           "1. Open Excel" & vbCrLf & _
           "2. File > Options > Trust Center > Trust Center Settings" & vbCrLf & _
           "3. Macro Settings > check 'Trust access to the VBA project object model'" & vbCrLf & _
           "4. Click OK, then double-click this script again.", _
           vbExclamation, "CTFN Setup"
    WScript.Quit
End If
On Error GoTo 0

' Import the .bas file
vbProj.VBComponents.Import basFile

' Save as .xlam (file format 55 = xlOpenXMLAddIn)
wb.SaveAs xlamFile, 55
wb.Close False

' Also copy to the Add-ins folder
Dim addinsPath
addinsPath = fso.BuildPath(addinsDir, "CTFN.xlam")
fso.CopyFile xlamFile, addinsPath, True

xlApp.Quit
Set xlApp = Nothing

MsgBox "CTFN.xlam has been created and installed!" & vbCrLf & vbCrLf & _
       "Saved to:" & vbCrLf & xlamFile & vbCrLf & vbCrLf & _
       "Also installed to:" & vbCrLf & addinsPath & vbCrLf & vbCrLf & _
       "Next step:" & vbCrLf & _
       "1. Open Excel" & vbCrLf & _
       "2. File > Options > Add-ins > Go" & vbCrLf & _
       "3. Check 'CTFN' and click OK" & vbCrLf & _
       "4. Type =CTFN(""headline"",""ROKU"") in any cell", _
       vbInformation, "CTFN Setup Complete"
