Attribute VB_Name = "modAddIn"
Option Explicit

Declare Function WritePrivateProfileString& Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)

Public Const ADDIN_NAME = "add_logerr"

'To define the caption that appears in the Add-In Manager window go to the
'Object Browser (Ctrl+R), select clsConnect, right click, select "Properties ..."
'VB's "Member Options" dialog should appear.  In the "Description" text box
'enter the caption you want to appear in the Add-Manager window.

Public theConnection As clsConnect
 
Public Sub InitializeAddIn()
  On Error Resume Next
  AddToINI
  'ɾ���ͱ༭ c:\windows\vbaddin.ini
End Sub

'This procedure must be executed before VB's Add-In Manager will
'recognize the add-in as available.  Normally the procedure should be
'executed by the setup program.  During program development you will need
'to run it once in the immediate window to make the add-in available in
'your local environment.
Sub AddToINI()
  Dim ErrCode As Long
  'д�� c:\windows\vbaddin.ini
  ErrCode = WritePrivateProfileString("Add-Ins32", ADDIN_NAME & ".clsConnect", "0", "vbaddin.ini")
End Sub

Public Sub CreateError(ByVal sErrorDescription As String)
    MsgBox "Error: " & "[" & ADDIN_NAME & " Add-In] " & sErrorDescription
    'Debug.Print "[" & ADDIN_NAME & " Add-In] " & sErrorDescription
End Sub
 
Public Sub DisconnectAddIn()
'
End Sub
