VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Dim mIDE                        As vbide.VBE
Dim mcolMenus                   As Collection               'store all menus for easy deletion
Public WithEvents MenuHandler   As CommandBarEvents         'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Dim menuControlPanel As Office.CommandBarControl  'command bar object
Const m_err_add = "e自动添加错误处理 (&E) "

Private Sub Class_Initialize()
    Set mIDE = modAddIn.theConnection.theIDE
End Sub

Friend Sub SetupMenus()
On Error GoTo LocalError
    Dim cbFileMenu As Office.CommandBarControl, cbMakeMenu As Office.CommandBarControl
    Dim ctFileCommandBars As Long, idxCommandBar As Long
    Set mcolMenus = New Collection

    ' 添加代码右键菜单项  m_err_add
    Set cbContextMenu = mIDE.CommandBars(15).Controls("定义(&D)")
    If Not cbContextMenu Is Nothing Then
        contextMenuIndex = cbContextMenu.Index + 3
        Dim contextMenuItem As Office.CommandBarControl
        Set contextMenuItem = mIDE.CommandBars(15).Controls.Add(, , , contextMenuIndex)
        contextMenuItem.Caption = m_err_add
        mcolMenus.Add contextMenuItem
        Set Me.MenuHandler = mIDE.Events.CommandBarEvents(contextMenuItem)
    End If

    Exit Sub

LocalError:
    CreateError "Unexpected error while attempting to add " & ADDIN_NAME & " menu to VB IDE: " & Err.Description & " " & Erl
End Sub

Friend Sub RemoveMenus()
    For Each CurrentMenu In mcolMenus
        CurrentMenu.Delete
    Next
End Sub

'Event occurs when the menu is clicked in the IDE
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Select Case CommandBarControl.Caption
        Case m_err_add
          'MsgBox "自动 添加 错误处理"
          添加错误处理
        Case Else
            'not a handled menu
    End Select
End Sub

Function 读取当前编辑的函数_内容() As String
    On Error Resume Next
    Dim codeModule As codeModule
    Dim procName As String
    Dim procKind As vbext_ProcKind
    Dim startLine As Long, startCol As Long, endLine As Long, endCol As Long
    Dim procStartLine As Long, procBodyLineCount As Long
    
    Set codeModule = mIDE.ActiveCodePane.codeModule
    If codeModule Is Nothing Then Exit Function
    
    ' 获取当前光标所在行
    mIDE.ActiveCodePane.GetSelection startLine, startCol, endLine, endCol
    
    ' 获取当前函数或子程序的名称
    procName = codeModule.ProcOfLine(startLine, procKind)
    If procName = "" Then Exit Function
    
    ' 获取函数或子程序的起始行和行数
    procStartLine = codeModule.procStartLine(procName, procKind)
    procBodyLineCount = codeModule.ProcCountLines(procName, procKind)
    
    ' 读取函数或子程序的内容
    读取当前编辑的函数_内容 = codeModule.Lines(procStartLine, procBodyLineCount)
    
End Function

Function 读取当前编辑的函数_name() As String
  On Error Resume Next
    Dim codeModule As codeModule
    Dim currentLine As Long
    Dim procName As String
    Dim procKind As vbext_ProcKind
     Dim startLine As Long, startCol As Long, endLine As Long, endCol As Long
    
    Set codeModule = mIDE.ActiveCodePane.codeModule
    If codeModule Is Nothing Then Exit Function
    
    'currentLine = mIDE.ActiveCodePane.TopLine
    'currentLine = mIDE.ActiveCodePane.CodePane.currentLine
    mIDE.ActiveCodePane.GetSelection startLine, startCol, endLine, endCol
    currentLine = startLine
    
    Dim i As Long
    procName = codeModule.ProcOfLine(currentLine, procKind)
    For i = 0 To 9
      If Left(procName, 1) = "'" Then
        currentLine = currentLine + 1
        procName = codeModule.ProcOfLine(currentLine, procKind)
      Else
        Exit For
      End If
    Next

    If procName <> "" Then
        读取当前编辑的函数_name = procName
    End If
End Function

Function 读取当前编辑的代码() As String
  On Error Resume Next
  Dim codeModule As codeModule
  Set codeModule = mIDE.ActiveCodePane.codeModule
  If Not codeModule Is Nothing Then
      读取当前编辑的代码 = codeModule.Lines(1, codeModule.CountOfLines)
  End If
End Function

Sub 在当前编辑的代码中插入错误处理(name As String)
    On Error Resume Next
    Dim codeModule As codeModule
    Dim procName As String
    Dim procKind As vbext_ProcKind
    Dim startLine As Long, startCol As Long, endLine As Long, endCol As Long
    Dim actualStartLine As Long
    Dim procStartLine As Long, procBodyLineCount As Long
    Dim procContent As String
    Dim exitStatement As String, s As String
    
    ' 获取当前代码模块
    Set codeModule = mIDE.ActiveCodePane.codeModule
    If codeModule Is Nothing Then Exit Sub
    
    ' 获取当前光标所在位置
    mIDE.ActiveCodePane.GetSelection startLine, startCol, endLine, endCol
    
    ' 获取当前函数/子程序名称和类型
    procName = codeModule.ProcOfLine(startLine, procKind)
    If procName = "" Then Exit Sub
    
    ' 获取函数起始行和行数
    procStartLine = codeModule.procStartLine(procName, procKind)
    procBodyLineCount = codeModule.ProcCountLines(procName, procKind)
    
    ' 跳过空行，找到实际的起始行
    actualStartLine = procStartLine
    For i = procStartLine To procStartLine + 7  ' 最多检查n行
        If Trim(codeModule.Lines(i, 1)) <> "" Then
            actualStartLine = i
            Exit For
        End If
    Next i
    
    ' 确定退出语句类型
    If procKind = 0 Then  ' 0 表示 Sub
        exitStatement = "Exit Sub"
    Else
        exitStatement = "Exit Function"
    End If
    
    ' 在函数第一行插入错误处理头
    codeModule.InsertLines actualStartLine + 1, vbTab & "On Error Goto EH"
    
    ' 在函数末尾插入错误处理尾
    s = vbCrLf & vbTab & exitStatement & vbCrLf & "EH:" & vbCrLf & vbTab & "logerr err.Description"
    s = s & " & " & """ @ " & Left(name, 1) & "_" & Mid(name, 2) & """"
    codeModule.InsertLines procStartLine + procBodyLineCount, s
End Sub

Sub 添加错误处理()
  Static name As String
'  s = 读取当前编辑的代码
'  s = 读取当前编辑的函数or_sub
  name = 读取当前编辑的函数_name
  s = 读取当前编辑的函数_内容
  'MsgBox "当前函数: " & s
  'MsgBox "内容: " & 读取当前编辑的函数_内容
  If InStr(1, Left(s, 70), "On Err") > 0 Then
    MsgBox "已存在: " & s, vbInformation
    Exit Sub
  End If
  在当前编辑的代码中插入错误处理 name
End Sub


