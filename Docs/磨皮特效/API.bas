Attribute VB_Name = "API"



Option Explicit


Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type



Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long


Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long


Public Const ERRORINDEX                 As Long = -1
Public SystemFrequency                      As Currency

Public Function GetCurrentTime() As Currency
    On Error Resume Next
    If SystemFrequency = 0 Then '未初始化
        If QueryPerformanceFrequency(SystemFrequency) = 0 Then
            SystemFrequency = ERRORINDEX '无高精度计数器
        End If
    End If

    If SystemFrequency <> ERRORINDEX Then
        Dim CurCount As Currency
        QueryPerformanceCounter CurCount
        GetCurrentTime = CurCount * 1000@ / SystemFrequency
    Else
        GetCurrentTime = GetTickCount()
    End If
End Function



'*************************************************************************
'**    作    者 ：    laviewpbt
'**    函 数 名 ：    ShowOpen
'**    输    入 ：    hwnd  (long)    -   窗口句柄
'**                   Filter(String)  -   过滤选项
'**                   Title (String)  -   对话框标题
'**                   IniPath         -   首选路径
'**    输    出 ：    返回一路径字符串
'**    功能描述 ：    通用打开对话框
'**    日    期 ：    2005-10-21 11:27.25
'**    修 改 人 ：
'**    日    期 ：
'**    版    本 ：    Version 1.1.1
'*************************************************************************

Public Function ShowOpen(Hwnd As Long, Filter As String, Title As String, Optional IniPath = "") As String
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Filter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = Title
    OFName.lpstrInitialDir = IniPath
    OFName.Flags = 0
    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function

'*************************************************************************
'**    作    者 ：    laviewpbt
'**    函 数 名 ：    ShowSave
'**    输    入 ：    hwnd  (long)    -   窗口句柄
'**                   Filter(String)  -   过滤选项
'**                   Title (String)  -   对话框标题
'**                   IniPath         -   首选路径
'**    输    出 ：    返回一路径字符串
'**    功能描述 ：    通用保存对话框
'**    日    期 ：    2005-10-21 11:29.42
'**    修 改 人 ：
'**    日    期 ：
'**    版    本 ：    Version 1.1.1
'*************************************************************************

Public Function ShowSave(Hwnd As Long, Filter As String, Title As String, FilterIndex As Integer, Optional IniPath = "") As String
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = Filter
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = Title
    OFName.lpstrInitialDir = IniPath
    OFName.Flags = &H2                '有相同的文件是提示
    If GetSaveFileName(OFName) Then
        ShowSave = Trim$(OFName.lpstrFile)
        ShowSave = Left(ShowSave, Len(ShowSave) - 1)
        FilterIndex = OFName.nFilterIndex
    Else
        ShowSave = ""
    End If
End Function



Public Sub GetBestFitInfoEx(ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByRef FitX As Long, ByRef FitY As Long, ByRef FitWidth As Long, ByRef FitHeight As Long)
    Dim WidthRatio As Double
    Dim HeightRatio As Double
    If SrcWidth > DestWidth Or SrcHeight > DestHeight Then
        WidthRatio = DestWidth / SrcWidth
        HeightRatio = DestHeight / SrcHeight
        If WidthRatio < HeightRatio Then
            FitWidth = DestWidth
            FitHeight = CInt(SrcHeight * WidthRatio + 0.5)
        Else
            FitHeight = DestHeight
            FitWidth = CInt(SrcWidth * HeightRatio + 0.5)
        End If
    Else
        FitWidth = SrcWidth
        FitHeight = SrcHeight
    End If
    FitX = (DestWidth - FitWidth) \ 2
    FitY = (DestHeight - FitHeight) \ 2
End Sub

