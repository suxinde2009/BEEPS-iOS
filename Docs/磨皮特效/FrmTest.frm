VERSION 5.00
Begin VB.Form FrmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "磨皮测试"
   ClientHeight    =   10260
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   18360
   Icon            =   "FrmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   684
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1224
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox CmbWhitenMethod 
      Height          =   300
      Left            =   10770
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   9840
      Width           =   2115
   End
   Begin VB.HScrollBar WhitenLevel 
      Height          =   285
      Left            =   5820
      Max             =   10
      TabIndex        =   12
      Top             =   9810
      Value           =   2
      Width           =   3015
   End
   Begin VB.HScrollBar Amount 
      Height          =   285
      Left            =   5820
      Max             =   10
      TabIndex        =   10
      Top             =   9390
      Value           =   8
      Width           =   3015
   End
   Begin VB.ComboBox CmbRetouchMethod 
      Height          =   300
      Left            =   10770
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   9390
      Width           =   2115
   End
   Begin VB.CheckBox ChkPreview 
      Caption         =   " 完整预览"
      Height          =   345
      Left            =   3210
      TabIndex        =   7
      Top             =   9360
      Width           =   1425
   End
   Begin VB.PictureBox PicDstFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   9270
      ScaleHeight     =   600
      ScaleMode       =   0  'User
      ScaleWidth      =   600
      TabIndex        =   5
      Top             =   120
      Width           =   9000
      Begin VB.PictureBox PicDst 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6840
         Left            =   1560
         MousePointer    =   99  'Custom
         Picture         =   "FrmTest.frx":1CFA
         ScaleHeight     =   456
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   579
         TabIndex        =   6
         Top             =   1500
         Width           =   8685
      End
   End
   Begin VB.PictureBox PicSrcFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   120
      ScaleHeight     =   600
      ScaleMode       =   0  'User
      ScaleWidth      =   600
      TabIndex        =   3
      Top             =   120
      Width           =   9000
      Begin VB.PictureBox PicSrc 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6840
         Left            =   1560
         MousePointer    =   99  'Custom
         Picture         =   "FrmTest.frx":21503
         ScaleHeight     =   456
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   579
         TabIndex        =   4
         Top             =   1500
         Width           =   8685
      End
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存图像"
      Height          =   375
      Left            =   1620
      TabIndex        =   1
      Top             =   9330
      Width           =   1275
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "选择图像"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   9330
      Width           =   1275
   End
   Begin VB.Label LblB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "基于算法："
      Height          =   180
      Index           =   5
      Left            =   9750
      TabIndex        =   16
      Top             =   9900
      Width           =   900
   End
   Begin VB.Label LblB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "美白程度："
      Height          =   180
      Index           =   4
      Left            =   4830
      TabIndex        =   14
      Top             =   9870
      Width           =   900
   End
   Begin VB.Label LblB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   180
      Index           =   3
      Left            =   9030
      TabIndex        =   13
      Top             =   9870
      Width           =   90
   End
   Begin VB.Label LblB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   180
      Index           =   2
      Left            =   9000
      TabIndex        =   11
      Top             =   9450
      Width           =   90
   End
   Begin VB.Label LblB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "基于算法："
      Height          =   180
      Index           =   0
      Left            =   9720
      TabIndex        =   9
      Top             =   9450
      Width           =   900
   End
   Begin VB.Label LblB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "磨皮程度："
      Height          =   180
      Index           =   1
      Left            =   4830
      TabIndex        =   2
      Top             =   9450
      Width           =   900
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************************
'**    开发日期 ：  2009-7-21
'**    作    者 ：  laviewpbt
'**    联系方式：   33184777
'**    修改日期 ：   2009-7-21
'**    版    本 ：  Version 1.3.1
'**    转载请不要删除以上信息
'****************************************************************************************



Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)


Private Declare Sub DermabrasionFilter Lib "ImageProcessing.dll" (ByVal Src As Long, ByVal dest As Long, ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal RetouchMethod As Long, ByVal WhiteningMethod As Long, ByVal RetouchLevel As Long, ByVal WhiteningLevel As Long)


Private Img         As New Cimage
Attribute Img.VB_VarHelpID = -1
Private ImgC        As New Cimage
Private FitX        As Long
Private FitY        As Long
Private FitWidth    As Long
Private FitHeight   As Long

Private OriginalX   As Long
Private OriginalY   As Long
Private NewX        As Long
Private NewY        As Long

Private Init        As Boolean
Private Busy        As Boolean

Private Sub Amount_Change()
    UpdateImage
    RefreshDst
    LblB(2).Caption = Amount.Value
End Sub

Private Sub Amount_Scroll()
    Amount_Change
End Sub

Private Sub ChkPreview_Click()
    If Img.hdc <> 0 Then
        PicSrc.Cls
        PicDst.Cls
        If ChkPreview.Value = vbChecked Then
            PicSrc.Move 0, 0, 600, 600
            PicDst.Move 0, 0, 600, 600
            API.GetBestFitInfoEx Img.Width, Img.Height, 600, 600, FitX, FitY, FitWidth, FitHeight
            Img.Render PicSrc.hdc, FitX, FitY, FitWidth, FitHeight, 0, 0, Img.Width, Img.Height
            ImgC.Render PicDst.hdc, FitX, FitY, FitWidth, FitHeight, 0, 0, ImgC.Width, ImgC.Height
        Else
            PicSrc.Width = Img.Width: PicSrc.Height = Img.Height
            PicSrc.Move (PicSrcFrame.ScaleWidth - PicSrc.ScaleWidth) / 2, (PicSrcFrame.ScaleHeight - PicSrc.ScaleHeight) / 2
            PicDst.Width = ImgC.Width: PicDst.Height = ImgC.Height
            PicDst.Move (PicDstFrame.ScaleWidth - PicDst.ScaleWidth) / 2, (PicDstFrame.ScaleHeight - PicDst.ScaleHeight) / 2
            Img.Render PicSrc.hdc, 0, 0, Img.Width, Img.Height, 0, 0, Img.Width, Img.Height
            ImgC.Render PicDst.hdc, 0, 0, ImgC.Width, ImgC.Height, 0, 0, ImgC.Width, ImgC.Height
        End If
        PicSrc.Refresh
        PicDst.Refresh
    End If
End Sub


Private Sub CmbRetouchMethod_Change()
    If Init = True Then
        UpdateImage
        RefreshDst
    End If
End Sub

Private Sub CmbRetouchMethod_Click()
    CmbRetouchMethod_Change
End Sub

Private Sub CmbWhitenMethod_Change()
    If Init = True Then
        UpdateImage
        RefreshDst
    End If
End Sub

Private Sub CmbWhitenMethod_Click()
    CmbWhitenMethod_Change
End Sub

Private Sub Form_Load()


    CmbRetouchMethod.AddItem "基于BEEP算法"
    CmbRetouchMethod.AddItem "基于导向滤波"
    CmbRetouchMethod.AddItem "基于快速双边滤波"
    CmbRetouchMethod.AddItem "基于Domain  Transform 滤波"
    CmbRetouchMethod.AddItem "基于局部均方差滤波"
     
    CmbRetouchMethod.ListIndex = 4
    
    CmbWhitenMethod.AddItem "基于log曲线"
    CmbWhitenMethod.AddItem "基于色彩平衡"
    CmbWhitenMethod.ListIndex = 0
    
    Init = True
    
    Img.LoadPictureFromStdPicture PicSrc.Image
    ImgC.LoadPictureFromStdPicture PicDst.Image
    PicSrc.Picture = LoadPicture
    PicDst.Picture = LoadPicture
   
    UpdateImage
    ChkPreview_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Img.DisposeResource
    ImgC.DisposeResource
    Set Img = Nothing
    Set ImgC = Nothing
End Sub



Private Sub CmdOpen_Click()
    On Error GoTo ErrHandle:
    Dim FileName            As String
    FileName = API.ShowOpen(Me.Hwnd, "All Suppported Images |*.bmp;*.jpg;|BMP Images|*.bmp|JPG Images|*.jpg", "选择图像", App.Path)
    If FileName <> "" Then
        If Img.handle <> 0 Then
            Img.DisposeResource
            ImgC.DisposeResource
        End If
        Img.LoadPictureFromFile (FileName)
        Set ImgC = Img.Clone

        UpdateImage
        ChkPreview_Click
    End If
    Exit Sub
ErrHandle:
    MsgBox "不支持的图像格式。", vbCritical
End Sub

Private Sub CmdSave_Click()
    Dim FileName            As String
    FileName = API.ShowSave(Me.Hwnd, "Png Images|*.Png", "保存图像", 0)
    If FileName <> "" Then
        ImgC.SavePictureToFile (FileName + ".png")
    End If
End Sub


Private Sub UpdateImage()
    If Img.handle <> 0 And Busy = False Then
        Busy = True
        Dim TimeElpase As Currency
        TimeElpase = API.GetCurrentTime
        DermabrasionFilter Img.Pointer, ImgC.Pointer, Img.Width, Img.Height, Img.Stride, CmbRetouchMethod.ListIndex, CmbWhitenMethod.ListIndex, Amount.Value, WhitenLevel.Value
        Me.Caption = "美化算法用时 " & API.GetCurrentTime - TimeElpase & " ms"
        Busy = False
    End If
End Sub

Private Sub RefreshDst()
    If ChkPreview.Value = vbChecked Then
        ImgC.Render PicDst.hdc, FitX, FitY, FitWidth, FitHeight, 0, 0, ImgC.Width, ImgC.Height
    Else
        ImgC.Render PicDst.hdc, 0, 0, ImgC.Width, ImgC.Height, 0, 0, ImgC.Width, ImgC.Height
    End If
    PicDst.Refresh
End Sub



Private Sub PicDst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicSrc_MouseDown Button, Shift, X, Y
End Sub

Private Sub PicDst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PicSrc_MouseMove Button, Shift, X, Y
End Sub

Private Sub PicSrc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        OriginalX = X                   '记录起点坐标
        OriginalY = Y
    End If
End Sub

Private Sub PicSrc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And ChkPreview.Value <> vbChecked Then
        NewX = PicSrc.Left + X - OriginalX  '新的X坐标
        NewY = PicSrc.Top + Y - OriginalY
        If PicSrc.Left <= 0 Then         '如果>0,则没有移动的必要
            If NewX > 0 Then             '越界处理
                PicSrc.Left = 0
                PicDst.Left = 0
            ElseIf NewX >= PicSrcFrame.ScaleWidth - PicSrc.ScaleWidth Then
                PicSrc.Left = NewX
                PicDst.Left = NewX
            End If
        End If
        If PicSrc.Top <= 0 Then
            If NewY > 0 Then
                PicSrc.Top = 0
                PicDst.Top = 0
            ElseIf NewY >= PicSrcFrame.ScaleHeight - PicSrc.ScaleHeight Then
                PicSrc.Top = NewY
                PicDst.Top = NewY
            End If
        End If
    End If
End Sub

Private Sub WhitenLevel_Change()
    UpdateImage
    RefreshDst
    LblB(3).Caption = WhitenLevel.Value
End Sub

Private Sub WhitenLevel_Scroll()
    WhitenLevel_Change
End Sub
