VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cimage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit


Private Type BITMAPFILEHEADER
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits  As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth        As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression  As Long
    biSizeImage    As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed      As Long
    biClrImportant  As Long
End Type

Private Type Bitmap
    bmType      As Long
    bmWidth      As Long
    bmHeight    As Long
    bmWidthBytes As Long
    bmPlanes    As Integer
    bmBitsPixel  As Integer
    BmBits      As Long
End Type

Private Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    reserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type


Private Const BI_bitfields = 3&                '带掩码的
Private Const BI_RGB = 0                        '正常
Private Const DIB_RGB_COLORS = 0                '真彩色
Private Const OBJ_BITMAP = 7                    '位图对象
Private Const SRCCOPY = &HCC0020                '直接拷贝
Private Const IMAGE_BITMAP = 0                  'LoadImage函数的载入类型，位图
Private Const LR_LOADFROMFILE = &H10            '从文件载入
Private Const LR_CREATEDIBSECTION = &H2000      '如果指定了IMAGE_BITMAP，就返回DIBSection的句柄，而不是位图的句柄
Private Const STRETCH_ANDSCANS = 1              '默认设置。剔除的线段与剩下的线段进行AND运算。这个模式通常应用于采用了白色背景的单色位图
Private Const STRETCH_ORSCANS = 2              '剔除的线段被简单的清除。这个模式通常用于彩色位图
Private Const STRETCH_DELETESCANS = 3          '剔除的线段与剩下的线段进行OR运算。这个模式通常应用于采用了白色背景的单色位图
Private Const STRETCH_HALFTONE = 4              '目标位图上的像素块被设为源位图上大致近似的块。这个模式要明显慢于其他模式

'******************************************** 用于图像方面的相关API ********************************************

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, lpBits As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

'******************************************** 用于系统输出的相关API ********************************************

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

'******************************************** 用于内存处理的相关API ********************************************

Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal ByteLength As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)




'******************************************** 公共常用的API函数 ********************************************

Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long

Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long





Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type RECTF
    nLeft               As Single
    nTop                As Single
    nWidth              As Single
    nHeight             As Single
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Private Const UnitPixel As Long = &H2&

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, hImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, Graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipGetImageBounds Lib "gdiplus.dll" (ByVal nImage As Long, SrcRect As RECTF, srcUnit As Long) As Long

Private Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6
    EncoderParameterValueTypeUndefined = 7
    EncoderParameterValueTypeRationalRange = 8
End Enum

Private Type EncoderParameter
    GUID(0 To 3)        As Long
    NumberOfValues      As Long
    Type                As EncoderParameterValueType
    Value               As Long
End Type

Private Type EncoderParameters
    Count               As Long
    Parameter           As EncoderParameter
End Type

Private Type ImageCodecInfo
    ClassID(0 To 3)     As Long
    FormatID(0 To 3)    As Long
    CodecName           As Long
    DllName             As Long
    FormatDescription   As Long
    FilenameExtension   As Long
    MimeType            As Long
    Flags               As Long
    Version             As Long
    SigCount            As Long
    SigSize             As Long
    SigPattern          As Long
    SigMask             As Long
End Type

Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFilename As Long, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Function GdipGetImageEncodersSize Lib "gdiplus" (numEncoders As Long, Size As Long) As Long
Private Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal numEncoders As Long, ByVal Size As Long, Encoders As Any) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, Bitmap As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal psString As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long



Private Type SAFEARRAYBOUND
    Elements            As Long
    lLbound             As Long
End Type


Private Type SAFEARRAY
    Dimension           As Integer
    Features            As Integer
    Element             As Long
    Locks               As Long
    Pointer             As Long
    Bounds              As SAFEARRAYBOUND
End Type


Private m_Left              As Long                     '层在画布中的左侧坐标
Private m_Top               As Long                     '层在画布中的顶部坐标
Private m_Width             As Long                     '层的宽度
Private m_Height            As Long                     '层的高度
Private m_Stride            As Long                     '层数据每个扫描行的大小
Private m_Resolution        As Double
Private m_Hdc               As Long                     '层的内存DC
Private m_Pointer           As Long                     '层数据在内存的首地址


Private m_Handle            As Long                     'DIBSection的句柄
Private m_OldHandle         As Long                     '原始设备环境的句柄
Private m_Image             As Long                     '用于输出的GDI+的Image对象


Private Sa                  As SAFEARRAY
Private ImageData()         As Byte



Public Property Get Width() As Long
    Width = m_Width
End Property

Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Get Stride() As Long
    Stride = m_Stride
End Property


Public Property Get Pointer() As Long
    Pointer = m_Pointer
End Property


Public Property Get hdc() As Long
    hdc = m_Hdc
End Property

Public Property Get handle() As Long
    handle = m_Handle
End Property

Private Sub Class_Terminate()
    DisposeResource
End Sub


Public Function CreateNewImage(ByVal Width As Long, ByVal Height As Long) As Boolean

    Dim ScreenDC            As Long, BmpInfo            As BITMAPINFOHEADER
    Dim X                   As Long, Y                  As Long

    Dim Red                 As Long, Green              As Long
    Dim Blue                As Long, Alpha              As Long
    Dim Gray                As Long

    If Width <= 0 Or Height <= 0 Then Exit Function

    DisposeResource                        '删除原始的内存资源

    With BmpInfo
        .biBitCount = 24
        .biHeight = -Height                '为了和GDI对象的坐标系统（起点坐标在左上角），建立一个倒序的DIB
        .biWidth = Width
        .biPlanes = 1
        .biSize = 40
        m_Stride = ((Width * 3 + 3) And &HFFFFFFFC)
         .biSizeImage = m_Stride * Height
    End With
    ScreenDC = GetDC(0)                                                               '得到屏幕DC
    m_Hdc = CreateCompatibleDC(ScreenDC)
    ReleaseDC 0, ScreenDC                                                             '释放屏幕DC
    m_Handle = CreateDIBSection(m_Hdc, BmpInfo, DIB_RGB_COLORS, m_Pointer, 0, 0)
    If m_Handle <> 0 Then                                                              '希望系统能够让我们成功创建DIB吧
        m_OldHandle = SelectObject(m_Hdc, m_Handle)
        m_Width = Width: m_Height = Height
        With Sa
            .Element = 1
            .Dimension = 1
            .Bounds.Elements = m_Stride * m_Height
            .Pointer = m_Pointer
        End With
        CopyMemory ByVal VarPtrArray(ImageData()), VarPtr(Sa), 4
        CreateNewImage = True
    End If
End Function

Public Function OutPut(ByVal OutDC As Long, Optional ByVal X As Long = 0, Optional ByVal Y As Long = 0, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional ByVal SrcX As Long = 0, Optional ByVal SrcY As Long = 0, Optional ByVal dwRop As RasterOpConstants = vbSrcCopy) As Boolean
    If Width = 0 Then Width = m_Width
    If Height = 0 Then Height = m_Height
    OutPut = BitBlt(OutDC, X, Y, Width, Height, m_Hdc, SrcX, SrcY, dwRop)
End Function

Public Function Render(ByVal DestDC As Long, _
                       Optional ByVal DestX As Long, _
                       Optional ByVal DestY As Long, _
                       Optional ByVal DestWidth As Long, _
                       Optional ByVal DestHeight As Long, _
                       Optional ByVal SrcX As Long, _
                       Optional ByVal SrcY As Long, _
                       Optional ByVal SrcWidth As Long, _
                       Optional ByVal SrcHeight As Long) As Boolean
    If DestWidth = 0 Then DestWidth = m_Width
    If DestHeight = 0 Then DestHeight = m_Height
    If SrcX < 0 Then SrcX = 0  ' 源X,Y不能为负，但目的X,Y可以
    If SrcY < 0 Then SrcY = 0
    If SrcWidth = 0 Then
        SrcWidth = m_Width
    ElseIf SrcWidth < 0 Then
        DestWidth = -DestWidth
        SrcWidth = -SrcWidth
    End If
    If SrcHeight = 0 Then
        SrcHeight = m_Height
    ElseIf SrcHeight < 0 Then
        DestHeight = -DestHeight
        SrcHeight = -SrcHeight
    End If
    SetStretchBltMode DestDC, STRETCH_DELETESCANS
    StretchBlt DestDC, DestX, DestY, DestWidth, DestHeight, m_Hdc, SrcX, SrcY, SrcWidth, SrcHeight, vbSrcCopy
End Function


Public Function Clone() As Cimage
    Dim C As New Cimage
    C.CreateNewImage m_Width, m_Height
    CopyMemory ByVal C.Pointer, ByVal m_Pointer, Height * Stride
    Set Clone = C
End Function



Public Sub DisposeResource()
    If m_Hdc <> 0 Then
        SelectObject m_Hdc, m_OldHandle
        DeleteDC m_Hdc
        DeleteObject m_Handle
    
        CopyMemory ByVal VarPtrArray(ImageData()), 0&, 4

        m_Width = 0: m_Height = 0                                                            '重置其他的图像相关属性
        m_Handle = 0: m_OldHandle = 0
        m_Pointer = 0: m_Hdc = 0
    End If
End Sub


Public Function LoadPictureFromFile(ByVal FileName As String) As Boolean
    Dim Image               As Long, Graphics           As Long
    Dim Dimensions          As RECTF, Rct               As RECT
    Dim Resolution          As Single, Token            As Long
    Dim GdipInput           As GdiplusStartupInput
    GdipInput.GdiplusVersion = 1
    GdiplusStartup Token, GdipInput
    GdipLoadImageFromFile StrPtr(FileName), Image
    If Image <> 0 Then
        GdipGetImageBounds Image, Dimensions, UnitPixel
        If CreateNewImage(CLng(Dimensions.nWidth), CLng(Dimensions.nHeight)) = True Then
            Rct.Right = Dimensions.nWidth
            Rct.Bottom = Dimensions.nHeight
            GdipCreateFromHDC m_Hdc, Graphics
            GdipDrawImageRectRectI Graphics, Image, 0, 0, m_Width, m_Height, 0, 0, m_Width, m_Height, UnitPixel, 0, 0, 0
            GdipDeleteGraphics Graphics
        End If
    End If
    GdipDisposeImage Image
    GdiplusShutdown Token
End Function


Public Function LoadPictureFromStdPicture(StdPic As StdPicture) As Boolean
    If Not StdPic Is Nothing Then
        Dim Width       As Long '
        Dim Height      As Long
        Width = ConvertHimetrix2Pixels(StdPic.Width, True)      'StdPicture宽度中的单位是Himetrics
        Height = ConvertHimetrix2Pixels(StdPic.Height, False)
        If CreateNewImage(Width, Height) = True Then            '创建一个空白的Dib
            StdPic.Render m_Hdc + 0, 0, 0, Width + 0, Height + 0, 0, StdPic.Height, StdPic.Width, -StdPic.Height, ByVal 0  '类似于BMP的逆序存储，所以用-StdPic.Height
        End If
    End If
End Function


Public Function SavePictureToFile(FileName As String) As Boolean
    Dim uEncCLSID(0 To 3)   As Long
    Dim Token               As Long
    Dim GdipInput           As GdiplusStartupInput
    GdipInput.GdiplusVersion = 1
    GdiplusStartup Token, GdipInput
    
    If m_Hdc <> 0 Then
        If Not GetEncoderClsID("Image/png", uEncCLSID) = -1 Then
            Dim Image    As Long
            GdipCreateBitmapFromScan0 m_Width, m_Height, m_Stride, &H21808, ByVal m_Pointer, Image
            GdipSaveImageToFile Image, StrPtr(FileName), uEncCLSID(0), ByVal 0
            GdipDisposeImage Image
        End If
    End If
    GdiplusShutdown Token
End Function


Private Function GetEncoderClsID(strMimeType As String, ClassID() As Long) As Long
    Dim Num         As Long
    Dim Size        As Long
    Dim I           As Long
    Dim Info()      As ImageCodecInfo
    Dim Buffer()    As Byte
    GetEncoderClsID = -1
    GdipGetImageEncodersSize Num, Size               '得到解码器数组的大小
    If Size <> 0 Then
       ReDim Info(1 To Num) As ImageCodecInfo       '给数组动态分配内存
       ReDim Buffer(1 To Size) As Byte
       GdipGetImageEncoders Num, Size, Buffer(1)            '得到数组和字符数据
       CopyMemory Info(1), Buffer(1), (Len(Info(1)) * Num)     '复制类头
       For I = 1 To Num             '循环检测所有解码
           If (StrComp(PtrToStrW(Info(I).MimeType), strMimeType, vbTextCompare) = 0) Then         '必须把指针转换成可用的字符
               CopyMemory ClassID(0), Info(I).ClassID(0), 16  '保存类的ID
               GetEncoderClsID = I      '返回成功的索引值
               Exit For
           End If
       Next
    End If
End Function

Private Function PtrToStrW(ByVal lpsz As Long) As String
    Dim Out         As String
    Dim Length      As Long
    Length = lstrlenW(lpsz)
    If Length > 0 Then
        Out = StrConv(String$(Length, vbNullChar), vbUnicode)
        CopyMemory ByVal Out, ByVal lpsz, Length * 2
        PtrToStrW = StrConv(Out, vbFromUnicode)
    End If
End Function


' 将Himetrics转变为 Pixels
Private Function ConvertHimetrix2Pixels(HiMetrix As Long, Horizontally As Boolean) As Long
    If Horizontally Then
        ConvertHimetrix2Pixels = HiMetrix * 1440 / 2540 / Screen.TwipsPerPixelX
    Else
        ConvertHimetrix2Pixels = HiMetrix * 1440 / 2540 / Screen.TwipsPerPixelY
    End If
End Function

