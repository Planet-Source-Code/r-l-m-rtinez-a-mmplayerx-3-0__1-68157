VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWindowSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SendmFormPadressage Lib "user32" Alias "SendmFormPadressageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetRegionDataByte Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Byte) As Long
Private Declare Function GetRegionDataLong Lib "gdi32" Alias "GetRegionData" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Long) As Long
Private Declare Function ExtCreateRegionByte Lib "gdi32" Alias "ExtCreateRegion" (lpXform As Long, ByVal nCount As Long, lpRgnData As Byte) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

 Const RGN_OR = 2

'
' GDI declarations
'

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long

' The function name is GetObject, but VB has a simlar named
' internal function...
Private Declare Function GDIGetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

 Const IMAGE_BITMAP = 0
 Const LR_DEFAULTCOLOR = &H0
 Const LR_LOADFROMFILE = &H10
 Const SRCCOPY = &HCC0020
 Const WM_NCLBUTTONDOWN = &HA1
 Const HTCAPTION = 2
Const API_NULL_HANDLE = 0

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Private Enum FormEdges
    FE_TOP_LEFT
    FE_TOP_RIGHT
    FE_BOTTOM_LEFT
    FE_BOTTOM_RIGHT
    FE_TOP_H_SEGMENT
    FE_BOTTOM_H_SEGMENT
    FE_RIGHT_V_SEGMENT
    FE_LEFT_V_SEGMENT
    FE_TOP_HEADER
    FE_LAST = FE_TOP_HEADER
End Enum

Private Type RegionDataType
    RegionData() As Byte
    DataLength As Long
End Type



Private mDC As Long         ' Memory DC
Private mBitmap As Long     ' Bitmap handle
Private mOldBitmap As Long  ' "Original" Bitmap handle

Private mWidth As Long
Private mHeight As Long

Private Type RECT
    Left As Long
    Top As Long
    Height As Long
    Width As Long
End Type

Private mFormPadre As Form
Private LoadedRegionsFromFile As Boolean
Private bAjustando As Boolean
Private AreaTrabajo As RECT
Private lColorInv As Long

 ' Current number of horizontal/vertical segments
Private NumXSlices As Long
Private NumYSlices As Long

' Minimum needed number of X slices so we don't mess-up
' the button positions
Private MinXSlices As Long
Private MinYSlices As Long
                        
' Width/height of pad w/o any horizontal segment
Private BaseXSize As Long
Private BaseYSize As Long

' Used when resizing the window -
' X/Y distance of the mouse pointer from the form's edge
Private XDistance As Long
Private YDistance As Long

' Boolean flags - the current state of the form
Private InXDrag As Boolean ' In horizontal resize
Private InYDrag As Boolean ' In vertical resize

Private NoRedraw As Boolean

' Set to TRUE when in ListSkins(), to prevent lstSkins_Click()
' events from being handled while the list is created
Private InListSkins As Boolean

' Size of right/bottom segments
Private XEdgeSize As Single
Private YEdgeSize As Single

' Holds the actual edge skin bitmaps
Private EdgeImages(FE_LAST) As cWindowSkin

' Holds the region data for each of the skin bitmaps
Private EdgeRegions(FE_LAST) As RegionDataType

Private WindowRegion As Long ' Current window region

Private EdgeImageFileNames(FE_LAST) As String

'//////////////////////////////////////////////////////////////////////////////
Private mUpImage As New cWindowSkin
Private mDownImage As New cWindowSkin

Private mButtonWidth As Long
Private mButtonHeight As Long

Private mX As Long
Private mY As Long
Private bExitClick As Boolean

Public Sub Formulario_Down(X As Single, Y As Single)
        ' Test whether the user has pressed a "button",
        ' and show the 'down button' image if so
        If HitTest(CLng(X), CLng(Y)) Then
           PaintDownImage
           Exit Sub
        End If
     
        
        YDistance = Y - mFormPadre.ScaleHeight
        XDistance = X - mFormPadre.ScaleWidth
        
        ' If the mouse pointer is on the the bottom edge,
        ' flag Y (vertical) drag
        If Abs(YDistance) < YEdgeSize Then
            InYDrag = True
            bAjustando = True
        End If
        
        ' If the mouse pointer is on the the right edge,
        ' flag X drag. Don't start drag if wer'e in the window
        ' title area
        If Abs(XDistance) < XEdgeSize And _
           Y > EdgeImages(FE_TOP_RIGHT).Height Then
            InXDrag = True
            bAjustando = True
        End If
End Sub

Public Sub Formulario_MouseMove(iBoton As Integer, X As Single, Y As Single)
Dim NewYSlices As Single
Dim NewXSlices As Single
Dim ShowXResizeCursor As Boolean
Dim ShowYResizeCursor As Boolean
Dim ResizingNeeded As Boolean

    ' Determine what kind of cursor should be shown
    
    If Abs(Y - mFormPadre.ScaleHeight) < YEdgeSize Or InYDrag Then
        ShowYResizeCursor = True
    End If
    
    If (Abs(X - mFormPadre.ScaleWidth) < XEdgeSize And _
        Y > EdgeImages(FE_TOP_RIGHT).Height) Or InXDrag Then
        
        ShowXResizeCursor = True
    End If
    
    If ShowXResizeCursor And ShowYResizeCursor Then
        mFormPadre.MousePointer = vbSizeNWSE
        
    ElseIf ShowXResizeCursor Then
        mFormPadre.MousePointer = vbSizeWE
    
    ElseIf ShowYResizeCursor Then
        mFormPadre.MousePointer = vbSizeNS
    
    Else
        mFormPadre.MousePointer = vbDefault
    End If
    If iBoton <> vbLeftButton Then Exit Sub

    If InXDrag Then
        ' Compute new number of horizontal segments
        NewXSlices = (X - BaseXSize - XDistance) / EdgeImages(FE_TOP_H_SEGMENT).Width
        
        If NewXSlices < MinXSlices Then NewXSlices = MinXSlices
           
        
        ' Check if we should actually do the resize. Not every
        ' slightest mouse drag should cause a resize
        If (NewXSlices - NumXSlices >= 0.5) Or _
           (NewXSlices - NumXSlices < -0.5) Then
            
            NumXSlices = NewXSlices
            ResizingNeeded = True
        End If
    End If

    ' Same handling for vertical resize-drag
    If InYDrag Then
        
        NewYSlices = (Y - BaseYSize - YDistance) / EdgeImages(FE_LEFT_V_SEGMENT).Height
        If NewYSlices < MinYSlices Then NewYSlices = MinYSlices
                
        If NewYSlices - NumYSlices >= 0.5 Or _
           (NewYSlices - NumYSlices < -0.5) Then
            
            NumYSlices = NewYSlices
            ResizingNeeded = True
        End If
    End If

    If ResizingNeeded Then
      AjustarAncho
      

    End If

    
End Sub

Public Sub Formulario_MouseUp(X As Single, Y As Single)
    ' Clear window dragging/resizing flags
    PaintUpImage
    ' Test whether the user has released a "button",
    ' and commit the appropriate action if so
    bExitClick = HitTest(CLng(X), CLng(Y))

    InXDrag = False
    InYDrag = False
    bAjustando = False
End Sub
Public Sub Formulario_Paint()
    
    If Not NoRedraw Then
'        mFormPadre.Cls
        DrawEdges mFormPadre, EdgeImages, NumXSlices, NumYSlices, False
        PaintUpImage
    End If

End Sub


Public Sub CargarSkin(Direccion As String)
Dim i As Long
Dim PrevXSliceSize As Long, PrevYSliceSize As Long
Dim filename As String

    EdgeImageFileNames(FE_TOP_LEFT) = "top_left.bmp"
    EdgeImageFileNames(FE_TOP_RIGHT) = "top_right.bmp"
    EdgeImageFileNames(FE_BOTTOM_LEFT) = "bottom_left.bmp"
    EdgeImageFileNames(FE_BOTTOM_RIGHT) = "bottom_right.bmp"
    EdgeImageFileNames(FE_TOP_H_SEGMENT) = "hsegment_top.bmp"
    EdgeImageFileNames(FE_BOTTOM_H_SEGMENT) = "hsegment_bottom.bmp"
    EdgeImageFileNames(FE_RIGHT_V_SEGMENT) = "vsegment_right.bmp"
    EdgeImageFileNames(FE_LEFT_V_SEGMENT) = "vsegment_left.bmp"
    EdgeImageFileNames(FE_TOP_HEADER) = "hsegment_top_header.bmp"


    ' Save for later. You'll see.
    If Not EdgeImages(0) Is Nothing Then
        PrevXSliceSize = EdgeImages(FE_TOP_H_SEGMENT).Width
        PrevYSliceSize = EdgeImages(FE_LEFT_V_SEGMENT).Height
    End If
    
    ' Initialize bitmaps array
    For i = 0 To FE_LAST
        Set EdgeImages(i) = New cWindowSkin
    Next
    
    ' Load skin bitmaps. Check that the files actally  exist
    If Right(Direccion, 1) <> "\" Then Direccion = Direccion & "\"
    
    For i = 0 To FE_LAST
        filename = Direccion & EdgeImageFileNames(i)
       
        If Dir(filename) = "" Then
            err.Raise 1, , "Image file " & filename & " not found!"
                        
        ElseIf EdgeImages(i).LoadFile(filename) = False Then
            err.Raise 1, , "Could not load image file: " & filename
        End If
    Next
        
    ' See documentation in start of file for all those variables
    BaseXSize = EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_RIGHT).Width
    BaseYSize = EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height

    XEdgeSize = EdgeImages(FE_RIGHT_V_SEGMENT).Width
    YEdgeSize = EdgeImages(FE_BOTTOM_H_SEGMENT).Height

    ' Here we compute how much horizontal/vertical segments
    ' sould be drawn
    If PrevXSliceSize <> 0 Then
        ' Skin was changed, match number of x/y slices
        ' according to the currect/previous sizes of the slices
        NumXSlices = Round(NumXSlices * PrevXSliceSize / EdgeImages(FE_TOP_H_SEGMENT).Width)
        NumYSlices = Round(NumYSlices * PrevYSliceSize / EdgeImages(FE_LEFT_V_SEGMENT).Height)
    Else
        ' Program was just loaded, match number of x/y slices
        ' to the default client width/height
        NumXSlices = Round(mFormPadre.ScaleWidth / EdgeImages(FE_TOP_H_SEGMENT).Width)
        NumYSlices = Round(mFormPadre.ScaleHeight / EdgeImages(FE_LEFT_V_SEGMENT).Height)
    End If
    
    NumXSlices = IIf(NumXSlices < MinXSlices, MinXSlices, NumXSlices)
    NumYSlices = IIf(NumYSlices < MinYSlices, MinYSlices, NumYSlices)
    
    ' Initialize exit/minimize buttons
    mUpImage.LoadFile Direccion & "exitbutton_up.bmp"
    mDownImage.LoadFile Direccion & "exitbutton_down.bmp"
    
    mButtonWidth = mUpImage.Width
    mButtonHeight = mUpImage.Height

    ' Create and store region data for each of the skin bitmaps,
    ' for use whenever creating the window region
    
    If Not LoadedRegionsFromFile Then
        For i = 0 To FE_LAST
            CreateRegionData EdgeImages(i), EdgeRegions(i)
        Next
    
        SaveEdgeRegions EdgeRegions, Direccion & "trans.dat"
    End If
  
    AjustarAncho
End Sub

Public Property Let CargarRegionDeArchivo(bValor As Boolean)
 LoadedRegionsFromFile = bValor
End Property

Public Property Get CargarRegionDeArchivo() As Boolean
CargarRegionDeArchivo = LoadedRegionsFromFile
End Property



Public Sub AjustarAncho()
Dim NewSize As Long
    
    ' We don't want form redraws when in middle of new size
    ' setting, before the new region was set
    NoRedraw = True
    
    ' Compute width/height of form accodring to the number of
    ' x/y slices
    mFormPadre.Width = (EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices + EdgeImages(FE_TOP_RIGHT).Width) * Screen.TwipsPerPixelX
    mFormPadre.Height = (EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices + EdgeImages(FE_BOTTOM_LEFT).Height) * Screen.TwipsPerPixelY

    ' Compute size of client area
    DrawEdges mFormPadre, EdgeImages, NumXSlices, NumYSlices, False
    
    NoRedraw = False
    
    ' Create new window region. Also triggers a redraw, now that
    ' wer'e done setting the new form shape
    
    BuildWindowRegion
    
End Sub
Public Property Get Ajustando() As Boolean
Ajustando = bAjustando
End Property

Public Property Get AreaLeft() As Integer
 AreaLeft = AreaTrabajo.Left
End Property

Public Property Get AreaTop() As Integer
 AreaTop = AreaTrabajo.Top
End Property

Public Property Get AreaWidth() As Integer
 AreaWidth = AreaTrabajo.Width
End Property

Public Property Get AreaHeight() As Integer
 AreaHeight = AreaTrabajo.Height
End Property


Private Sub BuildWindowRegion()
Dim PrevRegion As Long
Dim iLeft As Integer
Dim iTop As Integer
Dim iWidth As Integer
Dim iHeight As Integer


  PrevRegion = WindowRegion
  
    iTop = EdgeImages(FE_TOP_LEFT).Height
    iLeft = EdgeImages(FE_LEFT_V_SEGMENT).Width
    iHeight = mFormPadre.ScaleHeight - (EdgeImages(FE_TOP_H_SEGMENT).Height + EdgeImages(FE_BOTTOM_H_SEGMENT).Height)
    iWidth = mFormPadre.ScaleWidth - (EdgeImages(FE_LEFT_V_SEGMENT).Width + EdgeImages(FE_RIGHT_V_SEGMENT).Width)

    AreaTrabajo.Top = iTop
    AreaTrabajo.Left = iLeft
    AreaTrabajo.Height = iHeight
    AreaTrabajo.Width = iWidth
    
    ' Create initial region that covers the client area
'    WindowRegion = CreateRectRgn(iLeft, iTop, iWidth, iHeight)
    WindowRegion = CreateRectRgn(iLeft, iTop, iLeft + iWidth, iTop + iHeight)

    ' Add to it the window region of the form edges
    BuildEdgesRegion WindowRegion, EdgeImages, EdgeRegions, NumXSlices, NumYSlices

    ' Finally - set the full region
    SetWindowRgn mFormPadre.hwnd, WindowRegion, True
    
    ' Don't forget - delete old window region
    DeleteObject PrevRegion
    
End Sub



' This function builds the window region of the form's edges -
' the corners and the sides, using the pre-created regions data
' Each created region is combined with the full window region
Private Sub BuildEdgesRegion(WindowRegion As Long, EdgeImages() As cWindowSkin, EdgeRegions() As RegionDataType, NumXSlices As Long, NumYSlices As Long)
Dim i As Long

    ' Make region for top-left corner. That's an easy one
    MakeRegionWithOffset EdgeRegions(FE_TOP_LEFT), 0, 0, WindowRegion

    ' Top-right corner
    MakeRegionWithOffset EdgeRegions(FE_TOP_RIGHT), _
        EdgeImages(FE_TOP_LEFT).Width + (EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices), 0, _
        WindowRegion
    
    ' Bottom-left corner
    MakeRegionWithOffset EdgeRegions(FE_BOTTOM_LEFT), 0, _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices, _
        WindowRegion
    
    ' Bottom-right corner
    MakeRegionWithOffset EdgeRegions(FE_BOTTOM_RIGHT), _
        EdgeImages(FE_TOP_LEFT).Width + (EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices), _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices, _
        WindowRegion

    ' Create the regions for the top and bottom sides,
    ' By the number of X slices.
    For i = 1 To NumXSlices
        MakeRegionWithOffset EdgeRegions(FE_TOP_H_SEGMENT), EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), 0, WindowRegion
        
        MakeRegionWithOffset EdgeRegions(FE_BOTTOM_H_SEGMENT), _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices - EdgeImages(FE_BOTTOM_H_SEGMENT).Height, WindowRegion
    Next

    ' Create the regions for the left and right sides,
    ' By the number of Y slices.
    For i = 1 To NumYSlices
        MakeRegionWithOffset EdgeRegions(FE_LEFT_V_SEGMENT), 0, EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1), WindowRegion
        
        MakeRegionWithOffset EdgeRegions(FE_RIGHT_V_SEGMENT), _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices + EdgeImages(FE_TOP_RIGHT).Width - EdgeImages(FE_RIGHT_V_SEGMENT).Width, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1), _
            WindowRegion
    Next
    
End Sub

' This fucntion is almost identical to MakeEdgesRegion,
' excepts that it actually DRAWS the edges.
Private Sub DrawEdges(DestForm As Form, _
                     EdgeImages() As cWindowSkin, _
                     NumXSlices As Long, NumYSlices As Long, _
                     SimpleStyle As Boolean, _
                     Optional DrawActiveState As Boolean = True)
Dim i As Long
Dim iLeft As Long

        
    EdgeImages(FE_TOP_LEFT).Paint DestForm.hDC, 0, 0

    EdgeImages(FE_TOP_RIGHT).Paint DestForm.hDC, _
        EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices, 0
    
    EdgeImages(FE_BOTTOM_LEFT).Paint DestForm.hDC, _
        0, EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
    
    EdgeImages(FE_BOTTOM_RIGHT).Paint DestForm.hDC, _
        EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * NumXSlices, _
        EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices
    
    For i = 1 To NumXSlices
        EdgeImages(FE_TOP_H_SEGMENT).Paint DestForm.hDC, _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), 0
        EdgeImages(FE_BOTTOM_H_SEGMENT).Paint DestForm.hDC, _
            EdgeImages(FE_TOP_LEFT).Width + EdgeImages(FE_TOP_H_SEGMENT).Width * (i - 1), _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_BOTTOM_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * NumYSlices - EdgeImages(FE_BOTTOM_H_SEGMENT).Height
    Next

    For i = 1 To NumYSlices
        EdgeImages(FE_LEFT_V_SEGMENT).Paint DestForm.hDC, 0, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1)
        EdgeImages(FE_RIGHT_V_SEGMENT).Paint DestForm.hDC, _
            DestForm.ScaleWidth - EdgeImages(FE_RIGHT_V_SEGMENT).Width, _
            EdgeImages(FE_TOP_LEFT).Height + EdgeImages(FE_LEFT_V_SEGMENT).Height * (i - 1)
    Next
    
    iLeft = (DestForm.ScaleWidth / 2) - (EdgeImages(FE_TOP_HEADER).Width / 2)
        
    EdgeImages(FE_TOP_HEADER).Paint DestForm.hDC, iLeft, 0
    
End Sub

' Added in v1.2 of demo
'
' Save the region data of all edges to a file. When loading the skin later on,
' we won't have to compute the region again (using MakeRegion()), only load
' the file. This method typically yields about 50% speed increase in loading skins.
' In kewlpAd 1.2, which has more complex skins, it's very noticeable!
'
' NOTE: If you change the transparent areas in a bitmap, the file will be outdated.
' That's why I added the 'Load region data from file' checkbox
'
Private Sub SaveEdgeRegions(EdgeRegions() As RegionDataType, _
                           filename As String)

Dim i As Long

    Open filename For Binary As #1

    For i = 0 To FE_LAST
        Put 1, , EdgeRegions(i).DataLength
        Put 1, , EdgeRegions(i).RegionData
    Next

    Close
    
End Sub

' Load the edges' region data from file
Private Function LoadEdgeRegions(EdgeRegions() As RegionDataType, _
                                filename As String) As Boolean

Dim i As Long
    
    If Dir(filename) = "" Then Exit Function
    
    Open filename For Binary As #1
    
    For i = 0 To FE_LAST
        Get 1, , EdgeRegions(i).DataLength
        ReDim EdgeRegions(i).RegionData(EdgeRegions(i).DataLength + 32)
        Get 1, , EdgeRegions(i).RegionData
    Next
    
    Close
    
    LoadEdgeRegions = True

End Function
''/////////////////////////////////////////////////////////////////////////////////
Private Function MakeRegion(hDC As Long, Width As Long, Height As Long) As Long
    
    ' The usual "Make region by bitmap" procedure
    ' TransparentColor is total pink &HFF00FF
    
    Dim X As Long, Y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean  ' Flags whether we are in a non-tranparent pixel sequence
    
    InFirstRegion = True: InLine = False
    X = Y = StartLineX = 0
    
    TransparentColor = lColorInv
    
    For Y = 0 To Height - 1
        For X = 0 To Width
            
            If GetPixel(hDC, X, Y) = TransparentColor Or X = Width Then
                ' We reached a transparent pixel
                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, Y, X, Y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR
                        ' Always clean up your mess
                        DeleteObject LineRegion
                    End If
                End If
            Else
                ' We reached a non-transparent pixel
                If Not InLine Then
                    InLine = True
                    StartLineX = X
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function


' Given a bitmap, this procedure fills a RegionDataType
' with binary data describing its region.
' This can be described as "Serializing" a region.

Private Sub CreateRegionData(EdgeImage As cWindowSkin, EdgeRegion As RegionDataType)
Dim WinRegion As Long
Dim Ret As Long

    ' First create the region for the bitmap
    WinRegion = MakeRegion(EdgeImage.hDC, EdgeImage.Width, EdgeImage.Height)
    
    ' Get the size needed for the region data buffer
    EdgeRegion.DataLength = GetRegionDataLong(WinRegion, 0&, ByVal 0&)

    If EdgeRegion.DataLength = 0 Then
        err.Raise 1, , "Could not create region data"
    Else
        ' Actually get the data into the buffer - a byte array
        ' of the proper size.
        ' You need 32 bytes more, because the API call attaches
        ' a 32-byte structure called RGNDATAHEADER before the
        ' data itself
        ReDim EdgeRegion.RegionData(EdgeRegion.DataLength + 32)
        
        Ret = GetRegionDataByte(WinRegion, EdgeRegion.DataLength, EdgeRegion.RegionData(0))
    End If
    
    DeleteObject WinRegion

End Sub


' Create a new region by binary region data, then "move" the
' region by the given offset, and combine it with a previous
' region, if supplied.
' The region data was previously created by a call to
' CreateRegionData()
Private Sub MakeRegionWithOffset(Region As RegionDataType, _
                                  XOffset As Long, _
                                  YOffset As Long, _
                                  PrevRegion As Long)
    
Dim NewRegion As Long
    
    ' The API call requires the address of the region data,
    ' so we pass the first cell in the array. VB passes arrays
    ' ByRef, so here's our address.
    NewRegion = ExtCreateRegionByte(ByVal 0&, Region.DataLength, Region.RegionData(0))
    
    OffsetRgn NewRegion, XOffset, YOffset
    
    If PrevRegion = API_NULL_HANDLE Then
        PrevRegion = NewRegion
    Else
        CombineRgn PrevRegion, PrevRegion, NewRegion, RGN_OR
        DeleteObject NewRegion
    End If

End Sub

Public Property Let MinimoAncho(Valor As Integer)
   MinXSlices = Valor
End Property

Public Property Get MinimoAncho() As Integer
    MinimoAncho = MinXSlices
End Property

Public Property Let MinimoAlto(Valor As Integer)
   MinYSlices = Valor
End Property

Public Property Get MinimoAlto() As Integer
    MinimoAlto = MinYSlices
End Property

Public Property Let ColorInvisible(Valor As Long)
   lColorInv = Valor
End Property


Public Property Set FormularioPadre(fFormulario As Form)
    Set mFormPadre = fFormulario
    MinXSlices = 10
    MinYSlices = 10
    mX = 1
    mY = 1
    '/color a extraer rosa
    lColorInv = RGB(255, 0, 255)
End Property


Public Property Get FormularioPadre() As Form
    Set FormularioPadre = mFormPadre
End Property

Public Function LoadFile(filename As String) As Boolean

    ' Clear up previous DC/bitmap
    ClearAll
    
    mBitmap = LoadImage(API_NULL_HANDLE, filename, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
    
    If (mBitmap <> API_NULL_HANDLE) Then
        LoadFile = LoadBitmapIntoDC
    End If
    
End Function

' Note: that function will NOT function properly when the
' project is run under the IDE, because then the resources
' loaded will be VB's resources, not your app's resources.
' You may want to watch some interesting side-effects though...
Public Function LoadResource(ResourceID As Long) As Boolean
    
    ' Clear up previous DC/bitmap
    ClearAll
    
    mBitmap = LoadImage(App.hInstance, ResourceID, IMAGE_BITMAP, 0, 0, LR_DEFAULTCOLOR)
    
    If (mBitmap <> API_NULL_HANDLE) Then
        LoadResource = LoadBitmapIntoDC
    End If
    
End Function

Private Function LoadBitmapIntoDC() As Boolean
Dim ScreenDC As Long
Dim BitmapData As BITMAP

    ' Create a compatible memory DC to hold the bitmap
    ScreenDC = GetDC(API_NULL_HANDLE)
    mDC = CreateCompatibleDC(ScreenDC)
    ReleaseDC API_NULL_HANDLE, ScreenDC
    
    If (mDC <> API_NULL_HANDLE) Then
        ' If the DC was created successfully,
        ' select the bitmap into it
        mOldBitmap = SelectObject(mDC, mBitmap)
        
        ' Get the dimensions of the bitmap
        GDIGetObject mBitmap, Len(BitmapData), BitmapData
        mWidth = BitmapData.bmWidth
        mHeight = BitmapData.bmHeight
        
        LoadBitmapIntoDC = True
    End If

End Function

Property Get Width() As Long
    Width = mWidth
End Property

Property Get Height() As Long
    Height = mHeight
End Property

Property Get hDC() As Long
    hDC = mDC
End Property

Public Sub Paint(DestDC As Long, _
                 DestX As Long, _
                 DestY As Long)
    
    BitBlt DestDC, DestX, DestY, _
        mWidth, mHeight, mDC, _
        0, 0, vbSrcCopy
        
End Sub

Private Sub ClearAll()
    
    If (mDC <> API_NULL_HANDLE) Then
        If (mBitmap <> API_NULL_HANDLE) Then
            ' Select the original bitmap into the DC,
            ' and delete our bitmap
            SelectObject mDC, mOldBitmap
            DeleteObject mBitmap
            mBitmap = API_NULL_HANDLE
        End If
        
        ' Delete the memory DC
        DeleteObject mDC
        mDC = API_NULL_HANDLE
    End If

End Sub

Private Sub Class_Terminate()
    ClearAll
End Sub

'#############################################################################
'# PROPIEDADES PARA COLOCAR LOS BOTONES EN SU POSICION CORRECTA
Public Property Get ClickExitButton() As Boolean
    ClickExitButton = bExitClick
End Property

Public Sub ButtonExitXY(X As Long, Y As Long)
    mX = X
    mY = Y

End Sub


' Test whether the given (x,y) coordinate is inside the
' button area, and return TRUE is so.
' Used by the parent form in its MouseDown/MouseUp events to
' determine if a button was pressed/released.
Private Function HitTest(X As Long, Y As Long) As Boolean

    If (X >= AbsX() And X < AbsX() + mButtonWidth) And _
       (Y >= AbsY() And Y < AbsY() + mButtonHeight) Then
        HitTest = True
    End If

End Function

Private Sub PaintUpImage()
    ' Note that the 'real' x/y values are used
    mUpImage.Paint mFormPadre.hDC, AbsX(), AbsY()
End Sub

Private Sub PaintDownImage()
    mDownImage.Paint mFormPadre.hDC, AbsX(), AbsY()
End Sub

' Get the real X position of the button.
' If the X coordinate is negative, compute AbsX as the
' distance from the right side
Private Property Get AbsX() As Long

    If mX >= 0 Then
        AbsX = mX
    Else
        AbsX = mFormPadre.ScaleWidth + mX
    End If

End Property

Private Property Get AbsY() As Long

    If mY >= 0 Then
        AbsY = mY
    Else
        AbsY = mFormPadre.ScaleHeight + mY
    End If

End Property





