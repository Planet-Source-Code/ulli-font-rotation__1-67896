VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'print rotated text

Option Explicit

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Const LOGPIXELSY    As Long = 90
Private Const _
        ANTIALIASED_QUALITY As Long = 4

Private Enum FontWeight 'not used; just FYI
    FW_DONTCARE = 0
    FW_THIN = 100
    FW_EXTRALIGHT = 200
    FW_LIGHT = 300
    FW_NORMAL = 400
    FW_MEDIUM = 500
    FW_SEMIBOLD = 600
    FW_BOLD = 700
    FW_EXTRABOLD = 800
    FW_HEAVY = 900
End Enum
#If False Then
Private FW_DONTCARE, FW_THIN, FW_EXTRALIGHT, FW_LIGHT, FW_NORMAL, FW_MEDIUM, FW_SEMIBOLD, FW_BOLD, FW_EXTRABOLD, FW_HEAVY, ANTIALIASED_QUALITY
#End If

Private Type LOGFONT
    lfHeight                As Long
    lfWidth                 As Long
    lfEscapement            As Long
    lfOrientation           As Long 'tenth of degrees
    lfWeight                As FontWeight
    lfItalic                As Byte
    lfUnderline             As Byte
    lfStrikeOut             As Byte
    lfCharSet               As Byte
    lfOutPrecision          As Byte
    lfClipPrecision         As Byte
    lfQuality               As Byte
    lfPitchAndFamily        As Byte
    lfFacename              As String * 33
End Type

Public Sub PrintRotated(Where As Object, Text As String, ByVal Rotation As Long, ByVal OriginX As Long, ByVal OriginY As Long, Optional Color As Long = -1)
Attribute PrintRotated.VB_UserMemId = 0

  'the object must support printing

  Dim hPrevFont As Long
  Dim PrevColor As Long
  Dim PrevOrigX As Long
  Dim PrevOrigY As Long
  Dim Idx       As String

    With Where
        On Error Resume Next
            PrevOrigX = .CurrentX
            If Err Then 'object does not support printing
                Idx = "(" & .Index & ")"
                MsgBox .Name & Idx & " is unsuitable for rotated printing.", vbExclamation
              Else 'ERR = FALSE/0
                PrevOrigY = .CurrentY
                PrevColor = .ForeColor

                .CurrentX = OriginX
                .CurrentY = OriginY
                If Color <> -1 Then
                    .ForeColor = Color
                End If

                hPrevFont = SelectObject(.hDC, CreateFontIndirect(SetupRotatedFont(Rotation, Where)))
                Where.Print Text;
                DeleteObject SelectObject(.hDC, hPrevFont)

                .ForeColor = PrevColor
                .CurrentX = PrevOrigX
                .CurrentY = PrevOrigY
            End If
        On Error GoTo 0
    End With 'WHERE

End Sub

Private Function SetupRotatedFont(ByVal Angle As Long, CurrObj As Object) As LOGFONT 'rotation angle, in degrees * 10, counterclockwise

  'using the current font properties of object; so changing the object's font will be honored

    With SetupRotatedFont
        .lfFacename = CurrObj.Font.Name + Chr$(0) 'null terminated
        .lfHeight = MulDiv(CurrObj.Font.Size, GetDeviceCaps(CurrObj.hDC, LOGPIXELSY), 60)
        .lfEscapement = Angle
        .lfOrientation = .lfEscapement
        .lfQuality = ANTIALIASED_QUALITY
        .lfItalic = CurrObj.Font.Italic
        .lfUnderline = CurrObj.Font.Underline
        .lfStrikeOut = CurrObj.Font.Strikethrough
        .lfWeight = CurrObj.Font.Weight
    End With 'SETUPROTATEDFONT

End Function

':) Ulli's VB Code Formatter V2.22.15 (2007-Feb-18 01:50)  Decl: 46  Code: 60  Total: 106 Lines
':) CommentOnly: 3 (2,8%)  Commented: 7 (6,6%)  Empty: 17 (16%)  Max Logic Depth: 4
