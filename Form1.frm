VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Font Rotation Example"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Ausgefüllt
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   2100
      Left            =   7380
      ScaleHeight     =   2100
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   120
      Width           =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6795
      TabIndex        =   0
      Top             =   7185
      Width           =   720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PR                  As clsPR 'print rotated

Private CenterX             As Single
Private CenterY             As Single
Private PrevAngle           As Single
Private ThisAngle           As Single
Private Pi                  As Single
Private TwoPi               As Single
Private Const Radius        As Long = 45

Private Function Atn2(ByVal X As Single, ByVal Y As Single) As Single

  'computes the angle in degrees from (relative) mouse coords

  'quadrants are numbered counterclockwise(!) as follows; the o indicating the center

  '            |
  '            |
  '     II     |     I
  '            |
  '            |
  ' -----------o-----------
  '            |
  '            |
  '    III     |     IV
  '            |
  '            |

    If X = 0 Then
        X = 1E-16  'prevent infinity
    End If

    Atn2 = Atn(Abs(Y) / Abs(X)) 'returns the correct value for quadrant 1 only
    Select Case True

      Case X < 0 And Y >= 0     'quadrant II
        Atn2 = Pi - Atn2        'adjust for q2

      Case X < 0 And Y < 0      'quadrant III
        Atn2 = Atn2 + Pi        'adjust for q3

      Case X >= 0 And Y < 0     'quadrant IV
        Atn2 = TwoPi - Atn2     'adjust for q4

    End Select
    Atn2 = Atn2 * 180 / Pi      'convert to degrees

End Function

Private Sub Command1_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Set PR = New clsPR 'print rotated

    CenterX = ScaleWidth / 2
    CenterY = ScaleHeight / 2
    Pi = 4 * Atn(1)
    TwoPi = 8 * Atn(1)
    Form_MouseMove 0, 0, CenterX + 240, CenterY + 240

    PR.PrintRotated Picture1, "Move mouse around...", 2700, Picture1.ScaleWidth - 15, 120

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Sqr((X - CenterX) ^ 2 + (CenterY - Y) ^ 2) >= Radius Then 'not too close to center

        ThisAngle = Atn2(X - CenterX, CenterY - Y) 'y increases towards the bottom

        If ThisAngle <> PrevAngle Then 'mouse was moved

            Cls

            FillStyle = vbFSSolid
            Circle (CenterX, CenterY), Radius 'draw cemter

            FillStyle = vbFSTransparent
            Circle (CenterX, CenterY), TextWidth(Caption) + 330 'draw outer circle

            PR.PrintRotated Me, Caption, ThisAngle * 10, CenterX, CenterY, vbRed 'print caption rotated
            PR.PrintRotated Me, Format$(ThisAngle, "000.0°"), (ThisAngle + 90) * 10, CenterX, CenterY 'print rotation angle

            PrevAngle = ThisAngle

        End If

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set PR = Nothing 'kill class

End Sub

':) Ulli's VB Code Formatter V2.22.15 (2007-Feb-18 01:43)  Decl: 11  Code: 95  Total: 106 Lines
':) CommentOnly: 13 (12,3%)  Commented: 14 (13,2%)  Empty: 36 (34%)  Max Logic Depth: 3
