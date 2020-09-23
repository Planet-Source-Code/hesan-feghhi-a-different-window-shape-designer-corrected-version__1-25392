VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FSh 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "HN Form Shaper ! Plus"
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10485
   Icon            =   "FSh.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Caption         =   "Mask Background"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Shaper Procedure  Name"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   1560
      ScaleHeight     =   7065
      ScaleWidth      =   8625
      TabIndex        =   7
      Top             =   480
      Width           =   8655
      Begin VB.Line Line1 
         DrawMode        =   6  'Mask Pen Not
         Visible         =   0   'False
         X1              =   0
         X2              =   120
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Img 
      Left            =   9240
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Pictures|*.bmp;*.jpg;*.jpe;*.gif;*.ico;*.wmf"
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   9240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Form Files (*.frm)|*.frm"
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Browse Image"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Affect Form"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   120
      Picture         =   "FSh.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   375
   End
   Begin VB.Image ImgColor 
      Height          =   375
      Left            =   120
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Settings"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"FSh.frx":074C
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Polyline Designer"
      Height          =   495
      Left            =   600
      TabIndex        =   8
      Top             =   820
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "BackGround"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Form Shape"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "HN Form Shape Designer For Microsoft Visual Basic 6.0 (With no ActiveX Required)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   -120
      TabIndex        =   3
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "FSh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private polyline(15, 200) As POINTAPI
Private rects(15) As RECT
Private curact$, linenum, polylinenum, circnum, rrectnum, rectnum, fname$
Private curbgpic As String

Private Sub Command1_Click()
 curact = "Polyline"
End Sub

Private Sub Command2_Click()
 R$ = InputBox("Enter/Change Procedure name here :", "Procedure Name", fname$)
 If StrPtr(R$) <> 0 Then fname$ = R$
End Sub

Private Sub Command3_Click()
 End
End Sub

Private Sub Command4_Click()
 On Error GoTo 2
 CD.DialogTitle = "Affect Form File"
 CD.ShowOpen
 If StrPtr(CD.FileName) <> 0 Then
  Open CD.FileName For Append As #1
  Print #1, "'You must just add the file: region.bas to your project to activate"
  Print #1, "'this function"
  Print #1, "'You can simply change the shape of your form by calling:"
  Print #1, "'    Form_DrawShape()"
  Print #1, "Private Sub " + fname$
  Print #1, " Me.Picture = LoadPicture(""" + curbgpic + """)"
  If Check1.Value = 1 Then
   Print #1, " MaskColor =" + Str$(Shape1.FillColor)
   Print #1, " lRgn = CreateRectRgn(0, 0, Me.Width, Me.Height)"
   Print #1, " For X = 0 To Me.Width/Screen.TwipsPerPixelX"
   Print #1, "  For Y = 0 To Me.Height/Screen.TwipsPerPixelY"
   Print #1, "   If Me.point(X, Y) = MaskColor Then"
   Print #1, "    lRgn2 = CreateRectRgn(X, Y, X + 1, Y + 1)"
   Print #1, "    CombineRgn lRgn, lRgn, lRgn2, RGN_DIFF"
   Print #1, "    DeleteObject lRgn2"
   Print #1, "   End If"
   Print #1, "  Next Y"
   Print #1, " Next X"
  End If
   For pln = 0 To polylinenum
    If polyline(pln, 0).X = -1 Then GoTo 1
    cnt = 0
    Print #1, " Dim pnt" + Right$(Str$(pln), Len(Str$(pln)) - 1) + "(200) as POINTAPI"
    For ln = 0 To 100
     If polyline(pln, ln).X > -1 Then
      Print #1, " pnt" + Right$(Str$(pln), Len(Str$(pln)) - 1) + "(" + Str$(ln) + ").X =" + Str$(polyline(pln, ln).X / 15)
      Print #1, " pnt" + Right$(Str$(pln), Len(Str$(pln)) - 1) + "(" + Str$(ln) + ").Y =" + Str$(polyline(pln, ln).Y / 15)
      cnt = cnt + 1
     Else
      Exit For
     End If
    Next ln
    Print #1, " Rgn" + Right$(Str$(pln), Len(Str$(pln)) - 1) + "&=CreatePolygonRgn(pnt" + Right$(Str$(pln), Len(Str$(pln)) - 1) + "(0)," + Str$(cnt) + ",1)"
    If trn = 1 Then Print #1, " CombineRgn Rgn" + Right$(Str$(pln), Len(Str$(pln)) - 1) + ", Rgn" + Right$(Str$(pln), Len(Str$(pln)) - 1) + ", Rgn" + Right$(Str$(pln - 1), Len(Str$(pln - 1)) - 1) + ", RGN_OR"
    trn = 1
    Print #1, " hRgn& = Rgn" + Right$(Str$(pln), Len(Str$(pln)) - 1)
1
   Next pln
   Print #1, " SetWindowRgn Me.hWnd, hRgn&, True"
   Print #1, "End Sub"
  Print #1, ""
  Print #1, "'You can also copy-paste this code into a module:"
  Print #1, "'"
  Print #1, "'Public Type POINTAPI"
  Print #1, "'        X As Long"
  Print #1, "'        Y As Long"
  Print #1, "'End Type"
  Print #1, "'Public Type RECT"
  Print #1, "'        Left As Long"
  Print #1, "'        Top As Long"
  Print #1, "'        Right As Long"
  Print #1, "'        Bottom As Long"
  Print #1, "'End Type"
  Print #1, "'"
  Print #1, "'Public Const RGN_AND = 1"
  Print #1, "'Public Const RGN_COPY = 5"
  Print #1, "'Public Const RGN_DIFF = 4"
  Print #1, "'Public Const RGN_MAX = RGN_COPY"
  Print #1, "'Public Const RGN_MIN = RGN_AND"
  Print #1, "'Public Const RGN_OR = 2"
  Print #1, "'Public Const RGN_XOR = 3"
  Print #1, "'"
  Print #1, "'Public Declare Function SetWindowRgn Lib ""user32"" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long"
  Print #1, "'Public Declare Function CreateRoundRectRgn Lib ""gdi32"" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long"
  Print #1, "'Public Declare Function CombineRgn Lib ""gdi32"" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long"
  Print #1, "'Public Declare Function CreateRectRgn Lib ""gdi32"" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long"
  Print #1, "'Public Declare Function DeleteObject Lib ""gdi32"" (ByVal hObject As Long) As Long"
  Print #1, "'Public Declare Function GetDesktopWindow Lib ""user32"" () As Long"
  Print #1, "'Public Declare Function CreatePolygonRgn Lib ""gdi32"" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long"
  Close #1
 End If
2  Close #1
End Sub

Private Sub Command5_Click()
 If Picture1.Tag = "" Then
  Img.DialogTitle = "Choose BackGround Picture"
  Img.ShowOpen
  If StrPtr(Img.FileName) <> 0 Then
   Picture1.Picture = LoadPicture(Img.FileName)
   Picture1.Tag = Img.FileName
  End If
  Command5.Caption = "Remove BG"
 Else
  Picture1.Picture = LoadPicture("")
  Picture1.Tag = ""
  Command5.Caption = "Browse Image"
 End If
 'Reset the Polygons
 linenum = 0
 polylinenum = 0
End Sub

Private Sub Command6_Click()
 R$ = InputBox("Enter the name of the function here:", "Function Name", fname$)
 If StrPtr(R$) <> 0 Then fname$ = R$
End Sub

Private Sub Form_Load()
 For i = 0 To 15
  For j = 0 To 200
   polyline(i, j).X = -1
 Next j, i
 fname$ = "Form_DrawShape()"
 hRgn& = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 30, 30)
 SetWindowRgn Me.hwnd, hRgn, True
End Sub

Private Sub ImgColor_Click()
 Img.ShowColor
 Shape1.FillColor = Img.Color
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Line1.x1 = Picture1.CurrentX
 Line1.y1 = Picture1.CurrentY
 Line1.x2 = X
 Line1.y2 = Y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Static cstep
 If Button = 1 Then
   'For Polyline
   If curact = "Polyline" Then
    If linenum = 0 Then
     Picture1.PSet (X, Y)
     polyline(polylinenum, linenum).X = X
     polyline(polylinenum, linenum).Y = Y
     linenum = linenum + 1
     Line1.Visible = True
    Else
     Picture1.Line -(X, Y)
     polyline(polylinenum, linenum).X = X
     polyline(polylinenum, linenum).Y = Y
     linenum = linenum + 1
    End If
   End If
 ElseIf Button = 2 Then
   If curact = "Polyline" Then
    Line1.Visible = False
    Picture1.Line -(polyline(polylinenum, 0).X, polyline(polylinenum, 0).Y)
    polyline(polylinenum, linenum).X = polyline(polylinenum, 0).X
    polyline(polylinenum, linenum).Y = polyline(polylinenum, 0).Y
    linenum = 0
    polylinenum = polylinenum + 1
    curact = ""
   End If
 End If
End Sub

