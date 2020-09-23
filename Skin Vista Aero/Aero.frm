VERSION 5.00
Begin VB.Form Aero 
   BorderStyle     =   0  'None
   Caption         =   "Skin Vista Aero The mere idea"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   0
   ClientWidth     =   13845
   Icon            =   "Aero.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Vista_Aero.aicAlphaImage Label2 
      Height          =   660
      Left            =   2040
      Top             =   4920
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   1164
      Image           =   "Aero.frx":AE5A
      Props           =   5
   End
   Begin Vista_Aero.aicAlphaImage Command1 
      Height          =   315
      Left            =   840
      Top             =   7200
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   556
      Image           =   "Aero.frx":C3B4
      Props           =   5
   End
   Begin Vista_Aero.aicAlphaImage AlphaImage2 
      Height          =   3150
      Left            =   3360
      Top             =   1440
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   5556
      Image           =   "Aero.frx":C851
      Enabled         =   0   'False
      Props           =   5
   End
   Begin Vista_Aero.aicAlphaImage AlphaImage1 
      Height          =   3150
      Left            =   3360
      Top             =   1440
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   5556
      Image           =   "Aero.frx":2056A
   End
   Begin Vista_Aero.aicAlphaImage ucAlphaImage4 
      Height          =   540
      Left            =   8820
      Top             =   390
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   953
      Image           =   "Aero.frx":33A0A
      Props           =   5
   End
   Begin Vista_Aero.aicAlphaImage ucAlphaImage2 
      Height          =   540
      Left            =   9600
      Top             =   390
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   953
      Image           =   "Aero.frx":34592
      Props           =   5
   End
   Begin Vista_Aero.aicAlphaImage ucAlphaImage5 
      Height          =   450
      Left            =   9210
      Top             =   420
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   794
      Image           =   "Aero.frx":353AF
      Props           =   5
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   9765
      Top             =   520
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   9000
      Top             =   520
      Width           =   435
   End
   Begin VB.Image Image3 
      Height          =   285
      Left            =   9375
      Top             =   520
      Width           =   435
   End
   Begin Vista_Aero.aicAlphaImage Picture1 
      Height          =   8250
      Left            =   60
      Top             =   0
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   14552
      Image           =   "Aero.frx":3DB80
   End
End
Attribute VB_Name = "Aero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************
'* Karmba_a@hotmail.com *
'*      MSLE 2008       *
'************************
Private X1 As Integer, Y1 As Integer

Private Sub AlphaImage1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

End Sub
Private Sub AlphaImage1_MouseEnter()
    ' on mouse enter, we will grayscale and fade the bubble in
    AlphaImage1.ShadowEnabled = False
    AlphaImage1.grayScale = aiNoGrayScale
    AlphaImage2.FadeInOut 100
End Sub

Private Sub AlphaImage1_MouseExit()
    ' on mouse exit we will fade bubble out, then when it
    ' is completely faded out, we will grayscale cheetah
    AlphaImage1.ShadowEnabled = True
    AlphaImage2.FadeInOut 0
End Sub

Private Sub Command1_Click(ByVal Button As Integer)
End
End Sub

Private Sub Form_Click()
On Error Resume Next
        Call SystrayOn(Me, "New Skin Vista Aero The mere idea")
        Call SetForegroundWindow(Me.hwnd)
        Me.Hide
        Me.WindowState = 1
End Sub

Private Sub Form_Load()
On Error Resume Next
AlphaImage2.Opacity = 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    Static lngMsg As Long
    Dim blnflag As Boolean, lngResult As Long
    
    lngMsg = x / Screen.TwipsPerPixelX
    If blnflag = False Then
        blnflag = True
        Select Case lngMsg
        Case WM_LBUTTONCLK      'to popup on left-click
            Call SystrayOff(Me)
            Call SetForegroundWindow(Me.hwnd)
            Me.WindowState = 2
            TakeScreenShot Me, ""
            Me.Show
        Case WM_LBUTTONDBLCLK   'open on left-dblclick
            PopupMenu mnuRestore
        End Select
    End If
    ucAlphaImage2.Visible = False
    ucAlphaImage4.Visible = False
    ucAlphaImage5.Visible = False
End Sub

Private Sub Form_Resize()
    TakeScreenShot Me, ""
    Command1.Visible = True
    Picture1.Visible = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ucAlphaImage5.Visible = False
    ucAlphaImage4.Visible = False
    ucAlphaImage2.ZOrder 0
    ucAlphaImage2.Visible = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ucAlphaImage5.Visible = False
    ucAlphaImage2.Visible = False
    ucAlphaImage4.ZOrder 0
    ucAlphaImage4.Visible = True
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ucAlphaImage2.Visible = False
    ucAlphaImage4.Visible = False
    ucAlphaImage5.ZOrder 0
    ucAlphaImage5.Visible = True
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Picture1.ZOrder 0
    X1 = x
    Y1 = y
    Label2.ZOrder 0
    Command1.ZOrder 0
    AlphaImage1.ZOrder 0
    AlphaImage2.ZOrder 0
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Picture1.Left = Picture1.Left + x - X1
        Picture1.Top = Picture1.Top + y - Y1
        Command1.Left = Command1.Left + x - X1
        Command1.Top = Command1.Top + y - Y1
        AlphaImage1.Left = AlphaImage1.Left + x - X1
        AlphaImage1.Top = AlphaImage1.Top + y - Y1
        AlphaImage2.Left = AlphaImage2.Left + x - X1
        AlphaImage2.Top = AlphaImage2.Top + y - Y1
        Image1.Left = Image1.Left + x - X1
        Image1.Top = Image1.Top + y - Y1
        Image2.Left = Image2.Left + x - X1
        Image2.Top = Image2.Top + y - Y1
        Image3.Left = Image3.Left + x - X1
        Image3.Top = Image3.Top + y - Y1
        ucAlphaImage2.Left = ucAlphaImage2.Left + x - X1
        ucAlphaImage2.Top = ucAlphaImage2.Top + y - Y1
        ucAlphaImage4.Left = ucAlphaImage4.Left + x - X1
        ucAlphaImage4.Top = ucAlphaImage4.Top + y - Y1
        ucAlphaImage5.Left = ucAlphaImage5.Left + x - X1
        ucAlphaImage5.Top = ucAlphaImage5.Top + y - Y1
        Label2.Left = Label2.Left + x - X1
        Label2.Top = Label2.Top + y - Y1
    End If
        ucAlphaImage2.Visible = False
        ucAlphaImage4.Visible = False
        ucAlphaImage5.Visible = False
        Label2.ZOrder 0
        Command1.ZOrder 0
        AlphaImage1.ZOrder 0
        AlphaImage2.ZOrder 0
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim a As Integer, B As Integer, A1 As Integer, b1 As Integer
    a = Val(Left(Picture1.Tag, 7))
    B = Val(Right(Picture1.Tag, 7))
    A1 = Picture1.Left
    b1 = Picture1.Top
    If a + 300 > A1 And a - 300 < A1 And B + 300 > b1 And B - 300 < b1 Then
        Picture1.Left = a
        Picture1.Top = B
    End If
    B = 1
    For a = 0 To Sqa * Sqb - 1
        A1 = Val(Left(Picture1.Tag, 7))
        b1 = Val(Right(Picture1.Tag, 7))
        If A1 <> Picture1.Left Or b1 <> Picture1.Top Then
            B = 0
        End If
    Next a
        Image1.ZOrder 0
        Image2.ZOrder 0
        Image3.ZOrder 0
        
        
End Sub


Private Sub ucAlphaImage2_Click(ByVal Button As Integer)
End
End Sub

Private Sub ucAlphaImage4_Click(ByVal Button As Integer)
On Error Resume Next
        Call SystrayOn(Me, "New Skin Vista Aero The mere idea")
        Call SetForegroundWindow(Me.hwnd)
        Me.Hide
        Me.WindowState = 1
End Sub



