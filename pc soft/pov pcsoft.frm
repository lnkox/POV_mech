VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton pic_to_wpic 
      Caption         =   "^^^"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0C000&
      Height          =   330
      Left            =   240
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   4
      Top             =   1320
      Width           =   3780
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      Height          =   495
      Left            =   10200
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "build"
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox code_out 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0C000&
      Height          =   630
      Left            =   240
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   503
      TabIndex        =   0
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
code_out.Text = ""
a = 0
b = 0
c = 0
i = 0
For X = 1 To 500 Step 2
t = 0
i = i + 1
For Y = 0 To 7
If Picture1.Point(X, 1 + Y * 2) <> &HFFFFFF Then t = t + 2 ^ Y
z = Picture1.Point(X, Y * 2)
Next
a = a & "," & Trim(Str(t))
t = 0
For Y = 0 To 7
If Picture1.Point(X, 17 + Y * 2) <> &HFFFFFF Then t = t + 2 ^ Y
Next
b = b & "," & Trim(Str(t))
t = 0
For Y = 0 To 3
If Picture1.Point(X, 33 + Y * 2) <> &HFFFFFF Then t = t + 2 ^ Y
Next
c = c & "," & Trim(Str(t))
t = 0

Next
code_out.Text = code_out.Text & "flash char a[]={" & a & "};" & vbNewLine
code_out.Text = code_out.Text & "flash char b[]={" & b & "};" & vbNewLine
code_out.Text = code_out.Text & "flash char c[]={" & c & "};" & vbNewLine
End Sub

Private Sub Form_Load()
Picture1.ForeColor = &HFFFFFF
For X = 1 To 720
For Y = 1 To 40
Picture1.PSet (X, Y)
Next
Next
End Sub

Private Sub pic_to_wpic_Click()
Picture1.ForeColor = &HFFFFFF
For X = 1 To 720
For Y = 1 To 40
Picture1.PSet (X, Y)
Next
Next
For X = 0 To 250
For Y = 0 To 20
If Picture2.Point(X, Y) <> &HFFFFFF Then
Yw = Y * 2
Xw = X * 2
If Yw > 20 Then Picture1.ForeColor = &HFFFF& Else Picture1.ForeColor = &HC0C000
If Yw \ 2 <> Yw / 2 Then Yw = Int(Yw \ 2) * 2
If Xw \ 2 <> Xw / 2 Then Xw = Int(Xw \ 2) * 2
Picture1.PSet (Xw, Yw)
Picture1.PSet (Xw + 1, Yw)
Picture1.PSet (Xw, Yw + 1)
Picture1.PSet (Xw + 1, Yw + 1)
End If
Next
Next
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture1_MouseMove Button, Shift, X, Y
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
If Y > 20 Then Picture1.ForeColor = &HFFFF& Else Picture1.ForeColor = &HC0C000
If Y \ 2 <> Y / 2 Then Y = Int(Y \ 2) * 2
If X \ 2 <> X / 2 Then X = Int(X \ 2) * 2
Picture1.PSet (X, Y)
Picture1.PSet (X + 1, Y)
Picture1.PSet (X, Y + 1)
Picture1.PSet (X + 1, Y + 1)
End If
If Button = 2 Then
Picture1.ForeColor = &HFFFFFF
If Y \ 2 <> Y / 2 Then Y = Int(Y \ 2) * 2
If X \ 2 <> X / 2 Then X = Int(X \ 2) * 2
Picture1.PSet (X, Y)
Picture1.PSet (X + 1, Y)
Picture1.PSet (X, Y + 1)
Picture1.PSet (X + 1, Y + 1)
End If
End Sub

Private Sub Picture2_DblClick()
On Error Resume Next
CommonDialog1.ShowOpen
tf = CommonDialog1.FileName
Picture2.Picture = LoadPicture(tf)
End Sub
