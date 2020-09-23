VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTake 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tutorial v1.0.0 (Take)"
   ClientHeight    =   4815
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6885
   Icon            =   "frmTake.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6885
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmm 
      Left            =   3240
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   4095
      TabIndex        =   6
      Top             =   480
      Width           =   4095
      Begin VB.TextBox txtText 
         Height          =   4215
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   0
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   2055
      Left            =   4320
      TabIndex        =   0
      Top             =   2640
      Width           =   2535
      Begin VB.CommandButton cmdEnlarge 
         Caption         =   "Enlarge Pic"
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<< Back"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >>"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Pic:"
      Height          =   255
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   2535
   End
   Begin VB.Image imgPic 
      Height          =   2055
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   480
      Width           =   2535
   End
   Begin VB.Menu FI 
      Caption         =   "File"
      Begin VB.Menu OP 
         Caption         =   "Open"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MA 
         Caption         =   "Make"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu EX 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTake"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim on_page As Long
Dim pic(99) As String, text(99) As String, xx() As String, pic2() As String, text2() As String

Private Sub cmdBack_Click()
If on_page = 1 Then
    Me.Caption = "Tutorial v1.0.0 (Take) - Error lowest number page alowed is 1!"
    Exit Sub
End If
on_page = on_page - 1
imgPic.Picture = LoadPicture(pic(on_page - 1))
txtText.text = text(on_page - 1)
End Sub

Private Sub cmdEnlarge_Click()
frmEnl.go imgPic.Picture
End Sub

Private Sub cmdNext_Click()
If on_page = 100 Then
    Me.Caption = "Tutorial v1.0.0 (Take) - Error only 100 pages alowed!"
    Exit Sub
End If
on_page = on_page + 1
imgPic.Picture = LoadPicture(pic(on_page - 1))
txtText.text = text(on_page - 1)
End Sub

Private Sub Form_Load()
on_page = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMake
Unload Me
End
End Sub

Private Sub MA_Click()
Me.Hide
frmMake.Show
End Sub

Private Sub OP_Click()
cmm.Filter = "Tutorial Files (*.tut)|*.tut"
cmm.ShowOpen
Open cmm.FileName For Input As #1
    Line Input #1, a
    xx = Split(a, "~")
    pic2 = Split(xx(0), "è")
    text2 = Split(xx(1), "è")
Close #1
For x = 0 To 99
    pic(x) = pic2(x)
    If Not FileExist(pic(x)) Then
        pic(x) = App.Path & "\" & GetFileNameFromPath(pic(x))
    End If
    text(x) = text2(x)
Next x
on_page = 1
cmdNext_Click
cmdBack_Click
Me.Caption = "Tutorial v1.0.0 (Take)"
End Sub


