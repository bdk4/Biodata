VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Login"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Enter"
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Username"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "admin" And Text2 = "1234" Then
Form1.Show 'Perintah Menampilkan Form 2
Form1.Jurusan.Enabled = True
Form2.Visible = False 'Menyembunyikan Form 1
Form1.Login.Enabled = False
Form1.Logout.Enabled = True
Form1.Command1.Enabled = True
Form1.Command2.Enabled = False
Form1.Command3.Enabled = True
Form1.Command5.Enabled = False
Form1.DataGrid1.Enabled = True
Unload Me 'Menutup Form 1
Else
MsgBox "User Name atau Password yang Anda Masukkan salah" _
& vbNewLine & "Silahkan Coba lagi !!", vbInformation, "NAH LOH"
Text1 = ""
Text2 = ""
Text1.SetFocus
End If
End Sub
