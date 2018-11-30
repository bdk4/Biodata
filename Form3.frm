VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Jurusan"
   ClientHeight    =   6135
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   5565
   LinkTopic       =   "Form3"
   ScaleHeight     =   6135
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option3 
      Caption         =   "SK"
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   3720
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Caption         =   "TI"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3720
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SI"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Text            =   "Jurusan"
      Top             =   4320
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   1800
      TabIndex        =   5
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label5 
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "Nama Kepala Jurusan :"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Konsentrasi"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Jurusan"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Id Jurusan"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Menu Login 
      Caption         =   "Login"
   End
   Begin VB.Menu Biodata 
      Caption         =   "Biodata"
   End
   Begin VB.Menu Logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Biodata_Click()
Form1.Show
Form3.Hide
End Sub

Private Sub Logout_Click()
Form1.Show
Form3.Hide
Form1.Jurusan.Enabled = False
Form2.Visible = False 'Menyembunyikan Form 1
Form1.Login.Enabled = True
Form1.Logout.Enabled = False
Form1.Command1.Enabled = False
Form1.Command2.Enabled = False
Form1.Command3.Enabled = False
Form1.Command5.Enabled = False
Form1.DataGrid1.Enabled = False
End Sub

Private Sub Option1_Click()
Option1.Value = True
Text1.Text = "Sistem Informasi"
List1.Clear
List1.AddItem ("Business Intelligence")
List1.AddItem ("Computer Accountancy")
List1.AddItem ("Management Information System")
Image1.Picture = LoadPicture(App.Path & "\foto\kajur1" & ".jpg")
Label5.Caption = "Nur Azizah, M.Akt.,M.Kom"
End Sub
Private Sub Option2_Click()
Option2.Value = True
Text1.Text = "Teknik Informasi"
List1.Clear
List1.AddItem ("MAVIB")
List1.AddItem ("Software Engineering")
Image1.Picture = LoadPicture(App.Path & "\foto\kajur2" & ".jpg")
Label5.Caption = "Junaidi, M.Kom"
End Sub

Private Sub Option3_Click()
Option3.Value = True
List1.Clear
Text1.Text = "Pilih Jurusan"
List1.AddItem ("CCIT")
List1.AddItem ("Computer System")
Image1.Picture = LoadPicture(App.Path & "\foto\kajur3" & ".jpg")
Label5.Caption = "Ferry Sudarto, S.Kom,M.Pd"
End Sub
Private Sub Form_Load()
Text1.Enabled = False
Login.Enabled = False
End Sub
