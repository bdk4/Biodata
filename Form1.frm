VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Biodata"
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   8400
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Update"
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Batal"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cari Foto"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   1575
      Left            =   480
      TabIndex        =   11
      Top             =   6720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2778
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ubah"
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1920
      MaxLength       =   12
      TabIndex        =   7
      Top             =   5520
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   1920
      MaxLength       =   35
      TabIndex        =   6
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1920
      MaxLength       =   25
      TabIndex        =   5
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3600
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4440
      Top             =   7920
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Bayu\pemrograman1\biodata2\Database1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Bayu\pemrograman1\biodata2\Database1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Table1"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3135
      Left            =   600
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "No. Hp"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Alamat"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "NIM"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Menu Login 
      Caption         =   "Login"
   End
   Begin VB.Menu Jurusan 
      Caption         =   "Jurusan"
   End
   Begin VB.Menu Logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text1.SetFocus
Command1.Enabled = False
Command2.Enabled = True
Command3.Enabled = False
Command5.Enabled = True
Command6.Enabled = True
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Then
MsgBox "Data harus di isi!"
Else
    Adodc1.Recordset.Find "NIM='" + Text1.Text + "'", , adSearchForward, 1
    If Not Adodc1.Recordset.EOF Then
     MsgBox "Maaf, NIM sudah ada!"
     Text1.Text = ""
     Text1.SetFocus
    Else
    Adodc1.Recordset.AddNew
    Adodc1.Recordset!NIM = Text1.Text
    Adodc1.Recordset!nama = Text2.Text
    Adodc1.Recordset!Alamat = Text3.Text
    Adodc1.Recordset!No_Hp = Text4.Text
    Adodc1.Recordset.Update
    ' code berikut berfungsi untuk menyimpan gambar ke dalam folder foto
    ' nama file gambar didepannya ada kata NIP, contoh: nama foto = NIP_12
    SavePicture Image1.Picture, App.Path & "\foto\NIM_" & Text1.Text & ".jpg"
    'membuat laporan
    MsgBox "Data & Foto telah di Simpan!", vbInformation + vbOKOnly = vbIgnore
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Image1.Picture = LoadPicture()
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Visible = False
    Command5.Enabled = False
    Command6.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Text4.Enabled = False
    DataGrid1.Enabled = True
    Command2.Enabled = False
    Command6.Enabled = False
    End If
End If
End Sub
Private Sub Command3_Click()
If Text1.Text = "" Then
MsgBox "Pilih Data yang Akan Diubah!", vbInformation + vbOKOnly = vbIgnore
Else
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Visible = True
Command5.Enabled = True
Command6.Enabled = True
Text1.SetFocus
End If
End Sub

Private Sub Command4_Click()
    Adodc1.Recordset.Update
    Adodc1.Recordset!NIM = Text1.Text
    Adodc1.Recordset!nama = Text2.Text
    Adodc1.Recordset!Alamat = Text3.Text
    Adodc1.Recordset!No_Hp = Text4.Text
Adodc1.Recordset.Update
DataGrid1.Refresh
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command4.Visible = True
Command5.Enabled = False
Command6.Enabled = False
Image1.Picture = LoadPicture()
MsgBox "Data & Foto telah di Update!", vbInformation + vbOKOnly = vbIgnore
End Sub

Private Sub Command5_Click()
With CommonDialog1
    .FileName = ""
    .Filter = "Image (*.jpg)|*.jpg"
    .ShowOpen
        Image1.Picture = LoadPicture(.FileName)
End With
End Sub
Private Sub Command6_Click()
Command1.Enabled = True
Command2.Enabled = False
Command3.Enabled = True
Command4.Visible = False
Command5.Enabled = False
Command6.Enabled = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Image1.Picture = LoadPicture()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

Private Sub DataGrid1_Click()
If Adodc1.Recordset.RecordCount <= 0 Then Exit Sub
    If Not Adodc1.Recordset.BOF And Not Adodc1.Recordset.EOF Then
        Text1.Text = Adodc1.Recordset.Fields("NIM")
        Text2.Text = Adodc1.Recordset.Fields("Nama")
        Text3.Text = Adodc1.Recordset.Fields("Alamat")
        Text4.Text = Adodc1.Recordset.Fields("No_Hp")
        Image1.Picture = LoadPicture(App.Path & "\foto\NIM_" & Text1.Text & ".jpg")
        Command3.Enabled = True
        
    End If
End Sub

Private Sub DataGrid1_DblClick()
Dim hapus As String
Dim a
hapus = DataGrid1.Columns(0).Text
a = MsgBox("Hapus Data...?", vbQuestion + vbYesNo)
If a = vbYes Then
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveFirst
DataGrid1.ReBind
DataGrid1.Refresh
MsgBox "Data Berhasil Di Hapus", vbInformation
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Image1.Picture = LoadPicture()
Else
Exit Sub
End If
End Sub

Private Sub Form_Load()
Form1.Jurusan.Enabled = False
Form1.Logout.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Visible = False
Command5.Enabled = False
Command6.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
DataGrid1.Enabled = False
End Sub

Private Sub Jurusan_Click()
Form3.Show
Form1.Hide

End Sub

Private Sub Login_Click()
Form2.Show
End Sub

Private Sub Logout_Click()
Form1.Jurusan.Enabled = False
Form1.Logout.Enabled = False
Form1.Login.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command5.Enabled = False
Command6.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
DataGrid1.Enabled = False
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Image1.Picture = LoadPicture()
End Sub
