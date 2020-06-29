VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   LinkTopic       =   "Form3"
   Picture         =   "Formsiswa.frx":0000
   ScaleHeight     =   5955
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtnis 
      Height          =   285
      Left            =   3240
      TabIndex        =   8
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txtnamasiswa 
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox Txtkelas 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox Comjeniskelamin 
      Height          =   315
      Left            =   3240
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   1560
      TabIndex        =   0
      Top             =   3720
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3413
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
            LCID            =   14345
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
            LCID            =   14345
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DATA SISWA"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "NIS"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Siswa"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Kelas"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Jenis Kelamin"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rssiswa As New ADODB.Recordset

Private Sub Command1_Click()
If Txtnis = "" Then
MsgBox "Nis Kosong", vbExclamation, "pesan"
Txtnis.SetFocus
Exit Sub
End If
    If txtnamasiswa = "" Then
    MsgBox "Nama Kosong", vbExclamation, "pesan"
    txtnamasiswa.SetFocus
    Exit Sub
    End If
If Txtkelas = "" Then
MsgBox "kelas Kosong", vbExclamation, "pesan"
Txtkelas.SetFocus
Exit Sub
End If
    If Comjeniskelamin = "" Then
    MsgBox "Jenis kelamin Kosong", vbExclamation, "pesan"
    Comjeniskelamin.SetFocus
    Exit Sub
    End If
Set rssiswa = New ADODB.Recordset
rssiswa.Open "select*from tb_siswa where nis='" & Txtnis & "'", koneksidb
If Not rssiswa.EOF Then
MsgBox "nis sudah ada", vbCritical, "pesan"
Txtnis = ""
Txtnis.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tb_siswa(nis,nmsiswa,kelas,jnskel) value ('" & Txtnis & "','" & txtnamasiswa & "','" & Txtkelas & "','" & Comjeniskelamin & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = rssiswa
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub Command2_Click()
koneksidb.Execute "update tb_siswa set nmsiswa='" & txtnamasiswa.Text & "',kelas='" & Txtkelas & "',jnskel='" & Comjeniskelamin & "' where nis='" & Txtnis & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub Command3_Click()
koneksidb.Execute "delete from tb_siswa where nis='" & Txtnis & "'"
Call refreshh
Call kosong
Txtnis.SetFocus
End Sub

Private Sub Command4_Click()
Call kosong
End Sub

Private Sub DataGrid1_Click()
Txtnis.Text = rssiswa!nis
txtnamasiswa.Text = rssiswa!nmsiswa
Txtkelas.Text = rssiswa!kelas
Comjeniskelamin.Text = rssiswa!jnskel
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rssiswa
With rssiswa
With Comjeniskelamin
    .AddItem "laki-laki"
    .AddItem "perempuan"
End With
Call edit_grid
End With
End Sub

Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "NIS"
    .Columns(1).Caption = "Nama Siswa"
    .Columns(2).Caption = "Kelas"
    .Columns(3).Caption = "jenis kelamin"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
End With
End Sub
Sub tampil_data()
Set rssiswa = New ADODB.Recordset
rssiswa.ActiveConnection = koneksidb
rssiswa.CursorLocation = adUseClient
rssiswa.LockType = adLockOptimistic
rssiswa.Source = "select * from tb_siswa"
rssiswa.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rssiswa
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rssiswa
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
Txtkelas = ""
txtnamasiswa = ""
Txtnis = ""
Comjeniskelamin = ""
Txtnis.SetFocus
End Sub


