VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11190
   LinkTopic       =   "Form6"
   Picture         =   "Formpenilaian.frx":0000
   ScaleHeight     =   7260
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   8400
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   6000
      TabIndex        =   20
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   3600
      TabIndex        =   19
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   1200
      TabIndex        =   18
      Top             =   4440
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   240
      TabIndex        =   17
      Top             =   5040
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   2990
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
   Begin VB.TextBox txtnilai 
      Height          =   375
      Left            =   5760
      TabIndex        =   16
      Top             =   3840
      Width           =   2895
   End
   Begin VB.TextBox txtkriteria 
      Height          =   405
      Left            =   5760
      TabIndex        =   15
      Top             =   3360
      Width           =   2895
   End
   Begin VB.TextBox txtkodekriteria 
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox txtupenilaian 
      Height          =   405
      Left            =   5760
      TabIndex        =   13
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtnamaguru 
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtnip 
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtnamasiswa 
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox Txtnis 
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Nilai"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Kriteria"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Kodekriteria"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Unsur Penilaian"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Nama Guru"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "NIP Guru"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Siswa"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "NIS"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DATA PENILAIAN"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rspenilaian As New ADODB.Recordset

Private Sub Command1_Click()
If Txtnis = "" Then
MsgBox "Nis Kosong", vbExclamation, "pesan"
Txtnis.SetFocus
Exit Sub
End If
    If txtnamasiswa = "" Then
    MsgBox "Nama Siswa Kosong", vbExclamation, "pesan"
    txtnamasiswa.SetFocus
    Exit Sub
    End If
If txtnip = "" Then
MsgBox "Nip Kosong", vbExclamation, "pesan"
txtnip.SetFocus
Exit Sub
End If
    If txtnamaguru = "" Then
    MsgBox "Nama Guru Kosong", vbExclamation, "pesan"
    txtnamaguru.SetFocus
    Exit Sub
    End If
If txtupenilaian = "" Then
MsgBox "Unsur Penilaian Kosong", vbExclamation, "pesan"
txtupenilaian.SetFocus
Exit Sub
End If
    If txtkodekriteria = "" Then
    MsgBox "Kode Kriteria Kosong", vbExclamation, "pesan"
    txtkodekriteria.SetFocus
    Exit Sub
    End If
If txtkriteria = "" Then
MsgBox "Kriteria Kosong", vbExclamation, "pesan"
txtkriteria.SetFocus
Exit Sub
End If
    If txtnilai = "" Then
    MsgBox "Nilai Kosong", vbExclamation, "pesan"
    txtnilai.SetFocus
    Exit Sub
    End If
Set rspenilaian = New ADODB.Recordset
rspenilaian.Open "select*from tb_penilaian where kdkriteria='" & txtkodekriteria & "'", koneksidb
If Not rspenilaian.EOF Then
MsgBox "Kode Kriteria sudah ada", vbCritical, "pesan"
txtkodekriteria = ""
txtkodekriteria.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tb_penilaian(nis,nmsiswa,nip,nmguru,unsurpenilaian,kdkriteria,nmkriteria,nilai) value ('" & Txtnis & "','" & txtnamasiswa & "','" & txtnip & "','" & txtnamaguru & "','" & txtupenilaian & "','" & txtkodekriteria & "','" & txtkriteria & "','" & txtnilai & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = rspenilaian
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub Command2_Click()
koneksidb.Execute "update tb_penilaian set nis='" & Txtnis.Text & "',nmsiswa='" & txtnamasiswa & "',nip='" & txtnip & "',nmguru='" & txtnamaguru & "',unsurpenilaian='" & txtupenilaian & "',nmkriteria='" & txtkriteria & "',nilai='" & txtnilai & "' where kdkriteria='" & txtkodekriteria & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub Command3_Click()
koneksidb.Execute "delete from tb_penilaian where kdkriteria='" & txtkodekriteria & "'"
Call refreshh
Call kosong
txtnip.SetFocus
End Sub

Private Sub Command4_Click()
Call kosong
End Sub

Private Sub DataGrid1_Click()
Txtnis.Text = rspenilaian!nis
txtnamasiswa.Text = rspenilaian!nmsiswa
txtnip.Text = rspenilaian!nip
txtnamaguru.Text = rspenilaian!nmguru
txtupenilaian.Text = rspenilaian!unsurpenilaian
txtkodekriteria.Text = rspenilaian!kdkriteria
txtkriteria.Text = rspenilaian!nmkriteria
txtnilai.Text = rspenilaian!nilai
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rspenilaian
With rspenilaian
End With
Call edit_grid
End Sub

Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "NIS"
    .Columns(1).Caption = "Nama Siswa"
    .Columns(2).Caption = "NIP Guru"
    .Columns(3).Caption = "Nama Guru"
    .Columns(4).Caption = "Unsur Penilaian"
    .Columns(5).Caption = "Kode Kriteria"
    .Columns(6).Caption = "Kriteria"
    .Columns(7).Caption = "Nilai"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
    .Columns(4).Width = 1200
    .Columns(5).Width = 1200
    .Columns(6).Width = 1200
    .Columns(7).Width = 1200
End With
End Sub

Sub tampil_data()
Set rspenilaian = New ADODB.Recordset
rspenilaian.ActiveConnection = koneksidb
rspenilaian.CursorLocation = adUseClient
rspenilaian.LockType = adLockOptimistic
rspenilaian.Source = "select * from tb_penilaian"
rspenilaian.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rspenilaian
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rspenilaian
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
Txtnis = ""
txtnamasiswa = ""
txtnip = ""
txtnamaguru = ""
txtupenilaian = ""
txtkodekriteria = ""
txtkriteria = ""
txtnilai = ""
txtkodekriteria.SetFocus
End Sub

