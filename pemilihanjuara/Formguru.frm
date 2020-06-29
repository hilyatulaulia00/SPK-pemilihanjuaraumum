VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11565
   LinkTopic       =   "Form4"
   Picture         =   "Formguru.frx":0000
   ScaleHeight     =   6195
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Txtjabatan 
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Txtnamaguru 
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox txtnip 
      Height          =   405
      Left            =   4560
      TabIndex        =   9
      Top             =   600
      Width           =   3495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   2040
      TabIndex        =   8
      Top             =   3840
      Width           =   6255
      _ExtentX        =   11033
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
   Begin VB.CommandButton Command4 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Jabatan"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Guru"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "NIP"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DATA GURU"
      Height          =   255
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsguru As New ADODB.Recordset


Private Sub Command1_Click()
If txtnip = "" Then
MsgBox "Nip Kosong", vbExclamation, "pesan"
txtnip.SetFocus
Exit Sub
End If
    If Txtnamaguru = "" Then
    MsgBox "Nama Kosong", vbExclamation, "pesan"
    txtnamasiswa.SetFocus
    Exit Sub
    End If
If Txtjabatan = "" Then
MsgBox "Jabatan Kosong", vbExclamation, "pesan"
Txtjabatan.SetFocus
Exit Sub
End If
Set guru = New ADODB.Recordset
guru.Open "select*from tb_guru where nip='" & txtnip & "'", koneksidb
If Not guru.EOF Then
MsgBox "nip sudah ada", vbCritical, "pesan"
txtnip = ""
txtnip.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tb_guru(nip,nmguru,jabatan) value ('" & txtnip & "','" & Txtnamaguru & "','" & Txtjabatan & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = rsguru
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub Command2_Click()
koneksidb.Execute "update tb_guru set nmguru='" & Txtnamaguru.Text & "',jabatan='" & Txtjabatan & "' where nip='" & txtnip & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub Command3_Click()
koneksidb.Execute "delete from tb_guru where nip='" & txtnip & "'"
Call refreshh
Call kosong
txtnip.SetFocus
End Sub

Private Sub Command4_Click()
Call kosong
End Sub

Private Sub DataGrid1_Click()
txtnip.Text = rsguru!nip
Txtnamaguru.Text = rsguru!nmguru
Txtjabatan.Text = rsguru!jabatan
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rsguru
With rsguru
End With
Call edit_grid
End Sub

Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "NIP"
    .Columns(1).Caption = "Nama Guru"
    .Columns(2).Caption = "Jabatan"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
End With
End Sub

Sub tampil_data()
Set rsguru = New ADODB.Recordset
rsguru.ActiveConnection = koneksidb
rsguru.CursorLocation = adUseClient
rsguru.LockType = adLockOptimistic
rsguru.Source = "select * from tb_guru"
rsguru.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rsguru
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rsguru
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
Txtjabatan = ""
Txtnamaguru = ""
txtnip = ""
txtnip.SetFocus
End Sub

