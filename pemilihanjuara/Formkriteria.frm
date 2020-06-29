VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11010
   LinkTopic       =   "Form5"
   Picture         =   "Formkriteria.frx":0000
   ScaleHeight     =   5190
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "BATAL"
      Height          =   495
      Left            =   8640
      TabIndex        =   9
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
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
   Begin VB.TextBox txtkriteria 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox txtkodekriteria 
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Kriteria"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Kode Kriteria"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "DATA KRITERIA"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rskriteria As New ADODB.Recordset

Private Sub Command1_Click()
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
Set kriteria = New ADODB.Recordset
kriteria.Open "select*from tb_kriteria where kdkriteria='" & txtkodekriteria & "'", koneksidb
If Not kriteria.EOF Then
MsgBox "Kode Kriteria sudah ada", vbCritical, "pesan"
txtkodekriteria = ""
txtkodekriteria.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tb_kriteria(kdkriteria,nmkriteria) value ('" & txtkodekriteria & "','" & txtkriteria & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = rskriteria
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub Command2_Click()
koneksidb.Execute "update tb_kriteria set nmkriteria='" & txtkriteria.Text & "' where kdkriteria='" & txtkodekriteria & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub Command3_Click()
koneksidb.Execute "delete from tb_kriteria where kdkriteria='" & txtkodekriteria & "'"
Call refreshh
Call kosong
txtkodekriteria.SetFocus
End Sub

Private Sub Command4_Click()
Call kosong
End Sub

Private Sub DataGrid1_Click()
txtkodekriteria.Text = rskriteria!kdkriteria
txtkriteria.Text = rskriteria!nmkriteria
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rskriteria
With rskriteria
End With
Call edit_grid
End Sub

Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "Kode Kriteria"
    .Columns(1).Caption = "Kriteria"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
End With
End Sub

Sub tampil_data()
Set rskriteria = New ADODB.Recordset
rskriteria.ActiveConnection = koneksidb
rskriteria.CursorLocation = adUseClient
rskriteria.LockType = adLockOptimistic
rskriteria.Source = "select * from tb_kriteria"
rskriteria.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rskriteria
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = rskriteria
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
txtkriteria = ""
txtkodekriteria = ""
txtkodekriteria.SetFocus
End Sub


