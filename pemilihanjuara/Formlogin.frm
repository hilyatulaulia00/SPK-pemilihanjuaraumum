VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10965
   LinkTopic       =   "Form2"
   Picture         =   "Formlogin.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtnamaadm 
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   1800
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000000&
      Caption         =   "LOG-IN"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000000&
      Caption         =   "KELUAR"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000002&
      Caption         =   "SISTEM PENDUKUNG KEPUTUSAN PEMILIHAN SISWA JUARA UMUM KENAIKAN KELAS DI SMP NEGERI 03 KOTA BENGKULU"
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000002&
      Caption         =   "NAMA ADMIN"
      ForeColor       =   &H80000001&
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000002&
      Caption         =   "PASSWORD"
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1800
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rslogin As New ADODB.Recordset

Private Sub Command1_Click()
nmadmin = txtnamaadm.Text
Password = txtpass.Text
If nmadmin = "Aulia" And Password = "01703042" Then
MsgBox "Selamat Datang"
Form1.Show
Form2.Hide
Else
login = login + 1
MsgBox "Anda salah memasukkan password" & login & " kali "
If login = 2 Then
MsgBox "Kesempatan Anda satu kali lagi", vbExclamation
End If
If login = 3 Then
MsgBox "anda sudah salah memasukkan password 3 kali,maka program kali ini akan kami tutup!"
End
End If
End If
End Sub

Private Sub Command2_Click()
X = MsgBox("Yakin Keluar?", vbQuestion + vbYesNo, "informasi")
If X = vbYes Then End
End Sub

