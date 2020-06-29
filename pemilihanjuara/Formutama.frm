VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   Picture         =   "Formutama.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Data 
      Caption         =   "DATA"
      Begin VB.Menu Datasiswa 
         Caption         =   "Data Siswa"
      End
      Begin VB.Menu Dataguru 
         Caption         =   "Data Guru"
      End
      Begin VB.Menu Datakriteria 
         Caption         =   "Data Kriteria"
      End
   End
   Begin VB.Menu Penilaian 
      Caption         =   "PENILAIAN"
   End
   Begin VB.Menu Laporan 
      Caption         =   "LAPORAN"
      Begin VB.Menu Hasiljuara 
         Caption         =   "Hasil Juara"
      End
   End
   Begin VB.Menu KELUAR 
      Caption         =   "KELUAR"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
