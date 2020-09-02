VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1455
   ScaleWidth      =   3000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox t1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Nama File disimpan pada drive c."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    SavePicture Form1.Picture1.Image, App.Path + "\hasil_simpan\file" & ".bmp"
    MsgBox "Data tersebut telah tersimpan.", vbInformation, "Pemberitahuan"
    Unload Me
End Sub
