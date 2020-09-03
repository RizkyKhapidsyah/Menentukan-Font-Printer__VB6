VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menentukan Font Printer"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2160
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Printer.FontName = "Arial"
    Printer.FontUnderline = False
    Printer.FontBold = False
    Printer.FontItalic = True
    Printer.FontSize = "30"
    Printer.Print "hello"
    'Gunakan perintah EndDoc jika teks ini adalah yang
    'terakhir yang dicetak ke dalam kertas.
    Printer.EndDoc
End Sub

