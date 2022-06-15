VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menu"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   12960
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnuRetencionIVa 
      Caption         =   "Retencio IVA"
      Begin VB.Menu MnuProcesarFacturaIva 
         Caption         =   "Registar Facturas"
      End
   End
   Begin VB.Menu MnuISLR 
      Caption         =   "ISLR"
      Begin VB.Menu MnuProcesarISLR 
         Caption         =   "Registar ISLR"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub

Private Sub MnuRetencionIVa_Click()

End Sub
