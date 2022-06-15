VERSION 5.00
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Begin VB.Form PopUpEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7500
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbskfr2.Skinner Skinner1 
      Left            =   630
      Top             =   0
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
      SysEnableSkinCaption=   "Habilitar &Skin"
      SysDisableSkinCaption=   "Deshabilitar &Skin"
   End
   Begin VB.CommandButton Command2 
      Height          =   945
      Left            =   6210
      Picture         =   "PopUpEditor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   990
   End
   Begin VB.TextBox txteditor 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   5955
   End
End
Attribute VB_Name = "PopUpEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command2_Click()
Variablepopupeditor = txteditor.Text
Unload Me
End Sub

Private Sub Form_Load()
Set Skinner1.Forms = Forms
PopUpEditor.Caption = TXTSQL
 If Variablepopupeditor <> "" Then
    txteditor.Width = (Len(Variablepopupeditor) * 250) + 800
    PopUpEditor.Width = txteditor.Width + 2000
    Command2.Left = txteditor.Width + txteditor.Left + 30
    txteditor.Text = Variablepopupeditor
End If
End Sub

Private Sub txteditor_GotFocus()
    With txteditor
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txteditor_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Dim FUNCIONES As Integer
If KeyAscii = 44 Or KeyAscii = 46 Then
    FUNCIONES = DECIMALES(KeyAscii)
End If
If FUNCIONES <> 0 Then
    KeyAscii = FUNCIONES
End If
If KeyAscii = 13 Then
Variablepopupeditor = txteditor.Text
Unload Me
End If
End Sub
