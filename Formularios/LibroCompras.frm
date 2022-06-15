VERSION 5.00
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Begin VB.Form FrmLibCompra 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Libro de Compras"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4575
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin vbskfr2.Skinner Skinner1 
      Left            =   720
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   465
      TabIndex        =   0
      Top             =   330
      Width           =   3645
      Begin VB.CommandButton CmdImpImpresora 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   1830
         MaskColor       =   &H00E0E0E0&
         Picture         =   "LibroCompras.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Imprime en la impresora las retenciones del Proveedor"
         Top             =   1080
         Width           =   1260
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "LibroCompras.frx":0857
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   510
         MaskColor       =   &H00E0E0E0&
         Picture         =   "LibroCompras.frx":1099
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "1"
         Top             =   1080
         Width           =   1260
      End
      Begin VB.ComboBox Cboano 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1710
         TabIndex        =   2
         Text            =   "Año"
         Top             =   420
         Width           =   1695
      End
      Begin VB.ComboBox CboMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   450
         TabIndex        =   1
         Text            =   "Mes"
         Top             =   420
         Width           =   1245
      End
   End
End
Attribute VB_Name = "FrmLibCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command6_Click()
On Error Resume Next
AnoFiscaLibComp = Cboano.Text
MesFiscaLibComp = CboMes.Text
TipoDeConsulta = "LibroCompras"
FmReporte.Show 1, Me
End Sub

Private Sub Form_Load()
Set Skinner1.Forms = Forms
TXTSQL = "select distinct (MesPeriodoFiscal) from Comprobante"
    Set REC = DB.Execute(TXTSQL)
    CboMes.Clear
    Do Until REC.EOF
        CboMes.AddItem REC.Fields("MesPeriodoFiscal")

        REC.MoveNext
    Loop


TXTSQL = "select distinct (AnoPeriodoFiscal) from Comprobante"
    Set REC = DB.Execute(TXTSQL)
    Cboano.Clear
    Do Until REC.EOF
        Cboano.AddItem REC.Fields("AnoPeriodoFiscal")

        REC.MoveNext
    Loop
    
    CboMes.Text = Format(Date, "mm")
    Cboano.Text = Format(Date, "yyyy")

End Sub
