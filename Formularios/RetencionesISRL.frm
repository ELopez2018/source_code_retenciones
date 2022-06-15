VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form RetencionesISLR 
   Appearance      =   0  'Flat
   BackColor       =   &H00C00000&
   Caption         =   "Retenciones al ISLR"
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   10680
   Icon            =   "RetencionesISRL.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   10950
      TabIndex        =   42
      Top             =   2400
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C00000&
      Height          =   10065
      Left            =   308
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      Begin VB.Frame Frame3 
         BackColor       =   &H00C00000&
         Caption         =   "Datos de la Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   180
         TabIndex        =   2
         Top             =   4440
         Width           =   9465
         Begin VB.TextBox txtIva 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7230
            TabIndex        =   50
            Text            =   "0"
            Top             =   990
            Width           =   2115
         End
         Begin VB.TextBox TxtTotalaPagar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7230
            TabIndex        =   49
            Text            =   "0"
            Top             =   1890
            Width           =   2115
         End
         Begin VB.TextBox TxtMontoRetencion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7230
            TabIndex        =   48
            Text            =   "0"
            Top             =   1440
            Width           =   2115
         End
         Begin VB.TextBox TxtBaseImponible 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3690
            TabIndex        =   45
            Text            =   "0"
            Top             =   990
            Width           =   2115
         End
         Begin VB.TextBox txtExento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3690
            TabIndex        =   43
            Text            =   "0"
            Top             =   1440
            Width           =   2115
         End
         Begin VB.TextBox txtMontoTotalFactura 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3690
            TabIndex        =   10
            Text            =   "0"
            Top             =   1890
            Width           =   2115
         End
         Begin VB.TextBox txtNoControl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3615
            TabIndex        =   9
            Top             =   300
            Width           =   2265
         End
         Begin VB.TextBox txtNoFactura 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   7080
            TabIndex        =   8
            Top             =   300
            Width           =   2265
         End
         Begin VB.TextBox TxtFechafa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   150
            TabIndex        =   7
            Top             =   300
            Width           =   2265
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Imponible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2310
            TabIndex        =   46
            Top             =   1110
            Width           =   1305
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   3015
            TabIndex        =   44
            Top             =   1560
            Width           =   600
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "La  Factura Existe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   300
            TabIndex        =   39
            Top             =   1590
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retencion Bs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6000
            TabIndex        =   38
            Top             =   1553
            Width           =   1155
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total a Pagar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   5985
            TabIndex        =   19
            Top             =   2003
            Width           =   1170
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Iva"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6870
            TabIndex        =   18
            Top             =   1103
            Width           =   285
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Total"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2580
            TabIndex        =   17
            Top             =   2010
            Width           =   1035
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Control"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4290
            TabIndex        =   16
            Top             =   750
            Width           =   915
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No de Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   7597
            TabIndex        =   15
            Top             =   750
            Width           =   1230
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1012
            TabIndex        =   14
            Top             =   750
            Width           =   540
         End
      End
      Begin VB.CommandButton CmdExportPDF 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "RetencionesISRL.frx":0CCA
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   5355
         MaskColor       =   &H00E0E0E0&
         Picture         =   "RetencionesISRL.frx":1E65
         Style           =   1  'Graphical
         TabIndex        =   41
         Tag             =   "1"
         ToolTipText     =   "Exportar a PDF"
         Top             =   7770
         Width           =   1260
      End
      Begin VB.CommandButton CmdNuevo 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "RetencionesISRL.frx":2525
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   2775
         MaskColor       =   &H00E0E0E0&
         Picture         =   "RetencionesISRL.frx":36C0
         Style           =   1  'Graphical
         TabIndex        =   40
         Tag             =   "1"
         ToolTipText     =   "Nuevo"
         Top             =   7770
         Width           =   1260
      End
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
         Height          =   1000
         Left            =   6645
         MaskColor       =   &H00E0E0E0&
         Picture         =   "RetencionesISRL.frx":3C79
         Style           =   1  'Graphical
         TabIndex        =   37
         Tag             =   "1"
         ToolTipText     =   "Imprime en la impresora las retenciones del Proveedor"
         Top             =   7770
         Width           =   1260
      End
      Begin MSFlexGridLib.MSFlexGrid MsFProveedores 
         Height          =   555
         Left            =   450
         TabIndex        =   25
         Top             =   900
         Visible         =   0   'False
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   979
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16711680
         ForeColorFixed  =   16777215
         BackColorSel    =   16777215
         ForeColorSel    =   16711680
         BackColorBkg    =   16777215
         GridColorFixed  =   14707546
         FocusRect       =   0
         GridLines       =   0
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Cmdimpr 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "RetencionesISRL.frx":44D0
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   4080
         MaskColor       =   &H00E0E0E0&
         Picture         =   "RetencionesISRL.frx":4D12
         Style           =   1  'Graphical
         TabIndex        =   24
         Tag             =   "1"
         ToolTipText     =   "Guardar"
         Top             =   7770
         Width           =   1260
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C00000&
         Caption         =   "Datos del Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2145
         Left            =   180
         TabIndex        =   3
         Top             =   2295
         Width           =   9465
         Begin VB.TextBox txtCodXml 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1950
            TabIndex        =   33
            Top             =   1080
            Width           =   795
         End
         Begin VB.TextBox TxtPorcentaRetenc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   1155
            TabIndex        =   32
            Top             =   1080
            Width           =   795
         End
         Begin VB.TextBox txtDescripcion 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   2745
            MultiLine       =   -1  'True
            TabIndex        =   31
            Top             =   1080
            Width           =   6615
         End
         Begin VB.TextBox txtfechaCompr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   30
            Top             =   330
            Width           =   2745
         End
         Begin VB.TextBox txtNumeCompr 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3400
            TabIndex        =   29
            Top             =   322
            Width           =   2745
         End
         Begin VB.TextBox txtMes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   6560
            TabIndex        =   28
            Top             =   322
            Width           =   855
         End
         Begin VB.TextBox txtAno 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   7800
            TabIndex        =   27
            Top             =   322
            Width           =   1575
         End
         Begin VB.ComboBox CboCodCta 
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
            Left            =   240
            TabIndex        =   26
            Text            =   "Combo1"
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   5580
            TabIndex        =   36
            Top             =   1875
            Width           =   1020
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "XML"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2145
            TabIndex        =   35
            Top             =   1605
            Width           =   390
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1335
            TabIndex        =   34
            Top             =   1605
            Width           =   420
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   8400
            TabIndex        =   23
            Top             =   720
            Width           =   330
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6810
            TabIndex        =   22
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numero"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4440
            TabIndex        =   21
            Top             =   720
            Width           =   660
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   1335
            TabIndex        =   20
            Top             =   720
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C00000&
         Caption         =   "Datos del Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   210
         TabIndex        =   1
         Top             =   120
         Width           =   9465
         Begin VB.TextBox TxtDireccionFiscal 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   1020
            Width           =   9195
         End
         Begin VB.TextBox TxtRazonSocial 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1965
            TabIndex        =   5
            Top             =   360
            Width           =   7455
         End
         Begin VB.TextBox TxtRif 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direccion Fiscal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4125
            TabIndex        =   13
            Top             =   1875
            Width           =   1380
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre o Razon Social"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   4650
            TabIndex        =   12
            Top             =   795
            Width           =   2010
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RIF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   930
            TabIndex        =   11
            Top             =   795
            Width           =   315
         End
      End
      Begin VB.Label lblDestino 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "....Seleccione Una carpeta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2228
         TabIndex        =   47
         Top             =   8820
         Width           =   6225
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9810
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MnuReportes 
      Caption         =   "Imprimir"
      Begin VB.Menu MnuPorMes 
         Caption         =   "Por Mes"
      End
      Begin VB.Menu MnuCompro 
         Caption         =   "Comprobantes"
      End
   End
   Begin VB.Menu mnuver 
      Caption         =   "Ver"
      Begin VB.Menu MnuListas 
         Caption         =   "Listas"
         Begin VB.Menu MnuComproba 
            Caption         =   "Comprobantes"
         End
      End
   End
End
Attribute VB_Name = "RetencionesISLR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Alicuota As Double
Dim GrabarDireccion As Boolean
Dim IncluirProveedor As Boolean
Dim Consulta As Boolean
Dim MONTOTOTALFACT As Currency
Dim BASEIMPONIBLE As Currency
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean
Private mstrParametro1 As String
Private mlngParametro2 As Long
Private Sub Command1_Click()
 TXTSQL = "select distinct (rif) from Comprobante"
 Set REC = DB.Execute(TXTSQL)
 Do Until REC.EOF
    TXTSQL = "select * from Comprobante where rif='" & REC.Fields("rif") & "'"
    Set REC1 = DB.Execute(TXTSQL)
        TXTSQL = "select * from PROVEEDORES WHERE rif='" & REC.Fields("rif") & "'"
        Set REC2 = DB.Execute(TXTSQL)
        If REC2.EOF Then
            TXTSQL = "INSERT INTO PROVEEDORES (RIF,RazonSocial,DireccionF) VALUES"
            TXTSQL = TXTSQL & "("
            TXTSQL = TXTSQL & "'" & REC1("RIF") & "'"
            TXTSQL = TXTSQL & ",'" & REC1("Razon") & "'"
            TXTSQL = TXTSQL & ",' '"
            TXTSQL = TXTSQL & ")"
            DB.Execute (TXTSQL)
        End If

 REC.MoveNext
 Loop
 MsgBox "Datos Actualizados"
End Sub

Private Sub CboCodCta_Click()
TXTSQL = "SELECT * FROM TABLAISRL WHERE ID =" & CboCodCta.Text & ""
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
    TxtPorcentaRetenc.Text = REC.Fields("PorcentajeRetencion")
    txtCodXml.Text = REC.Fields("ARCHIVOXML")
    txtDescripcion.Text = REC.Fields("Descripcion")
End If
End Sub

Private Sub CmdExportPDF_Click()
ImprimirAIntervaloISLR.Show 1, Me
End Sub

Private Sub CmdImpImpresora_Click()
If MsgBox("¿Desea Imprimir el Comprobante?", vbYesNo + vbQuestion, "Impresion") = vbNo Then
    Exit Sub
End If
        'Abrir el reporte

        TXTSQL = "SELECT * FROM ISLR WHERE NoComprobanteISLR='" & txtNumeCompr.Text & "'"
        Set REC = DB.Execute(TXTSQL)
        If Not REC.EOF Then
            Screen.MousePointer = vbHourglass
            Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionISRL.rpt", 1)
            crReport.RecordSelectionFormula = "{ISLR.NoComprobanteISLR}='" & REC.Fields("NoComprobanteISLR") & "'"
            crReport.PrintOut (False)
            Screen.MousePointer = vbDefault
        Else
            MsgBox "No se encontro el Comprobante"
        End If
        'MsgBox "El Archivo a sido exportado Satisfactoriamente " & Chr(13) & direcciondestinocarpeta
        Screen.MousePointer = vbDefault
        Exit Sub
ErrHandler:
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", vbCritical + vbOKOnly
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly
    End If
    mflgContinuar = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Cmdimpr_Click()
'On Error GoTo errores
IncluirProvee
GuardarComprobante
MsgBox "Comprobante Guardado", vbInformation
TXTSQL = "UPDATE CONFIGURACION SET VALOR='" & Val(txtNumeCompr.Text) & "' WHERE DESCRIPCION='UltimoCompRentecionISLR'"
DB.Execute (TXTSQL)
Exit Sub
errores:
MsgBox "Se produjeron Errores", vbCritical
End Sub

Private Sub CmdNuevo_Click()
If TxtRif.Text <> "" Then
    If MsgBox("¿Realzara retencion a un nuevo proveedor?", vbQuestion + vbYesNo) = vbNo Then
        txtNoFactura.Text = ""
        txtNoControl.Text = ""
        txtMontoTotalFactura.Text = ""
    Else
        TxtRif.Text = ""
        TxtRazonSocial.Text = ""
        TxtDireccionFiscal.Text = ""
        TxtFechafa.Text = ""
        txtNoFactura.Text = ""
        txtNoControl.Text = ""
        txtMontoTotalFactura.Text = ""
    End If
End If
TXTSQL = "SELECT * FROM CONFIGURACION WHERE Descripcion='UltimoCompRentecionISLR'"
Set REC = DB.Execute(TXTSQL)
txtNumeCompr.Text = Format(Val(REC.Fields("VALOR")) + 1, "00000000")
End Sub

Private Sub Form_Load()
'Fechas
txtfechaCompr.Text = Format(Date, "dd/mm/yyyy")
txtMes.Text = Format(Date, "mm")
txtAno.Text = Format(Date, "yyyy")
Consulta = True
'Grillas
Forma_Grid_Rif
'Impuesto
Alicuota = 12 / 100

'NUMERO DE COMPROBANTE
TXTSQL = "SELECT * FROM CONFIGURACION WHERE Descripcion='UltimoCompRentecionISLR'"
Set REC = DB.Execute(TXTSQL)
txtNumeCompr.Text = Format(Val(REC.Fields("VALOR")) + 1, "00000000")

'Codigo de Cuentas
TXTSQL = "SELECT * FROM TablaISRL"
Set REC = DB.Execute(TXTSQL)
CboCodCta.Clear
Do Until REC.EOF
    CboCodCta.AddItem REC.Fields("ID")
    REC.MoveNext
Loop
CboCodCta.Text = 2
CboCodCta_Click
'===== En caso de Facturas Registardas ======'
TXTSQL = "SELECT * FROM COMPROBANTE WHERE Rif='" & NoRifProveedor_G & "' AND NoFactura='" & NoFacturaProveedor_G & "'"
Set REC1 = DB.Execute(TXTSQL)
If Not REC1.EOF Then
    Consulta = False
    TXTSQL = "SELECT * FROM PROVEEDORES WHERE RIF ='" & NoRifProveedor_G & "'"
    Set REC = DB.Execute(TXTSQL)
    If Trim(REC.Fields("DireccionF")) = "" Then
        TxtDireccionFiscal.Text = "No hay registrado"
    Else
        
        TxtDireccionFiscal.Text = REC.Fields("DireccionF")
        GrabarDireccion = True
    End If
    
    TxtRif.Text = NoRifProveedor_G
    TxtRazonSocial.Text = REC1.Fields("RAZON")
    TxtFechafa.Text = REC1.Fields("FechaFactura")
    txtNoFactura.Text = REC1.Fields("NoFactura")
    txtNoControl.Text = REC1.Fields("NoControl")
    txtMontoTotalFactura.Text = Format(REC1.Fields("MontoTotal"), FORMATOMIL)
    txtExento.Text = Format(REC1.Fields("Exento"), FORMATOMIL)
   ' CboCodCta.SetFocus
    Consulta = True
End If ' if del Rec

End Sub

Public Sub Forma_Grid_Rif()
     With MsFProveedores
        .Clear
        .Rows = 1
        .Cols = 3
        .ColWidth(0) = 1700
        .ColWidth(1) = 35
        .ColWidth(2) = 7400
    End With
End Sub

Private Sub Form_Resize()
 Frame9.Left = (Retenciones.Width - Frame9.Width) / 2
End Sub

Private Sub MnuCompro_Click()
ImprimirAIntervaloISLR.Show 1, Me
End Sub



Private Sub MnuComproba_Click()
If MsgBox("¿Desea Ver las Lista de Comprobantes?", vbYesNo + vbQuestion, "Listado de Comprobantes") = vbNo Then
    Exit Sub
End If

End Sub

Private Sub MnuPorMes_Click()
FrmIslrporMes.Show 1, Me
End Sub

Private Sub MsFProveedores_DblClick()
    TxtRif.Text = MsFProveedores.TextMatrix(MsFProveedores.RowSel, 0)
    TxtRazonSocial.Text = MsFProveedores.TextMatrix(MsFProveedores.RowSel, 2)
    MsFProveedores.Visible = False
    TXTSQL = "SELECT * FROM PROVEEDORES WHERE RIF ='" & TxtRif.Text & "'"
    Set REC = DB.Execute(TXTSQL)
    If Trim(REC.Fields("DireccionF")) = "" Then
        TxtDireccionFiscal.Text = "No hay registrado"
    Else
        TxtDireccionFiscal.Text = REC.Fields("DireccionF")
        GrabarDireccion = True
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TxtDireccionFiscal_Change()
GrabarDireccion = True
End Sub

Private Sub TxtDireccionFiscal_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtExento_Change()
txtMontoTotalFactura_Change
End Sub

Private Sub TxtFechafa_KeyPress(KeyAscii As Integer)
  If KeyAscii > 26 Then ' si no es un codigo de control
      If InStr(NUMEROSFECHA & Chr(13), Chr(KeyAscii)) = 0 Then
            MsgBox "Disculpe, ingrese solo Numeros", vbInformation
            KeyAscii = 0
      End If
  End If
End Sub

Private Sub txtMontoTotalFactura_Change()
On Error Resume Next
Dim PorctjdeRetencion As Single
PorctjdeRetencion = (TxtPorcentaRetenc.Text / 100)
BASEIMPONIBLE = Format((MONTOTOTALFACT - CCur(txtExento.Text)) / (1 + Alicuota), FORMATOMIL)
txtIva.Text = Format((MONTOTOTALFACT - CCur(txtExento.Text)) - BASEIMPONIBLE, FORMATOMIL)
TxtBaseImponible.Text = Format(BASEIMPONIBLE, FORMATOMIL)
MONTOTOTALFACT = CCur(txtMontoTotalFactura)
TxtMontoRetencion.Text = Format((BASEIMPONIBLE * PorctjdeRetencion), FORMATOMIL)
TxtTotalaPagar.Text = Format(MONTOTOTALFACT - Format(BASEIMPONIBLE * (PorctjdeRetencion), FORMATOMIL), FORMATOMIL)
End Sub

Private Sub txtMontoTotalFactura_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Dim FUNCIONES As Integer
  If KeyAscii = 44 Or KeyAscii = 46 Then
      FUNCIONES = DECIMALES(KeyAscii)
  End If
  If FUNCIONES <> 0 Then
      KeyAscii = FUNCIONES
  End If
  If KeyAscii > 26 Then ' si no es un codigo de control
      If InStr(SOLONUMEROS & Chr(13), Chr(KeyAscii)) = 0 Then
            MsgBox "Disculpe, ingrese solo Numeros", vbInformation
          KeyAscii = 0
      End If
  End If
End Sub

Private Sub txtNoControl_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNoFactura_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNoFactura_LostFocus()
TXTSQL = "select * from ISLR where rif='" & TxtRif.Text & "' and NoFactura='" & txtNoFactura.Text & "'"
Set REC = DB.Execute(TXTSQL)
Label18.Visible = Not REC.EOF
If Not REC.EOF Then
    txtNoFactura.ForeColor = &HFF&
Else
    txtNoFactura.ForeColor = &H80000008
End If

End Sub

Private Sub TxtPorcentaRetenc_Change()
txtMontoTotalFactura_Change
End Sub

Private Sub TxtRazonSocial_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtRif_Change()
'On Error Resume Next
If Consulta Then
    MsFProveedores.Visible = False
    If TxtRif.Text = "" Then
        Exit Sub
    End If
    MsFProveedores.Rows = 0
    MsFProveedores.Height = MsFProveedores.Rows * 400
    TXTSQL = "select distinct (rif) from Proveedores where rif like('" & Trim(TxtRif.Text) & "%')"
    Set REC = DB.Execute(TXTSQL)
    Do Until REC.EOF
        TXTSQL = "Select * from Proveedores where rif='" & REC!rif & "'"
        Set REC1 = DB.Execute(TXTSQL)
        MsFProveedores.AddItem REC1!rif & vbTab & vbTab & REC1!RazonSocial
        MsFProveedores.Row = MsFProveedores.Rows - 1
        MsFProveedores.Col = 1
        MsFProveedores.CellBackColor = &HC00000
    REC.MoveNext
        If MsFProveedores.Height <= 4925 Then
            MsFProveedores.Height = MsFProveedores.Rows * 400
        End If
    Loop
    TXTSQL = "select distinct (rif) from Comprobante where rif like('" & Trim(TxtRif.Text) & "%')order by RIF"
    Set REC = DB.Execute(TXTSQL)
    MsFProveedores.Visible = Not REC.EOF
End If
End Sub

Private Sub TxtRif_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtRif_LostFocus()
TXTSQL = "SELECT * FROM PROVEEDORES WHERE RIF='" & TxtRif.Text & "'"
Set REC = DB.Execute(TXTSQL)
IncluirProveedor = REC.EOF
End Sub

Public Sub IncluirProvee()
If IncluirProveedor Then
    TXTSQL = "INSERT INTO PROVEEDORES (RIF,RazonSocial,DireccionF) VALUES"
    TXTSQL = TXTSQL & "("
    TXTSQL = TXTSQL & "'" & TxtRif.Text & "'"
    TXTSQL = TXTSQL & ",'" & TxtRazonSocial.Text & "'"
    TXTSQL = TXTSQL & ",'" & TxtDireccionFiscal.Text & "'"
    TXTSQL = TXTSQL & ")"
    DB.Execute (TXTSQL)
End If
If GrabarDireccion = True Then
    TXTSQL = "UPDATE PROVEEDORES SET DireccionF='" & TxtDireccionFiscal.Text & "' WHERE RIF='" & TxtRif.Text & "'"
    DB.Execute (TXTSQL)
End If
End Sub

Public Sub GuardarComprobante()
'Policias
If CboCodCta.Text = "" Then
    MsgBox "Falta clasificar el tipo de Gasto"
    Exit Sub
End If
If txtfechaCompr.Text = "" Then
    MsgBox "Falta Fecha del Comprobante"
    Exit Sub
End If
If txtNumeCompr.Text = "" Then
    MsgBox "Falta Numero del Comprobante"
    Exit Sub
End If
If txtMes.Text = "" Then
    MsgBox "Falta mes del Comprobante"
    Exit Sub
End If
If txtAno.Text = "" Then
    MsgBox "Falta Año del Comprobante"
    Exit Sub
End If
If TxtRif.Text = "" Then
    MsgBox "Falta el Rif del Proveedor"
    Exit Sub
End If
If TxtRazonSocial.Text = "" Then
    MsgBox "Falta la Razon Social del Proveedor"
    Exit Sub
End If
If Trim(TxtDireccionFiscal.Text) = "" Or Trim(TxtDireccionFiscal.Text) = "No hay registrado" Then
    MsgBox "Falta la Direccion Fiscal Social del Proveedor"
    Exit Sub
End If
If TxtFechafa.Text = "" Then
    MsgBox "Falta la Fecha de la factura"
    Exit Sub
End If
If txtNoFactura.Text = "" Then
    MsgBox "Falta el Numero de la factura"
    Exit Sub
End If
If txtNoControl.Text = "" Then
    MsgBox "Falta el Numero de control de la factura"
    Exit Sub
End If
If txtMontoTotalFactura.Text = "" Or CCur(txtMontoTotalFactura.Text) = 0 Then
    MsgBox "Falta el Monto Total de la factura"
    Exit Sub
End If



TXTSQL = "select * from ISLR where rif='" & TxtRif.Text & "' and NoFactura='" & txtNoFactura.Text & "'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
    MsgBox "La Factura Existe" & Chr(13) & "Consulte el Comprobante Nº " & REC.Fields("NoComprobanteISLR")
    Exit Sub
End If
'Genera el Numero de Comprobante si ya ha uno con el numero
 TXTSQL = "select * from ISLR where NoComprobanteISLR='" & txtNumeCompr.Text & "'"
 Set REC = DB.Execute(TXTSQL)
 If Not REC.EOF Then
     Do Until REC.EOF
         txtNumeCompr.Text = Format(Val(txtNumeCompr.Text) + 1, "00000000")
         TXTSQL = "select * from ISLR where NoComprobanteISLR='" & txtNumeCompr.Text & "'"
         Set REC = DB.Execute(TXTSQL)
     Loop
 End If

    TXTSQL = "INSERT INTO ISLR (NoComprobanteISLR,RIF,FechaFactura,NoControl,NoFactura,MontoFactura,BaseImponible,Iva,PorcentajeRetencion,ArchivoXML,DATIMPUESTORETENIDO,ANO,MES,FechaComprobante,IVARETENIDO) VALUES"
    TXTSQL = TXTSQL & "("
    TXTSQL = TXTSQL & "'" & txtNumeCompr.Text & "'"
    TXTSQL = TXTSQL & ",'" & TxtRif.Text & "'"
    TXTSQL = TXTSQL & ",'" & TxtFechafa.Text & "'"
    TXTSQL = TXTSQL & ",'" & txtNoControl.Text & "'"
    TXTSQL = TXTSQL & ",'" & txtNoFactura.Text & "'"
    TXTSQL = TXTSQL & ",'" & txtMontoTotalFactura.Text & "'"
    TXTSQL = TXTSQL & ",'" & BASEIMPONIBLE & "'"
    TXTSQL = TXTSQL & ",'" & txtIva.Text & "'"
    TXTSQL = TXTSQL & ",'" & TxtPorcentaRetenc.Text & "'"
    TXTSQL = TXTSQL & ",'" & txtCodXml.Text & "'"
    TXTSQL = TXTSQL & ",'" & txtDescripcion.Text & "'"
    TXTSQL = TXTSQL & ",'" & txtAno.Text & "'"
    TXTSQL = TXTSQL & ",'" & txtMes.Text & "'"
    TXTSQL = TXTSQL & ",'" & txtfechaCompr.Text & "'"
    TXTSQL = TXTSQL & ",'" & TxtMontoRetencion.Text & "'"
    TXTSQL = TXTSQL & ")"
    DB.Execute (TXTSQL)
    
End Sub
