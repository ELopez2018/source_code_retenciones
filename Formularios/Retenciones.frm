VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form Retenciones 
   Appearance      =   0  'Flat
   BackColor       =   &H000000C0&
   Caption         =   "Retenciones al IVA"
   ClientHeight    =   10470
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   12885
   Icon            =   "Retenciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   12885
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   69
      Top             =   10065
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9816
            Text            =   "Sistema de Retenciones"
            TextSave        =   "Sistema de Retenciones"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9816
            Text            =   "Direccion de Base de datos"
            TextSave        =   "Direccion de Base de datos"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "06/04/2015"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   11670
      Top             =   2880
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Height          =   9615
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   11355
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   120
         TabIndex        =   2
         Top             =   7710
         Width           =   11085
         Begin VB.TextBox TxtBaseImp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   60
            Locked          =   -1  'True
            TabIndex        =   67
            Text            =   "0"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   3080
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   8550
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "0"
            Top             =   945
            Width           =   2415
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   9120
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "0"
            Top             =   420
            Width           =   1845
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   6100
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "0"
            Top             =   420
            Width           =   1845
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Base Imponible:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   135
            TabIndex        =   68
            Top             =   150
            Width           =   1695
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "I.V.A:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   3080
            TabIndex        =   12
            Top             =   150
            Width           =   555
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Retenido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   6990
            TabIndex        =   8
            Top             =   975
            Width           =   1560
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Retencion:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   6100
            TabIndex        =   7
            Top             =   150
            Width           =   1125
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total a Pagar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   9060
            TabIndex        =   6
            Top             =   150
            Width           =   1440
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datos del Comprobante"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   915
         Left            =   1830
         TabIndex        =   23
         Top             =   180
         Width           =   9375
         Begin VB.TextBox TxtNoComprobante 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7560
            TabIndex        =   24
            Text            =   "0000"
            Top             =   240
            Width           =   1725
         End
         Begin MSMask.MaskEdBox MaskEdBox3 
            Height          =   495
            Left            =   2340
            TabIndex        =   25
            Top             =   300
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Comprobante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   240
            Left            =   150
            TabIndex        =   27
            Top             =   420
            Width           =   2115
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comprobante No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   240
            Left            =   5790
            TabIndex        =   26
            Top             =   360
            Width           =   1770
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Periodo Fiscal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   945
         Left            =   90
         TabIndex        =   13
         Top             =   1110
         Width           =   11115
         Begin VB.CommandButton Command8 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":0CCA
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2940
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":150C
            Style           =   1  'Graphical
            TabIndex        =   18
            Tag             =   "1"
            ToolTipText     =   "Ver las Retenciones de la Quincena"
            Top             =   165
            Width           =   750
         End
         Begin VB.TextBox TxtPeriodo 
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
            Left            =   1380
            TabIndex        =   17
            Top             =   390
            Width           =   1485
         End
         Begin VB.CommandButton Command9 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":19C6
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   3720
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":2208
            Style           =   1  'Graphical
            TabIndex        =   16
            Tag             =   "1"
            ToolTipText     =   "Exporta  las Retenciones de la Quincena"
            Top             =   165
            Width           =   750
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":317A
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   4500
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":39BC
            Style           =   1  'Graphical
            TabIndex        =   15
            Tag             =   "1"
            ToolTipText     =   "Resumen"
            Top             =   165
            Width           =   750
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":3F67
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5280
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":47A9
            Style           =   1  'Graphical
            TabIndex        =   14
            Tag             =   "1"
            ToolTipText     =   "Imprime en la impresora las retenciones de la quincena"
            Top             =   165
            Width           =   750
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   405
            Left            =   10215
            TabIndex        =   19
            Top             =   405
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   714
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####"
            PromptChar      =   "-"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   405
            Left            =   9720
            TabIndex        =   20
            Top             =   405
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   714
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##"
            PromptChar      =   "-"
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quincena"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   300
            Left            =   150
            TabIndex        =   22
            Top             =   510
            Width           =   1155
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo Fiscal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   300
            Left            =   7950
            TabIndex        =   21
            Top             =   510
            Width           =   1710
         End
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Periodo Fiscal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   300
         Width           =   1620
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   405
         Left            =   10800
         TabIndex        =   9
         Top             =   9120
         Width           =   495
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   5655
         Left            =   120
         TabIndex        =   28
         Top             =   2055
         Width           =   11085
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   555
            Left            =   1590
            TabIndex        =   30
            Top             =   1560
            Visible         =   0   'False
            Width           =   9345
            _ExtentX        =   16484
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
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
            Height          =   495
            Left            =   1590
            TabIndex        =   31
            Top             =   990
            Visible         =   0   'False
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   873
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
            HighLight       =   2
            GridLines       =   0
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8130
            TabIndex        =   65
            Text            =   "0"
            Top             =   4665
            Width           =   2805
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7200
            TabIndex        =   63
            Top             =   2138
            Width           =   3735
         End
         Begin VB.CommandButton Command4 
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
            Height          =   1035
            Left            =   2250
            Picture         =   "Retenciones.frx":4C28
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   4380
            Width           =   975
         End
         Begin VB.TextBox Text12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1590
            TabIndex        =   49
            Text            =   "0"
            Top             =   3255
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox Text8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10260
            TabIndex        =   48
            Top             =   510
            Width           =   645
         End
         Begin VB.CommandButton Cmdimpr 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":6BE2
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   1275
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":7424
            Style           =   1  'Graphical
            TabIndex        =   47
            Tag             =   "1"
            ToolTipText     =   "Guardar"
            Top             =   4380
            Width           =   975
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":7A86
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   3240
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":82C8
            Style           =   1  'Graphical
            TabIndex        =   46
            Tag             =   "1"
            Top             =   4380
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":8C3D
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   3900
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":947F
            Style           =   1  'Graphical
            TabIndex        =   45
            Tag             =   "1"
            ToolTipText     =   "Ver retenciones del la quincena del Proveedor"
            Top             =   405
            Width           =   750
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":9939
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   4680
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":A17B
            Style           =   1  'Graphical
            TabIndex        =   44
            Tag             =   "1"
            ToolTipText     =   "Exporta  las retenciones del la quincena del Proveedor"
            Top             =   405
            Width           =   750
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":B0ED
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   5460
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":B92F
            Style           =   1  'Graphical
            TabIndex        =   43
            Tag             =   "1"
            ToolTipText     =   "Consulta de Facturas del Proveedor"
            Top             =   405
            Width           =   750
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":BDC1
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   6990
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":C603
            Style           =   1  'Graphical
            TabIndex        =   42
            Tag             =   "1"
            ToolTipText     =   "Imprime en la impresora las retenciones del Proveedor"
            Top             =   420
            Width           =   750
         End
         Begin VB.CommandButton CmdBuscaFactura 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":CA82
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   6240
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":D2C4
            Style           =   1  'Graphical
            TabIndex        =   41
            Tag             =   "1"
            ToolTipText     =   "Busca Factura del Proveedor"
            Top             =   420
            Width           =   750
         End
         Begin VB.CommandButton CmdIsrl 
            BackColor       =   &H00FFFFFF&
            DisabledPicture =   "Retenciones.frx":DC39
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   4200
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":E47B
            Style           =   1  'Graphical
            TabIndex        =   40
            Tag             =   "1"
            Top             =   4380
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            ItemData        =   "Retenciones.frx":EC6A
            Left            =   1590
            List            =   "Retenciones.frx":EC6C
            TabIndex        =   39
            Text            =   "01"
            Top             =   2715
            Width           =   885
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1590
            TabIndex        =   38
            Top             =   510
            Width           =   2235
         End
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1590
            TabIndex        =   37
            Text            =   "0"
            Top             =   3255
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.ComboBox Combo2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            ItemData        =   "Retenciones.frx":EC6E
            Left            =   1590
            List            =   "Retenciones.frx":EC78
            TabIndex        =   36
            Text            =   "ORDINARIO"
            Top             =   2175
            Width           =   2625
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   8130
            TabIndex        =   34
            Text            =   "0"
            Top             =   4031
            Width           =   2805
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7200
            TabIndex        =   33
            Top             =   3400
            Width           =   3735
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   7200
            TabIndex        =   32
            Top             =   2769
            Width           =   3735
         End
         Begin VB.CommandButton Command2 
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
            Height          =   1035
            Left            =   300
            MaskColor       =   &H00E0E0E0&
            Picture         =   "Retenciones.frx":EC91
            Style           =   1  'Graphical
            TabIndex        =   29
            Tag             =   "1"
            ToolTipText     =   "Guardar"
            Top             =   4380
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1590
            TabIndex        =   35
            Top             =   1620
            Width           =   9345
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1590
            TabIndex        =   51
            Top             =   1065
            Width           =   9345
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   6720
            TabIndex        =   66
            Top             =   4815
            Width           =   1395
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha  Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   5610
            TabIndex        =   64
            Top             =   2340
            Width           =   1560
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N de Credito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   225
            TabIndex        =   62
            Top             =   3382
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N. de Debito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   225
            TabIndex        =   61
            Top             =   3382
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RIF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   1170
            TabIndex        =   60
            Top             =   637
            Width           =   375
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razon Social"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   150
            TabIndex        =   59
            Top             =   1192
            Width           =   1395
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No de Control"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   5730
            TabIndex        =   58
            Top             =   2970
            Width           =   1440
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No Factura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   6015
            TabIndex        =   57
            Top             =   3570
            Width           =   1155
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto exento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   6720
            TabIndex        =   56
            Top             =   4185
            Width           =   1395
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "% de Ret."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   9180
            TabIndex        =   55
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Transa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   255
            TabIndex        =   54
            Top             =   2835
            Width           =   1290
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Contr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   240
            TabIndex        =   53
            Top             =   2295
            Width           =   1095
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Direcion Fisc"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   240
            Left            =   150
            TabIndex        =   52
            Top             =   1740
            Width           =   1380
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
         Left            =   120
         TabIndex        =   10
         Top             =   9150
         Width           =   10665
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2970
      Top             =   6060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu MunPPal 
      Caption         =   "Principal"
      Begin VB.Menu MnuDecla 
         Caption         =   "Declaracion"
      End
      Begin VB.Menu MnuCalFiscal 
         Caption         =   "Calendario Fiscal"
      End
      Begin VB.Menu MnuEdicion 
         Caption         =   "Edicion"
      End
   End
   Begin VB.Menu MnuProveedores 
      Caption         =   "Proveedores"
      Begin VB.Menu MnuLista 
         Caption         =   "Lista"
      End
      Begin VB.Menu MnuConsulFacturas 
         Caption         =   "Consulta de Facturas"
      End
   End
   Begin VB.Menu Mnureportes 
      Caption         =   "Reportes"
      Begin VB.Menu MnuComprobante 
         Caption         =   "Comprobante"
         Begin VB.Menu MnuUno 
            Caption         =   "Uno"
         End
         Begin VB.Menu MnuDeXhastaX 
            Caption         =   "De X Comprobante hasta X"
         End
      End
      Begin VB.Menu MnuConsulta 
         Caption         =   "Lista Comprobantes"
      End
      Begin VB.Menu MnuLibComprs 
         Caption         =   "Libro de Compras"
      End
   End
   Begin VB.Menu MnuBqda 
      Caption         =   "Busqueda"
      Begin VB.Menu MnuFactura 
         Caption         =   "Factura"
      End
   End
   Begin VB.Menu MnuISLR 
      Caption         =   "ISLR"
   End
End
Attribute VB_Name = "Retenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EMPRESA As String
Dim MontoTotalFactura As Currency
Dim MontoRetenido As Currency
Dim BASEIMPONIBLE As Currency
Dim Iva As Currency
Dim Alicuota As Currency
Dim Consultarif As Boolean
Dim Consultarazon As Boolean
'Dim Periodo As String
Dim UltimaDeclaracion As Date
Dim proximaDeclaracion As Date
Dim NoComprobante, Condicion_UC As String
Dim TipoContribuyente As String
Dim D_fiscal_g As Boolean
Dim NombreArchivoTxt As String
'variables reporte
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean
Dim Quincena As Integer
Dim Mes_Declaracion As String

'******** variable para consulta ***************
Dim Esconsulta As Boolean
'******** variable para Condicion de Contribuyente Especial ***************
Dim MinimoRetencion As Currency

Private Sub CmdBuscaFactura_Click()
On Error Resume Next
RifProveedor = Text1.Text
NocomproBanteConsulta = InputBox("Por Favor Ingrese el No de Factura", "Consulta de Factura")
If Not IsNumeric(NocomproBanteConsulta) Or NocomproBanteConsulta = 0 Then
    Exit Sub
End If
TipoDeConsulta = "ConsultaPorComprobante_F"
FmReporte.Show 1, Me
End Sub

Private Sub Cmdimpr_Click()
Cmdimpr.Enabled = False
Dim fecha1 As Date
Dim fecha2  As Date
fecha1 = MaskEdBox3.Text
fecha2 = proximaDeclaracion
If fecha1 > fecha2 Then
    MsgBox "La Fecha del comprobante es Mayor a la Fecha de la Declaracion", vbExclamation
    Cmdimpr.Enabled = True
Exit Sub
End If

Guardaproveedor
If Combo1.Text = "02" Then
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE Ndebito='" & Trim(Text12.Text) & "'"
    Set REC = DB.Execute(TXTSQL)
    If REC.EOF = False Then
        MsgBox "El Numero de Nota de Debito Existe", vbExclamation, "Nota de Debito Repetida"
        Cmdimpr.Enabled = True
        Exit Sub
    End If
ElseIf Combo1.Text = "03" Then
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE Ncredito='" & Trim(Text13.Text) & "'"
    Set REC = DB.Execute(TXTSQL)
    If REC.EOF = False Then
    MsgBox "El Numero de Nota de Credito Existe", vbExclamation, "Nota de Credito Repetida"
        Cmdimpr.Enabled = True
        Exit Sub
    End If
End If

Guardar
If Esconsulta Then
    ActualizarCambios
    'UltimoCompRentecionIva
    TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='UltimoCompRentecionIva'"
    Set REC = DB.Execute(TXTSQL)
    If Not REC.EOF Then
        NoComprobante = REC.Fields("vALOR")
    End If
    TxtNoComprobante.Text = NoComprobante + 1
End If
'ActualizaComprobante
SumarIvadePeriodo
If MsgBox("¿Desea Conservar los Datos?", vbQuestion + vbYesNo) = vbNo Then
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = 0
    Text7.Text = 0
End If
Text1.SetFocus
Cmdimpr.Enabled = True
End Sub

Private Sub CmdIsrl_Click()
NoRifProveedor_G = Text1.Text
NoFacturaProveedor_G = Text5.Text
RetencionesISLR.Show 1, Me
End Sub

Private Sub CmdOrgCal_Click()
Dim Quincena_L As String
Quincena = InputBox("Numero de Quincena")
TXTSQL = "UPDATE COMPROBANTE SET CALENDARIO='" & Quincena_L & "-2013' WHERE CALENDARIO='" & Quincena & "'"
DB.Execute (TXTSQL)

End Sub

Private Sub Combo1_Click()
'01 Factura 02 Debito 03 Credito
Select Case Combo1.Text
Case Is = "01"
    Text12.Text = "0" 'Debito
    Text13.Text = "0" 'Credito
    
    Label16.Visible = False 'Debito
    Label17.Visible = False 'Credito
    
    Text12.Visible = False
    Text13.Visible = False
Case Is = "02"
    Text12.Text = "0" 'Debito
    Text13.Text = "0" 'Credito
    Text12.Visible = True
    Text13.Visible = False
    
    Label16.Visible = True 'Debito
    Label17.Visible = False 'Credito
Case Is = "03"
    Text12.Text = "0" 'Debito
    Text13.Text = "0" 'Credito
    Text12.Visible = False
    Text13.Visible = True
        
    Label16.Visible = False 'Debito
    Label17.Visible = True 'Credito
End Select
End Sub

Private Sub Command1_Click()
TipoDeConsulta = "FACTURASDELPROVEEDOR"
RifProveedor = Text1.Text
FmReporte.Show 1, Me
End Sub

Private Sub Command10_Click()
If MsgBox("¿Desea Exporta  las retenciones del la quincena del Proveedor?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If
Dim direcciondestinocarpeta As String
Dim Periodos_Consultado  As String
RifProveedor = Text1.Text
PeriodoConsulta = TxtPeriodo.Text
    direcciondestinocarpeta = Buscar_Carpeta(" ... Seleccione una carpeta ")
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    'On Error GoTo ErrHandler
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE Calendario='" & PeriodoConsulta & "' and Rif='" & RifProveedor & "' AND NoComprobante<>'N/A'"
    Set REC1 = DB.Execute(TXTSQL)
    
    Do Until REC1.EOF
        'Abrir el reporte
        Screen.MousePointer = vbHourglass
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionCF.rpt", 1)
        TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & REC1.Fields("NoComprobante") & "'"
        Set REC = DB.Execute(TXTSQL)
        If Not REC.EOF Then
        Periodos_Consultado = REC.Fields("Calendario")
        
        If Dir(direcciondestinocarpeta & "\Quincena " & Periodos_Consultado, vbDirectory) = "" Then
            'MsgBox "La carpeta no existe"
            Call MkDir(direcciondestinocarpeta & "\Quincena " & Periodos_Consultado)
            'MsgBox "La carpeta Creada" & Chr(13) & direcciondestinocarpeta & "\Quincena " & Periodos_Consultado
        End If
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionCF.rpt", 1)
        crReport.RecordSelectionFormula = "{comprobante.Nocomprobante}='" & REC1.Fields("Nocomprobante") & "'"
        crReport.ExportOptions.FormatType = crEFTPortableDocFormat
        crReport.ExportOptions.DestinationType = crEDTDiskFile
        crReport.ExportOptions.DiskFileName = direcciondestinocarpeta & "\Quincena " & Periodos_Consultado & "\IVA-" & REC.Fields("NoComprobante") & "-" & REC.Fields("NoFactura") & "-" & REC.Fields("Razon") & ".pdf"
        crReport.ExportOptions.PDFExportAllPages = True
        crReport.Export (False)
        Screen.MousePointer = vbDefault
        Set crParamDef = Nothing
        End If
        
        REC1.MoveNext
    Loop
    MsgBox "El Archivo a sido exportado Satisfactoriamente " & Chr(13) & direcciondestinocarpeta, vbInformation
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

Private Sub Command11_Click()
If MsgBox("¿Desea ver el Resumen?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If
On Error Resume Next
RifProveedor = TxtPeriodo.Text
If Not IsNumeric(NocomproBanteConsulta) Then
    Exit Sub
End If
TipoDeConsulta = "RESUMEN"
FmReporte.Show 1, Me
End Sub

Private Sub Command12_Click()
If MsgBox("¿Desea Imprimir las Retenciones de la Quincena?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If
Dim Periodos_Consultado  As String
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE Calendario='" & TxtPeriodo.Text & "'"
    Set REC1 = DB.Execute(TXTSQL)
    
    Do Until REC1.EOF
        'Abrir el reporte
        Screen.MousePointer = vbHourglass
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
        TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & REC1.Fields("NoComprobante") & "'"
        Set REC = DB.Execute(TXTSQL)
        If Not REC.EOF Then
            Periodos_Consultado = REC.Fields("Calendario")
            Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
            crReport.RecordSelectionFormula = "{comprobante.Nocomprobante}='" & REC1.Fields("NoComprobante") & "'"
            crReport.PrintOut (False)
            Screen.MousePointer = vbDefault
        End If
        REC1.MoveNext
    Loop
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

Private Sub Command13_Click()
If MsgBox("¿Desea Imprimir las retenciones del la quincena del Proveedor?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If

Dim direcciondestinocarpeta As String
Dim Periodos_Consultado  As String
RifProveedor = Text1.Text
PeriodoConsulta = TxtPeriodo.Text

    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE Calendario='" & PeriodoConsulta & "' and Rif='" & RifProveedor & "'"
    Set REC1 = DB.Execute(TXTSQL)
    Do Until REC1.EOF
        'Abrir el reporte
        Screen.MousePointer = vbHourglass
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
        TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & REC1.Fields("NoComprobante") & "'"
        Set REC = DB.Execute(TXTSQL)
        If Not REC.EOF Then
            Periodos_Consultado = REC.Fields("Calendario")
            Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
            crReport.RecordSelectionFormula = "{comprobante.Nocomprobante}='" & REC1.Fields("Nocomprobante") & "'"
            crReport.PrintOut False
            Screen.MousePointer = vbDefault
        End If
        REC1.MoveNext
    Loop
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

Private Sub Command2_Click()
Consultarif = False
Consultarazon = False
TXTSQL = ""
Forma_Grid_Rif
Forma_Grid_razon

MaskEdBox3.Text = Format(Date, "dd/mm/yyyy")
TxtNoComprobante.Text = 1

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='IVA'"
Set REC = DB.Execute(TXTSQL)
If Not IsNull(REC.Fields("vALOR")) Then
    Alicuota = CCur(REC.Fields("vALOR")) / 100
End If

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='CarpetaDestino'"
Set REC = DB.Execute(TXTSQL)
If Not REC.Fields("cadena") = "" Then
    lblDestino.Caption = REC.Fields("cadena")
End If

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='Periodo'"
Set REC = DB.Execute(TXTSQL)
If Not IsNull(REC.Fields("vALOR")) Then
    Periodo = REC.Fields("vALOR")
End If
TxtPeriodo.Text = Periodo

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='Empresa'"
Set REC = DB.Execute(TXTSQL)
If Not IsNull(REC.Fields("vALOR")) Then
    EMPRESA = REC.Fields("vALOR")
End If

'UltimoCompRentecionIva
TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='UltimoCompRentecionIva'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
NoComprobante = REC.Fields("vALOR")
End If
TxtNoComprobante.Text = NoComprobante + 1

'Tipo De transaccion
TXTSQL = "select distinct(TipoTansaccion) from Comprobante "
Set REC = DB.Execute(TXTSQL)
Combo1.Clear
Do Until REC.EOF
    Combo1.AddItem REC.Fields("TipoTansaccion")
    REC.MoveNext
Loop
Combo1.Text = "01"
Text8.Text = "75"


'Declaracion
TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='UltimaDeclaracion'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
UltimaDeclaracion = REC.Fields("VALOR")
End If

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='proximaDeclaracion'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
    proximaDeclaracion = REC.Fields("VALOR")
    MaskEdBox1.Text = Format(proximaDeclaracion, "mm")
    MaskEdBox5.Text = Format(proximaDeclaracion, "yyyy")
End If

'quincena actual
TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='Quincena'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
Quincena = REC.Fields("VALOR")
End If
SumarIvadePeriodo
Esconsulta = False

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='CondicionRetencionesUnidadTrib'"
Set REC = DB.Execute(TXTSQL)
MinimoRetencion = REC.Fields("valor") * Val(REC.Fields("Cadena"))

End Sub

Private Sub Command3_Click()
Dim InicioQuincena As Date
Dim FinQuincena As Date
'Policia
     If MsgBox("¿Desea pasar a la Proxima Quincena?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
     End If
     
    TXTSQL = "Select * from CalendarioFiscal where Id=" & Quincena + 1 & ""
    Set REC = DB.Execute(TXTSQL)
    If REC.EOF Then
        MsgBox "Por favor, configure el calendario fiscal"
        Exit Sub
    End If
    
    TXTSQL = "Select * from CalendarioFiscal where Id=" & Quincena & ""
    Set REC1 = DB.Execute(TXTSQL)
    If REC1.EOF Then
        MsgBox "Por favor, configure el calendario fiscal"
        Exit Sub
    End If
    
    'Usamos la funcion DAteSerial para obtener el primero y el ultimo dia
    'Primer = DateSerial(Year(FECHA), Month(FECHA) + 0, 1)'Primer Dia del Mes
    'Ultimo = DateSerial(Year(FECHA), Month(FECHA) + 1, 0) 'Ultimo dia del mes
    
    InicioQuincena = REC1.Fields("FechadeDeclaracion") + 1 'VALOR DEL INICIO DE LA QUINCENA
    FinQuincena = REC.Fields("FechadeDeclaracion") 'VALOR DEL FINAL DE LA QUINCENA
    MsgBox InicioQuincena & Chr(13) & FinQuincena
    
    'Actualiza La Quicena
    TXTSQL = "update Configuracion set Valor='" & Quincena + 1 & "' WHERE DESCRIPCION='Quincena'"
    DB.Execute (TXTSQL)
    
    'Actualiza Los Comprobantes del periodo a declarados
    TXTSQL = "update Comprobante set Declarado=true where FechaComprobante>cdate('" & UltimaDeclaracion & "') and FechaComprobante<cdate('" & proximaDeclaracion & "')"
    DB.Execute (TXTSQL)
    
    'Periodo, configura la variable al siguiente periodo
    TXTSQL = "update Configuracion set Valor='" & REC.Fields("Periodo") & "' WHERE DESCRIPCION='Periodo'"
    DB.Execute (TXTSQL)
    

    
    'UltimaDeclaracion, actualiza la fecha de la ultima declaracion
    TXTSQL = "update Configuracion set Valor='" & REC1.Fields("FechadeDeclaracion") & "' WHERE DESCRIPCION='UltimaDeclaracion'"
    DB.Execute (TXTSQL)
    
    'ProximaDeclaracion, actualiza la fecha de la proxima declaracion
    TXTSQL = "Select * from CalendarioFiscal where id=" & Quincena + 1 & ""
    Set REC = DB.Execute(TXTSQL)
    
    TXTSQL = "update Configuracion set Valor='" & FinQuincena & "' WHERE DESCRIPCION='ProximaDeclaracion'"
    DB.Execute (TXTSQL)
    
    'Actualizar
    TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='Periodo'"
    Set REC = DB.Execute(TXTSQL)
    If Not REC.EOF Then
    Periodo = REC.Fields("vALOR")
    End If
    
    TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='UltimaDeclaracion'"
    Set REC = DB.Execute(TXTSQL)
    If Not REC.EOF Then
        UltimaDeclaracion = REC.Fields("VALOR")
    End If
    
    TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='proximaDeclaracion'"
    Set REC = DB.Execute(TXTSQL)
    If Not REC.EOF Then
        proximaDeclaracion = REC.Fields("VALOR")
    End If

    TxtPeriodo.Text = Periodo
    MsgBox "Se han realizado los cambios Satisfactoriamente", vbInformation
    
    SumarIvadePeriodo
End Sub

Private Sub Command4_Click()
On Error GoTo ErrHandler
    Dim Periodos_Consultado As String
    'Set VLman_arch = New FileSystemObject
    If lblDestino.Caption = "....Seleccione Una carpeta" Then
        MsgBox "....Seleccione Una carpeta", vbCritical
        Command5_Click
        Exit Sub
    End If
    
    NocomproBanteConsulta = TxtNoComprobante.Text
    NocomproBanteConsulta = InputBox("Por Favor Ingrese el No de comprobante", "Consulta de Comprobante", NocomproBanteConsulta)
    If Not IsNumeric(NocomproBanteConsulta) Then
        Exit Sub
    End If
    
     Dim TipoImpresionComprobante As String
    If MsgBox("¿Desea que la Retencion Contenga La Firma y sello?", vbQuestion + vbYesNo) = vbYes Then
        TipoImpresionComprobante = "\ComproRetencionCF.rpt"
    Else
    
        TipoImpresionComprobante = "\ComproRetencion.rpt"
    End If
    
    
    
    NoComprobante = Format(NocomproBanteConsulta, "00000000")
    'Abrir el reporte
    Screen.MousePointer = vbHourglass
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & TipoImpresionComprobante, 1)
    crReport.RecordSelectionFormula = "{comprobante.nocomprobante}='" & Format(NoComprobante, "00000000") & "'"
    
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & Format(NoComprobante, "00000000") & "'"
    Set REC = DB.Execute(TXTSQL)
    Periodos_Consultado = REC.Fields("Calendario")
    
    If Dir(lblDestino.Caption & "\Quincena " & Periodos_Consultado, vbDirectory) = "" Then
        'MsgBox "La carpeta no existe"
        Call MkDir(lblDestino.Caption & "\Quincena " & Periodos_Consultado)
        MsgBox "La carpeta Creada" & Chr(13) & lblDestino.Caption & "\Quincena " & Periodos_Consultado
    End If
    crReport.ExportOptions.FormatType = crEFTPortableDocFormat
    crReport.ExportOptions.DestinationType = crEDTDiskFile
    crReport.ExportOptions.DiskFileName = lblDestino.Caption & "\Quincena " & Periodos_Consultado & "\IVA-" & REC.Fields("NoComprobante") & "-" & REC.Fields("NoFactura") & "-" & REC.Fields("Razon") & ".pdf"
    crReport.ExportOptions.PDFExportAllPages = True
    crReport.Export (False)
    Screen.MousePointer = vbDefault
    MsgBox "El Archivo a sido exportado Satisfactoriamente " & Chr(13) & lblDestino.Caption, vbInformation
    
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

Private Sub Command5_Click()
lblDestino.Caption = Buscar_Carpeta(" ... Seleccione una carpeta ")
TXTSQL = "UPDATE CONFIGURACION SET CADENA='" & lblDestino.Caption & "'  WHERE DESCRIPCION='CarpetaDestino'"
DB.Execute (TXTSQL)
End Sub

Private Sub Command6_Click()
On Error Resume Next
NocomproBanteConsulta = TxtNoComprobante.Text
NocomproBanteConsulta = InputBox("Por Favor Ingrese el No de comprobante", "Consulta de Comprobante", NocomproBanteConsulta)
If Not IsNumeric(NocomproBanteConsulta) Or NocomproBanteConsulta = 0 Then
    Exit Sub
End If
TipoDeConsulta = "PORCOMPROBANTE"
FmReporte.Show 1, Me
End Sub

Private Sub Command7_Click()
TipoDeConsulta = "COMPROBANTESDELPROVEEDOR"
RifProveedor = Text1.Text
PeriodoConsulta = TxtPeriodo.Text
FmReporte.Show 1, Me
End Sub

Private Sub Command8_Click()
If MsgBox("¿Desea Ver las Retenciones de la Quincena?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If

On Error Resume Next
RifProveedor = TxtPeriodo.Text
If Not IsNumeric(NocomproBanteConsulta) Then
Exit Sub
End If
TipoDeConsulta = "PORQUINCENA"
FmReporte.Show 1, Me
End Sub

Private Sub Command9_Click()
If MsgBox("¿Desea Exporta  las Retenciones de la Quincena?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If

Dim direcciondestinocarpeta As String
Dim Periodos_Consultado  As String
    direcciondestinocarpeta = Buscar_Carpeta(" ... Seleccione una carpeta ")
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    'On Error GoTo ErrHandler
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE Calendario='" & TxtPeriodo.Text & "'"
    Set REC1 = DB.Execute(TXTSQL)
    
    Do Until REC1.EOF
        'Abrir el reporte
        Screen.MousePointer = vbHourglass
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionCF.rpt", 1)
        TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & REC1.Fields("NoComprobante") & "'"
        Set REC = DB.Execute(TXTSQL)
        If Not REC.EOF Then
        Periodos_Consultado = REC.Fields("Calendario")
        
        If Dir(direcciondestinocarpeta & "\Quincena " & Periodos_Consultado, vbDirectory) = "" Then
            'MsgBox "La carpeta no existe"
            Call MkDir(direcciondestinocarpeta & "\Quincena " & Periodos_Consultado)
            'MsgBox "La carpeta Creada" & Chr(13) & direcciondestinocarpeta & "\Quincena " & Periodos_Consultado
        End If
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionCF.rpt", 1)
        crReport.RecordSelectionFormula = "{comprobante.Nocomprobante}='" & REC1.Fields("NoComprobante") & "'"
        crReport.ExportOptions.FormatType = crEFTPortableDocFormat
        crReport.ExportOptions.DestinationType = crEDTDiskFile
        crReport.ExportOptions.DiskFileName = direcciondestinocarpeta & "\Quincena " & Periodos_Consultado & "\IVA-" & REC.Fields("NoComprobante") & "-" & REC.Fields("NoFactura") & "-" & REC.Fields("Razon") & ".pdf"
        crReport.ExportOptions.PDFExportAllPages = True
        crReport.Export (False)
        Screen.MousePointer = vbDefault
        Set crParamDefs = Nothing
        Set crParamDef = Nothing
        End If
        
        REC1.MoveNext
    Loop
    MsgBox "El Archivo a sido exportado Satisfactoriamente " & Chr(13) & direcciondestinocarpeta
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    ConsulCompobuscar = False
    TituloVentana = "Resultados de la Busqueda segun Numero de Factura"
    ConsultaMotosVendidas3
ElseIf KeyCode = vbKeyF4 Then
    TituloVentana = "Resultados de la Busqueda segun Numero de Control"
    ConsulCompobuscar = False
    ConsultaRetencionNoControl
ElseIf KeyCode = vbKeyF5 Then
    MsgBox MinimoRetencion
End If
End Sub

Private Sub Form_Load()
Consultarif = False
Consultarazon = False
Frame9.BorderStyle = False
TXTSQL = ""
Forma_Grid_Rif
Forma_Grid_razon

MaskEdBox3.Text = Format(Date, "dd/mm/yyyy")
TxtNoComprobante.Text = 1

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='IVA'"
Set REC = DB.Execute(TXTSQL)
If Not IsNull(REC.Fields("vALOR")) Then
    Alicuota = CCur(REC.Fields("vALOR")) / 100
End If

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='CarpetaDestino'"
Set REC = DB.Execute(TXTSQL)
If Not REC.Fields("cadena") = "" Then
    lblDestino.Caption = REC.Fields("cadena")
End If

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='Periodo'"
Set REC = DB.Execute(TXTSQL)
If Not IsNull(REC.Fields("vALOR")) Then
    Periodo = REC.Fields("vALOR")
End If
TxtPeriodo.Text = Periodo

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='Empresa'"
Set REC = DB.Execute(TXTSQL)
If Not IsNull(REC.Fields("vALOR")) Then
    EMPRESA = REC.Fields("vALOR")
End If

'UltimoCompRentecionIva
TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='UltimoCompRentecionIva'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
NoComprobante = REC.Fields("vALOR")
End If
TxtNoComprobante.Text = NoComprobante + 1

'Tipo De transaccion
TXTSQL = "select distinct(TipoTansaccion) from Comprobante "
Set REC = DB.Execute(TXTSQL)
Combo1.Clear
Do Until REC.EOF
    Combo1.AddItem REC.Fields("TipoTansaccion")
    REC.MoveNext
Loop
Combo1.Text = "01"
Text8.Text = "75"


'Declaracion
TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='UltimaDeclaracion'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
UltimaDeclaracion = REC.Fields("VALOR")
End If

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='proximaDeclaracion'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
    proximaDeclaracion = REC.Fields("VALOR")
    MaskEdBox1.Text = Format(proximaDeclaracion, "mm")
    MaskEdBox5.Text = Format(proximaDeclaracion, "yyyy")
End If

'quincena actual
TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='Quincena'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
Quincena = REC.Fields("VALOR")
End If
SumarIvadePeriodo
Esconsulta = False

TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='CondicionRetencionesUnidadTrib'"
Set REC = DB.Execute(TXTSQL)
MinimoRetencion = REC.Fields("valor") * Val(REC.Fields("Cadena"))
StatusBar1.Panels(2).Text = DireccionBaseDatos
End Sub

Public Sub Guardar()
Dim NumeroComprobante_V As String
'POLICIAS
If Esconsulta Then 'Controla si se esta consultando
    Exit Sub
End If

If Text1.Text = "" Then
    MsgBox "Falta Rif"
    Text1.SetFocus
    Exit Sub
End If
If Text2.Text = "" Then
    MsgBox "Razon Social"
    Text2.SetFocus
    Exit Sub
End If
If Text5.Text = "" Then
    MsgBox "Falta No de Factura"
    Text5.SetFocus
    Exit Sub
End If

'Genera el Numero de Comprobante si ya ha uno con el numero
TXTSQL = "select * from COMPROBANTE where NoComprobante='" & TxtNoComprobante.Text & "'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
    Do Until REC.EOF
        TxtNoComprobante.Text = Val(TxtNoComprobante.Text) + 1
        TXTSQL = "select * from COMPROBANTE where NoComprobante='" & TxtNoComprobante.Text & "'"
        Set REC = DB.Execute(TXTSQL)
    Loop
End If

TXTSQL = "select * from COMPROBANTE where Rif='" & Text1.Text & "' and NoFactura='" & Text5.Text & "'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
    MsgBox "Ya se le ha hecho retencion a esa factura" & Chr(13) & "Comprobante No " & REC.Fields("NoComprobante"), vbExclamation, "Comprobante No " & REC.Fields("NoComprobante")
    If MsgBox("Ya se le ha hecho retencion a esa factura" & Chr(13) & "¿Va a hacer una Nota de Credito o debito?", vbQuestion + vbYesNo) = vbNo Then
        If MsgBox("¿Desea que el sistema le asigne un Numero de Comprobante?", vbQuestion + vbYesNo) = vbNo Then
          Exit Sub
        Else
            'Genera el Numero de Comprobante si ya ha uno con el numero
            TXTSQL = "select * from COMPROBANTE where NoComprobante='" & TxtNoComprobante.Text & "'"
            Set REC = DB.Execute(TXTSQL)
            If Not REC.EOF Then
                Do Until REC.EOF
                    TxtNoComprobante.Text = Val(TxtNoComprobante.Text) + 1
                    TXTSQL = "select * from COMPROBANTE where NoComprobante='" & TxtNoComprobante.Text & "'"
                    Set REC = DB.Execute(TXTSQL)
                Loop
            End If
        End If
    End If
End If

If MontoRetenido <= 0 Then
    NumeroComprobante_V = "N/A"
Else
    NumeroComprobante_V = Format(TxtNoComprobante.Text, "00000000")
End If

NoComprobante = Val(TxtNoComprobante.Text)
TXTSQL = "INSERT INTO COMPROBANTE"
    TXTSQL = TXTSQL & "(NoComprobante,Rif,Razon,FechaComprobante,FechaFactura,NoControl,NoFactura,MontoTotal"
    TXTSQL = TXTSQL & ",Exento,IvaRetenido,TipoTansaccion,Empresa,PorcentajeIva,MesPeriodoFiscal, AnoPeriodoFiscal,PorcentajeReten,Ndebito,Ncredito,calendario,declarado)"
    TXTSQL = TXTSQL & " VALUES "
    TXTSQL = TXTSQL & "("
    TXTSQL = TXTSQL & "'" & NumeroComprobante_V & "'"
    TXTSQL = TXTSQL & ",'" & Text1.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text2.Text & "'"
    TXTSQL = TXTSQL & ",'" & MaskEdBox3.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text3.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text4.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text5.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text7.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text6.Text & "'"
    TXTSQL = TXTSQL & ",'" & MontoRetenido & "'"
    TXTSQL = TXTSQL & ",'" & Combo1.Text & "'"
    TXTSQL = TXTSQL & ",'" & EMPRESA & "'"
    TXTSQL = TXTSQL & ",'" & Alicuota * 100 & "'"
    TXTSQL = TXTSQL & ",'" & MaskEdBox1.Text & "'"
    TXTSQL = TXTSQL & ",'" & MaskEdBox5.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text8.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text12.Text & "'"
    TXTSQL = TXTSQL & ",'" & Text13.Text & "'"
    TXTSQL = TXTSQL & ",'" & TxtPeriodo.Text & "'"
    TXTSQL = TXTSQL & ",false"
    TXTSQL = TXTSQL & ")"
    DB.Execute (TXTSQL)
    MsgBox "Se guardo bajo el No " & NumeroComprobante_V, vbInformation, "Guardado sin errores"
    'Actualizar No Comprobante
    If MsgBox("¿Desea imprimir el comprobante?", vbQuestion + vbYesNo, "Comprobante Nº" & TxtNoComprobante.Text) = vbYes Then

        '///////////////////////////////////////////////NUEV0
        
        Dim numerocopias As Integer
        Dim repuesta_l As String
        repuesta_l = "empezar"
        numerocopias = 0
        Do Until numerocopias > 0 Or UCase(repuesta_l) = "SALIR" Or repuesta_l = ""
            repuesta_l = InputBox("No de Compias a Imprimir", "Comprobante Nº " & TxtNoComprobante.Text, 1)
            If IsNumeric(repuesta_l) Then
            numerocopias = repuesta_l
            Else
            MsgBox repuesta_l & " No es un Numero, vuela a Intentarlo o Escriba 'Salir'"
            End If
        Loop
        
        Dim I_L As Integer
        I_L = 1
        Do Until I_L > numerocopias
            Screen.MousePointer = vbHourglass
            Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
            crReport.RecordSelectionFormula = "{Comprobante.NoComprobante}='" & Format(TxtNoComprobante.Text, "00000000") & "'"
            crReport.PrintOut (False)
            Screen.MousePointer = vbDefault
        I_L = I_L + 1
        Loop
        '///////////////////////////////////////////////FIN NUEVO
    End If
    
    TXTSQL = "update Configuracion set Valor='" & TxtNoComprobante.Text & "'  WHERE DESCRIPCION='UltimoCompRentecionIva'"
    DB.Execute (TXTSQL)
    TxtNoComprobante.Text = NoComprobante + 1
End Sub

Private Sub Form_Resize()
 Frame9.Left = (Retenciones.Width - Frame9.Width) / 2
End Sub

Private Sub MnuCalFiscal_Click()
    CalendarioFiscal.Show 1
End Sub

Private Sub MnuConsulta_Click()
FrmConsulta.Show 1, Me
End Sub

Private Sub MnuDecla_Click()
'Dim QuincenaDeclarar As Date
'Adodc1.ConnectionString = ConexionDatasABaseD
'Adodc1.RecordSource = "SELECT * FROM COMPROBANTE WHERE calendario='" & TxtPeriodo.Text & "' and IvaRetenido<>0 order by val(NoComprobante)"
'Adodc1.Refresh
'QuincenaDeclarar = Adodc1.Recordset.Fields("FechaComprobante")
'Mes_Declaracion = UCase(Format(QuincenaDeclarar, "MMMM") & " DE " & Format(QuincenaDeclarar, "YYYY"))
'If Format(QuincenaDeclarar, "DD") > 15 Then
    'NombreArchivoTxt = "2DA QNA DE "
'Else
    'NombreArchivoTxt = "1RA QNA DE "
'End If
'GenerarTxt
FrmDeclaracion.Show 1, Me
End Sub

Private Sub MnuDeXhastaX_Click()
ImprimirAIntervalo.Show 1
End Sub

Private Sub MnuEdicion_Click()
Editor.Show 1, Me
End Sub


Private Sub MnuExportar_Click()
'FrmParamExportar.Show 1, Me
ComprobanteInicial = 1807
ComprobanteFinal = 1942

Dim direcciondestinocarpeta As String
Dim Periodos_Consultado  As String
    direcciondestinocarpeta = Buscar_Carpeta(" ... Seleccione una carpeta ")
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    'On Error GoTo ErrHandler
    Do Until ComprobanteInicial > ComprobanteFinal
    'Abrir el reporte
    Screen.MousePointer = vbHourglass
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & Format(ComprobanteInicial, "00000000") & "'"
    Set REC = DB.Execute(TXTSQL)
    If Not REC.EOF Then
    Periodos_Consultado = REC.Fields("Calendario")
    
    If Dir(direcciondestinocarpeta & "\Quincena " & Periodos_Consultado, vbDirectory) = "" Then
        'MsgBox "La carpeta no existe"
        Call MkDir(direcciondestinocarpeta & "\Quincena " & Periodos_Consultado)
        'MsgBox "La carpeta Creada" & Chr(13) & direcciondestinocarpeta & "\Quincena " & Periodos_Consultado
    End If
    
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
    TXTSQL = "select * from comprobante where nocomprobante='" & Format(ComprobanteInicial, "00000000") & "'"
    Set REC = DB.Execute(TXTSQL)
    crReport.RecordSelectionFormula = "{comprobante.Nocomprobante}='" & Format(REC.Fields("NoComprobante"), "00000000") & "'"
    crReport.ExportOptions.FormatType = crEFTPortableDocFormat
    crReport.ExportOptions.DestinationType = crEDTDiskFile
    crReport.ExportOptions.DiskFileName = direcciondestinocarpeta & "\Quincena " & Periodos_Consultado & "\IVA-" & REC.Fields("NoComprobante") & "-" & REC.Fields("NoFactura") & "-" & REC.Fields("Razon") & ".pdf"
    crReport.ExportOptions.PDFExportAllPages = True
    crReport.Export (False)
    Screen.MousePointer = vbDefault
End If
        ComprobanteInicial = ComprobanteInicial + 1
    Loop
    MsgBox "El Archivo a sido exportado Satisfactoriamente " & Chr(13) & lblDestino.Caption, vbInformation
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

Private Sub MnuFactura_Click()
On Error Resume Next
NoFacturaPaConsultar = Text5.Text
NoFacturaPaConsultar = InputBox("Por Favor Ingrese el No de Factura", "Consulta de Comprobante", NoFacturaPaConsultar)
If NoFacturaPaConsultar = "" Then
    Exit Sub
End If
TipoDeConsulta = "BUSCARFACTURA"
FmReporte.Show 1, Me
'410846
End Sub

Private Sub MnuISLR_Click()
RetencionesISLR.Show 1
End Sub

Private Sub MnuLibComprs_Click()
FrmLibCompra.Show 1
End Sub

Private Sub MnuUno_Click()
On Error Resume Next
NocomproBanteConsulta = InputBox("Por Favor Ingrese el No de comprobante", "Consulta de Comprobante")
If Not IsNumeric(NocomproBanteConsulta) Then
Exit Sub
End If
TipoDeConsulta = "PORCOMPROBANTE"
FmReporte.Show 1, Me
End Sub

Private Sub MSFlexGrid1_Click()
'On Error Resume Next
    Text2.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0)
    Text2_KeyPress 13
    MSFlexGrid1.Visible = False
    TXTSQL = "SELECT * FROM Proveedores WHERE razonSOCIAL='" & Text2.Text & "'"
    Set REC = DB.Execute(TXTSQL)
    Text1.Text = REC.Fields("Rif")
    Text10.Text = REC.Fields("DireccionF")
    Combo2.Text = REC.Fields("TipoContrubuyente") & ""
    Condicion_UC = " razon='" & Text2.Text & "'"
    If Trim(Text10.Text) = "" Then
        D_fiscal_g = True
    End If
    UltimoComprobanteporProvee
End Sub

Private Sub MSFlexGrid3_Click()
    Text1.Text = MSFlexGrid3.TextMatrix(MSFlexGrid3.RowSel, 0)
    Condicion_UC = " Rif='" & Text1.Text & "'"
    Text1_KeyPress 13
    MSFlexGrid3.Visible = False
    If Text10.Text = "" Then
        D_fiscal_g = True
    End If
    UltimoComprobanteporProvee
End Sub

Private Sub Text1_Change()
'On Error Resume Next
If Consultarif Then
    MSFlexGrid3.Visible = False
    If Text1.Text = "" Then
        Exit Sub
    End If
    MSFlexGrid3.Rows = 0
    MSFlexGrid3.Height = MSFlexGrid3.Rows * 400
    TXTSQL = "select * from Proveedores where rif like('" & Trim(Text1.Text) & "%')"
    Set REC = DB.Execute(TXTSQL)
    Do Until REC.EOF
        MSFlexGrid3.AddItem REC!rif
        REC.MoveNext
        If MSFlexGrid3.Height <= 5925 Then
            MSFlexGrid3.Height = MSFlexGrid3.Rows * 400
        End If
    Loop
    TXTSQL = "select * from Proveedores where rif like('" & Trim(Text1.Text) & "%')order by RIF"
    Set REC = DB.Execute(TXTSQL)
    MSFlexGrid3.Visible = Not REC.EOF
End If
End Sub
Public Sub Forma_Grid_Rif()
     With MSFlexGrid3
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 1800
    End With
End Sub
Public Sub Forma_Grid_razon()
     With MSFlexGrid1
        .Clear
        .Rows = 1
        .Cols = 1
        .ColWidth(0) = 8000
    End With
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Consultarif = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
Dim FUNCIONES As Integer
  If KeyAscii = 44 Or KeyAscii = 46 Then
      FUNCIONES = DECIMALES(KeyAscii)
  End If
  If FUNCIONES <> 0 Then
      KeyAscii = FUNCIONES
  End If
  If KeyAscii > 26 Then ' si no es un codigo de control
      If InStr(NumerosCedulaRif & Chr(13), Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
      End If
  End If
    If KeyAscii = 13 Then ' si no es un codigo de control
            TXTSQL = "select * from Proveedores where rif ='" & Trim(Text1.Text) & "'"
            Set REC = DB.Execute(TXTSQL)
            If Not REC.EOF Then
                Dim rif_vl As String
                Dim razon_vl As String
                Dim TipoC_vl As String
                Dim DirFisc_vl As String
                
                rif_vl = REC.Fields("Rif")
                razon_vl = REC.Fields("RazonSocial")
                DirFisc_vl = REC.Fields("DireccionF")
                TipoC_vl = REC.Fields("TipoContrubuyente") & ""
                
                Text1.Text = rif_vl
                Text2.Text = razon_vl
                Text10.Text = DirFisc_vl
                Combo2.Text = TipoC_vl
            End If
    End If
End Sub

Private Sub Text1_LostFocus()
Consultarif = False
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text2_Change()
If Consultarazon Then
    MSFlexGrid1.Visible = False
    If Text2.Text = "" Then
        Exit Sub
    End If
    MSFlexGrid1.Rows = 0
    MSFlexGrid1.Height = MSFlexGrid1.Rows * 400
    TXTSQL = "select distinct (Razon) from Comprobante where Razon like('" & Trim(Text2.Text) & "%')order by Razon"
    Set REC = DB.Execute(TXTSQL)
    Do Until REC.EOF
        MSFlexGrid1.AddItem REC!Razon
        REC.MoveNext
        If MSFlexGrid1.Height <= 5925 Then
            MSFlexGrid1.Height = MSFlexGrid1.Rows * 400
        End If
    Loop
    TXTSQL = "select distinct (Razon) from Comprobante where Razon like('" & Trim(Text2.Text) & "%')order by Razon"
    Set REC = DB.Execute(TXTSQL)
    MSFlexGrid1.Visible = Not REC.EOF
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Consultarazon = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text2_LinkOpen(Cancel As Integer)
Consultarazon = False
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text4_LostFocus()
TXTSQL = "select * from comprobante where RIF='" & Text1.Text & "'and Nocontrol='" & Text4.Text & "'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
Label6.ForeColor = &HFF&
Else
Label6.ForeColor = &H808080
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text5_LostFocus()
TXTSQL = "select * from comprobante where RIF='" & Text1.Text & "'and NoFactura='" & Text5.Text & "'"
Set REC = DB.Execute(TXTSQL)
If Not REC.EOF Then
    Label7.ForeColor = &HFF&
    Else
    Label7.ForeColor = &H808080
End If
End Sub

Private Sub Text6_Change()
Text7_Change
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
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
          KeyAscii = 0
      End If
  End If
End Sub

Private Sub Text7_Change()
On Error Resume Next
MontoTotalFactura = CCur(Text7.Text)
BASEIMPONIBLE = ((MontoTotalFactura - CCur(Text6.Text)) / (1 + Alicuota))
TxtBaseImp.Text = Format(BASEIMPONIBLE, FORMATOMIL)
Iva = (MontoTotalFactura - CCur(Text6.Text)) - BASEIMPONIBLE


If MontoTotalFactura >= MinimoRetencion Then
    MontoRetenido = Format(Iva * (Val(Text8.Text) / 100), FORMATOMIL)
    Else
    MontoRetenido = 0
End If

Text15.Text = Format(MontoRetenido, FORMATOMIL)
Text9.Text = Format(MontoTotalFactura - MontoRetenido, FORMATOMIL)
Text16.Text = Format(Iva, FORMATOMIL)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
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
          KeyAscii = 0
      End If
  End If
End Sub

Private Sub Text8_Change()
Text7_Change
End Sub

Private Sub TxtNoComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxtNoComprobante.Text <> "" Then
    TXTSQL = "Select * from Comprobante where NoComprobante='" & Format(TxtNoComprobante.Text, "00000000") & "'"
    Set REC = DB.Execute(TXTSQL)
    If REC.EOF Then
        MsgBox "El Comprobante no existe", vbCritical
        Exit Sub
    Else
        Esconsulta = True
        If Not REC.EOF Then
            Dim rif_vl As String
            rif_vl = REC.Fields("Rif")
            Text1.Text = rif_vl
            Text1_KeyPress 13
        End If
        TXTSQL = "Select * from Comprobante where NoComprobante='" & Format(TxtNoComprobante.Text, "00000000") & "'"
        Set REC = DB.Execute(TXTSQL)
        Esconsulta = Not REC.EOF
        If Not REC.EOF Then
            TxtNoComprobante.Text = Val(REC.Fields("NoComprobante"))
            Text8.Text = REC.Fields("PorcentajeReten")
            Text3.Text = REC.Fields("FechaFactura")
            Text4.Text = REC.Fields("NoControl")
            Text5.Text = REC.Fields("NoFactura")
            Text6.Text = REC.Fields("Exento")
            Text7.Text = REC.Fields("MontoTotal")
            MaskEdBox3.Text = REC.Fields("FechaComprobante")
            Combo1.Text = REC.Fields("TipoTansaccion")
        End If
    End If
End If

End Sub

Public Sub UltimoComprobanteporProvee()
'Ultimo Comprobante
TXTSQL = "select * From Comprobante where " & Condicion_UC & " order by NoComprobante DESC"
Set REC = DB.Execute(TXTSQL)
If REC.EOF Then
    Exit Sub
End If
Text1.Text = REC.Fields("Rif")
Text2.Text = REC.Fields("Razon")
Text8.Text = REC.Fields("PorcentajeReten")
Text3.Text = REC.Fields("FechaFactura")
Text4.Text = REC.Fields("NoControl")
Text5.Text = REC.Fields("NoFactura")
Text6.Text = REC.Fields("Exento")
Text7.Text = REC.Fields("MontoTotal")
Combo1.Text = REC.Fields("TipoTansaccion")
Text12.Text = REC.Fields("Ndebito")
Text13.Text = REC.Fields("NCredito")
End Sub

Public Sub ActualizaComprobante()
    'UltimoCompRentecionIva
    TXTSQL = "SELECT * FROM CONFIGURACION WHERE DESCRIPCION='UltimoCompRentecionIva'"
    Set REC = DB.Execute(TXTSQL)
    If Not REC.EOF Then
        NoComprobante = REC.Fields("vALOR")
    End If
    TxtNoComprobante.Text = NoComprobante + 1
    'suma de retenciones
    TXTSQL = "SELECT SUM(IvaRetenido) from comprobante where FechaComprobante>cdate('" & UltimaDeclaracion & "') and FechaComprobante<cdate('" & proximaDeclaracion & "')"
    Set REC = DB.Execute(TXTSQL)
    If IsNull(REC.Fields(0)) Then
        Text11.Text = 0
    Else
        Text11.Text = Format(REC.Fields(0), FORMATOMIL)
    End If
End Sub
Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String

On Local Error GoTo errFunction
    
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
    
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
    
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
    
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
    
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path

Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString

End Function


Public Sub ActualizarCambios()
        TXTSQL = "UPDATE COMPROBANTE SET"
        TXTSQL = TXTSQL & " RIF='" & Text1.Text & "'"
        TXTSQL = TXTSQL & ", RAZON='" & Text2.Text & "'"
        'TXTSQL = TXTSQL & ", PorcentajeIva=" & Text8.Text & ""
        TXTSQL = TXTSQL & ", FechaFactura='" & Text3.Text & "'"
        TXTSQL = TXTSQL & ", NoControl='" & Text4.Text & "'"
        TXTSQL = TXTSQL & ", NoFactura='" & Text5.Text & "'"
        TXTSQL = TXTSQL & ", Exento='" & Text6.Text & "'"
        TXTSQL = TXTSQL & ", MontoTotal='" & Text7.Text & "'"
        TXTSQL = TXTSQL & ", FechaComprobante='" & MaskEdBox3.Text & "'"
        TXTSQL = TXTSQL & ", TipoTansaccion='" & Combo1.Text & "'"
        TXTSQL = TXTSQL & ", IvaRetenido='" & Text15.Text & "'"
        TXTSQL = TXTSQL & " WHERE NoComprobante='" & Format(TxtNoComprobante.Text, "00000000") & "'"
        DB.Execute (TXTSQL)
        MsgBox "Cambios Guardados", vbInformation
        Esconsulta = False
End Sub

Public Sub SumarIvadePeriodo()
TXTSQL = "SELECT SUM(IvaRetenido) from comprobante where Calendario= '" & TxtPeriodo.Text & "' and MontoTotal>=" & MinimoRetencion & ""
Set REC = DB.Execute(TXTSQL)
If IsNull(REC.Fields(0)) Then
    Text11.Text = 0
Else
    Text11.Text = Format(REC.Fields(0), FORMATOMIL)
End If
Label14.Caption = "Total " & "desde el " & UltimaDeclaracion + 1 & " Hasta el " & proximaDeclaracion
End Sub
Public Sub ConsultaMotosVendidas3()
Dim N_Factura As String
Dim contarencontrada As Integer
Dim Detalles As String
N_Factura = "No Factura"

Do Until N_Factura = "" Or N_Factura = "Q"
contarencontrada = 0
Detalles = ""
N_Factura = InputBox("Ingrese el No de Factura", "Busqueda de Facturas", N_Factura)
    If N_Factura <> "" Then
         TXTSQL = "select * from Comprobante where  Nofactura like('%" & N_Factura & "')"
         Set REC = DB.Execute(TXTSQL)
         If REC.EOF Then
            MsgBox "No se ha encontrado " & Chr(13) & N_Factura, vbCritical, "Busqueda de Facturas Registradas"
         Else
            BusqRetenciones.Show 1, Me
            If ConsulCompobuscar = True Then
                TXTSQL = "Select * from Comprobante where NoComprobante='" & Format(TXTSQL, "00000000") & "'"
                Set REC = DB.Execute(TXTSQL)
                Esconsulta = Not REC.EOF
                If Not REC.EOF Then
                    Text1.Text = REC.Fields("Rif")
                    TxtNoComprobante.Text = Val(REC.Fields("NoComprobante"))
                    Text2.Text = REC.Fields("Razon")
                    Text8.Text = REC.Fields("PorcentajeReten")
                    Text3.Text = REC.Fields("FechaFactura")
                    Text4.Text = REC.Fields("NoControl")
                    Text5.Text = REC.Fields("NoFactura")
                    Text6.Text = REC.Fields("Exento")
                    Text7.Text = REC.Fields("MontoTotal")
                    MaskEdBox3.Text = REC.Fields("FechaComprobante")
                    Combo1.Text = REC.Fields("TipoTansaccion")
                End If
                N_Factura = "Q"
            End If
         End If
    End If
Loop
End Sub
Public Sub ConsultaRetencionNoControl()
Dim N_Factura As String
Dim contarencontrada As Integer
Dim Detalles As String
N_Factura = "No Factura"

Do Until N_Factura = "" Or N_Factura = "Q"
contarencontrada = 0
Detalles = ""
N_Factura = InputBox("Ingrese el No de Control", "Busqueda de Facturas", N_Factura)
    If N_Factura <> "" Then
         TXTSQL = "select * from Comprobante where  NoControl like('%" & N_Factura & "')"
         Set REC = DB.Execute(TXTSQL)
         If REC.EOF Then
            MsgBox "No se ha encontrado " & Chr(13) & N_Factura, vbCritical, "Busqueda de Facturas Registradas"
         Else
            BusqRetenciones.Show 1, Me
            If ConsulCompobuscar = True Then
                TXTSQL = "Select * from Comprobante where NoComprobante='" & Format(TXTSQL, "00000000") & "'"
                Set REC = DB.Execute(TXTSQL)
                Esconsulta = Not REC.EOF
                If Not REC.EOF Then
                    Text1.Text = REC.Fields("Rif")
                    TxtNoComprobante.Text = Val(REC.Fields("NoComprobante"))
                    Text2.Text = REC.Fields("Razon")
                    Text8.Text = REC.Fields("PorcentajeReten")
                    Text3.Text = REC.Fields("FechaFactura")
                    Text4.Text = REC.Fields("NoControl")
                    Text5.Text = REC.Fields("NoFactura")
                    Text6.Text = REC.Fields("Exento")
                    Text7.Text = REC.Fields("MontoTotal")
                    MaskEdBox3.Text = REC.Fields("FechaComprobante")
                    Combo1.Text = REC.Fields("TipoTansaccion")
                End If
                N_Factura = "Q"
            End If
         End If
    End If
Loop
End Sub

Public Sub GenerarTxt()
Dim RifPropio As String
Dim PeriodoImpositivo As String
Dim FechaFactura As String
Dim TipoOperac As String
Dim TipoDocumento As String
Dim RifProvee As String
Dim NumeroDocumento As String
Dim NumeroControl_D As String
Dim MontoTotal_D As String
Dim BaseImponible_D As String
Dim MontoIvaRetenido_D As String
Dim NumeroDocumentoAfctado As String
Dim NumeroComprobanteRetencio As String
Dim MontoExentoIva  As String
Dim AlicuotaD As String
Dim Numeroexpediente_d As String


Dim CambiaMonto As String
Dim strLine As String
Dim I As Integer
CommonDialog1.FileName = "IVA " & NombreArchivoTxt & " " & Mes_Declaracion
CommonDialog1.Filter = "Archivos de texto txt|*.txt"
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then
   'salimos ya que no se ha seleccionado ningún archivo
   Exit Sub
End If
Screen.MousePointer = vbHourglass
    
    Open CommonDialog1.FileName For Output As #1
    Adodc1.Refresh
        With Adodc1.Recordset
            .MoveFirst
            'rif,periodo impositivo,fechafactura,tipodeoperacion,tipodocumento,rifproveedor,numero documento, numero control, montototal, base imponible, iva retenido, numero del documento afectado, numerocomprobante retencion, monto excento, alicuota, numero expediente
            Do While Not .EOF
                strLine = ""
                RifPropio = .Fields("Empresa")
                PeriodoImpositivo = .Fields("AnoPeriodoFiscal") & .Fields("MesPeriodoFiscal")
                FechaFactura = Format(.Fields("FechaFactura"), "yyyy-mm-dd")
                TipoOperac = "C"
                TipoDocumento = .Fields("TipoTansaccion")
                RifProvee = .Fields("Rif")
                NumeroDocumento = .Fields("NoFactura")
                NumeroControl_D = .Fields("NoControl")
                
                MontoExentoIva = Format(.Fields("Exento"), "#.00")
                MontoExentoIva = Replace(MontoExentoIva, ",", ".")
                
                MontoTotal_D = Format(.Fields("MontoTotal"), "#.00")
                MontoTotal_D = Replace(MontoTotal_D, ",", ".")
                
                BaseImponible_D = Format((.Fields("MontoTotal") - .Fields("Exento")) / (1 + (.Fields("PorcentajeIva") / 100)), "#.00")
                BaseImponible_D = Replace(BaseImponible_D, ",", ".")
                
                MontoIvaRetenido_D = Format(.Fields("IvaRetenido"), "#.00")
                MontoIvaRetenido_D = Replace(MontoIvaRetenido_D, ",", ".")
                
                NumeroDocumentoAfctado = 0
                If .Fields("Ndebito") <> 0 Or .Fields("Ncredito") <> 0 Then
                    NumeroDocumentoAfctado = .Fields("NoFactura")
                End If
                NumeroComprobanteRetencio = .Fields("AnoPeriodoFiscal") & .Fields("MesPeriodoFiscal") & .Fields("NoComprobante")
                
                MontoExentoIva = Format(.Fields("Exento"), "0.00")
                MontoExentoIva = Replace(MontoExentoIva, ",", ".")
                
                AlicuotaD = .Fields("PorcentajeIva")
                Numeroexpediente_d = 0
                CambiaMonto = Replace(CambiaMonto, ",", "")
                strLine = strLine & "" & RifPropio & vbTab & PeriodoImpositivo & vbTab & FechaFactura & vbTab & TipoOperac & vbTab & TipoDocumento & vbTab & RifProvee & vbTab & NumeroDocumento & vbTab & NumeroControl_D & vbTab & MontoTotal_D & vbTab & BaseImponible_D & vbTab & MontoIvaRetenido_D & vbTab & NumeroDocumentoAfctado & vbTab & NumeroComprobanteRetencio & vbTab & MontoExentoIva & vbTab & AlicuotaD & vbTab & Numeroexpediente_d & Chr(13)
                Print #1, strLine
                
                .MoveNext
            Loop
            .MoveFirst
        End With
    Close #1
    
    Screen.MousePointer = vbDefault
MsgBox "Archivo Generado Exitosamente", vbInformation
Exit Sub
errores:
MsgBox "No se genero la Nomina, Hubo errores en el proceso Verifique los datos" & Chr(13) & Err.Description
End Sub
Public Sub Guardaproveedor()
    TXTSQL = "select * from PROVEEDORES WHERE rif='" & Trim(Text1.Text) & "'"
    Set REC2 = DB.Execute(TXTSQL)
    If REC2.EOF Then
        TXTSQL = "INSERT INTO PROVEEDORES (RIF,RazonSocial,DireccionF) VALUES"
        TXTSQL = TXTSQL & "("
        TXTSQL = TXTSQL & "'" & Trim(Text1.Text) & "'"
        TXTSQL = TXTSQL & ",'" & Trim(Text2.Text) & "'"
        TXTSQL = TXTSQL & ",'" & Trim(Text10.Text) & "'"
        TXTSQL = TXTSQL & ")"
        DB.Execute (TXTSQL)
    Else
        If D_fiscal_g Then
            TXTSQL = "update PROVEEDORES set DireccionF='" & Text10.Text & "' where rif='" & Trim(Text1.Text) & "'"
            DB.Execute (TXTSQL)
        End If
    End If
End Sub
