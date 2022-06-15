VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CalendarioFiscal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Calendario Fiscal"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   270
      TabIndex        =   8
      Top             =   930
      Visible         =   0   'False
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483643
      Appearance      =   1
      StartOfWeek     =   41746434
      CurrentDate     =   41095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Periodo Fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   8475
      Begin VB.CommandButton Command1 
         Height          =   555
         Left            =   1965
         Picture         =   "CalendarioFiscal.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   323
         Width           =   585
      End
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   405
         Left            =   4935
         TabIndex        =   3
         Top             =   398
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   714
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   405
         Left            =   4260
         TabIndex        =   4
         Top             =   398
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   714
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   405
         Left            =   180
         TabIndex        =   7
         Top             =   398
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   714
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   405
         Left            =   7320
         TabIndex        =   9
         Top             =   398
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   714
         _Version        =   393216
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quincena Nº"
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
         Left            =   5820
         TabIndex        =   10
         Top             =   450
         Width           =   1500
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
         Left            =   2550
         TabIndex        =   5
         Top             =   450
         Width           =   1710
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar"
      Height          =   615
      Left            =   7350
      TabIndex        =   1
      Top             =   7170
      Width           =   1245
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5925
      Left            =   120
      TabIndex        =   0
      Top             =   1110
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   10451
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   24
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Id"
         Caption         =   "No"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FechadeDeclaracion"
         Caption         =   "Declaración"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "TotalaDeclarar"
         Caption         =   "Total Bs."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Observacion"
         Caption         =   "Observacion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   555
      Left            =   8640
      Top             =   780
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   979
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
      RecordSource    =   "SELECT * FROM CALENDARIOFISCAL"
      Caption         =   "Calendario"
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
End
Attribute VB_Name = "CalendarioFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
MonthView1.Visible = True
End Sub

Private Sub Command2_Click()
Dim FECHA As String
FECHA = MaskEdBox1.Text
If Not IsDate(MaskEdBox1.Text) Then
    MsgBox "Error en Mes Fiscal"
    Exit Sub
End If

If Not IsNumeric(MaskEdBox5.Text) Then
    MsgBox "Error en Año Fiscal"
    Exit Sub
End If
If Not IsNumeric(MaskEdBox3.Text) Then
    MsgBox "Error en Quincena"
    Exit Sub
End If
If IsDate(FECHA) Then
    TXTSQL = "INSERT INTO CALENDARIOFISCAL"
    TXTSQL = TXTSQL & "(FechadeDeclaracion,MesFiscal,AnoFiscal,Periodo"
    TXTSQL = TXTSQL & ")"
    TXTSQL = TXTSQL & " VALUES "
    TXTSQL = TXTSQL & "("
    TXTSQL = TXTSQL & "cdate('" & FECHA & "')"
    TXTSQL = TXTSQL & ",'" & MaskEdBox2.Text & "'"
    TXTSQL = TXTSQL & ",'" & MaskEdBox5.Text & "'"
    TXTSQL = TXTSQL & "," & MaskEdBox3.Text & "-" & MaskEdBox5.Text & ""
    TXTSQL = TXTSQL & ")"
    DB.Execute (TXTSQL)
    MsgBox "Se guardo sin errores"
Else
MsgBox "Verifique que la fecha tiene el Formato Correcto, dd/mm/yyyy"
End If
Adodc1.Refresh
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = ConexionDatasABaseD
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
MaskEdBox1.Text = DateClicked
MonthView1.Visible = Not MonthView1.Visible
End Sub
