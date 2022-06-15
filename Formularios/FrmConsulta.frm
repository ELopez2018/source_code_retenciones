VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmConsulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Impresion de Comprobantes"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   10500
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox TxtRazonSoc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2850
      TabIndex        =   24
      Top             =   450
      Width           =   7185
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir Todas"
      Height          =   705
      Left            =   7410
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4440
      Width           =   2565
   End
   Begin VB.Frame Frame3 
      Caption         =   "Por Numero de  Factura"
      Height          =   1125
      Left            =   2910
      TabIndex        =   12
      Top             =   3210
      Width           =   7125
      Begin VB.CommandButton Command7 
         Caption         =   "Ver"
         Height          =   585
         Left            =   5895
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   330
         Width           =   1125
      End
      Begin VB.TextBox TxtNofac 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   60
         TabIndex        =   16
         Top             =   270
         Width           =   3225
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir"
         Height          =   585
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   330
         Width           =   1125
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Sin Firma"
         Height          =   285
         Left            =   3390
         TabIndex        =   14
         Top             =   240
         Width           =   1035
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Con Firma"
         Height          =   285
         Left            =   3390
         TabIndex        =   13
         Top             =   540
         Width           =   1035
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No Factura"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1275
         TabIndex        =   17
         Top             =   810
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fecha de Comprobantes"
      Height          =   1125
      Left            =   2910
      TabIndex        =   7
      Top             =   2085
      Width           =   7125
      Begin VB.ComboBox CboAno 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2220
         TabIndex        =   25
         Text            =   "Año"
         Top             =   330
         Width           =   1125
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Ver"
         Height          =   585
         Left            =   5895
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   270
         Width           =   1125
      End
      Begin VB.ComboBox CboMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   11
         Text            =   "Mes"
         Top             =   330
         Width           =   2025
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir"
         Height          =   585
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   1125
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Sin Firma"
         Height          =   285
         Left            =   3420
         TabIndex        =   9
         Top             =   240
         Width           =   1005
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Con Firma"
         Height          =   285
         Left            =   3420
         TabIndex        =   8
         Top             =   540
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Por Fecha de Factura"
      Height          =   1125
      Left            =   2910
      TabIndex        =   1
      Top             =   960
      Width           =   7125
      Begin VB.CommandButton Command5 
         Caption         =   "Ver"
         Height          =   585
         Left            =   5925
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   270
         Width           =   1125
      End
      Begin MSMask.MaskEdBox MskDesdeNFact 
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yy"
         Mask            =   "##/##/##"
         PromptChar      =   "-"
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Con Firma"
         Height          =   285
         Left            =   3420
         TabIndex        =   6
         Top             =   540
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sin Firma"
         Height          =   285
         Left            =   3420
         TabIndex        =   5
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Imprimir"
         Height          =   585
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   270
         Width           =   1125
      End
      Begin MSMask.MaskEdBox MskHastaNFact 
         Height          =   435
         Left            =   1860
         TabIndex        =   20
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yy"
         Mask            =   "##/##/##"
         PromptChar      =   "-"
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   600
         TabIndex        =   3
         Top             =   795
         Width           =   465
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2400
         TabIndex        =   2
         Top             =   780
         Width           =   420
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   300
      TabIndex        =   0
      Text            =   "Rif"
      Top             =   450
      Width           =   2565
   End
End
Attribute VB_Name = "FrmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean

Option Explicit

Private Sub CboAno_Click()
On Error Resume Next
Dim Primer As String
Dim Ultimo As String
Dim MES_VL As String
 MES_VL = ""
Select Case CboMes.Text
    Case Is = "ENERO"
    MES_VL = "01"
    Case Is = "FEBRERO"
    MES_VL = "02"
    Case Is = "MARZO"
    MES_VL = "03"
    Case Is = "ABRIL"
    MES_VL = "04"
    Case Is = "MAYO"
    MES_VL = "05"
    Case Is = "JUNIO"
    MES_VL = "06"
    Case Is = "JULIO"
    MES_VL = "07"
    Case Is = "AGOSTO"
    MES_VL = "08"
    Case Is = "SEPTIEMBRE"
    MES_VL = "09"
    Case Is = "OCTUBRE"
    MES_VL = "10"
    Case Is = "NOVIEMBRE"
    MES_VL = "11"
    Case Is = "DICIEMBRE"
    MES_VL = "12"
End Select

If MES_VL = "" Then
    MsgBox "No se reconoce la fecha", vbExclamation
End If

MES_VL = "01/" & MES_VL & "/" & CboAno
Primer = DateSerial(Year(MES_VL), Month(MES_VL) + 0, 1)   'Primer Dia del Mes
Ultimo = DateSerial(Year(MES_VL), Month(MES_VL) + 1, 0) 'Ultimo dia del mes
VG_Desde = Primer
VG_Hasta = Ultimo
End Sub

Private Sub CboMes_Click()
On Error Resume Next
Dim Primer As String
Dim Ultimo As String
Dim MES_VL As String
Select Case CboMes.Text
    Case Is = "ENERO"
    MES_VL = "01"
    Case Is = "FEBRERO"
    MES_VL = "02"
    Case Is = "MARZO"
    MES_VL = "03"
    Case Is = "ABRIL"
    MES_VL = "04"
    Case Is = "MAYO"
    MES_VL = "05"
    Case Is = "JUNIO"
    MES_VL = "06"
    Case Is = "JULIO"
    MES_VL = "07"
    Case Is = "AGOSTO"
    MES_VL = "08"
    Case Is = "SEPTIEMBRE"
    MES_VL = "09"
    Case Is = "OCTUBRE"
    MES_VL = "10"
    Case Is = "NOVIEMBRE"
    MES_VL = "11"
    Case Is = "DICIEMBRE"
    MES_VL = "12"
End Select
MES_VL = "01/" & MES_VL & "/" & CboAno
Primer = DateSerial(Year(MES_VL), Month(MES_VL) + 0, 1)   'Primer Dia del Mes
Ultimo = DateSerial(Year(MES_VL), Month(MES_VL) + 1, 0) 'Ultimo dia del mes
VG_Desde = Primer
VG_Hasta = Ultimo
End Sub

Private Sub Combo1_Click()
If Combo1.Text <> "" Then
    TXTSQL = "SELECT * FROM PROVEEDORES WHERE RIF='" & Combo1.Text & "'"
    Set REC = DB.Execute(TXTSQL)
    TxtRazonSoc.Text = REC.Fields("RazonSocial")
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Or Combo1.Text = "Ingrese Rif" Then
    MsgBox "Seleccione un Proveedor", vbExclamation
    Exit Sub
End If

Dim TipoComprobante As String
Dim direcciondestinocarpeta As String
Dim Periodos_Consultado  As String

If Check1.Value = 1 Then
    TipoComprobante = "ComproRetencion.rpt"
ElseIf Check2.Value = 1 Then
    TipoComprobante = "ComproRetencionCF.rpt"
Else
    MsgBox "Por favor selecciones que tipo de Comprobante quiere" & Chr(13) & "Sin Firma o Con Firma", vbCritical
    Exit Sub
End If
If MsgBox("¿Desea Imprimir las Retenciones del Proveedor?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If

RifProveedor = Combo1.Text
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE FechaComprobante>=cdate('" & MskDesdeNFact.Text & "') and FechaComprobante<=cdate('" & MskHastaNFact.Text & "') and Rif='" & RifProveedor & "'"
    Set REC1 = DB.Execute(TXTSQL)
    Do Until REC1.EOF
        'Abrir el reporte
        Screen.MousePointer = vbHourglass
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\" & TipoComprobante, 1)
        TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & REC1.Fields("NoComprobante") & "'"
        Set REC = DB.Execute(TXTSQL)
        If Not REC.EOF Then
            Periodos_Consultado = REC.Fields("Calendario")
            Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\" & TipoComprobante, 1)
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
If Combo1.Text = "" Or Combo1.Text = "Ingrese Rif" Then
    MsgBox "Seleccione un Proveedor", vbExclamation
    Exit Sub
End If

Dim TipoComprobante As String
Dim direcciondestinocarpeta As String
Dim Periodos_Consultado  As String

If MsgBox("¿Desea Continuar con la Impresion de Los Comprobantes?", vbYesNo + vbQuestion, "Impresion de Comprobante") = vbNo Then
    Exit Sub
End If
If Check1.Value = 1 Then
TipoComprobante = "ComproRetencion.rpt"

ElseIf Check2.Value = 1 Then
    TipoComprobante = "ComproRetencionCF.rpt"
Else
    MsgBox "Por favor selecciones que tipo de Comprobante quiere" & Chr(13) & "Sin Firma o Con Firma", vbCritical
    Exit Sub
End If

RifProveedor = Combo1.Text
TXTSQL = "SELECT * FROM COMPROBANTE WHERE FechaComprobante>=cdate('" & VG_Desde & "') and FechaComprobante<=cdate('" & VG_Hasta & "') and Rif='" & RifProveedor & "'"
Set REC1 = DB.Execute(TXTSQL)
    If REC.EOF Then
        MsgBox "No se encontraron Registros", vbInformation
        Exit Sub
    End If
Do Until REC1.EOF
    Screen.MousePointer = vbHourglass
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\" & TipoComprobante, 1)
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & REC1.Fields("NoComprobante") & "'"
    Set REC = DB.Execute(TXTSQL)
    If Not REC.EOF Then
        Periodos_Consultado = REC.Fields("Calendario")
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\" & TipoComprobante, 1)
        crReport.RecordSelectionFormula = "{comprobante.Nocomprobante}='" & REC1.Fields("Nocomprobante") & "'"
        crReport.PrintOut False
        Screen.MousePointer = vbDefault
    End If
    REC1.MoveNext
Loop
End Sub

Private Sub Command3_Click()
If Combo1.Text = "" Or Combo1.Text = "Ingrese Rif" Then
    MsgBox "Seleccione un Proveedor", vbExclamation
    Exit Sub
End If
Dim TipoComprobante As String
Dim direcciondestinocarpeta As String
Dim Periodos_Consultado  As String

If MsgBox("¿Desea Imprimir las retenciones del la quincena del Proveedor?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If
If Check1.Value = 1 Then
TipoComprobante = "ComproRetencion.rpt"
ElseIf Check1.Value = 1 Then
TipoComprobante = "ComproRetencionCF.rpt"
Else
    MsgBox "Por favor selecciones que tipo de Comprobante quiere, Sin Firma o Con Firma", vbCritical
End If
End Sub

Private Sub Command5_Click()
If Combo1.Text = "" Or Combo1.Text = "Ingrese Rif" Then
    MsgBox "Seleccione un Proveedor", vbExclamation
    Exit Sub
End If
TipoDeConsulta = "ConsultaListaReportexFechaFact"
RifProveedor = Combo1.Text
VG_Desde = MskDesdeNFact.Text
VG_Hasta = MskHastaNFact.Text
FmReporte.Show 1, Me
End Sub

Private Sub Command6_Click()

If Combo1.Text = "" Or Combo1.Text = "Ingrese Rif" Then
MsgBox "Seleccione un Proveedor", vbExclamation
Exit Sub
End If

TipoDeConsulta = "ConsultaListaReportexFechaFact"
RifProveedor = Combo1.Text
FmReporte.Show 1, Me
End Sub

Private Sub Command7_Click()
If Combo1.Text = "" Or Combo1.Text = "Ingrese Rif" Then
    MsgBox "Seleccione un Proveedor", vbExclamation
    Exit Sub
End If
TipoDeConsulta = "ConsultadeUnaFactura"
RifProveedor = Combo1.Text
NoFacturaPaConsultar = TxtNofac.Text
FmReporte.Show 1, Me
End Sub

Private Sub Command8_Click()
MsgBox "Desde" & VG_Desde & Chr(13) & "Hasta " & VG_Hasta
End Sub

Private Sub Form_Load()
LlenarComboRif
LlenarComboMes
LlenarComboAno
End Sub

Public Sub LlenarComboRif()
TXTSQL = "SELECT * FROM PROVEEDORES ORDER bY razonsocial"
Set REC = DB.Execute(TXTSQL)
Combo1.Clear
TxtRazonSoc.Clear
Do Until REC.EOF
    Combo1.AddItem REC.Fields("rif")
    TxtRazonSoc.AddItem REC.Fields("razonsocial")
    REC.MoveNext
Loop
Combo1.Text = "Ingrese Rif"
End Sub

Private Sub TxtRazonSoc_Click()
If TxtRazonSoc.Text <> "" Then
    TXTSQL = "SELECT * FROM PROVEEDORES WHERE RazonSocial='" & TxtRazonSoc.Text & "'"
    Set REC = DB.Execute(TXTSQL)
    Combo1.Text = REC.Fields("rif")
End If
End Sub

Public Sub LlenarComboMes()
Dim I As Integer
Dim fecha_vl As String
I = 1
Do Until I > 12
    fecha_vl = "01/" & I & "/2014"
    CboMes.AddItem UCase(Format(fecha_vl, "MMMM"))
    I = I + 1
Loop
End Sub

Public Sub LlenarComboAno()
Dim I As Integer
Dim Ano_Vl As Integer
I = 1
Ano_Vl = Format(Date, "YYYY")
Do Until I > 4
    CboAno.AddItem Ano_Vl
    Ano_Vl = Ano_Vl - 1
    I = I + 1
Loop
End Sub

Private Sub TxtRazonSoc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
