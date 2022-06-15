VERSION 5.00
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Begin VB.Form FrmIslrporMes 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Comprobantes de ISLR"
   ClientHeight    =   3105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   4860
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
      Top             =   345
      Width           =   4425
      Begin VB.CommandButton CmdExportPDF 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "ImprISLRmes.frx":0000
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
         Left            =   1530
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ImprISLRmes.frx":119B
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "1"
         ToolTipText     =   "Exportar a PDF"
         Top             =   1080
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
         Height          =   1155
         Left            =   2820
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ImprISLRmes.frx":185B
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Imprime en la impresora las retenciones del Proveedor"
         Top             =   1080
         Width           =   1260
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         DisabledPicture =   "ImprISLRmes.frx":20B2
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
         Left            =   240
         MaskColor       =   &H00E0E0E0&
         Picture         =   "ImprISLRmes.frx":28F4
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
         Left            =   1950
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
         Left            =   690
         TabIndex        =   1
         Text            =   "Mes"
         Top             =   420
         Width           =   1245
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
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   9135
   End
End
Attribute VB_Name = "FrmIslrporMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report

Private Sub CmdExportPDF_Click()
On Error GoTo ErrHandler
    If MsgBox("Desea Exportar los Comprobantes correspondiente al " & CboMes.Text & " de " & Cboano.Text, vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
        Exit Sub
    End If
    
    Dim Periodos_Consultado As String
    If lblDestino.Caption = "....Seleccione Una carpeta" Then
        MsgBox "....Seleccione Una carpeta", vbCritical
    End If
    lblDestino.Caption = Buscar_Carpeta(" ... Seleccione una carpeta ")
    TXTSQL = "SELECT * FROM ISLR WHERE mes='" & CboMes.Text & "' and ano='" & Cboano.Text & "'"
    Set REC1 = DB.Execute(TXTSQL)
    Dim NoComprobante As String
    Do Until REC1.EOF
        NoComprobante = Format(REC1.Fields("NoComprobanteISLR"), "00000000")
        
        'Abrir el reporte
        Screen.MousePointer = vbHourglass
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionISRL.rpt", 1)
        crReport.RecordSelectionFormula = "{ISLR.NoComprobanteISLR}='" & Format(NoComprobante, "00000000") & "'"
        
        TXTSQL = "SELECT * FROM ISLR WHERE NoComprobanteISLR='" & Format(NoComprobante, "00000000") & "'"
        Set REC = DB.Execute(TXTSQL)
        Periodos_Consultado = REC.Fields("MES") & "-" & REC.Fields("ANO")
        Dim Proveedor_ISLR As String
        If Dir(lblDestino.Caption & "\Islr Periodo " & Periodos_Consultado, vbDirectory) = "" Then
            Call MkDir(lblDestino.Caption & "\Islr Periodo " & Periodos_Consultado)
            MsgBox "La carpeta" & lblDestino.Caption & "\Islr Periodo " & Periodos_Consultado & " fue Creada"
        End If
        
        TXTSQL = "SELECT * FROM PROVEEDORES WHERE RIF='" & REC.Fields("RIF") & "'"
        Set REC2 = DB.Execute(TXTSQL)
        Proveedor_ISLR = REC2.Fields("RAZONSOCIAL")
        
        crReport.ExportOptions.FormatType = crEFTPortableDocFormat
        crReport.ExportOptions.DestinationType = crEDTDiskFile
        crReport.ExportOptions.DiskFileName = lblDestino.Caption & "\Islr Periodo " & Periodos_Consultado & "\ISLR-" & REC.Fields("NoComprobanteISLR") & "-" & REC.Fields("NoFactura") & "-" & Proveedor_ISLR & ".pdf"
        crReport.ExportOptions.PDFExportAllPages = True
        crReport.Export (False)
        Screen.MousePointer = vbDefault
    REC1.MoveNext
    Loop
    
    MsgBox "El Archivo a sido exportado Satisfactoriamente " & Chr(13) & lblDestino.Caption & "\Periodo " & Periodos_Consultado, vbInformation
    
Exit Sub
ErrHandler:
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", vbCritical + vbOKOnly
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub CmdImpImpresora_Click()
On Error GoTo ErrHandler
    Dim Periodos_Consultado As String
    If MsgBox("¿Desea Imprimir las Retenciones Correspondiente al  " & CboMes.Text & " de " & Cboano.Text & "?", vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
        Exit Sub
    End If
    
    TXTSQL = "SELECT * FROM ISLR WHERE mes='" & CboMes.Text & "' and ano='" & Cboano.Text & "'"
    Set REC1 = DB.Execute(TXTSQL)
    Dim NoComprobante As String
    Do Until REC1.EOF
        NoComprobante = REC1.Fields("NoComprobanteISLR")
        'Abrir el reporte
        Screen.MousePointer = vbHourglass
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionISRL.rpt", 1)
        crReport.RecordSelectionFormula = "{ISLR.NoComprobanteISLR}='" & NoComprobante & "'"
        crReport.PrintOut (False)
        Screen.MousePointer = vbDefault
    REC1.MoveNext
    Loop
Exit Sub
ErrHandler:
    If Err.Number = -2147206461 Then
        MsgBox "El archivo de reporte no se encuentra, restáurelo de los discos de instalación", vbCritical + vbOKOnly
    Else
        MsgBox Err.Description, vbCritical + vbOKOnly
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command6_Click()
On Error Resume Next
AnoFiscaLibComp = Cboano.Text
MesFiscaLibComp = CboMes.Text
TipoDeConsulta = "ComprobantesISLR"
FmReporte.Show 1, Me
End Sub

Private Sub Form_Load()
Set Skinner1.Forms = Forms
TXTSQL = "select distinct (MES) from ISLR"
    Set REC = DB.Execute(TXTSQL)
    CboMes.Clear
    Do Until REC.EOF
        CboMes.AddItem REC.Fields("MES")
        REC.MoveNext
    Loop


TXTSQL = "select distinct (ANO) from ISLR"
    Set REC = DB.Execute(TXTSQL)
    Cboano.Clear
    Do Until REC.EOF
        Cboano.AddItem REC.Fields("ANO")

        REC.MoveNext
    Loop
    
    CboMes.Text = Format(Date, "mm")
    Cboano.Text = Format(Date, "yyyy")

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

