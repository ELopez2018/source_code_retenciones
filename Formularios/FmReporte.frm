VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form FmReporte 
   Caption         =   "Comprobante de Retencion"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer 
      Height          =   6825
      Left            =   1620
      TabIndex        =   0
      Top             =   930
      Width           =   8235
      lastProp        =   500
      _cx             =   14526
      _cy             =   12039
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "FmReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean
Private mstrParametro1 As String
Private mlngParametro2 As Long


Private Sub Form_Load()
    Select Case TipoDeConsulta
    Case Is = "PORCOMPROBANTE"
        ConsultaPorComprobante
    Case Is = "COMPROBANTESDELPROVEEDOR"
        ConsultadelosComprobantes
    Case Is = "PORQUINCENA"
        ConsultaporQuincena
    Case Is = "RESUMEN"
        ResumenQuincenal
    Case Is = "FACTURASDELPROVEEDOR"
        FacturasdelProveedor
    Case Is = "BUSCARFACTURA"
        FacturasdelProveedoraBuscar
    Case Is = "ConsultaPorComprobante_F"
        ConsultaPorComprobante_F
    Case Is = "LibroCompras"
        LibroCompras
    Case Is = "ComprobantesISLR"
        ComprobantesISLR
    Case Is = "ConsultaListaReportexFechaFact"
        ConsultaListaReportexFechaFact
    Case Is = "ConsultadeUnaFactura"
        ConsultadeUnaFactura
    End Select
End Sub

Private Sub Form_Resize()
    CRViewer.Top = 0
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.Width = ScaleWidth
    
End Sub

Public Sub ConsultaPorComprobante()
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    On Error GoTo ErrHandler
     'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.nocomprobante}='" & Format(NocomproBanteConsulta, "00000000") & "'"
    ' Parametros del reporte
    Set crParamDefs = crReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Parametro1"
               crParamDef.AddCurrentValue (mstrParametro1)
            Case "Parametro2"
               crParamDef.AddCurrentValue (mlngParametro2)
        End Select
    Next
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
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

Public Sub ConsultadelosComprobantes()
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    On Error GoTo ErrHandler
     'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ListaRetencion.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.RIF}='" & RifProveedor & "' and {comprobante.calendario}='" & PeriodoConsulta & "'"
    ' Parametros del reporte
    Set crParamDefs = crReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Parametro1"
               crParamDef.AddCurrentValue (mstrParametro1)
            Case "Parametro2"
               crParamDef.AddCurrentValue (mlngParametro2)
        End Select
    Next
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
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

Public Sub ConsultaporQuincena()
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    On Error GoTo ErrHandler
     'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ListaRetencion.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.Calendario}='" & RifProveedor & "'"
    ' Parametros del reporte
    Set crParamDefs = crReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Parametro1"
               crParamDef.AddCurrentValue (mstrParametro1)
            Case "Parametro2"
               crParamDef.AddCurrentValue (mlngParametro2)
        End Select
    Next
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    'crReport.ExportOptions.FormatType = crEFTPortableDocFormat
    'crReport.ExportOptions.DestinationType = crEDTDiskFile
    'crReport.ExportOptions.DiskFileName = "C:\reporte 1.pdf"
   ' crReport.ExportOptions.PDFExportAllPages = True
    'crReport.Export (False)
    CRViewer.ViewReport
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
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

Public Sub ResumenQuincenal()
Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    On Error GoTo ErrHandler
     'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ListaRetencionTotales.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.Calendario}=" & RifProveedor & ""
    ' Parametros del reporte
    Set crParamDefs = crReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Parametro1"
               crParamDef.AddCurrentValue (mstrParametro1)
            Case "Parametro2"
               crParamDef.AddCurrentValue (mlngParametro2)
        End Select
    Next
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
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

Public Sub FacturasdelProveedor()
    
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    On Error GoTo ErrHandler
     'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ListaRetencion.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.rif}='" & RifProveedor & "'"
    ' Parametros del reporte
    Set crParamDefs = crReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Parametro1"
               crParamDef.AddCurrentValue (mstrParametro1)
            Case "Parametro2"
               crParamDef.AddCurrentValue (mlngParametro2)
        End Select
    Next
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
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
Public Sub FacturasdelProveedoraBuscar()
    Screen.MousePointer = vbHourglass
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ListaRetencion.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.NOFACTURA}='" & NoFacturaPaConsultar & "'"
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
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
Public Sub ConsultadeUnaFactura()
    Screen.MousePointer = vbHourglass
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ListaRetencion.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.NOFACTURA}='" & NoFacturaPaConsultar & "' and {comprobante.rif}='" & RifProveedor & "'"
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
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
Public Sub ConsultaPorComprobante_F()
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    On Error GoTo ErrHandler
     'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencion.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.RIF}='" & RifProveedor & "' and {comprobante.Nofactura}='" & NocomproBanteConsulta & "'"
    ' Parametros del reporte
    Set crParamDefs = crReport.ParameterFields
    For Each crParamDef In crParamDefs
        Select Case crParamDef.ParameterFieldName
            Case "Parametro1"
               crParamDef.AddCurrentValue (mstrParametro1)
            Case "Parametro2"
               crParamDef.AddCurrentValue (mlngParametro2)
        End Select
    Next
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
    Screen.MousePointer = vbDefault
    Set crParamDefs = Nothing
    Set crParamDef = Nothing
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
Public Sub LibroCompras()
Dim Fecha_V As String
    On Error GoTo ErrHandler
     'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\Libro de Compras.rpt", 1)
    crReport.RecordSelectionFormula = "{Comprobante.MesPeriodoFiscal}='" & MesFiscaLibComp & "' and {Comprobante.AnoPeriodoFiscal}='" & AnoFiscaLibComp & "'"
    Fecha_V = "01/" & MesFiscaLibComp & "/" & AnoFiscaLibComp & ""
    Fecha_V = Format(Fecha_V, "mmmm")
    Fecha_V = UCase(Fecha_V)
    Fecha_V = "'" & Fecha_V & " DE " & AnoFiscaLibComp & "'"
    crReport.FormulaFields(3).Text = Fecha_V
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
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

Public Sub ComprobantesISLR()
Dim Fecha_V As String
    On Error GoTo ErrHandler
     'Abrir el reporte
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ListaRetencionIsrl.rpt", 1)
    crReport.RecordSelectionFormula = "{ISLR.Mes}='" & MesFiscaLibComp & "' and {ISLR.Ano}='" & AnoFiscaLibComp & "'"
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
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

Public Sub ConsultaListaReportexFechaFact()
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    mflgContinuar = True
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ListaRetencion.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.RIF}='" & RifProveedor & "' and {comprobante.FechaComprobante}>=cdate('" & VG_Desde & "') and {comprobante.FechaComprobante}<=cdate('" & VG_Hasta & "')"
    CRViewer.ReportSource = crReport
    CRViewer.DisplayGroupTree = False
    CRViewer.ViewReport
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
