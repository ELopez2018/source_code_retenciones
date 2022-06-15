VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Begin VB.Form BusqRetenciones 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Busqueda Segun No de Factura"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14295
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5835
      Left            =   240
      TabIndex        =   0
      Top             =   285
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   10292
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbskfr2.Skinner Skinner1 
      Left            =   4290
      Top             =   330
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
      SysEnableSkinCaption=   "Habilitar &Skin"
      SysDisableSkinCaption=   "Deshabilitar &Skin"
   End
   Begin VB.Image ImageMod 
      Height          =   480
      Left            =   3330
      Picture         =   "BusqRetenciones.frx":0000
      Top             =   2880
      Width           =   450
   End
   Begin VB.Image ImageImprimir 
      DataField       =   "&H00C0FFC0&"
      Height          =   480
      Left            =   1800
      Picture         =   "BusqRetenciones.frx":03C9
      Top             =   2850
      Width           =   480
   End
   Begin VB.Image ImagePdf 
      DataSource      =   "&H00FFFFFF&"
      Height          =   450
      Left            =   1290
      Picture         =   "BusqRetenciones.frx":0848
      Top             =   2850
      Width           =   390
   End
   Begin VB.Label lblDestino 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   1500
      Width           =   480
   End
End
Attribute VB_Name = "BusqRetenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'variables reporte
Private crApp As New CRAXDRT.Application
Private crReport As New CRAXDRT.Report
Private mflgContinuar As Boolean
Dim Quincena As Integer

Private mstrParametro1 As String
Private mlngParametro2 As Long


Dim NoComprobanteSel As String

Private Sub Form_Load()
BusqRetenciones.Caption = TituloVentana
Forma_Grid_Buscar
Set Skinner1.Forms = Forms
Set REC = DB.Execute(TXTSQL)
Do Until REC.EOF
    MSFlexGrid1.AddItem REC.Fields("NoComprobante") & vbTab & REC.Fields("FechaComprobante") & vbTab & REC.Fields("NoFactura") & vbTab & REC.Fields("NoControl") & vbTab & REC.Fields("FechaFactura") & vbTab & REC.Fields("Razon") & vbTab & Format(REC.Fields("MontoTotal"), FORMATOMIL)
    MSFlexGrid1.Col = 7
     MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
    MSFlexGrid1.RowHeightMin = ImagePdf.Height + ScaleY(2, 3, 1)
    MSFlexGrid1.ColWidth(7) = MSFlexGrid1.RowHeightMin
    Set MSFlexGrid1.CellPicture = ImagePdf.Picture
    
    MSFlexGrid1.Col = 8
    MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
    MSFlexGrid1.RowHeightMin = ImageImprimir.Height + ScaleY(2, 3, 1)
    MSFlexGrid1.ColWidth(8) = MSFlexGrid1.RowHeightMin
    Set MSFlexGrid1.CellPicture = ImageImprimir.Picture
    
    MSFlexGrid1.Col = 9
    MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
    MSFlexGrid1.RowHeightMin = ImageMod.Height + ScaleY(2, 3, 1)
    MSFlexGrid1.ColWidth(9) = MSFlexGrid1.RowHeightMin
    Set MSFlexGrid1.CellPicture = ImageMod.Picture
    REC.MoveNext
Loop
End Sub
Public Sub Forma_Grid_Buscar()
     With MSFlexGrid1
        .Clear
        .Rows = 1
        .Cols = 7
        .FormatString = "^Comprobante|^Fecha C.|^No Factura|^No Control|^Fecha F.|Razon Social|Monto Total|PDF|IMP|BUS"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = 1100
        .ColWidth(5) = 4900
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
    End With
End Sub

Private Sub MSFlexGrid1_DblClick()
Dim LARGO As String
Dim ANCH As String
NoComprobanteSel = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0)
If MSFlexGrid1.ColSel = 7 Then
    If MsgBox("¿Desea exportar a PDF el Comprobante No " & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0), vbQuestion + vbYesNo) = vbYes Then
        ImprimirPdf
    End If
ElseIf MSFlexGrid1.ColSel = 8 Then
    If MsgBox("¿Desea Imprimir el Comprobante No " & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0), vbQuestion + vbYesNo) = vbYes Then
        ImprimirImpresora
    End If
ElseIf MSFlexGrid1.ColSel = 9 Then
    If MsgBox("¿Desea Buscar el Comprobante No " & MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0), vbQuestion + vbYesNo) = vbYes Then
    TXTSQL = Val(MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 0))
    ConsulCompobuscar = True
    Unload Me
    End If
End If
End Sub
Private Sub ImprimirPdf()
On Error GoTo ErrHandler
    Dim Periodos_Consultado As String
    lblDestino.Caption = DireccionCarpetaExportar
    If lblDestino.Caption <> "" Then
        lblDestino.Caption = Buscar_Carpeta(" ... Seleccione una carpeta ")
        DireccionCarpetaExportar = lblDestino.Caption
    End If
    lblDestino.Caption = DireccionCarpetaExportar
    NocomproBanteConsulta = NoComprobanteSel
    'Abrir el reporte
    Screen.MousePointer = vbHourglass
    Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionCF.rpt", 1)
    crReport.RecordSelectionFormula = "{comprobante.nocomprobante}='" & Format(NoComprobanteSel, "00000000") & "'"
    
    TXTSQL = "SELECT * FROM COMPROBANTE WHERE NoComprobante='" & Format(NoComprobanteSel, "00000000") & "'"
    Set REC = DB.Execute(TXTSQL)
    Periodos_Consultado = REC.Fields("Calendario")
    
    If Dir(lblDestino.Caption & "\Quincena " & Periodos_Consultado, vbDirectory) = "" Then
        Call MkDir(lblDestino.Caption & "\Quincena " & Periodos_Consultado)
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

Public Sub ImprimirImpresora()
 '///////////////////////////////////////////////NUEV0
        
        Dim numerocopias As Integer
        Dim repuesta_l As String
        repuesta_l = "empezar"
        numerocopias = 0
        Do Until numerocopias > 0 Or UCase(repuesta_l) = "SALIR" Or repuesta_l = ""
            repuesta_l = InputBox("No de Compias a Imprimir", "Comprobante Nº " & NoComprobanteSel, 1)
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
            crReport.RecordSelectionFormula = "{Comprobante.NoComprobante}='" & Format(NoComprobanteSel, "00000000") & "'"
            crReport.PrintOut (False)
            Screen.MousePointer = vbDefault
        I_L = I_L + 1
        Loop
        '///////////////////////////////////////////////FIN NUEVO
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

