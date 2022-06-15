VERSION 5.00
Begin VB.Form ImprimirAIntervaloISLR 
   Caption         =   "Impresion de Comprobantes"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   405
      Left            =   6420
      TabIndex        =   8
      Top             =   4020
      Width           =   495
   End
   Begin VB.CommandButton CmdImpPDF 
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
      Left            =   4410
      Picture         =   "ImprimirAIntervaloISLR.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Imprime en archivos PDF"
      Top             =   2415
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
      Left            =   5670
      MaskColor       =   &H00E0E0E0&
      Picture         =   "ImprimirAIntervaloISLR.frx":1FBA
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "1"
      ToolTipText     =   "Imprime en la impresora las retenciones del Proveedor"
      Top             =   2415
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Caption         =   "No Comprobantes"
      Height          =   2295
      Left            =   270
      TabIndex        =   0
      Top             =   60
      Width           =   6705
      Begin VB.TextBox TxtHasta 
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
         Height          =   675
         Left            =   3630
         TabIndex        =   2
         Top             =   480
         Width           =   2625
      End
      Begin VB.TextBox Txtdesde 
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
         Height          =   675
         Left            =   210
         TabIndex        =   1
         Top             =   480
         Width           =   2625
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   9
         Top             =   1380
         Width           =   6585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   195
         Left            =   4725
         TabIndex        =   4
         Top             =   1155
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         Height          =   195
         Left            =   1290
         TabIndex        =   3
         Top             =   1155
         Width           =   465
      End
   End
   Begin VB.Label lblDestino 
      Caption         =   "....Seleccione Una carpeta"
      Height          =   525
      Left            =   210
      TabIndex        =   7
      Top             =   3990
      Visible         =   0   'False
      Width           =   5865
   End
End
Attribute VB_Name = "ImprimirAIntervaloISLR"
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

Private Sub CmdImpImpresora_Click()

If MsgBox("Desea Imprimir las Retenciones desde " & Txtdesde.Text & " hasta " & TxtHasta.Text, vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
    Exit Sub
End If
Dim Periodos_Consultado  As String
    TXTSQL = "SELECT * FROM ISLR WHERE NoComprobanteISLR>=" & Txtdesde.Text & " and NoComprobante<=" & TxtHasta.Text
    Set REC1 = DB.Execute(TXTSQL)
    
    Do Until REC1.EOF
        'Abrir el reporte
        Screen.MousePointer = vbHourglass
        Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionISRL.rpt", 1)
        TXTSQL = "SELECT * FROM ISLR WHERE NoComprobanteISLR='" & REC1.Fields("NoComprobanteISLR") & "'"
        Set REC = DB.Execute(TXTSQL)
        If Not REC.EOF Then
            Set crReport = crApp.OpenReport(DireccionCarpetaReportes & "\ComproRetencionISRL.rpt", 1)
            crReport.RecordSelectionFormula = "{ISLR.NoComprobanteISLR}='" & REC1.Fields("NoComprobanteISLR") & "'"
            crReport.PrintOut (False)
            Screen.MousePointer = vbDefault
        End If
        REC1.MoveNext
    Loop
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

Private Sub CmdImpPDF_Click()
On Error GoTo ErrHandler
    Dim Periodos_Consultado As String
    'Set VLman_arch = New FileSystemObject
    If lblDestino.Caption = "....Seleccione Una carpeta" Then
        MsgBox "....Seleccione Una carpeta", vbCritical
        Command5_Click
    End If
    
    If MsgBox("Desea EXPORTAR los Comprobantes desde " & Txtdesde.Text & " hasta " & TxtHasta.Text, vbYesNo + vbQuestion, "Exportar Retenciones") = vbNo Then
        Exit Sub
    End If
    
    TXTSQL = "SELECT * FROM ISLR WHERE NoComprobanteISLR>='" & Format(Txtdesde.Text, "00000000") & "' and NoComprobanteISLR<='" & Format(TxtHasta.Text, "00000000") & "'"
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
        If Dir(lblDestino.Caption & "\FECHA " & Periodos_Consultado, vbDirectory) = "" Then
            'MsgBox "La carpeta no existe"
            Call MkDir(lblDestino.Caption & "\FECHA " & Periodos_Consultado)
            MsgBox "La carpeta Creada" & Chr(13) & lblDestino.Caption & "\PERIODO " & Periodos_Consultado
        End If
        TXTSQL = "SELECT * FROM PROVEEDORES WHERE RIF='" & REC.Fields("RIF") & "'"
        Set REC2 = DB.Execute(TXTSQL)
        Proveedor_ISLR = REC2.Fields("RAZONSOCIAL")
        
        crReport.ExportOptions.FormatType = crEFTPortableDocFormat
        crReport.ExportOptions.DestinationType = crEDTDiskFile
        crReport.ExportOptions.DiskFileName = lblDestino.Caption & "\FECHA " & Periodos_Consultado & "\ISLR-" & REC.Fields("NoComprobanteISLR") & "-" & REC.Fields("NoFactura") & "-" & Proveedor_ISLR & ".pdf"
        crReport.ExportOptions.PDFExportAllPages = True
        crReport.Export (False)
        Screen.MousePointer = vbDefault
    REC1.MoveNext
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

Private Sub Command5_Click()
lblDestino.Caption = Buscar_Carpeta(" ... Seleccione una carpeta ")
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

Private Sub Txtdesde_Change()
On Error GoTo errores
   TXTSQL = "SELECT SUM(IVARETENIDO) from islr where NoComprobanteISLR>='" & Format(Txtdesde.Text, "00000000") & "' and NoComprobanteISLR<='" & Format(TxtHasta.Text, "00000000") & "'"
    Set REC1 = DB.Execute(TXTSQL)
    Label3.Caption = Format(CCur(REC1(0)) + 0, FORMATOMIL)
Exit Sub
errores:
Label3.Caption = Format(0, FORMATOMIL)
End Sub

Private Sub TxtHasta_Change()
On Error GoTo errores
   TXTSQL = "SELECT SUM(IVARETENIDO) from islr where NoComprobanteISLR>='" & Format(Txtdesde.Text, "00000000") & "' and NoComprobanteISLR<='" & Format(TxtHasta.Text, "00000000") & "'"
    Set REC1 = DB.Execute(TXTSQL)
    Label3.Caption = Format(CCur(REC1(0)) + 0, FORMATOMIL)
Exit Sub
errores:
Label3.Caption = Format(0, FORMATOMIL)
End Sub
