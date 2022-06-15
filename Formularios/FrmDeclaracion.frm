VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Begin VB.Form FrmDeclaracion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Declaracion"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   ScaleHeight     =   2955
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   5145
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
         Left            =   180
         TabIndex        =   2
         Text            =   "Quincena"
         Top             =   150
         Width           =   1635
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
         Left            =   3720
         MaskColor       =   &H00E0E0E0&
         Picture         =   "FrmDeclaracion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Tag             =   "1"
         ToolTipText     =   "Imprime en la impresora las retenciones del Proveedor"
         Top             =   1350
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dd/mm/aaaa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   1950
         TabIndex        =   4
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dd/mm/aaaa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   300
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   1485
      End
   End
   Begin vbskfr2.Skinner Skinner1 
      Left            =   4980
      Top             =   3150
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
      SysEnableSkinCaption=   "Habilitar &Skin"
      SysDisableSkinCaption=   "Deshabilitar &Skin"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   60
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Generacion de Archivo TXT"
   End
End
Attribute VB_Name = "FrmDeclaracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mes_Declaracion  As String
Dim NombreArchivoTxt As String

Private Sub Command1_Click()

End Sub

Private Sub CmdImpImpresora_Click()
Dim QuincenaDeclarar As Date
Adodc1.ConnectionString = ConexionDatasABaseD
Adodc1.RecordSource = "SELECT * FROM COMPROBANTE WHERE calendario='" & Combo1.Text & "' and IvaRetenido<>0 order by val(NoComprobante)"
Adodc1.Refresh
If Adodc1.Recordset.BOF Then
    MsgBox "No se han encontrado registros"
    Exit Sub
End If
QuincenaDeclarar = Adodc1.Recordset.Fields("FechaComprobante")
Mes_Declaracion = UCase(Format(QuincenaDeclarar, "MMMM") & " DE " & Format(QuincenaDeclarar, "YYYY"))
If Format(QuincenaDeclarar, "DD") > 15 Then
NombreArchivoTxt = "2DA QNA DE "
Else
NombreArchivoTxt = "1RA QNA DE "
End If
GenerarTxt
End Sub

Private Sub Combo1_Change()
TXTSQL = "SELECT * FROM CALENDARIOFISCAL WHERE PERIODO='" & Combo1.Text & "'"
Set REC = DB.Execute(TXTSQL)
Label1.Caption = REC.Fields("FechadeDeclaracion")

TXTSQL = "SELECT * FROM CALENDARIOFISCAL WHERE id=" & REC.Fields("ID") - 1 & ""
Set REC = DB.Execute(TXTSQL)
Label2.Caption = REC.Fields("FechadeDeclaracion")
End Sub

Private Sub Combo1_Click()
TXTSQL = "SELECT * FROM CALENDARIOFISCAL WHERE PERIODO='" & Combo1.Text & "'"
Set REC = DB.Execute(TXTSQL)
Label1.Caption = REC.Fields("FechadeDeclaracion")

TXTSQL = "SELECT * FROM CALENDARIOFISCAL WHERE id=" & REC.Fields("ID") - 1 & ""
Set REC = DB.Execute(TXTSQL)
Label2.Caption = REC.Fields("FechadeDeclaracion")
End Sub

Private Sub Form_Load()
Set Skinner1.Forms = Forms
Combo1.Clear
TXTSQL = "SELECT * FROM CALENDARIOFISCAL"
Set REC = DB.Execute(TXTSQL)
Do Until REC.EOF
Combo1.AddItem REC.Fields("PERIODO")
REC.MoveNext
Loop
Combo1.Text = Periodo
End Sub
Public Sub GenerarTxt()
On Error GoTo errores
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
CommonDialog1.Filter = "Archivos de Texto txt|*.txt"
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
MsgBox "No se genero el Archivo"
End Sub

