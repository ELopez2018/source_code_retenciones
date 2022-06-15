Attribute VB_Name = "ModulodeInicio"
'*******************variables************
Public DB As New ADODB.Connection
'RECORDSET'S
Public REC As New ADODB.Recordset
Public REC1 As New ADODB.Recordset
Public REC2 As New ADODB.Recordset
'***************-Variables Para numeros **********
Public NumerosCedulaRif, Noreporte, SOLONUMEROS, Letras  As String
Public TXTSQL As String
Public FORMATOMIL As String
Public TipoDeConsulta As String
Public RifProveedor As String
Public PeriodoConsulta As String
Public NUMEROSFECHA As String

'++++++++++++++Direcciones++++++++++++++++++++++++
Public TipodeReporte As String
Public DirecciondelReporte As String
Public FiltrodelReporte As String
Public DireccionBaseDatos As String
Public DireccionCarpetaReportes As String
Public ConexionDatasABaseD As String
Public VG_Desde As String
Public VG_Hasta As String
Public DireccionCarpetaExportar As String
'++++++++++++++Variables exportar Pdf+++++++++++++
Public ComprobanteInicial As Integer
Public ComprobanteFinal As Integer

'++++++++++++++Variables exportar Pdf+++++++++++++
Public AnoFiscaLibComp As String
Public MesFiscaLibComp  As String

'++++++++++++++Variables exportar Pdf++++++++++++
Public NocomproBanteConsulta As Double
Public NoFacturaPaConsultar As String

'++++++++++++++Operaciones IVA <=> ISRL ++++++++++
Public NoRifProveedor_G As String
Public NoFacturaProveedor_G As String
Public ConsulCompobuscar As Boolean
'++++++++++++++Operaciones Consulta Busqueda ++++++++++
Public TituloVentana As String
Public Periodo As String
Sub Main()
On Error GoTo errores
    If App.PrevInstance Then
        End
    End If
    'FUENTE
    DireccionBaseDatos = "\\ventas\Retenciones\Datos\"
    DireccionCarpetaReportes = "\\ventas\Retenciones\Comprobante"
    
    'DireccionBaseDatos = "S:\Datos\"
    'DireccionCarpetaReportes = "S:\Comprobante"
    
    FORMATOMIL = "#,##0.00"
    SOLONUMEROS = "0123456789,."
    NUMEROSFECHA = "0123456789/-"
    NumerosCedulaRif = "VEJPGBC0123456789"
    Letras = "0123456789ABCDEFGHIJKLMN—OPQRSTUVWXYZ¡…Õ”⁄,.;-_\\\/ "
   ConexionDatasABaseD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DireccionBaseDatos & "\DatosRetenciones.mdb;Persist Security Info=True;Jet OLEDB:Database Password=181277"
    DB.Open (ConexionDatasABaseD)
    'RetencionesISLR.Show
    frmSplash.Show
    'FrmConsulta.Show
Exit Sub
errores:
MsgBox Err.Description
End Sub
Sub degrada(frm As Form)
Dim I As Integer, Y As Integer
With frm
          .DrawStyle = 6
          .AutoRedraw = True
          .DrawMode = 13
          .ScaleMode = 2
          .DrawWidth = 8
          .ScaleWidth = 99
For I = 1 To 200
  frm.Line (I, 0)-(I - 1, Screen.Height), RGB(0, 30 + (I + 1), I), BF
Next I
End With
End Sub

Public Sub Pausa(ByVal nSegundos As Single)
    Dim t0 As Single
    t0 = Timer
    Do While Timer - t0 < nSegundos
        Dim dummy As Integer
        dummy = DoEvents()
        ' si nos pasamos de medianoche, retrocedemos un dÌa
        If Timer < t0 Then
            t0 = t0 - 24 * 60 * 60
        End If
    Loop
End Sub
Public Function DECIMALES(SIGNOD As Integer)
    Dim CANTIDAD  As String
    Dim DECIMALESV  As String
    Dim OPERACION As String
    CANTIDAD = "10"
    OPERACION = CANTIDAD & Chr(SIGNOD) & "00"
    If CCur(OPERACION) > 10 Then
    Select Case SIGNOD
        Case Is = 44
            SIGNOD = 46
        Case Is = 46
            SIGNOD = 44
        End Select
    End If
End Function
Public Function EnLetras(Numero As String) As String

    Dim b, paso As Integer

    Dim expresion, entero, deci, flag As String

        

    flag = "N"

    For paso = 1 To Len(Numero)

        If Mid(Numero, paso, 1) = "." Then

            flag = "S"

        Else

            If flag = "N" Then

                entero = entero + Mid(Numero, paso, 1) 'Extae la parte entera del numero

            Else

                deci = deci + Mid(Numero, paso, 1) 'Extrae la parte decimal del numero

            End If

        End If

    Next paso

    

    If Len(deci) = 1 Then

        deci = deci & "0"

    End If

    

    flag = "N"

    If Val(Numero) >= -999999999 And Val(Numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999

        For paso = Len(entero) To 1 Step -1

            b = Len(entero) - (paso - 1)

            Select Case paso

            Case 3, 6, 9

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then

                            expresion = expresion & "cien "

                        Else

                            expresion = expresion & "ciento "

                        End If

                    Case "2"

                        expresion = expresion & "doscientos "

                    Case "3"

                        expresion = expresion & "trescientos "

                    Case "4"

                        expresion = expresion & "cuatrocientos "

                    Case "5"

                        expresion = expresion & "quinientos "

                    Case "6"

                        expresion = expresion & "seiscientos "

                    Case "7"

                        expresion = expresion & "setecientos "

                    Case "8"

                        expresion = expresion & "ochocientos "

                    Case "9"

                        expresion = expresion & "novecientos "

                End Select

                

            Case 2, 5, 8

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If Mid(entero, b + 1, 1) = "0" Then

                            flag = "S"

                            expresion = expresion & "diez "

                        End If

                        If Mid(entero, b + 1, 1) = "1" Then

                            flag = "S"

                            expresion = expresion & "once "

                        End If

                        If Mid(entero, b + 1, 1) = "2" Then

                            flag = "S"

                            expresion = expresion & "doce "

                        End If

                        If Mid(entero, b + 1, 1) = "3" Then

                            flag = "S"

                            expresion = expresion & "trece "

                        End If

                        If Mid(entero, b + 1, 1) = "4" Then

                            flag = "S"

                            expresion = expresion & "catorce "

                        End If

                        If Mid(entero, b + 1, 1) = "5" Then

                            flag = "S"

                            expresion = expresion & "quince "

                        End If

                        If Mid(entero, b + 1, 1) > "5" Then

                            flag = "N"

                            expresion = expresion & "dieci"

                        End If

                

                    Case "2"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "veinte "

                            flag = "S"

                        Else

                            expresion = expresion & "veinti"

                            flag = "N"

                        End If

                    

                    Case "3"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "treinta "

                            flag = "S"

                        Else

                            expresion = expresion & "treinta y "

                            flag = "N"

                        End If

                

                    Case "4"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "cuarenta "

                            flag = "S"

                        Else

                            expresion = expresion & "cuarenta y "

                            flag = "N"

                        End If

                

                    Case "5"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "cincuenta "

                            flag = "S"

                        Else

                            expresion = expresion & "cincuenta y "

                            flag = "N"

                        End If

                

                    Case "6"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "sesenta "

                            flag = "S"

                        Else

                            expresion = expresion & "sesenta y "

                            flag = "N"

                        End If

                

                    Case "7"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "setenta "

                            flag = "S"

                        Else

                            expresion = expresion & "setenta y "

                            flag = "N"

                        End If

                

                    Case "8"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "ochenta "

                            flag = "S"

                        Else

                            expresion = expresion & "ochenta y "

                            flag = "N"

                        End If

                

                    Case "9"

                        If Mid(entero, b + 1, 1) = "0" Then

                            expresion = expresion & "noventa "

                            flag = "S"

                        Else

                            expresion = expresion & "noventa y "

                            flag = "N"

                        End If

                End Select

                

            Case 1, 4, 7

                Select Case Mid(entero, b, 1)

                    Case "1"

                        If flag = "N" Then

                            If paso = 1 Then

                                expresion = expresion & "uno "

                            Else

                                expresion = expresion & "un "

                            End If

                        End If

                    Case "2"

                        If flag = "N" Then

                            expresion = expresion & "dos "

                        End If

                    Case "3"

                        If flag = "N" Then

                            expresion = expresion & "tres "

                        End If

                    Case "4"

                        If flag = "N" Then

                            expresion = expresion & "cuatro "

                        End If

                    Case "5"

                        If flag = "N" Then

                            expresion = expresion & "cinco "

                        End If

                    Case "6"

                        If flag = "N" Then

                            expresion = expresion & "seis "

                        End If

                    Case "7"

                        If flag = "N" Then

                            expresion = expresion & "siete "

                        End If

                    Case "8"

                        If flag = "N" Then

                            expresion = expresion & "ocho "

                        End If

                    Case "9"

                        If flag = "N" Then

                            expresion = expresion & "nueve "

                        End If

                End Select

            End Select

            If paso = 4 Then

                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And Len(entero) <= 6) Then
                    expresion = expresion & "mil "
                End If

            End If

            If paso = 7 Then

                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then

                    expresion = expresion & "millÛn "

                Else

                    expresion = expresion & "millones "

                End If

            End If

        Next paso

        

        If deci <> "" Then

            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo

                EnLetras = "menos " & expresion & "con " & deci ' & "/100"

            Else

                EnLetras = expresion & "con " & deci ' & "/100"

            End If

        Else

            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo

                EnLetras = "menos " & expresion

            Else

                EnLetras = expresion

            End If

        End If

    Else 'si el numero a convertir esta fuera del rango superior e inferior

        EnLetras = ""

    End If

End Function


Public Function CREARTABLA()
On Error GoTo errores
TXTSQL = "SELECT * FROM RIF"
Set REC = FacturaS.Execute(TXTSQL)
Exit Function
errores:
If Err.Number = "-2147217865" Then
    TXTSQL = "CREATE TABLE RIF(RIF CHAR(20), nombre char(200))"
    FacturaS.Execute (TXTSQL)
    TXTSQL = "SELECT DISTINCT CEDULARIF FROM FACTURAS"
    Set REC = FacturaS.Execute(TXTSQL)
    Do Until REC.EOF
    TXTSQL = "SELECT * FROM FACTURAS WHERE CEDULARIF='" & REC("CEDULARIF") & "'"
    Set REC1 = FacturaS.Execute(TXTSQL)
    TXTSQL = "Insert Into RIF (RIF,NOMBRE)"
    TXTSQL = TXTSQL & " VALUES "
    TXTSQL = TXTSQL & "("
    TXTSQL = TXTSQL & "'" & REC1("CEDULARIF") & "'"
    TXTSQL = TXTSQL & ",'" & REC1("NOMBRE") & "'"
    TXTSQL = TXTSQL & ")"
    FacturaS.Execute (TXTSQL)
    REC.MoveNext
    Loop
End If
End Function
Public Function NUEVORIF()
    TXTSQL = "Insert Into RIF (RIF,NOMBRE)"
    TXTSQL = TXTSQL & " VALUES "
    TXTSQL = TXTSQL & "("
    TXTSQL = TXTSQL & "'" & CedulaP & "'"
    TXTSQL = TXTSQL & ",'" & NombreP & "'"
    TXTSQL = TXTSQL & ")"
    FacturaS.Execute (TXTSQL)
End Function
Public Function InformeDeErrores(Informe As String)
On Error GoTo cancelar 'en caso de que halla un error ir a cancelar
Dim archivo As String 'declarar una variable para identificar 'el nonbre del archivo
archivo = "c:\Errores\" & Format(Date, "DD") & Format(Date, "MMMMMMMMM") & Format(Now, "HH-nn-ss") & ".TXT"
Open archivo For Output As #1 'sintax para guardar una archivo
'open nombre_archivo for tipo as numero el cual se va a
'conocer en el codigo el archivo abierto

Print #1, Informe + vbCrLf 'anadimos lo que queremos en el archivo
'vbCrlf = orden del archivo quedara acomodado tal como
'esta en el textbox

Close 'cerramos el archivo

cancelar: 'accion que va a tomar si hay algun error
End Function
