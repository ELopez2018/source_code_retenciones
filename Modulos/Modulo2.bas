Attribute VB_Name = "ModulodeInicio"
'Para leer desde el archivo ini
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As _
String) As Long

'Para Escribir en el archivo Ini
Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As _
Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'*******************variables************
Public DB As New ADODB.Connection
Public REC As New ADODB.Recordset
Public REC1 As New ADODB.Recordset
Public REC2 As New ADODB.Recordset
Public TXTSQL As String
Public FacturaS As New ADODB.Connection
Public DBRespaldo As New ADODB.Connection
Public MesActivo As Integer
'***************-Variables Para numeros ******************
Public Numeros As String
Public Letras As String
Public CedulaP As String
Public NombreP As String
Public Cede As String
Public ImpuestoG As Currency
Public CodigoDelModelo, NoCertificado, fechaCer, NFACTURA, Empresa, RifEmpresa  As String



Sub Main()
On Error GoTo Errores
    If App.PrevInstance Then
        MsgBox "Ya est· en Ejecusion", vbExclamation
        End
    End If
    NUMEROSFECHA = "0123456789/-"
    Numeros = "VEJPG-0123456789."
    Letras = "0123456789ABCDEFGHIJKLMN—OPQRSTUVWXYZ¡…Õ”⁄,.;-_\\\/ "
    TXTSQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\siv\DATOS\DATOS.mdb;Persist Security Info=True;Jet OLEDB:Database Password=181277"
    DB.Open (TXTSQL)
    'TXTSQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Facturas.mdb;Persist Security Info=True;Jet OLEDB:Database Password=181277"
    'FacturaS.Open (TXTSQL)
    'CREARTABLA
        TXTSQL = "Select * From ContadoresClaves where Nombre='Cede'"
    Set REC = DB.Execute(TXTSQL)
    Cede = REC!numero
    
    TXTSQL = "Select * From sucursales where CEDE=" & Cede & ""
    Set REC = DB.Execute(TXTSQL)
    Empresa = REC!Nombre
    
    TXTSQL = "Select * From ContadoresClaves where Nombre='Impuesto'"
    Set REC = DB.Execute(TXTSQL)
    ImpuestoG = REC!numero
    
    TXTSQL = "Select * From ContadoresClaves where Nombre='MesActivo'"
    Set REC = DB.Execute(TXTSQL)
    
    'MesActivo = REC!numero
    
    MenuPrincipal.Show
Exit Sub
Errores:
MsgBox Err.Description
End Sub
Sub degrada(frm As Form)
Dim i As Integer, Y As Integer
With frm
          .DrawStyle = 6
          .AutoRedraw = True
          .DrawMode = 13
          .ScaleMode = 2
          .DrawWidth = 8
          .ScaleWidth = 99
For i = 1 To 200
  frm.Line (i, 0)-(i - 1, Screen.Height), RGB(0, 30 + (i + 1), i), BF
Next i
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
Public Function EnLetras(numero As String) As String

    Dim b, paso As Integer

    Dim expresion, entero, deci, flag As String

        

    flag = "N"

    For paso = 1 To Len(numero)

        If Mid(numero, paso, 1) = "." Then

            flag = "S"

        Else

            If flag = "N" Then

                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero

            Else

                deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero

            End If

        End If

    Next paso

    

    If Len(deci) = 1 Then

        deci = deci & "0"

    End If

    

    flag = "N"

    If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999

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
On Error GoTo Errores
TXTSQL = "SELECT * FROM RIF"
Set REC = FacturaS.Execute(TXTSQL)
Exit Function
Errores:
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
