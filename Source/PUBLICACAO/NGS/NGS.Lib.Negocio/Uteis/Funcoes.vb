Imports System.IO
Imports System.Data
Imports System.Text
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Drawing
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Uteis
Imports System.Net.Mail
Imports System.Web.Configuration
Imports System.Net.Configuration
Imports System.Configuration
Imports System.Net
Imports System.Web.HttpServerUtility
Imports System.Collections.Generic
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports System.Security.Cryptography
Imports System.Threading.Tasks
Imports PdfiumViewer
Imports System.Drawing.Printing
Imports Ionic.Zip
Imports DocsBr.Validation
Imports DocsBr
Imports System.Xml

Public Class Funcoes

    Private Shared ChaveSecreta As String = "0e73dbd8-d49d-42d7-849c-671d175901c8"
    '"0987612345!@#$%¨&*"

#Region "Validação Código de Barras"
    Public Shared Function FormataLinhaDigitavelOriginal(ByVal _codigo As String) As String
        Dim cmplivre As String
        Dim campo1 As String
        Dim campo2 As String
        Dim campo3 As String
        Dim campo4 As String
        Dim campo5 As String
        Dim icampo5 As Long
        Dim digitoMod As Integer

        _codigo = _codigo.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "")

        cmplivre = Mid(_codigo, 5, 5) & Mid(_codigo, 11, 10) & Mid(_codigo, 22, 10)
        campo1 = Strings.Left(_codigo, 4) + Strings.Mid(cmplivre, 1, 5)
        digitoMod = Mid(_codigo, 10, 1)
        campo1 = campo1 + digitoMod.ToString()
        campo1 = (Strings.Mid(campo1, 1, 5) & ".") + Strings.Mid(campo1, 6, 5)

        campo2 = Strings.Mid(cmplivre, 6, 10)
        digitoMod = Mid(_codigo, 21, 1)
        campo2 = campo2 + digitoMod.ToString()
        campo2 = (Strings.Mid(campo2, 1, 5) & ".") + Strings.Mid(campo2, 6, 6)

        campo3 = Strings.Mid(cmplivre, 16, 10)
        digitoMod = Mid(_codigo, 32, 1)
        campo3 = campo3 + digitoMod.ToString
        campo3 = (Strings.Mid(campo3, 1, 5) & ".") + Strings.Mid(campo3, 6, 6)

        campo4 = Strings.Mid(_codigo, 33, 1)

        icampo5 = Convert.ToInt64(Strings.Mid(_codigo, 34, 14))

        If icampo5 = 0 Then
            campo5 = "000"
        Else
            campo5 = icampo5.ToString()
        End If

        Return ((((campo1 & " ") + campo2 & " ") + campo3 & " ") + campo4 & " ") + campo5
    End Function

    Public Shared Function ValidaCodigoBarras(ByVal pCodigo As String, ByVal Digitado As Boolean, ByVal Vencimento As Date, ByVal Valor As Decimal, ByVal Cliente As String, ByVal EndCliente As String, ByRef Banco As AcessaBanco) As Boolean
        If Digitado = True Then
            If pCodigo <> FormataLinhaDigitavel(pCodigo) Then
                HttpContext.Current.Session("ssMessage") = "Código de Barras Inválido"
                Return False
            End If
        ElseIf Digitado = False Then
            If FormataCodigoBarra(pCodigo) = False Then
                HttpContext.Current.Session("ssMessage") = "Código de Barras Inválido"
                Return False
            End If
        End If

        If ValidaFatorDeVencimento(pCodigo, Digitado, Vencimento, Cliente, EndCliente, Banco) = False Then
            HttpContext.Current.Session("ssMessage") = "Vencimento não confere com Código de Barras"
            Return False
        End If

        If ValidaValorVencimento(pCodigo, Valor, Digitado) = False Then
            HttpContext.Current.Session("ssMessage") = "Valor não confere com Código de Barras"
            Return False
        End If

        If Not ValidaDigitoVerificador(pCodigo) Then
            HttpContext.Current.Session("ssMessage") = "Código de Barras Inválido!"
            Return False
        End If

        Return True
    End Function

    Friend Shared Function ValidaValorVencimento(ByVal pCodigo As String, ByVal Valor As Decimal, ByVal Digitado As Boolean) As Boolean
        pCodigo = pCodigo.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "")

        If Digitado Then
            If Valor <> (Mid(pCodigo, 38, 10) / 100) Then
                Return False
            End If
        Else
            If Valor <> (Mid(pCodigo, 10, 10) / 100) Then
                Return False
            End If
        End If

        Return True
    End Function

    Friend Shared Function FormataLinhaDigitavel(ByVal _codigo As String) As String
        Try
            Dim cmplivre As String
            Dim campo1 As String
            Dim campo2 As String
            Dim campo3 As String
            Dim campo4 As String
            Dim campo5 As String
            Dim icampo5 As Long
            Dim digitoMod As Integer

            _codigo = _codigo.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "")

            cmplivre = Mid(_codigo, 5, 5) & Mid(_codigo, 11, 10) & Mid(_codigo, 22, 10)
            campo1 = Strings.Left(_codigo, 4) + Strings.Mid(cmplivre, 1, 5)
            digitoMod = Mod10(campo1)
            campo1 = campo1 + digitoMod.ToString()
            campo1 = (Strings.Mid(campo1, 1, 5) & ".") + Strings.Mid(campo1, 6, 5)

            campo2 = Strings.Mid(cmplivre, 6, 10)
            digitoMod = Mod10(campo2)
            campo2 = campo2 + digitoMod.ToString()
            campo2 = (Strings.Mid(campo2, 1, 5) & ".") + Strings.Mid(campo2, 6, 6)

            campo3 = Strings.Mid(cmplivre, 16, 10)
            digitoMod = Mod10(campo3)
            campo3 = campo3 + digitoMod.ToString
            campo3 = (Strings.Mid(campo3, 1, 5) & ".") + Strings.Mid(campo3, 6, 6)

            campo4 = Strings.Mid(_codigo, 33, 1)

            icampo5 = Convert.ToInt64(Strings.Mid(_codigo, 34, 14))

            If icampo5 = 0 Then
                campo5 = "000"
            Else
                campo5 = icampo5.ToString()
            End If

            Return ((((campo1 & " ") + campo2 & " ") + campo3 & " ") + campo4 & " ") + campo5
        Catch
            Throw New Exception("Código de barras inválido")
        End Try
    End Function

    Friend Shared Function Mod10(ByVal seq As String) As Integer
        Dim Digito As Integer, Soma As Integer = 0, Peso As Integer = 2, res As Integer

        For i As Integer = seq.Length To 1 Step -1

            res = (Convert.ToInt32(Strings.Mid(seq, i, 1)) * Peso)

            If res > 9 Then
                res = CInt(res.ToString.Substring(0, 1)) + CInt(res.ToString.Substring(1, 1))
            End If

            Soma += res

            If Peso = 2 Then
                Peso = 1
            Else
                Peso = Peso + 1
            End If
        Next
        Digito = ((10 - (Soma Mod 10)) Mod 10)
        Return Digito
    End Function

    Friend Shared Function FormataCodigoBarra(ByVal _codigo As String) As Boolean
        Dim pdigito As Integer

        pdigito = CInt(Mid(_codigo, 5, 1))

        _codigo = _codigo.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "")

        _codigo = Mid(_codigo, 1, 4) & Mid(_codigo, 6, _codigo.Length)

        Dim Digito As Integer, Soma As Integer = 0, Peso As Integer = 2, res As Integer
        Dim arr As New ArrayList

        Dim aux As String = ""
        Dim n As Integer = 2

        For i As Integer = _codigo.Length To 1 Step -1
            aux = n.ToString & aux
            If n = 9 Then
                n = 2
            Else
                n = n + 1
            End If
        Next

        For i As Integer = _codigo.Length To 1 Step -1
            res = (CInt(_codigo.Substring(i - 1, 1)) * CInt(aux.Chars(i - 1).ToString))
            Soma += res
        Next

        If (Soma Mod 11 <> 0 And Soma Mod 11 <> 1) Or Soma Mod 11 > 9 Then
            Digito = 11 - Soma Mod 11
        ElseIf Soma Mod 11 = 0 Or Soma Mod 11 = 1 Or Soma Mod 11 > 9 Then
            Digito = 1
        End If

        If Digito = pdigito Then
            Return True
        Else
            Return False
        End If
        Return Digito
    End Function

    Friend Shared Function ValidaFatorDeVencimento(ByVal pCodigo As String, ByVal Digitado As Boolean, ByVal Vencimento As Date, ByVal Cliente As String, ByVal EndCliente As String, ByRef Banco As AcessaBanco) As Boolean

        pCodigo = pCodigo.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "")

        If Digitado = True AndAlso Len(pCodigo) <> 47 Then
            HttpContext.Current.Session("ssMessage") = "Linha Digitável inválida! Deve ter 47 dígitos."
            Return False
        ElseIf Digitado = False AndAlso Len(pCodigo) <> 44 Then
            HttpContext.Current.Session("ssMessage") = "Código de Barras inválido! Deve ter 44 dígitos."
            Return False
        End If

        'Regra de fator de vencimento se aplica a códigos de barra que começam com 2, 3, 4, 6 ou 7;
        Dim codigosValidos As String() = {"2", "3", "4", "6", "7"}

        If codigosValidos.Contains(Left(pCodigo, 1)) Then

            ' Definir a base de 1997 (usada para boletos gerados antes de 22/02/2025)
            Dim DataBase1997 As Date = #10/07/1997#

            ' Base de 2025 (usada para boletos gerados após 22/02/2025)
            Dim DataBase2025 As Date = #02/22/2025#

            ' Obter o fator de vencimento
            Dim FatorDeVencimento As Integer = 0
            If Digitado = True Then
                FatorDeVencimento = CInt(Mid(pCodigo, 34, 4))
            Else
                FatorDeVencimento = CInt(Mid(pCodigo, 6, 4))
            End If

            ' Ajustar o fator se o boleto for gerado após 22/02/2025
            If Vencimento >= DataBase2025 AndAlso FatorDeVencimento >= 1000 Then
                FatorDeVencimento -= 1000
            End If

            ' Validar com a base de 1997
            Dim DataVencimento1997 As Date = DataBase1997.AddDays(FatorDeVencimento)
            While VerificaDatasNaoProgramaveisBoleto(DataVencimento1997, "Movimento", Cliente, EndCliente, Banco) = False
                DataVencimento1997 = DataVencimento1997.AddDays(1)
            End While

            ' Validar com a base de 2025
            Dim DataVencimento2025 As Date = DataBase2025.AddDays(FatorDeVencimento)
            While VerificaDatasNaoProgramaveisBoleto(DataVencimento2025, "Movimento", Cliente, EndCliente, Banco) = False
                DataVencimento2025 = DataVencimento2025.AddDays(1)
            End While

            ' Se nenhuma das duas datas de vencimento for válida, retornamos False
            If (DataVencimento1997.Equals(Vencimento) = False AndAlso DateDiff(DateInterval.Day, DataVencimento1997, Vencimento) >= 1) And
           (DataVencimento2025.Equals(Vencimento) = False AndAlso DateDiff(DateInterval.Day, DataVencimento2025, Vencimento) >= 1) Then
                Return False
            End If

        End If

        Return True

    End Function

    Friend Shared Function VerificaDatasNaoProgramaveisBoleto(ByVal Data As String, ByVal Tipo As String, ByVal Empresa As String, ByVal EndEmpresa As String, ByRef Banco As AcessaBanco) As Boolean
        If IsDate(Data) Then
        Else
            Return False
        End If

        If CDate(Data).DayOfWeek = 6 Then
            Return False
        End If

        If CDate(Data).DayOfWeek = 0 Then
            Return False
        End If

        Dim Sql As String = ""
        Sql = "  SELECT * From DatasNaoProgramaveis"
        Sql &= " Where Empresa_Id = '99999999999999' And EndEmpresa_ID = 0 And Data_ID = '" & CDate(Data).ToString("yyyy/MM/dd") & "'"
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Return False
        Next

        If Empresa <> "" Then
            Sql = "  SELECT * From DatasNaoProgramaveis"
            Sql &= " Where Empresa_Id = '" & Empresa & "' And EndEmpresa_ID = " & EndEmpresa & " And Data_ID = '" & CDate(Data).ToString("yyyy/MM/dd") & "'"
            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                Return False
            Next
        End If
        Return True
    End Function

    Friend Shared Function ValidaDigitoVerificador(ByVal codigoDeBarra As String) As Boolean
        codigoDeBarra = Regex.Replace(codigoDeBarra, "[^0-9]", "")
        'Codigo de Barras
        If codigoDeBarra.Length.Equals(44) Then
            codigoDeBarra = codigoDeBarra.Substring(0, 4) + codigoDeBarra.Substring(19, 1) + codigoDeBarra.Substring(20, 4) + Mod10(codigoDeBarra.Substring(0, 4) + codigoDeBarra.Substring(19, 1) + codigoDeBarra.Substring(20, 4)).ToString +
                    codigoDeBarra.Substring(24, 5) + codigoDeBarra.Substring(29, 5) + Mod10(codigoDeBarra.Substring(24, 5) + codigoDeBarra.Substring(29, 5)).ToString +
                    codigoDeBarra.Substring(34, 5) + codigoDeBarra.Substring(39, 5) + Mod10(codigoDeBarra.Substring(34, 5) + codigoDeBarra.Substring(39, 5)).ToString +
                    codigoDeBarra.Substring(4, 1) + codigoDeBarra.Substring(5, 14)
        End If
        'Preenche com zeros a direita
        If codigoDeBarra.Length < 47 Then
            codigoDeBarra = String.Concat(codigoDeBarra, New String("0", 47 - codigoDeBarra.Length))
        End If
        codigoDeBarra = codigoDeBarra.Substring(0, 4) + codigoDeBarra.Substring(32, 15) +
                    codigoDeBarra.Substring(4, 5) + codigoDeBarra.Substring(10, 10) +
                    codigoDeBarra.Substring(21, 10)
        Dim digitoVerificador = codigoDeBarra.Substring(4, 1)
        codigoDeBarra = codigoDeBarra.Substring(0, 4) + codigoDeBarra.Substring(5, 39)
        Dim soma As Integer = 0, peso As Integer = 2, base As Integer = 9, resto As Integer = 0
        For i As Integer = codigoDeBarra.Length - 1 To 0 Step -1
            soma = soma + (codigoDeBarra.Substring(i, 1) * peso)
            If peso < base Then
                peso += 1
            Else
                peso = 2
            End If
        Next
        Dim digito As Integer = 11 - (soma Mod 11)
        If digito > 9 Then digito = 0
        If digito = 0 Then digito = 1

        If (digito <> digitoVerificador) Then Return False

        Return True
    End Function
#End Region

#Region "Retorna Último dia do Mês corrente"
    Public Shared Function UltimoDiaDoMes() As DateTime
        Return CDate(DateTime.DaysInMonth(Now.Year, Now.Month) & "/" & Now.Month.ToString & "/" & Now.Year.ToString)
    End Function

    Public Shared Function getUltimoDiaMes(ByVal data As DateTime) As Integer
        Return DateTime.DaysInMonth(data.Year, data.Month)
    End Function
#End Region

#Region "Datas não programáveis, sábado, domingo e feriados"
    Public Shared Function ValidaDataUtil(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pData As Date) As Date
        Dim dsDatas As New DataSet
        Dim DataAux As Date
        DataAux = pData

        While DataAux.DayOfWeek = DayOfWeek.Saturday Or DataAux.DayOfWeek = DayOfWeek.Sunday Or (pEmpresa.Length > 0 AndAlso VerificaDatasNaoProgramaveis(pEmpresa, pEndEmpresa, DataAux) = True)
            DataAux = DataAux.AddDays(1)
        End While
        Return DataAux
    End Function

    Public Shared Function VerificaDatasNaoProgramaveis(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pdata As Date) As Boolean
        Dim Banco As New AcessaBanco
        Dim sql As String
        Dim dsdatas As New DataSet
        Dim dr As DataRow
        sql = "select * from DatasNaoProgramaveis where Empresa_Id='" & pEmpresa & "' and EndEmpresa_Id=" & pEndEmpresa & " and Data_Id = '" & CDate(pdata).ToString("yyyy-MM-dd") & "'"
        dsdatas = Banco.ConsultaDataSet(sql, "DatasNaoProgramaveis")
        For Each dr In dsdatas.Tables(0).Rows
            Return True
        Next
        Return False
    End Function
#End Region

#Region "Converte Array String em Byte"
    Public Shared Function ConvertStringToByteArray(ByVal s As [String]) As [Byte]()
        Return (New UnicodeEncoding).GetBytes(s)
    End Function
#End Region

#Region "Verifica Permissoes"
    Public Shared Function VerificaPermissao(ByVal NomePrograma As String, ByVal TipoDeAcesso As String, Optional ByVal Usuario As String = "") As Boolean
        Dim banco As New AcessaBanco

        'If UserIsAdmin(IIf(String.IsNullOrWhiteSpace(Usuario), UsuarioServidor.NomeUsuario, Usuario)) Then Return True

        Dim sql As String = " SELECT Permissoes.Permissao_Id, Usuarios.Usuario_Id, GruposXUsuarios.Grupo_Id" & vbCrLf &
              "  FROM GruposXProcessosXPermissoes " & vbCrLf &
              " INNER JOIN GruposXUsuarios " & vbCrLf &
              "    ON GruposXProcessosXPermissoes.Grupo_Id = GruposXUsuarios.Grupo_Id" & vbCrLf &
              " INNER JOIN Usuarios" & vbCrLf &
              "    ON GruposXUsuarios.Usuario_Id = Usuarios.Usuario_Id" & vbCrLf &
              " INNER JOIN Permissoes" & vbCrLf &
              "    ON GruposXProcessosXPermissoes.Permissao_Id = Permissoes.Permissao_Id" & vbCrLf &
              " WHERE GruposXProcessosXPermissoes.Processo_Id = '" & NomePrograma & "'" & vbCrLf &
              "   AND GruposXUsuarios.Usuario_Id = '" & IIf(String.IsNullOrWhiteSpace(Usuario), UsuarioServidor.NomeUsuario, Usuario) & "'" & vbCrLf &
              "   And Permissoes.Permissao_Id = '" & TipoDeAcesso.ToUpper & "'" & vbCrLf &
              "   And Usuarios.Ativo = 1" & vbCrLf &
              "Select upp.Permissao_Id " & vbCrLf &
              "  From UsuariosXProcessosXPermissoes	upp" & vbCrLf &
              " Inner Join Usuarios u" & vbCrLf &
              "    on u.Usuario_Id = upp.Usuario_Id" & vbCrLf &
              " Where upp.Processo_Id = '" & NomePrograma & "'" & vbCrLf &
              "   And upp.Usuario_Id = '" & IIf(String.IsNullOrWhiteSpace(Usuario), UsuarioServidor.NomeUsuario, Usuario) & "'" & vbCrLf &
              "   And upp.Permissao_Id = '" & TipoDeAcesso.ToUpper & "'" & vbCrLf &
              "   And u.Ativo = 1" & vbCrLf

        Dim ds As DataSet = banco.ConsultaDataSet(sql, "GruposXProcessosXPermissoes")

        If (ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0) OrElse ds.Tables(1).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Shared Function UserIsAdmin(ByVal Usuario As String) As Boolean
        Dim banco As New AcessaBanco
        Dim sql As String = " select gu.Grupo_Id, gu.Usuario_Id       " & vbCrLf &
                            "   from GruposXUsuarios gu               " & vbCrLf &
                            "  Inner Join Usuarios u                  " & vbCrLf &
                            "     on u.Usuario_Id  = gu.Usuario_Id    " & vbCrLf &
                            "    and u.Ativo       = 1                " & vbCrLf &
                            "  where gu.Grupo_Id   = 'Admin'          " & vbCrLf &
                            "    and gu.Usuario_Id = '" & Usuario & "'" & vbCrLf

        Dim ds As DataSet = banco.ConsultaDataSet(sql, "UsuarioAdmin")

        If ds.Tables(0).Rows.Count > 0 Then
            Return True
        End If

        Return False
    End Function
#End Region

#Region "Alinhar Texto"
    Public Shared Sub InserirLinhaEmBranco(ByVal Combo As DropDownList)
        InserirLinhaEmBranco(Combo, 0)
    End Sub

    Public Shared Sub InserirLinhaEmBranco(ByVal Combo As DropDownList, ByVal IndiceSelecionado As Integer)
        Combo.Items.Insert(0, "")
        Combo.SelectedIndex = IndiceSelecionado
    End Sub

    Public Shared Sub ApenasDecimais(ByVal Controle As TextBox)
        Controle.Attributes.Add("onkeypress", "return ValidarNumeroDecimal(this, event);")
    End Sub

    Public Shared Function AlinharDireita(ByVal str As String, ByVal len As Integer, ByVal chr As String) As String
        Dim cont As Integer

        str = EliminarCaracteresEspeciais(str)

        For cont = str.Length To len
            str = chr & str
        Next

        If str.Length > len Then
            str = str.Substring(str.Length - len, len)
        End If

        Return str
    End Function

    Public Shared Function AlinharEsquerda(ByVal str As String, ByVal len As Integer, ByVal chr As String) As String
        Dim cont As Integer

        str = EliminarCaracteresEspeciais(str)

        For cont = str.Length To len
            str = str & chr
        Next

        If str.Length > len Then
            str = str.Substring(0, len)
        End If

        Return str
    End Function

    Public Shared Function EliminarCaracteresEspeciaisNF(ByVal texto As String) As String
        texto = texto.ToUpper()
        Dim s As String = texto.Normalize(NormalizationForm.FormKD)
        Dim sb As New StringBuilder()

        For k As Integer = 0 To s.Length - 1
            If s(k) = "." OrElse s(k) = "," OrElse s(k) = "-" OrElse s(k) = "|" OrElse s(k) = "/" OrElse s(k) = "\" OrElse s(k) = "$" OrElse s(k) = ":" OrElse s(k) = "+" OrElse s(k) = "%" OrElse s(k) = "(" OrElse s(k) = ")" OrElse s(k) = "'" OrElse s(k) = "`" OrElse s(k) = "´" Then
                sb.Append(s(k))
            Else
                Dim uc As UnicodeCategory = CharUnicodeInfo.GetUnicodeCategory(s(k))
                If uc = UnicodeCategory.SpaceSeparator OrElse uc = UnicodeCategory.UppercaseLetter OrElse uc = UnicodeCategory.DecimalDigitNumber Then
                    sb.Append(s(k))
                End If
            End If
        Next

        Return sb.ToString()
    End Function

    Public Shared Function EliminarCaracteresEspeciais(ByVal texto As String) As String
        texto = Trim(SubstituirCaracteresEspeciais(texto)).ToUpper()
        Dim s As String = texto.Normalize(NormalizationForm.FormKD)
        Dim sb As New StringBuilder()

        For k As Integer = 0 To s.Length - 1
            If s(k) = "." OrElse s(k) = "," OrElse s(k) = "-" OrElse s(k) = "|" OrElse s(k) = "/" OrElse s(k) = "\" OrElse s(k) = "$" OrElse s(k) = ":" OrElse s(k) = "+" OrElse s(k) = "%" OrElse s(k) = "(" OrElse s(k) = ")" OrElse s(k) = "'" OrElse s(k) = "`" OrElse s(k) = "´" OrElse s(k) = "=" Then
                sb.Append(s(k))
            Else
                Dim uc As UnicodeCategory = CharUnicodeInfo.GetUnicodeCategory(s(k))
                If uc = UnicodeCategory.SpaceSeparator OrElse uc = UnicodeCategory.UppercaseLetter OrElse uc = UnicodeCategory.DecimalDigitNumber Then
                    sb.Append(s(k))
                End If
            End If
        Next

        Return sb.ToString()
    End Function

    Public Shared Function EliminarAspasSimples(ByVal texto As String) As String
        texto = Trim(SubstituirCaracteresEspeciais(texto)).ToUpper()
        Dim s As String = texto.Normalize(NormalizationForm.FormKD)
        Dim sb As New StringBuilder()

        For k As Integer = 0 To s.Length - 1
            If s(k) = "'" Then
                Dim uc As UnicodeCategory = CharUnicodeInfo.GetUnicodeCategory(s(k))
                If uc = UnicodeCategory.SpaceSeparator OrElse uc = UnicodeCategory.UppercaseLetter OrElse uc = UnicodeCategory.DecimalDigitNumber Then
                    sb.Append(s(k))
                End If
            Else
                sb.Append(s(k))
            End If
        Next

        Return sb.ToString()
    End Function

    Public Shared Function RemoverLetrasDoNumero(ByVal str As String) As String
        If str Is Nothing Then Return ""

        Return str.Replace("A", "").Replace("B", "").Replace("C", "").Replace("D", "").Replace("E", "").Replace("F", "").Replace("G", "").Replace("H", "").Replace("I", "").Replace("J", "").Replace("K", "").Replace("L", "").Replace("M", "").Replace("N", "").Replace("O", "").Replace("P", "").Replace("Q", "").Replace("R", "").Replace("S", "").Replace("T", "").Replace("U", "").Replace("X", "").Replace("Y", "").Replace("W", "").Replace("Z", "").Replace("-", "").Replace("/", "").Replace(".", "").Replace("*", "").Replace("'", "").Replace(" ", "")
    End Function

    Public Shared Function SubstituirCaracteresEspeciaisMunicipio(ByVal str As String) As String
        If str Is Nothing Then Return ""

        Return str.Replace("´", "").Replace("`", "").Replace("Á", "A").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Ú", "U").Replace("Ã", "A").Replace("Õ", "O").Replace("À", "A").Replace("È", "E").Replace("Ì", "I").Replace("Ò", "O").Replace("Ù", "U").Replace("Â", "A").Replace("Ê", "E").Replace("Î", "I").Replace("Ô", "O").Replace("Û", "U").Replace("Ä", "A").Replace("Ë", "E").Replace("Ï", "I").Replace("Ö", "O").Replace("Ü", "U").Replace("'", "").Replace("Ç", "C")
    End Function

    Public Shared Function SubstituirCaracteresEspeciais(ByVal str As String) As String
        If str Is Nothing Then Return ""

        Return str.Replace("º", "").Replace("ª", "").Replace("´", "").Replace("§", "").Replace("`", "").Replace("€", "").Replace("ç", "c").Replace("Ç", "C").Replace("Á", "A").Replace("É", "E").Replace("Í", "I").Replace("Ó", "O").Replace("Ú", "U").Replace("Ã", "A").Replace("Õ", "O").Replace("À", "A").Replace("È", "E").Replace("Ì", "I").Replace("Ò", "O").Replace("Ù", "U").Replace("Â", "A").Replace("Ê", "E").Replace("Î", "I").Replace("Ô", "O").Replace("Û", "U").Replace("Ä", "A").Replace("Ë", "E").Replace("Ï", "I").Replace("Ö", "O").Replace("Ü", "U").Replace("'", "")
    End Function

    Public Shared Function SubstituirCaracteresEspeciaisALL(ByVal str As String) As String
        If String.IsNullOrEmpty(str) Then Return ""

        ' Substituições para letras maiúsculas e minúsculas
        Dim caracteresEspeciais As New Dictionary(Of String, String) From {
        {"´", ""}, {"`", ""}, {"'", ""}, {"Á", "A"}, {"É", "E"}, {"Í", "I"}, {"Ó", "O"}, {"Ú", "U"},
        {"Ã", "A"}, {"Õ", "O"}, {"À", "A"}, {"È", "E"}, {"Ì", "I"}, {"Ò", "O"}, {"Ù", "U"},
        {"Â", "A"}, {"Ê", "E"}, {"Î", "I"}, {"Ô", "O"}, {"Û", "U"}, {"Ä", "A"}, {"Ë", "E"}, {"Ï", "I"},
        {"Ö", "O"}, {"Ü", "U"}, {"Ç", "C"}, {"á", "a"}, {"é", "e"}, {"í", "i"}, {"ó", "o"}, {"ú", "u"},
        {"ã", "a"}, {"õ", "o"}, {"à", "a"}, {"è", "e"}, {"ì", "i"}, {"ò", "o"}, {"ù", "u"},
        {"â", "a"}, {"ê", "e"}, {"î", "i"}, {"ô", "o"}, {"û", "u"}, {"ä", "a"}, {"ë", "e"}, {"ï", "i"},
        {"ö", "o"}, {"ü", "u"}, {"ç", "c"}
    }

        ' Realizar as substituições
        For Each kvp In caracteresEspeciais
            str = str.Replace(kvp.Key, kvp.Value)
        Next

        Return str
    End Function

    Public Shared Function RemoveAllEnterKey(ByVal str As String) As String
        Dim ptr As Integer = str.Length
        Dim result As String = ""
        While ptr > 0
            If (Asc(str.Chars(ptr - 1)) = 13 Or Asc(str.Chars(ptr - 1)) = 10) Then
                result = " " & result
            Else
                result = str.Substring(ptr - 1, 1) & result
            End If
            ptr -= 1
        End While
        Return result
    End Function

    Public Shared Function ValidarSerie(ByVal Serie As String, ByVal ES As String) As String
        Select Case Serie
            Case "1", "1A", "1B", "1C", "1D", "1-A", "1-B", "1-C", "1-D", "M1", "M-1", "A-1", "B-1", "C-1", "D-1", "E-1", "F-1", "S-1", "U-1"
                Serie = "1"
            Case "2", "2A", "2B", "2C", "2D", "2-A", "2-B", "2-C", "2-D", "M2", "M-2", "A-2", "B-2", "C-2", "D-2", "E-2", "F-2", "S-2", "U-2"
                Serie = "2"
            Case "3", "M3", "M-3", "A-3", "B-3", "C-3", "D-3", "E-3", "F-3", "S-3", "U-3"
                Serie = "3"
            Case "4", "M4", "M-4", "A-4", "B-4", "C-4", "D-4", "E-4", "F-4", "S-4", "U-4"
                Serie = "4"
            Case Else
                If ES = "E" Then
                    Serie = "2"
                Else
                    Serie = "1"
                End If
        End Select

        Return Serie

    End Function
#End Region

#Region "Retorna o nome do arquivo para os relatorios"
    Public Shared Function GeraNomeArquivo() As String
        Dim strNomeArquivo As String = HttpContext.Current.Session("ssNomeUsuario") & DateTime.Now.ToString("ddMMyyHHmmss")
        Return strNomeArquivo
    End Function
#End Region

#Region "Formatar CNPJ e CPF"
    Public Shared Function FormatarCpfCnpj(ByVal Numero As String) As String
        Dim Campo As String = Numero.Trim()

        If Not IsDBNull(Campo) AndAlso Not String.IsNullOrWhiteSpace(Campo) Then
            If Len(Campo) > 11 Then
                Campo = Left(Campo, 2) & "." & Mid(Campo, 3, 3) & "." & Mid(Campo, 6, 3) & "/" & Mid(Campo, 9, 4) & "-" & Mid(Campo, 13, 2)
            Else
                Campo = Left(Campo, 3) & "." & Mid(Campo, 4, 3) & "." & Mid(Campo, 7, 3) & "-" & Mid(Campo, 10, 2)
            End If
            FormatarCpfCnpj = Campo
        Else
            FormatarCpfCnpj = String.Empty
        End If
    End Function

    Public Shared Function FormatarCPFeCNPJ(ByVal pcodigo As String) As String
        Dim format_codigo As String = ""
        For Each number As String In pcodigo
            If Not IsNumeric(number) Then
            Else
                format_codigo += number
            End If
        Next
        Return format_codigo
    End Function
#End Region

#Region "FormatarIE"
    Public Shared Function FormatarInscricaoEstadual(inscricao As String) As String
        ' Verifica se a inscrição é nula ou vazia
        If String.IsNullOrWhiteSpace(inscricao) Then
            Return String.Empty
        End If

        ' Remove quaisquer caracteres não numéricos
        inscricao = New String(inscricao.Where(Function(c) Char.IsDigit(c)).ToArray())

        Select Case inscricao.Length
            Case 8 ' Exemplo: 12345678 -> 12.345.678
                Return $"{inscricao.Substring(0, 2)}.{inscricao.Substring(2, 3)}.{inscricao.Substring(5, 3)}"
            Case 9 ' Exemplo: 123456789 -> 123.456.78-9
                Return $"{inscricao.Substring(0, 3)}.{inscricao.Substring(3, 3)}.{inscricao.Substring(6, 2)}-{inscricao.Substring(8, 1)}"
            Case 10 ' Exemplo: 1234567890 -> 123.456.789-0
                Return $"{inscricao.Substring(0, 3)}.{inscricao.Substring(3, 3)}.{inscricao.Substring(6, 3)}-{inscricao.Substring(9, 1)}"
            Case 12 ' Exemplo: 123456789012 -> 123.456.789.012
                Return $"{inscricao.Substring(0, 3)}.{inscricao.Substring(3, 3)}.{inscricao.Substring(6, 3)}.{inscricao.Substring(9, 3)}"
            Case 13 ' Exemplo: 1234567890123 -> 123.456.789.012-3
                Return $"{inscricao.Substring(0, 3)}.{inscricao.Substring(3, 3)}.{inscricao.Substring(6, 3)}.{inscricao.Substring(9, 3)}-{inscricao.Substring(12, 1)}"
            Case 14 ' Exemplo: 12345678901234 -> 123.456.789.0123-4
                Return $"{inscricao.Substring(0, 3)}.{inscricao.Substring(3, 3)}.{inscricao.Substring(6, 3)}.{inscricao.Substring(9, 4)}-{inscricao.Substring(13, 1)}"
            Case Else
                ' Retorna o número original caso não tenha o tamanho esperado
                Return inscricao
        End Select

    End Function
#End Region

#Region "FormatarTELEFONE"

    Public Shared Function FormatarTelefone(telefone As String) As String
        ' Remove caracteres inválidos como espaços, parênteses, traços, etc.
        Dim telefoneLimpo As String = New String(telefone.Where(Function(c) Char.IsDigit(c)).ToArray())

        ' Verifica o tamanho do telefone para decidir a formatação
        Select Case telefoneLimpo.Length
            Case 8
                ' Formata como telefone fixo: XXXX-XXXX
                Return telefoneLimpo.Insert(4, "-")
            Case 9
                ' Formata como telefone celular: X XXXX-XXXX
                Return telefoneLimpo.Insert(1, " ").Insert(6, "-")
            Case 10
                ' Formata como DDD + telefone fixo: (XX) XXXX-XXXX
                Return $"({telefoneLimpo.Substring(0, 2)}) {telefoneLimpo.Substring(2, 4)}-{telefoneLimpo.Substring(6)}"
            Case 11
                ' Formata como DDD + telefone celular: (XX) X XXXX-XXXX
                Return $"({telefoneLimpo.Substring(0, 2)}) {telefoneLimpo.Substring(2, 1)} {telefoneLimpo.Substring(3, 4)}-{telefoneLimpo.Substring(7)}"
            Case Else
                ' Retorna o telefone original se o tamanho não for compatível
                Return telefone
        End Select

    End Function

#End Region

#Region "Formatar Chave NFe"
    Public Shared Function FormatarChaveNFe(ByVal Numero As String) As String
        Dim Campo As String = Numero.Trim()

        If Not String.IsNullOrWhiteSpace(Campo) AndAlso Campo.Length = 44 Then
            Campo = Left(Campo, 4) & "." & Mid(Campo, 5, 4) & "." & Mid(Campo, 9, 4) & "." & Mid(Campo, 13, 4) & "." & Mid(Campo, 17, 4) & "." &
                Mid(Campo, 21, 4) & "." & Mid(Campo, 25, 4) & "." & Mid(Campo, 29, 4) & "." & Mid(Campo, 33, 4) & Mid(Campo, 37, 4) & "." & Mid(Campo, 41, 4)

        End If
        FormatarChaveNFe = Campo
    End Function
#End Region

#Region "Retorna um valor em extenso"

    Public Shared Function Extenso(ByVal Numero As String, ByVal MoedaSingular As String, ByVal MoedaPlural As String) As String
        Dim stringValor As String
        Dim valor As Decimal
        Dim Negativo As Boolean
        Dim buffer As String
        Dim parcial As Integer
        Dim posicao As Integer
        Dim numExtenso As String = String.Empty
        Dim CultureInfo As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("pt-BR")

        Dim Unidades() As String = {String.Empty, "Um", "Dois",
                                    "Três", "Quatro", "Cinco",
                                    "Seis", "Sete", "Oito", "Nove",
                                    "Dez", "Onze", "Doze", "Treze",
                                    "Quatorze", "Quinze", "Dezesseis",
                                    "Dezessete", "Dezoito", "Dezenove"}

        Dim Dezenas() As String = {String.Empty, String.Empty,
                                   "Vinte", "Trinta", "Quarenta",
                                   "Cinquenta", "Sessenta", "Setenta",
                                   "Oitenta", "Noventa"}

        Dim Centenas() As String = {String.Empty, "Cento",
                                    "Duzentos", "Trezentos",
                                    "Quatrocentos", "Quinhentos",
                                    "Seiscentos", "Setecentos",
                                    "Oitocentos", "Novecentos"}

        Dim PotenciasSingular() As String = {String.Empty, " Mil",
                                             " Milhão", " Bilhão",
                                             " Trilhão", " Quatrilhão"}

        Dim PotenciasPlural() As String = {String.Empty, " Mil",
                                           " Milhões", " Bilhões",
                                           " Trilhões", " Quatrilhões"}

        Dim ret As Boolean = Decimal.TryParse(Numero, System.Globalization.NumberStyles.Currency, CultureInfo, valor)
        Dim semDe As Boolean = False
        If ret Then
            Negativo = (valor < 0)
            If Negativo Then
                valor = Math.Abs(valor)
            End If
            stringValor = valor.ToString(New String("0"c, 18) & ".000").Substring(0, 18)
            For posicao = 1 To 18 Step 3
                parcial = CType(stringValor.Substring(posicao - 1, 3), Integer)
                If parcial > 0 Then
                    If parcial = 1 Then
                        buffer = "Um" & PotenciasSingular((18 - posicao) \ 3)
                    ElseIf parcial = 100 Then
                        buffer = "Cem" & PotenciasPlural((18 - posicao) \ 3) 'Convert.ToString( _
                        'IIf(posicao > 16, PotenciasSingular((18 - posicao) \ 3), PotenciasPlural((18 - posicao) \ 3)))
                    Else
                        buffer = Centenas(parcial \ 100)
                        parcial = parcial Mod 100
                        If parcial <> 0 AndAlso buffer <> String.Empty Then
                            buffer = buffer & " e "
                        End If
                        If parcial < 20 Then
                            buffer = buffer & Unidades(parcial)
                        Else
                            buffer = buffer & Dezenas(parcial \ 10)
                            parcial = parcial Mod 10
                            If parcial <> 0 AndAlso buffer <> String.Empty Then
                                buffer = buffer & " e "
                            End If
                            buffer = buffer & Unidades(parcial)
                        End If
                        buffer = buffer & PotenciasPlural((18 - posicao) \ 3)
                    End If
                    If buffer <> String.Empty Then
                        If numExtenso <> String.Empty Then
                            parcial = CType(stringValor.Substring(posicao - 1, 3), Integer)
                            If posicao = 16 AndAlso (parcial < 100 OrElse (parcial Mod 100) = 0) Then
                                numExtenso = numExtenso & " e "
                            Else
                                numExtenso = numExtenso & ", "
                            End If
                            semDe = True
                        End If
                        numExtenso = numExtenso & buffer
                    End If
                Else
                    If Not semDe AndAlso posicao = 16 AndAlso (valor >= 1000000) Then
                        numExtenso = numExtenso & " de"
                    End If
                End If
            Next
            If numExtenso <> String.Empty Then
                If Negativo Then
                    numExtenso = "Menos " & numExtenso
                End If
                If Math.Truncate(valor) = 1 Then
                    numExtenso = numExtenso & " " & MoedaSingular
                Else
                    numExtenso = numExtenso & " " & MoedaPlural
                End If
            End If
            parcial = CType(((valor - Math.Truncate(valor)) * 100 + 0.1), Integer)
            If parcial > 0 Then
                buffer = Extenso(Convert.ToString(parcial), "Centavo", "Centavos")
                If numExtenso <> String.Empty Then
                    numExtenso = numExtenso & " e "
                End If
                numExtenso = numExtenso & buffer
            End If
        End If
        Return numExtenso
    End Function

    Public Shared Function ExtensoGeneroFeminino(ByVal Numero As String, ByVal MoedaSingular As String, ByVal MoedaPlural As String) As String
        Dim stringValor As String
        Dim valor As Decimal
        Dim Negativo As Boolean
        Dim buffer As String
        Dim parcial As Integer
        Dim posicao As Integer
        Dim numExtenso As String = String.Empty
        Dim CultureInfo As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("pt-BR")

        Dim Unidades() As String = {String.Empty, "Uma", "Duas",
                                "Três", "Quatro", "Cinco",
                                "Seis", "Sete", "Oito", "Nove",
                                "Dez", "Onze", "Doze", "Treze",
                                "Quatorze", "Quinze", "Dezesseis",
                                "Dezessete", "Dezoito", "Dezenove"}

        Dim Dezenas() As String = {String.Empty, String.Empty,
                               "Vinte", "Trinta", "Quarenta",
                               "Cinquenta", "Sessenta", "Setenta",
                               "Oitenta", "Noventa"}

        Dim Centenas() As String = {String.Empty, "Cento",
                                "Duzentas", "Trezentas",
                                "Quatrocentas", "Quinhentas",
                                "Seiscentas", "Setecentas",
                                "Oitocentas", "Novecentas"}

        Dim PotenciasSingular() As String = {String.Empty, " Mil",
                                         " Milhão", " Bilhão",
                                         " Trilhão", " Quatrilhão"}

        Dim PotenciasPlural() As String = {String.Empty, " Mil",
                                       " Milhões", " Bilhões",
                                       " Trilhões", " Quatrilhões"}

        Dim ret As Boolean = Decimal.TryParse(Numero, System.Globalization.NumberStyles.Currency, CultureInfo, valor)
        Dim semDe As Boolean = False
        If ret Then
            Negativo = (valor < 0)
            If Negativo Then valor = Math.Abs(valor)

            stringValor = valor.ToString(New String("0"c, 18) & ".000").Substring(0, 18)
            For posicao = 1 To 18 Step 3
                parcial = CType(stringValor.Substring(posicao - 1, 3), Integer)
                If parcial > 0 Then
                    If parcial = 1 Then
                        buffer = Unidades(1) & PotenciasSingular((18 - posicao) \ 3)
                    ElseIf parcial = 100 Then
                        buffer = "Cem" & PotenciasPlural((18 - posicao) \ 3)
                    Else
                        buffer = Centenas(parcial \ 100)
                        parcial = parcial Mod 100
                        If parcial <> 0 AndAlso buffer <> String.Empty Then
                            buffer &= " e "
                        End If
                        If parcial < 20 Then
                            buffer &= Unidades(parcial)
                        Else
                            buffer &= Dezenas(parcial \ 10)
                            parcial = parcial Mod 10
                            If parcial <> 0 AndAlso buffer <> String.Empty Then
                                buffer &= " e "
                            End If
                            buffer &= Unidades(parcial)
                        End If
                        buffer &= PotenciasPlural((18 - posicao) \ 3)
                    End If
                    If buffer <> String.Empty Then
                        If numExtenso <> String.Empty Then
                            parcial = CType(stringValor.Substring(posicao - 1, 3), Integer)
                            If posicao = 16 AndAlso (parcial < 100 OrElse (parcial Mod 100) = 0) Then
                                numExtenso &= " e "
                            Else
                                numExtenso &= ", "
                            End If
                            semDe = True
                        End If
                        numExtenso &= buffer
                    End If
                Else
                    If Not semDe AndAlso posicao = 16 AndAlso (valor >= 1000000) Then
                        numExtenso &= " de"
                    End If
                End If
            Next
            If numExtenso <> String.Empty Then
                If Negativo Then numExtenso = "Menos " & numExtenso
                If Math.Truncate(valor) = 1 Then
                    numExtenso &= " " & MoedaSingular
                Else
                    numExtenso &= " " & MoedaPlural
                End If
            End If
            parcial = CType(((valor - Math.Truncate(valor)) * 100 + 0.1), Integer)
            If parcial > 0 Then
                buffer = ExtensoGeneroFeminino(parcial.ToString(), "Centavo", "Centavos")
                If numExtenso <> String.Empty Then
                    numExtenso &= " e "
                End If
                numExtenso &= buffer
            End If
        End If
        Return numExtenso
    End Function

#End Region

    Public Shared Function VerificaAcessoMensal(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal Data As String, ByVal Processo As String) As Boolean
        Dim sql As String
        Dim banco As New AcessaBanco
        Dim ds As New DataSet
        Dim drMovimento
        Dim TemMovimento As Boolean = False
        Dim TemProcesso As Boolean = False
        Dim TemUsuarios As Boolean = False
        sql = "select Situacao from AcessosXMovimento where Empresa_Id = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and month(Movimento_Id) = '" & CDate(Data).ToString("MM") & "'" & " and year(Movimento_Id) = '" & CDate(Data).ToString("yyyy") & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXMovimento")
        If ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each drMovimento In ds.Tables(0).Rows
                If drMovimento("Situacao") = "BLOQUEADO" Then
                    TemMovimento = True
                End If
            Next
        Else
            Return False
        End If

        sql = "select * from AcessosXProcessos where Empresa_Id = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and Movimento_Id = '" & CDate(Data).ToString("yyyy-MM-dd") & "' and Processo_Id = '" & Processo & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXProcessos")
        If ds.Tables(0).Rows.Count = 0 Then
            TemProcesso = False
        Else
            TemProcesso = True
        End If

        sql = "select * from AcessosXUsuarios where Empresa_Id = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and Movimento_Id = '" & CDate(Data).ToString("yyyy-MM-dd") & "' and Processo_Id = '" & Processo & "' and Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXUsuarios")
        If ds.Tables(0).Rows.Count = 0 Then
            TemUsuarios = False
        Else
            TemUsuarios = True

        End If

        Return ((TemProcesso And TemUsuarios) Or (TemMovimento And TemUsuarios) Or (Not TemMovimento And Not TemProcesso And Not TemUsuarios))
    End Function


    Public Shared Function VerificaAcesso(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal Data As String, ByVal Processo As String) As Boolean
        Dim sql As String
        Dim banco As New AcessaBanco
        Dim ds As New DataSet
        Dim drMovimento
        Dim TemMovimento As Boolean = False
        Dim TemProcesso As Boolean = False
        Dim TemUsuarios As Boolean = False
        sql = "select Situacao from AcessosXMovimento where Empresa_Id = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and Movimento_Id = '" & CDate(Data).ToString("yyyy-MM-dd") & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXMovimento")
        If ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each drMovimento In ds.Tables(0).Rows
                If drMovimento("Situacao") = "LIBERADO" Then
                    TemMovimento = False
                Else
                    TemMovimento = True
                End If
            Next
        Else
            Return False
        End If

        sql = "select * from AcessosXProcessos where Empresa_Id = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and Movimento_Id = '" & CDate(Data).ToString("yyyy-MM-dd") & "' and Processo_Id = '" & Processo & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXProcessos")
        If ds.Tables(0).Rows.Count = 0 Then
            TemProcesso = False
        Else
            TemProcesso = True
        End If

        sql = "select * from AcessosXUsuarios where Empresa_Id = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and Movimento_Id = '" & CDate(Data).ToString("yyyy-MM-dd") & "' and Processo_Id = '" & Processo & "' and Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXUsuarios")
        If ds.Tables(0).Rows.Count = 0 Then
            TemUsuarios = False
        Else
            TemUsuarios = True

        End If

        Return ((TemProcesso And TemUsuarios) Or (TemMovimento And TemUsuarios) Or (Not TemMovimento And Not TemProcesso And Not TemUsuarios))
    End Function

    Public Shared Function verificaAcessoCusto(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal mes As String, ByVal ano As String, ByVal Processo As String) As Boolean
        Dim sql As String
        Dim banco As New AcessaBanco
        Dim ds As New DataSet
        Dim drMovimento
        Dim TemMovimento As Boolean = False
        Dim TemProcesso As Boolean = False
        Dim TemUsuarios As Boolean = False

        Dim Data As String

        If (mes.Length = 1) Then
            mes = "0" & mes
        End If

        Data = Date.DaysInMonth(Today.Year, mes) & "/" & mes & "/" & ano

        sql = "select Situacao from AcessosXMovimento where left(Empresa_Id,8) = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and Movimento_Id = '" & CDate(Data).ToString("yyyy-MM-dd") & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXMovimento")
        If ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each drMovimento In ds.Tables(0).Rows
                If drMovimento("Situacao") = "LIBERADO" Then
                    TemMovimento = False
                Else
                    TemMovimento = True
                End If
            Next
        Else
            Return False
        End If

        sql = "select * from AcessosXProcessos where left(Empresa_Id,8) = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and Movimento_Id = '" & CDate(Data).ToString("yyyy-MM-dd") & "' and Processo_Id = '" & Processo & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXProcessos")
        If ds.Tables(0).Rows.Count = 0 Then
            TemProcesso = False
        Else
            TemProcesso = True
        End If

        sql = "select * from AcessosXUsuarios where left(Empresa_Id,8) = '" & Empresa & "' and EndEmpresa_Id = " & EndEmpresa & " and Movimento_Id = '" & CDate(Data).ToString("yyyy-MM-dd") & "' and Processo_Id = '" & Processo & "' and Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
        ds = banco.ConsultaDataSet(sql, "AcessosXUsuarios")
        If ds.Tables(0).Rows.Count = 0 Then
            TemUsuarios = False
        Else
            TemUsuarios = True

        End If

        Return ((TemProcesso And TemUsuarios) Or (TemMovimento And TemUsuarios) Or (Not TemMovimento And Not TemProcesso And Not TemUsuarios))
    End Function

#Region "Funcao utilizada pela valida cpf"
    Private Shared Function Modulo11(ByVal numero As String) As String
        Dim I As Integer
        Dim Produto As Integer
        Dim Multiplicador As Integer
        Dim Digito As Integer
        ' Válida Argumento
        If Not IsNumeric(numero) Then
            Modulo11 = ""
            Exit Function
        End If
        ' Cálcula Dígito no Módulo 11
        Multiplicador = 2
        For I = Len(numero) To 1 Step -1
            Produto = Produto + Val(Mid(numero, I, 1)) * Multiplicador
            Multiplicador = IIf(Multiplicador = 9, 2, Multiplicador + 1)
        Next
        ' Exceção
        Digito = 11 - Int(Produto Mod 11)
        Digito = IIf(Digito = 10 Or Digito = 11, 0, Digito)
        ' Retorna
        Modulo11 = Trim(Str(Digito))
    End Function
#End Region

#Region " Válida CNPJ"
    Public Shared Function ValidaCNPJ(ByVal cnpj As String) As Boolean
        'Obs. Os parametros devem ser passados sem nenhuma pontuação!

        ' Válida argumento
        If Len(cnpj) <> 14 Then
            ValidaCNPJ = False
            Exit Function
        End If
        ' Válida Primeiro Dígito
        If Modulo11(Left(cnpj, 12)) <> Mid(cnpj, 13, 1) Then
            ValidaCNPJ = False
            Exit Function
        End If
        ' Válida Segundo Dígito
        If Modulo11(Left(cnpj, 13)) <> Mid(cnpj, 14, 1) Then
            ValidaCNPJ = False
            Exit Function
        End If
        ValidaCNPJ = True
        'CodigoCliente = Cgc
    End Function
#End Region

#Region "Válida o CPF"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Cpf">Número do CPF.</param>
    ''' <returns>Retorna verdadeiro ou falso para validação do cpf.</returns>
    ''' <remarks>Observação: Parametro deve ser passado com apenas números.</remarks>
    Public Shared Function ValidaCPF(ByVal Cpf As String) As Boolean
        'Obs. Os parametros devem ser passados sem nenhuma pontuação!
        'Dim WdigitoDoCPF
        ValidaCPF = True
        Dim wSomaDosProdutos
        Dim wResto
        Dim wDigitChk1
        Dim wDigitChk2
        'Dim wStatus
        Dim wI
        wSomaDosProdutos = 0
        For wI = 1 To 9
            wSomaDosProdutos = wSomaDosProdutos + Val(Mid(Cpf, wI, 1)) * (11 - wI)
        Next wI
        wResto = wSomaDosProdutos - Int(wSomaDosProdutos / 11) * 11
        wDigitChk1 = IIf(wResto = 0 Or wResto = 1, 0, 11 - wResto)

        wSomaDosProdutos = 0
        For wI = 1 To 9
            wSomaDosProdutos = wSomaDosProdutos + (Val(Mid(Cpf, wI, 1)) * (12 - wI))
        Next wI
        wSomaDosProdutos = wSomaDosProdutos + (2 * wDigitChk1)
        wResto = wSomaDosProdutos - Int(wSomaDosProdutos / 11) * 11
        wDigitChk2 = IIf(wResto = 0 Or wResto = 1, 0, 11 - wResto)

        If Mid(Cpf, 10, 1) = Mid(Trim(Str(wDigitChk1)), 1, 1) And Mid(Cpf, 11, 1) = Mid(Trim(Str(wDigitChk2)), 1, 1) Then
            ValidaCPF = True
            'CodigoCliente = Cpf
        Else
            ValidaCPF = False
        End If
    End Function
#End Region

#Region "Valida CPF"
    Public Shared Function VerificaCPF(ByVal strCPFCliente As String) As Boolean

        '--Declaração das Variáveis
        Dim strCPFOriginal As String = strCPFCliente.Replace(".", "").Replace("-", "")
        Dim strCPF As String = Mid(strCPFOriginal, 1, 9)
        Dim strCPFTemp As String
        Dim intSoma As Integer
        Dim intResto As Integer
        Dim strDigito As String
        Dim intMultiplicador As Integer = 10
        Const constIntMultiplicador As Integer = 11
        Dim i As Integer
        '--------------------------

        For i = 0 To strCPF.ToString.Length - 1
            intSoma += CInt(strCPF.ToString.Chars(i).ToString) * intMultiplicador
            intMultiplicador -= 1
        Next

        If (intSoma Mod constIntMultiplicador) < 2 Then
            intResto = 0
        Else
            intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
        End If

        strDigito = intResto
        intSoma = 0

        strCPFTemp = strCPF & strDigito
        intMultiplicador = 11

        For i = 0 To strCPFTemp.Length - 1
            intSoma += CInt(strCPFTemp.Chars(i).ToString) * intMultiplicador
            intMultiplicador -= 1
        Next

        If (intSoma Mod constIntMultiplicador) < 2 Then
            intResto = 0
        Else
            intResto = constIntMultiplicador - (intSoma Mod constIntMultiplicador)
        End If

        strDigito &= intResto

        If strDigito = Mid(strCPFOriginal, 10, strCPFOriginal.Length) Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

#Region "Mes Por Extenso"
    Public Shared Function MesPorExtenso(ByVal Mes As String) As String
        Select Case Mes
            Case "1"
                Return "Janeiro"
            Case "2"
                Return "Fevereiro"
            Case "3"
                Return "Março"
            Case "4"
                Return "Abril"
            Case "5"
                Return "Maio"
            Case "6"
                Return "Junho"
            Case "7"
                Return "Julho"
            Case "8"
                Return "Agosto"
            Case "9"
                Return "Setembro"
            Case "10"
                Return "Outubro"
            Case "11"
                Return "Novembro"
            Case "12"
                Return "Dezembro"
        End Select
        Return String.Empty
    End Function
#End Region

#Region "Retorna Nome Cliente formatado com Cidade/UF e CPF/CNPJ"

    Public Shared Function FormatarListItemCliente(ByVal ObjetoCliente As Cliente, Optional ByVal ComEstado As Boolean = False) As System.Web.UI.WebControls.ListItem
        Dim strDescricao As String = Funcoes.AlinharEsquerda(ObjetoCliente.Nome, 28, "-")
        strDescricao &= "- " & Funcoes.AlinharEsquerda(ObjetoCliente.Cidade, 20, "-") & " " & ObjetoCliente.CodigoEstado
        strDescricao &= " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(ObjetoCliente.Codigo), 18, "-")
        strDescricao &= "-" & ObjetoCliente.CodigoEndereco.ToString() & IIf(Not String.IsNullOrWhiteSpace(ObjetoCliente.Reduzido), "-" & ObjetoCliente.Reduzido, "")

        Return New System.Web.UI.WebControls.ListItem(strDescricao, ObjetoCliente.Codigo & "-" & ObjetoCliente.CodigoEndereco.ToString() & IIf(ComEstado, "-" & ObjetoCliente.CodigoEstado, ""))
    End Function

    Public Shared Function getDescricaoCliente(ByVal codigo As String, ByVal endereco As Integer) As String
        Dim obj As New Cliente(codigo, endereco)
        Dim nome As String = String.Empty
        If Not String.IsNullOrWhiteSpace(obj.Codigo) Then
            nome = FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " - " & obj.CodigoEstado
        End If

        Return nome
    End Function
#End Region

#Region " Valida IE "

    '<DllImport("DllInscE32.dll")>
    'Public Shared Function ConsisteInscricaoEstadual(ByVal vInsc As String, ByVal vUF As String) As Integer
    'End Function

    'Shared Function ValidarIE(ByVal ie As String, ByVal uf As String) As Boolean
    '    Try
    '        Select Case ConsisteInscricaoEstadual(EliminarCaracteresEspeciais(ie), uf.ToUpper())
    '            Case 0
    '                ' IE correta
    '                Return True
    '            Case 1
    '                ' IE incorreta
    '                Return IIf(uf.ToUpper() = "PE" And ie.Replace(",", "").Replace(".", "").Replace("/", "").Replace("\", "").Replace("-", "").Replace(" ", "").Length = 9, ValidarIEPE(ie.Replace(",", "").Replace(".", "").Replace("/", "").Replace("\", "").Replace("-", "").Replace(" ", ""), uf.ToUpper()), False)
    '            Case Else
    '                ' Parametros incorretos
    '                Return False
    '        End Select
    '    Catch ex As Exception
    '        Throw New Exception("Problemas ao carregar a DLL DllInscE32.dll, utilizada na validação de Inscrição Estadual. Contate o CPD informando o erro. " & "[" & ex.Message & "]")
    '        Return False
    '    End Try
    'End Function

    Shared Function ValidarIE(ByVal ie As String, ByVal ufString As String) As Boolean
        Try
            Dim uf As UF
            If [Enum].TryParse(ufString, uf) Then
                Dim codigoUF As Integer = CInt(uf)
                Console.WriteLine(codigoUF) ' Saída: 41
            Else
                Throw New Exception("Sigla de UF inválida!")
            End If

            Dim validador As New IEValidator(ie, uf)
            Return validador.IsValid()

        Catch ex As Exception
            'Throw New Exception("Problemas ao carregar a DLL DllInscE32.dll, utilizada na validação de Inscrição Estadual. Contate o CPD informando o erro. " & "[" & ex.Message & "]")
            Throw New Exception("Problemas ao validar a inscrição estadual! Erro: " & ex.Message)
            Return False
        End Try
    End Function

    Shared Function ValidarIEPE(ByVal ie As String, ByVal uf As String) As Boolean
        'Dim ChecaInscrE As Boolean
        Dim strBase As String
        Dim strDigito1 As String
        Dim strDigito2 As String
        Dim intPos As Integer
        Dim intValor As Integer
        Dim intSoma As Integer
        Dim intResto As Integer


        strBase = Left(Trim(ie) & "0000000", 7)

        'Primeiro Digito Verificador
        intSoma = 0
        For intPos = 1 To 7
            intValor = Val(Mid$(strBase, intPos, 1))
            intValor = intValor * (9 - intPos)
            intSoma += intValor
        Next
        If intSoma < 11 Then
            strDigito1 = 0
        Else
            intResto = intSoma Mod 11
            strDigito1 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
        End If

        'Segundo Digito Verificador
        intSoma = 0
        For intPos = 1 To 8
            intValor = Val(Mid$(strBase & strDigito1, intPos, 1))
            intValor = intValor * (10 - intPos)
            intSoma += intValor
        Next
        If intSoma < 11 Then
            strDigito2 = 0
        Else
            intResto = intSoma Mod 11
            strDigito2 = Right(IIf(intResto < 2, "0", Str(11 - intResto)), 1)
        End If

        If ie = strBase & strDigito1 & strDigito2 Then
            Return True
        Else
            Return False
        End If

    End Function
#End Region

#Region "Busca Empresa do Menu"
    Public Shared Sub VerificaUnidadeDeNegocio(ByRef ddlUnidade As DropDownList, ByRef ddlEmpresa As DropDownList, Optional ByVal InsereLinhaemBranco As Boolean = True)
        Dim Banco As New AcessaBanco
        Dim Sql As String = ""
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade,"
        Sql &= "        isnull(AcessoEmpresa, '') as AcessoEmpresa,"
        Sql &= "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa"
        Sql &= " from Usuarios"
        Sql &= " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "UnidadeDeNegocio").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            If Dr("AcessoUnidade").ToString.Length > 0 Then
                Dim Codigo As String
                Dim Descricao As String
                Dim Nome As String
                Dim Cidade As String
                Dim Cnpj As String
                Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado "
                Sql &= " FROM   GruposXEmpresas INNER JOIN"
                Sql &= " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id"
                Sql &= " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "' "

                ddlEmpresa.Items.Clear()

                For Each DrE As DataRow In Banco.ConsultaDataSet(Sql, "Empresa").Tables(0).Rows
                    Codigo = DrE("Codigo") & "-" & CStr(DrE("Endereco_Id"))

                    Cnpj = Funcoes.FormatarCpfCnpj(DrE("Codigo"))
                    Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

                    Nome = Funcoes.AlinharEsquerda(DrE("Nome"), 28, ".")
                    Cidade = Funcoes.AlinharEsquerda(DrE("Cidade"), 20, ".")
                    Descricao = Nome & " - " & Cidade & " " & DrE("Estado") & " " & Cnpj & "-" & CStr(DrE("Endereco_Id")) & "-" & DrE("Reduzido")

                    ddlEmpresa.Items.Add(New System.Web.UI.WebControls.ListItem(Descricao, Codigo))
                Next

                If InsereLinhaemBranco Then ddlEmpresa.Items.Insert(0, "")
            End If

            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Public Shared Sub VerificaUnidadeDeNegocio(ByRef ddlUnidade As DropDownList, ByRef lstEmpresa As ListBox, Optional ByVal InsereLinhaemBranco As Boolean = True)
        Dim Banco As New AcessaBanco
        Dim Sql As String = ""
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade,"
        Sql &= "        isnull(AcessoEmpresa, '') as AcessoEmpresa,"
        Sql &= "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa"
        Sql &= " from Usuarios"
        Sql &= " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "UnidadeDeNegocio").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            If Dr("AcessoUnidade").ToString.Length > 0 Then
                Dim Codigo As String
                Dim Descricao As String
                Dim Nome As String
                Dim Cidade As String
                Dim Cnpj As String
                Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado "
                Sql &= " FROM   GruposXEmpresas INNER JOIN"
                Sql &= " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id"
                Sql &= " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "' "

                For Each DrE As DataRow In Banco.ConsultaDataSet(Sql, "Empresa").Tables(0).Rows
                    Codigo = DrE("Codigo") & "-" & CStr(DrE("Endereco_Id"))

                    Cnpj = Funcoes.FormatarCpfCnpj(DrE("Codigo"))
                    Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

                    Nome = Funcoes.AlinharEsquerda(DrE("Nome"), 28, ".")
                    Cidade = Funcoes.AlinharEsquerda(DrE("Cidade"), 20, ".")
                    Descricao = Nome & " - " & Cidade & " " & DrE("Estado") & " " & Cnpj & "-" & CStr(DrE("Endereco_Id")) & "-" & DrE("Reduzido")

                    lstEmpresa.Items.Add(New System.Web.UI.WebControls.ListItem(Descricao, Codigo))
                Next

                If InsereLinhaemBranco Then lstEmpresa.Items.Insert(0, "")
            End If

            lstEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Public Shared Sub VerificaEmpresa(ByRef ddlEmpresa As DropDownList)
        Dim Banco As New AcessaBanco
        Dim Sql As String = ""
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              " from Usuarios" & vbCrLf &
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Public Shared Sub VerificaEmpresa(ByRef lstEmpresa As ListBox)
        Dim Banco As New AcessaBanco
        Dim Sql As String = ""
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              " from Usuarios" & vbCrLf &
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            lstEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub
#End Region

#Region "Verifica Conexão com a Internet"
    Shared Function VerificaConexaoInternet() As Boolean
        Return My.Computer.Network.IsAvailable
    End Function
#End Region

#Region "Verifica CEP"
    Public Shared Function BuscaCep(ByVal cep As String) As DataSet
        Try
            Dim ds As New DataSet
            If VerificaPing("https://viacep.com.br/") Then
                ds.ReadXml("https://viacep.com.br/ws/" + cep.Replace("-", "").Trim() + "/xml")
            ElseIf VerificaPing("viacep.com.br") Then
                ds.ReadXml("http://viacep.com.br/ws/" & cep.Replace("-", "").Trim() & "/xml")
            End If
            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function ConsultaCep(ByVal estado As String, ByVal cidade As String, ByVal endereco As String) As DataSet
        Try
            Dim ds As New DataSet
            If VerificaPing("https://viacep.com.br/") Then
                ds.ReadXml(String.Format("https://viacep.com.br/ws/{0}/{1}/{2}/xml", estado.Trim(), cidade.Trim(), endereco.Trim()))
            ElseIf VerificaPing("viacep.com.br") Then
                ds.ReadXml(String.Format("https://viacep.com.br/ws/{0}/{1}/{2}/xml", estado.Trim(), cidade.Trim(), endereco.Trim()))
            End If
            Return ds
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Shared Function VerificaPing(ByVal url) As Boolean
        Try
            My.Computer.Network.Ping(url)
            Return True
        Catch pingException As System.Net.NetworkInformation.PingException
            Return False
        Catch ex As Exception
            Return False
        End Try
    End Function

#End Region

    Public Shared Function FormatarValor(ByVal Valor As Decimal, ByVal Tamanho As Integer, ByVal CasasDecimais As Integer) As String
        Dim cd As Integer = CasasDecimais
        Dim n As String = "0123456789" '**  Valores Validos  ** 

        Dim d As String = Math.Round(Valor, CasasDecimais).ToString     '**  Valor do Campo   ** 
        Dim l As Integer = d.Length      '** Tamanho do Campo  **
        Dim r As String = ""             '**    Resultado      ** 

        '** 1 ******* caso o tamanho seja maior que zero ira percorrer a string retirando valores invalidos
        Dim s As String = ""
        Dim P1 As Integer = 0
        While P1 <= (l - 1)
            If n.Contains(d.Substring(P1, 1)) Then
                s &= d.Substring(P1, 1)
            End If
            P1 += 1
        End While

        l = s.Length
        '*** 1 *********************************************************************************************

        ' caso a string seja maior q o tamanho passado ele conta a string no tamanho referenciado
        If (l > Tamanho) Then
            l = Tamanho
            s = s.Substring(0, Tamanho)
        End If

        '*** se o tamanho for maior q o numero de casas decimais, já coloca a virgula
        If cd > 0 Then
            If l > cd Then
                r = s.Substring(0, l - cd) & "," & s.Substring(l - cd, cd)
            Else
                '**** se o tamanho for igual as casas decimais coloca o 0 e a virgula na frente do valor
                If (l = cd) Then
                    r = "0," + s
                Else
                    Dim P2 As Integer = l
                    While P2 < cd
                        s = "0" & s
                        P2 += 1
                    End While
                    r = "0," + s
                End If
            End If
        Else
            r = s
        End If

        l = r.Length
        If (l > cd + 3) Then
            Dim j As Integer
            Dim w As String
            Dim wa As String
            Dim wb As String
            If (cd = 0) Then
                j = l Mod 3
                w = r.Substring(0, j)
                wa = r.Substring(j, l - j)
                wb = ""
            Else
                j = (l - cd - 1) Mod 3
                w = r.Substring(0, j)
                wa = r.Substring(j, l - j - cd - 1)
                wb = r.Substring(l - cd - 1, cd + 1)
            End If

            Dim wat As Integer = wa.Length
            Dim k As Decimal = (wat / 3)

            If k > 0 Then
                Dim P3 As Integer = 0
                While P3 < k
                    w &= "." & wa.Substring(P3 * 3, 3)
                    P3 += 1
                End While
            Else
                w &= wa
            End If

            If w.Substring(0, 1) = "." Then
                Dim x As Integer = w.Length
                w = w.Substring(1, x - 1)
            End If
            r = w + wb
        End If

        Return r
    End Function

    Public Shared Sub AbrirArquivo(ByVal page As System.Web.UI.Page, ByVal NomeArquivo As String)
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo & "');", True)
    End Sub

    Public Shared Sub ImprimirArquivo(ByVal page As System.Web.UI.Page, ByVal fileName As String, ByVal nf As [Lib].Negocio.NotaFiscal)

        Try

            Dim fileInfo As FileInfo = New FileInfo(fileName)
            If Not fileInfo.Exists Then
                Throw New Exception("Não foi possível encontrar o arquivo!")
                Exit Sub
            End If

            Dim impressora As String
            If nf.CodigoEmpresa = "04854422000266" Then
                impressora = "\\10.1.1.71\HP LaserJet P1005"
            ElseIf Mid(nf.CodigoEmpresa, 1, 8) = "04854422" Then
                impressora = "\\10.1.1.46\HP LaserJet M1120 MFP"
            ElseIf Mid(nf.CodigoEmpresa, 1, 8) = "40938762" Or Mid(nf.CodigoEmpresa, 1, 8) = "49673784" Then
                impressora = "Brother DCP-8157DN FISCAL"
            ElseIf Mid(nf.CodigoEmpresa, 1, 8) = "05366261" Or Mid(nf.CodigoEmpresa, 1, 8) = "38198213" Then
                impressora = "Brother DCP-8157DN FISCAL"
            Else
                impressora = "Laser"
            End If

            For Each printer As String In PrinterSettings.InstalledPrinters
                Console.WriteLine(printer)
            Next

            ' Carregar o documento PDF
            'Metodo abaixo imprime na impressora padrão do windows
            Using document As PdfDocument = PdfDocument.Load(fileName)

                ' Criar o documento de impressão
                Using printDocument As PrintDocument = document.CreatePrintDocument()
                    ' Configurar as definições da impressora
                    printDocument.PrinterSettings = New PrinterSettings() With {
                    .PrintFileName = fileInfo.Name,
                    .PrinterName = impressora
                }

                    printDocument.DocumentName = fileInfo.Name
                    printDocument.PrintController = New StandardPrintController() ' Usar o controlador padrão para impressão

                    ' Imprimir o documento
                    printDocument.Print()

                    ' Adicionar um atraso (por exemplo, 5 segundos) para a próxima impressão
                    System.Threading.Thread.Sleep(3000) ' Atraso de 5000 milissegundos (5 segundos)

                End Using
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Public Shared Sub CriarPastaLivroFiscal(diretorio As String)
        ' Verifica se a pasta já existe
        If Not Directory.Exists(diretorio) Then
            ' Cria a pasta
            Directory.CreateDirectory(diretorio)
        End If
    End Sub

    Public Shared Sub AbrirExcel(ByVal page As System.Web.UI.Page, ByVal NomeArquivo As String)
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, "window.location = '" & NomeArquivo & "';", True)
    End Sub

    Public Shared Sub DownloadZIP(ByVal page As System.Web.UI.Page, ByVal directoryPath As String)
        Try
            If Not Directory.Exists(directoryPath) Then
                Throw New Exception("Não foi possível encontrar o diretório!")
                Exit Sub
            End If

            Dim url As String = HttpContext.Current.Request.Url.AbsoluteUri
            Dim domainServer As String
            Dim bServidor As Boolean

            If url.ToUpper().Contains("/NGS/") Or url.ToUpper().Contains("/NGSTESTE/") Then
                bServidor = True
            End If

            If HttpContext.Current.Request.Url.Query.Length > 0 Then
                domainServer = url.Replace(HttpContext.Current.Request.Url.Query, "").Replace(HttpContext.Current.Request.Url.LocalPath, "")
            Else
                domainServer = url.Replace(HttpContext.Current.Request.Url.LocalPath, "")
            End If

            Dim dirEncode As String = EncodeBase64(directoryPath)
            Dim urlZIP As String

            If bServidor Then
                urlZIP = String.Format("{0}/ngsApi/api/Download/Zip?path={1}", domainServer, dirEncode)
            Else
                urlZIP = String.Format("{0}/api/Download/Zip?path={1}", domainServer, dirEncode)
            End If

            'Chamar API
            AbrirArquivo(page, urlZIP)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Shared Function DecodeBase64(ByVal base64EncodedData As String) As String
        Dim base64EncodedBytes As Byte() = Convert.FromBase64String(base64EncodedData)
        Return Encoding.UTF8.GetString(base64EncodedBytes)
    End Function

    Public Shared Function EncodeBase64(ByVal plainText As String) As String
        Dim plainTextBytes As Byte() = Encoding.UTF8.GetBytes(plainText)
        Return Convert.ToBase64String(plainTextBytes)
    End Function

    Private Sub Download(ByVal page As System.Web.UI.Page, ByVal fileName As String)
        Try
            Dim fileInfo As FileInfo = New FileInfo(fileName)
            If Not fileInfo.Exists Then
                MsgBox(page, "Não foi possível encontrar o arquivo!")
                Exit Sub
            End If
            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=" & fileInfo.Name)
            HttpContext.Current.Response.AddHeader("Content-Length", fileInfo.Length.ToString())
            HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            HttpContext.Current.Response.WriteFile(fileInfo.FullName)
            HttpContext.Current.Response.End()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Shared Sub OpenManyPDF(ByVal page As System.Web.UI.Page, ByVal strJavaScript As String)
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

#Region "Conversão de Moedas"

    Public Shared Function PegarValorConversao(ByVal CodigoIndexador As Integer, ByVal Data As DateTime) As Double
        Dim objBanco As New AcessaBanco()
        Dim Indice As Decimal = 0

        Dim strSQL As String = "SELECT Indice " &
                               "  FROM Cotacoes " &
                               " WHERE Indexador_Id = " & CodigoIndexador.ToString() & " " &
                               "   AND Data_Id = '" & Data.ToString("yyyy-MM-dd") & "'"

        Dim dsConversao As DataSet = objBanco.ConsultaDataSet(strSQL, "Cotacoes")

        If dsConversao.Tables(0).Rows.Count > 0 Then
            Indice = dsConversao.Tables(0).Rows(0)("Indice")
        End If

        If Indice = 0 Then
            Throw New Exception(" |Não existe data de cotação cadastrada para o dia " & Data.ToString("dd/MM/yyyy") & "|")
        End If
        Return Indice
    End Function

    Public Shared Function ConverteParaMoedaOficial(ByVal Valor As Double, ByVal CodigoIndexador As Integer, ByVal Data As DateTime, Optional ByVal Arredondar As Boolean = False, Optional ByVal Trucar As Boolean = False, Optional ByVal CasasDecimais As Integer = 0) As Double
        If Arredondar Then
            Return Math.Round((Valor * PegarValorConversao(CodigoIndexador, Data)), CasasDecimais, MidpointRounding.AwayFromZero)
        ElseIf Trucar Then
            Return Math.Round((Valor * PegarValorConversao(CodigoIndexador, Data)), CasasDecimais, MidpointRounding.ToEven)
        Else
            Return Valor * PegarValorConversao(CodigoIndexador, Data)
        End If
    End Function

    Public Shared Function ConverteParaMoedaExtrangeira(ByVal Valor As Double, ByVal CodigoIndexador As Integer, ByVal Data As DateTime, Optional ByVal Arredondar As Boolean = False, Optional ByVal Trucar As Boolean = False, Optional ByVal CasasDecimais As Integer = 0) As Double
        If Arredondar Then
            Return Math.Round((Valor / PegarValorConversao(CodigoIndexador, Data)), CasasDecimais, MidpointRounding.AwayFromZero)
        ElseIf Trucar Then
            Return Math.Round((Valor / PegarValorConversao(CodigoIndexador, Data)), CasasDecimais, MidpointRounding.ToEven)
        Else
            Return Valor / PegarValorConversao(CodigoIndexador, Data)
        End If
    End Function

    Public Shared Function ConverteMoeda(ByVal Valor As Decimal, ByVal IndiceCotacao As Decimal, ByVal MoedaDestino As eTiposMoeda, Optional ByVal Arredondar As Boolean = False, Optional ByVal Trucar As Boolean = False, Optional ByVal CasasDecimais As Integer = 0) As Decimal

        'Evitar erro de divisão por Zero.
        If IndiceCotacao > 0 Then
            If MoedaDestino = eTiposMoeda.MoedaEstrangeira Then
                If Arredondar Then
                    Return Math.Round((Valor / IndiceCotacao), CasasDecimais, MidpointRounding.AwayFromZero)
                ElseIf Trucar Then
                    Return Math.Round((Valor / IndiceCotacao), CasasDecimais, MidpointRounding.ToEven)
                Else
                    Return Valor / IndiceCotacao
                End If
            Else
                If Arredondar Then
                    Return Math.Round((Valor * IndiceCotacao), CasasDecimais, MidpointRounding.AwayFromZero)
                ElseIf Trucar Then
                    Return Math.Round((Valor * IndiceCotacao), CasasDecimais, MidpointRounding.ToEven)
                Else
                    Return Valor * IndiceCotacao
                End If
            End If
        End If
        Return 0
    End Function

#End Region

    Public Shared Sub AdicionarClienteAoDDL(ByVal pddl As DropDownList, ByVal pCodigoCliente As String, ByVal pEndCliente As Integer)
        Dim cli As New Cliente(pCodigoCliente, pEndCliente)
        pddl.Items.Add(New System.Web.UI.WebControls.ListItem(Funcoes.AlinharEsquerda(cli.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(cli.Cidade, 20, ".") & " " & cli.CodigoEstado & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(cli.Codigo), 18, ".") & "-" & CStr(cli.CodigoEndereco), cli.Codigo & "-" & CStr(cli.CodigoEndereco)))
    End Sub

    Public Shared Sub AdicionarClienteAoDDL(ByVal pddl As DropDownList, ByVal cli As Cliente)
        pddl.Items.Add(New System.Web.UI.WebControls.ListItem(Funcoes.AlinharEsquerda(cli.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(cli.Cidade, 20, ".") & " " & cli.CodigoEstado & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(cli.Codigo), 18, ".") & "-" & CStr(cli.CodigoEndereco), cli.Codigo & "-" & CStr(cli.CodigoEndereco)))
    End Sub

    Public Shared Sub FormatarClienteTXT(ByVal txt As TextBox, ByVal cli As Cliente)
        txt.Text = Funcoes.AlinharEsquerda(cli.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(cli.Cidade, 20, ".") & " " & cli.CodigoEstado & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(cli.Codigo), 18, ".") & "-" & CStr(cli.CodigoEndereco)
    End Sub

    Public Shared Sub BindExcelOffice(ByVal page As Page, ByVal ds As DataSet, ByVal TituloAba As String, ByVal colunas As MDTwoValues(Of String, String, eTipoCampo), Optional ByVal Cabecalho As String() = Nothing)
        If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Nenhum resultado encontrado.")
        Else
            'Emitir Excel.xsls do office/ Relatorio padrao em lista.
            Dim rowIndex As Integer = 1
            Dim columnIndex As Integer = 1

            Dim rowIndexCabecalhoColunas As Integer = 1

            Try
                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then File.Delete(fileName)

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)

                        'criando aba da planilha.
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)

                        For i As Integer = 0 To Cabecalho.Length - 1
                            'Inserindo o cabeçalho.

                            worksheet.Cells(rowIndex, columnIndex).Value = Cabecalho(i)
                            rowIndex += 1
                        Next

                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        Next

                        rowIndexCabecalhoColunas = rowIndex

                        rowIndex += 1

                        For i As Integer = 1 To Cabecalho.Length
                            'aplicando formatação nas células do cabeçalho

                            Using range = worksheet.Cells(i, 1, i, ds.Tables(0).Columns.Count)
                                range.Merge = True
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(System.Drawing.Color.White)
                                range.Style.Font.Size = 16
                                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                                range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(0, 0, 0))
                            End Using
                        Next

                        Using range = worksheet.Cells(Cabecalho.Length + 1, 1, Cabecalho.Length + 1, ds.Tables(0).Columns.Count)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(System.Drawing.Color.White)
                        End Using

                        ' Exportando conteúdo da planilha com os dados da Tabela.
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                If IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Contains(",") Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName).ToString().Replace(".", ""))
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                End If

                                'Formatacoes de valores
                                If colunas IsNot Nothing Then
                                    For Each key In colunas.Keys
                                        If key = col.ColumnName Then
                                            If colunas.GetValue2(key) = eTipoCampo.Numerico Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                                If rowIndexCabecalhoColunas > 1 Then
                                                    worksheet.Cells(rowIndexCabecalhoColunas, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                                End If
                                                Exit For
                                            ElseIf colunas.GetValue2(key) = eTipoCampo.ValorComTotalizador OrElse colunas.GetValue2(key) = eTipoCampo.ValorSemTotalizador Then
                                                If Not String.IsNullOrWhiteSpace(colunas.GetValue1(key)) Then
                                                    worksheet.Cells(rowIndex, columnIndex).Formula = String.Format(colunas.GetValue1(key), rowIndex)
                                                End If
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                If rowIndexCabecalhoColunas > 1 Then
                                                    worksheet.Cells(rowIndexCabecalhoColunas, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                                End If
                                                Exit For
                                            ElseIf colunas.GetValue2(key) = eTipoCampo.Data Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                                Exit For
                                            ElseIf colunas.GetValue2(key) = eTipoCampo.ValorSemTotalizadorCom3CasasDecimais Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,###0.000_ ;[Red]-#,###0.000"
                                                If rowIndexCabecalhoColunas > 1 Then
                                                    worksheet.Cells(rowIndexCabecalhoColunas, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                                End If
                                                Exit For
                                            ElseIf colunas.GetValue2(key) = eTipoCampo.ValorSemTotalizadorCom4CasasDecimais Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,###0.0000_ ;[Red]-#,###0.0000"
                                                If rowIndexCabecalhoColunas > 1 Then
                                                    worksheet.Cells(rowIndexCabecalhoColunas, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                                End If
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                                columnIndex += 1
                            Next

                            rowIndexCabecalhoColunas = 0

                            'Formatacoes de celulas
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        'aplicando formatação nas células do RODAPÉ
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(System.Drawing.Color.White)
                            range.Style.Border.Top.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(0, 0, 0))
                            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
                            range.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(0, 0, 0))
                        End Using

                        'Soma dos campos de valores
                        columnIndex = 1
                        worksheet.Cells(rowIndex, columnIndex).Value = "TOTAL"
                        For Each col As DataColumn In ds.Tables(0).Columns
                            For Each key In colunas.Keys
                                If key = col.ColumnName Then
                                    If colunas.GetValue2(key) = eTipoCampo.ValorComTotalizador Then
                                        worksheet.Cells(rowIndex, columnIndex).Formula = String.Format("SUM({0}:{1})", worksheet.Cells(2, columnIndex).Address, worksheet.Cells(rowIndex - 1, columnIndex).Address)
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;-#,##0.00"
                                        Exit For
                                    End If
                                End If
                            Next
                            columnIndex += 1
                        Next

                        'criando auto filtro na planilha
                        'worksheet.Cells(Cabecalho.Length + 1, 1, Cabecalho.Length + 1, columnIndex - 1).AutoFilter = True

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando primeira linham
                        worksheet.View.FreezePanes(Cabecalho.Length + 2, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                AbrirExcel(page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                Throw New Exception(ex.Message & " - linha: " & rowIndex & " - coluna: " & columnIndex)
            End Try
        End If
    End Sub

    ''' <summary>
    ''' FAZ O CARREGAMENTO DO EXCELOFFICE EM LISTA SIMPLES
    ''' </summary>
    ''' <param name="page">PAGINA ATUAL - ME.PAGE</param>
    ''' <param name="ds">DATASET RESPONSÁVEL PELO CARREGAMENTO DO RELATORIO - AS COLUNAS DEVEM ESTAR RENOMEADAS</param>
    ''' <param name="TituloAba">TITULO DA ABA DO EXCELOFFICE.</param>
    ''' <param name="colunas">SE EXISTIR COLUNAS COM TIPOS(NUMERICO, VALOR OU DATA), 
    ''' CRIE UM DICIONARIO CONTENDO O NOME DA COLUNA E O TIPO DE COLUNA. POIS ASSIM O CAMPO SERÁ PREENCHIDO COM A FORMATAÇÃO CORRETA
    ''' SE O DICIONARIO CONTER CAMPOS DO TIPO VALOR O MESMO TERÁ UM TOTALIZADOR AO FINAL.</param>
    ''' <remarks></remarks>
    Public Shared Sub BindExcelOffice(ByVal page As Page, ByVal ds As DataSet, ByVal TituloAba As String, Optional ByVal colunas As Dictionary(Of String, eTipoCampo) = Nothing, Optional ByVal viewAllTables As Boolean = False, Optional ByVal cabecalho As String = "")
        If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Nenhum resultado encontrado!")
        Else
            Try
                Dim rowIndex As Integer = 1
                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then File.Delete(fileName)

                'emitir excel.xsls do office / relatório padrão em lista
                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando aba da planilha
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)
                        Dim columnIndex As Integer = 1

                        'Inserindo o cabeçalho.
                        If Not String.IsNullOrWhiteSpace(cabecalho) Then
                            worksheet.Cells(rowIndex, columnIndex).Value = cabecalho
                            worksheet.Cells(1, 1, 1, ds.Tables(0).Columns.Count).Merge = True
                            worksheet.Cells(1, 1, 1, ds.Tables(0).Columns.Count).Style.Font.Bold = True
                            rowIndex += 2
                        End If

                        For Each dt As DataTable In ds.Tables
                            If viewAllTables Then
                                worksheet.Cells(rowIndex, columnIndex).Value = dt.TableName
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                                worksheet.Cells(rowIndex, columnIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                                worksheet.Cells(rowIndex, columnIndex).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189))
                                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(System.Drawing.Color.White)
                                rowIndex += 1
                            End If

                            'inserindo o cabeçalho
                            For Each col As DataColumn In dt.Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            Next

                            'aplicando formatação nas células do cabeçalho
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(System.Drawing.Color.White)
                            End Using

                            rowIndex += 1

                            'exportando conteúdo da planilha com os dados da tabela
                            For Each row As DataRow In dt.Rows
                                columnIndex = 1
                                For Each col As DataColumn In dt.Columns

                                    If col.ColumnName.ToUpper = "CHAVENFE" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    ElseIf IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Trim.Contains(",") Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName).ToString().Replace(".", ""))
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    End If

                                    'formatações de valores
                                    If colunas IsNot Nothing Then
                                        For Each coluna In colunas
                                            If coluna.Key = col.ColumnName Then
                                                If coluna.Value = eTipoCampo.Numerico Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                                    Exit For
                                                ElseIf coluna.Value = eTipoCampo.ValorComTotalizador OrElse coluna.Value = eTipoCampo.ValorSemTotalizador Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                    Exit For
                                                ElseIf coluna.Value = eTipoCampo.Data Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                                    Exit For
                                                ElseIf coluna.Value = eTipoCampo.ValorSemTotalizadorCom3CasasDecimais Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,###0.000_ ;[Red]-#,###0.000"
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    Else
                                        If IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString.Contains(",") Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                        ElseIf IsDate(row(col.ColumnName)) Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                        End If
                                    End If
                                    columnIndex += 1
                                Next

                                'formatações de celulas
                                If rowIndex Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(153, 204, 255))
                                    End Using
                                End If
                                rowIndex += 1
                            Next

                            'aplicando formatação nas células do rodapé
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(System.Drawing.Color.White)
                            End Using

                            columnIndex = 1

                            'soma dos campos de valores
                            If colunas IsNot Nothing Then

                                For Each col In colunas
                                    If col.Value = eTipoCampo.ValorComTotalizador Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "TOTAL"
                                        Exit For
                                    End If
                                Next

                                For Each col As DataColumn In dt.Columns
                                    For Each coluna In colunas
                                        If coluna.Key = col.ColumnName Then
                                            If coluna.Value = eTipoCampo.ValorComTotalizador Then
                                                worksheet.Cells(rowIndex, columnIndex).Formula = "SUM(" & worksheet.Cells(1, columnIndex).Address & ":" & worksheet.Cells(rowIndex - 1, columnIndex).Address & ")"
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    columnIndex += 1
                                Next
                            End If

                            If Not viewAllTables Then
                                Exit For
                            End If

                            rowIndex += 2
                        Next

                        'criando auto filtro na planilha
                        If Not viewAllTables Then
                            worksheet.Cells(IIf(Not String.IsNullOrWhiteSpace(cabecalho), 3, 1), 1, IIf(Not String.IsNullOrWhiteSpace(cabecalho), 3, 1), ds.Tables(0).Columns.Count).AutoFilter = True
                        End If

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando primeira linham
                        worksheet.View.FreezePanes(IIf(Not String.IsNullOrEmpty(cabecalho), 4, 2), 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                AbrirExcel(page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
    End Sub

    Public Shared Sub BindWordOffice(ByVal page As Page, ByVal NomeDoArquivoModelo As String, ByVal Substituicoes As Dictionary(Of String, Object), ByRef CaminhoDoArqModificado As String)

        Dim CaminhoDoModelo As String = HttpContext.Current.Server.MapPath("~/Arquivos/" & NomeDoArquivoModelo & ".docx")
        CaminhoDoArqModificado = HttpContext.Current.Server.MapPath("~/Files/" & NomeDoArquivoModelo & "-" & Now.ToString("dd.MM.yyy-HH.mm.ss") & ".docx")

        If Not System.IO.File.Exists(CaminhoDoModelo) Then
            Throw New Exception("Arquivo: " & CaminhoDoModelo & " não encontrado!")
        End If

        File.Copy(CaminhoDoModelo, CaminhoDoArqModificado)

        Using WordDoc As WordprocessingDocument = WordprocessingDocument.Open(CaminhoDoArqModificado, True)

            Dim DocText As String = String.Empty

            Using sr As StreamReader = New StreamReader(WordDoc.MainDocumentPart.GetStream())
                DocText = sr.ReadToEnd()
            End Using

            Dim RegexText As Regex
            For Each it In Substituicoes
                RegexText = New Regex(it.Key, RegexOptions.IgnoreCase)
                If it.Key = "#TABELA_DE_MERCADORIAS#" Then
                    DocText = RegexText.Replace(DocText, InserirTabelaNoDoc(it.Value))
                Else
                    'DocText = RegexText.Replace(DocText, it.Value)
                    DocText = RegexText.Replace(DocText, InserirQuebrasDeLinhaNoDoc(it.Value))
                End If
            Next

            Using sw As StreamWriter = New StreamWriter(WordDoc.MainDocumentPart.GetStream(FileMode.Create))
                sw.Write(DocText)
            End Using
        End Using
    End Sub

    Private Shared Function InserirQuebrasDeLinhaNoDoc(ByVal texto As String) As String
        ' Verifica se o texto contém quebras de linha
        If texto.Contains(vbCrLf) OrElse texto.Contains(vbLf) Then
            ' Divide o texto em partes com base nas quebras de linha
            Dim partes As String() = texto.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)

            ' Reconstrói o texto usando quebras de linha OpenXML
            Dim resultado As New StringBuilder()
            For i As Integer = 0 To partes.Length - 1
                resultado.Append(partes(i)) ' Adiciona a parte atual
                If i < partes.Length - 1 Then
                    resultado.AppendLine("<w:br/>") ' Adiciona a marcação de quebra de linha do Word
                End If
            Next

            ' Retorna o texto processado
            Return resultado.ToString()
        End If

        ' Se não houver quebras, retorna o texto original
        Return texto
    End Function

    Private Shared Function InserirTabelaNoDoc(ByVal ds As DataSet) As String

        'Cria um novo parágrafo'
        Dim p As New DocumentFormat.OpenXml.Wordprocessing.Paragraph()

        'Cria uma tabela.
        Dim tbl As New DocumentFormat.OpenXml.Wordprocessing.Table()

        'Cria uma Propriedade e estilo para a tabela.
        Dim tableProp As New DocumentFormat.OpenXml.Wordprocessing.TableProperties()
        Dim tableStyle As New DocumentFormat.OpenXml.Wordprocessing.TableStyle() With {.Val = "TableGrid"}

        'Cria uma largura de tabela 100% para que utilize toda a largura da página.
        Dim tableWidth As New TableWidth() With {.Width = "5000", .Type = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues.Pct}

        'Aplica na tabela
        tableProp.Append(tableStyle, tableWidth)
        tbl.AppendChild(tableProp)

        'Cria a 1 linha que serão os títulos das colunas
        Dim tr As New DocumentFormat.OpenXml.Wordprocessing.TableRow
        'Adiciona as células com os títulos das colunas'
        For Each row In ds.Tables(0).Columns
            Dim tc As New DocumentFormat.OpenXml.Wordprocessing.TableCell
            tc.Append(New Paragraph(New Run(New Text(row.ToString))))
            ' Assume you want columns that are automatically sized.
            tc.Append(New TableCellProperties(
                New TableCellWidth With {.Type = TableWidthUnitValues.Auto}))
            tr.Append(tc)
        Next
        tbl.Append(tr)

        'Adiciona as células com o conteúdo das colunas
        For i = 0 To ds.Tables(0).Rows.Count - 1
            Dim trC As New DocumentFormat.OpenXml.Wordprocessing.TableRow
            For Each row In ds.Tables(0).Columns
                Dim tc As New DocumentFormat.OpenXml.Wordprocessing.TableCell
                tc.Append(New Paragraph(New Run(New Text(ds.Tables(0).Rows(i)(row.ToString)))))
                tc.Append(New TableCellProperties(
                    New TableCellWidth With {.Type = TableWidthUnitValues.Auto}))
                trC.Append(tc)
            Next
            tbl.Append(trC)
        Next

        p.Append(tbl)

        Return p.InnerXml
    End Function

    ''' <summary>
    ''' FAZ O CARREGAMENTO DO CRYSTAL
    ''' </summary>
    ''' <param name="page">PAGINA ATUAL - ME.PAGE</param>
    ''' <param name="ds">DATASET RESPONSÁVEL PELO CARREGAMENTO DO RELATORIO, AS TABELAS E COLUNAS DEVEM CONTER O MESMO NOME E/OU TIPO DO DATASET FÍSICO (CRYSTAL)</param>
    ''' <param name="NameCrystal">INFORME O NOME DO RELATÓRIO A SER PREENCHIDO SEM O .RPT EXº CR_EXEMPLO</param>
    ''' <param name="tipo">INFORME O TIPO DE RELATORIO A SER PREENCHIDO PELO CRYSTAL - EEXPORTTYPE.PDF, EEXPORTTYPE.EXCELCRYSTAL OU EEXPORTTYPE.EXCELOFFICE</param>
    ''' <param name="paramentros">OPCIONAL: SE O RELATÓRIO CONTER PARÂMETROS, FAZER O PREENCHIMENTO DO MESMO.</param>
    ''' <remarks></remarks>
    Public Shared Sub BindReport(ByVal page As Page, ByVal ds As DataSet, ByVal NameCrystal As String, ByVal tipo As eExportType, Optional ByVal paramentros As Dictionary(Of String, Object) = Nothing, Optional mostrarEmBranco As Boolean = False, Optional dtLogo As String = "",
                                 Optional ByVal codEmp As String = "", Optional ByVal endEmp As String = "")
        If Not mostrarEmBranco AndAlso (ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0) Then
            Throw New Exception("NENHUM RESULTADO ENCONTRADO!")
        Else
            'ADD LOGOTIPO
            Dim dt As DataTable = ds.Tables.Add(IIf(String.IsNullOrWhiteSpace(dtLogo), "Logo", dtLogo))
            Dim objEmpresa As Cliente = New Cliente(IIf(Not String.IsNullOrWhiteSpace(codEmp), codEmp, UsuarioServidor.CodigoEmpresa), IIf(Not String.IsNullOrWhiteSpace(endEmp), endEmp, UsuarioServidor.EnderecoEmpresa))

            dt.Columns.Add("Imagem", GetType(System.Byte()))
            dt.Columns.Add("Empresa", GetType(System.String))
            dt.Columns.Add("Cidade", GetType(System.String))
            dt.Columns.Add("Cnpj", GetType(System.String))
            dt.Columns.Add("Usuario", GetType(System.String))


            Dim drImagem As DataRow = dt.NewRow()
            Dim strCaminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & IIf(Not String.IsNullOrWhiteSpace(codEmp), objEmpresa.Imagem, HttpContext.Current.Session("ssImagemEmpresa")))
            drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)

            drImagem("Empresa") = objEmpresa.Nome.Trim()

            If objEmpresa.Nome.Contains("NGS") Then

                drImagem("Cidade") = "CURITIBA/PR" & objEmpresa.Estado.Codigo.Trim()
                drImagem("Cnpj") = "CNPJ: 27.153.202/0001-20"
            Else

                drImagem("Cidade") = objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim()
                drImagem("Cnpj") = "CNPJ: " & FormatarCpfCnpj(objEmpresa.Codigo)
            End If

            drImagem("Usuario") = UsuarioServidor.NomeUsuario
            dt.Rows.Add(drImagem)

            Dim rpt As New ReportDocument()

            Try
                rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/" & NameCrystal & ".rpt")
                rpt.Load(rpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim extensao As String = ""
                If tipo = eExportType.PDF Then
                    extensao = ".pdf"
                ElseIf tipo = eExportType.DOC Then
                    extensao = ".doc"
                Else
                    extensao = ".xls"
                End If

                'If eExportType.PDF Then
                Dim fileName As String = "Files/" & Funcoes.GeraNomeArquivo & extensao

                Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(fileName)
                rpt.SetDataSource(ds)

                'rpt.SetDataSource() 

                If paramentros IsNot Nothing Then BindParameters(rpt, paramentros)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                If tipo = eExportType.PDF Then
                    rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)
                ElseIf tipo = eExportType.DOC Then
                    rpt.ExportToDisk(ExportFormatType.WordForWindows, NomeArquivo)
                Else
                    rpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, NomeArquivo)
                End If

                If System.IO.File.Exists(NomeArquivo) Then
                    If tipo = eExportType.PDF Then
                        AbrirArquivo(page, fileName)
                    Else
                        AbrirExcel(page, fileName)
                    End If
                End If
            Catch ex As Exception
                Dim msg As String = ex.Message
                If ex.InnerException IsNot Nothing Then
                    msg &= " - " & ex.InnerException.Message
                End If
                Throw New Exception(msg)
            Finally
                rpt.Close()
                rpt.Dispose()
            End Try
        End If
    End Sub

    Public Shared Sub BindParameters(ByRef rpt As ReportDocument, ByVal parameters As Dictionary(Of String, Object))
        Dim parameterValues As New ParameterValues
        Dim parameterDiscreteValue As New ParameterDiscreteValue
        Dim parameterFieldDefinitions As ParameterFieldDefinitions
        Dim parameterFieldDefinition As ParameterFieldDefinition

        parameterFieldDefinitions = rpt.DataDefinition.ParameterFields()

        Try
            For Each p In parameters
                If parameterFieldDefinitions.Count > 0 Then
                    parameterFieldDefinition = parameterFieldDefinitions.Item(p.Key)
                    parameterValues = parameterFieldDefinition.CurrentValues
                    parameterDiscreteValue.Value = p.Value.ToString
                    parameterValues.Add(parameterDiscreteValue)
                    parameterFieldDefinition.ApplyCurrentValues(parameterValues)
                End If
            Next
        Catch ex As Exception
            Dim teste = "ERRO EM PARAMETROS"
        End Try

    End Sub

    Public Function OnlyChars(ByVal str As String) As String
        Return OnlyChars(str, "abcdefghijklmnopqrstuvxwyzABCDEFGHIJKLMNOPQRSTUVXWYZ")
    End Function

    Public Shared Function OnlyNumbers(ByVal str As String) As String
        Return OnlyChars(str, "0123456789")
    End Function

    Public Shared Function OnlyChars(ByVal str As String, ByVal charListToKeep As String) As String
        Dim newStr = New StringBuilder()
        str.ToCharArray().Where(Function(c) charListToKeep.Any(Function(cc) cc = c)).ToList().ForEach(Function(c) newStr.Append(c))
        Return newStr.ToString()
    End Function

    Public Shared Sub BindRomaneio(ByVal page As Page, ByVal Empresa As String, ByVal Romaneio As String, ByVal Produto As String)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim sql As String = String.Empty
        Dim myDataRow As DataRow

        sql = "SELECT  0 AS Laudo, r.Produto, isnull(nXt.Placa,'') AS Placa, r.EntradaSaida,  " & vbCrLf &
                "           isnull(r.PrimeiraPesagem,0) AS PrimeiraPesagem, isnull(r.SegundaPesagem,0) AS SegundaPesagem, r.PesoBruto AS BrutoBalanca,  " & vbCrLf &
                "           r.Desconto AS Descontos, r.PesoLiquido AS Liquido, CURRENT_TIMESTAMP AS EntradaPatio,  " & vbCrLf &
                "           CURRENT_TIMESTAMP AS EntradaBalanca, CURRENT_TIMESTAMP AS  SaidaBalanca, r.Movimento, nXr.Nota_Id AS NotaFiscal, " & vbCrLf &
                "           nXr.Serie_Id AS SerieNota, nXi.QuantidadeFiscal as PesoFiscal, '' AS Observacoes, prd.Nome AS NomeProduto,  " & vbCrLf &
                "           ClientePedido.Cliente_Id AS CodigoCliente, ClientePedido.Endereco_Id AS EndCliente, ClientePedido.Nome AS NomeCliente, " & vbCrLf &
                "           ClientePedido.Reduzido AS ReduzidoCliente, ClientePedido.Endereco AS EnderecoCliente, ClientePedido.Cidade AS CidadeCliente,  " & vbCrLf &
                "           ClientePedido.Estado AS EstadoCliente, isnull(Transportes.Cliente_Id,'') AS CodigoTransportador, isnull(Transportes.Endereco_Id,'') AS EndTransportador,  " & vbCrLf &
                "           isnull(Transportes.Nome,'') AS NomeTransportador, isnull(Transportes.Reduzido,'') AS ReduzidoTransportador, isnull(Transportes.Endereco,'') AS EnderecoTransportador,  " & vbCrLf &
                "           isnull(Transportes.Cidade,'') AS CidadeTransportador, isnull(Transportes.Estado,'') AS EstadoTransportador, Depositos.Cliente_Id AS CodigoDeposito,  " & vbCrLf &
                "           Depositos.Endereco_Id AS EndDeposito, Depositos.Nome AS NomeDeposito, Depositos.Reduzido AS ReduzidoDeposito, Depositos.Endereco AS EnderecoDeposito, " & vbCrLf &
                "           Depositos.Cidade AS CidadeDeposito, Depositos.Estado AS EstadoDeposito, Depositos.Inscricao AS InscricaoDeposito, isnull(Placas.Placa01,'') AS Placa01,  " & vbCrLf &
                "           isnull(Placas.Placa02,'') AS Placa02, isnull(Placas.Placa03,'') AS Placa03, isnull(Placas.CidadePlaca,'') AS CidadePlaca, isnull(Placas.EstadoPlaca,'') AS EstadoPlaca, isnull(Motorista.Nome,'') AS NomeMotorista,  " & vbCrLf &
                "           isnull(Motorista.Cidade,'') AS CidadeMotorista, isnull(Motorista.Estado,'') AS EstadoMotorista, isnull(Placas.Habilitacao,'') AS Habilitacao, CASE WHEN Placas.CpfMotorista IS NULL  " & vbCrLf &
                "           THEN '' WHEN Placas.CpfMotorista = '' THEN '' ELSE SUBSTRING(Placas.CpfMotorista, 1, 3) + '.' + SUBSTRING(Placas.CpfMotorista, 4, 3)  " & vbCrLf &
                "           + '.' + SUBSTRING(Placas.CpfMotorista, 7, 3) + '-' + SUBSTRING(Placas.CpfMotorista, 10, 2) END AS CpfMotorista, isnull(EstadoPlaca.Descricao,'') AS NomeEstadoPlaca, " & vbCrLf &
                "           isnull(EstadoMotorista.Descricao,'') AS NomeEstadoMotorista, r.Pedido, Depositos.Numero AS NumeroDeposito, Depositos.Complemento AS ComplementoDeposito,  " & vbCrLf &
                "           Depositos.Bairro AS BairroDeposito, Clientes.Numero AS NumeroCliente, Clientes.Complemento AS ComplementoCliente, Clientes.Bairro AS BairroCliente,  " & vbCrLf &
                "           Clientes.Inscricao AS InscricaoCliente, r.Empresa_Id as CodigoEmpresa, r.EndEmpresa_id as EndEmpresa, Empresa.Nome as NomeEmpresa,  " & vbCrLf &
                "           Empresa.Reduzido as ReduzidoEmpresa, Empresa.Endereco as EnderecoEmpresa, Empresa.Cidade as CidadeEmpresa, Empresa.Estado as EstadoEmpresa,  " & vbCrLf &
                "           Empresa.Inscricao as InscricaoEmpresa, Empresa.Numero as NumeroEmpresa, Empresa.Complemento as ComplementoEmpresa, Empresa.Bairro as BairroEmpresa, " & vbCrLf &
                "           r.Romaneio_Id AS Romaneio " & vbCrLf &
                "        FROM  Romaneios r " & vbCrLf &
                "              INNER JOIN NotasFiscaisXRomaneios nXr " & vbCrLf &
                "					   ON nXr.Empresa_Id    = r.Empresa_Id " & vbCrLf &
                "                     AND nXr.EndEmpresa_Id = r.EndEmpresa_Id  " & vbCrLf &
                "                     AND nXr.Romaneio_Id   = r.Romaneio_Id  " & vbCrLf &
                "              INNER JOIN NotasFiscaisXItens nXi " & vbCrLf &
                "					   ON nXi.Empresa_Id      = nXr.Empresa_Id " & vbCrLf &
                "                     AND nXi.EndEmpresa_Id   = nXr.EndEmpresa_Id  " & vbCrLf &
                "                     AND nXi.Cliente_Id      = nXr.Cliente_Id  " & vbCrLf &
                "                     AND nXi.EndCliente_Id   = nXr.EndCliente_Id  " & vbCrLf &
                "                     AND nXi.EntradaSaida_Id = nXr.EntradaSaida_Id  " & vbCrLf &
                "                     AND nXi.Serie_Id        = nXr.Serie_Id " & vbCrLf &
                "                     AND nXi.Nota_Id         = nXr.Nota_Id " & vbCrLf &
                "              INNER JOIN Clientes as Empresa  " & vbCrLf &
                "					   ON Empresa.Cliente_Id  = nXr.Empresa_Id  " & vbCrLf &
                "                     And Empresa.Endereco_Id = nXr.EndEmpresa_Id " & vbCrLf &
                "              LEFT JOIN NotasFiscaisXTransportadores nXt " & vbCrLf &
                "					   ON nXt.Empresa_Id    = nXr.Empresa_Id " & vbCrLf &
                "                     AND nXt.EndEmpresa_Id = nXr.EndEmpresa_Id  " & vbCrLf &
                "                     AND nXt.Cliente_Id    = nXr.Cliente_Id  " & vbCrLf &
                "                     AND nXt.EndCliente_Id = nXr.EndCliente_Id  " & vbCrLf &
                "                     AND nXt.EntradaSaida_Id = nXr.EntradaSaida_Id  " & vbCrLf &
                "                     AND nXt.Serie_Id = nXr.Serie_Id " & vbCrLf &
                "                     AND nXt.Nota_Id = nXr.Nota_Id  " & vbCrLf &
                "              INNER JOIN Pedidos p " & vbCrLf &
                "                      ON p.Empresa_Id    = r.Empresa_Id  " & vbCrLf &
                "                     AND p.EndEmpresa_Id = r.EndEmpresa_Id  " & vbCrLf &
                "                     AND p.Pedido_Id     = r.Pedido " & vbCrLf &
                "              INNER JOIN Clientes AS ClientePedido  " & vbCrLf &
                "                      ON p.Cliente    = ClientePedido.Cliente_Id  " & vbCrLf &
                "                     AND p.EndCliente = ClientePedido.Endereco_Id  " & vbCrLf &
                "              INNER JOIN Clientes  " & vbCrLf &
                "                      ON  Clientes.Cliente_Id  = nXr.Cliente_Id " & vbCrLf &
                "                     AND  Clientes.Endereco_Id = nXr.EndCliente_Id " & vbCrLf &
                "              INNER JOIN Clientes AS Depositos  " & vbCrLf &
                "                      ON Depositos.Cliente_Id  = r.Deposito " & vbCrLf &
                "                     AND Depositos.Endereco_Id = r.EndDeposito " & vbCrLf &
                "              INNER JOIN Produtos prd " & vbCrLf &
                "                  ON prd.Produto_Id = r.Produto  " & vbCrLf &
                "              LEFT JOIN Placas  " & vbCrLf &
                "                      ON Placas.Placa_Id = nxt.Placa  " & vbCrLf &
                "              LEFT JOIN Clientes AS Transportes  " & vbCrLf &
                "                      ON Transportes.Cliente_Id  = nXt.Proprietario " & vbCrLf &
                "                     AND Transportes.Endereco_Id = nXt.EndProprietario  " & vbCrLf &
                "              LEFT JOIN Clientes AS Motorista  " & vbCrLf &
                "                      ON Motorista.Cliente_Id  = nXt.Motorista " & vbCrLf &
                "                     AND Motorista.Endereco_Id = nXt.EndMotorista " & vbCrLf &
                "              LEFT JOIN Estados AS EstadoPlaca  " & vbCrLf &
                "                      ON EstadoPlaca.Estado_Id = Placas.EstadoPlaca " & vbCrLf &
                "              LEFT JOIN Estados AS EstadoMotorista " & vbCrLf &
                "                      ON EstadoMotorista.Estado_Id = Motorista.Estado " & vbCrLf &
                "        WHERE (r.Empresa_Id = '" & Empresa & "')  " & vbCrLf &
                "          AND (r.EndEmpresa_Id = 0)  " & vbCrLf &
                "          AND (r.Romaneio_Id = " & Romaneio & ")"

        ds = Banco.ConsultaDataSet(sql, "Laudo")

        sql = "SELECT RomaneiosXDescontos.Analise_Id AS Analise, " & vbCrLf &
              "	   CASE " & vbCrLf &
              "		  WHEN RomaneiosXDescontos.Analise_Id = 6 " & vbCrLf &
              "				THEN Analises.Descricao + ' ' + " & vbCrLf &
              "					  (SELECT     Observacoes " & vbCrLf &
              "						 FROM    Classificacoes " & vbCrLf &
              "						WHERE       (Analise_Id = 6) AND (Sequencia_Id = RomaneiosXDescontos.Indice) " & vbCrLf &
              "						  AND (Produto_Id = '" & Produto & "')) " & vbCrLf &
              "				ELSE " & vbCrLf &
              "				   CASE " & vbCrLf &
              "				      WHEN RomaneiosXDescontos.Analise_Id = 12 " & vbCrLf &
              "				         THEN " & vbCrLf &
              "				            CASE " & vbCrLf &
              "				               WHEN RomaneiosXDescontos.Percentual = 2 " & vbCrLf &
              "				                  THEN Analises.Descricao + ' - TESTE POSITIVO' " & vbCrLf &
              "				                  ELSE " & vbCrLf &
              "				                     CASE " & vbCrLf &
              "				                        WHEN RomaneiosXDescontos.Percentual = 1 " & vbCrLf &
              "				                           THEN Analises.Descricao + ' - SIM' " & vbCrLf &
              "				                           ELSE Analises.Descricao + ' - NÃO' " & vbCrLf &
              "				                        END " & vbCrLf &
              "				            END " & vbCrLf &
              "				         ELSE Analises.Descricao " & vbCrLf &
              "				      END " & vbCrLf &
              "		  END AS Descricao, " & vbCrLf &
              "       RomaneiosXDescontos.Percentual, RomaneiosXDescontos.Indice, RomaneiosXDescontos.Desconto " & vbCrLf &
              "FROM RomaneiosXDescontos " & vbCrLf &
              "	 INNER JOIN Analises " & vbCrLf &
              "			 ON RomaneiosXDescontos.Analise_Id = Analises.Analise_Id " & vbCrLf &
              "WHERE (RomaneiosXDescontos.Empresa_Id    = '" & Empresa & "') " & vbCrLf &
              "AND   (RomaneiosXDescontos.EndEmpresa_Id = 0) " & vbCrLf &
              "AND   (RomaneiosXDescontos.Romaneio_Id   = " & Romaneio & ") " & vbCrLf &
              "ORDER BY Analise"

        ds.Merge(Banco.ConsultaDataSet(sql, "Analises"))

        myDataRow = ds.Tables(0).NewRow
        myDataRow("Laudo") = ds.Tables(0).Rows(0).Item("Laudo")
        myDataRow("Produto") = ds.Tables(0).Rows(0).Item("Produto")
        myDataRow("Placa") = ds.Tables(0).Rows(0).Item("Placa")
        myDataRow("EntradaSaida") = ds.Tables(0).Rows(0).Item("EntradaSaida")
        myDataRow("PrimeiraPesagem") = ds.Tables(0).Rows(0).Item("PrimeiraPesagem")
        myDataRow("SegundaPesagem") = ds.Tables(0).Rows(0).Item("SegundaPesagem")
        myDataRow("BrutoBalanca") = ds.Tables(0).Rows(0).Item("BrutoBalanca")
        myDataRow("Descontos") = ds.Tables(0).Rows(0).Item("Descontos")
        myDataRow("Liquido") = ds.Tables(0).Rows(0).Item("Liquido")
        myDataRow("EntradaPatio") = ds.Tables(0).Rows(0).Item("EntradaPatio")
        myDataRow("EntradaBalanca") = ds.Tables(0).Rows(0).Item("EntradaBalanca")
        myDataRow("SaidaBalanca") = ds.Tables(0).Rows(0).Item("SaidaBalanca")
        myDataRow("Movimento") = ds.Tables(0).Rows(0).Item("Movimento")
        myDataRow("NotaFiscal") = ds.Tables(0).Rows(0).Item("NotaFiscal")
        myDataRow("SerieNota") = ds.Tables(0).Rows(0).Item("SerieNota")
        myDataRow("PesoFiscal") = ds.Tables(0).Rows(0).Item("PesoFiscal")
        myDataRow("Observacoes") = ds.Tables(0).Rows(0).Item("Observacoes")
        myDataRow("NomeProduto") = ds.Tables(0).Rows(0).Item("NomeProduto")
        myDataRow("CodigoCliente") = ds.Tables(0).Rows(0).Item("CodigoCliente")
        myDataRow("EndCliente") = ds.Tables(0).Rows(0).Item("EndCliente")
        myDataRow("NomeCliente") = ds.Tables(0).Rows(0).Item("NomeCliente")
        myDataRow("ReduzidoCliente") = ds.Tables(0).Rows(0).Item("ReduzidoCliente")
        myDataRow("EnderecoCliente") = ds.Tables(0).Rows(0).Item("EnderecoCliente")
        myDataRow("CidadeCliente") = ds.Tables(0).Rows(0).Item("CidadeCliente")
        myDataRow("EstadoCliente") = ds.Tables(0).Rows(0).Item("EstadoCliente")
        myDataRow("CodigoTransportador") = ds.Tables(0).Rows(0).Item("CodigoTransportador")
        myDataRow("EndTransportador") = ds.Tables(0).Rows(0).Item("EndTransportador")
        myDataRow("NomeTransportador") = ds.Tables(0).Rows(0).Item("NomeTransportador")
        myDataRow("ReduzidoTransportador") = ds.Tables(0).Rows(0).Item("ReduzidoTransportador")
        myDataRow("EnderecoTransportador") = ds.Tables(0).Rows(0).Item("EnderecoTransportador")
        myDataRow("CidadeTransportador") = ds.Tables(0).Rows(0).Item("CidadeTransportador")
        myDataRow("EstadoTransportador") = ds.Tables(0).Rows(0).Item("EstadoTransportador")
        myDataRow("CodigoDeposito") = ds.Tables(0).Rows(0).Item("CodigoDeposito")
        myDataRow("EndDeposito") = ds.Tables(0).Rows(0).Item("EndDeposito")
        myDataRow("NomeDeposito") = ds.Tables(0).Rows(0).Item("NomeDeposito")
        myDataRow("ReduzidoDeposito") = ds.Tables(0).Rows(0).Item("ReduzidoDeposito")
        myDataRow("EnderecoDeposito") = ds.Tables(0).Rows(0).Item("EnderecoDeposito")
        myDataRow("CidadeDeposito") = ds.Tables(0).Rows(0).Item("CidadeDeposito")
        myDataRow("EstadoDeposito") = ds.Tables(0).Rows(0).Item("EstadoDeposito")
        myDataRow("InscricaoDeposito") = ds.Tables(0).Rows(0).Item("InscricaoDeposito")

        myDataRow("CodigoEmpresa") = ds.Tables(0).Rows(0).Item("CodigoEmpresa")
        myDataRow("EndEmpresa") = ds.Tables(0).Rows(0).Item("EndEmpresa")
        myDataRow("NomeEmpresa") = ds.Tables(0).Rows(0).Item("NomeEmpresa")
        myDataRow("ReduzidoEmpresa") = ds.Tables(0).Rows(0).Item("ReduzidoEmpresa")
        myDataRow("EnderecoEmpresa") = ds.Tables(0).Rows(0).Item("EnderecoEmpresa")
        myDataRow("CidadeEmpresa") = ds.Tables(0).Rows(0).Item("CidadeEmpresa")
        myDataRow("EstadoEmpresa") = ds.Tables(0).Rows(0).Item("EstadoEmpresa")
        myDataRow("InscricaoEmpresa") = ds.Tables(0).Rows(0).Item("InscricaoEmpresa")
        myDataRow("NumeroEmpresa") = ds.Tables(0).Rows(0).Item("NumeroEmpresa")
        myDataRow("ComplementoEmpresa") = ds.Tables(0).Rows(0).Item("ComplementoEmpresa")
        myDataRow("BairroEmpresa") = ds.Tables(0).Rows(0).Item("BairroEmpresa")

        myDataRow("Placa01") = ds.Tables(0).Rows(0).Item("Placa01")
        myDataRow("Placa02") = ds.Tables(0).Rows(0).Item("Placa02")
        myDataRow("Placa03") = ds.Tables(0).Rows(0).Item("Placa03")
        myDataRow("CidadePlaca") = ds.Tables(0).Rows(0).Item("CidadePlaca")
        myDataRow("EstadoPlaca") = ds.Tables(0).Rows(0).Item("EstadoPlaca")
        myDataRow("NomeMotorista") = ds.Tables(0).Rows(0).Item("NomeMotorista")
        myDataRow("CidadeMotorista") = ds.Tables(0).Rows(0).Item("CidadeMotorista")
        myDataRow("EstadoMotorista") = ds.Tables(0).Rows(0).Item("EstadoMotorista")
        myDataRow("Habilitacao") = ds.Tables(0).Rows(0).Item("Habilitacao")
        myDataRow("CpfMotorista") = ds.Tables(0).Rows(0).Item("CpfMotorista")
        myDataRow("NomeEstadoPlaca") = ds.Tables(0).Rows(0).Item("NomeEstadoPlaca")
        myDataRow("NomeEstadoMotorista") = ds.Tables(0).Rows(0).Item("NomeEstadoMotorista")
        myDataRow("Pedido") = ds.Tables(0).Rows(0).Item("Pedido")
        myDataRow("NumeroDeposito") = ds.Tables(0).Rows(0).Item("NumeroDeposito")
        myDataRow("ComplementoDeposito") = ds.Tables(0).Rows(0).Item("ComplementoDeposito")
        myDataRow("BairroDeposito") = ds.Tables(0).Rows(0).Item("BairroDeposito")
        myDataRow("NumeroCliente") = ds.Tables(0).Rows(0).Item("NumeroCliente")
        myDataRow("ComplementoCliente") = ds.Tables(0).Rows(0).Item("ComplementoCliente")
        myDataRow("BairroCliente") = ds.Tables(0).Rows(0).Item("BairroCliente")
        myDataRow("InscricaoCliente") = ds.Tables(0).Rows(0).Item("InscricaoCliente")

        myDataRow("Romaneio") = ds.Tables(0).Rows(0).Item("Romaneio")

        ds.Tables(0).Rows.Add(myDataRow)

        Dim param As New Dictionary(Of String, Object)
        param.Add("Reemissao", "")

        BindReport(page, ds, "Cr_LaudoDePesagem", eExportType.PDF, param)
    End Sub

#Region "Envio de E-mail"

    Public Shared Function SendMail(from As [String], friendlyName As [String], [lstTo] As List(Of [String]), subject As [String], bodyHMTL As [String], smtpNetworkSettings As SmtpNetworkElement,
                                    ByRef errorMessage As [String], Optional file As String = "", Optional stream As MemoryStream = Nothing, Optional fileName As [String] = "") As [Boolean]
        Dim aux = True
        Try
            Dim objMail = New MailMessage()
            For Each toMail As String In lstTo
                objMail.[To].Add(toMail)
            Next

            objMail.Subject = subject

            If Not [String].IsNullOrEmpty(file) Then
                objMail.Attachments.Add(New Attachment(file))
            End If

            If stream IsNot Nothing AndAlso Not [String].IsNullOrEmpty(fileName) Then
                objMail.Attachments.Add(New Attachment(stream, fileName))
            End If

            objMail.IsBodyHtml = True
            objMail.Body = bodyHMTL

            Dim configurationFile = WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Request.Url.AbsolutePath)
            Dim mailSettings = TryCast(configurationFile.GetSectionGroup("system.net/mailSettings"), MailSettingsSectionGroup)

            Dim port = 587
            Dim host = [String].Empty
            Dim password = [String].Empty

            Dim userName = [String].Empty
            Dim enableSsl = False
            Dim defaultCredentials = False

            If smtpNetworkSettings IsNot Nothing AndAlso Not [String].IsNullOrEmpty(smtpNetworkSettings.Host) Then
                port = smtpNetworkSettings.Port
                host = smtpNetworkSettings.Host
                password = smtpNetworkSettings.Password
                userName = smtpNetworkSettings.UserName
                enableSsl = smtpNetworkSettings.EnableSsl
                defaultCredentials = mailSettings.Smtp.Network.DefaultCredentials
            ElseIf mailSettings IsNot Nothing Then
                port = mailSettings.Smtp.Network.Port
                host = mailSettings.Smtp.Network.Host
                password = mailSettings.Smtp.Network.Password
                userName = mailSettings.Smtp.Network.UserName
                enableSsl = mailSettings.Smtp.Network.EnableSsl
                defaultCredentials = mailSettings.Smtp.Network.DefaultCredentials
            End If

            If [String].IsNullOrEmpty(from) Then
                If mailSettings IsNot Nothing Then
                    objMail.From = New MailAddress(mailSettings.Smtp.From)
                End If
            Else
                objMail.From = If(Not [String].IsNullOrEmpty(friendlyName), New MailAddress(from, friendlyName), New MailAddress(from))
            End If

            Dim smtp = New SmtpClient() With {
             .Port = port,
             .Host = host,
             .EnableSsl = enableSsl,
             .UseDefaultCredentials = defaultCredentials,
             .Credentials = New NetworkCredential(userName, password)
            }

            'NÀO REMOVER ESSA LINHA COMENTADA - FURLAN 19/05/2025
            'ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
            ServicePointManager.SecurityProtocol = CType(&HC00, SecurityProtocolType)
            smtp.Send(objMail)
            errorMessage = "E-mail enviado com sucesso!"
        Catch ex As SmtpException
            aux = False
            errorMessage = ex.Message
        End Try
        Return aux
    End Function

    Public Shared Function GetFromMail() As [String]
        Dim config As Configuration = WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Request.ApplicationPath)
        If config IsNot Nothing Then
            Dim settings As MailSettingsSectionGroup = DirectCast(config.GetSectionGroup("system.net/mailSettings"), MailSettingsSectionGroup)
            If settings IsNot Nothing Then
                Return settings.Smtp.From
            End If
        End If
        Return String.Empty
    End Function

    Public Shared Function GetSmtpSettings() As SmtpNetworkElement
        Dim config As Configuration = WebConfigurationManager.OpenWebConfiguration(HttpContext.Current.Request.ApplicationPath)
        If config IsNot Nothing Then
            Dim settings As MailSettingsSectionGroup = DirectCast(config.GetSectionGroup("system.net/mailSettings"), MailSettingsSectionGroup)
            If settings IsNot Nothing Then
                Return New SmtpNetworkElement() With {
                 .Host = settings.Smtp.Network.Host,
                 .Port = settings.Smtp.Network.Port,
                 .Password = settings.Smtp.Network.Password,
                 .UserName = settings.Smtp.Network.UserName,
                 .EnableSsl = settings.Smtp.Network.EnableSsl,
                 .DefaultCredentials = settings.Smtp.Network.DefaultCredentials
                }
            End If
        End If
        Return Nothing
    End Function
#End Region

    Public Shared Sub Ajuda(ByRef page As System.Web.UI.Control, ByVal processo As String)
        Dim objProcesso As New Processo(eServidor.ServidorLocal, processo)
        'Verifica Conexão com a internet para atualização local.
        If Funcoes.VerificaConexaoInternet() Then
            Dim objProcessoUOL As New Processo(eServidor.ServidorUOL, processo)

            'encontrou no servidor ngs.
            If objProcessoUOL IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objProcessoUOL.Processo) Then
                'não encontrou no servidor local.
                If objProcesso Is Nothing OrElse String.IsNullOrWhiteSpace(objProcesso.Processo) Then
                    objProcesso = objProcessoUOL
                    objProcesso.IUD = "I"
                    objProcesso.Salvar(eServidor.ServidorLocal)
                    'encontrou no servidor local, porém desatualizado.
                ElseIf objProcesso.DataAtualizacao < objProcessoUOL.DataAtualizacao Then
                    objProcesso = objProcessoUOL
                    objProcesso.IUD = "U"
                    objProcesso.Salvar(eServidor.ServidorLocal)
                End If
            End If
        End If

        If objProcesso IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objProcesso.Manual) Then
            ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, "$(document).ready(function () { $('#mceu_44').css('visibility','visible'); $('#mceu_44-body').html('" & HttpUtility.HtmlDecode(objProcesso.Manual.ToString.Replace(Environment.NewLine, "")) & "'); });", True)
        Else
            Throw New Exception("Não foi possivel exibir a ajuda do processo: " & processo & ". OBS: em elaboração.")
        End If
    End Sub

    Public Shared Sub AbrirAccordion(ByRef Page As System.Web.UI.Control, ByVal AccordionId As String, ByVal AbrirAccordion As Boolean, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & AccordionId & "').accordion('activate',  " & IIf(AbrirAccordion, 0, 1) & ");"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Shared Sub WriteLogFile(fileName As String, methodName As String, message As String)
        Try
            If Not [String].IsNullOrEmpty(message) Then
                Dim sw As New StreamWriter(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log.txt"), True)
                sw.WriteLine("--------------------------------------------------------------------------------------------------------------------------------------------")
                sw.WriteLine(Convert.ToString(((Convert.ToString(((Convert.ToString((System.DateTime.Now + " - ")) & fileName) + " - ")) & methodName) + " - ")) & message)
                sw.Close()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Shared Function AtribuirirValorFormatado(ByVal Campo As String) As String
        Dim Valor As String = 0

        If String.IsNullOrWhiteSpace(Campo) Then
            Valor = FormatarValor(0, 18, 2)
        Else
            Valor = FormatarValor(CDec(Campo), 18, 2)
        End If

        Return Valor

    End Function

    'Public Shared Function Cifrar(ByVal vstrTextToBeEncrypted As String, ByVal vstrEncryptionKey As String) As String
    'Public Shared Function Cifrar(ByVal vstrTextToBeEncrypted As String) As String

    '    Dim vstrEncryptionKey As String = ChaveSecreta

    '    Dim bytValue() As Byte
    '    Dim bytKey() As Byte
    '    Dim bytEncoded() As Byte = {0, 0}
    '    Dim bytIV() As Byte = {121, 241, 10, 1, 132, 74, 11, 39, 255, 91, 45, 78, 14, 211, 22, 62}
    '    Dim intLength As Integer
    '    Dim intRemaining As Integer
    '    Dim objMemoryStream As New MemoryStream
    '    Dim objCryptoStream As CryptoStream
    '    Dim objRijndaelManaged As RijndaelManaged


    '    '   **********************************************************************
    '    '   ****** Descarta todos os caracteres nulos da palavra a ser cifrada****  
    '    '   **********************************************************************

    '    vstrTextToBeEncrypted = TiraCaracteresNulos(vstrTextToBeEncrypted)

    '    '   **********************************************************************
    '    '   ******  O valor deve estar dentro da tabela ASCII (i.e., no DBCS chars)
    '    '   **********************************************************************

    '    bytValue = Encoding.ASCII.GetBytes(vstrTextToBeEncrypted.ToCharArray)

    '    intLength = Len(vstrEncryptionKey)

    '    '   ********************************************************************
    '    '   ******   A chave cifrada será de 256 bits long (32 bytes)     ******
    '    '   ******   Se for maior que 32 bytes então será truncado.       ******
    '    '   ******   Se for menor que 32 bytes será alocado.              ******
    '    '   ******   Usando upper-case Xs.                                ****** 
    '    '   ********************************************************************

    '    If intLength >= 32 Then
    '        vstrEncryptionKey = Strings.Left(vstrEncryptionKey, 32)
    '    Else
    '        intLength = Len(vstrEncryptionKey)
    '        intRemaining = 32 - intLength
    '        vstrEncryptionKey = vstrEncryptionKey & Strings.StrDup(intRemaining, "X")
    '    End If

    '    bytKey = Encoding.ASCII.GetBytes(vstrEncryptionKey.ToCharArray)

    '    objRijndaelManaged = New RijndaelManaged

    '    '   ***********************************************************************
    '    '   ******  Cria o valor a ser crifrado e depois escreve             ******
    '    '   ******  Convertido em uma disposição do byte                     ******
    '    '   ***********************************************************************

    '    Try

    '        objCryptoStream = New CryptoStream(objMemoryStream, objRijndaelManaged.CreateEncryptor(bytKey, bytIV), CryptoStreamMode.Write)
    '        objCryptoStream.Write(bytValue, 0, bytValue.Length)

    '        objCryptoStream.FlushFinalBlock()

    '        bytEncoded = objMemoryStream.ToArray
    '        objMemoryStream.Close()
    '        objCryptoStream.Close()
    '    Catch

    '    End Try

    '    '   ***********************************************************************
    '    '   ****** Retorna o valor cifrado (convertido de byte para base64 )  *****
    '    '   ***********************************************************************

    '    Return HttpUtility.UrlEncode(Convert.ToBase64String(bytEncoded))

    'End Function

    ''   *********************************************************************
    ''   ***** Função Responsável por Decifrar a sua String Cifrada       *****
    ''   ***** Use da seguinte forma:                                    *****
    ''   ***** Call Decifrar("Palavra", "SuaChaveSecreta(Ex.2345)")      *****
    ''   *********************************************************************

    ''Public Shared Function Decifrar(ByVal vstrStringToBeDecrypted As String, ByVal vstrDecryptionKey As String) As String
    'Public Shared Function Decifrar(ByVal vstrStringToBeDecrypted As String) As String

    '    Dim vstrDecryptionKey As String = ChaveSecreta

    '    Dim bytDataToBeDecrypted() As Byte
    '    Dim bytTemp() As Byte
    '    Dim bytIV() As Byte = {121, 241, 10, 1, 132, 74, 11, 39, 255, 91, 45, 78, 14, 211, 22, 62}
    '    Dim objRijndaelManaged As New RijndaelManaged
    '    Dim objMemoryStream As MemoryStream
    '    Dim objCryptoStream As CryptoStream
    '    Dim bytDecryptionKey() As Byte

    '    Dim intLength As Integer
    '    Dim intRemaining As Integer
    '    Dim strReturnString As String = String.Empty

    '    Try



    '    '   *****************************************************************
    '    '   ******   Convert base64 cifrada para byte array            ******
    '    '   ******   Convert base64 cifrada para byte array            ******
    '    '   *****************************************************************

    '    'Dim strTemp = Convert.ToBase64String(vstrStringToBeDecrypted)

    '    Dim strTemp = HttpUtility.UrlDecode(vstrStringToBeDecrypted)

    '    bytDataToBeDecrypted = Convert.FromBase64String(strTemp)

    '    '   ********************************************************************
    '    '   ******   A chave cifrada sera de 256 bits long (32 bytes)     ******
    '    '   ******   Se for maior que 32 bytes então será truncado.       ******
    '    '   ******   Se for menor que 32 bytes será alocado.              ******
    '    '   ******   Usando upper-case Xs.                                ****** 
    '    '   ********************************************************************

    '    intLength = Len(vstrDecryptionKey)

    '    If intLength >= 32 Then
    '        vstrDecryptionKey = Strings.Left(vstrDecryptionKey, 32)
    '    Else
    '        intLength = Len(vstrDecryptionKey)
    '        intRemaining = 32 - intLength
    '        vstrDecryptionKey = vstrDecryptionKey & Strings.StrDup(intRemaining, "X")
    '    End If

    '    bytDecryptionKey = Encoding.ASCII.GetBytes(vstrDecryptionKey.ToCharArray)

    '    ReDim bytTemp(bytDataToBeDecrypted.Length)

    '    objMemoryStream = New MemoryStream(bytDataToBeDecrypted)

    '    '   ***********************************************************************
    '    '   ******  Escrever o valor decifrado depois que é convertido       ******
    '    '   ***********************************************************************

    '    Catch ex As Exception

    '    End Try


    '    Try

    '        objCryptoStream = New CryptoStream(objMemoryStream, _
    '           objRijndaelManaged.CreateDecryptor(bytDecryptionKey, bytIV), _
    '           CryptoStreamMode.Read)

    '        objCryptoStream.Read(bytTemp, 0, bytTemp.Length)

    '        objCryptoStream.FlushFinalBlock()
    '        objMemoryStream.Close()
    '        objCryptoStream.Close()

    '    Catch

    '    End Try

    '    '   ***********************************************************************
    '    '   ******  Retorna o valor decifrado                                ******
    '    '   ***********************************************************************

    '    Return TiraCaracteresNulos(Encoding.ASCII.GetString(bytTemp))

    'End Function

    '   *********************************************************************
    '   ***** Função responvel por tirar os espaços em branco da        *****
    '   ***** variável a ser cifrada                                    *****
    '   ***** Esta função é chamada internamente                        *****
    '   *********************************************************************


    Public Shared Function Cifrar(clearText As String) As String
        Try


            Dim EncryptionKey As String = ChaveSecreta  '"MAKV2SPBNI99212"
            Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
            Using encryptor As Aes = Aes.Create()
                Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
                 &H65, &H64, &H76, &H65, &H64, &H65,
                 &H76})
                encryptor.Key = pdb.GetBytes(32)
                encryptor.IV = pdb.GetBytes(16)
                Using ms As New MemoryStream()
                    Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                        cs.Write(clearBytes, 0, clearBytes.Length)
                        cs.Close()
                    End Using
                    clearText = Convert.ToBase64String(ms.ToArray())
                End Using
            End Using


        Catch ex As Exception
            Return HttpUtility.UrlDecode(clearText)
        End Try

        Return HttpUtility.UrlDecode(clearText)
    End Function

    Public Shared Function Decifrar(cipherText As String) As String
        Try
            Dim EncryptionKey As String = ChaveSecreta '"MAKV2SPBNI99212"
            cipherText = cipherText.Replace(" ", "+")
            Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
            Using encryptor As Aes = Aes.Create()
                Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
                 &H65, &H64, &H76, &H65, &H64, &H65,
                 &H76})
                encryptor.Key = pdb.GetBytes(32)
                encryptor.IV = pdb.GetBytes(16)
                Using ms As New MemoryStream()
                    Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                        cs.Write(cipherBytes, 0, cipherBytes.Length)
                        cs.Close()
                    End Using
                    cipherText = Encoding.Unicode.GetString(ms.ToArray())
                End Using
            End Using
        Catch ex As Exception
            Return HttpUtility.UrlEncode(cipherText)
        End Try
        Return HttpUtility.UrlEncode(cipherText)
    End Function

    Public Shared Function TiraCaracteresNulos(ByVal vstrStringWithNulls As String) As String

        Dim intPosition As Integer
        Dim strStringWithOutNulls As String

        intPosition = 1
        strStringWithOutNulls = vstrStringWithNulls

        Do While intPosition > 0
            intPosition = InStr(intPosition, vstrStringWithNulls, vbNullChar)

            If intPosition > 0 Then
                strStringWithOutNulls = Left$(strStringWithOutNulls, intPosition - 1) &
                                  Right$(strStringWithOutNulls, Len(strStringWithOutNulls) - intPosition)
            End If

            If intPosition > strStringWithOutNulls.Length Then
                Exit Do
            End If
        Loop


        Return strStringWithOutNulls

    End Function

    Public Shared Function LimparErroLeituraXML(ByVal caminhoXml As String, ByVal DsXml As DataSet) As Boolean

        Try

            ' Ler como texto
            Dim xmlOriginal As String = File.ReadAllText(caminhoXml)

            ' Remover namespace
            xmlOriginal = System.Text.RegularExpressions.Regex.Replace(xmlOriginal, "xmlns=""[^""]*""", "")

            ' Regex para remover atributos Id="..." e versao="..."
            Dim pattern As String = "\s(Id|versao)=""[^""]*"""
            xmlOriginal = System.Text.RegularExpressions.Regex.Replace(xmlOriginal, pattern, "")

            ' Carregar no DataSet
            Using reader As New StringReader(xmlOriginal)
                Using xmlReader As New XmlTextReader(reader)
                    DsXml.ReadXml(xmlReader)
                End Using
            End Using

            If DsXml.Tables.Count = 0 Then
                Throw New Exception("XML lido, mas sem tabelas.")
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        Return True

    End Function

    Public Shared Function BuscarEmailPorUsuario(ByVal sUsuario As String)
        Try
            Dim sql As String = "
            SELECT Email FROM Usuarios
            WHERE Usuario_Id = '" & sUsuario & "'"

            Dim banco As New AcessaBanco()
            Dim dt As DataSet = banco.ConsultaDataSet(sql, "Usuarios")

            If dt.Tables.Count = 0 Then
                Throw New Exception("Usuário não encontrado no sistema.")
            End If

            Dim lsUsuarios As New List(Of String)

            For Each rowDt As DataRow In dt.Tables(0).Rows
                lsUsuarios.Append(rowDt("Email"))

                Return lsUsuarios.Item(0)
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function Sha256(text As String) As String
        Using sha As SHA256 = SHA256Cng.Create()
            Dim bytes() As Byte = Encoding.UTF8.GetBytes(text)
            Dim hash() As Byte = sha.ComputeHash(bytes)
            Return BitConverter.ToString(hash).Replace("-", "").ToLower()
        End Using
    End Function
End Class