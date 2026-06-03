Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class DadosBoleto
    Implements IBaseEntity

    Dim pCodBanco As String
    Dim pCodMoeda As String
    Dim pDigitoVerificador As String
    Dim pFatorDeVencimento As String
    Dim pValor As String
    Dim pCampoLivre As String
    Dim pAgencia As String
    Dim pDigAgencia As String
    Dim pCarteira As String
    Dim pNossoNumero As String
    Dim pContaCorrente As String
    Dim pDigContaCorrente As String
    Dim pDacCodigo As String
    Dim pEnviarVlrDocumento As Boolean = True

    Public Property CodBanco() As String
        Get
            Return pCodBanco
        End Get
        Set(ByVal Value As String)
            pCodBanco = Value
        End Set
    End Property

    Public Property CodMoeda() As String
        Get
            Return pCodMoeda
        End Get
        Set(ByVal Value As String)
            pCodMoeda = Value
        End Set
    End Property

    Public Property DigitoVerificador() As String
        Get
            Return pDigitoVerificador
        End Get
        Set(ByVal Value As String)
            pDigitoVerificador = Value
        End Set
    End Property

    Public Property FatorDeVencimento() As String
        Get
            Return pFatorDeVencimento
        End Get
        Set(ByVal Value As String)
            pFatorDeVencimento = Value
        End Set
    End Property

    Public Property Valor() As String
        Get
            Return pValor
        End Get
        Set(ByVal Value As String)
            pValor = Value
        End Set
    End Property

    Public Property CampoLivre() As String
        Get
            Return pCampoLivre
        End Get
        Set(ByVal Value As String)
            pCampoLivre = Value
        End Set
    End Property

    Public Property Agencia() As String
        Get
            Return pAgencia
        End Get
        Set(ByVal Value As String)
            pAgencia = Value
        End Set
    End Property

    Public ReadOnly Property DigAgencia() As String
        Get
            If Agencia = Nothing Then
                Return ""
            Else
                Return Mod11(Agencia.Trim, 11).ToString
            End If
        End Get
    End Property

    Public Property Carteira() As String
        Get
            Return pCarteira
        End Get
        Set(ByVal Value As String)
            pCarteira = Value
        End Set
    End Property

    Public Property NossoNumero() As String
        Get
            Return pNossoNumero
        End Get
        Set(ByVal Value As String)
            pNossoNumero = Value
        End Set
    End Property

    Public Property ContaCorrente() As String
        Get
            Return pContaCorrente
        End Get
        Set(ByVal Value As String)
            pContaCorrente = Value
        End Set
    End Property

    Public ReadOnly Property DigContaCorrente() As String
        Get
            If ContaCorrente = Nothing Then
                Return ""
            Else
                Return Mod11(ContaCorrente.Trim, 11).ToString
            End If
        End Get
    End Property

    Public Property DacCodigo() As String
        Get
            Return pDacCodigo
        End Get
        Set(ByVal Value As String)
            pDacCodigo = Value
        End Set
    End Property

    Public Property EnviarVlrDocumento() As Boolean
        Get
            Return pEnviarVlrDocumento
        End Get
        Set(ByVal Value As Boolean)
            pEnviarVlrDocumento = Value
        End Set
    End Property

    Sub New(ByVal pCodigo As String, ByVal Digitado As Boolean)
        pCodigo = pCodigo.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "")
        CodBanco = Mid(pCodigo, 1, 3)
        If CodBanco <> "237" And Digitado = True Then

            CodMoeda = Mid(pCodigo, 4, 1)
            CampoLivre = Mid(pCodigo, 5, 5) & Mid(pCodigo, 11, 10) & Mid(pCodigo, 22, 10)
            DigitoVerificador = Mid(pCodigo, 33, 1)
            Try
                Valor = CDbl(Mid(pCodigo, 38, 10)) / 100
                If Valor = 0 Then
                    EnviarVlrDocumento = False
                End If
            Catch ex As Exception
                Valor = "0"
                EnviarVlrDocumento = False
            End Try

            Try
                FatorDeVencimento = Mid(pCodigo.Trim, 34, 4)
            Catch ex As Exception
                FatorDeVencimento = "0000"
                EnviarVlrDocumento = False
            End Try



        ElseIf CodBanco <> "237" And Digitado = False Then


            CodMoeda = Mid(pCodigo, 4, 1)
            DigitoVerificador = Mid(pCodigo, 5, 1)
            FatorDeVencimento = Mid(pCodigo, 6, 4)

            'CampoLivre = pCodigo.Substring(5, 5).Substring(11, 10).Substring(22, 10)
            CampoLivre = Mid(pCodigo, 20, 25)
            Try
                Valor = Mid(pCodigo, 10, 10) / 100
                If CDbl(Valor) / 100 = 0 Then
                    EnviarVlrDocumento = False
                End If
            Catch ex As Exception
                Valor = "0"
                EnviarVlrDocumento = False
            End Try



        ElseIf CodBanco = "237" And Digitado = True Then
            CodMoeda = Mid(pCodigo, 4, 1)
            Agencia = Mid(pCodigo, 5, 4)
            Carteira = Mid(pCodigo, 10, 2)
            NossoNumero = Mid(pCodigo, 12, 10)
            ContaCorrente = Mid(pCodigo, 24, 7)
            DacCodigo = Mid(pCodigo, 31, 1)

            CampoLivre = Mid(pCodigo, 5, 5) & Mid(pCodigo, 11, 10) & Mid(pCodigo, 22, 10)
            DigitoVerificador = Mid(pCodigo, 33, 1)
            Try
                Valor = Mid(pCodigo, 38, 10) / 100
                If CDbl(Valor) = 0 Then
                    EnviarVlrDocumento = False
                End If
            Catch ex As Exception
                Valor = "0"
                EnviarVlrDocumento = False
            End Try

            Try
                FatorDeVencimento = Mid(pCodigo.Trim, 34, 4)
            Catch ex As Exception
                FatorDeVencimento = "0000"
                EnviarVlrDocumento = False
            End Try

        ElseIf CodBanco = "237" And Digitado = False Then
            CodMoeda = Mid(pCodigo, 4, 1)
            DigitoVerificador = Mid(pCodigo, 5, 1)
            FatorDeVencimento = Mid(pCodigo, 6, 4)

            Agencia = Mid(pCodigo, 20, 4)
            Carteira = Mid(pCodigo, 24, 2)
            NossoNumero = Mid(pCodigo, 26, 11)
            ContaCorrente = Mid(pCodigo, 37, 7)
            CampoLivre = Mid(pCodigo, 20, 25)
            DigitoVerificador = Mid(pCodigo, 5, 1)
            Try
                Valor = Mid(pCodigo, 10, 10) / 100
                If CDbl(Valor) / 100 = 0 Then
                    EnviarVlrDocumento = False
                End If
            Catch ex As Exception
                EnviarVlrDocumento = False
                Valor = "0"
            End Try

        End If
    End Sub

    Sub New()

    End Sub

    Friend Shared Function Mod11(ByVal value As String, ByVal Base As Integer) As Integer

        Dim Digito As Integer, Soma As Integer = 0, Peso As Integer = 2


        For i As Integer = value.Length To 1 Step -1
            Soma = Soma + (Convert.ToInt32(Strings.Mid(value, i, 1)) * Peso)
            If Peso = Base Then
                Peso = 2
            Else
                Peso = Peso + 1
            End If
        Next

        Digito = 11 - (Soma Mod 11)
        If Digito > 9 Then Digito = 0

        'If (Digito > 9) OrElse (Digito = 0) OrElse (Digito = 1) Then
        '    Digito = 1
        'End If

        Return Digito
    End Function

    Public Function ValidaCodigoBarras(ByVal pCodigo As String, ByVal Digitado As Boolean) As Boolean
        If Digitado = True Then
            If pCodigo <> FormataLinhaDigitavel(pCodigo) Then
                MsgBox("Codigo de Barras Invalido", MsgBoxStyle.Information, "Financeiro")
                Return False
            End If
        ElseIf Digitado = False Then
            If FormataCodigoBarra(pCodigo) = False Then
                MsgBox("Codigo de Barras Invalido", MsgBoxStyle.Information, "Financeiro")
                Return False
            End If

        End If
        Return True


    End Function

    Private Function FormataLinhaDigitavel(ByVal _codigo As String) As String
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
                'res = (res / 10) + (res Mod 10)
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

    Private Function FormataCodigoBarra(ByVal _codigo As String) As Boolean

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


        'If Soma Mod 11 > 1 Or Soma Mod 11 < 10 Then
        '    Digito = 11 - Soma Mod 11
        'Else
        '    Digito = 1
        'End If


        If Digito = pdigito Then
            Return True
        Else
            Return False
        End If
        Return Digito


    End Function

    Public Function ValidaFatorDeVencimento(ByVal pCodigo As String, ByVal Digitado As Boolean, ByVal Vencimento As Date) As Boolean
        pCodigo = pCodigo.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "")
        Dim DataBase As Date = "07/10/1997"
        Dim FatorDeVencimento As Integer = 0
        If Digitado = True Then
            FatorDeVencimento = CInt(Mid(pCodigo, 34, 4))
        End If
        DataBase = DataBase.AddDays(FatorDeVencimento)
        'While DataBase.DayOfWeek = DayOfWeek.Saturday Or DataBase.DayOfWeek = DayOfWeek.Sunday Or VerificaDatasNaoProgramaveis(pTitulo, DataBase) = True
        '    DataBase = DataBase.AddDays(1)
        'End While
        If DataBase.Equals(Vencimento) = False Then
            Return False
        End If
        Return True


    End Function

    'Public Function VerificaDatasNaoProgramaveis(ByVal pTitulo As FrmTitulosAgrupados.TituloAgrupado, ByVal data As Date) As Boolean
    '    Dim sql As String
    '    Dim dsdatas As New DataSet
    '    Dim banco As New Utilitarios.AcessaBanco
    '    Dim dr As DataRow
    '    sql = "select * from DatasNaoProgramaveis where Empresa_Id='" & pTitulo.EmpresaCredito & "' and EndEmpresa_Id=" & pTitulo.EnderecoCredito & " and Data_Id = '" & Format(CDate(data), "yyyy-MM-dd") & "'"
    '    dsdatas = banco.ConsultaDataSet(sql, "DatasNaoProgramaveis")
    '    For Each dr In dsdatas.Tables(0).Rows
    '        Return True
    '    Next
    '    Return False

    'End Function

    Public Function ValidaValorVencimento(ByVal pCodigo As String, ByVal Valor As Decimal) As Boolean
        pCodigo = pCodigo.Replace(".", "").Replace("-", "").Replace(" ", "").Replace(",", "")

        If Valor <> (Mid(pCodigo, 38, 10) / 100) Then
            Return False
        End If
        Return True
    End Function

End Class
