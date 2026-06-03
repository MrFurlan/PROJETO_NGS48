Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListGerenciadorFinanceiro
    Inherits List(Of GerenciadorFinanceiro)

#Region "Construtor"
    Public Sub New(Optional ByVal Carregar As Boolean = False)
        If Not Carregar Then Exit Sub
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = ""

        sql = "select isnull(SbR.ano,SbP.Ano)     as Ano," & vbCrLf & _
              "       isnull(SbR.Moeda,SbP.Moeda) as CodigoMoeda ," & vbCrLf & _
              "       M.Descricao                 as DescricaoMoeda," & vbCrLf & _
              "       isnull(Receber01,0)         as Receber01," & vbCrLf & _
              "       isnull(Pagar01,0)           as Pagar01," & vbCrLf & _
              "       isnull(Receber02,0)         as Receber02," & vbCrLf & _
              "       isnull(Pagar02,0)           as Pagar02," & vbCrLf & _
              "       isnull(Receber03,0)         as Receber03," & vbCrLf & _
              "       isnull(Pagar03,0)           as Pagar03," & vbCrLf & _
              "       isnull(Receber04,0)         as Receber04," & vbCrLf & _
              "       isnull(Pagar04,0)           as Pagar04," & vbCrLf & _
              "       isnull(Receber05,0)         as Receber05," & vbCrLf & _
              "       isnull(Pagar05,0)           as Pagar05," & vbCrLf & _
              "       isnull(Receber06,0)         as Receber06," & vbCrLf & _
              "       isnull(Pagar06,0)           as Pagar06," & vbCrLf & _
              "       isnull(Receber07,0)         as Receber07," & vbCrLf & _
              "       isnull(Pagar07,0)           as Pagar07," & vbCrLf & _
              "       isnull(Receber08,0)         as Receber08," & vbCrLf & _
              "       isnull(Pagar08,0)           as Pagar08," & vbCrLf & _
              "       isnull(Receber09,0)         as Receber09," & vbCrLf & _
              "       isnull(Pagar09,0)           as Pagar09," & vbCrLf & _
              "       isnull(Receber10,0)         as Receber10," & vbCrLf & _
              "       isnull(Pagar10,0)           as Pagar10," & vbCrLf & _
              "       isnull(Receber11,0)         as Receber11," & vbCrLf & _
              "       isnull(Pagar11,0)           as Pagar11," & vbCrLf & _
              "       isnull(Receber12,0)         as Receber12," & vbCrLf & _
              "       isnull(Pagar12,0)           as Pagar12" & vbCrLf & _
              " from (Select year(Movimento) as ano," & vbCrLf & _
              "	  	     Moeda," & vbCrLf & _
              "		     sum(case when month(Movimento) =  1 then ValorLiquido else 0 end) as Receber01," & vbCrLf & _
              "		     sum(case when month(Movimento) =  2 then ValorLiquido else 0 end) as Receber02," & vbCrLf & _
              "		     sum(case when month(Movimento) =  3 then ValorLiquido else 0 end) as Receber03," & vbCrLf & _
              "		     sum(case when month(Movimento) =  4 then ValorLiquido else 0 end) as Receber04," & vbCrLf & _
              "		     sum(case when month(Movimento) =  5 then ValorLiquido else 0 end) as Receber05," & vbCrLf & _
              "		     sum(case when month(Movimento) =  6 then ValorLiquido else 0 end) as Receber06," & vbCrLf & _
              "		     sum(case when month(Movimento) =  7 then ValorLiquido else 0 end) as Receber07," & vbCrLf & _
              "		     sum(case when month(Movimento) =  8 then ValorLiquido else 0 end) as Receber08," & vbCrLf & _
              "		     sum(case when month(Movimento) =  9 then ValorLiquido else 0 end) as Receber09," & vbCrLf & _
              "		     sum(case when month(Movimento) = 10 then ValorLiquido else 0 end) as Receber10," & vbCrLf & _
              "		     sum(case when month(Movimento) = 11 then ValorLiquido else 0 end) as Receber11," & vbCrLf & _
              "		     sum(case when month(Movimento) = 12 then ValorLiquido else 0 end) as Receber12" & vbCrLf & _
              "	    from ContasaReceber" & vbCrLf & _
              "	   where situacao = 1" & vbCrLf & _
              "	     and provisao <> 1" & vbCrLf & _
              "	   group by year(Movimento), Moeda" & vbCrLf & _
              "     ) SbR" & vbCrLf & _
              " full join (Select year(Movimento) as ano," & vbCrLf & _
              "	  	     Moeda," & vbCrLf & _
              "		     sum(case when month(Movimento) =  1 then ValorLiquido else 0 end) as Pagar01," & vbCrLf & _
              "		     sum(case when month(Movimento) =  2 then ValorLiquido else 0 end) as Pagar02," & vbCrLf & _
              "		     sum(case when month(Movimento) =  3 then ValorLiquido else 0 end) as Pagar03," & vbCrLf & _
              "		     sum(case when month(Movimento) =  4 then ValorLiquido else 0 end) as Pagar04," & vbCrLf & _
              "		     sum(case when month(Movimento) =  5 then ValorLiquido else 0 end) as Pagar05," & vbCrLf & _
              "		     sum(case when month(Movimento) =  6 then ValorLiquido else 0 end) as Pagar06," & vbCrLf & _
              "		     sum(case when month(Movimento) =  7 then ValorLiquido else 0 end) as Pagar07," & vbCrLf & _
              "		     sum(case when month(Movimento) =  8 then ValorLiquido else 0 end) as Pagar08," & vbCrLf & _
              "		     sum(case when month(Movimento) =  9 then ValorLiquido else 0 end) as Pagar09," & vbCrLf & _
              "		     sum(case when month(Movimento) = 10 then ValorLiquido else 0 end) as Pagar10," & vbCrLf & _
              "		     sum(case when month(Movimento) = 11 then ValorLiquido else 0 end) as Pagar11," & vbCrLf & _
              "		     sum(case when month(Movimento) = 12 then ValorLiquido else 0 end) as Pagar12" & vbCrLf & _
              "	    from ContasaPagar" & vbCrLf & _
              "	   where situacao = 1" & vbCrLf & _
              "	     and provisao <> 1" & vbCrLf & _
              "	   group by year(Movimento), Moeda" & vbCrLf & _
              "     ) SbP" & vbCrLf & _
              "   on SbR.ano   = sbP.ano" & vbCrLf & _
              "   and SbR.Moeda = sbP.Moeda" & vbCrLf & _
              " Inner join Moedas M" & vbCrLf & _
              "    on M.Moeda_Id = isnull(SbR.Moeda,SbP.Moeda)" & vbCrLf & _
              " order by isnull(SbR.ano,SbP.Ano)," & vbCrLf & _
              "          isnull(SbR.Moeda,SbP.Moeda)" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "Financeiro")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim GF As New GerenciadorFinanceiro
            GF.Ano = row("ANO")
            GF.CodigoMoeda = row("CodigoMoeda")
            GF.DescricaoMoeda = row("DescricaoMoeda")
            GF.Receber01 = row("Receber01")
            GF.Receber02 = row("Receber02")
            GF.Receber03 = row("Receber03")
            GF.Receber04 = row("Receber04")
            GF.Receber05 = row("Receber05")
            GF.Receber06 = row("Receber06")
            GF.Receber07 = row("Receber07")
            GF.Receber08 = row("Receber08")
            GF.Receber09 = row("Receber09")
            GF.Receber10 = row("Receber10")
            GF.Receber11 = row("Receber11")
            GF.Receber12 = row("Receber12")

            GF.Pagar01 = row("Pagar01")
            GF.Pagar02 = row("Pagar02")
            GF.Pagar03 = row("Pagar03")
            GF.Pagar04 = row("Pagar04")
            GF.Pagar05 = row("Pagar05")
            GF.Pagar06 = row("Pagar06")
            GF.Pagar07 = row("Pagar07")
            GF.Pagar08 = row("Pagar08")
            GF.Pagar09 = row("Pagar09")
            GF.Pagar10 = row("Pagar10")
            GF.Pagar11 = row("Pagar11")
            GF.Pagar12 = row("Pagar12")

            Add(GF)
        Next
        carregarJuros()
    End Sub
#End Region

    Private _Titulos As ListTitulo

#Region "Methods"
    Private Sub carregarJuros()
        Dim Where As String
        Where = "     Provisao <> 1" & _
                " and Situacao = 1" & _
                " and DATEDIFF(DAY,Vencimento,GETDATE()) > 0 "

        _Titulos = New ListTitulo(Where)

        For Each GF As GerenciadorFinanceiro In Me
            For Each tit As Titulo In _Titulos
                If GF.Ano = tit.Vencimento.Year And GF.CodigoMoeda = tit.CodigoMoeda Then
                    If tit.ReceberPagar = "R" Then tit.CalculaJuros(Now.Date)
                    Select Case tit.Prorrogacao.Month
                        Case 1
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros01 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros01 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 2
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros02 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros02 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 3
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros03 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros03 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 4
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros04 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros04 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 5
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros05 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros05 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 6
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros06 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros06 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 7
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros07 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros07 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 8
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros08 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros08 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 9
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros09 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros09 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 10
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros10 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros10 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 11
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros11 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros11 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                        Case 12
                            If tit.ReceberPagar = "R" Then
                                GF.ReceberJuros12 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            Else
                                GF.PagarJuros12 += IIf(tit.Moeda.Classificacao = eTiposMoeda.Oficial, tit.Juros, tit.MoedaJuros)
                            End If
                    End Select
                End If
            Next
        Next

    End Sub
#End Region

End Class

<Serializable()> _
Public Class GerenciadorFinanceiro
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _CasasDecimais As Integer = 0
    Private _Ano As Integer
    Private _CodigoMoeda As Integer
    Private _Moeda As Moeda
    Private _DescricaoMoeda As String

    Private _Receber01 As Decimal
    Private _Receber02 As Decimal
    Private _Receber03 As Decimal
    Private _Receber04 As Decimal
    Private _Receber05 As Decimal
    Private _Receber06 As Decimal
    Private _Receber07 As Decimal
    Private _Receber08 As Decimal
    Private _Receber09 As Decimal
    Private _Receber10 As Decimal
    Private _Receber11 As Decimal
    Private _Receber12 As Decimal

    Private _ReceberJuros01 As Decimal
    Private _ReceberJuros02 As Decimal
    Private _ReceberJuros03 As Decimal
    Private _ReceberJuros04 As Decimal
    Private _ReceberJuros05 As Decimal
    Private _ReceberJuros06 As Decimal
    Private _ReceberJuros07 As Decimal
    Private _ReceberJuros08 As Decimal
    Private _ReceberJuros09 As Decimal
    Private _ReceberJuros10 As Decimal
    Private _ReceberJuros11 As Decimal
    Private _ReceberJuros12 As Decimal

    Private _Pagar01 As Decimal
    Private _Pagar02 As Decimal
    Private _Pagar03 As Decimal
    Private _Pagar04 As Decimal
    Private _Pagar05 As Decimal
    Private _Pagar06 As Decimal
    Private _Pagar07 As Decimal
    Private _Pagar08 As Decimal
    Private _Pagar09 As Decimal
    Private _Pagar10 As Decimal
    Private _Pagar11 As Decimal
    Private _Pagar12 As Decimal

    Private _PagarJuros01 As Decimal
    Private _PagarJuros02 As Decimal
    Private _PagarJuros03 As Decimal
    Private _PagarJuros04 As Decimal
    Private _PagarJuros05 As Decimal
    Private _PagarJuros06 As Decimal
    Private _PagarJuros07 As Decimal
    Private _PagarJuros08 As Decimal
    Private _PagarJuros09 As Decimal
    Private _PagarJuros10 As Decimal
    Private _PagarJuros11 As Decimal
    Private _PagarJuros12 As Decimal
#End Region

#Region "Property"
    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property CodigoMoeda() As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(ByVal value As Integer)
            _CodigoMoeda = value
            _Moeda = Nothing
        End Set
    End Property

    Public ReadOnly Property Moeda() As Moeda
        Get
            If _Moeda Is Nothing And _CodigoMoeda > 0 Then _Moeda = New Moeda(_CodigoMoeda)
            Return _Moeda
        End Get
    End Property

    Public Property DescricaoMoeda() As String
        Get
            Return _DescricaoMoeda
        End Get
        Set(ByVal value As String)
            _DescricaoMoeda = value
        End Set
    End Property

    Public Property Receber01() As Decimal
        Get
            Return _Receber01
        End Get
        Set(ByVal value As Decimal)
            _Receber01 = value
        End Set
    End Property

    Public ReadOnly Property R01() As String
        Get
            Return Funcoes.FormatarValor(_Receber01, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber02() As Decimal
        Get
            Return _Receber02
        End Get
        Set(ByVal value As Decimal)
            _Receber02 = value
        End Set
    End Property

    Public ReadOnly Property R02() As String
        Get
            Return Funcoes.FormatarValor(_Receber02, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber03() As Decimal
        Get
            Return _Receber03
        End Get
        Set(ByVal value As Decimal)
            _Receber03 = value
        End Set
    End Property

    Public ReadOnly Property R03() As String
        Get
            Return Funcoes.FormatarValor(_Receber03, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber04() As Decimal
        Get
            Return _Receber04
        End Get
        Set(ByVal value As Decimal)
            _Receber04 = value
        End Set
    End Property

    Public ReadOnly Property R04() As String
        Get
            Return Funcoes.FormatarValor(_Receber04, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber05() As Decimal
        Get
            Return _Receber05
        End Get
        Set(ByVal value As Decimal)
            _Receber05 = value
        End Set
    End Property

    Public ReadOnly Property R05() As String
        Get
            Return Funcoes.FormatarValor(_Receber05, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber06() As Decimal
        Get
            Return _Receber06
        End Get
        Set(ByVal value As Decimal)
            _Receber06 = value
        End Set
    End Property

    Public ReadOnly Property R06() As String
        Get
            Return Funcoes.FormatarValor(_Receber06, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber07() As Decimal
        Get
            Return _Receber07
        End Get
        Set(ByVal value As Decimal)
            _Receber07 = value
        End Set
    End Property

    Public ReadOnly Property R07() As String
        Get
            Return Funcoes.FormatarValor(_Receber07, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber08() As Decimal
        Get
            Return _Receber08
        End Get
        Set(ByVal value As Decimal)
            _Receber08 = value
        End Set
    End Property

    Public ReadOnly Property R08() As String
        Get
            Return Funcoes.FormatarValor(_Receber08, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber09() As Decimal
        Get
            Return _Receber09
        End Get
        Set(ByVal value As Decimal)
            _Receber09 = value
        End Set
    End Property

    Public ReadOnly Property R09() As String
        Get
            Return Funcoes.FormatarValor(_Receber09, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber10() As Decimal
        Get
            Return _Receber10
        End Get
        Set(ByVal value As Decimal)
            _Receber10 = value
        End Set
    End Property

    Public ReadOnly Property R10() As String
        Get
            Return Funcoes.FormatarValor(_Receber10, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber11() As Decimal
        Get
            Return _Receber11
        End Get
        Set(ByVal value As Decimal)
            _Receber11 = value
        End Set
    End Property

    Public ReadOnly Property R11() As String
        Get
            Return Funcoes.FormatarValor(_Receber11, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Receber12() As Decimal
        Get
            Return _Receber12
        End Get
        Set(ByVal value As Decimal)
            _Receber12 = value
        End Set
    End Property

    Public ReadOnly Property R12() As String
        Get
            Return Funcoes.FormatarValor(_Receber12, 18, _CasasDecimais)
        End Get
    End Property

    Public ReadOnly Property ReceberTotal() As String
        Get
            Return Funcoes.FormatarValor(_Receber01 + _Receber02 + _Receber03 + _Receber04 + _Receber05 + _Receber06 + _Receber07 + _Receber08 + _Receber09 + _Receber10 + _Receber11 + _Receber12, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros01() As Decimal
        Get
            Return _ReceberJuros01
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros01 = value
        End Set
    End Property

    Public ReadOnly Property RJ01() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros01, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros02() As Decimal
        Get
            Return _ReceberJuros02
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros02 = value
        End Set
    End Property

    Public ReadOnly Property RJ02() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros02, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros03() As Decimal
        Get
            Return _ReceberJuros03
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros03 = value
        End Set
    End Property

    Public ReadOnly Property RJ03() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros03, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros04() As Decimal
        Get
            Return _ReceberJuros04
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros04 = value
        End Set
    End Property

    Public ReadOnly Property RJ04() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros04, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros05() As Decimal
        Get
            Return _ReceberJuros05
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros05 = value
        End Set
    End Property

    Public ReadOnly Property RJ05() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros05, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros06() As Decimal
        Get
            Return _ReceberJuros06
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros06 = value
        End Set
    End Property

    Public ReadOnly Property RJ06() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros06, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros07() As Decimal
        Get
            Return _ReceberJuros07
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros07 = value
        End Set
    End Property

    Public ReadOnly Property RJ07() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros07, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros08() As Decimal
        Get
            Return _ReceberJuros08
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros08 = value
        End Set
    End Property

    Public ReadOnly Property RJ08() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros08, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros09() As Decimal
        Get
            Return _ReceberJuros09
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros09 = value
        End Set
    End Property

    Public ReadOnly Property RJ09() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros09, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros10() As Decimal
        Get
            Return _ReceberJuros10
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros10 = value
        End Set
    End Property

    Public ReadOnly Property RJ10() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros10, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros11() As Decimal
        Get
            Return _ReceberJuros11
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros11 = value
        End Set
    End Property

    Public ReadOnly Property RJ11() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros11, 18, _CasasDecimais)
        End Get
    End Property

    Public Property ReceberJuros12() As Decimal
        Get
            Return _ReceberJuros12
        End Get
        Set(ByVal value As Decimal)
            _ReceberJuros12 = value
        End Set
    End Property

    Public ReadOnly Property RJ12() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros12, 18, _CasasDecimais)
        End Get
    End Property

    Public ReadOnly Property ReceberJurosTotal() As String
        Get
            Return Funcoes.FormatarValor(_ReceberJuros01 + _ReceberJuros02 + _ReceberJuros03 + _ReceberJuros04 + _ReceberJuros05 + _ReceberJuros06 + _ReceberJuros07 + _ReceberJuros08 + _ReceberJuros09 + _ReceberJuros10 + _ReceberJuros11 + _ReceberJuros12, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar01() As Decimal
        Get
            Return _Pagar01
        End Get
        Set(ByVal value As Decimal)
            _Pagar01 = value
        End Set
    End Property

    Public ReadOnly Property P01() As String
        Get
            Return Funcoes.FormatarValor(_Pagar01, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar02() As Decimal
        Get
            Return _Pagar02
        End Get
        Set(ByVal value As Decimal)
            _Pagar02 = value
        End Set
    End Property

    Public ReadOnly Property P02() As String
        Get
            Return Funcoes.FormatarValor(_Pagar02, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar03() As Decimal
        Get
            Return _Pagar03
        End Get
        Set(ByVal value As Decimal)
            _Pagar03 = value
        End Set
    End Property

    Public ReadOnly Property P03() As String
        Get
            Return Funcoes.FormatarValor(_Pagar03, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar04() As Decimal
        Get
            Return _Pagar04
        End Get
        Set(ByVal value As Decimal)
            _Pagar04 = value
        End Set
    End Property

    Public ReadOnly Property P04() As String
        Get
            Return Funcoes.FormatarValor(_Pagar04, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar05() As Decimal
        Get
            Return _Pagar05
        End Get
        Set(ByVal value As Decimal)
            _Pagar05 = value
        End Set
    End Property

    Public ReadOnly Property P05() As String
        Get
            Return Funcoes.FormatarValor(_Pagar05, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar06() As Decimal
        Get
            Return _Pagar06
        End Get
        Set(ByVal value As Decimal)
            _Pagar06 = value
        End Set
    End Property

    Public ReadOnly Property P06() As String
        Get
            Return Funcoes.FormatarValor(_Pagar06, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar07() As Decimal
        Get
            Return _Pagar07
        End Get
        Set(ByVal value As Decimal)
            _Pagar07 = value
        End Set
    End Property

    Public ReadOnly Property P07() As String
        Get
            Return Funcoes.FormatarValor(_Pagar07, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar08() As Decimal
        Get
            Return _Pagar08
        End Get
        Set(ByVal value As Decimal)
            _Pagar08 = value
        End Set
    End Property

    Public ReadOnly Property P08() As String
        Get
            Return Funcoes.FormatarValor(_Pagar08, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar09() As Decimal
        Get
            Return _Pagar09
        End Get
        Set(ByVal value As Decimal)
            _Pagar09 = value
        End Set
    End Property

    Public ReadOnly Property P09() As String
        Get
            Return Funcoes.FormatarValor(_Pagar09, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar10() As Decimal
        Get
            Return _Pagar10
        End Get
        Set(ByVal value As Decimal)
            _Pagar10 = value
        End Set
    End Property

    Public ReadOnly Property P10() As String
        Get
            Return Funcoes.FormatarValor(_Pagar10, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar11() As Decimal
        Get
            Return _Pagar11
        End Get
        Set(ByVal value As Decimal)
            _Pagar11 = value
        End Set
    End Property

    Public ReadOnly Property P11() As String
        Get
            Return Funcoes.FormatarValor(_Pagar11, 18, _CasasDecimais)
        End Get
    End Property

    Public Property Pagar12() As Decimal
        Get
            Return _Pagar12
        End Get
        Set(ByVal value As Decimal)
            _Pagar12 = value
        End Set
    End Property

    Public ReadOnly Property P12() As String
        Get
            Return Funcoes.FormatarValor(_Pagar12, 18, _CasasDecimais)
        End Get
    End Property

    Public ReadOnly Property PagarTotal() As String
        Get
            Return Funcoes.FormatarValor(_Pagar01 + _Pagar02 + _Pagar03 + _Pagar04 + _Pagar05 + _Pagar06 + _Pagar07 + _Pagar08 + _Pagar09 + _Pagar10 + _Pagar11 + _Pagar12, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros01() As Decimal
        Get
            Return _PagarJuros01
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros01 = value
        End Set
    End Property

    Public ReadOnly Property PJ01() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros01, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros02() As Decimal
        Get
            Return _PagarJuros02
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros02 = value
        End Set
    End Property

    Public ReadOnly Property PJ02() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros02, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros03() As Decimal
        Get
            Return _PagarJuros03
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros03 = value
        End Set
    End Property

    Public ReadOnly Property PJ03() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros03, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros04() As Decimal
        Get
            Return _PagarJuros04
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros04 = value
        End Set
    End Property

    Public ReadOnly Property PJ04() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros04, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros05() As Decimal
        Get
            Return _PagarJuros05
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros05 = value
        End Set
    End Property

    Public ReadOnly Property PJ05() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros05, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros06() As Decimal
        Get
            Return _PagarJuros06
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros06 = value
        End Set
    End Property

    Public ReadOnly Property PJ06() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros06, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros07() As Decimal
        Get
            Return _PagarJuros07
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros07 = value
        End Set
    End Property

    Public ReadOnly Property PJ07() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros07, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros08() As Decimal
        Get
            Return _PagarJuros08
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros08 = value
        End Set
    End Property

    Public ReadOnly Property PJ08() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros08, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros09() As Decimal
        Get
            Return _PagarJuros09
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros09 = value
        End Set
    End Property

    Public ReadOnly Property PJ09() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros09, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros10() As Decimal
        Get
            Return _PagarJuros10
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros10 = value
        End Set
    End Property

    Public ReadOnly Property PJ10() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros10, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros11() As Decimal
        Get
            Return _PagarJuros11
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros11 = value
        End Set
    End Property

    Public ReadOnly Property PJ11() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros11, 18, _CasasDecimais)
        End Get
    End Property

    Public Property PagarJuros12() As Decimal
        Get
            Return _PagarJuros12
        End Get
        Set(ByVal value As Decimal)
            _PagarJuros12 = value
        End Set
    End Property

    Public ReadOnly Property PJ12() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros12, 18, _CasasDecimais)
        End Get
    End Property

    Public ReadOnly Property PagarJurosTotal() As String
        Get
            Return Funcoes.FormatarValor(_PagarJuros01 + _PagarJuros02 + _PagarJuros03 + _PagarJuros04 + _PagarJuros05 + _PagarJuros06 + _PagarJuros07 + _PagarJuros08 + _PagarJuros09 + _PagarJuros10 + _PagarJuros11 + _PagarJuros12, 18, _CasasDecimais)
        End Get
    End Property
#End Region

End Class







