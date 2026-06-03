Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Math
Imports NGS.Lib.Uteis

Namespace Novo
    <Serializable()> _
    Public Class ListAdiantamentoBaixaNovo
        Inherits List(Of AdiantamentoBaixaNovo)

#Region "Construtor"
        Public Sub New()

        End Sub

        'Mostra as Baixas q o Adiantamento Sofreu
        Public Sub New(ByVal pAdiantamento As Novo.AdiantamentoNovo)
            Dim Sql As String
            Sql = "SELECT axb.Titulo_Id," & vbCrLf & _
                  "       axb.Sequencia_Id," & vbCrLf & _
                  "       tAd.DataBaixa as DataAdiantamento," & vbCrLf & _
                  "       TxC.ValorOficial as ValorOficialAdiantamento," & vbCrLf & _
                  "       TxC.ValorMoeda as ValorMoedaAdiantamento," & vbCrLf & _
                  "       Tad.Indice as IndiceAdiantamento," & vbCrLf & _
                  "       axb.TituloBaixa," & vbCrLf & _
                  "       axb.Movimento," & vbCrLf & _
                  "       TBxAd.Indice as indiceBaixa," & vbCrLf & _
                  "       axb.Lancamento," & vbCrLf & _
                  "       axb.TipoLancamento," & vbCrLf & _
                  "       axb.ValorOficial," & vbCrLf & _
                  "       axb.ValorMoeda" & vbCrLf & _
                  "  FROM AdiantamentosXBaixas axb" & vbCrLf & _
                  " Inner Join Titulos TAd " & vbCrLf & _
                  "    on TAd.Titulo_id = axb.Titulo_Id" & vbCrLf & _
                  " Inner Join titulosxcontacontabil txc" & vbCrLf & _
                  "    on Tad.Titulo_id           = txc.Titulo_id" & vbCrLf & _
                  "   and Tad.ContaContabilRecPag = txc.Conta_id" & vbCrLf & _
                  " Inner Join Titulos TBxAd" & vbCrLf & _
                  "    on TBxAd.Titulo_id = axb.TituloBaixa" & vbCrLf & _
                  " Where axb.Titulo_id = " & pAdiantamento.CodigoTitulo

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "AdiantamentoBaixa")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim AB As New AdiantamentoBaixaNovo(pAdiantamento)
                AB.Sequencia = row("Sequencia_Id")
                AB.CodigoTituloBaixa = row("TituloBaixa")
                AB.DataAdiantamento = row("DataAdiantamento")
                AB.ValorAdiantamentoOficial = row("ValorOficialAdiantamento")
                AB.ValorAdiantamentoMoeda = row("ValorMoedaAdiantamento")
                AB.IndiceAdiantamento = row("IndiceAdiantamento")
                AB.Movimento = row("Movimento")
                AB.IndiceBaixa = row("indiceBaixa")
                AB.Lancamento = row("Lancamento")
                AB.TipoLancamento = row("TipoLancamento")
                AB.ValorOficial = row("ValorOficial")
                AB.ValorMoeda = row("ValorMoeda")
                Me.Add(AB)
            Next
        End Sub

        'Mostra as Baixas nos Adiantamentos q o Titulo Efetuou
        Public Sub New(ByRef pTituloBaixa As Novo.TituloNovo)
            Dim Sql As String
            Sql = "SELECT axb.Titulo_Id," & vbCrLf & _
                  "       axb.Sequencia_Id," & vbCrLf & _
                  "       tAd.DataBaixa as DataAdiantamento," & vbCrLf & _
                  "       TxC.ValorOficial as ValorOficialAdiantamento," & vbCrLf & _
                  "       TxC.ValorMoeda as ValorMoedaAdiantamento," & vbCrLf & _
                  "       Tad.Indice as IndiceAdiantamento," & vbCrLf & _
                  "       axb.TituloBaixa," & vbCrLf & _
                  "       axb.Movimento," & vbCrLf & _
                  "       TBxAd.Indice as indiceBaixa," & vbCrLf & _
                  "       axb.Lancamento," & vbCrLf & _
                  "       axb.TipoLancamento," & vbCrLf & _
                  "       axb.ValorOficial," & vbCrLf & _
                  "       axb.ValorMoeda" & vbCrLf & _
                  "  FROM AdiantamentosXBaixas axb" & vbCrLf & _
                  " Inner Join Titulos TAd " & vbCrLf & _
                  "    on TAd.Titulo_id = axb.Titulo_Id" & vbCrLf & _
                  " Inner Join titulosxcontacontabil txc" & vbCrLf & _
                  "    on Tad.Titulo_id           = txc.Titulo_id" & vbCrLf & _
                  "   and Tad.ContaContabilRecPag = txc.Conta_id" & vbCrLf & _
                  " Inner Join Titulos TBxAd" & vbCrLf & _
                  "    on TBxAd.Titulo_id = axb.TituloBaixa" & vbCrLf & _
                  " Where axb.TituloBaixa = " & pTituloBaixa.Codigo

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "AdiantamentoBaixa")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Ad As New AdiantamentoNovo(New Novo.TituloNovo(row("Titulo_Id")))

                Dim AB As New Novo.AdiantamentoBaixaNovo(Ad)
                AB.Sequencia = row("Sequencia_Id")
                AB.CodigoTituloBaixa = row("TituloBaixa")
                AB.DataAdiantamento = row("DataAdiantamento")
                AB.ValorAdiantamentoOficial = row("ValorOficialAdiantamento")
                AB.ValorAdiantamentoMoeda = row("ValorMoedaAdiantamento")
                AB.IndiceAdiantamento = row("IndiceAdiantamento")
                AB.Movimento = row("Movimento")
                AB.IndiceBaixa = row("indiceBaixa")
                AB.Lancamento = row("Lancamento")
                AB.TipoLancamento = row("TipoLancamento")
                AB.ValorOficial = row("ValorOficial")
                AB.ValorMoeda = row("ValorMoeda")
                Me.Add(AB)
            Next
        End Sub
#End Region

#Region "Property"
        '*********************** BAIXAS ***************************
        Public ReadOnly Property TotalOficialBaixado()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "B" And bx.IUD <> "D"
                        Select bx.ValorOficial).Sum
            End Get
        End Property

        Public ReadOnly Property TotalMoedaBaixado()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "B" And bx.IUD <> "D"
                        Select bx.ValorMoeda).Sum
            End Get
        End Property

        '*********************** JUROS ***************************

        Public ReadOnly Property TotalOficialJuroRecebido()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "J" And bx.TipoLancamento = "R"
                        Select bx.ValorOficial).Sum
            End Get
        End Property

        Public ReadOnly Property TotalMoedaJuroRecebido()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "J" And bx.TipoLancamento = "R"
                        Select bx.ValorMoeda).Sum
            End Get
        End Property

        Public ReadOnly Property TotalOficialJuroPago()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "J" And bx.TipoLancamento = "P"
                        Select bx.ValorOficial).Sum
            End Get
        End Property

        Public ReadOnly Property TotalMoedaJuroPago()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "J" And bx.TipoLancamento = "P"
                        Select bx.ValorMoeda).Sum
            End Get
        End Property

        '********************* VARIACAO ***************************

        Public ReadOnly Property TotalOficialVariacaoAtiva()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "V" And bx.TipoLancamento = "A"
                        Select bx.ValorOficial).Sum
            End Get
        End Property

        Public ReadOnly Property TotalMoedaVariacaoAtiva()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "V" And bx.TipoLancamento = "A"
                        Select bx.ValorMoeda).Sum
            End Get
        End Property

        Public ReadOnly Property TotalOficialVariacaoPassiva()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "V" And bx.TipoLancamento = "P"
                        Select bx.ValorOficial).Sum
            End Get
        End Property

        Public ReadOnly Property TotalMoedaVariacaoPassiva()
            Get
                Return (From bx In Me
                        Where bx.Lancamento = "V" And bx.TipoLancamento = "P"
                        Select bx.ValorMoeda).Sum
            End Get
        End Property

        '************ NAO CONSEGUI FAZER O MAX DO LINQ RETORNA DIFERENTE DE NULL ENTAO FIZ ESSA MERDA PQ TO SEM TEMPO  *************
        Public ReadOnly Property ProximaSequenciaLivre As Integer
            Get
                Dim seq As Integer = 0
                For Each bx As AdiantamentoBaixaNovo In Me
                    If bx.Sequencia > seq Then seq = bx.Sequencia
                Next
                Return seq + 1
            End Get
        End Property
#End Region

#Region "Methods"
        Public Function Salvar() As Boolean
            Dim Banco As New AcessaBanco
            Dim Sqls As New ArrayList

            Sqls.Clear()
            Me.SalvarSql(Sqls)

            If Banco.GravaBanco(Sqls) Then
                Return True
            End If
            Return False
        End Function

        Public Sub SalvarSql(ByRef Sqls As ArrayList)
            For Each row As Novo.AdiantamentoBaixaNovo In Me
                If row.IUD <> "" Then
                    row.SalvarSql(Sqls)
                End If
            Next
        End Sub
#End Region
    End Class

    <Serializable()> _
    Public Class AdiantamentoBaixaNovo

#Region "Construtor"
        Public Sub New()

        End Sub

        Public Sub New(ByRef pAdiantamento As Novo.AdiantamentoNovo)
            _Adiantamento = pAdiantamento
            _CodigoTituloAdiantamento = pAdiantamento.CodigoTitulo
        End Sub
#End Region

#Region "fields"
        Private _IUD As String = ""
        Private _Adiantamento As Novo.AdiantamentoNovo
        Private _CodigoTituloAdiantamento As Integer
        Private _TituloAdiantamento As Novo.TituloNovo
        Private _Sequencia As Integer
        '*********** ADIANTAMENTO ****************
        Private _DataAdiantamento As DateTime
        Private _ValorAdiantamentoOficial As Decimal
        Private _ValorAdiantamentoMoeda As Decimal
        Private _IndiceAdiantamento As Decimal
        '************** BAIXA  *******************
        Private _CodigoTituloBaixa As Integer
        Private _TituloBaixa As Novo.TituloNovo
        Private _IndiceBaixa As Decimal

        Private _Movimento As DateTime
        Private _Lancamento As Char
        Private _TipoLancamento As Char

        Private _ValorOficial As Decimal
        Private _ValorMoeda As Decimal
#End Region

#Region "Property"
        Public Property IUD() As String
            Get
                Return _IUD
            End Get
            Set(ByVal value As String)
                _IUD = value
            End Set
        End Property

        Public Property CodigoTituloAdiantamento As Integer
            Get
                Return _CodigoTituloAdiantamento
            End Get
            Set(value As Integer)
                _CodigoTituloAdiantamento = value
                _TituloAdiantamento = Nothing
            End Set
        End Property

        Public ReadOnly Property TituloAdiantamento As Novo.TituloNovo
            Get
                If _TituloAdiantamento Is Nothing And Me.CodigoTituloAdiantamento > 0 Then
                    _TituloAdiantamento = New Novo.TituloNovo(Me.CodigoTituloAdiantamento)
                End If
                Return Nothing
            End Get
        End Property

        Public Property Adiantamento() As Novo.AdiantamentoNovo
            Get
                If _Adiantamento Is Nothing And _CodigoTituloAdiantamento > 0 Then _Adiantamento = New Novo.AdiantamentoNovo(TituloAdiantamento)
                Return _Adiantamento
            End Get
            Set(ByVal value As Novo.AdiantamentoNovo)
                _Adiantamento = value
            End Set
        End Property

        Public Property Sequencia() As Integer
            Get
                Return _Sequencia
            End Get
            Set(ByVal value As Integer)
                _Sequencia = value
            End Set
        End Property
        '***********************************************************
        '*******  ADIANTAMENTO *************************************
        '***********************************************************
        Public Property DataAdiantamento As DateTime
            Get
                Return _DataAdiantamento
            End Get
            Set(value As DateTime)
                _DataAdiantamento = value
            End Set
        End Property

        Public Property ValorAdiantamentoOficial As Decimal
            Get
                Return _ValorAdiantamentoOficial
            End Get
            Set(value As Decimal)
                _ValorAdiantamentoOficial = value
            End Set
        End Property

        Public Property ValorAdiantamentoMoeda As Decimal
            Get
                Return _ValorAdiantamentoMoeda
            End Get
            Set(value As Decimal)
                _ValorAdiantamentoMoeda = value
            End Set
        End Property

        Public Property IndiceAdiantamento As Decimal
            Get
                Return _IndiceAdiantamento
            End Get
            Set(value As Decimal)
                _IndiceAdiantamento = value
            End Set
        End Property

        Public ReadOnly Property SaldoOficialAdiantamento As Decimal
            Get
                Return Adiantamento.SaldoValorOficial
            End Get
        End Property

        Public ReadOnly Property SaldoMoedaAdiantamento As Decimal
            Get
                Return Adiantamento.SaldoValorMoeda
            End Get
        End Property

        '***********************************************************
        '*******  BAIXA ********************************************
        '***********************************************************
        Public Property CodigoTituloBaixa() As Integer
            Get
                Return _CodigoTituloBaixa
            End Get
            Set(ByVal value As Integer)
                _CodigoTituloBaixa = value
                _TituloBaixa = Nothing
            End Set
        End Property

        Public Property TituloBaixa() As Novo.TituloNovo
            Get
                If _TituloBaixa Is Nothing And _CodigoTituloBaixa > 0 Then _TituloBaixa = New Novo.TituloNovo(_CodigoTituloBaixa)
                Return _TituloBaixa
            End Get
            Set(ByVal value As Novo.TituloNovo)
                _TituloBaixa = value
            End Set
        End Property

        Public Property Movimento() As DateTime
            Get
                Return _Movimento
            End Get
            Set(ByVal value As DateTime)
                _Movimento = value
            End Set
        End Property

        Public Property IndiceBaixa As Decimal
            Get
                Return _IndiceBaixa
            End Get
            Set(value As Decimal)
                _IndiceBaixa = value
            End Set
        End Property

        Public Property Lancamento() As Char
            Get
                Return _Lancamento
            End Get
            Set(ByVal value As Char)
                _Lancamento = value
            End Set
        End Property

        Public Property TipoLancamento() As Char
            Get
                Return _TipoLancamento
            End Get
            Set(ByVal value As Char)
                _TipoLancamento = value
            End Set
        End Property

        Public Property ValorOficial() As Decimal
            Get
                Return _ValorOficial
            End Get
            Set(ByVal value As Decimal)
                _ValorOficial = value
            End Set
        End Property

        Public Property ValorMoeda() As Decimal
            Get
                Return _ValorMoeda
            End Get
            Set(ByVal value As Decimal)
                _ValorMoeda = value
            End Set
        End Property

        'Segue a moeda do TituloBaixa
        Public Property Valor() As Decimal
            Get
                If TituloBaixa.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Return _ValorOficial
                Else
                    Return _ValorMoeda
                End If
            End Get
            Set(value As Decimal)
                If TituloBaixa.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    _ValorOficial = value
                    _ValorMoeda = _ValorOficial / _IndiceBaixa
                Else
                    _ValorMoeda = value
                    _ValorOficial = _ValorMoeda * _IndiceBaixa
                End If
            End Set
        End Property

#End Region

#Region "Methods"
        Public Function Salvar() As Boolean
            If IUD = Nothing Then Return True
            Dim Banco As New AcessaBanco
            Dim Sqls As New ArrayList

            Sqls.Clear()
            SalvarSql(Sqls)
            Return Banco.GravaBanco(Sqls)
        End Function

        Public Sub SalvarSql(ByRef Sqls As ArrayList)
            Dim Sql As String = ""
            Select Case Me.IUD
                Case "I"
                    Sql = " INSERT AdiantamentosXBaixas (Titulo_Id, Sequencia_Id, TituloBaixa, Movimento, Lancamento, TipoLancamento, ValorOficial, ValorMoeda)" & vbCrLf & _
                          " values (" & Me.Adiantamento.CodigoTitulo & "," & Me.Sequencia & "," & Me.CodigoTituloBaixa & ",'" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Me.Lancamento & "','" & Me.TipoLancamento & "'," & Str(Math.Round(Me.ValorOficial, 2)) & "," & Str(Math.Round(Me.ValorMoeda, 2)) & ")"
                    Sqls.Add(Sql)
                Case "U"
                    Sql = " UPDATE AdiantamentosXBaixas SET" & vbCrLf & _
                          "    Movimento       ='" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                          "   ,Lancamento      ='" & Me.Lancamento & "'" & vbCrLf & _
                          "   ,TipoLancamento  ='" & Me.TipoLancamento & "'" & vbCrLf & _
                          "   ,ValorOficial    = " & Str(Math.Round(_ValorOficial, 2)) & vbCrLf & _
                          "   ,ValorMoeda      = " & Str(Math.Round(_ValorMoeda, 2)) & vbCrLf & _
                          "  WHERE Titulo_Id       = " & Me.Adiantamento.CodigoTitulo & vbCrLf & _
                          "    And Sequencia_Id    = " & Me.Sequencia
                    Sqls.Add(Sql)
                Case "D", "C"
                    Sql = " DELETE AdiantamentosXBaixas" & vbCrLf & _
                          "  WHERE Titulo_Id       = " & Me.Adiantamento.CodigoTitulo & vbCrLf & _
                          "    And Sequencia_Id    = " & Me.Sequencia
                    Sqls.Add(Sql)
            End Select
        End Sub
#End Region

    End Class
End Namespace

