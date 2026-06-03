Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

Namespace Novo
    <Serializable()> _
    Public Class ListAdiantamentoNovo
        Inherits List(Of AdiantamentoNovo)

#Region "Contrutor"
        Public Sub New(Parametros As Hashtable)
            'CodigoCliente     string obg
            'EndCliente        int    obg
            'ConsolidarCliente bool   False
            'SomenteComSaldo   bool   False
            'CodigoEmpresa     string 
            'EndEmpresa        int
            'ConsolidarEmpresa bool   False

            'CodigoPedido      int    Ordena para que os adiantamentos do proprio pedido sejam liquidados primeiro
            'ContaAdiantamento int    obg, Conta de Adiantamento definida na subOperacao

            'isTroca           bool   False

            'CodigoEmpresaTroca     string obg 
            'EndEmpresaTroca        int obg
            'CodigoPedidoTroca      int obg


            Dim Troca As Boolean = False
            If Parametros.ContainsKey("isTroca") Then Troca = Parametros("isTroca")

            Dim Sql As String

            If Troca Then
                Sql = "Select case when T.Pedido in (" & Parametros("CodigoPedido") & "," & Parametros("CodigoPedidoTroca") & ") then 1 else 0 end as DoPedido,"
            ElseIf Parametros.ContainsKey("CodigoPedido") Then
                Sql = "Select case when T.Pedido = " & Parametros("CodigoPedido") & " then 1 else 0 end as DoPedido,"
            Else
                Sql = "Select 0 as DoPedido,"
            End If

            Sql &= "       A.Titulo_Id, " & vbCrLf & _
                   "       A.Vencimento," & vbCrLf & _
                   "	      A.Taxa," & vbCrLf & _
                   "       A.Atualizacao" & vbCrLf & _
                   "  FROM Adiantamentos A " & vbCrLf & _
                   " Inner Join Titulos T" & vbCrLf & _
                   "    on A.Titulo_Id = T.Titulo_Id" & vbCrLf & _
                   "    and T.situacao = 1 " & vbCrLf & _
                   " Inner Join TitulosXContaContabil TC" & vbCrLf & _
                   "    on T.Titulo_Id           = TC.Titulo_id" & vbCrLf & _
                   "   and T.ContaContabilCliFor = TC.Conta_Id" & vbCrLf & _
                   "   and Case RecPag" & vbCrLf & _
                   "         When 'P' then 'D'" & vbCrLf & _
                   "         When 'R' then 'C'" & vbCrLf & _
                   "         When 'C' then 'D'" & vbCrLf & _
                   "       end                   = TC.DC_Id" & vbCrLf & _
                   "  LEFT JOIN (" & vbCrLf & _
                   "            SELECT Titulo_Id," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'B'                          then ValorOficial else 0 end) as BaixaOficial," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'J' and TipoLancamento = 'R' then ValorOficial else 0 end) as JuroOficialRecebido," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'J' and TipoLancamento = 'P' then ValorOficial else 0 end) as JuroOficialPago," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'V' and TipoLancamento = 'A' then ValorOficial else 0 end) as VariacaoOficialAtiva," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'V' and TipoLancamento = 'P' then ValorOficial else 0 end) as VariacaoOficialPassiva," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'B'                          then ValorMoeda   else 0 end) as BaixaMoeda," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'J' and TipoLancamento = 'R' then ValorMoeda   else 0 end) as JuroMoedaRecebido," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'J' and TipoLancamento = 'P' then ValorMoeda   else 0 end) as JuroMoedaPago," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'V' and TipoLancamento = 'A' then ValorMoeda   else 0 end) as VariacaoMoedaAtiva," & vbCrLf & _
                   "                   Sum(Case When Lancamento = 'V' and TipoLancamento = 'P' then ValorMoeda   else 0 end) as VariacaoMoedaPassiva" & vbCrLf & _
                   "                FROM AdiantamentosXBaixas AS AxB" & vbCrLf & _
                   "               group by Titulo_Id" & vbCrLf & _
                   "      ) sb" & vbCrLf & _
                   "    ON A.Titulo_Id = sb.Titulo_Id" & vbCrLf & _
                   " Inner join Moedas M" & vbCrLf & _
                   "    on M.Moeda_Id = T.Moeda" & vbCrLf & _
                   "  WHERE T.Situacao = 1 " & vbCrLf

            If Not String.IsNullOrWhiteSpace(Parametros("ContaAdiantamento")) Then
                Sql &= "   And T.ContaContabilCliFor = " & Parametros("ContaAdiantamento") & vbCrLf
            End If

            If Parametros("SomenteComSaldo") Then
                Sql &= "   and case" & vbCrLf & _
                       "          when M.Classificacao = 'O'" & vbCrLf & _
                       "            then (TC.ValorOficial + isnull(sb.JuroOficialRecebido,0) + isnull(sb.JurooficialPago,0)) - isnull(sb.Baixaoficial,0)" & vbCrLf & _
                       "            else (TC.ValorMoeda   + isnull(sb.JuroMoedaRecebido,0)   + isnull(sb.JuroMoedaPago,0))   - isnull(sb.BaixaMoeda,0)" & vbCrLf & _
                       "        end > 0" & vbCrLf
            End If

            If Troca Then
                Sql &= "  and (" & vbCrLf & _
                       "       (    T.Empresa    ='" & Parametros("CodigoEmpresa") & "'" & vbCrLf & _
                       "        and T.EndEmpresa = " & Parametros("EndEmpresa") & vbCrLf & _
                       "        and T.Pedido     = " & Parametros("CodigoPedido") & vbCrLf & _
                       "       )" & vbCrLf & _
                       "       or " & vbCrLf & _
                       "       (" & vbCrLf & _
                       "            T.Empresa    ='" & Parametros("CodigoEmpresaTroca") & "'" & vbCrLf & _
                       "        and T.EndEmpresa = " & Parametros("EndEmpresaTroca") & vbCrLf & _
                       "        and T.Pedido     = " & Parametros("CodigoPedidoTroca") & vbCrLf & _
                       "       )" & vbCrLf & _
                       "      )"
            Else
                If Parametros.ContainsKey("CodigoEmpresa") Then
                    If Parametros("ConsolidarEmpresa") Then
                        Sql &= "    and left(T.Empresa,8) ='" & Parametros("CodigoEmpresa").Substring(0, 8) & "'" & vbCrLf
                    Else
                        Sql &= "    and T.Empresa    ='" & Parametros("CodigoEmpresa") & "'" & vbCrLf & _
                               "    and T.endEmpresa = " & Parametros("EndEmpresa") & vbCrLf
                    End If
                End If

                If Parametros("ConsolidarCliente") Then
                    If Parametros("CodigoCliente").ToString.Length < 8 Then
                        Sql &= "    and T.clifor ='" & Parametros("CodigoCliente") & "'" & vbCrLf
                    Else
                        Sql &= "    and left(T.clifor,8) ='" & Parametros("CodigoCliente").Substring(0, 8) & "'" & vbCrLf
                    End If
                Else
                    Sql &= "    and T.clifor         ='" & Parametros("CodigoCliente") & "'" & vbCrLf & _
                           "    and T.enderecoclifor = " & Parametros("EndCliente") & vbCrLf
                End If

                If Parametros.ContainsKey("CodigoPedido") Then
                    Sql &= "        and isnull(T.Pedido ,0) = 0 or T.Pedido = " & Parametros("CodigoPedido") & vbCrLf
                End If
            End If

            If Parametros("SomenteComSaldo") Then
                Sql &= "   and case" & vbCrLf & _
                       "          when M.Classificacao = 'O'" & vbCrLf & _
                       "            then (TC.ValorOficial + isnull(sb.JuroOficialRecebido,0) + isnull(sb.JurooficialPago,0)) - isnull(sb.Baixaoficial,0)" & vbCrLf & _
                       "            else (TC.ValorMoeda   + isnull(sb.JuroMoedaRecebido,0)   + isnull(sb.JuroMoedaPago,0))   - isnull(sb.BaixaMoeda,0)" & vbCrLf & _
                       "        end > 0" & vbCrLf
            End If

            If Troca Then
                Sql &= " order by case when T.Pedido in (" & Parametros("CodigoPedido") & "," & Parametros("CodigoPedidoTroca") & ") then 1 else 0 end, A.Vencimento"
            ElseIf Parametros.ContainsKey("CodigoPedido") Then
                Sql &= " order by case when T.pedido = " & Parametros("CodigoPedido") & " then 1 else 2 end, A.Vencimento"
            Else
                Sql &= " order by A.Vencimento"
            End If

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Adiantamentos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim AD As New AdiantamentoNovo
                AD.DoPedido = row("DoPedido")
                AD.CodigoTitulo = row("Titulo_ID")
                AD.Vencimento = row("Vencimento")
                AD.Taxa = row("Taxa")
                AD.Atualizacao = row("Atualizacao")
                Me.Add(AD)
            Next

        End Sub
#End Region

#Region "Fields"
        Private _Titulo As Novo.TituloNovo
#End Region

#Region "Property"
        Public ReadOnly Property ValorTotalDisponivelParaBaixa
            Get
                If Titulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Return (From x In Me
                            Select x.SaldoValorOficial).Sum
                Else
                    Return (From x In Me
                            Select x.SaldoValorMoeda).Sum
                End If
            End Get
        End Property

        Public ReadOnly Property ValorTotalInformadoParaBaixa
            Get
                Return (From x In Me
                        Select x.VlrBaixa).Sum
            End Get
        End Property

        Public Property Titulo() As Novo.TituloNovo
            Get
                Return _Titulo
            End Get
            Set(ByVal value As Novo.TituloNovo)
                _Titulo = value
            End Set
        End Property
#End Region

#Region "Methods"
        Public Sub DistribuirValorParaBaixarAdiantamentos(Valor As Decimal, Moeda As eTiposMoeda, Optional distribuirValor As Boolean = False)
            Dim VlrADistribuir As Decimal = Valor
            If VlrADistribuir = 0 Then Exit Sub
            For Each row As Novo.AdiantamentoNovo In Me
                If row.DoPedido Or distribuirValor Then
                    If Moeda = eTiposMoeda.Oficial Then
                        If Moeda = eTiposMoeda.Oficial Then
                            If VlrADistribuir > row.SaldoValorOficial Then
                                row.VlrBaixa = row.SaldoValorOficial
                                VlrADistribuir -= row.SaldoValorOficial
                            Else
                                row.VlrBaixa = VlrADistribuir
                                VlrADistribuir = 0
                            End If
                        Else
                            If Math.Round(VlrADistribuir / row.Titulo.IndiceTitulo, 2) > row.SaldoValorMoeda Then
                                row.VlrBaixa = row.SaldoValorMoeda
                                VlrADistribuir -= Math.Round(row.SaldoValorMoeda / row.Titulo.IndiceTitulo, 2)
                            Else
                                row.VlrBaixa = Math.Round(VlrADistribuir / row.Titulo.IndiceTitulo, 2)
                                VlrADistribuir = 0
                            End If
                        End If

                    Else

                        If row.Titulo.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                            If VlrADistribuir > row.SaldoValorMoeda Then
                                row.VlrBaixa = row.SaldoValorMoeda
                                VlrADistribuir -= row.SaldoValorMoeda
                            Else
                                row.VlrBaixa = VlrADistribuir
                                VlrADistribuir = 0
                            End If
                        Else
                            If Math.Round(VlrADistribuir * row.Titulo.IndiceTitulo, 2) > row.SaldoValorOficial Then
                                row.VlrBaixa = row.SaldoValorOficial
                                VlrADistribuir -= Math.Round(row.SaldoValorOficial * row.Titulo.IndiceTitulo, 2)
                            Else
                                row.VlrBaixa = Math.Round(VlrADistribuir * row.Titulo.IndiceTitulo, 2)
                                VlrADistribuir = 0
                            End If
                        End If
                    End If
                End If
            Next
        End Sub

        Public Sub DistribuirValorParaBaixaAdiantamento(Valor As Decimal, Moeda As eTiposMoeda, row As Novo.AdiantamentoNovo)
            Dim VlrADistribuir As Decimal = Valor
            If VlrADistribuir = 0 Then Exit Sub
            If Moeda = eTiposMoeda.Oficial Then
                If Moeda = eTiposMoeda.Oficial Then
                    If VlrADistribuir > row.SaldoValorOficial Then
                        row.VlrBaixa = row.SaldoValorOficial
                        VlrADistribuir -= row.SaldoValorOficial
                    Else
                        row.VlrBaixa = VlrADistribuir
                        VlrADistribuir = 0
                    End If
                Else
                    If Math.Round(VlrADistribuir / row.Titulo.IndiceTitulo, 2) > row.SaldoValorMoeda Then
                        row.VlrBaixa = row.SaldoValorMoeda
                        VlrADistribuir -= Math.Round(row.SaldoValorMoeda / row.Titulo.IndiceTitulo, 2)
                    Else
                        row.VlrBaixa = Math.Round(VlrADistribuir / row.Titulo.IndiceTitulo, 2)
                        VlrADistribuir = 0
                    End If
                End If

            Else
                If row.Titulo.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                    If VlrADistribuir > row.SaldoValorMoeda Then
                        row.VlrBaixa = row.SaldoValorMoeda
                        VlrADistribuir -= row.SaldoValorMoeda
                    Else
                        row.VlrBaixa = VlrADistribuir
                        VlrADistribuir = 0
                    End If
                Else
                    If Math.Round(VlrADistribuir * row.Titulo.IndiceTitulo, 2) > row.SaldoValorOficial Then
                        row.VlrBaixa = row.SaldoValorOficial
                        VlrADistribuir -= Math.Round(row.SaldoValorOficial * row.Titulo.IndiceTitulo, 2)
                    Else
                        row.VlrBaixa = Math.Round(VlrADistribuir * row.Titulo.IndiceTitulo, 2)
                        VlrADistribuir = 0
                    End If
                End If
            End If

        End Sub


        Public Function AtualizarJuroEVariacaoDosAdiantamentos() As Boolean
            For Each row In Me
                If Not row.AtualizarJuroEVariacao() Then Return False
            Next
            Return True
        End Function
#End Region
    End Class

    <Serializable()> _
    Public Class AdiantamentoNovo
#Region "Contrutor"
        Public Sub New()

        End Sub

        Public Sub New(ByVal pTitulo As Novo.TituloNovo)
            _Titulo = pTitulo

            Dim Sql As String
            Sql = " Select A.Titulo_Id," & vbCrLf & _
                  "        A.Vencimento," & vbCrLf & _
                  "	       A.Taxa," & vbCrLf & _
                  "        A.Atualizacao " & vbCrLf & _
                  "   FROM Adiantamentos AS A " & vbCrLf & _
                  "   INNER JOIN Titulos AS T " & vbCrLf & _
                  "         ON A.Titulo_Id = T.Titulo_Id " & vbCrLf & _
                  "  Where T.Situacao = 1 And A.Titulo_Id = " & pTitulo.Codigo

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Adiantamentos")

            If ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)

            CodigoTitulo = row("Titulo_Id")
            Vencimento = row("Vencimento")
            Taxa = row("Taxa")
            Atualizacao = row("Atualizacao")
        End Sub
#End Region

#Region "Fields"
        Private _IUD As String

        Private _DoPedido As String

        Private _DescMoeda As String = ""
        Private _CodigoTitulo As Integer
        Private _Titulo As Novo.TituloNovo

        Private _Vencimento As DateTime
        Private _Taxa As Decimal
        Private _Atualizacao As DateTime

        Private _VlrBaixa As Decimal

        Private _Baixas As Novo.ListAdiantamentoBaixaNovo
#End Region

#Region "Property"
        Public Property DoPedido() As String
            Get
                Return _DoPedido
            End Get
            Set(ByVal value As String)
                _DoPedido = value
            End Set
        End Property

        Public Property IUD() As String
            Get
                Return _IUD
            End Get
            Set(ByVal value As String)
                _IUD = value
            End Set
        End Property

        Public ReadOnly Property DescMoeda() As String
            Get
                If _DescMoeda.Length = 0 And _CodigoTitulo > 0 Then _DescMoeda = Titulo.Moeda.Descricao
                Return _DescMoeda
            End Get
        End Property

        Public Property CodigoTitulo() As Integer
            Get
                Return _CodigoTitulo
            End Get
            Set(ByVal value As Integer)
                _CodigoTitulo = value
                _DescMoeda = ""
            End Set
        End Property

        Public ReadOnly Property Titulo() As Novo.TituloNovo
            Get
                If _Titulo Is Nothing And _CodigoTitulo > 0 Then _Titulo = New Novo.TituloNovo(_CodigoTitulo)
                Return _Titulo
            End Get
        End Property

        Public ReadOnly Property Movimento() As DateTime
            Get
                Return Titulo.Movimento
            End Get
        End Property

        Public Property Vencimento() As DateTime
            Get
                Return _Vencimento
            End Get
            Set(ByVal value As DateTime)
                _Vencimento = value
            End Set
        End Property

        Public Property Taxa() As Decimal
            Get
                Return _Taxa
            End Get
            Set(ByVal value As Decimal)
                _Taxa = value
            End Set
        End Property

        Public Property Atualizacao() As DateTime
            Get
                Return _Atualizacao
            End Get
            Set(ByVal value As DateTime)
                _Atualizacao = value
            End Set
        End Property

        '*************  VALORES ****************************
        Public ReadOnly Property ValorOficial() As Decimal
            Get
                Return Titulo.Valores.EncargoValorLiquido.ValorOficial
            End Get
        End Property

        Public ReadOnly Property ValorMoeda() As Decimal
            Get
                Return Titulo.Valores.EncargoValorLiquido.ValorMoeda
            End Get
        End Property

        '**********  BAIXAS VALORES ************************
        Public Property Baixas() As Novo.ListAdiantamentoBaixaNovo
            Get
                If _Baixas Is Nothing Then _Baixas = New Novo.ListAdiantamentoBaixaNovo(Me)
                Return _Baixas
            End Get
            Set(ByVal value As Novo.ListAdiantamentoBaixaNovo)
                _Baixas = value
            End Set
        End Property

        Public ReadOnly Property TotalOficialBaixado() As Decimal
            Get
                Return Baixas.TotalOficialBaixado
            End Get
        End Property

        Public ReadOnly Property TotalMoedaBaixado() As Decimal
            Get
                Return Baixas.TotalMoedaBaixado
            End Get
        End Property

        '************** JUROS ******************************
        Public ReadOnly Property TotalOficialJuroRecebido() As Decimal
            Get
                Return Baixas.TotalOficialJuroRecebido
            End Get
        End Property

        Public ReadOnly Property TotalMoedaJuroRecebido() As Decimal
            Get
                Return Baixas.TotalMoedaJuroRecebido
            End Get
        End Property

        Public ReadOnly Property TotalOficialJuroPago() As Decimal
            Get
                Return Baixas.TotalOficialJuroPago
            End Get
        End Property

        Public ReadOnly Property TotalMoedaJuroPago() As Decimal
            Get
                Return Baixas.TotalMoedaJuroPago
            End Get
        End Property

        '***********  VARIACAO ***************************
        Public ReadOnly Property TotalOficialVariacaoAtiva() As Decimal
            Get
                Return Baixas.TotalOficialVariacaoAtiva
            End Get
        End Property

        Public ReadOnly Property TotalMoedaVariacaoAtiva() As Decimal
            Get
                Return Baixas.TotalMoedaVariacaoAtiva
            End Get
        End Property

        Public ReadOnly Property TotalOficialVariacaoPassiva() As Decimal
            Get
                Return Baixas.TotalOficialVariacaoPassiva
            End Get
        End Property

        Public ReadOnly Property TotalMoedaVariacaoPassiva() As Decimal
            Get
                Return Baixas.TotalMoedaVariacaoPassiva
            End Get
        End Property

        '*********  SALDO VALORES ************************
        Public ReadOnly Property SaldoValorOficial As Decimal
            Get
                Return ValorOficial + TotalOficialJuroRecebido + TotalOficialJuroPago - TotalOficialBaixado
            End Get
        End Property

        Public ReadOnly Property SaldoValorMoeda As Decimal
            Get
                Return ValorMoeda + TotalMoedaJuroRecebido + TotalMoedaJuroPago - TotalMoedaBaixado
            End Get
        End Property

        '*********  VALORES A SEREM BAIXADOS **********************
        Public Property VlrBaixa() As Decimal
            Get
                Return _VlrBaixa
            End Get
            Set(ByVal value As Decimal)
                _VlrBaixa = value
            End Set
        End Property


        '******************* Segue a Moeda  ***********************
        Public ReadOnly Property Valor() As Decimal
            Get
                If Titulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Return Titulo.Valores.EncargoValorLiquido.ValorOficial
                Else
                    Return Titulo.Valores.EncargoValorLiquido.ValorMoeda
                End If
            End Get
        End Property

        Public ReadOnly Property TotalBaixado() As Decimal
            Get
                If Titulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Return Baixas.TotalOficialBaixado
                Else
                    Return Baixas.TotalMoedaBaixado
                End If
            End Get
        End Property

        Public ReadOnly Property SaldoValor As Decimal
            Get
                If Titulo.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    Return ValorOficial + TotalOficialJuroRecebido + TotalOficialJuroPago - TotalOficialBaixado
                Else
                    Return ValorMoeda + TotalMoedaJuroRecebido + TotalMoedaJuroPago - TotalMoedaBaixado
                End If
            End Get
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
                    Sql = "Insert Into Adiantamentos(Titulo_Id, Vencimento, Taxa, Atualizacao)" & vbCrLf & _
                          " values (" & Me.CodigoTitulo & ",'" & Me.Vencimento.ToString("yyyy-MM-dd") & "'," & Str(Math.Round(Me.Taxa, 2)) & ",'" & Me.Atualizacao.ToString("yyyy-MM-dd") & "')" & vbCrLf
                    Sqls.Add(Sql)
                Case "U"
                    Sql = " UPDATE Adiantamentos SET" & vbCrLf & _
                          "    Vencimento   ='" & Me.Vencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                          "   ,Taxa         = " & Str(Math.Round(Me.Taxa, 2)) & vbCrLf & _
                          "   ,Atualizacao  ='" & Me.Atualizacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                          " WHERE Titulo_Id = " & Me.CodigoTitulo
                    Sqls.Add(Sql)
                Case "D"
                    Sql = " DELETE Adiantamentos " & vbCrLf & _
                          "  WHERE Titulo_Id = " & _CodigoTitulo
                    Sqls.Add(Sql)
            End Select
        End Sub
#End Region

#Region "Atualizacao dos Juros e Variacoes"
        Public Function AtualizarJuroEVariacao() As Boolean
            If Atualizacao = Date.Now.Date Then Return True

            Dim sqls As New ArrayList
            CalculaVariacao(sqls)
            CalculaJuro(sqls)

            Me.IUD = "U"
            Me.Atualizacao = Date.Now
            Me.SalvarSql(sqls)

            Dim Banco As New AcessaBanco

            If Banco.GravaBanco(sqls) Then
                Return True
            Else
                Me.Baixas = Nothing
                Return False
            End If
        End Function

        Private Sub CalculaJuro(ByRef sqls As ArrayList)
            If Date.Now < Vencimento Or Taxa = 0 Then Exit Sub

            Dim DiasCorridos As Integer = Now.Subtract(Atualizacao).Days

            Dim TxJuro As Decimal = (Taxa / 30) * DiasCorridos

            Dim vlrOficial As Decimal
            Dim vlrMoeda As Decimal

            If Titulo.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                vlrMoeda = Math.Round((SaldoValorMoeda * TxJuro) / 100, 2)
                vlrOficial = Math.Round(vlrMoeda * New Cotacao(Titulo.CodigoIndexador, Date.Now).Indice, 2)
            Else
                Dim BxVar As New AdiantamentoBaixaNovo(Me)
                vlrOficial = Math.Round((SaldoValorOficial * TxJuro) / 100, 2)
                vlrMoeda = Math.Round(vlrOficial / New Cotacao(Titulo.CodigoIndexador, Date.Now).Indice, 2)
            End If

            Dim Historico As String = "Atualizacao do Juro calculdado da data de ultima atualizacao de " & Me.Atualizacao.ToString("dd/MM/yyyy") & " ate " & Date.Now.ToString("dd/MM/yyyy") & " do Adiantamento Numero " & Me.CodigoTitulo

            If Titulo.ReceberPagar = "R" Then
                Dim BxVar As New AdiantamentoBaixaNovo(Me)
                BxVar.IUD = "I"
                BxVar.CodigoTituloBaixa = CodigoTitulo
                BxVar.Movimento = Date.Now
                BxVar.Lancamento = "J"
                BxVar.TipoLancamento = "P"
                BxVar.ValorOficial = vlrOficial
                BxVar.ValorMoeda = vlrMoeda
                BxVar.Sequencia = Me.Baixas.ProximaSequenciaLivre
                BxVar.SalvarSql(sqls)
                Me.Baixas.Add(BxVar)

                GeraSqlRazao("D", "J", "P", vlrOficial, vlrMoeda, 24, Historico, sqls)
                GeraSqlRazao("C", "A", "", vlrOficial, vlrMoeda, 24, Historico, sqls)
            Else
                Dim BxVar As New AdiantamentoBaixaNovo(Me)
                BxVar.IUD = "I"
                BxVar.CodigoTituloBaixa = CodigoTitulo
                BxVar.Movimento = Date.Now
                BxVar.Lancamento = "J"
                BxVar.TipoLancamento = "R"
                BxVar.ValorOficial = vlrOficial
                BxVar.ValorMoeda = vlrMoeda
                BxVar.Sequencia = Me.Baixas.ProximaSequenciaLivre
                BxVar.SalvarSql(sqls)
                Me.Baixas.Add(BxVar)

                GeraSqlRazao("C", "J", "R", vlrOficial, vlrMoeda, 24, Historico, sqls)
                GeraSqlRazao("D", "A", "", vlrOficial, vlrMoeda, 24, Historico, sqls)
            End If
        End Sub

        Private Sub CalculaVariacao(ByRef sqls As ArrayList)
            If Titulo.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                If SaldoValorMoeda = 0 Then Exit Sub
            Else
                Exit Sub
            End If

            Dim Reaishj As Decimal = Math.Round(SaldoValorMoeda * New Cotacao(2, Date.Now).Indice, 2)
            Dim VariacaoOficial As Decimal = Reaishj - SaldoValorOficial
            If VariacaoOficial = 0 Then Exit Sub

            'O Dolar é Zero só alimenta os reais
            Dim Historico As String = "Atualizacao da Variacao Monetaria calculdado da data de ultima atualizacao de " & Me.Atualizacao.ToString("dd/MM/yyyy") & " ate " & Date.Now.ToString("dd/MM/yyyy") & " do Adiantamento Numero " & Me.CodigoTitulo

            If (Titulo.ReceberPagar = "R" And VariacaoOficial < 0) Or (Titulo.ReceberPagar = "P" And VariacaoOficial > 0) Then
                Dim BxVar As New AdiantamentoBaixaNovo(Me)
                BxVar.IUD = "I"
                BxVar.CodigoTituloBaixa = CodigoTitulo
                BxVar.Movimento = Date.Now
                BxVar.Lancamento = "V"
                BxVar.TipoLancamento = "A"
                BxVar.ValorOficial = VariacaoOficial
                BxVar.ValorMoeda = 0
                BxVar.Sequencia = Me.Baixas.ProximaSequenciaLivre
                BxVar.SalvarSql(sqls)
                Me.Baixas.Add(BxVar)

                GeraSqlRazao("C", "V", "A", VariacaoOficial, 0, 25, Historico, sqls)
                GeraSqlRazao("D", "A", "", VariacaoOficial, 0, 25, Historico, sqls)

            ElseIf (Titulo.ReceberPagar = "R" And VariacaoOficial > 0) Or (Titulo.ReceberPagar = "P" And VariacaoOficial < 0) Then
                Dim BxVar As New AdiantamentoBaixaNovo(Me)
                BxVar.IUD = "I"
                BxVar.CodigoTituloBaixa = CodigoTitulo
                BxVar.Movimento = Date.Now
                BxVar.Lancamento = "V"
                BxVar.TipoLancamento = "P"
                BxVar.ValorOficial = VariacaoOficial
                BxVar.ValorMoeda = 0
                BxVar.Sequencia = Me.Baixas.ProximaSequenciaLivre
                BxVar.SalvarSql(sqls)
                Me.Baixas.Add(BxVar)

                GeraSqlRazao("D", "V", "P", VariacaoOficial, 0, 25, Historico, sqls)
                GeraSqlRazao("C", "A", "", VariacaoOficial, 0, 25, Historico, sqls)
            End If
        End Sub

        Private Sub GeraSqlRazao(DebitoCredito As String, JuroVariacaoAdiantamento As String, AtivoPassivoPagoRecebido As String, ValorOficial As Decimal, ValorMoeda As Decimal, Lote As Integer, Historico As String, ByRef sqls As ArrayList)
            Dim Lan As New Razao()
            Lan.IUD = "I"
            Lan.CodigoEmpresa = Titulo.CodigoEmpresa
            Lan.EndEmpresa = Titulo.EnderecoEmpresa

            If JuroVariacaoAdiantamento = "V" Then
                If AtivoPassivoPagoRecebido = "A" Then
                    Lan.CodigoConta = Titulo.Empresa.Empresa.CodigoContaVariacaoMonetariaAtiva
                Else
                    Lan.CodigoConta = Titulo.Empresa.Empresa.CodigoContaVariacaoMonetariaPassiva
                End If
            ElseIf JuroVariacaoAdiantamento = "A" Then
                Lan.CodigoConta = Titulo.CodigoContaContabilCliFor
            ElseIf JuroVariacaoAdiantamento = "J" Then
                If AtivoPassivoPagoRecebido = "P" Then
                    Lan.CodigoConta = Titulo.Empresa.Empresa.CodigoContaJuroPago
                Else
                    Lan.CodigoConta = Titulo.Empresa.Empresa.CodigoContaJuroRecebido
                End If
            End If

            If Lan.Conta.TemCliente Then
                Lan.CodigoCliente = Titulo.CodigoCliFor
                Lan.EnderecoCliente = Titulo.EnderecoCliFor
            End If

            Lan.Lote = Lote
            Lan.Sequencia = Titulo.Codigo
            Lan.CodigoUnidadeDeNegocio = Titulo.CodigoUnidadeDeNegocio
            Lan.CodigoTitulo = Titulo.Codigo
            Lan.CodigoPedido = Titulo.CodigoPedido
            Lan.CodigoIndexador = 2
            Lan.DataMoeda = Date.Now
            Lan.Movimento = Date.Now
            Lan.DataBaixa = Date.Now
            Lan.DebitoMoeda = 0
            Lan.CreditoMoeda = 0

            If DebitoCredito = "D" Then
                Lan.DebitoOficial = ValorOficial
                Lan.CreditoOficial = 0
                Lan.DebitoMoeda = ValorMoeda
                Lan.CreditoMoeda = 0
            Else
                Lan.DebitoOficial = 0
                Lan.CreditoOficial = ValorOficial
                Lan.DebitoMoeda = 0
                Lan.CreditoMoeda = ValorMoeda
            End If

            Lan.Historico = Historico
            Lan.PrevistoRealizado = "R"
            Lan.Processo = "ATUALIZACAO DE JUROS E VARIACAO"
            Lan.CodigoSituacao = 1
            sqls.Add(Lan.SalvarSql())
        End Sub
#End Region
    End Class

End Namespace