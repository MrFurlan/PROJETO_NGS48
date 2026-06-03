Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

'************************************************************************************************************************************************************
'***********************************************      LISTA DE ADIANTAMENTOS     ****************************************************************************
'************************************************************************************************************************************************************
<Serializable()> _
Public Class ListAdiantamento
    Inherits List(Of Adiantamento)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(pEmpresa As Cliente, pcliente As Cliente, Optional ptipo As Integer = 0, Optional Moeda As eTiposMoeda = eTiposMoeda.Todas, Optional pCodigoPedido As Integer = 0, Optional pConta As String = "")
        'ptipo
        '0 - Todos
        '1 - Com Saldo
        '2 - Sem Saldo

        Dim sql As String
        sql = "SELECT Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id, Moeda, isnull(RegistroPedido,0) as RegistroPedido, Titulo, isnull(Safra,'NENHUMA') Safra, Vencimento, Baixa as Movimento, ValorOficial, ValorMoeda," & vbCrLf & _
              "       Taxa, Periodicidade, isnull(UltimaAtualizacao,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UltimaAtualizacao, " & vbCrLf & _
              "       isnull(Informacoes,'') AS Informacoes, isnull(Observacoes,'') AS Observacoes, isnull(DataDeCalculo,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS DataDeCalculo, " & vbCrLf & _
              "       Cifrao, Conta_Id, Conta, Adiantamento, Variacao, Baixas, BaixaOficial, BaixaMoeda, Saldo, SaldoOficial, SaldoMoeda," & vbCrLf & _
              "       Case When Conta_Id ='" & pConta & "' then 1 else 0 end Elegivel" & vbCrLf & _
              "  FROM VW_Adiantamento " & vbCrLf & _
              " Where Empresa_ID    ='" & pEmpresa.Codigo & "'" & vbCrLf & _
              "   and EndEmpresa_ID = " & pEmpresa.CodigoEndereco & vbCrLf & _
              "   and Cliente_ID    ='" & pcliente.Codigo & "'" & vbCrLf & _
              "   and EndCliente_ID = " & pcliente.CodigoEndereco

        'If HttpContext.Current.Session("ssNomeUsuario") = "FURLAN" Then
        '    sql &= "   and Adiantamento_Id = 1265"
        'End If

        Select Case ptipo
            Case 1 : sql &= "   and saldo > 0"
            Case 2 : sql &= "   and saldo = 0"
        End Select

        If Moeda = eTiposMoeda.MoedaEstrangeira Then
            sql &= "   and Classificacao = 'M'"
        ElseIf Moeda = eTiposMoeda.Oficial Then
            sql &= "   and Classificacao = 'O'"
        End If


        If pCodigoPedido > 0 Then
            sql &= " Order by Case when isnull(RegistroPedido,0) =" & pCodigoPedido & " then 0 when isnull(RegistroPedido,0) = 0 then 1 else 2 end"
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Adiantamento")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ad As New Adiantamento()
            ad.Codigo = row("Adiantamento_Id")
            ad.CodigoMoeda = row("Moeda")
            ad.CodigoEmpresa = row("Empresa_ID")
            ad.EndEmpresa = row("EndEmpresa_ID")
            ad.CodigoCliente = row("Cliente_ID")
            ad.EndCliente = row("EndCliente_ID")
            ad.RegistroPedido = row("RegistroPedido")
            ad.CodigoTitulo = row("Titulo")
            ad.CodigoSafra = row("Safra")
            ad.Movimento = row("Movimento")
            ad.Vencimento = row("Vencimento")
            ad.ValorOficial = row("ValorOficial")
            ad.ValorMoeda = row("ValorMoeda")
            ad.Taxa = row("Taxa")
            ad.Periodicidade = row("Periodicidade")
            ad.UltimaAtualizacao = row("UltimaAtualizacao")
            ad.Informacoes = row("Informacoes")
            ad.Observacoes = row("Observacoes")
            ad.DataDeCalculo = row("DataDeCalculo")
            ad.Cifrao = row("cifrao")
            ad.CodigoContaAdto = row("Conta_Id")
            ad.DescContaAdto = row("Conta")
            ad.VlrAdto = row("Adiantamento")
            ad.VlrVariacao = row("Variacao")
            ad.VlrBaixa = row("Baixas")
            ad.VlrBaixaOficial = row("BaixaOficial")
            ad.VlrBaixaMoeda = row("BaixaMoeda")
            ad.VlrSaldo = row("Saldo")
            ad.VlrSaldoOficial = row("SaldoOficial")
            ad.VlrSaldoMoeda = row("SaldoMoeda")
            ad.Elegivel = row("Elegivel")

            Me.Add(ad)
        Next
    End Sub

    Public Sub New(ByRef NF As NotaFiscal)
        If NF.Pedido IsNot Nothing AndAlso NF.Pedido.Troca Then
            If Not NF.Pedido.PedidoTroca Is Nothing Then
                Dim Sql As String
                Sql = " select A.Empresa_ID,  A.EndEmpresa_ID, A.Cliente_ID, A.EndCliente_ID, A.Adiantamento_Id, A.Moeda,A.RegistroPedido, A.Titulo, A.Safra,  A.Vencimento, A.ValorOficial,  A.ValorMoeda,  " & vbCrLf & _
                      "        A.Taxa, A.Periodicidade, A.UltimaAtualizacao," & vbCrLf & _
                      "        A.Informacoes, A.Observacoes,   A.DataDeCalculo" & vbCrLf & _
                      "        A.ValorOficial - Sum(isnull(AxB.ValorOficial,0)) as SaldoOficial," & vbCrLf & _
                      "        A.ValorMoeda   - Sum(isnull(AxB.ValorMoeda,0))   as SaldoMoeda" & vbCrLf & _
                      "   from VW_Adiantamento A" & vbCrLf & _
                      "   LEFT JOIN VW_AdiantamentosXBaixas AxB" & vbCrLf & _
                      "     ON A.Titulo      = AxB.TituloAdiantamento " & vbCrLf & _
                      "  where A.Empresa_Id    ='" & NF.Pedido.PedidoTroca.CodigoEmpresa & "'" & vbCrLf & _
                      "    and A.EndEmpresa_Id = " & NF.Pedido.PedidoTroca.EnderecoEmpresa & vbCrLf & _
                      "    and A.Pedido_id     = " & NF.Pedido.PedidoTroca.Codigo & vbCrLf & _
                      "  group by A.Empresa_ID,  A.EndEmpresa_ID, A.Cliente_ID, A.EndCliente_ID, A.Adiantamento_Id, A.RegistroPedido, A.Titulo, A.Safra, A.Vencimento, A.ValorOficial,  A.ValorMoeda,  " & vbCrLf & _
                      "           A.Taxa, A.Periodicidade, A.UltimaAtualizacao," & vbCrLf & _
                      "           A.Informacoes, A.Observacoes,   A.DataDeCalculo" & vbCrLf & _
                      " having A.ValorOficial - Sum(isnull(AxB.ValorOficial,0)) > 0" & vbCrLf

                Dim Banco As New AcessaBanco
                Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Adiantamentos")

                For Each row As DataRow In ds.Tables(0).Rows
                    Dim AD As New Adiantamento
                    AD.Codigo = row("Adiantamento_Id")
                    AD.CodigoMoeda = row("Moeda")
                    AD.CodigoEmpresa = row("Empresa_ID")
                    AD.EndEmpresa = row("EndEmpresa_ID")
                    AD.CodigoCliente = row("Cliente_ID")
                    AD.EndCliente = row("EndCliente_ID")
                    AD.RegistroPedido = row("RegistroPedido")
                    AD.CodigoTitulo = row("Titulo")
                    AD.Safra = row("Safra")
                    AD.Vencimento = row("Vencimento")
                    AD.ValorOficial = row("ValorOficial")
                    AD.ValorMoeda = row("ValorMoeda")
                    AD.Taxa = row("Taxa")
                    AD.Periodicidade = row("Periodicidade")
                    AD.UltimaAtualizacao = row("UltimaAtualizacao")
                    AD.Informacoes = row("Informacoes")
                    AD.Observacoes = row("Observacoes")
                    AD.DataDeCalculo = row("DataDeCalculo")
                    Me.Add(AD)
                Next

            End If
        End If


    End Sub

    Public Sub New(ByVal pAdiantento As Adiantamento)
        Dim sql As String
        sql = " SELECT Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id, moeda, RegistroPedido, Titulo, Safra, Vencimento, Baixa as movimento, ValorOficial, ValorMoeda," & vbCrLf & _
              "        Taxa, Periodicidade, UltimaAtualizacao, " & vbCrLf & _
              "        Informacoes, Observacoes, DataDeCalculo " & vbCrLf & _
              "   FROM VW_Adiantamento" & vbCrLf & _
              "  Where Empresa_ID      ='" & pAdiantento.CodigoEmpresa & "'" & vbCrLf & _
              "    and EndEmpresa_ID   = " & pAdiantento.EndEmpresa & vbCrLf & _
              "    and Cliente_ID      ='" & pAdiantento.CodigoCliente & vbCrLf & _
              "    and EndCliente_ID   = " & pAdiantento.EndCliente & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Adiantamentos")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim AD As New Adiantamento
            AD.Codigo = row("Adiantamento_Id")
            AD.CodigoMoeda = row("Moeda")
            AD.CodigoEmpresa = row("Empresa_ID")
            AD.EndEmpresa = row("EndEmpresa_ID")
            AD.CodigoCliente = row("Cliente_ID")
            AD.EndCliente = row("EndCliente_ID")
            AD.RegistroPedido = row("RegistroPedido")
            AD.CodigoTitulo = row("Titulo")
            AD.Safra = row("Safra")
            AD.Vencimento = row("Vencimento")
            AD.Movimento = row("Movimento")
            AD.ValorOficial = row("ValorOficial")
            AD.ValorMoeda = row("ValorMoeda")
            AD.Taxa = row("Taxa")
            AD.Periodicidade = row("Periodicidade")
            AD.UltimaAtualizacao = row("UltimaAtualizacao")
            AD.Informacoes = row("Informacoes")
            AD.Observacoes = row("Observacoes")
            AD.DataDeCalculo = row("DataDeCalculo")
            Me.Add(AD)
        Next
    End Sub
#End Region

End Class


'************************************************************************************************************************************************************
'*******************************************      CLASSE BASE DE ADIANTAMENTOS     **************************************************************************
'************************************************************************************************************************************************************
<Serializable()> _
Public Class Adiantamento
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal NrTituloAdiantamento As Integer, Optional nrAdiantamento As Integer = 0, Optional Empresa As String = "", Optional EndEmpresa As String = "")
        Dim sql As String
        sql = "SELECT Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id, Moeda, isnull(RegistroPedido,0) as RegistroPedido, Titulo, isnull(Safra,'NENHUMA') Safra, Vencimento, Baixa as movimento, ValorOficial, ValorMoeda," & vbCrLf & _
              "       Taxa, Periodicidade, isnull(UltimaAtualizacao,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UltimaAtualizacao, " & vbCrLf & _
              "       isnull(Informacoes,'') AS Informacoes, isnull(Observacoes,'') AS Observacoes, isnull(DataDeCalculo,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS DataDeCalculo, " & vbCrLf & _
              "       Cifrao, Conta_Id, Conta, Adiantamento, Variacao, Baixas, BaixaOficial, BaixaMoeda, Saldo, SaldoOficial, SaldoMoeda" & vbCrLf & _
              "  FROM VW_Adiantamento " & vbCrLf & _
              " Where " & IIf(nrAdiantamento > 0, "Empresa_ID = " & Empresa & " AND EndEmpresa_ID = " & EndEmpresa & " AND Adiantamento_Id= " & nrAdiantamento, "Titulo = " & NrTituloAdiantamento) & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Adiantamento")

        For Each row As DataRow In ds.Tables(0).Rows
            Codigo = row("Adiantamento_Id")
            CodigoMoeda = row("Moeda")

            CodigoEmpresa = row("Empresa_ID")
            EndEmpresa = row("EndEmpresa_ID")
            CodigoCliente = row("Cliente_ID")
            EndCliente = row("EndCliente_ID")
            RegistroPedido = row("RegistroPedido")
            CodigoTitulo = row("Titulo")
            CodigoSafra = row("Safra")
            Movimento = row("Movimento")
            Vencimento = row("Vencimento")
            ValorOficial = row("ValorOficial")
            ValorMoeda = row("ValorMoeda")
            Taxa = row("Taxa")
            Periodicidade = row("Periodicidade")
            UltimaAtualizacao = row("UltimaAtualizacao")
            Informacoes = row("Informacoes")
            Observacoes = row("Observacoes")
            DataDeCalculo = row("DataDeCalculo")

            Cifrao = row("cifrao")
            CodigoContaAdto = row("Conta_Id")
            DescContaAdto = row("Conta")
            VlrAdto = row("Adiantamento")
            VlrVariacao = row("Variacao")
            VlrBaixa = row("Baixas")
            VlrBaixaOficial = row("BaixaOficial")
            VlrBaixaMoeda = row("BaixaMoeda")

            VlrSaldo = row("Saldo")
            VlrSaldoOficial = row("SaldoOficial")
            VlrSaldoMoeda = row("SaldoMoeda")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer

    Private _CodigoMoeda As Integer
    Private _Moeda As Moeda

    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente

    Private _CodigoCliente As String
    Private _EndCliente As Integer
    Private _Cliente As Cliente

    Private _RegistroPedido As Integer
    Private _Pedido As Pedido

    Private _CodigoTitulo As Integer
    Private _Titulo As Titulo

    Private _CodigoSafra As String
    Private _safra As Safra

    Private _Vencimento As DateTime
    Private _Movimento As DateTime

    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal

    Private _Taxa As Decimal
    Private _Periodicidade As Integer
    Private _UltimaAtualizacao As DateTime

    Private _Informacoes As String
    Private _Observacoes As String
    Private _DataDeCalculo As DateTime

    Private _Baixas As ListAdiantamentoBaixa

    '***************************************************************
    '************************  Saldos  *****************************
    '***************************************************************
    Private _Cifrao As String
    Private _CodigoContaAdto As String
    Private _DescContaAdto As String
    Private _VlrAdto As Decimal
    Private _VlrVariacao As Decimal
    Private _VlrBaixa As Decimal
    Private _VlrBaixaOficial As Decimal
    Private _VlrBaixaMoeda As Decimal
    Private _VlrSaldoOficial As Decimal
    Private _VlrSaldoMoeda As Decimal
    Private _VlrSaldo As Decimal

    '***************************************************************
    '***********************  CONTROLE  ****************************
    '***************************************************************
    Private _VlrABaixarOficial As Decimal
    Private _VlrABaixarMoeda As Decimal
    Private _Elegivel As Boolean
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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
            _Baixas = Nothing
        End Set
    End Property

    Public Property CodigoMoeda As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(value As Integer)
            _CodigoMoeda = value
            Me.Moeda = Nothing
        End Set
    End Property

    Public Property Moeda As Moeda
        Get
            If _Moeda Is Nothing AndAlso Me.CodigoMoeda > 0 Then _Moeda = New Moeda(Me.CodigoMoeda)
            Return _Moeda
        End Get
        Set(value As Moeda)
            _Moeda = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            Me.Empresa = Nothing
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            Me.Empresa = Nothing
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            Me.Cliente = Nothing
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
            Me.Cliente = Nothing
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property RegistroPedido() As Integer
        Get
            Return _RegistroPedido
        End Get
        Set(ByVal value As Integer)
            _RegistroPedido = value
            Me.Pedido = Nothing
        End Set
    End Property

    Public Property Pedido() As Pedido
        Get
            If _Pedido Is Nothing And _RegistroPedido > 0 And _CodigoEmpresa.Length > 0 Then _Pedido = New Pedido(_CodigoEmpresa, _EndEmpresa, _RegistroPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property CodigoTitulo() As Integer
        Get
            Return _CodigoTitulo
        End Get
        Set(ByVal value As Integer)
            _CodigoTitulo = value
            Me.Titulo = Nothing
        End Set
    End Property

    Public Property Titulo() As Titulo
        Get
            If _Titulo Is Nothing And Me.CodigoTitulo > 0 Then _Titulo = New Titulo(Me.CodigoTitulo)
            Return _Titulo
        End Get
        Set(ByVal value As Titulo)
            _Titulo = value
        End Set
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property Safra() As Safra
        Get
            Return _safra
        End Get
        Set(ByVal value As Safra)
            _safra = value
        End Set
    End Property

    Public Property Vencimento() As DateTime
        Get
            Return _Vencimento
        End Get
        Set(ByVal value As DateTime)
            _Vencimento = value
        End Set
    End Property

    Public Property Movimento As DateTime
        Get
            Return _Movimento
        End Get
        Set(value As DateTime)
            _Movimento = value
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

    Public Property Taxa() As Decimal
        Get
            Return _Taxa
        End Get
        Set(ByVal value As Decimal)
            _Taxa = value
        End Set
    End Property

    Public Property Periodicidade() As Integer
        Get
            Return _Periodicidade
        End Get
        Set(ByVal value As Integer)
            _Periodicidade = value
        End Set
    End Property

    Public Property UltimaAtualizacao() As DateTime
        Get
            Return _UltimaAtualizacao
        End Get
        Set(ByVal value As DateTime)
            _UltimaAtualizacao = value
        End Set
    End Property

    Public Property Informacoes() As String
        Get
            Return _Informacoes
        End Get
        Set(ByVal value As String)
            _Informacoes = value
        End Set
    End Property

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property DataDeCalculo() As DateTime
        Get
            Return _DataDeCalculo
        End Get
        Set(ByVal value As DateTime)
            _DataDeCalculo = value
        End Set
    End Property

    '***************************************************************
    '************************  Listas  *****************************
    '***************************************************************
    Public Property Baixas() As ListAdiantamentoBaixa
        Get
            If _Baixas Is Nothing Then _Baixas = New ListAdiantamentoBaixa(Me)
            Return _Baixas
        End Get
        Set(ByVal value As ListAdiantamentoBaixa)
            _Baixas = value
        End Set
    End Property

    '***************************************************************
    '************************  Saldos  *****************************
    '***************************************************************
    Public Property Cifrao As String
        Get
            Return _Cifrao
        End Get
        Set(value As String)
            _Cifrao = value
        End Set
    End Property
    Public Property CodigoContaAdto As String
        Get
            Return _CodigoContaAdto
        End Get
        Set(value As String)
            _CodigoContaAdto = value
        End Set
    End Property
    Public Property DescContaAdto As String
        Get
            Return _DescContaAdto
        End Get
        Set(value As String)
            _DescContaAdto = value
        End Set
    End Property
    Public Property VlrAdto As Decimal
        Get
            Return _VlrAdto
        End Get
        Set(value As Decimal)
            _VlrAdto = value
        End Set
    End Property
    Public Property VlrVariacao As Decimal
        Get
            Return _VlrVariacao
        End Get
        Set(value As Decimal)
            _VlrVariacao = value
        End Set
    End Property
    Public Property VlrBaixa As Decimal
        Get
            Return _VlrBaixa
        End Get
        Set(value As Decimal)
            _VlrBaixa = value
        End Set
    End Property
    Public Property VlrSaldo As Decimal
        Get
            Return _VlrSaldo
        End Get
        Set(value As Decimal)
            _VlrSaldo = value
        End Set
    End Property

    Public Property VlrBaixaOficial As Decimal
        Get
            Return _VlrBaixaOficial
        End Get
        Set(value As Decimal)
            _VlrBaixaOficial = value
        End Set
    End Property
    Public Property VlrBaixaMoeda As Decimal
        Get
            Return _VlrBaixaMoeda
        End Get
        Set(value As Decimal)
            _VlrBaixaMoeda = value
        End Set
    End Property
    Public Property VlrSaldoOficial As Decimal
        Get
            Return _VlrSaldoOficial
        End Get
        Set(value As Decimal)
            _VlrSaldoOficial = value
        End Set
    End Property
    Public Property VlrSaldoMoeda As Decimal
        Get
            Return _VlrSaldoMoeda
        End Get
        Set(value As Decimal)
            _VlrSaldoMoeda = value
        End Set
    End Property

    '***************************************************************
    '************************  CONTROLE  ***************************
    '***************************************************************
    Public Property Elegivel As Boolean
        Get
            Return _Elegivel
        End Get
        Set(value As Boolean)
            _Elegivel = value
        End Set
    End Property
    Public Property VlrABaixarOficial As Decimal
        Get
            Return _VlrABaixarOficial
        End Get
        Set(value As Decimal)
            _VlrABaixarOficial = value
        End Set
    End Property
    Public Property VlrABaixarMoeda As Decimal
        Get
            Return _VlrABaixarMoeda
        End Get
        Set(value As Decimal)
            _VlrABaixarMoeda = value
        End Set
    End Property

    Public ReadOnly Property VlrRealBaixaOficial
        Get
            If Me.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return VlrSaldoOficial
            End If

            If VlrSaldoMoeda = 0 Then Return 0

            If VlrABaixarMoeda = VlrSaldoMoeda Then
                Return VlrSaldoOficial
            End If
            Return Math.Round((VlrABaixarMoeda * (VlrSaldoOficial / VlrSaldoMoeda)), 2)
        End Get
    End Property
    Public ReadOnly Property VlrCalcVariacao
        Get
            If Moeda.Classificacao = eTiposMoeda.Oficial Then
                Return 0
            Else
                Return VlrABaixarOficial - VlrRealBaixaOficial
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
                Dim Num As New Numerador(1)
                _Codigo = Num.Sequencia
                Sqls.Add(Num.IncrementarNumeradorSql)

                Sql = "Insert Into Adiantamentos(Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id, RegistroPedido, Titulo, Safra, Vencimento, Taxa, Periodicidade, UltimaAtualizacao, Informacoes, Observacoes, DataDeCalculo)" & vbCrLf & _
                      " values ('" & Me.CodigoEmpresa & "'," & Me.EndEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EndCliente & "," & Me.Codigo & "," & Me.RegistroPedido & "," & Me.CodigoTitulo & ",'" & Me.CodigoSafra & "','" & Me.Vencimento.ToSqlDate & "'," & vbCrLf & _
                      "         ," & Str(Me.Taxa) & "," & Me.Periodicidade & "," & Me.UltimaAtualizacao.ToSqlDateNULL & ",'" & Me.Informacoes & "','" & Me.Observacoes & "'," & Me.DataDeCalculo.ToSqlDateNULL & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Adiantamentos SET" & vbCrLf & _
                      "    RegistroPedido    = " & Me.RegistroPedido & vbCrLf & _
                      "   ,Safra             ='" & Me.CodigoSafra & "'" & vbCrLf & _
                      "   ,Vencimento        ='" & Me.Vencimento.ToSqlDate & "'" & vbCrLf & _
                      "   ,Taxa              = " & Me.Taxa & vbCrLf & _
                      "   ,Periodicidade     = " & Me.Periodicidade & vbCrLf & _
                      "   ,UltimaAtualizacao ='" & Me.UltimaAtualizacao.ToSqlDate & "'" & vbCrLf & _
                      "   ,Informacoes       ='" & Me.Informacoes & "'" & vbCrLf & _
                      "   ,Observacoes       ='" & Me.Observacoes & "'" & vbCrLf & _
                      "   ,DataDeCalculo     = " & Me.DataDeCalculo.ToSqlDateNULL & vbCrLf & _
                      " WHERE Titulo = " & Me.CodigoTitulo
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE Adiantamentos " & vbCrLf & _
                      " WHERE Titulo = " & Me.CodigoTitulo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class