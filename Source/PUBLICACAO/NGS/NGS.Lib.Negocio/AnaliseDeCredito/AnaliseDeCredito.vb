Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAnaliseDeCredito
    Inherits List(Of AnaliseDeCredito)

End Class

<Serializable()> _
Public Class AnaliseDeCredito
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pParametrosAnaliseDeCredito As ParametrosAnaliseDeCredito)
        _ParametrosAnaliseDeCredito = pParametrosAnaliseDeCredito
    End Sub

    Public Sub New(ByVal pCodigoAnalise As Integer)
        Dim sql As String
        sql = "SELECT Analise_Id, Ano_Id, DefinicaoAno_Id, DataAnalise," & vbCrLf & _
              "       AreaArrendada, CustoArrendamento, OutrasReceitas, OutrasDespesas," & vbCrLf & _
              "       LimiteCredito, PercRedutorRiscoCultura, Situacao, " & vbCrLf & _
              "       isnull(UsuarioAprovacao,'') as UsuarioAprovacao,  DataAprovacao," & vbCrLf & _
              "       isnull(UsuarioLiberacao,'') as UsuarioLiberacao, DataLiberacao, " & vbCrLf & _
              "       isnull(UsuarioCancelamento,'') as UsuarioCancelamento , DataCancelamento" & vbCrLf & _
              "  FROM AnaliseDeCredito" & vbCrLf & _
              " Where Analise_Id = " & pCodigoAnalise

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "AnaliseDeCredito")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _ParametrosAnaliseDeCredito = New ParametrosAnaliseDeCredito(row("Ano_Id"), row("DefinicaoAno_Id"))
        Me.CodigoAnalise = row("Analise_Id")
        Me.DataAnalise = row("DataAnalise")
        Me.AreaArrendada = row("AreaArrendada")
        Me.OutrasReceitas = row("OutrasReceitas")
        Me.OutrasDespesas = row("OutrasDespesas")
        Me.LimiteCredito = row("LimiteCredito")
        Me.Situacao = row("Situacao")
        Me.UsuarioAprovacao = row("UsuarioAprovacao")
        Me.UsuarioLiberacao = row("UsuarioLiberacao")
        Me.UsuarioCancelamento = row("UsuarioCancelamento")
        If Me.UsuarioAprovacao.Length > 0 Then Me.DataAprovacao = row("DataAprovacao")
        If Me.UsuarioLiberacao.Length > 0 Then Me.DataLiberacao = row("DataLiberacao")
        If Me.UsuarioCancelamento.Length > 0 Then Me.DataCancelamento = row("DataCancelamento")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _ParametrosAnaliseDeCredito As ParametrosAnaliseDeCredito
    Private _CodigoAnalise As Integer
    Private _DataAnalise As Date
    Private _AreaPlantioTotal As Decimal
    Private _AreaArrendada As Decimal
    Private _OutrasReceitas As Decimal
    Private _OutrasDespesas As Decimal
    Private _LimiteCredito As Char
    Private _Situacao As String = "NORMAL"

    Private _UsuarioAprovacao As String = ""
    Private _DataAprovacao As Date = Date.Now
    Private _UsuarioLiberacao As String = ""
    Private _DataLiberacao As Date = Date.Now
    Private _UsuarioCancelamento As String = ""
    Private _DataCancelamento As Date = Date.Now

    Private _Clientes As ListAnaliseDeCreditoCliente
    Private _Culturas As ListAnaliseDeCreditoCultura
    Private _Perguntas As ListAnaliseDeCreditoPergunta
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

    Public Property ParametrosAnaliseDeCredito() As ParametrosAnaliseDeCredito
        Get
            Return _ParametrosAnaliseDeCredito
        End Get
        Set(ByVal value As ParametrosAnaliseDeCredito)
            _ParametrosAnaliseDeCredito = value
        End Set
    End Property

    Public Property CodigoAnalise() As Integer
        Get
            Return _CodigoAnalise
        End Get
        Set(ByVal value As Integer)
            _CodigoAnalise = value
        End Set
    End Property

    Public Property DataAnalise() As Date
        Get
            Return _DataAnalise
        End Get
        Set(ByVal value As Date)
            _DataAnalise = value
        End Set
    End Property

    Public Property AreaPlantioTotal() As Decimal
        Get
            Return _AreaPlantioTotal
        End Get
        Set(ByVal value As Decimal)
            _AreaPlantioTotal = value
        End Set
    End Property

    Public Property AreaArrendada() As Decimal
        Get
            Return _AreaArrendada
        End Get
        Set(ByVal value As Decimal)
            _AreaArrendada = value
        End Set
    End Property

    Public ReadOnly Property CustoArrendamento() As Decimal
        Get
            Return AreaArrendada * ParametrosAnaliseDeCredito.CustoArrendamentoHa
        End Get
    End Property

    Public Property OutrasReceitas() As Decimal
        Get
            Return _OutrasReceitas
        End Get
        Set(ByVal value As Decimal)
            _OutrasReceitas = value
        End Set
    End Property

    Public Property OutrasDespesas() As Decimal
        Get
            Return _OutrasDespesas
        End Get
        Set(ByVal value As Decimal)
            _OutrasDespesas = value
        End Set
    End Property

    Public ReadOnly Property TotalReceitas() As Decimal
        Get
            Return Culturas.TotalReceitaCultura + OutrasReceitas
        End Get
    End Property

    Public ReadOnly Property TotalCustos() As Decimal
        Get
            Return Culturas.TotalCustoCultura + OutrasDespesas
        End Get
    End Property

    Public ReadOnly Property SaldoCliente() As Decimal
        Get
            Return TotalReceitas - TotalCustos
        End Get
    End Property

    Public ReadOnly Property CreditoConcedido() As Decimal
        Get
            If SaldoCliente > 0 Then
                Select Case Perguntas.RatingCreditoCliente
                    Case "A"
                        Return (Culturas.TotalCustoPortifolio * (ParametrosAnaliseDeCredito.PercLimiteCreditoA - Culturas.CoefRedutorRiscoCultura)) / 100
                    Case "B"
                        Return (Culturas.TotalCustoPortifolio * (ParametrosAnaliseDeCredito.PercLimiteCreditoB - Culturas.CoefRedutorRiscoCultura)) / 100
                    Case "C"
                        Return (Culturas.TotalCustoPortifolio * (ParametrosAnaliseDeCredito.PercLimiteCreditoC - Culturas.CoefRedutorRiscoCultura)) / 100
                End Select
            End If
            Return 0
        End Get
    End Property

    Public Property LimiteCredito() As Char
        Get
            Return _LimiteCredito
        End Get
        Set(ByVal value As Char)
            _LimiteCredito = value
        End Set
    End Property

    Public Property Situacao() As String
        Get
            Return _Situacao
        End Get
        Set(ByVal value As String)
            _Situacao = value
        End Set
    End Property

    Public Property UsuarioAprovacao() As String
        Get
            Return _UsuarioAprovacao
        End Get
        Set(ByVal value As String)
            _UsuarioAprovacao = value
        End Set
    End Property

    Public Property DataAprovacao() As Date
        Get
            Return _DataAprovacao
        End Get
        Set(ByVal value As Date)
            _DataAprovacao = value
        End Set
    End Property

    Public Property UsuarioLiberacao() As String
        Get
            Return _UsuarioLiberacao
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacao = value
        End Set
    End Property

    Public Property DataLiberacao() As Date
        Get
            Return _DataLiberacao
        End Get
        Set(ByVal value As Date)
            _DataLiberacao = value
        End Set
    End Property

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property DataCancelamento() As Date
        Get
            Return _DataCancelamento
        End Get
        Set(ByVal value As Date)
            _DataCancelamento = value
        End Set
    End Property

    'LISTAS
    Public Property Clientes() As ListAnaliseDeCreditoCliente
        Get
            If _Clientes Is Nothing Then _Clientes = New ListAnaliseDeCreditoCliente(Me)
            Return _Clientes
        End Get
        Set(ByVal value As ListAnaliseDeCreditoCliente)
            _Clientes = value
        End Set
    End Property

    Public Property Culturas() As ListAnaliseDeCreditoCultura
        Get
            If _Culturas Is Nothing Then _Culturas = New ListAnaliseDeCreditoCultura(Me)
            Return _Culturas
        End Get
        Set(ByVal value As ListAnaliseDeCreditoCultura)
            _Culturas = value
        End Set
    End Property

    Public Property Perguntas() As ListAnaliseDeCreditoPergunta
        Get
            If _Perguntas Is Nothing Then _Perguntas = New ListAnaliseDeCreditoPergunta(Me)
            Return _Perguntas
        End Get
        Set(ByVal value As ListAnaliseDeCreditoPergunta)
            _Perguntas = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function JaExisteCreditoLiberadoNesteAno() As Boolean
        Dim sql As String
        sql = "SELECT COUNT(*) AS num" & vbCrLf & _
              "  FROM AnaliseDeCredito " & vbCrLf & _
              " INNER JOIN AnaliseDeCreditoCliente" & vbCrLf & _
              "    ON AnaliseDeCredito.Analise_Id      = AnaliseDeCreditoCliente.Analise_Id" & vbCrLf & _
              "   AND AnaliseDeCredito.Ano_Id          = AnaliseDeCreditoCliente.Ano_Id" & vbCrLf & _
              "   AND AnaliseDeCredito.DefinicaoAno_Id = AnaliseDeCreditoCliente.DefinicaoAno_Id" & vbCrLf & _
              " WHERE AnaliseDeCredito.Ano_Id   = " & ParametrosAnaliseDeCredito.Ano & vbCrLf & _
              "   AND AnaliseDeCredito.Situacao = 'LIBERADA'" & vbCrLf & _
              "   AND AnaliseDeCreditoCliente.Cliente_id   in (" & Clientes.ClientesParaSql & ")" & vbCrLf
        Dim ds As DataSet
        Dim banco As New AcessaBanco
        ds = banco.ConsultaDataSet(sql, "Existe")
        If ds.Tables(0).Rows(0)(0) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub AtualizarAnalise()
        AtualizarCulturas()
        AtualizarAreaArrendada()
        AtualizarOutrasReceitasDespesas()
    End Sub

    Private Sub AtualizarCulturas()
        Culturas.Clear()

        Dim sql As String = ""
        sql = "SELECT CxS.Safra_Id, " & vbCrLf & _
              "       CxS.Cultura_Id," & vbCrLf & _
              "       sum(CxS.AreaPlantada) as AreaPlantio" & vbCrLf & _
              "  FROM ClienteXSafra CxS" & vbCrLf & _
              " Inner Join Safras S" & vbCrLf & _
              "    on S.Safra_Id = CxS.Safra_Id" & vbCrLf & _
              " Where year(S.vencimento) = " & ParametrosAnaliseDeCredito.Ano & vbCrLf & _
              "   and Cliente_Id in (" & Clientes.ClientesParaSql & ")" & vbCrLf & _
              " Group By CxS.Safra_Id," & vbCrLf & _
              "          CxS.Cultura_Id" & vbCrLf

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "clientexcultura")

        Dim TotalArea As Decimal
        For Each row As DataRow In ds.Tables(0).Rows
            Dim AnaXCult As New AnaliseDeCreditoCultura(Me)
            AnaXCult.CodigoSafra = row("Safra_Id")
            AnaXCult.CodigoCultura = row("Cultura_Id")
            AnaXCult.AreaPlantio = row("AreaPlantio")
            If AnaXCult.ParametrosCultura.CodigoCultura > 0 Then
                TotalArea += row("AreaPlantio")
                AnaXCult.PercCultura = -1
            End If
            Culturas.Add(AnaXCult)
        Next
        Me.AreaPlantioTotal = TotalArea

        For Each row As AnaliseDeCreditoCultura In Culturas
            If row.PercCultura = -1 Then
                row.PercCultura = Math.Round((row.AreaPlantio * 100) / TotalArea, 2)
            End If
        Next

    End Sub

    Private Sub AtualizarAreaArrendada()
        Dim sql As String = ""
        sql = "Select isnull(sum(area),0) as Area " & vbCrLf & _
              "  from ClienteXArrendante" & vbCrLf & _
              " where year(datacontrato_id) <= " & ParametrosAnaliseDeCredito.Ano & vbCrLf & _
              "   and year(dataVencimento)  >= " & ParametrosAnaliseDeCredito.Ano & vbCrLf & _
              "   and cliente_Id in (" & Clientes.ClientesParaSql & ")" & vbCrLf

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "ClienteXArrendante")

        _AreaArrendada = ds.Tables(0).Rows(0)(0)
    End Sub

    Public Sub AtualizarOutrasReceitasDespesas()
        Dim sql As String = ""
        sql = "select isnull(sum(case" & vbCrLf & _
              "					 when ReceitaDespesa = 'D'" & vbCrLf & _
              "					   then valorano" & vbCrLf & _
              "					   else 0" & vbCrLf & _
              "					end),0) as Despesas," & vbCrLf & _
              "       isnull(sum(case" & vbCrLf & _
              "					 when ReceitaDespesa = 'R'" & vbCrLf & _
              "					   then valorano" & vbCrLf & _
              "					   else 0" & vbCrLf & _
              "					end),0) as Receitas" & vbCrLf & _
              "  from ClientesXReceitasDespesas" & vbCrLf & _
              " Where cliente_Id in (" & Clientes.ClientesParaSql & ")" & vbCrLf & _
              "   and Ano_Id = " & ParametrosAnaliseDeCredito.Ano & vbCrLf

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "ClientesXReceitasDespesas")

        _OutrasDespesas = ds.Tables(0).Rows(0)("Despesas")
        _OutrasReceitas = ds.Tables(0).Rows(0)("Receitas")
    End Sub

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String = ""
        Select Case Me.IUD
            Case "I"
                Dim num As New Numerador(3)
                Me.CodigoAnalise = num.Sequencia + 1
                Sqls.Add(num.IncrementarNumeradorSql(True))

                strSQL &= " Insert Into AnaliseDeCredito (Analise_Id, Ano_Id, DefinicaoAno_Id, DataAnalise, AreaArrendada, CustoArrendamento, OutrasReceitas, OutrasDespesas, LimiteCredito, PercRedutorRiscoCultura, Situacao, UsuarioAprovacao, DataAprovacao, UsuarioLiberacao, DataLiberacao, UsuarioCancelamento, DataCancelamento)" & vbCrLf & _
                          " Values (" & Me.CodigoAnalise & "," & Me.ParametrosAnaliseDeCredito.Ano & "," & Me.ParametrosAnaliseDeCredito.DefinicaoAno & ",'" & Me.DataAnalise.ToString("yyyy-MM-dd") & "'," & Str(Me.AreaArrendada) & "," & Str(Me.CustoArrendamento) & "," & Str(Me.OutrasReceitas) & "," & Str(Me.OutrasDespesas) & ",'" & Me.LimiteCredito & "'," & Str(Me.Culturas.CoefRedutorRiscoCultura) & ",'" & Me.Situacao & "','" & Me.UsuarioAprovacao & "'," & IIf(Me.UsuarioAprovacao = "", "NULL", "'" & Me.DataAprovacao.ToString("yyyy-MM-dd") & "'") & ",'" & Me.UsuarioLiberacao & "'," & IIf(Me.UsuarioLiberacao = "", "NULL", "'" & Me.DataLiberacao.ToString("yyyy-MM-dd") & "'") & ",'" & Me.UsuarioCancelamento & "'," & IIf(Me.UsuarioCancelamento = "", "NULL", "'" & Me.DataCancelamento.ToString("yyyy-MM-dd") & "'") & ")"
                Sqls.Add(strSQL)

                Clientes.SalvarSql(Sqls)
                Culturas.SalvarSql(Sqls)
                Perguntas.SalvarSql(Sqls)
            Case "U"
                strSQL = " Update AnaliseDeCredito set " & vbCrLf & _
                         "      DataAnalise             ='" & Me.DataAnalise.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                         "     ,AreaArrendada           = " & Str(Me.AreaArrendada) & vbCrLf & _
                         "     ,CustoArrendamento       = " & Str(Me.CustoArrendamento) & vbCrLf & _
                         "     ,OutrasReceitas          = " & Str(Me.OutrasReceitas) & vbCrLf & _
                         "     ,OutrasDespesas          = " & Str(Me.OutrasDespesas) & vbCrLf & _
                         "     ,LimiteCredito           ='" & Me.LimiteCredito & "'" & vbCrLf & _
                         "     ,PercRedutorRiscoCultura = " & Str(Me.Culturas.CoefRedutorRiscoCultura) & vbCrLf & _
                         "     ,Situacao                ='" & Me.Situacao & "'" & vbCrLf & _
                         "     ,UsuarioAprovacao        ='" & Me.UsuarioAprovacao & "'" & vbCrLf & _
                         "     ,DataAprovacao           = " & IIf(Me.UsuarioAprovacao = "", "NULL", "'" & Me.DataAprovacao.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                         "     ,UsuarioLiberacao        ='" & Me.UsuarioLiberacao & "'" & vbCrLf & _
                         "     ,DataLiberacao           = " & IIf(Me.UsuarioLiberacao = "", "NULL", "'" & Me.DataLiberacao.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                         "     ,UsuarioCancelamento     ='" & Me.UsuarioCancelamento & "'" & vbCrLf & _
                         "     ,DataCancelamento        = " & IIf(Me.UsuarioCancelamento = "", "NULL", "'" & Me.DataCancelamento.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                         "  Where Analise_Id      = " & Me.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & Me.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Me.ParametrosAnaliseDeCredito.DefinicaoAno
                Sqls.Add(strSQL)
                Clientes.SalvarSql(Sqls)
                Culturas.SalvarSql(Sqls)
                Perguntas.SalvarSql(Sqls)
            Case "D"
                Perguntas.SalvarSql(Sqls)
                Culturas.SalvarSql(Sqls)
                Clientes.SalvarSql(Sqls)
                strSQL = " Delete AnaliseDeCredito" & vbCrLf & _
                         "  Where Analise_Id      = " & Me.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & Me.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Me.ParametrosAnaliseDeCredito.DefinicaoAno
                Sqls.Add(strSQL)
            Case "C"
                strSQL = " Update AnaliseDeCredito set " & vbCrLf & _
                         "      Situacao                ='" & Me.Situacao & "'" & vbCrLf & _
                         "     ,UsuarioCancelamento    ='" & Me.UsuarioCancelamento & "'" & vbCrLf & _
                         "     ,DataCancelamento        = " & IIf(Me.UsuarioCancelamento = "", "NULL", "'" & Me.DataCancelamento.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                         "  Where Analise_Id      = " & Me.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & Me.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Me.ParametrosAnaliseDeCredito.DefinicaoAno
                Sqls.Add(strSQL)
            Case "A"
                strSQL = " Update AnaliseDeCredito set " & vbCrLf & _
                         "      Situacao                ='" & Me.Situacao & "'" & vbCrLf & _
                         "     ,UsuarioAprovacao        ='" & Me.UsuarioAprovacao & "'" & vbCrLf & _
                         "     ,DataAprovacao           = " & IIf(Me.UsuarioAprovacao = "", "NULL", "'" & Me.DataAprovacao.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                         "  Where Analise_Id      = " & Me.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & Me.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Me.ParametrosAnaliseDeCredito.DefinicaoAno
                Sqls.Add(strSQL)
            Case "L"
                strSQL = " Update AnaliseDeCredito set " & vbCrLf & _
                         "      Situacao                ='" & Me.Situacao & "'" & vbCrLf & _
                         "     ,UsuarioLiberacao       ='" & Me.UsuarioLiberacao & "'" & vbCrLf & _
                         "     ,DataLiberacao           = " & IIf(Me.UsuarioLiberacao = "", "NULL", "'" & Me.DataLiberacao.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                         "  Where Analise_Id      = " & Me.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & Me.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Me.ParametrosAnaliseDeCredito.DefinicaoAno
                Sqls.Add(strSQL)

                strSQL = " Update AnaliseDeCredito set " & vbCrLf & _
                         "      Situacao                ='CANCELADA'" & vbCrLf & _
                         "     ,UsuarioCancelamento    ='" & Me.UsuarioAprovacao & "'" & vbCrLf & _
                         "     ,DataCancelamento        = " & IIf(Me.UsuarioAprovacao = "", "NULL", "'" & Me.DataAprovacao.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                         "   from AnaliseDeCredito " & vbCrLf & _
                         "  Where AnaliseDeCredito.Analise_id in (Select distinct Analise_id" & vbCrLf & _
                         "                                          From AnaliseDeCreditoCliente" & vbCrLf & _
                         "                                         where Ano_id = " & Me.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "                                           and cliente_id in (Select cliente_id" & vbCrLf & _
                         "                                                                from AnaliseDeCreditoCliente" & vbCrLf & _
                         "                                                               where Analise_id = " & Me.CodigoAnalise & "))" & vbCrLf & _
                         "   and AnaliseDeCredito.Analise_id <> " & Me.CodigoAnalise
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class
