Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListDisponibilidade
    Inherits List(Of Disponibilidade)

    Public Sub New(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer)
        Dim sql As String = ""
        sql = "Select Disponibilidade_Id, Empresa, EndEmpresa, Consolidado, DataInicialApuracao, isnull(DataInicialCargaPatio,DataInicialApuracao) as DataInicialCargaPatio" & vbCrLf & _
              "  from SaldoInicialDisponibilidade" & vbCrLf & _
              " Where Empresa    ='" & pCodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa = " & pEndEmpresa

        Dim banco As New AcessaBanco
        Dim ds As DataSet
        ds = banco.ConsultaDataSet(sql, "Disponibilidade")

        For Each row In ds.Tables(0).Rows
            Dim Dis As New Disponibilidade
            Dis.CodigoDisponibilidade = row("Disponibilidade_Id")
            Dis.CodigoEmpresa = row("Empresa")
            Dis.EndEmpresa = row("EndEmpresa")
            Dis.Consolidado = row("Consolidado")
            Dis.DataInicial = row("DataInicialApuracao")
            Dis.DataInicialCargaPatio = row("DataInicialCargaPatio")
            Me.Add(Dis)
        Next
    End Sub

End Class

<Serializable()> _
Public Class Disponibilidade
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoDisponibilidade As Integer)
        Dim sql As String = ""
        sql = "Select Disponibilidade_Id, Empresa, EndEmpresa, Consolidado, DataInicialApuracao, isnull(DataInicialCargaPatio,DataInicialApuracao) as DataInicialCargaPatio " & vbCrLf & _
              "  from SaldoInicialDisponibilidade" & vbCrLf & _
              " Where Disponibilidade_id = " & pCodigoDisponibilidade

        Dim banco As New AcessaBanco
        Dim ds As DataSet
        ds = banco.ConsultaDataSet(sql, "Disponibilidade")

        For Each row In ds.Tables(0).Rows
            CodigoDisponibilidade = row("Disponibilidade_Id")
            CodigoEmpresa = row("Empresa")
            EndEmpresa = row("EndEmpresa")
            Consolidado = row("Consolidado")
            DataInicial = row("DataInicialApuracao")
            DataInicialCargaPatio = row("DataInicialCargaPatio")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoDisponibilidade As Integer
    Private _CodigoEmpresa As String = ""
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente
    Private _Consolidado As Boolean
    Private _DataInicial As Date
    Private _DataInicialCargaPatio As Date
    Private _Produtos As ListDisponibilidadeProduto
    Private _Depositos As ListDisponibilidadeDeposito
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoDisponibilidade As Integer
        Get
            Return _CodigoDisponibilidade
        End Get
        Set(ByVal value As Integer)
            _CodigoDisponibilidade = value
        End Set
    End Property

    Public Property CodigoEmpresa As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property EndEmpresa As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public ReadOnly Property Empresa As Cliente
        Get
            If _Empresa Is Nothing And Me.CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(Me.CodigoEmpresa, Me.EndEmpresa)
            Return _Empresa
        End Get
    End Property

    Public ReadOnly Property NomeEmpresa As String
        Get
            If Me.Empresa Is Nothing Then Return ""
            Return Me.Empresa.Nome & "...." & Me.Empresa.Cidade & "-" & Me.Empresa.CodigoEstado
        End Get
    End Property

    Public Property Consolidado As Boolean
        Get
            Return _Consolidado
        End Get
        Set(ByVal value As Boolean)
            _Consolidado = value
        End Set
    End Property

    Public Property DataInicial As Date
        Get
            Return _DataInicial
        End Get
        Set(ByVal value As Date)
            _DataInicial = value
        End Set
    End Property

    Public Property DataInicialCargaPatio As Date
        Get
            Return _DataInicialCargaPatio
        End Get
        Set(ByVal value As Date)
            _DataInicialCargaPatio = value
        End Set
    End Property


    Public Property Produtos As ListDisponibilidadeProduto
        Get
            If _Produtos Is Nothing Then _Produtos = New ListDisponibilidadeProduto(Me)
            Return _Produtos
        End Get
        Set(ByVal value As ListDisponibilidadeProduto)
            _Produtos = value
        End Set
    End Property

    Public Property Depositos As ListDisponibilidadeDeposito
        Get
            If _Depositos Is Nothing Then _Depositos = New ListDisponibilidadeDeposito(Me)
            Return _Depositos
        End Get
        Set(ByVal value As ListDisponibilidadeDeposito)
            _Depositos = value
        End Set
    End Property

    Public ReadOnly Property DescricaoProdutos As String
        Get
            If Me.Produtos Is Nothing Then Return ""
            Return Me.Produtos.DescricaoProdutos
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

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String = ""

        Dim ds As DataSet
        Dim banco As New AcessaBanco

        If IUD = "I" Then
            ds = banco.ConsultaDataSet("select isnull(max(disponibilidade_ID),0) + 1 as ID from SaldoInicialDisponibilidade", "Consulta")
            Me.CodigoDisponibilidade = ds.Tables(0).Rows(0)("ID")
        End If
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into SaldoInicialDisponibilidade(Disponibilidade_Id,Empresa, EndEmpresa, Consolidado, DataInicialApuracao, DataInicialCargaPatio) " & vbCrLf & _
                      " Values(" & Me.CodigoDisponibilidade & ",'" & Me.CodigoEmpresa & "'," & Me.EndEmpresa & "," & IIf(Me.Consolidado, 1, 0) & ",'" & Me.DataInicial.ToString("yyyy-MM-dd") & "','" & Me.DataInicialCargaPatio.ToString("yyyy-MM-dd") & "')"
                Sqls.Add(sql)
                SalvarTabelasRelacionada(Sqls)
            Case "D"
                SalvarTabelasRelacionada(Sqls)
                sql = " Delete SaldoInicialDisponibilidade" & vbCrLf & _
                      "  Where Disponibilidade_Id    = " & Me.CodigoDisponibilidade & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub

    Public Sub SalvarTabelasRelacionada(ByRef Sqls As ArrayList)
        Me.Produtos.SalvarSql(Sqls)
        Me.Depositos.SalvarSql(Sqls)
    End Sub

    Public Function getSqlProduto() As String
        If Produtos.Count = 0 Then Return ""
        Dim pPRD As String = ""
        For Each row In Produtos
            pPRD &= IIf(pPRD = "", "", ",") & "'" & row.CodigoProduto & "'"
        Next
        Return pPRD
    End Function
#End Region

End Class
