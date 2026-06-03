Public Class ListOrdemParaProducaoXEmbalagens
    Inherits List(Of OrdemParaProducaoXEmbalagens)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objOrdemDeProducao As OrdemParaProducao)

        Me.OrdemParaProducao = objOrdemDeProducao

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Empresa_Id, EndEmpresa_Id, Ordem_Id, Embalagem_Id, Unidade_Id, Quantidade, Capacidade " & vbCrLf & _
                     "From OrdemDeProducaoXEmbalagens " & vbCrLf & _
                     "Where Empresa_id    = '" & Me.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf & _
                     "  and EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & vbCrLf

            If Me.OrdemParaProducao.Codigo > 0 Then strSQL &= "  and Ordem_Id      = " & Me.OrdemParaProducao.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "OrdemParaProducaoXEmbalagens")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPE As New OrdemParaProducaoXEmbalagens(Me.OrdemParaProducao)

                oPE.CodigoEmbalagem = row("Embalagem_Id")
                oPE.CodigoUnidade = row("Unidade_Id")
                oPE.Quantidade = row("Quantidade")
                oPE.Capacidade = row("Capacidade")

                Me.Add(oPE)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As OrdemParaProducaoXEmbalagens In Me
            If item.IUD = "I" OrElse item.IUD = "D" Then item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _OrdemParaProducao As OrdemParaProducao
#End Region

#Region "Property"
    Public Property OrdemParaProducao() As OrdemParaProducao
        Get
            Return _OrdemParaProducao
        End Get
        Set(ByVal value As OrdemParaProducao)
            _OrdemParaProducao = value
        End Set
    End Property

#End Region

End Class

<Serializable()> _
Public Class OrdemParaProducaoXEmbalagens
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _OrdemParaProducao As OrdemParaProducao

    Private _CodigoEmbalagem As Integer
    Private _Embalagem As Embalagem

    Private _CodigoUnidade As String
    Private _UnidadeDeMedida As UnidadeDeMedida

    Private _Quantidade As Integer
    Private _Capacidade As Decimal

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

    Public Property OrdemParaProducao() As OrdemParaProducao
        Get
            Return _OrdemParaProducao
        End Get
        Set(ByVal value As OrdemParaProducao)
            _OrdemParaProducao = value
        End Set
    End Property

    Public Property CodigoEmbalagem As Integer
        Get
            Return _CodigoEmbalagem
        End Get
        Set(value As Integer)
            _CodigoEmbalagem = value
        End Set
    End Property

    Public Property Embalagem As Embalagem
        Get
            If _Embalagem Is Nothing Then _Embalagem = New Embalagem(_CodigoEmbalagem)
            Return _Embalagem
        End Get
        Set(value As Embalagem)
            _Embalagem = value
        End Set
    End Property

    Public Property CodigoUnidade As String
        Get
            Return _CodigoUnidade
        End Get
        Set(value As String)
            _CodigoUnidade = value
        End Set
    End Property

    Public Property UnidadeDeMedida As UnidadeDeMedida
        Get
            If _UnidadeDeMedida Is Nothing Then _UnidadeDeMedida = New UnidadeDeMedida(_CodigoUnidade)
            Return _UnidadeDeMedida
        End Get
        Set(value As UnidadeDeMedida)
            _UnidadeDeMedida = value
        End Set
    End Property

    Public Property Quantidade As Integer
        Get
            Return _Quantidade
        End Get
        Set(value As Integer)
            _Quantidade = value
        End Set
    End Property

    Public Property Capacidade() As Decimal
        Get
            Return _Capacidade
        End Get
        Set(ByVal value As Decimal)
            _Capacidade = value
        End Set
    End Property

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal OrdemParaProducao As OrdemParaProducao)
        Me.OrdemParaProducao = OrdemParaProducao
    End Sub
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSql As String = ""

        Select Case Me.IUD
            Case "I"
                strSql = "INSERT INTO OrdemDeProducaoXEmbalagens(Empresa_Id, EndEmpresa_Id, Ordem_Id, Embalagem_Id, Unidade_Id, Quantidade, Capacidade)" & vbCrLf & _
                         "Values('" & OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducao.Codigo & "," & Me.CodigoEmbalagem & ",'" & Me.CodigoUnidade & "'," & Me.Quantidade & "," & Str(Me.Capacidade) & ")"

            Case "D"
                strSql = "Delete OrdemDeProducaoXEmbalagens " & vbCrLf & _
                         " Where Empresa_Id             = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id          = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf & _
                         "   and Ordem_Id               = " & OrdemParaProducao.Codigo & vbCrLf & _
                         "   and Embalagem_Id           = " & Me.CodigoEmbalagem & vbCrLf & _
                         "   and Unidade_Id             = '" & Me.CodigoUnidade & "'"
        End Select

        Sqls.Add(strSql)
    End Sub
#End Region

End Class
