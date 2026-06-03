Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAutCarregamentoXItens
    Inherits List(Of AutCarregamentoXItens)

    Public Sub New()
    End Sub

    Public Sub New(ByVal objAutCarregamento As AutCarregamento)
        _ParentAutCarregamento = objAutCarregamento

        Dim sql As String = ""
        sql &= "Select Carregamento_Id, Empresa_Id, EndEmpresa_Id, Pedido_Id, Entrega_Id, EndEntrega_Id, NrCotacao, Produto_Id, QuantidadeProgramado, PesoProgramado, VolumesProgramado" & vbCrLf & _
            "           From AutCarregamentoXItens " & vbCrLf & _
            "       Where   Carregamento_Id = '" & ParentAutCarregamento.Carregamento_Id & "'" & vbCrLf & _
            "               And Empresa_Id = '" & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
            "               And EndEmpresa_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
            "               And Pedido_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
            "               And Entrega_Id = '" & ParentAutCarregamento.ParentEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
            "               And EndEntrega_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.EndClienteEntrega & vbCrLf

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "AutCarregamento")

        For Each row In ds.Tables(0).Rows
            Dim obj As New AutCarregamentoXItens(ParentAutCarregamento)
            obj.CodigoProduto = row("Produto_Id")
            obj.NrCotacao_Id = row("NrCotacao")
            'obj.QuantidadeProgramado = row("QuantidadeProgramado")
            'obj.PesoProgramado = row("PesoProgramado")
            'obj.VolumesProgramado = row("VolumesProgramado")

            Me.Add(obj)
        Next
    End Sub

    Public Sub New(ByVal objAutCarregamento As AutCarregamento, ByVal nrCotacao As Integer)
        _ParentAutCarregamento = objAutCarregamento

        Dim sql As String = ""
        sql &= "Select Carregamento_Id, Empresa_Id, EndEmpresa_Id, Pedido_Id, Entrega_Id, EndEntrega_Id, NrCotacao, Produto_Id, QuantidadeProgramado, PesoProgramado, VolumesProgramado" & vbCrLf & _
            "           From AutCarregamentoXItens " & vbCrLf & _
            "       Where   Empresa_Id = '" & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
            "               And EndEmpresa_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
            "               And Pedido_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
            "               And Entrega_Id = '" & ParentAutCarregamento.ParentEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
            "               And EndEntrega_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.EndClienteEntrega & vbCrLf & _
            "               And NrCotacao = " & nrCotacao

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "AutCarregamento")

        For Each row In ds.Tables(0).Rows
            Dim obj As New AutCarregamentoXItens(ParentAutCarregamento)
            obj.CodigoProduto = row("Produto_Id")
            obj.NrCotacao_Id = row("NrCotacao")
            'obj.QuantidadeProgramado = row("QuantidadeProgramado")
            'obj.PesoProgramado = row("PesoProgramado")
            'obj.VolumesProgramado = row("VolumesProgramado")
            Me.Add(obj)
        Next
    End Sub

    'Public Sub New(ByVal objAutCarregamento As AutCarregamento, ByVal verificaCarregados As Boolean)
    '    _ParentAutCarregamento = objAutCarregamento

    '    Dim sql As String = "   SELECT	ISNULL(SUM(ACI.QuantidadeProgramado),0) as SomaCarregados" & vbCrLf & _
    '        "       	FROM	AutCarregamento AC                                   " & vbCrLf & _
    '        "       		INNER JOIN  AutCarregamentoXItens ACI                    " & vbCrLf & _
    '        "       			ON	ACI.Empresa_Id = AC.Empresa                      " & vbCrLf & _
    '        "       			AND AC.Carregamento_Id = ACI.Carregamento_Id         " & vbCrLf & _
    '        "       			AND AC.EndEmpresa = ACI.EndEmpresa_Id                " & vbCrLf & _
    '        "       			AND AC.Pedido = ACI.Pedido_Id                        " & vbCrLf & _
    '        "       			AND AC.Entrega = ACI.Entrega_Id                      " & vbCrLf & _
    '        "       			AND AC.EndEntrega = ACI.EndEntrega_Id                " & vbCrLf & _
    '        "           Where	And AC.Situacao    = 1                        " & vbCrLf & _
    '        "               And Empresa_Id = '" & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
    '        "               And EndEmpresa_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
    '        "               And Pedido_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
    '        "       		AND ACI.Produto_Id = " & _ParentAutCarregamento.ListCarregamentoXItens.pro & "                     " & vbCrLf & _
    '        "       			                               " & vbCrLf
    '    sql &= "Select Carregamento_Id, Empresa_Id, EndEmpresa_Id, Pedido_Id, Entrega_Id, EndEntrega_Id, NrCotacao, Produto_Id, QuantidadeProgramado, PesoProgramado, VolumesProgramado" & vbCrLf & _
    '        "           From AutCarregamentoXItens " & vbCrLf & _
    '        "       Where   Carregamento_Id = '" & ParentAutCarregamento.Carregamento_Id & "'" & vbCrLf & _

    '        "               And Entrega_Id = '" & ParentAutCarregamento.ParentEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
    '        "               And EndEntrega_Id = " & ParentAutCarregamento.ParentEmbarqueEntrega.EndClienteEntrega & vbCrLf

    '    Dim db As New AcessaBanco
    '    Dim ds As DataSet = db.ConsultaDataSet(sql, "AutCarregamento")

    '    For Each row In ds.Tables(0).Rows
    '        Dim obj As New AutCarregamentoXItens(ParentAutCarregamento)
    '        obj.CodigoProduto = row("Produto_Id")
    '        obj.NrCotacao_Id = row("NrCotacao")
    '        obj.QuantidadeProgramado = row("QuantidadeProgramado")
    '        obj.PesoProgramado = row("PesoProgramado")
    '        obj.VolumesProgramado = row("VolumesProgramado")

    '        Me.Add(obj)
    '    Next
    'End Sub

#Region "Fields"
    Private _ParentAutCarregamento As AutCarregamento
#End Region

#Region "Property"
    Public ReadOnly Property ParentAutCarregamento() As AutCarregamento
        Get
            Return _ParentAutCarregamento
        End Get
    End Property
#End Region

End Class

<Serializable()> _
Public Class AutCarregamentoXItens
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New(pAutCarregamento As AutCarregamento)
        _ParentAutCarregamento = pAutCarregamento
    End Sub

#End Region

#Region "Fields"

    Private _IUD As String
    Private _ParentAutCarregamento As AutCarregamento
    Private _NrCotacao_Id As Integer
    Private _Produto As Produto
    Private _CodigoProduto As String
    Private _QuantidadeProgramado As String
    Private _PesoProgramado As String
    Private _VolumesProgramado As String
    Private _ListCarregamentoXNotas As ListAutCarregamentoXNotaFiscal
#End Region

#Region "Properties"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property ParentAutCarregamento() As AutCarregamento
        Get
            Return _ParentAutCarregamento
        End Get
    End Property

    Public Property NrCotacao_Id As Integer
        Get
            Return _NrCotacao_Id
        End Get
        Set(value As Integer)
            _NrCotacao_Id = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto As Produto
        Get
            If _Produto Is Nothing And Me.CodigoProduto > 0 Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
        Set(value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property QuantidadeProgramado() As Decimal
        Get
            Return _QuantidadeProgramado
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeProgramado = value
        End Set
    End Property

    Public Property PesoProgramado() As Decimal
        Get
            Return _PesoProgramado
        End Get
        Set(ByVal value As Decimal)
            _PesoProgramado = value
        End Set
    End Property

    Public Property VolumesProgramado() As Decimal
        Get
            Return _VolumesProgramado
        End Get
        Set(ByVal value As Decimal)
            _VolumesProgramado = value
        End Set
    End Property

#End Region

#Region "Methods"

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim db As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If db.GravaBanco(Sqls) Then
            IUD = ""
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""
        Select Case _IUD
            Case "I"
                sql = "INSERT INTO AutCarregamentoXItens (Carregamento_Id, Empresa_Id, EndEmpresa_Id, Pedido_Id, Entrega_Id, EndEntrega_Id, " & vbCrLf & _
                    "                                       NrCotacao, Produto_Id, QuantidadeProgramado, PesoProgramado, VolumesProgramado) " & vbCrLf & _
                    "               VALUES (" & ParentAutCarregamento.Carregamento_Id & ", '" & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "', " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & ", " & vbCrLf & _
                    ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & ", '" & ParentAutCarregamento.ParentEmbarqueEntrega.CodigoClienteEntrega & "', " & ParentAutCarregamento.ParentEmbarqueEntrega.EndClienteEntrega & ", " & vbCrLf & _
                    Me.NrCotacao_Id & ", '" & Me.CodigoProduto & "', " & Str(Me.QuantidadeProgramado) & ", " & Str(Me.PesoProgramado) & ", " & Str(Me.VolumesProgramado) & ")"
                Sqls.Add(sql)

            Case "U"
                sql = "UPDATE AutCarregamentoXItens SET " & _
                "       QuantidadeProgramado = " & Str(Me.QuantidadeProgramado) & vbCrLf & _
                "     , PesoProgramado = " & Str(Me.PesoProgramado) & vbCrLf & _
                "     , VolumesProgramado = " & Str(Me.VolumesProgramado) & vbCrLf & _
                "   WHERE   Carregamento_Id = " & ParentAutCarregamento.Carregamento_Id & vbCrLf & _
                "       And Empresa_Id      = '" & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                "       And EndEmpresa_Id   = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                "       And Pedido_Id       = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                "       And Entrega_Id      = '" & ParentAutCarregamento.ParentEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
                "       And EndEntrega_Id   = " & ParentAutCarregamento.ParentEmbarqueEntrega.EndClienteEntrega & vbCrLf & _
                "       And NrCotacao       = " & Me.NrCotacao_Id & vbCrLf & _
                "       And Produto_Id      = " & Me.CodigoProduto & vbCrLf
                Sqls.Add(sql)

            Case "D"
                sql = "DELETE AutCarregamentoXItens " & _
                    "   WHERE   Carregamento_Id = " & ParentAutCarregamento.Carregamento_Id & vbCrLf & _
                    "       And Empresa_Id      = '" & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                    "       And EndEmpresa_Id   = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                    "       And Pedido_Id       = " & ParentAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                    "       And Entrega_Id      = '" & ParentAutCarregamento.ParentEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
                    "       And EndEntrega_Id   = " & ParentAutCarregamento.ParentEmbarqueEntrega.EndClienteEntrega & vbCrLf & _
                    "       And NrCotacao       = " & Me.NrCotacao_Id & vbCrLf & _
                    "       And Produto_Id      = '" & Me.CodigoProduto & "'" & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub

#End Region

End Class