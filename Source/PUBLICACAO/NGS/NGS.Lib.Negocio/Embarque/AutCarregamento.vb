Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAutCarregamento
    Inherits List(Of AutCarregamento)

    Public Sub New()
    End Sub

    Public Sub New(pEmbarqueEntrega As EmbarqueXEntrega)
        _ParentEmbarqueEntrega = pEmbarqueEntrega

        Dim sql As String = ""
        sql &= "Select Carregamento_Id, Empresa, EndEmpresa, Pedido, Entrega, EndEntrega, Placa, PercAdiantamento, Situacao, " & vbCrLf & _
            "           isnull(MotivoCancelamento, '') as MotivoCancelamento, NrCarregamentoTerceiro, UsuarioInclusao, UsuarioInclusaoData, isnull(UsuarioCancelamento, '') as UsuarioCancelamento, isnull(UsuarioCancelamentoData, '') as UsuarioCancelamentoData, Movimento, " & vbCrLf & _
            "           Transportador, EndTransportador, Motorista, EndMotorista" & vbCrLf & _
            "       From AutCarregamento " & vbCrLf & _
            "       Where       Empresa     = '" & ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
            "           And     EndEmpresa  = " & ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
            "           And     Pedido      = " & ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
            "           And     Entrega     = '" & ParentEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
            "           And     EndEntrega  = " & ParentEmbarqueEntrega.EndClienteEntrega & vbCrLf & _
            "           And     Situacao    = 1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ParentEmbarqueEntrega.Placa) Then
            sql &= "           And     Placa       = '" & ParentEmbarqueEntrega.Placa & "'" & vbCrLf
        End If

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "AutCarregamento")

        For Each row In ds.Tables(0).Rows
            Dim obj As New AutCarregamento(ParentEmbarqueEntrega)
            obj.Carregamento_Id = row("Carregamento_Id")
            obj.Placa = row("Placa")
            obj.PercAdiantamento = row("PercAdiantamento")
            obj.Situacao = row("Situacao")
            obj.MotivoCancelamento = row("MotivoCancelamento")
            obj.NrCarregamentoTerceiro = row("NrCarregamentoTerceiro")
            obj.UsuarioInclusao = row("UsuarioInclusao")
            obj.UsuarioInclusaoData = row("UsuarioInclusaoData")
            obj.UsuarioCancelamento = row("UsuarioCancelamento")
            obj.UsuarioCancelamentoData = row("UsuarioCancelamentoData")
            obj.CodigoTransportador = row("Transportador")
            obj.EndTransportador = row("EndTransportador")
            obj.CodigoMotorista = row("Motorista")
            obj.EndMotorista = row("EndMotorista")
            obj.Movimento = row("Movimento")
            Me.Add(obj)
        Next
    End Sub

    Public Sub New(ByVal pedido As String, Optional ByVal placa As String = "")
        Dim sql As String = ""
        sql &= "Select Carregamento_Id, Empresa, EndEmpresa, Pedido, Entrega, EndEntrega, Placa, PercAdiantamento, Situacao, " & vbCrLf & _
            "           isnull(MotivoCancelamento, '') as MotivoCancelamento, NrCarregamentoTerceiro, UsuarioInclusao, UsuarioInclusaoData, isnull(UsuarioCancelamento, '') as UsuarioCancelamento, isnull(UsuarioCancelamentoData, '') as UsuarioCancelamentoData, Movimento, " & vbCrLf & _
            "           Transportador, EndTransportador, Motorista, EndMotorista" & vbCrLf & _
            "       From AutCarregamento " & vbCrLf & _
            "       Where       1=1     " & vbCrLf & _
            "           And     Pedido  = " & pedido & vbCrLf & _
            "           And     Situacao    = 1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(placa) Then
            sql &= "    And     Placa       = '" & placa & "'" & vbCrLf
        End If

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "AutCarregamento")

        For Each row In ds.Tables(0).Rows
            Dim obj As New AutCarregamento(ParentEmbarqueEntrega)
            obj.Carregamento_Id = row("Carregamento_Id")
            obj.Placa = row("Placa")
            obj.PercAdiantamento = row("PercAdiantamento")
            obj.Situacao = row("Situacao")
            obj.MotivoCancelamento = row("MotivoCancelamento")
            obj.NrCarregamentoTerceiro = row("NrCarregamentoTerceiro")
            obj.UsuarioInclusao = row("UsuarioInclusao")
            obj.UsuarioInclusaoData = row("UsuarioInclusaoData")
            obj.UsuarioCancelamento = row("UsuarioCancelamento")
            obj.UsuarioCancelamentoData = row("UsuarioCancelamentoData")
            obj.CodigoTransportador = row("Transportador")
            obj.EndTransportador = row("EndTransportador")
            obj.CodigoMotorista = row("Motorista")
            obj.EndMotorista = row("EndMotorista")
            obj.Movimento = row("Movimento")
            Me.Add(obj)
        Next
    End Sub

#Region "Fields"
    Private _ParentEmbarqueEntrega As EmbarqueXEntrega
#End Region

#Region "Property"
    Public ReadOnly Property ParentEmbarqueEntrega As EmbarqueXEntrega
        Get
            Return _ParentEmbarqueEntrega
        End Get
    End Property
#End Region

End Class

<Serializable()> _
Public Class AutCarregamento
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(pEmbarqueEntrega As EmbarqueXEntrega)
        _ParentEmbarqueEntrega = pEmbarqueEntrega
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _ParentEmbarqueEntrega As EmbarqueXEntrega
    Private _Carregamento_Id As String
    Private _Movimento As DateTime
    Private _Placa As String
    Private _PercAdiantamento As Decimal
    Private _Situacao As String
    Private _MotivoCancelamento As String
    Private _NrCarregamentoTerceiro As String
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioCancelamento As String
    Private _UsuarioCancelamentoData As DateTime?
    Private _CodigoTransportador As String
    Private _EndTransportador As Integer
    Private _CodigoMotorista As String
    Private _EndMotorista As Integer
    Private _ListCarregamentoXItens As ListAutCarregamentoXItens
    'Private _TotalCarregado As Decimal

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

    Public ReadOnly Property ParentEmbarqueEntrega As EmbarqueXEntrega
        Get
            Return _ParentEmbarqueEntrega
        End Get
    End Property

    Public Property Carregamento_Id() As String
        Get
            Return _Carregamento_Id
        End Get
        Set(ByVal value As String)
            _Carregamento_Id = value
        End Set
    End Property

    Public Property Placa() As String
        Get
            Return _Placa
        End Get
        Set(ByVal value As String)
            _Placa = value
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

    Public Property PercAdiantamento() As Decimal
        Get
            Return _PercAdiantamento
        End Get
        Set(ByVal value As Decimal)
            _PercAdiantamento = value
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

    Public Property MotivoCancelamento() As String
        Get
            Return _MotivoCancelamento
        End Get
        Set(ByVal value As String)
            _MotivoCancelamento = value
        End Set
    End Property

    Public Property NrCarregamentoTerceiro() As String
        Get
            Return _NrCarregamentoTerceiro
        End Get
        Set(ByVal value As String)
            _NrCarregamentoTerceiro = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property UsuarioInclusaoData() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioInclusaoData = value
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

    Public Property UsuarioCancelamentoData() As DateTime?
        Get
            Return _UsuarioCancelamentoData
        End Get
        Set(ByVal value As DateTime?)
            _UsuarioCancelamentoData = value
        End Set
    End Property

    Public Property CodigoTransportador() As String
        Get
            Return _CodigoTransportador
        End Get
        Set(ByVal value As String)
            _CodigoTransportador = value
        End Set
    End Property

    Public Property EndTransportador() As Integer
        Get
            Return _EndTransportador
        End Get
        Set(ByVal value As Integer)
            _EndTransportador = value
        End Set
    End Property

    Public Property CodigoMotorista() As String
        Get
            Return _CodigoMotorista
        End Get
        Set(ByVal value As String)
            _CodigoMotorista = value
        End Set
    End Property

    Public Property EndMotorista() As Integer
        Get
            Return _EndMotorista
        End Get
        Set(ByVal value As Integer)
            _EndMotorista = value
        End Set
    End Property

    Public Property ListCarregamentoXItens() As ListAutCarregamentoXItens
        Get
            If _ListCarregamentoXItens Is Nothing Then _ListCarregamentoXItens = New ListAutCarregamentoXItens(Me)
            Return _ListCarregamentoXItens
        End Get
        Set(ByVal value As ListAutCarregamentoXItens)
            _ListCarregamentoXItens = value
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
                sql = "INSERT INTO AutCarregamento (Carregamento_Id, Empresa, EndEmpresa, Pedido, Entrega, EndEntrega, Placa, " & vbCrLf & _
                    "                               PercAdiantamento, Situacao, NrCarregamentoTerceiro, UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf & _
                    "                               Movimento, Transportador, EndTransportador, Motorista, EndMotorista) " & vbCrLf & _
                    "       VALUES ( " & Me.Carregamento_Id & ", '" & ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "', " & ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & ", " & vbCrLf & _
                    "               " & ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & ", '" & ParentEmbarqueEntrega.CodigoClienteEntrega & "', " & ParentEmbarqueEntrega.EndClienteEntrega & "," & vbCrLf & _
                    "               '" & Me.Placa & "', " & Str(Me.PercAdiantamento) & ", " & Me.Situacao & ", '" & Me.NrCarregamentoTerceiro & "', '" & Me.UsuarioInclusao & "', " & vbCrLf & _
                    "               '" & Me.UsuarioInclusaoData.ToString("yyyy-MM-dd") & "', '" & Me.Movimento.ToString("yyyy-MM-dd") & "', '" & Me.CodigoTransportador & "', " & vbCrLf & _
                    "                " & Me.EndTransportador & ", '" & Me.CodigoMotorista & "', " & Me.EndMotorista & ")" & vbCrLf
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = "UPDATE AutCarregamento SET" & vbCrLf & _
                "                         Placa = '" & Me.Placa & "'" & vbCrLf & _
                "                       , PercAdiantamento        = " & Str(Me.PercAdiantamento) & vbCrLf & _
                "                       , Situacao                = " & Me.Situacao & vbCrLf & _
                "                       , MotivoCancelamento      = '" & Me.MotivoCancelamento & "'" & vbCrLf & _
                "                       , NrCarregamentoTerceiro  = '" & Me.NrCarregamentoTerceiro & "'" & vbCrLf & _
                "                       , UsuarioInclusao         = '" & Me.UsuarioInclusao & "'" & vbCrLf & _
                "                       , UsuarioInclusaoData     = '" & Me.UsuarioInclusaoData.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                "                       , UsuarioCancelamento     = '" & Me.UsuarioCancelamento & "'" & vbCrLf & _
                "                       , UsuarioCancelamentoData = '" & IIf(String.IsNullOrWhiteSpace(Me.UsuarioCancelamento), "null", Me.UsuarioCancelamentoData.ToString("yyyy-MM-dd")) & "'" & vbCrLf & _
                "                       , Movimento               = '" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                "                       , Transportador           = '" & Me.CodigoTransportador & "'" & vbCrLf & _
                "                       , EndTransportador        = " & Me.EndTransportador & vbCrLf & _
                "                       , Motorista               = '" & Me.CodigoMotorista & "'" & vbCrLf & _
                "                       , EndMotorista            = " & Me.EndMotorista & vbCrLf & _
                "       WHERE   Carregamento_Id     = " & Me.Carregamento_Id & vbCrLf & _
                "           And Empresa             = '" & ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                "           And EndEmpresa          = " & ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                "           And Pedido              = " & ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                "           And Entrega             = '" & ParentEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
                "           And EndEntrega          = " & ParentEmbarqueEntrega.EndClienteEntrega & vbCrLf


                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                sql = "DELETE AutCarregamento " & vbCrLf & _
                      "  WHERE  Carregamento_Id = " & Me.Carregamento_Id & vbCrLf & _
                      "     And Empresa     = '" & ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "     And EndEmpresa  = " & ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "     And Pedido      = " & ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "     And Entrega     = '" & ParentEmbarqueEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
                      "     And EndEntrega  = " & ParentEmbarqueEntrega.EndClienteEntrega & vbCrLf
                SalvarTabelasRelacionadasSql(Sqls)
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Me.ListCarregamentoXItens IsNot Nothing AndAlso Me.ListCarregamentoXItens.Count > 0 Then
            For Each item As AutCarregamentoXItens In Me.ListCarregamentoXItens
                If Me.IUD <> "U" Then
                    item.IUD = Me.IUD
                ElseIf item.IUD = "" Then
                    item.IUD = "U"
                End If
                item.SalvarSql(Sqls)
            Next
        End If
    End Sub
#End Region

End Class