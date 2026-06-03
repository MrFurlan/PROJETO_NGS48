Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

'****************************************************************************************************************************************************
'******************************************  LISTA CLASS BASE EMBARQUE X ENTREGA X ROTEIRO  *********************************************************
'****************************************************************************************************************************************************
<Serializable()> _
Public Class ListEmbarqueRoteiro
    Inherits List(Of EmbarqueRoteiro)

#Region "Construtor"
    Public Sub New(pEmbarquexEntrega As EmbarqueXEntrega)
        _ParentEntrega = pEmbarquexEntrega
        Dim sql As String = ""

        sql &= "Select Empresa_Id, EndEmpresa_Id, Pedido_Id, Roteiro_Id, Origem, EndOrigem, Destino, EndDestino, Entrega, EndEntrega, isnull(ViaDeTransporte,0) as ViaDeTransporte" & vbCrLf & _
               "  From AutEmbarqueXRoteiro" & vbCrLf & _
               " where Empresa_Id    ='" & pEmbarquexEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
               "   and EndEmpresa_Id = " & pEmbarquexEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
               "   and Pedido_Id     = " & pEmbarquexEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
               "   and Entrega       ='" & pEmbarquexEntrega.CodigoClienteEntrega & "'" & vbCrLf & _
               "   and EndEntrega    = " & pEmbarquexEntrega.EndClienteEntrega & vbCrLf

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "EmbarqueRoteiro")

        For Each row In ds.Tables(0).Rows
            Dim r As New EmbarqueRoteiro(pEmbarquexEntrega)
            r.CodigoRoteiro = row("Roteiro_Id")
            r.CodigoOrigem = row("Origem")
            r.EndOrigem = row("EndOrigem")
            r.CodigoDestino = row("Destino")
            r.EndDestino = row("EndDestino")
            r.CodigoViaDeTransporte = row("ViaDeTransporte")
            Me.Add(r)
        Next
    End Sub

#End Region

#Region "Fields"
    Private _ParentEntrega As EmbarqueXEntrega
#End Region

#Region "Property"
    Public ReadOnly Property ParentEntrega As EmbarqueXEntrega
        Get
            Return _ParentEntrega
        End Get
    End Property
#End Region

End Class

'****************************************************************************************************************************************************
'************************************************  CLASS BASE EMBARQUE X ENTREGA X ROTEIRO  *********************************************************
'****************************************************************************************************************************************************
<Serializable()> _
Public Class EmbarqueRoteiro
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(pEmbarqueXEntrega As EmbarqueXEntrega)
        _ParentEntrega = pEmbarqueXEntrega
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _ParentEntrega As EmbarqueXEntrega
    Private _CodigoRoteiro As Integer

    Private _CodigoViaDeTransporte As Integer
    Private _ViaDeTransporte As ViaDeTransporte

    Private _CodigoOrigem As String = ""
    Private _EndOrigem As Integer
    Private _Origem As Cliente

    Private _CodigoDestino As String = ""
    Private _EndDestino As Integer
    Private _Destino As Cliente

    Private _OrigemCompleto As String
    Private _DestinoCompleto As String

    Private _OrigemDescricao As String
    Private _DestinoDescricao As String

    Private _Transportadores As ListEmbarqueTransportador
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

    Public ReadOnly Property ParentEntrega As EmbarqueXEntrega
        Get
            Return _ParentEntrega
        End Get
    End Property

    Public Property CodigoRoteiro() As Integer
        Get
            Return _CodigoRoteiro
        End Get
        Set(ByVal value As Integer)
            _CodigoRoteiro = value
        End Set
    End Property

    'VIA DE TRANSPORTE
    Public Property CodigoViaDeTransporte As Integer
        Get
            Return _CodigoViaDeTransporte
        End Get
        Set(value As Integer)
            _CodigoViaDeTransporte = value
            _ViaDeTransporte = Nothing
        End Set
    End Property

    Public ReadOnly Property ViaDeTransporte As ViaDeTransporte
        Get
            If _ViaDeTransporte Is Nothing Then _ViaDeTransporte = New ViaDeTransporte(Me.CodigoViaDeTransporte)
            Return _ViaDeTransporte
        End Get
    End Property

    'ORIGEM
    Public Property CodigoOrigem() As String
        Get
            Return _CodigoOrigem
        End Get
        Set(ByVal value As String)
            _CodigoOrigem = value
        End Set
    End Property

    Public Property EndOrigem() As Integer
        Get
            Return _EndOrigem
        End Get
        Set(ByVal value As Integer)
            _EndOrigem = value
        End Set
    End Property

    Public Property Origem() As Cliente
        Get
            If _Origem Is Nothing Then _Origem = New Cliente(Me.CodigoOrigem, Me.EndOrigem)
            Return _Origem
        End Get
        Set(ByVal value As Cliente)
            _Origem = value
        End Set
    End Property

    Public Property OrigemCompleto() As String
        Get
            Return String.Format("{0}-{1}", _CodigoOrigem, _EndOrigem)
        End Get
        Set(ByVal value As String)
            _OrigemCompleto = value
        End Set
    End Property

    Public Property OrigemDescricao() As String
        Get
            Return String.Format("CPF/CNPJ: {0}<br/>{1}<br/>{2}/{3}", Funcoes.FormatarCpfCnpj(Origem.Codigo), Origem.Nome, Origem.Cidade, Origem.CodigoEstado)
        End Get
        Set(ByVal value As String)
            _OrigemDescricao = value
        End Set
    End Property

    'DESTINO
    Public Property CodigoDestino() As String
        Get
            Return _CodigoDestino
        End Get
        Set(ByVal value As String)
            _CodigoDestino = value
        End Set
    End Property

    Public Property EndDestino() As Integer
        Get
            Return _EndDestino
        End Get
        Set(ByVal value As Integer)
            _EndDestino = value
        End Set
    End Property

    Public Property Destino() As Cliente
        Get
            If _Destino Is Nothing Then _Destino = New Cliente(Me.CodigoDestino, Me.EndDestino)
            Return _Destino
        End Get
        Set(ByVal value As Cliente)
            _Destino = value
        End Set
    End Property

    Public Property DestinoCompleto() As String
        Get
            Return String.Format("{0}-{1}", _CodigoDestino, _EndDestino)
        End Get
        Set(ByVal value As String)
            _DestinoCompleto = value
        End Set
    End Property

    Public Property DestinoDescricao() As String
        Get
            Return String.Format("CPF/CNPJ: {0}<br/>{1}<br/>{2}/{3}", Funcoes.FormatarCpfCnpj(Destino.Codigo), Destino.Nome, Destino.Cidade, Destino.CodigoEstado)
        End Get
        Set(ByVal value As String)
            _DestinoDescricao = value
        End Set
    End Property

    'TRANSPORTADORES
    Public Property Transportadores() As ListEmbarqueTransportador
        Get
            If _Transportadores Is Nothing Then _Transportadores = New ListEmbarqueTransportador(Me)
            Return _Transportadores
        End Get
        Set(ByVal value As ListEmbarqueTransportador)
            _Transportadores = value
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
                sql = "INSERT INTO AutEmbarqueXRoteiro (Empresa_Id, EndEmpresa_Id, Pedido_Id, Roteiro_Id, Entrega, EndEntrega, Origem, EndOrigem, Destino, EndDestino, ViaDeTransporte) " & vbCrLf & _
                      "                         VALUES ('" & ParentEntrega.ParentEmbPedido.CodigoEmpresa & "', " & ParentEntrega.ParentEmbPedido.EndEmpresa & ", " & ParentEntrega.ParentEmbPedido.CodigoPedido & ", " & vbCrLf & _
                      "                                  (select ISNULL(MAX(roteiro_Id), 0) + 1 " & vbCrLf & _
                      "                                    From AutEmbarqueXRoteiro " & vbCrLf & _
                      "                                   Where Empresa_Id    = '" & ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "                                     And EndEmpresa_Id = " & ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "                                     And Pedido_Id     = " & ParentEntrega.ParentEmbPedido.CodigoPedido & "), " & vbCrLf & _
                      "                                 '" & ParentEntrega.CodigoClienteEntrega & "', " & ParentEntrega.EndClienteEntrega & ", '" & Me.CodigoOrigem & "', " & Me.EndOrigem & "," & vbCrLf & _
                      "                                 '" & Me.CodigoDestino & "', " & Me.EndDestino & "," & Me.CodigoViaDeTransporte & ")"
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = "UPDATE AutEmbarqueXRoteiro SET" & _
                      "     Origem           ='" & Me.CodigoOrigem & "'" & vbCrLf & _
                      "   , EndOrigem        = " & Me.EndOrigem & vbCrLf & _
                      "   , Destino          ='" & Me.CodigoDestino & "'" & vbCrLf & _
                      "   , EndDestino       = " & Me.EndDestino & vbCrLf & _
                      " WHERE Empresa_Id     ='" & ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id  = " & ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id      = " & ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Roteiro_Id     = " & Me.CodigoRoteiro
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                sql = "DELETE AutEmbarqueXRoteiro " & _
                      " WHERE Empresa_Id     ='" & ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id  = " & ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id      = " & ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Roteiro_Id     = " & Me.CodigoRoteiro
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Me.Transportadores IsNot Nothing AndAlso Me.Transportadores.Count > 0 Then
            For Each item As EmbarqueTransportador In Me.Transportadores
                If Me.IUD <> "D" Then
                    item.IUD = Me.IUD
                End If

                If Not String.IsNullOrWhiteSpace(item.IUD) Then
                    item.SalvarSql(Sqls)
                End If
            Next
        End If
    End Sub
#End Region

End Class
