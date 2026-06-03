Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

'******************************************************************************************************************************************
'*******************************************  LISTA CLASSE BASE EMBARQUE X ENTREGA  *******************************************************
'******************************************************************************************************************************************
<Serializable()> _
Public Class ListEmbarqueXEntrega
    Inherits List(Of EmbarqueXEntrega)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(pEmbarquePedido As EmbarquePedido)
        _ParentEmbPedido = pEmbarquePedido

        Dim sql As String = ""

        sql &= "Select Empresa_Id, EndEmpresa_Id, Pedido_Id, Entrega_Id, EndEntrega_Id, EmitirNota, Observacao" & vbCrLf & _
               "  from AutEmbarqueXEntrega" & vbCrLf & _
               " where Empresa_Id    ='" & pEmbarquePedido.CodigoEmpresa & "'" & vbCrLf & _
               "   and EndEmpresa_Id = " & pEmbarquePedido.EndEmpresa & vbCrLf & _
               "   and Pedido_Id     = " & pEmbarquePedido.CodigoPedido & vbCrLf

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "EmbarqueXEntrega")

        For Each row In ds.Tables(0).Rows
            Dim obj As New EmbarqueXEntrega(pEmbarquePedido)
            obj.CodigoClienteEntrega = row("Entrega_Id")
            obj.EndClienteEntrega = row("EndEntrega_Id")
            obj.EmitirNota = row("EmitirNota")
            obj.Observacao = row("Observacao")
            Me.Add(obj)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _ParentEmbPedido As EmbarquePedido
#End Region

#Region "Property"
    Public ReadOnly Property ParentEmbPedido As EmbarquePedido
        Get
            Return _ParentEmbPedido
        End Get
    End Property
#End Region

End Class

'******************************************************************************************************************************************
'************************************************  CLASS BASE EMBARQUE X ENTREGA  *********************************************************
'******************************************************************************************************************************************
<Serializable()> _
Public Class EmbarqueXEntrega
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New(pEmbarquePedido As EmbarquePedido)
        _ParentEmbPedido = pEmbarquePedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _ParentEmbPedido As EmbarquePedido
    Private _CodigoClienteEntrega As String
    Private _EndClienteEntrega As Integer
    Private _ClienteEntrega As Cliente
    Private _DescClienteEntrega As String = ""
    Private _EmitirNota As Boolean = False
    Private _Observacao As String = ""
    Private _Placa As String = ""

    Private _Roteiros As ListEmbarqueRoteiro
    Private _Produtos As ListEmbarqueXEntregaProduto
    Private _ListCarregamentos As ListAutCarregamento
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

    Public ReadOnly Property ParentEmbPedido As EmbarquePedido
        Get
            Return _ParentEmbPedido
        End Get
    End Property

    Public Property CodigoClienteEntrega() As String
        Get
            Return _CodigoClienteEntrega
        End Get
        Set(ByVal value As String)
            _CodigoClienteEntrega = value
        End Set
    End Property

    Public Property EndClienteEntrega() As Integer
        Get
            Return _EndClienteEntrega
        End Get
        Set(ByVal value As Integer)
            _EndClienteEntrega = value
        End Set
    End Property

    Public ReadOnly Property DescClienteEntrega As String
        Get
            If String.IsNullOrWhiteSpace(_DescClienteEntrega) Then _DescClienteEntrega = ClienteEntrega.Cidade & "-" & ClienteEntrega.CodigoEstado & "-" & ClienteEntrega.Nome
            Return _DescClienteEntrega
        End Get

    End Property

    Public ReadOnly Property ClienteEntrega() As Cliente
        Get
            If _ClienteEntrega Is Nothing And Me.CodigoClienteEntrega.Length > 0 Then _ClienteEntrega = New Cliente(Me.CodigoClienteEntrega, Me.EndClienteEntrega)
            Return _ClienteEntrega
        End Get
    End Property

    Public ReadOnly Property EntregaFormatado() As String
        Get
            Return ClienteEntrega.Cidade & " - " & ClienteEntrega.Estado.Descricao & " - " & ClienteEntrega.Nome & " - " & Funcoes.FormatarCpfCnpj(ClienteEntrega.Codigo)
        End Get
    End Property

    Public Property EmitirNota As Boolean
        Get
            Return _EmitirNota
        End Get
        Set(value As Boolean)
            _EmitirNota = value
        End Set
    End Property

    Public Property Observacao As String
        Get
            Return _Observacao
        End Get
        Set(value As String)
            _Observacao = value
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

    Public Property Roteiros() As ListEmbarqueRoteiro
        Get
            If _Roteiros Is Nothing Then _Roteiros = New ListEmbarqueRoteiro(Me)
            Return _Roteiros
        End Get
        Set(ByVal value As ListEmbarqueRoteiro)
            _Roteiros = value
        End Set
    End Property

    Public Property Produtos As ListEmbarqueXEntregaProduto
        Get
            If _Produtos Is Nothing Then _Produtos = New ListEmbarqueXEntregaProduto(Me)
            Return _Produtos
        End Get
        Set(value As ListEmbarqueXEntregaProduto)
            _Produtos = value
        End Set
    End Property

    Public Property ListCarregamentos() As ListAutCarregamento
        Get
            If _ListCarregamentos Is Nothing Then _ListCarregamentos = New ListAutCarregamento(Me)
            Return _ListCarregamentos
        End Get
        Set(ByVal value As ListAutCarregamento)
            _ListCarregamentos = value
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
                sql = "INSERT INTO AutEmbarqueXEntrega (Empresa_Id, EndEmpresa_Id, Pedido_Id, Entrega_Id, EndEntrega_Id, EmitirNota, Observacao) " & vbCrLf & _
                      "VALUES ('" & ParentEmbPedido.CodigoEmpresa & "', " & ParentEmbPedido.EndEmpresa & ", " & ParentEmbPedido.CodigoPedido & ", '" & Me.CodigoClienteEntrega & "', " & Me.EndClienteEntrega & "," & IIf(Me.EmitirNota, 1, 0) & ",'" & Me.Observacao & "')"
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = "Update AutEmbarqueXEntrega set" & vbCrLf & _
                      "   Observacao ='" & Me.Observacao & "'" & vbCrLf & _
                      " WHERE Empresa_Id    ='" & ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id     = " & ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Entrega_Id    ='" & Me.CodigoClienteEntrega & "'" & vbCrLf & _
                      "   AND EndEntrega_Id = " & Me.EndClienteEntrega
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                sql = "DELETE AutEmbarqueXEntrega " & _
                      " WHERE Empresa_Id    ='" & ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id     = " & ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Entrega_Id    ='" & Me.CodigoClienteEntrega & "'" & vbCrLf & _
                      "   AND EndEntrega_Id = " & Me.EndClienteEntrega
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Me.Roteiros IsNot Nothing AndAlso Me.Roteiros.Count > 0 Then
            For Each item As EmbarqueRoteiro In Me.Roteiros
                If IUD = "D" Then
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
