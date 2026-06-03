Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

'*******************************************************************************************************************************************************************
'*******************************************  LISTA CLASSE BASE EMBARQUE X ENTREGA X ROTEIRO X TRANSPORTADOR *******************************************************
'*******************************************************************************************************************************************************************
<Serializable()> _
Public Class ListEmbarqueTransportador
    Inherits List(Of EmbarqueTransportador)

#Region "Construtor"
    Public Sub New(ByVal pEmbarquexRoteiro As EmbarqueRoteiro)
        _ParentRoteiro = pEmbarquexRoteiro

        Dim sql As String = ""
        sql &= "Select aet.Transportador_Id," & vbCrLf & _
               "       aet.EndTransportar_Id," & vbCrLf & _
               "       aet.Quota," & vbCrLf & _
               "       aet.Ativo," & vbCrLf & _
               "       isnull(aet.PesoQuantidade,'P') as PesoQuantidade" & vbCrLf & _
               "  From AutEmbarqueXTransportador aet " & vbCrLf & _
               " inner join AutEmbarqueXRoteiro aer" & vbCrLf & _
               "    on aet.Empresa_Id    = aer.Empresa_Id" & vbCrLf & _
               "   and aet.EndEmpresa_Id = aer.EndEmpresa_Id" & vbCrLf & _
               "   and aet.Pedido_Id     = aer.Pedido_Id" & vbCrLf & _
               "   and aet.Roteiro_Id    = aer.Roteiro_Id " & vbCrLf & _
               " where aet.Empresa_id    ='" & pEmbarquexRoteiro.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
               "   and aet.EndEmpresa_id = " & pEmbarquexRoteiro.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
               "   and aet.Pedido_id     = " & pEmbarquexRoteiro.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
               "   and aet.Roteiro_Id    ='" & pEmbarquexRoteiro.CodigoRoteiro & "'"


        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "EmbarqueTransportador")

        For Each row In ds.Tables(0).Rows
            Dim t As New EmbarqueTransportador(pEmbarquexRoteiro)
            t.CodigoTransportador = row("Transportador_Id")
            t.EndTransportador = row("EndTransportar_Id")
            t.Quota = row("Quota")
            t.Ativo = row("Ativo")
            t.PesoQuantidade = row("PesoQuantidade")
            Me.Add(t)
        Next
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each trans As Negocio.EmbarqueTransportador In Me
            If _ParentRoteiro.IUD = "D" Then trans.IUD = _ParentRoteiro.IUD
            If trans.IUD <> "" Then
                trans.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

#Region "Fields"
    Private _ParentRoteiro As EmbarqueRoteiro
#End Region

#Region "Property"
    Public ReadOnly Property ParentRoteiro As EmbarqueRoteiro
        Get
            Return _ParentRoteiro
        End Get
    End Property
#End Region

End Class


'*******************************************************************************************************************************************************************
'*******************************************   CLASSE BASE EMBARQUE X ENTREGA X ROTEIRO X TRANSPORTADOR ************************************************************
'*******************************************************************************************************************************************************************
<Serializable()> _
Public Class EmbarqueTransportador
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pEmbarquexRoteiro As EmbarqueRoteiro)
        _ParentRoteiro = pEmbarquexRoteiro
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _ParentRoteiro As EmbarqueRoteiro
    Private _CodigoTransportador As String
    Private _EndTransportador As Integer
    Private _Transportador As Cliente
    Private _Quota As Decimal
    Private _Ativo As Boolean
    Private _PesoQuantidade As String
    Private _ViaTransporte As Integer
    Private _Precos As ListEmbarquePrecoFrete
    Private _CnpjTransp As String
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

    Public ReadOnly Property ParentRoteiro As EmbarqueRoteiro
        Get
            Return _ParentRoteiro
        End Get
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

    Public Property Transportador() As Cliente
        Get
            If _Transportador Is Nothing Then
                _Transportador = New Cliente(_CodigoTransportador, _EndTransportador)
            End If
            Return _Transportador
        End Get
        Set(ByVal value As Cliente)
            _Transportador = value
        End Set
    End Property

    Public Property Quota() As Decimal
        Get
            Return _Quota
        End Get
        Set(ByVal value As Decimal)
            _Quota = value
        End Set
    End Property

    Public Property Ativo() As Boolean
        Get
            Return _Ativo
        End Get
        Set(ByVal value As Boolean)
            _Ativo = value
        End Set
    End Property

    Public Property PesoQuantidade As String
        Get
            Return _PesoQuantidade
        End Get
        Set(value As String)
            _PesoQuantidade = value
        End Set
    End Property

    Public ReadOnly Property DescPesoQuantidade As String
        Get
            If _PesoQuantidade = "P" Then
                Return ParentRoteiro.ParentEntrega.ParentEmbPedido.Produto.Unidade
            Else
                Return "FRETES"
            End If
        End Get
    End Property

    Public Property CnpjTransp() As String
        Get
            Return String.Format("{0}-{1}", CodigoTransportador, EndTransportador)
        End Get
        Set(ByVal value As String)
            _CnpjTransp = value
        End Set
    End Property

    Public Property Precos() As ListEmbarquePrecoFrete
        Get
            If _Precos Is Nothing Then _Precos = New ListEmbarquePrecoFrete(Me)
            Return _Precos
        End Get
        Set(ByVal value As ListEmbarquePrecoFrete)
            _Precos = value
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
                sql = "INSERT INTO AutEmbarqueXTransportador (Empresa_Id, EndEmpresa_Id, Pedido_Id, Transportador_Id, EndTransportar_Id, Roteiro_Id, Quota, Ativo, PesoQuantidade) " & vbCrLf & _
                      "VALUES ('" & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "', " & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.EndEmpresa & "," & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoPedido & ",'" & Me.CodigoTransportador & "', " & Me.EndTransportador & ", " & Me.ParentRoteiro.CodigoRoteiro & "," & Str(Me.Quota) & ", " & IIf(Me.Ativo, 1, 0) & ",'" & Me.PesoQuantidade & "')"
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = "UPDATE AutEmbarqueXTransportador SET" & _
                      "     Quota          = " & Str(Me.Quota) & vbCrLf & _
                      "   , Ativo          = " & IIf(Me.Ativo, 1, 0) & vbCrLf & _
                      "   , PesoQuantidade ='" & Me.PesoQuantidade & "'" & vbCrLf & _
                      " WHERE Empresa_Id        ='" & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id     = " & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id         = " & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Transportador_Id  ='" & Me.CodigoTransportador & "'" & vbCrLf & _
                      "   AND EndTransportar_Id = " & Me.EndTransportador & vbCrLf & _
                      "   AND Roteiro_Id        = " & Me.ParentRoteiro.CodigoRoteiro
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                sql = "DELETE AutEmbarqueXTransportador " & _
                      " WHERE Empresa_Id        ='" & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id     = " & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id         = " & Me.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Transportador_Id  ='" & Me.CodigoTransportador & "'" & vbCrLf & _
                      "   AND EndTransportar_Id = " & Me.EndTransportador & vbCrLf & _
                      "   AND Roteiro_Id        = " & Me.ParentRoteiro.CodigoRoteiro
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Me.Precos IsNot Nothing AndAlso Me.Precos.Count > 0 Then
            For Each item As EmbarquePrecoFrete In Me.Precos
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

