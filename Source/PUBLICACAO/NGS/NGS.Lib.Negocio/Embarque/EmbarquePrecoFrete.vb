Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

'***************************************************************************************************************************************************************************
'*******************************************  LISTA CLASSE BASE EMBARQUE X ENTREGA X ROTEIRO X TRANSPORTADOR X PRECO *******************************************************
'***************************************************************************************************************************************************************************
<Serializable()> _
Public Class ListEmbarquePrecoFrete
    Inherits List(Of EmbarquePrecoFrete)

#Region "Construtor"
    Public Sub New(ByVal pEmbarqueTransportador As EmbarqueTransportador)
        _ParentTransportador = pEmbarqueTransportador

        Dim sql As String = ""

        sql &= "Select aep.NrCotacao_Id," & vbCrLf & _
               "       aep.Movimento," & vbCrLf & _
               "       aep.ValorFrete," & vbCrLf & _
               "       aep.Quota," & vbCrLf & _
               "       aep.ValorTon, " & vbCrLf & _
               "       aep.UsuarioInclusao" & vbCrLf & _
               "  From AutEmbarqueXPrecoFrete aep " & vbCrLf & _
               " Inner join AutEmbarqueXTransportador aet" & vbCrLf & _
               "    on aet.Empresa_Id        = aep.Empresa_Id  " & vbCrLf & _
               "   and aet.EndEmpresa_Id     = aep.EndEmpresa_Id  " & vbCrLf & _
               "   and aet.Pedido_Id         = aep.Pedido_Id  " & vbCrLf & _
               "   and aet.Transportador_Id  = aep.Transportador_Id  " & vbCrLf & _
               "   and aet.EndTransportar_Id = aep.EndTransportador_Id  " & vbCrLf & _
               "   and aet.Roteiro_Id        = aep.Roteiro_Id " & vbCrLf & _
               " where aep.Empresa_Id          ='" & pEmbarqueTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
               "   and aep.EndEmpresa_Id       = " & pEmbarqueTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
               "   and aep.Pedido_Id           = " & pEmbarqueTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
               "   and aep.Transportador_Id    ='" & pEmbarqueTransportador.CodigoTransportador & "'" & vbCrLf & _
               "   and aep.EndTransportador_Id = " & pEmbarqueTransportador.EndTransportador & vbCrLf & _
               "   and aep.Roteiro_Id          = " & pEmbarqueTransportador.ParentRoteiro.CodigoRoteiro & vbCrLf


        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "EmbarquePrecoFrete")

        For Each row In ds.Tables(0).Rows
            Dim p As New EmbarquePrecoFrete(pEmbarqueTransportador)
            p.NrCotacao = row("NrCotacao_Id")
            p.Movimento = row("Movimento")
            p.ValorFrete = row("ValorFrete")
            p.Quota = row("Quota")
            p.ValorTon = row("ValorTon")
            p.UsuarioInclusao = row("UsuarioInclusao")
            Me.Add(p)
        Next
    End Sub

    Public Sub New(ByVal pAutCarregamento As AutCarregamento)
        _ParentAutCarregamento = pAutCarregamento

        Dim sql As String = ""

        sql &= "Select aep.NrCotacao_Id," & vbCrLf & _
               "       aep.Movimento," & vbCrLf & _
               "       aep.ValorFrete," & vbCrLf & _
               "       aep.Quota," & vbCrLf & _
               "       aep.ValorTon " & vbCrLf & _
               "  From AutEmbarqueXPrecoFrete aep " & vbCrLf & _
               " Inner join AutEmbarqueXTransportador aet" & vbCrLf & _
               "    on aet.Empresa_Id        = aep.Empresa_Id  " & vbCrLf & _
               "   and aet.EndEmpresa_Id     = aep.EndEmpresa_Id  " & vbCrLf & _
               "   and aet.Pedido_Id         = aep.Pedido_Id  " & vbCrLf & _
               "   and aet.Transportador_Id  = aep.Transportador_Id  " & vbCrLf & _
               "   and aet.EndTransportar_Id = aep.EndTransportador_Id  " & vbCrLf & _
               "   and aet.Roteiro_Id        = aep.Roteiro_Id " & vbCrLf & _
               " where aep.Empresa_Id          ='" & pAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
               "   and aep.EndEmpresa_Id       = " & pAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
               "   and aep.Pedido_Id           = " & pAutCarregamento.ParentEmbarqueEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
               "   and aep.Transportador_Id    ='" & pAutCarregamento.CodigoTransportador & "'" & vbCrLf & _
               "   and aep.EndTransportador_Id = " & pAutCarregamento.EndTransportador & vbCrLf


        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "EmbarquePrecoFrete")

        For Each row In ds.Tables(0).Rows
            Dim p As New EmbarquePrecoFrete(pAutCarregamento)
            p.NrCotacao = row("NrCotacao_Id")
            p.Movimento = row("Movimento")
            p.ValorFrete = row("ValorFrete")
            p.Quota = row("Quota")
            p.ValorTon = row("ValorTon")
            Me.Add(p)
        Next
    End Sub

#End Region

#Region "Fields"
    Private _ParentTransportador As EmbarqueTransportador
    Private _ParentAutCarregamento As AutCarregamento
#End Region

#Region "Property"
    Public ReadOnly Property ParentTransportador As EmbarqueTransportador
        Get
            Return _ParentTransportador
        End Get
    End Property
#End Region

End Class

'***************************************************************************************************************************************************************************
'*******************************************   CLASSE BASE EMBARQUE X ENTREGA X ROTEIRO X TRANSPORTADOR x PRECO ************************************************************
'***************************************************************************************************************************************************************************
<Serializable()> _
Public Class EmbarquePrecoFrete
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New(ByVal pEmbarqueTransportador As EmbarqueTransportador)
        _ParentTransportador = pEmbarqueTransportador
    End Sub
    Public Sub New(ByVal pAutCarregamento As AutCarregamento)
        _ParentAutCarregamento = pAutCarregamento
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _ParentTransportador As EmbarqueTransportador
    Private _ParentAutCarregamento As AutCarregamento
    Private _NrCotacao As Integer
    Private _Movimento As DateTime
    Private _ValorFrete As Decimal
    Private _Quota As Decimal
    Private _ValorTon As Decimal
    Private _UsuarioInclusao As String = ""
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

    Public ReadOnly Property ParentTransportador As EmbarqueTransportador
        Get
            Return _ParentTransportador
        End Get
    End Property

    Public ReadOnly Property ParentAutCarregamento As AutCarregamento
        Get
            Return _ParentAutCarregamento
        End Get
    End Property

    Public Property NrCotacao() As Integer
        Get
            Return _NrCotacao
        End Get
        Set(ByVal value As Integer)
            _NrCotacao = value
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

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property ValorFrete() As Decimal
        Get
            Return _ValorFrete
        End Get
        Set(ByVal value As Decimal)
            _ValorFrete = value
        End Set
    End Property

    Public Property ValorTon() As Decimal
        Get
            Return _ValorTon
        End Get
        Set(ByVal value As Decimal)
            _ValorTon = value
        End Set
    End Property

    Public Property UsuarioInclusao As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(value As String)
            _UsuarioInclusao = value
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
                sql = "INSERT INTO AutEmbarqueXPrecoFrete " & vbCrLf & _
                    "(Empresa_Id, EndEmpresa_Id, Pedido_Id, Transportador_Id, EndTransportador_Id, Roteiro_Id, NrCotacao_Id, Movimento, ValorFrete, Quota, ValorTon, UsuarioInclusao) " & vbCrLf & _
                    "VALUES ('" & Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'," & vbCrLf & _
                    Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.EndEmpresa & "," & vbCrLf & _
                    Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoPedido & "," & vbCrLf & _
                    "'" & Me.ParentTransportador.CodigoTransportador & "'," & vbCrLf & _
                    Me.ParentTransportador.EndTransportador & "," & vbCrLf & _
                    Me.ParentTransportador.ParentRoteiro.CodigoRoteiro & "," & vbCrLf & _
                    Me.NrCotacao & "," & vbCrLf & _
                    "'" & Me.Movimento.ToString("yyyy-MM-dd") & "', " & Str(ValorFrete) & ", " & Str(Me.Quota) & ", " & Str(Me.ValorTon) & ",'" & Me.UsuarioInclusao & "')" & vbCrLf
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE AutEmbarqueXPrecoFrete SET" & _
                      "     Quota      = " & Str(Me.Quota) & "" & vbCrLf & _
                      "   , Movimento  ='" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   , ValorFrete = " & Str(Me.ValorFrete) & vbCrLf & _
                      "   , ValorTon   = " & Str(Me.ValorTon) & vbCrLf & _
                      " WHERE Empresa_Id          ='" & Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id       = " & Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id           = " & Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Transportador_Id    ='" & Me.ParentTransportador.CodigoTransportador & "'" & vbCrLf & _
                      "   AND EndTransportador_Id = " & Me.ParentTransportador.EndTransportador & vbCrLf & _
                      "   AND Roteiro_Id          = " & Me.ParentTransportador.ParentRoteiro.CodigoRoteiro & vbCrLf & _
                      "   And NrCotacao_Id        = " & Me.NrCotacao & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE AutEmbarqueXPrecoFrete " & _
                      " WHERE Empresa_Id          ='" & Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id       = " & Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.EndEmpresa & vbCrLf & _
                      "   AND Pedido_Id           = " & Me.ParentTransportador.ParentRoteiro.ParentEntrega.ParentEmbPedido.CodigoPedido & vbCrLf & _
                      "   AND Transportador_Id    ='" & Me.ParentTransportador.CodigoTransportador & "'" & vbCrLf & _
                      "   AND EndTransportador_Id = " & Me.ParentTransportador.EndTransportador & vbCrLf & _
                      "   AND Roteiro_Id          = " & Me.ParentTransportador.ParentRoteiro.CodigoRoteiro & vbCrLf & _
                      "   And NrCotacao_Id        = " & Me.NrCotacao & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class