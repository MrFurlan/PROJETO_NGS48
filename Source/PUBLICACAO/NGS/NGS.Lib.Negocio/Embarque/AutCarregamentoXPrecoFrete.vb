Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAutCarregamentoXPrecoFrete
    Inherits List(Of AutCarregamentoXPrecoFrete)

    Public Sub New()

    End Sub

    Public Sub New(ByVal objCarregamento As AutCarregamento)
        Dim sql As String = ""
        sql &= ""

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "AutCarregamentoXPrecoFrete")

        For Each row In ds.Tables(0).Rows
            Dim obj As New AutCarregamentoXPrecoFrete
            obj.Empresa_Id = row("Empresa_Id")
            obj.EndEmpresa_Id = row("EndEmpresa_Id")
            obj.Pedido_Id = row("Pedido_Id")
            obj.Transportador_Id = row("Transportador_Id")
            obj.EndTransportar_Id = row("EndTransportar_Id")
            obj.Roteiro_Id = row("Roteiro_Id")
            obj.NrCotacao_Id = row("NrCotacao_Id")
            obj.Carregamento_Id = row("Carregamento_Id")
            obj.Quantidade = row("Quantidade")
            Me.Add(obj)
        Next
    End Sub

End Class

<Serializable()> _
Public Class AutCarregamentoXPrecoFrete
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String
    Private _Empresa_Id As String
    Private _EndEmpresa_Id As Integer
    Private _Pedido_Id As Integer
    Private _Transportador_Id As String
    Private _EndTransportar_Id As Integer
    Private _Transportador As Cliente
    Private _Roteiro_Id As Integer
    Private _NrCotacao_Id As Integer
    Private _Carregamento_Id As Integer
    Private _Quantidade As Decimal
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

    Public Property Empresa_Id() As String
        Get
            Return _Empresa_Id
        End Get
        Set(ByVal value As String)
            _Empresa_Id = value
        End Set
    End Property

    Public Property EndEmpresa_Id() As Integer
        Get
            Return _EndEmpresa_Id
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa_Id = value
        End Set
    End Property

    Public Property Pedido_Id() As Integer
        Get
            Return _Pedido_Id
        End Get
        Set(ByVal value As Integer)
            _Pedido_Id = value
        End Set
    End Property

    Public Property Transportador_Id() As String
        Get
            Return _Transportador_Id
        End Get
        Set(ByVal value As String)
            _Transportador_Id = value
        End Set
    End Property

    Public Property EndTransportar_Id() As Integer
        Get
            Return _EndTransportar_Id
        End Get
        Set(ByVal value As Integer)
            _EndTransportar_Id = value
        End Set
    End Property

    Public Property Transportador() As Cliente
        Get
            If _Transportador Is Nothing Then
                _Transportador = New Cliente(_Transportador_Id, _EndTransportar_Id)
            End If
            Return _Transportador
        End Get
        Set(ByVal value As Cliente)
            _Transportador = value
        End Set
    End Property

    Public Property NrCotacao_Id() As Integer
        Get
            Return _NrCotacao_Id
        End Get
        Set(ByVal value As Integer)
            _NrCotacao_Id = value
        End Set
    End Property

    Public Property Roteiro_Id() As Integer
        Get
            Return _Roteiro_Id
        End Get
        Set(ByVal value As Integer)
            _Roteiro_Id = value
        End Set
    End Property

    Public Property Carregamento_Id() As Integer
        Get
            Return _Carregamento_Id
        End Get
        Set(ByVal value As Integer)
            _Carregamento_Id = value
        End Set
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
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
                sql = "INSERT INTO AutEmbarqueXPrecoFrete (Empresa_Id, EndEmpresa_Id, Pedido_Id, Transportador_Id, EndTransportador_Id, Roteiro_Id, NrCotacao_Id, Carregamento_Id, Quantidade) " & vbCrLf & _
                      "VALUES ('" & Me.Empresa_Id & "', " & Me.EndEmpresa_Id & "," & Me.Pedido_Id & ",'" & Me.Transportador_Id & "', " & Me.EndTransportar_Id & ", " & Me.Roteiro_Id & ", " & Me.NrCotacao_Id & ", " & Me.Carregamento_Id & ", " & Str(Me.Quantidade) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE AutEmbarqueXPrecoFrete SET" & _
                      "     Quantidade = " & Str(Me.Quantidade) & "" & vbCrLf & _
                      " WHERE Empresa_Id    ='" & Me.Empresa_Id & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Me.EndEmpresa_Id & vbCrLf & _
                      "   AND Pedido_Id     = " & Me.Pedido_Id & vbCrLf & _
                      "   AND Transportador_Id     = '" & Me.Transportador_Id & "'" & vbCrLf & _
                      "   AND EndTransportador_Id     = " & Me.EndTransportar_Id & vbCrLf & _
                      "   AND Roteiro_Id     = " & Me.Roteiro_Id & vbCrLf & _
                      "   AND NrCotacao_Id = " & Me.Roteiro_Id & vbCrLf & _
                      "   AND Carregamento_Id = " & Me.Carregamento_Id
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE AutEmbarqueXPrecoFrete " & _
                      " WHERE Empresa_Id    ='" & Me.Empresa_Id & "'" & vbCrLf & _
                    "   AND EndEmpresa_Id = " & Me.EndEmpresa_Id & vbCrLf & _
                      "   AND Pedido_Id     = " & Me.Pedido_Id & vbCrLf & _
                      "   AND Transportador_Id     = '" & Me.Transportador_Id & "'" & vbCrLf & _
                      "   AND EndTransportador_Id     = " & Me.EndTransportar_Id & vbCrLf & _
                      "   AND Roteiro_Id     = " & Me.Roteiro_Id & vbCrLf & _
                      "   AND NrCotacao_Id = " & Me.Roteiro_Id & vbCrLf & _
                      "   AND Carregamento_Id = " & Me.Carregamento_Id
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
