
<Serializable()> _
Public Class ListGruposXEmpresas
    Inherits List(Of GruposXEmpresas)

End Class


<Serializable()> _
Public Class GruposXEmpresas
    Implements IBaseEntity


#Region "Variáveis locais"

    Private _CodigoUnidade As String
    Private _EnderecoUnidade As Integer
    Private _CodigoEmpresa As String
    Private _EnderecoEmpresa As Integer
    Private _Nome As String

#End Region

#Region "Propriedades"

    Public Property CodigoUnidade() As String
        Get
            Return _CodigoUnidade
        End Get
        Set(ByVal value As String)
            _CodigoUnidade = value
        End Set
    End Property
    Public Property EnderecoUnidade() As Integer
        Get
            Return _EnderecoUnidade
        End Get
        Set(ByVal value As Integer)
            _EnderecoUnidade = value
        End Set
    End Property
    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property
    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
        End Set
    End Property

    Public Property Nome() As String
        Get
            Return _Nome
        End Get
        Set(ByVal value As String)
            _Nome = value
        End Set
    End Property

#End Region

#Region "Construtores"
    Public Sub New()

    End Sub

    Sub New(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer)
        Dim objBanco As New AcessaBanco()

        Try
            Dim Sql As String = " Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Descricao " & vbCrLf & _
                                "   FROM GruposXEmpresas " & vbCrLf & _
                                " WHERE Cliente_Id    = '" & CodigoEmpresa & "'" & vbCrLf & _
                                "   AND EndCliente_Id = " & EndEmpresa
           
            Dim ds As DataSet = objBanco.ConsultaDataSet(Sql, "GruposXEmpresas")

            If ds.Tables(0).Rows.Count > 0 Then
                _CodigoUnidade = ds.Tables(0).Rows(0)("Empresa_Id")
                _EnderecoUnidade = ds.Tables(0).Rows(0)("EndEmpresa_Id")
                _CodigoEmpresa = ds.Tables(0).Rows(0)("Cliente_Id")
                _EnderecoEmpresa = ds.Tables(0).Rows(0)("EndCliente_Id")
                _Nome = ds.Tables(0).Rows(0)("Descricao")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            objBanco = Nothing
        End Try
    End Sub

#End Region

#Region "Métodos"

#End Region

End Class
