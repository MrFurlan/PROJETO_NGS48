
Imports System.Web
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListClienteXSubstitutoTributario
    Inherits List(Of ClienteXSubstitutoTributario)

#Region "Construtor"
    Public Sub New(ByVal pCliente As Negocio.Cliente)

        _Cliente = pCliente

        If String.IsNullOrWhiteSpace(_Cliente.Codigo) Then Exit Sub

        Dim Banco As New AcessaBanco
        Try
            Dim SQL As String = ""
            SQL = "SELECT Estado_Id, IESubstitutoTributario " &
                  "  FROM ClienteXSubstitutoTributario " &
                  " WHERE Cliente_Id    ='" & Me.Cliente.Codigo & "' " &
                  "   AND EndCliente_Id = " & Me.Cliente.CodigoEndereco.ToString()

            Dim ds As DataSet = Banco.ConsultaDataSet(SQL, "SubstitutoTributario")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim cXst As New Negocio.ClienteXSubstitutoTributario(Cliente)

                cXst.Estado_Id = row("Estado_Id")
                cXst.IESubstitutoTributario = row("IESubstitutoTributario")

                Me.Add(cXst)
            Next
        Catch ex As Exception
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Field"
    Private _Cliente As Negocio.Cliente
#End Region

#Region "Property"
    Public Property Cliente As Negocio.Cliente
        Get
            Return _Cliente
        End Get
        Set(value As Negocio.Cliente)
            _Cliente = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each pXst As ClienteXSubstitutoTributario In Me
            pXst.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

<Serializable()>
Public Class ClienteXSubstitutoTributario

#Region "Contrutor"
    Public Sub New(ByVal pCliente As Negocio.Cliente)
        _Cliente = pCliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Negocio.Cliente
    Private _Estado_Id As String
    Private _IESubstitutoTributario As String
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As String
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As String
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property Cliente As Negocio.Cliente
        Get
            Return _Cliente
        End Get
    End Property

    Public Property Estado_Id() As String
        Get
            Return _Estado_Id
        End Get
        Set(ByVal value As String)
            _Estado_Id = value
        End Set
    End Property

    Public Property IESubstitutoTributario() As String
        Get
            Return _IESubstitutoTributario
        End Get
        Set(ByVal value As String)
            _IESubstitutoTributario = value
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

    Public Property UsuarioInclusaoData As String
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(value As String)
            _UsuarioInclusaoData = value
        End Set
    End Property

    Public Property UsuarioAlteracao As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property UsuarioAlteracaoData As String
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(value As String)
            _UsuarioAlteracaoData = value
        End Set
    End Property
#End Region

#Region "Methods"

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = "INSERT INTO ClienteXSubstitutoTributario " & vbCrLf &
                      "(Cliente_Id, EndCliente_Id, Estado_Id, IESubstitutoTributario, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf &
                      " VALUES('" & Me.Cliente.Codigo & "'," & Me.Cliente.CodigoEndereco & "," & vbCrLf &
                              "'" & Me.Estado_Id & "'," & vbCrLf &
                              "'" & Me.IESubstitutoTributario & "'," & vbCrLf &
                              "'" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                              "'" & Now().ToString("yyyy-MM-dd") & "')" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update ClienteXSubstitutoTributario Set" & vbCrLf &
                              "IESubstitutoTributario = '" & Me.IESubstitutoTributario & "'," & vbCrLf &
                              "UsuarioAlteracao       = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                              "UsuarioAlteracaoData   = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                       "WHERE Cliente_Id              = '" & Me.Cliente.Codigo & "'" & vbCrLf &
                       "  AND EndCliente_Id           =  " & Me.Cliente.CodigoEndereco & vbCrLf &
                       "  AND Estado_Id               = '" & Me.Estado_Id & "'" &
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE ClienteXSubstitutoTributario " &
                      " WHERE Cliente_Id     ='" & Me.Cliente.Codigo & "'" & vbCrLf &
                      "   AND EndCliente_Id  = " & Me.Cliente.CodigoEndereco & vbCrLf &
                      "   AND Estado_Id      = '" & Me.Estado_Id & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
