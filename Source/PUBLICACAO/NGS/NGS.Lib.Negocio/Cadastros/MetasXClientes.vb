Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMetasXClientes
    Inherits List(Of MetasXClientes)
End Class

<Serializable()> _
Public Class MetasXClientes
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoMeta As Integer
    Private _CodigoSafra As String = ""
    Private _CodigoCultura As Integer
    Private _CodigoCliente As String = ""
    Private _SocVenda As Decimal
#End Region

#Region "Propriedades"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoMeta() As Integer
        Get
            Return _CodigoMeta
        End Get
        Set(ByVal value As Integer)
            _CodigoMeta = value
        End Set
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property CodigoCultura() As String
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As String)
            _CodigoCultura = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property SocVenda() As Decimal
        Get
            Return _SocVenda
        End Get
        Set(ByVal value As Decimal)
            _SocVenda = value
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
                Sql = " INSERT INTO MetasXClientes (Meta_Id, Safra_Id, Cultura_Id, Cliente_Id, Soc_Meta, Soc_Venda)" & vbCrLf & _
                      " VALUES (" & Me.CodigoMeta & ", '" & Me.CodigoSafra & "', " & Me.CodigoCultura & ", '" & Me.CodigoCliente & "', " & Me.SocVenda & ")" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE MetasXClientes SET" & vbCrLf & _
                      "    Soc_Venda         = " & Me.SocVenda & vbCrLf & _
                      "  WHERE  Meta_Id      = " & Me.CodigoMeta & vbCrLf & _
                      "     And Safra_Id     = " & Str(Me.CodigoSafra) & vbCrLf & _
                      "     And Cultura_Id   = " & Me.CodigoCultura & vbCrLf & _
                      "     And Cliente_Id   = " & Str(Me.CodigoCliente) & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE MetasXClientes" & vbCrLf & _
                      "  WHERE  Meta_Id      = " & Me.CodigoMeta & vbCrLf & _
                      "     And Safra_Id     = " & Str(Me.CodigoSafra) & vbCrLf & _
                      "     And Cultura_Id   = " & Me.CodigoCultura & vbCrLf & _
                      "     And Cliente_Id   = " & Str(Me.CodigoCliente) & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
