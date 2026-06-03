Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMetasXSafras
    Inherits List(Of MetasXSafras)
End Class

<Serializable()> _
Public Class MetasXSafras
    Implements IBaseEntity

#Region "Fields"

    Private _IUD As String
    Private _CodigoMeta As Integer
    Private _CodigoSafra As String
    Private _CodigoCultura As Integer

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

    Public Property CodigoCultura() As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
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
                Sql = " INSERT INTO MetasXSafras (Meta_Id, Safra_Id, Cultura_Id) " & vbCrLf & _
                      "     VALUES (" & Me.CodigoMeta & ", '" & Me.CodigoSafra & "', " & Me.CodigoCultura & ")" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE MetasXSafras" & vbCrLf & _
                      "  WHERE      Meta_Id       = " & Me.CodigoMeta & vbCrLf & _
                      "     And     Safra_Id    = " & Str(Me.CodigoSafra) & vbCrLf & _
                      "     And     Cultura_Id    = " & Str(Me.CodigoCultura) & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

#End Region
End Class
