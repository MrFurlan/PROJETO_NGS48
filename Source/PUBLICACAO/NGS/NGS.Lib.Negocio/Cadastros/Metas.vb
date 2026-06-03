Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMetas
    Inherits List(Of Metas)
End Class

<Serializable()> _
Public Class Metas
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoMeta As Integer
    Private _Empresa As String = ""
    Private _EndEmpresa As Integer
    Private _Ano As Integer
    Private _CotacaoDollar As Decimal
#End Region

#Region "Propierts"
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

    Public Property Empresa() As String
        Get
            Return _Empresa
        End Get
        Set(ByVal value As String)
            _Empresa = value
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property CotacaoDollar() As Decimal
        Get
            Return _CotacaoDollar
        End Get
        Set(ByVal value As Decimal)
            _CotacaoDollar = value
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
                Sql = " INSERT INTO Metas (Meta_Id, Empresa, End_Empresa, Ano, CotacaoDollar) " & vbCrLf & _
                      " VALUES ('" & Me.CodigoMeta & ", " & Me.Empresa & "', '" & Me.EndEmpresa & "', " & Me.Ano & ", " & Me.CotacaoDollar & ")" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Metas SET" & vbCrLf & _
                      "    Empresa           = " & Me.Empresa & vbCrLf & _
                      "    End_Empresa       = " & Me.EndEmpresa & vbCrLf & _
                      "    Ano               = " & Me.Ano & vbCrLf & _
                      "    CotacaoDollar     = " & Me.CotacaoDollar & vbCrLf & _
                      "  WHERE      Meta_Id  = " & Me.CodigoMeta & vbCrLf

                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Metas" & vbCrLf & _
                      "  WHERE      Meta_Id  = " & Me.CodigoMeta & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
