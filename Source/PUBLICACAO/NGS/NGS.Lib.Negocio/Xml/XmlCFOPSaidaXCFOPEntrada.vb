Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

<Serializable()> _
Public Class ListXmlCFOPSaidaXCFOPEntrada
    Inherits List(Of XmlCFOPSaidaXCFOPEntrada)
    Implements IBaseEntity

    Public Erro As Exception

#Region "Contructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoCFOPSaida As Integer, Optional ByVal pCodigoCFOPEntrada As String = "")
        Dim sql As String = String.Empty

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        sql = "SELECT CFOPSaida_Id, CFOPEntrada_Id" & vbCrLf & _
              "  FROM XmlCFOPSaidaXCFOPEntrada" & vbCrLf
        If Not String.IsNullOrWhiteSpace(pCodigoCFOPSaida) Then
            sql &= " WHERE CFOPSaida_Id = " & pCodigoCFOPSaida
        End If

        If Not String.IsNullOrWhiteSpace(pCodigoCFOPEntrada) Then
            sql &= " WHERE CFOPEntrada_Id = " & pCodigoCFOPEntrada
        End If

        ds = Banco.ConsultaDataSet(sql, "XmlCFOPSaidaXCFOPEntrada")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ObjCfop As New XmlCFOPSaidaXCFOPEntrada()
            ObjCfop.CodigoCFOPSaida = row("CFOPSaida_Id")
            ObjCfop.CodigoCFOPSaida = row("CFOPEntrada_Id")
            Me.Add(ObjCfop)
        Next

    End Sub

#End Region
End Class


Public Class XmlCFOPSaidaXCFOPEntrada

#Region "Constructors"
    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoCFOPSaida As Integer
    Private _CFOPSaida As CFOP
    Private _CodigoCFOPEntrada As Integer
    Private _CFOPEntrada As CFOP
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

    Public Property CodigoCFOPSaida() As Integer
        Get
            Return _CodigoCFOPSaida
        End Get
        Set(ByVal value As Integer)
            _CodigoCFOPSaida = value
        End Set
    End Property

    Public Property CFOPSaida() As CFOP
        Get
            Return _CFOPSaida
        End Get
        Set(ByVal value As CFOP)
            _CFOPSaida = value
        End Set
    End Property

    Public Property CodigoCFOPEntrada() As Integer
        Get
            Return _CodigoCFOPEntrada
        End Get
        Set(ByVal value As Integer)
            _CodigoCFOPEntrada = value
        End Set
    End Property

    Public Property CFOPEntrada() As CFOP
        Get
            Return _CFOPEntrada
        End Get
        Set(ByVal value As CFOP)
            _CFOPEntrada = value
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
                Sql = " INSERT INTO XmlCFOPSaidaXCFOPEntrada(CFOPSaida_Id, CFOPEntrada_Id) " & vbCrLf & _
                      " VALUES (" & Me._CodigoCFOPSaida & "," & Me.CodigoCFOPEntrada & ")"
                Sqls.Add(Sql)
                'Case "U"
                '    Sql = " "
                '    Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE XmlCFOPSaidaXCFOPEntrada" & vbCrLf & _
                      "  WHERE CFOPSaida_Id   = " & Me.CodigoCFOPSaida & vbCrLf & _
                      "    AND CFOPEntrada_Id = " & Me.CodigoCFOPEntrada & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
