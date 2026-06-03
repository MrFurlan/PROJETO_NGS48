Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListFitoXEpi
    Inherits List(Of FitoXEpi)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoFito As Integer)

        Dim sql As String = "Select Fito_Id, EPI_Id From FitoxEPI Where Fito_Id = " & pCodigoFito
        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "FxE")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim FxE As New FitoXEpi
            FxE.CodigoFito = row("Fito_Id")
            FxE.CodigoEPI = row("EPI_Id")
            Me.Add(FxE)
        Next

    End Sub

End Class

<Serializable()> _
Public Class FitoXEpi

#Region "Fields"
    Private _IUD As String
    Private _CodigoFito As Integer
    Private _EPI As EPI
    Private _CodigoEPI As Integer
#End Region

#Region "Property"
    Public Property IUD()
        Get
            Return _IUD
        End Get
        Set(ByVal value)
            _IUD = value
        End Set
    End Property

    Public Property CodigoFito() As Integer
        Get
            Return _CodigoFito
        End Get
        Set(ByVal value As Integer)
            _CodigoFito = value
        End Set
    End Property

    Public ReadOnly Property EPI() As EPI
        Get
            If _EPI Is Nothing And _CodigoFito > 0 Then _EPI = New EPI(_CodigoEPI)
            Return _EPI
        End Get
    End Property

    Public Property CodigoEPI() As Integer
        Get
            Return _CodigoEPI
        End Get
        Set(ByVal value As Integer)
            _CodigoEPI = value
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
                Sql = " INSERT INTO FitoXEPI (Fito_Id,EPI_Id) " & vbCrLf & _
                      " VALUES (" & _CodigoFito & ",'" & _CodigoEPI & "')"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE FitoXEPI" & vbCrLf & _
                      "  WHERE Fito_Id  =" & _CodigoFito & vbCrLf & _
                      "    AND EPI_Id   =" & _CodigoEPI
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class