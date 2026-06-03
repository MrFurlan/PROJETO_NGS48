Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListRegioes
    Inherits List(Of Regioes)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select Regiao_Id, Descricao from Regioes "

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Regioes")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Rg As New Regioes
                Rg.Codigo = row("Regiao_Id")
                Rg.Descricao = row("Descricao")
                Me.Add(Rg)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class Regioes
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Regiao_Id, Descricao from Regioes where Regiao_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "Regioes")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Regiao_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
    Private _MicroRegioes As ListMicroRegiao
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property



    Public Property MicroRegioes() As ListMicroRegiao
        Get
            If _MicroRegioes Is Nothing And Me.Codigo > 0 Then _MicroRegioes = New ListMicroRegiao(Me)
            Return _MicroRegioes
        End Get
        Set(ByVal value As ListMicroRegiao)
            _MicroRegioes = value
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
                Sql = " INSERT INTO Regioes(Regiao_Id, Descricao) " & vbCrLf & _
                      " VALUES ((SELECT ISNULL(MAX(Regiao_ID),0)+1 FROM Regioes)" & ",'" & Me.Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Regioes SET" & vbCrLf & _
                      "    Descricao     = '" & Me.Descricao & "'" & vbCrLf & _
                      "  WHERE Regiao_Id = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Regioes" & vbCrLf & _
                      "  WHERE Regiao_Id = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class