Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMicroRegiao
    Inherits List(Of MicroRegiao)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pRegiao As Regioes)
        Dim sql As String = "Select MicroRegiao_Id, Regiao_Id, Descricao from MicroRegioes " & vbCrLf & _
                            " where Regiao_Id = " & pRegiao.Codigo

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "MicroRegioes")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Rg As New MicroRegiao(pRegiao)
            Rg.Codigo = row("MicroRegiao_Id")
            Rg.Descricao = row("Descricao")
            Me.Add(Rg)
        Next
    End Sub

End Class

<Serializable()> _
Public Class MicroRegiao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal Regiao As Regioes)
        Me.Regiao = Regiao
    End Sub

    Public Sub New(ByVal pRegiao As Integer, ByVal pMicroRegiao As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select MicroRegiao_Id, Regiao_Id, Descricao " & vbCrLf & _
                            "  from MicroRegioes " & vbCrLf & _
                            " where Regiao_Id = " & pRegiao & vbCrLf & _
                            "   and MicroRegiao_Id = " & pMicroRegiao

        ds = Banco.ConsultaDataSet(sql, "MicroRegioes")
        If ds.Tables(0).Rows.Count > 0 Then
            Me.Regiao = New Regioes(ds.Tables(0).Rows(0)("Regiao_Id"))
            Me.Codigo = ds.Tables(0).Rows(0)("MicroRegiao_Id")
            Me.Descricao = ds.Tables(0).Rows(0)("Descricao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Regiao As Regioes
    Private _Descricao As String = ""
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

    Public Property Regiao() As Regioes
        Get
            Return _Regiao
        End Get
        Set(ByVal value As Regioes)
            _Regiao = value
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
                Sql = " INSERT INTO MicroRegioes(MicroRegiao_Id, Regiao_Id, Descricao) " & vbCrLf & _
                      " VALUES ( (SELECT isnull(MAX(MicroRegiao_Id),0)+1 FROM MicroRegioes WHERE Regiao_ID= " & Me.Regiao.Codigo & ") , " & Me.Regiao.Codigo & ",'" & Me.Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE MicroRegioes      " & vbCrLf & _
                      "    SET Descricao      = '" & Me.Descricao & "'" & vbCrLf & _
                      "  WHERE MicroRegiao_Id =  " & Me.Codigo & vbCrLf & _
                        "  AND Regiao_Id      =  " & Me.Regiao.Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE MicroRegioes" & vbCrLf & _
                      "  WHERE MicroRegiao_Id = " & Me.Codigo & vbCrLf & _
                      "    AND Regiao_Id      = " & Me.Regiao.Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class