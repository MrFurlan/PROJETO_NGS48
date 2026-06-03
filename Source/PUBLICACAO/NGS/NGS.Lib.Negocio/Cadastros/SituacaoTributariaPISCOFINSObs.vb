Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListSituacaoTributariaPISCOFINSObs
    Inherits List(Of SituacaoTributariaPISCOFINSObs)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pSitTribPISCOFINS As SituacaoTributariaPISCOFINS)
        Me.SituacaoTributaria = pSitTribPISCOFINS
        Dim sql As String
        sql = "Select SituacaoTributariaPISCOFINS_Id, Observacao_Id, Descricao " & vbCrLf & _
              "  from SituacaoTributariaPISCOFINSObs" & vbCrLf & _
              " where SituacaoTributariaPISCOFINS_Id= " & SituacaoTributaria.Codigo & vbCrLf & _
              " Order By Observacao_Id " & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "SituacaoTributariaPISCOFINS")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ST As New SituacaoTributariaPISCOFINSObs(Me.SituacaoTributaria)
            ST.CodigoObservacao = row("Observacao_Id")
            ST.Descricao = row("Descricao")
            Me.Add(ST)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _SituacaoTributaria As SituacaoTributariaPISCOFINS
#End Region

#Region "Property"
    Public Property SituacaoTributaria() As SituacaoTributariaPISCOFINS
        Get
            Return _SituacaoTributaria
        End Get
        Set(ByVal value As SituacaoTributariaPISCOFINS)
            _SituacaoTributaria = value
        End Set
    End Property
#End Region

End Class


<Serializable()> _
Public Class SituacaoTributariaPISCOFINSObs
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pSitTribPISCOFINS As SituacaoTributariaPISCOFINS)
        _SituacaoTributaria = pSitTribPISCOFINS
    End Sub

    Public Sub New(ByVal pCodigo As String, ByVal pCodigoST As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "select * from dbo.SituacaoTributariaPISCOFINSObs where SituacaoTributariaPISCOFINS_Id = " & pCodigoST & " and Observacao_Id = " & pCodigo
        ds = Banco.ConsultaDataSet(sql, "SituacaoTributariaPISCOFINS")
        If ds.Tables(0).Rows.Count > 0 Then
            _SituacaoTributaria = New SituacaoTributariaPISCOFINS(ds.Tables(0).Rows(0)("SituacaoTributariaPISCOFINS_Id"))
            _CodigoObservacao = ds.Tables(0).Rows(0)("Observacao_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _SituacaoTributaria As SituacaoTributariaPISCOFINS
    Private _CodigoObservacao As Integer = 0
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

    Public Property SituacaoTributaria() As SituacaoTributariaPISCOFINS
        Get
            Return _SituacaoTributaria
        End Get
        Set(ByVal value As SituacaoTributariaPISCOFINS)
            _SituacaoTributaria = value
        End Set
    End Property

    Public Property CodigoObservacao() As Integer
        Get
            Return _CodigoObservacao
        End Get
        Set(ByVal value As Integer)
            _CodigoObservacao = value
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

    Public ReadOnly Property DescricaoObj() As String
        Get
            Return Format(_CodigoObservacao, "000") & " - " & _Descricao
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim retorno As Boolean = False

        Sqls.Clear()

        SalvarSql(Sqls)

        If IUD = "D" Then
            If Not JaUsada() Then
                retorno = Banco.GravaBanco(Sqls)
            Else
                retorno = False
            End If
        Else
            retorno = Banco.GravaBanco(Sqls)
        End If


        Return retorno
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO SituacaoTributariaPISCOFINSObs(SituacaoTributariaPISCOFINS_Id,Observacao_Id, Descricao) " & vbCrLf & _
                      " VALUES (" & SituacaoTributaria.Codigo & "," & _CodigoObservacao & ",'" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE SituacaoTributariaPISCOFINSObs SET" & vbCrLf & _
                      "    Descricao     = '" & _Descricao & "'" & vbCrLf & _
                      "  WHERE SituacaoTributariaPISCOFINS_Id = " & SituacaoTributaria.Codigo & vbCrLf & _
                      "    AND Observacao_id                  = " & Me.CodigoObservacao & vbCrLf
                Sqls.Add(Sql)
            Case "D"

                Sql = " DELETE SituacaoTributariaPISCOFINSObs" & vbCrLf & _
                      "  WHERE SituacaoTributariaPISCOFINS_Id = " & SituacaoTributaria.Codigo & vbCrLf & _
                      "    AND Observacao_id                  = " & Me.CodigoObservacao & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
    Public Function JaUsada() As Boolean
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = ""

        sql = "SELECT top 1 ObsPISCOFINS" & vbCrLf & _
              "  from OperacaoxEstado " & vbCrLf & _
              " WHERE STPISCOFINS  = " & SituacaoTributaria.Codigo & vbCrLf & _
              "   AND ObsPISCOFINS = " & _CodigoObservacao

        ds = Banco.ConsultaDataSet(sql, "STPC")

        Return ds.Tables(0).Rows.Count > 0
    End Function

#End Region

End Class