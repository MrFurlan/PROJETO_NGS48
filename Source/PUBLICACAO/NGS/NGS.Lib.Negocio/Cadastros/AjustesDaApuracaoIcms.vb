Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListAjustesDaApuracaoIcms
    Inherits List(Of AjustesDaApuracaoIcms)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select Codigo_Id, Descricao, isnull(DescricaoDetalhadaObservacoes,'') AS DescricaoDetalhadaObservacoes, isnull(DescricaoRAICMS_Id,0) as DescricaoRAICMS_Id" & vbCrLf & _
                                "from AjustesDaApuracaoIcms"

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "AjustesDaApuracaoIcms")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim aIcms As New AjustesDaApuracaoIcms
                aIcms.Codigo = row("Codigo_Id")
                aIcms.Descricao = row("Descricao")
                aIcms.DescricaoDetalhadaObservacoes = row("DescricaoDetalhadaObservacoes")
                aIcms.DescricaoRAICMS = row("DescricaoRAICMS_Id")
                Me.Add(aIcms)
            Next
        End If
    End Sub

    Public Sub New(ByVal UF As String)
        Dim sql As String = "Select Codigo_Id, Descricao, isnull(DescricaoDetalhadaObservacoes,'') AS DescricaoDetalhadaObservacoes, isnull(DescricaoRAICMS_Id,0) as DescricaoRAICMS_Id " & vbCrLf & _
                            "from AjustesDaApuracaoIcms " & vbCrLf & _
                            "where left(Codigo_Id,2) = '" & UF & "'"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "AjustesDaApuracaoIcms")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim aIcms As New AjustesDaApuracaoIcms
            aIcms.Codigo = row("Codigo_Id")
            aIcms.Descricao = row("Descricao")
            aIcms.DescricaoDetalhadaObservacoes = row("DescricaoDetalhadaObservacoes")
            aIcms.DescricaoRAICMS = row("DescricaoRAICMS_Id")
            Me.Add(aIcms)
        Next
    End Sub

End Class

<Serializable()>
Public Class AjustesDaApuracaoIcms
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Codigo_Id, Descricao, isnull(DescricaoDetalhadaObservacoes,'') AS DescricaoDetalhadaObservacoes, isnull(DescricaoRAICMS_Id,0) as DescricaoRAICMS_Id" & vbCrLf & _
                            "from AjustesDaApuracaoIcms" & vbCrLf & _
                            "where Codigo_Id = '" & pCodigo & "'"

        ds = Banco.ConsultaDataSet(sql, "AjustesDaApuracaoIcms")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Codigo_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
            _DescricaoDetalhadaObservacoes = ds.Tables(0).Rows(0)("DescricaoDetalhadaObservacoes")
            _DescricaoRAICMS = ds.Tables(0).Rows(0)("DescricaoRAICMS_Id")
        End If
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As String = ""
    Private _Descricao As String = ""
    Private _DescricaoDetalhadaObservacoes As String = ""
    Private _DescricaoRAICMS As Integer = 0
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

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
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

    Public Property DescricaoDetalhadaObservacoes() As String
        Get
            Return _DescricaoDetalhadaObservacoes
        End Get
        Set(ByVal value As String)
            _DescricaoDetalhadaObservacoes = value
        End Set
    End Property

    Public Property DescricaoRAICMS() As Integer
        Get
            Return _DescricaoRAICMS
        End Get
        Set(ByVal value As Integer)
            _DescricaoRAICMS = value
        End Set
    End Property
#End Region

#Region "Methods"

    Public Function CodigoUtilizado() As Boolean
        Dim Banco As New AcessaBanco

        Dim Sql As String = "select Codigo_Id" & vbCrLf & _
                            "from OperacaoXEstado" & vbCrLf & _
                            "where CodigoBeneficio = '" & Me.Codigo & "'"

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(Sql, "OperacaoXEstado")

        If ds.Tables(0).Rows.Count > 0 Then Return True

        Sql = "Select AjustesApuracaoIcms_Id" & vbCrLf & _
                "from ResumoItensRAICMS" & vbCrLf & _
                "where AjustesApuracaoIcms_Id = '" & Codigo & "'"
        ds = Banco.ConsultaDataSet(Sql, "ResumoItensRAICMS")

        If ds.Tables(0).Rows.Count > 0 Then Return True

        Return False

    End Function



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
                Sql = " INSERT INTO AjustesDaApuracaoIcms(Codigo_Id, Descricao) " & vbCrLf & _
                      " VALUES ('" & Me.Codigo & "','" & Me.Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE AjustesDaApuracaoIcms SET" & vbCrLf & _
                      "     Descricao ='" & _Descricao & "'" & vbCrLf & _
                      "  WHERE Codigo_Id = '" & _Codigo & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE AjustesDaApuracaoIcms" & vbCrLf & _
                      "  WHERE Codigo_Id = '" & _Codigo & "'"
                Sqls.Add(Sql)
        End Select
    End Sub

#End Region

End Class
