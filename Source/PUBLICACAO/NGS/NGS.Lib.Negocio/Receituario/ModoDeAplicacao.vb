Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListModoDeAplicacao
    Inherits List(Of ModoDeAplicacao)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoCulturaPragaFito As Integer, ByVal pformadeaplicacao As Integer, Optional ByVal pCodigoFito As Integer = 0, Optional ByVal PCodigoCultura As Integer = 0, Optional ByVal PCodigoPraga As Integer = 0)
        If pCodigoCulturaPragaFito = 0 And pformadeaplicacao = 0 Then Exit Sub
        If pCodigoFito = 0 Or PCodigoCultura = 0 Or PCodigoPraga = 0 Or pformadeaplicacao = 0 Then Exit Sub

        Dim sql As String
        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        sql = "SELECT MA.CulturaPragaFito_Id, MA.FormaDeAplicacao_Id, isnull(MA.Descricao,'') as Descricao" & vbCrLf & _
              "  FROM ModoDeAplicacao MA" & vbCrLf & _
              " Inner Join CulturaXPragaXFito CPF" & vbCrLf & _
              "    on MA.CulturaPragaFito_Id = CPF.CulturaPragaFito_Id" & vbCrLf

        If pCodigoCulturaPragaFito > 0 Then
            sql &= " Where MA.CulturaPragaFito_Id = " & pCodigoCulturaPragaFito & vbCrLf & _
                   "   And MA.FormaDeAplicacao_Id = " & pformadeaplicacao
        Else
            sql &= " Where CPF.Fito               = " & pCodigoFito & vbCrLf & _
                   "   and CPF.Cultura            = " & PCodigoCultura & vbCrLf & _
                   "   and CPF.Praga              = " & PCodigoPraga & vbCrLf & _
                   "   and MA.FormaDeAplicacao_Id = " & pformadeaplicacao
        End If

        ds = Banco.ConsultaDataSet(sql, "ModoDeAplicacao")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim MA As New ModoDeAplicacao
            MA.CodigoCulturaPragaFito = row("CulturaPragaFito_Id")
            MA.CodigoFormaDeAplicacao = row("FormaDeAplicacao_Id")
            MA.Descricao = row("Descricao")
            Me.Add(MA)
        Next

    End Sub

End Class

<Serializable()> _
Public Class ModoDeAplicacao

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoCulturaPragaFito As Integer, ByVal pformadeaplicacao As Integer, Optional ByVal pCodigoFito As Integer = 0, Optional ByVal PCodigoCultura As Integer = 0, Optional ByVal PCodigoPraga As Integer = 0)
        If pCodigoCulturaPragaFito = 0 And pformadeaplicacao = 0 Then Exit Sub
        If pCodigoFito = 0 Or PCodigoCultura = 0 Or PCodigoPraga = 0 Or pformadeaplicacao = 0 Then Exit Sub

        Dim sql As String
        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        sql = "SELECT MA.CulturaPragaFito_Id, MA.FormaDeAplicacao_Id, isnull(MA.Descricao,'') as Descricao" & vbCrLf & _
              "  FROM ModoDeAplicacao MA" & vbCrLf & _
              " Inner Join CulturaXPragaXFito CPF" & vbCrLf & _
              "    on MA.CulturaPragaFito_Id = CPF.CulturaPragaFito_Id" & vbCrLf

        If pCodigoCulturaPragaFito > 0 Then
            sql &= " Where MA.CulturaPragaFito_Id = " & pCodigoCulturaPragaFito & vbCrLf & _
                   "   And MA.FormaDeAplicacao_Id = " & pformadeaplicacao
        Else
            sql &= " Where CPF.Fito               = " & pCodigoFito & vbCrLf & _
                   "   and CPF.Cultura            = " & PCodigoCultura & vbCrLf & _
                   "   and CPF.Praga              = " & PCodigoPraga & vbCrLf & _
                   "   and MA.FormaDeAplicacao_Id = " & pformadeaplicacao
        End If

        ds = Banco.ConsultaDataSet(sql, "Dosagem")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim row As DataRow = ds.Tables(0).Rows(0)

        _CodigoCulturaPragaFito = row("CulturaPragaFito_Id")
        _CodigoFormaDeAplicacao = row("FormaDeAplicacao_Id")
        _Descricao = row("Descricao")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoCulturaPragaFito As Integer
    Private _FormaDeAplicacao As FormaDeAplicacao
    Private _CodigoFormaDeAplicacao As Integer
    Private _Descricao As String
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

    Public Property CodigoCulturaPragaFito() As Integer
        Get
            Return _CodigoCulturaPragaFito
        End Get
        Set(ByVal value As Integer)
            _CodigoCulturaPragaFito = value
        End Set
    End Property

    Public ReadOnly Property FormaDeAplicacao() As FormaDeAplicacao
        Get
            If _FormaDeAplicacao Is Nothing And _CodigoFormaDeAplicacao > 0 Then _FormaDeAplicacao = New FormaDeAplicacao(_CodigoFormaDeAplicacao)
            Return _FormaDeAplicacao
        End Get
    End Property

    Public Property CodigoFormaDeAplicacao() As Integer
        Get
            Return _CodigoFormaDeAplicacao
        End Get
        Set(ByVal value As Integer)
            _CodigoFormaDeAplicacao = value
            _FormaDeAplicacao = Nothing
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
                Sql = " INSERT INTO ModoDeAplicacao(CulturaPragaFito_Id,FormaDeAplicacao_Id, Descricao) " & vbCrLf & _
                      " VALUES (" & _CodigoCulturaPragaFito & "," & _CodigoFormaDeAplicacao & ",'" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ModoDeAplicacao SET" & vbCrLf & _
                      "    Descricao ='" & _Descricao & "'" & vbCrLf & _
                      "  WHERE CulturaPragaFito_Id = " & _CodigoCulturaPragaFito & vbCrLf & _
                      "    AND FormaDeAplicacao_Id = " & _CodigoFormaDeAplicacao
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ModoDeAplicacao" & vbCrLf & _
                      "  WHERE CulturaPragaFito_Id = " & _CodigoCulturaPragaFito & vbCrLf & _
                      "    AND FormaDeAplicacao_Id = " & _CodigoFormaDeAplicacao
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class