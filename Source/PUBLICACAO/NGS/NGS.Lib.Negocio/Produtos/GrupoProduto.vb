Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListGrupoProduto
    Inherits List(Of GrupoProduto)

    Public Erro As Exception

    Public Sub New(Optional ByVal CarregarDados As Boolean = False, Optional ByVal Agrupar As String = "", Optional ByVal ControlarLote As String = "", Optional ByVal ControlarEmbalagem As String = "")
        If CarregarDados Then
            Dim Banco As New AcessaBanco

            Dim strSQL As String
            strSQL = "SELECT DISTINCT GE.Grupo_Id, GE.Descricao, GE.Custo, isnull(Ge.AgruparFinanceiro,0) as AgruparFinanceiro, isnull(MapaDeEstoque, 0) as MapaDeEstoque" & _
                     "  FROM GruposDeEstoques GE " & _
                     " LEFT JOIN Produtos P " & _
                     "    ON GE.Grupo_Id = P.Grupo " & _
                     " WHERE LEN(GE.Grupo_Id) >= 5 "

            If Agrupar.Length > 0 Then
                strSQL &= " And P.Agrupar = '" & Agrupar & "'"
            End If

            If ControlarLote.Length > 0 Then
                strSQL &= " And P.ControlarLote = '" & ControlarLote & "'"
            End If

            If ControlarEmbalagem.Length > 0 Then
                If ControlarLote.Length > 0 Then
                    strSQL &= " Or P.ControlarEmbalagem = '" & ControlarEmbalagem & "'"
                Else
                    strSQL &= " And P.ControlarEmbalagem = '" & ControlarEmbalagem & "'"
                End If
            End If

            strSQL &= " ORDER BY GE.Descricao"

            Dim dsGrupos As DataSet = Banco.ConsultaDataSet(strSQL, "Grupos")

            For Each drGrupo As DataRow In dsGrupos.Tables(0).Rows
                Dim objGrupo As New GrupoProduto()
                objGrupo.Codigo = drGrupo("Grupo_Id")
                objGrupo.Descricao = drGrupo("Descricao")
                If Not IsDBNull(drGrupo("Custo")) Then
                    objGrupo.Custo = drGrupo("Custo")
                End If
                objGrupo.AgrupaFinanceiro = drGrupo("AgruparFinanceiro")
                objGrupo.RelatorioEstoque = drGrupo("MapaDeEstoque")
                Me.Add(objGrupo)
            Next
        End If
    End Sub

    Public Function SelecionarComProdutos() As Boolean
        Dim objBanco As New AcessaBanco()
        Try
            Dim strSQL As String
            strSQL = "SELECT DISTINCT GE.Grupo_Id, GE.Descricao " & _
                     "  FROM GruposDeEstoques GE " & _
                     " INNER JOIN Produtos P " & _
                     "    ON GE.Grupo_Id = P.Grupo " & _
                     " WHERE LEN(GE.Grupo_Id) >= 5 " & _
                     " ORDER BY GE.Descricao"

            Dim dsGrupos As DataSet = objBanco.ConsultaDataSet(strSQL, "GruposDeEstoques")

            For Each drGrupo As DataRow In dsGrupos.Tables(0).Rows
                Dim objGrupo As New GrupoProduto()

                objGrupo.Codigo = drGrupo("Grupo_Id").ToString()
                objGrupo.Descricao = drGrupo("Descricao").ToString()

                Me.Add(objGrupo)
            Next

            Return True
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

End Class

<Serializable()> _
Public Class GrupoProduto
    Public Erro As Exception

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Grupo_Id, Descricao, Custo, isnull(AgruparFinanceiro,0) as AgruparFinanceiro, MapaDeEstoque  " & _
                                   "  FROM GruposDeEstoques " & _
                                   "  WHERE Grupo_Id = '" & pCodigo & "'"

            Dim dsGrupos As DataSet = objBanco.ConsultaDataSet(strSQL, "GruposDeEstoques")

            If dsGrupos.Tables(0).Rows.Count > 0 Then
                _Codigo = dsGrupos.Tables(0).Rows(0)("Grupo_Id").ToString()
                _Descricao = dsGrupos.Tables(0).Rows(0)("Descricao").ToString()
                _Custo = dsGrupos.Tables(0).Rows(0)("Custo")
                _AgrupaFinanceiro = dsGrupos.Tables(0).Rows(0)("AgruparFinanceiro")
                _RelatorioEstoque = dsGrupos.Tables(0).Rows(0)("MapaDeEstoque")
            End If


        Catch ex As Exception
            Me.Erro = ex
        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _Selecionado As Boolean = False
    Private _IUD As String
    Private _Codigo As String
    Private _Descricao As String
    Private _Custo As Boolean
    Private _AgrupaFinanceiro As Boolean
    Private _RelatorioEstoque As Boolean
    Private _Produtos As ListProduto
#End Region

#Region "Property"
    Public Property Selecionado() As Boolean
        Get
            Return _Selecionado
        End Get
        Set(ByVal value As Boolean)
            _Selecionado = value
        End Set
    End Property

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

    Public Property Custo() As Boolean
        Get
            Return _Custo
        End Get
        Set(ByVal value As Boolean)
            _Custo = value
        End Set
    End Property

    Public Property AgrupaFinanceiro() As Boolean
        Get
            Return _AgrupaFinanceiro
        End Get
        Set(ByVal value As Boolean)
            _AgrupaFinanceiro = value
        End Set
    End Property

    Public Property RelatorioEstoque() As Boolean
        Get
            Return _RelatorioEstoque
        End Get
        Set(ByVal value As Boolean)
            _RelatorioEstoque = value
        End Set
    End Property

    Public Property Produtos() As ListProduto
        Get
            If _Produtos Is Nothing And _Codigo.Length > 0 Then _Produtos = New ListProduto(_Codigo, "", "", "", "", "", True)
            Return _Produtos
        End Get
        Set(ByVal value As ListProduto)
            _Produtos = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        salvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub salvarSql(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        Select Case IUD
            Case "I"
                strSql = "INSERT Into GruposDeEstoques(Grupo_id, Descricao, Custo, AgruparFinanceiro, MapaDeEstoque) " & vbCrLf & _
                  " Values('" & Me.Codigo & "','" & UCase(Me.Descricao) & "','" & Me.Custo & "','" & Me.AgrupaFinanceiro & "','" & Me.RelatorioEstoque & "')"
                sqls.Add(strSql)
            Case "U"
                strSql = "UPDATE GruposDeEstoques Set Descricao = '" & UCase(Me.Descricao) & "'," & vbCrLf & _
                                                         "Custo = '" & Me.Custo & "'," & vbCrLf & _
                                             "AgruparFinanceiro = '" & Me.AgrupaFinanceiro & "'," & vbCrLf & _
                                                 "MapaDeEstoque = '" & Me.RelatorioEstoque & "'" & vbCrLf & _
                         " WHERE Grupo_Id = '" & Me.Codigo & "'"
                sqls.Add(strSql)
            Case "D"
                strSql = "DELETE FROM GruposDeEstoques" & vbCrLf & _
                      " WHERE Grupo_Id = '" & Me.Codigo & "'"
                sqls.Add(strSql)
        End Select
    End Sub
#End Region
End Class