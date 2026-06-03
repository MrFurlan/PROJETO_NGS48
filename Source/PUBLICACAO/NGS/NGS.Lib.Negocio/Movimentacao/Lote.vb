Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListLotes
    Inherits List(Of Lote)

End Class

<Serializable()> _
Public Class Lote

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pLote As String, ByVal pProduto As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim Sql As String

        Sql = "Select Lote_id, Produto_Id, DataLancamento, DataValidade, Fornecedor, EndFornecedor, Renasem, TermoConformidade, Boletim, Germinacao, isnull(Safra,'NENHUMA') as Safra, " & _
              " isnull(Pureza,0) as Pureza, Tipo " & vbCrLf & _
              "  From Lote" & vbCrLf & _
              " Where Lote_Id    ='" & pLote & "'" & vbCrLf & _
              "   and Produto_Id = " & pProduto

        ds = Banco.ConsultaDataSet(Sql, "Lote")

        If ds.Tables(0).Rows.Count > 0 Then
            _Lote = ds.Tables(0).Rows(0)("Lote_id")
            _CodigoProduto = ds.Tables(0).Rows(0)("Produto_Id")
            _DataLancamento = ds.Tables(0).Rows(0)("DataLancamento")
            _DataValidade = ds.Tables(0).Rows(0)("DataValidade")
            _CodigoFornecedor = ds.Tables(0).Rows(0)("Fornecedor")
            _CodigoFornecedorEndereco = ds.Tables(0).Rows(0)("EndFornecedor")
            _Renasem = ds.Tables(0).Rows(0)("Renasem")
            _TermoConformidade = ds.Tables(0).Rows(0)("TermoConformidade")
            _Boletim = ds.Tables(0).Rows(0)("Boletim")
            _Germinacao = ds.Tables(0).Rows(0)("Germinacao")
            _Pureza = ds.Tables(0).Rows(0)("Pureza")
            _CodigoSafra = ds.Tables(0).Rows(0)("Safra")
            _Tipo = ds.Tables(0).Rows(0)("Tipo")
            _Classificacoes = New ListLoteXClassificacao(Me)
        End If

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Lote As String = ""
    Private _Produto As Produto
    Private _DataLancamento As DateTime = Date.Now
    Private _DataValidade As DateTime = Date.Now
    Private _CodigoFornecedor As String = ""
    Private _CodigoFornecedorEndereco As Integer = 0
    Private _Fornecedor As Cliente
    Private _CodigoProduto As String
    Private _Renasem As String
    Private _TermoConformidade As String
    Private _Boletim As String
    Private _Germinacao As Decimal
    Private _Pureza As Decimal

    Private _Classificacoes As ListLoteXClassificacao
    Private _CodigoSafra As String
    Private _Safra As Safra

    Private _Banco As New AcessaBanco
    Private _msg As String
    Private _Tipo As Integer

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

    Public Property Lote() As String
        Get
            Return _Lote
        End Get
        Set(ByVal value As String)
            _Lote = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property DataLancamento() As DateTime
        Get
            Return _DataLancamento
        End Get
        Set(ByVal value As DateTime)
            _DataLancamento = value
        End Set
    End Property

    Public Property DataValidade() As DateTime
        Get
            Return _DataValidade
        End Get
        Set(ByVal value As DateTime)
            _DataValidade = value
        End Set
    End Property

    Public Property CodigoFornecedor() As String
        Get
            Return _CodigoFornecedor
        End Get
        Set(ByVal value As String)
            _CodigoFornecedor = value
        End Set
    End Property

    Public Property CodigoFornecedorEndereco() As Integer
        Get
            Return _CodigoFornecedorEndereco
        End Get
        Set(ByVal value As Integer)
            _CodigoFornecedorEndereco = value
        End Set
    End Property

    Public Property Fornecedor() As Cliente
        Get
            If _Fornecedor Is Nothing Then _Fornecedor = New Cliente(_CodigoFornecedor, CodigoFornecedorEndereco)
            Return _Fornecedor
        End Get
        Set(ByVal value As Cliente)
            _Fornecedor = value
        End Set
    End Property

    Public Property Renasem() As String
        Get
            Return _Renasem
        End Get
        Set(ByVal value As String)
            _Renasem = value
        End Set
    End Property

    Public Property TermoConformidade() As String
        Get
            Return _TermoConformidade
        End Get
        Set(ByVal value As String)
            _TermoConformidade = value
        End Set
    End Property

    Public Property Boletim() As String
        Get
            Return _Boletim
        End Get
        Set(ByVal value As String)
            _Boletim = value
        End Set
    End Property

    Public Property Germinacao() As Decimal
        Get
            Return _Germinacao
        End Get
        Set(ByVal value As Decimal)
            _Germinacao = value
        End Set
    End Property

    Public Property Pureza() As Decimal
        Get
            Return _Pureza
        End Get
        Set(ByVal value As Decimal)
            _Pureza = value
        End Set
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
            _Safra = Nothing
        End Set
    End Property

    Public Property Safra() As Safra
        Get
            If _Safra Is Nothing And _CodigoSafra.Length > 0 Then _Safra = New Safra(_CodigoSafra)
            Return _Safra
        End Get
        Set(ByVal value As Safra)
            _Safra = value
        End Set
    End Property

    Public Property Classificacoes() As ListLoteXClassificacao
        Get
            If _Classificacoes Is Nothing Then _Classificacoes = New ListLoteXClassificacao(Me)
            Return _Classificacoes
        End Get
        Set(ByVal value As ListLoteXClassificacao)
            _Classificacoes = value
        End Set
    End Property

    Public ReadOnly Property MensagemDeControle() As String
        Get
            Return _msg
        End Get
    End Property

    Public Property Tipo As Integer
        Get
            Return _Tipo
        End Get
        Set(ByVal value As Integer)
            _Tipo = value
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
                Sql = "Insert Into Lote(Lote_id, Produto_Id, DataLancamento, DataValidade, Fornecedor, EndFornecedor, Renasem, TermoConformidade, Boletim, Germinacao, Pureza, Safra, Tipo)" & vbCrLf & _
                      " values('" & Me.Lote & "','" & Me.CodigoProduto & "','" & Me.DataLancamento.ToSqlDate & "','" & Me.DataValidade.ToSqlDate & "','" & _CodigoFornecedor & "'," & _
                      Me.CodigoFornecedorEndereco & ",'" & Me.Renasem & "','" & Me.TermoConformidade & "','" & Me.Boletim & "'," & Str(Me.Germinacao) & "," & Str(Me.Pureza) & ",'" & Me.CodigoSafra & "'," & Me.Tipo & ")"
                Sqls.Add(Sql)
                SalvarTabelasRelacionadas(Sqls)
            Case "U"
                Sql = "Update Lote Set  " & _
                      "   DataLancamento    ='" & Me.DataLancamento.ToSqlDate & "'" & _
                      "  ,DataValidade      ='" & Me.DataValidade.ToSqlDate & "'" & _
                      "  ,Fornecedor        ='" & Me.CodigoFornecedor & "'" & _
                      "  ,EndFornecedor     = " & Me.CodigoFornecedorEndereco & _
                      "  ,Renasem           ='" & Me.Renasem & "'" & _
                      "  ,TermoConformidade ='" & Me.TermoConformidade & "'" & _
                      "  ,Boletim           ='" & Me.Boletim & "'" & _
                      "  ,Germinacao        = " & Str(Me.Germinacao) & _
                      "  ,Pureza            = " & Str(Me.Pureza) & _
                      "  ,Safra             ='" & Me.CodigoSafra & "'" & _
                      " WHERE Lote_id = '" & Me.Lote & "'" & _
                      "   AND Produto_Id = '" & Me.CodigoProduto & "'"

                Sqls.Add(Sql)
                SalvarTabelasRelacionadas(Sqls)
            Case "D"
                SalvarTabelasRelacionadas(Sqls)
                Sql = "Delete Lote " & vbCrLf & _
                      " WHERE Lote_id = '" & Me.Lote & "' " & _
                      "   AND Produto_Id = '" & Me.CodigoProduto & "'"

                Sqls.Add(Sql)
        End Select
    End Sub

    Public Sub SalvarTabelasRelacionadas(ByRef Sqls As ArrayList)
        Me.Classificacoes.SalvarSql(Sqls)
    End Sub

    Public Function JaExisteMovimentacao() As Boolean
        Dim sql As String

        '************************************************
        '************   NOTAS FISCAIS   *****************
        '************************************************

        sql = "select Case" & _
              "         When exists(select * from notasfiscaisxitens where produto_id = '" & _CodigoProduto & "' and lote = '" & _Lote & "')" & _
              "           Then 'TRUE'" & _
              "           Else 'FALSE'" & _
              "       End Existe"

        Dim ds As DataSet = _Banco.ConsultaDataSet(sql, "Existe")
        If ds.Tables(0).Rows(0)("Existe") = "TRUE" Then
            _msg = "Existe notas lançadas com este cadastro"
            Return True
        Else
            _msg = "Não existe notas lançadas com este cadastro"
        End If

        '************************************************
        '************      Produção     *****************
        '************************************************

        sql = "select Case" & _
              "         When exists(select * from Producao where produto_id = '" & _CodigoProduto & "' and lote_Id = '" & _Lote & "')" & _
              "           Then 'TRUE'" & _
              "           Else 'FALSE'" & _
              "       End Existe"

        ds = _Banco.ConsultaDataSet(sql, "Existe")
        If ds.Tables(0).Rows(0)("Existe") = "TRUE" Then
            _msg = "Existe Produção lançadas com este Lote"
            Return True
        Else
            _msg &= " - Não existe Produção lançadas com este Lote"
            Return False
        End If
    End Function
#End Region

End Class