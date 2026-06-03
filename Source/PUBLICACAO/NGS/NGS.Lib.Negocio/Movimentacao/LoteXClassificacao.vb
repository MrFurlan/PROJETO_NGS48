Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListLoteXClassificacao
    Inherits List(Of LoteXClassificacao)

    Private _RowId As Integer
    Private _Banco As New AcessaBanco
    Private _msg As String
    Private _Classificacao As LoteXClassificacao
    Private _Parent As Lote

    Public Sub New()

    End Sub

    Public Sub New(ByRef pLote As Lote)
        _Parent = pLote
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim Sql As String

        Sql = "Select LxC.Lote_Id, LxC.Produto_id, LxC.Classificacao_Id, isnull(LxC.PesoSaco,0) as PesoSaco " & vbCrLf & _
              "  From LoteXClassificacao LxC     " & vbCrLf & _
              " Where LxC.Lote_Id    ='" & pLote.Lote & "'" & vbCrLf & _
              "   and LxC.Produto_Id ='" & pLote.CodigoProduto & "'"
        ds = Banco.ConsultaDataSet(Sql, "Classificacao")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim LxP As New LoteXClassificacao(pLote)
            LxP.Classificacao = row("Classificacao_Id")
            LxP.PesoSaco = row("PesoSaco")
            Me.Add(LxP)
        Next
    End Sub

    Public Sub PraQServeIsso(ByRef pProduto As String, ByVal data As Date)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim Sql As String

        Sql = " SELECT LoteXClassificacao.Lote_Id, LoteXClassificacao.Classificacao_Id, isnull(LoteXClassificacao.PesoSaco,0) as PesoSaco " & vbCrLf & _
              "   FROM Lote " & vbCrLf & _
              "  INNER JOIN LoteXClassificacao " & vbCrLf & _
              "     ON Lote.Produto_id = LoteXClassificacao.Produto_id " & vbCrLf & _
              "    AND Lote.Lote_id    = LoteXClassificacao.Lote_Id " & vbCrLf & _
              "  Where LoteXClassificacao.Produto_id  = '" & pProduto & "'" & vbCrLf & _
              "    and Lote.DataLancamento           <= '" & data.ToSqlDate() & "'" & vbCrLf & _
              "    and Lote.DataValidade             >= '" & data.ToSqlDate() & "'"

        ds = Banco.ConsultaDataSet(Sql, "Lote")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim LxP As New LoteXClassificacao()
            LxP.Lote.Lote = row("Lote_Id")
            LxP.Classificacao = row("Classificacao_Id")
            LxP.PesoSaco = row("PesoSaco")
            Me.Add(LxP)
        Next
    End Sub

    Public Function JaExisteClassificacao(ByVal pClassificacao As String) As Boolean
        _RowId = 0

        For Each p As LoteXClassificacao In Me
            If p.Classificacao = pClassificacao Then
                _Classificacao = p
                _msg = "Classificação já Existe"
                Return True
            End If
            _RowId += 1
        Next
        _msg = "Classificação não Encontrada"
        _RowId = -1
        Return False
    End Function

    Public Property ClassificacaoExistente() As LoteXClassificacao
        Get
            Return Me(_RowId)
        End Get
        Set(ByVal value As LoteXClassificacao)
            Me(_RowId) = value
        End Set
    End Property

    Public Function JaExisteMovimentacao(ByVal pClassificacao As LoteXClassificacao) As Boolean
        Dim sql As String

        '************************************************
        '************   NOTAS FISCAIS   *****************
        '************************************************

        sql = "select Case" & _
              "         When exists(select * from notasfiscaisxitens where produto_id = '" & pClassificacao.Lote.CodigoProduto & "' and lote = '" & pClassificacao.Lote.Lote & "' and Classificacao = '" & pClassificacao.Classificacao & "')" & _
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
              "         When exists(select * from Producao where produto_id = '" & pClassificacao.Lote.CodigoProduto & "' and lote_Id = '" & pClassificacao.Lote.Lote & "' and Classificacao_Id = '" & pClassificacao.Classificacao & "')" & _
              "           Then 'TRUE'" & _
              "           Else 'FALSE'" & _
              "       End Existe"

        ds = _Banco.ConsultaDataSet(sql, "Existe")
        If ds.Tables(0).Rows(0)("Existe") = "TRUE" Then
            _msg = "Existe Produção lançadas com esta Classificaçao"
            Return True
        Else
            _msg &= " - Não existe Produção lançadas com esta Classificação"
            Return False
        End If
    End Function

    Public Function Remover(ByVal pClassificacao As String) As Boolean
        If JaExisteClassificacao(pClassificacao) Then
            If Me(_RowId).IUD = "" Then
                If JaExisteMovimentacao(_Classificacao) Then
                    Return False
                Else
                    Me(_RowId).IUD = "D"
                    If Me(_RowId).Salvar() Then
                        _msg = "Classificáção excluida com Sucesso"
                        Me.RemoveAt(_RowId)
                        Return True
                    Else
                        _msg = "Erro durante ao excluir a Classificação.. Tente Novamente."
                        Return False
                    End If
                End If
            Else
                _msg = "Classificação excluida com Sucesso"
                Me.RemoveAt(_RowId)
                Return True
            End If
        End If
        Return False
    End Function

    Public ReadOnly Property MensagemControle() As String
        Get
            Return _msg
        End Get
    End Property

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As LoteXClassificacao In Me
            If _Parent.IUD = "D" Or _Parent.IUD = "I" Then item.IUD = _Parent.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub

End Class

'*******************************************************************************************
'*************************   CLASSE BASE LOTEXCLASSIFICACAO    *****************************
'*******************************************************************************************
Public Class LoteXClassificacao

#Region "Contrutor"
    Public Sub New(ByRef pLote As Lote)
        _Lote = pLote
    End Sub

    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoProduto As String, ByVal pLote As String, ByVal pClassificacao As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim Sql As String

        Sql = "Select LxC.Lote_Id, LxC.Produto_id, LxC.Classificacao_Id, isnull(LxC.PesoSaco,0) as PesoSaco " & vbCrLf & _
              "  From LoteXClassificacao LxC     " & vbCrLf & _
              " Where LxC.Lote_Id          ='" & pLote & "'" & vbCrLf & _
              "   And LXC.Classificacao_Id ='" & pClassificacao & "'" & vbCrLf & _
              "   and LxC.Produto_Id       ='" & pCodigoProduto & "'"

        ds = Banco.ConsultaDataSet(Sql, "Classificacao")

        Dim row As DataRow = ds.Tables(0).Rows(0)
        Me.Classificacao = row("Classificacao_Id")
        Me.PesoSaco = row("PesoSaco")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Lote As Lote
    Private _Classificacao As String
    Private _PesoSaco As Decimal
    Private _Entrada As Double
    Private _Saida As Double
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

    Public Property Lote() As Lote
        Get
            If _Lote Is Nothing Then _Lote = New Lote()
            Return _Lote
        End Get
        Set(ByVal value As Lote)
            _Lote = value
        End Set
    End Property

    Public ReadOnly Property CodigoLote() As String
        Get
            Return _Lote.Lote
        End Get
    End Property

    Public Property Classificacao() As String
        Get
            Return _Classificacao
        End Get
        Set(ByVal value As String)
            _Classificacao = value
        End Set
    End Property

    Public Property PesoSaco() As Decimal
        Get
            Return _PesoSaco
        End Get
        Set(ByVal value As Decimal)
            _PesoSaco = value
        End Set
    End Property

    Public Property Entrada() As Double
        Get
            Return _Entrada
        End Get
        Set(ByVal value As Double)
            _Entrada = value
        End Set
    End Property

    Public Property Saida() As Double
        Get
            Return _Saida
        End Get
        Set(ByVal value As Double)
            _Saida = value
        End Set
    End Property

    Public ReadOnly Property Saldo() As Double
        Get
            Return _Entrada - _Saida
        End Get
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

        If Lote.IUD = "D" Then Me.IUD = Lote.IUD

        Select Case Me.IUD
            Case "I"
                Sql = "Insert Into LoteXClassificacao(Lote_Id, Produto_Id, Classificacao_Id, PesoSaco)" & vbCrLf & _
                      " values('" & Lote.Lote & "','" & Lote.CodigoProduto & "','" & _Classificacao & "'," & Str(Me.PesoSaco) & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update LoteXClassificacao set" & vbCrLf & _
                      "  PesoSaco = " & Str(Me.PesoSaco) & vbCrLf & _
                      " Where Lote_Id          ='" & Lote.Lote & "'" & vbCrLf & _
                      "   And Produto_id       ='" & Lote.CodigoProduto & "'" & vbCrLf & _
                      "   And Classificacao_Id ='" & _Classificacao & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = "Delete LoteXClassificacao" & vbCrLf & _
                      " Where Lote_Id          ='" & Lote.Lote & "'" & vbCrLf & _
                      "   And Produto_id       ='" & Lote.CodigoProduto & "'" & vbCrLf & _
                      "   And Classificacao_Id ='" & _Classificacao & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class