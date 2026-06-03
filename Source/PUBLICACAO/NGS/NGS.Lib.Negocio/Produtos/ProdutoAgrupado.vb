Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListProdutoAgrupado
    Inherits List(Of ProdutoAgrupado)

#Region "Construtor"
    Public Sub New(pObjProduto As Negocio.Produto, Tipo As Integer)
        '0 - Produtos que sao agrupados pelo produto passado no pobjProduto
        '1 - Produto que esta sendo agrupado o pobjProduto
        Produto = pObjProduto
        If Produto.IUD = "I" Then Exit Sub

        Dim sql As String = ""
        sql = "Select Produto_id, ProdutoAgrupado_id" & vbCrLf & _
              "  from ProdutosAgrupados" & vbCrLf

        If Tipo = 0 Then
            sql &= " Where Produto_Id = " & Produto.Codigo
        Else
            sql &= " Where ProdutoAgrupado_id = " & Produto.Codigo
        End If


        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "ProdutosAgrupados")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        For Each row As DataRow In ds.Tables(0).Rows
            If Tipo = 1 Then
                Produto = New Produto(row("Produto_id"))
            End If

            Dim pa As New ProdutoAgrupado(Produto)

            If Tipo = 0 Then
                pa.CodigoProdutoAgrupado = row("ProdutoAgrupado_Id")
            Else
                pa.CodigoProdutoAgrupado = row("Produto_id")
            End If

            Me.Add(pa)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Produto As Negocio.Produto
#End Region

#Region "Property"
    Public Property Produto As Negocio.Produto
        Get
            Return _Produto
        End Get
        Set(value As Negocio.Produto)
            _Produto = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList

        sqls.Clear()
        Me.SalvarSQL(sqls)

        If sqls.Count = 0 OrElse Banco.GravaBanco(sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSQL(ByRef Sqls As ArrayList, Optional ByVal UsaNumerador As Boolean = True)
        For Each row As ProdutoAgrupado In Me
            If Produto.IUD = "I" Or Produto.IUD = "D" Then row.IUD = Produto.IUD
            If Not Produto.IUD = "" Then
                row.salvarSql(Sqls)
            End If
        Next
    End Sub

    Public Sub AdicionarProdutosAgrupados(pSql As String)
        'Recebe o sql do user control ucSelecaoProduto
        Dim sql As String = ""
        sql = "Select Produto_Id" & vbCrLf & _
              "  from Produtos P" & vbCrLf & _
              " where not exists (select 1" & vbCrLf & _
              "                     from ProdutosAgrupados PA" & vbCrLf & _
              "                    where PA.ProdutoAgrupado_Id = P.Produto_Id)" & vbCrLf & _
              "   AND Produto_id = '" & pSql & "'"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "ProdutosAgrupados")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim pa As New ProdutoAgrupado(Produto)
            pa.IUD = "I"
            pa.CodigoProdutoAgrupado = row("Produto_Id")
            If Produto.IUD = "U" Then
                If pa.Salvar Then pa.IUD = ""
            End If
            Me.Add(pa)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ProdutoAgrupado

#Region "Construtor"
    Public Sub New(pObjProduto As Negocio.Produto)
        _Produto = pObjProduto
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Produto As Negocio.Produto
    Private _CodigoProdutoAgrupado As String = ""
    Private _ProdutoAgrupado As Negocio.Produto
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property
    Public Property Produto As Negocio.Produto
        Get
            Return _Produto
        End Get
        Set(value As Negocio.Produto)
            _Produto = value
        End Set
    End Property
    Public ReadOnly Property NomeProduto As String
        Get
            Return Produto.Nome
        End Get
    End Property

    Public Property CodigoProdutoAgrupado As String
        Get
            Return _CodigoProdutoAgrupado
        End Get
        Set(value As String)
            _CodigoProdutoAgrupado = value
        End Set
    End Property
    Public Property ProdutoAgrupado As Negocio.Produto
        Get
            If _ProdutoAgrupado Is Nothing And _CodigoProdutoAgrupado.Length > 0 Then _ProdutoAgrupado = New Negocio.Produto(_CodigoProdutoAgrupado)
            Return _ProdutoAgrupado
        End Get
        Set(value As Negocio.Produto)
            _ProdutoAgrupado = value
        End Set
    End Property
    Public ReadOnly Property NomeProdutoAgrupado As String
        Get
            Return ProdutoAgrupado.Nome
        End Get
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
                strSql = "INSERT INTO ProdutosAgrupados(Produto_Id, ProdutoAgrupado_Id) " & vbCrLf & _
                         "Values('" & Produto.Codigo & "','" & _CodigoProdutoAgrupado & "')"
            Case "D"
                strSql = "DELETE ProdutosAgrupados " & vbCrLf & _
                         " Where Produto_Id         ='" & Produto.Codigo & "'" & vbCrLf & _
                         "   and ProdutoAgrupado_Id ='" & _CodigoProdutoAgrupado & "'"
        End Select

        sqls.Add(strSql)
    End Sub
#End Region

End Class
