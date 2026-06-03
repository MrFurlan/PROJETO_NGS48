Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListProdutoXEmbalagem
    Inherits List(Of ProdutoXEmbalagem)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pProduto As Produto, Optional ByVal pEmbalagem As Integer = 0, Optional ByVal pTipoDeEmbalagem As String = "")
        Dim sql As String
        sql = "Select Produto_id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_id, PesoBruto, PesoLiquido, isnull(PesoVariavel,0) as PesoVariavel " & vbCrLf & _
              "  from ProdutoXEmbalagem" & vbCrLf & _
              " Where Produto_id ='" & pProduto.Codigo & "'"
        If pEmbalagem > 0 Then
            sql &= " And Embalagem_Id = " & pEmbalagem & vbCrLf
        End If

        If pTipoDeEmbalagem.Length > 0 Then
            sql &= " And TipoDeEmbalagem_Id = '" & pEmbalagem & "'" & vbCrLf
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "ProdutoXEmbalagem")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim PE As New ProdutoXEmbalagem(pProduto)
            PE.CodigoEmbalagem = row("Embalagem_Id")
            PE.CodigoTipoDeEmbalagem = row("TipoDeEmbalagem_Id")
            PE.Capacidade = row("CapacidadeEmbalagem_Id")
            PE.PesoBruto = row("PesoBruto")
            PE.PesoLiquido = row("PesoLiquido")
            PE.PesoVariavel = row("PesoVariavel")
            Me.Add(PE)
        Next
    End Sub

End Class

<Serializable()> _
Public Class ProdutoXEmbalagem

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pProduto As Produto)
        Me.Produto = pProduto
    End Sub

    Public Sub New(ByVal pProduto As String, ByVal pEmbalagem As Integer, ByVal pTipoDeEmbalagem As String, ByVal pCapacidadeEmbalagem As Decimal)
        Dim sql As String
        sql = "Select Produto_id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_id, PesoBruto, PesoLiquido, isnull(PesoVariavel,0) as PesoVariavel " & vbCrLf & _
              "  from ProdutoXEmbalagem" & vbCrLf & _
              " Where Produto_id             ='" & pProduto & "'" & vbCrLf & _
              "   And Embalagem_Id           = " & pEmbalagem & vbCrLf & _
              "   And TipoDeEmbalagem_Id     ='" & pTipoDeEmbalagem & "'" & vbCrLf & _
              "   And CapacidadeEmbalagem_id = " & Str(pCapacidadeEmbalagem)

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "ProdutoXEmbalagem")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow
            row = ds.Tables(0).Rows(0)

            Me.Produto = New Produto(row("Produto_Id"))
            Me.CodigoEmbalagem = row("Embalagem_Id")
            Me.CodigoTipoDeEmbalagem = row("TipoDeEmbalagem_Id")
            Me.Capacidade = row("CapacidadeEmbalagem_Id")
            Me.PesoBruto = row("PesoBruto")
            Me.PesoLiquido = row("PesoLiquido")
            Me.PesoVariavel = row("PesoVariavel")
        End If

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Produto As Produto
    Private _CodigoEmbalagem As Integer
    Private _Embalagem As Embalagem
    Private _CodigoTipoDeEmbalagem As String = ""
    Private _TipoDeEmbalagem As TipoDeEmbalagem
    Private _Capacidade As Double
    Private _PesoBruto As Double
    Private _PesoLiquido As Double
    Private _SaldoEstoque As Double
    Private _PesoVariavel As Boolean
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

    Public Property Produto() As Produto
        Get
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property CodigoEmbalagem() As Integer
        Get
            Return _CodigoEmbalagem
        End Get
        Set(ByVal value As Integer)
            _CodigoEmbalagem = value
        End Set
    End Property

    Public Property Embalagem() As Embalagem
        Get
            If _Embalagem Is Nothing And _CodigoEmbalagem > 0 Then _Embalagem = New Embalagem(_CodigoEmbalagem)
            Return _Embalagem
        End Get
        Set(ByVal value As Embalagem)
            _Embalagem = value
        End Set
    End Property

    Public ReadOnly Property DescricaoEmbalagem() As String
        Get
            Return Embalagem.Descricao
        End Get
    End Property

    Public ReadOnly Property EmbalagemIndea() As String
        Get
            Return Embalagem.EmbalagemIndea
        End Get
    End Property

    Public Property CodigoTipoDeEmbalagem() As String
        Get
            Return _CodigoTipoDeEmbalagem
        End Get
        Set(ByVal value As String)
            _CodigoTipoDeEmbalagem = value
        End Set
    End Property

    Public Property TipoDeEmbalagem() As TipoDeEmbalagem
        Get
            If _TipoDeEmbalagem Is Nothing And _CodigoTipoDeEmbalagem.Length > 0 Then _TipoDeEmbalagem = New TipoDeEmbalagem(_CodigoTipoDeEmbalagem)
            Return _TipoDeEmbalagem
        End Get
        Set(ByVal value As TipoDeEmbalagem)
            _TipoDeEmbalagem = value
        End Set
    End Property

    Public ReadOnly Property DescricaoTipoDeEmbalagem() As String
        Get
            Return TipoDeEmbalagem.Descricao
        End Get
    End Property

    Public Property Capacidade() As Double
        Get
            Return _Capacidade
        End Get
        Set(ByVal value As Double)
            _Capacidade = value
        End Set
    End Property

    Public Property SaldoEstoque() As Double
        Get
            Return _SaldoEstoque
        End Get
        Set(ByVal value As Double)
            _SaldoEstoque = value
        End Set
    End Property

    Public Property PesoBruto() As Double
        Get
            Return _PesoBruto
        End Get
        Set(ByVal value As Double)
            _PesoBruto = value
        End Set
    End Property

    Public Property PesoLiquido() As Double
        Get
            Return _PesoLiquido
        End Get
        Set(ByVal value As Double)
            _PesoLiquido = value
        End Set
    End Property

    Public Property PesoVariavel() As Boolean
        Get
            Return _PesoVariavel
        End Get
        Set(ByVal value As Boolean)
            _PesoVariavel = value
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
                Sql = " INSERT INTO ProdutoXEmbalagem(Produto_Id,Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id, PesoBruto, PesoLiquido, PesoVariavel) " & vbCrLf & _
                      " VALUES ('" & Produto.Codigo & "'," & Me.CodigoEmbalagem & ",'" & Me.CodigoTipoDeEmbalagem & "'," & Str(Me.Capacidade) & "," & Str(Me.PesoBruto) & "," & Str(Me.PesoLiquido) & "," & IIf(Me.PesoVariavel, "1", "0") & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " Update ProdutoXEmbalagem set" & vbCrLf & _
                      "   PesoBruto    = " & Str(Me.PesoBruto) & vbCrLf & _
                      "   PesoLiquido  = " & Str(Me.PesoLiquido) & vbCrLf & _
                      "   PesoVariavel = " & IIf(Me.PesoVariavel, "1", "0") & _
                      "  WHERE Produto_Id             ='" & Produto.Codigo & "'" & vbCrLf & _
                      "    AND Embalagem_Id           = " & Me.CodigoEmbalagem & vbCrLf & _
                      "    AND TipoDeEmbalagem_Id     ='" & Me.CodigoTipoDeEmbalagem & "'" & vbCrLf & _
                      "    AND CapacidadeEmbalagem_Id = " & Str(Me.Capacidade) & vbCrLf
            Case "D"
                Sql = " DELETE ProdutoXEmbalagem" & vbCrLf & _
                      "  WHERE Produto_Id             ='" & Produto.Codigo & "'" & vbCrLf & _
                      "    AND Embalagem_Id           = " & Me.CodigoEmbalagem & vbCrLf & _
                      "    AND TipoDeEmbalagem_Id     ='" & Me.CodigoTipoDeEmbalagem & "'" & vbCrLf & _
                      "    AND CapacidadeEmbalagem_Id = " & Str(Me.Capacidade) & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function EmbalagemXNotaFiscal() As Boolean
        Dim strSQL As String = "Select Top 1 Nota_Id " & vbCrLf & _
                               "  From NotasFiscaisXItens " & vbCrLf & _
                               " where Produto_id      ='" & Produto.Codigo & "'" & vbCrLf & _
                               "   and Embalagem       = " & Me.CodigoEmbalagem & vbCrLf & _
                               "   and TipoDeEmbalagem ='" & Me.CodigoTipoDeEmbalagem & "'"
        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet
        ds = objBanco.ConsultaDataSet(strSQL, "EmbalagemXNotaFiscal")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

End Class