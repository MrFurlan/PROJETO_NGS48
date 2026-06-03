Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListDisponibilidadeProduto
    Inherits List(Of DisponibilidadeProduto)

#Region "Construtor"
    Public Sub New(ByVal pDisponibilidade As Disponibilidade, ByVal pwhere As String)
        _Disponibilidade = pDisponibilidade
        Dim sql As String = ""
        sql = "Select Produto_Id " & vbCrLf & _
              "  from Produtos" & vbCrLf & _
              " Where " & pwhere

        Dim banco As New AcessaBanco
        Dim ds As DataSet
        ds = banco.ConsultaDataSet(sql, "Produto")

        For Each row In ds.Tables(0).Rows
            Dim Prd As New DisponibilidadeProduto(pDisponibilidade)
            Prd.CodigoProduto = row("Produto_Id")
            Me.Add(Prd)
        Next
    End Sub

    Public Sub New(ByVal pDisponibilidade As Disponibilidade)
        _Disponibilidade = pDisponibilidade
        Dim sql As String = ""
        sql = "Select Disponibilidade_Id, Produto_Id " & vbCrLf & _
              "  from SaldoInicialDisponibilidadeProduto" & vbCrLf & _
              " Where Disponibilidade_id = " & pDisponibilidade.CodigoDisponibilidade

        Dim banco As New AcessaBanco
        Dim ds As DataSet
        ds = banco.ConsultaDataSet(sql, "Produto")

        For Each row In ds.Tables(0).Rows
            Dim Prd As New DisponibilidadeProduto(pDisponibilidade)
            Prd.CodigoProduto = row("Produto_Id")
            Me.Add(Prd)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Disponibilidade As Disponibilidade
#End Region

#Region "Property"
    Public ReadOnly Property Disponibilidade As Disponibilidade
        Get
            Return _Disponibilidade
        End Get
    End Property

    Public ReadOnly Property DescricaoProdutos As String
        Get
            If Me.Count = 0 Then Return ""
            Dim DescProduto As String = ""
            For Each row In Me
                DescProduto &= IIf(DescProduto.Length > 0, ", ", "") & row.CodigoProduto & "-" & row.NomeProduto
            Next
            Return DescProduto
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Item As DisponibilidadeProduto In Me
            If Disponibilidade.IUD = "D" Or Disponibilidade.IUD = "I" Then Item.IUD = Disponibilidade.IUD
            If Item.IUD <> "" Then
                Item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class DisponibilidadeProduto
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pDisponibilidade As Disponibilidade)
        _Disponibilidade = pDisponibilidade
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Disponibilidade As Disponibilidade
    Private _CodigoProduto As String
    Private _Produto As Produto
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property Disponibilidade As Disponibilidade
        Get
            Return _Disponibilidade
        End Get
    End Property

    Public Property CodigoProduto As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Produto = Nothing
        End Set
    End Property

    Public Property Produto As Produto
        Get
            If _Produto Is Nothing AndAlso Me.CodigoProduto.Length > 0 Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public ReadOnly Property NomeProduto As String
        Get
            If Me.Produto Is Nothing Then Return ""
            Return Me.Produto.Nome
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

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String = ""
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into SaldoInicialDisponibilidadeProduto(Disponibilidade_Id, Produto_Id) " & vbCrLf & _
                      " Values(" & Disponibilidade.CodigoDisponibilidade & ",'" & Me.CodigoProduto & "')"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete SaldoInicialDisponibilidadeProduto" & vbCrLf & _
                      "  Where Disponibilidade_Id = " & Disponibilidade.CodigoDisponibilidade & vbCrLf & _
                      "    And Produto_Id         = '" & Me.CodigoProduto & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
