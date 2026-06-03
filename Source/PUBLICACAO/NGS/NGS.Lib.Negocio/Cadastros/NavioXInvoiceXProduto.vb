Public Class ListNavioXInvoiceXProduto
    Inherits List(Of NavioXInvoiceXProduto)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objNavioXInvoice As NavioXInvoice)

        Me.NavioXInvoice = objNavioXInvoice

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Codigo_Id, Produto_Id, Quantidade " & vbCrLf & _
                     "From NavioXInvoiceXProduto " & vbCrLf

            If Me.NavioXInvoice.Codigo > 0 Then strSQL &= "  and Codigo_Id = " & Me.NavioXInvoice.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "NavioXInvoiceXProduto")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPRD As New NavioXInvoiceXProduto(Me.NavioXInvoice)

                oPRD.CodigoProduto = row("Produto_Id")
                oPRD.Quantidade = row("Quantidade")

                Me.Add(oPRD)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As NavioXInvoiceXProduto In Me
            If item.IUD = "I" OrElse item.IUD = "D" Then item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _NavioXInvoice As NavioXInvoice
#End Region

#Region "Property"
    Public Property NavioXInvoice() As NavioXInvoice
        Get
            Return _NavioXInvoice
        End Get
        Set(ByVal value As NavioXInvoice)
            _NavioXInvoice = value
        End Set
    End Property

#End Region

End Class

<Serializable()> _
Public Class NavioXInvoiceXProduto
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _NavioXInvoice As NavioXInvoice

    Private _CodigoProduto As String = ""
    Private _DescricaoProduto As String = ""
    Private _Produto As Produto
    Private _Quantidade As Decimal

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

    Public Property NavioXInvoice() As NavioXInvoice
        Get
            Return _NavioXInvoice
        End Get
        Set(ByVal value As NavioXInvoice)
            _NavioXInvoice = value
        End Set
    End Property

    Public Property CodigoProduto As String
        Get
            Return _CodigoProduto
        End Get
        Set(value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property DescricaoProduto As String
        Get
            Return _DescricaoProduto
        End Get
        Set(value As String)
            _DescricaoProduto = value
        End Set
    End Property

    Public Property Produto As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property Quantidade As Decimal
        Get
            Return _Quantidade
        End Get
        Set(value As Decimal)
            _Quantidade = value
        End Set
    End Property

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal NavioXInvoice As NavioXInvoice)
        Me.NavioXInvoice = NavioXInvoice

        Dim Banco As New AcessaBanco
        Dim sql As String = " SELECT Produto_id, Quantidade" & vbCrLf &
                            " FROM NavioXInvoiceXProduto " & vbCrLf &
                            " WHERE Codigo_Id = " & Me.NavioXInvoice.Codigo

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NavioXProduto")
        If ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim Row As DataRow = ds.Tables(0).Rows(0)

        _CodigoProduto = Row("Produto_id")
        _Quantidade = Row("Quantidade")
        _DescricaoProduto = _CodigoProduto & "-" & Me.Produto.Nome & "(" & Me.Produto.RegistroMinisterioAgricultura & ")"

    End Sub
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSql As String = ""

        Select Case Me.IUD
            Case "I"
                strSql = "INSERT INTO NavioXInvoiceXProduto(Codigo_Id, Produto_Id, Quantidade)" & vbCrLf & _
                         "Values(@cod,'" & Me.CodigoProduto & "'," & Me.Quantidade & ");"

            Case "U"
                strSql = "Update NavioXInvoiceXProduto " & vbCrLf &
                         " Set Quantidade   = " & Me.Quantidade.ToString().Replace(",", ".") & vbCrLf &
                         " Where Codigo_Id  = " & NavioXInvoice.Codigo & vbCrLf &
                         "   and Produto_Id = " & Me.CodigoProduto

        End Select

        Sqls.Add(strSql)
    End Sub
#End Region

End Class
