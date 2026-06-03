Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAdiantamentoBaixa
    Inherits List(Of AdiantamentoBaixa)

#Region "Construtor"
    Public Sub New()

    End Sub
    Public Sub New(ByVal NumeroDoTituloDaBaixa As Integer)
        Dim Sql As String
        Sql = "SELECT TituloAdiantamento, Sequencia_Id, ISNULL(RegistroPedido,0) RegistroPedido, Titulo, ValorOficial, ValorMoeda," & vbCrLf & _
              "       ISNULL(VariacaoOficial,0) VariacaoOficial, DataBaixa " & vbCrLf & _
              "  FROM vw_AdiantamentosXBaixas" & vbCrLf & _
              " Where Titulo = " & NumeroDoTituloDaBaixa

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "AdiantamentoBaixa")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim AB As New AdiantamentoBaixa(New Adiantamento(row("TituloAdiantamento")))
            AB.Sequencia = row("Sequencia_Id")
            AB.RegistroPedido = row("RegistroPedido")
            AB.CodigoTitulo = row("Titulo")
            AB.ValorOficial = row("ValorOficial")
            AB.ValorMoeda = row("ValorMoeda")
            AB.VariacaoOficial = row("VariacaoOficial")
            AB.DataBaixa = row("DataBaixa")
            Me.Add(AB)
        Next

    End Sub

    Public Sub New(ByVal pAdiantamento As Adiantamento)
        Dim Sql As String
        Sql = "SELECT Sequencia_Id, ISNULL(RegistroPedido,0) RegistroPedido, Titulo, ValorOficial, ValorMoeda," & vbCrLf & _
              "       ISNULL(VariacaoOficial,0) VariacaoOficial, DataBaixa " & vbCrLf & _
              "  FROM vw_AdiantamentosXBaixas" & vbCrLf & _
              " Where Empresa_Id      ='" & pAdiantamento.CodigoEmpresa & "' " & vbCrLf & _
              "   and EndEmpresa_Id   = " & pAdiantamento.EndEmpresa & vbCrLf & _
              "   and Cliente_Id      ='" & pAdiantamento.CodigoCliente & "'" & vbCrLf & _
              "   and EndCliente_Id   = " & pAdiantamento.EndCliente & vbCrLf & _
              "   and Adiantamento_Id = " & pAdiantamento.Codigo
        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "AdiantamentoBaixa")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim AB As New AdiantamentoBaixa(pAdiantamento)
            AB.Sequencia = row("Sequencia_Id")
            AB.RegistroPedido = row("RegistroPedido")
            AB.CodigoTitulo = row("Titulo")
            AB.ValorOficial = row("ValorOficial")
            AB.ValorMoeda = row("ValorMoeda")
            AB.VariacaoOficial = row("VariacaoOficial")
            AB.DataBaixa = row("DataBaixa")
            Me.Add(AB)
        Next

    End Sub
#End Region

#Region "Fields"
    Private _TotalOficialBaixado As Decimal
    Private _TotalMoedaBaixado As Decimal
#End Region

#Region "Property"
    Public ReadOnly Property TotalOficialBaixado()
        Get
            _TotalOficialBaixado = 0
            For Each row As AdiantamentoBaixa In Me
                _TotalOficialBaixado += row.ValorOficial
            Next
            Return _TotalOficialBaixado
        End Get
    End Property

    Public ReadOnly Property TotalMoedaBaixado()
        Get
            _TotalMoedaBaixado = 0
            For Each row As AdiantamentoBaixa In Me
                _TotalMoedaBaixado += row.ValorMoeda
            Next
            Return _TotalMoedaBaixado
        End Get
    End Property
#End Region

End Class

<Serializable()> _
Public Class AdiantamentoBaixa
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByRef pAdiantamento As Adiantamento)
        _Adiantamento = pAdiantamento
    End Sub
#End Region

#Region "fields"
    Private _IUD As String = ""
    Private _Adiantamento As Adiantamento
    Private _Sequencia As Integer
    Private _RegistroPedido As Integer
    Private _Pedido As Pedido
    Private _CodigoTitulo As Integer
    Private _Titulo As Titulo
    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
    Private _VariacaoOficial As Decimal
    Private _DataBaixa As DateTime
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

    Public Property Adiantamento() As Adiantamento
        Get
            Return _Adiantamento
        End Get
        Set(ByVal value As Adiantamento)
            _Adiantamento = value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property RegistroPedido() As Integer
        Get
            Return _RegistroPedido
        End Get
        Set(ByVal value As Integer)
            _RegistroPedido = value
            _Pedido = Nothing
        End Set
    End Property

    Public Property Pedido() As Pedido
        Get
            If _Pedido Is Nothing And _RegistroPedido > 0 Then _Pedido = New Pedido(_Adiantamento.CodigoEmpresa, _Adiantamento.EndEmpresa, _RegistroPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property CodigoTitulo() As Integer
        Get
            Return _CodigoTitulo
        End Get
        Set(ByVal value As Integer)
            _CodigoTitulo = value
            _Titulo = Nothing
        End Set
    End Property

    Public Property Titulo() As Titulo
        Get
            If _Titulo Is Nothing And _CodigoTitulo > 0 Then _Titulo = New Titulo(_CodigoTitulo)
            Return _Titulo
        End Get
        Set(ByVal value As Titulo)
            _Titulo = value
        End Set
    End Property

    Public Property ValorOficial() As Decimal
        Get
            Return _ValorOficial
        End Get
        Set(ByVal value As Decimal)
            _ValorOficial = value
        End Set
    End Property

    Public Property ValorMoeda() As Decimal
        Get
            Return _ValorMoeda
        End Get
        Set(ByVal value As Decimal)
            _ValorMoeda = value
        End Set
    End Property

    Public Property VariacaoOficial() As Decimal
        Get
            Return _VariacaoOficial
        End Get
        Set(ByVal value As Decimal)
            _VariacaoOficial = value
        End Set
    End Property

    Public Property DataBaixa() As DateTime
        Get
            Return _DataBaixa
        End Get
        Set(ByVal value As DateTime)
            _DataBaixa = value
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
                Sql = "INSERT INTO AdiantamentosXBaixas (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Adiantamento_Id, Sequencia_Id, RegistroPedido, Titulo, ValorOficial, ValorMoeda, VariacaoOficial, DataBaixa)" & vbCrLf & _
                      "values ('" & Me.Adiantamento.CodigoEmpresa & "'," & Me.Adiantamento.EndEmpresa & ",'" & Me.Adiantamento.CodigoCliente & "'," & Me.Adiantamento.EndCliente & "," & Me.Adiantamento.Codigo & "," & Me.Sequencia & "," & vbCrLf & _
                      "         " & Me.RegistroPedido & "," & Me.CodigoTitulo & "," & Str(Me.ValorOficial) & "," & Str(Me.ValorMoeda) & "," & Str(Me.VariacaoOficial) & ",'" & Me.DataBaixa.ToSqlDate & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = "UPDATE VW_AdiantamentosXBaixas SET" & vbCrLf & _
                      "   RegistroPedido  = " & Me.RegistroPedido & vbCrLf & _
                      "  ,ValorOficial    = " & Str(Me.ValorOficial) & vbCrLf & _
                      "  ,ValorMoeda      = " & Str(Me.ValorMoeda) & vbCrLf & _
                      "  ,VariacaoOficial = " & Str(Me.VariacaoOficial) & vbCrLf & _
                      "  ,DataBaixa       ='" & Me.DataBaixa.ToSqlDate & "'" & vbCrLf & _
                      " WHERE TituloAdiantamento = " & Me.Adiantamento.CodigoTitulo & vbCrLf & _
                      "   And Titulo             = " & Me.CodigoTitulo
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE VW_AdiantamentosXBaixas" & vbCrLf & _
                      " WHERE TituloAdiantamento = " & Me.Adiantamento.CodigoTitulo & vbCrLf & _
                      "   And Titulo             = " & Me.CodigoTitulo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class