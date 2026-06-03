Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'***********************************************************************************************
'******************************  LISTA CONDICAO DE PAGAMENTO  **********************************
'***********************************************************************************************
<Serializable()> _
Public Class ListCondicaoPagamento
    Inherits List(Of CondicaoPagamento)

    Public Sub New(Optional Carregar As Boolean = False)
        If Not Carregar Then Exit Sub

        Dim objBanco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT Pagamento_Id, Descricao, Parcelas, isnull(Avista,0) as AVista, ISNULL(Antecipado,0) AS Antecipado, ISNULL(VencimentoPedido,0) AS VencimentoPedido " &
                  "  FROM Pagamentos " &
                  " ORDER BY Pagamento_Id"

            Dim dsCondicoes As DataSet = objBanco.ConsultaDataSet(sql, "Pagamentos")

            For Each row As DataRow In dsCondicoes.Tables(0).Rows
                Dim objCondicao As New CondicaoPagamento()

                objCondicao.Codigo = row("Pagamento_Id")
                objCondicao.Descricao = row("Descricao")
                objCondicao.Parcelas = row("Parcelas")
                objCondicao.AVista = row("AVista")
                objCondicao.Antecipado = row("Antecipado")
                objCondicao.VencimentoPedido = row("VencimentoPedido")

                Me.Add(objCondicao)
            Next

        Catch ex As Exception

        Finally
            objBanco = Nothing
        End Try
    End Sub
End Class


'***********************************************************************************************
'**************************  CLASSE BASE CONDICAO DE PAGAMENTO  ********************************
'***********************************************************************************************
<Serializable()> _
Public Class CondicaoPagamento
    Implements IBaseEntity

#Region "Fields"
    Private _Codigo As Integer
    Private _Descricao As String
    Private _Parcelas As Integer
    Private _AVista As Boolean
    Private _Periodo As ArrayList
    Private _Antecipado As Boolean = False
    Private _VencimentoPedido As Boolean = False

#End Region

#Region "Property"
    Public Property Codigo As Integer
        Get
            Return _Codigo
        End Get
        Set(value As Integer)
            _Codigo = value
        End Set
    End Property
    Public Property Descricao As String
        Get
            Return _Descricao
        End Get
        Set(value As String)
            _Descricao = value
        End Set
    End Property
    Public Property Parcelas As Integer
        Get
            Return _Parcelas
        End Get
        Set(value As Integer)
            _Parcelas = value
        End Set
    End Property
    Public Property Antecipado As Boolean
        Get
            Return _Antecipado
        End Get
        Set(value As Boolean)
            _Antecipado = value
        End Set
    End Property
    Public Property AVista As Boolean
        Get
            Return _AVista
        End Get
        Set(value As Boolean)
            _AVista = value
        End Set
    End Property
    Public Property VencimentoPedido As Boolean
        Get
            Return _VencimentoPedido
        End Get
        Set(value As Boolean)
            _VencimentoPedido = value
        End Set
    End Property
    Public Property Periodo As ArrayList
        Get
            If _Periodo Is Nothing Then CarregarPeriodo()
            Return _Periodo
        End Get
        Set(value As ArrayList)
            _Periodo = value
        End Set
    End Property
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Me.Codigo = Codigo
        Dim Banco As New AcessaBanco
        Try
            Dim sql As String
            sql = "SELECT Pagamento_Id, Descricao, Parcelas, isnull(AVista,0) as Avista, ISNULL(Antecipado,0) AS Antecipado, ISNULL(VencimentoPedido,0) AS VencimentoPedido " &
                  "  FROM Pagamentos " &
                  " WHERE Pagamento_Id = " & Codigo.ToString() & " " &
                  " ORDER BY Pagamento_Id"

            Dim dsCondicoes As DataSet = Banco.ConsultaDataSet(sql, "Produtos")

            If dsCondicoes.Tables(0).Rows.Count > 0 Then
                Dim row As DataRow = dsCondicoes.Tables(0).Rows(0)
                Me.Codigo = row("Pagamento_Id")
                Me.Descricao = row("Descricao")
                Me.Parcelas = row("Parcelas")
                Me.AVista = row("AVista")
                Me.Antecipado = row("Antecipado")
                Me.VencimentoPedido = row("VencimentoPedido")

            End If
        Catch ex As Exception

        Finally
            Banco = Nothing
        End Try
    End Sub

    ''' <summary>
    ''' Construtor para busca de numero de parcelas
    ''' </summary>
    ''' <param name="Codigo">Codigo do Pagamento</param>
    ''' <param name="Parcela">Numero de Parcelas</param>
    Public Sub New(ByVal Codigo As Integer, ByVal Parcelas As Integer)
        Me.Codigo = Codigo
        Dim Banco As New AcessaBanco
        Try
            Dim sql As String
            sql = "SELECT Pagamento_Id, Descricao, Parcelas, isnull(AVista,0) as Avista, ISNULL(Antecipado,0) AS Antecipado, ISNULL(VencimentoPedido,0) AS VencimentoPedido " &
                  " FROM Pagamentos " &
                  " WHERE AVista = 0 "

            If Codigo <> 0 Then
                sql &= " AND Pagamento_Id = " & Codigo & ""
            End If

            If Parcelas <> 0 Then
                sql &= " AND Parcelas = " & Parcelas & ""
            End If

            sql &= " ORDER BY Pagamento_Id"


            Dim dsCondicoes As DataSet = Banco.ConsultaDataSet(sql, "Produtos")

            If dsCondicoes.Tables(0).Rows.Count > 0 Then
                Dim row As DataRow = dsCondicoes.Tables(0).Rows(0)
                Me.Codigo = row("Pagamento_Id")
                Me.Descricao = row("Descricao")
                Me.Parcelas = row("Parcelas")
                Me.AVista = row("AVista")
                Me.Antecipado = row("Antecipado")
                Me.VencimentoPedido = row("VencimentoPedido")

            End If
        Catch ex As Exception

        Finally
            Banco = Nothing
        End Try
    End Sub

#End Region

#Region "Methods"
    Private Sub CarregarPeriodo()
        If Me.Codigo <> 0 Then
            Dim objBanco As New AcessaBanco()

            Dim sql As String
            sql = "SELECT Dias " & _
                  "  FROM PagamentosXParcelas " & _
                  " WHERE Pagamento_Id = " & Me.Codigo & _
                  " ORDER BY Sequencia_Id"

            Dim dsPeriodos As DataSet = objBanco.ConsultaDataSet(sql, "PagamentosXParcelas")

            _Periodo = New ArrayList
            For Each row As DataRow In dsPeriodos.Tables(0).Rows
                _Periodo.Add(row("Dias"))
            Next

            objBanco = Nothing
        End If
    End Sub
#End Region


End Class