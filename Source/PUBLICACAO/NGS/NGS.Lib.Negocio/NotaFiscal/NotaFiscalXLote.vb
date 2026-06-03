Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListNotaFiscalXLote
    Inherits List(Of NotaFiscalXLote)

#Region "Fields"
    Private _itemNota As NotaFiscalXItem
#End Region


#Region "Property"
    Public ReadOnly Property ItemNota As NotaFiscalXItem
        Get
            Return _itemNota
        End Get
    End Property
#End Region


#Region "Construtor"
    Public Sub New(ByRef nfxi As NotaFiscalXItem)
        _itemNota = nfxi
    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        If _itemNota.IUD = "D" Then
            Dim sql As String
            sql = " Delete NotaFiscalXLote " & vbCrLf &
                  " Where Empresa_Id      ='" & _itemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                  "   and EndEmpresa_Id   = " & _itemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
                  "   and Cliente_Id      ='" & _itemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                  "   and EndCliente_Id   = " & _itemNota.NotaFiscal.EnderecoCliente & vbCrLf &
                  "   and EntradaSaida_Id ='" & _itemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  "   and Serie_Id        ='" & _itemNota.NotaFiscal.Serie & "'" & vbCrLf &
                  "   and Nota_Id         = " & _itemNota.NotaFiscal.Codigo & vbCrLf &
                  "   and Produto_Id      ='" & _itemNota.CodigoProduto & "'" & vbCrLf &
                  "   and CFOP_Id         = " & _itemNota.CFOP & vbCrLf &
                  "   and Sequencia_Id    = " & _itemNota.Sequencia & vbCrLf

            Sqls.Add(sql)
        End If

        For Each nfl As NotaFiscalXLote In _itemNota.Lotes
            If _itemNota.IUD <> "D" Then nfl.IUD = _itemNota.IUD

            If nfl.IUD <> "" Then
                nfl.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Sub CarregarLotes()
        Dim Sql As String
        Sql = "SELECT Lote_Id," & vbCrLf & _
              "       Quantidade," & vbCrLf & _
              "       Fabricado," & vbCrLf & _
              "       Validade," & vbCrLf & _
              "       isnull(QuantidadeDeConsumo,0) AS QuantidadeDeConsumo" & vbCrLf & _
              "	 FROM NotaFiscalXLote " & vbCrLf & _
              " WHERE Empresa_Id      ='" & _itemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
              "   AND EndEmpresa_Id   = " & _itemNota.NotaFiscal.EnderecoEmpresa & vbCrLf & _
              "   AND Cliente_Id      ='" & _itemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
              "   AND EndCliente_Id   = " & _itemNota.NotaFiscal.EnderecoCliente & vbCrLf & _
              "   AND EntradaSaida_Id ='" & IIf(_itemNota.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   AND Serie_Id        ='" & _itemNota.NotaFiscal.Serie & "'" & vbCrLf & _
              "   AND Nota_Id         = " & _itemNota.NotaFiscal.Codigo & vbCrLf & _
              "   AND Produto_id      ='" & _itemNota.CodigoProduto & "'" & vbCrLf & _
              "   AND CFOP_Id         = " & _itemNota.CFOP & vbCrLf & _
              "   AND Sequencia_Id    = " & _itemNota.Sequencia
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(Sql, "Notas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Nfl As New NotaFiscalXLote(_itemNota)

            Nfl.Lote = row("Lote_Id")
            Nfl.LoteOld = row("Lote_Id")
            Nfl.Quantidade = row("Quantidade")
            Nfl.Fabricado = CDate(row("Fabricado"))
            Nfl.Validade = CDate(row("Validade"))
            Nfl.ValidadeOld = CDate(row("Validade"))
            Nfl.QuantidadeDeConsumo = row("QuantidadeDeConsumo")

            Me.Add(Nfl)
        Next
    End Sub

#End Region

End Class


'**************************************************************************************************
'*********************************   Classe Base NotaFiscalXLote  *********************************
'**************************************************************************************************
<Serializable()> _
Public Class NotaFiscalXLote

#Region "Construtor"
    Public Sub New()

    End Sub
    Public Sub New(ByRef nfxi As NotaFiscalXItem)
        _ItemNota = nfxi
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _ItemNota As NotaFiscalXItem
    Private _Lote As String
    Private _Quantidade As Decimal
    Private _Fabricado As DateTime = Date.Today
    Private _Validade As DateTime = Date.Today
    Private _QuantidadeDeConsumo As Decimal
    Private _LoteOld As String
    Private _ValidadeOld As DateTime = Date.Today

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

    Public ReadOnly Property ItemNota As NotaFiscalXItem
        Get
            Return _ItemNota
        End Get
    End Property

    Public Property Lote() As String
        Get
            Return _Lote
        End Get
        Set(ByVal value As String)
            _Lote = value
        End Set
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property Fabricado() As DateTime
        Get
            Return _Fabricado
        End Get
        Set(ByVal value As DateTime)
            _Fabricado = value
        End Set
    End Property

    Public Property Validade() As DateTime
        Get
            Return _Validade
        End Get
        Set(ByVal value As DateTime)
            _Validade = value
        End Set
    End Property

    Public Property QuantidadeDeConsumo() As Decimal
        Get
            Return _QuantidadeDeConsumo
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeDeConsumo = value
        End Set
    End Property

    Public Property LoteOld As String
        Get
            Return _LoteOld
        End Get
        Set(value As String)
            _LoteOld = value
        End Set
    End Property

    Public Property ValidadeOld As Date
        Get
            Return _ValidadeOld
        End Get
        Set(value As Date)
            _ValidadeOld = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case IUD
            Case "I"
                sql = "INSERT INTO NotaFiscalXLote" & vbCrLf & _
                      "  (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id, Lote_Id, Quantidade, Fabricado, Validade, QuantidadeDeConsumo) " & vbCrLf & _
                      " VALUES ('" & _ItemNota.NotaFiscal.CodigoEmpresa & "'," & _ItemNota.NotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
                      "'" & _ItemNota.NotaFiscal.CodigoCliente & "', " & _ItemNota.NotaFiscal.EnderecoCliente & ", " & vbCrLf & _
                      "'" & _ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & _ItemNota.NotaFiscal.Serie & "', " & _ItemNota.NotaFiscal.Codigo & "," & vbCrLf & _
                       "'" & _ItemNota.CodigoProduto & "', " & _ItemNota.CFOP & "," & _ItemNota.Sequencia & ",'" & _Lote & "'," & Str(_Quantidade) & ",'" & _Fabricado.ToString("yyyy-MM-dd") & "','" & _Validade.ToString("yyyy-MM-dd") & "'," & Str(_QuantidadeDeConsumo) & ")" & vbCrLf
                Sqls.Add(sql)
            Case "U"
                sql = "Update NotaFiscalXLote set" & vbCrLf & _
                      "     Quantidade          = " & Str(_Quantidade) & vbCrLf & _
                      "    ,Fabricado           = '" & _Fabricado.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "    ,Validade            = '" & _Validade.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "    ,QuantidadeDeConsumo = " & Str(_QuantidadeDeConsumo) & vbCrLf & _
                      " Where Empresa_Id      ='" & _ItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id   = " & _ItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id      ='" & _ItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id   = " & _ItemNota.NotaFiscal.EnderecoCliente & vbCrLf & _
                      "   and EntradaSaida_Id ='" & _ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "   and Serie_Id        ='" & _ItemNota.NotaFiscal.Serie & "'" & vbCrLf & _
                      "   and Nota_Id         = " & _ItemNota.NotaFiscal.Codigo & vbCrLf & _
                      "   and Produto_Id      ='" & _ItemNota.CodigoProduto & "'" & vbCrLf & _
                      "   and CFOP_Id         = " & _ItemNota.CFOP & vbCrLf & _
                      "   and Sequencia_Id    = " & _ItemNota.Sequencia & vbCrLf & _
                      "   and Lote_Id         = '" & _Lote & "'" & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotaFiscalXLote " & vbCrLf & _
                      " Where Empresa_Id      ='" & _ItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id   = " & _ItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id      ='" & _ItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id   = " & _ItemNota.NotaFiscal.EnderecoCliente & vbCrLf & _
                      "   and EntradaSaida_Id ='" & _ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "   and Serie_Id        ='" & _ItemNota.NotaFiscal.Serie & "'" & vbCrLf & _
                      "   and Nota_Id         = " & _ItemNota.NotaFiscal.Codigo & vbCrLf & _
                      "   and Produto_Id      ='" & _ItemNota.CodigoProduto & "'" & vbCrLf & _
                      "   and Sequencia_Id    = " & _ItemNota.Sequencia & vbCrLf & _
                      "   and CFOP_Id         = " & _ItemNota.CFOP & vbCrLf & _
                      "   and Lote_Id         = '" & _Lote & "'" & vbCrLf
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class
