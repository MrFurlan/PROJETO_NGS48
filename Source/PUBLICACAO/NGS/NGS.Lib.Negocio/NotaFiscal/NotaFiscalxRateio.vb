Imports NGS.Lib.Uteis


'***************************************************************************************************************************
'**********************************  Lista da Classe Nota Fiscal x Rateio   ************************************************
'***************************************************************************************************************************
<Serializable()> _
Public Class ListNotaFiscalxRateio
    Inherits List(Of NotaFiscalxRateio)

#Region "Fields"
    Private _itemNota As Negocio.NotaFiscalXItem
#End Region

#Region "Property"
    Public ReadOnly Property ItemNota As Negocio.NotaFiscalXItem
        Get
            Return _itemNota
        End Get
    End Property
#End Region

#Region "Construtor"
    Public Sub New(pItemNota As Negocio.NotaFiscalXItem)
        Dim Banco As New AcessaBanco
        Dim Ds As DataSet
        Dim Sql As String = ""

        Sql &= "SELECT UnidadeDeNegocioRateio_Id, EmpresaRateio_Id, EndEmpresaRateio_Id, CentroDeCusto_Id, Valor" & vbCrLf & _
               "  FROM NotasFiscaisXItensXRateio" & vbCrLf & _
               " where Empresa_Id      ='" & pItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
               "   and EndEmpresa_Id   = " & pItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf & _
               "   and Cliente_Id      ='" & pItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
               "   and EndCliente_Id   = " & pItemNota.NotaFiscal.EnderecoCliente & vbCrLf & _
               "   and EntradaSaida_Id ='" & pItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
               "   and Serie_Id        ='" & pItemNota.NotaFiscal.Serie & "'" & vbCrLf & _
               "   and Nota_Id         = " & pItemNota.NotaFiscal.Codigo & vbCrLf & _
               "   and Produto_Id      ='" & pItemNota.CodigoProduto & "'" & vbCrLf & _
               "   and CFOP_Id         = " & pItemNota.CFOP & vbCrLf & _
               "   and Sequencia_Id    = " & pItemNota.Sequencia & _
               "   and Nota_id > 0"

        Ds = Banco.ConsultaDataSet(Sql, "NotasFiscaisXItensXRateio")

        For Each row In Ds.Tables(0).Rows
            Dim rat As New Negocio.NotaFiscalxRateio(pItemNota)
            rat.CodigoUnidadeDeNegocio = row("UnidadeDeNegocioRateio_Id")
            rat.EndUnidadeDeNegocio = 0
            rat.CodigoEmpresaRateio = row("EmpresaRateio_Id")
            rat.EndEmpresaRateio = row("EndEmpresaRateio_Id")
            rat.CodigoCentroDeCusto = row("CentroDeCusto_Id")
            rat.Valor = row("Valor")
            Me.Add(rat)
        Next

    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each row As NotaFiscalxRateio In Me
            If row.ItemNota.IUD = "D" OrElse row.ItemNota.IUD = "I" Then row.IUD = row.ItemNota.IUD
            If row.IUD <> "" Then
                row.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Function LancamentoRateado() As Boolean
        'VR = soma dos Valores dos Lancamentos Rateados
        If Me.Count = 0 Then Return True

        Dim VR As Decimal = Me.Where(Function(s) s.IUD <> "D").Sum(Function(x) x.Valor)

        If VR <> ItemNota.ValorTotal Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region

End Class


'***************************************************************************************************************************
'**************************************  Classe Base Nota Fiscal x Rateio   ************************************************
'***************************************************************************************************************************
<Serializable()> _
Public Class NotaFiscalxRateio

#Region "Contrutor"
    Public Sub New(pItemNota As Negocio.NotaFiscalXItem)
        _ItemNota = pItemNota
    End Sub
#End Region

#Region "fields"
    Private _IUD As String = ""
    Private _ItemNota As Negocio.NotaFiscalXItem
    Private _CodigoUnidadeDeNegocio As String
    Private _EndUnidadeDeNegocio As Integer
    Private _UnidadeDeNegocio As Cliente
    Private _NomeUnidadeDeNegocio As String
    Private _NomeEmpresa As String
    Private _CodigoEmpresaRateio As String
    Private _EndEmpresaRateio As Integer
    Private _EmpresaRateio As Cliente
    Private _CodigoCentroDeCusto As String
    Private _CentroDeCusto As CentroDeCusto
    Private _DescCentroDeCusto As String
    Private _Valor As Decimal
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

    Public ReadOnly Property ItemNota As Negocio.NotaFiscalXItem
        Get
            Return _ItemNota
        End Get
    End Property

    Public Property CodigoUnidadeDeNegocio As String
        Get
            Return _CodigoUnidadeDeNegocio
        End Get
        Set(value As String)
            _CodigoUnidadeDeNegocio = value
            _UnidadeDeNegocio = Nothing
        End Set
    End Property

    Public Property EndUnidadeDeNegocio As Integer
        Get
            Return _EndUnidadeDeNegocio
        End Get
        Set(value As Integer)
            _EndUnidadeDeNegocio = value
            _UnidadeDeNegocio = Nothing
        End Set
    End Property

    Public ReadOnly Property UnidadeDeNegocio As Cliente
        Get
            If _UnidadeDeNegocio Is Nothing And Me.CodigoUnidadeDeNegocio.Length > 0 Then _UnidadeDeNegocio = New Cliente(Me.CodigoUnidadeDeNegocio, Me.EndUnidadeDeNegocio)
            Return _UnidadeDeNegocio
        End Get
    End Property

    Public Property NomeUnidadeDeNegocio() As String
        Get
            Return UnidadeDeNegocio.Nome
        End Get
        Set(ByVal value As String)
            _NomeUnidadeDeNegocio = value
        End Set
    End Property


    Public Property CodigoEmpresaRateio As String
        Get
            Return _CodigoEmpresaRateio
        End Get
        Set(value As String)
            _CodigoEmpresaRateio = value
            _EmpresaRateio = Nothing
        End Set
    End Property

    Public Property EndEmpresaRateio As Integer
        Get
            Return _EndEmpresaRateio
        End Get
        Set(value As Integer)
            _EndEmpresaRateio = value
            _EmpresaRateio = Nothing
        End Set
    End Property

    Public Property NomeEmpresa() As String
        Get
            Return Me.EmpresaRateio.Nome
        End Get
        Set(ByVal value As String)
            _NomeEmpresa = value
        End Set
    End Property


    Public ReadOnly Property EmpresaRateio As Cliente
        Get
            If _EmpresaRateio Is Nothing And Me.CodigoEmpresaRateio.Length > 0 Then _EmpresaRateio = New Cliente(Me.CodigoEmpresaRateio, Me.EndEmpresaRateio)
            Return _EmpresaRateio
        End Get
    End Property

    Public Property CodigoCentroDeCusto As String
        Get
            Return _CodigoCentroDeCusto
        End Get
        Set(value As String)
            _CodigoCentroDeCusto = value
            _CentroDeCusto = Nothing
            _DescCentroDeCusto = ""
        End Set
    End Property

    Public ReadOnly Property CentroDeCusto As CentroDeCusto
        Get
            If _CentroDeCusto Is Nothing And Me.CodigoCentroDeCusto.Length > 0 Then _CentroDeCusto = New CentroDeCusto(Me.CodigoCentroDeCusto)
            Return _CentroDeCusto
        End Get
    End Property

    Public Property DescCentroDeCusto As String
        Get
            Return ""
        End Get
        Set(value As String)

        End Set
    End Property

    Public Property Valor As Decimal
        Get
            Return _Valor
        End Get
        Set(value As Decimal)
            _Valor = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case IUD
            Case "I"
                sql = "INSERT INTO NotasFiscaisXItensXRateio (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Produto_Id, CFOP_Id, Sequencia_Id," & vbCrLf &
                      "                                       UnidadeDeNegocioRateio_Id, EmpresaRateio_Id, EndEmpresaRateio_Id, CentroDeCusto_Id, Valor)" & vbCrLf &
                      "VALUES ('" & Me.ItemNota.NotaFiscal.CodigoEmpresa & "'," & Me.ItemNota.NotaFiscal.EnderecoEmpresa & ",'" & Me.ItemNota.NotaFiscal.CodigoCliente & "'," & Me.ItemNota.NotaFiscal.EnderecoCliente & ",'" & Me.ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "','" & Me.ItemNota.NotaFiscal.Serie & "'," & Me.ItemNota.NotaFiscal.Codigo & ",'" & Me.ItemNota.CodigoProduto & "'," & Me.ItemNota.CFOP & "," & Me.ItemNota.Sequencia & "," & vbCrLf &
                      "        '" & Me.CodigoUnidadeDeNegocio & "','" & Me.CodigoEmpresaRateio & "'," & Me.EndEmpresaRateio & ",'" & Me.CodigoCentroDeCusto & "'," & Str(Me.Valor) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = " UPDATE NotasFiscaisXItensXRateio SET Valor = " & Str(Me.Valor) & vbCrLf &
                     "  Where Empresa_Id                ='" & Me.ItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                     "    and EndEmpresa_Id             = " & Me.ItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
                     "    and Cliente_Id                ='" & Me.ItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                     "    and EndCliente_Id             = " & Me.ItemNota.NotaFiscal.EnderecoCliente & vbCrLf &
                     "    and EntradaSaida_Id           ='" & Me.ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                     "    and Serie_Id                  ='" & Me.ItemNota.NotaFiscal.Serie & "'" & vbCrLf &
                     "    and Nota_Id                   = " & Me.ItemNota.NotaFiscal.Codigo & vbCrLf &
                     "    and Produto_Id                ='" & Me.ItemNota.CodigoProduto & "'" & vbCrLf &
                     "    and CFOP_Id                   = " & Me.ItemNota.CFOP & vbCrLf &
                     "    and Sequencia_Id              = " & Me.ItemNota.Sequencia & vbCrLf &
                     "    and UnidadeDeNegocioRateio_Id ='" & Me.CodigoUnidadeDeNegocio & "'" & vbCrLf &
                     "    and EmpresaRateio_Id          ='" & Me.CodigoEmpresaRateio & "'" & vbCrLf &
                     "    and EndEmpresaRateio_Id       = " & Me.EndEmpresaRateio & vbCrLf &
                     "    and CentroDeCusto_Id          ='" & Me.CodigoCentroDeCusto & "'"
                Sqls.Add(sql)
            Case "D"

                If Me.ItemNota.CodigoProdutoOld Is Nothing OrElse Me.ItemNota.CodigoProdutoOld.Length = 0 Then
                    sql = " Delete NotasFiscaisXItensXRateio " & vbCrLf &
                      "  Where Empresa_Id                ='" & Me.ItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "    and EndEmpresa_Id             = " & Me.ItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
                      "    and Cliente_Id                ='" & Me.ItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "    and EndCliente_Id             = " & Me.ItemNota.NotaFiscal.EnderecoCliente & vbCrLf &
                      "    and EntradaSaida_Id           ='" & Me.ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "    and Serie_Id                  ='" & Me.ItemNota.NotaFiscal.Serie & "'" & vbCrLf &
                      "    and Nota_Id                   = " & Me.ItemNota.NotaFiscal.Codigo & vbCrLf &
                      "    and Produto_Id                ='" & Me.ItemNota.CodigoProduto & "'" & vbCrLf &
                      "    and CFOP_Id                   = " & Me.ItemNota.CFOP & vbCrLf &
                      "    and Sequencia_Id              = " & Me.ItemNota.Sequencia & vbCrLf &
                      "    and UnidadeDeNegocioRateio_Id ='" & Me.CodigoUnidadeDeNegocio & "'" & vbCrLf &
                      "    and EmpresaRateio_Id          ='" & Me.CodigoEmpresaRateio & "'" & vbCrLf &
                      "    and EndEmpresaRateio_Id       = " & Me.EndEmpresaRateio & vbCrLf &
                      "    and CentroDeCusto_Id          ='" & Me.CodigoCentroDeCusto & "'"
                Else
                    sql = " Delete NotasFiscaisXItensXRateio " & vbCrLf &
                      "  Where Empresa_Id                ='" & Me.ItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "    and EndEmpresa_Id             = " & Me.ItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
                      "    and Cliente_Id                ='" & Me.ItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "    and EndCliente_Id             = " & Me.ItemNota.NotaFiscal.EnderecoCliente & vbCrLf &
                      "    and EntradaSaida_Id           ='" & Me.ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "    and Serie_Id                  ='" & Me.ItemNota.NotaFiscal.Serie & "'" & vbCrLf &
                      "    and Nota_Id                   = " & Me.ItemNota.NotaFiscal.Codigo & vbCrLf &
                      "    and Produto_Id                ='" & Me.ItemNota.CodigoProdutoOld & "'" & vbCrLf &
                      "    and CFOP_Id                   = " & Me.ItemNota.CFOPOld & vbCrLf &
                      "    and Sequencia_Id              = " & Me.ItemNota.Sequencia & vbCrLf &
                      "    and UnidadeDeNegocioRateio_Id ='" & Me.CodigoUnidadeDeNegocio & "'" & vbCrLf &
                      "    and EmpresaRateio_Id          ='" & Me.CodigoEmpresaRateio & "'" & vbCrLf &
                      "    and EndEmpresaRateio_Id       = " & Me.EndEmpresaRateio & vbCrLf &
                      "    and CentroDeCusto_Id          ='" & Me.CodigoCentroDeCusto & "'"
                End If

                Sqls.Add(sql)

        End Select
    End Sub
#End Region

End Class
