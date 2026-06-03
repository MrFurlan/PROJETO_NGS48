Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

Public Class ListNotaFiscalReferencial
    Inherits List(Of NotaFiscalReferencial)
    Implements IBaseEntity

    Public Sub New()
    End Sub

    Public Sub New(ByVal pNotaFiscal As NotaFiscal, Optional ByVal Tipo As eTipoReferencial = eTipoReferencial.CTE)
        Dim sql As String
        sql = "SELECT *" & vbCrLf & _
              "  FROM NotaFiscalReferencial" & vbCrLf & _
              " WHERE EmpresaReferencial_Id      ='" & pNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
              "   AND EndEmpresaReferencial_Id   ='" & pNotaFiscal.EnderecoEmpresa & "'" & vbCrLf & _
              "   AND ClienteReferencial_Id      ='" & pNotaFiscal.CodigoCliente & "'" & vbCrLf & _
              "   AND EndClienteReferencial_Id   ='" & pNotaFiscal.EnderecoCliente & "'" & vbCrLf & _
              "   AND EntradaSaidaReferencial_Id ='" & IIf(pNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   AND SerieReferencial_Id        ='" & pNotaFiscal.Serie & "'" & vbCrLf & _
              "   AND NotaReferencial_Id         ='" & pNotaFiscal.Codigo & "'" & vbCrLf & _
              "   AND CFOPReferencial_Id         ='" & pNotaFiscal.Itens(0).CFOP & "'" & vbCrLf & _
              "   AND SequenciaReferencial_Id    ='" & pNotaFiscal.Itens(0).Sequencia & "'" & vbCrLf & _
              "   AND TipoReferencial_Id         ='" & Tipo.ToString & "'" & vbCrLf & _
              "   AND NotaReferencial_Id         > 0"

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "NotaFiscalReferencial")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New NotaFiscalReferencial()
                obj.Parent = pNotaFiscal.Itens(0)
                obj.EmpresaReferencial_Id = row("EmpresaReferencial_Id")
                obj.EndEmpresaReferencial_Id = row("EndEmpresaReferencial_Id")
                obj.ClienteReferencial_Id = row("ClienteReferencial_Id")
                obj.EndClienteReferencial_Id = row("EndClienteReferencial_Id")
                obj.EntradaSaidaReferencial_Id = IIf(row("EntradaSaidaReferencial_Id") = "S", 1, 0)
                obj.SerieReferencial_Id = row("SerieReferencial_Id")
                obj.NotaReferencial_Id = row("NotaReferencial_Id")
                obj.ProdutoReferencial_Id = row("ProdutoReferencial_Id")
                obj.CFOPReferencial_Id = row("CFOPReferencial_Id")
                obj.SequenciaReferencial_Id = row("SequenciaReferencial_Id")
                obj.TipoReferencial_Id = [Enum].Parse(GetType(eTipoReferencial), row("TipoReferencial_Id"))
                obj.Empresa_Id = row("Empresa_Id")
                obj.EndEmpresa_Id = row("EndEmpresa_Id")
                obj.Cliente_Id = row("Cliente_Id")
                obj.EndCliente_Id = row("EndCliente_Id")
                obj.EntradaSaida_Id = IIf(row("EntradaSaida_Id") = "S", 1, 0)
                obj.Serie_Id = row("Serie_Id")
                obj.Nota_Id = row("Nota_Id")
                obj.Produto_Id = row("Produto_Id")
                obj.CFOP_Id = row("CFOP_Id")
                obj.Sequencia_Id = row("Sequencia_Id")
                obj.Quantidade = row("Quantidade")
                obj.Valor = row("Valor")
                Me.Add(obj)
            Next
        End If
    End Sub

    Public Sub New(ByVal pNotaFiscalXItem As NotaFiscalXItem, Optional ByVal Tipo As eTipoReferencial = eTipoReferencial.CTE)
        Dim sql As String
        sql = "SELECT *" & vbCrLf & _
              "  FROM NotaFiscalReferencial" & vbCrLf & _
              " WHERE EmpresaReferencial_Id      ='" & pNotaFiscalXItem.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
              "   AND EndEmpresaReferencial_Id   ='" & pNotaFiscalXItem.NotaFiscal.EnderecoEmpresa & "'" & vbCrLf & _
              "   AND ClienteReferencial_Id      ='" & pNotaFiscalXItem.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
              "   AND EndClienteReferencial_Id   ='" & pNotaFiscalXItem.NotaFiscal.EnderecoCliente & "'" & vbCrLf & _
              "   AND EntradaSaidaReferencial_Id ='" & IIf(pNotaFiscalXItem.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   AND SerieReferencial_Id        ='" & pNotaFiscalXItem.NotaFiscal.Serie & "'" & vbCrLf & _
              "   AND NotaReferencial_Id         ='" & pNotaFiscalXItem.NotaFiscal.Codigo & "'" & vbCrLf & _
              "   AND CFOPReferencial_Id         ='" & pNotaFiscalXItem.CFOP & "'" & vbCrLf & _
              "   AND SequenciaReferencial_Id    ='" & pNotaFiscalXItem.Sequencia & "'" & vbCrLf & _
              "   AND TipoReferencial_Id         ='" & Tipo.ToString & "'" & vbCrLf & _
              "   AND NotaReferencial_Id         > 0"

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "NotaFiscalReferencial")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New NotaFiscalReferencial()
                obj.Parent = pNotaFiscalXItem
                obj.EmpresaReferencial_Id = row("EmpresaReferencial_Id")
                obj.EndEmpresaReferencial_Id = row("EndEmpresaReferencial_Id")
                obj.ClienteReferencial_Id = row("ClienteReferencial_Id")
                obj.EndClienteReferencial_Id = row("EndClienteReferencial_Id")
                obj.EntradaSaidaReferencial_Id = IIf(row("EntradaSaidaReferencial_Id") = "S", 1, 0)
                obj.SerieReferencial_Id = row("SerieReferencial_Id")
                obj.NotaReferencial_Id = row("NotaReferencial_Id")
                obj.ProdutoReferencial_Id = row("ProdutoReferencial_Id")
                obj.CFOPReferencial_Id = row("CFOPReferencial_Id")
                obj.SequenciaReferencial_Id = row("SequenciaReferencial_Id")
                obj.TipoReferencial_Id = row("TipoReferencial_Id")
                obj.Empresa_Id = row("Empresa_Id")
                obj.EndEmpresa_Id = row("EndEmpresa_Id")
                obj.Cliente_Id = row("Cliente_Id")
                obj.EndCliente_Id = row("EndCliente_Id")
                obj.EntradaSaida_Id = IIf(row("EntradaSaida_Id") = "S", 1, 0)
                obj.Serie_Id = row("Serie_Id")
                obj.Nota_Id = row("Nota_Id")
                obj.Produto_Id = row("Produto_Id")
                obj.CFOP_Id = row("CFOP_Id")
                obj.Sequencia_Id = row("Sequencia_Id")
                obj.Quantidade = row("Quantidade")
                obj.Valor = row("Valor")
                Me.Add(obj)
            Next
        End If
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As NotaFiscalReferencial In Me
            If String.IsNullOrWhiteSpace(item.IUD) Then
                item.IUD = item.Parent.IUD
            End If

            If item.IUD Is Nothing OrElse String.IsNullOrWhiteSpace(item.IUD) Then
                item.IUD = "I"
            End If

            If Not String.IsNullOrWhiteSpace(item.IUD) Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public ReadOnly Property QuantidadeTotal() As Decimal
        Get
            Return (From s In Me Select s.Quantidade).Sum
        End Get
    End Property

    Public ReadOnly Property ValorTotal() As Decimal
        Get
            Return (From s In Me Select s.Valor).Sum
        End Get
    End Property


End Class

<Serializable()> _
Public Class NotaFiscalReferencial
    Implements IBaseEntity

#Region "Atributos"

    Private _IUD As String
    Private _Parent As NotaFiscalXItem
    Private _EmpresaReferencial_Id As String
    Private _EndEmpresaReferencial_Id As String
    Private _ClienteReferencial_Id As String
    Private _EndClienteReferencial_Id As String
    Private _EntradaSaidaReferencial_Id As String
    Private _SerieReferencial_Id As String
    Private _NotaReferencial_Id As String
    Private _ProdutoReferencial_Id As String
    Private _CFOPReferencial_Id As String
    Private _SequenciaReferencial_Id As String
    Private _TipoReferencial_Id As eTipoReferencial
    Private _NotaFiscalXItem As NotaFiscalXItem
    Private _Empresa_Id As String
    Private _EndEmpresa_Id As String
    Private _Cliente_Id As String
    Private _EndCliente_Id As String
    Private _EntradaSaida_Id As String
    Private _Serie_Id As String
    Private _Nota_Id As String
    Private _Produto_Id As String
    Private _CFOP_Id As String
    Private _Sequencia_Id As String
    Private _Quantidade As String
    Private _Valor As String

#End Region

#Region "Propriedades"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Parent() As NotaFiscalXItem
        Get
            Return _Parent
        End Get
        Set(ByVal value As NotaFiscalXItem)
            _Parent = value
        End Set
    End Property

    Public Property EmpresaReferencial_Id() As String
        Get
            Return _EmpresaReferencial_Id
        End Get
        Set(ByVal value As String)
            _EmpresaReferencial_Id = value
        End Set
    End Property

    Public Property EndEmpresaReferencial_Id() As String
        Get
            Return _EndEmpresaReferencial_Id
        End Get
        Set(ByVal value As String)
            _EndEmpresaReferencial_Id = value
        End Set
    End Property

    Public Property ClienteReferencial_Id() As String
        Get
            Return _ClienteReferencial_Id
        End Get
        Set(ByVal value As String)
            _ClienteReferencial_Id = value
        End Set
    End Property

    Public Property EndClienteReferencial_Id() As String
        Get
            Return _EndClienteReferencial_Id
        End Get
        Set(ByVal value As String)
            _EndClienteReferencial_Id = value
        End Set
    End Property

    Public Property EntradaSaidaReferencial_Id() As Negocio.eEntradaSaida
        Get
            Return _EntradaSaidaReferencial_Id
        End Get
        Set(ByVal value As Negocio.eEntradaSaida)
            _EntradaSaidaReferencial_Id = value
        End Set
    End Property

    Public Property SerieReferencial_Id() As String
        Get
            Return _SerieReferencial_Id
        End Get
        Set(ByVal value As String)
            _SerieReferencial_Id = value
        End Set
    End Property

    Public Property NotaReferencial_Id() As String
        Get
            Return _NotaReferencial_Id
        End Get
        Set(ByVal value As String)
            _NotaReferencial_Id = value
        End Set
    End Property

    Public Property ProdutoReferencial_Id() As String
        Get
            Return _ProdutoReferencial_Id
        End Get
        Set(ByVal value As String)
            _ProdutoReferencial_Id = value
        End Set
    End Property

    Public Property CFOPReferencial_Id() As String
        Get
            Return _CFOPReferencial_Id
        End Get
        Set(ByVal value As String)
            _CFOPReferencial_Id = value
        End Set
    End Property

    Public Property SequenciaReferencial_Id() As String
        Get
            Return _SequenciaReferencial_Id
        End Get
        Set(ByVal value As String)
            _SequenciaReferencial_Id = value
        End Set
    End Property

    Public Property TipoReferencial_Id() As eTipoReferencial
        Get
            Return _TipoReferencial_Id
        End Get
        Set(ByVal value As eTipoReferencial)
            _TipoReferencial_Id = value
        End Set
    End Property

    Public Property Empresa_Id() As String
        Get
            Return _Empresa_Id
        End Get
        Set(ByVal value As String)
            _Empresa_Id = value
        End Set
    End Property

    Public Property EndEmpresa_Id() As String
        Get
            Return _EndEmpresa_Id
        End Get
        Set(ByVal value As String)
            _EndEmpresa_Id = value
        End Set
    End Property

    Public Property Cliente_Id() As String
        Get
            Return _Cliente_Id
        End Get
        Set(ByVal value As String)
            _Cliente_Id = value
        End Set
    End Property

    Public Property EndCliente_Id() As String
        Get
            Return _EndCliente_Id
        End Get
        Set(ByVal value As String)
            _EndCliente_Id = value
        End Set
    End Property

    Public Property EntradaSaida_Id() As Negocio.eEntradaSaida
        Get
            Return _EntradaSaida_Id
        End Get
        Set(ByVal value As Negocio.eEntradaSaida)
            _EntradaSaida_Id = value
        End Set
    End Property

    Public Property Serie_Id() As String
        Get
            Return _Serie_Id
        End Get
        Set(ByVal value As String)
            _Serie_Id = value
        End Set
    End Property

    Public Property Nota_Id() As String
        Get
            Return _Nota_Id
        End Get
        Set(ByVal value As String)
            _Nota_Id = value
        End Set
    End Property


    Public Property ParentOrigem() As NotaFiscalXItem
        Get

            Dim objnfTemp As New NotaFiscal()
            objnfTemp.CodigoEmpresa = Me.Empresa_Id
            objnfTemp.EnderecoEmpresa = Me.EndEmpresa_Id
            objnfTemp.CodigoCliente = Me.Cliente_Id
            objnfTemp.EnderecoCliente = Me.EndCliente_Id
            objnfTemp.EntradaSaida = Me.EntradaSaida_Id
            objnfTemp.Serie = Me.Serie_Id
            objnfTemp.Codigo = Me.Nota_Id

            Dim objnf As New NotaFiscal(objnfTemp)

            _NotaFiscalXItem = New NotaFiscalXItem(objnf)
            Return _NotaFiscalXItem
        End Get
        Set(ByVal value As NotaFiscalXItem)
            _NotaFiscalXItem = value
        End Set
    End Property


    Public Property Produto_Id() As String
        Get
            Return _Produto_Id
        End Get
        Set(ByVal value As String)
            _Produto_Id = value
        End Set
    End Property

    Public Property CFOP_Id() As String
        Get
            Return _CFOP_Id
        End Get
        Set(ByVal value As String)
            _CFOP_Id = value
        End Set
    End Property

    Public Property Sequencia_Id() As String
        Get
            Return _Sequencia_Id
        End Get
        Set(ByVal value As String)
            _Sequencia_Id = value
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

    Public Property Valor() As Decimal
        Get
            Return _Valor
        End Get
        Set(ByVal value As Decimal)
            _Valor = value
        End Set
    End Property

#End Region

#Region "Construtor"

    Public Sub New()
    End Sub

    Public Sub New(ByVal item As NotaFiscalXItem)
        Me.Parent = item
    End Sub

#End Region

#Region "Métodos"

    Public Function Salvar() As Boolean
        If String.IsNullOrWhiteSpace(IUD) Then
            Return True
        End If

        Dim db As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If db.GravaBanco(Sqls) Then
            IUD = ""
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""
        Select Case _IUD
            Case "I"
                sql = " INSERT INTO NotaFiscalReferencial" & vbCrLf & _
                      " (EmpresaReferencial_Id" & vbCrLf & _
                      " ,EndEmpresaReferencial_Id" & vbCrLf & _
                      " ,ClienteReferencial_Id" & vbCrLf & _
                      " ,EndClienteReferencial_Id" & vbCrLf & _
                      " ,EntradaSaidaReferencial_Id" & vbCrLf & _
                      " ,SerieReferencial_Id" & vbCrLf & _
                      " ,NotaReferencial_Id" & vbCrLf & _
                      " ,ProdutoReferencial_Id" & vbCrLf & _
                      " ,CFOPReferencial_Id" & vbCrLf & _
                      " ,SequenciaReferencial_Id" & vbCrLf & _
                      " ,TipoReferencial_Id" & vbCrLf & _
                      " ,Empresa_Id" & vbCrLf & _
                      " ,EndEmpresa_Id" & vbCrLf & _
                      " ,Cliente_Id" & vbCrLf & _
                      " ,EndCliente_Id" & vbCrLf & _
                      " ,EntradaSaida_Id" & vbCrLf & _
                      " ,Serie_Id" & vbCrLf & _
                      " ,Nota_Id" & vbCrLf & _
                      " ,Produto_Id" & vbCrLf & _
                      " ,CFOP_Id" & vbCrLf & _
                      " ,Sequencia_Id" & vbCrLf & _
                      " ,Quantidade" & vbCrLf & _
                      " ,Valor)" & vbCrLf & _
                      " VALUES" & vbCrLf & _
                      " ('" & Me.Parent.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      " ,'" & Me.Parent.NotaFiscal.EnderecoEmpresa & "'" & vbCrLf & _
                      " ,'" & Me.Parent.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      " ,'" & Me.Parent.NotaFiscal.EnderecoCliente & "'" & vbCrLf & _
                      " ,'" & IIf(Me.Parent.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'  " & vbCrLf & _
                      " ,'" & Me.Parent.NotaFiscal.Serie & "'" & vbCrLf & _
                      " ,'" & Me.Parent.NotaFiscal.Codigo & "'" & vbCrLf & _
                      " ,'" & Me.Parent.CodigoProduto & "'" & vbCrLf & _
                      " ,'" & Me.Parent.CFOP & "'" & vbCrLf & _
                      " ,'" & Me.Parent.Sequencia & "'" & vbCrLf & _
                      " ,'" & Me.TipoReferencial_Id.ToString & "'" & vbCrLf & _
                      " ,'" & Me.Empresa_Id & "'" & vbCrLf & _
                      " ,'" & Me.EndEmpresa_Id & "'" & vbCrLf & _
                      " ,'" & Me.Cliente_Id & "'" & vbCrLf & _
                      " ,'" & Me.EndCliente_Id & "'" & vbCrLf & _
                      " ,'" & IIf(Me.EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "'  " & vbCrLf & _
                      " ,'" & Me.Serie_Id & "'" & vbCrLf & _
                      " ,'" & Me.Nota_Id & "'" & vbCrLf & _
                      " ,'" & Me.Produto_Id & "'" & vbCrLf & _
                      " ,'" & Me.CFOP_Id & "'" & vbCrLf & _
                      " ,'" & Me.Sequencia_Id & "'" & vbCrLf & _
                      " ," & Str(Me.Quantidade) & "" & vbCrLf & _
                      " ," & Str(Me.Valor) & ")" & vbCrLf
                Sqls.Add(sql)

            Case "U"
                sql = " UPDATE NotaFiscalReferencial" & vbCrLf & _
                      " SET" & vbCrLf & _
                      " Quantidade = " & Str(Me.Quantidade) & "" & vbCrLf & _
                      " ,Valor = " & Str(Me.Valor) & "" & vbCrLf & _
                      " WHERE 1=1" & vbCrLf & _
                      " AND EmpresaReferencial_Id = '" & Me.Parent.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      " AND EndEmpresaReferencial_Id = '" & Me.Parent.NotaFiscal.EnderecoEmpresa & "'" & vbCrLf & _
                      " AND ClienteReferencial_Id = '" & Me.Parent.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      " AND EndClienteReferencial_Id = '" & Me.Parent.NotaFiscal.EnderecoCliente & "'" & vbCrLf & _
                      " AND EntradaSaidaReferencial_Id = '" & IIf(Me.Parent.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'  " & vbCrLf & _
                      " AND SerieReferencial_Id = '" & Me.Parent.NotaFiscal.Serie & "'" & vbCrLf & _
                      " AND NotaReferencial_Id = '" & Me.Parent.NotaFiscal.Codigo & "'" & vbCrLf & _
                      " AND ProdutoReferencial_Id = '" & Me.Parent.CodigoProduto & "'" & vbCrLf & _
                      " AND CFOPReferencial_Id = '" & Me.Parent.CFOP & "'" & vbCrLf & _
                      " AND SequenciaReferencial_Id = '" & Me.Parent.Sequencia & "'" & vbCrLf & _
                      " AND TipoReferencial_Id = '" & Me.TipoReferencial_Id.ToString & "'" & vbCrLf & _
                      " AND Empresa_Id = '" & Me.Empresa_Id & "'" & vbCrLf & _
                      " AND EndEmpresa_Id = '" & Me.EndEmpresa_Id & "'" & vbCrLf & _
                      " AND Cliente_Id = '" & Me.Cliente_Id & "'" & vbCrLf & _
                      " AND EndCliente_Id = '" & Me.EndCliente_Id & "'" & vbCrLf & _
                      " AND EntradaSaida_Id = '" & IIf(Me.EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      " AND Serie_Id = '" & Me.Serie_Id & "'" & vbCrLf & _
                      " AND Nota_Id = '" & Me.Nota_Id & "'" & vbCrLf & _
                      " AND Produto_Id = '" & Me.Produto_Id & "'" & vbCrLf & _
                      " AND CFOP_Id = '" & Me.CFOP_Id & "'" & vbCrLf & _
                      " AND Sequencia_Id = '" & Me.Sequencia_Id & "'" & vbCrLf
                Sqls.Add(sql)

            Case "D"
                sql = " DELETE FROM NotaFiscalReferencial " & vbCrLf & _
                      " WHERE 1=1  " & vbCrLf & _
                      " AND EmpresaReferencial_Id = '" & Me.Parent.NotaFiscal.CodigoEmpresa & "' " & vbCrLf & _
                      " AND EndEmpresaReferencial_Id = '" & Me.Parent.NotaFiscal.EnderecoEmpresa & "'" & vbCrLf & _
                      " AND ClienteReferencial_Id = '" & Me.Parent.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      " AND EndClienteReferencial_Id = '" & Me.Parent.NotaFiscal.EnderecoCliente & "'" & vbCrLf & _
                      " AND EntradaSaidaReferencial_Id = '" & IIf(Me.Parent.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'  " & vbCrLf & _
                      " AND SerieReferencial_Id = '" & Me.Parent.NotaFiscal.Serie & "'" & vbCrLf & _
                      " AND NotaReferencial_Id = '" & Me.Parent.NotaFiscal.Codigo & "'" & vbCrLf & _
                      " AND ProdutoReferencial_Id = '" & Me.Parent.CodigoProduto & "'" & vbCrLf & _
                      " AND CFOPReferencial_Id = '" & Me.Parent.CFOP & "'" & vbCrLf & _
                      " AND SequenciaReferencial_Id = '" & Me.Parent.Sequencia & "'" & vbCrLf & _
                      " AND TipoReferencial_Id = '" & Me.TipoReferencial_Id.ToString & "'" & vbCrLf & _
                      " AND Empresa_Id = '" & Me.Empresa_Id & "'" & vbCrLf & _
                      " AND EndEmpresa_Id = '" & Me.EndEmpresa_Id & "'" & vbCrLf & _
                      " AND Cliente_Id = '" & Me.Cliente_Id & "'" & vbCrLf & _
                      " AND EndCliente_Id = '" & Me.EndCliente_Id & "'" & vbCrLf & _
                      " AND EntradaSaida_Id = '" & IIf(Me.EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      " AND Serie_Id = '" & Me.Serie_Id & "'" & vbCrLf & _
                      " AND Nota_Id = '" & Me.Nota_Id & "'" & vbCrLf & _
                      " AND Produto_Id = '" & Me.Produto_Id & "'" & vbCrLf & _
                      " AND CFOP_Id = '" & Me.CFOP_Id & "'" & vbCrLf & _
                      " AND Sequencia_Id = '" & Me.Sequencia_Id & "'" & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub

#End Region

End Class
