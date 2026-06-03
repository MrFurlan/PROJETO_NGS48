Imports System.Web

<Serializable()>
Public Class ListHistoricoFinanceiro
    Inherits List(Of HistoricoFinanceiro)

    Public Sub New()

    End Sub

    Public Sub New(ByVal nf As NotaFiscal)
        Dim sql As String

        sql = "SELECT [Empresa_Id]" & vbCrLf &
              ",[EndEmpresa_Id] " & vbCrLf &
              ",[Historico] " & vbCrLf &
              ",[EntradaSaida] " & vbCrLf &
              ",[Serie] " & vbCrLf &
              ",[Nota] " & vbCrLf &
              ",[Pedido] " & vbCrLf &
              ",[Provisao] " & vbCrLf &
              ",[Titulo] " & vbCrLf &
              ",[TituloOrigem] " & vbCrLf &
              ",[Valor]" & vbCrLf &
              ",[Data]" & vbCrLf &
              ",[Usuario]" & vbCrLf &
              "FROM [dbo].[HistoricoFinanceiro]" & vbCrLf &
              "WHERE Empresa_Id = " & nf.CodigoEmpresa & vbCrLf &
              "  AND Nota       = " & nf.Codigo & vbCrLf &
              "  AND Serie        = '" & nf.Serie & "'" & vbCrLf &
              "  AND Pedido       = " & nf.CodigoPedido & vbCrLf &
              "  AND EntradaSaida = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Parametros")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim hf As New HistoricoFinanceiro

            hf.CodigoEmpresa = row("Empresa_Id")
            hf.EnderecoEmpresa = row("EndEmpresa_Id")
            hf.Historico = row("Historico")
            hf.Serie = row("Serie")
            hf.EntradaSaida = row("EntradaSaida")
            hf.Nota = CInt(row("Nota"))
            hf.Pedido = CInt(row("Pedido"))
            hf.CodigoProvisao = CInt(row("Provisao"))
            hf.Titulo = row("Titulo")
            hf.TituloOrigem = row("TituloOrigem")
            hf.Valor = CDec(row("Valor"))
            hf.Data = CDate(row("Data"))
            hf.Usuario = row("Usuario")

            Me.Add(hf)
        Next
    End Sub
End Class

<Serializable()>
Public Class HistoricoFinanceiro
    Implements IBaseEntity

#Region "Contrutores"
    Public Sub New()

    End Sub

    Public Sub New(ByVal tit As Titulo, ByVal nf As NotaFiscal, ByVal titOrigem As Integer, ByVal valor As Decimal, Optional ByVal historico As String = "")
        Me.CodigoEmpresa = tit.CodigoEmpresa
        Me.EnderecoEmpresa = tit.EnderecoEmpresa
        Me.Titulo = tit.Codigo
        Me.TituloOrigem = titOrigem
        Me.Valor = valor
        Me.Historico = IIf(String.IsNullOrEmpty(historico), tit.Historico, historico)
        Me.CodigoProvisao = tit.CodigoProvisao
        Me.Pedido = tit.CodigoPedido
        Me.Usuario = tit.UsuarioInclusao
        Me.Nota = nf.Codigo
        Me.Serie = nf.Serie
        Me.EntradaSaida = IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
    End Sub

#End Region

#Region "Fields"
    Private _Codigo As Integer
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _Historico As String
    Private _Titulo As Integer
    Private _CodigoProvisao As Integer
    Private _Pedido As Integer
    Private _TituloOrigem As Integer
    Private _Valor As Decimal
    Private _Usuario As String
    Private _Data As DateTime
    Private _Nota As Integer
    Private _Serie As String
    Private _EntradaSaida As String

#End Region

#Region "Property"
    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property
    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property TituloOrigem() As Integer
        Get
            Return _TituloOrigem
        End Get
        Set(ByVal value As Integer)
            _TituloOrigem = value
        End Set
    End Property

    Public Property Pedido() As Integer
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Integer)
            _Pedido = value
        End Set
    End Property

    Public Property Titulo() As Integer
        Get
            Return _Titulo
        End Get
        Set(ByVal value As Integer)
            _Titulo = value
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

    Public Property Historico() As String
        Get
            Return _Historico
        End Get
        Set(ByVal value As String)
            _Historico = value
        End Set
    End Property

    Public Property Usuario() As String
        Get
            Return _Usuario
        End Get
        Set(ByVal value As String)
            _Usuario = value
        End Set
    End Property

    Public Property Data() As DateTime
        Get
            Return _Data
        End Get
        Set(ByVal value As DateTime)
            _Data = value
        End Set
    End Property

    Public Property CodigoProvisao() As Integer
        Get
            Return _CodigoProvisao
        End Get
        Set(ByVal value As Integer)
            _CodigoProvisao = value
        End Set
    End Property

    Public Property Nota() As Integer
        Get
            Return _Nota
        End Get
        Set(ByVal value As Integer)
            _Nota = value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
        End Set
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Incluir(h As HistoricoFinanceiro) As String
        Dim Banco As New AcessaBanco

        Dim strSql = "INSERT INTO [dbo].[HistoricoFinanceiro]([Empresa_Id], [EndEmpresa_Id], [Historico], [EntradaSaida], [Serie], [Nota], [Pedido], [Provisao], [Titulo], [TituloOrigem], [Valor], [Usuario])" & vbCrLf &
                         "Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.Historico & "','" & Me.EntradaSaida & "','" & Me.Serie & "'," & Me.Nota & "," & Me.Pedido & "," & Me.CodigoProvisao & "," & Me.Titulo & "," & Me.TituloOrigem & "," & Str(Me.Valor) & ",'" & Me.Usuario & "')"

        Return strSql
    End Function

    Public Function AdicionarHistoricoPedido(mf As MovimentacaoFinanceira) As String
        Me.CodigoEmpresa = mf.CodigoEmpresa
        Me.EnderecoEmpresa = mf.EnderecoEmpresa
        Me.Titulo = mf.Codigo
        Me.TituloOrigem = 0
        Me.Valor = mf.ValorDocumentoOficial
        Me.Historico = "Adicionado Provisão do Pedido " & mf.CodigoPedido
        Me.CodigoProvisao = eProvisao.Provisao
        Me.Nota = 0
        Me.Pedido = mf.CodigoPedido
        Return Incluir(Me)
    End Function

    Public Function AdicionarHistoricoNF() As String
        Return Incluir(Me)
    End Function
#End Region

End Class
