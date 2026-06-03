Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMemorandoDeExportacao
    Inherits List(Of MemorandoDeExportacao)

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Where As String)
        Dim Banco As New AcessaBanco
        Dim sql As String = ""
        sql = "Select Empresa_Id, EndEmpresa_Id, Memorando_Id, DataMemorando, NrDespachoExp, DataDespachoExp, DataAverbacao" & vbCrLf & _
              "       PaisDestino, Navio, Representante, Cliente, EndCliente, Nota, Serie, EntradaSaida, Produto, ExportadorEquiparado, " & vbCrLf & _
              "       isnull(EndExportadorEquiparado,0) as  EndExportadorEquiparado,  isnull(MemorandoEquiparado,0) as  MemorandoEquiparado, " & vbCrLf & _
              "       isnull(UsuarioInclusao,'') as  UsuarioInclusao,  isnull(DataInclusao,'') as  DataInclusao, " & vbCrLf & _
              "       isnull(UsuarioAlteracao,'') as  UsuarioAlteracao,  isnull(DataAlteracao,'') as  DataAlteracao, " & vbCrLf & _
              "       isnull(DataDaNota,'') as DataDaNota, ValorNota, Moeda, Indexador, TipoDocumento  " & vbCrLf & _
              "  From MemorandoDeExportacao " & vbCrLf & _
              " Where " & Where

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Memorando")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Mem As New MemorandoDeExportacao
            Mem.CodigoEmpresa = row("Empresa_Id")
            Mem.EnderecoEmpresa = row("EndEmpresa_Id")
            Mem.CodigoMemorando = row("Memorando_Id")
            Mem.DataMemorando = row("DataMemorando")
            Mem.NrDespachoExp = row("NrDespachoExp")
            Mem.DataDespachoExp = IIf(row("NrDespachoExp").ToString.Length = 0, row("DataMemorando"), row("DataDespachoExp"))
            Mem.DataAverbacao = row("DataAverbacao")
            Mem.CodigoPaisDestino = row("PaisDestino")
            Mem.Navio = row("Navio")
            Mem.Representante = row("Representante")
            Mem.CodigoCliente = row("Cliente")
            Mem.EnderecoCliente = row("EndCliente")
            Mem.NumeroNota = row("Nota")
            Mem.Serie = row("Serie")
            Mem.EntradaSaida = row("EntradaSaida")
            Mem.CodigoProduto = row("Produto")
            Mem.CodExportadorEquiparado = row("ExportadorEquiparado")
            Mem.EnderecoExportadorEquiparado = row("EndExportadorEquiparado")
            Mem.CodigoMemorandoEquiparado = row("MemorandoEquiparado")
            Mem.UsuarioInclusao = row("UsuarioInclusao")
            Mem.DataInclusao = row("DataInclusao")
            Mem.UsuarioAlteracao = row("UsuarioAlteracao")
            Mem.DataAlteracao = row("DataAlteracao")
            Mem.DataNota = row("DataNota")
            Mem.ValorNota = row("ValorNota")
            Mem.Moeda = row("Moeda")
            Mem.Indexador = row("Indexador")
            Mem.TipoDocumento = row("TipoDocumento")

            Me.Add(Mem)
        Next

        Banco = Nothing
    End Sub

#End Region

End Class

<Serializable()> _
Public Class MemorandoDeExportacao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pMemorando As String)
        Dim Banco As New AcessaBanco
        Dim sql As String = ""
        sql = "Select EmpresaMemorando_Id, EndEmpresaMemorando_Id, Memorando_Id, ClienteMemorando, EndClienteMemorando, DataMemorando, NrDespachoExp, isnull(DataDespachoExp,'') as DataDespachoExp, isnull(DataAverbacao,'') as DataAverbacao," & vbCrLf & _
              "       PaisDestino, Navio, isnull(Representante,'') as Representante, Empresa, EndEmpresa, Cliente, EndCliente, Nota, Serie, EntradaSaida, Produto, NossaEmissao, isnull(ExportadorEquiparado,'') as ExportadorEquiparado,  " & vbCrLf & _
              "       isnull(EndExportadorEquiparado,0) as  EndExportadorEquiparado, isnull(MemorandoEquiparado,0) as  MemorandoEquiparado, " & vbCrLf & _
              "       isnull(UsuarioInclusao,'') as  UsuarioInclusao,  isnull(DataInclusao,'') as  DataInclusao, " & vbCrLf & _
              "       isnull(UsuarioAlteracao,'') as  UsuarioAlteracao,  isnull(DataAlteracao,'') as  DataAlteracao, " & vbCrLf & _
              "       isnull(DataDaNota,'') as DataDaNota, isnull(ValorNota,0) as ValorNota, isnull(Moeda,0) as Moeda, isnull(Indexador,0) as Indexador, isnull(TipoDocumento,0) as TipoDocumento, EndRepresentante  " & vbCrLf & _
              "  From MemorandoDeExportacao " & vbCrLf & _
              " Where EmpresaMemorando_Id    ='" & pEmpresa & "'" & vbCrLf & _
              "   and EndEmpresaMemorando_Id = " & pEndEmpresa & vbCrLf & _
              "   and Memorando_Id           = '" & pMemorando & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Memorando")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _CodigoEmpresaMemorando = row("EmpresaMemorando_Id")
        _EnderecoEmpresaMemorando = row("EndEmpresaMemorando_Id")
        _CodigoMemorando = row("Memorando_Id")
        _CodigoClienteMemorando = row("ClienteMemorando")
        _EnderecoClienteMemorando = row("EndClienteMemorando")
        _DataMemorando = row("DataMemorando")
        _NrDespachoExp = row("NrDespachoExp")
        _DataDespachoExp = IIf(row("NrDespachoExp").ToString.Length = 0, row("DataMemorando"), row("DataDespachoExp"))
        _DataAverbacao = row("DataAverbacao")
        _CodigoPaisDestino = row("PaisDestino")
        _Navio = row("Navio")
        _Representante = row("Representante")
        _CodigoRepresentante = row("Representante")
        _EnderecoRepresentante = IIf(IsDBNull(row("EndRepresentante")), 0, row("EndRepresentante"))
        _CodigoEmpresa = row("Empresa")
        _EnderecoEmpresa = row("EndEmpresa")
        _CodigoCliente = row("Cliente")
        _EnderecoCliente = row("EndCliente")
        _NumeroNota = row("Nota")
        _Serie = row("Serie")
        _EntradaSaida = row("EntradaSaida")
        _CodigoProduto = row("Produto")
        _NossaEmissao = row("NossaEmissao")
        _CodExportadorEquiparado = row("ExportadorEquiparado")
        _EnderecoExportadorEquiparado = row("EndExportadorEquiparado")
        _CodigoMemorandoEquiparado = row("MemorandoEquiparado")
        _UsuarioInclusao = row("UsuarioInclusao")
        _DataInclusao = row("DataInclusao")
        _UsuarioAlteracao = row("UsuarioAlteracao")
        _DataAlteracao = row("DataAlteracao")
        _DataNota = row("DataDaNota")
        _ValorNota = row("ValorNota")
        _Moeda = row("Moeda")
        _Indexador = row("Indexador")
        _TipoDocumento = row("TipoDocumento")

        Banco = Nothing
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoEmpresaMemorando As String = ""
    Private _EnderecoEmpresaMemorando As Integer
    Private _EmpresaMemorando As Cliente

    Private _CodExportadorEquiparado As String = ""
    Private _EnderecoExportadorEquiparado As Integer
    Private _ExportadorEquiparadoMemorando As Cliente

    Private _CodigoClienteMemorando As String = ""
    Private _EnderecoClienteMemorando As Integer
    Private _ClienteMemorando As Cliente

    Private _CodigoMemorando As String

    Private _CodigoMemorandoEquiparado As String

    Private _DataMemorando As Date = Now
    Private _NrDespachoExp As String = ""
    Private _DataDespachoExp As Date = Now
    Private _DataAverbacao As Date = Now

    Private _DataAverbacaoDecEmbarque As Date = Now
    Private _CodigoTipoConhecEmbarque As Integer

    Private _CodigoPaisDestino As Integer
    Private _Navio As String

    Private _Representante As String = ""

    Private _NossaEmissao As Boolean = False
    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Cliente

    Private _CodigoCliente As String = ""
    Private _EnderecoCliente As Integer
    Private _Cliente As Cliente

    Private _CodigoRepresentante As String = ""
    Private _EnderecoRepresentante As String = ""
    Private _RepresentanteMemorando As Cliente

    Private _NumeroNota As Integer
    Private _Serie As String = ""
    Private _EntradaSaida As String = ""

    Private _CodigoProduto As String = ""
    Private _Produto As Produto

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As Date = Now

    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As Date = Now

    Private _DataNota As Date = Now
    Private _Moeda As Integer
    Private _Indexador As Integer

    Private _ValorNota As Decimal

    Private _NumAtoConcessorio As String = ""
    Private _DtaRegAtoConcessorio As Date = Now
    Private _DtaValidAtoConcessorio As Date = Now

    Private _TipoDocumento As String = ""
    Private _DescTipoDocumento As String = ""

    Private _NotasComprovadas As ListMemorandoDeExportacaoXNotaFiscal
    Private _ConhecimentosDeEmbarque As ListMemorandoDeExportacaoXConhecimento
    Private _RegistrosDeExportacao As ListMemorandoDeExportacaoXRegistroDeExportacao
    Private _NotaDeComprovacao As NotaDeExportacao

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

    Public Property CodigoEmpresaMemorando() As String
        Get
            Return _CodigoEmpresaMemorando
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaMemorando = value
            _EmpresaMemorando = Nothing
            _NotasComprovadas = Nothing
            _RegistrosDeExportacao = Nothing
        End Set
    End Property

    Public Property EnderecoEmpresaMemorando() As Integer
        Get
            Return _EnderecoEmpresaMemorando
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresaMemorando = value
            _EmpresaMemorando = Nothing
            _NotasComprovadas = Nothing
            _RegistrosDeExportacao = Nothing
        End Set
    End Property

    Public ReadOnly Property EmpresaMemorando() As Cliente
        Get
            If _EmpresaMemorando Is Nothing And _CodigoEmpresaMemorando.Length > 0 Then _EmpresaMemorando = New Cliente(_CodigoEmpresaMemorando, _EnderecoEmpresaMemorando)
            Return _EmpresaMemorando
        End Get
    End Property

    Public Property CodigoMemorando() As String
        Get
            Return _CodigoMemorando
        End Get
        Set(ByVal value As String)
            _CodigoMemorando = value
        End Set
    End Property

    Public Property CodExportadorEquiparado() As String
        Get
            Return _CodExportadorEquiparado
        End Get
        Set(ByVal value As String)
            _CodExportadorEquiparado = value
            _ExportadorEquiparadoMemorando = Nothing
            _NotasComprovadas = Nothing
            _RegistrosDeExportacao = Nothing
        End Set
    End Property

    Public Property EnderecoExportadorEquiparado() As Integer
        Get
            Return _EnderecoExportadorEquiparado
        End Get
        Set(ByVal value As Integer)
            _EnderecoExportadorEquiparado = value
            _ExportadorEquiparadoMemorando = Nothing
            _NotasComprovadas = Nothing
            _RegistrosDeExportacao = Nothing
        End Set
    End Property

    Public ReadOnly Property ExportadorEquiparadoMemorando() As Cliente
        Get
            If _ExportadorEquiparadoMemorando Is Nothing And _CodExportadorEquiparado.Length > 0 Then _ExportadorEquiparadoMemorando = New Cliente(_CodExportadorEquiparado, _EnderecoExportadorEquiparado)
            Return _ExportadorEquiparadoMemorando
        End Get
    End Property

    Public Property CodigoClienteMemorando() As String
        Get
            Return _CodigoClienteMemorando
        End Get
        Set(ByVal value As String)
            _CodigoClienteMemorando = value
            _ClienteMemorando = Nothing
            _NotasComprovadas = Nothing
            _RegistrosDeExportacao = Nothing
        End Set
    End Property

    Public Property EnderecoClienteMemorando() As Integer
        Get
            Return _EnderecoClienteMemorando
        End Get
        Set(ByVal value As Integer)
            _EnderecoClienteMemorando = value
            _ClienteMemorando = Nothing
            _NotasComprovadas = Nothing
            _RegistrosDeExportacao = Nothing
        End Set
    End Property

    Public ReadOnly Property ClienteMemorando() As Cliente
        Get
            If _ClienteMemorando Is Nothing And _CodigoClienteMemorando.Length > 0 Then _ClienteMemorando = New Cliente(_CodigoClienteMemorando, _EnderecoClienteMemorando)
            Return _ClienteMemorando
        End Get
    End Property

    Public Property DataMemorando() As Date
        Get
            Return _DataMemorando
        End Get
        Set(ByVal value As Date)
            _DataMemorando = value
        End Set
    End Property

    Public Property NrDespachoExp() As String
        Get
            Return _NrDespachoExp
        End Get
        Set(ByVal value As String)
            _NrDespachoExp = value
        End Set
    End Property

    Public Property DataDespachoExp() As Date
        Get
            Return _DataDespachoExp
        End Get
        Set(ByVal value As Date)
            _DataDespachoExp = value
        End Set
    End Property


    Public Property DataAverbacao() As Date
        Get
            Return _DataAverbacao
        End Get
        Set(ByVal value As Date)
            _DataAverbacao = value
        End Set
    End Property

    Public Property CodigoPaisDestino() As Integer
        Get
            Return _CodigoPaisDestino
        End Get
        Set(ByVal value As Integer)
            _CodigoPaisDestino = value
        End Set
    End Property

    Public Property DataAverbacaoDecEmbarque() As Date
        Get
            Return _DataAverbacaoDecEmbarque
        End Get
        Set(ByVal value As Date)
            _DataAverbacaoDecEmbarque = value
        End Set
    End Property

    Public Property CodigoTipoConhecEmbarque() As Integer
        Get
            Return _CodigoTipoConhecEmbarque
        End Get
        Set(ByVal value As Integer)
            _CodigoTipoConhecEmbarque = value
        End Set
    End Property

    Public Property Navio() As String
        Get
            Return _Navio
        End Get
        Set(ByVal value As String)
            _Navio = value
        End Set
    End Property

    Public Property Representante() As String
        Get
            Return _Representante
        End Get
        Set(ByVal value As String)
            _Representante = value
        End Set
    End Property

    Public Property CodigoRepresentante() As String
        Get
            Return _CodigoRepresentante
        End Get
        Set(ByVal value As String)
            _CodigoRepresentante = value
        End Set
    End Property

    Public Property EnderecoRepresentante() As String
        Get
            Return _EnderecoRepresentante
        End Get
        Set(ByVal value As String)
            _EnderecoRepresentante = value
        End Set
    End Property

    Public ReadOnly Property RepresentanteMemorando() As Cliente
        Get
            If _RepresentanteMemorando Is Nothing And _CodigoRepresentante.Length > 0 Then _RepresentanteMemorando = New Cliente(_CodigoRepresentante, _EnderecoRepresentante)
            Return _RepresentanteMemorando
        End Get
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
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
        End Set
    End Property

    Public ReadOnly Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EnderecoEmpresa)
            Return _Empresa
        End Get
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public ReadOnly Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EnderecoCliente)
            Return _Cliente
        End Get
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value.ToUpper
        End Set
    End Property

    Public Property NumeroNota() As Integer
        Get
            Return _NumeroNota
        End Get
        Set(ByVal value As Integer)
            _NumeroNota = value
        End Set
    End Property

    Public Property NossaEmissao() As Boolean
        Get
            Return _NossaEmissao
        End Get
        Set(ByVal value As Boolean)
            _NossaEmissao = value
        End Set
    End Property

    Public Property DataNota() As Date
        Get
            Return _DataNota
        End Get
        Set(ByVal value As Date)
            _DataNota = value

        End Set
    End Property

    Public Property Moeda() As Integer
        Get
            Return _Moeda
        End Get
        Set(ByVal value As Integer)
            _Moeda = value

        End Set
    End Property

    Public Property Indexador() As Integer
        Get
            Return _Indexador
        End Get
        Set(ByVal value As Integer)
            _Indexador = value

        End Set
    End Property

    Public Property ValorNota() As Decimal
        Get
            Return _ValorNota
        End Get
        Set(ByVal value As Decimal)
            _ValorNota = value
        End Set
    End Property

    Public Property NumAtoConcessorio() As String
        Get
            Return _NumAtoConcessorio
        End Get
        Set(ByVal value As String)
            _NumAtoConcessorio = value.ToUpper
        End Set
    End Property

    Public Property DtaRegAtoConcessorio() As Date
        Get
            Return _DtaRegAtoConcessorio
        End Get
        Set(ByVal value As Date)
            _DtaRegAtoConcessorio = value

        End Set
    End Property

    Public Property DtaValidAtoConcessorio() As Date
        Get
            Return _DtaValidAtoConcessorio
        End Get
        Set(ByVal value As Date)
            _DtaValidAtoConcessorio = value

        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Produto = Nothing
        End Set
    End Property

    Public ReadOnly Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
    End Property

    Public Property NotasComprovadas() As ListMemorandoDeExportacaoXNotaFiscal
        Get
            If _NotasComprovadas Is Nothing Then _NotasComprovadas = New ListMemorandoDeExportacaoXNotaFiscal(Me)
            Return _NotasComprovadas
        End Get
        Set(ByVal value As ListMemorandoDeExportacaoXNotaFiscal)
            _NotasComprovadas = value
        End Set
    End Property

    Public Property ConhecimentosDeEmbarque() As ListMemorandoDeExportacaoXConhecimento
        Get
            If _ConhecimentosDeEmbarque Is Nothing Then _ConhecimentosDeEmbarque = New ListMemorandoDeExportacaoXConhecimento(Me)
            Return _ConhecimentosDeEmbarque
        End Get
        Set(ByVal value As ListMemorandoDeExportacaoXConhecimento)
            _ConhecimentosDeEmbarque = value
        End Set
    End Property

    Public Property RegistrosDeExportacao() As ListMemorandoDeExportacaoXRegistroDeExportacao
        Get
            If _RegistrosDeExportacao Is Nothing Then _RegistrosDeExportacao = New ListMemorandoDeExportacaoXRegistroDeExportacao(Me)
            Return _RegistrosDeExportacao
        End Get
        Set(ByVal value As ListMemorandoDeExportacaoXRegistroDeExportacao)
            _RegistrosDeExportacao = value
        End Set
    End Property

    Public Property NotaDeComprovacao() As NotaDeExportacao
        Get
            If _NotaDeComprovacao Is Nothing Then _NotaDeComprovacao = New NotaDeExportacao(Me)
            Return _NotaDeComprovacao
        End Get
        Set(ByVal value As NotaDeExportacao)
            _NotaDeComprovacao = value
        End Set
    End Property

    Public ReadOnly Property QuantidadeMemorando() As Decimal
        Get
            If NotasComprovadas Is Nothing Then Return 0
            Dim qtde As Decimal = 0
            For Each nf As MemorandoDeExportacaoXNotaFiscal In NotasComprovadas
                qtde += nf.QuantidadeMemorando
            Next
            Return qtde
        End Get
    End Property

    Public Property CodigoMemorandoEquiparado() As String
        Get
            Return _CodigoMemorandoEquiparado
        End Get
        Set(ByVal value As String)
            _CodigoMemorandoEquiparado = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As Date
        Get
            Return _DataInclusao
        End Get
        Set(ByVal value As Date)
            _DataInclusao = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property DataAlteracao() As Date
        Get
            Return _DataAlteracao
        End Get
        Set(ByVal value As Date)
            _DataAlteracao = value
        End Set
    End Property

    Public Property TipoDocumento() As String
        Get
            Return _TipoDocumento
        End Get
        Set(ByVal value As String)
            _TipoDocumento = value
        End Set
    End Property

    Public ReadOnly Property DescTipoDocumento() As String
        Get
            Select Case _DescTipoDocumento
                Case "0" : Return "Declaração De Exportação"
                Case "1" : Return "Declaração Simplificada De Exportação"
            End Select
            Return String.Empty
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
                sql = " Insert Into MemorandoDeExportacao(EmpresaMemorando_Id, EndEmpresaMemorando_Id, Memorando_Id, ClienteMemorando, EndClienteMemorando, DataMemorando, NrDespachoExp, DataDespachoExp, DataAverbacao, " & vbCrLf & _
                      "                                   PaisDestino, Navio, Representante, Empresa, EndEmpresa, Cliente, EndCliente, Nota, Serie, EntradaSaida, Produto,  NossaEmissao, ExportadorEquiparado, EndExportadorEquiparado, MemorandoEquiparado, " & vbCrLf & _
                      "                                   UsuarioInclusao, DataInclusao, UsuarioAlteracao, DataAlteracao, DataDaNota, ValorNota, Moeda, Indexador, TipoDocumento, NumAtoConcessorio, DtaRegAtoConcessorio, DtaValidAtoConcessorio, EndRepresentante) " & vbCrLf & _
                      " Values('" & _CodigoEmpresaMemorando & "'," & _EnderecoEmpresaMemorando & ",'" & _CodigoMemorando & "','" & _CodigoClienteMemorando & "'," & _EnderecoClienteMemorando & ",'" & _DataMemorando.ToString("yyyy-MM-dd") & "','" & _NrDespachoExp & "'," & "'" & _DataDespachoExp.ToString("yyyy-MM-dd") & "'" & "," & "'" & _DataAverbacao.ToString("yyyy-MM-dd") & "'" & "," & vbCrLf & _
                      "         " & _CodigoPaisDestino & ",'" & _Navio & "','" & _CodigoRepresentante & "','" & _CodigoEmpresa & "'," & _EnderecoEmpresa & ",'" & _CodigoCliente & "'," & _EnderecoCliente & "," & _NumeroNota & ",'" & _Serie & "','" & _EntradaSaida & "','" & _CodigoProduto & "'," & CByte(_NossaEmissao) & ",'" & _CodExportadorEquiparado & "'," & _EnderecoExportadorEquiparado & ",'" & _CodigoMemorandoEquiparado & "'," & vbCrLf & _
                      "         '" & _UsuarioInclusao & "','" & _DataInclusao.ToString("yyyy-MM-dd") & "','" & _UsuarioAlteracao & "','" & _DataAlteracao.ToString("yyyy-MM-dd") & "', '" & _DataNota.ToString("yyyy-MM-dd") & "', " & Str(_ValorNota) & ", " & _Moeda & ", " & _Indexador & " , '" & _TipoDocumento & "','" & _NumAtoConcessorio & "'," & "'" & _DtaRegAtoConcessorio.ToString("yyyy-MM-dd") & "'" & "," & "'" & _DtaValidAtoConcessorio.ToString("yyyy-MM-dd") & "'," & _EnderecoRepresentante & ")"
                Sqls.Add(sql)
                NotasComprovadas.SalvarSql(Sqls)
                ConhecimentosDeEmbarque.SalvarSql(Sqls)
                RegistrosDeExportacao.SalvarSql(Sqls)
            Case "U"
                sql = "Update MemorandoDeExportacao set" & vbCrLf & _
                      "   DataMemorando         ='" & _DataMemorando.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "  ,NrDespachoExp         ='" & _NrDespachoExp & "'" & vbCrLf & _
                      "  ,DataDespachoExp       = " & IIf(_NrDespachoExp.Length = 0, "NULL", "'" & _DataDespachoExp.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                      "  ,DataAverbacao         ='" & _DataAverbacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "  ,PaisDestino           = " & _CodigoPaisDestino & vbCrLf & _
                      "  ,Navio                 ='" & _Navio & "'" & vbCrLf & _
                      "  ,Representante         ='" & _CodigoRepresentante & "'" & vbCrLf & _
                      "  ,EndRepresentante      ='" & _EnderecoRepresentante & "'" & vbCrLf & _
                      "  ,Empresa               ='" & _CodigoEmpresa & "'" & vbCrLf & _
                      "  ,EndEmpresa            =  " & _EnderecoEmpresa & vbCrLf & _
                      "  ,Cliente               ='" & _CodigoCliente & "'" & vbCrLf & _
                      "  ,EndCliente            = " & _EnderecoCliente & vbCrLf & _
                      "  ,Nota                  = " & _NumeroNota & vbCrLf & _
                      "  ,Serie                 ='" & _Serie & "'" & vbCrLf & _
                      "  ,EntradaSaida          ='" & _EntradaSaida & "'" & vbCrLf & _
                      "  ,Produto               ='" & _CodigoProduto & "'" & vbCrLf & _
                      "  ,NossaEmissao          = " & CByte(_NossaEmissao) & vbCrLf & _
                      "  ,NumAtoConcessorio     ='" & _NumAtoConcessorio & "'" & vbCrLf & _
                      "  ,DtaRegAtoConcessorio  = " & IIf(_NumAtoConcessorio.Length = 0, "NULL", "'" & _DtaRegAtoConcessorio.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                      "  ,DtaValidAtoConcessorio= " & IIf(_NumAtoConcessorio.Length = 0, "NULL", "'" & _DtaValidAtoConcessorio.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                      "  ,ExportadorEquiparado  ='" & _CodExportadorEquiparado & "'" & vbCrLf & _
                      "  ,EndExportadorEquiparado = " & _EnderecoExportadorEquiparado & vbCrLf & _
                      "  ,MemorandoEquiparado  = '" & _CodigoMemorandoEquiparado & "'" & vbCrLf & _
                      "  ,UsuarioAlteracao     = '" & _UsuarioAlteracao & "'" & vbCrLf & _
                      "  ,DataAlteracao        ='" & _DataAlteracao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "  ,DataDaNota           = '" & _DataNota.ToString("yyyy-MM-dd") & "' " & vbCrLf & _
                      "  ,Moeda                = " & _Moeda & " " & vbCrLf & _
                      "  ,Indexador            = " & _Indexador & " " & vbCrLf & _
                      "  ,ValorNota            = " & Str(_ValorNota) & " " & vbCrLf & _
                      "  ,TipoDocumento        = '" & _TipoDocumento & "' " & vbCrLf & _
                      " Where EmpresaMemorando_Id    ='" & _CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and EndEmpresaMemorando_Id = " & _EnderecoEmpresaMemorando & vbCrLf & _
                      "   and Memorando_Id           = '" & _CodigoMemorando & "' " & vbCrLf
                Sqls.Add(sql)
                NotasComprovadas.SalvarSql(Sqls)
                ConhecimentosDeEmbarque.SalvarSql(Sqls)
                RegistrosDeExportacao.SalvarSql(Sqls)
            Case "D"
                NotasComprovadas.SalvarSql(Sqls)
                ConhecimentosDeEmbarque.SalvarSql(Sqls)
                RegistrosDeExportacao.SalvarSql(Sqls)
                sql = " Delete MemorandoDeExportacao" & vbCrLf & _
                      "  Where EmpresaMemorando_Id    ='" & _CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "    and EndEmpresaMemorando_Id = " & _EnderecoEmpresaMemorando & vbCrLf & _
                      "    and Memorando_Id           = '" & _CodigoMemorando & "' " & vbCrLf

                Sqls.Add(sql)
        End Select
    End Sub

#End Region

End Class
