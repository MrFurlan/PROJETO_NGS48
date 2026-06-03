Imports System.Web
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListFaturasDeFretesXItens
    Inherits List(Of FaturaDeFreteXItens)
    Implements IBaseEntity

    Private Parent As FaturaDeFrete

    Public Sub New()

    End Sub

    Public Sub New(ByRef fatura As FaturaDeFrete)
        Parent = fatura
        Dim sql As String = String.Empty
        Dim ds As New DataSet
        Dim Banco As New AcessaBanco

        sql &= " SELECT FFxI.EmpresaPagadora_Id, FFxI.EndEmpresaPagadora_Id, FFxI.Conveniado_Id, FFxI.EndConveniado_Id, FFxI.Fatura_Id," & vbCrLf & _
               "        FFxI.Empresa_Id, FFxI.EndEmpresa_Id,      FFxI.Cliente_Id,            FFxI.EndCliente_Id, FFxI.EntradaSaida_Id,  FFxI.Serie_Id," & vbCrLf & _
               "        FFxI.Nota_Id,    FFxI.Encargo_Id,         FFxI.PesoDeChegada,         NfxI.QuantidadeFiscal as Peso, NfxI.Valor, NFxI.Unitario," & vbCrLf & _
               "        SB_NxE.FreteCombinado," & vbCrLf & _
               "        SB_NxE.ValorAdiantamento," & vbCrLf & _
               "        SB_NxE.PercAdiantamento," & vbCrLf & _
               "        case " & vbCrLf & _
               "           when FFxI.Encargo_Id = 'BAIXAADTO'" & vbCrLf & _
               "             then SB_NxE.ValorLancadoBaixaAdto" & vbCrLf & _
               "             else SB_NxE.ValorLancadoLiquidoAPagar" & vbCrLf & _
               "        end ValorLancado  " & vbCrLf & _
               "  FROM FaturasDeFretesXItens AS FFxI" & vbCrLf & _
               " INNER JOIN NotasFiscaisXItens as NfxI" & vbCrLf & _
               "    ON NfxI.Empresa_Id      = FFxI.Empresa_Id" & vbCrLf & _
               "   AND NfxI.EndEmpresa_Id   = FFxI.EndEmpresa_Id" & vbCrLf & _
               "   AND NfxI.Cliente_Id      = FFxI.Cliente_Id" & vbCrLf & _
               "   AND NfxI.EndCliente_Id   = FFxI.EndCliente_Id" & vbCrLf & _
               "   AND NfxI.EntradaSaida_Id = FFxI.EntradaSaida_Id" & vbCrLf & _
               "   AND NfxI.Serie_Id        = FFxI.Serie_Id" & vbCrLf & _
               "   AND NfxI.Nota_Id         = FFxI.Nota_Id" & vbCrLf & _
               " INNER JOIN (SELECT NxE.Empresa_Id, NxE.EndEmpresa_Id, NxE.Cliente_Id, NxE.EndCliente_Id, NxE.EntradaSaida_Id, NxE.Serie_Id, NxE.Nota_Id, " & vbCrLf & _
               "                     NxE.Produto_Id," & vbCrLf & _
               "                     SUM(case" & vbCrLf & _
               "						  when NxE.Encargo_id = 'ADIANTAMENTO'" & vbCrLf & _
               "							then NxE.Valor" & vbCrLf & _
               "							else 0" & vbCrLf & _
               "						end) as ValorAdiantamento," & vbCrLf & _
               "					SUM(case" & vbCrLf & _
               "						  when NxE.Encargo_id = 'ADIANTAMENTO'" & vbCrLf & _
               "							then NxE.PERCENTUAL" & vbCrLf & _
               "							else 0" & vbCrLf & _
               "						end) AS PercAdiantamento, " & vbCrLf & _
               "					SUM(case" & vbCrLf & _
               "						  when NxE.Encargo_id = 'BAIXAADTO'" & vbCrLf & _
               "							then NxE.Valor" & vbCrLf & _
               "							else 0" & vbCrLf & _
               "						end) AS ValorLancadoBaixaAdto," & vbCrLf & _
               "                    SUM(case" & vbCrLf & _
               "						  when NxE.Encargo_id = 'LIQUIDO'" & vbCrLf & _
               "							then NxE.Valor" & vbCrLf & _
               "							else 0" & vbCrLf & _
               "						end) AS ValorLancadoLiquidoAPagar," & vbCrLf & _
               "					SUM(case" & vbCrLf & _
               "						  when NxE.Encargo_id = 'PRODUTO'" & vbCrLf & _
               "							then NxE.Valor" & vbCrLf & _
               "							else 0" & vbCrLf & _
               "						end) AS FreteCombinado" & vbCrLf & _
               "                from NotasFiscaisXEncargos as NxE" & vbCrLf & _
               "               group by NxE.Empresa_Id, NxE.EndEmpresa_Id, NxE.Cliente_Id, NxE.EndCliente_Id, NxE.EntradaSaida_Id, NxE.Serie_Id, NxE.Nota_Id, NxE.Produto_Id" & vbCrLf & _
               "             ) SB_NxE" & vbCrLf & _
               "    ON NfxI.Empresa_Id      = SB_NxE.Empresa_Id" & vbCrLf & _
               "   AND NfxI.EndEmpresa_Id   = SB_NxE.EndEmpresa_Id" & vbCrLf & _
               "   AND NfxI.Cliente_Id      = SB_NxE.Cliente_Id" & vbCrLf & _
               "   AND NfxI.EndCliente_Id   = SB_NxE.EndCliente_Id" & vbCrLf & _
               "   AND NfxI.EntradaSaida_Id = SB_NxE.EntradaSaida_Id" & vbCrLf & _
               "   AND NfxI.Serie_Id        = SB_NxE.Serie_Id" & vbCrLf & _
               "   AND NfxI.Nota_Id         = SB_NxE.Nota_Id" & vbCrLf & _
               "   AND NfxI.Produto_Id      = SB_NxE.Produto_Id" & vbCrLf & _
               " Where FFxI.Fatura_Id             ='" & fatura.CodigoFatura & "' " & vbCrLf & _
               "   AND FFxI.EmpresaPagadora_Id    ='" & fatura.CodigoEmpresa & "' " & vbCrLf & _
               "   AND FFxI.EndEmpresaPagadora_Id = " & fatura.EnderecoEmpresa & " " & vbCrLf & _
               "   AND FFxI.Conveniado_Id         ='" & fatura.CodigoConveniado & "' " & vbCrLf & _
               "   AND FFxI.EndConveniado_Id      = " & fatura.EnderecoConveniado & " " & vbCrLf


        ds = Banco.ConsultaDataSet(sql, "FaturasDeFretesXItens")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim FFXI As New FaturaDeFreteXItens(fatura)
            FFXI.ClienteCnpj = row("cliente_id")
            FFXI.ClienteEnd = row("endcliente_id")
            FFXI.CTRC = IIf(row("serie_id") = "UN", True, False)
            FFXI.ReciboFrete = IIf(row("serie_id") = "REC", True, False)
            FFXI.EmpresaCnpj = row("empresa_id")
            FFXI.EmpresaEnd = row("endempresa_id")
            FFXI.CodigoEncargo = row("encargo_id")
            FFXI.EntradaSaida = row("entradasaida_id")
            FFXI.NumeroNota = row("Nota_Id")
            FFXI.Serie = row("serie_id")
            FFXI.Peso = row("Peso")
            FFXI.ViaAdiantamento = IIf(row("encargo_id") = "BAIXAADTO", True, False)
            FFXI.ViaAmortizacao = IIf(row("encargo_id") = "AMORTIZAADTO", True, False)
            FFXI.ViaCartaFrete = IIf(row("encargo_id") = "LIQUIDOAPAGAR", True, False)
            FFXI.FreteCombinado = row("Unitario")
            FFXI.ValorFrete = row("Valor")
            FFXI.PesoDeChegada = row("PesoDeChegada")
            FFXI.PercentAdiantamento = row("PercAdiantamento")
            FFXI.ValorAdiantamento = row("ValorAdiantamento")
            FFXI.ValorLancadoNota = row("ValorLancado")
            Me.Add(FFXI)
        Next
    End Sub

    Public Sub New(ByVal nf As NotaFiscal)
        Dim sql As String = ""
        Dim db As New AcessaBanco
        Dim ds As New DataSet

        sql &= "SELECT EmpresaPagadora_Id, EndEmpresaPagadora_Id, Conveniado_Id, EndConveniado_Id, Fatura_Id, Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf & _
               "       EntradaSaida_Id, Serie_Id, Nota_Id, Encargo_Id, PesoDeChegada " & vbCrLf & _
               "  FROM FaturasDeFretesXItens " & vbCrLf & _
               " WHERE 1=1 " & vbCrLf & _
               "   AND Empresa_Id = '" & nf.CodigoEmpresa & "' " & vbCrLf & _
               "   AND EndEmpresa_Id = '" & nf.EnderecoEmpresa & "' " & vbCrLf & _
               "   AND Cliente_Id = '" & nf.CodigoCliente & "' " & vbCrLf & _
               "   AND EndCliente_Id = '" & nf.EnderecoCliente & "' " & vbCrLf & _
               "   AND EntradaSaida_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf & _
               "   AND Nota_Id = '" & nf.Codigo & "' " & vbCrLf & _
               "   AND Serie_Id = '" & nf.Serie & "'"

        ds = db.ConsultaDataSet(sql, "FaturasDeFretesXItens")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ff As New FaturaDeFrete()
            ff.CodigoEmpresa = row("EmpresaPagadora_Id")
            ff.EnderecoEmpresa = row("EndEmpresaPagadora_Id")
            ff.CodigoConveniado = row("Conveniado_Id")
            ff.EnderecoConveniado = row("EndConveniado_Id")
            ff.CodigoFatura = row("Fatura_Id")

            Dim item As New FaturaDeFreteXItens(ff)
            item.PagadoraCnpj = row("EmpresaPagadora_Id")
            item.PagadoraEnd = row("EndEmpresaPagadora_Id")
            item.ConveniadoCnpj = row("Conveniado_Id")
            item.ConveniadoEnd = row("EndConveniado_Id")
            item.ClienteCnpj = row("Cliente_Id")
            item.ClienteEnd = row("EndCliente_Id")
            item.CTRC = IIf(row("Serie_Id") = "UN", True, False)
            item.ReciboFrete = IIf(row("Serie_Id") = "REC", True, False)
            item.EmpresaCnpj = row("Empresa_Id")
            item.EmpresaEnd = row("EndEmpresa_Id")
            item.CodigoEncargo = row("Encargo_Id")
            item.EntradaSaida = row("EntradaSaida_Id")
            item.NumeroNota = row("Nota_Id")
            item.Serie = row("Serie_Id")
            item.ViaAdiantamento = IIf(row("Encargo_Id") = "BAIXAADTO", True, False)
            item.ViaAmortizacao = IIf(row("Encargo_Id") = "AMORTIZAADTO", True, False)
            item.ViaCartaFrete = IIf(row("Encargo_Id") = "LIQUIDOAPAGAR", True, False)
            item.PesoDeChegada = row("PesoDeChegada")
            item.CodigoFatura = row("Fatura_Id")
            Me.Add(item)
        Next
    End Sub

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        Me.SalvarSql(Sqls)

        If Sqls.Count = 0 OrElse Banco.GravaBanco(Sqls) Then
            For Each x As FaturaDeFreteXItens In Me
                x.IUD = ""
            Next
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As FaturaDeFreteXItens In Me
            If Not item.IUD = Nothing Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub

End Class

<Serializable()> _
Public Class FaturaDeFreteXItens
    Implements IBaseEntity

#Region "Construtor"

    Private _Parent As FaturaDeFrete

    Public Sub New(Optional ByRef PFrete As FaturaDeFrete = Nothing)
        If Not PFrete Is Nothing Then
            _Parent = PFrete
        Else
            _Parent = New FaturaDeFrete
        End If
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoFatura As String = ""
    Private _PagadoraCnpj As String = ""
    Private _PagadoraEnd As Integer
    Private _ConveniadoCnpj As String = ""
    Private _ConveniadoEnd As Integer
    Private _EmpresaCnpj As String = ""
    Private _EmpresaEnd As Integer
    Private _EmpresaReduzido As String
    Private _EmpresaNome As String
    Private _PesoDeChegada As Decimal
    Private _ClienteCnpj As String = ""
    Private _ClienteEnd As Integer
    Private _ClienteReduzido As String
    Private _ClienteNome As String
    Private _EntradaSaida As String
    Private _Serie As String
    Private _NumeroNota As Integer
    Private _CodigoEncargo As String
    Private _ReciboFrete As Boolean
    Private _CTRC As Boolean
    Private _ViaAdiantamento As Boolean = True
    Private _ViaAmortizacao As Boolean
    Private _ViaCartaFrete As Boolean
    Private _Peso As Decimal
    Private _FreteCombinado As Double
    Private _ValorFrete As Double
    Private _FreteChegada As Double
    Private _PercentAdiantamento As String
    Private _ValorAdiantamento As Double
    Private _ValorLancadoNota As Double
    Private _Empresa As Cliente = Nothing
    Private _Cliente As Cliente = Nothing
    Private _Nota As NotaFiscal
#End Region

#Region "Propriedades"
    Public Property CodigoFatura() As String
        Get
            Return _CodigoFatura
        End Get
        Set(ByVal value As String)
            _CodigoFatura = value
        End Set
    End Property

    Public Property Nota() As NotaFiscal
        Get
            If _Nota Is Nothing Then
                Dim nf As New NotaFiscal
                nf.CodigoEmpresa = _EmpresaCnpj
                nf.EnderecoEmpresa = _EmpresaEnd
                nf.CodigoCliente = _ClienteCnpj
                nf.EnderecoCliente = _ClienteEnd
                nf.EntradaSaida = IIf(_EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                nf.Codigo = _NumeroNota
                nf.Serie = _Serie
                _Nota = New NotaFiscal(nf)
            End If
            Return _Nota
        End Get
        Set(ByVal value As NotaFiscal)
            _Nota = value
        End Set
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing AndAlso _EmpresaCnpj.Length > 0 Then _Empresa = New Cliente(Funcoes.EliminarCaracteresEspeciais(_EmpresaCnpj), _EmpresaEnd)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing AndAlso _ClienteCnpj.Length > 0 Then _Cliente = New Cliente(Funcoes.EliminarCaracteresEspeciais(_ClienteCnpj), _ClienteEnd)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property EmpresaCnpj() As String
        Get
            Return _EmpresaCnpj
        End Get
        Set(ByVal value As String)
            _EmpresaCnpj = value
        End Set
    End Property

    Public Property EmpresaEnd() As Integer
        Get
            Return _EmpresaEnd
        End Get
        Set(ByVal value As Integer)
            _EmpresaEnd = value
        End Set
    End Property

    Public Property PagadoraCnpj() As String
        Get
            Return _PagadoraCnpj
        End Get
        Set(ByVal value As String)
            _PagadoraCnpj = value
        End Set
    End Property

    Public Property PagadoraEnd() As Integer
        Get
            Return _PagadoraEnd
        End Get
        Set(ByVal value As Integer)
            _PagadoraEnd = value
        End Set
    End Property

    Public Property ConveniadoCnpj() As String
        Get
            Return _ConveniadoCnpj
        End Get
        Set(ByVal value As String)
            _ConveniadoCnpj = value
        End Set
    End Property

    Public Property ConveniadoEnd() As Integer
        Get
            Return _ConveniadoEnd
        End Get
        Set(ByVal value As Integer)
            _ConveniadoEnd = value
        End Set
    End Property

    Public ReadOnly Property EmpresaReduzido() As String
        Get
            If Empresa Is Nothing Then
                Return ""
            Else
                Return Empresa.Reduzido
            End If
        End Get
    End Property

    Public ReadOnly Property EmpresaNome() As String
        Get
            If Empresa Is Nothing Then
                Return ""
            Else
                Return Empresa.Nome
            End If
        End Get
    End Property

    Public Property ClienteCnpj() As String
        Get
            Return _ClienteCnpj
        End Get
        Set(ByVal value As String)
            _ClienteCnpj = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property ClienteEnd() As Integer
        Get
            Return _ClienteEnd
        End Get
        Set(ByVal value As Integer)
            _ClienteEnd = value
        End Set
    End Property

    Public ReadOnly Property ClienteReduzido() As String
        Get
            If Cliente Is Nothing Then
                Return ""
            Else
                Return Cliente.Reduzido
            End If
        End Get
    End Property

    Public ReadOnly Property ClienteNome() As String
        Get
            If Cliente Is Nothing Then
                Return ""
            Else
                Return Cliente.Nome
            End If
        End Get
    End Property

    Public Property NumeroNota() As Integer
        Get
            Return _NumeroNota
        End Get
        Set(ByVal value As Integer)
            _NumeroNota = value
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

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
        End Set
    End Property

    Public Property CodigoEncargo() As String
        Get
            Return _CodigoEncargo
        End Get
        Set(ByVal value As String)
            _CodigoEncargo = value
        End Set
    End Property

    Public Property ReciboFrete() As Boolean
        Get
            Return _ReciboFrete
        End Get
        Set(ByVal value As Boolean)
            _ReciboFrete = value
            _CTRC = Not value
        End Set
    End Property

    Public Property CTRC() As Boolean
        Get
            Return _CTRC
        End Get
        Set(ByVal value As Boolean)
            _CTRC = value
            _ReciboFrete = Not value
        End Set
    End Property

    Public Property ViaAdiantamento() As Boolean
        Get
            Return _ViaAdiantamento
        End Get
        Set(ByVal value As Boolean)
            _ViaAdiantamento = value
            _ViaAmortizacao = Not value
            _ViaCartaFrete = Not value
        End Set
    End Property

    Public Property ViaAmortizacao() As Boolean
        Get
            Return _ViaAmortizacao
        End Get
        Set(ByVal value As Boolean)
            _ViaAmortizacao = value
            _ViaAdiantamento = Not value
            _ViaCartaFrete = Not value
        End Set
    End Property

    Public Property ViaCartaFrete() As Boolean
        Get
            Return _ViaCartaFrete
        End Get
        Set(ByVal value As Boolean)
            _ViaCartaFrete = value
            _ViaAdiantamento = Not value
            _ViaAmortizacao = Not value
        End Set
    End Property

    Public Property Peso() As Decimal
        Get
            Return _Peso
        End Get
        Set(ByVal value As Decimal)
            _Peso = value
        End Set
    End Property

    Public Property PesoDeChegada() As Decimal
        Get
            Return _PesoDeChegada
        End Get
        Set(ByVal value As Decimal)
            _PesoDeChegada = value
            _FreteChegada = Math.Round(IIf(_PesoDeChegada > _Peso, _Peso * _FreteCombinado / 1000, _PesoDeChegada * _FreteCombinado / 1000), 2, MidpointRounding.AwayFromZero)
        End Set
    End Property

    Public Property FreteCombinado() As Double
        Get
            Return _FreteCombinado
        End Get
        Set(ByVal value As Double)
            _FreteCombinado = value
        End Set
    End Property

    Public Property ValorFrete() As Double
        Get
            Return _ValorFrete
        End Get
        Set(ByVal value As Double)
            _ValorFrete = value
        End Set
    End Property

    Public Property FreteChegada() As Double
        Get
            Return _FreteChegada
        End Get
        Set(ByVal value As Double)
            _FreteChegada = value
        End Set
    End Property

    Public ReadOnly Property Diferenca() As Double
        Get
            Return IIf(_FreteChegada = 0, 0, _FreteChegada - _ValorFrete)
        End Get
    End Property

    Public Property PercentAdiantamento() As String
        Get
            Return _PercentAdiantamento
        End Get
        Set(ByVal value As String)
            _PercentAdiantamento = value
        End Set
    End Property

    Public Property ValorAdiantamento() As Double
        Get
            Return _ValorAdiantamento
        End Get
        Set(ByVal value As Double)
            _ValorAdiantamento = value
        End Set
    End Property

    Public Property ValorLancadoNota() As Double
        Get
            Return _ValorLancadoNota
        End Get
        Set(ByVal value As Double)
            _ValorLancadoNota = value
        End Set
    End Property

    Public ReadOnly Property Parent() As FaturaDeFrete
        Get
            Return _Parent
        End Get

    End Property
#End Region

#Region "Métodos"

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        Me.SalvarSql(Sqls)

        If Sqls.Count = 0 Then Return True

        If Banco.GravaBanco(Sqls) Then
            Me.IUD = ""
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String
        Select Case Me.IUD

            Case "I"

                strSQL = "INSERT INTO FaturasDeFretesXItens " & vbCrLf &
                         "        (EmpresaPagadora_Id, EndEmpresaPagadora_Id, Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, " & vbCrLf &
                         "         Nota_Id, Encargo_Id, PesoDeChegada, Conveniado_Id, EndConveniado_Id, Fatura_Id)" & vbCrLf &
                         "VALUES ('" & Parent.CodigoEmpresa & "'," & Parent.EnderecoEmpresa & ",'" & Me.EmpresaCnpj & "'," & Me.EmpresaEnd & ",'" & Me.ClienteCnpj & "'," & vbCrLf &
                         "         " & Me.ClienteEnd & ",'" & Me.EntradaSaida & "','" & Me.Serie & "'," & Me.NumeroNota & ",'" & IIf(ViaAdiantamento, "BAIXAADTO", IIf(ViaAmortizacao, "AMORTIZAADTO", "LIQUIDOAPAGAR")) & "', " & vbCrLf &
                         "         " & Str(Me.PesoDeChegada) & ", '" & Parent.CodigoConveniado & "', " & Parent.EnderecoConveniado & ", " & Parent.CodigoFatura & "); "
                Sqls.Add(strSQL)

            Case "U"
                strSQL = "UPDATE FaturasDeFretesXItens  " & vbCrLf &
                         "   SET PesoDeChegada          = " & Me.PesoDeChegada & vbCrLf &
                         " WHERE EmpresaPagadora_Id    ='" & Parent.CodigoEmpresa & "'" & vbCrLf &
                         "   AND EndEmpresaPagadora_Id ='" & Parent.EnderecoEmpresa & "'" & vbCrLf &
                         "   AND Empresa_Id            ='" & Me.EmpresaCnpj & "'" & vbCrLf &
                         "   AND EndEmpresa_Id         = " & Me.EmpresaEnd & vbCrLf &
                         "   AND Cliente_Id            ='" & Me.ClienteCnpj & "'" & vbCrLf &
                         "   AND EndCliente_Id         ='" & Me.ClienteEnd & "'" & vbCrLf &
                         "   AND EntradaSaida_Id       ='" & Me.EntradaSaida & "'" & vbCrLf &
                         "   AND Serie_Id              ='" & Me.Serie & "'" & vbCrLf &
                         "   AND Nota_Id               ='" & Me.NumeroNota & "'" & vbCrLf &
                         "   AND Encargo_Id            ='" & IIf(Me.ViaAdiantamento, "BAIXAADTO", IIf(ViaAmortizacao, "AMORTIZAADTO", "LIQUIDOAPAGAR")) & "'" & vbCrLf &
                         "   AND Fatura_Id             ='" & Parent.CodigoFatura & "'" & vbCrLf &
                         "   AND Conveniado_Id         ='" & Parent.CodigoConveniado & "'" & vbCrLf &
                         "   AND EndConveniado_Id      = " & Parent.EnderecoConveniado & ";"
                Sqls.Add(strSQL)

            Case "D"

                strSQL = " DELETE FaturasDeFretesXItens" & vbCrLf &
                         "	WHERE EmpresaPagadora_Id    ='" & Parent.CodigoEmpresa & "'" & vbCrLf &
                         "	  AND EndEmpresaPagadora_Id ='" & Parent.EnderecoEmpresa & "'" & vbCrLf &
                         "	  AND Empresa_Id            ='" & Me.EmpresaCnpj & "'" & vbCrLf &
                         "	  AND EndEmpresa_Id         ='" & Me.EmpresaEnd & "'" & vbCrLf &
                         "	  AND Cliente_Id            ='" & Me.ClienteCnpj & "'" & vbCrLf &
                         "	  AND EndCliente_Id         ='" & Me.ClienteEnd & "'" & vbCrLf &
                         "	  AND EntradaSaida_Id       ='" & Me.EntradaSaida & "'" & vbCrLf &
                         "	  AND Serie_Id              ='" & Me.Serie & "'" & vbCrLf &
                         "	  AND Nota_Id               ='" & Me.NumeroNota & "'" & vbCrLf &
                         "	  AND Encargo_Id            ='" & IIf(Me.ViaAdiantamento, "BAIXAADTO", IIf(Me.ViaAmortizacao, "AMORTIZAADTO", "LIQUIDOAPAGAR")) & "'" & vbCrLf &
                         "	  AND Fatura_Id             ='" & Parent.CodigoFatura & "'" & vbCrLf &
                         "	  AND Conveniado_Id         ='" & Parent.CodigoConveniado & "'" & vbCrLf &
                         "	  AND EndConveniado_Id      = " & Parent.EnderecoConveniado & ";"
                Sqls.Add(strSQL)

        End Select
    End Sub

    Public Function JaFaturada_old() As Boolean
        Dim sql As String = ""
        sql = "     SELECT FF.Fatura_Id, C.Nome AS DescConveniado" & vbCrLf &
              "     FROM FaturasDeFretesXItens FxI " & vbCrLf &
              "     INNER JOIN Clientes C" & vbCrLf &
              "         ON FxI.Conveniado_Id    = C.Cliente_Id " & vbCrLf &
              "         AND FxI.EndConveniado_Id = C.Endereco_Id " & vbCrLf &
              "     INNER JOIN FaturasDeFretes FF " & vbCrLf &
              "         ON FxI.Empresa_Id = FF.Empresa_Id " & vbCrLf &
              "         AND FxI.EndEmpresa_Id = FF.EndEmpresa_Id " & vbCrLf &
              "         AND FxI.Fatura_Id = FF.Fatura_Id " & vbCrLf &
              "         AND NOT ISNULL(FF.RegistroMestre,0) = 0 " & vbCrLf &
              "	    WHERE FxI.Empresa_Id       ='" & Me.EmpresaCnpj & "'" & vbCrLf &
              "	        AND FxI.EndEmpresa_Id    ='" & Me.EmpresaEnd & "'" & vbCrLf &
              "	        AND FxI.Cliente_Id       ='" & Me.ClienteCnpj & "'" & vbCrLf &
              "	        AND FxI.EndCliente_Id    ='" & Me.ClienteEnd & "'" & vbCrLf &
              "	        AND FxI.EntradaSaida_Id  ='" & Me.EntradaSaida & "'" & vbCrLf &
              "	        AND FxI.Serie_Id         ='" & Me.Serie & "'" & vbCrLf &
              "	        AND FxI.Nota_Id          ='" & Me.NumeroNota & "'" & vbCrLf &
              "	        AND FxI.Encargo_Id  in(" & IIf(Me.ViaAdiantamento Or ViaAmortizacao, "'BAIXAADTO','AMORTIZAADTO'", "'LIQUIDOAPAGAR'") & ")"

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "JaFaturada")
        If ds.Tables(0).Rows.Count > 0 Then
            HttpContext.Current.Session("ssMessage") = "Esse CTRC já foi vinculado a fatura " & ds.Tables(0).Rows(0)("Fatura_Id").ToString & " para " & ds.Tables(0).Rows(0)("DescConveniado").ToString & "!"
            Return True
        End If
        Return False
    End Function

    Public Function JaFaturada() As Boolean

        ' Monta a lista de encargos conforme sua regra
        Dim encargoIn As String = If(Me.ViaAdiantamento Or ViaAmortizacao, "'BAIXAADTO','AMORTIZAADTO'", "'LIQUIDOAPAGAR'")

        Dim sql As String =
        "SELECT TOP 1 " & vbCrLf &
        "       FF.Fatura_Id, " & vbCrLf &
        "       FF.RegistroMestre, " & vbCrLf &
        "       C.Nome AS DescConveniado " & vbCrLf &
        "FROM FaturasDeFretesXItens FxI " & vbCrLf &
        "INNER JOIN Clientes C " & vbCrLf &
        "   ON FxI.Conveniado_Id = C.Cliente_Id " & vbCrLf &
        "  AND FxI.EndConveniado_Id = C.Endereco_Id " & vbCrLf &
        "INNER JOIN FaturasDeFretes FF " & vbCrLf &
        "   ON FxI.Empresa_Id    = FF.Empresa_Id " & vbCrLf &
        "  AND FxI.EndEmpresa_Id = FF.EndEmpresa_Id " & vbCrLf &
        "  AND FxI.Fatura_Id     = FF.Fatura_Id " & vbCrLf &
        "WHERE FxI.Empresa_Id      = '" & Me.EmpresaCnpj & "' " & vbCrLf &
        "  AND FxI.EndEmpresa_Id   = '" & Me.EmpresaEnd & "' " & vbCrLf &
        "  AND FxI.Cliente_Id      = '" & Me.ClienteCnpj & "' " & vbCrLf &
        "  AND FxI.EndCliente_Id   = '" & Me.ClienteEnd & "' " & vbCrLf &
        "  AND FxI.EntradaSaida_Id = '" & Me.EntradaSaida & "' " & vbCrLf &
        "  AND FxI.Serie_Id        = '" & Me.Serie & "' " & vbCrLf &
        "  AND FxI.Nota_Id         = '" & Me.NumeroNota & "' " & vbCrLf &
        "  AND FxI.Encargo_Id IN (" & encargoIn & ") "

        ' Observação: não filtramos mais por RegistroMestre aqui; vamos decidir no VB.

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "JaFaturada")

        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Function JaFaturadaRegistroMestre() As Boolean

        ' Monta a lista de encargos conforme sua regra
        Dim encargoIn As String = If(Me.ViaAdiantamento Or ViaAmortizacao, "'BAIXAADTO','AMORTIZAADTO'", "'LIQUIDOAPAGAR'")

        Dim sql As String =
        "SELECT TOP 1 " & vbCrLf &
        "       FF.Fatura_Id, " & vbCrLf &
        "       FF.RegistroMestre, " & vbCrLf &
        "       C.Nome AS DescConveniado " & vbCrLf &
        "FROM FaturasDeFretesXItens FxI " & vbCrLf &
        "INNER JOIN Clientes C " & vbCrLf &
        "   ON FxI.Conveniado_Id = C.Cliente_Id " & vbCrLf &
        "  AND FxI.EndConveniado_Id = C.Endereco_Id " & vbCrLf &
        "INNER JOIN FaturasDeFretes FF " & vbCrLf &
        "   ON FxI.Empresa_Id    = FF.Empresa_Id " & vbCrLf &
        "  AND FxI.EndEmpresa_Id = FF.EndEmpresa_Id " & vbCrLf &
        "  AND FxI.Fatura_Id     = FF.Fatura_Id " & vbCrLf &
        "WHERE FxI.Empresa_Id      = '" & Me.EmpresaCnpj & "' " & vbCrLf &
        "  AND FxI.EndEmpresa_Id   = '" & Me.EmpresaEnd & "' " & vbCrLf &
        "  AND FxI.Cliente_Id      = '" & Me.ClienteCnpj & "' " & vbCrLf &
        "  AND FxI.EndCliente_Id   = '" & Me.ClienteEnd & "' " & vbCrLf &
        "  AND FxI.EntradaSaida_Id = '" & Me.EntradaSaida & "' " & vbCrLf &
        "  AND FxI.Serie_Id        = '" & Me.Serie & "' " & vbCrLf &
        "  AND FxI.Nota_Id         = '" & Me.NumeroNota & "' " & vbCrLf &
        "  AND FxI.Encargo_Id IN (" & encargoIn & ") " & vbCrLf &
        "  AND ISNULL(RegistroMestre, 0) > 0 "

        ' Observação: não filtramos mais por RegistroMestre aqui; vamos decidir no VB.

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "JaFaturada")

        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
            Return False
        End If

        Dim faturaId As Integer = 0
        Dim descConveniado As String = ""
        Dim regMestre As Integer = 0

        For Each row As DataRow In ds.Tables(0).Rows

            faturaId = Convert.ToInt32(row("Fatura_Id"))
            descConveniado = row("DescConveniado").ToString()
            If Not IsDBNull(row("RegistroMestre")) Then
                regMestre = Convert.ToInt32(row("RegistroMestre"))
                Exit For
            End If

        Next

        If Me.Parent.LancamentoManual = 1 Then
            If regMestre <> Me.Parent.CodigoFatura Then
                HttpContext.Current.Session("ssMessage") =
            "Este CTRC já está vinculado à fatura mestre " & regMestre & ". Processo bloqueado."
                Return True
            Else
                'Se for a mesma fatura, não tem problema
                Return False
            End If
        Else

            HttpContext.Current.Session("ssMessage") =
        "Este CTRC está na fatura filha " & faturaId &
        " (vinculada à fatura mestre " & regMestre & ") para " & descConveniado & ". Exclusão bloqueada."

            Return True

        End If

        Return False

    End Function

#End Region

End Class