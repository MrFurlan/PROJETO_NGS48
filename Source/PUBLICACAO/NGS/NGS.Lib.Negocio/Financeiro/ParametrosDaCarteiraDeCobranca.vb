<Serializable()>
Public Class ListParametrosDaCarteiraDeCobranca
    Inherits List(Of ParametrosDaCarteiraDeCobranca)

    Public Sub New()

    End Sub

    Public Sub New(ByVal Where As String)
        Dim sql As String

        sql = "SELECT  Empresa_Id, EndEmpresa_Id, Banco_Id, Agencia_Id, DigitoDaAgencia_Id, Conta_Id, DigitoConta_Id, SubConta_Id," & vbCrLf &
                "		AgenciaCobradora, CodigoDaEmpresaNoBanco, CobrancaRegistrada, Instrucao01, Instrucao02," & vbCrLf &
                "       MensagemDeInstrucao, isnull(DiasParaProtesto, 0) as DiasParaProtesto, JurosMes, MoraDiaria, Convenio, Aceite, EmiteBoletoNaEmpresa," & vbCrLf &
                "		SequenciaDeRemessa, SequenciaDeRetorno, CarteiraSimples, CarteiraDesconto, CarteiraCaucao," & vbCrLf &
                "       InformarNumeroNoBoleto, GravarTituloNoArquivoDeRemessa, BaixarPor, ContaContabil, DiasBaixaDecursoPrazo," & vbCrLf &
                "		DiasParaLiberacaoFinanceira, NomeDaEmpresaNoBanco, SequenciaDePagamentos, SequenciaDeRegistroDePagamento," & vbCrLf &
                "       CodigoDaEmpresaPagamento, TipoArquivoCNAB, CarteiraPadrao, NomeNoArquivo, isnull(FolhaDePagamento,0) as FolhaDePagamento, isnull(BeneficiarioFinal,'') as BeneficiarioFinal " & vbCrLf &
                "FROM ParametrosDaCarteiraDeCobranca "

        If Not String.IsNullOrEmpty(Where) Then
            sql = sql + Where
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Parametros")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim p As New ParametrosDaCarteiraDeCobranca
            p.CodigoEmpresa = row("Empresa_Id").ToString()
            p.EnderecoEmpresa = row("EndEmpresa_Id").ToString()
            If Not IsDBNull(row("Banco_Id")) Then p.CodigoBanco = row("Banco_Id")
            If Not IsDBNull(row("Agencia_Id")) Then p.CodigoAgencia = row("Agencia_Id")
            p.DigitoAgencia = row("DigitoDaAgencia_Id").ToString()
            If Not IsDBNull(row("Conta_Id")) Then p.Conta = row("Conta_Id")
            p.DigitoConta = row("DigitoConta_Id").ToString()
            p.SubConta = row("SubConta_Id").ToString()
            If Not IsDBNull(row("AgenciaCobradora")) Then p.AgenciaCobradora = row("AgenciaCobradora")
            p.CodigoDaEmpresaNoBanco = row("CodigoDaEmpresaNoBanco").ToString()
            p.CobrancaRegistrada = row("CobrancaRegistrada").ToString()
            p.Instrucao01 = row("Instrucao01").ToString()
            p.Instrucao02 = row("Instrucao02").ToString()
            p.MensagemDeInstrucao = row("MensagemDeInstrucao").ToString()
            If Not IsDBNull(row("DiasParaProtesto")) Then p.DiasParaProtesto = row("DiasParaProtesto")
            If Not IsDBNull(row("JurosMes")) Then p.JurosMes = row("JurosMes")
            If Not IsDBNull(row("MoraDiaria")) Then p.MoraDiaria = row("MoraDiaria")
            If Not IsDBNull(row("Convenio")) Then p.Convenio = row("Convenio")
            p.Aceite = row("Aceite").ToString()
            p.EmiteBoletoNaEmpresa = row("EmiteBoletoNaEmpresa").ToString()
            If Not IsDBNull(row("SequenciaDeRemessa")) Then p.SequenciaDeRemessa = row("SequenciaDeRemessa")
            If Not IsDBNull(row("SequenciaDeRetorno")) Then p.SequenciaDeRetorno = row("SequenciaDeRetorno")
            If Not IsDBNull(row("CarteiraSimples")) Then p.CarteiraSimples = row("CarteiraSimples")
            If Not IsDBNull(row("CarteiraDesconto")) Then p.CarteiraDesconto = row("CarteiraDesconto")
            If Not IsDBNull(row("CarteiraCaucao")) Then p.CarteiraCaucao = row("CarteiraCaucao")
            p.InformarNumeroNoBoleto = row("InformarNumeroNoBoleto").ToString()
            p.GravarTituloNoArquivoDeRemessa = row("GravarTituloNoArquivoDeRemessa").ToString()
            p.BaixarPor = row("BaixarPor").ToString()
            p.ContaContabil = row("ContaContabil").ToString()
            If Not IsDBNull(row("DiasBaixaDecursoPrazo")) Then p.DiasBaixaDecursoPrazo = row("DiasBaixaDecursoPrazo")
            If Not IsDBNull(row("DiasParaLiberacaoFinanceira")) Then p.DiasParaLiberacaoFinanceira = row("DiasParaLiberacaoFinanceira")
            p.NomeDaEmpresaNoBanco = row("NomeDaEmpresaNoBanco").ToString()
            If Not IsDBNull(row("SequenciaDePagamentos")) Then p.SequenciaDePagamentos = row("SequenciaDePagamentos")
            If Not IsDBNull(row("SequenciaDeRegistroDePagamento")) Then p.SequenciaDeRegistroDePagamento = row("SequenciaDeRegistroDePagamento")
            p.CodigoDaEmpresaPagamento = row("CodigoDaEmpresaPagamento").ToString()
            p.TipoArquivoCNAB = row("TipoArquivoCNAB")
            If Not IsDBNull(row("CarteiraPadrao")) Then p.CarteiraPadrao = row("CarteiraPadrao")
            p.NomeNoArquivo = row("NomeNoArquivo").ToString()
            If Not IsDBNull(row("FolhaDePagamento")) Then p.FolhaDePagamento = row("FolhaDePagamento")
            p.BeneficiarioFinal = row("BeneficiarioFinal")

            Me.Add(p)
        Next
    End Sub
End Class


<Serializable()>
Public Class ParametrosDaCarteiraDeCobranca
    Implements IBaseEntity

#Region "Contrutores"
    Public Sub New()

    End Sub

    Public Sub New(ByVal where As String)
        Dim sql As String

        sql = "SELECT  Empresa_Id, EndEmpresa_Id, Banco_Id, Agencia_Id, DigitoDaAgencia_Id, Conta_Id, DigitoConta_Id, SubConta_Id," & vbCrLf &
                "		AgenciaCobradora, CodigoDaEmpresaNoBanco, CobrancaRegistrada, Instrucao01, Instrucao02," & vbCrLf &
                "       MensagemDeInstrucao, isnull(DiasParaProtesto, 0) as DiasParaProtesto, JurosMes, MoraDiaria, Convenio, Aceite, EmiteBoletoNaEmpresa," & vbCrLf &
                "		SequenciaDeRemessa, SequenciaDeRetorno, CarteiraSimples, CarteiraDesconto, CarteiraCaucao," & vbCrLf &
                "       InformarNumeroNoBoleto, GravarTituloNoArquivoDeRemessa, BaixarPor, ContaContabil, DiasBaixaDecursoPrazo," & vbCrLf &
                "		DiasParaLiberacaoFinanceira, NomeDaEmpresaNoBanco, SequenciaDePagamentos, SequenciaDeRegistroDePagamento, " & vbCrLf &
                "       CodigoDaEmpresaPagamento, TipoArquivoCNAB, CarteiraPadrao, NomeNoArquivo, isnull(FolhaDePagamento,0) as FolhaDePagamento, isnull(BeneficiarioFinal,'') as BeneficiarioFinal" & vbCrLf &
                "FROM ParametrosDaCarteiraDeCobranca "

        If Not String.IsNullOrEmpty(where) Then
            sql = sql + where
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Parametros")

        For Each row As DataRow In ds.Tables(0).Rows

            _CodigoEmpresa = row("Empresa_Id").ToString()
            _EndEmpresa = row("EndEmpresa_Id").ToString()
            If Not IsDBNull(row("Banco_Id")) Then _CodigoBanco = row("Banco_Id")
            If Not IsDBNull(row("Agencia_Id")) Then _CodigoAgencia = row("Agencia_Id")
            _DigitoAgencia = row("DigitoDaAgencia_Id").ToString()
            If Not IsDBNull(row("Conta_Id")) Then _Conta = row("Conta_Id")
            _DigitoConta = row("DigitoConta_Id").ToString()
            _SubConta = row("SubConta_Id").ToString()
            If Not IsDBNull(row("AgenciaCobradora")) Then _AgenciaCobradora = row("AgenciaCobradora")
            _CodigoDaEmpresaNoBanco = row("CodigoDaEmpresaNoBanco").ToString()
            _CobrancaRegistrada = row("CobrancaRegistrada").ToString()
            _Instrucao01 = row("Instrucao01").ToString()
            _Instrucao02 = row("Instrucao02").ToString()
            _MensagemDeInstrucao = row("MensagemDeInstrucao").ToString()
            If Not IsDBNull(row("DiasParaProtesto")) Then _DiasParaProtesto = row("DiasParaProtesto")
            If Not IsDBNull(row("JurosMes")) Then _JurosMes = row("JurosMes")
            If Not IsDBNull(row("MoraDiaria")) Then _MoraDiaria = row("MoraDiaria")
            If Not IsDBNull(row("Convenio")) Then _Convenio = row("Convenio")
            _Aceite = row("Aceite").ToString()
            _EmiteBoletoNaEmpresa = row("EmiteBoletoNaEmpresa").ToString()
            If Not IsDBNull(row("SequenciaDeRemessa")) Then _SequenciaDeRemessa = row("SequenciaDeRemessa")
            If Not IsDBNull(row("SequenciaDeRetorno")) Then _SequenciaDeRetorno = row("SequenciaDeRetorno")
            If Not IsDBNull(row("CarteiraSimples")) Then _CarteiraSimples = row("CarteiraSimples")
            If Not IsDBNull(row("CarteiraDesconto")) Then _CarteiraDesconto = row("CarteiraDesconto")
            If Not IsDBNull(row("CarteiraCaucao")) Then _CarteiraCaucao = row("CarteiraCaucao")
            _InformarNumeroNoBoleto = row("InformarNumeroNoBoleto").ToString()
            _GravarTituloNoArquivoDeRemessa = row("GravarTituloNoArquivoDeRemessa").ToString()
            _BaixarPor = row("BaixarPor").ToString()
            _ContaContabil = row("ContaContabil").ToString()
            If Not IsDBNull(row("DiasBaixaDecursoPrazo")) Then _DiasBaixaDecursoPrazo = row("DiasBaixaDecursoPrazo")
            If Not IsDBNull(row("DiasParaLiberacaoFinanceira")) Then _DiasParaLiberacaoFinanceira = row("DiasParaLiberacaoFinanceira")
            _NomeDaEmpresaNoBanco = row("NomeDaEmpresaNoBanco").ToString()
            If Not IsDBNull(row("SequenciaDePagamentos")) Then _SequenciaDePagamentos = row("SequenciaDePagamentos")
            If Not IsDBNull(row("SequenciaDeRegistroDePagamento")) Then _SequenciaDeRegistroDePagamento = row("SequenciaDeRegistroDePagamento")
            _CodigoDaEmpresaPagamento = row("CodigoDaEmpresaPagamento").ToString()
            _TipoArquivoCNAB = row("TipoArquivoCNAB")
            If Not IsDBNull(row("CarteiraPadrao")) Then _CarteiraPadrao = row("CarteiraPadrao")
            _NomeNoArquivo = row("NomeNoArquivo").ToString()
            If Not IsDBNull(row("FolhaDePagamento")) Then _FolhaDePagamento = row("FolhaDePagamento")
            _BeneficiarioFinal = row("BeneficiarioFinal")

        Next
    End Sub
#End Region

#Region "Fields"
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer 'EndEmpresa

    Private _CodigoBanco As Integer
    Private _CodigoAgencia As Integer
    Private _DigitoAgencia As String
    Private _Conta As Decimal
    Private _DigitoConta As String
    Private _SubConta As String
    Private _AgenciaCobradora As Integer
    Private _CodigoDaEmpresaNoBanco As String
    Private _CobrancaRegistrada As String
    Private _Instrucao01 As String
    Private _Instrucao02 As String
    Private _MensagemDeInstrucao As String
    Private _DiasParaProtesto As Integer
    Private _JurosMes As Decimal
    Private _MoraDiaria As Decimal
    Private _Convenio As Integer
    Private _Aceite As String
    Private _EmiteBoletoNaEmpresa As String
    Private _SequenciaDeRemessa As Integer
    Private _SequenciaDeRetorno As Integer
    Private _CarteiraSimples As Integer
    Private _CarteiraDesconto As Integer
    Private _CarteiraCaucao As Integer
    Private _InformarNumeroNoBoleto As String
    Private _GravarTituloNoArquivoDeRemessa As String
    Private _BaixarPor As String
    Private _ContaContabil As String
    Private _DiasBaixaDecursoPrazo As Integer
    Private _DiasParaLiberacaoFinanceira As Integer
    Private _NomeDaEmpresaNoBanco As String
    Private _SequenciaDePagamentos As Integer
    Private _SequenciaDeRegistroDePagamento As Integer
    Private _CodigoDaEmpresaPagamento As String
    Private _TipoArquivoCNAB As Integer
    Private _CarteiraPadrao As String
    Private _NomeNoArquivo As String
    Private _FolhaDePagamento As Boolean
    Private _BeneficiarioFinal As String

#End Region

#Region "Property"
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

    Public Property CodigoBanco() As Integer
        Get
            Return _CodigoBanco
        End Get
        Set(ByVal value As Integer)
            _CodigoBanco = value
        End Set
    End Property

    Public Property CodigoAgencia() As Integer
        Get
            Return _CodigoAgencia
        End Get
        Set(ByVal value As Integer)
            _CodigoAgencia = value
        End Set
    End Property

    Public Property DigitoAgencia() As String
        Get
            Return _DigitoAgencia
        End Get
        Set(ByVal value As String)
            _DigitoAgencia = value
        End Set
    End Property

    Public Property Conta() As Decimal
        Get
            Return _Conta
        End Get
        Set(ByVal value As Decimal)
            _Conta = value
        End Set
    End Property

    Public Property DigitoConta() As String
        Get
            Return _DigitoConta
        End Get
        Set(ByVal value As String)
            _DigitoConta = value
        End Set
    End Property

    Public Property SubConta() As String
        Get
            Return _SubConta
        End Get
        Set(ByVal value As String)
            _SubConta = value
        End Set
    End Property

    Public Property AgenciaCobradora() As Integer
        Get
            Return _AgenciaCobradora
        End Get
        Set(ByVal value As Integer)
            _AgenciaCobradora = value
        End Set
    End Property

    Public Property CodigoDaEmpresaNoBanco() As String
        Get
            Return _CodigoDaEmpresaNoBanco
        End Get
        Set(ByVal value As String)
            _CodigoDaEmpresaNoBanco = value
        End Set
    End Property

    Public Property CobrancaRegistrada() As String
        Get
            Return _CobrancaRegistrada
        End Get
        Set(ByVal value As String)
            _CobrancaRegistrada = value
        End Set
    End Property

    Public Property Instrucao01() As String
        Get
            Return _Instrucao01
        End Get
        Set(ByVal value As String)
            _Instrucao01 = value
        End Set
    End Property

    Public Property Instrucao02() As String
        Get
            Return _Instrucao02
        End Get
        Set(ByVal value As String)
            _Instrucao02 = value
        End Set
    End Property

    Public Property MensagemDeInstrucao() As String
        Get
            Return _MensagemDeInstrucao
        End Get
        Set(ByVal value As String)
            _MensagemDeInstrucao = value
        End Set
    End Property

    Public Property DiasParaProtesto() As Integer
        Get
            Return _DiasParaProtesto
        End Get
        Set(ByVal value As Integer)
            _DiasParaProtesto = value
        End Set
    End Property

    Public Property JurosMes() As Decimal
        Get
            Return _JurosMes
        End Get
        Set(ByVal value As Decimal)
            _JurosMes = value
        End Set
    End Property

    Public Property MoraDiaria() As Decimal
        Get
            Return _MoraDiaria
        End Get
        Set(ByVal value As Decimal)
            _MoraDiaria = value
        End Set
    End Property

    Public Property Convenio() As Integer
        Get
            Return _Convenio
        End Get
        Set(ByVal value As Integer)
            _Convenio = value
        End Set
    End Property

    Public Property Aceite() As String
        Get
            Return _Aceite
        End Get
        Set(ByVal value As String)
            _Aceite = value
        End Set
    End Property

    Public Property EmiteBoletoNaEmpresa() As String
        Get
            Return _EmiteBoletoNaEmpresa
        End Get
        Set(ByVal value As String)
            _EmiteBoletoNaEmpresa = value
        End Set
    End Property

    Public Property SequenciaDeRemessa() As Integer
        Get
            Return _SequenciaDeRemessa
        End Get
        Set(ByVal value As Integer)
            _SequenciaDeRemessa = value
        End Set
    End Property

    Public Property SequenciaDeRetorno() As Integer
        Get
            Return _SequenciaDeRetorno
        End Get
        Set(ByVal value As Integer)
            _SequenciaDeRetorno = value
        End Set
    End Property

    Public Property CarteiraSimples() As Integer
        Get
            Return _CarteiraSimples
        End Get
        Set(ByVal value As Integer)
            _CarteiraSimples = value
        End Set
    End Property

    Public Property CarteiraDesconto() As Integer
        Get
            Return _CarteiraDesconto
        End Get
        Set(ByVal value As Integer)
            _CarteiraDesconto = value
        End Set
    End Property

    Public Property CarteiraCaucao() As Integer
        Get
            Return _CarteiraCaucao
        End Get
        Set(ByVal value As Integer)
            _CarteiraCaucao = value
        End Set
    End Property

    Public Property InformarNumeroNoBoleto() As String
        Get
            Return _InformarNumeroNoBoleto
        End Get
        Set(ByVal value As String)
            _InformarNumeroNoBoleto = value
        End Set
    End Property

    Public Property GravarTituloNoArquivoDeRemessa() As String
        Get
            Return _GravarTituloNoArquivoDeRemessa
        End Get
        Set(ByVal value As String)
            _GravarTituloNoArquivoDeRemessa = value
        End Set
    End Property

    Public Property BaixarPor() As String
        Get
            Return _BaixarPor
        End Get
        Set(ByVal value As String)
            _BaixarPor = value
        End Set
    End Property

    Public Property ContaContabil() As String
        Get
            Return _ContaContabil
        End Get
        Set(ByVal value As String)
            _ContaContabil = value
        End Set
    End Property

    Public Property DiasBaixaDecursoPrazo() As Integer
        Get
            Return _DiasBaixaDecursoPrazo
        End Get
        Set(ByVal value As Integer)
            _DiasBaixaDecursoPrazo = value
        End Set
    End Property

    Public Property DiasParaLiberacaoFinanceira() As Integer
        Get
            Return _DiasParaLiberacaoFinanceira
        End Get
        Set(ByVal value As Integer)
            _DiasParaLiberacaoFinanceira = value
        End Set
    End Property

    Public Property NomeDaEmpresaNoBanco() As String
        Get
            Return _NomeDaEmpresaNoBanco
        End Get
        Set(ByVal value As String)
            _NomeDaEmpresaNoBanco = value
        End Set
    End Property

    Public Property SequenciaDePagamentos() As Integer
        Get
            Return _SequenciaDePagamentos
        End Get
        Set(ByVal value As Integer)
            _SequenciaDePagamentos = value
        End Set
    End Property

    Public Property SequenciaDeRegistroDePagamento() As Integer
        Get
            Return _SequenciaDeRegistroDePagamento
        End Get
        Set(ByVal value As Integer)
            _SequenciaDeRegistroDePagamento = value
        End Set
    End Property

    Public Property CodigoDaEmpresaPagamento() As String
        Get
            Return _CodigoDaEmpresaPagamento
        End Get
        Set(ByVal value As String)
            _CodigoDaEmpresaPagamento = value
        End Set
    End Property

    Public Property TipoArquivoCNAB() As Integer
        Get
            Return _TipoArquivoCNAB
        End Get
        Set(ByVal value As Integer)
            _TipoArquivoCNAB = value
        End Set
    End Property

    Public Property CarteiraPadrao() As String
        Get
            Return _CarteiraPadrao
        End Get
        Set(ByVal value As String)
            _CarteiraPadrao = value
        End Set
    End Property

    Public Property NomeNoArquivo() As String
        Get
            Return _NomeNoArquivo
        End Get
        Set(ByVal value As String)
            _NomeNoArquivo = value
        End Set
    End Property

    Public Property FolhaDePagamento() As Boolean
        Get
            Return _FolhaDePagamento
        End Get
        Set(ByVal value As Boolean)
            _FolhaDePagamento = value
        End Set
    End Property

    Public Property BeneficiarioFinal() As String
        Get
            Return _BeneficiarioFinal
        End Get
        Set(ByVal value As String)
            _BeneficiarioFinal = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function IncrementarSequenciaDeRemessa(ByVal where As String) As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        Dim Sql = "  UPDATE ParametrosDaCarteiraDeCobranca SET " & vbCrLf &
                  "        SequenciaDeRemessa = SequenciaDeRemessa + 1 " & vbCrLf &
                  " " & where

        sqls.Add(Sql)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function IncrementarSequenciaDeRemessaDaycoval(ByVal where As String, ByVal nomeArq As String, ByVal seqPag As Integer) As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        Dim Sql = "  UPDATE ParametrosDaCarteiraDeCobranca SET " & vbCrLf &
                  "        SequenciaDePagamentos = " + seqPag.ToString() + " " & vbCrLf &
                  "        ,NomeNoArquivo =  '" & nomeArq & "'" & vbCrLf &
                  " " & where

        sqls.Add(Sql)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function AtualizarParametrosDaCarteira(ByVal where As String, ByVal param As ParametrosDaCarteiraDeCobranca) As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        Dim Sql = "  UPDATE ParametrosDaCarteiraDeCobranca SET " & vbCrLf &
                  "        SequenciaDeRemessa = SequenciaDeRemessa + 1 " & vbCrLf &
                  "        ,SequenciaDePagamentos = " + param.SequenciaDePagamentos.ToString() + " " & vbCrLf &
                  "        ,NomeNoArquivo =  '" & param.NomeNoArquivo & "'" & vbCrLf &
                  " " & where

        sqls.Add(Sql)
        Return Banco.GravaBanco(sqls)
    End Function
#End Region

End Class
