Imports System.Text

Public Class FolhaDePagamento

#Region "Fields"
    Private _Titulos As List(Of Titulo)

    Public Erro As Exception
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Property"
    Public Property Titulos() As List(Of Titulo)
        Get
            Return _Titulos
        End Get
        Set(ByVal value As List(Of Titulo))
            _Titulos = value
        End Set
    End Property

#End Region

#Region "Methods"

#Region "Bradesco"
    Public Function BradescoCNAB240(ByVal titulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of StringBuilder)
        Dim remessa As New RemessaBancaria()
        Dim sb As New List(Of StringBuilder)
        Dim i As Integer = 0

        Dim tituloHeader = titulos.First()

        'Header Arquivo
        remessa.Header.Append(remessa.Format(1, 3, True, 237)) 'Bradesco - 237
        remessa.Header.Append(remessa.Format(4, 7, True, 0)) 'Lote de Serviço - 0000
        remessa.Header.Append(remessa.Format(8, 8, True, 0)) 'Tipo de Registro - 0
        remessa.Header.Append(remessa.Format(9, 17, False, "")) 'Uso Exclusivo FEBRABAN / CNAB - Brancos
        remessa.Header.Append(remessa.Format(18, 18, True, IIf(param.CodigoEmpresa.Length = 14, 2, 1))) 'Tipo de Inscrição da Empresa (2 - CNPJ)
        remessa.Header.Append(remessa.Format(19, 32, True, param.CodigoEmpresa)) 'Número de Inscrição da Empresa 
        remessa.Header.Append(remessa.Format(33, 52, False, param.Convenio)) 'Código do Convênio no Banco
        remessa.Header.Append(remessa.Format(53, 57, True, param.CodigoAgencia)) 'Agência Mantenedora da Conta
        remessa.Header.Append(remessa.Format(58, 58, False, param.DigitoAgencia)) 'Dígito Verificador da Agência
        remessa.Header.Append(remessa.Format(59, 70, True, param.Conta)) 'Número da Conta Corrente
        remessa.Header.Append(remessa.Format(71, 71, False, param.DigitoConta)) 'Dígito Verificador da Conta
        remessa.Header.Append(remessa.Format(72, 72, False, "")) 'Dígito Verificador da Ag/Conta
        remessa.Header.Append(remessa.Format(73, 102, False, param.NomeDaEmpresaNoBanco)) 'Nome da Empresa
        remessa.Header.Append(remessa.Format(103, 132, False, "BANCO BRADESCO S.A")) 'Nome do Banco
        remessa.Header.Append(remessa.Format(133, 142, False, "")) 'Uso Exclusivo FEBRABAN / CNAB (Brancos)
        remessa.Header.Append(remessa.Format(143, 143, True, 1)) 'Código Remessa / Retorno '1' = Remessa (Cliente  Banco)
        remessa.Header.Append(remessa.Format(144, 151, True, DateTime.Now.ToString("ddMMyyyy"))) 'Data de Geração do Arquivo - DDMMAAAA
        remessa.Header.Append(remessa.Format(152, 157, True, DateTime.Now.ToString("hhmmss"))) 'Hora de Geração do Arquivo - HHMMSS
        remessa.Header.Append(remessa.Format(158, 163, True, param.SequenciaDePagamentos)) 'Número Seqüencial do Arquivo
        remessa.Header.Append(remessa.Format(164, 166, True, 89)) 'Numero da Versão do Layout do Arquivo - 089
        remessa.Header.Append(remessa.Format(167, 171, True, 1600)) 'Densidade de Gravação do Arquivo - 1600
        remessa.Header.Append(remessa.Format(172, 191, False, "")) 'Para Uso Reservado do Banco
        remessa.Header.Append(remessa.Format(192, 211, False, "")) 'Para Uso Reservado da Empresa
        remessa.Header.Append(remessa.Format(212, 240, False, "")) 'Uso Exclusivo FEBRABAN / CNAB - Brancos

        sb.Add(remessa.Header)

        'Header Lote
        remessa.Header = New StringBuilder()
        Dim contadorLote As Integer = 1

        remessa.Header.Append(remessa.Format(1, 3, True, 237)) 'Código do Banco na Compensação - 237
        remessa.Header.Append(remessa.Format(4, 7, True, contadorLote)) 'Lote de Serviço 
        remessa.Header.Append(remessa.Format(8, 8, True, 1)) 'Tipo de Registro - 1
        remessa.Header.Append(remessa.Format(9, 9, False, "C")) 'Tipo da Operação  - 'C'
        remessa.Header.Append(remessa.Format(10, 11, True, 30)) 'Tipo do Serviço - '30' = Pagamento Salários
        remessa.Header.Append(remessa.Format(12, 13, True, 1)) 'Forma de Lançamento - '01' = Crédito em Conta Corrente
        remessa.Header.Append(remessa.Format(14, 16, True, 45)) 'Nº da Versão do Layout do Lote - 040
        remessa.Header.Append(remessa.Format(17, 17, False, "")) 'Uso Exclusivo da FEBRABAN/CNAB - Brancos
        remessa.Header.Append(remessa.Format(18, 18, True, IIf(param.CodigoEmpresa.Length = 14, 2, 1))) 'Tipo de Inscrição da Empresa (2 - CNPJ)
        remessa.Header.Append(remessa.Format(19, 32, True, param.CodigoEmpresa)) 'Número de Inscrição da Empresa 
        remessa.Header.Append(remessa.Format(33, 52, False, param.Convenio)) 'Código do Convênio no Banco
        remessa.Header.Append(remessa.Format(53, 57, True, param.CodigoAgencia)) 'Agência Mantenedora da Conta
        remessa.Header.Append(remessa.Format(58, 58, False, param.DigitoAgencia)) 'Dígito Verificador da Agência
        remessa.Header.Append(remessa.Format(59, 70, True, param.Conta)) 'Número da Conta Corrente
        remessa.Header.Append(remessa.Format(71, 71, False, param.DigitoConta)) 'Dígito Verificador da Conta
        remessa.Header.Append(remessa.Format(72, 72, False, "")) 'Dígito Verificador da Ag/Conta
        remessa.Header.Append(remessa.Format(73, 102, False, param.NomeDaEmpresaNoBanco)) 'Nome da Empresa
        remessa.Header.Append(remessa.Format(103, 142, False, "")) 'Mensagem
        remessa.Header.Append(remessa.Format(143, 172, False, tituloHeader.Empresa.Endereco)) 'Nome da Rua, Av, Pça, Etc
        remessa.Header.Append(remessa.Format(173, 177, True, tituloHeader.Empresa.Numero)) 'Número do Local
        remessa.Header.Append(remessa.Format(178, 192, False, tituloHeader.Empresa.Complemento)) 'Casa, Apto, Sala, Etc
        remessa.Header.Append(remessa.Format(193, 212, False, tituloHeader.Empresa.Cidade)) 'Nome da Cidade
        remessa.Header.Append(remessa.Format(213, 217, True, tituloHeader.Empresa.CEP)) 'CEP
        remessa.Header.Append(remessa.Format(218, 220, False, "")) 'Complemento do CEP
        remessa.Header.Append(remessa.Format(221, 222, False, tituloHeader.Empresa.Estado.Codigo)) 'Sigla do Estado
        remessa.Header.Append(remessa.Format(223, 224, False, "01"))  'Indicativo de Forma de Pagamento Do Serviço - 01
        remessa.Header.Append(remessa.Format(225, 230, False, "")) 'Uso Exclusivo FEBRABAN/CNAB
        remessa.Header.Append(remessa.Format(231, 240, False, "")) 'Códigos das Ocorrências p/ Retorno

        sb.Add(remessa.Header)

        'Dim _detalheSegC As New StringBuilder

        Dim contadorRegistro As Integer = 0

        For Each tit In titulos

            contadorRegistro = contadorRegistro + 1
            'Segmento A
            Dim _detalheSegA As New StringBuilder

            _detalheSegA.Append(remessa.Format(1, 3, True, 237)) 'Código do Banco na Compensação - 237
            _detalheSegA.Append(remessa.Format(4, 7, True, contadorLote)) 'Lote de Serviço
            _detalheSegA.Append(remessa.Format(8, 8, True, 3)) 'Tipo de Registro - 3
            _detalheSegA.Append(remessa.Format(9, 13, True, contadorRegistro)) 'Nº Seqüencial do Registro no Lote
            _detalheSegA.Append(remessa.Format(14, 14, False, "A")) 'Código de Segmento do Reg. Detalhe - A
            _detalheSegA.Append(remessa.Format(15, 15, True, 0)) 'Tipo de Movimento - '0' = Indica INCLUSÃO
            _detalheSegA.Append(remessa.Format(16, 17, True, 0)) 'Código da Instrução p/ Movimento - 00' = Inclusão de Registro Detalhe Liberado
            _detalheSegA.Append(remessa.Format(18, 20, True, 0)) 'Código da Câmara Centralizadora - Para folha de pagamento 000
            _detalheSegA.Append(remessa.Format(21, 23, True, tit.BancoCliente.Codigo)) 'Código do Banco do Favorecido
            _detalheSegA.Append(remessa.Format(24, 28, True, tit.AgenciaCliente.Agencia)) 'Ag. Mantenedora da Cta do Favor.
            _detalheSegA.Append(remessa.Format(29, 29, False, tit.AgenciaCliente.DigitoAgencia)) 'Dígito Verificador da Agência
            _detalheSegA.Append(remessa.Format(30, 41, True, tit.ContaCliente)) 'Número da Conta Corrente
            _detalheSegA.Append(remessa.Format(42, 42, False, tit.DigitoContaCliente.Trim())) 'Dígito Verificador da Conta
            _detalheSegA.Append(remessa.Format(43, 43, False, "")) 'Dígito Verificador da AG/Conta
            _detalheSegA.Append(remessa.Format(44, 73, False, tit.Cliente.Nome)) 'Nome do Favorecido
            _detalheSegA.Append(remessa.Format(74, 93, False, tit.Codigo)) 'Nº do Docum. Atribuído p/ Empresa
            _detalheSegA.Append(remessa.Format(94, 101, True, tit.Prorrogacao.ToString("ddMMyyyy"))) 'Data do Pagamento - DDMMAAAA
            _detalheSegA.Append(remessa.Format(102, 104, False, "BRL")) 'Tipo da Moeda 'BRL' = Real
            _detalheSegA.Append(remessa.Format(105, 119, True, 1)) 'Quantidade da Moeda
            _detalheSegA.Append(remessa.Format(120, 134, True, tit.ValorLiquido.ToString().Replace(",", ""))) 'Valor do Pagamento
            _detalheSegA.Append(remessa.Format(135, 154, False, "")) 'Nº do Docum. Atribuído pelo Banco
            _detalheSegA.Append(remessa.Format(155, 162, True, tit.Prorrogacao.ToString("ddMMyyyy"))) 'Data Real da Efetivação Pagto
            _detalheSegA.Append(remessa.Format(163, 177, True, tit.ValorLiquido.ToString().Replace(",", ""))) 'Valor Real da Efetivação do Pagto
            _detalheSegA.Append(remessa.Format(178, 217, False, "")) 'Informação 2
            _detalheSegA.Append(remessa.Format(218, 219, False, "")) 'Compl. Tipo Serviço - '01' = Crédito em Conta
            _detalheSegA.Append(remessa.Format(220, 224, False, "")) 'Codigo finalidade da TED - 4 Pagamento de Salários 
            _detalheSegA.Append(remessa.Format(225, 226, False, "")) 'Compl. Finalidade de Pagamento
            _detalheSegA.Append(remessa.Format(227, 229, False, "")) 'Uso Exclusivo FEBRABAN/CNAB - Brancos
            _detalheSegA.Append(remessa.Format(230, 230, True, 0)) 'Aviso ao Favorecido '0' = Não Emite Aviso
            _detalheSegA.Append(remessa.Format(231, 240, False, "")) 'Códigos das Ocorrências p/ Retorno

            remessa.AddDetalhe(_detalheSegA)
            sb.Add(_detalheSegA)


            'Segmento B
            Dim _detalheSegB As New StringBuilder
            contadorRegistro = contadorRegistro + 1

            _detalheSegB.Append(remessa.Format(1, 3, True, 237)) 'Código do Banco na Compensação - 237 Bradesco
            _detalheSegB.Append(remessa.Format(4, 7, True, contadorLote)) 'Lote de Serviço
            _detalheSegB.Append(remessa.Format(8, 8, True, 3)) 'Tipo do Registro - 3
            _detalheSegB.Append(remessa.Format(9, 13, True, contadorRegistro)) 'Nº Seqüencial do Registro no Lote
            _detalheSegB.Append(remessa.Format(14, 14, False, "B")) 'Código de Segmento do Reg. Detalhe - B
            _detalheSegB.Append(remessa.Format(15, 17, False, "")) 'Uso Exclusivo FEBRABAN/CNAB - Brancos
            _detalheSegB.Append(remessa.Format(18, 18, True, IIf(tit.CodigoCliente.Length = 14, 2, 1))) 'Tipo de Inscrição do Favorecido - '1' = CPF '2' = CGC / CNPJ
            _detalheSegB.Append(remessa.Format(19, 32, True, tit.CodigoCliente)) 'Nº de Inscrição do Favorecido
            _detalheSegB.Append(remessa.Format(33, 62, False, tit.Cliente.Endereco.Replace(",", ""))) 'Nome da Rua, Av, Pça, Etc
            _detalheSegB.Append(remessa.Format(63, 67, True, tit.Cliente.Numero)) 'Nº do Local
            _detalheSegB.Append(remessa.Format(68, 82, False, tit.Cliente.Complemento)) 'Casa, Apto, Etc
            _detalheSegB.Append(remessa.Format(83, 97, False, tit.Cliente.Bairro)) 'Bairro
            _detalheSegB.Append(remessa.Format(98, 117, False, tit.Cliente.Cidade)) 'Nome da Cidade
            _detalheSegB.Append(remessa.Format(118, 122, True, tit.Cliente.CEP)) 'CEP
            _detalheSegB.Append(remessa.Format(123, 125, False, "")) 'Complemento do CEP
            _detalheSegB.Append(remessa.Format(126, 127, False, tit.Cliente.Estado.Codigo)) 'Sigla do Estado
            _detalheSegB.Append(remessa.Format(128, 135, True, tit.Prorrogacao.ToString("ddMMyyyy"))) 'Data do Vencimento (Nominal) - DDMMAAAA
            _detalheSegB.Append(remessa.Format(136, 150, True, tit.ValorLiquido.ToString().Replace(",", ""))) 'Valor do Documento (Nominal)
            _detalheSegB.Append(remessa.Format(151, 165, True, 0)) 'Valor do Abatimento
            _detalheSegB.Append(remessa.Format(166, 180, True, 0)) 'Valor do Desconto
            _detalheSegB.Append(remessa.Format(181, 195, True, 0)) 'Valor da Mora
            _detalheSegB.Append(remessa.Format(196, 210, True, 0)) 'Valor da Multa
            _detalheSegB.Append(remessa.Format(211, 225, False, tit.Codigo)) 'Código/Documento do Favorecido
            _detalheSegB.Append(remessa.Format(226, 226, True, 0)) 'Aviso ao Favorecido
            _detalheSegB.Append(remessa.Format(227, 232, True, 0)) 'Uso Exclusivo para o SIAPE
            _detalheSegB.Append(remessa.Format(233, 240, False, "")) 'Uso Exclusivo FEBRABAN/CNAB - Brancos

            remessa.AddDetalhe(_detalheSegB)
            sb.Add(_detalheSegB)
        Next

        'Trailler do Lote        
        Dim _trailerLote As New StringBuilder
        _trailerLote.Append(remessa.Format(1, 3, True, 237)) 'Código do Banco na Compensação - 237 Bradesco
        _trailerLote.Append(remessa.Format(4, 7, True, contadorLote)) 'Lote de Serviço
        _trailerLote.Append(remessa.Format(8, 8, True, 5)) 'Tipo de Registro - 5
        _trailerLote.Append(remessa.Format(9, 17, False, "")) 'Uso Exclusivo FEBRABAN/CNAB - Brancos
        _trailerLote.Append(remessa.Format(18, 23, True, (contadorRegistro + 2))) 'Quantidade de Registros do Lote
        _trailerLote.Append(remessa.Format(24, 41, True, titulos.Sum(Function(p) p.ValorLiquido).ToString().Replace(",", ""))) 'Somatória dos Valores
        _trailerLote.Append(remessa.Format(42, 59, True, 0)) 'Somatória de Quantidade de Moedas
        _trailerLote.Append(remessa.Format(60, 65, True, 0)) 'Número Aviso de Débito
        _trailerLote.Append(remessa.Format(66, 230, False, "")) 'Número Aviso de Débito - Brancos
        _trailerLote.Append(remessa.Format(231, 240, False, "")) 'Códigos das Ocorrências para Retorno
        sb.Add(_trailerLote)



        'Trailler do arquivo
        i = i + 1
        Dim _trailer As New StringBuilder
        remessa.Trailler.Append(remessa.Format(1, 3, True, 237)) 'Código do Banco na Compensação - Bradesco 237
        remessa.Trailler.Append(remessa.Format(4, 7, True, 9999)) 'Lote de Serviço
        remessa.Trailler.Append(remessa.Format(8, 8, True, 9)) 'Tipo de Registro - 9
        remessa.Trailler.Append(remessa.Format(9, 17, False, "")) 'Uso Exclusivo FEBRABAN/CNAB - Brancos
        remessa.Trailler.Append(remessa.Format(18, 23, True, contadorLote)) 'Quantidade de Lotes do Arquivo
        remessa.Trailler.Append(remessa.Format(24, 29, True, (contadorRegistro + 4))) 'Quantidade de Registros do Arquivo
        remessa.Trailler.Append(remessa.Format(30, 35, True, 0)) 'Qtde de Contas p/ Conc. (Lotes)
        remessa.Trailler.Append(remessa.Format(36, 240, False, "")) 'Uso Exclusivo FEBRABAN/CNAB - Brancos
        sb.Add(remessa.Trailler)

        Return sb
    End Function

#End Region

#End Region
End Class
