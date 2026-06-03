Imports NGS.Lib.Uteis
Imports System.Globalization
Imports System.Text

<Serializable()>
Public Class BoletoBancario
#Region "Fields"
    Private _Boleto As BoletoNet.Boleto

    Public Erro As Exception
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Property"
    Public Property Boleto() As BoletoNet.Boleto
        Get
            Return _Boleto
        End Get
        Set(ByVal value As BoletoNet.Boleto)
            _Boleto = value
        End Set
    End Property

#End Region

#Region "Daycoval"
    Public Function DaycovalArquivoCNAB400(ByVal Boletos As List(Of BoletoNet.Boleto)) As List(Of StringBuilder)
        Dim remessa As New RemessaBancaria()
        Dim sb As New List(Of StringBuilder)
        Dim i As Integer = 0

        Dim b = Boletos.First()

        'Header
        i = i + 1
        remessa.Header.Append(remessa.Format(1, 1, True, 0))
        remessa.Header.Append(remessa.Format(2, 2, True, 1))
        remessa.Header.Append(remessa.Format(3, 9, False, "REMESSA"))
        remessa.Header.Append(remessa.Format(10, 11, True, 1))
        remessa.Header.Append(remessa.Format(12, 26, False, "COBRANCA"))
        remessa.Header.Append(remessa.Format(27, 46, False, b.Cedente.Codigo))
        remessa.Header.Append(remessa.Format(47, 76, False, b.Cedente.Nome))
        remessa.Header.Append(remessa.Format(77, 79, True, b.Banco.Codigo))
        remessa.Header.Append(remessa.Format(80, 94, False, b.Banco.Nome))
        remessa.Header.Append(remessa.Format(95, 100, True, DateTime.Now.ToString("ddMMyy")))
        remessa.Header.Append(remessa.Format(101, 394, False, " "))
        remessa.Header.Append(remessa.Format(395, 400, True, i))

        sb.Add(remessa.Header)

        Dim _detalhe As New StringBuilder

        For Each tit As BoletoNet.Boleto In Boletos
            'Detalhe do Titulo
            i = i + 1
            _detalhe = New StringBuilder
            _detalhe.Append(remessa.Format(1, 1, True, 1))
            _detalhe.Append(remessa.Format(2, 3, True, 2))
            _detalhe.Append(remessa.Format(4, 17, True, tit.Cedente.CPFCNPJ))
            _detalhe.Append(remessa.Format(18, 37, False, tit.Cedente.Codigo))
            _detalhe.Append(remessa.Format(38, 62, True, tit.NumeroDocumento))
            _detalhe.Append(remessa.Format(63, 70, False, tit.NossoNumero))
            _detalhe.Append(remessa.Format(71, 83, False, " "))
            _detalhe.Append(remessa.Format(84, 107, False, " "))
            _detalhe.Append(remessa.Format(108, 108, True, 6))
            _detalhe.Append(remessa.Format(109, 110, True, 1))
            _detalhe.Append(remessa.Format(111, 120, False, tit.NumeroDocumento)) 'Tem que verificar
            _detalhe.Append(remessa.Format(121, 126, True, tit.DataVencimento.ToString("ddMMyy")))
            _detalhe.Append(remessa.Format(127, 139, True, tit.ValorBoleto.ToString().Replace(",", "")))
            _detalhe.Append(remessa.Format(140, 142, True, 707))
            _detalhe.Append(remessa.Format(143, 146, True, 0))
            _detalhe.Append(remessa.Format(147, 147, True, 0))
            _detalhe.Append(remessa.Format(148, 149, False, "01"))
            _detalhe.Append(remessa.Format(150, 150, False, "N"))
            _detalhe.Append(remessa.Format(151, 156, True, tit.DataDocumento.ToString("ddMMyy")))
            _detalhe.Append(remessa.Format(157, 158, True, 0))
            _detalhe.Append(remessa.Format(159, 160, True, 0))
            _detalhe.Append(remessa.Format(161, 173, True, 0))
            _detalhe.Append(remessa.Format(174, 179, True, 0))
            _detalhe.Append(remessa.Format(180, 192, True, 0))
            _detalhe.Append(remessa.Format(193, 205, True, 0))
            _detalhe.Append(remessa.Format(206, 218, True, 0))
            _detalhe.Append(remessa.Format(219, 220, True, IIf(tit.Sacado.CPFCNPJ.Length = 14, 2, 1)))
            _detalhe.Append(remessa.Format(221, 234, False, tit.Sacado.CPFCNPJ))
            _detalhe.Append(remessa.Format(235, 264, False, tit.Sacado.Nome))
            _detalhe.Append(remessa.Format(265, 274, False, " "))
            _detalhe.Append(remessa.Format(275, 314, False, tit.Sacado.Endereco.EndComNumeroEComplemento))
            _detalhe.Append(remessa.Format(315, 326, False, tit.Sacado.Endereco.Bairro))
            _detalhe.Append(remessa.Format(327, 334, False, tit.Sacado.Endereco.CEP))
            _detalhe.Append(remessa.Format(335, 349, False, tit.Sacado.Endereco.Cidade))
            _detalhe.Append(remessa.Format(350, 351, False, tit.Sacado.Endereco.UF))
            _detalhe.Append(remessa.Format(352, 381, False, tit.Cedente.Nome)) 'Tem q verificar
            _detalhe.Append(remessa.Format(382, 385, False, " "))
            _detalhe.Append(remessa.Format(386, 391, False, " "))
            _detalhe.Append(remessa.Format(392, 393, True, 0))
            _detalhe.Append(remessa.Format(394, 394, True, 0))
            _detalhe.Append(remessa.Format(395, 400, True, i))

            remessa.AddDetalhe(_detalhe)
            sb.Add(_detalhe)
            'Nota Fiscal
            Dim nf As New NotaFiscalXTitulo(New TituloV With {.Codigo = tit.NumeroDocumento})
            _detalhe = New StringBuilder
            i = i + 1
            _detalhe.Append(remessa.Format(1, 1, True, 4))
            _detalhe.Append(remessa.Format(2, 16, False, nf.NotaFiscal.Codigo)) 'Número da Nota Fiscal 
            _detalhe.Append(remessa.Format(17, 29, True, nf.NotaFiscal.TotalNota.ToString().Replace(",", ""))) 'Valor da Nota Fiscal
            _detalhe.Append(remessa.Format(30, 37, True, nf.NotaFiscal.DataNota.ToString("ddMMyyyy"))) 'Data de Emissão da Nota
            _detalhe.Append(remessa.Format(38, 81, True, nf.NotaFiscal.ChaveNFE)) 'Chave de Acesso DANFE
            _detalhe.Append(remessa.Format(82, 394, False, " ")) '
            _detalhe.Append(remessa.Format(395, 400, True, i))

            remessa.AddDetalhe(_detalhe)
            sb.Add(_detalhe)

        Next


        'Trailler
        i = i + 1
        Dim _trailer As New StringBuilder
        remessa.Trailler.Append(remessa.Format(1, 1, True, 9))
        remessa.Trailler.Append(remessa.Format(2, 394, False, " "))
        remessa.Trailler.Append(remessa.Format(395, 400, True, i))
        sb.Add(remessa.Trailler)

        Return sb
    End Function

    Public Function DaycovalPdf(ByVal Boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        Dim _codigoDeBarras = DaycovalCodigoDeBarras(Boleto)
        Dim DV = _codigoDeBarras.Last()
        Dim imagemCodBarras = "<img src='" + ImagemCodigoDeBarras(_codigoDeBarras) + "' alt='Código de Barras' />"
        Dim linhaDigitavel = DaycovalLinhaDigitavel(Boleto, _codigoDeBarras)

        Dim pdf = Html()
        pdf = pdf.Replace("@CODIGOBANCO", Boleto.Banco.Codigo.ToString())
        pdf = pdf.Replace("@DIGITOBANCO", 2)
        pdf = pdf.Replace("@NOMEBANCO", Boleto.Banco.Nome)
        pdf = pdf.Replace("@LOCALPAGAMENTO", "PAGÁVEL EM QUALQUER AGÊNCIA BANCÁRIA, MESMO APÓS VENCIMENTO")
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "-" + Boleto.Cedente.ContaBancaria.DigitoAgencia + "/" + r.Format(1, 7, True, Boleto.Cedente.ContaBancaria.Conta) + "-" + Boleto.Cedente.ContaBancaria.DigitoConta)
        'pdf = pdf.Replace("@URLIMGCEDENTE", "")
        pdf = pdf.Replace("@URLIMAGEMBARRA", "")
        pdf = pdf.Replace("@LINHADIGITAVEL", linhaDigitavel)
        pdf = pdf.Replace("@LOCALPAGAMENTO", Boleto.LocalPagamento)
        pdf = pdf.Replace("@DATAVENCIMENTO", Boleto.DataVencimento)
        pdf = pdf.Replace("@CEDENTE_BOLETO", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome, String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Cedente.Nome, Convert.ToUInt64(Boleto.Cedente.CPFCNPJ).ToString("00\.000\.000\/0000\-00"))))
        pdf = pdf.Replace("@CEDENTE", Boleto.Cedente.Nome)
        pdf = pdf.Replace("@DATADOCUMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NUMERODOCUMENTO", Boleto.NumeroDocumento)
        pdf = pdf.Replace("@ESPECIEDOCUMENTO", "DM")
        pdf = pdf.Replace("@DATAPROCESSAMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NOSSONUMERO", Boleto.Carteira + "/" + Boleto.NossoNumero + "-" + DV)
        pdf = pdf.Replace("@CARTEIRA", Boleto.Carteira)
        pdf = pdf.Replace("@ESPECIE", Boleto.Especie)
        pdf = pdf.Replace("@QUANTIDADE", IIf(Boleto.QuantidadeMoeda = 0, "", Boleto.QuantidadeMoeda.ToString()))
        pdf = pdf.Replace("@VALORDOCUMENTO", Boleto.ValorMoeda)
        pdf = pdf.Replace("@=VALORDOCUMENTO", Boleto.ValorBoleto.ToString("C", CultureInfo.GetCultureInfo("PT-BR")))
        pdf = pdf.Replace("@VALORCOBRADO", IIf(Boleto.ValorCobrado = 0, "", Boleto.ValorCobrado.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@OUTROSACRESCIMOS", "")
        pdf = pdf.Replace("@OUTRASDEDUCOES", "")
        pdf = pdf.Replace("@DESCONTOS", IIf(Boleto.ValorDesconto = 0, "", Boleto.ValorDesconto.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@AGENCIACONTA", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "-" + Boleto.Cedente.ContaBancaria.DigitoAgencia + "/" + r.Format(1, 7, True, Boleto.Cedente.ContaBancaria.Conta) + "-" + Boleto.Cedente.ContaBancaria.DigitoConta)
        pdf = pdf.Replace("@SACADO", String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Sacado.Nome, Convert.ToUInt64(Boleto.Sacado.CPFCNPJ).ToString("00\.000\.000\/0000\-00")))
        pdf = pdf.Replace("@INFOSACADO", Boleto.Sacado.Endereco.EndComNumeroEComplemento.ToString() + " - " + Boleto.Sacado.Endereco.Cidade + ", " + Boleto.Sacado.Endereco.UF)
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", Boleto.ContaBancaria.Agencia)
        pdf = pdf.Replace("@CPFCNPJ", Boleto.Cedente.CPFCNPJ)
        pdf = pdf.Replace("@MORAMULTA", IIf(Boleto.ValorMulta = 0, "", Boleto.ValorMulta.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@AUTENTICACAOMECANICA", "")
        pdf = pdf.Replace("@USODOBANCO", Boleto.UsoBanco)
        pdf = pdf.Replace("@IMAGEMCODIGOBARRA", imagemCodBarras) 'Imagem do código de barras
        pdf = pdf.Replace("@ACEITE", Boleto.Aceite).ToString()
        pdf = pdf.Replace("@ENDERECOCEDENTE", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Endereco, ""))
        pdf = pdf.Replace("@AVALISTA", "")

        Dim texto As String = String.Empty

        For Each i In Boleto.Instrucoes
            texto += Environment.NewLine + i.Descricao
        Next
        pdf = pdf.Replace("@INSTRUCOES", texto)
        '          String.Format(
        '              "{0} - {1}",
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.Nome, String.Empty),
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.CPFCNPJ, String.Empty)))
        'pdf.Replace("Ar\" > R$", RemoveSimboloMoedaValorDocumento ? "Ar\">" :  "Ar\">R$")
        pdf = pdf.Replace("@PARCELATOTAL", "") 'IIf(Boleto.NumeroParcela <> 0 And Boleto.TotalParcela <> 0, Boleto.NumeroParcela + " / " + Boleto.TotalParcela, String.Empty))
        pdf = pdf.Replace("@DADOSCEDENTE", "") 'IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome + " - " + Boleto.Cedente.CPFCNPJ + "</br>" + Boleto.Cedente.Endereco.EndComNumeroEComplemento, ""))

        Return pdf
    End Function

    Private Function ImagemCodigoDeBarras(ByVal codigoDeBarras) As String
        Dim fnCodigoBarras = IO.Path.GetTempFileName()
        Dim cb = New BoletoNet.C2of5i(codigoDeBarras, 1, 50, codigoDeBarras.Length)
        cb.ToBitmap().Save(fnCodigoBarras)

        Return fnCodigoBarras
    End Function

    Private Function DaycovalCodigoDeBarras(ByVal Boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        'Agência 3 posições
        Dim codigoDeBarras = Boleto.Banco.Codigo.ToString()
        'Moeda 1 posição
        codigoDeBarras += "9"
        'Fator de Vencimento 4 posições
        Dim ts = Convert.ToDateTime(Boleto.DataDocumento).Subtract(Convert.ToDateTime("03/07/2000"))
        codigoDeBarras += ts.Days.ToString()
        'Valor 10 posições
        codigoDeBarras += r.Format(1, 10, True, Boleto.ValorBoleto.ToString().Replace(",", ""))
        'Campo Livre 25
        'Agência sem Digito 4 \
        Dim agencia = r.Format(1, 4, True, Boleto.Cedente.ContaBancaria.Agencia)
        codigoDeBarras += agencia
        'Carteira 3 posições
        codigoDeBarras += r.Format(1, 3, True, Boleto.Carteira)
        'Operação 7 posições
        codigoDeBarras += r.Format(1, 7, True, Boleto.NumeroControle)
        'Nosso Número + DV 11 posições
        Dim nossoNumero = r.Format(1, 10, True, Boleto.NossoNumero.ToString())
        Dim dvNN = DVNossoNumero((agencia.ToString() + Boleto.Carteira.ToString() + nossoNumero.ToString()))
        codigoDeBarras += r.Format(1, 11, True, Boleto.NossoNumero.ToString() + dvNN.ToString())

        If codigoDeBarras.Length() <> 43 Then
            Throw New Exception("Codigo de Barras incorreto: " + codigoDeBarras.Length())
        End If
        Dim dv = Mod11Daycoval(codigoDeBarras)

        codigoDeBarras = codigoDeBarras.Insert(4, dv).ToString()

        Return codigoDeBarras
    End Function

    Private Function DaycovalLinhaDigitavel(ByVal Boleto As BoletoNet.Boleto, ByVal codBarras As String) As String
        Dim r As New RemessaBancaria()
        Dim campoLivre = codBarras.Substring(19, 25)
        'Campo 1
        Dim campo1 = Boleto.Banco.Codigo.ToString() + Boleto.Moeda.ToString() + campoLivre.Substring(0, 5)
        campo1 += Mod10Daycoval(campo1, 2)
        Dim campo2 = campoLivre.Substring(5, 10)
        campo2 += Mod10Daycoval(campo2, 1)
        Dim campo3 = campoLivre.Substring(15, 10)
        campo3 += Mod10Daycoval(campo3, 1)
        Dim campo4 = codBarras.Substring(4, 1)
        Dim campo5 = codBarras.Substring(5, 4) + codBarras.Substring(9, 10)

        Return (campo1 + "." + campo2 + "." + campo3 + " " + campo4 + " " + campo5)
    End Function

    Public Function Mod11Daycoval(ByVal pCodigoDeBarra As String) As String
        Dim soma = 0
        'Dim pesos = Split("4,3,2,9,8,7,6,5,4,3,2,9,8,7,6,5,4,5,3,2,9,8,7,6,5,4,5,3,2,9,8,7,6,5,4,5,3,2,9,8,7,6,5,4,5,3,2", ",")
        Dim peso = 4
        Dim arr = pCodigoDeBarra.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            soma += res
            'Pesos
            If peso = 2 Then
                peso = 9
            Else
                peso = (peso - 1)
            End If
        Next

        Dim dv = (11 - (soma Mod 11))

        If dv = 0 OrElse dv = 1 OrElse dv > 9 Then
            Return 1
        End If

        Return dv
    End Function

    Public Function Mod10Daycoval(ByVal campo As String, ByVal peso As Integer) As String
        Dim soma = 0

        Dim arr = campo.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            For Each numAux In res.ToString().ToCharArray()
                soma += Integer.Parse(numAux)
            Next

            peso = IIf(peso = 2, 1, 2)
        Next

        Dim dv = ((soma * 9) Mod 10)

        Return dv
    End Function

    Public Function DVNossoNumero(ByVal nossoNumero) As String
        Dim soma = 0

        Dim peso = 2
        Dim arr = nossoNumero.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            For Each numAux In res.ToString().ToCharArray()
                soma += Integer.Parse(numAux)
            Next

            peso = IIf(peso = 2, 1, 2)
        Next

        Dim dv = (soma Mod 10)

        If dv > 0 Then
            dv = (10 - dv)
        End If

        Return dv
    End Function

    Public Function Html() As String
        Dim sb As New StringBuilder
        sb.AppendLine("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
        sb.AppendLine("<html xmlns=""http://www.w3.org/1999/xhtml"">")
        sb.AppendLine("<meta http-equiv='Content-Type' content=""text/html; charset=utf-8"" />")
        sb.AppendLine("<meta charset='utf-8' />")
        sb.AppendLine("<head><title>Boleto.Net</title>")
        sb.AppendLine("<style>")
        sb.AppendLine("body { color: #000000; background-color: #ffffff; margin-top: 0; margin-right: 0; }")
        sb.AppendLine("* { margin: 0px; padding: 0px; }")
        sb.AppendLine("table { border: 0; border-collapse: collapse; padding: 0; }")
        sb.AppendLine("img { border: 0; }")
        sb.AppendLine(".cp { font: bold 10px arial; color: black; }")
        sb.AppendLine(".ti { font: 9px arial, helvetica, sans-serif; }")
        sb.AppendLine(".ld { font: bold 15px arial; color: #000000; }")
        sb.AppendLine(".ct { font: 9px 'arial narrow'; color: #000033; }")
        sb.AppendLine(".cn { font: 9px arial; color: black; }")
        sb.AppendLine(".bc { font: bold 20px arial; color: #000000; }")
        sb.AppendLine(".cut { width: 665px; height: 1px; border-top: dashed 1px #000; }")
        sb.AppendLine(".Ac { text-align: center; }")
        sb.AppendLine(".Ar { text-align: right; }")
        sb.AppendLine(".Al { text-align: left; }")
        sb.AppendLine(".At { vertical-align: top; }")
        sb.AppendLine(".Ab { vertical-align: bottom; }")
        sb.AppendLine(".ct td, .cp td { padding-left: 6px; border-left: solid 1px #000; }")
        sb.AppendLine(".cpN { font: bold 10px arial; color: black; }")
        sb.AppendLine(".ctN { font: 9px 'arial narrow'; color: #000033; }")
        sb.AppendLine(".pL0 { padding-left: 0px; }")
        sb.AppendLine(".pL6 { padding-left: 6px; }")
        sb.AppendLine(".pL10 { padding-left: 10px; }")
        sb.AppendLine(".imgLogo { width: 120px; }")
        sb.AppendLine(".imgLogo img { width: 120px; height: 20px; }")
        sb.AppendLine(".barra { width: 3px; height: 22px; vertical-align: bottom; }")
        sb.AppendLine(".barra img { width: 2px; height: 22px; }")
        sb.AppendLine(".rBb td { border-bottom: solid 1px #000; }")
        sb.AppendLine(".BB { border-bottom: solid 1px #000; }")
        sb.AppendLine(".BL { border-left: solid 1px #000; }")
        sb.AppendLine(".BR { border-right: solid 1px #000; }")
        sb.AppendLine(".BT { border-top: solid 1px #000; }")
        sb.AppendLine(".BT1 { border-top: dashed 1px #000; }")
        sb.AppendLine(".BT2 { border-top: solid 2px #000; }")
        sb.AppendLine(".h1 { height: 1px; }")
        sb.AppendLine(".h13 { height: 13px; }")
        sb.AppendLine(".h12 { height: 12px; }")
        sb.AppendLine(".h13 td { vertical-align: top; }")
        sb.AppendLine(".h12 td { vertical-align: top; }")
        sb.AppendLine(".w6 { width: 6px; }")
        sb.AppendLine(".w7 { width: 7px; }")
        sb.AppendLine(".w34 { width: 34px; }")
        sb.AppendLine(".w45 { width: 45px; }")
        sb.AppendLine(".w53 { width: 53px; }")
        sb.AppendLine(".w62 { width: 62px; }")
        sb.AppendLine(".w65 { width: 65px; }")
        sb.AppendLine(".w72 { width: 72px; }")
        sb.AppendLine(".w83 { width: 83px; }")
        sb.AppendLine(".w88 { width: 88px; }")
        sb.AppendLine(".w104 { width: 104px; }")
        sb.AppendLine(".w105 { width: 105px; }")
        sb.AppendLine(".w106 { width: 106px; }")
        sb.AppendLine(".w113 { width: 113px; }")
        sb.AppendLine(".w112 { width: 112px; }")
        sb.AppendLine(".w123 { width: 123px; }")
        sb.AppendLine(".w126 { width: 126px; }")
        sb.AppendLine(".w128 { width: 128px; }")
        sb.AppendLine(".w132 { width: 132px; }")
        sb.AppendLine(".w134 { width: 134px; }")
        sb.AppendLine(".w150 { width: 150px; }")
        sb.AppendLine(".w163 { width: 163px; }")
        sb.AppendLine(".w164 { width: 164px; }")
        sb.AppendLine(".w180 { width: 180px; }")
        sb.AppendLine(".w182 { width: 182px; }")
        sb.AppendLine(".w186 { width: 186px; }")
        sb.AppendLine(".w192 { width: 192px; }")
        sb.AppendLine(".w250 { width: 250px; }")
        sb.AppendLine(".w298 { width: 298px; }")
        sb.AppendLine(".w409 { width: 409px; }")
        sb.AppendLine(".w472 { width: 472px; }")
        sb.AppendLine(".l12 { font: bold 12px arial; color: #000000; }")
        sb.AppendLine(".pr5 { padding-right: 5px; }")
        sb.AppendLine(".w478 { width: 478px; }")
        sb.AppendLine(".w233 { width: 233px; }")
        sb.AppendLine(".w424 { width: 424px; }")
        sb.AppendLine(".w240 { width: 240px; }")
        sb.AppendLine(".w200 { width: 200px; }")
        sb.AppendLine(".w213 { width: 213px; }")
        sb.AppendLine(".w430 { width: 430px; }")
        sb.AppendLine(".w184 { width: 184px; }")
        sb.AppendLine(".w500 { width: 500px; }")
        sb.AppendLine(".w544 { width: 544px; }")
        sb.AppendLine(".w564 { width: 564px; }")
        sb.AppendLine(".w659 { width: 659px; }")
        sb.AppendLine(".w666 { width: 666px; }")
        sb.AppendLine(".w667 { width: 667px; }")
        sb.AppendLine(".BHead td { border-bottom: solid 2px #000; }")
        sb.AppendLine(".EcdBar { height: 50px; vertical-align: bottom; }")
        sb.AppendLine(".rc6 td { vertical-align: top; border-bottom: solid 1px #000; border-left: solid 1px #000; }")
        sb.AppendLine(".rc6 div { padding-left: 6px; }")
        sb.AppendLine(".rc6 .t { font: 9px 'arial narrow'; color: #000033; height: 13px; }")
        sb.AppendLine(".rc6 .c { font: bold 10px arial; color: black; height: 12px; text-align: right; }")
        sb.AppendLine(".mt23 { margin-top: 23px; }")
        sb.AppendLine(".pb4 { padding-bottom: 14px; }")
        sb.AppendLine(".ebc { width: 4px; height: 400px; border-right: dotted 1px #000000; margin-right: 4px; ")
        sb.AppendLine(".break { display: block; clear: both; page-break-after: always; }</style>")
        sb.AppendLine("    </head>")
        sb.AppendLine("    <body>")
        sb.AppendLine("        <table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='cpN'>")
        sb.AppendLine("                <td class='At Ac'>Instruções de Impressão</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='ti'>")
        sb.AppendLine("                <td class='At Ac'>Imprimir em impressora jato de tinta (ink jet) ou laser em qualidade normal. (Não use modo econômico).<br>Utilize folha A4 (210 x 297 mm) ou Carta (216 x 279 mm) - Corte na linha indicada<br></td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><br/><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr>")
        sb.AppendLine("                <td class='ctN cut'></td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr>")
        sb.AppendLine("                <td class='cpN Ar'>Recibo do Pagador</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><br/><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='BHead'>")
        sb.AppendLine("                <td class='imgLogo Al'><img src='@URLIMGCEDENTE'/></td>")
        sb.AppendLine("                <td class='barra'><img src='@URLIMAGEMBARRA'/></td>")
        sb.AppendLine("                <td class='w65 Ab bc Ac'>@CODIGOBANCO-@DIGITOBANCO</td>")
        sb.AppendLine("                <td class='barra'><img src='@URLIMAGEMBARRA'/></td>")
        sb.AppendLine("                <td class='w500 Ar Ab ld'>@LINHADIGITAVEL</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13 At'>")
        sb.AppendLine("                <td class='w298'>Beneficiário</td>")
        sb.AppendLine("                <td class='w126'>Agência / Código do Beneficiário</td>")
        sb.AppendLine("                <td class='w34'>Espécie</td>")
        sb.AppendLine("                <td class='w45'>Quantidade</td>")
        sb.AppendLine("                <td class='w128'>Carteira / Nosso número</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12 At rBb'>")
        sb.AppendLine("                <td>@CEDENTE</td>")
        sb.AppendLine("                <td>@AGENCIACODIGOCEDENTE</td>")
        sb.AppendLine("                <td>@ESPECIE</td>")
        sb.AppendLine("                <td>@QUANTIDADE</td>")
        sb.AppendLine("                <td class='Ar'>@NOSSONUMERO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13'>")
        sb.AppendLine("                <td class='w192'>Número do documento</td>")
        sb.AppendLine("                <td class='w132'>CPF/CNPJ</td>")
        sb.AppendLine("                <td class='w134'>Vencimento</td>")
        sb.AppendLine("                <td class='w180'>Valor documento</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12 rBb'>")
        sb.AppendLine("                <td>@NUMERODOCUMENTO</td>")
        sb.AppendLine("                <td>@CPFCNPJ</td>")
        sb.AppendLine("                <td>@DATAVENCIMENTO</td>")
        sb.AppendLine("                <td class='Ar'>@=VALORDOCUMENTO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13'>")
        sb.AppendLine("                <td class='w113'>(-) Desconto / Abatimentos</td>")
        sb.AppendLine("                <td class='w112'>(-) Outras deduções</td>")
        sb.AppendLine("                <td class='w113'>(+) Mora / Multa</td>")
        sb.AppendLine("                <td class='w113'>(+) Outros acréscimos</td>")
        sb.AppendLine("                <td class='w180'>(=) Valor cobrado</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12 rBb Ab'>")
        sb.AppendLine("                <td>@DESCONTOS</td>")
        sb.AppendLine("                <td>@OUTRASDEDUCOES</td>")
        sb.AppendLine("                <td>@MORAMULTA</td>")
        sb.AppendLine("                <td>@OUTROSACRESCIMOS</td>")
        sb.AppendLine("                <td class='Ar'>&nbsp;@VALORCOBRADO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13'>")
        sb.AppendLine("                <td class='w659'>Pagador</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12'>")
        sb.AppendLine("                <td>@SACADO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12 rBb'>")
        sb.AppendLine("                <td>@INFOSACADO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ctN h13'>")
        sb.AppendLine("                <td class='pL6'>Instruções</td>")
        sb.AppendLine("                <td class='w180 Ar'>Autenticação mecânica</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cpN h12'>")
        sb.AppendLine("                <td class='pL6 it'></td>")
        sb.AppendLine("                <td class='pL6 Ar'>@AUTENTICACAOMECANICA</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='ctN w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='h13'><td><td/></tr>")
        sb.AppendLine("            <tr class='h13'><td><td/></tr>")
        sb.AppendLine("            <tr><td class='Ar'>Corte na linha pontilhada</td></tr>")
        sb.AppendLine("            <tr><td class='cut'><td/></tr>")
        sb.AppendLine("            <tr class='h13'><td><td/></tr>")
        sb.AppendLine("            <tr class='h13'><td><td/></tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='BHead'>")
        sb.AppendLine("                <td class='imgLogo Al'><img src='@URLIMAGEMLOGO'/></td>")
        sb.AppendLine("                <td class='barra'><img src='@URLIMAGEMBARRA'/></td>")
        sb.AppendLine("                <td class='w65 Ab bc Ac'>@CODIGOBANCO-@DIGITOBANCO</td>")
        sb.AppendLine("                <td class='barra'><img src='@URLIMAGEMBARRA'/></td>")
        sb.AppendLine("                <td class='w500 Ar ld' valign='bottom'>@LINHADIGITAVEL</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13'>")
        sb.AppendLine("                <td class='w472'>Local de pagamento</td>")
        sb.AppendLine("                <td class='w180'>Vencimento</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12 rBb'>")
        sb.AppendLine("                <td>@LOCALPAGAMENTO</td>")
        sb.AppendLine("                <td class='Ar'>@DATAVENCIMENTO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13'>")
        sb.AppendLine("                <td class='w472'>Beneficiário</td>")
        sb.AppendLine("                <td class='w180'>Agência / Código do Beneficiário</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12 rBb'>")
        sb.AppendLine("                <td>@CEDENTE_BOLETO</td>")
        sb.AppendLine("                <td class='Ar'>@AGENCIACONTA</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13'>")
        sb.AppendLine("                <td class='w113'>Data do documento</td>")
        sb.AppendLine("                <td class='w163'>N<u>o</u> documento</td>")
        sb.AppendLine("                <td class='w62'>Espécie doc.</td>")
        sb.AppendLine("                <td class='w34'>Aceite</td>")
        sb.AppendLine("                <td class='w72'>Data processamento</td>")
        sb.AppendLine("                <td class='w180'>Carteira / Nosso número</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12 rBb'>")
        sb.AppendLine("                <td>@DATADOCUMENTO</td>")
        sb.AppendLine("                <td>@NUMERODOCUMENTO</td>")
        sb.AppendLine("                <td>@ESPECIEDOCUMENTO</td>")
        sb.AppendLine("                <td>@ACEITE</td>")
        sb.AppendLine("                <td>@DATAPROCESSAMENTO</td>")
        sb.AppendLine("                <td class='Ar'>@NOSSONUMERO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13'>")
        sb.AppendLine("                <td class='w113'>Uso do banco</td>")
        sb.AppendLine("                <td class='w83'>Carteira</td>")
        sb.AppendLine("                <td class='w53'>Espécie</td>")
        sb.AppendLine("                <td class='w123'>Quantidade</td>")
        sb.AppendLine("                <td class='w72'>(x) Valor</td>")
        sb.AppendLine("                <td class='w180'>(=) Valor documento</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12 rBb'>")
        sb.AppendLine("                <td>@USODOBANCO</td>")
        sb.AppendLine("                <td class='Al'>@CARTEIRA</td>")
        sb.AppendLine("                <td class='Al'>@ESPECIE</td>")
        sb.AppendLine("                <td>@QUANTIDADE</td>")
        sb.AppendLine("                <td>@VALORDOCUMENTO</td>")
        sb.AppendLine("                <td class='Ar'>@=VALORDOCUMENTO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='rc6'>")
        sb.AppendLine("                <td class='w478'>")
        sb.AppendLine("                    <div class='ctN pL10'>Instruções (Texto de responsabilidade do beneficiário)</div>")
        sb.AppendLine("                    <div class='cpN pL10 it' style='height:105px; overflow:hidden'>@INSTRUCOES</div>")
        sb.AppendLine("                </td>")
        sb.AppendLine("                <td class='w186'>")
        sb.AppendLine("                    <div class='t'>(-) Desconto / Abatimentos</div>")
        sb.AppendLine("                    <div class='c BB'>@DESCONTOS</div>")
        sb.AppendLine("                    <div class='t'>(-) Outras deduções</div>")
        sb.AppendLine("                    <div class='c BB'>@OUTRASDEDUCOES</div>")
        sb.AppendLine("                    <div class='t'>(+) Mora / Multa</div>")
        sb.AppendLine("                    <div class='c BB'>@MORAMULTA</div>")
        sb.AppendLine("                    <div class='t'>(+) Outros acréscimos</div>")
        sb.AppendLine("                    <div class='c BB'>@OUTROSACRESCIMOS</div>")
        sb.AppendLine("                    <div class='t'>(=) Valor cobrado</div>")
        sb.AppendLine("                    <div class='c'>@VALORCOBRADO</div>")
        sb.AppendLine("                </td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='ct h13'>")
        sb.AppendLine("                <td class='w659'>Pagador</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("            <tr class='cp h12'>")
        sb.AppendLine("                <td class='At'>@SACADO</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr class='rBb'>")
        sb.AppendLine("                <td class='w478 BL'>")
        sb.AppendLine("                    <div class='cpN pL6'>@INFOSACADO</div>")
        sb.AppendLine("                </td>")
        sb.AppendLine("                <td class='Ab BL'>")
        sb.AppendLine("                    <div class='ctN pL6'>Cód. baixa</div>")
        sb.AppendLine("                </td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666 ctN' style='margin: auto;'>")
        sb.AppendLine("            <tr>")
        sb.AppendLine("                <td class='pL6  w409'>Sacador / Avalista: @AVALISTA</td>")
        sb.AppendLine("                <td class='w250 Ar'>Autenticação mecânica - <b class='cpN'>Ficha de Compensação</b></td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='w666' style='margin: auto;'>")
        sb.AppendLine("            <tr>")
        sb.AppendLine("                <td class='EcdBar Al pL10'>@IMAGEMCODIGOBARRA</td>")
        sb.AppendLine("            </tr>")
        sb.AppendLine("        </table><table class='ctN w666' style='margin: auto;'>")
        sb.AppendLine("            <tr><td class='Ar'>Corte na linha pontilhada</td></tr>")
        sb.AppendLine("            <tr><td class='cut'></td></tr>")
        sb.AppendLine("        </table>")
        sb.AppendLine("    </body>")
        sb.AppendLine("</html>")

        Return sb.ToString()
    End Function

#End Region

#Region "Sicredi"
    Public Function SicrediArquivoCNAB400(ByVal Boletos As List(Of BoletoNet.Boleto)) As List(Of StringBuilder)
        Dim remessa As New RemessaBancaria()
        Dim sb As New List(Of StringBuilder)
        Dim i As Integer = 0
        Dim b = Boletos.First()

        'Header
        i = i + 1
        remessa.Header.Append(remessa.Format(1, 1, True, 0))
        remessa.Header.Append(remessa.Format(2, 2, True, 1))
        remessa.Header.Append(remessa.Format(3, 9, False, "REMESSA"))
        remessa.Header.Append(remessa.Format(10, 11, True, 1))
        remessa.Header.Append(remessa.Format(12, 19, False, "COBRANCA"))
        remessa.Header.Append(remessa.Format(20, 26, False, ""))
        remessa.Header.Append(remessa.Format(27, 31, True, b.Cedente.Codigo))
        remessa.Header.Append(remessa.Format(32, 45, True, b.Cedente.CPFCNPJ))
        remessa.Header.Append(remessa.Format(46, 76, False, ""))
        remessa.Header.Append(remessa.Format(77, 79, True, b.Banco.Codigo))
        remessa.Header.Append(remessa.Format(80, 94, False, "SICREDI"))
        remessa.Header.Append(remessa.Format(95, 102, True, DateTime.Now.ToString("yyyyMMdd")))
        remessa.Header.Append(remessa.Format(103, 110, False, ""))
        remessa.Header.Append(remessa.Format(111, 117, True, b.Cedente.CodigoTransmissao))
        remessa.Header.Append(remessa.Format(118, 390, False, ""))
        remessa.Header.Append(remessa.Format(391, 394, False, "2.00"))
        remessa.Header.Append(remessa.Format(395, 400, True, i))


        sb.Add(remessa.Header)

        Dim _detalhe As New StringBuilder

        For Each tit As BoletoNet.Boleto In Boletos

            Dim nossoNumero = NossoNumeroSicredi(tit)

            'Detalhe do Titulo
            i = i + 1
            _detalhe = New StringBuilder
            _detalhe.Append(remessa.Format(1, 1, True, 1))
            _detalhe.Append(remessa.Format(2, 2, False, "A"))
            _detalhe.Append(remessa.Format(3, 3, False, "A"))
            _detalhe.Append(remessa.Format(4, 4, False, "A"))
            _detalhe.Append(remessa.Format(5, 16, False, ""))
            _detalhe.Append(remessa.Format(17, 17, False, "A"))
            _detalhe.Append(remessa.Format(18, 18, False, "A"))
            _detalhe.Append(remessa.Format(19, 19, False, "B"))
            _detalhe.Append(remessa.Format(20, 47, False, ""))
            _detalhe.Append(remessa.Format(48, 56, True, nossoNumero))
            _detalhe.Append(remessa.Format(57, 62, False, ""))
            _detalhe.Append(remessa.Format(63, 70, True, DateTime.Now.ToString("yyyyMMdd")))
            _detalhe.Append(remessa.Format(71, 71, False, ""))
            _detalhe.Append(remessa.Format(72, 72, False, "N"))
            _detalhe.Append(remessa.Format(73, 73, False, ""))
            _detalhe.Append(remessa.Format(74, 74, False, "B"))
            _detalhe.Append(remessa.Format(75, 76, True, 1))
            _detalhe.Append(remessa.Format(77, 78, True, 1))
            _detalhe.Append(remessa.Format(79, 82, False, ""))
            _detalhe.Append(remessa.Format(83, 92, True, 0))
            _detalhe.Append(remessa.Format(93, 96, True, tit.PercMulta.ToString().Replace(",", "")))
            _detalhe.Append(remessa.Format(97, 108, False, ""))
            _detalhe.Append(remessa.Format(109, 110, True, 1))
            _detalhe.Append(remessa.Format(111, 120, False, tit.NumeroDocumento))
            _detalhe.Append(remessa.Format(121, 126, True, tit.DataVencimento.ToString("ddMMyy")))
            _detalhe.Append(remessa.Format(127, 139, True, tit.ValorBoleto.ToString().Replace(",", "")))
            _detalhe.Append(remessa.Format(140, 148, False, ""))
            _detalhe.Append(remessa.Format(149, 149, False, "A"))
            _detalhe.Append(remessa.Format(150, 150, False, "S"))
            _detalhe.Append(remessa.Format(151, 156, True, tit.DataDocumento.ToString("ddMMyy")))
            _detalhe.Append(remessa.Format(157, 158, True, 6))
            _detalhe.Append(remessa.Format(159, 160, True, 3))
            _detalhe.Append(remessa.Format(161, 173, True, tit.PercJurosMora.ToString().Replace(",", "")))
            _detalhe.Append(remessa.Format(174, 179, True, 0))
            _detalhe.Append(remessa.Format(180, 192, True, 0))
            _detalhe.Append(remessa.Format(193, 194, True, 0)) '
            _detalhe.Append(remessa.Format(195, 196, True, 0))
            _detalhe.Append(remessa.Format(197, 205, True, 0))
            _detalhe.Append(remessa.Format(206, 218, True, 0))
            'Sacado
            _detalhe.Append(remessa.Format(219, 219, True, IIf(tit.Sacado.CPFCNPJ.Length = 14, 2, 1)))
            _detalhe.Append(remessa.Format(220, 220, True, 0))
            _detalhe.Append(remessa.Format(221, 234, True, tit.Sacado.CPFCNPJ))
            _detalhe.Append(remessa.Format(235, 274, False, tit.Sacado.Nome))
            _detalhe.Append(remessa.Format(275, 314, False, tit.Sacado.Endereco.EndComNumeroEComplemento))
            _detalhe.Append(remessa.Format(315, 319, True, 0))
            _detalhe.Append(remessa.Format(320, 325, True, 0))
            _detalhe.Append(remessa.Format(326, 326, False, ""))
            _detalhe.Append(remessa.Format(327, 334, True, tit.Sacado.Endereco.CEP))
            _detalhe.Append(remessa.Format(335, 339, True, 0))
            _detalhe.Append(remessa.Format(340, 353, False, ""))
            _detalhe.Append(remessa.Format(354, 394, False, ""))
            _detalhe.Append(remessa.Format(395, 400, True, i))

            remessa.AddDetalhe(_detalhe)
            sb.Add(_detalhe)
        Next

        'Trailler
        i = i + 1
        Dim _trailer As New StringBuilder
        remessa.Trailler.Append(remessa.Format(1, 1, True, 9))
        remessa.Trailler.Append(remessa.Format(2, 2, True, 1))
        remessa.Trailler.Append(remessa.Format(3, 5, True, b.Banco.Codigo))
        remessa.Trailler.Append(remessa.Format(6, 10, True, b.Cedente.Codigo))
        remessa.Trailler.Append(remessa.Format(11, 394, False, " "))
        remessa.Trailler.Append(remessa.Format(395, 400, True, i))
        sb.Add(remessa.Trailler)

        Return sb
    End Function

    Public Function SicrediPdf(ByVal Boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        Dim _codigoDeBarras = SicrediCodigoDeBarras(Boleto)
        Dim DV = _codigoDeBarras.Last()
        Dim imagemCodBarras = "<img src='" + ImagemCodigoDeBarras(_codigoDeBarras) + "' alt='Código de Barras' />"
        Dim linhaDigitavel = SicrediLinhaDigitavel(Boleto, _codigoDeBarras)
        Dim nossoNumero = NossoNumeroSicredi(Boleto)
        'Formato Nosso Número AA/BXXXXX-D
        nossoNumero = nossoNumero.Insert(2, "/")
        nossoNumero = nossoNumero.Insert(9, "-")

        Dim pdf = Html()
        pdf = pdf.Replace("@CODIGOBANCO", Boleto.Banco.Codigo.ToString())
        pdf = pdf.Replace("@DIGITOBANCO", Boleto.Banco.Digito)
        pdf = pdf.Replace("@NOMEBANCO", Boleto.Banco.Nome)
        pdf = pdf.Replace("@LOCALPAGAMENTO", "PAGAVEL PREFERENCIALMENTE EM CANAIS ELETRONICOS DA SUA INSTITUICAO FINANCEIRA")
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "." + r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.OperacaConta) + "." + r.Format(1, 4, True, Boleto.Cedente.Codigo))
        'pdf = pdf.Replace("@URLIMGCEDENTE", "")
        pdf = pdf.Replace("@URLIMAGEMBARRA", "")
        pdf = pdf.Replace("@LINHADIGITAVEL", linhaDigitavel)
        pdf = pdf.Replace("@LOCALPAGAMENTO", Boleto.LocalPagamento)
        pdf = pdf.Replace("@DATAVENCIMENTO", Boleto.DataVencimento)
        pdf = pdf.Replace("@CEDENTE_BOLETO", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome, String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Cedente.Nome, Boleto.Cedente.CPFCNPJcomMascara)))
        pdf = pdf.Replace("@CEDENTE", Boleto.Cedente.Nome)
        pdf = pdf.Replace("@DATADOCUMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NUMERODOCUMENTO", Boleto.NumeroDocumento)
        pdf = pdf.Replace("@ESPECIEDOCUMENTO", "DMI")
        pdf = pdf.Replace("@DATAPROCESSAMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NOSSONUMERO", nossoNumero)
        pdf = pdf.Replace("@CARTEIRA", "")
        pdf = pdf.Replace("@ESPECIE", "REAL")
        pdf = pdf.Replace("@QUANTIDADE", IIf(Boleto.QuantidadeMoeda = 0, "", Boleto.QuantidadeMoeda.ToString()))
        pdf = pdf.Replace("@VALORDOCUMENTO", Boleto.ValorMoeda)
        pdf = pdf.Replace("@=VALORDOCUMENTO", Boleto.ValorBoleto.ToString("C", CultureInfo.GetCultureInfo("PT-BR")))
        pdf = pdf.Replace("@VALORCOBRADO", IIf(Boleto.ValorCobrado = 0, "", Boleto.ValorCobrado.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@OUTROSACRESCIMOS", "")
        pdf = pdf.Replace("@OUTRASDEDUCOES", "")
        pdf = pdf.Replace("@DESCONTOS", IIf(Boleto.ValorDesconto = 0, "", Boleto.ValorDesconto.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@AGENCIACONTA", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "-" + Boleto.Cedente.ContaBancaria.DigitoAgencia + "/" + r.Format(1, 7, True, Boleto.Cedente.ContaBancaria.Conta) + "-" + Boleto.Cedente.ContaBancaria.DigitoConta)
        pdf = pdf.Replace("@SACADO", String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Sacado.Nome, Convert.ToUInt64(Boleto.Sacado.CPFCNPJ).ToString("00\.000\.000\/0000\-00")))
        pdf = pdf.Replace("@INFOSACADO", Boleto.Sacado.Endereco.EndComNumeroEComplemento.ToString() + " - " + Boleto.Sacado.Endereco.Cidade + ", " + Boleto.Sacado.Endereco.UF)
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", Boleto.ContaBancaria.Agencia)
        pdf = pdf.Replace("@CPFCNPJ", Boleto.Cedente.CPFCNPJ)
        pdf = pdf.Replace("@MORAMULTA", IIf(Boleto.ValorMulta = 0, "", Boleto.ValorMulta.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@AUTENTICACAOMECANICA", "")
        pdf = pdf.Replace("@USODOBANCO", Boleto.UsoBanco)
        pdf = pdf.Replace("@IMAGEMCODIGOBARRA", imagemCodBarras) 'Imagem do código de barras
        pdf = pdf.Replace("@ACEITE", "N")
        pdf = pdf.Replace("@ENDERECOCEDENTE", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Endereco, ""))
        pdf = pdf.Replace("@AVALISTA", "")

        Dim texto As String = String.Empty

        For Each i In Boleto.Instrucoes
            texto += Environment.NewLine + i.Descricao
        Next
        pdf = pdf.Replace("@INSTRUCOES", texto)
        '          String.Format(
        '              "{0} - {1}",
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.Nome, String.Empty),
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.CPFCNPJ, String.Empty)))
        'pdf.Replace("Ar\" > R$", RemoveSimboloMoedaValorDocumento ? "Ar\">" :  "Ar\">R$")
        pdf = pdf.Replace("@PARCELATOTAL", "") 'IIf(Boleto.NumeroParcela <> 0 And Boleto.TotalParcela <> 0, Boleto.NumeroParcela + " / " + Boleto.TotalParcela, String.Empty))
        pdf = pdf.Replace("@DADOSCEDENTE", "") 'IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome + " - " + Boleto.Cedente.CPFCNPJ + "</br>" + Boleto.Cedente.Endereco.EndComNumeroEComplemento, ""))

        Return pdf
    End Function

    Private Function SicrediCodigoDeBarras(ByVal Boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        'Banco 3 posições
        Dim codigoDeBarras = Boleto.Banco.Codigo.ToString()
        'Moeda 1 posição
        codigoDeBarras += "9"
        'Digito Verificador 1 posição
        'codigoDeBarras += "1"
        'Fator de Vencimento 4 posições
        'Dim ts = Convert.ToDateTime(Boleto.DataVencimento).Subtract(Convert.ToDateTime("07/10/1997"))
        Dim fatorVencimento = BoletoNet.AbstractBanco.FatorVencimento(
            New BoletoNet.Boleto With
            {
                .DataVencimento = Boleto.DataVencimento
            })
        codigoDeBarras += fatorVencimento.ToString()
        'Valor 10 posições
        codigoDeBarras += r.Format(1, 10, True, Boleto.ValorBoleto.ToString().Replace(",", ""))

        'Campo Livre 25
        '1 – Com Registro 1 posição
        Dim campoLivre = "1"
        '1 - Carteira simples 1 posiçaõ
        campoLivre += "1"
        'Nosso número 9 posições
        campoLivre += r.Format(1, 9, True, NossoNumeroSicredi(Boleto))
        'Cooperativa 4 posições
        Dim agencia = r.Format(1, 4, True, Boleto.Cedente.ContaBancaria.Agencia)
        campoLivre += agencia
        'Posto 2 posições
        campoLivre += r.Format(1, 2, True, Boleto.Cedente.ContaBancaria.OperacaConta)
        'Código do beneficiário/cedente 5 posições
        campoLivre += r.Format(1, 5, True, Boleto.Cedente.Codigo)
        'valor expresso no campo “valor do documento” 1 posição
        campoLivre += "1"
        'Filler – “0” zeros 1 posição
        campoLivre += "0"
        'Digito Verificador do Campo Livre 1 posição
        campoLivre += Mod11SicrediCampoLivre(campoLivre)

        'Digito Verificador do Cod de Barras
        codigoDeBarras += campoLivre
        If codigoDeBarras.Length() <> 43 Then
            Throw New Exception("Codigo de Barras incorreto: " + codigoDeBarras.Length())
        End If

        Dim dv = Mod11Daycoval(codigoDeBarras)

        codigoDeBarras = codigoDeBarras.Insert(4, dv).ToString()

        Return codigoDeBarras
    End Function

    Private Function NossoNumeroSicredi(ByVal b As BoletoNet.Boleto) As String
        Dim remessa As New RemessaBancaria()
        'Digito Verificador
        'Agência Cooperativa
        Dim Ag = remessa.Format(1, 4, True, b.Cedente.ContaBancaria.Agencia)
        'Posto Cooperativa
        Dim posto = remessa.Format(1, 2, True, b.Cedente.ContaBancaria.OperacaConta)
        'Beneficiário
        Dim beneficiario = remessa.Format(1, 5, True, b.Cedente.Codigo)
        'Ano
        Dim ano = remessa.Format(1, 2, True, b.DataVencimento.ToString("yy"))
        'Byte
        Dim byteSicredi = "2" 'Verificar
        'Sequencial de Pagamentos
        Dim seq = remessa.Format(1, 5, True, b.NossoNumero.ToString())

        Dim nossoNumero = String.Format("{0}{1}{2}{3}{4}{5}", Ag, posto, beneficiario, ano, byteSicredi, seq)
        Dim dv = Mod11Sicredi(nossoNumero)

        Return String.Format("{0}{1}{2}{3}", ano, 2, seq, dv)
    End Function

    Private Function Mod11SicrediCampoLivre(ByVal campoLivre As String) As String
        Dim soma = 0
        'Dim pesos = Split("4,3,2,9,8,7,6,5,4,3,2,9,8,7,6,5,4,5,3,2,9,8,7,6,5,4,5,3,2,9,8,7,6,5,4,5,3,2,9,8,7,6,5,4,5,3,2", ",")
        Dim peso = 9
        Dim arr = campoLivre.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            soma += res
            'Pesos
            If peso = 2 Then
                peso = 9
            Else
                peso = (peso - 1)
            End If
        Next
        Dim parteInteiraDivisao = (soma \ 11)
        Dim resto = (soma - (11 * parteInteiraDivisao))

        If resto = 0 Or resto = 1 Then Return 0

        Dim dv = (11 - resto)
        Return dv
    End Function

    Private Function SicrediLinhaDigitavel(ByVal Boleto As BoletoNet.Boleto, ByVal codBarras As String) As String
        Dim r As New RemessaBancaria()
        Dim campoLivre = codBarras.Substring(19, 25)
        'Campo 1
        Dim campo1 = Boleto.Banco.Codigo.ToString() + Boleto.Moeda.ToString() + campoLivre.Substring(0, 5)
        campo1 += Mod10Sicredi(campo1, 2)
        Dim campo2 = campoLivre.Substring(5, 10)
        campo2 += Mod10Sicredi(campo2, 1)
        Dim campo3 = campoLivre.Substring(15, 10)
        campo3 += Mod10Sicredi(campo3, 1)
        Dim campo4 = codBarras.Substring(4, 1)
        Dim campo5 = codBarras.Substring(5, 4) + codBarras.Substring(9, 10)

        Return (campo1.Insert(5, ".") + " " + campo2.Insert(5, ".") + " " + campo3.Insert(5, ".") + " " + campo4 + " " + campo5)
    End Function

    Private Function Mod10Sicredi(ByVal campo As String, ByVal peso As Integer) As String
        Dim soma = 0

        Dim arr = campo.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            For Each numAux In res.ToString().ToCharArray()
                soma += Integer.Parse(numAux)
            Next

            peso = IIf(peso = 2, 1, 2)
        Next

        If (soma Mod 10) = 0 Then
            Return "0"
        End If

        Dim multiploDe10 = (Integer.Parse(soma.ToString().ToCharArray()(0)) + 1) * 10

        Dim dv = multiploDe10 - soma

        Return dv
    End Function

    Private Function Mod11Sicredi(ByVal numero As String) As String
        Dim soma = 0
        Dim peso = 4
        Dim arr = numero.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            soma += res
            'Pesos
            If peso = 2 Then
                peso = 9
            Else
                peso = (peso - 1)
            End If
        Next

        Dim dv = (11 - (soma Mod 11))

        If dv > 9 Then
            Return 0
        End If

        Return dv
    End Function

#End Region

#Region "Sicoob"
    Public Function SicoobArquivoCNAB240(ByVal Boletos As List(Of BoletoNet.Boleto)) As List(Of StringBuilder)
        Dim remessa As New RemessaBancaria()
        Dim sb As New List(Of StringBuilder)
        Dim i As Integer = 0
        Dim b = Boletos.First()
        Dim headerLoteContador As Integer = 0

        'Header Arquivo
        remessa.Header.Append(remessa.Format(1, 3, True, b.Banco.Codigo))
        remessa.Header.Append(remessa.Format(4, 7, True, 0))
        remessa.Header.Append(remessa.Format(8, 8, True, 0))
        remessa.Header.Append(remessa.Format(9, 17, False, ""))
        remessa.Header.Append(remessa.Format(18, 18, True, IIf(b.Cedente.CPFCNPJ.Length = 11, 1, 2)))
        remessa.Header.Append(remessa.Format(19, 32, True, b.Cedente.CPFCNPJ))
        remessa.Header.Append(remessa.Format(33, 52, False, ""))
        remessa.Header.Append(remessa.Format(53, 57, True, b.Cedente.ContaBancaria.Agencia))
        remessa.Header.Append(remessa.Format(58, 58, False, b.Cedente.ContaBancaria.DigitoAgencia))
        remessa.Header.Append(remessa.Format(59, 70, True, b.Cedente.ContaBancaria.Conta))
        remessa.Header.Append(remessa.Format(71, 71, True, b.Cedente.ContaBancaria.DigitoConta))
        remessa.Header.Append(remessa.Format(72, 72, False, "0"))
        remessa.Header.Append(remessa.Format(73, 102, False, b.Cedente.Nome))
        remessa.Header.Append(remessa.Format(103, 132, False, "SICOOB"))
        remessa.Header.Append(remessa.Format(133, 142, False, ""))
        remessa.Header.Append(remessa.Format(143, 143, False, "1"))
        remessa.Header.Append(remessa.Format(144, 151, True, DateTime.Now.ToString("ddMMyyyy")))
        remessa.Header.Append(remessa.Format(152, 157, True, DateTime.Now.ToString("HHss")))
        remessa.Header.Append(remessa.Format(158, 163, True, b.Cedente.CodigoTransmissao))
        remessa.Header.Append(remessa.Format(164, 166, True, 81))
        remessa.Header.Append(remessa.Format(167, 171, True, 0))
        remessa.Header.Append(remessa.Format(172, 191, False, ""))
        remessa.Header.Append(remessa.Format(192, 211, False, ""))
        remessa.Header.Append(remessa.Format(212, 240, False, ""))

        sb.Add(remessa.Header)

        'Header Lote
        remessa.Header = New StringBuilder()
        headerLoteContador = headerLoteContador + 1

        remessa.Header.Append(remessa.Format(1, 3, True, b.Banco.Codigo))
        remessa.Header.Append(remessa.Format(4, 7, True, headerLoteContador))
        remessa.Header.Append(remessa.Format(8, 8, True, 1))
        remessa.Header.Append(remessa.Format(9, 9, False, "R"))
        remessa.Header.Append(remessa.Format(10, 11, True, 1))
        remessa.Header.Append(remessa.Format(12, 13, False, ""))
        remessa.Header.Append(remessa.Format(14, 16, True, 40)) 'FIXO 040.
        remessa.Header.Append(remessa.Format(17, 17, False, ""))
        remessa.Header.Append(remessa.Format(18, 18, True, IIf(b.Cedente.CPFCNPJ.Length = 11, 1, 2)))
        remessa.Header.Append(remessa.Format(19, 33, True, b.Cedente.CPFCNPJ))
        remessa.Header.Append(remessa.Format(34, 53, False, ""))
        remessa.Header.Append(remessa.Format(54, 58, True, b.Cedente.ContaBancaria.Agencia))
        remessa.Header.Append(remessa.Format(59, 59, False, b.Cedente.ContaBancaria.DigitoAgencia))
        remessa.Header.Append(remessa.Format(60, 71, True, b.Cedente.ContaBancaria.Conta))
        remessa.Header.Append(remessa.Format(72, 72, True, b.Cedente.ContaBancaria.DigitoConta))
        remessa.Header.Append(remessa.Format(73, 73, False, ""))
        remessa.Header.Append(remessa.Format(74, 103, False, b.Banco.Nome))
        remessa.Header.Append(remessa.Format(104, 143, False, ""))
        remessa.Header.Append(remessa.Format(144, 183, False, ""))
        remessa.Header.Append(remessa.Format(184, 191, True, b.Cedente.CodigoTransmissao))
        remessa.Header.Append(remessa.Format(192, 199, True, DateTime.Now.ToString("ddMMyyyy")))
        remessa.Header.Append(remessa.Format(200, 207, True, 0))
        remessa.Header.Append(remessa.Format(208, 240, False, ""))

        sb.Add(remessa.Header)

        Dim _detalheSegP As New StringBuilder
        Dim _detalheSegQ As New StringBuilder
        Dim contadorLote As Integer = 0
        For Each tit As BoletoNet.Boleto In Boletos
            'Detalhe do Titulo SEGMENTO P
            i = i + 1
            contadorLote = contadorLote = 1
            _detalheSegP = New StringBuilder
            _detalheSegP.Append(remessa.Format(1, 3, True, tit.Banco.Codigo))
            _detalheSegP.Append(remessa.Format(4, 7, True, headerLoteContador))
            _detalheSegP.Append(remessa.Format(8, 8, True, 3)) 'Fixo 3
            _detalheSegP.Append(remessa.Format(9, 13, True, i))
            _detalheSegP.Append(remessa.Format(14, 14, False, "P"))
            _detalheSegP.Append(remessa.Format(15, 15, False, ""))
            _detalheSegP.Append(remessa.Format(16, 17, True, 1)) 'Fixo '01' = Entrada de Títulos
            _detalheSegP.Append(remessa.Format(18, 22, True, tit.Cedente.ContaBancaria.Agencia))
            _detalheSegP.Append(remessa.Format(23, 23, False, tit.Cedente.ContaBancaria.DigitoAgencia))
            _detalheSegP.Append(remessa.Format(24, 35, True, tit.Cedente.ContaBancaria.Conta))
            _detalheSegP.Append(remessa.Format(36, 36, True, tit.Cedente.ContaBancaria.DigitoConta))
            _detalheSegP.Append(remessa.Format(37, 37, False, ""))

            'Nosso numero
            'NumTitulo -10 posições (1 a 10): Vide planilha "02.Especificações do Boleto" deste arquivo item 3.13
            Dim nossoNumero = NossoNumeroSicoob(tit)
            'nossoNumero = nossoNumero.Substring(0, nossoNumero.Length - 1)
            _detalheSegP.Append(remessa.Format(38, 47, True, nossoNumero))
            'Parcela -2 posições (11 a 12) - "01" se parcela única
            _detalheSegP.Append(remessa.Format(48, 49, False, "01")) 'Fixo "01" se parcela única
            'Modalidade -2 posições (13 a 14) - vide planilha "Contracapa" deste arquivo
            _detalheSegP.Append(remessa.Format(50, 51, False, "01")) 'Fixo 01 - Simples Com Registro
            'Tipo Formulário - 1 posição  (15 a 15):
            '"1" -auto-copiativo
            '"3"-auto-envelopável
            '"4"-A4 sem envelopamento
            '"6"-A4 sem envelopamento 3 vias
            _detalheSegP.Append(remessa.Format(52, 52, False, "4"))
            'Em branco - 5 posições (16 a 20)
            _detalheSegP.Append(remessa.Format(53, 57, False, ""))
            'Fim Nosso numero

            _detalheSegP.Append(remessa.Format(58, 58, True, 1)) 'Fixo 1 - Carteira
            _detalheSegP.Append(remessa.Format(59, 59, True, 0))
            _detalheSegP.Append(remessa.Format(60, 60, False, ""))
            _detalheSegP.Append(remessa.Format(61, 61, True, 2)) 'Fixo '2' -  Beneficiário Emite
            _detalheSegP.Append(remessa.Format(62, 62, False, "2")) 'Fixo '2' - Beneficiário Distribui
            _detalheSegP.Append(remessa.Format(63, 77, False, tit.NumeroDocumento))
            _detalheSegP.Append(remessa.Format(78, 85, True, tit.DataVencimento.ToString("ddMMyyyy")))
            _detalheSegP.Append(remessa.Format(86, 100, True, tit.ValorBoleto.ToString().Replace(",", "")))
            _detalheSegP.Append(remessa.Format(101, 105, True, 0))
            _detalheSegP.Append(remessa.Format(106, 106, False, ""))
            _detalheSegP.Append(remessa.Format(107, 108, True, 2)) 'Fixo '02'  =  DM Duplicata Mercantil
            _detalheSegP.Append(remessa.Format(109, 109, False, "A"))
            _detalheSegP.Append(remessa.Format(110, 117, True, Now.ToString("ddMMyyyy")))
            _detalheSegP.Append(remessa.Format(118, 118, True, 1)) 'Fixo '1' - Valor por Dia
            _detalheSegP.Append(remessa.Format(119, 126, True, tit.DataJurosMora.AddDays(1).ToString("ddMMyyyy")))
            _detalheSegP.Append(remessa.Format(127, 141, True, tit.PercJurosMora.ToString().Replace(",", "")))
            _detalheSegP.Append(remessa.Format(142, 142, True, 0)) 'Fixo '0' - Não Conceder desconto
            _detalheSegP.Append(remessa.Format(143, 150, True, 0)) '
            _detalheSegP.Append(remessa.Format(151, 165, True, 0))
            _detalheSegP.Append(remessa.Format(166, 180, True, 0))
            _detalheSegP.Append(remessa.Format(181, 195, True, 0))
            _detalheSegP.Append(remessa.Format(196, 220, False, tit.NumeroDocumento))
            _detalheSegP.Append(remessa.Format(221, 221, True, 3)) 'Fixo '3' = Não Protestar
            _detalheSegP.Append(remessa.Format(222, 223, True, 0))
            _detalheSegP.Append(remessa.Format(224, 224, True, 0))
            _detalheSegP.Append(remessa.Format(225, 227, False, ""))
            _detalheSegP.Append(remessa.Format(228, 229, False, "09"))
            _detalheSegP.Append(remessa.Format(230, 239, True, 0))
            _detalheSegP.Append(remessa.Format(240, 240, False, ""))

            remessa.AddDetalhe(_detalheSegP)
            sb.Add(_detalheSegP)

            'Detalhe do Titulo SEGMENTO Q   
            i = i + 1
            contadorLote = contadorLote + 1
            _detalheSegQ = New StringBuilder
            _detalheSegQ.Append(remessa.Format(1, 3, True, tit.Banco.Codigo))
            _detalheSegQ.Append(remessa.Format(4, 7, True, headerLoteContador))
            _detalheSegQ.Append(remessa.Format(8, 8, True, 3)) 'Fixo 3
            _detalheSegQ.Append(remessa.Format(9, 13, True, i))
            _detalheSegQ.Append(remessa.Format(14, 14, False, "Q"))
            _detalheSegQ.Append(remessa.Format(15, 15, False, ""))
            _detalheSegQ.Append(remessa.Format(16, 17, True, 1)) 'Fixo '01' = Entrada de Títulos
            _detalheSegQ.Append(remessa.Format(18, 18, True, IIf(tit.Sacado.CPFCNPJ.Length = 11, 1, 2)))
            _detalheSegQ.Append(remessa.Format(19, 33, True, tit.Sacado.CPFCNPJ))
            _detalheSegQ.Append(remessa.Format(34, 73, True, tit.Sacado.Nome))
            _detalheSegQ.Append(remessa.Format(74, 113, False, tit.Sacado.Endereco.EndComNumeroEComplemento))
            _detalheSegQ.Append(remessa.Format(114, 128, False, tit.Sacado.Endereco.Bairro))
            _detalheSegQ.Append(remessa.Format(129, 133, True, tit.Sacado.Endereco.CEP))
            _detalheSegQ.Append(remessa.Format(134, 136, False, tit.Sacado.Endereco.CEP.Substring(5, 3)))
            _detalheSegQ.Append(remessa.Format(137, 151, True, tit.Sacado.Endereco.Cidade))
            _detalheSegQ.Append(remessa.Format(152, 153, True, tit.Sacado.Endereco.UF))
            _detalheSegQ.Append(remessa.Format(154, 154, True, IIf(tit.Sacado.CPFCNPJ.Length = 11, 1, 2)))
            _detalheSegQ.Append(remessa.Format(155, 169, True, tit.Sacado.CPFCNPJ))
            _detalheSegQ.Append(remessa.Format(170, 209, True, tit.Sacado.Nome))
            _detalheSegQ.Append(remessa.Format(210, 212, True, 0))
            _detalheSegQ.Append(remessa.Format(213, 232, False, ""))
            _detalheSegQ.Append(remessa.Format(233, 240, False, ""))

            remessa.AddDetalhe(_detalheSegQ)
            sb.Add(_detalheSegQ)

            'Detalhe do Titulo SEGMENTO Q   
            i = i + 1
            Dim _detalheSegR = New StringBuilder
            _detalheSegR.Append(remessa.Format(1, 3, True, tit.Banco.Codigo))
            _detalheSegR.Append(remessa.Format(4, 7, True, headerLoteContador))
            _detalheSegR.Append(remessa.Format(8, 8, True, 3)) 'Fixo 3
            _detalheSegR.Append(remessa.Format(9, 13, True, i))
            _detalheSegR.Append(remessa.Format(14, 14, False, "R"))
            _detalheSegR.Append(remessa.Format(15, 15, False, ""))
            _detalheSegR.Append(remessa.Format(16, 17, True, 1)) 'Fixo '01' = Entrada de Títulos
            _detalheSegR.Append(remessa.Format(18, 18, True, 0)) ' '0'  =  Não Conceder desconto
            _detalheSegR.Append(remessa.Format(19, 26, True, 0))
            _detalheSegR.Append(remessa.Format(27, 41, True, 0))
            _detalheSegR.Append(remessa.Format(42, 42, True, 0)) ' '0'  =  Não Conceder desconto
            _detalheSegR.Append(remessa.Format(43, 50, True, 0))
            _detalheSegR.Append(remessa.Format(51, 65, True, 0))
            _detalheSegR.Append(remessa.Format(66, 66, True, 1)) 'Multa '1'  =  Valor Fixo
            _detalheSegR.Append(remessa.Format(67, 74, True, tit.DataMulta.AddDays(1).ToString("ddMMyyyy")))
            _detalheSegR.Append(remessa.Format(75, 89, True, tit.ValorMulta.ToString().Replace(",", "")))
            _detalheSegR.Append(remessa.Format(90, 99, False, ""))
            _detalheSegR.Append(remessa.Format(100, 139, False, ""))
            _detalheSegR.Append(remessa.Format(140, 179, False, ""))
            _detalheSegR.Append(remessa.Format(180, 199, False, ""))
            _detalheSegR.Append(remessa.Format(200, 207, True, tit.DataVencimento.AddDays(1).ToString("ddMMyyyy")))
            _detalheSegR.Append(remessa.Format(208, 210, True, 0))
            _detalheSegR.Append(remessa.Format(211, 215, True, 0))
            _detalheSegR.Append(remessa.Format(216, 216, False, ""))
            _detalheSegR.Append(remessa.Format(217, 228, True, 0))
            _detalheSegR.Append(remessa.Format(229, 229, False, ""))
            _detalheSegR.Append(remessa.Format(230, 230, False, ""))
            _detalheSegR.Append(remessa.Format(231, 231, True, 0))
            _detalheSegR.Append(remessa.Format(232, 240, False, ""))

            remessa.AddDetalhe(_detalheSegR)
            sb.Add(_detalheSegR)
        Next

        'Trailler do Lote
        contadorLote = contadorLote + 1
        Dim _trailerLote As New StringBuilder
        _trailerLote.Append(remessa.Format(1, 3, True, b.Banco.Codigo))
        _trailerLote.Append(remessa.Format(4, 7, True, headerLoteContador))
        _trailerLote.Append(remessa.Format(8, 8, True, 5))
        _trailerLote.Append(remessa.Format(9, 17, False, ""))
        _trailerLote.Append(remessa.Format(18, 23, True, Boletos.Count()))
        _trailerLote.Append(remessa.Format(24, 29, True, Boletos.Count()))
        _trailerLote.Append(remessa.Format(30, 46, True, Boletos.Sum(Function(x) x.ValorBoleto).ToString().Replace(",", "")))
        _trailerLote.Append(remessa.Format(47, 52, True, 0))
        _trailerLote.Append(remessa.Format(53, 69, True, 0))
        _trailerLote.Append(remessa.Format(70, 75, True, 0))
        _trailerLote.Append(remessa.Format(76, 92, True, 0))
        _trailerLote.Append(remessa.Format(93, 98, True, 0))
        _trailerLote.Append(remessa.Format(99, 115, True, 0))
        _trailerLote.Append(remessa.Format(116, 123, False, ""))
        _trailerLote.Append(remessa.Format(124, 240, False, ""))
        sb.Add(_trailerLote)

        'Trailler do arquivo
        i = i + 1
        Dim _trailer As New StringBuilder
        remessa.Trailler.Append(remessa.Format(1, 3, True, b.Banco.Codigo))
        remessa.Trailler.Append(remessa.Format(4, 7, True, 9999))
        remessa.Trailler.Append(remessa.Format(8, 8, True, 9))
        remessa.Trailler.Append(remessa.Format(9, 17, False, ""))
        remessa.Trailler.Append(remessa.Format(18, 23, True, 3))
        remessa.Trailler.Append(remessa.Format(24, 29, True, i))
        remessa.Trailler.Append(remessa.Format(30, 35, True, 0))
        remessa.Trailler.Append(remessa.Format(36, 240, False, ""))
        sb.Add(remessa.Trailler)

        Return sb
    End Function

    Private Function SicoobLinhaDigitavel(ByVal Boleto As BoletoNet.Boleto, ByVal codBarras As String) As String
        'Campo 1		 	    	     	 
        'AAABC.DDDDE    			     
        Dim r As New RemessaBancaria()
        Dim nossoNumero = NossoNumeroSicoob(Boleto)
        'A = Código Do Sicoob na câmara de compensação - "756"	
        Dim campo1 = Boleto.Banco.Codigo.ToString()
        'B = Código da moeda - "9"	
        campo1 += Boleto.Moeda.ToString()
        'C = Código da carteira - verificar na planilha "Capa" deste arquivo
        campo1 += Boleto.Carteira
        'D = Código da agência/cooperativa - verificar na planilha "Capa" deste arquivo	
        campo1 += Boleto.Cedente.ContaBancaria.Agencia
        'E = Dígito verificador Do Campo 1 - vide demonstrativo de cálculo a seguir	
        campo1 += Mod10Sicredi(campo1, 2)

        'Campo 2
        'FFGGG.GGGGHI
        'F = Código da modalidade - verificar na planilha "Capa" deste arquivo
        Dim campo2 = r.Format(1, 2, True, Boleto.Carteira)
        'G = Código Do beneficiário/cliente - verificar na planilha "Capa" deste arquivo
        campo2 += Boleto.Cedente.Codigo.ToString()
        'H = Nosso número Do boleto		
        campo2 += nossoNumero.Substring(0, 1).ToString()
        'I = Dígito verificador Do Campo 2 - vide demonstrativo de cálculo a seguir	
        campo2 += Mod10Sicredi(campo2, 1)

        'Campo 3
        'HHHHH.HHJJJK
        'H = Nosso número Do boleto	
        Dim campo3 = nossoNumero.Substring(1, (nossoNumero.Length - 1)).ToString()
        'J = Número da parcela a que o boleto se refere - "001" se parcela única	
        campo3 += r.Format(1, 3, True, 1)
        'K = Dígito verificador Do Campo 3 - vide demonstrativo de cálculo a seguir
        campo3 += Mod10Sicoob(campo3, 1)

        'Campo 4
        'L
        'L = Dígito verificador Do Código de Barras - vide demonstrativo de cálculo a seguir	
        Dim campo4 = codBarras.Substring(4, 1)

        'Campo 5
        'MMMMNNNNNNNNNN
        'M = Fator de vencimento - vide demonstrativo de cálculo a seguir	
        Dim campo5 = codBarras.Substring(5, 4)
        'N = Valor Do boleto - Em casos de cobrança com valor em aberto (o valor a ser pago é preenchido pelo próprio pagador) ou cobrança em moeda variável, deve ser preenchido com zeros				
        campo5 += codBarras.Substring(9, 10)

        Return (campo1.Insert(5, ".") + " " + campo2.Insert(5, ".") + " " + campo3.Insert(5, ".") + " " + campo4 + " " + campo5)
    End Function

    Public Function SicoobPdf(ByVal Boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        Dim _codigoDeBarras = SicoobCodigoDeBarras(Boleto)
        Dim DV = _codigoDeBarras.Last()
        Dim imagemCodBarras = "<img src='" + ImagemCodigoDeBarras(_codigoDeBarras) + "' alt='Código de Barras' />"
        Dim linhaDigitavel = SicoobLinhaDigitavel(Boleto, _codigoDeBarras)
        Dim nossoNumero = NossoNumeroSicoob(Boleto)

        Dim pdf = Html()
        pdf = pdf.Replace("@CODIGOBANCO", Boleto.Banco.Codigo.ToString())
        pdf = pdf.Replace("@DIGITOBANCO", Boleto.Banco.Digito)
        pdf = pdf.Replace("@NOMEBANCO", Boleto.Banco.Nome)
        pdf = pdf.Replace("@LOCALPAGAMENTO", "PAGÁVEL PREFERENCIALMENTE NO SICOOB")
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "." + r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.OperacaConta) + "." + r.Format(1, 4, True, Boleto.Cedente.Codigo))
        'pdf = pdf.Replace("@URLIMGCEDENTE", "")
        pdf = pdf.Replace("@URLIMAGEMBARRA", "")
        pdf = pdf.Replace("@LINHADIGITAVEL", linhaDigitavel)
        pdf = pdf.Replace("@LOCALPAGAMENTO", Boleto.LocalPagamento)
        pdf = pdf.Replace("@DATAVENCIMENTO", Boleto.DataVencimento)
        pdf = pdf.Replace("@CEDENTE_BOLETO", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome, String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Cedente.Nome, Boleto.Cedente.CPFCNPJcomMascara)))
        pdf = pdf.Replace("@CEDENTE", Boleto.Cedente.Nome)
        pdf = pdf.Replace("@DATADOCUMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NUMERODOCUMENTO", Boleto.NumeroDocumento)
        pdf = pdf.Replace("@ESPECIEDOCUMENTO", "DM")
        pdf = pdf.Replace("@DATAPROCESSAMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NOSSONUMERO", String.Format("{0}/{1}", Boleto.Carteira, nossoNumero))
        pdf = pdf.Replace("@CARTEIRA", Boleto.Carteira)
        pdf = pdf.Replace("@ESPECIE", "REAL")
        pdf = pdf.Replace("@QUANTIDADE", IIf(Boleto.QuantidadeMoeda = 0, "", Boleto.QuantidadeMoeda.ToString()))
        pdf = pdf.Replace("@VALORDOCUMENTO", Boleto.ValorMoeda)
        pdf = pdf.Replace("@=VALORDOCUMENTO", Boleto.ValorBoleto.ToString("C", CultureInfo.GetCultureInfo("PT-BR")))
        pdf = pdf.Replace("@VALORCOBRADO", IIf(Boleto.ValorCobrado = 0, "", Boleto.ValorCobrado.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@OUTROSACRESCIMOS", "")
        pdf = pdf.Replace("@OUTRASDEDUCOES", "")
        pdf = pdf.Replace("@DESCONTOS", IIf(Boleto.ValorDesconto = 0, "", Boleto.ValorDesconto.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@AGENCIACONTA", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "-" + Boleto.Cedente.ContaBancaria.DigitoAgencia + "/" + r.Format(1, 7, True, Boleto.Cedente.ContaBancaria.Conta) + "-" + Boleto.Cedente.ContaBancaria.DigitoConta)
        pdf = pdf.Replace("@SACADO", String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Sacado.Nome, Convert.ToUInt64(Boleto.Sacado.CPFCNPJ).ToString("00\.000\.000\/0000\-00")))
        pdf = pdf.Replace("@INFOSACADO", Boleto.Sacado.Endereco.EndComNumeroEComplemento.ToString() + " - " + Boleto.Sacado.Endereco.Cidade + ", " + Boleto.Sacado.Endereco.UF)
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", Boleto.ContaBancaria.Agencia)
        pdf = pdf.Replace("@CPFCNPJ", Boleto.Cedente.CPFCNPJ)
        pdf = pdf.Replace("@MORAMULTA", IIf(Boleto.ValorMulta = 0, "", Boleto.ValorMulta.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        'pdf = pdf.Replace("@MORAMULTA", "")
        pdf = pdf.Replace("@AUTENTICACAOMECANICA", "")
        pdf = pdf.Replace("@USODOBANCO", Boleto.UsoBanco)
        pdf = pdf.Replace("@IMAGEMCODIGOBARRA", imagemCodBarras) 'Imagem do código de barras
        pdf = pdf.Replace("@ACEITE", "N")
        pdf = pdf.Replace("@ENDERECOCEDENTE", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Endereco, ""))
        pdf = pdf.Replace("@AVALISTA", "")

        Dim texto As String = String.Empty

        For Each i In Boleto.Instrucoes
            texto += Environment.NewLine + i.Descricao
        Next
        pdf = pdf.Replace("@INSTRUCOES", texto)
        '          String.Format(
        '              "{0} - {1}",
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.Nome, String.Empty),
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.CPFCNPJ, String.Empty)))
        'pdf.Replace("Ar\" > R$", RemoveSimboloMoedaValorDocumento ? "Ar\">" :  "Ar\">R$")
        pdf = pdf.Replace("@PARCELATOTAL", "") 'IIf(Boleto.NumeroParcela <> 0 And Boleto.TotalParcela <> 0, Boleto.NumeroParcela + " / " + Boleto.TotalParcela, String.Empty))
        pdf = pdf.Replace("@DADOSCEDENTE", "") 'IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome + " - " + Boleto.Cedente.CPFCNPJ + "</br>" + Boleto.Cedente.Endereco.EndComNumeroEComplemento, ""))

        Return pdf
    End Function

    Private Function SicoobCodigoDeBarras(ByVal Boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        'Deve conter 44 posições, disposto da seguinte forma
        'Posição     Tamanho     Conteúdo
        '01 a 03      03                    Código do Banco na Câmara de Compensação = '756'
        Dim codigoDeBarras = Boleto.Banco.Codigo.ToString()
        '04 a 04      01                    Código da Moeda = 9 (Real)
        codigoDeBarras += "9"
        '05 a 05      01                    Digito Verificador (DV) do Código de Barras
        '06 a 09      04                    Fator de Vencimento
        Dim days = DateDiff(DateInterval.Day, New DateTime(2000, 7, 3), New DateTime(Boleto.DataVencimento.Year, Boleto.DataVencimento.Month, Boleto.DataVencimento.Day))
        codigoDeBarras += (CInt(days) + 1000).ToString()
        '10 a 19      10                    Valor nominal do documento
        codigoDeBarras += r.Format(1, 10, True, Boleto.ValorBoleto.ToString().Replace(",", ""))

        '20 a 44      03                    Campo Livre
        'Composição do Campo Livre no Sicoob
        'Posição     Tamanho     Conteúdo
        '20 a 20      01                     Código da carteira de cobrança - vide planilha "Capa" deste arquivo
        Dim campoLivre = Boleto.Carteira
        '21 a 24      04                     Código da agência/cooperativa - verificar na planilha "Capa" deste arquivo
        campoLivre += Boleto.Cedente.ContaBancaria.Agencia
        '25 a 26      02                     Código da modalidade - verificar na planilha "Capa" deste arquivo
        campoLivre += r.Format(1, 2, True, Boleto.Carteira)
        '27 a 33      07                     Código do associado/cliente - verificar na planilha "Capa" deste arquivo
        campoLivre += r.Format(1, 7, True, Boleto.Cedente.Convenio)
        '34 a 41      08                     Nosso número do boleto
        Dim nossoNumero = NossoNumeroSicoob(Boleto)
        campoLivre += r.Format(1, 8, True, nossoNumero)
        '41 a 44      03                     Número da parcela a que o boleto se refere - "001" se parcela única
        campoLivre += r.Format(1, 3, True, 1)

        'campoLivre += Mod11SicrediCampoLivre(campoLivre)

        'Digito Verificador do Cod de Barras
        codigoDeBarras += campoLivre
        If codigoDeBarras.Length() <> 43 Then
            Throw New Exception("Codigo de Barras incorreto: " + codigoDeBarras.Length())
        End If

        Dim dv = Mod11Daycoval(codigoDeBarras)

        codigoDeBarras = codigoDeBarras.Insert(4, dv).ToString()

        Return codigoDeBarras
    End Function

    Private Function NossoNumeroSicoob(ByVal b As BoletoNet.Boleto) As String
        Dim remessa As New RemessaBancaria()
        'Digito Verificador
        'Agência Cooperativa
        Dim Ag = remessa.Format(1, 4, True, b.Cedente.ContaBancaria.Agencia)
        'Beneficiário
        Dim beneficiario = remessa.Format(1, 10, True, b.Cedente.Codigo)
        'Sequencial de Pagamentos
        Dim seq = remessa.Format(1, 7, True, b.NossoNumero.ToString())

        Dim nossoNumero = String.Format("{0}{1}{2}", Ag, beneficiario, seq)

        Dim dv = DVNossoNumeroSicoob(nossoNumero)

        Return String.Format("{0}{1}", remessa.Format(1, 7, True, seq), dv)
    End Function

    Private Function DVNossoNumeroSicoob(ByVal numero As String) As String
        Dim soma = 0
        Dim peso = 3197.ToString().ToCharArray()
        Dim contador = 0
        Dim arr = numero.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * Integer.Parse(peso(contador))
            soma += res
            'Pesos
            If contador = 3 Then
                contador = 0
            Else
                contador = (contador + 1)
            End If
        Next

        Dim dv = (11 - (soma Mod 11))

        If dv > 9 Then
            Return 0
        End If

        Return dv
    End Function

    Private Function Mod10Sicoob(ByVal campo As String, ByVal peso As Integer) As String
        Dim soma = 0

        Dim arr = campo.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            For Each numAux In res.ToString().ToCharArray()
                soma += Integer.Parse(numAux)
            Next

            peso = IIf(peso = 2, 1, 2)
        Next

        If (soma Mod 10) = 0 Then
            Return "0"
        End If

        Dim multiploDe10 = (Integer.Parse(soma.ToString().ToCharArray()(0)) + 1) * 10

        Dim dv = multiploDe10 - soma

        Return dv
    End Function
#End Region

#Region "Money Plus - Flow"
    Public Function MoneyPlusCNAB400(ByVal Boletos As List(Of BoletoNet.Boleto)) As List(Of StringBuilder)
        Dim remessa As New RemessaBancaria()
        Dim sb As New List(Of StringBuilder)
        Dim i As Integer = 0
        Dim b = Boletos.First()

        'Header
        i = i + 1
        remessa.Header.Append(remessa.Format(1, 1, True, 0))
        remessa.Header.Append(remessa.Format(2, 2, True, 1))
        remessa.Header.Append(remessa.Format(3, 9, False, "REMESSA"))
        remessa.Header.Append(remessa.Format(10, 11, True, 1))
        remessa.Header.Append(remessa.Format(12, 26, False, "COBRANCA"))
        remessa.Header.Append(remessa.Format(27, 46, True, b.Cedente.Codigo))
        remessa.Header.Append(remessa.Format(47, 76, False, b.Cedente.Nome))
        remessa.Header.Append(remessa.Format(77, 79, True, 274))
        remessa.Header.Append(remessa.Format(80, 94, False, "BMP Money Plus"))
        remessa.Header.Append(remessa.Format(95, 100, True, DateTime.Now.ToString("ddMMyyyy")))
        remessa.Header.Append(remessa.Format(101, 108, False, ""))
        remessa.Header.Append(remessa.Format(109, 110, False, "MX"))
        remessa.Header.Append(remessa.Format(111, 117, True, b.Cedente.CodigoTransmissao))
        remessa.Header.Append(remessa.Format(118, 394, False, ""))
        remessa.Header.Append(remessa.Format(395, 400, True, i))


        sb.Add(remessa.Header)

        Dim _detalhe As New StringBuilder

        For Each tit As BoletoNet.Boleto In Boletos
            'Detalhe do Titulo
            i = i + 1
            _detalhe = New StringBuilder
            '001 a 001 Identificação do Registro 001 1 X
            _detalhe.Append(remessa.Format(1, 1, True, 1))
            '002 a 006 Agência de Débito (opcional)005 00000 X
            _detalhe.Append(remessa.Format(2, 6, True, 0))
            '007 a 007 Dígito da Agência de Débito (opcional)001 Branco X
            _detalhe.Append(remessa.Format(7, 7, False, ""))
            '008 a 012 Razão da Conta Corrente (opcional) 005 00000 X
            _detalhe.Append(remessa.Format(8, 12, True, 0))
            '013 a 019 Conta Corrente (opcional) 007 0000000 X
            _detalhe.Append(remessa.Format(13, 19, True, 0))
            '020 a 020 Dígito da Conta Corrente (opcional) 001 Branco X
            _detalhe.Append(remessa.Format(20, 20, False, ""))

            '021 a 037 Identificação da Empresa Beneficiária no Banco 017 Detalhes página 12 X
            '21 a 21 - Zero
            _detalhe.Append(remessa.Format(21, 21, True, 0))
            '22 a 24 - Códigos da carteira
            _detalhe.Append(remessa.Format(22, 24, True, tit.Carteira))
            '25 a 29 - Códigos da Agência Beneficiários, sem o dígito
            _detalhe.Append(remessa.Format(25, 29, True, tit.Cedente.ContaBancaria.Agencia))
            '30 a 36 - Contas Corrente
            _detalhe.Append(remessa.Format(30, 36, True, tit.Cedente.ContaBancaria.Conta))
            '37 a 37 – Dígitos da Conta
            _detalhe.Append(remessa.Format(37, 37, True, tit.Cedente.ContaBancaria.DigitoConta))

            '038 a 052 No Controle do Participante 015 Uso da Empresa. Detalhes página 12 X
            _detalhe.Append(remessa.Format(38, 52, False, tit.NumeroDocumento))
            '053 a 062 Complemento ao No Controle do  Participante 010 Detalhes página 12 X
            _detalhe.Append(remessa.Format(53, 62, False, ""))
            '063 a 065 Código do Banco a ser debitado na Câmara de Compensação 003 000 X
            _detalhe.Append(remessa.Format(63, 65, True, 0))
            '066 a 066 Campo de Multa 001 Se = 2 considerar percentual de multa. Se = 0, sem multa. Vide Obs. página 12 X
            _detalhe.Append(remessa.Format(66, 66, True, 2))
            '067 a 070 Percentual de multa 004 Percentual de multa a ser considerado vide Obs. página 12 X
            _detalhe.Append(remessa.Format(67, 70, True, tit.PercMulta.ToString().Replace(",", "")))
            '071 a 081 Identificação do Título no Banco (Nosso Número) 
            Dim nossoNumero = NossoNumeroMoneyPlus(tit)
            _detalhe.Append(remessa.Format(71, 81, True, nossoNumero.Substring(0, 11)))
            '082 a 082 Digito de Auto Conferência do Número Bancário. Vide Obs. página 12 
            _detalhe.Append(remessa.Format(82, 82, False, nossoNumero.Substring(11, 1)))
            '083 a 092 Desconto Bonificação por dia 010 Valor do desconto bonif./dia. X
            _detalhe.Append(remessa.Format(83, 92, True, 0))
            '093 a 093 Condição para Emissão da Papeleta de Cobrança 001 2 
            _detalhe.Append(remessa.Format(93, 93, True, 2))
            '094 a 094 Ident. se emite Boleto para Débito Automático 001 N 
            _detalhe.Append(remessa.Format(94, 94, False, "N"))
            '095 a 104 Identificação da Operação do Banco 010 Brancos 
            _detalhe.Append(remessa.Format(95, 104, False, ""))
            '105 a 105 Indicador Rateio Crédito (opcional) 001 Branco 
            _detalhe.Append(remessa.Format(105, 105, False, ""))
            '106 a 106 Endereçamento para Aviso do Débito Automático em Conta Corrente (opcional) 001 0 X
            _detalhe.Append(remessa.Format(106, 106, True, 0))
            '107 a 108 Quantidade de pagamentos 002 Branco X
            _detalhe.Append(remessa.Format(107, 108, False, ""))
            '109 a 110 Identificação da ocorrência 002 Códigos de ocorrência Vide Obs. página 15 X
            '01 - Remessa
            '02 - Pedido de baixa
            '06 - Alteração de vencimento
            '07 - Alteração do controle do participante
            '20 – Alteração do valor
            _detalhe.Append(remessa.Format(109, 110, True, 1))
            '111 a 120 No do Documento 010 Documento X 
            _detalhe.Append(remessa.Format(111, 120, False, tit.NumeroDocumento))
            '121 a 126 Data do Vencimento do Título 006 DDMMAA. Vide Obs. página 15 X
            _detalhe.Append(remessa.Format(121, 126, True, tit.DataVencimento.ToString("ddMMyy")))
            '127 a 139 Valor do Título 013 Valor do Título (preencher sem ponto e sem vírgula) X
            _detalhe.Append(remessa.Format(127, 139, True, tit.ValorBoleto.ToString().Replace(",", "")))
            '140 a 142 Banco Encarregado da Cobrança 003 000 N
            _detalhe.Append(remessa.Format(140, 142, True, 0))
            '143 a 147 Agência Depositária 005 00000 N
            _detalhe.Append(remessa.Format(143, 147, True, 0))
            '148 a 149 Espécie de Título 002 Códigos Espécie Título Vide Obs. página 15  X
            '02 - Duplicata Mercantil (DM)
            _detalhe.Append(remessa.Format(148, 149, True, 2))
            '150 a 150 Identificação 001 Sempre = N X
            _detalhe.Append(remessa.Format(150, 150, False, "N"))
            '151 a 156 Data da emissão do Título 006 DDMMAA X
            _detalhe.Append(remessa.Format(151, 156, True, tit.DataDocumento.ToString("ddMMyy")))
            '157 a 158 1a instrução 002 00 X
            _detalhe.Append(remessa.Format(157, 158, True, 0))
            '159 a 160 2a instrução 002 00 X
            _detalhe.Append(remessa.Format(159, 160, True, 0))
            '161 a 173 Valor a ser cobrado por Dia de Atraso 013 Mora por Dia de Atraso Vide obs. página 16 X
            _detalhe.Append(remessa.Format(161, 173, True, tit.JurosMora.ToString().Replace(",", "")))
            '174 a 179 Data Limite P/Concessão de Desconto 006 DDMMAA X
            _detalhe.Append(remessa.Format(174, 179, True, tit.DataVencimento.ToString("ddMMyy"))) 'dúvida ??
            '180 a 192 Valor do Desconto 013 Valor Desconto (preenchersem ponto e sem vírgula) X Posição De/a Nome do Campo Tamanho do Campo Conteúdo A N
            _detalhe.Append(remessa.Format(180, 192, True, 0))
            '193 a 205 Valor do IOF 013 0000000000000 X
            _detalhe.Append(remessa.Format(193, 205, True, 0))
            '206 a 218 Valor do Abatimento a ser concedido ou cancelado 013 Valor Abatimento(preencher sem ponto e sem vírgula) X
            _detalhe.Append(remessa.Format(206, 218, True, 0))
            '219 a 220 Identificação do Tipo de Inscrição do Pagador 002 01-CPF 02- CNPJ X
            _detalhe.Append(remessa.Format(219, 220, True, IIf(tit.Sacado.CPFCNPJ.Length = 14, 2, 1)))
            '221 a 234 No Inscrição do Pagador 014 CNPJ/ CPF(Preenchimento obrigatório) X
            _detalhe.Append(remessa.Format(221, 234, True, tit.Sacado.CPFCNPJ))

            '235 a 274 Nome do Pagador 040 Nome do Pagador X
            _detalhe.Append(remessa.Format(235, 274, False, tit.Sacado.Nome))
            '275 a 314 Endereço Completo 040 Endereço do Pagador X
            _detalhe.Append(remessa.Format(275, 314, False, tit.Sacado.Endereco.EndComNumeroEComplemento.Replace(":", "")))
            '315 a 326 1a Mensagem 012 Vide Obs. página 16 X
            _detalhe.Append(remessa.Format(315, 326, False, "")) 'dúvida ??
            '327 a 331 CEP 005 CEP Pagador X
            _detalhe.Append(remessa.Format(327, 331, True, tit.Sacado.Endereco.CEP))
            '332 a 334 Sufixo do CEP 003 Sufixo X
            _detalhe.Append(remessa.Format(332, 334, True, tit.Sacado.Endereco.CEP))
            '335 a 335 Tipo Sacador Avalista 001 Obs. página 16 X
            _detalhe.Append(remessa.Format(335, 335, True, 0))
            '336 a 350 Identificador Sacador Avalista 015 Obs. página 16 X
            _detalhe.Append(remessa.Format(336, 350, True, 0))
            '351 a 394 Nome/Razão Social Sacador Avalista 044 Nome/Razão Social Sacador Avalista X
            _detalhe.Append(remessa.Format(351, 394, False, ""))
            '395 a 400 No Sequencial do Registro 006 No Sequencial do Registro
            _detalhe.Append(remessa.Format(395, 400, True, i))

            remessa.AddDetalhe(_detalhe)
            sb.Add(_detalhe)
        Next

        'Trailler
        i = i + 1
        Dim _trailer As New StringBuilder
        remessa.Trailler.Append(remessa.Format(1, 1, True, 9))
        remessa.Trailler.Append(remessa.Format(2, 394, False, ""))
        remessa.Trailler.Append(remessa.Format(395, 400, True, i))
        sb.Add(remessa.Trailler)

        Return sb
    End Function

    Public Function MoneyPlusPdf(ByVal Boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        Dim _codigoDeBarras = MoneyPlusCodigoDeBarras(Boleto)
        Dim DV = _codigoDeBarras.Last()
        Dim imagemCodBarras = "<img src='" + ImagemCodigoDeBarras(_codigoDeBarras) + "' alt='Código de Barras' />"
        Dim linhaDigitavel = MoneyPlusLinhaDigitavel(Boleto, _codigoDeBarras)
        Dim nossoNumero = NossoNumeroMoneyPlus(Boleto)
        'Formato Nosso Número AA/BXXXXX-D
        nossoNumero = nossoNumero.Insert(2, "/")
        nossoNumero = nossoNumero.Insert(9, "-")

        Dim pdf = Html()
        pdf = pdf.Replace("@CODIGOBANCO", 274)
        pdf = pdf.Replace("@DIGITOBANCO", 7)
        pdf = pdf.Replace("@NOMEBANCO", Boleto.Banco.Nome)
        pdf = pdf.Replace("@LOCALPAGAMENTO", "PAGAVEL PREFERENCIALMENTE EM CANAIS ELETRONICOS DA SUA INSTITUICAO FINANCEIRA")
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "." + r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.OperacaConta) + "." + r.Format(1, 4, True, Boleto.Cedente.Codigo))
        'pdf = pdf.Replace("@URLIMGCEDENTE", "")
        pdf = pdf.Replace("@URLIMAGEMBARRA", "")
        pdf = pdf.Replace("@LINHADIGITAVEL", linhaDigitavel)
        pdf = pdf.Replace("@LOCALPAGAMENTO", Boleto.LocalPagamento)
        pdf = pdf.Replace("@DATAVENCIMENTO", Boleto.DataVencimento)
        pdf = pdf.Replace("@CEDENTE_BOLETO", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome, String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Cedente.Nome, Boleto.Cedente.CPFCNPJcomMascara)))
        pdf = pdf.Replace("@CEDENTE", Boleto.Cedente.Nome)
        pdf = pdf.Replace("@DATADOCUMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NUMERODOCUMENTO", Boleto.NumeroDocumento)
        pdf = pdf.Replace("@ESPECIEDOCUMENTO", "DMI")
        pdf = pdf.Replace("@DATAPROCESSAMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NOSSONUMERO", nossoNumero)
        pdf = pdf.Replace("@CARTEIRA", Convert.ToInt32(Boleto.Cedente.Carteira).ToString("000"))
        pdf = pdf.Replace("@ESPECIE", "REAL")
        pdf = pdf.Replace("@QUANTIDADE", IIf(Boleto.QuantidadeMoeda = 0, "", Boleto.QuantidadeMoeda.ToString()))
        pdf = pdf.Replace("@VALORDOCUMENTO", Boleto.ValorMoeda)
        pdf = pdf.Replace("@=VALORDOCUMENTO", Boleto.ValorBoleto.ToString("C", CultureInfo.GetCultureInfo("PT-BR")))
        pdf = pdf.Replace("@VALORCOBRADO", IIf(Boleto.ValorCobrado = 0, "", Boleto.ValorCobrado.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@OUTROSACRESCIMOS", "")
        pdf = pdf.Replace("@OUTRASDEDUCOES", "")
        pdf = pdf.Replace("@DESCONTOS", IIf(Boleto.ValorDesconto = 0, "", Boleto.ValorDesconto.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@AGENCIACONTA", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "-" + Boleto.Cedente.ContaBancaria.DigitoAgencia + "/" + r.Format(1, 7, True, Boleto.Cedente.ContaBancaria.Conta) + "-" + Boleto.Cedente.ContaBancaria.DigitoConta)
        pdf = pdf.Replace("@SACADO", String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Sacado.Nome, Convert.ToUInt64(Boleto.Sacado.CPFCNPJ).ToString("00\.000\.000\/0000\-00")))
        pdf = pdf.Replace("@INFOSACADO", Boleto.Sacado.Endereco.EndComNumeroEComplemento.ToString() + " - " + Boleto.Sacado.Endereco.Cidade + ", " + Boleto.Sacado.Endereco.UF)
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", Boleto.ContaBancaria.Agencia)
        pdf = pdf.Replace("@CPFCNPJ", Boleto.Cedente.CPFCNPJ)
        pdf = pdf.Replace("@MORAMULTA", "")
        pdf = pdf.Replace("@AUTENTICACAOMECANICA", "")
        pdf = pdf.Replace("@USODOBANCO", Boleto.UsoBanco)
        pdf = pdf.Replace("@IMAGEMCODIGOBARRA", imagemCodBarras) 'Imagem do código de barras
        pdf = pdf.Replace("@ACEITE", "N")
        pdf = pdf.Replace("@ENDERECOCEDENTE", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Endereco, ""))
        pdf = pdf.Replace("@AVALISTA", "")

        Dim texto As String = String.Empty

        For Each i In Boleto.Instrucoes
            texto += Environment.NewLine + i.Descricao
        Next
        pdf = pdf.Replace("@INSTRUCOES", texto)
        '          String.Format(
        '              "{0} - {1}",
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.Nome, String.Empty),
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.CPFCNPJ, String.Empty)))
        'pdf.Replace("Ar\" > R$", RemoveSimboloMoedaValorDocumento ? "Ar\">" :  "Ar\">R$")
        pdf = pdf.Replace("@PARCELATOTAL", "") 'IIf(Boleto.NumeroParcela <> 0 And Boleto.TotalParcela <> 0, Boleto.NumeroParcela + " / " + Boleto.TotalParcela, String.Empty))
        pdf = pdf.Replace("@DADOSCEDENTE", "") 'IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome + " - " + Boleto.Cedente.CPFCNPJ + "</br>" + Boleto.Cedente.Endereco.EndComNumeroEComplemento, ""))

        Return pdf
    End Function

    Private Function MoneyPlusCodigoDeBarras(ByVal boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        '01 a 03 3 Identificação do Banco
        Dim codigoDeBarras = boleto.Banco.Codigo.ToString()
        '04 a 04 1 Código da Moeda (Real = 9, Outras=0)
        codigoDeBarras += "9"
        '05 a 05 1 Dígito verificador do Código de Barras
        '06 a 09 4 Fator de Vencimento (Vide Nota)
        'Dim ts = Convert.ToDateTime(boleto.DataVencimento).Subtract(Convert.ToDateTime("07/10/1997"))

        Dim fatorVencimento = BoletoNet.AbstractBanco.FatorVencimento(
            New BoletoNet.Boleto With
            {
                .DataVencimento = boleto.DataVencimento
            })

        codigoDeBarras += fatorVencimento.ToString()
        '10 a 19 10 Valor
        codigoDeBarras += r.Format(1, 10, True, boleto.ValorBoleto.ToString().Replace(",", ""))

        '20 a 44 25 Campo Livre
        '20 a 23 4 Agência Beneficiária (Sem o digito verificador, completar com zeros a esquerda quando necessário)
        Dim agencia = r.Format(1, 4, True, boleto.Cedente.ContaBancaria.Agencia)
        Dim campoLivre = agencia
        '24 a 25 2 Carteira
        campoLivre += boleto.Carteira
        '26 a 36 11 Número do Nosso Número (Sem o digito verificador)
        campoLivre += r.Format(1, 11, True, NossoNumeroMoneyPlus(boleto))
        '37 a 43 7 Conta do Beneficiário (Sem o digito verificador, completar com zeros a esquerda quando necessário)
        campoLivre += r.Format(1, 7, True, boleto.Cedente.ContaBancaria.Conta)
        '44 a 44 1 Zero
        campoLivre += r.Format(1, 1, True, 0)

        'Digito Verificador do Cod de Barras
        codigoDeBarras += campoLivre
        If codigoDeBarras.Length() <> 43 Then
            Throw New Exception("Codigo de Barras incorreto: " + codigoDeBarras.Length())
        End If

        Dim dv = Mod11MoneyPlusCampoLivre(codigoDeBarras)

        codigoDeBarras = codigoDeBarras.Insert(4, dv).ToString()

        Return codigoDeBarras
    End Function

    Private Function NossoNumeroMoneyPlus(ByVal b As BoletoNet.Boleto) As String
        Dim remessa As New RemessaBancaria()

        'Sequencial de Pagamentos
        Dim seq = remessa.Format(1, 11, True, b.NossoNumero.ToString())
        Dim carteira = remessa.Format(1, 2, True, b.Carteira.ToString())

        Dim dv = Mod11MoneyPlus(String.Format("{0}{1}", carteira, seq))

        Dim nossoNumero = String.Format("{0}{1}", seq, dv)
        Return nossoNumero
    End Function

    Private Function Mod11MoneyPlusCampoLivre(ByVal campoLivre As String) As String
        Dim soma = 0
        Dim peso = 2
        Dim arr = campoLivre.ToCharArray().Reverse()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            soma += res
            'Pesos
            If peso = 9 Then
                peso = 2
            Else
                peso = (peso + 1)
            End If
        Next
        Dim parteInteiraDivisao = (soma \ 11)
        Dim resto = (soma - (11 * parteInteiraDivisao))

        If resto = 0 Or resto = 1 Or resto > 9 Then Return 1

        Dim dv = (11 - resto)
        Return dv
    End Function

    Private Function MoneyPlusLinhaDigitavel(ByVal Boleto As BoletoNet.Boleto, ByVal codBarras As String) As String
        Dim r As New RemessaBancaria()
        Dim campoLivre = codBarras.Substring(19, 25)
        '1° campo Composto pelo código de Banco, código da moeda, as cinco primeiras posições do campo livre e o dígito verificador deste campo;
        Dim campo1 = Boleto.Banco.Codigo.ToString() + Boleto.Moeda.ToString() + campoLivre.Substring(0, 5)
        campo1 += Mod10MoneyPlus(campo1, 2)
        '2° Campo composto pelas posições 6 a 15 do campo livre e o dígito verificador deste campo;
        Dim campo2 = campoLivre.Substring(5, 10)
        campo2 += Mod10MoneyPlus(campo2, 1)
        '3° campo composto pelas posições 16 a 25 do campo livre e o dígito verificador deste campo;
        Dim campo3 = campoLivre.Substring(15, 10)
        campo3 += Mod10MoneyPlus(campo3, 1)
        '4° campo composto pelo dígito verificador do código de barras, ou seja, a 5a posição do código de barras;
        Dim campo4 = codBarras.Substring(4, 1)
        '5° Composto pelo fator de vencimento com 4(quatro) caracteres e o valor do documento com 10(dez) caracteres, sem separadores e sem edição.
        Dim campo5 = codBarras.Substring(5, 4) + codBarras.Substring(9, 10)


        Return (campo1.Insert(5, ".") + " " + campo2.Insert(5, ".") + " " + campo3.Insert(5, ".") + " " + campo4 + " " + campo5)
    End Function

    Private Function Mod10MoneyPlus(ByVal campo As String, ByVal peso As Integer) As String
        Dim soma = 0

        Dim arr = campo.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            For Each numAux In res.ToString().ToCharArray()
                soma += Integer.Parse(numAux)
            Next

            peso = IIf(peso = 2, 1, 2)
        Next

        If (soma Mod 10) = 0 Then
            Return "0"
        End If

        Dim multiploDe10 = (Integer.Parse(soma.ToString().ToCharArray()(0)) + 1) * 10

        Dim dv = multiploDe10 - soma

        Return dv
    End Function

    Private Function Mod11MoneyPlus(ByVal numero As String) As String
        Dim soma = 0
        Dim peso = 2
        Dim arr = numero.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            soma += res
            'Pesos
            If peso = 2 Then
                peso = 7
            Else
                peso = (peso - 1)
            End If
        Next

        Dim resto = (soma Mod 11)

        If resto = 0 Then
            Return "0"
        End If

        Dim dv = (11 - resto)

        If dv > 9 Then
            Return "0"
        End If

        Return dv
    End Function

#End Region

#Region "Unavanti"
    Public Function UnavantiCNAB400(ByVal Boletos As List(Of BoletoNet.Boleto)) As List(Of StringBuilder)
        Dim remessa As New RemessaBancaria()
        Dim sb As New List(Of StringBuilder)
        Dim i As Integer = 0
        Dim b = Boletos.First()

        'Header
        i = i + 1
        remessa.Header.Append(remessa.Format(1, 1, True, 0))
        remessa.Header.Append(remessa.Format(2, 2, True, 1))
        remessa.Header.Append(remessa.Format(3, 9, False, "REMESSA"))
        remessa.Header.Append(remessa.Format(10, 11, True, 1))
        remessa.Header.Append(remessa.Format(12, 26, False, "COBRANCA"))
        remessa.Header.Append(remessa.Format(27, 46, True, b.Cedente.Convenio))
        remessa.Header.Append(remessa.Format(47, 76, False, Funcoes.SubstituirCaracteresEspeciais(b.Cedente.Nome)))
        remessa.Header.Append(remessa.Format(77, 79, True, 460))
        remessa.Header.Append(remessa.Format(80, 94, False, "UNA"))
        remessa.Header.Append(remessa.Format(95, 100, True, DateTime.Now.ToString("ddMMyy")))
        remessa.Header.Append(remessa.Format(101, 108, False, ""))
        remessa.Header.Append(remessa.Format(109, 110, False, "MX"))
        remessa.Header.Append(remessa.Format(111, 117, True, b.Cedente.CodigoTransmissao))
        remessa.Header.Append(remessa.Format(118, 394, False, ""))
        remessa.Header.Append(remessa.Format(395, 400, True, i))


        sb.Add(remessa.Header)

        Dim _detalhe As New StringBuilder

        For Each tit As BoletoNet.Boleto In Boletos
            'Detalhe do Titulo
            i = i + 1
            _detalhe = New StringBuilder
            '001 a 001 Identificação do Registro 001 1 X
            _detalhe.Append(remessa.Format(1, 1, True, 1))
            '002 a 021 Zeros 00000 X
            _detalhe.Append(remessa.Format(2, 21, True, 0))
            '22 a 24 - Códigos da carteira
            _detalhe.Append(remessa.Format(22, 24, True, tit.Carteira))
            '25 a 29 - PREENCHER COM 00001
            _detalhe.Append(remessa.Format(25, 29, True, 1))
            '30 a 36 - Número da Conta-Corrente do cliente sem dígito
            _detalhe.Append(remessa.Format(30, 36, True, tit.Cedente.ContaBancaria.Conta))
            '37 a 37 – Dígito da Conta Corrente do cliente
            _detalhe.Append(remessa.Format(37, 37, True, tit.Cedente.ContaBancaria.DigitoConta))
            '38 a 62 - Brancos
            _detalhe.Append(remessa.Format(38, 62, False, ""))
            '063 a 065 Zeros
            _detalhe.Append(remessa.Format(63, 65, True, 0))
            '066 a 066 Se = 2 considerar percentual de multa. Se = 0 sem multa.
            _detalhe.Append(remessa.Format(66, 66, True, 2))
            '067 a 070 Percentual de multa 004 Percentual de multa a ser considerado vide Obs. página 12 X
            _detalhe.Append(remessa.Format(67, 70, True, tit.PercMulta.ToString().Replace(",", "")))
            '071 a 081 Identificação do Título no Banco (Nosso Número) 
            Dim nossoNumero = NossoNumeroUnavanti(tit)
            _detalhe.Append(remessa.Format(71, 81, True, nossoNumero.Substring(0, 11)))
            '082 a 082 Dígito de Auto conferência do Nosso Número
            _detalhe.Append(remessa.Format(82, 82, False, nossoNumero.Substring(11, 1)))
            '083 a 092 Zeros
            _detalhe.Append(remessa.Format(83, 92, True, 0))
            '093 a 093 Preencher com 2 
            _detalhe.Append(remessa.Format(93, 93, True, 2))
            '094 a 094 Zeros
            _detalhe.Append(remessa.Format(94, 94, True, 0))
            '095 a 108 Brancos 
            _detalhe.Append(remessa.Format(95, 108, False, ""))
            '109 a 110 Identificação da ocorrência 002 Códigos de ocorrência Vide Obs. página 15 X
            '01.. Remessa para Registro
            '02.. Pedido de baixa
            '04.. Concessão de Abatimento
            '05.. Cancelamento de Abatimento
            '06.. Alteração de Vencimento
            _detalhe.Append(remessa.Format(109, 110, True, 1))
            '111 a 120 No do Documento 010 Documento X 
            _detalhe.Append(remessa.Format(111, 120, False, tit.NumeroDocumento))
            '121 a 126 Data do Vencimento do Título 006 DDMMAA. Vide Obs. página 15 X
            _detalhe.Append(remessa.Format(121, 126, True, tit.DataVencimento.ToString("ddMMyy")))
            '127 a 139 Valor do Título 013 Valor do Título (preencher sem ponto e sem vírgula) X
            _detalhe.Append(remessa.Format(127, 139, True, tit.ValorBoleto.ToString().Replace(",", "")))
            '140 a 142 Banco Encarregado da Cobrança 003 000 N
            _detalhe.Append(remessa.Format(140, 142, True, 0))
            '143 a 147 Zeros
            _detalhe.Append(remessa.Format(143, 147, True, 0))
            '148 a 149 Espécie de Título 002 Códigos Espécie Título Vide Obs. página 15  X
            '01 - Duplicata Mercantil (DM)
            _detalhe.Append(remessa.Format(148, 149, True, 1))
            '150 a 150 Identificação 001 Sempre = N X
            _detalhe.Append(remessa.Format(150, 150, False, "N"))
            '151 a 156 Data da emissão do Título 006 DDMMAA 
            _detalhe.Append(remessa.Format(151, 156, True, tit.DataDocumento.ToString("ddMMyy")))
            '157 a 160 Zeros
            _detalhe.Append(remessa.Format(157, 160, True, 0))
            '161 a 173 Valor a ser cobrado por Dia de Atraso 013 Mora por Dia de Atraso Vide obs. página 16 X
            _detalhe.Append(remessa.Format(161, 173, True, Math.Round(tit.JurosMora, 2).ToString().Replace(",", "")))
            '174 a 179 Data Limite P/Concessão de Desconto 006 DDMMAA X
            _detalhe.Append(remessa.Format(174, 179, True, 0))
            '180 a 192 Valor do Desconto 013 Valor Desconto (preenchersem ponto e sem vírgula) X Posição De/a Nome do Campo Tamanho do Campo Conteúdo A N
            _detalhe.Append(remessa.Format(180, 192, True, 0))
            '193 a 205 Zeros
            _detalhe.Append(remessa.Format(193, 205, True, 0))
            '206 a 218 Valor do Abatimento a ser concedido ou cancelado 
            _detalhe.Append(remessa.Format(206, 218, True, 0))
            '219 a 220 Identificação do Tipo de Inscrição do Pagador 002 01-CPF 02- CNPJ X
            _detalhe.Append(remessa.Format(219, 220, True, IIf(tit.Sacado.CPFCNPJ.Length = 14, 2, 1)))
            '221 a 234 No Inscrição do Pagador 014 CNPJ/ CPF(Preenchimento obrigatório) X
            _detalhe.Append(remessa.Format(221, 234, True, tit.Sacado.CPFCNPJ))

            '235 a 274 Nome do Pagador 040 Nome do Pagador X
            _detalhe.Append(remessa.Format(235, 274, False, tit.Sacado.Nome))
            '275 a 314 Endereço Completo 040 Endereço do Pagador X
            Dim logradouro = Funcoes.SubstituirCaracteresEspeciais(tit.Sacado.Endereco.EndComNumeroEComplemento.Replace(":", ""))
            _detalhe.Append(remessa.Format(275, 314, False, logradouro))
            '315 a 326 Brancos
            _detalhe.Append(remessa.Format(315, 326, False, ""))
            '327 a 331 CEP 005 CEP Pagador X
            _detalhe.Append(remessa.Format(327, 331, True, tit.Sacado.Endereco.CEP))
            '332 a 334 Sufixo do CEP 003 Sufixo X
            _detalhe.Append(remessa.Format(332, 334, True, tit.Sacado.Endereco.CEP))
            '335 a 394 Sacador/Avalista ou 2ª Mensagem
            Dim messagem = String.Format("{0} - {1}", tit.Instrucoes(4).Descricao.Replace("<br />", ""), tit.Instrucoes(3).Descricao.Replace("<br />", ""))
            _detalhe.Append(remessa.Format(335, 394, False, Funcoes.SubstituirCaracteresEspeciais(messagem)))
            '395 a 400 No Sequencial do Registro 006 No Sequencial do Registro
            _detalhe.Append(remessa.Format(395, 400, True, i))

            remessa.AddDetalhe(_detalhe)
            sb.Add(_detalhe)
        Next

        'Trailler
        i = i + 1
        Dim _trailer As New StringBuilder
        remessa.Trailler.Append(remessa.Format(1, 1, True, 9))
        remessa.Trailler.Append(remessa.Format(2, 394, False, ""))
        remessa.Trailler.Append(remessa.Format(395, 400, True, i))
        sb.Add(remessa.Trailler)

        Return sb
    End Function

    Public Function UnavantiPdf(ByVal Boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        Dim _codigoDeBarras = UnavantiCodigoDeBarras(Boleto)
        Dim DV = _codigoDeBarras.Last()
        Dim imagemCodBarras = "<img src='" + ImagemCodigoDeBarras(_codigoDeBarras) + "' alt='Código de Barras' />"
        Dim linhaDigitavel = UnavantiLinhaDigitavel(Boleto, _codigoDeBarras)
        Dim nossoNumero = NossoNumeroUnavanti(Boleto)
        'Formato Nosso Número AA/BXXXXX-D
        nossoNumero = nossoNumero.Insert(11, "-")


        Dim pdf = Html()
        pdf = pdf.Replace("@CODIGOBANCO", 460)
        pdf = pdf.Replace("@DIGITOBANCO", 1)
        pdf = pdf.Replace("@NOMEBANCO", Boleto.Banco.Nome)
        pdf = pdf.Replace("@LOCALPAGAMENTO", "PAGAVEL PREFERENCIALMENTE EM CANAIS ELETRONICOS DA SUA INSTITUICAO FINANCEIRA")
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "." + r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.OperacaConta) + "." + r.Format(1, 4, True, Boleto.Cedente.Codigo))
        'pdf = pdf.Replace("@URLIMGCEDENTE", "")
        pdf = pdf.Replace("@URLIMAGEMBARRA", "")
        pdf = pdf.Replace("@LINHADIGITAVEL", linhaDigitavel)
        pdf = pdf.Replace("@LOCALPAGAMENTO", Boleto.LocalPagamento)
        pdf = pdf.Replace("@DATAVENCIMENTO", Boleto.DataVencimento)
        pdf = pdf.Replace("@CEDENTE_BOLETO", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome, String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Cedente.Nome, Boleto.Cedente.CPFCNPJcomMascara)))
        pdf = pdf.Replace("@CEDENTE", Boleto.Cedente.Nome)
        pdf = pdf.Replace("@DATADOCUMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NUMERODOCUMENTO", Boleto.NumeroDocumento)
        pdf = pdf.Replace("@ESPECIEDOCUMENTO", "DMI")
        pdf = pdf.Replace("@DATAPROCESSAMENTO", Boleto.DataDocumento.ToString("dd/MM/yyyy"))
        pdf = pdf.Replace("@NOSSONUMERO", String.Format("{0}/{1}", Boleto.Carteira, nossoNumero))
        pdf = pdf.Replace("@CARTEIRA", Convert.ToInt32(Boleto.Carteira).ToString("000"))
        pdf = pdf.Replace("@ESPECIE", "REAL")
        pdf = pdf.Replace("@QUANTIDADE", IIf(Boleto.QuantidadeMoeda = 0, "", Boleto.QuantidadeMoeda.ToString()))
        pdf = pdf.Replace("@VALORDOCUMENTO", Boleto.ValorMoeda)
        pdf = pdf.Replace("@=VALORDOCUMENTO", Boleto.ValorBoleto.ToString("C", CultureInfo.GetCultureInfo("PT-BR")))
        pdf = pdf.Replace("@VALORCOBRADO", IIf(Boleto.ValorCobrado = 0, "", Boleto.ValorCobrado.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@OUTROSACRESCIMOS", "")
        pdf = pdf.Replace("@OUTRASDEDUCOES", "")
        pdf = pdf.Replace("@DESCONTOS", IIf(Boleto.ValorDesconto = 0, "", Boleto.ValorDesconto.ToString("C", CultureInfo.GetCultureInfo("PT-BR"))))
        pdf = pdf.Replace("@AGENCIACONTA", r.Format(1, 3, True, Boleto.Cedente.ContaBancaria.Agencia) + "-" + Boleto.Cedente.ContaBancaria.DigitoAgencia + "/" + r.Format(1, 7, True, Boleto.Cedente.ContaBancaria.Conta) + "-" + Boleto.Cedente.ContaBancaria.DigitoConta)
        pdf = pdf.Replace("@SACADO", String.Format("{0}&nbsp;&nbsp;&nbsp;CNPJ: {1}", Boleto.Sacado.Nome, Convert.ToUInt64(Boleto.Sacado.CPFCNPJ).ToString("00\.000\.000\/0000\-00")))
        pdf = pdf.Replace("@INFOSACADO", Boleto.Sacado.Endereco.EndComNumeroEComplemento.ToString() + " - " + Boleto.Sacado.Endereco.Cidade + ", " + Boleto.Sacado.Endereco.UF)
        pdf = pdf.Replace("@AGENCIACODIGOCEDENTE", Boleto.ContaBancaria.Agencia)
        pdf = pdf.Replace("@CPFCNPJ", Boleto.Cedente.CPFCNPJ)
        pdf = pdf.Replace("@MORAMULTA", "")
        pdf = pdf.Replace("@AUTENTICACAOMECANICA", "")
        pdf = pdf.Replace("@USODOBANCO", Boleto.UsoBanco)
        pdf = pdf.Replace("@IMAGEMCODIGOBARRA", imagemCodBarras) 'Imagem do código de barras
        pdf = pdf.Replace("@ACEITE", "N")
        pdf = pdf.Replace("@ENDERECOCEDENTE", IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Endereco, ""))
        pdf = pdf.Replace("@AVALISTA", "")

        Dim texto As String = String.Empty

        For Each i In Boleto.Instrucoes
            texto += Environment.NewLine + i.Descricao
        Next
        pdf = pdf.Replace("@INSTRUCOES", texto)
        '          String.Format(
        '              "{0} - {1}",
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.Nome, String.Empty),
        '                IIf(Boleto.Avalista IsNot Nothing, Boleto.Avalista.CPFCNPJ, String.Empty)))
        'pdf.Replace("Ar\" > R$", RemoveSimboloMoedaValorDocumento ? "Ar\">" :  "Ar\">R$")
        pdf = pdf.Replace("@PARCELATOTAL", "") 'IIf(Boleto.NumeroParcela <> 0 And Boleto.TotalParcela <> 0, Boleto.NumeroParcela + " / " + Boleto.TotalParcela, String.Empty))
        pdf = pdf.Replace("@DADOSCEDENTE", "") 'IIf(Boleto.Cedente.MostrarCNPJnoBoleto, Boleto.Cedente.Nome + " - " + Boleto.Cedente.CPFCNPJ + "</br>" + Boleto.Cedente.Endereco.EndComNumeroEComplemento, ""))

        Return pdf
    End Function

    Private Function UnavantiCodigoDeBarras(ByVal boleto As BoletoNet.Boleto) As String
        Dim r = New RemessaBancaria()
        '01 a 03 3 Identificação do Banco
        Dim codigoDeBarras = "460"
        '04 a 04 1 Código da Moeda (Real = 9, Outras=0)
        codigoDeBarras += "9"
        '05 a 05 1 Dígito verificador do Código de Barras
        '06 a 09 4 Fator de Vencimento (Vide Nota)
        Dim fatorVencimento = BoletoNet.AbstractBanco.FatorVencimento(
            New BoletoNet.Boleto With
            {
                .DataVencimento = boleto.DataVencimento
            })

        codigoDeBarras += fatorVencimento.ToString()
        '10 a 19 10 Valor
        codigoDeBarras += r.Format(1, 10, True, boleto.ValorBoleto.ToString().Replace(",", ""))

        '20 a 44 25 Campo Livre
        '20 a 23 4 Agência Beneficiária (Sem o digito verificador, completar com zeros a esquerda quando necessário)
        Dim agencia = r.Format(1, 4, True, boleto.Cedente.ContaBancaria.Agencia)
        Dim campoLivre = agencia
        '24 a 25 2 Carteira
        campoLivre += boleto.Carteira
        '26 a 36 11 Número do Nosso Número (Sem o digito verificador)
        Dim nossoNumero = r.Format(1, 11, True, NossoNumeroUnavanti(boleto))
        campoLivre += nossoNumero
        '37 a 43 7 Número Convênio, fornecido pelo banco (Completar com zeros a esquerda quando necessário)
        campoLivre += r.Format(1, 7, True, boleto.Cedente.Convenio)
        '44 a 44 1 Zero
        campoLivre += r.Format(1, 1, True, 0)

        'Digito Verificador do Cod de Barras
        codigoDeBarras += campoLivre
        If codigoDeBarras.Length() <> 43 Then
            Throw New Exception("Codigo de Barras incorreto: " + codigoDeBarras.Length())
        End If

        Dim dv = Mod11MoneyPlusCampoLivre(codigoDeBarras)

        codigoDeBarras = codigoDeBarras.Insert(4, dv).ToString()

        Return codigoDeBarras
    End Function

    Private Function NossoNumeroUnavanti(ByVal b As BoletoNet.Boleto) As String
        Dim remessa As New RemessaBancaria()

        'Sequencial de Pagamentos
        Dim seq = remessa.Format(1, 11, True, b.NossoNumero.ToString())
        Dim carteira = remessa.Format(1, 2, True, b.Carteira.ToString())

        Dim dv = Mod11Unavanti(String.Format("{0}{1}", carteira, seq))

        Dim nossoNumero = String.Format("{0}{1}", seq, dv)
        Return nossoNumero
    End Function

    Private Function Mod11UnavantiCampoLivre(ByVal campoLivre As String) As String
        Dim soma = 0
        Dim peso = 2
        Dim arr = campoLivre.ToCharArray().Reverse()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            soma += res
            'Pesos
            If peso = 9 Then
                peso = 2
            Else
                peso = (peso + 1)
            End If
        Next
        Dim parteInteiraDivisao = (soma \ 11)
        Dim resto = (soma - (11 * parteInteiraDivisao))

        If resto = 0 Or resto = 1 Or resto > 9 Then Return 1

        Dim dv = (11 - resto)
        Return dv
    End Function

    Private Function UnavantiLinhaDigitavel(ByVal Boleto As BoletoNet.Boleto, ByVal codBarras As String) As String
        Dim r As New RemessaBancaria()
        Dim campoLivre = codBarras.Substring(19, 25)

        '1° campo Composto pelo código de Banco, código da moeda, as cinco primeiras posições do campo livre e o dígito verificador deste campo;
        Dim campo1 = "460" + Boleto.Moeda.ToString() + campoLivre.Substring(0, 5)
        campo1 += Mod10Unavanti(campo1, 2)

        '2° Campo composto pelas posições 6 a 15 do campo livre e o dígito verificador deste campo;
        Dim campo2 = campoLivre.Substring(5, 10)
        campo2 += Mod10Unavanti(campo2, 1)

        '3° campo composto pelas posições 16 a 25 do campo livre e o dígito verificador deste campo;
        Dim campo3 = campoLivre.Substring(15, 10)
        campo3 += Mod10Unavanti(campo3, 1)

        '4° campo composto pelo dígito verificador do código de barras, ou seja, a 5a posição do código de barras;
        Dim campo4 = codBarras.Substring(4, 1)

        '5° Composto pelo fator de vencimento com 4(quatro) caracteres e o valor do documento com 10(dez) caracteres, sem separadores e sem edição.
        Dim campo5 = codBarras.Substring(5, 4) + codBarras.Substring(9, 10)

        Return (campo1.Insert(5, ".") + " " + campo2.Insert(5, ".") + " " + campo3.Insert(5, ".") + " " + campo4 + " " + campo5)
    End Function

    Private Function Mod10Unavanti(ByVal campo As String, ByVal peso As Integer) As String
        Dim soma = 0

        Dim arr = campo.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            For Each numAux In res.ToString().ToCharArray()
                soma += Integer.Parse(numAux)
            Next

            peso = IIf(peso = 2, 1, 2)
        Next

        If (soma Mod 10) = 0 Then
            Return "0"
        End If

        Dim multiploDe10 = (10 - soma Mod 10) + soma

        Dim dv = multiploDe10 - soma

        Return dv
    End Function

    Private Function Mod11Unavanti(ByVal numero As String) As String
        Dim soma = 0
        Dim peso = 2
        Dim arr = numero.ToCharArray()
        For Each num In arr
            Dim res = Integer.Parse(num) * (peso)
            soma += res
            'Pesos
            If peso = 2 Then
                peso = 7
            Else
                peso = (peso - 1)
            End If
        Next

        Dim resto = (soma Mod 11)

        If resto = 0 Then
            Return "0"
        End If

        Dim dv = (11 - resto)

        If dv > 9 Then
            Return "P"
        End If

        Return dv
    End Function

#End Region

End Class
