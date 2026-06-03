Imports System.Data
Imports System.Net
Imports System.Net.Mail
Imports System.Web
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports AjaxControlToolkit
Imports Boleto2Net
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ReemissaoBoletoBancario
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ReemissaoBoletoBancario", "ACESSAR") Then
                Limpar()
                BuncarUnidadeDeNegocio()
                CarregarBancos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
    End Sub

    Private Sub Limpar()

        lnkImpressao.Parent.Visible = False

        txtCodigoCliente.Value = 0
        txtClientes.Text = String.Empty
        txtNotaFiscal.Text = String.Empty
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

        chkPeriodo.Checked = True

        gridBoletoBancario.DataSource = Nothing
        gridBoletoBancario.DataBind()
    End Sub

    Private Sub Download(pPath As String, pNomeArquivo As String, Optional pType As String = "text/plain")
        Try
            Response.Clear()
            Response.ContentType = pType
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & pNomeArquivo)
            Response.TransmitFile(Server.MapPath("~/" & pPath & "/" & pNomeArquivo))
            Response.Flush()
            Response.End()
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao fazer o download do arquivo! contate o suporte")
        End Try
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlEmpresa.SelectedIndex > 0 Then
                CarregarBancos()
            Else
                ddlBanco.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarBancos()

        Dim Empresa As String() = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim where As String = "Where Ativo = 1 AND left(Empresa_Id,8) = '" & Left(Empresa(0), 8) + "'"
        Dim objReceber As New [Lib].Negocio.ListParametrosDaCarteiraDeCobranca(where)

        ddlBanco.Items.Clear()

        For Each row As [Lib].Negocio.ParametrosDaCarteiraDeCobranca In objReceber

            Dim banco As New [Lib].Negocio.Banco(row.CodigoBanco)

            Dim agencia = IIf(row.DigitoAgencia.Length > 0, Funcoes.AlinharDireita(row.CodigoAgencia, 4, " ") & "-" & row.DigitoAgencia, Funcoes.AlinharDireita(row.CodigoAgencia, 6, "0"))
            Dim conta = IIf(row.DigitoConta.Length > 0, Funcoes.AlinharDireita(row.Conta, 10, " ") & "-" & row.DigitoConta, Funcoes.AlinharDireita(row.Conta, 10, " "))

            ddlBanco.Items.Add(New ListItem(Funcoes.AlinharEsquerda(banco.Descricao, 35, ".") & " - AG: " & agencia & " - CTA: " & conta & " CONVÊNIO: " & row.Convenio,
                                 row.CodigoBanco & ";" & row.CodigoAgencia & ";" & row.DigitoAgencia & ";" & row.Conta & ";" & row.DigitoConta & ";" & row.Convenio))
        Next

        Funcoes.InserirLinhaEmBranco(ddlBanco)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objCliente" & HID.Value) IsNot Nothing Then
            Dim cli As Cliente = Session("objCliente" & HID.Value)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Function BoletosDaycoval(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(param.CodigoBanco)
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()
            b.NumeroControle = param.SubConta
            b.NossoNumero = param.SequenciaDePagamentos
            b.DataDocumento = DateTime.Now.Date
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.EspecieDocumento = New BoletoNet.EspecieDocumento(param.CodigoBanco, "12")
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(param.CodigoBanco)
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao.AddDays(1)
                b.PercMulta = param.JurosMes
                b.ValorMulta = CDec(b.ValorBoleto * b.PercMulta / 100)

                Dim instrucao2 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If

            boletos.Add(b)
            param.SequenciaDePagamentos += 1
        Next
        Return boletos
    End Function

    Protected Function BoletosSicoob(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        Dim i As Integer = 0
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(param.CodigoBanco)
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta,
                    .OperacaConta = param.SubConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio,
                .CodigoTransmissao = param.SequenciaDeRemessa
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()

            If (tit.NossoNumero <> "0" AndAlso (Not String.IsNullOrWhiteSpace(tit.NossoNumero))) Then
                b.NossoNumero = tit.NossoNumero
            Else
                b.NossoNumero = (param.SequenciaDePagamentos + i)
                i = (i + 1)
            End If

            b.DataDocumento = tit.UsuarioInclusaoData
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(param.CodigoBanco)
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao
                b.PercMulta = param.JurosMes
                b.ValorMulta = Math.Round(CDec(b.ValorBoleto * b.PercMulta / 100), 2)

                Dim instrucao2 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If
            'Mora dia
            If (param.MoraDiaria > 0) Then
                'Furlan - 26-07-2021 - Química é diferente pois cobra 3,50% ao mês dividido por 30.  
                If param.CodigoEmpresa = "05272759000147" Then
                    b.PercJurosMora = Math.Round((param.MoraDiaria / 30), 2)
                    b.JurosMora = Math.Round(b.ValorBoleto * b.PercJurosMora / 100, 2)
                Else
                    b.PercJurosMora = param.MoraDiaria
                    b.JurosMora = (b.ValorBoleto / 30) * b.PercJurosMora / 100
                End If
                b.DataJurosMora = tit.Prorrogacao
                Dim instrucao3 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao3.Descricao = "<br /> APOS VENCIMENTO COBRAR R$ " & b.JurosMora.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"
                b.Instrucoes.Add(instrucao3)

            End If
            'Precisa gravar no titulo o nosso numero para gerar o pdf
            'tit.NossoNumero = param.SequenciaDePagamentos
            boletos.Add(b)
        Next
        Return boletos
    End Function

    Protected Function BoletosUnavanti(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(237) 'Unavanti usa o codigo do Bradesco
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta,
                    .OperacaConta = param.SubConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio,
                .CodigoTransmissao = param.SequenciaDeRemessa
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()
            b.NossoNumero = IIf(tit.NossoNumero = "0", param.SequenciaDePagamentos, tit.NossoNumero)
            b.DataDocumento = tit.UsuarioInclusaoData
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao.AddDays(1)
                b.PercMulta = param.JurosMes
                b.ValorMulta = CDec(b.ValorBoleto * b.PercMulta / 100)

                Dim instrucao2 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If
            'Mora dia
            If (param.MoraDiaria > 0) Then
                b.PercJurosMora = param.MoraDiaria
                b.JurosMora = (b.ValorBoleto / 30) * b.PercJurosMora / 100
                Dim instrucao3 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
                instrucao3.Descricao = "<br /> APOS VENCIMENTO COBRAR R$ " & b.JurosMora.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"
                b.Instrucoes.Add(instrucao3)

            End If

            Dim instrucao4 = New BoletoNet.Instrucao(237) 'Usa igual ao Bradesco
            instrucao4.Descricao = "<br /> PROTESTAR APÓS " + param.DiasParaProtesto.ToString() + " DIAS DO VENCIMENTO."
            b.Instrucoes.Add(instrucao4)

            'Precisa gravar no titulo o nosso numero para gerar o pdf
            tit.NossoNumero = param.SequenciaDePagamentos
            boletos.Add(b)
            param.SequenciaDePagamentos += 1
        Next
        Return boletos
    End Function

    Protected Function BoletosSicredi(ByRef pTitulos As List(Of Titulo), ByVal param As ParametrosDaCarteiraDeCobranca) As List(Of BoletoNet.Boleto)
        Dim boletos As New List(Of BoletoNet.Boleto)
        For Each tit In pTitulos
            Dim b As New BoletoNet.Boleto
            b.Banco = New BoletoNet.Banco(param.CodigoBanco)
            b.Cedente = New BoletoNet.Cedente With
            {
                .CPFCNPJ = param.CodigoEmpresa,
                .Nome = param.NomeDaEmpresaNoBanco,
                .ContaBancaria = New BoletoNet.ContaBancaria With
                {
                    .Agencia = param.CodigoAgencia,
                    .DigitoAgencia = param.DigitoAgencia,
                    .Conta = param.Conta,
                    .DigitoConta = param.DigitoConta,
                    .OperacaConta = param.SubConta
                },
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Empresa.Endereco,
                    .Numero = tit.Empresa.Numero,
                    .Complemento = tit.Empresa.Complemento,
                    .Bairro = tit.Empresa.Bairro,
                    .Cidade = tit.Empresa.Cidade,
                    .UF = tit.Empresa.Estado.Codigo,
                    .CEP = tit.Empresa.CEP
                 },
                .Codigo = param.CodigoDaEmpresaNoBanco,
                .Convenio = param.Convenio,
                .CodigoTransmissao = param.SequenciaDeRemessa
            }
            b.Sacado = New BoletoNet.Sacado With
            {
                .CPFCNPJ = tit.Cliente.Codigo,
                .Nome = tit.Cliente.Nome,
                .Endereco = New BoletoNet.Endereco With
                {
                    .[End] = tit.Cliente.Endereco,
                    .Numero = tit.Cliente.Numero,
                    .Complemento = tit.Cliente.Complemento,
                    .Bairro = tit.Cliente.Bairro,
                    .Cidade = tit.Cliente.Cidade,
                    .UF = tit.Cliente.Estado.Codigo,
                    .CEP = tit.Cliente.CEP
                }
            }

            b.NumeroDocumento = tit.Codigo.ToString()
            b.NossoNumero = IIf(tit.NossoNumero = "0", param.SequenciaDePagamentos, tit.NossoNumero)
            b.DataDocumento = tit.UsuarioInclusaoData
            b.DataVencimento = tit.Prorrogacao
            b.ValorBoleto = tit.ValorLiquido
            b.Aceite = "N"
            b.Carteira = param.CarteiraPadrao

            Dim instrucao1 = New BoletoNet.Instrucao(param.CodigoBanco)
            instrucao1.Descricao = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."
            b.Instrucoes.Add(instrucao1)

            If (param.JurosMes > 0) Then
                b.DataMulta = tit.Prorrogacao.AddDays(1)
                b.PercMulta = param.JurosMes
                b.ValorMulta = CDec(b.ValorBoleto * b.PercMulta / 100)

                Dim instrucao2 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao2.Descricao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & b.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."
                b.Instrucoes.Add(instrucao2)

            End If
            'Mora dia
            If (param.MoraDiaria > 0) Then
                b.JurosMora = (b.ValorBoleto * param.MoraDiaria) / 100
                b.PercJurosMora = param.MoraDiaria
                Dim instrucao3 = New BoletoNet.Instrucao(param.CodigoBanco)
                instrucao3.Descricao = "<br /> APOS VENCIMENTO COBRAR R$ " & b.JurosMora.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"
                b.Instrucoes.Add(instrucao3)

            End If
            'Precisa gravar no titulo o nosso numero para gerar o pdf
            tit.NossoNumero = param.SequenciaDePagamentos
            boletos.Add(b)
            param.SequenciaDePagamentos += 1
        Next
        Return boletos
    End Function

    Protected Function NomeNoPdf(ByRef ListaDeTitulos As List(Of Titulo)) As String
        Dim nomePDF As String = ""
        If ListaDeTitulos.Count() = 1 Then

            Dim tit = ListaDeTitulos(0)
            Dim historico As String() = tit.Historico.Split(",")
            nomePDF = historico(0).Replace("REF.", "").Replace("-1-S", "").Replace(Environment.NewLine, "")
            If Not nomePDF.Contains("AGRUPAMENTO") Then
                nomePDF += " " & historico(1).Replace("/", " de ").Replace(Environment.NewLine, "") & " " & tit.Cliente.Nome.Replace("/", "").Replace(".", "") & " - " & tit.Prorrogacao.ToString("dd-MM-yyyy")
            Else
                nomePDF += " " & tit.Cliente.Nome.Replace("/", "").Replace(".", "") & " - " & tit.Prorrogacao.ToString("dd-MM-yyyy")
            End If
        End If

        Return nomePDF
    End Function

    Protected Sub GerarBoleto(ByRef boletos As List(Of BoletoNet.Boleto), ByVal nomePDF As String, ByVal baixar As Boolean, Optional ByVal nBanco As String = "")
        Dim html = New StringBuilder()
        Dim remessa As New NGS.Lib.Negocio.BoletoBancario
        'Gerar Boletos HTML
        For Each b In boletos

            html.Append("<div style=""page-break-after: always;"">")
            Dim htmlAux As String = ""

            If b.Banco.Codigo = 748 Then
                htmlAux = remessa.SicrediPdf(b)
            ElseIf b.Banco.Codigo = 756 Then
                htmlAux = remessa.SicoobPdf(b)
            ElseIf b.Banco.Codigo = 237 AndAlso nBanco = 460 Then
                htmlAux = remessa.UnavantiPdf(b)
            Else
                htmlAux = remessa.DaycovalPdf(b)
            End If

            If nBanco = 460 Then
                htmlAux = htmlAux.Replace("@URLIMAGEMLOGO", Server.MapPath("Images/" + nBanco + ".png"))
                htmlAux = htmlAux.Replace("@URLIMGCEDENTE", Server.MapPath("Images/" + nBanco + ".png"))
            Else
                htmlAux = htmlAux.Replace("@URLIMAGEMLOGO", Server.MapPath("Images/" + b.Banco.Codigo.ToString() + ".png"))
                htmlAux = htmlAux.Replace("@URLIMGCEDENTE", Server.MapPath("Images/" + b.Banco.Codigo.ToString() + ".png"))
            End If

            html.Append(htmlAux)
            html.Append("</div>")
        Next

        Dim fileName As String = "Boletos/" + nomePDF + ".pdf"
        Dim pathPDF = Server.MapPath("Boletos/" + nomePDF + ".pdf")
        Dim pdf = New NReco.PdfGenerator.HtmlToPdfConverter().GeneratePdf(html.ToString)

        Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
            fs.Write(pdf, 0, pdf.Length)
            fs.Close()
        End Using

        'If baixar Then Download("Boletos", nomePDF & ".pdf", "application/pdf")

        If baixar Then ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), Guid.NewGuid().ToString, "window.open('" & fileName & "');", True)

    End Sub

    Private Sub GerarBoletos(ByRef Boletos As Boleto2Net.Boletos, ByRef ListaDeTitulos As List(Of Titulo), ByVal tit As Titulo)

        Dim where = "Where Ativo = 1 AND Banco_Id = " & tit.CodigoBancoPagador & " AND Agencia_Id = " & tit.CodigoAgenciaPagadora & " AND Conta_Id = " & tit.ContaPagadora
        Dim parametros As New ParametrosDaCarteiraDeCobranca(where)

        Dim empresa As New Cliente(parametros.CodigoEmpresa, parametros.EnderecoEmpresa)
        'Cabeçalho
        Boletos.Banco = Boleto2Net.Banco.Instancia(parametros.CodigoBanco)
        Boletos.Banco.Cedente = New Boleto2Net.Cedente With
        {
            .CPFCNPJ = parametros.CodigoEmpresa,
            .Nome = parametros.NomeDaEmpresaNoBanco,
            .ContaBancaria = New Boleto2Net.ContaBancaria With
            {
                .Agencia = parametros.CodigoAgencia,
                .DigitoAgencia = parametros.DigitoAgencia,
                .Conta = IIf(parametros.CodigoBanco = Boleto2Net.Bancos.Bradesco, parametros.Conta, parametros.Convenio),
                .DigitoConta = parametros.DigitoConta,
                .CarteiraPadrao = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, parametros.CarteiraPadrao.ToArray()(0).ToString(), parametros.CarteiraPadrao),
                .TipoCarteiraPadrao = TipoCarteira.CarteiraCobrancaSimples,
                .TipoFormaCadastramento = TipoFormaCadastramento.ComRegistro,
                .TipoImpressaoBoleto = TipoImpressaoBoleto.Empresa,
                .TipoDocumento = TipoDocumento.Tradicional,
                .VariacaoCarteiraPadrao = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, parametros.CarteiraPadrao.ToArray()(1).ToString(), ""),
                .OperacaoConta = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, String.Format("{0:00}", Integer.Parse(parametros.SubConta)), "")
            },
            .Endereco = New Boleto2Net.Endereco With
            {
                .LogradouroEndereco = empresa.Endereco,
                .LogradouroNumero = empresa.Numero,
                .LogradouroComplemento = empresa.Complemento,
                .Bairro = empresa.Bairro,
                .Cidade = empresa.Cidade,
                .UF = empresa.Estado.Codigo,
                .CEP = empresa.CEP
             },
            .Codigo = parametros.Convenio,
            .CodigoDV = parametros.DigitoConta,
            .CodigoTransmissao = String.Empty
        }

        Boletos.Banco.FormataCedente()

        Dim cedente = New Boleto2Net.Cedente()


        For Each Titulo In ListaDeTitulos
            'Adiciona os Títulos
            Dim Boleto As New Boleto(Boletos.Banco)
            Boleto.Sacado = New Sacado With
            {
                .CPFCNPJ = Titulo.Cliente.Codigo,
                .Nome = Titulo.Cliente.Nome,
                .Endereco = New Boleto2Net.Endereco With
                {
                    .LogradouroEndereco = Titulo.Cliente.Endereco,
                    .LogradouroNumero = Titulo.Cliente.Numero,
                    .LogradouroComplemento = Titulo.Cliente.Complemento,
                    .Bairro = Titulo.Cliente.Bairro,
                    .Cidade = Titulo.Cliente.Cidade,
                    .UF = Titulo.Cliente.Estado.Codigo,
                    .CEP = Titulo.Cliente.CEP
                }
            }

            Boleto.CodigoOcorrencia = "01"
            Boleto.DescricaoOcorrencia = "Remessa Registrar"

            Boleto.NumeroDocumento = Titulo.Codigo.ToString()
            Boleto.NumeroControleParticipante = Titulo.Codigo.ToString()
            Boleto.NossoNumero = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, String.Format("{0:00000}", (parametros.SequenciaDeRegistroDePagamento + 1)), Titulo.Codigo.ToString())

            Boleto.DataEmissao = DateTime.Now.Date

            'Furlan - 26/07/2021 - Alterado para prorrogação pois podem ter mudado a data antes de gerar os Boletos
            'Boleto.DataVencimento = Titulo.Vencimento
            Boleto.DataVencimento = Titulo.Prorrogacao

            Boleto.ValorTitulo = Titulo.ValorLiquido
            Boleto.Aceite = "N"
            Boleto.EspecieDocumento = IIf(parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756, TipoEspecieDocumento.DMI, TipoEspecieDocumento.DM)

            'Boleto.CodigoInstrucao1 = IIf(parametros.CodigoBanco = 104, "91", "05")

            Boleto.MensagemInstrucoesCaixa = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."

            If (parametros.JurosMes > 0) Then
                Boleto.DataMulta = Titulo.Prorrogacao.AddDays(1)
                Boleto.PercentualMulta = parametros.JurosMes
                Boleto.ValorMulta = CDec(Boleto.ValorTitulo * Boleto.PercentualMulta / 100)

                Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & Boleto.ValorMulta.ToString("C2", CultureInfo.CurrentCulture) & "."

                If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
                    Boleto.MensagemInstrucoesCaixa = Instrucao
                Else
                    Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
                End If
            End If

            If (parametros.MoraDiaria > 0) Then
                Boleto.DataJuros = Titulo.Prorrogacao.AddDays(1)

                'Furlan - 26-07-2021 - Química é diferente pois cobra 3,50% ao mês dividido por 30.  
                If parametros.CodigoEmpresa = "05272759000147" Then
                    Boleto.PercentualJurosDia = (parametros.MoraDiaria / 30)
                    Boleto.ValorJurosDia = Boleto.ValorTitulo * Boleto.PercentualJurosDia / 100
                Else
                    Boleto.PercentualJurosDia = parametros.MoraDiaria
                    Boleto.ValorJurosDia = (Boleto.ValorTitulo / 30) * Boleto.PercentualJurosDia / 100
                End If

                Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR R$ " & Boleto.ValorJurosDia.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO."

                If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
                    Boleto.MensagemInstrucoesCaixa = Instrucao
                Else
                    Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
                End If
            End If

            'Protesto boleto
            Boleto.CodigoProtesto = IIf(parametros.DiasParaProtesto > 0, TipoCodigoProtesto.ProtestarDiasUteis, TipoCodigoProtesto.NaoProtestar)
            If parametros.DiasParaProtesto > 0 Then
                Boleto.CodigoInstrucao1 = "06"
                Boleto.CodigoInstrucao2 = parametros.DiasParaProtesto.ToString()
                Boleto.DiasProtesto = parametros.DiasParaProtesto

                If Left(parametros.CodigoEmpresa, 8) = "40938762" OrElse Left(parametros.CodigoEmpresa, 8) = "49673784" OrElse Left(parametros.CodigoEmpresa, 8) = "05366261" OrElse Left(parametros.CodigoEmpresa, 8) = "62780383" Then
                    Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS DO VENCIMENTO."
                Else
                    Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS ÚTEIS."
                End If
            End If

            Boleto.CodigoBaixaDevolucao = IIf(parametros.CodigoBanco = 104, TipoCodigoBaixaDevolucao.BaixarDevolver, TipoCodigoBaixaDevolucao.NaoBaixarNaoDevolver)
            'Boleto.DiasBaixaDevolucao = 0

            Boleto.ValidarDados()
            Boletos.Add(Boleto)
        Next
    End Sub

    Protected Sub btnBoleto_Click(ByVal sender As Object, ByVal e As CommandEventArgs)
        Try
            Dim ListaDeTitulos = New List(Of Titulo)
            Dim Boletos = New Boleto2Net.Boletos()
            Dim tit = New Titulo(CInt(e.CommandName))
            ListaDeTitulos.Add(tit)

            Dim where = "Where Ativo = 1 AND Banco_Id = " & tit.CodigoBancoPagador & " AND Agencia_Id = " & tit.CodigoAgenciaPagadora & " AND Conta_Id = " & tit.ContaPagadora
            Dim parametros As New ParametrosDaCarteiraDeCobranca(where)

            If tit.CodigoBancoPagador = 707 Then
                Dim daycovalBoletos = BoletosDaycoval(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoleto(daycovalBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 748 Then
                Dim sicredilBoletos = BoletosSicredi(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoleto(sicredilBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 756 Then
                Dim sicredilBoletos = BoletosSicoob(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoleto(sicredilBoletos, _nomePdf, True)
                Exit Sub
            End If

            If tit.CodigoBancoPagador = 460 Then
                Dim unavantiBoletos = BoletosUnavanti(ListaDeTitulos, parametros)
                Dim _nomePdf = NomeNoPdf(ListaDeTitulos)
                GerarBoleto(unavantiBoletos, _nomePdf, True, 460)
                Exit Sub
            End If

            GerarBoletos(Boletos, ListaDeTitulos, tit)

            Dim Titulo = ListaDeTitulos.FirstOrDefault(Function(x) x.Codigo = CInt(e.CommandName))
            Dim nomePDF As String = ""

            'Gerar Boletos PDF
            For Each b In Boletos
                Dim BoletoBancario = New Boleto2Net.BoletoBancario With {.Boleto = b}
                Dim pdf = BoletoBancario.MontaBytesPDF(False)
                nomePDF = NomeNoPdf(ListaDeTitulos)

                Dim pathPDF = Server.MapPath("Boletos/" & nomePDF & ".pdf")
                File.WriteAllBytes(pathPDF, pdf)
            Next

            Download("Boletos", nomePDF & ".pdf", "application/pdf")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub chkAllTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If gridBoletoBancario.Rows.Count > 0 Then
                Dim chkAllTitulos As CheckBox = CType(sender, CheckBox)

                For Each rowgrid As GridViewRow In gridBoletoBancario.Rows
                    Dim chkTitulo As CheckBox = CType(rowgrid.FindControl("chkImpBoletos"), CheckBox)

                    If chkAllTitulos.Checked Then
                        chkTitulo.Checked = True
                    Else
                        chkTitulo.Checked = False
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ImpressaoDeBoletos()
        Try
            Dim numTitulo As Integer = 0
            For Each row As GridViewRow In gridBoletoBancario.Rows
                Dim chk As CheckBox = row.Cells(0).FindControl("chkImpBoletos")
                If chk.Checked Then
                    numTitulo += 1
                End If
            Next

            If numTitulo < 2 Then
                MsgBox(Me.Page, "Para impressão em bloco deve ser selecionado no mínimo 2 registros.", eTitulo.Info)
                Exit Sub
            End If

            Dim strConta As String() = ddlBanco.SelectedValue.Split(";")
            Dim where = "Where Ativo = 1 AND Banco_Id = " & strConta(0) & " AND Agencia_Id = " & strConta(1) & " AND Conta_Id = " & strConta(3) & " AND Convenio = " & strConta(5)
            Dim Boletos = New Boleto2Net.Boletos()

            Dim parametros As New ParametrosDaCarteiraDeCobranca(where)
            Dim empresa As New Cliente(parametros.CodigoEmpresa, parametros.EnderecoEmpresa)

            'Daycoval 707 - Sicredi 748 - SICOOB 756
            If parametros.CodigoBanco = 707 Or parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756 Or parametros.CodigoBanco = 274 Or parametros.CodigoBanco = 460 Then
                Dim listaDeTitulos As New List(Of Titulo)
                For Each row As GridViewRow In gridBoletoBancario.Rows
                    Dim chk As CheckBox = row.Cells(0).FindControl("chkImpBoletos")
                    If chk.Checked Then
                        Dim tit As New Titulo(CInt(CType(row.FindControl("hpTitulo"), HyperLink).Text))
                        listaDeTitulos.Add(tit)
                    End If
                Next

                If parametros.CodigoBanco = 707 Then
                    Dim daycovalBoletos = BoletosDaycoval(listaDeTitulos, parametros)
                    GerarBoleto(daycovalBoletos, "Boletos", True)
                    Exit Sub
                ElseIf parametros.CodigoBanco = 756 Then
                    Dim sicoobBoletos = BoletosSicoob(listaDeTitulos, parametros)
                    GerarBoleto(sicoobBoletos, "Boletos", True)
                    Exit Sub
                ElseIf parametros.CodigoBanco = 460 Then
                    Dim unavantiBoletos = BoletosUnavanti(listaDeTitulos, parametros)
                    GerarBoleto(unavantiBoletos, "Boletos", True, "460")
                    Exit Sub
                Else
                    'Sicredi
                    Dim sicrediBoletos = BoletosSicredi(listaDeTitulos, parametros)
                    GerarBoleto(sicrediBoletos, "Boletos", True)
                    Exit Sub
                End If

            End If

            'Cabeçalho
            Boletos.Banco = Boleto2Net.Banco.Instancia(parametros.CodigoBanco)
            Boletos.Banco.Cedente = New Boleto2Net.Cedente With
            {
                .CPFCNPJ = parametros.CodigoEmpresa,
            .Nome = parametros.NomeDaEmpresaNoBanco,
            .ContaBancaria = New Boleto2Net.ContaBancaria With
            {
                .Agencia = parametros.CodigoAgencia,
                .DigitoAgencia = parametros.DigitoAgencia,
                .Conta = IIf(parametros.CodigoBanco = Boleto2Net.Bancos.Bradesco, parametros.Conta, parametros.Convenio),
                .DigitoConta = parametros.DigitoConta,
                .CarteiraPadrao = parametros.CarteiraPadrao,
                .TipoCarteiraPadrao = TipoCarteira.CarteiraCobrancaSimples,
                .TipoFormaCadastramento = TipoFormaCadastramento.ComRegistro,
                .TipoImpressaoBoleto = TipoImpressaoBoleto.Empresa,
                .TipoDocumento = TipoDocumento.Tradicional
            },
            .Endereco = New Boleto2Net.Endereco With
            {
                .LogradouroEndereco = empresa.Endereco,
                .LogradouroNumero = empresa.Numero,
                .LogradouroComplemento = empresa.Complemento,
                .Bairro = empresa.Bairro,
                .Cidade = empresa.Cidade,
                .UF = empresa.Estado.Codigo,
                .CEP = empresa.CEP
             },
            .Codigo = parametros.Convenio,
            .CodigoDV = parametros.DigitoConta,
            .CodigoTransmissao = String.Empty
            }

            Boletos.Banco.FormataCedente()

            For Each row As GridViewRow In gridBoletoBancario.Rows
                Dim chk As CheckBox = row.Cells(0).FindControl("chkImpBoletos")
                If chk.Checked Then

                    Dim Titulo As New Titulo(CInt(CType(row.FindControl("hpTitulo"), HyperLink).Text))

                    'Adiciona os Títulos
                    Dim Boleto As New Boleto(Boletos.Banco)
                    Boleto.Sacado = New Sacado With
                    {
                        .CPFCNPJ = Titulo.Cliente.Codigo,
                        .Nome = Titulo.Cliente.Nome,
                        .Endereco = New Boleto2Net.Endereco With
                        {
                            .LogradouroEndereco = Titulo.Cliente.Endereco,
                            .LogradouroNumero = Titulo.Cliente.Numero,
                            .LogradouroComplemento = Titulo.Cliente.Complemento,
                            .Bairro = Titulo.Cliente.Bairro,
                            .Cidade = Titulo.Cliente.Cidade,
                            .UF = Titulo.Cliente.Estado.Codigo,
                            .CEP = Titulo.Cliente.CEP
                        }
                    }

                    Boleto.CodigoOcorrencia = "01"
                    Boleto.DescricaoOcorrencia = "Remessa Registrar"

                    Boleto.NumeroDocumento = Titulo.Codigo.ToString()
                    Boleto.NumeroControleParticipante = Titulo.Codigo.ToString()
                    Boleto.NossoNumero = Titulo.Codigo.ToString()

                    Boleto.DataEmissao = DateTime.Now.Date

                    'Furlan - 26/07/2021 - Alterado para prorrogação pois podem ter mudado a data antes de gerar os Boletos
                    'Boleto.DataVencimento = Titulo.Vencimento
                    Boleto.DataVencimento = Titulo.Prorrogacao

                    Boleto.ValorTitulo = Titulo.ValorLiquido
                    Boleto.Aceite = "N"
                    Boleto.EspecieDocumento = TipoEspecieDocumento.DM
                    Boleto.MensagemInstrucoesCaixa = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."

                    If (parametros.JurosMes > 0) Then
                        Boleto.DataMulta = Titulo.Prorrogacao.AddDays(1)
                        Boleto.PercentualMulta = parametros.JurosMes
                        Boleto.ValorMulta = CDec(Boleto.ValorTitulo * Boleto.PercentualMulta / 100)

                        Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & Boleto.ValorMulta.ToString("C2", CultureInfo.CurrentCulture)

                        If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
                            Boleto.MensagemInstrucoesCaixa = Instrucao
                        Else
                            Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
                        End If
                    End If

                    If (parametros.MoraDiaria > 0) Then
                        Boleto.DataJuros = Titulo.Prorrogacao.AddDays(1)

                        'Furlan - 26-07-2021 - Química é diferente pois cobra 3,50% ao mês dividido por 30. 
                        If parametros.CodigoEmpresa = "05272759000147" Then
                            Boleto.PercentualJurosDia = (parametros.MoraDiaria / 30)
                            Boleto.ValorJurosDia = Boleto.ValorTitulo * Boleto.PercentualJurosDia / 100
                        Else
                            Boleto.PercentualJurosDia = parametros.MoraDiaria
                            Boleto.ValorJurosDia = (Boleto.ValorTitulo / 30) * Boleto.PercentualJurosDia / 100
                        End If

                        Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR R$ " & Boleto.ValorJurosDia.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"

                        If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
                            Boleto.MensagemInstrucoesCaixa = Instrucao
                        Else
                            Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
                        End If
                    End If

                    'Protesto boleto
                    Boleto.CodigoProtesto = IIf(parametros.DiasParaProtesto > 0, TipoCodigoProtesto.ProtestarDiasUteis, TipoCodigoProtesto.NaoProtestar)
                    If parametros.DiasParaProtesto > 0 Then
                        Boleto.CodigoInstrucao1 = "06"
                        Boleto.CodigoInstrucao2 = parametros.DiasParaProtesto.ToString()
                        Boleto.DiasProtesto = parametros.DiasParaProtesto

                        If Left(parametros.CodigoEmpresa, 8) = "40938762" OrElse Left(parametros.CodigoEmpresa, 8) = "49673784" OrElse Left(parametros.CodigoEmpresa, 8) = "05366261" OrElse Left(parametros.CodigoEmpresa, 8) = "62780383" Then
                            Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS DO VENCIMENTO."
                        Else
                            Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS ÚTEIS."
                        End If
                    End If

                    Boleto.CodigoBaixaDevolucao = IIf(parametros.CodigoBanco = 104, TipoCodigoBaixaDevolucao.BaixarDevolver, TipoCodigoBaixaDevolucao.NaoBaixarNaoDevolver)
                    'Boleto.DiasBaixaDevolucao = 0

                    Boleto.ValidarDados()
                    Boletos.Add(Boleto)
                End If
            Next

            Dim html = New StringBuilder()

            'Gerar Boletos HTML
            For Each b In Boletos
                Using BoletoBancario As Boleto2Net.BoletoBancario = New Boleto2Net.BoletoBancario With {.Boleto = b, .OcultarInstrucoes = False, .MostrarComprovanteEntrega = False, .MostrarEnderecoCedente = True}
                    html.Append("<div style=""page-break-after: always;"">")
                    html.Append(BoletoBancario.MontaHtml())
                    html.Append("</div>")
                End Using
            Next

            Dim fileName As String = "Boletos/Boletos.pdf"
            Dim pathPDF = Server.MapPath("Boletos/Boletos.pdf")

            Dim pdf = New NReco.PdfGenerator.HtmlToPdfConverter().GeneratePdf(html.ToString)

            Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
                fs.Write(pdf, 0, pdf.Length)
                fs.Close()
            End Using

            ' Download("Boletos", "Boletos.pdf", "application/pdf")

            ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), Guid.NewGuid().ToString, "window.open('" & fileName & "');", True)

        Catch ex As Exception

        End Try
    End Sub

    'Protected Sub btnBoletos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        Dim t As New Titulo
    '        For Each row As GridViewRow In gridBoletoBancario.Rows
    '            Dim chk As CheckBox = row.Cells(0).FindControl("chkImpBoletos")
    '            If chk.Checked Then
    '                t = New Titulo(CInt(CType(row.FindControl("hpTitulo"), HyperLink).Text))
    '            End If
    '        Next

    '        Dim where = "Where Ativo = 1 AND Banco_Id = " & t.CodigoBancoPagador & " AND Agencia_Id = " & t.CodigoAgenciaPagadora & " AND Conta_Id = " & t.ContaPagadora
    '        Dim Boletos = New Boleto2Net.Boletos()
    '        Dim parametros As New ParametrosDaCarteiraDeCobranca(where)
    '        Dim empresa As New Cliente(parametros.CodigoEmpresa, parametros.EnderecoEmpresa)

    '        'Daycoval 707 - Sicredi 748 - SICOOB 756
    '        If parametros.CodigoBanco = 707 Or parametros.CodigoBanco = 748 Or parametros.CodigoBanco = 756 Or parametros.CodigoBanco = 460 Then
    '            Dim listaDeTitulos As New List(Of Titulo)
    '            For Each row As GridViewRow In gridBoletoBancario.Rows
    '                Dim chk As CheckBox = row.Cells(0).FindControl("chkImpBoletos")
    '                If chk.Checked Then
    '                    Dim tit As New Titulo(CInt(CType(row.FindControl("hpTitulo"), HyperLink).Text))
    '                    listaDeTitulos.Add(tit)
    '                End If
    '            Next

    '            If listaDeTitulos.Count < 2 Then
    '                MsgBox(Me.Page, "Para impressão em bloco deve ser selecionado no mínimo 2 registros.", eTitulo.Info)
    '                Exit Sub
    '            End If
    '            If parametros.CodigoBanco = 707 Then
    '                Dim daycovalBoletos = BoletosDaycoval(listaDeTitulos, parametros)
    '                GerarBoleto(daycovalBoletos, "Boletos", True)
    '                Exit Sub
    '            ElseIf parametros.CodigoBanco = 756 Then
    '                Dim sicoobBoletos = BoletosSicoob(listaDeTitulos, parametros)
    '                GerarBoleto(sicoobBoletos, "Boletos", True)
    '                Exit Sub
    '            ElseIf parametros.CodigoBanco = 460 Then
    '                Dim unavantiBoletos = BoletosUnavanti(listaDeTitulos, parametros)
    '                GerarBoleto(unavantiBoletos, "Boletos", True)
    '                Exit Sub
    '            Else
    '                'Sicredi
    '                Dim sicrediBoletos = BoletosSicredi(listaDeTitulos, parametros)
    '                GerarBoleto(sicrediBoletos, "Boletos", True)
    '                Exit Sub
    '            End If
    '        End If

    '        'Cabeçalho
    '        Boletos.Banco = Boleto2Net.Banco.Instancia(parametros.CodigoBanco)
    '        Boletos.Banco.Cedente = New Boleto2Net.Cedente With
    '        {
    '            .CPFCNPJ = parametros.CodigoEmpresa,
    '            .Nome = parametros.NomeDaEmpresaNoBanco,
    '            .ContaBancaria = New Boleto2Net.ContaBancaria With
    '        {
    '            .Agencia = parametros.CodigoAgencia,
    '            .DigitoAgencia = parametros.DigitoAgencia,
    '            .Conta = IIf(parametros.CodigoBanco = Boleto2Net.Bancos.Bradesco, parametros.Conta, parametros.Convenio),
    '            .DigitoConta = parametros.DigitoConta,
    '            .CarteiraPadrao = parametros.CarteiraPadrao,
    '            .TipoCarteiraPadrao = TipoCarteira.CarteiraCobrancaSimples,
    '            .TipoFormaCadastramento = TipoFormaCadastramento.ComRegistro,
    '            .TipoImpressaoBoleto = TipoImpressaoBoleto.Empresa,
    '            .TipoDocumento = TipoDocumento.Tradicional
    '        },
    '            .Endereco = New Boleto2Net.Endereco With
    '        {
    '            .LogradouroEndereco = empresa.Endereco,
    '            .LogradouroNumero = empresa.Numero,
    '            .LogradouroComplemento = empresa.Complemento,
    '            .Bairro = empresa.Bairro,
    '            .Cidade = empresa.Cidade,
    '            .UF = empresa.Estado.Codigo,
    '            .CEP = empresa.CEP
    '         },
    '            .Codigo = parametros.Convenio,
    '            .CodigoDV = parametros.DigitoConta,
    '            .CodigoTransmissao = String.Empty
    '        }

    '        Boletos.Banco.FormataCedente()

    '        For Each row As GridViewRow In gridBoletoBancario.Rows
    '            Dim chk As CheckBox = row.Cells(0).FindControl("chkImpBoletos")
    '            If chk.Checked Then

    '                Dim Titulo As New Titulo(CInt(CType(row.FindControl("hpTitulo"), HyperLink).Text))

    '                'Adiciona os Títulos
    '                Dim Boleto As New Boleto(Boletos.Banco)
    '                Boleto.Sacado = New Sacado With
    '                {
    '                    .CPFCNPJ = Titulo.Cliente.Codigo,
    '                    .Nome = Titulo.Cliente.Nome,
    '                    .Endereco = New Boleto2Net.Endereco With
    '                    {
    '                        .LogradouroEndereco = Titulo.Cliente.Endereco,
    '                        .LogradouroNumero = Titulo.Cliente.Numero,
    '                        .LogradouroComplemento = Titulo.Cliente.Complemento,
    '                        .Bairro = Titulo.Cliente.Bairro,
    '                        .Cidade = Titulo.Cliente.Cidade,
    '                        .UF = Titulo.Cliente.Estado.Codigo,
    '                        .CEP = Titulo.Cliente.CEP
    '                    }
    '                }

    '                Boleto.CodigoOcorrencia = "01"
    '                Boleto.DescricaoOcorrencia = "Remessa Registrar"
    '                Boleto.NumeroDocumento = Titulo.Codigo.ToString()
    '                Boleto.NumeroControleParticipante = Titulo.Codigo.ToString()
    '                Boleto.NossoNumero = Titulo.Codigo.ToString()
    '                Boleto.DataEmissao = DateTime.Now.Date
    '                Boleto.DataVencimento = Titulo.Prorrogacao
    '                Boleto.ValorTitulo = Titulo.ValorLiquido
    '                Boleto.Aceite = "N"
    '                Boleto.EspecieDocumento = TipoEspecieDocumento.DM
    '                Boleto.MensagemInstrucoesCaixa = "Instruções de responsabilidade do BENEFICIÁRIO. Qualquer dúvida sobre este boleto, contate o BENEFICIÁRIO."

    '                If (parametros.JurosMes > 0) Then
    '                    Boleto.DataMulta = Titulo.Prorrogacao.AddDays(1)
    '                    Boleto.PercentualMulta = parametros.JurosMes
    '                    Boleto.ValorMulta = CDec(Boleto.ValorTitulo * Boleto.PercentualMulta / 100)

    '                    Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR MULTA DE R$ " & Boleto.ValorMulta.ToString("C2", CultureInfo.CurrentCulture)

    '                    If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
    '                        Boleto.MensagemInstrucoesCaixa = Instrucao
    '                    Else
    '                        Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
    '                    End If
    '                End If

    '                If (parametros.MoraDiaria > 0) Then
    '                    Boleto.DataJuros = Titulo.Prorrogacao.AddDays(1)

    '                    'Furlan - 26-07-2021 - "05272759000147" Química é diferente pois cobra 3,50% ao mês dividido por 30. 
    '                    If parametros.CodigoEmpresa = "05272759000147" Then
    '                        Boleto.PercentualJurosDia = (parametros.MoraDiaria / 30)
    '                        Boleto.ValorJurosDia = Boleto.ValorTitulo * Boleto.PercentualJurosDia / 100
    '                    Else
    '                        Boleto.PercentualJurosDia = parametros.MoraDiaria
    '                        Boleto.ValorJurosDia = (Boleto.ValorTitulo / 30) * Boleto.PercentualJurosDia / 100
    '                    End If

    '                    Dim Instrucao = "<br /> APOS VENCIMENTO COBRAR R$ " & Boleto.ValorJurosDia.ToString("C2", CultureInfo.CurrentCulture) & " POR DIA DE ATRASO"

    '                    If (String.IsNullOrEmpty(Boleto.MensagemInstrucoesCaixa)) Then
    '                        Boleto.MensagemInstrucoesCaixa = Instrucao
    '                    Else
    '                        Boleto.MensagemInstrucoesCaixa += Environment.NewLine + Instrucao
    '                    End If
    '                End If

    '                'Protesto boleto
    '                Boleto.CodigoProtesto = IIf(parametros.DiasParaProtesto > 0, TipoCodigoProtesto.ProtestarDiasUteis, TipoCodigoProtesto.NaoProtestar)
    '                If parametros.DiasParaProtesto > 0 Then
    '                    Boleto.CodigoInstrucao1 = "06"
    '                    Boleto.CodigoInstrucao2 = parametros.DiasParaProtesto.ToString()
    '                    Boleto.DiasProtesto = parametros.DiasParaProtesto

    '                    If Left(parametros.CodigoEmpresa, 8) = "40938762" OrElse Left(parametros.CodigoEmpresa, 8) = "49673784" OrElse Left(parametros.CodigoEmpresa, 8) = "05366261" Then
    '                        Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS DO VENCIMENTO."
    '                    Else
    '                        Boleto.MensagemInstrucoesCaixa += Environment.NewLine + "<br /> PROTESTAR APÓS " + parametros.DiasParaProtesto.ToString() + " DIAS ÚTEIS."
    '                    End If
    '                End If

    '                Boleto.CodigoBaixaDevolucao = IIf(parametros.CodigoBanco = 104, TipoCodigoBaixaDevolucao.BaixarDevolver, TipoCodigoBaixaDevolucao.NaoBaixarNaoDevolver)
    '                Boleto.ValidarDados()
    '                Boletos.Add(Boleto)
    '            End If
    '        Next

    '        If Boletos.Count < 2 Then
    '            MsgBox(Me.Page, "Para impressão em bloco deve ser selecionado no mínimo 2 registros.", eTitulo.Info)
    '            Exit Sub
    '        End If

    '        Dim html = New StringBuilder()

    '        'Gerar Boletos HTML
    '        For Each b In Boletos
    '            Using BoletoBancario As Boleto2Net.BoletoBancario = New Boleto2Net.BoletoBancario With {.Boleto = b, .OcultarInstrucoes = False, .MostrarComprovanteEntrega = False, .MostrarEnderecoCedente = True}
    '                html.Append("<div style=""page-break-after: always;"">")
    '                html.Append(BoletoBancario.MontaHtml())
    '                html.Append("</div>")
    '            End Using
    '        Next

    '        Dim pathPDF = Server.MapPath("Boletos/Boletos.pdf")
    '        Dim pdf = New NReco.PdfGenerator.HtmlToPdfConverter().GeneratePdf(html.ToString)

    '        Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
    '            fs.Write(pdf, 0, pdf.Length)
    '            fs.Close()
    '        End Using

    '        Download("Boletos", "Boletos.pdf", "application/pdf")

    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message)
    '    End Try
    'End Sub

    Public Sub CarregarTitulosEnviadosAoBanco()

        If ddlBanco.SelectedValue.Length = 0 Then
            MsgBox(Me.Page, "Banco não foi selecionado.", eTitulo.Info)
            Exit Sub
        End If

        Dim empresa = ddlEmpresa.SelectedValue.Split("-")

        Dim where = "UnidadeDeNegocio = " & ddlUnidadeNegocio.SelectedValue & " AND Empresa = " & empresa(0) & " AND EndEmpresa = " & empresa(1)

        If txtClientes.Text.Length > 0 Then
            Dim cliente = txtCodigoCliente.Value.ToString.Split("-")
            where += " AND Cliente = '" & cliente(0) & "' AND EndCliente = '" & cliente(1) & "'"
        End If
        If txtNotaFiscal.Text.Length > 0 Then
            where += " AND Historico like 'REF. NF " & txtNotaFiscal.Text & "-%'"
        End If
        If chkPeriodo.Checked = True Then
            If txtDataInicial.Text.Length > 0 AndAlso txtDataFinal.Text.Length > 0 Then
                If radMovimento.Checked = True Then
                    where += "   AND Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' "
                Else
                    where += "   AND Prorrogacao BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' "
                End If
            Else
                MsgBox(Me.Page, "Período não foi preenchido")
                Limpar()
                Exit Sub
            End If
        End If
        where += " AND BoletoBancario = 1"

        Dim strConta As String() = ddlBanco.SelectedValue.Split(";")
        where += " AND BancoPagador = " & strConta(0)

        Dim titulos As New ListTitulo(where)

        If titulos.Count > 0 Then

            If titulos.Count > 1 Then
                lnkImpressao.Parent.Visible = True
            Else
                lnkImpressao.Parent.Visible = False
            End If

            For Each tit In titulos
                If tit.SituacaoBancaria = CInt(eSituacaoBancaria.AguardandoArquivoDeRetorno) Then
                    tit.Observacoes = CInt(eSituacaoBancaria.AguardandoArquivoDeRetorno).ToString("00") & " - " & Textos.GetEnumDescription(eSituacaoBancaria.AguardandoArquivoDeRetorno)
                End If
            Next

            gridBoletoBancario.DataSource = From tit In titulos
                                            Select New With {.Codigo = tit.Codigo, .Cidade = tit.Empresa.Cidade, .Prorrogacao = tit.Prorrogacao, .Observacoes = tit.Observacoes, .Historico = tit.Historico, .ValorLiquido = tit.ValorLiquido, .Juros = tit.Juros, .CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)}
            gridBoletoBancario.DataBind()
        Else
            MsgBox(Me.Page, "Registro(s) não encontrado(s) para a seleção.")
            Limpar()
        End If
    End Sub

    Protected Sub lnkImpressao_Click(sender As Object, e As EventArgs) Handles lnkImpressao.Click
        Try
            If Funcoes.VerificaPermissao("ReemissaoBoletoBancario", "LEITURA") Then
                ImpressaoDeBoletos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir registro!")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("ReemissaoBoletoBancario", "LEITURA") Then
                CarregarTitulosEnviadosAoBanco()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class