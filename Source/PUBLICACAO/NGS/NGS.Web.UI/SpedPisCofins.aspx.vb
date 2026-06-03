Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class SpedPisCofins
    Inherits BasePage

#Region "Variáveis"

    Dim Sql As String = ""
    Dim Sqla As String = ""
    Dim Row As DataRow

    Dim SqlArray As New ArrayList
    Dim Empresa() As String
    Dim Cliente() As String
    Dim SequenciaLote As Integer = 0

    Dim Condicao As String = ""
    Dim Codigo As String = ""
    Dim Descricao As String = ""
    Dim Crc As String = ""

    Dim Endereco As Integer = 0
    Dim Mensagem As String = ""

    Dim PeriodoInicial As String
    Dim PeriodoFinal As String

    Dim IcmsARecolher As Decimal

    Dim Processo As Integer = 0
    Dim Livro As Integer = 0
    Dim Folha As Integer = 0
    Dim Opcao As String = ""

    Dim EmpresaNome As String = ""
    Dim EmpresaCNPJ As String = ""
    Dim EmpresaInscricao As String = ""
    Dim EmpresaMestre As String

    Dim xPop As String = ""
    Dim yPop As String = ""

    Dim sqlAux As String = ""
    Dim ds As New DataSet
    Dim dsaux As New DataSet

    Dim dtInicialOriginal As Date
    Dim dtFinalOriginal As Date
    Dim Conv As String = ""
    Dim NatOp As String = ""
    Dim ArqMag As String = ""

    Dim ArquivoAux As New ArrayList
    Dim ArquivoPadrao As New ArrayList

    Dim linha As String
    Dim IntSequencia As Integer
    Dim RegistroGeral As Integer = 0    'Registro Geral

    Dim ReceitaTotal As Double
    Dim Trib_MI As Double
    Dim Trib_ME As Double
    Dim Trib_NT As Double
    Dim Percentual As Double

    Dim Registro0000 As Integer      'Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária
    Dim Registro0001 As Integer      'Abertura do Bloco 0
    Dim Registro0100 As Integer      'Dados do Contabilista
    Dim Registro0110 As Integer      'Regimes de apuracao da contribuiçao social e de apropriaçao de credito
    Dim Registro0111 As Integer      'Tabela de receita bruta mensal para fins de rateio de creditos comuns
    Dim Registro0140 As Integer      'Tabela de cadastrato de estabelecimento
    Dim Registro0145 As Integer      'Regime de Apuração da Contribuição Prevedenciaria sobre a Receita Bruta
    Dim Registro0150 As Integer      'Tabela de Cadastro do Participante
    Dim Registro0190 As Integer      'Identificao das unidades de medida
    Dim Registro0200 As Integer      'Tabela de Identificao do Item (Produtos e Servicos)
    Dim Registro0400 As Integer      'Tabela de Natureza da Operação
    Dim Registro0450 As Integer      'Tabela de Informação Complementar do documento fiscal
    Dim Registro0500 As Integer      'Tabela de Plano de contas
    Dim Registro0600 As Integer      'Tabela de Centro de Custos

    Dim Registro0007 As Integer      'Outras Inscrições Cadstrais do empresário ou Sociedade Empresária
    Dim Registro0020 As Integer      'Escrituração Contábil Descentralizada
    Dim Registro0180 As Integer      'Identificação do Relacionamento com o Participante
    Dim Registro0990 As Integer      'Encerramento do Bloco 0

    Dim RegistroA001 As Integer      'Abertura do Bloco A
    Dim RegistroA010 As Integer      'Identificação do Estabelecimento
    Dim RegistroA100 As Integer      'Documento - Nota Fiscal De Serviço
    Dim RegistroA170 As Integer      'Complemento do Documento = Itens do Documento
    Dim RegistroA990 As Integer      'Fechamento bloco A

    Dim RegistroC001 As Integer      'Abertura Bloco C
    Dim RegistroC010 As Integer      'Identificacao estabelecimento bloco C
    Dim RegistroC100 As Integer      'Documento - Nota Fiscal Bloco C
    Dim RegistroC170 As Integer      'Itens do Documento Bloco C
    Dim RegistroC180 As Integer      'Consolidação de Notas Fiscais Eletronicas de Saidas
    Dim RegistroC181 As Integer      'Detalhamento da consolidação - Operacacoes de vendas pis/pasep
    Dim RegistroC185 As Integer      'Detalhamento da consolidação - Operacacoes de vendas cofins
    Dim RegistroC190 As Integer      'Consolidado - Operacacoes de devolucao de compras e vendas
    Dim RegistroC191 As Integer      'Consolidado - Operacacoes de devolucao de compras e vendas
    Dim RegistroC195 As Integer      'Consolidado - Operacacoes de devolucao de compras e vendas
    Dim RegistroC500 As Integer      'Notas Energia Eletrica
    Dim RegistroC501 As Integer      'Notas Energia Eletrica - Complemento da Operacao Pis
    Dim RegistroC505 As Integer      'Notas Energia Eletrica - Complemento da Operacao Cofins
    Dim RegistroC990 As Integer      'Encerramento do Bloco C

    Dim RegistroD001 As Integer      'Abertura Bloco D
    Dim RegistroD010 As Integer      'Identificacao estabelecimento bloco D
    Dim RegistroD100 As Integer      'Aquisição de Serviços de Transporte
    Dim RegistroD101 As Integer      'Complemento de documento de transporte Pis
    Dim RegistroD105 As Integer      'Complemento de documento de transporte Cofins

    Dim RegistroD200 As Integer      'Receita de Serviços de Transporte
    Dim RegistroD201 As Integer      'Pis
    Dim RegistroD205 As Integer      'Cofins

    Dim RegistroD500 As Integer      'Nota Fiscal de Serviço e Comunicação
    Dim RegistroD501 As Integer      'Nota Fiscal de Serviço e Comunicação Complemento da operacao Pis
    Dim RegistroD505 As Integer      'Nota Fiscal de Serviço e Comunicação Complemento da operacao Cofins
    Dim RegistroD990 As Integer      'Encerramento bloco D

    Dim RegistroF001 As Integer      'Abertura Bloco F
    Dim RegistroF010 As Integer      'Identificacao estabelecimento bloco F
    Dim RegistroF100 As Integer      'Demais documentos e Operaçao geradora creditos
    Dim RegistroF120 As Integer      'Depreciações do Ativo Imobilizado
    Dim RegistroF130 As Integer      'Aquisições do ativo Imobilizado
    Dim RegistroF550 As Integer      'Contribuição Retida na Fonte

    Dim RegistroF600 As Integer      'Contribuição Retida na Fonte

    Dim RegistroF990 As Integer      'Encerramento do Bloco F

    Dim RegistroM001 As Integer      'Abertura Bloco M
    Dim RegistroM100 As Integer      'Credito de Pis Relativo ao Periodo
    Dim RegistroM105 As Integer      'Detalhamento da Base de Calculo do Crédito Apurado no Periodo
    Dim RegistroM110 As Integer      'Ajuste do Crédito de Pis/Apurado

    Dim RegistroM200 As Integer      'Credito de Pis Relativo ao Periodo
    Dim RegistroM205 As Integer      'Credito de Pis Relativo ao Periodo

    Dim RegistroM210 As Integer      'Detalhamento da Base de Calculo do Crédito Apurado no Periodo
    Dim RegistroM220 As Integer      'Ajustes de Debitos/Creditos de Pis

    Dim RegistroM400 As Integer      'Credito de Pis Relativo ao Periodo
    Dim RegistroM410 As Integer      'Detalhamento da Base de Calculo do Crédito Apurado no Periodo

    Dim RegistroM500 As Integer      'Credito de Pis Relativo ao Periodo
    Dim RegistroM505 As Integer      'Detalhamento da Base de Calculo do Crédito Apurado no Periodo
    Dim RegistroM510 As Integer      'Ajuste do Créditop do Cofins Apurado

    Dim RegistroM600 As Integer      'Credito de Pis Relativo ao Periodo
    Dim RegistroM605 As Integer      'Credito de Pis Relativo ao Periodo

    Dim RegistroM610 As Integer      'Detalhamento da Base de Calculo do Crédito Apurado no Periodo
    Dim RegistroM620 As Integer      'Ajustes de Débitos/ Créditos de Cofins
    Dim RegistroM800 As Integer      'Credito de Pis Relativo ao Periodo
    Dim RegistroM810 As Integer      'Detalhamento da Base de Calculo do Crédito Apurado no Periodo

    Dim RegistroM990 As Integer      'Encerramento do Bloco M

    Dim RegistroP001 As Integer
    Dim RegistroP010 As Integer
    Dim REgistroP100 As Integer
    Dim RegistroP200 As Integer
    Dim RegistroP990 As Integer

    Dim Registro1001 As Integer      'Abertura Bloco 1
    Dim Registro1100 As Integer      'Controle de Saldos Pis
    Dim Registro1300 As Integer      'Controle de Saldo Pis Retido na Fonte

    Dim Registro1500 As Integer      'Controle de saldo Cofins
    Dim Registro1700 As Integer      'Controle de Saldo Cofins Retido na Fonte

    Dim Registro1900 As Integer      'Encerramento Bloco 1

    Dim Registro1990 As Integer      'Encerramento Bloco 1

    Dim Registro9001 As Integer      'Abertura do Bloco 9
    Dim Registro9900 As Integer      'Registros do Arquivo
    Dim Registro9990 As Integer      'Encerramento do Bloco 9
    Dim Registro9999 As Integer      'Encerramento do Arquivo Digital

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Fiscal)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Fiscal.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("SpedPisCofins", "ACESSAR") Then
                If Not Directory.Exists(Server.MapPath("SpedPisCofins")) Then
                    Directory.CreateDirectory(Server.MapPath("SpedPisCofins"))
                End If

                txtDataInicial.Text = "01/" & Today.AddMonths(-1).ToString("MM") & "/" & Today.ToString("yyyy")
                txtDataFinal.Text = Date.DaysInMonth(Today.Year, Today.AddMonths(-1).Month) & "/" & Today.AddMonths(-1).ToString("MM") & "/" & Today.ToString("yyyy")

                CargaUnidade()
                VerificaUnidade()
                imdDownload.Visible = False
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Fiscal.aspx")
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf &
              "  FROM Clientes C " & vbCrLf &
              " INNER JOIN ClientesXTipos CT " & vbCrLf &
              "    ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf &
              " WHERE CT.Tipo_Id = 050 " & vbCrLf &
              " ORDER BY Nome" & vbCrLf

        Dim dsUnidades As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
        If dsUnidades IsNot Nothing AndAlso dsUnidades.Tables.Count > 0 Then
            For Each Dr As DataRow In dsUnidades.Tables(0).Rows
                ddlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
            Next
        End If

        ddlUnidade.Items.Insert(0, "")
        ddlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "       isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "       isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              "  from Usuarios" & vbCrLf &
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        Dim dsUnidades As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")
        If dsUnidades IsNot Nothing AndAlso dsUnidades.Tables.Count > 0 Then
            For Each Dr As DataRow In dsUnidades.Tables(0).Rows
                ddlUnidade.SelectedValue = Dr("AcessoUnidade")
                CargaEmpresas()
                DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            Next
        End If
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()
        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "SELECT Cli.Cliente_Id as Codigo, Cli.Endereco_Id, Cli.Reduzido, Cli.Nome, Cli.Cidade, Cli.Estado " & vbCrLf &
              "  FROM GruposXEmpresas AS GxE" & vbCrLf &
              " INNER JOIN Clientes AS Cli " & vbCrLf &
              "    ON GxE.Cliente_Id = Cli.Cliente_Id " & vbCrLf &
              "   AND GxE.EndCliente_Id = Cli.Endereco_Id " & vbCrLf &
              " WHERE GxE.Empresa_Id = '" & ddlUnidade.SelectedValue & "' " & vbCrLf

        Dim dsEmpresas As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
        If dsEmpresas IsNot Nothing AndAlso dsEmpresas.Tables.Count > 0 Then
            For Each Dr As DataRow In dsEmpresas.Tables(0).Rows
                Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
                Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
                Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
                Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
                Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
                Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
                DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
            Next
        End If

        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function Validar()
        If ddlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a unidade de negocio.", eTitulo.Info)
            Return False
        ElseIf DdlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a empresa.", eTitulo.Info)
            Return False
        ElseIf txtDataInicial.Text = "" Then
            MsgBox(Me.Page, "Informe o periodo inicial.", eTitulo.Info)
            Return False
        ElseIf txtDataFinal.Text = "" Then
            MsgBox(Me.Page, "Informe o período final.", eTitulo.Info)
            Return False
        ElseIf txt_NUM_ORD.Text = "" Then
            MsgBox(Me.Page, "Informe o número recibo anterior.", eTitulo.Info)
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        ddlIndicadorAtividade.Items.Clear()
        ddlIndicadorDeNatureza.Items.Clear()
        txtDataInicial.Text = Today
        txtDataFinal.Text = Today
        txt_NUM_ORD.Text = ""
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmdLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            ddlUnidade.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imdDownload_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imdDownload.Click
        Try
            Response.ContentType = "text/plain"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & txtArquivoDeSaida.Text)
            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(Server.MapPath("~/SpedPisCofins/" & txtArquivoDeSaida.Text), FileMode.Open, FileAccess.Read)
                Do
                    If Response.IsClientConnected Then
                        length = download.Read(buffer, 0, bufferLength)
                        Response.OutputStream.Write(buffer, 0, length)
                        buffer = New Byte(bufferLength - 1) {}
                    Else
                        length = -1
                    End If
                Loop While length > 0
                Response.Flush()
                Response.End()
            Finally
                If download IsNot Nothing Then
                    download.Close()
                End If
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString)
        End Try
    End Sub

    Protected Sub Processar()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        EmpresaMestre = Left(Empresa(0), 8)
        PeriodoInicial = DateValue(txtDataInicial.Text).ToString("yyyy-MM-dd")
        PeriodoFinal = DateValue(txtDataFinal.Text).ToString("yyyy-MM-dd")

        Dim NomeArquivo As String = ""

        txtArquivoDeSaida.Text = "EFD-Contribuicoes-" & DateValue(txtDataInicial.Text).ToString("MM-yyyy") & "-" & Empresa(0) & ".txt"

        Dim NomeArquivo2 As String = "SpedPisCofins/" & txtArquivoDeSaida.Text.ToUpper
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        Using sw As New StreamWriter(Server.MapPath("~/SpedPisCofins/" & txtArquivoDeSaida.Text.Trim().ToUpper()), True)
            Try
                Dim CGC As String = ""
                Dim Inscricao As String = ""
                Dim CpfContador As String = ""
                Dim CRCContador As String = ""

                Dim Nota As Integer = 0
                Dim Serie As String = ""
                Dim EntradaSaida As String = ""
                Dim Cliente As String = ""
                Dim EndCliente As Integer = 0
                Dim Emissao As String = ""
                Dim Modelo As String = ""
                Dim Cancelada As String = ""

                Dim Array As New ArrayList
                Dim ds As DataSet

                ''Registro 0000 - Abertura do Arquivo Digital e Identificação da Pessoa Juridica
                ds = Consulta0000()

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Cadastro da Empresa com problemas.", eTitulo.Info)
                    Exit Sub
                End If

                CompoeRegistro0000(ds, ArquivoAux, Registro0000, RegistroGeral)
                '****************

                ''Registro 0001  - Abertura do Bloco 0
                CompoeRegistro0001(ArquivoAux, Registro0001, RegistroGeral)
                '****************

                ' Registro 0100  - Dados do Contabilista
                CompoeRegistro0100(ds, ArquivoAux, Registro0100, RegistroGeral)

                ''Registro 0110  - Regime de apuracao da contribuição social e de apropriação de credito
                CompoeRegistro0110(ArquivoAux, Registro0100, RegistroGeral)

                ' Registro 0111  - Tabela de receita bruta mensal para fins de rateio de cretitos comuns
                If cboIncidencia.SelectedValue <> 2 Then

                    ds = ConsultaRegistro0111()

                    CompoeRegistro0111(ds, ArquivoAux, Registro0111, RegistroGeral)

                End If

                ' Registro 0140  - Tabela de cadastrato de estabelecimento
                ds = ConsultaRegistro0140()

                Dim EmpresaRegistro0140 As String = String.Empty

                For Each dr In ds.Tables(0).Rows
                    With dr
                        EmpresaRegistro0140 = .item("Cliente_Id")
                        linha = "|0140"
                        linha &= "|" & RTrim(.Item("Reduzido"))
                        linha &= "|" & .Item("Nome")
                        linha &= "|" & .Item("Cliente_Id")
                        linha &= "|" & .Item("Estado")
                        linha &= "|" & .Item("Inscricao").ToString.Replace("ISENTO", "")
                        linha &= "|" & .Item("EstadoIbge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                        linha &= "|" & .Item("InscricaoMunicipal")
                        linha &= "|" & .Item("Suframa")
                        linha &= "|"
                        ArquivoAux.Add(linha)
                        Registro0140 += 1
                        RegistroGeral += 1

                        'Registro 0145 - Regime de Apuração da contribuição
                        If (EmpresaMestre = "11111111") Then 'isolado temporariamente 03189063 utilizado somente no Curtume
                            dsaux = ConsultaRegistro0145()
                            CompoeRegistro0145(ds, ArquivoAux, Registro0145, RegistroGeral)
                        End If

                        'Registro 0150 - Tabela de Cadastro do Participante
                        ds = ConsultaRegistro0150()
                        CompoeRegistro0150(ds, ArquivoAux, Registro0150, RegistroGeral, EmpresaRegistro0140)
                        '**************************************************

                        'Registro 0190 - Identificação das Unidades de Medidas
                        Sql = "SELECT Unidade_Id AS Unidade," & vbCrLf &
                              "       Descricao" & vbCrLf &
                              "  FROM UnidadeDeMedida" & vbCrLf

                        'Registro 0190 - Identificação das Unidades de Medidas  
                        'Sql = "/*Registro 0190*/" & vbCrLf & _
                        '       "SELECT Unidade, UN.Descricao " & vbCrLf & _
                        '       " into #Unidades " & vbCrLf & _
                        '       "  FROM (" & vbCrLf & _
                        '       "        SELECT Produtos.Unidade " & vbCrLf & _
                        '       "          FROM NotasFiscais" & vbCrLf & _
                        '       "         INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                        '       "            ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                        '       "         INNER JOIN Produtos " & vbCrLf & _
                        '       "            ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                        '       "          LEFT JOIN NFERealizadas" & vbCrLf & _
                        '       "            ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                        '       "           AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                        '       "         INNER JOIN SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id " & vbCrLf & _
                        '       "           AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                        '       "         Inner Join OperacaoXEstado OE" & vbCrLf & _
                        '       "            ON OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf & _
                        '       "         WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
                        '       "           AND NotasFiscais.Empresa_Id =  '" & .Item("Cliente_Id") & "'" & vbCrLf & _
                        '       "           AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                        '       "           AND NotasFiscais.TipoDeDocumento NOT IN (15)" & vbCrLf &
                        '       "           AND NotasFiscais.Finalidade NOT IN (30)" & vbCrLf &
                        '       "           AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
                        '       "           AND (OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                        '       "           AND (OE.CodigoFiscal NOT BETWEEN 2300 AND 2360)" & vbCrLf & _
                        '       "           AND (OE.CodigoFiscal NOT IN (1932,2932))" & vbCrLf & _
                        '       "           AND (OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                        '       "           AND (OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                        '       "           AND (OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
                        '       "           AND (OE.CodigoFiscal NOT BETWEEN 1302 AND 1303)" & vbCrLf & _
                        '       "           AND Not (NotasFiscais.EntradaSaida_Id = 'E'   " & vbCrLf & _
                        '       "           AND OE.CodigoFiscal IN (1933,2933)) " & vbCrLf & _
                        '       "           AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf & _
                        '       "           AND Not (NotasFiscais.EntradaSaida_Id = 'S'   " & vbCrLf & _
                        '       "           AND OE.CodigoFiscal IN (5933,6933)) " & vbCrLf & _
                        '       "           AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949)) " & vbCrLf & _
                        '       "           AND ((NFERealizadas.Chavenfe is null) or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null) )" & vbCrLf & _
                        '       "            OR (NotasFiscais.Empresa_Id =  '" & .Item("Cliente_Id") & "' " & vbCrLf & _
                        '       "                AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' " & vbCrLf & _
                        '       "                AND NotasFiscais.TipoDeDocumento = 2" & vbCrLf & _
                        '       "                AND NotasFiscais.NossaEmissao    ='S') " & vbCrLf & _
                        '       "         GROUP BY Produtos.Unidade " & vbCrLf & _
                        '       "        having sum(NotasFiscaisxitens.quantidadefiscal) > 0) AS Consulta " & vbCrLf & _
                        '       " INNER JOIN UnidadeDeMedida UN " & vbCrLf & _
                        '       "    ON Unidade = UN.Unidade_Id " & vbCrLf & _
                        '       " GROUP BY Unidade, UN.Descricao " & vbCrLf

                        'If CDate(PeriodoFinal).Year > 2017 Then
                        '    Sql &= "INSERT INTO #Unidades" & vbCrLf & _
                        '            "SELECT Produtos.Unidade, UN.Descricao" & vbCrLf & _
                        '            "FROM ApuracaoDeCustos  " & vbCrLf & _
                        '            "		INNER JOIN Produtos " & vbCrLf & _
                        '            "				ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                        '            "		INNER JOIN UnidadeDeMedida UN " & vbCrLf & _
                        '            "				ON Produtos.Unidade = UN.Unidade_Id" & vbCrLf & _
                        '            "Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                        '            "    And ApuracaoDeCustos.Empresa_Id = '" & .Item("Cliente_Id") & "'" & vbCrLf & _
                        '            "    And ApuracaoDeCustos.Mes_Id =  " & CDate(PeriodoFinal).Month & vbCrLf & _
                        '            "    And ApuracaoDeCustos.Ano_Id = " & CDate(PeriodoFinal).Year & vbCrLf & _
                        '            "    And ApuracaoDeCustos.Quantidade > 0 " & vbCrLf & _
                        '            "    And ApuracaoDeCustos.ValorDoProduto > 0" & vbCrLf & _
                        '            "    And  ISNULL(Produtos.TipoDoItem, 0) in (00, 01, 02, 03, 04, 05, 06, 10)" & vbCrLf & _
                        '            "GROUP BY Produtos.Unidade, UN.Descricao" & vbCrLf
                        'End If

                        'Sql &= "select Unidade, Descricao from #Unidades " & vbCrLf & _
                        '       "GROUP BY Unidade, Descricao" & vbCrLf

                        ds = Banco.ConsultaDataSet(Sql, "Consulta0190")

                        If ds.Tables(0).Rows.Count > 0 Then
                            For Each dr2 In ds.Tables(0).Rows
                                With dr2
                                    linha = "|0190"
                                    linha &= "|" & RTrim(.Item("Unidade"))
                                    linha &= "|" & RTrim(.Item("Descricao"))
                                    linha &= "|"
                                End With

                                ArquivoAux.Add(linha)
                                Registro0190 += 1
                                RegistroGeral += 1
                            Next
                        End If

                        'Registro 0200 - Tabela de Identificação do Item (Produtos e Serviços)
                        Sql = " SELECT DISTINCT Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS" & vbCrLf &
                              " FROM (SELECT Produto_Id, Nome, Unidade, ISNULL(TipoDoItem, 0) AS TipoDoItem, ISNULL(NCM, N'') AS NCM, ISNULL(CodigoEX, N'') AS CodigoEX," & vbCrLf &
                              "              ISNULL(CodigoDoGenero, 0) AS CodigoDoGenero, ISNULL(CodigoDoServico, 0) AS CodigoDoServico, ISNULL(ICMS, 0) AS ICMS" & vbCrLf &
                              "         FROM Produtos" & vbCrLf &
                              "        WHERE (Len(NCM) > 7)" & vbCrLf &
                              "          AND Situacao = 1" & vbCrLf &
                              "        GROUP BY Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS" & vbCrLf &
                              "      ) AS Consulta" & vbCrLf &
                              " GROUP BY Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS" & vbCrLf

                        'Registro 0200 - Tabela de Identificação do Item (Produtos e Serviços)
                        'Sql = "/*Registro 0200*/" & vbCrLf & _
                        '       "SELECT P.Produto_Id," & vbCrLf & _
                        '       "       P.Nome," & vbCrLf & _
                        '       "       P.Unidade," & vbCrLf & _
                        '       "       ISNULL(P.TipoDoItem, 0) AS TipoDoItem," & vbCrLf & _
                        '       "       ISNULL(P.NCM, '') AS NCM," & vbCrLf & _
                        '       "       ISNULL(P.CodigoEX, '') AS CodigoEX," & vbCrLf & _
                        '       "       ISNULL(P.CodigoDoGenero, 0) AS CodigoDoGenero," & vbCrLf & _
                        '       "       ISNULL(P.CodigoDoServico, 0) AS CodigoDoServico," & vbCrLf & _
                        '       "       ISNULL(P.ICMS, 0) AS ICMS" & vbCrLf & _
                        '       " into #BASEPRODUTOS" & vbCrLf & _
                        '       "  FROM PRODUTOS P " & vbCrLf & _
                        '       " Where produto_Id" & vbCrLf & _
                        '       "    in (" & vbCrLf & _
                        '       "          --Select isnull(pa.Produto_Id, p.Produto_Id)" & vbCrLf & _
                        '       "           -- from Produtos p" & vbCrLf & _
                        '       "           -- left join ProdutosAgrupados pa" & vbCrLf & _
                        '       "             -- on p.Produto_Id = pa.ProdutoAgrupado_Id" & vbCrLf & _
                        '       "           --Where p.Produto_id in (" & vbCrLf & _
                        '       "                                    SELECT ISNULL(PA.Produto_id, NotasFiscaisXItens.Produto_Id) AS Produto_Id" & vbCrLf & _
                        '       "                                      FROM NotasFiscais" & vbCrLf & _
                        '       "                                     INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                        '       "                                        ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                        '       "                                     Inner Join OperacaoXEstado OE" & vbCrLf & _
                        '       "                                        ON OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf & _
                        '       "                                      LEFT JOIN NotaFiscalXExportacao AS NFEx " & vbCrLf & _
                        '       "                                        ON NotasFiscais.Empresa_Id       = NFEx.Empresa_Id " & vbCrLf & _
                        '       "                                       AND NotasFiscais.EndEmpresa_Id    = NFEx.EndEmpresa_Id " & vbCrLf & _
                        '       "                                       AND NotasFiscais.Cliente_Id       = NFEx.Cliente_Id " & vbCrLf & _
                        '       "	                                   AND NotasFiscais.EndCliente_Id    = NFEx.EndCliente_Id " & vbCrLf & _
                        '       "                                       AND NotasFiscais.EntradaSaida_Id  = NFEx.EntradaSaida_Id " & vbCrLf & _
                        '       "                                       AND NotasFiscais.Serie_Id         = NFEx.Serie_Id " & vbCrLf & _
                        '       "                                       AND NotasFiscais.Nota_Id          = NFEx.Nota_Id " & vbCrLf & _
                        '       "                                      LEFT JOIN MemorandoDeExportacao AS ME " & vbCrLf & _
                        '       "                                        ON NotasFiscais.Empresa_Id = ME.Empresa " & vbCrLf & _
                        '       "                                       AND NotasFiscais.EndEmpresa_Id = ME.EndEmpresa " & vbCrLf & _
                        '       "                                       AND NotasFiscais.Cliente_Id = ME.Cliente " & vbCrLf & _
                        '       "                                       AND NotasFiscais.EndCliente_Id = ME.EndCliente" & vbCrLf & _
                        '       "                                       AND NotasFiscais.EntradaSaida_Id = ME.EntradaSaida " & vbCrLf & _
                        '       "                                       AND NotasFiscais.Serie_Id = ME.Serie " & vbCrLf & _
                        '       "                                       AND NotasFiscais.Nota_Id = ME.Nota" & vbCrLf & _
                        '       "                                      --LEFT JOIN ProdutosAgrupados PA" & vbCrLf & _
                        '       "                                        --ON PA.ProdutoAgrupado_Id = NotasFiscaisXItens.Produto_Id" & vbCrLf & _
                        '       "                                     --INNER JOIN Produtos " & vbCrLf & _
                        '       "                                        --ON Produtos.Produto_Id = isnull(PA.Produto_id,NotasFiscaisXItens.Produto_Id)" & vbCrLf & _
                        '       "                                      LEFT JOIN ProdutosAgrupados pa" & vbCrLf & _
                        '       "                                        ON NotasFiscaisXItens.Produto_Id = pa.ProdutoAgrupado_Id" & vbCrLf & _
                        '       "                                      LEFT JOIN NFERealizadas" & vbCrLf & _
                        '       "                                        ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                        '       "                                       AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                        '       "                                     INNER JOIN SubOperacoes" & vbCrLf & _
                        '       "                                        ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id " & vbCrLf & _
                        '       "                                       AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                        '       "                                     WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
                        '       "                                       AND NotasFiscais.Situacao not in(2,9,10) " & vbCrLf & _
                        '       "                                       AND NotasFiscais.Empresa_Id =  '" & .Item("Cliente_Id") & "'" & vbCrLf & _
                        '       "                                       AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                        '       "                                       AND NotasFiscais.TipoDeDocumento NOT IN (15)" & vbCrLf &
                        '       "                                       AND NotasFiscais.Finalidade NOT IN (30)" & vbCrLf &
                        '       "                                       AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
                        '       "                                       AND (OE.CodigoFiscal NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                        '       "                                       AND (OE.CodigoFiscal NOT BETWEEN 2300 AND 2360)" & vbCrLf & _
                        '       "                                       AND (OE.CodigoFiscal NOT IN(1932,2932))" & vbCrLf & _
                        '       "                                       AND (OE.CodigoFiscal NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                        '       "                                       AND (OE.CodigoFiscal NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                        '       "                                       AND (OE.CodigoFiscal NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
                        '       "                                       AND (OE.CodigoFiscal NOT BETWEEN 1302 AND 1303)" & vbCrLf & _
                        '       "                                       AND Not (NotasFiscais.EntradaSaida_Id = 'E'  AND OE.CodigoFiscal IN (1933,2933))" & vbCrLf & _
                        '       "                                       AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (1949,2949))" & vbCrLf & _
                        '       "                                       --AND Not (NotasFiscais.EntradaSaida_Id = 'S'  AND OE.CodigoFiscal IN (5933,6933))" & vbCrLf & _
                        '       "                                       AND NOT (SubOperacoes.Classe IN ('" & eClassesOperacoes.SERVICOS.ToString & "') AND OE.CodigoFiscal IN (5949,6949)) " & vbCrLf & _
                        '       "                                       AND ((NFERealizadas.Chavenfe is null) " & vbCrLf & _
                        '       "                                              or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null)" & vbCrLf & _
                        '       "                                              or (ISNULL(ME.DataAverbacao, NFEx.DataAverbacao) BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'))" & vbCrLf & _
                        '       "                                        OR (NotasFiscais.Empresa_Id =  '" & .Item("Cliente_Id") & "' AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' AND NotasFiscais.TipoDeDocumento=2 AND NotasFiscais.NossaEmissao='S') " & vbCrLf & _
                        '       "                                     group by ISNULL(PA.Produto_id, NotasFiscaisXItens.Produto_Id), OE.CodigoFiscal" & vbCrLf & _
                        '       "                                    having sum(NotasFiscaisxitens.quantidadefiscal) > 0)" & vbCrLf

                        'If CDate(PeriodoFinal).Year > 2017 Then
                        '    Sql &= "INSERT INTO #BASEPRODUTOS" & vbCrLf & _
                        '            "SELECT P.Produto_Id," & vbCrLf & _
                        '            "       P.Nome," & vbCrLf & _
                        '            "       P.Unidade," & vbCrLf & _
                        '            "       ISNULL(P.TipoDoItem, 0) AS TipoDoItem," & vbCrLf & _
                        '            "       ISNULL(P.NCM, '') AS NCM," & vbCrLf & _
                        '            "       ISNULL(P.CodigoEX, '') AS CodigoEX," & vbCrLf & _
                        '            "       ISNULL(P.CodigoDoGenero, 0) AS CodigoDoGenero," & vbCrLf & _
                        '            "       ISNULL(P.CodigoDoServico, 0) AS CodigoDoServico," & vbCrLf & _
                        '            "       ISNULL(P.ICMS, 0) AS ICMS" & vbCrLf & _
                        '            "FROM ApuracaoDeCustos " & vbCrLf & _
                        '            "		INNER JOIN PRODUTOS P" & vbCrLf & _
                        '            "				ON P.Produto_id = ApuracaoDeCustos.Produto_Id" & vbCrLf & _
                        '            "Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                        '            "    And ApuracaoDeCustos.Empresa_Id = '" & .Item("Cliente_Id") & "'" & vbCrLf & _
                        '            "    And ApuracaoDeCustos.Mes_Id =  " & CDate(PeriodoFinal).Month & vbCrLf & _
                        '            "    And ApuracaoDeCustos.Ano_Id = " & CDate(PeriodoFinal).Year & vbCrLf & _
                        '            "    And ApuracaoDeCustos.Quantidade > 0 " & vbCrLf & _
                        '            "    And ApuracaoDeCustos.ValorDoProduto > 0" & vbCrLf & _
                        '            "    And  ISNULL(P.TipoDoItem, 0) in (00, 01, 02, 03, 04, 05, 06, 10)" & vbCrLf
                        'End If

                        'Sql &= "INSERT INTO #BASEPRODUTOS" & vbCrLf & _
                        '       "     Select sb.Produto_id," & vbCrLf & _
                        '       "            Prd.Nome," & vbCrLf & _
                        '       "            Prd.Unidade," & vbCrLf & _
                        '       "            ISNULL(Prd.TipoDoItem, 0) AS TipoDoItem," & vbCrLf & _
                        '       "            ISNULL(Prd.NCM, '') AS NCM," & vbCrLf & _
                        '       "            ISNULL(Prd.CodigoEX, '') AS CodigoEX," & vbCrLf & _
                        '       "            ISNULL(Prd.CodigoDoGenero, 0) AS CodigoDoGenero," & vbCrLf & _
                        '       "            ISNULL(Prd.CodigoDoServico, 0) AS CodigoDoServico," & vbCrLf & _
                        '       "            ISNULL(Prd.ICMS, 0) AS ICMS" & vbCrLf & _
                        '       "     from (" & vbCrLf & _
                        '       "   		select nfi.Empresa_Id," & vbCrLf & _
                        '       "   		       nfi.EndEmpresa_id," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
                        '       "   				  when SO.ProdutoDeTerceiro =  1                                                                    then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
                        '       "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
                        '       "   			   end as Classificacao," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                        '       "   				   then nf.destino" & vbCrLf & _
                        '       "   				   else nfi.cliente_id" & vbCrLf & _
                        '       "   			   end Cliente_id," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                        '       "   				   then nf.Enddestino" & vbCrLf & _
                        '       "   				   else nfi.Endcliente_id" & vbCrLf & _
                        '       "   			   end EndCliente_id," & vbCrLf & _
                        '       "   			   nfi.Produto_Id," & vbCrLf & _
                        '       "   			   isnull(nf.pedido,0) as Pedido," & vbCrLf & _
                        '       "                  convert(numeric(18,4),0) as QtdeRazao," & vbCrLf & _
                        '       "   			   SUM(Case" & vbCrLf & _
                        '       "   					 When So.Devolucao = 'S'" & vbCrLf & _
                        '       "   					   then nfi.QuantidadeFiscal * -1" & vbCrLf & _
                        '       "   					   else nfi.QuantidadeFiscal" & vbCrLf & _
                        '       "   				   end) QtdeFiscal," & vbCrLf & _
                        '       "               convert(numeric(18,4),0) as ValorRazao," & vbCrLf & _
                        '       "   			   SUM(Case" & vbCrLf & _
                        '       "   					 When So.Devolucao = 'S'" & vbCrLf & _
                        '       "   					   then nfe.valor * -1" & vbCrLf & _
                        '       "   					   else nfe.valor" & vbCrLf & _
                        '       "   				   end) as valorFiscal" & vbCrLf & _
                        '       "   		  from notasfiscais NF" & vbCrLf & _
                        '       "   		 inner join Notasfiscaisxitens NFI" & vbCrLf & _
                        '       "   			on NF.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
                        '       "   		   and NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                        '       "   		   and NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                        '       "   		   and NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                        '       "   		   and NF.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
                        '       "   		   and NF.Nota_id         = NFI.Nota_id" & vbCrLf & _
                        '       "   		   and NF.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                        '       "   		 Inner Join NotasFiscaisxEncargos Nfe" & vbCrLf & _
                        '       "   			on Nfe.Empresa_id      = NFI.Empresa_Id" & vbCrLf & _
                        '       "   		   and Nfe.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                        '       "   		   and Nfe.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                        '       "   		   and Nfe.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                        '       "   		   and Nfe.EntradaSaida_Id = NFI.EntradaSaida_id" & vbCrLf & _
                        '       "   		   and Nfe.Nota_id         = NFI.Nota_id" & vbCrLf & _
                        '       "   		   and Nfe.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                        '       "   		   and Nfe.Produto_id      = NFI.Produto_id" & vbCrLf & _
                        '       "   		   and Nfe.CFOP_Id         = NFI.CFOP_Id" & vbCrLf & _
                        '       "   		   and Nfe.Sequencia_Id    = NFI.Sequencia_Id" & vbCrLf & _
                        '       "   		   and Nfe.Encargo_Id      = 'LIQUIDO'" & vbCrLf & _
                        '       "   		 Inner join produtos p" & vbCrLf & _
                        '       "   			on NFI.Produto_id = P.Produto_id" & vbCrLf & _
                        '       "   		 Inner Join Suboperacoes SO" & vbCrLf & _
                        '       "   			on SO.Operacao_Id        = NFI.Operacao" & vbCrLf & _
                        '       "   		   and SO.Suboperacoes_Id    = NFI.Suboperacao" & vbCrLf & _
                        '       "   		 where (SO.produtodeterceiro = 1 or SO.deposito = 'S')" & vbCrLf & _
                        '       "   		   and (SO.EstoqueFisico = 'S' or SO.EstoqueFiscal = 'S' or SO.Devolucao = 'S')" & vbCrLf & _
                        '       "   		   and SO.Operacao_id <> 80" & vbCrLf & _
                        '       "   		   and nf.movimento  <= '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                        '       "           and nf.Empresa_Id = '" & .Item("Cliente_Id") & "'" & vbCrLf & _
                        '       "   		  group by nfi.Empresa_Id," & vbCrLf & _
                        '       "   		           nfi.EndEmpresa_id," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				  when So.Deposito = 'S'  or  (so.consignacao = 1 and so.EntradaSaida = 'E' and so.Devolucao = 'S') then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
                        '       "   				  when SO.ProdutoDeTerceiro = 1                                                                     then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
                        '       "   				  ELSE 'SEM CLASSIFICACAO'" & vbCrLf & _
                        '       "   			   end," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				 When left(nf.empresa_id, 8) = left(nfi.Cliente_id, 8)" & vbCrLf & _
                        '       "   				   then nf.destino" & vbCrLf & _
                        '       "   				   else nfi.cliente_id" & vbCrLf & _
                        '       "   			   end," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				 When left(nf.empresa_id,8) = left(nfi.Cliente_id,8)" & vbCrLf & _
                        '       "   				   then nf.Enddestino" & vbCrLf & _
                        '       "   				   else nfi.Endcliente_id" & vbCrLf & _
                        '       "   			   end," & vbCrLf & _
                        '       "               nf.pedido," & vbCrLf & _
                        '       "               nfi.Produto_Id" & vbCrLf & _
                        '       "        union all" & vbCrLf & _
                        '       "   		Select R.Empresa_Id," & vbCrLf & _
                        '       "   			   R.Endempresa_id," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then 'TERCEIROS EM NOSSO PODER'" & vbCrLf & _
                        '       "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then 'NOSSO EM PODER TERCEIROS'" & vbCrLf & _
                        '       "   			   end Classificacao," & vbCrLf & _
                        '       "   			   Case" & vbCrLf & _
                        '       "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
                        '       "                   then R.Deposito" & vbCrLf & _
                        '       "                   else R.Cliente_id" & vbCrLf & _
                        '       "               end Cliente_id," & vbCrLf & _
                        '       "   			   Case" & vbCrLf & _
                        '       "                 when R.Empresa_id = R.cliente_Id" & vbCrLf & _
                        '       "                   then R.EndDeposito" & vbCrLf & _
                        '       "                   else R.EndCliente_id" & vbCrLf & _
                        '       "               end EndCliente_id," & vbCrLf & _
                        '       "   			   R.Produto," & vbCrLf & _
                        '       "   			   isnull(R.Pedido,0) as Pedido," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then Isnull(R.CreditoQuantidade, 0.0000) - Isnull(R.DebitoQuantidade, 0.0000)" & vbCrLf & _
                        '       "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoQuantidade, 0.0000)  - Isnull(R.CreditoQuantidade, 0.0000)" & vbCrLf & _
                        '       "   			   end," & vbCrLf & _
                        '       "   			   convert(numeric(18,4),0)," & vbCrLf & _
                        '       "   			   case" & vbCrLf & _
                        '       "   				  when R.Conta_id = CxE.ContaEstoqueEmNossoPoder                       then isnull(R.CreditoOficial, 0.00) - isnull(R.DebitoOficial, 0.00)" & vbCrLf & _
                        '       "   				  when left(R.Conta_id,5) = Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) then isnull(R.DebitoOficial, 0.00)  - isnull(R.CreditoOficial, 0.00)" & vbCrLf & _
                        '       "   			   end," & vbCrLf & _
                        '       "               convert(numeric(18,4),0)" & vbCrLf & _
                        '       "   		  from Razao R" & vbCrLf & _
                        '       "   		 inner join ClientesxEmpresas CxE" & vbCrLf & _
                        '       "   			on R.Empresa_id    = CxE.Empresa_Id" & vbCrLf & _
                        '       "   		   and R.EndEmpresa_id = CxE.EndEmpresa_Id" & vbCrLf & _
                        '       "   		   and (left(R.Conta_Id,7) = CxE.ContaEstoqueEmNossoPoder or Left(CxE.ContaEstoqueEmPoderDeTerceiros,5) = Left(cxe.ContaestoqueemPoderdeTerceiros,5))" & vbCrLf & _
                        '       "   		 inner join Clientes C" & vbCrLf & _
                        '       "   			on R.Cliente_id    = C.Cliente_id" & vbCrLf & _
                        '       "   		   and R.EndCliente_id = C.Endereco_id" & vbCrLf & _
                        '       "   		 where lote_id not in (9,10,11)" & vbCrLf & _
                        '       "           and R.Movimento_Id <= '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                        '       "           and year(R.Movimento_Id) > 2010 " & vbCrLf & _
                        '       "           and R.Empresa_Id = '" & .Item("Cliente_Id") & "'" & vbCrLf & _
                        '       "   		   and len(isnull(produto,'')) > 0" & vbCrLf & _
                        '       "   		) as Sb" & vbCrLf & _
                        '       "     inner join Clientes Emp" & vbCrLf & _
                        '       "        on Sb.Empresa_id    = Emp.Cliente_id" & vbCrLf & _
                        '       "       and Sb.EndEmpresa_id = Emp.Endereco_id" & vbCrLf & _
                        '       "     inner join Clientes Cli" & vbCrLf & _
                        '       "        on Sb.cliente_id    = Cli.Cliente_id" & vbCrLf & _
                        '       "       and Sb.EndCliente_id = Cli.Endereco_id" & vbCrLf & _
                        '       "     inner Join Produtos Prd" & vbCrLf & _
                        '       "        on Prd.Produto_id = sb.Produto_id" & vbCrLf & _
                        '       "     Where Prd.TipoDoItem Not in(7,8,9,10,99)" & vbCrLf & _
                        '       "    Group by sb.Produto_id," & vbCrLf & _
                        '       "             Prd.Nome," & vbCrLf & _
                        '       "             Prd.Unidade," & vbCrLf & _
                        '       "             Prd.TipoDoItem," & vbCrLf & _
                        '       "             Prd.NCM," & vbCrLf & _
                        '       "             Prd.CodigoEX," & vbCrLf & _
                        '       "             Prd.CodigoDoGenero," & vbCrLf & _
                        '       "             Prd.CodigoDoServico," & vbCrLf & _
                        '       "             Prd.ICMS" & vbCrLf & _
                        '       "    having SUM(sb.QtdeRazao + sb.QtdeFiscal)   <> 0  Or Sum(sb.ValorRazao + sb.ValorFiscal) <> 0" & vbCrLf

                        'Sql &= "SELECT Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS" & vbCrLf & _
                        '        "FROM #BASEPRODUTOS" & vbCrLf & _
                        '        "GROUP BY Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS" & vbCrLf

                        ds = Banco.ConsultaDataSet(Sql, "Consulta0200")

                        If ds.Tables(0).Rows.Count > 0 Then
                            For Each dr3 In ds.Tables(0).Rows
                                With dr3
                                    linha = "|0200"                                                    ' 1 - Texto fixo contendo "0200" 
                                    linha &= "|" & RTrim(.item("Produto_Id"))                          ' 2 - Código do item
                                    linha &= "|" & RTrim(.item("Nome"))                                ' 3 - Descrição do item
                                    linha &= "|"                                                       ' 4 - Código de Barra
                                    linha &= "|"                                                       ' 5 - Código Anterior
                                    linha &= "|" & RTrim(.item("Unidade"))                             ' 6 - Unidade de medida utilizada na quantificação de estoques
                                    linha &= "|" & Funcoes.AlinharDireita(.item("TipoDoItem"), 2, "0") ' 7 - TIPO_ITEM
                                    linha &= "|" & Left(Replace(.item("NCM"), ".", ""), 8)             ' 8 - Código da Nomenclatura Comum do Mercosul
                                    linha &= "|" & RTrim(.item("CodigoEX"))                            ' 9 - Código EX, conforme a TIPI
                                    If .item("CodigoDoGenero") = 0 Then                                '10 - Código do gênero do item, conforme a Tabela 4.2.1.
                                        linha &= "|"
                                    Else
                                        linha &= "|" & Format(RTrim(.item("CodigoDoGenero")), "00")
                                    End If
                                    'linha &= "|" & RTrim(.Item("CodigoDoServico"))
                                    linha &= "|"                                                       '11 - Código do serviço conforme lista do Anexo I da Lei Complementar Federal nº 116/03.
                                    linha &= "|"                                                       '12 - Alíquota de ICMS aplicável ao item nas operações internas
                                    linha &= "|"
                                End With

                                ArquivoAux.Add(linha)
                                Registro0200 += 1
                                RegistroGeral += 1
                            Next
                        End If

                        'Registro 0400 - Tabela de natureza da operação/prestação
                        Sql = "SELECT Operacao_Id, SubOperacoes_Id, Descricao" & vbCrLf &
                              "  FROM (SELECT SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, SubOperacoes.Descricao" & vbCrLf &
                              "          FROM NotasFiscais" & vbCrLf &
                              "         INNER JOIN NotasFiscaisXItens" & vbCrLf &
                              "            ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                              "           AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                              "           AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                              "           AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                              "           AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                              "           AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                              "           AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                              "         INNER JOIN SubOperacoes" & vbCrLf &
                              "            ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                              "           AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                              "         WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F'))) " & vbCrLf &
                              "           and NotasFiscais.Empresa_Id = '" & .Item("Cliente_Id") & "'" & vbCrLf &
                              "           AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'" & vbCrLf &
                              "         GROUP BY SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, SubOperacoes.Descricao" & vbCrLf &
                              "       ) AS Consulta" & vbCrLf &
                              "  GROUP BY Operacao_Id, SubOperacoes_Id, Descricao" & vbCrLf

                        ds = Banco.ConsultaDataSet(Sql, "Consulta")

                        If ds.Tables(0).Rows.Count > 0 Then
                            For Each dr4 In ds.Tables(0).Rows
                                With dr4
                                    linha = "|0400"
                                    linha &= "|" & Format(.Item("Operacao_Id"), "00") & Format(.Item("SubOperacoes_Id"), "00")
                                    linha &= "|" & RTrim(.Item("Descricao"))
                                    linha &= "|"
                                End With

                                ArquivoAux.Add(linha)
                                Registro0400 += 1
                                RegistroGeral += 1
                            Next
                        End If


                        'Registro 0450 - Tabela de informações complementares do documento fiscal
                        Sql = " SELECT Distinct Codigo_Id, Descricao " & vbCrLf &
                              "   FROM ObservacoesFiscais " & vbCrLf &
                              "  INNER JOIN NotasFiscais" & vbCrLf &
                              "     ON ObservacoesFiscais.Codigo_Id = isnull(NotasFiscais.Situacao,1) " & vbCrLf &
                              "    AND NotasFiscais.Empresa_Id ='" & .Item("Cliente_Id") & "'" & vbCrLf &
                              "    AND isnull(NotasFiscais.Situacao,1) not in(2,3) " & vbCrLf &
                              "    AND NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf

                        ds = Banco.ConsultaDataSet(Sql, "Consulta")

                        If ds.Tables(0).Rows.Count > 0 Then
                            For Each dr5 In ds.Tables(0).Rows
                                With dr5
                                    linha = "|0450"
                                    linha &= "|" & .Item("Codigo_Id")
                                    linha &= "|" & RTrim(.Item("Descricao"))
                                    linha &= "|"
                                End With

                                ArquivoAux.Add(linha)
                                Registro0450 += 1
                                RegistroGeral += 1
                            Next
                        End If



                    End With
                Next
                ''Fim Cadastro de Estabelecimento



                'Registro 0500 - Plano de Contas

                Sql = " SELECT Conta_Id, Titulo " & vbCrLf &
                      "   FROM PlanoDeContas" & vbCrLf &
                      "  WHERE  PlanoDeContas.titulo <> ''" & vbCrLf &
                      "  Order by Conta_Id" & vbCrLf


                ds = Banco.ConsultaDataSet(Sql, "PlanoDeContas")

                Dim Nivel1 As String = ""
                Dim Nivel2 As String = ""

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        With dr
                            linha = "|0500"
                            linha = linha & "|01012012"

                            Nivel1 = Left(.item("Conta_Id"), 1)
                            Nivel2 = Left(.item("Conta_Id"), 3)

                            If Nivel2 = "204" Then
                                linha = linha & "|03"
                            ElseIf Nivel2 = "105" Then
                                linha = linha & "|05"
                            ElseIf Nivel2 = "205" Then
                                linha = linha & "|05"
                            Else
                                If Nivel1 = "1" Then
                                    linha = linha & "|01"
                                ElseIf Nivel1 = "2" Then
                                    linha = linha & "|02"
                                ElseIf Nivel1 = "3" Then
                                    linha = linha & "|04"
                                ElseIf Nivel1 = "4" Then
                                    linha = linha & "|04"
                                ElseIf Nivel1 = "5" Then
                                    linha = linha & "|04"
                                End If
                            End If

                            If Len(.item("Conta_Id")) > 5 Then
                                linha = linha & "|A"
                            Else
                                linha = linha & "|S"
                            End If

                            If Len(.item("Conta_Id")) > 7 Then
                                linha = linha & "|5"
                            ElseIf Len(.item("Conta_Id")) = 7 Then
                                linha = linha & "|4"
                            ElseIf Len(.item("Conta_Id")) = 5 Then
                                linha = linha & "|3"
                            ElseIf Len(.item("Conta_Id")) = 3 Then
                                linha = linha & "|2"
                            ElseIf Len(.item("Conta_Id")) = 1 Then
                                linha = linha & "|1"
                            End If

                            linha = linha & "|" & .item("Conta_Id")
                            linha = linha & "|" & Left(RTrim(.item("Titulo")), 50)

                            If Len(.item("Conta_Id")) > 7 Then
                                linha = linha & "|" & Left(.item("Conta_Id"), 7)
                            ElseIf Len(.item("Conta_Id")) = 7 Then
                                linha = linha & "|" & Left(.item("Conta_Id"), 5)
                            ElseIf Len(.item("Conta_Id")) = 5 Then
                                linha = linha & "|" & Left(.item("Conta_Id"), 3)
                            ElseIf Len(.item("Conta_Id")) = 3 Then
                                linha = linha & "|" & Left(.item("Conta_Id"), 1)
                            ElseIf Len(.item("Conta_Id")) = 1 Then
                                linha = linha & "|"
                            End If
                            linha = linha & "||"

                        End With

                        ArquivoAux.Add(linha)
                        Registro0500 += 1
                        RegistroGeral += 1
                    Next
                End If

                'Registro 0600 - Centro De Custos
                Sql = " SELECT CentroDeCusto_Id, Descricao" & vbCrLf &
                      "   FROM CentrosDeCustos" & vbCrLf &
                      "  WHERE CentroDeCusto_Id <> 0" & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "CentrosDeCustos")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        With dr
                            linha = "|0600"
                            linha &= "|01012009"
                            linha &= "|" & .Item("CentroDeCusto_Id")
                            linha &= "|" & RTrim(.Item("Descricao"))
                            linha &= "|"
                        End With

                        ArquivoAux.Add(linha)
                        Registro0600 += 1
                        RegistroGeral += 1
                    Next
                End If


                '' Registro 0990  - Encerramento do Bloco 0
                Registro0990 += 1

                linha = "|0990"
                linha &= "|" & Registro0000 + Registro0001 + Registro0100 + Registro0110 + Registro0111 + Registro0140 + Registro0145 + Registro0150 + Registro0190 + Registro0200 + Registro0400 + Registro0450 + Registro0500 + Registro0600 + Registro0990
                linha &= "|"

                ArquivoAux.Add(linha)
                RegistroGeral += 1

                '  Fim do Bloco 0

                If cboIncidencia.SelectedValue = 1 Then

                    ' Registro A001  - Abertura do Bloco A
                    Sql = "SELECT Distinct" & vbCrLf &
                          "       Clientes.Nome, Clientes.Cliente_Id, Clientes.Estado, Clientes.Inscricao," & vbCrLf &
                          "       Clientes.CodigoDoMunicipio, ClientesXEmpresas.InscricaoMunicipal, Clientes.Suframa" & vbCrLf &
                          "  FROM NotasFiscais" & vbCrLf &
                          " INNER JOIN NotasFiscaisXItens" & vbCrLf &
                          "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                          "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                          "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                          "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                          "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                          "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                          "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                          " INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                          "    ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id " & vbCrLf &
                          "   AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                          "   AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                          "   AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                          "   AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                          "   AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                          "   AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id " & vbCrLf &
                          "   AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                          "   AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                          "   AND NotasFiscaisXItens.Sequencia_id    = NotasFiscaisXEncargos.Sequencia_id " & vbCrLf &
                          "   AND NotasFiscaisXEncargos.Encargo_Id   = 'PIS'" & vbCrLf &
                          "   AND NotasFiscaisXEncargos.Valor        > 0" & vbCrLf &
                          " INNER JOIN SubOperacoes" & vbCrLf &
                          "    ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                          "   AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                          " INNER JOIN Clientes" & vbCrLf &
                          "    ON NotasFiscais.Empresa_Id    = Clientes.Cliente_Id" & vbCrLf &
                          "   AND NotasFiscais.EndEmpresa_Id = Clientes.Endereco_Id " & vbCrLf &
                          " INNER JOIN ClientesXEmpresas" & vbCrLf &
                          "    ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id" & vbCrLf &
                          "   AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id " & vbCrLf &
                          " WHERE (NotasFiscais.Serie_Id           <> 'D' OR NotasFiscais.TipoDedocumento IN(9))" & vbCrLf &
                          "   and left(NotasFiscais.Empresa_Id,8) = '" & EmpresaMestre & "'" & vbCrLf &
                          "   And NotasFiscais.Movimento           BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'" & vbCrLf &
                          "   AND SubOperacoes.Classe             = '" & eClassesOperacoes.SERVICOS.ToString & "'" & vbCrLf &
                          "   And NotasFiscais.Situacao           = 1" & vbCrLf


                    ds = Banco.ConsultaDataSet(Sql, "Consulta")

                    linha = "|A001"

                    If ds.Tables(0).Rows.Count > 0 Then
                        linha &= "|" & "0"
                    Else
                        linha &= "|" & "1"
                    End If

                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroA001 += 1
                    RegistroGeral += 1


                    ' Registro A010  - Identificação do Estabelecimento bloco A
                    For Each dr010 In ds.Tables(0).Rows
                        With dr010

                            linha = "|A010"
                            linha &= "|" & .Item("Cliente_Id")
                            linha &= "|"

                            ArquivoAux.Add(linha)
                            RegistroA010 += 1
                            RegistroGeral += 1

                            'Registro A100
                            Sql = "SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Situacao, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id," & vbCrLf &
                                  "       NFE.ChaveNfe, NotasFiscais.DataDaNota, NotasFiscais.Movimento," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO'   THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS Valor," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.')       THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END ELSE 0 END) AS BasePis," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.')       THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END ELSE 0 END) AS ValorPis," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END ELSE 0 END) AS BaseCofins," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END ELSE 0 END) AS ValorCofins," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS_RF'    THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorPisRetFonte," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS_RF' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorCofinsRetFonte," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ISS'       THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorISS," & vbCrLf &
                                  "       SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTOS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ValorDesconto" & vbCrLf &
                                  "  FROM NotasFiscais " & vbCrLf &
                                  " INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                  "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                  "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                                  "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                  "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf &
                                  "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                  "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                  "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                                  " INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                  "    ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id " & vbCrLf &
                                  "   AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                  "   AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id " & vbCrLf &
                                  "   AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                  "   AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                  "   AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                  "   AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id " & vbCrLf &
                                  "   AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                  "   AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id " & vbCrLf &
                                  "   AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisxEncargos.Sequencia_Id" & vbCrLf &
                                  " INNER JOIN SubOperacoes" & vbCrLf &
                                  "    ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id " & vbCrLf &
                                  "   AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                                  "  LEFT OUTER JOIN NFERealizadas AS NFE" & vbCrLf &
                                  "    ON NotasFiscais.Empresa_Id      = NFE.Empresa_Id" & vbCrLf &
                                  "   AND NotasFiscais.EndEmpresa_Id   = NFE.EndEmpresa_Id " & vbCrLf &
                                  "   AND NotasFiscais.Cliente_Id      = NFE.Cliente_Id" & vbCrLf &
                                  "   AND NotasFiscais.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf &
                                  "   AND NotasFiscais.EntradaSaida_Id = NFE.EntradaSaida_Id " & vbCrLf &
                                  "   AND NotasFiscais.Serie_Id        = NFE.Serie_Id" & vbCrLf &
                                  "   AND NotasFiscais.Nota_Id         = NFE.Nota_Id" & vbCrLf &
                                  " WHERE NotasFiscais.Empresa_Id ='" & .Item("Cliente_Id") & "'" & vbCrLf &
                                  "   and (NotasFiscais.Movimento  BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                  "   And NotasFiscais.Situacao = 1" & vbCrLf

                            If (EmpresaMestre = "05366261") OrElse (EmpresaMestre = "62747840") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "49673784") OrElse (EmpresaMestre = "53267147") OrElse (EmpresaMestre = "62780383") OrElse (EmpresaMestre = "63358210") Then
                                Sql &= "   AND NotasFiscais.TipoDeDocumento = 9" & vbCrLf &
                                       "   AND (NotasFiscaisXItens.Cfop_Id IN (1933,2933,5933,6933))" & vbCrLf
                            Else
                                Sql &= "   AND (SubOperacoes.Classe = '" & eClassesOperacoes.SERVICOS.ToString & "')" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 1252 AND 1255" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 1352 AND 1355" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 2252 AND 2255" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 2352 AND 2355" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 1555 AND 1555" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 2555 AND 2555" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 1555 AND 1556" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 2555 AND 2556" & vbCrLf &
                                "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 1653 AND 1653" & vbCrLf &
                                "   And (NotasFiscaisXItens.Cfop_Id NOT IN (5206))" & vbCrLf
                            End If

                            Sql &= " GROUP BY NotasFiscais.EntradaSaida_Id,NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Situacao," & vbCrLf &
                                  "          NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NFE.ChaveNfe, NotasFiscais.DataDaNota,NotasFiscais.Movimento " & vbCrLf &
                                  " having SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) <> 0" & vbCrLf

                            ds = Banco.ConsultaDataSet(Sql, "Clientes")

                            If ds.Tables(0).Rows.Count > 0 Then
                                For Each drA100 In ds.Tables(0).Rows
                                    With drA100

                                        Nota = .Item("Nota_Id")
                                        Serie = .Item("Serie_Id")
                                        EntradaSaida = .Item("EntradaSaida_Id")
                                        Cliente = Trim(.Item("Cliente_Id"))
                                        EndCliente = .Item("EndCliente_Id")


                                        linha = "|A100"
                                        If .Item("EntradaSaida_Id") = "E" Then
                                            linha &= "|" & "0"
                                        Else
                                            linha &= "|" & "1"
                                        End If
                                        If .Item("EntradaSaida_Id") = "E" Then
                                            linha &= "|" & "1"
                                        Else
                                            linha &= "|" & "0"
                                        End If
                                        linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                                        If .Item("Situacao") = "1" Then
                                            linha &= "|" & "00"
                                        Else
                                            linha &= "|" & "01"
                                        End If
                                        linha &= "|" & .Item("Serie_Id")
                                        linha &= "|" & ""
                                        linha &= "|" & .Item("Nota_Id")
                                        linha &= "|" & .Item("ChaveNfe")
                                        linha &= "|" & CDate(.Item("DataDaNota")).ToStrDate().Replace("-", "")
                                        linha &= "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                                        linha &= "|" & .Item("Valor")
                                        linha &= "|" & "1"
                                        linha &= "|" & .Item("ValorDesconto")
                                        linha &= "|" & .Item("BasePis")
                                        linha &= "|" & .Item("ValorPis")
                                        linha &= "|" & .Item("BaseCofins")
                                        linha &= "|" & .Item("ValorCofins")
                                        linha &= "|" & .Item("ValorPisRetFonte")
                                        linha &= "|" & .Item("ValorCofinsRetFonte")
                                        linha &= "|" & .Item("ValorISS")
                                        linha &= "|"

                                        ArquivoAux.Add(linha)
                                        RegistroA100 += 1
                                        RegistroGeral += 1
                                    End With

                                    '' Bloco A registro A170
                                    Sql = "SELECT nfi.Produto_Id," & vbCrLf &
                                          "       Produtos.Nome," & vbCrLf &
                                          "       nfi.Valor AS VL_Item," & vbCrLf &
                                          "       ISNULL(SUM(CASE WHEN Nfe.Encargo_Id = 'Descontos' THEN Nfe.Valor ELSE 0 END), 0) AS VL_Desc," & vbCrLf &
                                          "		  '03' AS Nat_BC_Cred," & vbCrLf &
                                          "       0 AS Ind_Orig_Cred," & vbCrLf &
                                          "       ISNULL(OE.STPISCOFINS, 0) AS CST_PIS," & vbCrLf &
                                          "       ISNULL(SUM(CASE WHEN Nfe.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(Nfe.BaseNova,0) > 0 then Nfe.BaseNova else Nfe.Base END ELSE 0 END), 0) AS VL_Bc_Pis," & vbCrLf &
                                          "       ISNULL(SUM(CASE WHEN Nfe.Encargo_Id IN ('PIS','PIS RECUP.') THEN Nfe.Percentual ELSE 0 END), 0) AS Aliq_Cred," & vbCrLf &
                                          "       ISNULL(SUM(CASE WHEN Nfe.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(Nfe.ValorNovo,0) > 0 then Nfe.ValorNovo else Nfe.Valor END ELSE 0 END), 0) AS VL_Pis," & vbCrLf &
                                          "       ISNULL(OE.STPISCOFINS, 0) AS CST_Cofins," & vbCrLf &
                                          "       ISNULL(SUM(CASE WHEN Nfe.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(Nfe.BaseNova,0) > 0 then Nfe.BaseNova else Nfe.Base END ELSE 0 END), 0) AS VL_Bc_Cofins," & vbCrLf &
                                          "       ISNULL(SUM(CASE WHEN Nfe.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN Nfe.Percentual ELSE 0 END), 0) AS Aliq_Cofins," & vbCrLf &
                                          "       ISNULL(SUM(CASE WHEN Nfe.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(Nfe.ValorNovo,0) > 0 then Nfe.ValorNovo else Nfe.Valor END ELSE 0 END), 0) AS VL_Cofins," & vbCrLf &
                                          "		  case" & vbCrLf &
                                          "			when NF.EntradaSaida_Id = 'E'" & vbCrLf &
                                          "				then OEE.DebitaConta" & vbCrLf &
                                          "				else OEE.CreditaConta" & vbCrLf &
                                          "		  end AS Cod_Cta," & vbCrLf &
                                          "       '' AS Cod_CCus," & vbCrLf &
                                          "       NF.Nota_Id," & vbCrLf &
                                          "       NF.EntradaSaida_Id as EntradaSaida" & vbCrLf &
                                          "  FROM NotasFiscais NF" & vbCrLf &
                                          " INNER JOIN NotasFiscaisXItens nfi" & vbCrLf &
                                          "    ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf &
                                          "   AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf &
                                          "   AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf &
                                          "   AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf &
                                          "   AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf &
                                          "   AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf &
                                          "   AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf &
                                          " inner join OperacaoXEstado OE" & vbCrLf &
                                          "    on OE.Codigo_id = nfi.OperacaoxEstado" & vbCrLf &
                                          "  inner join OperacaoXEstadoXEncargo OEE" & vbCrLf &
                                          "     on OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
                                          "    and OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
                                          " INNER JOIN NotasFiscaisXEncargos Nfe" & vbCrLf &
                                          "    ON nfi.Empresa_Id      = Nfe.Empresa_Id" & vbCrLf &
                                          "   AND nfi.EndEmpresa_Id   = Nfe.EndEmpresa_Id" & vbCrLf &
                                          "   AND nfi.Cliente_Id      = Nfe.Cliente_Id" & vbCrLf &
                                          "   AND nfi.EndCliente_Id   = Nfe.EndCliente_Id" & vbCrLf &
                                          "   AND nfi.EntradaSaida_Id = Nfe.EntradaSaida_Id" & vbCrLf &
                                          "   AND nfi.Serie_Id        = Nfe.Serie_Id" & vbCrLf &
                                          "   AND nfi.Nota_Id         = Nfe.Nota_Id" & vbCrLf &
                                          "   AND nfi.Produto_Id      = Nfe.Produto_Id" & vbCrLf &
                                          "   AND nfi.CFOP_Id         = Nfe.CFOP_Id" & vbCrLf &
                                          "   AND nfi.Sequencia_id    = Nfe.Sequencia_id" & vbCrLf &
                                          " INNER JOIN Produtos" & vbCrLf &
                                          "    ON nfi.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                                          " INNER JOIN SubOperacoes SO" & vbCrLf &
                                          "    ON nfi.Operacao    = SO.Operacao_Id" & vbCrLf &
                                          "   AND nfi.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                                          " INNER JOIN Clientes AS Empresa" & vbCrLf &
                                          "    ON NF.Empresa_Id    = Empresa.Cliente_Id" & vbCrLf &
                                          "   AND NF.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf &
                                          " INNER JOIN Clientes" & vbCrLf &
                                          "    ON NF.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
                                          "   AND NF.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                                          "  LEFT OUTER JOIN Estados" & vbCrLf &
                                          "    ON Clientes.Estado = Estados.Estado_Id" & vbCrLf &
                                          " WHERE NF.Empresa_Id      ='" & .Item("Cliente_Id") & "'" & vbCrLf &
                                          "   AND NF.Movimento       BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'" & vbCrLf &
                                          "   AND NF.Cliente_Id      ='" & Cliente & "'" & vbCrLf &
                                          "   AND NF.EndCliente_Id   = " & EndCliente & vbCrLf &
                                          "   And NF.Nota_Id         = " & Nota & vbCrLf &
                                          "   AND NF.Serie_Id        ='" & Serie & "'" & vbCrLf &
                                          "   And NF.Situacao        = 1" & vbCrLf &
                                          "   And NF.EntradaSaida_Id ='" & EntradaSaida & "'" & vbCrLf

                                    If (EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "49673784") OrElse (EmpresaMestre = "62747840") OrElse (EmpresaMestre = "62780383") OrElse (EmpresaMestre = "63358210") Then
                                        Sql &= "   AND NF.TipoDeDocumento = 9" & vbCrLf &
                                                "   AND (nfi.Cfop_Id IN (1933,2933,5933,6933))" & vbCrLf
                                    Else
                                        Sql &= "   AND so.Classe          ='" & eClassesOperacoes.SERVICOS.ToString & "'" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 1252 AND 1255" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 1352 AND 1355" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 2252 AND 2255" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 2352 AND 2355" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 1555 AND 1555" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 2555 AND 2555" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 1555 AND 1556" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 2555 AND 2556" & vbCrLf &
                                              "   And nfi.cfop_Id Not BETWEEN 1653 AND 1653" & vbCrLf &
                                              "   And (nfi.Cfop_Id NOT IN (5206))" & vbCrLf
                                    End If

                                    Sql &= " GROUP BY nfi.Produto_Id, Produtos.Nome, nfi.Valor, ISNULL(OE.STPISCOFINS, 0) , NF.Nota_Id, NF.EntradaSaida_Id, OEE.DebitaConta, OEE.CreditaConta" & vbCrLf &
                                          " having  ISNULL(SUM(CASE WHEN Nfe.Encargo_Id IN ('PIS','PIS RECUP.') THEN Nfe.Valor ELSE 0 END), 0) > 0" & vbCrLf &
                                          " ORDER BY NF.Nota_Id" & vbCrLf

                                    Dim Sequencia As Integer = 0

                                    ds = Banco.ConsultaDataSet(Sql, "Clientes")

                                    If ds.Tables(0).Rows.Count > 0 Then
                                        For Each drA170 In ds.Tables(0).Rows
                                            With drA170
                                                linha = "|A170"

                                                Sequencia += 1

                                                linha &= "|" & Sequencia
                                                linha &= "|" & .Item("Produto_Id")
                                                linha &= "|" & .Item("Nome")
                                                linha &= "|" & .Item("VL_Item")
                                                linha &= "|" & .Item("VL_Desc")

                                                If .Item("EntradaSaida") = "S" Then
                                                    linha &= "|"
                                                    linha &= "|"
                                                Else
                                                    linha &= "|" & .Item("Nat_BC_Cred")
                                                    linha &= "|" & .Item("Ind_Orig_Cred")
                                                End If

                                                linha &= "|" & Format(CInt(.Item("CST_Pis")), "00")
                                                linha &= "|" & .Item("Vl_Bc_Pis")
                                                linha &= "|" & CDbl(.Item("Aliq_Cred")).ToString("N2")
                                                linha &= "|" & .Item("VL_Pis")
                                                linha &= "|" & Format(CInt(.Item("CST_Cofins")), "00")
                                                linha &= "|" & .Item("VL_Bc_Cofins")
                                                linha &= "|" & CDbl(.Item("Aliq_Cofins")).ToString("N2")
                                                linha &= "|" & .Item("VL_Cofins")
                                                linha &= "|" & .Item("Cod_Cta")
                                                linha &= "|" & .Item("Cod_CCus")
                                                linha &= "|"


                                                ArquivoAux.Add(linha)
                                                RegistroA170 += 1
                                                RegistroGeral += 1
                                            End With
                                        Next
                                    End If


                                Next
                            End If
                            '' Fim da Identificação do Estabelecimento bloco A
                        End With
                    Next



                    '' Registro A990  - Encerramento do Bloco A
                    RegistroA990 += 1

                    linha = "|A990"
                    linha &= "|" & RegistroA001 + RegistroA010 + RegistroA100 + RegistroA170 + RegistroA990
                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroGeral += 1




                    ' Registro C001  - Abertura do Bloco C
                    Dim TemMovimento As String = "N"


                    'Colocados os registros da tabela NotasFiscaisXEncargos em uma tabela temporária para melhorar a performance da consulta'

                    Sql = "SELECT E.Encargo_Id, E.Empresa_Id, E.EndEmpresa_Id, E.Cliente_Id, E.EndCliente_Id, E.EntradaSaida_Id, " & vbCrLf &
                          "E.Serie_Id, E.Nota_Id, E.CFOP_Id, E.Sequencia_id, E.Base, E.Valor" & vbCrLf &
                          "INTO #Encargos" & vbCrLf &
                          "FROM  NotasFiscaisXEncargos E" & vbCrLf &
                          "WHERE E.Encargo_Id IN ('PIS', 'COFINS','PRODUTO')" & vbCrLf &
                          "  AND Left(E.Empresa_Id, 8)  = '" & EmpresaMestre & "'" & vbCrLf &
                          " " & vbCrLf &
                          " SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.NossaEmissao," & vbCrLf &
                          "        '' AS ChaveNfe, NotasFiscais.DataDaNota, NotasFiscais.Movimento," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'Produto' THEN #Encargos.Valor END, 0)) AS ValorTotal," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'ICMS'    THEN #Encargos.Base  END, 0)) AS BaseIcms," & vbCrLf &
                          "        0 AS Aliquota," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'ICMS'    THEN #Encargos.Valor END, 0)) AS ValorIcms," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'IPI'     THEN #Encargos.Valor END, 0)) AS ValorIPI," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'PIS'     THEN #Encargos.Valor END, 0)) AS ValorPIS," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'COFINS'  THEN #Encargos.Valor END, 0)) AS ValorCofins," & vbCrLf &
                          "        MAX(ISNULL(CASE WHEN Encargo_Id = 'PRODUTO' THEN NotasFiscaisXItens.Cfop_Id  END, 0)) AS Cfop," & vbCrLf &
                          "        ISNULL(NotasFiscais.Situacao, 1) AS Situacao" & vbCrLf &
                          "   FROM NotasFiscais" & vbCrLf &
                          "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                          "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                          "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                          "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                          "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                          "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                          "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf &
                          "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                          "  INNER JOIN #Encargos" & vbCrLf &
                          "     ON NotasFiscaisXItens.Empresa_Id      = #Encargos.Empresa_Id " & vbCrLf &
                          "    AND NotasFiscaisXItens.EndEmpresa_Id   = #Encargos.EndEmpresa_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.Cliente_Id      = #Encargos.Cliente_Id " & vbCrLf &
                          "    AND NotasFiscaisXItens.EndCliente_Id   = #Encargos.EndCliente_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.EntradaSaida_Id = #Encargos.EntradaSaida_Id " & vbCrLf &
                          "    AND NotasFiscaisXItens.Serie_Id        = #Encargos.Serie_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.Nota_Id         = #Encargos.Nota_Id " & vbCrLf &
                          "    AND NotasFiscaisXItens.CFOP_Id         = #Encargos.CFOP_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.Sequencia_id    = #Encargos.Sequencia_id" & vbCrLf &
                          "  INNER JOIN SubOperacoes" & vbCrLf &
                          "     ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                          "  WHERE isnull(NotasFiscais.Situacao,1) = 1" & vbCrLf &
                          "    AND Left(NotasFiscais.Empresa_Id, 8)  = '" & EmpresaMestre & "'" & vbCrLf &
                          "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                          "    AND (#Encargos.Encargo_Id IN ('PIS', 'COFINS','PRODUTO'))" & vbCrLf &
                          "    And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "'" & vbCrLf &
                          "    And NotasFiscais.EntradaSaida_ID  = 'E'" & vbCrLf &
                          "    And (NotasFiscaisXItens.Cfop_Id Not Between 1350 and 1360)" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 2350 and 2360)" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not IN (1932,2932))" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 5350 and 5360)" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 6350 and 6360)" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 1302 and 1303)" & vbCrLf &
                          "  GROUP BY NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.NossaEmissao," & vbCrLf &
                          "           NotasFiscais.DataDaNota, NotasFiscais.Movimento, NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Situacao" & vbCrLf &
                          " having SUM(ISNULL(CASE WHEN Encargo_Id = 'PIS' THEN #Encargos.Valor END, 0)) >0" & vbCrLf &
                          " Union" & vbCrLf &
                          " SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.NossaEmissao," & vbCrLf &
                          "        '' AS ChaveNfe, NotasFiscais.DataDaNota, NotasFiscais.Movimento," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'Produto' THEN #Encargos.Valor END, 0)) AS ValorTotal," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'ICMS'    THEN #Encargos.Base  END, 0)) AS BaseIcms," & vbCrLf &
                          "        0 AS Aliquota," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'ICMS'    THEN #Encargos.Valor END, 0)) AS ValorIcms," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'IPI'     THEN #Encargos.Valor END, 0)) AS ValorIPI," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'PIS'     THEN #Encargos.Valor END, 0)) AS ValorPIS," & vbCrLf &
                          "        SUM(ISNULL(CASE WHEN Encargo_Id = 'COFINS'  THEN #Encargos.Valor END, 0)) AS ValorCofins," & vbCrLf &
                          "        MAX(ISNULL(CASE WHEN Encargo_Id = 'PRODUTO' THEN NotasFiscaisXItens.Cfop_Id END, 0))  AS Cfop," & vbCrLf &
                          "        ISNULL(NotasFiscais.Situacao, 1) AS Situacao" & vbCrLf &
                          "   FROM NotasFiscais" & vbCrLf &
                          "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                          "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                          "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                          "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                          "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                          "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                          "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                          "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                          "  INNER JOIN #Encargos" & vbCrLf &
                          "     ON NotasFiscaisXItens.Empresa_Id      = #Encargos.Empresa_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.EndEmpresa_Id   = #Encargos.EndEmpresa_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.Cliente_Id      = #Encargos.Cliente_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.EndCliente_Id   = #Encargos.EndCliente_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.EntradaSaida_Id = #Encargos.EntradaSaida_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.Serie_Id        = #Encargos.Serie_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.Nota_Id         = #Encargos.Nota_Id " & vbCrLf &
                          "    AND NotasFiscaisXItens.CFOP_Id         = #Encargos.CFOP_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.Sequencia_id    = #Encargos.Sequencia_id" & vbCrLf &
                          "  INNER JOIN SubOperacoes" & vbCrLf &
                          "     ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                          "    AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                          "  WHERE isnull(NotasFiscais.Situacao,1) = 1" & vbCrLf &
                          "    AND Left(NotasFiscais.Empresa_Id, 8)  = '" & EmpresaMestre & "'" & vbCrLf &
                          "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                          "    AND (#Encargos.Encargo_Id IN ('PIS', 'COFINS','PRODUTO'))" & vbCrLf &
                          "    And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "'" & vbCrLf &
                          "    And NotasFiscais.EntradaSaida_ID  = 'S'" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 1350 and 1360)" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 2350 and 2360)" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not IN (1932,2932))" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 5350 and 5360)" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 6350 and 6360)" & vbCrLf &
                          "    and (NotasFiscaisXItens.Cfop_Id Not Between 1302 and 1303)" & vbCrLf &
                          "  GROUP BY NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.NossaEmissao," & vbCrLf &
                          "           NotasFiscais.DataDaNota, NotasFiscais.Movimento, NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Situacao" & vbCrLf

                    ds = Banco.ConsultaDataSet(Sql, "Clientes")

                    RegistroC001 += 1

                    linha = "|C001"
                    If ds.Tables(0).Rows.Count > 0 Then
                        linha &= "|" & "0"
                        TemMovimento = "S"
                    Else
                        linha &= "|" & "1"
                        TemMovimento = "N"
                    End If

                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroGeral += 1


                    ' Registro C010  - Identificação do Estabelecimento bloco C
                    If TemMovimento = "S" Then

                        Sql = "SELECT Clientes.Reduzido, Clientes.Nome, Clientes.Cliente_Id, Clientes.Estado, Clientes.Inscricao, Clientes.CodigoDoMunicipio, ClientesXEmpresas.InscricaoMunicipal, Clientes.Suframa" & vbCrLf &
                              "  FROM ClientesXEmpresas " & vbCrLf &
                              " RIGHT OUTER JOIN Clientes" & vbCrLf &
                              "    ON ClientesXEmpresas.Empresa_Id   = Clientes.Cliente_Id " & vbCrLf &
                              "   AND ClientesXEmpresas.EndEmpresa_Id = Clientes.Endereco_Id " & vbCrLf &
                              " RIGHT OUTER JOIN NotasFiscais" & vbCrLf &
                              "    ON Clientes.Cliente_Id  = NotasFiscais.Empresa_Id" & vbCrLf &
                              "   AND Clientes.Endereco_Id = NotasFiscais.EndEmpresa_Id" & vbCrLf &
                              "  LEFT OUTER JOIN SubOperacoes" & vbCrLf &
                              "    ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                              "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                              " WHERE (NOT (NotasFiscais.Serie_Id      IN ('D', 'F')))" & vbCrLf &
                              "   and (LEFT(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & " ') " & vbCrLf &
                              "   And NotasFiscais.Movimento      BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'" & vbCrLf &
                              " GROUP BY Clientes.Reduzido, Clientes.Nome, Clientes.Cliente_Id, Clientes.Estado, Clientes.Inscricao, Clientes.CodigoDoMunicipio, ClientesXEmpresas.InscricaoMunicipal, Clientes.Suframa" & vbCrLf

                        ds = Banco.ConsultaDataSet(Sql, "Clientes")

                        For Each drC010 In ds.Tables(0).Rows
                            With drC010

                                Dim EmpresaBlocoC010 As String = .Item("Cliente_Id")

                                linha = "|C010"
                                linha &= "|" & EmpresaBlocoC010

                                If rbAnalitico.Checked Then
                                    linha &= "|" & "2"
                                Else
                                    linha &= "|" & "1"
                                End If

                                linha &= "|"

                                ArquivoAux.Add(linha)
                                RegistroC010 += 1
                                RegistroGeral += 1

                                Dim dsAux As New DataSet

                                Dim codItem As String

                                If rbConsolidado.Checked Then
                                    '' c180 Consolidação de Notas Fiscais Eletronicas

                                    Sql = " SELECT 'C180' AS Reg," & vbCrLf &
                                          "        55 AS Cod_Mod," & vbCrLf &
                                          "        '" & PeriodoInicial & "' AS DT_Doc_Ini," & vbCrLf &
                                          "        '" & PeriodoFinal & "'   AS DT_Doc_Fin," & vbCrLf &
                                          "        NotasFiscaisXItens.Produto_Id AS Cod_Item," & vbCrLf &
                                          "        Produtos.NCM AS COd_Ncm," & vbCrLf &
                                          "        ISNULL(Produtos.CodigoEX, N'') AS EX_IPI," & vbCrLf &
                                          "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PRODUTO','FRETES','SEGURO') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Tot_Item " & vbCrLf &
                                          "   FROM NotasFiscais " & vbCrLf &
                                          "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                          "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                          "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                                          "  INNER JOIN Produtos" & vbCrLf &
                                          "     ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                                          "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                          "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Sequencia_id    = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
                                          "  INNER JOIN SubOperacoes" & vbCrLf &
                                          "     ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                                          "  inner join OperacaoXEstado OE" & vbCrLf &
                                          "     on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                                          "  WHERE (NotasFiscais.EntradaSaida_Id = 'S')" & vbCrLf &
                                          "    AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                          "    AND (NotasFiscais.Empresa_Id= '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                          "    AND (NotasFiscaisXEncargos.Encargo_Id IN ('PIS', 'COFINS','FRETES','SEGURO','PRODUTO'))" & vbCrLf &
                                          "    AND not SubOperacoes.Classe in ('" & eClassesOperacoes.SERVICOS.ToString & "', '" & eClassesOperacoes.CONTAEORDEM.ToString & "', '" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "', '" & eClassesOperacoes.DOACAO.ToString & "') " & vbCrLf &
                                          "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5150 AND 5160)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6150 AND 6160)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf &
                                          "    AND NotasFiscaisXEncargos.Valor > 0" & vbCrLf &
                                          "    AND OE.STPISCOFINS <> 49" & vbCrLf &
                                          "  GROUP BY NotasFiscaisXItens.Produto_Id, Produtos.NCM, Produtos.CodigoEX" & vbCrLf



                                    dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                    For Each drC180 In dsAux.Tables(0).Rows
                                        With drC180
                                            codItem = .Item("COd_Item")
                                            linha = "|" & .Item("Reg")
                                            linha &= "|" & .Item("Cod_Mod")
                                            linha &= "|" & CDate(.Item("DT_Doc_Ini")).ToStrDate().Replace("-", "")
                                            linha &= "|" & CDate(.Item("DT_Doc_Fin")).ToStrDate().Replace("-", "")
                                            linha &= "|" & .Item("COd_Item")
                                            linha &= "|" & .Item("COd_Ncm").ToString.Replace(".", "")
                                            linha &= "|" & .Item("EX_IPI")
                                            linha &= "|" & .Item("Vl_Tot_Item")
                                            linha &= "|"
                                        End With

                                        ArquivoAux.Add(linha)
                                        RegistroC180 += 1
                                        RegistroGeral += 1

                                        '' c181 Detalhamento da consolidação - Operaçoes de Vendas PIS/Pasep
                                        Sql = " SELECT 'C181' AS Reg," & vbCrLf &
                                              "        CST_PIS," & vbCrLf &
                                              "        CFOP, SUM(Vl_Item) AS Vl_Item," & vbCrLf &
                                              "        SUM(vL_DESC) AS vL_DESC," & vbCrLf &
                                              "        SUM(VL_BC_Pis) AS VL_BC_Pis," & vbCrLf &
                                              "        SUM(Aliq_Pis) AS Aliq_Pis," & vbCrLf &
                                              "        Quant_Bc_Pis," & vbCrLf &
                                              "        Aliq_Pis_Quant," & vbCrLf &
                                              "        SUM(VL_Pis) AS VL_Pis," & vbCrLf &
                                              "        Cod_Cta" & vbCrLf &
                                              "  From (" & vbCrLf &
                                              "        SELECT 'C181' AS Reg," & vbCrLf &
                                              "               ISNULL(OE.STPISCOFINS, 0) AS CST_PIS, NotasFiscaisXItens.CFOP_Id AS CFOP, " & vbCrLf &
                                              "               SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PRODUTO','FRETES','SEGURO') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Item," & vbCrLf &
                                              "               SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTO' THEN NotasFiscaisXEncargos.Valor END, 0)) AS vL_DESC," & vbCrLf &
                                              "               SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id in('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS VL_BC_Pis," & vbCrLf &
                                              "               ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id in('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Percentual END, 0) AS Aliq_Pis, 0 AS Quant_Bc_Pis," & vbCrLf &
                                              "               0 AS Aliq_Pis_Quant," & vbCrLf &
                                              "               SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id in('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END END, 0)) AS VL_Pis," & vbCrLf &
                                              "               SubOperacoes.GrupoDeContas AS Cod_Cta" & vbCrLf &
                                              "          FROM NotasFiscais" & vbCrLf &
                                              "         INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                              "            ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                              "           AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                              "           AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                              "           AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                              "           AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                              "           AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                              "           AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                                              "         INNER JOIN Produtos" & vbCrLf &
                                              "            ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                                              "         INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                              "            ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
                                              "         INNER JOIN SubOperacoes" & vbCrLf &
                                              "            ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                                              "           AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                                              "         inner join OperacaoXEstado OE" & vbCrLf &
                                              "           on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                                              "         WHERE NotasFiscaisXItens.Produto_Id = '" & codItem & "'" & vbCrLf &
                                              "           AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                              "           AND (NotasFiscaisXEncargos.Valor > 0)" & vbCrLf &
                                              "           And (NotasFiscais.Empresa_Id = '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                              "           AND (NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.','FRETES','SEGURO','PRODUTO'))" & vbCrLf &
                                              "           And not SubOperacoes.Classe in ('" & eClassesOperacoes.SERVICOS.ToString & "', '" & eClassesOperacoes.CONTAEORDEM.ToString & "', '" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "')  " & vbCrLf &
                                              "           AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                              "           And (NotasFiscais.EntradaSaida_Id = 'S')" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5150 AND 5160)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6150 AND 6160)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6910 AND 6920)" & vbCrLf &
                                              "           and (NotasFiscaisXItens.Cfop_Id Not Between 5910 and 5920)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
                                              "           and (NotasFiscaisXItens.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
                                              "           and (NotasFiscaisXItens.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
                                              "           and (NotasFiscaisXItens.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
                                              "           AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf &
                                              "           AND OE.STPISCOFINS <> 49" & vbCrLf &
                                              "         GROUP BY OE.STPISCOFINS, NotasFiscaisXEncargos.Encargo_Id," & vbCrLf &
                                              "                  NotasFiscaisXItens.CFOP_Id, NotasFiscaisXEncargos.Percentual , SubOperacoes.GrupoDeContas" & vbCrLf &
                                              "        ) AS Consulta" & vbCrLf &
                                              "   GROUP BY CST_PIS, CFOP, Quant_Bc_Pis, Aliq_Pis_Quant, Cod_Cta" & vbCrLf

                                        dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                        For Each drC181 In dsAux.Tables(0).Rows
                                            With drC181
                                                linha = "|" & .Item("Reg")
                                                linha &= "|" & Format(CInt(.Item("Cst_Pis")), "00")
                                                linha &= "|" & .Item("Cfop")
                                                linha &= "|" & .Item("Vl_Item")
                                                linha &= "|" & .Item("Vl_Desc")
                                                linha &= "|" & .Item("Vl_Bc_Pis")
                                                linha &= "|" & CDbl(.Item("Aliq_Pis")).ToString("N2")
                                                linha &= "|" ''& .Item("Quant_Bc_Pis") ver na usina
                                                linha &= "|" ''& .Item("Aliq_Pis_Quant")
                                                linha &= "|" & .Item("Vl_Pis")
                                                linha &= "|" & .Item("Cod_Cta")
                                                linha &= "|"
                                            End With

                                            ArquivoAux.Add(linha)
                                            RegistroC181 += 1
                                            RegistroGeral += 1
                                        Next


                                        '' c185 Detalhamento da consolidação - Operaçoes de Vendas Cofins

                                        Sql = " SELECT 'C185' AS Reg, " & vbCrLf &
                                              "        CST_PIS," & vbCrLf &
                                              "        CFOP, SUM(Vl_Item) AS Vl_Item," & vbCrLf &
                                              "        SUM(vL_DESC) AS vL_DESC," & vbCrLf &
                                              "        SUM(VL_BC_Pis) AS VL_BC_Pis," & vbCrLf &
                                              "        SUM(Aliq_Pis) AS Aliq_Pis," & vbCrLf &
                                              "        Quant_Bc_Pis," & vbCrLf &
                                              "        Aliq_Pis_Quant," & vbCrLf &
                                              "        SUM(VL_Pis) AS VL_Pis," & vbCrLf &
                                              "        Cod_Cta" & vbCrLf &
                                              " From (" & vbCrLf &
                                              "       SELECT 'C185' AS Reg," & vbCrLf &
                                              "              ISNULL(OE.STPISCOFINS, 0) AS CST_PIS," & vbCrLf &
                                              "              NotasFiscaisXItens.CFOP_Id AS CFOP, " & vbCrLf &
                                              "              SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PRODUTO','FRETES','SEGURO') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Item," & vbCrLf &
                                              "              SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTO' THEN NotasFiscaisXEncargos.Valor END, 0)) AS vL_DESC," & vbCrLf &
                                              "              SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id in('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS VL_BC_Pis," & vbCrLf &
                                              "              ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id in('COFINS','COFINS RECUP.','COFINS A REC.') THEN NotasFiscaisXEncargos.Percentual END, 0) AS Aliq_Pis, 0 AS Quant_Bc_Pis," & vbCrLf &
                                              "              0 AS Aliq_Pis_Quant, SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id in('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END END, 0)) AS VL_Pis," & vbCrLf &
                                              "              SubOperacoes.GrupoDeContas AS Cod_Cta" & vbCrLf &
                                              "         FROM NotasFiscais" & vbCrLf &
                                              "        INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                              "           ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                              "          AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                              "          AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                              "          AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                              "          AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                              "          AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                              "          AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                                              "        INNER JOIN Produtos" & vbCrLf &
                                              "           ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                                              "        INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                              "           ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id " & vbCrLf &
                                              "          AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
                                              "        INNER JOIN SubOperacoes" & vbCrLf &
                                              "           ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                                              "          AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                                              "        inner join OperacaoXEstado OE" & vbCrLf &
                                              "           on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                                              "        WHERE NotasFiscaisXItens.Produto_Id = '" & codItem & "'" & vbCrLf &
                                              "          AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                              "          AND (NotasFiscaisXEncargos.Valor > 0)" & vbCrLf &
                                              "          AND (NotasFiscais.Empresa_Id = '" & EmpresaBlocoC010 & "') " & vbCrLf &
                                              "          AND (NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.','FRETES','SEGURO','PRODUTO'))" & vbCrLf &
                                              "          And not SubOperacoes.Classe in ('" & eClassesOperacoes.SERVICOS.ToString & "', '" & eClassesOperacoes.CONTAEORDEM.ToString & "', '" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "') " & vbCrLf &
                                              "          AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5150 AND 5160)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6150 AND 6160)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                              "          And (NotasFiscais.EntradaSaida_Id = 'S')" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6910 AND 6920)" & vbCrLf &
                                              "          and (NotasFiscaisXItens.Cfop_Id Not Between 5910 and 5920)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
                                              "          and (NotasFiscaisXItens.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
                                              "          and (NotasFiscaisXItens.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
                                              "          and (NotasFiscaisXItens.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
                                              "          AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf &
                                              "          AND OE.STPISCOFINS <> 49" & vbCrLf &
                                              "        GROUP BY OE.STPISCOFINS, NotasFiscaisXEncargos.Encargo_Id," & vbCrLf &
                                              "                 NotasFiscaisXItens.CFOP_Id, NotasFiscaisXEncargos.Percentual , SubOperacoes.GrupoDeContas" & vbCrLf &
                                              "       ) AS Consulta" & vbCrLf &
                                              " GROUP BY CST_PIS, CFOP, Quant_Bc_Pis, Aliq_Pis_Quant, Cod_Cta" & vbCrLf

                                        dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                        For Each drC185 In dsAux.Tables(0).Rows
                                            With drC185
                                                linha = "|" & .Item("Reg")
                                                linha &= "|" & Format(CInt(.Item("Cst_Pis")), "00")
                                                linha &= "|" & .Item("Cfop")
                                                linha &= "|" & .Item("Vl_Item")
                                                linha &= "|" & .Item("Vl_Desc")
                                                linha &= "|" & .Item("Vl_Bc_Pis")
                                                linha &= "|" & CDbl(.Item("Aliq_Pis")).ToString("N2")
                                                linha &= "|" ''& .Item("Quant_Bc_Pis")   ver na usina
                                                linha &= "|" ''& .Item("Aliq_Pis_Quant")
                                                linha &= "|" & .Item("Vl_Pis")
                                                linha &= "|" & .Item("Cod_Cta")
                                                linha &= "|"
                                            End With

                                            ArquivoAux.Add(linha)
                                            RegistroC185 += 1
                                            RegistroGeral += 1
                                        Next
                                    Next

                                    '' c190 Consolidação de Notas Fiscais Eletronicas - operacoes de compras  

                                    Sql = " SELECT 'C190' AS Reg," & vbCrLf &
                                          "        55 AS Cod_Mod, '" & PeriodoInicial & "' AS DT_Doc_Ini, '" & PeriodoFinal & "' AS DT_Doc_Fin, NotasFiscaisXItens.Produto_Id AS Cod_Item, Produtos.NCM AS COd_Ncm," & vbCrLf &
                                          "        ISNULL(Produtos.CodigoEX, N'') AS EX_IPI," & vbCrLf &
                                          "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS Vl_Tot_Item" & vbCrLf &
                                          "   FROM NotasFiscais" & vbCrLf &
                                          "   LEFT OUTER JOIN NFERealizadas AS NFE" & vbCrLf &
                                          "     ON NotasFiscais.Empresa_Id      = NFE.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndEmpresa_Id   = NFE.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.Cliente_Id      = NFE.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf &
                                          "    AND NotasFiscais.Serie_Id        = NFE.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscais.Nota_Id         = NFE.Nota_Id " & vbCrLf &
                                          "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                          "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                          "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                                          "  INNER JOIN Produtos" & vbCrLf &
                                          "     ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                                          "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                          "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id " & vbCrLf &
                                          "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
                                          "  INNER JOIN SubOperacoes" & vbCrLf &
                                          "     ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                                          "  inner join OperacaoXEstado OE" & vbCrLf &
                                          "     on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                                          "  WHERE (NotasFiscais.EntradaSaida_Id = 'E') " & vbCrLf &
                                          "    AND (NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.','COFINS','COFINS RECUP.','COFINS A REC.','PRODUTO')) " & vbCrLf &
                                          "    And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "' " & vbCrLf &
                                          "    AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                          "    AND (NotasFiscais.Empresa_Id = '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                          "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') " & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                           "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5550 AND 5555)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6550 AND 6555)" & vbCrLf &
                                          "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 3000 AND 4000)" & vbCrLf &
                                          "    AND Not OE.STPISCOFINS in (98, 99)" & vbCrLf &
                                          "  GROUP BY NotasFiscaisXItens.Produto_Id, Produtos.NCM, Produtos.CodigoEX" & vbCrLf &
                                          " having SUM(ISNULL(CASE WHEN Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor END, 0)) >0" & vbCrLf
                                    'Removido a pedido da Insol
                                    '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1653 AND 1653)" & vbCrLf & _
                                    '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1556 AND 1556)" & vbCrLf & _
                                    '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2653 AND 2653)" & vbCrLf & _
                                    '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2556 AND 2556)" & vbCrLf & _


                                    dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                    For Each drC190 In dsAux.Tables(0).Rows
                                        With drC190
                                            codItem = .Item("Cod_Item")

                                            linha = "|" & .Item("Reg")
                                            linha &= "|" & .Item("Cod_Mod")
                                            linha &= "|" & CDate(.Item("DT_Doc_Ini")).ToStrDate().Replace("-", "")
                                            linha &= "|" & CDate(.Item("DT_Doc_Fin")).ToStrDate().Replace("-", "")
                                            linha &= "|" & .Item("COd_Item")
                                            linha &= "|" & Left(.Item("COd_Ncm").ToString.Replace(".", ""), 8)
                                            linha &= "|" & .Item("EX_IPI")
                                            linha &= "|" & .Item("Vl_Tot_Item")
                                            linha &= "|"
                                        End With

                                        ArquivoAux.Add(linha)
                                        RegistroC190 += 1
                                        RegistroGeral += 1

                                        'Registro C191 - Pis -----------------------------------

                                        Sql = " SELECT 'C191' AS Reg," & vbCrLf &
                                              "        Case when Clientes.Estado = 'EX' then '99999999999999' else Clientes.Cliente_Id end AS CNPJ_CPF_PART, " & vbCrLf &
                                              "        ISNULL(OE.STPISCOFINS, 0) AS CST_PIS," & vbCrLf &
                                              "        NotasFiscaisXItens.CFOP_Id AS CFOP, " & vbCrLf &
                                              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS Vl_Item," & vbCrLf &
                                              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTO' THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_DESC, " & vbCrLf &
                                              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS VL_BC_Pis, " & vbCrLf &
                                              "        ISNULL(NotasFiscaisXEncargos.Percentual, 0) AS Aliq_Pis, 0 AS Quant_Bc_Pis, 0 AS Aliq_Pis_Quant, " & vbCrLf &
                                              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END END, 0)) AS VL_Pis, " & vbCrLf &
                                              "        SubOperacoes.GrupoDeContas AS Cod_Cta" & vbCrLf &
                                              "   FROM NotasFiscais" & vbCrLf &
                                              "   LEFT OUTER JOIN NfeRealizadas NFE" & vbCrLf &
                                              "     ON NotasFiscais.Empresa_Id = NFE.Empresa_Id" & vbCrLf &
                                              "    AND NotasFiscais.EndEmpresa_Id = NFE.EndEmpresa_Id" & vbCrLf &
                                              "    AND NotasFiscais.Cliente_Id = NFE.Cliente_Id" & vbCrLf &
                                              "    AND NotasFiscais.EndCliente_Id = NFE.EndCliente_Id" & vbCrLf &
                                              "    AND NotasFiscais.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf &
                                              "    AND NotasFiscais.Serie_Id = NFE.Serie_Id" & vbCrLf &
                                              "    AND NotasFiscais.Nota_Id = NFE.Nota_Id" & vbCrLf &
                                              "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                              "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                              "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                              "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                              "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                              "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                              "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                              "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                                              "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                              "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
                                              "  INNER JOIN Produtos" & vbCrLf &
                                              "     ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                                              "  INNER JOIN SubOperacoes" & vbCrLf &
                                              "     ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                                              "    AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                                              "  INNER JOIN Clientes" & vbCrLf &
                                              "     ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
                                              "    AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                                              "  inner join OperacaoXEstado OE" & vbCrLf &
                                              "     on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                                              "  WHERE (NotasFiscais.EntradaSaida_Id = 'E')" & vbCrLf &
                                              "    And  NotasFiscaisXItens.Produto_Id = '" & codItem & "'" & vbCrLf &
                                              "    AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                              "    AND (NotasFiscais.Empresa_Id = '" & EmpresaBlocoC010 & "') " & vbCrLf &
                                              "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') " & vbCrLf &
                                              "    AND (NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.','PRODUTO'))" & vbCrLf &
                                              "    And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "'  " & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5550 AND 5555)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6550 AND 6555)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 3000 AND 4000)" & vbCrLf &
                                              "    AND Not OE.STPISCOFINS in (98, 99)" & vbCrLf &
                                              "  GROUP BY Clientes.Cliente_Id, OE.STPISCOFINS, NotasFiscaisXItens.Cfop_Id," & vbCrLf &
                                              "           NotasFiscaisXEncargos.Percentual, SubOperacoes.GrupoDeContas, Clientes.Estado" & vbCrLf &
                                              " HAVING (SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor END, 0)) > 0)" & vbCrLf &
                                              "  ORDER BY CNPJ_CPF_PART, OE.STPISCOFINS, CFOP" & vbCrLf

                                        'Removido a pedido da Insol
                                        '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1653 AND 1653)" & vbCrLf & _
                                        '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1556 AND 1556)" & vbCrLf & _
                                        '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2653 AND 2653)" & vbCrLf & _
                                        '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2556 AND 2556)" & vbCrLf & _


                                        dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                        For Each drC191 In dsAux.Tables(0).Rows
                                            With drC191
                                                linha = "|" & .Item("Reg")
                                                linha &= "|" & .Item("CNPJ_CPF_PART")
                                                linha &= "|" & Format(CInt(.Item("Cst_Pis")), "00")
                                                linha &= "|" & .Item("Cfop")
                                                linha &= "|" & .Item("Vl_Item")
                                                linha &= "|" & .Item("Vl_Desc")
                                                linha &= "|" & .Item("Vl_Bc_Pis")

                                                If .item("Aliq_Pis") = 0.83 Then
                                                    linha = linha & "|0,8250"
                                                ElseIf .item("Aliq_Pis") = 0.58 Then
                                                    linha = linha & "|0,5775"
                                                Else
                                                    linha &= "|" & CDbl(.Item("Aliq_Pis")).ToString("N4")
                                                End If

                                                linha &= "|" ''& .Item("Quant_Bc_Pis")
                                                linha &= "|" ''& .Item("Aliq_Pis_Quant")
                                                linha &= "|" & .Item("Vl_Pis")
                                                linha &= "|" & .Item("Cod_Cta")
                                                linha &= "|"
                                            End With

                                            ArquivoAux.Add(linha)
                                            RegistroC191 += 1
                                            RegistroGeral += 1

                                        Next

                                        'Registro C195 Cofins --------------------------------------------------

                                        Sql = " SELECT 'C195' AS Reg, Case when Clientes.Estado = 'EX' then '99999999999999' else Clientes.Cliente_Id end AS CNPJ_CPF_PART, " & vbCrLf &
                                              "        ISNULL(OE.STPISCOFINS, 0) AS CST_PIS," & vbCrLf &
                                              "        NotasFiscaisXItens.CFOP_Id AS CFOP, " & vbCrLf &
                                              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS Vl_Item," & vbCrLf &
                                              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTO' THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_DESC, " & vbCrLf &
                                              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS VL_BC_Pis, " & vbCrLf &
                                              "        ISNULL(NotasFiscaisXEncargos.Percentual, 0) AS Aliq_Pis, 0 AS Quant_Bc_Pis, 0 AS Aliq_Pis_Quant, " & vbCrLf &
                                              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END END, 0)) AS VL_Pis, " & vbCrLf &
                                              "        SubOperacoes.GrupoDeContas AS Cod_Cta" & vbCrLf &
                                              "   FROM NotasFiscais " & vbCrLf &
                                              "   LEFT OUTER JOIN NfeRealizadas NFE" & vbCrLf &
                                              "     ON NotasFiscais.Empresa_Id = NFE.Empresa_Id" & vbCrLf &
                                              "    AND NotasFiscais.EndEmpresa_Id = NFE.EndEmpresa_Id" & vbCrLf &
                                              "    AND NotasFiscais.Cliente_Id = NFE.Cliente_Id  " & vbCrLf &
                                              "    AND NotasFiscais.EndCliente_Id = NFE.EndCliente_Id" & vbCrLf &
                                              "    AND NotasFiscais.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf &
                                              "    AND NotasFiscais.Serie_Id = NFE.Serie_Id" & vbCrLf &
                                              "    AND NotasFiscais.Nota_Id = NFE.Nota_Id " & vbCrLf &
                                              "   INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                              "      ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                              "     AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                              "     AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id " & vbCrLf &
                                              "     AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                              "     AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                              "     AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                              "     AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                                               "   INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                              "      ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                              "     AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id " & vbCrLf &
                                              "     AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                              "     AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                              "     AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                              "     AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id " & vbCrLf &
                                              "     AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                              "     AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id " & vbCrLf &
                                              "     AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                              "     AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
                                              "   INNER JOIN Produtos" & vbCrLf &
                                              "      ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id " & vbCrLf &
                                              "   INNER JOIN SubOperacoes" & vbCrLf &
                                              "      ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                                              "     AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                                              "   INNER JOIN Clientes" & vbCrLf &
                                              "      ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
                                              "     AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                                              "   inner join OperacaoXEstado OE" & vbCrLf &
                                              "      on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                                              "   WHERE NotasFiscaisXItens.Produto_Id = '" & codItem & "'" & vbCrLf &
                                              "     AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                              "     AND (NotasFiscais.Empresa_Id = '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                              "     AND (NotasFiscais.EntradaSaida_Id = 'E') " & vbCrLf &
                                              "     AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') " & vbCrLf &
                                              "     AND (NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.','PRODUTO'))" & vbCrLf &
                                              "     And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "' " & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5550 AND 5555)" & vbCrLf &
                                              "     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6550 AND 6555)" & vbCrLf &
                                              "    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 3000 AND 4000)" & vbCrLf &
                                              "    AND Not OE.STPISCOFINS in (98, 99)" & vbCrLf &
                                              "   GROUP BY Clientes.Cliente_Id, OE.STPISCOFINS, NotasFiscaisXItens.cfop_Id," & vbCrLf &
                                              "            NotasFiscaisXEncargos.Percentual, SubOperacoes.GrupoDeContas, Clientes.Estado" & vbCrLf &
                                              "  HAVING (SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NotasFiscaisXEncargos.Valor END, 0)) > 0)" & vbCrLf &
                                              "   ORDER BY CNPJ_CPF_PART, OE.STPISCOFINS, CFOP" & vbCrLf

                                        'Removido a pedido da Insol
                                        '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1653 AND 1653)" & vbCrLf & _
                                        '"    AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1556 AND 1556)" & vbCrLf & _
                                        '"     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2653 AND 2653)" & vbCrLf & _
                                        '"     AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2556 AND 2556)" & vbCrLf & _

                                        dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                        For Each drC195 In dsAux.Tables(0).Rows
                                            With drC195
                                                linha = "|" & .Item("Reg")
                                                linha &= "|" & .Item("CNPJ_CPF_PART")
                                                linha &= "|" & Format(CInt(.Item("Cst_Pis")), "00")
                                                linha &= "|" & .Item("Cfop")
                                                linha &= "|" & .Item("Vl_Item")
                                                linha &= "|" & .Item("Vl_Desc")
                                                linha &= "|" & .Item("Vl_Bc_Pis")


                                                linha &= "|" & CDbl(.Item("Aliq_Pis")).ToString("N2")
                                                linha &= "|" ''& .Item("Quant_Bc_Pis")
                                                linha &= "|" ''& .Item("Aliq_Pis_Quant")
                                                linha &= "|" & .Item("Vl_Pis")
                                                linha &= "|" & .Item("Cod_Cta")
                                                linha &= "|"
                                            End With

                                            ArquivoAux.Add(linha)
                                            RegistroC195 += 1
                                            RegistroGeral += 1
                                        Next
                                    Next 'Fim Bloco 190 ------

                                ElseIf rbAnalitico.Checked Then

                                    'C100 - Documento Nota Fiscal Bloco C
                                    ''"    AND (ni.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf & _ Removi 5556 das Saídas cfe. Solicitação Luiz(Nutri) - Furlan - 09/08/2022

                                    Sql = "-- C100 - SAIDAS" & vbCrLf &
                                    "" & vbCrLf &
                                    "SELECT 'C100' AS Reg," & vbCrLf &
                                    "		1 AS Ind_Oper," & vbCrLf &
                                    "		case " & vbCrLf &
                                    "		    when ISNULL(DocXML.Tipo,'') = 'Mic'" & vbCrLf &
                                    "		        then 0" & vbCrLf &
                                    "		        else case " & vbCrLf &
                                    "		                when n.NossaEmissao = 'S' " & vbCrLf &
                                    "		                then 0 " & vbCrLf &
                                    "		                else 1 " & vbCrLf &
                                    "		             end " & vbCrLf &
                                    "		    end AS Ind_Emit," & vbCrLf &
                                    "		n.Cliente_id," & vbCrLf &
                                    "		n.EndCliente_id," & vbCrLf &
                                    "        55 AS Cod_Mod," & vbCrLf &
                                    "		'00' AS Cod_Sit," & vbCrLf &
                                    "		n.Serie_id AS Ser," & vbCrLf &
                                    "		n.Nota_id AS Num_Doc," & vbCrLf &
                                    "		nFe.ChaveNfe AS Chv_Nfe," & vbCrLf &
                                    "		n.Movimento AS Dt_Doc," & vbCrLf &
                                    "		n.DataDaNota AS Dt_ES," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PRODUTO','FRETES','SEGURO','IPI') THEN ne.Valor END, 0) - ISNULL(CASE WHEN ne.Encargo_Id IN ('DESCONTOS') THEN ne.Valor END, 0)) AS Vl_Doc," & vbCrLf &
                                    "		case" & vbCrLf &
                                    "			when n.Movimento < '2020-01-01'" & vbCrLf &
                                    "				then 1" & vbCrLf &
                                    "				else" & vbCrLf &
                                    "				   case" & vbCrLf &
                                    "				      when isnull(pg.AVista,0) = 0" & vbCrLf &
                                    "					     then 1" & vbCrLf &
                                    "						 else 0" & vbCrLf &
                                    "				      end" & vbCrLf &
                                    "			end AS Ind_Pgto, " & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('DESCONTOS') THEN ne.Valor END, 0)) AS Vl_Desc," & vbCrLf &
                                    "		SUM(0) AS Vl_Abat_Nt, " & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PRODUTO') THEN ne.Valor END, 0)) AS Vl_Merc," & vbCrLf &
                                    "		case" & vbCrLf &
                                    "			when n.CifFob = 'CIF'" & vbCrLf &
                                    "				then '0'" & vbCrLf &
                                    "				else" & vbCrLf &
                                    "					case" & vbCrLf &
                                    "						when n.CifFob = 'FOB'" & vbCrLf &
                                    "						then '1'" & vbCrLf &
                                    "				        else" & vbCrLf &
                                    "					        case" & vbCrLf &
                                    "						    when n.CifFob = 'NEN'" & vbCrLf &
                                    "						    then '9'" & vbCrLf &
                                    "						    else '9'" & vbCrLf &
                                    "					        end" & vbCrLf &
                                    "					end" & vbCrLf &
                                    "				end Ind_Frt," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('FRETES') THEN ne.Valor END, 0)) AS Vl_Frt," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('SEGURO') THEN ne.Valor END, 0)) AS Vl_Seg," & vbCrLf &
                                    "		SUM(0) AS Vl_Out_Da, " & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN ne.Base END, 0)) AS Vl_Bc_Icms," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN ne.Valor END, 0)) AS Vl_Icms," & vbCrLf &
                                    "		SUM(0) AS Vl_Bc_Icms_St," & vbCrLf &
                                    "		SUM(0) AS Vl_Icms_St," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('IPI') THEN ne.Valor END, 0)) AS Vl_Ipi," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN ne.Valor END, 0)) AS Vl_Pis," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN ne.Valor END, 0)) AS Vl_Cofins," & vbCrLf &
                                    "		SUM(0) AS Vl_Pis_St," & vbCrLf &
                                    "		SUM(0) AS Vl_Cofins_St" & vbCrLf &
                                    " INTO #TempPisCofins" & vbCrLf &
                                    "   FROM NotasFiscais n" & vbCrLf &
                                    "  INNER JOIN NotasFiscaisXItens ni" & vbCrLf &
                                    "     ON ni.Empresa_Id      = n.Empresa_Id" & vbCrLf &
                                    "    AND ni.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                                    "    AND ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                                    "    AND ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                                    "    AND ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                                    "    AND ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                                    "    AND ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                                    "  INNER JOIN NotasFiscaisXEncargos ne" & vbCrLf &
                                    "     ON ne.Empresa_Id      = ni.Empresa_Id" & vbCrLf &
                                    "    AND ne.EndEmpresa_Id   = ni.EndEmpresa_Id" & vbCrLf &
                                    "    AND ne.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                                    "    AND ne.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                                    "    AND ne.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                                    "    AND ne.Serie_Id        = ni.Serie_Id" & vbCrLf &
                                    "    AND ne.Nota_Id         = ni.Nota_Id" & vbCrLf &
                                    "    AND ne.Produto_Id      = ni.Produto_Id" & vbCrLf &
                                    "    AND ne.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                                    "    AND ne.Sequencia_id    = ni.Sequencia_id" & vbCrLf &
                                    "  INNER JOIN Produtos" & vbCrLf &
                                    "     ON Produtos.Produto_Id = ni.Produto_Id" & vbCrLf &
                                    "  INNER JOIN SubOperacoes so" & vbCrLf &
                                    "     ON so.Operacao_Id     = ni.Operacao" & vbCrLf &
                                    "    AND so.SubOperacoes_Id = ni.SubOperacao" & vbCrLf &
                                    "  inner join OperacaoXEstado OE" & vbCrLf &
                                    "     on OE.Codigo_id = ni.OperacaoxEstado" & vbCrLf &
                                    "  left JOIN Pedidos p" & vbCrLf &
                                    "     ON p.Empresa_id     = n.Empresa_id" & vbCrLf &
                                    "    AND p.EndEmpresa_Id  = n.EndEmpresa_Id" & vbCrLf &
                                    "    AND p.Pedido_Id      = n.Pedido" & vbCrLf &
                                    "  left JOIN Pagamentos pg" & vbCrLf &
                                    "     ON pg.Pagamento_Id     = p.CondicaoPagamento" & vbCrLf &
                                    "  left JOIN NfeRealizadas nFe" & vbCrLf &
                                    "     ON nFe.Empresa_Id      = n.Empresa_Id" & vbCrLf &
                                    "    AND nFe.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                                    "    AND nFe.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                                    "    AND nFe.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                                    "    AND nFe.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                                    "    AND nFe.Serie_Id        = n.Serie_Id" & vbCrLf &
                                    "    AND nFe.Nota_Id         = n.Nota_Id" & vbCrLf &
                                    "  left JOIN DocumentoXML DocXML" & vbCrLf &
                                    "     ON DocXML.Empresa_Id = n.Empresa_Id" & vbCrLf &
                                    "    AND DocXML.Cliente_Id = n.Cliente_Id" & vbCrLf &
                                    "    AND DocXML.Serie_Id   = n.Serie_Id" & vbCrLf &
                                    "    AND DocXML.Numero_Id  = n.Nota_Id" & vbCrLf &
                                    "  WHERE (n.EntradaSaida_Id = 'S')"

                                    If EmpresaBlocoC010.Contains("44979506") Then
                                        Sql &= "    AND (ISNULL(n.Situacao, 1) = 1)" & vbCrLf &
                                               "    AND (n.Serie_Id not in('D','F','REC'))" & vbCrLf &
                                               "    AND (n.Empresa_Id= '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                               "    AND (ne.Encargo_Id IN ('PIS','PIS RECUP.','COFINS','COFINS RECUP.','COFINS A REC.','FRETES','SEGURO','PRODUTO','DESCONTOS','IPI','ICMS'))" & vbCrLf &
                                               "    AND (not so.Classe in ('SERVICOS', 'CONTAEORDEM', 'DEPOSITOS','TRANSFERENCIAS', 'DOACAO') OR (ni.Cfop_Id IN (5206,6206))) " & vbCrLf &
                                               "    AND (n.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5150 AND 5160)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6150 AND 6160)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
                                               "    AND (ni.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
                                               "    AND (ni.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
                                               "    AND (ni.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT IN (5555))" & vbCrLf &
                                               "    AND NOT (n.TipoDeDocumento = 57)" & vbCrLf &
                                               "    AND ne.Valor > 0" & vbCrLf &
                                               "    AND OE.STPISCOFINS <> 49"
                                    Else
                                        Sql &= "    AND (ISNULL(n.Situacao, 1) = 1)" & vbCrLf &
                                               "    AND (n.Serie_Id not in('D','F','REC'))" & vbCrLf &
                                               "    AND (n.Empresa_Id= '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                               "    AND (ne.Encargo_Id IN ('PIS','PIS RECUP.','COFINS','COFINS RECUP.','COFINS A REC.','FRETES','SEGURO','PRODUTO','DESCONTOS','IPI','ICMS'))" & vbCrLf &
                                               "    AND (not so.Classe in ('SERVICOS', 'CONTAEORDEM', 'DEPOSITOS','TRANSFERENCIAS', 'DOACAO') OR (ni.Cfop_Id IN (5206,6206))) " & vbCrLf &
                                               "    AND (n.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5150 AND 5160)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6150 AND 6160)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
                                               "    AND (ni.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
                                               "    AND (ni.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
                                               "    AND (ni.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT IN (5555))" & vbCrLf &
                                               "    AND ne.Valor > 0" & vbCrLf &
                                               "    AND OE.STPISCOFINS <> 49"
                                    End If

                                    Sql &= " GROUP BY n.NossaEmissao, n.Cliente_id, n.EndCliente_id, n.Serie_id, n.Nota_Id, nFe.ChaveNfe, n.Movimento, n.DataDaNota, n.CifFob, pg.AVista, DocXML.Tipo" & vbCrLf &
                                    " -- C100 - Entradas" & vbCrLf &
                                    "" & vbCrLf &
                                    " INSERT INTO #TempPisCofins" & vbCrLf &
                                    " SELECT 'C100' AS Reg," & vbCrLf &
                                    "		0 AS Ind_Oper," & vbCrLf &
                                    "		case " & vbCrLf &
                                    "		    when ISNULL(DocXML.Tipo,'') = 'Mic'" & vbCrLf &
                                    "		        then 0" & vbCrLf &
                                    "		        else case " & vbCrLf &
                                    "		                when n.NossaEmissao = 'S' " & vbCrLf &
                                    "		                then 0 " & vbCrLf &
                                    "		                else 1 " & vbCrLf &
                                    "		             end " & vbCrLf &
                                    "		    end AS Ind_Emit," & vbCrLf &
                                    "		n.Cliente_id," & vbCrLf &
                                    "		n.EndCliente_id," & vbCrLf &
                                    "		case" & vbCrLf &
                                    "		   when nFe.ChaveNfe is null" & vbCrLf &
                                    "		      then 1" & vbCrLf &
                                    "			  else 55" & vbCrLf &
                                    "		   end AS Cod_Mod," & vbCrLf &
                                    "		'00' AS Cod_Sit," & vbCrLf &
                                    "		n.Serie_id AS Ser," & vbCrLf &
                                    "		n.Nota_id AS Num_Doc," & vbCrLf &
                                    "		nFe.ChaveNfe AS Chv_Nfe," & vbCrLf &
                                    "		n.Movimento AS Dt_Doc," & vbCrLf &
                                    "		n.DataDaNota AS Dt_ES," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PRODUTO','FRETES','SEGURO','IPI') THEN ne.Valor END, 0) - ISNULL(CASE WHEN ne.Encargo_Id IN ('DESCONTOS') THEN ne.Valor END, 0)) AS Vl_Doc," & vbCrLf &
                                    "		case" & vbCrLf &
                                    "			when n.Movimento < '2020-01-01'" & vbCrLf &
                                    "				then 1" & vbCrLf &
                                    "				else" & vbCrLf &
                                    "				   case" & vbCrLf &
                                    "				      when isnull(pg.AVista,0) = 0" & vbCrLf &
                                    "					     then 1" & vbCrLf &
                                    "						 else 0" & vbCrLf &
                                    "				      end" & vbCrLf &
                                    "			end AS Ind_Pgto, " & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('DESCONTOS') THEN ne.Valor END, 0)) AS Vl_Desc," & vbCrLf &
                                    "		SUM(0) AS Vl_Abat_Nt, " & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PRODUTO') THEN ne.Valor END, 0)) AS Vl_Merc," & vbCrLf &
                                    "		case" & vbCrLf &
                                    "			when n.CifFob = 'CIF'" & vbCrLf &
                                    "				then '0'" & vbCrLf &
                                    "				else" & vbCrLf &
                                    "					case" & vbCrLf &
                                    "						when n.CifFob = 'FOB'" & vbCrLf &
                                    "						then '1'" & vbCrLf &
                                    "				        else" & vbCrLf &
                                    "					        case" & vbCrLf &
                                    "						    when n.CifFob = 'NEN'" & vbCrLf &
                                    "						    then '9'" & vbCrLf &
                                    "						    else '9'" & vbCrLf &
                                    "					        end" & vbCrLf &
                                    "					end" & vbCrLf &
                                    "				end Ind_Frt," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('FRETES') THEN ne.Valor END, 0)) AS Vl_Frt," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('SEGURO') THEN ne.Valor END, 0)) AS Vl_Seg," & vbCrLf &
                                    "		SUM(0) AS Vl_Out_Da, " & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN ne.Base END, 0)) AS Vl_Bc_Icms," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN ne.Valor END, 0)) AS Vl_Icms," & vbCrLf &
                                    "		SUM(0) AS Vl_Bc_Icms_St," & vbCrLf &
                                    "		SUM(0) AS Vl_Icms_St," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('IPI') THEN ne.Valor END, 0)) AS Vl_Ipi," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN ne.Valor END, 0)) AS Vl_Pis," & vbCrLf &
                                    "       SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN ne.Valor END, 0)) AS Vl_Cofins," & vbCrLf &
                                    "		SUM(0) AS Vl_Pis_St," & vbCrLf &
                                    "		SUM(0) AS Vl_Cofins_St" & vbCrLf &
                                    "   FROM NotasFiscais n" & vbCrLf &
                                    "  INNER JOIN NotasFiscaisXItens ni" & vbCrLf &
                                    "     ON ni.Empresa_Id      = n.Empresa_Id" & vbCrLf &
                                    "    AND ni.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                                    "    AND ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                                    "    AND ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                                    "    AND ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                                    "    AND ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                                    "    AND ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                                    "  INNER JOIN NotasFiscaisXEncargos ne" & vbCrLf &
                                    "     ON ne.Empresa_Id      = ni.Empresa_Id" & vbCrLf &
                                    "    AND ne.EndEmpresa_Id   = ni.EndEmpresa_Id" & vbCrLf &
                                    "    AND ne.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                                    "    AND ne.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                                    "    AND ne.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                                    "    AND ne.Serie_Id        = ni.Serie_Id" & vbCrLf &
                                    "    AND ne.Nota_Id         = ni.Nota_Id" & vbCrLf &
                                    "    AND ne.Produto_Id      = ni.Produto_Id" & vbCrLf &
                                    "    AND ne.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                                    "    AND ne.Sequencia_id    = ni.Sequencia_id" & vbCrLf &
                                    "  INNER JOIN Produtos" & vbCrLf &
                                    "     ON Produtos.Produto_Id = ni.Produto_Id" & vbCrLf &
                                    "  INNER JOIN SubOperacoes so" & vbCrLf &
                                    "     ON so.Operacao_Id     = ni.Operacao" & vbCrLf &
                                    "    AND so.SubOperacoes_Id = ni.SubOperacao" & vbCrLf &
                                    "  inner join OperacaoXEstado OE" & vbCrLf &
                                    "     on OE.Codigo_id = ni.OperacaoxEstado" & vbCrLf &
                                    "  left JOIN Pedidos p" & vbCrLf &
                                    "     ON p.Empresa_id     = n.Empresa_id" & vbCrLf &
                                    "    AND p.EndEmpresa_Id  = n.EndEmpresa_Id" & vbCrLf &
                                    "    AND p.Pedido_Id      = n.Pedido" & vbCrLf &
                                    "  left JOIN Pagamentos pg" & vbCrLf &
                                    "     ON pg.Pagamento_Id     = p.CondicaoPagamento" & vbCrLf &
                                    "  left JOIN NfeRealizadas nFe" & vbCrLf &
                                    "     ON nFe.Empresa_Id      = n.Empresa_Id" & vbCrLf &
                                    "    AND nFe.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                                    "    AND nFe.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                                    "    AND nFe.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                                    "    AND nFe.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                                    "    AND nFe.Serie_Id        = n.Serie_Id" & vbCrLf &
                                    "    AND nFe.Nota_Id         = n.Nota_Id" & vbCrLf &
                                    "  left JOIN DocumentoXML DocXML" & vbCrLf &
                                    "     ON DocXML.Empresa_Id = n.Empresa_Id" & vbCrLf &
                                    "    AND DocXML.Cliente_Id = n.Cliente_Id" & vbCrLf &
                                    "    AND DocXML.Serie_Id   = n.Serie_Id" & vbCrLf &
                                    "    AND DocXML.Numero_Id  = n.Nota_Id" & vbCrLf &
                                    "  WHERE (n.EntradaSaida_Id = 'E')"

                                    If EmpresaBlocoC010.Contains("44979506") Then
                                        Sql &= "    AND (ISNULL(n.Situacao, 1) = 1)" & vbCrLf &
                                               "    AND (n.Serie_Id not in('D','F','REC'))" & vbCrLf &
                                               "    AND (n.Empresa_Id= '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                               "    AND (ne.Encargo_Id IN ('PIS','PIS RECUP.','COFINS','COFINS RECUP.','COFINS A REC.','FRETES','SEGURO','PRODUTO','DESCONTOS','IPI','ICMS'))" & vbCrLf &
                                               "    AND (not so.Classe in ('SERVICOS') OR (ni.Cfop_Id IN (5206))) " & vbCrLf &
                                               "    AND (n.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5550 AND 5555)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6550 AND 6555)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 3000 AND 4000)" & vbCrLf &
                                               "    AND NOT (n.TipoDeDocumento = 57)" & vbCrLf &
                                               "	AND Not OE.STPISCOFINS in (98, 99)" & vbCrLf &
                                               " GROUP BY n.NossaEmissao, n.Cliente_id, n.EndCliente_id, n.Serie_id, n.Nota_Id, nFe.ChaveNfe, n.Movimento, n.DataDaNota, n.CifFob, pg.AVista, DocXML.Tipo" & vbCrLf &
                                               " HAVING (SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN ne.Valor END, 0)) > 0)"
                                    Else
                                        Sql &= "    AND (ISNULL(n.Situacao, 1) = 1)" & vbCrLf &
                                               "    AND (n.Serie_Id not in('D','F','REC'))" & vbCrLf &
                                               "    AND (n.Empresa_Id= '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                               "    AND (ne.Encargo_Id IN ('PIS','PIS RECUP.','COFINS','COFINS RECUP.','COFINS A REC.','FRETES','SEGURO','PRODUTO','DESCONTOS','IPI','ICMS'))" & vbCrLf &
                                               "    AND (not so.Classe in ('SERVICOS') OR (ni.Cfop_Id IN (5206))) " & vbCrLf &
                                               "    AND (n.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 5550 AND 5555)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 6550 AND 6555)" & vbCrLf &
                                               "    AND (ni.Cfop_Id NOT BETWEEN 3000 AND 4000)" & vbCrLf &
                                               "	AND Not OE.STPISCOFINS in (98, 99)" & vbCrLf &
                                               " GROUP BY n.NossaEmissao, n.Cliente_id, n.EndCliente_id, n.Serie_id, n.Nota_Id, nFe.ChaveNfe, n.Movimento, n.DataDaNota, n.CifFob, pg.AVista, DocXML.Tipo" & vbCrLf &
                                               " HAVING (SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN ne.Valor END, 0)) > 0)"
                                    End If

                                    Sql &= " SELECT Reg, Ind_Oper, Ind_Emit, Cliente_id, EndCliente_id, Cod_Mod, Cod_Sit, Ser," & vbCrLf &
                                    "        Num_Doc, Chv_Nfe, Dt_Doc, Dt_ES, Vl_Doc, Ind_Pgto, Vl_Desc, Vl_Abat_Nt, " & vbCrLf &
                                    "        Vl_Merc, Ind_Frt, Vl_Frt, Vl_Seg, Vl_Out_Da, Vl_Bc_Icms, Vl_Icms, Vl_Bc_Icms_St," & vbCrLf &
                                    "		Vl_Icms_St, Vl_Ipi, Vl_Pis, Vl_Cofins, Vl_Pis_St, Vl_Cofins_St" & vbCrLf &
                                    " FROM #TempPisCofins" & vbCrLf &
                                    " ORDER BY Ind_Oper, Num_Doc"

                                    dsAux = Banco.ConsultaDataSet(Sql, "BlocoC100")

                                    For Each drC100 In dsAux.Tables(0).Rows
                                        With drC100

                                            'ESSE TESTE ESTÁ ERRADO, DEVEM IR TODAS AS NOTAS ENCONTRADAS
                                            'DESSE JEITO ESTAVA FICANDO C170 ORFÃO
                                            'CASO ALGUM CLIENTE APRESENTE PROBLEMA ME CHAMAR - FURLAN - 29/04/2025
                                            'If .Item("Vl_Pis") > 0 Then
                                            'End If

                                            linha = "|" & .Item("Reg")
                                            linha &= "|" & .Item("Ind_Oper")
                                            linha &= "|" & .Item("Ind_Emit")
                                            linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                                            linha &= "|" & Format(.Item("Cod_Mod"), "00")
                                            linha &= "|" & .Item("Cod_Sit")
                                            linha &= "|" & .Item("Ser")
                                            linha &= "|" & .Item("Num_Doc")
                                            linha &= "|" & .Item("Chv_Nfe")
                                            linha &= "|" & CDate(.Item("Dt_ES")).ToStrDate().Replace("-", "")
                                            linha &= "|" & CDate(.Item("Dt_Doc")).ToStrDate().Replace("-", "")
                                            linha &= "|" & .Item("Vl_Doc")
                                            linha &= "|" & .Item("Ind_Pgto")
                                            linha &= "|" & .Item("Vl_Desc")
                                            linha &= "|" & .Item("Vl_Abat_Nt")
                                            linha &= "|" & .Item("Vl_Merc")

                                            linha &= "|" & .Item("Ind_Frt")

                                            linha &= "|" & .Item("Vl_Frt")
                                            linha &= "|" & .Item("Vl_Seg")
                                            linha &= "|" & .Item("Vl_Out_Da")
                                            linha &= "|" & .Item("Vl_Bc_Icms")
                                            linha &= "|" & .Item("Vl_Icms")
                                            linha &= "|" & .Item("Vl_Bc_Icms_St")
                                            linha &= "|" & .Item("Vl_Icms_St")
                                            linha &= "|" & .Item("Vl_Ipi")
                                            linha &= "|" & .Item("Vl_Pis")
                                            linha &= "|" & .Item("Vl_Cofins")
                                            linha &= "|" & .Item("Vl_Pis_St")
                                            linha &= "|" & .Item("Vl_Cofins_St")
                                            linha &= "|"

                                            ArquivoAux.Add(linha)
                                            RegistroC100 += 1
                                            RegistroGeral += 1

                                            'C170 - Itens do Documento Bloco C

                                            If .Item("Ind_Oper") = "0" Then

                                                If .Item("Num_Doc") = 223 Then
                                                    Dim TESTE As String = "223"
                                                End If

                                                'ENTRADAS
                                                Sql = "SELECT 'C170' AS Reg," & vbCrLf &
                                                        "		ni.Sequencia_Id AS Num_Item," & vbCrLf &
                                                        "		ni.Produto_Id AS Cod_Item," & vbCrLf &
                                                        "		prd.Nome AS Descr_Compl," & vbCrLf &
                                                        "		SUM(ISNULL(ni.QuantidadeFiscal,0)) AS Qtd," & vbCrLf &
                                                        "		case" & vbCrLf &
                                                        "		   when isnull(pXi.Pedido_id,0) = 0" & vbCrLf &
                                                        "		      then prd.Unidade" & vbCrLf &
                                                        "			  else pXi.UnidadeComercializacao" & vbCrLf &
                                                        "		   end AS Unid," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PRODUTO') THEN ne.Valor END, 0)) AS Vl_Item," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('DESCONTOS') THEN ne.Valor END, 0)) AS Vl_Desc," & vbCrLf &
                                                        "		'0' AS Ind_Mov," & vbCrLf &
                                                        "		OE.StIcms AS Cst_Icms," & vbCrLf &
                                                        "		ni.Cfop_Id AS Cfop," & vbCrLf &
                                                        "		ni.Operacao," & vbCrLf &
                                                        "		ni.SubOperacao," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN CASE WHEN isnull(ne.BaseNova,0) > 0 then ne.BaseNova else ne.Base END END, 0)) AS Vl_Bc_Icms," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN ne.Percentual END, 0)) AS Aliq_Icms," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN CASE WHEN isnull(ne.ValorNovo,0) > 0 then ne.ValorNovo else ne.Valor END END, 0)) AS Vl_Icms," & vbCrLf &
                                                        "		SUM(0) AS Vl_Bc_Icms_St," & vbCrLf &
                                                        "		0 AS Aliq_St," & vbCrLf &
                                                        "		SUM(0) AS Vl_Icms_St," & vbCrLf &
                                                        "		0 AS Ind_Apur," & vbCrLf &
                                                        "        OE.StIpi AS Cst_Ipi," & vbCrLf &
                                                        "		'' AS Cod_Enq," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('IPI') THEN CASE WHEN isnull(ne.BaseNova,0) > 0 then ne.BaseNova else ne.Base END END, 0)) AS Vl_Bc_Ipi," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('IPI') THEN ne.Percentual END, 0)) AS Aliq_Ipi," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('IPI') THEN CASE WHEN isnull(ne.ValorNovo,0) > 0 then ne.ValorNovo else ne.Valor END END, 0)) AS Vl_Ipi," & vbCrLf &
                                                        "        OE.StPisCofins AS Cst_Pis," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(ne.BaseNova,0) > 0 then ne.BaseNova else ne.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN ne.Percentual END, 0)) AS Aliq_Pis," & vbCrLf &
                                                        "		SUM(0) AS Quant_Bc_Pis," & vbCrLf &
                                                        "		SUM(0) AS Aliq_Pis_Quant," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(ne.ValorNovo,0) > 0 then ne.ValorNovo else ne.Valor END END, 0)) AS Vl_Pis," & vbCrLf &
                                                        "        OE.StPisCofins AS Cst_Cofins," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(ne.BaseNova,0) > 0 then ne.BaseNova else ne.Base END END, 0)) AS Vl_Bc_Cofins," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN ne.Percentual END, 0)) AS Aliq_Cofins," & vbCrLf &
                                                        "		SUM(0) AS Quant_Bc_Cofins," & vbCrLf &
                                                        "		SUM(0) AS Aliq_Cofins_Quant," & vbCrLf &
                                                        "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(ne.ValorNovo,0) > 0 then ne.ValorNovo else ne.Valor END END, 0)) AS Vl_Cofins," & vbCrLf &
                                                        "		case" & vbCrLf &
                                                        "			when n.EntradaSaida_Id = 'E'" & vbCrLf &
                                                        "				then OEE.DebitaConta" & vbCrLf &
                                                        "				else OEE.CreditaConta" & vbCrLf &
                                                        "		end AS Cod_Cta" & vbCrLf &
                                                        "   FROM NotasFiscais n" & vbCrLf &
                                                        "  INNER JOIN NotasFiscaisXItens ni" & vbCrLf &
                                                        "     ON ni.Empresa_Id      = n.Empresa_Id" & vbCrLf &
                                                        "    AND ni.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                                                        "    AND ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                                                        "    AND ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                                                        "    AND ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                                                        "    AND ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                                                        "    AND ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                                                        "  INNER JOIN NotasFiscaisXEncargos ne" & vbCrLf &
                                                        "     ON ne.Empresa_Id      = ni.Empresa_Id" & vbCrLf &
                                                        "    AND ne.EndEmpresa_Id   = ni.EndEmpresa_Id" & vbCrLf &
                                                        "    AND ne.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                                                        "    AND ne.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                                                        "    AND ne.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                                                        "    AND ne.Serie_Id        = ni.Serie_Id" & vbCrLf &
                                                        "    AND ne.Nota_Id         = ni.Nota_Id" & vbCrLf &
                                                        "    AND ne.Produto_Id      = ni.Produto_Id" & vbCrLf &
                                                        "    AND ne.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                                                        "    AND ne.Sequencia_id    = ni.Sequencia_id" & vbCrLf &
                                                        "  INNER JOIN Produtos prd" & vbCrLf &
                                                        "     ON prd.Produto_Id = ni.Produto_Id" & vbCrLf &
                                                        "  INNER JOIN SubOperacoes so" & vbCrLf &
                                                        "     ON so.Operacao_Id     = ni.Operacao" & vbCrLf &
                                                        "    AND so.SubOperacoes_Id = ni.SubOperacao" & vbCrLf &
                                                        "  inner join OperacaoXEstado OE" & vbCrLf &
                                                        "     on OE.Codigo_id = ni.OperacaoxEstado" & vbCrLf &
                                                        "  inner join OperacaoXEstadoXEncargo OEE" & vbCrLf &
                                                        "     on OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
                                                        "    and OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
                                                        "  left JOIN Pedidos p" & vbCrLf &
                                                        "     ON p.Empresa_id     = n.Empresa_id" & vbCrLf &
                                                        "    AND p.EndEmpresa_Id  = n.EndEmpresa_Id" & vbCrLf &
                                                        "	 AND p.Pedido_Id      = n.Pedido" & vbCrLf &
                                                        "  left JOIN PedidoXItem pXi" & vbCrLf &
                                                        "     ON pXi.Empresa_id     = n.Empresa_id" & vbCrLf &
                                                        "    AND pXi.EndEmpresa_Id  = n.EndEmpresa_Id" & vbCrLf &
                                                        "	 AND pXi.Pedido_Id      = n.Pedido" & vbCrLf &
                                                        "	 AND pXi.Produto_Id     = ni.Produto_Id" & vbCrLf &
                                                        "  WHERE (n.EntradaSaida_Id = 'E')" & vbCrLf &
                                                        "    AND (ISNULL(n.Situacao, 1) = 1)" & vbCrLf &
                                                        "    AND (n.Empresa_Id= '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                                        "  	 AND (n.Cliente_id = '" & .Item("Cliente_Id") & "')" & vbCrLf &
                                                        "	 AND (n.EndCliente_Id = " & .Item("EndCliente_Id") & ")" & vbCrLf &
                                                        "	 AND (n.Serie_id = '" & .Item("Ser") & "')" & vbCrLf &
                                                        "	 AND (n.Nota_Id = " & .Item("Num_Doc") & ")" & vbCrLf &
                                                        "    AND (ne.Encargo_Id IN ('PIS','PIS RECUP.','COFINS','COFINS RECUP.','COFINS A REC.','FRETES','SEGURO','PRODUTO','DESCONTOS','IPI','ICMS'))" & vbCrLf &
                                                        "    AND (not so.Classe in ('SERVICOS') OR (ni.Cfop_Id IN (5206))) " & vbCrLf &
                                                        "    AND (n.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 5550 AND 5555)" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 6550 AND 6555)" & vbCrLf &
                                                        "    AND (ni.Cfop_Id NOT BETWEEN 3000 AND 4000)" & vbCrLf &
                                                        "    AND Not OE.STPISCOFINS in (98, 99)" & vbCrLf &
                                                        "Group by ni.Cliente_Id, n.EntradaSaida_Id, ni.nota_Id, ni.Sequencia_Id, ni.Produto_Id, prd.Nome, pXi.Pedido_id, prd.Unidade, pXi.UnidadeComercializacao, OE.StIcms, ni.Cfop_Id," & vbCrLf &
                                                        "         ni.Operacao, ni.SubOperacao, OE.StIpi, OE.StPisCofins, OEE.DebitaConta, OEE.CreditaConta "
                                                '" HAVING (SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN ne.Valor END, 0)) > 0)"
                                            Else
                                                'SAÍDAS

                                                '"    AND (ni.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf & _ Removi 5556 das Saídas cfe. Solicitação Luiz(Nutri) - Furlan - 09/08/2022

                                                Sql = "SELECT 'C170' AS Reg," & vbCrLf &
                                                "		ni.Sequencia_Id AS Num_Item," & vbCrLf &
                                                "		ni.Produto_Id AS Cod_Item," & vbCrLf &
                                                "		prd.Nome AS Descr_Compl," & vbCrLf &
                                                "		SUM(ISNULL(ni.QuantidadeFiscal,0)) AS Qtd," & vbCrLf &
                                                "		case" & vbCrLf &
                                                "		   when isnull(pXi.Pedido_id,0) = 0" & vbCrLf &
                                                "		      then prd.Unidade" & vbCrLf &
                                                "			  else pXi.UnidadeComercializacao" & vbCrLf &
                                                "		   end AS Unid," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PRODUTO') THEN ne.Valor END, 0)) AS Vl_Item," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('DESCONTOS') THEN ne.Valor END, 0)) AS Vl_Desc," & vbCrLf &
                                                "		'0' AS Ind_Mov," & vbCrLf &
                                                "		OE.StIcms AS Cst_Icms," & vbCrLf &
                                                "		ni.Cfop_Id AS Cfop," & vbCrLf &
                                                "		ni.Operacao," & vbCrLf &
                                                "		ni.SubOperacao," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN CASE WHEN isnull(ne.BaseNova,0) > 0 then ne.BaseNova else ne.Base END END, 0)) AS Vl_Bc_Icms," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN ne.Percentual END, 0)) AS Aliq_Icms," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('ICMS') THEN CASE WHEN isnull(ne.ValorNovo,0) > 0 then ne.ValorNovo else ne.Valor END END, 0)) AS Vl_Icms," & vbCrLf &
                                                "		SUM(0) AS Vl_Bc_Icms_St," & vbCrLf &
                                                "		0 AS Aliq_St," & vbCrLf &
                                                "		SUM(0) AS Vl_Icms_St," & vbCrLf &
                                                "		0 AS Ind_Apur," & vbCrLf &
                                                "        OE.StIpi AS Cst_Ipi," & vbCrLf &
                                                "		'' AS Cod_Enq," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('IPI') THEN CASE WHEN isnull(ne.BaseNova,0) > 0 then ne.BaseNova else ne.Base END END, 0)) AS Vl_Bc_Ipi," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('IPI') THEN ne.Percentual END, 0)) AS Aliq_Ipi," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('IPI') THEN CASE WHEN isnull(ne.ValorNovo,0) > 0 then ne.ValorNovo else ne.Valor END END, 0)) AS Vl_Ipi," & vbCrLf &
                                                "        OE.StPisCofins AS Cst_Pis," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(ne.BaseNova,0) > 0 then ne.BaseNova else ne.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN ne.Percentual END, 0)) AS Aliq_Pis," & vbCrLf &
                                                "		SUM(0) AS Quant_Bc_Pis," & vbCrLf &
                                                "		SUM(0) AS Aliq_Pis_Quant," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(ne.ValorNovo,0) > 0 then ne.ValorNovo else ne.Valor END END, 0)) AS Vl_Pis," & vbCrLf &
                                                "        OE.StPisCofins AS Cst_Cofins," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(ne.BaseNova,0) > 0 then ne.BaseNova else ne.Base END END, 0)) AS Vl_Bc_Cofins," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN ne.Percentual END, 0)) AS Aliq_Cofins," & vbCrLf &
                                                "		SUM(0) AS Quant_Bc_Cofins," & vbCrLf &
                                                "		SUM(0) AS Aliq_Cofins_Quant," & vbCrLf &
                                                "        SUM(ISNULL(CASE WHEN ne.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(ne.ValorNovo,0) > 0 then ne.ValorNovo else ne.Valor END END, 0)) AS Vl_Cofins," & vbCrLf &
                                                "		case" & vbCrLf &
                                                "			when n.EntradaSaida_Id = 'E'" & vbCrLf &
                                                "				then OEE.DebitaConta" & vbCrLf &
                                                "				else OEE.CreditaConta" & vbCrLf &
                                                "		end AS Cod_Cta" & vbCrLf &
                                                "   FROM NotasFiscais n" & vbCrLf &
                                                "  INNER JOIN NotasFiscaisXItens ni" & vbCrLf &
                                                "     ON ni.Empresa_Id      = n.Empresa_Id" & vbCrLf &
                                                "    AND ni.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                                                "    AND ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                                                "    AND ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                                                "    AND ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                                                "    AND ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                                                "    AND ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                                                "  INNER JOIN NotasFiscaisXEncargos ne" & vbCrLf &
                                                "     ON ne.Empresa_Id      = ni.Empresa_Id" & vbCrLf &
                                                "    AND ne.EndEmpresa_Id   = ni.EndEmpresa_Id" & vbCrLf &
                                                "    AND ne.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                                                "    AND ne.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                                                "    AND ne.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                                                "    AND ne.Serie_Id        = ni.Serie_Id" & vbCrLf &
                                                "    AND ne.Nota_Id         = ni.Nota_Id" & vbCrLf &
                                                "    AND ne.Produto_Id      = ni.Produto_Id" & vbCrLf &
                                                "    AND ne.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                                                "    AND ne.Sequencia_id    = ni.Sequencia_id" & vbCrLf &
                                                "  INNER JOIN Produtos prd" & vbCrLf &
                                                "     ON prd.Produto_Id = ni.Produto_Id" & vbCrLf &
                                                "  INNER JOIN SubOperacoes so" & vbCrLf &
                                                "     ON so.Operacao_Id     = ni.Operacao" & vbCrLf &
                                                "    AND so.SubOperacoes_Id = ni.SubOperacao" & vbCrLf &
                                                "  inner join OperacaoXEstado OE" & vbCrLf &
                                                "     on OE.Codigo_id = ni.OperacaoxEstado" & vbCrLf &
                                                "  inner join OperacaoXEstadoXEncargo OEE" & vbCrLf &
                                                "     on OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
                                                "    and OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
                                                "  left JOIN Pedidos p" & vbCrLf &
                                                "     ON p.Empresa_id     = n.Empresa_id" & vbCrLf &
                                                "    AND p.EndEmpresa_Id  = n.EndEmpresa_Id" & vbCrLf &
                                                "	AND p.Pedido_Id      = n.Pedido" & vbCrLf &
                                                "  left JOIN PedidoXItem pXi" & vbCrLf &
                                                "     ON pXi.Empresa_id     = n.Empresa_id" & vbCrLf &
                                                "    AND pXi.EndEmpresa_Id  = n.EndEmpresa_Id" & vbCrLf &
                                                "	AND pXi.Pedido_Id      = n.Pedido" & vbCrLf &
                                                "	AND pXi.Produto_Id     = ni.Produto_Id" & vbCrLf &
                                                "  WHERE (n.EntradaSaida_Id = 'S')" & vbCrLf &
                                                "    AND (ISNULL(n.Situacao, 1) = 1)" & vbCrLf &
                                                "    AND (n.Empresa_Id= '" & EmpresaBlocoC010 & "')" & vbCrLf &
                                                "	AND (n.Cliente_id = '" & .Item("Cliente_Id") & "')" & vbCrLf &
                                                "	AND (n.EndCliente_Id = " & .Item("EndCliente_Id") & ")" & vbCrLf &
                                                "	AND (n.Serie_id = '" & .Item("Ser") & "')" & vbCrLf &
                                                "	AND (n.Nota_Id = " & .Item("Num_Doc") & ")" & vbCrLf &
                                                "    AND (ne.Encargo_Id IN ('PIS','PIS RECUP.','COFINS','COFINS RECUP.','COFINS A REC.','FRETES','SEGURO','PRODUTO','DESCONTOS','IPI','ICMS'))" & vbCrLf &
                                                "    AND (not so.Classe in ('SERVICOS', 'CONTAEORDEM', 'DEPOSITOS','TRANSFERENCIAS', 'DOACAO') OR (ni.Cfop_Id IN (5206,6206))) " & vbCrLf &
                                                "    AND (n.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT IN (1932,2932))" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 5150 AND 5160)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 6150 AND 6160)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
                                                "    AND (ni.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
                                                "    AND (ni.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
                                                "    AND (ni.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
                                                "    AND (ni.Cfop_Id NOT IN (5555))" & vbCrLf &
                                                "    AND ne.Valor > 0" & vbCrLf &
                                                "    AND OE.STPISCOFINS <> 49" & vbCrLf &
                                                "Group by n.EntradaSaida_Id, ni.Sequencia_Id, ni.Produto_Id, prd.Nome, pXi.Pedido_id, prd.Unidade, pXi.UnidadeComercializacao, OE.StIcms, ni.Cfop_Id," & vbCrLf &
                                                "         ni.Operacao, ni.SubOperacao, OE.StIpi, OE.StPisCofins, OEE.DebitaConta, OEE.CreditaConta"
                                            End If

                                            dsAux = Banco.ConsultaDataSet(Sql, "BlocoC170")

                                            For Each drC170 In dsAux.Tables(0).Rows
                                                With drC170
                                                    linha = "|" & .Item("Reg")
                                                    linha &= "|" & .Item("Num_Item")
                                                    linha &= "|" & .Item("Cod_Item")
                                                    linha &= "|" & .Item("Descr_Compl")
                                                    linha &= "|" & CDbl(.Item("Qtd")).ToString("N5").Replace(".", "")
                                                    linha &= "|" & .Item("Unid")
                                                    linha &= "|" & .Item("Vl_Item")
                                                    linha &= "|" & .Item("Vl_Desc")
                                                    linha &= "|" & .Item("Ind_Mov")
                                                    linha &= "|" & Format(.Item("Cst_Icms"), "000")
                                                    linha &= "|" & .Item("Cfop")
                                                    linha &= "|" & Format(.Item("Operacao"), "00") & Format(.Item("SubOperacao"), "00")
                                                    linha &= "|" & .Item("Vl_Bc_Icms")
                                                    linha &= "|" & CDbl(.Item("Aliq_Icms")).ToString("N2")
                                                    linha &= "|" & .Item("Vl_Icms")
                                                    linha &= "|" & .Item("Vl_Bc_Icms_St")
                                                    linha &= "|" & CDbl(.Item("Aliq_St")).ToString("N2")
                                                    linha &= "|" & .Item("Vl_Icms_St")
                                                    'linha &= "|" & .Item("Ind_Apur")
                                                    'linha &= "|" & Format(.Item("Cst_Ipi"), "00")
                                                    'linha &= "|" & .Item("Cod_Enq")
                                                    linha &= "|" & ""
                                                    linha &= "|" & ""
                                                    linha &= "|" & ""
                                                    linha &= "|" & .Item("Vl_Bc_Ipi")
                                                    linha &= "|" & CDbl(.Item("Aliq_Ipi")).ToString("N2")
                                                    linha &= "|" & .Item("Vl_Ipi")
                                                    linha &= "|" & Format(.Item("Cst_Pis"), "00")
                                                    linha &= "|" & .Item("Vl_Bc_Pis")
                                                    linha &= "|" & CDbl(.Item("Aliq_Pis")).ToString("N4")
                                                    'linha &= "|" & CDbl(.Item("Quant_Bc_Pis")).ToString("N3")
                                                    'linha &= "|" & CDbl(.Item("Aliq_Pis_Quant")).ToString("N4")
                                                    linha &= "|" & ""
                                                    linha &= "|" & ""
                                                    linha &= "|" & .Item("Vl_Pis")
                                                    linha &= "|" & Format(.Item("Cst_Cofins"), "00")
                                                    linha &= "|" & .Item("Vl_Bc_Cofins")
                                                    linha &= "|" & CDbl(.Item("Aliq_Cofins")).ToString("N4")
                                                    'linha &= "|" & CDbl(.Item("Quant_Bc_Cofins")).ToString("N3")
                                                    'linha &= "|" & CDbl(.Item("Aliq_Cofins_Quant")).ToString("N4")
                                                    linha &= "|" & ""
                                                    linha &= "|" & ""
                                                    linha &= "|" & .Item("Vl_Cofins")
                                                    linha &= "|" & .Item("Cod_Cta")
                                                    linha &= "|"
                                                End With

                                                ArquivoAux.Add(linha)
                                                RegistroC170 += 1
                                                RegistroGeral += 1
                                            Next

                                        End With
                                    Next
                                End If

                                ''C500 Nota Fiscal De Energia Elétrica
                                '' WHEN NotasFiscaisXItens.cfop_Id IN (1555, 1556, 2555, 2556)   THEN '29'"  - Linha 2368
                                ''  "    AND NotasFiscaisXItens.cfop_Id in (1252, 1253, 2252, 2253, 1555, 1556, 2555, 2556)" & vbCrLf & _ - Linha 2403
                                ''Removi junto com Alessandro do Curtume .. se necessário vamos ter que criar um tipo de documento para tratar com esse CFOP

                                Sql = " SELECT 'C500' AS Reg," & vbCrLf &
                                      "         Clientes.Cliente_Id AS Cod_Part," & vbCrLf &
                                      "         Clientes.Endereco_Id," & vbCrLf &
                                      "         (CASE WHEN NotasFiscaisXItens.cfop_Id IN (1252, 1253, 2252, 2253)   THEN '06'" & vbCrLf &
                                      "               WHEN NotasFiscaisXItens.cfop_Id IN (1555, 1556, 2555, 2556)   THEN '29'" & vbCrLf &
                                      "               ELSE '28'" & vbCrLf &
                                      "         END) AS Cod_Mod," & vbCrLf &
                                      "        '0' AS Cod_Sit, NotasFiscais.Serie_Id AS Ser," & vbCrLf &
                                      "        NotasFiscais.Nota_Id AS Num_Doc, NotasFiscais.Movimento AS DT_Doc, NotasFiscais.DataDaNota AS DT_Ent," & vbCrLf &
                                      "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Doc," & vbCrLf &
                                      "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS'    THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Icms," & vbCrLf &
                                      "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Pis," & vbCrLf &
                                      "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.')  THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Cofins" & vbCrLf &
                                      "   FROM NotasFiscais" & vbCrLf &
                                      "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                      "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                      "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                      "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                      "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                      "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf &
                                      "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                      "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                                      "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                      "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                      "    AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
                                      "  INNER JOIN Clientes" & vbCrLf &
                                      "     ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
                                      "    AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                                      "  WHERE (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                      "    AND (NotasFiscais.Empresa_Id = '" & .Item("Cliente_Id") & "') " & vbCrLf &
                                      "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                      "    AND NotasFiscaisXItens.cfop_Id in (1252, 1253, 2252, 2253)" & vbCrLf &
                                      "    AND (NotasFiscais.EntradaSaida_Id = 'E')" & vbCrLf &
                                      "  GROUP BY Clientes.Cliente_Id ,Clientes.Endereco_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id," & vbCrLf &
                                      "           NotasFiscais.Movimento, NotasFiscais.DataDaNota, NotasFiscaisXItens.cfop_Id" & vbCrLf &
                                      " HAVING (SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor END, 0)) > 0)" & vbCrLf

                                dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                For Each drC500 In dsAux.Tables(0).Rows
                                    With drC500

                                        Nota = .Item("NUM_DOC")
                                        Serie = .Item("SER")
                                        Cliente = Trim(.Item("COD_PART"))
                                        EndCliente = .Item("Endereco_Id")

                                        linha = "|" & .Item("Reg")
                                        linha &= "|" & .Item("COD_PART") & .Item("Endereco_Id")
                                        linha &= "|" & .Item("COD_MOD")
                                        linha &= "|" & Format(CInt(.Item("COD_SIT")), "00")
                                        linha &= "|" & .Item("SER")
                                        linha &= "|" & ""
                                        linha &= "|" & .Item("NUM_DOC")
                                        linha &= "|" & CDate(.Item("DT_DOC")).ToStrDate().Replace("-", "")
                                        linha &= "|" & CDate(.Item("DT_ENT")).ToStrDate().Replace("-", "")
                                        linha &= "|" & .Item("VL_DOC")
                                        linha &= "|" & .Item("Vl_ICMS")
                                        linha &= "|" & ""
                                        linha &= "|" & .Item("VL_PIS")
                                        linha &= "|" & .Item("VL_COFINS")

                                        If CDate(txtDataFinal.Text) > "2019/12/31" Then
                                            linha &= "|" & ""
                                        End If

                                        linha &= "|"
                                    End With

                                    ArquivoAux.Add(linha)
                                    RegistroC500 += 1
                                    RegistroGeral += 1


                                    ' C501 Complemento Energia eletrica PIS ---------------------------

                                    Sql = " SELECT 'C501' AS Reg," & vbCrLf &
                                          "        ISNULL(OE.STPisCofins, 0) AS Cst_Pis," & vbCrLf &
                                          "        SUM(ISNULL(CASE WHEN NotasFiscaisXItens.Valor > 0 THEN NotasFiscaisXItens.Valor END, 0)) AS Vl_Item, '04' AS Nat_Bc_Cred," & vbCrLf &
                                          "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
                                          "        NotasFiscaisXEncargos.Percentual AS ALiq_Pis," & vbCrLf &
                                          "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END END, 0)) AS Vl_Pis," & vbCrLf &
                                          "		   case" & vbCrLf &
                                          "		   	when NotasFiscais.EntradaSaida_Id = 'E'" & vbCrLf &
                                          "				then OEE.DebitaConta" & vbCrLf &
                                          "				else OEE.CreditaConta" & vbCrLf &
                                          "		   end AS Cod_Cta" & vbCrLf &
                                          "   FROM NotasFiscais " & vbCrLf &
                                          "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                          "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                          "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                                          "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                          "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
                                          "  inner join OperacaoXEstado OE" & vbCrLf &
                                          "     on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                                          "  inner join OperacaoXEstadoXEncargo OEE" & vbCrLf &
                                          "     on OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
                                          "    and OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
                                          "  WHERE (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                          "    AND (NotasFiscais.Empresa_Id = '" & .Item("Cliente_Id") & "') " & vbCrLf &
                                          "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                          "    AND NotasFiscais.Cliente_Id     = '" & Cliente & "' " & vbCrLf &
                                          "    AND NotasFiscais.EndCliente_Id  = " & EndCliente & "  " & vbCrLf &
                                          "    And NotasFiscais.Nota_Id        = " & Nota & " " & vbCrLf &
                                          "    AND NotasFiscais.Serie_Id       = '" & Serie & "'" & vbCrLf &
                                          "    AND NotasFiscaisXItens.cfop_Id in (1252, 1253, 2252, 2253, 1555, 1556, 2555, 2556)" & vbCrLf &
                                          "    AND (NotasFiscais.EntradaSaida_Id = 'E')" & vbCrLf &
                                          "  GROUP BY OE.STPisCofins, NotasFiscaisXEncargos.Percentual, NotasFiscais.EntradaSaida_Id, OEE.DebitaConta, OEE.CreditaConta" & vbCrLf &
                                          " HAVING (SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor END, 0)) > 0) " & vbCrLf


                                    dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                    For Each drC501 In dsAux.Tables(0).Rows
                                        With drC501
                                            linha = "|" & .Item("Reg")
                                            linha &= "|" & Format(CInt(.Item("CST_PIS")), "00")
                                            linha &= "|" & .Item("VL_ITEM")
                                            linha &= "|" & .Item("NAT_BC_CRED")
                                            linha &= "|" & .Item("VL_BC_PIS")
                                            linha &= "|" & CDbl(.Item("ALIQ_PIS")).ToString("N2")
                                            linha &= "|" & .Item("VL_PIS")
                                            linha &= "|" & .Item("COD_CTA")
                                            linha &= "|"
                                        End With

                                        ArquivoAux.Add(linha)
                                        RegistroC501 += 1
                                        RegistroGeral += 1

                                    Next

                                    ' C505 Complemento Energia eletrica COFINS -------------------

                                    Sql = " SELECT 'C505' AS Reg," & vbCrLf &
                                          "        ISNULL(OE.STPisCofins, 0) AS Cst_Pis," & vbCrLf &
                                          "        SUM(ISNULL(CASE WHEN NotasFiscaisXItens.Valor > 0 THEN NotasFiscaisXItens.Valor END, 0)) AS Vl_Item," & vbCrLf &
                                          "        '04' AS Nat_Bc_Cred," & vbCrLf &
                                          "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
                                          "        NotasFiscaisXEncargos.Percentual AS ALiq_Pis," & vbCrLf &
                                          "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END END, 0)) AS Vl_Pis," & vbCrLf &
                                          "		   case" & vbCrLf &
                                          "		   	when NotasFiscais.EntradaSaida_Id = 'E'" & vbCrLf &
                                          "				then OEE.DebitaConta" & vbCrLf &
                                          "				else OEE.CreditaConta" & vbCrLf &
                                          "		   end AS Cod_Cta" & vbCrLf &
                                          "   FROM NotasFiscais " & vbCrLf &
                                          "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                                          "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                                          "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                                          "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
                                          "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
                                          "    AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
                                          "  inner join OperacaoXEstado OE" & vbCrLf &
                                          "     on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                                          "  inner join OperacaoXEstadoXEncargo OEE" & vbCrLf &
                                          "     on OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
                                          "    and OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
                                          "  WHERE (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
                                          "    AND (NotasFiscais.Empresa_Id   ='" & .Item("Cliente_Id") & "')" & vbCrLf &
                                          "    AND (NotasFiscais.Movimento    BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                          "    AND NotasFiscais.Cliente_Id    ='" & Cliente & "' " & vbCrLf &
                                          "    AND NotasFiscais.EndCliente_Id = " & EndCliente & "  " & vbCrLf &
                                          "    And NotasFiscais.Nota_Id       = " & Nota & " " & vbCrLf &
                                          "    AND NotasFiscais.Serie_Id      ='" & Serie & "'" & vbCrLf &
                                          "    AND NotasFiscaisXItens.cfop_Id in (1252, 1253, 2252, 2253, 1555, 1556, 2555, 2556)" & vbCrLf &
                                          "    AND (NotasFiscais.EntradaSaida_Id = 'E')" & vbCrLf &
                                          "  GROUP BY OE.STPisCofins, NotasFiscaisXEncargos.Percentual, NotasFiscais.EntradaSaida_Id, OEE.DebitaConta, OEE.CreditaConta" & vbCrLf &
                                          " HAVING (SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NotasFiscaisXEncargos.Valor END, 0)) > 0) " & vbCrLf

                                    dsAux = Banco.ConsultaDataSet(Sql, "Clientes")
                                    For Each drC505 In dsAux.Tables(0).Rows
                                        With drC505
                                            linha = "|" & .Item("Reg")
                                            linha &= "|" & Format(CInt(.Item("CST_PIS")), "00")
                                            linha &= "|" & .Item("VL_ITEM")
                                            linha &= "|" & .Item("NAT_BC_CRED")
                                            linha &= "|" & .Item("VL_BC_PIS")
                                            linha &= "|" & CDbl(.Item("ALIQ_PIS")).ToString("N2")
                                            linha &= "|" & .Item("VL_PIS")
                                            linha &= "|" & .Item("COD_CTA")
                                            linha &= "|"
                                        End With

                                        ArquivoAux.Add(linha)
                                        RegistroC505 += 1
                                        RegistroGeral += 1
                                    Next
                                Next
                            End With
                        Next   'Fim Estabelecimento Bloco C010
                    End If

                    '' Registro C990  - Encerramento do Bloco C
                    RegistroC990 += 1

                    linha = "|C990"
                    linha &= "|" & RegistroC001 + RegistroC010 + RegistroC100 + RegistroC170 + RegistroC180 + RegistroC181 + RegistroC185 + RegistroC190 + RegistroC191 + RegistroC195 + RegistroC500 + RegistroC501 + RegistroC505 + RegistroC990
                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroGeral += 1

                    '********************************** BLOCO D *************************************************
                    'Registro D001

                    ds = ConsultaD001()

                    RegistroD001 += 1

                    linha = "|D001"
                    If ds.Tables(0).Rows.Count > 0 Then
                        linha &= "|" & "0"
                        TemMovimento = "S"
                    Else
                        linha &= "|" & "1"
                        TemMovimento = "N"
                    End If

                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroGeral += 1

                    If TemMovimento = "S" Then

                        'Consulta dos Registros  D010, D100, D101, D105, D200, D201
                        ds = Consulta_D010_D100_D101_D105_D200_D201_D205_D500_D501_D505()

                        ' Registro D010  - Identificação do Estabelecimento bloco D
                        Dim EmpresaD As String = String.Empty

                        For Each drD010 In ds.Tables("D010").Rows
                            With drD010
                                Dim dsAux As New DataSet
                                linha = "|D010"
                                linha &= "|" & .Item("Cliente_Id")
                                linha &= "|"
                                ArquivoAux.Add(linha)
                                RegistroD010 += 1
                                RegistroGeral += 1

                                EmpresaD = .Item("Cliente_Id")

                                '********* REGISTRO D100 ********************************************
                                '#Registro D100#  - Aquisiçao de serviçoes de Transporte

                                Dim D100Empresa As String = String.Empty
                                Dim D100EndEmpresa As Integer = 0
                                Dim D100NF As Integer = 0
                                Dim D100Serie As String = String.Empty
                                Dim D100EntradaSaida As String = String.Empty
                                Dim D100Cliente As String = String.Empty
                                Dim D100EndCliente As Integer = 0

                                For Each drD100 In ds.Tables("D100").Select("Empresa_Id = " & EmpresaD)

                                    With drD100

                                        D100Empresa = .Item("Empresa_Id")
                                        D100EndEmpresa = .Item("EndEmpresa_Id")
                                        D100NF = .Item("Nota_Id")
                                        D100Serie = .Item("Serie_Id")
                                        D100EntradaSaida = .Item("EntradaSaida_Id")
                                        D100Cliente = Trim(.Item("Cliente_Id"))
                                        D100EndCliente = .Item("EndCliente_Id")

                                        linha = "|" & .Item("Reg")
                                        linha &= "|" & .Item("IND_OPER")
                                        linha &= "|" & .Item("IND_EMIT")
                                        linha &= "|" & LTrim(RTrim(.Item("Cliente_Id"))) & LTrim(RTrim(.Item("Endereco_Id"))) 'COD_PART

                                        'linha &= "|" & .Item("COD_MOD")
                                        If .Item("COD_MOD") = 2 AndAlso .Item("CHV_CTE").Equals("") Then
                                            linha &= "|08"
                                        ElseIf (.Item("COD_MOD") = 2 Or .Item("COD_MOD") = 58) AndAlso .Item("CHV_CTE").ToString.Length > 0 Then
                                            linha &= "|57"
                                        Else
                                            linha &= "|" & CInt(.Item("COD_MOD")).ToString("00")
                                        End If

                                        linha &= "|" & Format(CInt(.Item("COD_SIT")), "00")
                                        linha &= "|" & .Item("SER")
                                        linha &= "|" & ""
                                        linha &= "|" & .Item("NUM_DOC")

                                        If .Item("CHV_CTE").ToString.Length > 0 Then
                                            linha &= "|" & .Item("CHV_CTE")
                                        Else
                                            linha &= "|" '& .Item("CHV_CTE")
                                        End If

                                        linha &= "|" & CDate(.Item("DT_DOC")).ToStrDate().Replace("-", "")
                                        linha &= "|" & CDate(.Item("DT_A_P")).ToStrDate().Replace("-", "")
                                        linha &= "|" & .Item("TP_CT_e")
                                        linha &= "|" & .Item("CHV_CTE_REF")
                                        linha &= "|" & .Item("VL_DOC")
                                        linha &= "|" & .Item("VL_DESC")

                                        linha &= "|" & .Item("IND_FRT")

                                        linha &= "|" & .Item("VL_SERV")
                                        linha &= "|" & .Item("VL_BC_ICMS")
                                        linha &= "|" & .Item("Vl_Vl_ICMS")

                                        If (EmpresaMestre = "05366261" OrElse EmpresaMestre = "38198213" OrElse EmpresaMestre = "40938762" OrElse EmpresaMestre = "62747840" OrElse EmpresaMestre = "62780383" OrElse EmpresaMestre = "63358210") Then
                                            If .Item("VL_BC_ICMS") > .Item("VL_SERV") Then
                                                linha &= "|0,00"
                                            Else
                                                linha &= "|" & .Item("VL_SERV") - .Item("VL_BC_ICMS")
                                            End If
                                        Else
                                            linha &= "|" & .Item("VL_SERV") - .Item("VL_BC_ICMS")
                                        End If

                                        linha &= "|" & ""
                                        linha &= "|" & .Item("COD_CTA")
                                        linha &= "|"

                                    End With
                                    ArquivoAux.Add(linha)
                                    RegistroD100 += 1
                                    RegistroGeral += 1

                                    '****D101 Complemento do Documento de Transporte PIS sobre fretes **********
                                    If ds.Tables("D101").Rows.Count > 0 Then
                                        CompoeRegistroD101(ds, ArquivoAux, RegistroD101, RegistroGeral, D100Empresa, D100EndEmpresa, D100NF, D100Cliente, D100EndCliente, D100Serie, D100EntradaSaida)
                                    End If
                                    '**************************************************************

                                    '****D105 Complemento do Documento de Transporte COFINS sobre fretes **********
                                    If ds.Tables("D105").Rows.Count > 0 Then
                                        CompoeRegistroD105(ds, ArquivoAux, RegistroD105, RegistroGeral, D100Empresa, D100EndEmpresa, D100NF, D100Cliente, D100EndCliente, D100Serie, D100EntradaSaida)
                                    End If
                                    '**************************************************************

                                Next
                                '************************ Fim D100 ******************************************************


                                '********* REGISTRO D200 ********************************************

                                'Bloco D200 Receita com Prestação de Serviço de Transportes
                                Dim D200Empresa As String = String.Empty
                                Dim D200EndEmpresa As Integer = 0
                                Dim D200NF As Integer = 0
                                Dim D200Serie As String = String.Empty
                                Dim D200EntradaSaida As String = String.Empty
                                Dim D200Cliente As String = String.Empty
                                Dim D200EndCliente As Integer = 0

                                For Each drD200 In ds.Tables("D200").Select("Empresa_Id = " & EmpresaD)
                                    With drD200
                                        'Nota = .item("Nota_Id")
                                        'Serie = .item("Serie_Id")
                                        'EntradaSaida = .item("EntradaSaida_Id")

                                        D200Empresa = .Item("Empresa_Id")
                                        'D200EndEmpresa = .Item("EndEmpresa_Id")
                                        'D200NF = .Item("Nota_Id")
                                        'D200Serie = .Item("Serie_Id")
                                        'D200EntradaSaida = .Item("EntradaSaida_Id")
                                        'D200Cliente = Trim(.Item("Cliente_Id"))
                                        'D200EndCliente = .Item("EndCliente_Id")

                                        linha = "|" & .Item("Reg")
                                        linha &= "|" & .Item("COD_MOD")
                                        linha &= "|00" 'Código da Situação do Documento Fiscal
                                        linha &= "|" & .Item("SER")
                                        linha &= "|" & ""
                                        linha &= "|" & .Item("NUM_DOC_INI")
                                        linha &= "|" & .Item("NUM_DOC_FIM")
                                        linha &= "|" & .Item("CFOP")
                                        linha &= "|" & CDate(.Item("DT_REF")).ToString("ddMMyyyy")

                                        linha &= "|" & .Item("VL_DOC")
                                        linha &= "|" & .Item("VL_DESC")
                                        linha &= "|"

                                        ArquivoAux.Add(linha)
                                        RegistroD200 += 1
                                        RegistroGeral += 1
                                    End With

                                    '****D201 Complemento do Documento de Transporte PIS **********
                                    If ds.Tables("D201").Rows.Count > 0 Then
                                        CompoeRegistroD201(ds, ArquivoAux, RegistroD201, RegistroGeral, D200Empresa, drD200("NUM_DOC_INI"), drD200("NUM_DOC_FIM"), CDate(drD200("DT_REF")).ToString("yyyy-MM-dd"), drD200("CFOP"))
                                    End If

                                    '****D205 Complemento do Documento de Transporte Cofins**********
                                    If ds.Tables("D205").Rows.Count > 0 Then
                                        CompoeRegistroD205(ds, ArquivoAux, RegistroD205, RegistroGeral, D200Empresa, drD200("NUM_DOC_INI"), drD200("NUM_DOC_FIM"), CDate(drD200("DT_REF")).ToString("yyyy-MM-dd"), drD200("CFOP"))
                                    End If
                                Next
                                '************************ Fim D200 ******************************************************


                                '********* REGISTRO D500 ********************************************
                                ' D500 Nota Fiscal de Serviços de Comunicação  -------
                                Dim D500Empresa As String = String.Empty
                                Dim D500EndEmpresa As Integer = 0
                                Dim D500NF As Integer = 0
                                Dim D500Serie As String = String.Empty
                                Dim D500EntradaSaida As String = String.Empty
                                Dim D500Cliente As String = String.Empty
                                Dim D500EndCliente As Integer = 0


                                For Each drD500 In ds.Tables("D500").Select("Empresa_Id = " & EmpresaD)
                                    With drD500

                                        D500Empresa = .Item("Empresa_Id")
                                        D500EndEmpresa = .Item("EndEmpresa_Id")
                                        D500NF = .Item("Nota_Id")
                                        D500Serie = .Item("Serie_Id")
                                        D500EntradaSaida = .Item("EntradaSaida_Id")
                                        D500Cliente = Trim(.Item("Cliente_Id"))
                                        D500EndCliente = .Item("EndCliente_Id")

                                        linha = "|" & .Item("Reg")
                                        linha &= "|" & .Item("IND_OPER")
                                        linha &= "|" & .Item("IND_EMIT")
                                        linha &= "|" & .Item("COD_PART")
                                        linha &= "|" & .Item("COD_MOD")
                                        linha &= "|" & Format(CInt(.Item("COD_SIT")), "00")
                                        linha &= "|" & .Item("SER")
                                        linha &= "|" & ""
                                        linha &= "|" & .Item("NUM_DOC")
                                        linha &= "|" & CDate(.Item("DT_DOC")).ToString("ddMMyyyy")
                                        linha &= "|" & CDate(.Item("DT_A_P")).ToString("ddMMyyyy")
                                        linha &= "|" & .Item("VL_DOC")
                                        linha &= "|" & .Item("VL_DESC")
                                        linha &= "|" & .Item("VL_SERV")
                                        linha &= "|" & .Item("VL_SERV_NT")
                                        linha &= "|" & .Item("VL_TERC")
                                        linha &= "|" & .Item("VL_DA")
                                        linha &= "|" & .Item("VL_BC_ICMS")
                                        linha &= "|" & .Item("VL_ICMS")
                                        linha &= "|" & .Item("COD_INF")
                                        linha &= "|" & .Item("VL_PIS")
                                        linha &= "|" & .Item("VL_COFINS")
                                        linha &= "|"
                                    End With

                                    ArquivoAux.Add(linha)
                                    RegistroD500 += 1
                                    RegistroGeral += 1

                                    '****D501 Complemento da Operação Pis **********
                                    If ds.Tables("D501").Rows.Count > 0 Then
                                        CompoeRegistroD501(ds, ArquivoAux, RegistroD501, RegistroGeral, D500Empresa, D500EndEmpresa, D500NF, D500Cliente, D500EndCliente, D500Serie, D500EntradaSaida)
                                    End If

                                    '****D505 Complemento da Operação Cofins **********
                                    If ds.Tables("D501").Rows.Count > 0 Then
                                        CompoeRegistroD505(ds, ArquivoAux, RegistroD505, RegistroGeral, D500Empresa, D500EndEmpresa, D500NF, D500Cliente, D500EndCliente, D500Serie, D500EntradaSaida)
                                    End If
                                Next
                            End With
                        Next    ' Fim Estabelecimento Bloco D
                    End If

                    '' Registro D990  - Encerramento do Bloco D
                    RegistroD990 += 1

                    linha = "|D990"
                    linha &= "|" & RegistroD001 + RegistroD010 + RegistroD100 + RegistroD101 + RegistroD105 + RegistroD200 + RegistroD201 + RegistroD205 + RegistroD500 + RegistroD501 + RegistroD505 + RegistroD990
                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroGeral += 1

                End If 'Fim teste do Regime


                '********* REGISTRO F ********************************************

                ' Registro F001  - Abertura do Bloco F
                dsaux = ConsultaF001()
                linha = "|F001"
                If dsaux.Tables("F001").Rows.Count > 0 Then
                    linha &= "|" & "0"
                Else
                    linha &= "|" & "1"
                End If
                linha &= "|"
                RegistroF001 += 1
                ArquivoAux.Add(linha)
                RegistroGeral += 1
                ' Fim Registro F001


                '*****Consulta dos Registros 
                If dsaux.Tables("F001").Rows.Count > 0 Then
                    ds = Consulta_F010_F100_F120_F130_F550_F600()
                End If

                Dim F010Empresa As String = String.Empty

                'Registro F010 - Identificação do Estabelecimento
                For Each drF010 In dsaux.Tables("F001").Rows
                    With drF010

                        F010Empresa = .Item("Empresa_Id")

                        linha = "|F010"
                        linha &= "|" & .Item("Empresa_Id")
                        linha &= "|"

                        ArquivoAux.Add(linha)
                        RegistroF010 += 1
                        RegistroGeral += 1


                        'Registro F100 
                        'Outros Creditos e Razão Contábil -------------------------
                        CompoeRegistroF100(ds, ArquivoAux, RegistroF100, RegistroGeral, F010Empresa)

                        'Registro F120: 
                        'BENS INCORPORADOS AO ATIVO IMOBILIZADO – OPERAÇÕES GERADORAS DE CRÉDITOS COM BASE NOS ENCARGOS DE DEPRECIAÇÃO E AMORTIZAÇÃO Registro 
                        If cboIncidencia.SelectedValue <> 2 Then
                            CompoeRegistroF120(ds, ArquivoAux, RegistroF120, RegistroGeral, F010Empresa)
                        End If

                        'Registro F130 
                        '- Depreciações do Ativo Imobilizado --------
                        If cboIncidencia.SelectedValue <> 2 Then
                            CompoeRegistroF130(ds, ArquivoAux, RegistroF130, RegistroGeral, F010Empresa)
                        End If

                        'Registro F550 
                        'Receitas de Alugueis ------------------------
                        If cboIncidencia.SelectedValue = 2 Then
                            CompoeRegistroF550(ds, ArquivoAux, RegistroF550, RegistroGeral, F010Empresa)
                        End If

                        'Registro F600 
                        ' Contribuições retidas na fonte
                        CompoeRegistroF600(ds, ArquivoAux, RegistroF600, RegistroGeral, F010Empresa)

                    End With 'Fim Registro de Estabelecimentos F010
                Next 'Fim F010-----------------------------------------------

                '' Registro F990  - Encerramento do Bloco F
                RegistroF990 += 1

                linha = "|F990"
                linha &= "|" & RegistroF001 + RegistroF010 + RegistroF100 + RegistroF120 + RegistroF130 + RegistroF550 + RegistroF600 + RegistroF990
                linha &= "|"

                ArquivoAux.Add(linha)
                RegistroGeral += 1

                ' Registro M001  - Abertura do Bloco M
                If cboIncidencia.SelectedValue = 1 Then
                    linha = "|M001"
                    linha &= "|0"
                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroM001 += 1
                    RegistroGeral += 1

                    'Dim RegistroGeralAux As Integer

                    ArquivoPadrao = ArquivoAux.Clone

                    'BuscaValorRetidoFonte()
                    Bloco_M(ArquivoAux)

                    ' Registro M990  - Encerramento do Bloco M
                    RegistroM990 += 1
                Else
                    linha = "|M001"
                    linha &= "|1"
                    linha &= "|"
                    ArquivoAux.Add(linha)
                    RegistroM001 += 1
                    RegistroGeral += 1

                    RegistroM990 += 1
                End If


                linha = "|M990"
                linha &= "|" & RegistroM001 + RegistroM100 + RegistroM105 + RegistroM110 + RegistroM200 + RegistroM205 + RegistroM210 + RegistroM400 + RegistroM410 + RegistroM500 + RegistroM505 + RegistroM510 + RegistroM600 + RegistroM605 + RegistroM610 + RegistroM800 + RegistroM810 + RegistroM990
                linha &= "|"

                ArquivoAux.Add(linha)
                RegistroM990 += 1
                RegistroGeral += 1


                '-------------------------------------------
                ' Registro P001  - Abertura do Bloco P
                '-------------------------------------------
                If (EmpresaMestre = "11111111") Then ' Isolado Temporariamente - Utilizado somente para o Curtume.
                    Sql = "  Select	Empresa_Id, SUM(VL_REC_TOT_EST) as VL_REC_TOT_EST, COD_ATIV_ECON, SUM(VL_REC_ATIV_ESTAB) as VL_REC_ATIV_ESTAB" & vbCrLf &
                          " From (" & vbCrLf &
                          " Select	Razao.Empresa_Id, SUM(Razao.CreditoOficial - Razao.DebitoOficial) as  VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                          "        0 as  VL_REC_ATIV_ESTAB" & vbCrLf &
                          " From razao " & vbCrLf &
                          " Where Left(Empresa_Id, 8) Like '" & EmpresaMestre & "' And Conta_Id like '30101%' " & vbCrLf &
                          "    And (left(Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Produto in ('70101003',  '70101002'))" & vbCrLf &
                          "    And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                          "    And not Lote_Id in (7500) " & vbCrLf &
                          "    Group By Razao.Empresa_Id" & vbCrLf &
                          "    Union" & vbCrLf &
                          " SELECT  Razao.Empresa_Id, 	0 AS VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                          " 	     SUM(Razao.CreditoOficial - Razao.DebitoOficial) AS VL_REC_ATIV_ESTAB" & vbCrLf &
                          " FROM    Razao INNER JOIN" & vbCrLf &
                          "         Produtos ON Razao.Produto = Produtos.Produto_Id" & vbCrLf &
                          " Where	 Left(Empresa_Id, 8) Like '" & EmpresaMestre & "' And  (Razao.Conta_Id like '30101%' Or Razao.Conta_Id in ('301020101', '301020102', '301040106'))" & vbCrLf &
                          "        And Not Razao.Conta_Id like '3010102%'" & vbCrLf &
                          " 		And (left(Razao.Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Razao.Produto in ('70101003', '70101002'))" & vbCrLf &
                          " 		And (Razao.Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                          "        And Produtos.NCM Like '4104%'" & vbCrLf &
                          "        And not Lote_Id in (7500) " & vbCrLf &
                          " Group By Razao.Empresa_Id) as Consulta" & vbCrLf &
                          " Group By Empresa_Id, COD_ATIV_ECON " & vbCrLf


                    Dim dsP001 As New DataSet
                    Dim BlocoComMovimento As String

                    dsP001 = Banco.ConsultaDataSet(Sql, "RegistroP001")

                    RegistroP001 += 1

                    linha = "|P001"
                    If dsP001.Tables(0).Rows.Count > 0 Then
                        linha &= "|" & "0"
                        BlocoComMovimento = "S"
                    Else
                        linha &= "|" & "1"
                        BlocoComMovimento = "N"
                    End If

                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroGeral += 1

                    If BlocoComMovimento = "S" Then
                        '-------------------------------------
                        ' Registro P010  - Identificação do Estabelecimento bloco P
                        '-------------------------------------

                        Sql = "  Select	Empresa_Id, " & vbCrLf &
                              "        SUM(VL_REC_TOT_EST) as VL_REC_TOT_EST," & vbCrLf &
                              " 		COD_ATIV_ECON," & vbCrLf &
                              " 		SUM(VL_REC_ATIV_ESTAB) as VL_REC_ATIV_ESTAB," & vbCrLf &
                              " 	    SUM(VL_REC_TOT_EST - VL_REC_ATIV_ESTAB) AS VL_EXC," & vbCrLf &
                              " 	    SUM(VL_REC_ATIV_ESTAB) AS VL_BC_CONT," & vbCrLf &
                              " 	    1 AS ALIQ_CONT," & vbCrLf &
                              " 	    SUM((VL_REC_ATIV_ESTAB * 1) / 100) AS	VL_CONT_APU," & vbCrLf &
                              "        '301030104' AS COD_CTA" & vbCrLf &
                              " From (" & vbCrLf &
                              " Select	Razao.Empresa_Id, SUM(Razao.CreditoOficial - Razao.DebitoOficial) as  VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                              "        0 as  VL_REC_ATIV_ESTAB" & vbCrLf &
                              " From razao " & vbCrLf &
                              " Where Left(Empresa_Id, 8) Like '" & EmpresaMestre & "' And (Conta_Id like '30101%'  Or Conta_Id = '301040106')" & vbCrLf &
                              "    And (left(Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Produto in ('70101003',  '70101002'))" & vbCrLf &
                              "    And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                              "    And not Lote_Id in (7500) " & vbCrLf &
                              "    Group By Razao.Empresa_Id" & vbCrLf &
                              "    Union" & vbCrLf &
                              " SELECT  Razao.Empresa_Id, 	0 AS VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                              " 	     SUM(Razao.CreditoOficial - Razao.DebitoOficial) AS VL_REC_ATIV_ESTAB" & vbCrLf &
                              " FROM    Razao INNER JOIN" & vbCrLf &
                              "         Produtos ON Razao.Produto = Produtos.Produto_Id" & vbCrLf &
                              " Where	 Left(Empresa_Id, 8) Like '" & EmpresaMestre & "' And  (Razao.Conta_Id like '30101%' Or Razao.Conta_Id in ('301020101', '301020102', '301040106'))" & vbCrLf &
                              "        And Not Razao.Conta_Id like '3010102%'" & vbCrLf &
                              " 		And (left(Razao.Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Razao.Produto in ('70101003', '70101002'))" & vbCrLf &
                              " 		And (Razao.Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                              "        And Produtos.NCM Like '4104%'" & vbCrLf &
                              "        And not Lote_Id in (7500) " & vbCrLf &
                              " Group By Razao.Empresa_Id) as Consulta" & vbCrLf &
                              " Group By Empresa_Id, COD_ATIV_ECON " & vbCrLf

                        ds = Banco.ConsultaDataSet(Sql, "Clientes")

                        For Each drP010 In ds.Tables(0).Rows
                            With drP010
                                Dim dsAux As New DataSet
                                linha = "|P010"
                                linha &= "|" & .Item("Empresa_Id")
                                linha &= "|"

                                ArquivoAux.Add(linha)
                                RegistroP010 += 1
                                RegistroGeral += 1

                                Sql = "  Select	Empresa_Id, " & vbCrLf &
                                      "    SUM(VL_REC_TOT_EST) as VL_REC_TOT_EST," & vbCrLf &
                                      "    COD_ATIV_ECON," & vbCrLf &
                                      "    SUM(VL_REC_ATIV_ESTAB) as VL_REC_ATIV_ESTAB," & vbCrLf &
                                      "    SUM(VL_EXC) AS VL_EXC," & vbCrLf &
                                      "    SUM(VL_REC_ATIV_ESTAB - VL_EXC) AS VL_BC_CONT," & vbCrLf &
                                      "    1 AS ALIQ_CONT," & vbCrLf &
                                      "    SUM(((VL_REC_ATIV_ESTAB - VL_EXC) * 1) / 100) AS VL_CONT_APU," & vbCrLf &
                                      "    '301030104' AS COD_CTA" & vbCrLf &
                                      " From (" & vbCrLf &
                                      " Select	Razao.Empresa_Id, SUM(Razao.CreditoOficial - Razao.DebitoOficial) as  VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                                      "        0 as  VL_REC_ATIV_ESTAB, 0 as  VL_EXC" & vbCrLf &
                                      " From razao " & vbCrLf &
                                      " Where Empresa_Id = '" & drP010("Empresa_Id") & "' And (Conta_Id like '30101%' Or Conta_Id = '301040106')" & vbCrLf &
                                      "    And (left(Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Produto in ('70101003',  '70101002'))" & vbCrLf &
                                      "    And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                      "    And not Lote_Id in (7500) " & vbCrLf &
                                      "    Group By Razao.Empresa_Id" & vbCrLf &
                                      "    Union" & vbCrLf &
                                      " SELECT  Razao.Empresa_Id, 	0 AS VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                                      " 	     SUM(Razao.CreditoOficial) AS VL_REC_ATIV_ESTAB, 0 AS VL_EXC" & vbCrLf &
                                      " FROM    Razao INNER JOIN" & vbCrLf &
                                      "         Produtos ON Razao.Produto = Produtos.Produto_Id" & vbCrLf &
                                      " Where	 Empresa_Id = '" & drP010("Empresa_Id") & "'  And  (Razao.Conta_Id like '30101%' Or Razao.Conta_Id in ('301020101', '301020102', '301040106'))" & vbCrLf &
                                      " 		And (left(Razao.Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Razao.Produto in ('70101003', '70101002'))" & vbCrLf &
                                      " 		And (Razao.Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                      "        And Produtos.NCM Like '4104%'" & vbCrLf &
                                      "        And not Lote_Id in (7500) " & vbCrLf &
                                      " Group By Razao.Empresa_Id" & vbCrLf &
                                      " Union" & vbCrLf &
                                      " SELECT  Razao.Empresa_Id, " & vbCrLf &
                                      "		 0 AS VL_REC_TOT_EST, " & vbCrLf &
                                      "        '41040000' as COD_ATIV_ECON," & vbCrLf &
                                      "		 0 AS VL_REC_ATIV_ESTAB, " & vbCrLf &
                                      "		 SUM(Razao.CreditoOficial + Razao.DebitoOficial) AS VL_EXC" & vbCrLf &
                                      " FROM   Razao" & vbCrLf &
                                      " Where	 Empresa_Id = '" & drP010("Empresa_Id") & "'  And   left(Razao.Conta_Id, 7) in ('3010102', '3010201')" & vbCrLf &
                                      "          And (Razao.Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                      " And not Lote_Id in (7500) " & vbCrLf &
                                      " Group By Razao.Empresa_Id" & vbCrLf &
                                      " ) as Consulta Group By Empresa_Id, COD_ATIV_ECON " & vbCrLf

                                Dim dsP100 As New DataSet
                                dsP100 = Banco.ConsultaDataSet(Sql, "RegistroP100")
                                For Each drP100 In dsP100.Tables(0).Rows

                                    linha = "|" & "P100"
                                    linha &= "|" & CDate(PeriodoInicial).ToStrDate().Replace("-", "")
                                    linha &= "|" & CDate(PeriodoFinal).ToStrDate().Replace("-", "")
                                    linha &= "|" & drP100("Vl_Rec_Tot_Est")
                                    linha &= "|" & drP100("Cod_Ativ_Econ")
                                    linha &= "|" & drP100("Vl_Rec_Ativ_Estab")
                                    linha &= "|" & drP100("Vl_Exc")
                                    linha &= "|" & drP100("Vl_Bc_Cont")
                                    linha &= "|" & FormatNumber(drP100("Aliq_Cont"), 4)
                                    linha &= "|" & FormatNumber(drP100("Vl_Cont_Apu"), 2).ToString.Replace(".", "")
                                    linha &= "|" & drP100("Cod_Cta")
                                    linha &= "|" '& drP100("Info_Compl")
                                    linha &= "|"
                                    ArquivoAux.Add(linha)
                                    REgistroP100 += 1
                                    RegistroGeral += 1

                                Next

                            End With
                        Next

                        '---- P200 - Consolidação da Contribuição Previdenciária sobre a Receita Bruta

                        Sql = "  Select	Empresa_Id, " & vbCrLf &
                              "    SUM(VL_REC_TOT_EST) as VL_REC_TOT_EST," & vbCrLf &
                              "    COD_ATIV_ECON," & vbCrLf &
                              "    SUM(VL_REC_ATIV_ESTAB) as VL_REC_ATIV_ESTAB," & vbCrLf &
                              "    SUM(VL_EXC) AS VL_EXC," & vbCrLf &
                              "    SUM(VL_REC_ATIV_ESTAB - VL_EXC) AS VL_BC_CONT," & vbCrLf &
                              "    1 AS ALIQ_CONT," & vbCrLf &
                              "    SUM(((VL_REC_ATIV_ESTAB - VL_EXC) * 1) / 100) AS VL_CONT_APU," & vbCrLf &
                              "    '301030104' AS COD_CTA" & vbCrLf &
                              " From (" & vbCrLf &
                              " Select	Left(Razao.Empresa_Id, 8) as Empresa_Id, SUM(Razao.CreditoOficial - Razao.DebitoOficial) as  VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                              "        0 as  VL_REC_ATIV_ESTAB, 0 as  VL_EXC" & vbCrLf &
                              " From razao " & vbCrLf &
                              " Where Left(Empresa_Id, 8) Like '" & EmpresaMestre & "'  And (Conta_Id like '30101%'  Or Conta_Id = '301040106')" & vbCrLf &
                              "    And (left(Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Produto in ('70101003',  '70101002'))" & vbCrLf &
                              "    And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                              "    And not Lote_Id in (7500) " & vbCrLf &
                              "    Group By  Left(Razao.Empresa_Id, 8)" & vbCrLf &
                              "    Union" & vbCrLf &
                              " SELECT  Left(Razao.Empresa_Id, 8) as Empresa_Id, 	0 AS VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                              " 	     SUM(Razao.CreditoOficial) AS VL_REC_ATIV_ESTAB, 0 AS VL_EXC" & vbCrLf &
                              " FROM    Razao INNER JOIN" & vbCrLf &
                              "         Produtos ON Razao.Produto = Produtos.Produto_Id" & vbCrLf &
                              " Where	Left(Empresa_Id, 8) Like '" & EmpresaMestre & "'  And  (Razao.Conta_Id like '30101%' Or Razao.Conta_Id in ('301020101', '301020102'))" & vbCrLf &
                              " 		And (left(Razao.Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Razao.Produto in ('70101003', '70101002'))" & vbCrLf &
                              " 		And (Razao.Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                              "        And not Lote_Id in (7500) " & vbCrLf &
                              "        And Produtos.NCM Like '4104%'" & vbCrLf &
                              " Group By Left(Razao.Empresa_Id, 8)" & vbCrLf &
                              " Union" & vbCrLf &
                              " SELECT  Left(Razao.Empresa_Id, 8) as Empresa_Id," & vbCrLf &
                              "		 0 AS VL_REC_TOT_EST, " & vbCrLf &
                              "        '41040000' as COD_ATIV_ECON," & vbCrLf &
                              "		 0 AS VL_REC_ATIV_ESTAB, " & vbCrLf &
                              "		 SUM(Razao.CreditoOficial + Razao.DebitoOficial) AS VL_EXC" & vbCrLf &
                              " FROM   Razao" & vbCrLf &
                              " Where	 Left(Empresa_Id, 8) Like '" & EmpresaMestre & "' And   left(Razao.Conta_Id, 7) in ('3010102', '3010201')" & vbCrLf &
                              "          And Razao.Movimento_Id between '2014-07-01' And '2014-07-31'" & vbCrLf &
                              " And not Lote_Id in (7500) " & vbCrLf &
                              " Group By Razao.Empresa_Id" & vbCrLf &
                              " ) as Consulta Group By  Empresa_Id, COD_ATIV_ECON " & vbCrLf

                        Dim dsP200 As New DataSet
                        dsP200 = Banco.ConsultaDataSet(Sql, "RegistroP100")
                        For Each drP200 In dsP200.Tables(0).Rows

                            linha = "|" & "P200"
                            linha &= "|" & CDate(PeriodoInicial).ToString("MM/yyyy").Replace("/", "")
                            linha &= "|" & FormatNumber(drP200("Vl_Cont_Apu"), 2).ToString.Replace(".", "")
                            linha &= "|" & 0
                            linha &= "|" & 0
                            linha &= "|" & FormatNumber(drP200("Vl_Cont_Apu"), 2).ToString.Replace(".", "")
                            linha &= "|" & "299101"
                            linha &= "|"
                            ArquivoAux.Add(linha)
                            RegistroP200 += 1
                            RegistroGeral += 1
                        Next
                    End If

                    '-------------------------------------------
                    ' Registro P990  - Encerramento do Bloco P
                    '-------------------------------------------
                    RegistroP990 += 1
                    linha = "|P990"
                    linha &= "|" & RegistroP001 + RegistroP010 + REgistroP100 + RegistroP200 + RegistroP990
                    linha &= "|"

                    ArquivoAux.Add(linha)
                    RegistroGeral += 1

                End If 'Fim do Bloco P


                '*********************** BLOCO 1 *********************************************** 
                'COMPLEMENTO DA ESCRITURAÇÃO – CONTROLE DE SALDOS DE CRÉDITOS E DE RETENÇÕES, OPERAÇÕES EXTEMPORÂNEAS E OUTRAS INFORMAÇÕES  

                '**Registro 1001  - Abertura do Bloco 1**
                'Essa consulta será utilizada para verificação de conteúdo para o bloco 
                'e caso seja 2 - Incidência Acumulativa será utilizado para composição do registro 1900
                dsaux = ConsultaRegistro1001()

                linha = "|1001"
                If dsaux.Tables(0).Rows.Count > 0 Then
                    linha &= "|" & "0"
                Else
                    linha &= "|" & "1"
                End If
                linha &= "|"

                ArquivoAux.Add(linha)
                Registro1001 += 1
                RegistroGeral += 1
                '***************************************
                Dim Empresa1001 As String = String.Empty

                If dsaux.Tables(0).Rows.Count > 0 Then
                    ds = Consulta_1100_1300_1500_1700_1900()

                    For Each dt1001 In dsaux.Tables("1001").Rows
                        With dt1001
                            Empresa1001 = .Item("Empresa_Id")

                            '***REGISTRO 1100: CONTROLE DE CRÉDITOS FISCAIS – PIS/PASEP********
                            CompoeRegistro1100(ds, ArquivoAux, Registro1100, RegistroGeral, Empresa1001)
                            '************************

                            '***Registro 1300 - Controle de Saldos PIS Retido na Fonte*********
                            CompoeRegistro1300(ds, ArquivoAux, Registro1300, RegistroGeral, Empresa1001)
                            '************************

                            '***Registro 1500 - Controle de Créditos Fiscais - Cofins *********
                            CompoeRegistro1500(ds, ArquivoAux, Registro1500, RegistroGeral, Empresa1001)
                            '************************

                            '***Registro 1700 - Controle de Saldos Cofins Retido na Fonte *********
                            CompoeRegistro1700(ds, ArquivoAux, Registro1700, RegistroGeral, Empresa1001)
                            '************************

                            '***Registro 1900 - Empresas com Lucro Presumido ------
                            CompoeRegistro1900(ds, ArquivoAux, Registro1900, RegistroGeral, Empresa1001)
                            '************************
                        End With
                    Next
                End If '- Fim Do Regime Cumulativo

                ' Registro M990  - Encerramento do Bloco 1
                Registro1990 += 1
                Registro1990 = Registro1001 + Registro1100 + Registro1300 + Registro1500 + Registro1700 + Registro1900 + Registro1990

                linha = "|1990"
                linha = linha & "|" & Registro1990
                linha = linha & "|"

                ArquivoAux.Add(linha)
                Registro1990 += 1
                RegistroGeral += 1

                ' Registro 9001  - Abertura do Bloco 9
                linha = "|9001"
                linha &= "|0"
                linha &= "|"

                ArquivoAux.Add(linha)
                Registro9001 += 1
                RegistroGeral += 1

                ' Registro 9900  - Registros Do Arquivo

                For i = 1 To 91
                    linha = "|9900"
                    Select Case i
                        Case 1
                            linha &= "|0000"
                            linha &= "|" & Registro0000
                        Case 2
                            linha &= "|0001"
                            linha &= "|" & Registro0001
                        Case 3
                            linha &= "|0100"
                            linha &= "|" & Registro0100
                        Case 4
                            linha &= "|0110"
                            linha &= "|" & Registro0110
                        Case 5
                            linha &= "|0111"
                            linha &= "|" & Registro0111
                        Case 6
                            linha &= "|0140"
                            linha &= "|" & Registro0140
                        Case 7
                            linha &= "|0145"
                            linha &= "|" & Registro0145
                        Case 8
                            linha &= "|0150"
                            linha &= "|" & Registro0150
                        Case 9
                            linha &= "|0190"
                            linha &= "|" & Registro0190
                        Case 10
                            linha &= "|0200"
                            linha &= "|" & Registro0200
                        Case 11
                            linha &= "|0400"
                            linha &= "|" & Registro0400
                        Case 12
                            linha &= "|0450"
                            linha &= "|" & Registro0450
                        Case 13
                            linha &= "|0500"
                            linha &= "|" & Registro0500
                        Case 14
                            linha &= "|0600"
                            linha &= "|" & Registro0600
                        Case 15
                            linha &= "|0007"
                            linha &= "|" & Registro0007
                        Case 16
                            linha &= "|0020"
                            linha &= "|" & Registro0020
                        Case 17
                            linha &= "|0180"
                            linha &= "|" & Registro0180
                        Case 18
                            linha &= "|0990"
                            linha &= "|" & Registro0990
                        Case 19
                            linha &= "|A001"
                            linha &= "|" & RegistroA001
                        Case 20
                            linha &= "|A010"
                            linha &= "|" & RegistroA010
                        Case 21
                            linha &= "|A100"
                            linha &= "|" & RegistroA100
                        Case 22
                            linha &= "|A170"
                            linha &= "|" & RegistroA170
                        Case 23
                            linha &= "|A990"
                            linha &= "|" & RegistroA990
                        Case 24
                            linha &= "|C001"
                            linha &= "|" & RegistroC001
                        Case 25
                            linha &= "|C010"
                            linha &= "|" & RegistroC010
                        Case 26
                            linha &= "|C100"
                            linha &= "|" & RegistroC100
                        Case 27
                            linha &= "|C170"
                            linha &= "|" & RegistroC170
                        Case 28
                            linha &= "|C180"
                            linha &= "|" & RegistroC180
                        Case 29
                            linha &= "|C181"
                            linha &= "|" & RegistroC181
                        Case 30
                            linha &= "|C185"
                            linha &= "|" & RegistroC185
                        Case 31
                            linha &= "|C190"
                            linha &= "|" & RegistroC190
                        Case 32
                            linha &= "|C191"
                            linha &= "|" & RegistroC191
                        Case 33
                            linha &= "|C195"
                            linha &= "|" & RegistroC195
                        Case 34
                            linha &= "|C500"
                            linha &= "|" & RegistroC500
                        Case 35
                            linha &= "|C501"
                            linha &= "|" & RegistroC501
                        Case 36
                            linha &= "|C505"
                            linha &= "|" & RegistroC505
                        Case 37
                            linha &= "|C990"
                            linha &= "|" & RegistroC990
                        Case 38
                            linha &= "|D001"
                            linha &= "|" & RegistroD001
                        Case 39
                            linha &= "|D010"
                            linha &= "|" & RegistroD010
                        Case 40
                            linha &= "|D100"
                            linha &= "|" & RegistroD100
                        Case 41
                            linha &= "|D101"
                            linha &= "|" & RegistroD101
                        Case 42
                            linha &= "|D105"
                            linha &= "|" & RegistroD105

                        Case 43
                            linha &= "|D200"
                            linha &= "|" & RegistroD200
                        Case 44
                            linha &= "|D201"
                            linha &= "|" & RegistroD201
                        Case 45
                            linha &= "|D205"
                            linha &= "|" & RegistroD205

                        Case 46
                            linha &= "|D500"
                            linha &= "|" & RegistroD500
                        Case 47
                            linha &= "|D501"
                            linha &= "|" & RegistroD501
                        Case 48
                            linha &= "|D505"
                            linha &= "|" & RegistroD505
                        Case 49
                            linha &= "|D990"
                            linha &= "|" & RegistroD990
                        Case 50
                            linha &= "|F001"
                            linha &= "|" & RegistroF001
                        Case 51
                            linha &= "|F010"
                            linha &= "|" & RegistroF010
                        Case 52
                            linha &= "|F100"
                            linha &= "|" & RegistroF100
                        Case 53
                            linha &= "|F120"
                            linha &= "|" & RegistroF120

                        Case 54
                            linha &= "|F130"
                            linha &= "|" & RegistroF130
                        Case 55
                            linha &= "|F550"
                            linha &= "|" & RegistroF550
                        Case 56
                            linha &= "|F600"
                            linha &= "|" & RegistroF600
                        Case 57
                            linha &= "|F990"
                            linha &= "|" & RegistroF990

                        Case 58
                            linha &= "|M001"
                            linha &= "|" & RegistroM001
                        Case 59
                            linha &= "|M100"
                            linha &= "|" & RegistroM100
                        Case 60
                            linha &= "|M105"
                            linha &= "|" & RegistroM105
                        Case 61
                            linha &= "|M110"
                            linha &= "|" & RegistroM110
                        Case 62
                            linha &= "|M200"
                            linha &= "|" & RegistroM200

                        Case 63
                            linha &= "|M205"
                            linha &= "|" & RegistroM205

                        Case 64
                            linha &= "|M210"
                            linha &= "|" & RegistroM210
                        Case 65
                            linha &= "|M400"
                            linha &= "|" & RegistroM400
                        Case 66
                            linha &= "|M410"
                            linha &= "|" & RegistroM410
                        Case 67
                            linha &= "|M500"
                            linha &= "|" & RegistroM500
                        Case 68
                            linha &= "|M505"
                            linha &= "|" & RegistroM505
                        Case 69
                            linha &= "|M510"
                            linha &= "|" & RegistroM510
                        Case 70
                            linha &= "|M600"
                            linha &= "|" & RegistroM600
                        Case 71
                            linha &= "|M605"
                            linha &= "|" & RegistroM600

                        Case 72
                            linha &= "|M610"
                            linha &= "|" & RegistroM610
                        Case 73
                            linha &= "|M800"
                            linha &= "|" & RegistroM800
                        Case 74
                            linha &= "|M810"
                            linha &= "|" & RegistroM810
                        Case 75
                            linha &= "|M990"
                            linha &= "|" & RegistroM990

                        Case 76
                            linha &= "|P001"
                            linha &= "|" & RegistroP001
                        Case 77
                            linha &= "|P010"
                            linha &= "|" & RegistroP010
                        Case 78
                            linha &= "|P100"
                            linha &= "|" & REgistroP100
                        Case 79
                            linha &= "|P200"
                            linha &= "|" & RegistroP200
                        Case 80
                            linha &= "|P990"
                            linha &= "|" & RegistroP990

                        Case 81
                            linha &= "|1001"
                            linha &= "|" & Registro1001
                        Case 82
                            linha &= "|1100"
                            linha &= "|" & Registro1100
                        Case 83
                            linha &= "|1300"
                            linha &= "|" & Registro1300
                        Case 84
                            linha &= "|1500"
                            linha &= "|" & Registro1500
                        Case 85
                            linha &= "|1700"
                            linha &= "|" & Registro1700
                        Case 86
                            linha &= "|1900"
                            linha &= "|" & Registro1900
                        Case 87
                            linha &= "|1990"
                            linha &= "|1" '& Registro1990
                        Case 88
                            linha &= "|9001"
                            linha &= "|" & Registro9001
                        Case 89
                            linha &= "|9900"
                            linha &= "|" & Registro9900 + 3
                        Case 90
                            linha &= "|9990"
                            linha &= "|1"
                        Case 91
                            linha &= "|9999"
                            linha &= "|1"
                    End Select

                    linha &= "|"

                    ArquivoAux.Add(linha)
                    Registro9900 += 1
                    RegistroGeral += 1
                Next

                ' Registro 9990  - Encerramento do Bloco 9
                Registro9990 += 1

                linha = "|9990"
                linha &= "|" & Registro9001 + Registro9900 + Registro9990 + 1
                linha &= "|"

                ArquivoAux.Add(linha)
                RegistroGeral += 1

                ' Registro 9999  - Encerramento do Arquivo Digital

                Registro9999 += 1
                RegistroGeral += 1

                linha = "|9999"
                linha &= "|" & RegistroGeral
                linha &= "|"

                ArquivoAux.Add(linha)

                For Each aux In ArquivoAux
                    sw.WriteLine(aux)
                Next

                sw.Close()
                sw.Dispose()
                imdDownload.Visible = True
                ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "downloadArquivo();", True)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                sw.Close()
                sw.Dispose()
            End Try
        End Using
    End Sub

    Protected Sub lnkProcessar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkProcessar.Click
        Try
            If Funcoes.VerificaPermissao("SpedPisCofins", "RELATORIO") Then
                Processar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para processar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SpedPisCofins")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#Region "BLOCO 0"
    Private Function Consulta0000() As DataSet
        Dim SqlT As String = String.Empty

        'Registro 0000 - Abertura do Arquivo Digital e Identificação da Pessoa Juridica

        Dim objEmpresa = New Cliente(Empresa(0), 0)

        SqlT = " SELECT Cli.Cliente_Id, Cli.Nome, Cli.Suframa," & vbCrLf &
               "        Cli.CodigoDoMunicipio, Cli.Estado, Mun.EstadoIbge, " & vbCrLf &
               "        C.Nome AS ContadorNome," & vbCrLf &
               "        C.Cliente_Id AS ContadorCPF, " & vbCrLf &
               "        CxE.CrcContador AS ContadorCRC, " & vbCrLf &
               "        C.Cep AS ContadorCep," & vbCrLf &
               "        C.Endereco AS ContadorEndereco," & vbCrLf &
               "        C.Numero AS ContadorNumero," & vbCrLf &
               "        C.Complemento AS ContadorComplemento," & vbCrLf &
               "        C.Bairro AS ContadorBairro," & vbCrLf &
               "        C.Telefone AS ContadorTelefone," & vbCrLf &
               "        C.Fax AS ContadorFax," & vbCrLf &
               "        C.Email AS ContadorEmail," & vbCrLf &
               "        C.CodigoDoMunicipio AS ContadorCodigoDoMunicipio," & vbCrLf &
               "        CMun.EstadoIBGE AS ContadorEstadoIbge" & vbCrLf &
               "   FROM Clientes Cli " & vbCrLf &
               "  INNER JOIN ClientesXEmpresas CxE" & vbCrLf &
               "     ON Cli.Cliente_Id  = CxE.Empresa_Id  " & vbCrLf &
               "    AND Cli.Endereco_Id = CxE.EndEmpresa_Id" & vbCrLf &
               "  INNER JOIN Municipios Mun" & vbCrLf &
               "     ON Cli.Estado            = Mun.Estado_id" & vbCrLf &
               "    AND Cli.Cidade            = Mun.Municipio_id" & vbCrLf &
               "    AND Cli.CodigoDoMunicipio = Mun.Codigo_id " & vbCrLf &
               "  INNER JOIN Clientes C" & vbCrLf &
               "     ON CxE.CPFContador = C.Cliente_Id" & vbCrLf &
               "    AND C.Endereco_Id = 0" & vbCrLf &
               "  INNER JOIN Municipios CMun" & vbCrLf &
               "     ON C.Estado            = CMun.Estado_id" & vbCrLf &
               "    AND C.Cidade            = CMun.Municipio_id" & vbCrLf &
               "    AND C.CodigoDoMunicipio = CMun.Codigo_id " & vbCrLf &
               "  WHERE Cli.Cliente_Id  = '" & Empresa(0) & "'" & vbCrLf &
               "    AND Cli.Endereco_Id = 0" & vbCrLf
        Return Banco.ConsultaDataSet(SqlT, "Consulta")
    End Function

    Private Sub CompoeRegistro0000(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro0000 As Integer, ByRef RegistroGeral As Integer)

        ''Registro 0000 - Abertura do Arquivo Digital e Identificação da Pessoa Juridica
        For Each dr In ds.Tables(0).Rows
            With dr
                linha = "|0000"

                If CDate(txtDataFinal.Text) < "2012/07/01" Then
                    linha &= "|002"
                ElseIf CDate(txtDataFinal.Text) < "2018/06/01" Then
                    linha &= "|003"
                ElseIf CDate(txtDataFinal.Text) < "2019/01/01" Then
                    linha &= "|004"
                ElseIf CDate(txtDataFinal.Text) < "2020/01/01" Then
                    linha &= "|005"
                Else
                    linha &= "|006"
                End If

                linha &= "|" & cboTipoEscrituracao.SelectedValue
                linha &= "|" & DdlIndSituacaoEsp.SelectedValue

                If cboTipoEscrituracao.SelectedValue = 1 Then
                    linha &= "|" & txt_NUM_ORD.Text
                Else
                    linha &= "|"
                End If
                linha &= "|" & CDate(txtDataInicial.Text).ToString("ddMMyyyy")
                linha &= "|" & CDate(txtDataFinal.Text).ToString("ddMMyyyy")
                linha &= "|" & RTrim(.Item("Nome"))
                linha &= "|" & Left(.Item("Cliente_Id"), 14)
                linha &= "|" & .Item("Estado")
                linha &= "|" & .Item("EstadoIbge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                linha &= "|" & RTrim(Funcoes.AlinharEsquerda(.Item("Suframa").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))
                linha &= "|" & ddlIndicadorDeNatureza.SelectedValue
                linha &= "|" & ddlIndicadorAtividade.SelectedValue
                linha &= "|"
                'CpfContador = .Item("CpfContador")
                'CRCContador = .Item("CrcContador")
            End With

            ArquivoAux.Add(linha)
            Registro0000 += 1
            RegistroGeral += 1
        Next
    End Sub

    Private Sub CompoeRegistro0001(ByRef ArquivoAux As ArrayList, ByRef Registro0001 As Integer, ByRef RegistroGeral As Integer)

        ''Registro 0001  - Abertura do Bloco 0

        linha = "|0001"
        linha &= "|" & 0
        linha &= "|"

        ArquivoAux.Add(linha)

        Registro0001 += 1
        RegistroGeral += 1

    End Sub

    Private Sub CompoeRegistro0100(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro0100 As Integer, ByRef RegistroGeral As Integer)

        ' Registro 0100  - Dados do Contabilista

        For Each dr In ds.Tables(0).Rows
            With dr
                linha = "|0100"
                linha &= "|" & RTrim(.Item("ContadorNome"))

                If .Item("ContadorCPF").ToString.Length = 11 Then
                    linha &= "|" & .Item("ContadorCPF")
                Else
                    linha &= "|"
                End If

                linha &= "|" & .Item("ContadorCRC")

                If .Item("ContadorCPF").ToString.Length > 11 Then
                    linha &= "|" & .Item("ContadorCPF")
                Else
                    linha &= "|"
                End If

                linha &= "|" & .Item("ContadorCep")
                linha &= "|" & .Item("ContadorEndereco")
                linha &= "|" & .Item("ContadorNumero")
                linha &= "|" & .Item("ContadorComplemento")
                linha &= "|" & .Item("ContadorBairro")

                Dim largura As String = String.Empty
                largura = Funcoes.OnlyNumbers(.Item("ContadorTelefone")).Length

                linha &= "|" & Funcoes.OnlyNumbers(.Item("ContadorTelefone")).Substring(0, IIf(largura > 11, 11, largura))
                linha &= "|" & .Item("ContadorFax")
                linha &= "|" & .Item("ContadorEmail")
                linha &= "|" & .Item("ContadorEstadoIbge") & Funcoes.AlinharDireita(.Item("ContadorCodigoDoMunicipio"), 5, "0")
                linha &= "|"

                ArquivoAux.Add(linha)
                Registro0100 += 1
                RegistroGeral += 1
            End With
        Next

    End Sub

    Private Sub CompoeRegistro0110(ByRef ArquivoAux As ArrayList, ByRef Registro0110 As Integer, ByRef RegistroGeral As Integer)

        linha = "|0110"
        linha &= "|" & cboIncidencia.SelectedValue      'Codigo indicador do critério de escrituração e apuração adotado
        linha &= "|" & 2  'Codigo do metodo de apropriacao de creditos
        linha &= "|" & 1  'Codigo indicador do tipo de contribuicao apurada no periodo

        If cboIncidencia.SelectedValue = 2 Then
            linha &= "|2" '& cbo_Regime.value     'Codigo indicador do critério de escrituração e apuração adotado
        Else
            linha &= "|"
        End If

        If CDate(txtDataFinal.Text) > "2012/06/30" Then
            linha &= "|"
        End If

        ArquivoAux.Add(linha)
        Registro0110 += 1
        RegistroGeral += 1

    End Sub

    Private Function ConsultaRegistro0111() As DataSet
        Dim SqlT As String = String.Empty

        ' Registro 0111  - Tabela de receita bruta mensal para fins de rateio de cretitos comuns
        '    Where conta_Id in ('303030301', '303030302','303030303','303030304','303030305','403010104','403010106','403010107','403010108','402020303')" & vbCrLf & _


        SqlT = " DECLARE @SaldoRazao as Decimal DECLARE @SaldoRazaoNT as Decimal" & vbCrLf

        If EmpresaMestre = "05366261" OrElse EmpresaMestre = "38198213" OrElse EmpresaMestre = "40938762" OrElse EmpresaMestre = "62747840" OrElse EmpresaMestre = "62780383" OrElse EmpresaMestre = "63358210" Then
            SqlT &= " Select @SaldoRazao = 0" & vbCrLf
            SqlT &= " Select @SaldoRazaoNT = 0" & vbCrLf
        Else
            SqlT &= " Select @SaldoRazao = ISNULL(sum(Razao.CreditoOficial - Razao.DebitoOficial),0)" & vbCrLf &
                   "        From Razao INNER JOIN" & vbCrLf &
                   "                    PisCofinsXBlocos ON " & vbCrLf &
                   "                    Left(Razao.Empresa_Id, 8) = PisCofinsXBlocos.Empresa_Id" & vbCrLf &
                   "                    AND Razao.Conta_Id = PisCofinsXBlocos.Conta_Id" & vbCrLf &
                   "                    And PisCofinsXBlocos.Bloco_ID = '0111'" & vbCrLf &
                   "                    And PisCofinsXBlocos.Titulo_Id = 'TRIBUTADAS'" & vbCrLf &
                   "        Where   (Razao.Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "  	            AND Razao.lote_id not in (7500)" & vbCrLf &
                   "                AND (left(Razao.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   " Select @SaldoRazaoNT = ISNULL(sum(CreditoOficial - DebitoOficial),0)" & vbCrLf &
                   "		      from Razao" & vbCrLf &
                   "          Where conta_Id in ('402020401', '402020402')" & vbCrLf &
                   "		       AND (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "		   	   AND lote_id not in (7500)" & vbCrLf &
                   "            AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf
        End If

        SqlT &= " SELECT 'R0111' AS Registro," & vbCrLf &
                "        ISNULL(SUM(CASE" & vbCrLf &
                "			            WHEN oe.STPISCOFINS = 01" & vbCrLf &
                "				         THEN CASE WHEN isnull(nfe.BaseNova,0) > 0 then nfe.BaseNova else nfe.Base END" & vbCrLf &
                "				         ELSE 0 " & vbCrLf &
                "		              END), 0)  " & vbCrLf &
                "        + Case when @SaldoRazao > 0 then @SaldoRazao else 0 end" & vbCrLf &
                "                AS Rec_Bru_Ncum_Trib_Mi," & vbCrLf &
                "        ISNULL(SUM(CASE" & vbCrLf &
                "                    WHEN oe.STPISCOFINS IN (06, 09)" & vbCrLf &
                "			             THEN CASE WHEN isnull(nfe.BaseNova,0) > 0 then nfe.BaseNova else nfe.Base END" & vbCrLf &
                "                      ELSE 0" & vbCrLf &
                "			          END), 0) + Case when @SaldoRazaoNT > 0 then @SaldoRazaoNT else 0 end AS Rec_Bru_Ncum_Nt_Mi," & vbCrLf &
                "        ISNULL(SUM(CASE" & vbCrLf &
                "			           WHEN oe.STPISCOFINS = 08" & vbCrLf &
                "			  	         THEN CASE WHEN isnull(nfe.BaseNova,0) > 0 then nfe.BaseNova else nfe.Base END" & vbCrLf &
                "			  	         ELSE 0" & vbCrLf &
                "			           END), 0) AS Rec_Bru_Ncum_Exp," & vbCrLf &
                "        0 AS Rec_Bru_Cum," & vbCrLf &
                "        ISNULL(SUM(CASE" & vbCrLf &
                "			           WHEN oe.STPISCOFINS IN (01, 06, 08, 09)" & vbCrLf &
                "			        	 THEN CASE WHEN isnull(nfe.BaseNova,0) > 0 then nfe.BaseNova else nfe.Base END" & vbCrLf &
                "			        	 ELSE 0" & vbCrLf &
                "			          END), 0)" & vbCrLf &
                "		   + Case when @SaldoRazao > 0 then @SaldoRazao else 0 end + Case when @SaldoRazaoNT > 0 then @SaldoRazaoNT else 0 end" & vbCrLf &
                "		    AS Rec_Bru_Total " & vbCrLf &
                "   FROM NotasFiscaisXItens nfi" & vbCrLf &
                "  INNER JOIN NotasFiscais NF" & vbCrLf &
                "     ON nfi.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                "    AND nfi.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                "    AND nfi.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                "    AND nfi.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                "    AND nfi.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                "    AND nfi.Serie_Id        = NF.Serie_Id" & vbCrLf &
                "    AND nfi.Nota_Id         = NF.Nota_Id" & vbCrLf &
                "  INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
                "     ON nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
                "    AND nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
                "    AND nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
                "    AND nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
                "    AND nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
                "    AND nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
                "    AND nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
                "    AND nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
                "    AND nfi.CFOP_Id         = nfe.CFOP_Id" & vbCrLf &
                "    AND nfi.Sequencia_Id    = nfe.Sequencia_id" & vbCrLf &
                "  INNER JOIN SubOperacoes SO" & vbCrLf &
                "     ON SO.Operacao_Id     = nfi.Operacao" & vbCrLf &
                "    AND SO.SubOperacoes_Id = nfi.SubOperacao" & vbCrLf &
                "  inner join OperacaoXEstado OE" & vbCrLf &
                "     on OE.Codigo_id = nfi.OperacaoxEstado" & vbCrLf &
                "  WHERE (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') " & vbCrLf &
                "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                "    AND (ISNULL(NF.Situacao, 1) = 1)" & vbCrLf &
                "    AND ((NF.EntradaSaida_Id = 'S' and SO.Devolucao = 'N') or (NF.EntradaSaida_Id = 'E' and so.Devolucao = 'S')) " & vbCrLf &
                "    AND nfe.encargo_id = 'PRODUTO'" & vbCrLf &
                "    AND isnull(oe.ObsPISCOFINS, 0) <> 999" & vbCrLf

        Return Banco.ConsultaDataSet(SqlT, "Receitas")
    End Function

    Private Sub CompoeRegistro0111(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro0111 As Integer, ByRef RegistroGeral As Integer)

        If ds.Tables(0).Rows.Count < 1 Then
            linha = "|0111"
            linha &= "|0" '& .Item("Rec_Bru_Ncum_Trib_Mi")
            linha &= "|0" '& .Item("Rec_Bru_Ncum_Nt_Mi")
            linha &= "|0" '& .Item("Rec_Bru_Ncum_Exp")
            linha &= "|0" '& "0,00"
            linha &= "|0" '& .Item("Rec_Bru_Total")
            linha &= "|"
            ArquivoAux.Add(linha)
            Registro0111 += 1
            RegistroGeral += 1
        Else
            For Each dr In ds.Tables(0).Rows
                With dr
                    linha = "|0111"
                    linha &= "|" & .Item("Rec_Bru_Ncum_Trib_Mi")
                    linha &= "|" & .Item("Rec_Bru_Ncum_Nt_Mi")
                    linha &= "|" & .Item("Rec_Bru_Ncum_Exp")
                    linha &= "|" & "0,00"
                    linha &= "|" & .Item("Rec_Bru_Total")
                    linha &= "|"

                    ReceitaTotal = .Item("Rec_Bru_Total")
                    Trib_MI = .Item("Rec_Bru_Ncum_Trib_Mi")
                    Trib_ME = .Item("Rec_Bru_Ncum_Exp")
                    Trib_NT = .Item("Rec_Bru_Ncum_Nt_Mi")
                    ArquivoAux.Add(linha)
                    Registro0111 += 1
                    RegistroGeral += 1
                End With
            Next
        End If
    End Sub

    Private Function ConsultaRegistro0140() As DataSet
        Dim SqlT As String = String.Empty

        ' Registro 0140  - Tabela de cadastrato de estabelecimento

        SqlT = " SELECT Reduzido, Nome, Cliente_Id, Estado, Inscricao, CodigoDoMunicipio, InscricaoMunicipal,Suframa, Estadoibge From (" & vbCrLf &
                        " SELECT Cli.Reduzido, Cli.Nome, Cli.Cliente_Id, Cli.Estado, Cli.Inscricao," & vbCrLf &
                        "       Cli.CodigoDoMunicipio, CxE.InscricaoMunicipal," & vbCrLf &
                        "       Cli.Suframa , Mun.Estadoibge" & vbCrLf &
                        "  FROM Clientes Cli " & vbCrLf &
                        " INNER JOIN ClientesXEmpresas CxE" & vbCrLf &
                        "    ON Cli.Cliente_Id  = CxE.Empresa_Id" & vbCrLf &
                        "   AND Cli.Endereco_Id = CxE.EndEmpresa_Id" & vbCrLf &
                        " INNER JOIN Municipios Mun" & vbCrLf &
                        "    ON Cli.Estado            = Mun.Estado_id" & vbCrLf &
                        "   AND Cli.Cidade            = Mun.Municipio_id" & vbCrLf &
                        "   AND Cli.CodigoDoMunicipio = Mun.Codigo_id " & vbCrLf &
                        " where CxE.Matriz = 'S'" & vbCrLf &
                        "   AND (LEFT(Cli.Cliente_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                        " UNION " & vbCrLf &
                        " SELECT Cli.Reduzido, Cli.Nome, Cli.Cliente_Id, Cli.Estado, Cli.Inscricao," & vbCrLf &
                        "        Cli.CodigoDoMunicipio, CxE.InscricaoMunicipal," & vbCrLf &
                        "        Cli.Suframa , Mun.Estadoibge" & vbCrLf &
                        "   FROM Municipios Mun" & vbCrLf &
                        "  RIGHT OUTER JOIN Clientes Cli" & vbCrLf &
                        "     ON Mun.Estado_id    = Cli.Estado" & vbCrLf &
                        "    AND Mun.Municipio_id = Cli.Cidade" & vbCrLf &
                        "    AND Mun.Codigo_Id    = Cli.CodigoDoMunicipio" & vbCrLf &
                        "   LEFT OUTER JOIN NotasFiscais NF" & vbCrLf &
                        "     ON Cli.Cliente_Id  = NF.Empresa_Id" & vbCrLf &
                        "    AND Cli.Endereco_Id = NF.EndEmpresa_Id" & vbCrLf &
                        "   LEFT OUTER JOIN SubOperacoes SOP" & vbCrLf &
                        "     ON NF.Operacao    = SOP.Operacao_Id" & vbCrLf &
                        "    AND NF.SubOperacao = SOP.SubOperacoes_Id" & vbCrLf &
                        "   LEFT OUTER JOIN ClientesXEmpresas CxE" & vbCrLf &
                        "     ON Cli.Cliente_Id  = CxE.Empresa_Id" & vbCrLf &
                        "    AND Cli.Endereco_Id = CxE.EndEmpresa_Id" & vbCrLf &
                        "  WHERE NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'" & vbCrLf &
                        "    AND SUBSTRING(NF.Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
                        " UNION " & vbCrLf &
                        " SELECT     Cli.Reduzido, Cli.Nome, Cli.Cliente_Id, Cli.Estado, Cli.Inscricao, Cli.CodigoDoMunicipio, CxE.InscricaoMunicipal, Cli.Suframa, Mun.Estadoibge" & vbCrLf &
                        " FROM         Clientes AS Cli INNER JOIN" & vbCrLf &
                        " ClientesXEmpresas AS CxE ON Cli.Cliente_Id = CxE.Empresa_Id AND Cli.Endereco_Id = CxE.EndEmpresa_Id INNER JOIN" & vbCrLf &
                        " Municipios AS Mun ON Cli.Estado = Mun.Estado_id AND Cli.Cidade = Mun.Municipio_id AND Cli.CodigoDoMunicipio = Mun.Codigo_id INNER JOIN" & vbCrLf &
                        " Razao ON Cli.Cliente_Id = Razao.Empresa_Id AND Cli.Endereco_Id = Razao.EndEmpresa_Id" & vbCrLf &
                        " WHERE    (LEFT(Cli.Cliente_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                        " AND (LEFT(Razao.Conta_Id, 7) IN ('4020203', '4020204', '4030101')) AND (NOT (Razao.Conta_Id IN ('403010105', '403010103')))" & vbCrLf &
                        " AND Razao.Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'" & vbCrLf &
                        " AND (NOT (Razao.Lote_Id IN (7500))) " & vbCrLf &
                        " GROUP BY Cli.Reduzido, Cli.Nome, Cli.Cliente_Id, Cli.Estado, Cli.Inscricao, Cli.CodigoDoMunicipio, CxE.InscricaoMunicipal, Cli.Suframa, Mun.Estadoibge" & vbCrLf &
                        " ) as Consulta GROUP BY Reduzido, Nome, Cliente_Id, Estado, Inscricao, CodigoDoMunicipio, InscricaoMunicipal, Suframa, Estadoibge"


        Return Banco.ConsultaDataSet(SqlT, "Clientes")
    End Function

    Private Function ConsultaRegistro0145() As DataSet
        Dim SqlT As String = String.Empty

        'Registro 0145 - Regime de Apuração da contribuição
        '----------------------------------------------------------
        If (EmpresaMestre = "11111111") Then 'isolado temporariamente 03189063 utilizado somente no Curtume
            Sql = "  Select	Empresa_Id, " & vbCrLf &
                  "        SUM(VL_REC_TOT_EST) as VL_REC_TOT_EST," & vbCrLf &
                  " 		COD_ATIV_ECON," & vbCrLf &
                  " 		SUM(VL_REC_ATIV_ESTAB) as VL_REC_ATIV_ESTAB," & vbCrLf &
                  " 	    SUM(VL_REC_TOT_EST - VL_REC_ATIV_ESTAB) AS VL_EXC," & vbCrLf &
                  " 	    SUM(VL_REC_ATIV_ESTAB) AS VL_BC_CONT," & vbCrLf &
                  " 	    1 AS ALIQ_CONT," & vbCrLf &
                  " 	    SUM((VL_REC_ATIV_ESTAB * 1) / 100) AS	VL_CONT_APU," & vbCrLf &
                  "        '301030104' AS COD_CTA" & vbCrLf &
                  " From (" & vbCrLf &
                  " Select	Razao.Empresa_Id, SUM(Razao.CreditoOficial - Razao.DebitoOficial) as  VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                  "        0 as  VL_REC_ATIV_ESTAB" & vbCrLf &
                  " From razao " & vbCrLf &
                  " Where 1=1  And Conta_Id like '30101%' " & vbCrLf &
                  "    And (left(Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Produto in ('70101003',  '70101002'))" & vbCrLf &
                  "    And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                  "    AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                  "    And not Lote_Id in (7500) " & vbCrLf &
                  "    Group By Razao.Empresa_Id" & vbCrLf &
                  "    Union" & vbCrLf &
                  " SELECT  Razao.Empresa_Id, 	0 AS VL_REC_TOT_EST, '41040000' as COD_ATIV_ECON," & vbCrLf &
                  " 	     SUM(Razao.CreditoOficial - Razao.DebitoOficial) AS VL_REC_ATIV_ESTAB" & vbCrLf &
                  " FROM    Razao INNER JOIN" & vbCrLf &
                  "         Produtos ON Razao.Produto = Produtos.Produto_Id" & vbCrLf &
                  " Where	 1=1  And  (Razao.Conta_Id like '30101%' Or Razao.Conta_Id in ('301020101', '301020102'))" & vbCrLf &
                  "        And Not Razao.Conta_Id like '3010102%'" & vbCrLf &
                  " 		And (left(Razao.Produto, 5) in ('40102','20101', '10101','50101', '40103', '40104','70102') or Razao.Produto in ('70101003', '70101002'))" & vbCrLf &
                  " 		And (Razao.Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                  "         AND (left(Razao.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                  "        And Produtos.NCM Like '4104%'" & vbCrLf &
                  "        And not Lote_Id in (7500) " & vbCrLf &
                  " Group By Razao.Empresa_Id) as Consulta" & vbCrLf &
                  " Group By Empresa_Id, COD_ATIV_ECON " & vbCrLf
        End If

        Return Banco.ConsultaDataSet(SqlT, "Clientes")
    End Function

    Private Sub CompoeRegistro0145(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro0145 As Integer, ByRef RegistroGeral As Integer)

        'Registro 0145 - Regime de Apuração da contribuição
        For Each dr0145 In ds.Tables(0).Rows
            With dr0145
                linha = "|0145"
                linha &= "|" & 1
                linha &= "|" & .Item("Vl_Rec_Tot_Est")
                linha &= "|" & .Item("Vl_Rec_Ativ_Estab")
                linha &= "|" & 0
                linha &= "|"
                linha &= "|"
                ArquivoAux.Add(linha)
                Registro0145 += 1
                RegistroGeral += 1
            End With
        Next
    End Sub

    Private Function ConsultaRegistro0150() As DataSet
        Dim SqlT As String = String.Empty

        'Registro 0150 - Tabela de Cadastro do Participante
        SqlT = " SELECT DISTINCT NF.Empresa_Id, Cli.Cliente_Id, Cli.Endereco_Id, Cli.Regiao, Cli.Categoria, Cli.Estado, " & vbCrLf &
               "        Cli.Nome, Cli.Fantasia, Cli.Endereco, Cli.Bairro, Cli.Cep, Cli.Cidade, Cli.Inscricao, Cli.Telefone, " & vbCrLf &
               "        Cli.Fax, Cli.email, Cli.Reduzido, Cli.Numero, isnull(Cli.Complemento, '') as Complemento, Cli.Situacao, " & vbCrLf &
               "        Cli.CodigoDoMunicipio, ISNULL(Cli.Pais,'') as Pais_CLI, Cli.Habilitacao, " & vbCrLf &
               "        Cli.Suframa, Cli.EmailNFE, E.Estado_Id, E.Descricao, Mun.EstadoIbge as CodigoIBGE " & vbCrLf &
               "   FROM Clientes Cli " & vbCrLf &
               "  INNER JOIN NotasFiscais NF" & vbCrLf &
               "     ON Cli.Cliente_Id  = NF.Cliente_Id" & vbCrLf &
               "    AND Cli.Endereco_Id = NF.EndCliente_Id" & vbCrLf &
               "  INNER JOIN Estados E" & vbCrLf &
               "     ON Cli.Estado = E.Estado_Id" & vbCrLf &
               "   LEFT JOIN Municipios Mun" & vbCrLf &
               "     ON Cli.Estado            = Mun.Estado_id" & vbCrLf &
               "    AND Cli.CodigoDoMunicipio = Mun.Codigo_id  " & vbCrLf &
               "    AND Cli.Cidade            = Mun.Municipio_id " & vbCrLf &
               "  WHERE ISNULL(NF.Situacao,1) = 1 " & vbCrLf &
               "    AND NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "' " & vbCrLf &
               "    And LEN(Cli.cliente_Id) > 4" & vbCrLf &
               "    AND (LEFT(Cli.Cliente_Id, 10) <>  '0000000000')" & vbCrLf &
               "  UNION ALL " & vbCrLf &
               " SELECT R.Empresa_Id, Cli.Cliente_Id, Cli.Endereco_Id, Cli.Regiao, Cli.Categoria, Cli.Estado, Cli.Nome, Cli.Fantasia, Cli.Endereco, " & vbCrLf &
               "        Cli.Bairro, Cli.Cep, Cli.Cidade, Cli.Inscricao, Cli.Telefone, Cli.Fax, Cli.Email, Cli.Reduzido, Cli.Numero, " & vbCrLf &
               "        ISNULL(Cli.Complemento, N'') AS Complemento, Cli.Situacao, Cli.CodigoDoMunicipio, ISNULL(Cli.Pais, N'') AS Pais_CLI, Cli.Habilitacao, " & vbCrLf &
               "        Cli.Suframa, Cli.EmailNFE, E.Estado_Id, E.Descricao, Mun.Estadoibge AS CodigoIBGE" & vbCrLf &
               "   FROM Clientes Cli " & vbCrLf &
               "  INNER JOIN Estados E" & vbCrLf &
               "     ON Cli.Estado = E.Estado_Id " & vbCrLf &
               "  RIGHT OUTER JOIN Razao R " & vbCrLf &
               "     ON Cli.Cliente_Id  = R.Cliente_Id " & vbCrLf &
               "    AND Cli.Endereco_Id = R.EndCliente_Id " & vbCrLf &
               "   LEFT OUTER JOIN Municipios Mun " & vbCrLf &
               "     ON Cli.Estado            = Mun.Estado_id " & vbCrLf &
               "    AND Cli.CodigoDoMunicipio = Mun.Codigo_id " & vbCrLf &
               "    AND Cli.Cidade            = Mun.Municipio_id" & vbCrLf &
               "  WHERE LEFT(Conta_Id, 7) IN ('1010301', '2010101') " & vbCrLf &
               "    AND R.Lote_ID = 30 " & vbCrLf &
               "    AND R.Movimento_ID BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "' " & vbCrLf &
               "    AND R.DebitoOficial > 0" & vbCrLf &
               "    AND NOT R.Lote_Id in (7500) " & vbCrLf &
               " ORDER BY Nome, Estado, Cidade " & vbCrLf

        Return Banco.ConsultaDataSet(SqlT, "Clientes")
    End Function

    Private Sub CompoeRegistro0150(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro0150 As Integer, ByRef RegistroGeral As Integer, D100Empresa As String)

        'Registro 0150 - Tabela de Cadastro do Participante

        For Each dr1 In ds.Tables(0).Select("Empresa_Id =" & D100Empresa)
            With dr1
                linha = "|0150"
                linha &= "|" & LTrim(RTrim(.Item("Cliente_Id"))) & LTrim(RTrim(.Item("Endereco_Id")))
                linha &= "|" & RTrim(.Item("Nome"))
                linha &= "|" & .Item("Pais_CLI")
                If Microsoft.VisualBasic.Len(Trim(.Item("Cliente_Id"))) = 11 And .Item("Pais_CLI") = "1058" Then
                    linha &= "|"
                    linha &= "|" & Trim(.Item("Cliente_Id"))
                End If
                If Microsoft.VisualBasic.Len(Trim(.Item("Cliente_Id"))) = 14 And .Item("Pais_CLI") = "1058" Then
                    linha &= "|" & Trim(.Item("Cliente_Id"))
                    linha &= "|"
                End If
                If .Item("Pais_CLI") <> "1058" Then
                    linha &= "|"
                    linha &= "|"
                End If

                'linha &= "|" & .Item("Inscricao")
                linha &= "|" & IIf(IsNumeric(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", "").replace(" ", "")), .Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", "").replace(" ", ""), "")    'Inscricao Estadual

                If .Item("Pais_CLI") = "1058" Then
                    If IsDBNull(.Item("CodigoIBGE")) Then
                        linha &= "|00" & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                    Else
                        linha &= "|" & .Item("CodigoIBGE") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                    End If
                Else
                    linha &= "|"
                End If

                linha &= "|"                                'Suframa
                linha &= "|" & LTrim(RTrim(.Item("Endereco")))
                linha &= "|" & .Item("Numero")
                linha &= "|" & LTrim(RTrim(IIf(IsDBNull(.Item("Complemento")), "", .Item("Complemento"))))
                linha &= "|" & LTrim(RTrim(.Item("Bairro")))
                linha &= "|"
            End With

            ArquivoAux.Add(linha)
            Registro0150 += 1
            RegistroGeral += 1
        Next

    End Sub
#End Region

#Region "BLOCO M"
    Private Sub Bloco_M(ByRef pArray As ArrayList)
        Dim SqlArray As New ArrayList
        Dim ds As DataSet
        Dim dsaux As DataSet

        '----M100 ----- 101 Tributado_MI ----------------------

        Percentual = (Trib_MI * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M100|101|0|100,00|1,65|||1,65|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM100 += 1
        RegistroGeral += 1


        '' Ajuste Credito Pis - Devoluções
        ''-------Registro  M110 ---------------------

        Sql = "SELECT Case" & vbCrLf &
              "         when NotasFiscais.EntradaSaida_Id = 'S'" & vbCrLf &
              "           then 0 else 1" & vbCrLf &
              "       end as ES," & vbCrLf &
              "       NotasFiscais.Nota_Id," & vbCrLf &
              "       NotasFiscais.Movimento," & vbCrLf &
              "       NotasFiscaisXEncargos.Percentual," & vbCrLf &
              "       Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              "  FROM NotasFiscais" & vbCrLf &
              " INNER JOIN SubOperacoes" & vbCrLf &
              "    ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
              "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
              " INNER JOIN NotasFiscaisXItens" & vbCrLf &
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
              " INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
              "    ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
              "   AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
              "   And NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
              "   And NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
              "   And NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              " inner join OperacaoXEstado OE" & vbCrLf &
              "    on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "   And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "   And OE.STPisCofins in (49,  98, 99)" & vbCrLf &
              "   AND (NotasFiscaisXEncargos.Encargo_Id in ('Pis', 'PIS CRED PRE'))" & vbCrLf &
              "   And  NotasFiscaisXEncargos.Percentual > 1 And NotasFiscais.Situacao = 1 " & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual" & vbCrLf

        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux
                linha = "|" & "M110" '(01)
                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM110 = RegistroM110 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next


        '----M100 ----- 106 Presumido_MI ----------------------

        Percentual = (Trib_MI * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M100|106|0|100,00|0,4455|||0|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM100 += 1
        RegistroGeral += 1


        '' Ajuste Credito Pis - Devoluções
        ''-------Registro  M110 ---------------------

        Sql = "SELECT Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 0 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              "  FROM NotasFiscais" & vbCrLf &
              " INNER JOIN SubOperacoes" & vbCrLf &
              "    ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
              "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
              " INNER JOIN NotasFiscaisXItens" & vbCrLf &
              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
              " INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
              "    ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
              "   AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id " & vbCrLf &
              "   And NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id " & vbCrLf &
              "   And NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id " & vbCrLf &
              "   And NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
              "   And NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
              "   And NotasFiscaisXItens.Sequencia_Id     = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              " inner join OperacaoXEstado OE" & vbCrLf &
              "    on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "   And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "   And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "   And OE.STPisCofins in (49,   98, 99)" & vbCrLf &
              "   AND (NotasFiscaisXEncargos.Encargo_Id in ('Pis', 'PIS CRED PRE'))" & vbCrLf &
              "   And  NotasFiscaisXEncargos.Percentual < 1" & vbCrLf &
              "   And NotasFiscais.Situacao = 1 " & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual" & vbCrLf

        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux
                linha = "|" & "M110" '(01)
                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM110 = RegistroM110 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next

        ''--------Registro M100 - 201 Receitas NT -----------------------------

        Percentual = (Trib_NT * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M100|201|0|100,00|1,65|||1,65|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM100 += 1
        RegistroGeral += 1

        Sql = " SELECT Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 1 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              "   FROM NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49, 98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Pis', 'PIS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual > 1  And NotasFiscais.Situacao = 1 " & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual" & vbCrLf


        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux

                linha = "|" & "M110" '(01)

                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM110 = RegistroM110 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next

        ''--------Registro M100 - 206 Receitas NT -----------------------------

        Percentual = (Trib_NT * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M100|206|0|100,00|0,4455|||0|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM100 += 1
        RegistroGeral += 1


        Sql = " SELECT   Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 0 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49,   98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Pis', 'PIS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual < 1  And NotasFiscais.Situacao = 1 " & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual" & vbCrLf

        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux

                linha = "|" & "M110" '(01)

                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM110 = RegistroM110 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next



        ''--------Registro M100 - 301 Receitas ME -----------------------------

        Percentual = (Trib_ME * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M100|301|0|100,00|1,65|||1,65|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM100 += 1
        RegistroGeral += 1


        Sql = " SELECT  Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 1 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49,  98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Pis', 'PIS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual > 1  And NotasFiscais.Situacao = 1 " & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual" & vbCrLf

        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux

                linha = "|" & "M110" '(01)

                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM110 = RegistroM110 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next

        ''--------Registro M100 - 306 Receitas ME -----------------------------

        Percentual = (Trib_ME * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M100|306|0|100,00|0,4455|||0|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM100 += 1
        RegistroGeral += 1


        Sql = " SELECT  Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 0 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49, 98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Pis', 'PIS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual < 1  And NotasFiscais.Situacao = 1 " & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscaisXEncargos.Percentual" & vbCrLf

        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux

                linha = "|" & "M110" '(01)

                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM110 = RegistroM110 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next

        ''-------Registro M400 ---------------------
        Sql = " SELECT 'M400' AS Reg, STPISCOFINS, sum(Vl_Tot_Item) as  Vl_Tot_Item" & vbCrLf &
              "   From (" & vbCrLf &
              " SELECT 'M400' AS Reg," & vbCrLf &
              "        OE.STPISCOFINS," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PRODUTO','FRETES','SEGURO') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Tot_Item" & vbCrLf &
              " FROM         NotasFiscais INNER JOIN" & vbCrLf &
              "                  NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
              "                  NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
              "                  NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
              "                  NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "                  NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND" & vbCrLf &
              "                  NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND" & vbCrLf &
              "                  NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND" & vbCrLf &
              "                  NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND" & vbCrLf &
              "                  NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And " & vbCrLf &
              "                  NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id AND NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id INNER JOIN" & vbCrLf &
              "                  SubOperacoes ON NotasFiscaisXitens.Operacao = SubOperacoes.Operacao_Id AND" & vbCrLf &
              "                  NotasFiscaisXitens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
              "                 inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE                (NotasFiscais.EntradaSaida_Id = 'S') AND (ISNULL(NotasFiscais.Situacao, 1) = 1) AND (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "') AND" & vbCrLf &
              "                      (NotasFiscaisXEncargos.Encargo_Id IN ('PIS','FRETES','SEGURO','PRODUTO')) And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "' AND " & vbCrLf &
              "                      (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360) AND" & vbCrLf &
              "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360) AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932)) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360) AND" & vbCrLf &
              "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5150 AND 5160) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6150 AND 6160) AND" & vbCrLf &
              "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253) AND" & vbCrLf &
              "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303) " & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf &
              "                      AND OE.STPISCOFINS <> 0 " & vbCrLf &
              " GROUP     BY  OE.STPISCOFINS" & vbCrLf &
              " Having    SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Valor END, 0)) = 0" & vbCrLf

        '(SubOperacoes.Devolucao <> 'S') And  tirei a devolução do bloco acima - Furlan - 04-11-2020

        If (EmpresaMestre = "04376053") Or (EmpresaMestre = "04440724") Or (EmpresaMestre = "05954217") Or (EmpresaMestre = "06329316") Or (EmpresaMestre = "07090163") Or (EmpresaMestre = "07093052") And cboIncidencia.SelectedValue <> 2 Then
            Sql &= " Union " & vbCrLf &
                    " select   'M400' AS Reg, 6 as STPISCOFINS,  Sum(Vl_Tot_Item) as Vl_Tot_Item From (" & vbCrLf &
                    " Select 'M400' AS Reg, 6 as STPISCOFINS,  Empresa_Id, Sum(CreditoOficial - DebitoOficial) as Vl_Tot_Item, LEFT(Conta_Id, 7) AS Conta_Id" & vbCrLf &
                    " From Razao" & vbCrLf &
                    " Where (Left(Conta_Id, 7) in ('4020203', '4020204')) And not Conta_Id in ('403010105', '403010103')" & vbCrLf &
                    " And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                    " And (Left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                    " And not Lote_Id in (7500) " & vbCrLf &
                    " Group by Empresa_Id, LEFT(Conta_Id, 7) Having Sum(CreditoOficial - DebitoOficial) > 0 " & vbCrLf &
                    " ) as Consul" & vbCrLf &
                    " ) as Consulta Group By STPISCOFINS" & vbCrLf
        Else
            Sql &= "  ) as Consulta Group By STPISCOFINS"
        End If

        Dim CstPis As Integer = 0

        ds = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each dr In ds.Tables(0).Rows
            With dr

                linha = "|" & "M400"
                linha &= "|" & Format(.item("STPISCOFINS"), "00")
                linha &= "|" & .item("Vl_Tot_Item")
                linha &= "|" '& ""
                linha &= "|" '& ""
                linha &= "|"

                CstPis = .item("STPISCOFINS")

                pArray.Add(linha)
                RegistroM400 += 1
                RegistroGeral += 1


                ''---- M410 Detalhamento das Receitas Isentas

                Sql = " SELECT  'M410' AS Reg," & vbCrLf &
                      "     Case when oe.ObsPISCOFINS = 0 then 999 else oe.ObsPISCOFINS End  as ObsPISCOFINS," & vbCrLf &
                      "         SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PRODUTO','FRETES','SEGURO') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Tot_Item" & vbCrLf &
                      "  FROM  NotasFiscais INNER JOIN" & vbCrLf &
                      "     NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                      "     NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                      "     NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                      "     NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                      "     NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND" & vbCrLf &
                      "     NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND" & vbCrLf &
                      "     NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND" & vbCrLf &
                      "     NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND" & vbCrLf &
                      "     NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf &
                      "     NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id AND NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id INNER JOIN" & vbCrLf &
                      "     SubOperacoes ON NotasFiscaisXitens.Operacao = SubOperacoes.Operacao_Id AND" & vbCrLf &
                      "     NotasFiscaisXitens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                      "     inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                      " WHERE                (NotasFiscais.EntradaSaida_Id = 'S') AND (ISNULL(NotasFiscais.Situacao, 1) = 1) AND (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "') AND" & vbCrLf &
                      "                      (NotasFiscaisXEncargos.Encargo_Id IN ('PIS','FRETES','SEGURO','PRODUTO')) And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "' AND " & vbCrLf &
                      "                      (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360) AND" & vbCrLf &
                      "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360) AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932)) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360) AND" & vbCrLf &
                      "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5150 AND 5160) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6150 AND 6160) AND" & vbCrLf &
                      "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253) AND" & vbCrLf &
                      "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303) " & vbCrLf &
                      "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
                      "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
                      "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
                      "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
                      "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
                      "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
                      "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf &
                      "                      AND OE.STPisCofins = " & .item("STPisCofins") & vbCrLf &
                      " GROUP     BY  Case when oe.ObsPISCOFINS = 0 then 999 else oe.ObsPISCOFINS End" & vbCrLf &
                      " Having    SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Valor END, 0)) = 0" & vbCrLf

                '(SubOperacoes.Devolucao <> 'S') And  tirei a devolução do bloco acima - Furlan - 04-11-2020

                dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
                For Each draux In dsaux.Tables(0).Rows
                    With draux

                        linha = "|" & "M410"
                        linha &= "|" & Format(.item("ObsPISCOFINS"), "000")
                        linha &= "|" & .item("Vl_Tot_Item")
                        linha &= "|" '& ""
                        linha &= "|" '& ""
                        linha &= "|"

                        pArray.Add(linha)
                        RegistroM410 += 1
                        RegistroGeral += 1

                    End With
                Next

                'dsaux.Clear()

                If (EmpresaMestre = "04376053") Or (EmpresaMestre = "04440724") Or (EmpresaMestre = "05954217") Or (EmpresaMestre = "06329316") Or (EmpresaMestre = "07090163") Or (EmpresaMestre = "07093052") Then
                    If (CstPis = 6 And cboIncidencia.SelectedValue <> 2) Then
                        Sql = " select   'M410' AS Reg, 911 as ObsPISCOFINS,  Sum(Vl_Tot_Item) as Vl_Tot_Item From (" & vbCrLf &
                                   " Select 'M410' AS Reg, 911 as ObsPISCOFINS,  Empresa_Id, Sum(CreditoOficial - DebitoOficial) as Vl_Tot_Item, Left(Conta_Id, 7) as Conta_Id" & vbCrLf &
                                   " From Razao" & vbCrLf &
                                   " Where  (Left(Conta_Id, 7) in ('4020203', '4020204' ))  And not Conta_Id in ('403010105', '403010103')" & vbCrLf &
                                   " And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                   " And (Left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                                   " And not Lote_Id in (7500) " & vbCrLf &
                                   " Group by Empresa_Id, Left(Conta_Id, 7) Having Sum(CreditoOficial - DebitoOficial) > 0" & vbCrLf &
                                   " ) as Consul "

                        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
                        For Each draux In dsaux.Tables(0).Rows
                            With draux

                                linha = "|" & "M410"
                                linha &= "|" & Format(.item("ObsPISCOFINS"), "000")
                                linha &= "|" & .item("Vl_Tot_Item")
                                linha &= "|" '& ""
                                linha &= "|" '& ""
                                linha &= "|"

                                pArray.Add(linha)
                                RegistroM410 += 1
                                RegistroGeral += 1

                            End With
                        Next
                    End If
                End If
            End With
        Next

        '----M500 ----- 101 Tributado_MI ----------------------

        Percentual = (Trib_MI * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M500|101|0|100,00|7,6|||7,6|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM500 += 1
        RegistroGeral += 1


        Sql = " SELECT  Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 1 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49,  98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Cofins', 'COFINS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual > 5 AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento" & vbCrLf

        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux
                linha = "|" & "M510" '(01)
                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM510 = RegistroM510 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next

        '----M500 ----- 106 Tributado_MI ----------------------

        Percentual = (Trib_MI * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M500|106|0|100,00|2,0520|||0|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM500 += 1
        RegistroGeral += 1


        Sql = " SELECT  Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 0 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49,  98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Cofins', 'COFINS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual < 5 AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento" & vbCrLf

        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux
                linha = "|" & "M510" '(01)
                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM510 = RegistroM510 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next


        ''--------Registro M500 - 201 Receitas NT -----------------------------

        Percentual = (Trib_NT * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M500|201|0|100,00|7,6|||7,6|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM500 += 1
        RegistroGeral += 1


        Sql = " SELECT  Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 1 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49,  98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Cofins', 'COFINS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual > 5 AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento" & vbCrLf


        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux
                linha = "|" & "M510" '(01)
                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM510 = RegistroM510 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next


        ''--------Registro M500 - 206 Receitas NT -----------------------------

        Percentual = (Trib_NT * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M500|206|0|100,00|2,0520|||0|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM500 += 1
        RegistroGeral += 1


        Sql = " SELECT  Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 0 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49, 98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Cofins', 'COFINS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual < 5 AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento" & vbCrLf


        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux
                linha = "|" & "M510" '(01)
                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM510 = RegistroM510 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next


        ''--------Registro M500 - 301 Receitas ME -----------------------------

        Percentual = (Trib_ME * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M500|301|0|100,00|7,6|||7,6|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM500 += 1
        RegistroGeral += 1


        Sql = " SELECT  Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 1 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49, 98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Cofins', 'COFINS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual > 5 AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento" & vbCrLf


        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux
                linha = "|" & "M510" '(01)
                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM510 = RegistroM510 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next


        ''--------Registro M500 - 306 Receitas ME -----------------------------

        Percentual = (Trib_ME * 100) / ReceitaTotal
        If Not ReceitaTotal > 0 Then Percentual = 0

        linha = "|M500|306|0|100,00|2,0520|||0|0|0|0|0|1|0|0|"
        pArray.Add(linha)
        RegistroM500 += 1
        RegistroGeral += 1


        Sql = " SELECT  Case when NotasFiscais.EntradaSaida_Id = 'S' then 0 else 0 end as ES, NotasFiscais.Nota_Id, NotasFiscais.Movimento, Sum(NotasFiscaisXEncargos.Valor) as Valor" & vbCrLf &
              " FROM   NotasFiscais INNER JOIN" & vbCrLf &
              "        SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND " & vbCrLf &
              "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf &
              "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "        NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf &
              "        NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id And NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id And" & vbCrLf &
              "        NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id And NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id And" & vbCrLf &
              "        NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id And NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And" & vbCrLf &
              "        NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id And NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf &
              "        inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE  (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "        And (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 1551 AND 1551" & vbCrLf &
              "        And NotasFiscaisXItens.cfop_Id Not BETWEEN 2551 AND 2551" & vbCrLf &
              "        And OE.STPisCofins in (49, 98, 99) AND (NotasFiscaisXEncargos.Encargo_Id in ('Cofins', 'COFINS CRED PRE'))" & vbCrLf &
              "        And  NotasFiscaisXEncargos.Percentual < 5 AND (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              " Group  By NotasFiscais.EntradaSaida_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento" & vbCrLf

        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each draux In dsaux.Tables(0).Rows
            With draux
                linha = "|" & "M510" '(01)
                linha = linha & "|" & .item("ES") '(0 - Ajuste de Redução, 1 - Ajuste de Acrécimo
                linha = linha & "|" & FormatNumber((draux("Valor") * Percentual) / 100, 2, TriState.UseDefault, TriState.UseDefault, TriState.False)  '(03) Valor estorno
                linha = linha & "|" & "06" '(04) Codigo ajuste de estorno
                linha = linha & "|" & .item("Nota_Id") '(05) Numero documento
                linha = linha & "|" & IIf(.item("ES") = 0, "Devolucao de Compras", "Devolucao de Vendas") '(06) descricao
                linha = linha & "|" & CDate(.Item("Movimento")).ToStrDate().Replace("-", "")
                linha = linha & "|"

                pArray.Add(linha)
                RegistroM510 = RegistroM510 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next


        ''---- M800 Detalhamento das Receitas Isentas

        Sql = " SELECT  'M800' AS Reg, STPISCOFINS, sum(Vl_Tot_Item) as  Vl_Tot_Item From (" & vbCrLf &
              " SELECT  'M800' AS Reg," & vbCrLf &
              "     OE.STPISCOFINS," & vbCrLf &
              "         SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PRODUTO','FRETES','SEGURO') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Tot_Item" & vbCrLf &
              " FROM         NotasFiscais INNER JOIN" & vbCrLf &
              "                  NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
              "                  NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
              "                  NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
              "                  NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
              "                  NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND" & vbCrLf &
              "                  NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND" & vbCrLf &
              "                  NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND" & vbCrLf &
              "                  NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND" & vbCrLf &
              "                  NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id And " & vbCrLf &
              "                  NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id AND NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id INNER JOIN" & vbCrLf &
              "                  SubOperacoes ON NotasFiscaisXitens.Operacao = SubOperacoes.Operacao_Id AND" & vbCrLf &
              "                  NotasFiscaisXitens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
              "                  inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
              " WHERE                (NotasFiscais.EntradaSaida_Id = 'S') AND (ISNULL(NotasFiscais.Situacao, 1) = 1) " & vbCrLf &
              "                   AND (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "') AND" & vbCrLf &
              "                      (NotasFiscaisXEncargos.Encargo_Id IN ('PIS','FRETES','SEGURO','PRODUTO')) And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "' AND " & vbCrLf &
              "                      (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360) AND" & vbCrLf &
              "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360) AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932)) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360) AND" & vbCrLf &
              "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5150 AND 5160) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6150 AND 6160) AND" & vbCrLf &
              "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253) AND" & vbCrLf &
              "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303) " & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
              "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf &
              "                      AND OE.STPISCOFINS <> 0 " & vbCrLf &
              " GROUP     BY  OE.STPISCOFINS" & vbCrLf &
              " Having    SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Valor END, 0)) = 0" & vbCrLf

        '(SubOperacoes.Devolucao <> 'S') And  tirei a devolução do bloco acima - Furlan - 04-11-2020

        If (EmpresaMestre = "04376053") Or (EmpresaMestre = "04440724") Or (EmpresaMestre = "05954217") Or (EmpresaMestre = "06329316") Or (EmpresaMestre = "07090163") Or (EmpresaMestre = "07093052") And cboIncidencia.SelectedValue <> 2 Then
            Sql &= " Union " & vbCrLf &
                    " select   'M800' AS Reg, 6 as STPISCOFINS,  Sum(Vl_Tot_Item) as Vl_Tot_Item From (" & vbCrLf &
                    " Select 'M800' AS Reg, 6 as STPISCOFINS,  Empresa_Id, Sum(CreditoOficial - DebitoOficial) as Vl_Tot_Item, LEFT(Conta_Id, 7) AS Conta_Id" & vbCrLf &
                    " From Razao" & vbCrLf &
                    " Where (Left(Conta_Id, 7) in ('4020203', '4020204' ))  And not Conta_Id in ('403010105', '403010103')" & vbCrLf &
                    " And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                    " And (Left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                    " And not Lote_Id in (7500) " & vbCrLf &
                    " Group by Empresa_Id, LEFT(Conta_Id, 7) Having Sum(CreditoOficial - DebitoOficial) > 0 " & vbCrLf &
                    " ) as Consul" & vbCrLf &
                    " ) as Consulta Group By STPISCOFINS"
        Else
            Sql &= "  ) as Consulta Group By STPISCOFINS"
        End If

        Dim GCst As Integer = 0

        ds = Banco.ConsultaDataSet(Sql, "Clientes")
        For Each dr In ds.Tables(0).Rows
            With dr

                linha = "|" & "M800"
                linha &= "|" & Format(.item("STPISCOFINS"), "00")
                linha &= "|" & .item("Vl_Tot_Item")
                linha &= "|" '& ""
                linha &= "|" '& ""
                linha &= "|"

                GCst = .item("STPISCOFINS")

                pArray.Add(linha)
                RegistroM800 += 1
                RegistroGeral += 1

                ''---- M810 Detalhamento das Receitas Isentas

                Sql = " SELECT  'M810' AS Reg," & vbCrLf &
                      "     Case when OE.ObsPISCOFINS = 0 then 999 else OE.ObsPISCOFINS End  as ObsPISCOFINS," & vbCrLf &
                      "         SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PRODUTO','FRETES','SEGURO') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Tot_Item" & vbCrLf

                Sql &= "  FROM  NotasFiscais INNER JOIN" & vbCrLf &
                       "     NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                       "     NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                       "     NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                       "     NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                       "     NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND" & vbCrLf &
                       "     NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND" & vbCrLf &
                       "     NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND" & vbCrLf &
                       "     NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND" & vbCrLf &
                       "     NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf &
                       "     NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id AND NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id INNER JOIN" & vbCrLf &
                       "     SubOperacoes ON NotasFiscaisXitens.Operacao = SubOperacoes.Operacao_Id AND" & vbCrLf &
                       "     NotasFiscaisXitens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                       "     inner join OperacaoXEstado OE on OE.Codigo_id = NotasFiscaisXItens.OperacaoxEstado" & vbCrLf &
                       " WHERE                (NotasFiscais.EntradaSaida_Id = 'S') AND (ISNULL(NotasFiscais.Situacao, 1) = 1) " & vbCrLf &
                       "                      AND (Left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "') AND" & vbCrLf &
                       "                      (NotasFiscaisXEncargos.Encargo_Id IN ('PIS','FRETES','SEGURO','PRODUTO')) And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "' AND " & vbCrLf &
                       "                      (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1350 AND 1360) AND" & vbCrLf &
                       "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2350 AND 2360) AND (NotasFiscaisXItens.Cfop_Id NOT IN (1932,2932)) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5350 AND 5360) AND" & vbCrLf &
                       "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5150 AND 5160) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6150 AND 6160) AND" & vbCrLf &
                       "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 6350 AND 6360) AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1252 AND 1253) AND" & vbCrLf &
                       "                      (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1302 AND 1303) " & vbCrLf &
                       "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5901 AND 5921)" & vbCrLf &
                       "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6901 and 6921)" & vbCrLf &
                       "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5923 AND 5926)" & vbCrLf &
                       "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6923 and 6926)" & vbCrLf &
                       "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5929 AND 5949)" & vbCrLf &
                       "                      AND (NotasFiscaisXItens.Cfop_Id Not Between 6929 and 6949)" & vbCrLf &
                       "                      AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 5555 AND 5556)" & vbCrLf &
                       "                      AND OE.STPISCOFINS = " & .item("STPISCOFINS") & vbCrLf

                '(SubOperacoes.Devolucao <> 'S') And  tirei a devolução do bloco acima - Furlan - 04-11-2020

                Sql &= " GROUP     BY Case when OE.ObsPISCOFINS = 0 then 999 else OE.ObsPISCOFINS End" & vbCrLf &
                       " Having    SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Valor END, 0)) = 0" & vbCrLf

                dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
                For Each draux In dsaux.Tables(0).Rows
                    With draux

                        linha = "|" & "M810"
                        linha &= "|" & Format(.item("ObsPISCOFINS"), "000")
                        linha &= "|" & .item("Vl_Tot_Item")
                        linha &= "|" '& ""
                        linha &= "|" '& ""
                        linha &= "|"

                        pArray.Add(linha)
                        RegistroM810 += 1
                        RegistroGeral += 1

                    End With
                Next

                If (EmpresaMestre = "04376053") Or (EmpresaMestre = "04440724") Or (EmpresaMestre = "05954217") Or (EmpresaMestre = "06329316") Or (EmpresaMestre = "07090163") Or (EmpresaMestre = "07093052") Then
                    If (GCst = 6 And cboIncidencia.SelectedValue <> 2) Then
                        Sql = " select   'M810' AS Reg, 911 as ObsPISCOFINS,  Sum(Vl_Tot_Item) as Vl_Tot_Item From (" & vbCrLf &
                              " Select 'M810' AS Reg, 911 as ObsPISCOFINS,  Empresa_Id, Sum(CreditoOficial - DebitoOficial) as Vl_Tot_Item, Left(Conta_Id, 7) as Conta_Id" & vbCrLf &
                              " From Razao" & vbCrLf &
                              " Where  (Left(Conta_Id, 7) in ('4020203', '4020204'))  And not Conta_Id in ('403010105', '403010103')" & vbCrLf &
                              " And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                              " And (Left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                              " And not Lote_Id in (7500) " & vbCrLf &
                              " Group by Empresa_Id, Left(Conta_Id, 7) Having Sum(CreditoOficial - DebitoOficial) > 0" & vbCrLf &
                              " ) as Consul "

                        dsaux = Banco.ConsultaDataSet(Sql, "Clientes")
                        For Each draux In dsaux.Tables(0).Rows
                            With draux

                                linha = "|" & "M810"
                                linha &= "|" & Format(.item("ObsPISCOFINS"), "000")
                                linha &= "|" & .item("Vl_Tot_Item")
                                linha &= "|" '& ""
                                linha &= "|" '& ""
                                linha &= "|"

                                pArray.Add(linha)
                                RegistroM810 += 1
                                RegistroGeral += 1
                            End With
                        Next
                    End If
                End If
            End With
        Next
    End Sub
#End Region

#Region "BLOCO D"

    Private Function ConsultaD001() As DataSet
        Dim Sql As String = " SELECT TOP 1 1" & vbCrLf
        Sql &= GetSqlFROMRegistroD("")
        Return Banco.ConsultaDataSet(Sql, "D001")
    End Function

    Private Function Consulta_D010_D100_D101_D105_D200_D201_D205_D500_D501_D505() As DataSet
        Dim SqlT As String = String.Empty
        '******************************************************************************************************************************************************************************
        '*********************************************************************** D010 *************************************************************************************************
        '******************************************************************************************************************************************************************************
        SqlT = GetSqlRegistroD010()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** D100 *************************************************************************************************
        '******************************************************************************************************************************************************************************
        SqlT &= GetSqlRegistroD100()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** D101 *************************************************************************************************
        '******************************************************************************************************************************************************************************
        SqlT &= GetSqlRegistroD101()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** D105 *************************************************************************************************
        '******************************************************************************************************************************************************************************
        SqlT &= GetSqlRegistroD105()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** D200 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroD200()


        '******************************************************************************************************************************************************************************
        '*********************************************************************** D201 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroD201()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** D205 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroD205()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** D500 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroD500()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** D501 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroD501()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** D505 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroD505()

        '******************************************************************************************************************************************************************************
        '************************************************************************* RESULTADO ******************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= "Select * from #D010;" & vbCrLf &
               "Select * from #D100;" & vbCrLf &
               "Select * from #D101;" & vbCrLf &
               "Select * from #D105;" & vbCrLf &
               "Select * from #D200;" & vbCrLf &
               "Select * from #D201;" & vbCrLf &
               "Select * from #D205;" & vbCrLf &
               "Select * from #D500;" & vbCrLf &
               "Select * from #D501;" & vbCrLf &
               "Select * from #D505;" & vbCrLf

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(SqlT, "Consulta")
        ds.Tables(0).TableName = "D010"
        ds.Tables(1).TableName = "D100"
        ds.Tables(2).TableName = "D101"
        ds.Tables(3).TableName = "D105"
        ds.Tables(4).TableName = "D200"
        ds.Tables(5).TableName = "D201"
        ds.Tables(6).TableName = "D205"
        ds.Tables(7).TableName = "D500"
        ds.Tables(8).TableName = "D501"
        ds.Tables(9).TableName = "D505"

        Return ds
    End Function

    Private Function GetSqlRegistroD010() As String
        Sql = "SELECT *  " & vbCrLf &
              "  INTO #D010 " & vbCrLf &
              "  FROM ( " & vbCrLf &
              "        SELECT Cli.Reduzido, Cli.Nome, Cli.Cliente_Id, Cli.Estado, Cli.Inscricao, Cli.CodigoDoMunicipio, CxE.InscricaoMunicipal,Cli.Suframa" & vbCrLf &
              "          FROM Clientes Cli" & vbCrLf &
              "         INNER JOIN ClientesXEmpresas CxE " & vbCrLf &
              "            ON Cli.Cliente_Id  = CxE.Empresa_Id" & vbCrLf &
              "           AND Cli.Endereco_Id = CxE.EndEmpresa_Id" & vbCrLf &
              "         WHERE CxE.Matriz = 'True'" & vbCrLf &
              "         UNION " & vbCrLf &
              "        SELECT Cli.Reduzido, Cli.Nome, Cli.Cliente_Id, Cli.Estado, Cli.Inscricao, Cli.CodigoDoMunicipio, CxE.InscricaoMunicipal,Cli.Suframa" & vbCrLf &
              "          FROM NotasFiscais NF" & vbCrLf &
              "         INNER JOIN SubOperacoes Sop" & vbCrLf &
              "            ON NF.Operacao    = Sop.Operacao_Id" & vbCrLf &
              "           AND NF.SubOperacao = Sop.SubOperacoes_Id " & vbCrLf &
              "         INNER JOIN Clientes Cli" & vbCrLf &
              "            ON NF.Empresa_Id    = Cli.Cliente_Id" & vbCrLf &
              "           AND NF.EndEmpresa_Id = Cli.Endereco_Id" & vbCrLf &
              "         INNER JOIN ClientesXEmpresas CxE" & vbCrLf &
              "            ON Cli.Cliente_Id  = CxE.Empresa_Id" & vbCrLf &
              "           AND Cli.Endereco_Id = CxE.EndEmpresa_Id" & vbCrLf &
              "         WHERE NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "'  " & vbCrLf &
              "           AND LEFT(NF.Empresa_Id,8) = '" & EmpresaMestre & "'" & vbCrLf &
              "           AND (Sop.Classe IN ('" & eClassesOperacoes.VENDAS.ToString & "', '" & eClassesOperacoes.EXPORTACOES.ToString & "', '" & eClassesOperacoes.FRETES.ToString & "', '" & eClassesOperacoes.COMPRAS.ToString & "'))" & vbCrLf &
              "    ) AS X " & vbCrLf
        Return Sql
    End Function

    Private Function GetSqlRegistroD100() As String
        Sql = " SELECT 'D100' AS Reg," & vbCrLf &
              "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.Nota_Id, Cli.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, NF.Serie_Id,  " & vbCrLf &
              "        '0' AS Ind_Oper," & vbCrLf &
              "		   case " & vbCrLf &
              "		       when ISNULL(DocXML.Tipo,'') = 'Mic'" & vbCrLf &
              "		           then 0" & vbCrLf &
              "		           else case " & vbCrLf &
              "		                   when NF.NossaEmissao = 'S' " & vbCrLf &
              "		                   then 0 " & vbCrLf &
              "		                   else 1 " & vbCrLf &
              "		                end " & vbCrLf &
              "		       end AS Ind_Emit," & vbCrLf &
              "        Cli.Cliente_Id AS Cod_Part," & vbCrLf &
              "        Cli.Endereco_Id, NF.TipoDeDocumento AS Cod_Mod," & vbCrLf &
              "        0 AS Cod_Sit," & vbCrLf &
              "        NF.Serie_Id AS Ser, '' AS Sub," & vbCrLf &
              "        NF.Nota_Id AS Num_Doc," & vbCrLf &
              "		   case" & vbCrLf &
              "			   when NF.CifFob = 'CIF'" & vbCrLf &
              "				   then '0'" & vbCrLf &
              "				   else" & vbCrLf &
              "					   case" & vbCrLf &
              "						   when NF.CifFob = 'FOB'" & vbCrLf &
              "						   then '1'" & vbCrLf &
              "				           else" & vbCrLf &
              "					           case" & vbCrLf &
              "						       when NF.CifFob = 'NEN'" & vbCrLf &
              "						       then '9'" & vbCrLf &
              "						       else '9'" & vbCrLf &
              "					           end" & vbCrLf &
              "					   end" & vbCrLf &
              "		   end Ind_Frt," & vbCrLf &
              "		   isnull(NFER.ChaveNfe,'') AS CHV_CTE," & vbCrLf &
              "        NF.DataDaNota AS DT_DOC, NF.Movimento AS DT_A_P, '' AS TP_CT_e, '' AS CHV_CTE_REF," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxI.Valor END, 0)) AS Vl_Doc," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'DESCONTO' THEN NFxE.Valor END, 0)) AS Vl_Desc," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxI.Valor END, 0)) AS Vl_Serv," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS' THEN NFxE.Base END, 0)) AS Vl_Bc_ICMS," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS' THEN NFxE.Valor END, 0)) AS Vl_Vl_ICMS," & vbCrLf &
              "		   case" & vbCrLf &
              "			   when NF.EntradaSaida_Id = 'E'" & vbCrLf &
              "				   then OEE.DebitaConta" & vbCrLf &
              "				   else OEE.CreditaConta" & vbCrLf &
              "		   end AS Cod_Cta" & vbCrLf &
              "   INTO #D100 " & vbCrLf
        Sql &= GetSqlFROMRegistroD("E")

        Return Sql
    End Function

    Private Function GetSqlFROMRegistroD(ByVal EntSai As String) As String
        Dim SqlTemp As String = "   FROM NotasFiscais NF" & vbCrLf &
                                "  INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf &
                                "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                                "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                                "    AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
                                "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                                "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
                                "    AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
                                "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
                                "  INNER JOIN NotasFiscaisXEncargos NFxE" & vbCrLf &
                                "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                                "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                                "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
                                "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                                "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
                                "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
                                "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
                                "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
                                "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
                                "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                                "  INNER JOIN Clientes Cli" & vbCrLf &
                                "     ON NF.Cliente_Id    = Cli.Cliente_Id" & vbCrLf &
                                "    AND NF.EndCliente_Id = Cli.Endereco_Id" & vbCrLf &
                                "  inner join OperacaoXEstado OE" & vbCrLf &
                                "     on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
                                "  inner join OperacaoXEstadoXEncargo OEE" & vbCrLf &
                                "     on OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
                                "    and OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
                                "  LEFT JOIN NFERealizadas NFER" & vbCrLf &
                                "    ON NFER.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                                "   AND NFER.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                                "   AND NFER.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                                "   AND NFER.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                                "   AND NFER.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                                "   AND NFER.Serie_Id        = NF.Serie_Id" & vbCrLf &
                                "   AND NFER.Nota_Id         = NF.Nota_Id" & vbCrLf &
                                "  left JOIN DocumentoXML DocXML" & vbCrLf &
                                "     ON DocXML.Empresa_Id = NF.Empresa_Id" & vbCrLf &
                                "    AND DocXML.Cliente_Id = NF.Cliente_Id" & vbCrLf &
                                "    AND DocXML.Serie_Id   = NF.Serie_Id" & vbCrLf &
                                "    AND DocXML.Numero_Id  = NF.Nota_Id" & vbCrLf &
                                "  WHERE (ISNULL(NF.Situacao, 1) = 1)" & vbCrLf &
                                "    and (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                                "    AND (NFxI.Cfop_Id BETWEEN 1350 AND 1360 OR" & vbCrLf &
                                "         NFxI.Cfop_Id BETWEEN 2350 AND 2360 OR" & vbCrLf &
                                "         NFxI.Cfop_Id IN (1932,2932)        OR" & vbCrLf &
                                "         NFxI.Cfop_Id BETWEEN 5350 AND 5360 OR" & vbCrLf &
                                "         NFxI.Cfop_Id BETWEEN 6350 AND 6360) " & vbCrLf &
                                "    AND NF.Operacao <> 45 " & vbCrLf &
                                "    AND LEFT(NF.Empresa_Id,8) = '" & EmpresaMestre & "'" & vbCrLf &
                                "    AND NF.TipoDeDocumento in (2,8,10,11,14,57,58,67)" & vbCrLf

        If EntSai.Length > 0 Then
            SqlTemp &= "    and (NF.EntradaSaida_Id = '" & EntSai & "')" & vbCrLf
        End If

        SqlTemp &= "  GROUP BY NF.EntradaSaida_Id, Cli.Cliente_Id,Cli.Endereco_Id, NF.EndCliente_Id," & vbCrLf &
                                "           NF.Serie_Id, NF.Nota_Id, NF.DataDaNota, NF.Movimento, case when ISNULL(DocXML.Tipo,'') = 'Mic' then 0 else case when NF.NossaEmissao = 'S' then 0 else 1 end end," & vbCrLf &
                                "           NF.Empresa_Id, NF.EndEmpresa_Id, NF.CifFob, NF.TipoDeDocumento, NFER.ChaveNfe, OEE.DebitaConta, OEE.CreditaConta " & vbCrLf &
                                " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxE.Valor END, 0)) > 0)" & vbCrLf
        Return SqlTemp
    End Function

    Private Function GetSqlRegistroD101() As String
        Sql = " SELECT 'D101' AS Reg," & vbCrLf &
              "         NF.Empresa_Id, NF.EndEmpresa_Id, NF.Nota_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, Nf.Serie_Id, " & vbCrLf &
              "        '0' AS Ind_Nat_Frt," & vbCrLf &
              "         SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxI.Valor END, 0)) AS Vl_Item," & vbCrLf &
              "         ISNULL(OE.STPISCOFINS, 50) AS CST_PIS," & vbCrLf

        If EmpresaMestre = "04854422" Then
            Sql &= "         Case when Sop.Operacao_Id = 69 then 7 else 7 end as Nat_Bc_Cred," & vbCrLf
        Else
            Sql &= "         Case when Sop.Operacao_Id = 69 then 3 else 7 end as Nat_Bc_Cred," & vbCrLf
        End If

        Sql &= "         SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
               "         ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxE.Percentual END,0) AS Aliq_Pis," & vbCrLf &
               "         SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_PIS," & vbCrLf &
               "		 case" & vbCrLf &
               "		     when NF.EntradaSaida_Id = 'E'" & vbCrLf &
               "			     then OEE.DebitaConta" & vbCrLf &
               "				 else OEE.CreditaConta" & vbCrLf &
               "		     end AS Cod_Cta " & vbCrLf &
               "    INTO #D101 " & vbCrLf &
               "    FROM NotasFiscais AS NF" & vbCrLf &
               "   INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
               "      ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
               "     AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
               "     AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
               "     AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
               "     AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
               "     AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
               "     AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
               "   INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
               "      ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
               "     AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
               "     AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
               "     AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf &
               "     AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
               "     AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
               "     AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
               "     AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
               "     AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
               "     AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
               "   INNER JOIN SubOperacoes AS SOP" & vbCrLf &
               "      ON NFxI.Operacao    = SOP.Operacao_Id " & vbCrLf &
               "     AND NFxI.SubOperacao = SOP.SubOperacoes_Id " & vbCrLf &
               "   INNER JOIN OperacaoXEstado OE" & vbCrLf &
               "      ON OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
               "   inner join OperacaoXEstadoXEncargo OEE" & vbCrLf &
               "      on OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
               "     and OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
               "   WHERE (ISNULL(NF.Situacao, 1) = 1)" & vbCrLf &
               "     AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
               "     and (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
               "     AND (NFxI.Cfop_Id BETWEEN 1350 AND 1360 OR" & vbCrLf &
               "          NFxI.Cfop_Id BETWEEN 2350 AND 2360 OR" & vbCrLf &
               "          NFxI.Cfop_Id IN (1932,2932)        OR" & vbCrLf &
               "          NFxI.Cfop_Id BETWEEN 5350 AND 5360 OR" & vbCrLf &
               "          NFxI.Cfop_Id BETWEEN 6350 AND 6360) " & vbCrLf &
               "     And NF.Operacao <> 45 " & vbCrLf &
               "     And NF.EntradaSaida_Id = 'E'" & vbCrLf &
               "   GROUP BY NFxE.Encargo_Id, OE.STPISCOFINS, " & vbCrLf &
               "         NFxE.Percentual, SOP.GrupoDeContas, SOP.Operacao_Id," & vbCrLf &
               "         NF.Empresa_Id, NF.EndEmpresa_Id, NF.Nota_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, Nf.Serie_Id, OEE.DebitaConta, OEE.CreditaConta " & vbCrLf &
               "  HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxE.Valor END, 0)) > 0) " & vbCrLf

        Return Sql
    End Function

    Private Function GetSqlRegistroD105() As String
        Sql = " SELECT 'D105' AS Reg," & vbCrLf &
              "         NF.Empresa_Id, NF.EndEmpresa_Id, NF.Nota_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, Nf.Serie_Id, " & vbCrLf &
              "        '0' AS Ind_Nat_Frt," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxI.Valor END, 0)) AS Vl_Item," & vbCrLf &
              "        ISNULL(OE.STPISCOFINS, 50) AS CST_PIS," & vbCrLf

        If EmpresaMestre = "04854422" Then
            Sql &= "        Case when Sop.Operacao_Id = 69 then 7 else 7 end as Nat_Bc_Cred," & vbCrLf
        Else
            Sql &= "        Case when Sop.Operacao_Id = 69 then 3 else 7 end as Nat_Bc_Cred," & vbCrLf
        End If

        Sql &= "         SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
               "         ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NFxE.Percentual END,0) AS Aliq_Pis," & vbCrLf &
               "         SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_PIS," & vbCrLf &
               "		 case" & vbCrLf &
               "		     when NF.EntradaSaida_Id = 'E'" & vbCrLf &
               "			     then OEE.DebitaConta" & vbCrLf &
               "				 else OEE.CreditaConta" & vbCrLf &
               "		     end AS Cod_Cta" & vbCrLf &
               "    INTO #D105 " & vbCrLf &
               "    FROM NotasFiscais AS NF " & vbCrLf &
               "   INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
               "      ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
               "     AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
               "     AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
               "     AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
               "     AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
               "     AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
               "     AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
               "   INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
               "      ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
               "     AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
               "     AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
               "     AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
               "     AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
               "     AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
               "     AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
               "     AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
               "     AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
               "     AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
               "   INNER JOIN SubOperacoes AS SOP" & vbCrLf &
               "      ON NFxI.Operacao    = SOP.Operacao_Id" & vbCrLf &
               "     AND NFxI.SubOperacao = SOP.SubOperacoes_Id " & vbCrLf &
               "   INNER JOIN OperacaoXEstado OE" & vbCrLf &
               "      ON OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
               "   inner join OperacaoXEstadoXEncargo OEE" & vbCrLf &
               "      on OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
               "     and OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
               "   WHERE (ISNULL(NF.Situacao, 1) = 1)" & vbCrLf &
               "     AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
               "     and (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
               "     AND (NFxI.Cfop_Id BETWEEN 1350 AND 1360 OR" & vbCrLf &
               "          NFxI.Cfop_Id BETWEEN 2350 AND 2360 OR" & vbCrLf &
               "          NFxI.Cfop_Id IN (1932,2932)        OR" & vbCrLf &
               "          NFxI.Cfop_Id BETWEEN 5350 AND 5360 OR" & vbCrLf &
               "          NFxI.Cfop_Id BETWEEN 6350 AND 6360) " & vbCrLf &
               "     And NF.Operacao <> 45 " & vbCrLf &
               "     And NF.EntradaSaida_Id = 'E'" & vbCrLf &
               "   GROUP BY NFxE.Encargo_Id, ISNULL(OE.STPISCOFINS, 50)," & vbCrLf &
               "         NFxE.Percentual, SOP.GrupoDeContas, SOP.Operacao_Id," & vbCrLf &
               "         NF.Empresa_Id, NF.EndEmpresa_Id, NF.Nota_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, Nf.Serie_Id, OEE.DebitaConta, OEE.CreditaConta " & vbCrLf &
               "  HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NFxE.Valor END, 0)) > 0) " & vbCrLf
        Return Sql
    End Function

    Private Function GetSqlRegistroD200() As String
        'Sql = " SELECT 'D200' AS Reg," & vbCrLf &
        '      "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id," & vbCrLf &
        '      "        '0' AS Ind_Oper," & vbCrLf &
        '      "		   case " & vbCrLf &
        '      "		       when ISNULL(DocXML.Tipo,'') = 'Mic'" & vbCrLf &
        '      "		           then 0" & vbCrLf &
        '      "		           else case " & vbCrLf &
        '      "		                   when NF.NossaEmissao = 'S' " & vbCrLf &
        '      "		                   then 0 " & vbCrLf &
        '      "		                   else 1 " & vbCrLf &
        '      "		                end " & vbCrLf &
        '      "		       end AS Ind_Emit," & vbCrLf &
        '      "        NF.Cliente_Id AS Cod_Part," & vbCrLf &
        '      "        '08' AS Cod_Mod," & vbCrLf &
        '      "        0 AS Cod_Sit," & vbCrLf &
        '      "        NF.Serie_Id AS Ser," & vbCrLf &
        '      "        '' AS Sub," & vbCrLf &
        '      "        NF.Nota_Id AS Num_Doc," & vbCrLf &
        '      "		   case" & vbCrLf &
        '      "			   when NF.CifFob = 'CIF'" & vbCrLf &
        '      "				   then '0'" & vbCrLf &
        '      "				   else" & vbCrLf &
        '      "					   case" & vbCrLf &
        '      "						   when NF.CifFob = 'FOB'" & vbCrLf &
        '      "						   then '1'" & vbCrLf &
        '      "				           else" & vbCrLf &
        '      "					           case" & vbCrLf &
        '      "						       when NF.CifFob = 'NEN'" & vbCrLf &
        '      "						       then '9'" & vbCrLf &
        '      "						       else '9'" & vbCrLf &
        '      "					           end" & vbCrLf &
        '      "					   end" & vbCrLf &
        '      "		   end Ind_Frt," & vbCrLf &
        '      "        NF.DataDaNota AS DT_DOC, NF.Movimento AS DT_A_P, '' AS TP_CT_e, '' AS CHV_CTE_REF," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxI.Valor END, 0)) AS Vl_Doc," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'DESCONTO' THEN NFxE.Valor END, 0)) AS Vl_Desc," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxI.Valor END, 0)) AS Vl_Serv," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS' THEN NFxE.Base END, 0)) AS Vl_Bc_ICMS," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS' THEN NFxE.Valor END, 0)) AS Vl_Vl_ICMS," & vbCrLf &
        '      "        SOP.GrupoDeContas AS Cod_Cta " & vbCrLf &
        '      "   INTO #D200  " & vbCrLf &
        '      "   FROM NotasFiscais AS NF" & vbCrLf &
        '      "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
        '      "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
        '      "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
        '      "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
        '      "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
        '      "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
        '      "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
        '      "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
        '      "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
        '      "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
        '      "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
        '      "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
        '      "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
        '      "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
        '      "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
        '      "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
        '      "    AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
        '      "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
        '      "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
        '      "  INNER JOIN SubOperacoes AS SOP" & vbCrLf &
        '      "     ON NFxI.Operacao    = SOP.Operacao_Id" & vbCrLf &
        '      "    AND NFxI.SubOperacao = SOP.SubOperacoes_Id " & vbCrLf &
        '      "  left JOIN DocumentoXML DocXML" & vbCrLf &
        '      "     ON DocXML.Empresa_Id = NF.Empresa_Id" & vbCrLf &
        '      "    AND DocXML.Cliente_Id = NF.Cliente_Id" & vbCrLf &
        '      "    AND DocXML.Serie_Id   = NF.Serie_Id" & vbCrLf &
        '      "    AND DocXML.Numero_Id  = NF.Nota_Id" & vbCrLf &
        '      "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
        '      "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
        '      "    and (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf

        'If (EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "49673784") OrElse (EmpresaMestre = "53267147") Then
        '    Sql &= "    and (NFxI.Produto_Id Like '50101%') " & vbCrLf
        'Else
        '    Sql &= "    and (NF.Operacao = 45)  " & vbCrLf &
        '      "    and (NF.SubOperacao = 20)  " & vbCrLf &
        '      "    and (NFxI.Produto_Id Like '50201%') " & vbCrLf
        'End If

        'Sql &= "  GROUP BY NF.DataDaNota, NF.Movimento, SOP.GrupoDeContas, case when ISNULL(DocXML.Tipo,'') = 'Mic' then 0 else case when NF.NossaEmissao = 'S' then 0 else 1 end end, " & vbCrLf &
        '      "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, NF.CifFob" & vbCrLf &
        '      " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxE.Valor END, 0)) > 0)" & vbCrLf

        Sql = "SELECT" & vbCrLf &
                "    'D200' AS Reg," & vbCrLf &
                "    NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id," & vbCrLf &
                "    '0' AS Ind_Oper," & vbCrLf &
                "    CASE " & vbCrLf &
                "        WHEN ISNULL(DocXML.Tipo,'') = 'Mic' THEN 0" & vbCrLf &
                "        ELSE CASE WHEN NF.NossaEmissao = 'S' THEN 0 ELSE 1 END" & vbCrLf &
                "    END AS Ind_Emit," & vbCrLf &
                "    NF.Cliente_Id AS Cod_Part," & vbCrLf &
                "    NF.TipoDeDocumento AS Cod_Mod," & vbCrLf &
                "    0 AS Cod_Sit," & vbCrLf &
                "    NF.Serie_Id AS Ser," & vbCrLf &
                "    '' AS Sub," & vbCrLf &
                "    NF.Nota_Id AS Num_Doc," & vbCrLf &
                "    CASE" & vbCrLf &
                "        WHEN NF.CifFob = 'CIF' THEN '0'" & vbCrLf &
                "        WHEN NF.CifFob = 'FOB' THEN '1'" & vbCrLf &
                "        ELSE '9'" & vbCrLf &
                "    END AS Ind_Frt," & vbCrLf &
                "    NF.DataDaNota AS DT_DOC," & vbCrLf &
                "    NF.Movimento  AS DT_A_P," & vbCrLf &
                "    ''            AS TP_CT_e," & vbCrLf &
                "    ''            AS CHV_CTE_REF," & vbCrLf &
                "    NFxI.CFOP_ID," & vbCrLf &
                "    SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO'  THEN NFxI.Valor END, 0)) AS Vl_Doc," & vbCrLf &
                "    SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'DESCONTO' THEN NFxE.Valor END, 0)) AS Vl_Desc," & vbCrLf &
                "    SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO'  THEN NFxI.Valor END, 0)) AS Vl_Serv," & vbCrLf &
                "    SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS'     THEN NFxE.Base  END, 0)) AS Vl_Bc_ICMS," & vbCrLf &
                "    SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS'     THEN NFxE.Valor END, 0)) AS Vl_Vl_ICMS," & vbCrLf &
                "    SOP.GrupoDeContas AS Cod_Cta" & vbCrLf &
                "INTO #TEMPD200" & vbCrLf &
                "FROM NotasFiscais AS NF" & vbCrLf &
                "INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
                "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                "INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
                "    ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                "   AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                "   AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
                "   AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                "   AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
                "   AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
                "   AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
                "   AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
                "   AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
                "   AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                "INNER JOIN SubOperacoes AS SOP" & vbCrLf &
                "    ON NFxI.Operacao    = SOP.Operacao_Id" & vbCrLf &
                "   AND NFxI.SubOperacao = SOP.SubOperacoes_Id " & vbCrLf &
                "LEFT JOIN DocumentoXML DocXML" & vbCrLf &
                "    ON DocXML.Empresa_Id = NF.Empresa_Id" & vbCrLf &
                "   AND DocXML.Cliente_Id = NF.Cliente_Id" & vbCrLf &
                "   AND DocXML.Serie_Id   = NF.Serie_Id" & vbCrLf &
                "   AND DocXML.Numero_Id  = NF.Nota_Id" & vbCrLf &
                "WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
                "  AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                "  AND (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf

        If (EmpresaMestre = "05366261") OrElse (EmpresaMestre = "62747840") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "49673784") OrElse (EmpresaMestre = "53267147") OrElse (EmpresaMestre = "62780383") OrElse (EmpresaMestre = "63358210") Then
            Sql &= "  AND NF.TipoDeDocumento in(8,10,11,57,67)" & vbCrLf
        Else
            Sql &= "  AND (NF.Operacao = 45)  " & vbCrLf &
              "  AND (NF.SubOperacao = 20)  " & vbCrLf &
              "  AND (NFxI.Produto_Id Like '50201%') " & vbCrLf
        End If

        Sql &= "GROUP BY " & vbCrLf &
                "    NF.DataDaNota, NF.Movimento, SOP.GrupoDeContas," & vbCrLf &
                "    CASE WHEN ISNULL(DocXML.Tipo,'') = 'Mic' THEN 0 ELSE CASE WHEN NF.NossaEmissao = 'S' THEN 0 ELSE 1 END END," & vbCrLf &
                "    NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.TipoDeDocumento, NF.Serie_Id, NF.Nota_Id, NF.CifFob, NFxI.CFOP_ID;" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf &
                "SELECT" & vbCrLf &
                "    'D200' AS REG," & vbCrLf &
                "    d.Empresa_Id," & vbCrLf &
                "    d.Cod_Mod as COD_MOD," & vbCrLf &
                "    d.Cod_Sit as COD_SIT," & vbCrLf &
                "    d.Ser AS SER," & vbCrLf &
                "    d.Sub AS SUB," & vbCrLf &
                "    MIN(d.Num_Doc) AS NUM_DOC_INI," & vbCrLf &
                "    MAX(d.Num_Doc) AS NUM_DOC_FIM," & vbCrLf &
                "    CAST(DT_A_P AS date) AS DT_REF," & vbCrLf &
                "    d.CFOP_Id AS CFOP," & vbCrLf &
                "    SUM(d.Vl_Doc)     AS VL_DOC," & vbCrLf &
                "    SUM(d.Vl_Desc)    AS VL_DESC" & vbCrLf &
                "INTO #D200  " & vbCrLf &
                "FROM #TEMPD200 d" & vbCrLf &
                "GROUP BY" & vbCrLf &
                "    d.Empresa_Id," & vbCrLf &
                "    d.Cod_Mod," & vbCrLf &
                "    d.Cod_Sit," & vbCrLf &
                "    d.Ser," & vbCrLf &
                "    d.Sub," & vbCrLf &
                "    CAST(DT_A_P AS date)," & vbCrLf &
                "    d.CFOP_Id" & vbCrLf &
                "ORDER BY" & vbCrLf &
                "    DT_REF;"

        Return Sql
    End Function

    Private Function GetSqlRegistroD201() As String

        'Sql = " SELECT 'D201' AS Reg," & vbCrLf &
        '      "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, " & vbCrLf &
        '      "        '0' AS Ind_Nat_Frt," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxI.Valor END, 0)) AS Vl_Item," & vbCrLf &
        '      "        ISNULL(OE.STPISCOFINS, 0) AS CST_PIS," & vbCrLf &
        '      "        CASE WHEN NF.Operacao = 45 THEN 3 ELSE 7 END AS Nat_Bc_Cred," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
        '      "        ISNULL(CASE WHEN NFxE.Encargo_Id     IN ('PIS','PIS RECUP.') THEN NFxE.Percentual END, 0) AS Aliq_Pis," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_PIS, " & vbCrLf &
        '      "		   CASE" & vbCrLf &
        '      "			   WHEN (LEFT(NF.Empresa_Id,8) = '05366261' OR LEFT(NF.Empresa_Id,8) = '38198213' OR LEFT(NF.Empresa_Id,8) = '40938762' OR LEFT(NF.Empresa_Id,8) = '49673784' OR LEFT(NF.Empresa_Id,8) = '53267147')" & vbCrLf &
        '      "				   THEN iif(isnull(OEE.DebitaConta,'') ='',OEE.CreditaConta,2)" & vbCrLf &
        '      "				   ELSE '1010301'" & vbCrLf &
        '      "			   END AS Cod_Cta" & vbCrLf &
        '      "   INTO #D201" & vbCrLf &
        '      "   FROM NotasFiscais AS NF" & vbCrLf &
        '      "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
        '      "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
        '      "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
        '      "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
        '      "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
        '      "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
        '      "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
        '      "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
        '      "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
        '      "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
        '      "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
        '      "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
        '      "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
        '      "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
        '      "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
        '      "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
        '      "    AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
        '      "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
        '      "    AND NFxI.sequencia_ID    = NFxE.sequencia_ID" & vbCrLf &
        '      "  INNER join OperacaoXEstado OE" & vbCrLf &
        '      "     ON OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
        '      "  INNER join OperacaoXEstadoXEncargo OEE" & vbCrLf &
        '      "     ON OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
        '      "    AND OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
        '      "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
        '      "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
        '      "    and (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf

        'If (EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "49673784") OrElse (EmpresaMestre = "53267147") Then
        '    Sql &= "    and (NFxI.Produto_Id Like '50101%') " & vbCrLf
        'Else
        '    Sql &= "    and (NF.Operacao = 45)  " & vbCrLf &
        '      "    and (NF.SubOperacao = 20)  " & vbCrLf &
        '      "    and (NFxI.Produto_Id Like '50201%') " & vbCrLf
        'End If

        'Sql &= "  GROUP BY NFxE.Encargo_Id, ISNULL(OE.STPISCOFINS, 0), NFxE.Percentual, NF.Operacao, " & vbCrLf &
        '      "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, OEE.DebitaConta, OEE.CreditaConta" & vbCrLf &
        '      " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxE.Valor END, 0)) > 0) " & vbCrLf

        Sql = " SELECT 'D201' AS Reg," & vbCrLf &
              "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id AS Num_Doc, " & vbCrLf &
              "        NF.DataDaNota AS DT_DOC," & vbCrLf &
              "        NFxI.CFOP_ID," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxI.Valor END, 0)) AS Vl_Item," & vbCrLf &
              "        ISNULL(OE.STPISCOFINS, 0) AS CST_PIS," & vbCrLf &
              "        CASE WHEN NF.Operacao = 45 THEN 3 ELSE 7 END AS Nat_Bc_Cred," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
              "        ISNULL(CASE WHEN NFxE.Encargo_Id     IN ('PIS','PIS RECUP.') THEN NFxE.Percentual END, 0) AS Aliq_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_PIS, " & vbCrLf &
              "		   CASE" & vbCrLf &
              "			   WHEN (LEFT(NF.Empresa_Id,8) = '05366261' OR LEFT(NF.Empresa_Id,8) = '62747840' OR LEFT(NF.Empresa_Id,8) = '38198213' OR LEFT(NF.Empresa_Id,8) = '40938762' OR LEFT(NF.Empresa_Id,8) = '49673784' OR LEFT(NF.Empresa_Id,8) = '53267147')" & vbCrLf &
              "				   THEN iif(isnull(OEE.DebitaConta,'') ='',OEE.CreditaConta,2)" & vbCrLf &
              "				   ELSE '1010301'" & vbCrLf &
              "			   END AS Cod_Cta" & vbCrLf &
              "   INTO #TEMPD201" & vbCrLf &
              "   FROM NotasFiscais AS NF" & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
              "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
              "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
              "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
              "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
              "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
              "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
              "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
              "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
              "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
              "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
              "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
              "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
              "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
              "    AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
              "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
              "    AND NFxI.sequencia_ID    = NFxE.sequencia_ID" & vbCrLf &
              "  INNER join OperacaoXEstado OE" & vbCrLf &
              "     ON OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
              "  INNER join OperacaoXEstadoXEncargo OEE" & vbCrLf &
              "     ON OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
              "    AND OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
              "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
              "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "    and (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf

        If (EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "49673784") OrElse (EmpresaMestre = "53267147") OrElse (EmpresaMestre = "62747840") OrElse (EmpresaMestre = "62780383") OrElse (EmpresaMestre = "63358210") Then
            Sql &= "  AND NF.TipoDeDocumento in(8,10,11,57,67)" & vbCrLf
        Else
            Sql &= "  AND (NF.Operacao = 45)  " & vbCrLf &
              "  AND (NF.SubOperacao = 20)  " & vbCrLf &
              "  AND (NFxI.Produto_Id Like '50201%') " & vbCrLf
        End If

        Sql &= "  GROUP BY NFxE.Encargo_Id, ISNULL(OE.STPISCOFINS, 0), NFxE.Percentual, NF.Operacao, " & vbCrLf &
              "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, NF.DataDaNota, NFxI.CFOP_ID, OEE.DebitaConta, OEE.CreditaConta" & vbCrLf &
              " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxE.Valor END, 0)) > 0) " & vbCrLf

        Sql &= "" & vbCrLf &
                "" & vbCrLf &
                "SELECT" & vbCrLf &
                "    'D201' AS REG," & vbCrLf &
                "    d.Empresa_Id," & vbCrLf &
                "    MIN(d.Num_Doc) AS NUM_DOC_INI," & vbCrLf &
                "    MAX(d.Num_Doc) AS NUM_DOC_FIM," & vbCrLf &
                "    CAST(d.DT_DOC AS date) AS DT_REF," & vbCrLf &
                "    d.CFOP_Id AS CFOP," & vbCrLf &
                "    d.CST_PIS," & vbCrLf &
                "    SUM(d.Vl_Item)     AS VL_ITEM," & vbCrLf &
                "    SUM(d.Vl_Bc_Pis)   AS VL_BC_PIS," & vbCrLf &
                "    d.Aliq_Pis," & vbCrLf &
                "    SUM(d.Vl_PIS)     AS VL_PIS," & vbCrLf &
                "    d.Cod_Cta" & vbCrLf &
                "INTO #D201  " & vbCrLf &
                "FROM #TEMPD201 d" & vbCrLf &
                "GROUP BY" & vbCrLf &
                "    d.Empresa_Id," & vbCrLf &
                "    CAST(d.DT_DOC AS date)," & vbCrLf &
                "    d.CFOP_Id," & vbCrLf &
                "    d.CST_PIS," & vbCrLf &
                "    d.Aliq_Pis," & vbCrLf &
                "    d.Cod_Cta" & vbCrLf &
                "ORDER BY" & vbCrLf &
                "    DT_REF;"

        Return Sql
    End Function

    Private Function GetSqlRegistroD205() As String
        'Sql = " SELECT 'D205' AS Reg," & vbCrLf &
        '      "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, " & vbCrLf &
        '      "        '0' AS Ind_Nat_Frt," & vbCrLf &
        '      "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxI.Valor END, 0)) AS Vl_Item," & vbCrLf &
        '      "        ISNULL(OE.STPISCOFINS, 0) AS CST_PIS," & vbCrLf

        'If EmpresaMestre = "04854422" Then
        '    Sql &= "        Case when SOP.Operacao_Id = 69 then 7 else 7 end as Nat_Bc_Cred,"
        'Else
        '    Sql &= "        Case when SOP.Operacao_Id = 69 then 3 else 7 end as Nat_Bc_Cred,"
        'End If

        'Sql &= "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
        '       "        ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NFxE.Percentual END,0) AS Aliq_Pis," & vbCrLf &
        '       "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_PIS," & vbCrLf &
        '       "		CASE" & vbCrLf &
        '       "			WHEN (LEFT(NF.Empresa_Id,8) = '05366261' OR LEFT(NF.Empresa_Id,8) = '38198213' OR LEFT(NF.Empresa_Id,8) = '40938762' OR LEFT(NF.Empresa_Id,8) = '49673784' OR LEFT(NF.Empresa_Id,8) = '53267147')" & vbCrLf &
        '       "				THEN iif(isnull(OEE.DebitaConta,'') ='',OEE.CreditaConta,2)" & vbCrLf &
        '       "				ELSE '1010301'" & vbCrLf &
        '       "			END AS Cod_Cta" & vbCrLf &
        '       "   INTO #D205 " & vbCrLf &
        '       "   FROM NotasFiscais AS NF " & vbCrLf &
        '       "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
        '       "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
        '       "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
        '       "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
        '       "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
        '       "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
        '       "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
        '       "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
        '       "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
        '       "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
        '       "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
        '       "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
        '       "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
        '       "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
        '       "    AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
        '       "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
        '       "    AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
        '       "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
        '       "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
        '       "  INNER JOIN SubOperacoes AS SOP" & vbCrLf &
        '       "     ON NFxI.Operacao    = SOP.Operacao_Id" & vbCrLf &
        '       "    AND NFxI.SubOperacao = SOP.SubOperacoes_Id " & vbCrLf &
        '       "  inner join OperacaoXEstado OE" & vbCrLf &
        '       "     on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
        '      "  INNER join OperacaoXEstadoXEncargo OEE" & vbCrLf &
        '      "     ON OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
        '      "    AND OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
        '       "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
        '       "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
        '       "    and (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf

        'If (EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "49673784") OrElse (EmpresaMestre = "53267147") Then
        '    Sql &= "    and (NFxI.Produto_Id Like '50101%') " & vbCrLf
        'Else
        '    Sql &= "    and (NF.Operacao = 45)  " & vbCrLf &
        '       "    and (NF.SubOperacao = 20)  " & vbCrLf &
        '       "    and (NFxI.Produto_Id Like '50201%') " & vbCrLf
        'End If

        'Sql &= "  GROUP BY NFxE.Encargo_Id, ISNULL(OE.STPISCOFINS, 0)," & vbCrLf &
        '       "        NFxE.Percentual, SOP.GrupoDeContas,  SOP.Operacao_Id," & vbCrLf &
        '       "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, OEE.DebitaConta, OEE.CreditaConta " & vbCrLf &
        '       " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NFxE.Valor END, 0)) > 0) " & vbCrLf

        Sql = " SELECT 'D205' AS Reg," & vbCrLf &
              "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id AS Num_Doc, " & vbCrLf &
              "        NF.DataDaNota AS DT_DOC," & vbCrLf &
              "        NFxI.CFOP_ID," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NFxI.Valor END, 0)) AS Vl_Item," & vbCrLf &
              "        ISNULL(OE.STPISCOFINS, 0) AS CST_PIS," & vbCrLf &
              "        CASE WHEN NF.Operacao = 45 THEN 3 ELSE 7 END AS Nat_Bc_Cred," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
              "        ISNULL(CASE WHEN NFxE.Encargo_Id     IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NFxE.Percentual END, 0) AS Aliq_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_PIS, " & vbCrLf &
              "		   CASE" & vbCrLf &
              "			   WHEN (LEFT(NF.Empresa_Id,8) = '05366261' OR LEFT(NF.Empresa_Id,8) = '38198213' OR LEFT(NF.Empresa_Id,8) = '40938762' OR LEFT(NF.Empresa_Id,8) = '49673784' OR LEFT(NF.Empresa_Id,8) = '53267147' OR LEFT(NF.Empresa_Id,8) = '62747840')" & vbCrLf &
              "				   THEN iif(isnull(OEE.DebitaConta,'') ='',OEE.CreditaConta,2)" & vbCrLf &
              "				   ELSE '1010301'" & vbCrLf &
              "			   END AS Cod_Cta" & vbCrLf &
              "   INTO #TEMPD205" & vbCrLf &
              "   FROM NotasFiscais AS NF" & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
              "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
              "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
              "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
              "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
              "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
              "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
              "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
              "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
              "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
              "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
              "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
              "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
              "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
              "    AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
              "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
              "    AND NFxI.sequencia_ID    = NFxE.sequencia_ID" & vbCrLf &
              "  INNER join OperacaoXEstado OE" & vbCrLf &
              "     ON OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
              "  INNER join OperacaoXEstadoXEncargo OEE" & vbCrLf &
              "     ON OEE.Codigo_id  = OE.Codigo_id" & vbCrLf &
              "    AND OEE.Encargo_Id = 'PRODUTO'" & vbCrLf &
              "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
              "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "    and (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf

        If (EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "49673784") OrElse (EmpresaMestre = "53267147") OrElse (EmpresaMestre = "62747840") OrElse (EmpresaMestre = "62780383") OrElse (EmpresaMestre = "63358210") Then
            Sql &= "  AND NF.TipoDeDocumento in(8,10,11,57,67)" & vbCrLf
        Else
            Sql &= "  AND (NF.Operacao = 45)  " & vbCrLf &
              "  AND (NF.SubOperacao = 20)  " & vbCrLf &
              "  AND (NFxI.Produto_Id Like '50201%') " & vbCrLf
        End If

        Sql &= "  GROUP BY NFxE.Encargo_Id, ISNULL(OE.STPISCOFINS, 0), NFxE.Percentual, NF.Operacao, " & vbCrLf &
              "        NF.Empresa_Id, NF.EndEmpresa_Id, NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, NF.DataDaNota, NFxI.CFOP_ID, OEE.DebitaConta, OEE.CreditaConta" & vbCrLf &
              " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NFxE.Valor END, 0)) > 0) " & vbCrLf

        Sql &= "" & vbCrLf &
                "" & vbCrLf &
                "SELECT" & vbCrLf &
                "    'D205' AS REG," & vbCrLf &
                "    d.Empresa_Id," & vbCrLf &
                "    MIN(d.Num_Doc) AS NUM_DOC_INI," & vbCrLf &
                "    MAX(d.Num_Doc) AS NUM_DOC_FIM," & vbCrLf &
                "    CAST(d.DT_DOC AS date) AS DT_REF," & vbCrLf &
                "    d.CFOP_Id AS CFOP," & vbCrLf &
                "    d.CST_PIS," & vbCrLf &
                "    SUM(d.Vl_Item)     AS VL_ITEM," & vbCrLf &
                "    SUM(d.Vl_Bc_Pis)   AS VL_BC_PIS," & vbCrLf &
                "    d.Aliq_Pis," & vbCrLf &
                "    SUM(d.Vl_PIS)     AS VL_PIS," & vbCrLf &
                "    d.Cod_Cta" & vbCrLf &
                "INTO #D205  " & vbCrLf &
                "FROM #TEMPD205 d" & vbCrLf &
                "GROUP BY" & vbCrLf &
                "    d.Empresa_Id," & vbCrLf &
                "    CAST(d.DT_DOC AS date)," & vbCrLf &
                "    d.CFOP_Id," & vbCrLf &
                "    d.CST_PIS," & vbCrLf &
                "    d.Aliq_Pis," & vbCrLf &
                "    d.Cod_Cta" & vbCrLf &
                "ORDER BY" & vbCrLf &
                "    DT_REF;"

        Return Sql
    End Function

    Private Function GetSqlRegistroD500() As String
        Sql = " SELECT 'D500' AS Reg," & vbCrLf &
              "        NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id," & vbCrLf &
              "        '0' AS Ind_Oper," & vbCrLf &
              "        '1' AS Ind_Emit," & vbCrLf &
              "        Clientes.Cliente_Id + Convert(varchar, Clientes.Endereco_Id) AS Cod_Part," & vbCrLf &
              "        '22' AS Cod_Mod," & vbCrLf &
              "        '0' AS Cod_Sit," & vbCrLf &
              "        NotasFiscais.Serie_Id AS Ser," & vbCrLf &
              "        '' as Sub," & vbCrLf &
              "        NotasFiscais.Nota_Id AS Num_Doc," & vbCrLf &
              "        NotasFiscais.DataDaNota AS DT_Doc," & vbCrLf &
              "        NotasFiscais.Movimento AS DT_A_P," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO' THEN NotasFiscaisXItens.Valor END, 0)) AS Vl_Doc," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'DESCONTO' THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Desc," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'LIQUIDO' THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Serv," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor END, 0) -" & vbCrLf &
              "            ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Serv_Nt," & vbCrLf &
              "        0 AS VL_Terc," & vbCrLf &
              "        0 as VL_DA," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Base END, 0)) AS Vl_Bc_Icms," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Icms," & vbCrLf &
              "        '' as Cod_Inf," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NotasFiscaisXEncargos.Valor END, 0)) AS Vl_Cofins,Clientes.Endereco_Id" & vbCrLf &
              "   INTO #D500 " & vbCrLf &
              "   FROM NotasFiscais" & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
              "     ON NotasFiscais.Empresa_Id       = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
              "    AND NotasFiscais.EndEmpresa_Id    = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscais.Cliente_Id       = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id    = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscais.EntradaSaida_Id  = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscais.Serie_Id         = NotasFiscaisXItens.Serie_Id" & vbCrLf &
              "    AND NotasFiscais.Nota_Id          = NotasFiscaisXItens.Nota_Id" & vbCrLf &
              "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
              "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
              "  INNER JOIN Clientes" & vbCrLf &
              "     ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
              "  WHERE (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              "    AND (left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "    AND (NotasFiscaisXItens.Cfop_Id BETWEEN 1301 AND 1306 OR NotasFiscaisXItens.Cfop_Id BETWEEN 2301 AND 2306 )" & vbCrLf &
              "    AND (NotasFiscais.EntradaSaida_Id = 'E')" & vbCrLf &
              "  GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, Clientes.Cliente_Id,Clientes.Endereco_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.Movimento, NotasFiscais.DataDaNota" & vbCrLf &
              " HAVING (SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor END, 0)) > 0) " & vbCrLf

        '"    AND (NotasFiscais.Empresa_Id= '" & .Item("Cliente_Id") & "') " & vbCrLf & _

        Return Sql
    End Function

    Private Function GetSqlRegistroD501() As String
        Sql = " SELECT 'D501' AS Reg," & vbCrLf &
              "        NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id," & vbCrLf &
              "        ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN OxE.STPISCOFINS END, 0) AS CST_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO'THEN NotasFiscaisXItens.Valor END, 0)) AS Vl_Item," & vbCrLf &
              "        13 as Nat_Bc_Cred," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS Vl_Bc_Pis," & vbCrLf &
              "        ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Percentual END,0) AS Aliq_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END END, 0)) AS Vl_PIS," & vbCrLf &
              "        SubOperacoes.GrupoDeContas AS Cod_Cta" & vbCrLf &
              "   INTO #D501 " & vbCrLf &
              "   FROM NotasFiscais" & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
              "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
              "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
              "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf &
              "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
              "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id " & vbCrLf &
              "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
              "  INNER JOIN OperacaoXEstado OxE" & vbCrLf &
              "     ON NotasFiscaisXItens.OperacaoXEstado = OxE.Codigo_Id  " & vbCrLf &
              "  INNER JOIN SubOperacoes" & vbCrLf &
              "     ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
              "  INNER JOIN Clientes" & vbCrLf &
              "     ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id" & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf &
              "   LEFT OUTER JOIN NFEREALIZADAS NFE" & vbCrLf &
              "     ON NotasFiscais.Empresa_Id      = NFE.Empresa_Id" & vbCrLf &
              "    AND NotasFiscais.EndEmpresa_Id   = NFE.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscais.Cliente_Id      = NFE.Cliente_Id" & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscais.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscais.Serie_Id        = NFE.Serie_Id" & vbCrLf &
              "    AND NotasFiscais.Nota_Id         = NFE.Nota_Id" & vbCrLf &
              "  WHERE (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              "    AND (left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "    AND (NotasFiscaisXItens.Cfop_Id BETWEEN 1301 AND 1306 OR NotasFiscaisXItens.Cfop_Id BETWEEN 2301 AND 2306 )" & vbCrLf &
              "    AND (NotasFiscais.EntradaSaida_Id = 'E')" & vbCrLf &
              " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscaisXEncargos.Encargo_Id, OxE.STPISCOFINS, NotasFiscaisXEncargos.Percentual, SubOperacoes.GrupoDeContas" & vbCrLf &
              " HAVING (SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('PIS','PIS RECUP.') THEN NotasFiscaisXEncargos.Valor END, 0)) > 0)" & vbCrLf

        '"    AND (NotasFiscais.Empresa_Id = '" & .Item("Cliente_Id") & "') " & vbCrLf & _
        '     "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf & _
        '     "    AND NotasFiscais.Cliente_Id = '" & Cliente & "' " & vbCrLf & _
        '     "    AND NotasFiscais.EndCliente_Id = " & EndCliente & "  " & vbCrLf & _
        '     "    And NotasFiscais.Nota_Id = " & Nota & " " & vbCrLf & _
        '     "    AND NotasFiscais.Serie_Id = '" & Serie & "'  " & vbCrLf & _


        Return Sql
    End Function

    Private Function GetSqlRegistroD505() As String
        Sql = " SELECT 'D505' AS Reg," & vbCrLf &
            "        NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id," & vbCrLf &
              "        ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN OxE.STPISCOFINS END, 0) AS CST_Cofins," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO' THEN NotasFiscaisXItens.Valor END, 0)) AS Vl_Item," & vbCrLf &
              "        13 as Nat_Bc_Cred," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.BaseNova,0) > 0 then NotasFiscaisXEncargos.BaseNova else NotasFiscaisXEncargos.Base END END, 0)) AS Vl_Bc_Cofins," & vbCrLf &
              "        ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id     IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NotasFiscaisXEncargos.Percentual END,0) AS Aliq_Cofins," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NotasFiscaisXEncargos.ValorNovo,0) > 0 then NotasFiscaisXEncargos.ValorNovo else NotasFiscaisXEncargos.Valor END END, 0)) AS Vl_Cofins," & vbCrLf &
              "        SubOperacoes.GrupoDeContas AS Cod_Cta" & vbCrLf &
              "   INTO #D505 " & vbCrLf &
              "   FROM NotasFiscais " & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
              "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
              "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
              "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
              "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf &
              "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id " & vbCrLf &
              "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Sequencia_Id    = NotasFiscaisXEncargos.Sequencia_Id" & vbCrLf &
               "  INNER JOIN OperacaoXEstado OxE" & vbCrLf &
              "     ON NotasFiscaisXItens.OperacaoXEstado = OxE.Codigo_Id  " & vbCrLf &
              "  INNER JOIN SubOperacoes" & vbCrLf &
              "     ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
              "  INNER JOIN Clientes" & vbCrLf &
              "     ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf &
              "   LEFT OUTER JOIN NFEREALIZADAS NFE" & vbCrLf &
              "     ON NotasFiscais.Empresa_Id      = NFE.Empresa_Id" & vbCrLf &
              "    AND NotasFiscais.EndEmpresa_Id   = NFE.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscais.Cliente_Id      = NFE.Cliente_Id" & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscais.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscais.Serie_Id        = NFE.Serie_Id" & vbCrLf &
              "    AND NotasFiscais.Nota_Id         = NFE.Nota_Id" & vbCrLf &
              "  WHERE (ISNULL(NotasFiscais.Situacao, 1) = 1)" & vbCrLf &
              "    AND (left(NotasFiscais.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "    AND (NotasFiscaisXItens.Cfop_Id BETWEEN 1301 AND 1306 OR NotasFiscaisXItens.Cfop_Id BETWEEN 2301 AND 2306)" & vbCrLf &
              "    AND (NotasFiscais.EntradaSaida_Id = 'E')" & vbCrLf &
              "  GROUP BY NotasFiscais.Empresa_Id,NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscaisXEncargos.Encargo_Id, OxE.STPISCOFINS, NotasFiscaisXEncargos.Percentual, SubOperacoes.GrupoDeContas" & vbCrLf &
              " HAVING (SUM(ISNULL(CASE WHEN NotasFiscaisXEncargos.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NotasFiscaisXEncargos.Valor END, 0)) > 0)" & vbCrLf


        '"    AND (NotasFiscais.Empresa_Id = '" & .Item("Cliente_Id") & "')  " & vbCrLf & _
        '      "    AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf & _
        '      "    AND NotasFiscais.Cliente_Id = '" & Cliente & "' " & vbCrLf & _
        '      "    AND NotasFiscais.EndCliente_Id = " & EndCliente & "  " & vbCrLf & _
        '      "    AND NotasFiscais.Nota_Id = " & Nota & " " & vbCrLf & _
        '      "    AND NotasFiscais.Serie_Id = '" & Serie & "'  " & vbCrLf & _


        Return Sql
    End Function

    Private Sub CompoeRegistroD101(ByRef ds As DataSet, ByRef Arq As ArrayList, ByRef RegD101 As Integer, ByRef RegGeral As Integer,
                                  D100Empresa As String, D100EndEmpresa As Integer, D100NF As Integer, D100Cliente As String, D100EndCliente As Integer, D100Serie As String, D100EntradaSaida As String)

        Dim linhaTemp As String = String.Empty

        ' D101 Complemento do Documento de Transporte PIS ----------

        For Each drD101 As DataRow In ds.Tables("D101").Select("Empresa_Id =" & D100Empresa & " AND EndEmpresa_Id = " & D100EndEmpresa & " AND Nota_id = " & D100NF & " AND Cliente_id = '" & D100Cliente & "' AND EndCliente_id = " &
                                                                 D100EndCliente & " AND Serie_id = '" & D100Serie & "' AND EntradaSaida_id = '" & D100EntradaSaida & "'")
            With drD101
                linhaTemp = "|" & .Item("Reg")
                linhaTemp &= "|" & .Item("IND_NAT_FRT")
                linhaTemp &= "|" & .Item("VL_BC_PIS")

                If Format(CInt(.Item("CST_PIS")), "00") > 0 Then
                    linhaTemp &= "|" & Format(CInt(.Item("CST_PIS")), "00")
                Else
                    linhaTemp &= "|50"
                End If

                linhaTemp &= "|" & Format(.Item("NAT_BC_CRED"), "00")
                linhaTemp &= "|" & .Item("VL_BC_PIS")
                linhaTemp &= "|" & CDbl(.Item("ALIQ_PIS")).ToString("N2")
                linhaTemp &= "|" & .Item("VL_PIS")
                linhaTemp &= "|" & .Item("COD_CTA")
                linhaTemp &= "|"

            End With

            Arq.Add(linhaTemp)
            RegD101 += 1
            RegGeral += 1

        Next   ''Fim D101 - Pis Sobre Fretes ---------
    End Sub

    Private Sub CompoeRegistroD105(ByRef ds As DataSet, ByRef Arq As ArrayList, ByRef RegD105 As Integer, ByRef RegGeral As Integer,
                                  D100Empresa As String, D100EndEmpresa As Integer, D100NF As Integer, D100Cliente As String, D100EndCliente As Integer, D100Serie As String, D100EntradaSaida As String)

        Dim linhaTemp As String = String.Empty

        '' D105 Complemento do Documento de Transporte Cofins -------
        For Each drD105 As DataRow In ds.Tables("D105").Select("Empresa_Id =" & D100Empresa & " AND EndEmpresa_Id = " & D100EndEmpresa & " AND Nota_id = " & D100NF & " AND Cliente_id = '" & D100Cliente & "' AND EndCliente_id = " &
                                                                D100EndCliente & " AND Serie_id = '" & D100Serie & "' AND EntradaSaida_id = '" & D100EntradaSaida & "'")
            With drD105
                linhaTemp = "|" & .Item("Reg")
                linhaTemp &= "|" & .Item("IND_NAT_FRT")
                linhaTemp &= "|" & .Item("VL_BC_PIS")

                If Format(CInt(.Item("CST_PIS")), "00") > 0 Then
                    linhaTemp &= "|" & Format(CInt(.Item("CST_PIS")), "00")
                Else
                    linhaTemp &= "|50"
                End If

                linhaTemp = linhaTemp & "|" & Format(.Item("NAT_BC_CRED"), "00")
                linhaTemp &= "|" & .Item("VL_BC_PIS")
                linhaTemp &= "|" & CDbl(.Item("ALIQ_PIS")).ToString("N2")
                linhaTemp &= "|" & .Item("VL_PIS")
                linhaTemp &= "|" & .Item("COD_CTA")
                linhaTemp &= "|"
            End With

            ArquivoAux.Add(linhaTemp)
            RegD105 += 1
            RegGeral += 1
        Next
    End Sub

    Private Sub CompoeRegistroD201(ByRef ds As DataSet, ByRef Arq As ArrayList, ByRef RegD201 As Integer, ByRef RegGeral As Integer,
                        D200Empresa As String, NUM_DOC_INI As Integer, NUM_DOC_FIM As Integer, DT_REF As String, CFOP As Integer)

        Dim linhaTemp As String = String.Empty

        ' D201 Complemento do Documento de Transporte PIS -------

        'For Each drD201 In ds.Tables("D201").Rows
        For Each drD201 As DataRow In ds.Tables("D201").Select("Empresa_Id = " & D200Empresa & " AND NUM_DOC_INI = " & NUM_DOC_INI & " AND NUM_DOC_FIM = " & NUM_DOC_FIM & " AND DT_REF = '" &
                                                               DT_REF & "' AND CFOP = " & CFOP & "")
            With drD201

                linhaTemp = "|" & .Item("Reg")
                linhaTemp &= "|" & Format(CInt(.Item("CST_PIS")), "00")

                linhaTemp &= "|" & .Item("VL_ITEM")
                linhaTemp &= "|" & .Item("VL_BC_PIS")

                linhaTemp &= "|" & CDbl(.Item("ALIQ_PIS")).ToString("N2")
                linhaTemp &= "|" & .Item("VL_PIS")
                linhaTemp &= "|" & .Item("COD_CTA")
                linhaTemp &= "|"

                Arq.Add(linhaTemp)
                RegD201 += 1
                RegGeral += 1

            End With
        Next

    End Sub

    Private Sub CompoeRegistroD205(ByRef ds As DataSet, ByRef Arq As ArrayList, ByRef RegD205 As Integer, ByRef RegGeral As Integer,
                       D200Empresa As String, NUM_DOC_INI As Integer, NUM_DOC_FIM As Integer, DT_REF As String, CFOP As Integer)

        '' D205 Complemento do Documento de Transporte Cofins ---

        Dim linhaTemp As String = String.Empty

        For Each drD205 As DataRow In ds.Tables("D205").Select("Empresa_Id =" & D200Empresa & " AND NUM_DOC_INI = " & NUM_DOC_INI & " AND NUM_DOC_FIM = " & NUM_DOC_FIM & " AND DT_REF = '" &
                                                              DT_REF & "' AND CFOP = " & CFOP & "")
            With drD205

                linhaTemp = "|" & .Item("Reg")

                linhaTemp &= "|" & Format(CInt(.Item("CST_PIS")), "00")

                linhaTemp &= "|" & .Item("VL_ITEM")
                linhaTemp &= "|" & .Item("VL_BC_PIS")

                linhaTemp &= "|" & CDbl(.Item("ALIQ_PIS")).ToString("N2")
                linhaTemp &= "|" & .Item("VL_PIS")
                linhaTemp &= "|" & .Item("COD_CTA")
                linhaTemp &= "|"

                Arq.Add(linhaTemp)
                RegD205 += 1
                RegGeral += 1

            End With
        Next
    End Sub

    Public Sub CompoeRegistroD501(ByRef ds As DataSet, ByRef Arq As ArrayList, ByRef RegD501 As Integer, ByRef RegGeral As Integer,
                       D500Empresa As String, D500EndEmpresa As Integer, D500NF As Integer, D500Cliente As String, D500EndCliente As Integer, D500Serie As String, D500EntradaSaida As String)
        'D501 Complemento da Operação Pis -----

        For Each drD501 As DataRow In ds.Tables("D501").Select("Empresa_Id =" & D500Empresa & " AND EndEmpresa_Id = " & D500EndEmpresa & " AND Nota_id = " & D500NF & " AND Cliente_id = '" & D500Cliente & "' AND EndCliente_id = " &
                                                              D500EndCliente & " AND Serie_id = '" & D500Serie & "' AND EntradaSaida_id = '" & D500EntradaSaida & "'")
            With drD501
                linha = "|" & .Item("Reg")
                linha &= "|" & Format(CInt(.Item("CST_PIS")), "00")
                linha &= "|" & .Item("VL_ITEM")
                linha &= "|" & .Item("NAT_BC_CRED")
                linha &= "|" & .Item("VL_BC_PIS")
                linha &= "|" & CDbl(.Item("ALIQ_PIS")).ToString("N2")
                linha &= "|" & .Item("VL_PIS")
                linha &= "|" & .Item("COD_CTA")
                linha &= "|"
            End With

            Arq.Add(linha)
            RegD501 += 1
            RegGeral += 1
        Next

    End Sub

    Public Sub CompoeRegistroD505(ByRef ds As DataSet, ByRef Arq As ArrayList, ByRef RegD505 As Integer, ByRef RegGeral As Integer,
                      D500Empresa As String, D500EndEmpresa As Integer, D500NF As Integer, D500Cliente As String, D500EndCliente As Integer, D500Serie As String, D500EntradaSaida As String)

        'D505 Complemento da Operação Cofins ------------------------------

        For Each drD505 As DataRow In ds.Tables("D505").Select("Empresa_Id =" & D500Empresa & " AND EndEmpresa_Id = " & D500EndEmpresa & " AND Nota_id = " & D500NF & " AND Cliente_id = '" & D500Cliente & "' AND EndCliente_id = " &
                                                             D500EndCliente & " AND Serie_id = '" & D500Serie & "' AND EntradaSaida_id = '" & D500EntradaSaida & "'")
            With drD505
                linha = "|" & .Item("Reg")
                linha &= "|" & Format(CInt(.Item("CST_Cofins")), "00")
                linha &= "|" & .Item("VL_ITEM")
                linha &= "|" & .Item("NAT_BC_CRED")
                linha &= "|" & .Item("Vl_Bc_Cofins")
                linha &= "|" & CDbl(.Item("Aliq_Cofins")).ToString("N2")
                linha &= "|" & .Item("Vl_Cofins")
                linha &= "|" & .Item("COD_CTA")
                linha &= "||"
            End With

            Arq.Add(linha)
            RegD505 += 1
            RegGeral += 1
        Next
    End Sub

#End Region

#Region "BLOCO F"
    Private Function ConsultaF001() As DataSet
        Dim SqlT As String = String.Empty

        ' Registro F001  - Abertura do Bloco F
        SqlT = " SELECT NF.Empresa_Id" & vbCrLf &
               "   FROM NotasFiscais AS NF" & vbCrLf &
               "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
               "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
               "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
               "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
               "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
               "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
               "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
               "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
               "  INNER JOIN NotasFiscaisXEncargos AS NFxE " & vbCrLf &
               "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
               "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
               "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
               "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
               "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
               "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
               "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
               "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
               "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
               "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
               "  INNER JOIN Clientes" & vbCrLf &
               "     ON NF.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
               "    AND NF.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
               "  INNER join OperacaoXEstado OE" & vbCrLf &
               "     ON OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
               "  WHERE Substring(NF.Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
               "    AND (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
               "    AND (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
               "    AND (NFxI.CFOP_Id IN (1556, 1949, 2556, 1653, 2653, 3101, 5949))" & vbCrLf &
               "    And not OE.STPisCofins in  (49, 99) " & vbCrLf &
               "  GROUP BY NF.Empresa_Id" & vbCrLf &
               " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf &
               " UNION" & vbCrLf &
               " SELECT Empresa_Id" & vbCrLf &
               "   FROM PisCofins_Reg_F130" & vbCrLf &
               "  WHERE Substring(Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
               "    AND (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
               " Union" & vbCrLf &
               " SELECT Empresa_Id" & vbCrLf &
               "   FROM PisCofins_Reg_F120" & vbCrLf &
               "  WHERE Substring(Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
               "    And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
               "  UNION " & vbCrLf &
               " SELECT Empresa_Id" & vbCrLf &
               "   FROM Razao" & vbCrLf &
               "  WHERE (Left(Conta_Id, 7) in ('4020203') or Conta_Id in ('403010101', '403010102', '403010103', '403010104', '402020401', '402020402','303030301', '303030302', '303030303','303030304','303030305','402020399'))" & vbCrLf &
               "    AND (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
               "    AND Substring(Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
               "    AND NOT Lote_Id in (7500) " & vbCrLf &
               "  UNION " & vbCrLf &
               " SELECT Empresa_Id " & vbCrLf &
               "   FROM Razao" & vbCrLf &
               "  WHERE LEFT(Conta_Id, 7) in ('1010301', '2010101') " & vbCrLf &
               "    AND Lote_ID = 30" & vbCrLf &
               "    AND (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
               "    AND Substring(Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf

        If cboIncidencia.SelectedValue = 2 Then
            SqlT &= "  UNION " & vbCrLf &
                    " SELECT NF.Empresa_Id" & vbCrLf &
                    "   FROM NotasFiscais AS NF" & vbCrLf &
                    "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
                    "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                    "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                    "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                    "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                    "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                    "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                    "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                    "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
                    "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                    "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                    "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
                    "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                    "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
                    "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
                    "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
                    "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
                    "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
                    "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                    "  INNER JOIN Clientes" & vbCrLf &
                    "     ON NF.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
                    "    AND NF.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                    "  INNER JOIN OperacaoXEstado OE" & vbCrLf &
                    "     ON OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
                    "  WHERE Substring(NF.Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
                    "    AND (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
                    "    AND (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') " & vbCrLf &
                    "    AND not OE.STPisCofins in  (49, 99) " & vbCrLf &
                    "  GROUP BY NF.Empresa_ID" & vbCrLf &
                    " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf
        End If

        Return Banco.ConsultaDataSet(SqlT, "F001")
    End Function

    Private Function Consulta_F010_F100_F120_F130_F550_F600() As DataSet
        Dim SqlT As String = String.Empty

        '******************************************************************************************************************************************************************************
        '*********************************************************************** F010 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        'SqlT &= GetSqlRegistroF010()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** F100 e Razão Contábil ********************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroF100()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** F120 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroF120()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** F130 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroF130()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** F550 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroF550()



        '******************************************************************************************************************************************************************************
        '*********************************************************************** F600 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistroF600()


        '******************************************************************************************************************************************************************************
        '************************************************************************* RESULTADO ******************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= " --SELECT * FROM #F010 " & vbCrLf &
                " SELECT * FROM #F100 " & vbCrLf &
                " SELECT * FROM #F100RazaoContabil " & vbCrLf &
                " SELECT * FROM #F120 " & vbCrLf &
                " SELECT * FROM #F130 " & vbCrLf &
                " SELECT * FROM #F550 " & vbCrLf &
                " SELECT * FROM #F600 " & vbCrLf

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(SqlT, "Consulta")
        'ds.Tables(0).TableName = "F010"
        ds.Tables(0).TableName = "F100"
        ds.Tables(1).TableName = "F100RazaoContabil"
        ds.Tables(2).TableName = "F120"
        ds.Tables(3).TableName = "F130"
        ds.Tables(4).TableName = "F550"
        ds.Tables(5).TableName = "F600"

        Return ds
    End Function

    Private Function GetSqlRegistroF010() As String
        Dim SqlT As String = String.Empty

        ' Registro F010  - Identificação do Estabelecimento bloco F
        If cboIncidencia.SelectedValue <> 2 Then
            SqlT = " SELECT 'F010' AS Reg, Cliente_Id " & vbCrLf &
                   "   INTO #F010 " & vbCrLf &
                   "   FROM (" & vbCrLf &
                   " SELECT 'F010' AS Reg," & vbCrLf &
                   "        NF.Empresa_Id As Cliente_Id" & vbCrLf &
                   "   FROM NotasFiscais AS NF " & vbCrLf &
                   "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
                   "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                   "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                   "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                   "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                   "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                   "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                   "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                   "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
                   "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                   "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                   "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
                   "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                   "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
                   "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
                   "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
                   "    AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
                   "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
                   "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                   "   INNER JOIN Clientes" & vbCrLf &
                   "      ON NF.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
                   "     AND NF.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                   "   inner join OperacaoXEstado OE" & vbCrLf &
                   "      on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
                   "   WHERE Substring(NF.Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
                   "     And (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
                   "     And not OE.STPisCofins in  (49, 99) " & vbCrLf &
                   "     And (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') AND (NFxI.CFOP_Id IN (1556, 2556, 1949, 1653, 2653, 3101, 5949 ))" & vbCrLf &
                   "   GROUP BY   NF.Empresa_Id" & vbCrLf &
                   "  HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf &
                   "  Union " & vbCrLf &
                   " SELECT  'F010' AS Reg, Empresa_Id as Cliente_Id" & vbCrLf &
                   "   FROM PisCofins_Reg_F130 " & vbCrLf &
                   "  WHERE Substring(Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
                   "    And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "  Union " & vbCrLf &
                   " SELECT  'F010' AS Reg, Empresa_Id as Cliente_Id" & vbCrLf &
                   "   FROM PisCofins_Reg_F120 " & vbCrLf &
                   "  WHERE Substring(Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
                   "    And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   " Union" & vbCrLf &
                   " select 'F010' AS Reg, Empresa_Id as Cliente_Id" & vbCrLf &
                   " From Razao" & vbCrLf &
                   " Where (Left(Conta_Id, 7) in ('4020203') or  Conta_Id in ('403010101', '403010102', '403010103', '403010104', '402020401', '402020402','303030301', '303030302', '303030303','303030304','303030305','402020399'))" & vbCrLf &
                   " And (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   " And  Substring(Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
                   " And not Lote_Id in (7500) " & vbCrLf &
                   " Group By  Empresa_Id " & vbCrLf &
                   " Union" & vbCrLf &
                   " Select 'F010' AS Reg, Empresa_Id From Razao" & vbCrLf &
                   "        Where left(Conta_Id, 7) in ('1010301', '2010101') And Lote_ID = 30" & vbCrLf &
                   "        And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "        And not Lote_Id in (7500) " & vbCrLf &
                   "        AND Substring(Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
                   "        Group By Empresa_Id" & vbCrLf &
                   " ) as Consulta" & vbCrLf &
                   " Order By Consulta.Cliente_Id"
        ElseIf cboIncidencia.SelectedValue = 2 Then
            SqlT = " SELECT 'F010' AS Reg," & vbCrLf &
                   "        NF.Empresa_Id As Cliente_Id" & vbCrLf &
                   "   INTO #F010 " & vbCrLf &
                   "   FROM NotasFiscais AS NF " & vbCrLf &
                   "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
                   "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                   "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
                   "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                   "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                   "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                   "    AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
                   "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
                   "  INNER JOIN NotasFiscaisXEncargos AS NFxE " & vbCrLf &
                   "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                   "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                   "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf &
                   "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                   "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
                   "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
                   "    AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf &
                   "    AND NFxI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
                   "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
                   "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                   "  INNER JOIN Clientes" & vbCrLf &
                   "     ON NF.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf &
                   "    AND NF.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                   "   inner join OperacaoXEstado OE" & vbCrLf &
                   "      on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
                   "  WHERE Substring(NF.Empresa_Id,1,8) = '" & EmpresaMestre & "'" & vbCrLf &
                   "    And (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
                   "    And (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "    And not OE.STPisCofins in   (49, 99) " & vbCrLf &
                   "  GROUP BY NF.Empresa_ID" & vbCrLf &
                   " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf
        End If
        Return SqlT
    End Function

    Private Function GetSqlRegistroF100() As String
        Dim SqlT As String = String.Empty

        ' Registro F100 Outros Creditos -------------------------
        SqlT = " SELECT 'F100' AS Reg, NF.Empresa_ID, NF.Cliente_Id AS Cod_Part, NF.EndCliente_Id, " & vbCrLf &
              "        CASE WHEN NF.EntradaSaida_Id = 'S' THEN 1 ELSE 0 End AS Cod_Mod," & vbCrLf &
              "        ISNULL(OE.STPISCOFINS, 0) AS Cod_Sit," & vbCrLf &
              "        ISNULL( NFxI.cfop_Id, 0) AS Cfop, " & vbCrLf &
              "        NF.Serie_Id AS Ser, NF.Nota_Id AS Num_Doc, NF.Movimento AS DT_Doc, NF.DataDaNota AS DT_Ent," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id in ('PRODUTO', 'DESP.ADUANEIRAS') THEN NFxE.Valor      END, 0)) AS Vl_Doc," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS'    THEN NFxE.Valor      END, 0)) AS Vl_Icms," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Base_Pis," & vbCrLf &
              "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxE.Percentual END, 0)) AS Per_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Base_Cofins," & vbCrLf &
              "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN NFxE.Percentual END, 0)) AS Per_Cofins," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('COFINS','COFINS RECUP.','COFINS A REC.') THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Valor_Cofins," & vbCrLf &
              "        CASE" & vbCrLf &
              "        WHEN NFxI.cfop_Id IN(1949, 3101, 3556, 5949, 6949)" & vbCrLf &
              "        THEN '402020399'" & vbCrLf &
              "        Else ''" & vbCrLf &
              "        End As Conta" & vbCrLf &
              "   INTO #F100  " & vbCrLf &
              "   FROM NotasFiscais AS NF " & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
              "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
              "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
              "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
              "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
              "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
              "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
              "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
              "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
              "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
              "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
              "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
              "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
              "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
              "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
              "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
              "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
              "  INNER JOIN SubOperacoes" & vbCrLf &
              "     ON NF.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
              "    AND NF.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
              "   inner join OperacaoXEstado OE" & vbCrLf &
              "      on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
              "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
              "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "    AND (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "    AND NFxI.cfop_Id in (1949, 3101, 3556, 5949, 6949) And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "'" & vbCrLf &
              "    And not ISNULL(OE.STPISCOFINS, 0) in  (49, 98, 99) " & vbCrLf &
              "  GROUP BY NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, NF.Movimento, NF.DataDaNota, NFxI.cfop_Id,  ISNULL(OE.STPISCOFINS, 0), NF.Empresa_Id " & vbCrLf &
              " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PIS','PIS RECUP.') THEN NFxE.Valor END, 0)) > 0)" & vbCrLf &
              " Union " & vbCrLf &
              " SELECT 'F100' AS Reg, NF.Empresa_Id,  NF.Cliente_Id AS Cod_Part, NF.EndCliente_Id, " & vbCrLf &
              "        0  AS Cod_Mod," & vbCrLf &
              "        66 AS Cod_Sit," & vbCrLf &
              "        ISNULL(NFxI.cfop_Id, 0) AS Cfop, " & vbCrLf &
              "        NF.Serie_Id AS Ser, NF.Nota_Id AS Num_Doc, NF.Movimento AS DT_Doc, NF.DataDaNota AS DT_Ent," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id in ('PIS CRED PRE') THEN NFxE.Base  END, 0)) AS Vl_Doc," & vbCrLf &
              "        0 AS Vl_Icms," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS CRED PRE'     THEN NFxE.Base       END, 0)) AS Base_Pis," & vbCrLf &
              "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS CRED PRE'     THEN NFxE.Percentual END, 0)) AS Per_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS CRED PRE'     THEN NFxE.Valor      END, 0)) AS Vl_Pis," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS CRED PRE'  THEN NFxE.Base       END, 0)) AS Base_Cofins," & vbCrLf &
              "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS CRED PRE'  THEN NFxE.Percentual END, 0)) AS Per_Cofins," & vbCrLf &
              "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS CRED PRE'  THEN NFxE.Valor      END, 0)) AS Valor_Cofins," & vbCrLf &
              "        CASE" & vbCrLf &
              "        WHEN NFxI.cfop_Id IN(1949, 3101, 3556, 5949, 6949)" & vbCrLf &
              "        THEN '402020399'" & vbCrLf &
              "        Else ''" & vbCrLf &
              "        End As Conta" & vbCrLf &
              "   FROM NotasFiscais AS NF" & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
              "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
              "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
              "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
              "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
              "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
              "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
              "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
              "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
              "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
              "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
              "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
              "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
              "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
              "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
              "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
              "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
              "  INNER JOIN SubOperacoes" & vbCrLf &
              "     ON NF.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
              "    AND NF.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
              "  INNER join OperacaoXEstado OE" & vbCrLf &
              "     ON OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
              "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
              "    AND (Left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "    AND (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "    AND NFxE.Encargo_Id in ('PIS CRED PRE', 'COFINS CRED PRE')" & vbCrLf &
              "    And not OE.STPisCofins in  (49, 98, 99) " & vbCrLf &
              "  GROUP BY NF.EntradaSaida_Id, NF.Cliente_Id , NF.EndCliente_ID, NF.Serie_Id, NF.Nota_Id, NF.Movimento, NF.DataDaNota, NFxI.cfop_Id,  ISNULL(OE.STPISCOFINS, 0), NF.Empresa_Id " & vbCrLf &
              " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS CRED PRE' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf

        'Tirei o CFOP 1556,1653 - Furlan - 26-04-2017
        '"    AND NFxI.cfop_Id in (1556, 2556, 1653, 1949, 2653, 3101, 3556, 5949, 6949) And SubOperacoes.Classe <> '" & eClassesOperacoes.SERVICOS.ToString & "'" & vbCrLf & _

        '"  WHERE (ISNULL(NF.Situacao, 1) = 1) AND (NF.Empresa_Id = '" & .Item("Cliente_Id") & "') " & vbCrLf & _
        SqlT &= GetSqlRegistroF100_RazaoContabil()

        Return SqlT
    End Function

    Private Function GetSqlRegistroF100_RazaoContabil() As String
        Dim SqlT As String = String.Empty

        '------Razao Contabil
        SqlT = "Select 'F100' AS Reg, Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, CstPis, BaseDeCalculoPis, AliquotaPis, Pis, " & vbCrLf &
              "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta," & vbCrLf &
              "       Sum(Valor) as Valor " & vbCrLf &
              "   INTO #F100RazaoContabil  " & vbCrLf &
              " From ( " & vbCrLf &
              "       Select 'F100' AS Reg, 1 as Ind, Empresa_Id," & vbCrLf &
              "              Cliente_Id, EndCliente_Id," & vbCrLf &
              "              Month(Movimento_Id) as Mes, " & vbCrLf &
              "              year(Movimento_Id) as Ano, " & vbCrLf &
              "              Sum(DebitoOficial) as Valor, " & vbCrLf &
              "              01 as CstPis,  " & vbCrLf &
              "              Sum(DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
              "              1.65 as AliquotaPis, " & vbCrLf &
              "              ((Sum(DebitoOficial) * 1.65) / 100) as Pis, " & vbCrLf &
              "              01 as CstCofins, " & vbCrLf &
              "              Sum(DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
              "              7.6 as AliquotaCofins, " & vbCrLf &
              "              ((Sum(DebitoOficial) * 7.6) /100) as Cofins," & vbCrLf &
              "              13 as NatBaseCredito, " & vbCrLf &
              "              0 as IndOrigemCredito," & vbCrLf &
              "              Conta_Id as Conta" & vbCrLf &
              "         From Razao" & vbCrLf &
              "        Where left(Conta_Id, 7) in ('1010301')" & vbCrLf &
              "          AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "          And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
              "          And Lote_Id = 30" & vbCrLf &
              "          And not Lote_Id in (7500) " & vbCrLf &
              "        Group By Cliente_Id, EndCliente_Id, Month(Movimento_Id), year(Movimento_Id), Conta_Id, Empresa_Id" & vbCrLf &
              "        Union all" & vbCrLf &
              "       Select 'F100' AS Reg, 0 as Ind, Empresa_Id," & vbCrLf &
              "              Cliente_Id, EndCliente_Id," & vbCrLf &
              "              Month(Movimento_Id) as Mes, " & vbCrLf &
              "              year(Movimento_Id) as Ano, " & vbCrLf &
              "              Sum(DebitoOficial) as Valor, " & vbCrLf &
              "              56 as CstPis,  " & vbCrLf &
              "              Sum(DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
              "              1.65 as AliquotaPis, " & vbCrLf &
              "              ((Sum(DebitoOficial) * 1.65) / 100) as Pis, " & vbCrLf &
              "              56 as CstCofins, " & vbCrLf &
              "              Sum(DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
              "              7.6 as AliquotaCofins, " & vbCrLf &
              "              ((Sum(DebitoOficial) * 7.6) /100) as Cofins," & vbCrLf &
              "              5 as NatBaseCredito, " & vbCrLf &
              "              0 as IndOrigemCredito," & vbCrLf &
              "              Conta_Id as Conta" & vbCrLf &
              "         From Razao" & vbCrLf &
              "        Where left(Conta_Id, 7) in ('2010101')" & vbCrLf &
              "          AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
              "          And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') And Lote_Id = 30" & vbCrLf &
              "          And not Lote_Id in (7500) " & vbCrLf &
              "        Group By Cliente_Id, EndCliente_Id, Month(Movimento_Id), year(Movimento_Id), Conta_Id, Empresa_Id" & vbCrLf &
              "      ) Consulta " & vbCrLf &
              " Group By Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
              "          CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
              "          CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
              "          NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
              " Having(Sum(Valor) > 1)" & vbCrLf

        If (EmpresaMestre = "04376053") Or (EmpresaMestre = "04440724") Or (EmpresaMestre = "05954217") Or (EmpresaMestre = "06329316") Or (EmpresaMestre = "07090163") Or (EmpresaMestre = "07093052") And cboIncidencia.SelectedValue <> 2 Then
            SqlT = "SELECT 'F100' AS Reg, Ind, Empresa_Id,  Cliente_Id, EndCliente_Id, Mes, Ano, Sum(Valor) as Valor, CstPis, BaseDeCalculoPis, AliquotaPis, Pis, " & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   "  INTO #F100RazaoContabil " & vbCrLf &
                   "  FROM ( " & vbCrLf &
                   "         SELECT 'F100' AS Reg, 2 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                6 as CstPis,    0 as BaseDeCalculoPis,    0 as AliquotaPis, 0 as Pis, " & vbCrLf &
                   "                6 as CstCofins, 0 as BaseDeCalculoCofins, 0 as AliquotaCofins, 0 as Cofins," & vbCrLf &
                   "                2 as NatBaseCredito, 0 as IndOrigemCredito,  " & vbCrLf &
                   "                Left(Conta_Id, 7) as Conta" & vbCrLf &
                   "           FROM Razao" & vbCrLf &
                   "          WHERE (Left(Conta_Id, 7) in ('4020203', '4020204')) And not Conta_Id in ('403010105', '403010103')" & vbCrLf &
                   "            AND (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "            AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "            AND NOT Lote_Id in (7500) " & vbCrLf &
                   "          GROUP BY Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id), year(Movimento_Id), Conta_Id" & vbCrLf &
                   "          UNION " & vbCrLf &
                   "         SELECT 'F100' AS Reg, 1 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                01 as CstPis,  " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                1.65 as AliquotaPis, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 1.65) / 100) as Pis, " & vbCrLf &
                   "                01 as CstCofins, " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                7.6 as AliquotaCofins, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 7.6) /100) as Cofins," & vbCrLf &
                   "                5 as NatBaseCredito, " & vbCrLf &
                   "                0 as IndOrigemCredito," & vbCrLf &
                   "                Conta_Id as Conta" & vbCrLf &
                   "           FROM Razao" & vbCrLf &
                   "          WHERE (Left(Conta_Id, 7) in ('4030101'))  And not Conta_Id in ('403010105', '403010103')" & vbCrLf &
                   "            AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "            AND (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "            AND not Lote_Id in (7500) " & vbCrLf &
                   "          GROUP BY Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id), year(Movimento_Id), Conta_Id " & vbCrLf & vbCrLf &
                   "        ) Consulta " & vbCrLf &
                   " GROUP BY Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
                   "       CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
                   "       NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   "HAVING (SUM(Valor) > 1) " & vbCrLf
        End If

        'Panorama 
        If (EmpresaMestre = "03189063") And cboIncidencia.SelectedValue <> 2 Then
            SqlT = "SELECT 'F100' AS Reg, Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, Sum(Valor) as Valor,  CstPis,    BaseDeCalculoPis,   AliquotaPis, Pis, " & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta " & vbCrLf &
                   "  INTO #F100RazaoContabil " & vbCrLf &
                   "  FROM ( " & vbCrLf &
                   "         SELECT 'F100' AS Reg, 1 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                02 as CstPis,  " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                0.65 as AliquotaPis, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 0.65) / 100) as Pis, " & vbCrLf &
                   "                02 as CstCofins, " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                4.0 as AliquotaCofins, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 4.0) /100) as Cofins," & vbCrLf &
                   "                5 as NatBaseCredito, " & vbCrLf &
                   "                0 as IndOrigemCredito," & vbCrLf &
                   "                Conta_Id as Conta" & vbCrLf &
                   "                From Razao" & vbCrLf &
                   "                Where  (Left(Conta_Id, 7) in ('4020203')) And Conta_Id <> '402020303'" & vbCrLf &
                   "                AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "                And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "                And not Lote_Id in (7500) " & vbCrLf &
                   "          GROUP BY Empresa_Id, Month(Movimento_Id), Cliente_Id, EndCliente_Id, year(Movimento_Id), Conta_Id " & vbCrLf &
                   "        ) Consulta " & vbCrLf &
                   " GROUP BY Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
                   "       CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
                   "       NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   "Having (Sum(Valor) > 1) " & vbCrLf
        End If


        'Verde 
        If (EmpresaMestre = "44979506") And cboIncidencia.SelectedValue <> 2 Then
            SqlT = "SELECT 'F100' AS Reg, Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, Sum(Valor) as Valor,  CstPis,    BaseDeCalculoPis,   AliquotaPis, Pis, " & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta " & vbCrLf &
                   "  INTO #F100RazaoContabil " & vbCrLf &
                   "  FROM ( " & vbCrLf &
                   "         SELECT 'F100' AS Reg, 1 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                02 as CstPis,  " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                0.65 as AliquotaPis, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 0.65) / 100) as Pis, " & vbCrLf &
                   "                02 as CstCofins, " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                4.0 as AliquotaCofins, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 4.0) /100) as Cofins," & vbCrLf &
                   "                5 as NatBaseCredito, " & vbCrLf &
                   "                0 as IndOrigemCredito," & vbCrLf &
                   "                Conta_Id as Conta" & vbCrLf &
                   "                From Razao" & vbCrLf &
                   "                Where  (Left(Conta_Id, 7) in ('4020203')) And Conta_Id <> '402020303'" & vbCrLf &
                   "                AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "                And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "                And not Lote_Id in (9, 10, 7500) " & vbCrLf &
                   "          GROUP BY Empresa_Id, Month(Movimento_Id), Cliente_Id, EndCliente_Id, year(Movimento_Id), Conta_Id " & vbCrLf &
                   "        ) Consulta " & vbCrLf &
                   " GROUP BY Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
                   "       CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
                   "       NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   "Having (Sum(Valor) > 1) " & vbCrLf
        End If


        'Nutri
        'Jaber pediu para deixar igual do Curtume, qualquer problema rever a alteração - 16/03/2022
        ''"                Where  (Left(Conta_Id, 7) in ('4020203')) And Conta_Id <> '402020303' AND Conta_Id <> '402020304' AND Conta_Id <> '402020301'" & vbCrLf & _
        If ((EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "62747840") OrElse (EmpresaMestre = "62780383") OrElse (EmpresaMestre = "63358210")) And cboIncidencia.SelectedValue <> 2 Then
            SqlT = "SELECT 'F100' AS Reg, Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, Sum(Valor) as Valor,  CstPis,    BaseDeCalculoPis,   AliquotaPis, Pis, " & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta " & vbCrLf &
                   "  INTO #F100RazaoContabil " & vbCrLf &
                   "  FROM ( " & vbCrLf &
                   "         SELECT 'F100' AS Reg, 1 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                02 as CstPis,  " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                0.65 as AliquotaPis, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 0.65) / 100) as Pis, " & vbCrLf &
                   "                02 as CstCofins, " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                4.0 as AliquotaCofins, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 4.0) /100) as Cofins," & vbCrLf &
                   "                5 as NatBaseCredito, " & vbCrLf &
                   "                0 as IndOrigemCredito," & vbCrLf &
                   "                Conta_Id as Conta" & vbCrLf &
                   "                From Razao" & vbCrLf &
                   "                Where  (Left(Conta_Id, 7) in ('4020203')) And Conta_Id <> '402020303'" & vbCrLf &
                   "                AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "                And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "                And not Lote_Id in (7500) " & vbCrLf &
                   "          GROUP BY Empresa_Id, Month(Movimento_Id), Cliente_Id, EndCliente_Id, year(Movimento_Id), Conta_Id " & vbCrLf &
                   "        ) Consulta " & vbCrLf &
                   " GROUP BY Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
                   "       CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
                   "       NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   "Having (Sum(Valor) > 1) " & vbCrLf
        End If

        'Fronteira
        If (EmpresaMestre = "44005444") And cboIncidencia.SelectedValue <> 2 Then
            SqlT = "SELECT 'F100' AS Reg, Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, Sum(Valor) as Valor,  CstPis,    BaseDeCalculoPis,   AliquotaPis, Pis, " & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta " & vbCrLf &
                   "  INTO #F100RazaoContabil " & vbCrLf &
                   "  FROM ( " & vbCrLf &
                   "         SELECT 'F100' AS Reg, 1 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                02 as CstPis,  " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                0.65 as AliquotaPis, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 0.65) / 100) as Pis, " & vbCrLf &
                   "                02 as CstCofins, " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                4.0 as AliquotaCofins, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 4.0) /100) as Cofins," & vbCrLf &
                   "                5 as NatBaseCredito, " & vbCrLf &
                   "                0 as IndOrigemCredito," & vbCrLf &
                   "                Conta_Id as Conta" & vbCrLf &
                   "                From Razao" & vbCrLf &
                   "                Where  (Left(Conta_Id, 7) in ('4020203')) And Conta_Id <> '402020303'" & vbCrLf &
                   "                AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "                And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "                And not Lote_Id in (7500) " & vbCrLf &
                   "          GROUP BY Empresa_Id, Month(Movimento_Id), Cliente_Id, EndCliente_Id, year(Movimento_Id), Conta_Id " & vbCrLf &
                   "        ) Consulta " & vbCrLf &
                   " GROUP BY Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
                   "       CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
                   "       NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   "Having (Sum(Valor) > 1) " & vbCrLf
        End If

        'Quimica
        If (EmpresaMestre = "05272759") And cboIncidencia.SelectedValue <> 2 Then
            SqlT = "SELECT 'F100' AS Reg, Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, Sum(Valor) as Valor,  CstPis,    BaseDeCalculoPis,   AliquotaPis, Pis, " & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta " & vbCrLf &
                   "  INTO #F100RazaoContabil " & vbCrLf &
                   "  FROM ( " & vbCrLf &
                   "         SELECT 'F100' AS Reg, 1 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                02 as CstPis,  " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                0.65 as AliquotaPis, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 0.65) / 100) as Pis, " & vbCrLf &
                   "                02 as CstCofins, " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                4.0 as AliquotaCofins, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 4.0) /100) as Cofins," & vbCrLf &
                   "                5 as NatBaseCredito, " & vbCrLf &
                   "                0 as IndOrigemCredito," & vbCrLf &
                   "                Conta_Id as Conta" & vbCrLf &
                   "                From Razao" & vbCrLf &
                   "                Where  Conta_Id in ('402020301','402020302','402020304','402020401','402020402','402020399') " & vbCrLf &
                   "                AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "                And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "                And not Lote_Id in (7500) " & vbCrLf &
                   "          GROUP BY Empresa_Id, Month(Movimento_Id), Cliente_Id, EndCliente_Id, year(Movimento_Id), Conta_Id " & vbCrLf &
                   "        ) Consulta " & vbCrLf &
                   " GROUP BY Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
                   "       CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
                   "       NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   "Having (Sum(Valor) > 1) " & vbCrLf
        End If

        '----Alvorada-------------------

        If (EmpresaMestre = "04854422") And cboIncidencia.SelectedValue <> 2 Then
            '303030301 JUROS(AUFERIDOS / RECEBIDOS)
            '303030302 VARIACAO MONETARIA ATIVA
            '303030303 DESCONTOS(OBTIDOS)
            '303030304 JUROS AUFERIDOS - EMPRESTIMOS FORNECEDOR
            '303030305 RENDAS EM APLIC. FINANCEIRAS
            ' + CRÉDITOS DO lOTE 30 

            SqlT = "SELECT 'F100' AS Reg, Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, Sum(Valor) as Valor, CstPis, BaseDeCalculoPis, AliquotaPis, Pis, " & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta " & vbCrLf &
                   "  INTO #F100RazaoContabil " & vbCrLf &
                   "  FROM ( " & vbCrLf &
                   "         SELECT 'F100' AS Reg, 1 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                 02 as CstPis,  " & vbCrLf &
                   "                 Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                 0.65 as AliquotaPis, " & vbCrLf &
                   "                 ((Sum(CreditoOficial - DebitoOficial) * 0.65) / 100) as Pis, " & vbCrLf &
                   "                 02 as CstCofins, " & vbCrLf &
                   "                 Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                 4.0 as AliquotaCofins, " & vbCrLf &
                   "                 ((Sum(CreditoOficial - DebitoOficial) * 4.0) /100) as Cofins," & vbCrLf &
                   "                 0 as NatBaseCredito, " & vbCrLf &
                   "                 0 as IndOrigemCredito," & vbCrLf &
                   "                 Conta_Id as Conta" & vbCrLf &
                   "            FROM Razao" & vbCrLf &
                   "           WHERE Conta_Id in ('303030301', '303030302', '303030303','303030304','303030305' " & IIf(CDate(PeriodoFinal).Year >= 2016 AndAlso CDate(PeriodoFinal).Month >= 4, ",'303030312'", "") & " )  " & vbCrLf &
                   "             AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "             AND (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "             AND not Lote_Id in (7500) " & vbCrLf &
                   "           GROUP By Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id), year(Movimento_Id), Conta_Id " & vbCrLf &
                   "           UNION " & vbCrLf &
                   "          SELECT 'F100' AS Reg, 0 as Ind, Empresa_Id," & vbCrLf &
                   "                 Cliente_Id, EndCliente_Id," & vbCrLf &
                   "                 Month(Movimento_Id) as Mes, " & vbCrLf &
                   "                 year(Movimento_Id) as Ano, " & vbCrLf &
                   "                 Sum(CreditoOficial) as Valor, " & vbCrLf &
                   "                 56 as CstPis,  " & vbCrLf &
                   "                 Sum(CreditoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                 1.65 as AliquotaPis, " & vbCrLf &
                   "                 ((Sum(CreditoOficial) * 1.65) / 100) as Pis, " & vbCrLf &
                   "                 56 as CstCofins, " & vbCrLf &
                   "                 Sum(CreditoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                 7.6 as AliquotaCofins, " & vbCrLf &
                   "                 ((Sum(CreditoOficial) * 7.6) /100) as Cofins," & vbCrLf &
                   "                 5 as NatBaseCredito, " & vbCrLf &
                   "                 0 as IndOrigemCredito," & vbCrLf &
                   "                 Conta_Id as Conta" & vbCrLf &
                   "                 From Razao" & vbCrLf &
                   "                 Where  left(Conta_Id, 7) in ('2010101')" & vbCrLf &
                   "                 AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "                 And (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') And Lote_Id = 30" & vbCrLf &
                   "                 And not Lote_Id in (7500) " & vbCrLf &
                   "                 Group By Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id), year(Movimento_Id), Conta_Id" & vbCrLf &
                   "       ) Consulta " & vbCrLf &
                   "  GROUP BY Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
                   "        CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
                   "        CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
                   "        NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   " HAVING (Sum(Valor) > 1)" & vbCrLf
        End If

        '----Fex-------------------

        If (EmpresaMestre = "15204808") And cboIncidencia.SelectedValue <> 2 Then
            '303030301
            '303030303
            '303030305
            '+ CRÉDITOS DO LOTE 30

            SqlT = "Select 'F100' AS Reg, Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, Sum(Valor) as Valor,  CstPis,    BaseDeCalculoPis,   AliquotaPis, Pis, " & vbCrLf &
                   "        CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins, NatBaseCredito, IndOrigemCredito, Conta " & vbCrLf &
                   "  INTO #F100RazaoContabil " & vbCrLf &
                   "  FROM ( " & vbCrLf &
                   "        SELECT 'F100' AS Reg, 1 as Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id) as Mes, year(Movimento_Id) as Ano, Sum(CreditoOficial - DebitoOficial) as Valor,  " & vbCrLf &
                   "                02 as CstPis,  " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoPis, " & vbCrLf &
                   "                0.65 as AliquotaPis, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 0.65) / 100) as Pis, " & vbCrLf &
                   "                02 as CstCofins, " & vbCrLf &
                   "                Sum(CreditoOficial - DebitoOficial) as BaseDeCalculoCofins, " & vbCrLf &
                   "                4.0 as AliquotaCofins, " & vbCrLf &
                   "                ((Sum(CreditoOficial - DebitoOficial) * 4.0) /100) as Cofins," & vbCrLf &
                   "                0 as NatBaseCredito, " & vbCrLf &
                   "                0 as IndOrigemCredito," & vbCrLf &
                   "                Conta_Id as Conta" & vbCrLf &
                   "           FROM Razao" & vbCrLf &
                   "          WHERE Conta_Id in ('303030301', '303030303', '303030305')" & vbCrLf &
                   "            AND (left(Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "            AND (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "            AND not Lote_Id in (7500) " & vbCrLf &
                   "          GROUP By Empresa_Id, Cliente_Id, EndCliente_Id, Month(Movimento_Id), year(Movimento_Id), Conta_Id" & vbCrLf &
                   "        ) Consulta " & vbCrLf &
                   " GROUP BY Ind, Empresa_Id, Cliente_Id, EndCliente_Id, Mes, Ano, " & vbCrLf &
                   "       CstPis, BaseDeCalculoPis, AliquotaPis, Pis," & vbCrLf &
                   "       CstCofins, BaseDeCalculoCofins, AliquotaCofins, Cofins," & vbCrLf &
                   "       NatBaseCredito, IndOrigemCredito, Conta" & vbCrLf &
                   "HAVING (SUM(Valor) > 1)" & vbCrLf
        End If

        Return SqlT
    End Function

    Private Function GetSqlRegistroF120() As String
        Dim SqlT As String = String.Empty

        If cboIncidencia.SelectedValue <> 2 Then
            SqlT = " SELECT Empresa_Id, Movimento_Id, Nat_Bc_Cred, Ident_Bem_Imob, Ind_Orig_Cred, Ind_Util_Bem_Imob, Vl_Oper_Dep, Parc_Oper_Nao_Bc_Cred, Cst_Pis, Vl_Bc_Pis, " & vbCrLf &
                   "        Aliq_Pis,Vl_Pis, Cst_Cofins, Vl_Bc_Cofins, Aliq_Cofins, Vl_Cofins, Cod_Cta, Cod_Ccus, Desc_Bem_Imob" & vbCrLf &
                   "   INTO #F120  " & vbCrLf &
                   "   FROM PisCofins_Reg_F120" & vbCrLf &
                   "  WHERE (Movimento_Id between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')"
        Else
            SqlT = " SELECT 1 AS Cod " & vbCrLf &
                   "   INTO #F120  " & vbCrLf &
                   "  WHERE 1=2 " & vbCrLf
        End If
        Return SqlT
    End Function

    Private Function GetSqlRegistroF130() As String
        Dim SqlT As String = String.Empty

        If cboIncidencia.SelectedValue <> 2 Then
            SqlT = " SELECT Empresa_Id, Movimento_Id, Nat_Bc_Cred, Ident_Bem_Imob, Ind_Orig_Cred, Ind_Util_Bem_Imob, Mes_Ope_Aquis,Vl_Oper_Aquis, Parc_Oper_Nao_Bc_Cred, Vl_Bc_Cred, " & vbCrLf &
                   "        Ind_Nr_Parc, Cst_Pis, Vl_Bc_Pis, Aliq_Pis, Vl_Pis, Cst_Cofins, Vl_Bc_Cofins, Aliq_Cofins, Vl_Cofins, Cod_Cta, Cod_Ccus, Desc_Bem_Imob" & vbCrLf &
                   "   INTO #F130  " & vbCrLf &
                   "   FROM PisCofins_Reg_F130 " & vbCrLf &
                   "  WHERE (Movimento_Id BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf
        Else
            SqlT = " SELECT 1 AS Cod " & vbCrLf &
                   "   INTO #F130  " & vbCrLf &
                   "  WHERE 1=2 " & vbCrLf
        End If

        Return SqlT

    End Function

    Private Function GetSqlRegistroF550() As String
        Dim SqlT As String = String.Empty

        If cboIncidencia.SelectedValue = 2 Then
            'SqlT = " SELECT 'F550' AS Reg," & vbCrLf & _
            '       "        NF.Empresa_Id, " & vbCrLf & _
            '       "        Cli.Cliente_Id AS Cod_Part," & vbCrLf & _
            '       "        Cli.Endereco_Id, " & vbCrLf & _
            '       "        Case when NF.EntradaSaida_Id = 'S' THEN 1 ELSE 0 End AS Cod_Mod," & vbCrLf & _
            '       "        ISNULL(OE.STPISCOFINS, 0)  AS Cod_Sit," & vbCrLf & _
            '       "        NFxI.cfop_Id as CFOP," & vbCrLf & _
            '       "        NF.Serie_Id AS Ser," & vbCrLf & _
            '       "        NF.Nota_Id AS Num_Doc," & vbCrLf & _
            '       "        NF.Movimento AS DT_Doc," & vbCrLf & _
            '       "        NF.DataDaNota AS DT_Ent," & vbCrLf & _
            '       "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxE.Valor      END, 0)) AS Vl_Doc," & vbCrLf & _
            '       "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS'    THEN NFxE.Valor      END, 0)) AS Vl_Icms," & vbCrLf & _
            '       "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS'     THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Base_Pis," & vbCrLf & _
            '       "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS'     THEN NFxE.Percentual END, 0)) AS Per_Pis," & vbCrLf & _
            '       "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS'     THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_Pis," & vbCrLf & _
            '       "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS'  THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Base_Cofins," & vbCrLf & _
            '       "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS'  THEN NFxE.Percentual END, 0)) AS Per_Cofins," & vbCrLf & _
            '       "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS'  THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Valor_Cofins" & vbCrLf & _
            '       "   INTO #F550  " & vbCrLf & _
            '       "   FROM NotasFiscais NF " & vbCrLf & _
            '       "  INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
            '       "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
            '       "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
            '       "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
            '       "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
            '       "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
            '       "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
            '       "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
            '       "  INNER JOIN NotasFiscaisXEncargos NFxE" & vbCrLf & _
            '       "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf & _
            '       "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf & _
            '       "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf & _
            '       "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf & _
            '       "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf & _
            '       "    AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf & _
            '       "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf & _
            '       "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf & _
            '       "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf & _
            '       "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf & _
            '       "  INNER JOIN Clientes Cli" & vbCrLf & _
            '       "     ON NF.Cliente_Id    = Cli.Cliente_Id " & vbCrLf & _
            '       "    AND NF.EndCliente_Id = Cli.Endereco_Id" & vbCrLf & _
            '       "  inner join OperacaoXEstado OE" & vbCrLf & _
            '       "     on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf & _
            '       "  WHERE (ISNULL(NF.Situacao, 1) = 1)" & vbCrLf & _
            '       "    AND (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf & _
            '       "    AND NF.EntradaSaida_Id = 'S'" & vbCrLf & _
            '       "  GROUP BY NF.Empresa_Id, NF.EntradaSaida_Id, Cli.Cliente_Id, Cli.Endereco_Id, NF.Serie_Id, NF.Nota_Id, NF.Movimento, NF.DataDaNota, NFxI.cfop_Id, ISNULL(OE.STPISCOFINS, 0)" & vbCrLf & _
            '       " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf

            SqlT = " SELECT 'F550' AS Reg," & vbCrLf &
                   "        NF.Empresa_Id, " & vbCrLf &
                   "        Case when NF.EntradaSaida_Id = 'S' THEN 1 ELSE 0 End AS Cod_Mod," & vbCrLf &
                   "        ISNULL(OE.STPISCOFINS, 0)  AS Cod_Sit," & vbCrLf &
                   "        NF.TipoDeDocumento," & vbCrLf &
                   "        NF.Eletronica," & vbCrLf &
                   "        NFxI.cfop_Id as CFOP," & vbCrLf &
                   "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxE.Valor      END, 0)) AS Vl_Doc," & vbCrLf &
                   "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'ICMS'    THEN NFxE.Valor      END, 0)) AS Vl_Icms," & vbCrLf &
                   "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS'     THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Base_Pis," & vbCrLf &
                   "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS'     THEN NFxE.Percentual END, 0)) AS Per_Pis," & vbCrLf &
                   "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS'     THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Vl_Pis," & vbCrLf &
                   "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS'  THEN CASE WHEN isnull(NFxE.BaseNova,0) > 0 then NFxE.BaseNova else NFxE.Base END END, 0)) AS Base_Cofins," & vbCrLf &
                   "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS'  THEN NFxE.Percentual END, 0)) AS Per_Cofins," & vbCrLf &
                   "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'COFINS'  THEN CASE WHEN isnull(NFxE.ValorNovo,0) > 0 then NFxE.ValorNovo else NFxE.Valor END END, 0)) AS Valor_Cofins" & vbCrLf &
                   "   INTO #F550  " & vbCrLf &
                   "   FROM NotasFiscais NF " & vbCrLf &
                   "  INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf &
                   "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                   "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                   "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                   "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                   "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                   "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                   "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
                   "  INNER JOIN NotasFiscaisXEncargos NFxE" & vbCrLf &
                   "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                   "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                   "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
                   "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                   "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
                   "    AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
                   "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
                   "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
                   "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
                   "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                   "  INNER JOIN Clientes Cli" & vbCrLf &
                   "     ON NF.Cliente_Id    = Cli.Cliente_Id " & vbCrLf &
                   "    AND NF.EndCliente_Id = Cli.Endereco_Id" & vbCrLf &
                   "  inner join OperacaoXEstado OE" & vbCrLf &
                   "     on OE.Codigo_id = NFxI.OperacaoxEstado" & vbCrLf &
                   "  WHERE (ISNULL(NF.Situacao, 1) = 1)" & vbCrLf &
                   "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
                   "    AND (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
                   "    AND NF.EntradaSaida_Id = 'S'" & vbCrLf &
                   "  GROUP BY NF.Empresa_Id, NF.TipoDeDocumento, NF.Eletronica, NF.EntradaSaida_Id, NFxI.cfop_Id, ISNULL(OE.STPISCOFINS, 0)" & vbCrLf
            '" HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf
        Else
            SqlT = " SELECT 1 AS Cod " & vbCrLf &
                   "   INTO #F550 " & vbCrLf &
                   "  WHERE 1=2 " & vbCrLf
        End If

        Return SqlT
    End Function

    Private Function GetSqlRegistroF600() As String
        Dim SqlT As String = String.Empty

        SqlT = " SELECT NF.Empresa_Id, " & vbCrLf &
               "        NF.Cliente_Id, " & vbCrLf &
               "        NF.EndCliente_Id," & vbCrLf &
               "        NF.DataDaNota," & vbCrLf &
               "        NF.Movimento," & vbCrLf &
               "        SUM(CASE WHEN NFxE.Encargo_Id = 'PIS_Ret'    THEN NFxE.Base  ELSE 0 END) AS BasePis," & vbCrLf &
               "        SUM(CASE WHEN NFxE.Encargo_Id = 'PIS_Ret'    THEN NFxE.Valor ELSE 0 END) AS ValorPis," & vbCrLf &
               "        SUM(CASE WHEN NFxE.Encargo_Id = 'COFINS_Ret' THEN NFxE.Base  ELSE 0 END) AS BaseCofins," & vbCrLf &
               "        SUM(CASE WHEN NFxE.Encargo_Id = 'COFINS_Ret' THEN NFxE.Valor ELSE 0 END) AS ValorCofins" & vbCrLf &
               "   INTO #F600 " & vbCrLf &
               "   FROM NotasFiscais AS NF" & vbCrLf &
               "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
               "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
               "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
               "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
               "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
               "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
               "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
               "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
               "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
               "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
               "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
               "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
               "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
               "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
               "    AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
               "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
               "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
               "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
               "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
               "  INNER JOIN SubOperacoes AS Sop" & vbCrLf &
               "     ON NFxI.Operacao    = Sop.Operacao_Id" & vbCrLf &
               "    AND NFxI.SubOperacao = Sop.SubOperacoes_Id" & vbCrLf &
               "  WHERE isnull(NF.Situacao,1) = 1" & vbCrLf &
               "    AND (left(NF.Empresa_Id, 8) = '" & EmpresaMestre & "')" & vbCrLf &
               "    And (NF.Movimento  BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "')" & vbCrLf &
               "    AND (Sop.Classe = '" & eClassesOperacoes.SERVICOS.ToString & "')" & vbCrLf &
               "  GROUP BY NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.DataDaNota, NF.Movimento" & vbCrLf &
               " HAVING (SUM(CASE WHEN NFxE.Encargo_Id = 'PIS_Ret' THEN NFxE.Valor ELSE 0 END) > 0)" & vbCrLf
        Return SqlT
    End Function

    Private Sub CompoeRegistroF100(ByRef ds As DataSet, ByRef Arq As ArrayList, ByRef RegF100 As Integer, ByRef RegGeral As Integer, F010Empresa As String)

        Dim linhaTemp As String = String.Empty

        For Each drF100 In ds.Tables("F100").Select("Empresa_Id = " & F010Empresa)
            With drF100

                linhaTemp = "|" & .Item("Reg")
                linhaTemp &= "|" & .Item("Cod_Mod")
                linhaTemp &= "|" & .Item("Cod_Part") & .Item("EndCliente_Id")
                linhaTemp &= "|" & ""
                linhaTemp &= "|" & Format(CDate(.Item("DT_DOC")), "ddMMyyyy")
                linhaTemp &= "|" & .Item("Vl_Doc")
                linhaTemp &= "|" & Format(CInt(.Item("Cod_Sit")), "00")
                linhaTemp &= "|" & .Item("Base_Pis")
                linhaTemp &= "|" & CDbl(.Item("Per_Pis")).ToString("N4")
                linhaTemp &= "|" & .Item("Vl_Pis")
                linhaTemp &= "|" & Format(CInt(.Item("Cod_Sit")), "00")
                linhaTemp &= "|" & .Item("Base_Cofins")
                linhaTemp &= "|" & CDbl(.Item("Per_Cofins")).ToString("N4")
                linhaTemp &= "|" & .Item("Valor_Cofins")

                If .Item("Cfop") = 1949 Then
                    linhaTemp &= "|13"
                Else
                    linhaTemp &= "|02"
                End If

                If .Item("Cfop") > 3000 And .Item("Cfop") < 4000 Then
                    linhaTemp &= "|" & "1"
                Else
                    linhaTemp &= "|" & "0"
                End If

                linhaTemp &= "|" & .Item("Conta")
                linhaTemp &= "|" '& ""
                linhaTemp &= "|" '& rsF100("Descricao")
                linhaTemp &= "|"

                Arq.Add(linhaTemp)
                RegF100 += 1
                RegGeral += 1
            End With
        Next



        For Each drF100 In ds.Tables("F100RazaoContabil").Select("Empresa_Id = " & F010Empresa)
            With drF100

                If .Item("Conta") = "1010301" Or .Item("Conta") = "2010101" Then
                    linhaTemp = "|" & .Item("Reg")
                    linhaTemp &= "|" & .Item("Ind")
                    linhaTemp &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                    linhaTemp &= "|"

                    'linhaTemp &= "|01" & Format(CInt(.Item("Mes")), "00") & .Item("Ano")
                    linhaTemp &= "|" & DateValue(txtDataFinal.Text).ToString("dd") & Format(CInt(.Item("Mes")), "00") & .Item("Ano")

                    linhaTemp &= "|" & .Item("Valor")
                    linhaTemp &= "|" & Format(CInt(.Item("CstPis")), "00")
                    linhaTemp &= "|" & .Item("BaseDeCalculoPis")
                    linhaTemp &= "|" & CDbl(.Item("AliquotaPis")).ToString("N2")
                    linhaTemp &= "|" & Replace(CDbl(.Item("Pis")).ToString("N2"), ".", "")
                    linhaTemp &= "|" & Format(CInt(.Item("CstCofins")), "00")
                    linhaTemp &= "|" & Replace(CDbl(.Item("BaseDeCalculoCofins")).ToString("N2"), ".", "")
                    linhaTemp &= "|" & CDbl(.Item("AliquotaCofins")).ToString("N2")
                    linhaTemp &= "|" & Replace(CDbl(.Item("Cofins")).ToString("N2"), ".", "")

                    If .Item("NatBaseCredito") = 0 OrElse ((EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "44005444") OrElse (EmpresaMestre = "62747840") OrElse (EmpresaMestre = "62780383") OrElse (EmpresaMestre = "63358210")) Then
                        linhaTemp &= "|"
                    Else
                        linhaTemp &= "|" & Format(CInt(.Item("NatBaseCredito")), "00")
                    End If

                    linhaTemp &= "|" & .Item("IndOrigemCredito")
                    linhaTemp &= "|" & .Item("Conta")
                    linhaTemp &= "|"
                    linhaTemp &= "|"
                    linhaTemp &= "|"
                Else
                    linhaTemp = "|" & .Item("Reg")
                    linhaTemp &= "|" & .Item("Ind")
                    linhaTemp &= "|"
                    linhaTemp &= "|"

                    'linhaTemp &= "|01" & Format(CInt(.Item("Mes")), "00") & .Item("Ano")
                    linhaTemp &= "|" & DateValue(txtDataFinal.Text).ToString("dd") & Format(CInt(.Item("Mes")), "00") & .Item("Ano")

                    linhaTemp &= "|" & .Item("Valor")
                    linhaTemp &= "|" & Format(CInt(.Item("CstPis")), "00")
                    linhaTemp &= "|" & .Item("BaseDeCalculoPis")
                    linhaTemp &= "|" & CDbl(.Item("AliquotaPis")).ToString("N2")
                    linhaTemp &= "|" & Replace(CDbl(.Item("Pis")).ToString("N2"), ".", "")
                    linhaTemp &= "|" & Format(CInt(.Item("CstCofins")), "00")
                    linhaTemp &= "|" & Replace(CDbl(.Item("BaseDeCalculoCofins")).ToString("N2"), ".", "")
                    linhaTemp &= "|" & CDbl(.Item("AliquotaCofins")).ToString("N2")
                    linhaTemp &= "|" & Replace(CDbl(.Item("Cofins")).ToString("N2"), ".", "")

                    If .Item("NatBaseCredito") = 0 OrElse ((EmpresaMestre = "05366261") OrElse (EmpresaMestre = "38198213") OrElse (EmpresaMestre = "40938762") OrElse (EmpresaMestre = "44005444") OrElse (EmpresaMestre = "62747840") OrElse (EmpresaMestre = "62780383") OrElse (EmpresaMestre = "63358210")) Then
                        linhaTemp &= "|"
                    Else
                        linhaTemp &= "|" & Format(CInt(.Item("NatBaseCredito")), "00")
                    End If

                    If .Item("IndOrigemCredito") = 0 Then
                        linhaTemp &= "|"
                    Else
                        linhaTemp &= "|" & .Item("IndOrigemCredito")
                    End If

                    linhaTemp &= "|" & .Item("Conta")
                    linhaTemp &= "|"
                    linhaTemp &= "|"
                    linhaTemp &= "|"
                End If

                ArquivoAux.Add(linhaTemp)
                RegistroF100 = RegistroF100 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next

    End Sub

    Private Sub CompoeRegistroF120(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef RegistroF120 As Integer, ByRef RegistroGeral As Integer, F010Empresa As String)

        Dim linhaTemp As String = String.Empty

        ''Registro F120

        For Each drF120 In ds.Tables("F120").Select("Empresa_Id = " & F010Empresa)
            linhaTemp = "|" & "F120"
            linhaTemp &= "|" & Format(CInt(drF120("Nat_Bc_Cred")), "00")
            linhaTemp &= "|" & Format(CInt(drF120("Ident_Bem_Imob")), "00")
            linhaTemp &= "|" & drF120("Ind_Orig_Cred")
            linhaTemp &= "|" & drF120("Ind_Util_Bem_Imob")
            linhaTemp &= "|" & drF120("Vl_Oper_Dep")
            linhaTemp &= "|" & drF120("Parc_Oper_Nao_Bc_Cred")
            linhaTemp &= "|" & drF120("Cst_Pis")
            linhaTemp &= "|" & drF120("Vl_Bc_Pis")
            linhaTemp &= "|" & drF120("Aliq_Pis")
            linhaTemp &= "|" & drF120("Vl_Pis")
            linhaTemp &= "|" & drF120("Cst_Cofins")
            linhaTemp &= "|" & drF120("Vl_Bc_Cofins")
            linhaTemp &= "|" & drF120("Aliq_Cofins")
            linhaTemp &= "|" & drF120("Vl_Cofins")
            linhaTemp &= "|" & drF120("Cod_Cta")
            linhaTemp &= "|" & drF120("Cod_Ccus")
            linhaTemp &= "|" & drF120("Desc_Bem_Imob")
            linhaTemp &= "|"

            ArquivoAux.Add(linhaTemp)
            RegistroF120 += 1
            RegistroGeral += 1
        Next
    End Sub

    Private Sub CompoeRegistroF130(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef RegistroF130 As Integer, ByRef RegistroGeral As Integer, F010Empresa As String)

        Dim linhaTemp As String = String.Empty

        ''Registro F130

        For Each drF130 In ds.Tables("F130").Select("Empresa_Id = " & F010Empresa)
            With drF130
                linhaTemp = "|" & "F130"
                linhaTemp &= "|" & .Item("Nat_Bc_Cred")
                linhaTemp &= "|" & Format(.Item("Ident_Bem_Imob"), "00")
                linhaTemp &= "|" & .Item("Ind_Orig_Cred")
                linhaTemp &= "|" & .Item("Ind_Util_Bem_Imob")
                linhaTemp &= "|" & .Item("Mes_Ope_Aquis")
                linhaTemp &= "|" & .Item("Vl_Oper_Aquis")
                linhaTemp &= "|" & .Item("Parc_Oper_Nao_Bc_Cred")
                linhaTemp &= "|" & .Item("Vl_Bc_Cred")
                linhaTemp &= "|" & .Item("Ind_Nr_Parc")
                linhaTemp &= "|" & .Item("Cst_Pis")
                linhaTemp &= "|" & .Item("Vl_Bc_Pis")
                linhaTemp &= "|" & .Item("Aliq_Pis")
                linhaTemp &= "|" & .Item("Vl_Pis")
                linhaTemp &= "|" & .Item("Cst_Cofins")
                linhaTemp &= "|" & .Item("Vl_Bc_Cofins")
                linhaTemp &= "|" & .Item("Aliq_Cofins")
                linhaTemp &= "|" & .Item("Vl_Cofins")
                linhaTemp &= "|" & .Item("Cod_Cta")
                linhaTemp &= "|" & .Item("Cod_Ccus")
                linhaTemp &= "|" & .Item("Desc_Bem_Imob")
                linhaTemp &= "|"

                ArquivoAux.Add(linhaTemp)
                RegistroF130 = RegistroF130 + 1
                RegistroGeral = RegistroGeral + 1
            End With
        Next
    End Sub

    Private Sub CompoeRegistroF550(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef RegistroF550 As Integer, ByRef RegistroGeral As Integer, F010Empresa As String)

        Dim linhaTemp As String = String.Empty

        ''Registro F550

        For Each drF550 In ds.Tables("F550").Select("Empresa_Id= " & F010Empresa)
            With drF550
                linha = "|" & .Item("Reg")                           '01 - Texto fixo contendo "F550"
                linha &= "|" & .Item("Vl_Doc")                       '02 - Valor total da receita auferida, referente à combinação de CST e Alíquota.
                linha &= "|" & Format(CInt(drF550("Cod_Sit")), "00") '03 - Código da Situação Tributária referente ao PIS/PASEP
                linha &= "|"                                         '04 - Valor do desconto / exclusão da base de cálculo
                linha &= "|" & .Item("Base_Cofins")                  '05 - Valor da base de cálculo do PIS/PASEP
                linha &= "|" & FormatNumber(.Item("Per_Pis"), 2, TriState.UseDefault, TriState.UseDefault, TriState.False) '06 - Alíquota do PIS/PASEP (em percentual)
                linha &= "|" & .Item("Vl_Pis")                       '07 - Valor do PIS/PASEP
                linha &= "|" & Format(CInt(drF550("Cod_Sit")), "00") '08 - Código da Situação Tributária referente a COFINS
                linha &= "|"                                         '09 - Valor do desconto / exclusão da base de cálculo
                linha &= "|" & .Item("Base_Cofins")                  '10 - Valor da base de cálculo da COFINS
                linha &= "|" & FormatNumber(.Item("Per_Cofins"), 2, TriState.UseDefault, TriState.UseDefault, TriState.False) '11 - Alíquota da COFINS (em percentual)
                linha &= "|" & .Item("Valor_Cofins")                 '12 - Valor da COFINS

                If .Item("TipoDeDocumento") = 1 Then                 '13 - Código do modelo do documento fiscal conforme a Tabela 4.1.1
                    If .Item("Eletronica") = "S" Then
                        linha &= "|55"                'Nota Fiscal Eletrônica (NF-e)
                    Else
                        linha &= "|01"                'Nota Fiscal - 
                    End If
                ElseIf .Item("TipoDeDocumento") = 2 Then
                    If .Item("Eletronica") = "S" Then
                        linha &= "|57"                'Conhecimento de Transporte Eletrônico (CT-e)
                    Else
                        linha &= "|08"                'Conhecimento de Transporte Rodoviário de Cargas
                    End If
                ElseIf .Item("TipoDeDocumento") = 58 Then
                    If .Item("Eletronica") = "S" Then
                        linha &= "|57"                'Conhecimento de Transporte Eletrônico (CT-e)
                    Else
                        linha &= "|08"                'Conhecimento de Transporte Rodoviário de Cargas
                    End If
                Else
                    linha &= "|01"
                End If

                linha &= "|" & .Item("CFOP")                         '14 - Código fiscal de operação e prestação
                linha &= "|301010101"                                '15 - Código da conta analítica contábil debitada / creditada
                linha &= "|" '& .Item("Num_Doc")                     '16 - Informação complementar
                linha &= "|"

                ArquivoAux.Add(linha)
                RegistroF550 = RegistroF550 + 1
                RegistroGeral = RegistroGeral + 1

            End With
        Next

    End Sub

    Private Sub CompoeRegistroF600(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef RegistroF600 As Integer, ByRef RegistroGeral As Integer, F010Empresa As String)

        Dim linhaTemp As String = String.Empty

        ''Registro F600

        For Each drF600 In ds.Tables("F600").Select("Empresa_Id = " & F010Empresa)
            With drF600
                linhaTemp = "|" & "F600"
                linhaTemp &= "|" & "02"
                linhaTemp &= "|" & Format(CDate(.Item("DataDaNota")), "ddMMyyyy")
                linhaTemp &= "|" & .Item("BasePis")
                linhaTemp &= "|" & .Item("ValorPis") + .Item("ValorCofins")
                linhaTemp &= "|" & ""
                linhaTemp &= "|" & ""
                linhaTemp &= "|" & .Item("Cliente_Id")
                linhaTemp &= "|" & .Item("ValorPis")
                linhaTemp &= "|" & .Item("ValorCofins")
                linhaTemp &= "|" & "0"
                linhaTemp &= "|"

                ArquivoAux.Add(linhaTemp)
                RegistroF600 += 1
                RegistroGeral += 1

            End With
        Next
    End Sub
#End Region

#Region "BLOCO 1"

    Private Function ConsultaRegistro1001() As DataSet
        Return Banco.ConsultaDataSet(GetSqlRegistro1001(), "1001")
    End Function

    Private Function GetSqlRegistro1001() As String

        Dim SqlT As String = String.Empty

        SqlT = " SELECT DISTINCT '1001' AS Reg, Empresa_Id" & vbCrLf &
                      "   FROM PisCofins_Reg_1100 " & vbCrLf &
                      "  WHERE Empresa_Id LIKE '" & EmpresaMestre & "%'" & vbCrLf &
                      "    AND Ano           = " & Format(CDate(txtDataInicial.Text), "yyyy") & vbCrLf &
                      "    AND Mes           = " & Format(CDate(txtDataInicial.Text), "MM") & vbCrLf

        If cboIncidencia.SelectedValue = 2 Then

            SqlT &= " UNION " & vbCrLf &
                   " SELECT '1001' AS Reg," & vbCrLf &
                   "         NF.Empresa_Id" & vbCrLf &
                   "   FROM NotasFiscais AS NF" & vbCrLf &
                   "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
                   "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                   "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                   "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                   "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
                   "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                   "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                   "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                   "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
                   "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                   "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                   "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf &
                   "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                   "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id " & vbCrLf &
                   "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
                   "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
                   "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
                   "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
                   "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                   "  WHERE NF.Empresa_Id LIKE '" & EmpresaMestre & "%'" & vbCrLf &
                   "    And (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
                   "    And (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') " & vbCrLf &
                   "  GROUP BY NF.Empresa_ID" & vbCrLf &
                   " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf
        End If

        Return SqlT
    End Function

    Private Function Consulta_1100_1300_1500_1700_1900() As DataSet
        Dim SqlT As String = String.Empty

        '******************************************************************************************************************************************************************************
        '*********************************************************************** 1100 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT = GetSqlRegistro1100()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** 1300 ********************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistro1300()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** 1500 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistro1500()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** 1700 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistro1700()

        '******************************************************************************************************************************************************************************
        '*********************************************************************** 1900 *************************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= GetSqlRegistro1900()

        '******************************************************************************************************************************************************************************
        '************************************************************************* RESULTADO ******************************************************************************************
        '******************************************************************************************************************************************************************************

        SqlT &= " SELECT * FROM #1100 " & vbCrLf &
                " SELECT * FROM #1300 " & vbCrLf &
                " SELECT * FROM #1500 " & vbCrLf &
                " SELECT * FROM #1700 " & vbCrLf &
                " SELECT * FROM #1900 " & vbCrLf

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(SqlT, "Consulta")
        ds.Tables(0).TableName = "1100"
        ds.Tables(1).TableName = "1300"
        ds.Tables(2).TableName = "1500"
        ds.Tables(3).TableName = "1700"
        ds.Tables(4).TableName = "1900"

        Return ds
    End Function

    Private Function GetSqlRegistro1100() As String
        Dim SqlT As String = String.Empty

        SqlT = " SELECT Empresa_Id, Per_Apu_Cred, Orig_Cred, Cnpj_Suc, Cod_Cred, Vl_Cred_Apu, Vl_Cred_Ext_Apu, Vl_Tot_Cred_Apu, Vl_Cred_Desc_Pa_Ant, Vl_Cred_Per_Pa_Ant, " & vbCrLf &
               "        Vl_Cred_Dcomp_Pa_Ant, Sd_Cred_Disp_Efd, Vl_Cred_Desc_Efd, Vl_Cred_Per_Efd, Vl_Cred_Dcomp_Efd, Vl_Cred_Trans, Vl_Cred_Out, Sld_Cred_Fim" & vbCrLf &
               "   INTO #1100" & vbCrLf &
               "   FROM PisCofins_Reg_1100" & vbCrLf &
               "  WHERE Ano = " & Format(CDate(txtDataInicial.Text), "yyyy") & vbCrLf &
               "    AND Mes = " & Format(CDate(txtDataInicial.Text), "MM") & vbCrLf

        Return SqlT
    End Function

    Private Function GetSqlRegistro1300() As String
        Dim SqlT As String = String.Empty
        SqlT = " Select Empresa_Id, Ano_Id, Mes_Id, Ind_Nat_Ret, Per_Rec_Ret, Vl_Ret_Apu, Vl_Ret_Ded, Vl_Ret_Per, Vl_Ret_Dcomp, Sld_Ret " & vbCrLf &
               "   INTO #1300" & vbCrLf &
               "   FROM PisCofins_Reg_1300" & vbCrLf &
               "  WHERE Empresa_Id LIKE '" & EmpresaMestre & "%'" & vbCrLf &
               "    AND Mes_Id           = " & Format(CDate(txtDataInicial.Text), "MM") & vbCrLf &
               "    AND Ano_Id           = " & Format(CDate(txtDataInicial.Text), "yyyy") & vbCrLf


        Return SqlT
    End Function

    Private Function GetSqlRegistro1500() As String
        Dim SqlT As String = String.Empty
        SqlT = " SELECT Empresa_Id, Per_Apu_Cred, Orig_Cred, Cnpj_Suc, Cod_Cred, Vl_Cred_Apu, Vl_Cred_Ext_Apu, Vl_Tot_Cred_Apu, Vl_Cred_Desc_Pa_Ant, Vl_Cred_Per_Pa_Ant, " & vbCrLf &
               "        Vl_Cred_Dcomp_Pa_Ant, Sd_Cred_Disp_Efd, Vl_Cred_Desc_Efd, Vl_Cred_Per_Efd, Vl_Cred_Dcomp_Efd, Vl_Cred_Trans, Vl_Cred_Out, Sld_Cred_Fim" & vbCrLf &
               "   INTO #1500" & vbCrLf &
               "   FROM PisCofins_Reg_1500" & vbCrLf &
               "  WHERE Empresa_Id like '" & EmpresaMestre & "%'" & vbCrLf &
               "    AND Mes           = " & Format(CDate(txtDataInicial.Text), "MM") & vbCrLf &
               "    AND Ano           = " & Format(CDate(txtDataInicial.Text), "yyyy") & vbCrLf

        Return SqlT
    End Function

    Private Function GetSqlRegistro1700() As String
        Dim SqlT As String = String.Empty
        SqlT = " Select Empresa_Id, Ano_Id, Mes_Id, Ind_Nat_Ret, Per_Rec_Ret, Vl_Ret_Apu, Vl_Ret_Ded, Vl_Ret_Per, Vl_Ret_Dcomp, Sld_Ret " & vbCrLf &
               "   INTO #1700" & vbCrLf &
               "   from PisCofins_Reg_1700" & vbCrLf &
               "  where Empresa_Id like '" & EmpresaMestre & "%'" & vbCrLf &
               "    And Mes_Id           = " & Format(CDate(txtDataInicial.Text), "MM") & vbCrLf &
               "    And Ano_Id           = " & Format(CDate(txtDataInicial.Text), "yyyy") & vbCrLf


        Return SqlT
    End Function

    Private Function GetSqlRegistro1900() As String
        Dim SqlT As String = String.Empty

        If cboIncidencia.SelectedValue = 2 Then
            SqlT = " SELECT '1900' AS Reg," & vbCrLf &
                   "        NF.Empresa_Id," & vbCrLf &
                   "        99 AS Cod_Mod," & vbCrLf &
                   "        NF.Serie_Id AS Ser," & vbCrLf &
                   "        ISNULL(OE.STPISCOFINS, 0) AS CST_PIS, " & vbCrLf &
                   "        ISNULL(OE.STPISCOFINS, 0) AS CST_COFINS," & vbCrLf &
                   "        NFxI.CFOP_Id AS CFOP, " & vbCrLf &
                   "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PRODUTO' THEN NFxE.Valor END, 0)) AS Vl_Doc, " & vbCrLf &
                   "        COUNT(NF.EntradaSaida_Id) AS Quantos" & vbCrLf &
                   "   INTO #1900 " & vbCrLf &
                   "   FROM NotasFiscais AS NF" & vbCrLf &
                   "  INNER JOIN NotasFiscaisXItens AS NFxI" & vbCrLf &
                   "     ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                   "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                   "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                   "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
                   "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                   "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                   "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                   "  INNER JOIN NotasFiscaisXEncargos AS NFxE" & vbCrLf &
                   "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
                   "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
                   "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf &
                   "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
                   "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id " & vbCrLf &
                   "    AND NFxI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
                   "    AND NFxI.Nota_Id         = NFxE.Nota_Id" & vbCrLf &
                   "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
                   "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id" & vbCrLf &
                   "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
                   "   JOIN OperacaoXEstado OE " & vbCrLf &
                   "     ON NFxI.OperacaoXEstado = OE.Codigo_Id " & vbCrLf &
                   "  WHERE NF.Empresa_Id LIKE '" & EmpresaMestre & "%'" & vbCrLf &
                   "    And (ISNULL(NF.Situacao, 1) = 1) And NF.EntradaSaida_Id = 'S' " & vbCrLf &
                   "    And (NF.Movimento BETWEEN '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') " & vbCrLf &
                   "  GROUP BY NF.Empresa_ID, NF.Serie_Id, ISNULL(OE.STPISCOFINS, 0), NFxI.CFOP_Id" & vbCrLf &
                   " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id = 'PIS' THEN NFxE.Valor END, 0)) > 0)" & vbCrLf
        Else
            SqlT = " SELECT 1 AS Cod, '' Empresa_ID " & vbCrLf &
                   "   INTO #1900 " & vbCrLf &
                   "  WHERE 1=2 " & vbCrLf
        End If

        Return SqlT
    End Function

    Private Sub CompoeRegistro1100(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro1100 As Integer, ByRef RegistroGeral As Integer, R1001Empresa As String)

        Dim linhaTemp As String = String.Empty

        For Each draux In ds.Tables("1100").Select("Empresa_Id = " & R1001Empresa)
            With draux
                linhaTemp = "|" & "1100"
                linhaTemp &= "|" & Replace(.Item("Per_Apu_Cred"), "/", "")
                linhaTemp &= "|" & .Item("Orig_Cred")
                linhaTemp &= "|" & .Item("Cnpj_Suc")
                linhaTemp &= "|" & .Item("Cod_Cred")
                linhaTemp &= "|" & .Item("Vl_Cred_Apu")
                linhaTemp &= "|" & .Item("Vl_Cred_Ext_Apu")
                linhaTemp &= "|" & .Item("Vl_Tot_Cred_Apu")
                linhaTemp &= "|" & .Item("Vl_Cred_Desc_Pa_Ant")
                linhaTemp &= "|" & .Item("Vl_Cred_Per_Pa_Ant")
                linhaTemp &= "|" & .Item("Vl_Cred_Dcomp_Pa_Ant")
                linhaTemp &= "|" & .Item("Sd_Cred_Disp_Efd")
                linhaTemp &= "|" & .Item("Vl_Cred_Desc_Efd")
                linhaTemp &= "|" & .Item("Vl_Cred_Per_Efd")
                linhaTemp &= "|" & .Item("Vl_Cred_Dcomp_Efd")
                linhaTemp &= "|" & .Item("Vl_Cred_Trans")
                linhaTemp &= "|" & .Item("Vl_Cred_Out")
                linhaTemp &= "|" & .Item("Sld_Cred_Fim")
                linhaTemp &= "|"

                ArquivoAux.Add(linhaTemp)
                Registro1100 += 1
                RegistroGeral += 1

            End With
        Next ' Fim 1100 --------

    End Sub

    Private Sub CompoeRegistro1300(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro1300 As Integer, ByRef RegistroGeral As Integer, ByVal R1001Empresa As String)

        Dim linhaTemp As String = String.Empty

        For Each draux In ds.Tables("1300").Select("Empresa_Id = " & R1001Empresa)
            With draux

                linha = "|" & "1300"
                linha &= "|" & Format(.Item("Ind_Nat_Ret"), "00")
                linha &= "|" & Replace(.Item("Per_Rec_Ret"), "/", "")
                linha &= "|" & .Item("Vl_Ret_Apu")
                linha &= "|" & .Item("Vl_Ret_Ded")
                linha &= "|" & .Item("Vl_Ret_Per")
                linha &= "|" & .Item("Vl_Ret_Dcomp")
                linha &= "|" & .Item("Sld_Ret")
                linha &= "|"

                ArquivoAux.Add(linha)
                Registro1300 += 1
                RegistroGeral += 1

            End With
        Next ' Fim 1300 --------

    End Sub

    Private Sub CompoeRegistro1500(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro1500 As Integer, ByRef RegistroGeral As Integer, ByVal R1001Empresa As String)

        Dim linhaTemp As String = String.Empty

        For Each draux In ds.Tables("1500").Select("Empresa_Id = " & R1001Empresa)
            With draux
                linha = "|" & "1500"
                linha &= "|" & Replace(.Item("Per_Apu_Cred"), "/", "")
                linha &= "|" & .Item("Orig_Cred")
                linha &= "|" & .Item("Cnpj_Suc")
                linha &= "|" & .Item("Cod_Cred")
                linha &= "|" & .Item("Vl_Cred_Apu")
                linha &= "|" & .Item("Vl_Cred_Ext_Apu")
                linha &= "|" & .Item("Vl_Tot_Cred_Apu")
                linha &= "|" & .Item("Vl_Cred_Desc_Pa_Ant")
                linha &= "|" & .Item("Vl_Cred_Per_Pa_Ant")
                linha &= "|" & .Item("Vl_Cred_Dcomp_Pa_Ant")
                linha &= "|" & .Item("Sd_Cred_Disp_Efd")
                linha &= "|" & .Item("Vl_Cred_Desc_Efd")
                linha &= "|" & .Item("Vl_Cred_Per_Efd")
                linha &= "|" & .Item("Vl_Cred_Dcomp_Efd")
                linha &= "|" & .Item("Vl_Cred_Trans")
                linha &= "|" & .Item("Vl_Cred_Out")
                linha &= "|" & .Item("Sld_Cred_Fim")
                linha &= "|"

                ArquivoAux.Add(linha)
                Registro1500 += 1
                RegistroGeral += 1

            End With
        Next ' Fim 1500 --------

    End Sub

    Private Sub CompoeRegistro1700(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro1700 As Integer, ByRef RegistroGeral As Integer, ByVal R1001Empresa As String)

        Dim linhaTemp As String = String.Empty

        For Each draux In ds.Tables("1700").Select("Empresa_Id = " & R1001Empresa)
            With draux

                linhaTemp = "|" & "1700"
                linhaTemp &= "|" & Format(.Item("Ind_Nat_Ret"), "00")
                linhaTemp &= "|" & Replace(.Item("Per_Rec_Ret"), "/", "")
                linhaTemp &= "|" & .Item("Vl_Ret_Apu")
                linhaTemp &= "|" & .Item("Vl_Ret_Ded")
                linhaTemp &= "|" & .Item("Vl_Ret_Per")
                linhaTemp &= "|" & .Item("Vl_Ret_Dcomp")
                linhaTemp &= "|" & .Item("Sld_Ret")
                linhaTemp &= "|"


                ArquivoAux.Add(linhaTemp)
                Registro1700 += 1
                RegistroGeral += 1

            End With
        Next ' Fim 1700 --------

    End Sub

    Private Sub CompoeRegistro1900(ByRef ds As DataSet, ByRef ArquivoAux As ArrayList, ByRef Registro1900 As Integer, ByRef RegistroGeral As Integer, ByVal R1001Empresa As String)

        Dim linhaTemp As String = String.Empty
        For Each draux In ds.Tables("1900").Select("Empresa_Id = " & R1001Empresa)
            With draux
                linhaTemp = "|" & .Item("Reg")
                linhaTemp &= "|" & Empresa(0)
                linhaTemp &= "|99"
                linhaTemp &= "|" & .Item("Ser")
                linhaTemp &= "|"
                linhaTemp &= "|00"
                linhaTemp &= "|" & .Item("Vl_Doc")
                linhaTemp &= "|" & .Item("Quantos")
                linhaTemp &= "|" & Format(CInt(.Item("CST_Pis")), "00")
                linhaTemp &= "|" & Format(CInt(.Item("Cst_COFINS")), "00")
                linhaTemp &= "|" & .Item("CFOP")

                linhaTemp &= "|"
                linhaTemp &= "|301010101"
                linhaTemp &= "|"

                ArquivoAux.Add(linhaTemp)
                Registro1900 += 1
                RegistroGeral += 1

            End With
        Next ' Fim 1900 --------
    End Sub

#End Region

End Class