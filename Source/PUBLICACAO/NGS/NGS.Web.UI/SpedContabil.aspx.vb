Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class SpedContabil
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Ds As DataSet
    Dim Dsr As DataSet

    Dim Empresa() As String
    Dim Cliente() As String
    Dim Ciclo As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Contabil.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("SpedContabil", "ACESSAR") Then
                CarregarUnidade()
                txtDataInicial.Text = String.Format("01/01/{0}", DateTime.Now.Year - 1)
                txtDataFinal.Text = String.Format("31/12/{0}", DateTime.Now.Year - 1)

                ddlFormaDeTributacao.SelectedValue = 1 'PJ em Geral – Lucro Real

                ddlSituacaoInicioPeriodo.SelectedValue = 0 'Normal (Início no primeiro dia do ano ou do mês)

                If Not Directory.Exists(Server.MapPath("~/Sped/Ecd/" & Today().Year() - 1)) Then
                    Directory.CreateDirectory(Server.MapPath("~/Sped/Ecd/" & Today().Year() - 1))
                End If
                txtArquivoAuxiliar.Text = "c:\sped\ecd\2016\NomeDoArquivo.rtf"

                txtNumeroDoLivro.Text = "123"
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        'txtEstadoDaJunta.Text = ""
        'txtDataDaFundacao.Text = ""
    End Sub

    Protected Sub Processar()
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")

            Dim arquivo As String = txtArquivoDeSaida.Text
            Dim linha As String
            'Dim strm As StreamWriter
            Dim strmSecundario As StreamWriter
            Dim ConteudoDoArquivo As New StringBuilder()

            Dim Registro0000 As Integer = 0     'Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária
            Dim Registro0001 As Integer = 0     'Abertura do Bloco 0
            Dim Registro0007 As Integer = 0     'Outras Inscrições Cadastrais do empresário ou Sociedade Empresária
            Dim Registro0020 As Integer = 0     'Escrituração Contábil Descentralizada
            Dim Registro0150 As Integer = 0     'Tabela de Cadastro do Participante
            Dim Registro0180 As Integer = 0     'Identificação do Relacionamento com o Participante
            Dim Registro0990 As Integer = 0     'Encerramento do Bloco 0

            Dim RegistroI001 As Integer = 0     'Abertura do Broco I
            Dim RegistroI010 As Integer = 0     'Identificação da Escrituracao Contábil
            Dim RegistroI012 As Integer = 0     'Livros Auxiliares ao Diario
            Dim RegistroI015 As Integer = 0     'Identificação das Contas da escrituração Resumida a que se Refere a Escrituração Auxiliar
            Dim RegistroI020 As Integer = 0     'Campos Adicionais
            Dim RegistroI030 As Integer = 0     'Termo de Abertura do Livro
            Dim RegistroI050 As Integer = 0     'Plano de Contas
            Dim RegistroI051 As Integer = 0     'Plano de Contas Referencial
            Dim RegistroI052 As Integer = 0     'Indicação dos Códigos de Aglutinação
            Dim RegistroI075 As Integer = 0     'Tabela de Históricos Padronizados
            Dim RegistroI100 As Integer = 0     'Centros de Custos
            Dim RegistroI150 As Integer = 0     'Saldos Periódicos - Identificação do Periodo
            Dim RegistroI155 As Integer = 0     'Detalhe dos Saldos Periódicos
            Dim RegistroI200 As Integer = 0     'Lancamento Contábil
            Dim RegistroI250 As Integer = 0     'Partidas do Lançamento
            Dim RegistroI300 As Integer = 0     'Balancetes Diários - Identificação da Data
            Dim RegistroI310 As Integer = 0     'Detalhes do Balancete Diário
            Dim RegistroI350 As Integer = 0     'Saldo das Contas de Resultado Antes do Encerramento - Identificação da Data
            Dim RegistroI355 As Integer = 0     'Detalhes dos Saldos das Contas de Resultado Antes do Encerramento
            Dim RegistroI500 As Integer = 0     'Parametros de Impressão e Visualização do Livro Razao Auxiliar com Leiaute Parametrizavel
            Dim RegistroI510 As Integer = 0     'Definição de Campos do Livro Razão Auxiliar com Leiaute Parametrizavel
            Dim RegistroI550 As Integer = 0     'Detalhes do Livro Razao Auxiliar com Leioute Parametrizavel
            Dim RegistroI555 As Integer = 0     'Totais no Livro Razao Auxiliar com Leiaute Parametrizavel
            Dim RegistroI990 As Integer = 0     'Encerramento do Bloco I

            Dim RegistroJ001 As Integer = 0     'Abertura do Bloco J
            Dim RegistroJ005 As Integer = 0     'Demonstrações Contábeis
            Dim RegistroJ100 As Integer = 0     'Balanço Patrimonial
            Dim RegistroJ150 As Integer = 0     'Demonstração do Resultado do Exercicio
            Dim RegistroJ800 As Integer = 0     'Outras Informações
            Dim RegistroJ900 As Integer = 0     'Termo de Encerramento
            Dim RegistroJ930 As Integer = 0     'Identificação dos Signatários da Escrituração
            Dim RegistroJ990 As Integer = 0     'Encerramento do Bloco J

            Dim Registro9001 As Integer = 0     'Abertura do Bloco 9
            Dim Registro9900 As Integer = 0     'Registros do Arquivo
            Dim Registro9990 As Integer = 0     'Encerramento do Bloco 9
            Dim Registro9999 As Integer = 0     'Encerramento do Arquivo Digital

            Dim RegistroGeral As Integer = 0    'Registro Geral

            Dim CGC As String = ""
            Dim Inscricao As String = ""

            'Me.Cursor = Cursors.WaitCursor

            Dim PreArquivo As String = Server.MapPath(String.Format("~/Sped/Ecd/{0}/", DateTime.Now.Year - 1) & Left(Empresa(0), 8) & ".txt")
            arquivo = PreArquivo

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            Dim ds As New DataSet
            Dim Array As New ArrayList

            'strm = New StreamWriter(arquivo, True)

            'Registro 0000 - Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária
            'Leiaute Anterior--------------------------------------

            Sql = " SELECT  Top 1   Clientes.*, ClientesXEmpresas.*, Municipios.*" & vbCrLf &
                " FROM      Clientes INNER JOIN" & vbCrLf &
                "           ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf &
                "           Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                "           LEFT OUTER JOIN Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id" & vbCrLf &
                " WHERE     Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "|0000"     '01 -REGISTRO
                        linha &= "|LECD"    '02 -Texto fixo
                        linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy")   '03 - Data Inicial
                        linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")     '04 - Data Final
                        linha &= "|" & RTrim(.Item("Nome"))                             '05 - Nome da Empresa
                        linha &= "|" & Left(.Item("Cliente_Id"), 14)                    '06 - CNPJ
                        linha &= "|" & .Item("Estado")                                  '07 - Estado
                        linha &= "|" & RTrim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    '08 - Inscricao Estadual
                        linha &= "|" & .ITEM("EstadoIbge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")     '09 - Código do município
                        linha &= "|" & .Item("InscricaoMunicipal")                      '10 - Inscrição Municipal
                        linha &= "|" & DdlSituacaoEspecial.SelectedValue                '11 - Indicador de situação especial

                        linha &= "|" & ddlSituacaoInicioPeriodo.SelectedValue           '12 - Indicador de situação no início do período
                        linha &= "|1"   '13 - Indicador de existência de NIRE: 0 – Empresa não possui registro na Junta Comercial (não possui NIRE) 1 – Empresa possui registro na Junta Comercial (possui NIRE)
                        If chkRetifica.Checked Then '14 - Indicador de finalidade da escrituração: 0 - Original 1 – Substituta - Retificação - Furlan - 22-07-2021
                            linha &= "|1"
                        Else
                            linha &= "|0"
                        End If
                        linha &= "|"  '15 - Hash da escrituração substituída
                        linha &= "|0" '16 - Auditoria Independente
                        linha &= "|0" '17 - Indicador do tipo de ECD: 0 – ECD de empresa não participante de SCP como sócio ostensivo. 1 – ECD de empresa participante de SCP como sócio ostensivo. 2 – ECD da SCP.
                        linha &= "|"  '18 - COD_SCP
                        linha &= "|N" '19 - 
                        linha &= "|"  '20 - vide abaixo

                        If Format(CDate(txtDataFinal.Text), "yyyy") > 2014 Then '20 - campo
                            linha &= "N"   'Quem tem moeda funcional diferente de real (comercio exterior por ex) pode ter que marcar com "S" e não "N".
                        End If

                        If Format(CDate(txtDataFinal.Text), "yyyy") > 2018 Then
                            linha &= "|0"   '21 - Indicador da modalidade de escrituração centralizada ou descentralizada: 0 – Escrituração Centralizada 1 – Escrituração Descentralizada
                            linha &= "|0"   '22 - Indicador de mudança de plano de contas: 0 – Não houve mudança no plano de contas. 1 – Houve mudança no plano de contas

                            'Mudei para Selecionar a Forma de Tributação da Empresa na Tela - Furlan - 27-07-2021
                            'If cbTrimestral.Checked AndAlso Left(Empresa(0), 8) = "05366261" Then 'Nutri
                            '    linha &= "1|"   '23 - Código do Plano de Contas Referencial que será utilizado para o mapeamento de todas as contas analíticas
                            'ElseIf cbTrimestral.Checked Then
                            '    linha &= "2|"   '23 - Código do Plano de Contas Referencial que será utilizado para o mapeamento de todas as contas analíticas
                            'Else
                            '    linha &= "1|"   '23 - Código do Plano de Contas Referencial que será utilizado para o mapeamento de todas as contas analíticas
                            'End If
                            linha &= "|" & ddlFormaDeTributacao.SelectedValue  '23 - Código do Plano de Contas Referencial

                            linha &= "|" 'fechamento do registro

                        End If

                    End With

                    'strm.WriteLine(linha)
                    ConteudoDoArquivo.AppendLine(linha)
                    Registro0000 += 1
                    RegistroGeral += 1
                Next
            End If

            ' Registro 0001  - Abertura do Bloco 0

            linha = "|0001"
            linha &= "|" & DdlIndicadorDeMovimento.SelectedIndex
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            Registro0001 += 1
            RegistroGeral += 1

            'Registro 0007 - Outras Inscrições Cadastrais do empresário ou Sociedade Empresária

            Sql = "  SELECT     Clientes.*, ClientesXEmpresas.*" & vbCrLf &
                  "     FROM    Clientes INNER JOIN" & vbCrLf &
                  "             ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf &
                  "             Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                  "     WHERE   Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "|0007"                                                                                                'Tipo
                        linha &= "|" & .Item("Estado")
                        linha &= "|" & RTrim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    'Inscricao Estadual
                        linha &= "|"
                    End With

                    'strm.WriteLine(linha)
                    ConteudoDoArquivo.AppendLine(linha)
                    Registro0007 += 1
                    RegistroGeral += 1
                Next
            End If

            'Registro 0020 - Escrituração Contábil Descentralizada

            'Sql = "  SELECT    Clientes.*, ClientesXEmpresas.*" & vbCrLf & _
            '      "     FROM    Clientes INNER JOIN" & vbCrLf & _
            '      "             ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
            '      "             Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf & _
            '      "     WHERE   Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf


            ' Isolado em 22-05-2017 - Neri
            'Sql = " SELECT  Top 1   Clientes.*, ClientesXEmpresas.*, Municipios.*" & vbCrLf & _
            '    " FROM      Clientes INNER JOIN" & vbCrLf & _
            '    "           ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
            '    "           Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf & _
            '    "           LEFT OUTER JOIN Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id" & vbCrLf & _
            '    " WHERE     Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf

            'ds = Banco.ConsultaDataSet(Sql, "Clientes")

            'If ds.Tables(0).Rows.Count > 0 Then
            '    For Each dr In ds.Tables(0).Rows
            '        With dr
            '            linha = "|0020"
            '            If .Item("Matriz") = "S" Then
            '                linha &= "|0"
            '            Else
            '                linha &= "|1"
            '            End If
            '            linha &= "|" & Left(.Item("Cliente_Id"), 14)
            '            linha &= "|" & .Item("Estado")
            '            linha &= "|" & RTrim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    'Inscricao Estadual
            '            linha &= "|" & .ITEM("EstadoIbge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
            '            linha &= "|" & .Item("InscricaoMunicipal")
            '            linha &= "|" & Funcoes.AlinharDireita(.Item("RegistroNire"), 11, "0")
            '            linha &= "|"
            '        End With

            '        'strm.WriteLine(linha)
            '        ConteudoDoArquivo.AppendLine(linha)
            '        Registro0020 += 1
            '        RegistroGeral += 1
            '    Next
            'End If



            'Registro 0150 - Tabela de Cadastro do Participante
            'sql = "  SELECT    Clientes.*, ClientesXEmpresas.*"
            'sql &= " FROM      Clientes INNER JOIN"
            'sql &= "           ClientesXEmpresas ON Clientes.Pais_Id = ClientesXEmpresas.Pais_Id AND Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
            'sql &= "           Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id"
            'sql &= " WHERE     Clientes.Pais_Id = '" & Utilitarios.[Global].GPais & "' and Clientes.Cliente_Id = '" & txtCNPJEmpresa.Text & "' And Clientes.Endereco_Id = 0"

            'ds = banco.ConsultaDataSet(sql, "Clientes")

            'If ds.Tables(0).Rows.Count > 0 Then
            '    For Each dr In ds.Tables(0).Rows
            '        With dr
            '            linha = "|0150"
            '            linha &= "|" & Left(ddlCodigoDeRelacionamento.Text, 2)
            '            linha &= "|" & RTrim(.Item("Nome"))
            '            linha &= "|" & .Item("Pais_Id")                                                         'Pais do participante
            '            linha &= "|" & Left(.Item("Cliente_Id"), 14)
            '            linha &= "||"                                                                           'CPF do Participante
            '            linha &= "||"                                                                           'Numero de Identificação do trabalhador, Pis Pasep, SUS
            '            linha &= "|" & .Item("Estado_Id")
            '            linha &= "|" & RTrim(AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    'Inscricao Estadual
            '            linha &= "|" & RTrim(AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    'Inscricao Estadual
            '            linha &= "|" & AlinharDireita(.Item("CodigoDoMunicipio"), 7, "0")
            '            linha &= "|" & .Item("InscricaoMunicipal")
            '            linha &= "||"                                                                           'Suframa
            '            linha &= "|"
            '        End With

            '        strm = New StreamWriter(arquivo, True)
            '        strm.WriteLine(linha)
            '        strm.Close()
            '        Registro0150 += 1
            '        RegistroGeral += 1
            '    Next
            'End If

            ' Registro 0180  - Identificação do Relacionamento com o Participante

            'linha = "|0180"
            'linha &= "|" & Left(ddlCodigoDeRelacionamento.Text, 2)
            'linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy")                                         'Data do inicio do relacionamento
            'linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")                                           'Data do fim do relacionamento
            'linha &= "|"

            'strm = New StreamWriter(arquivo, True)
            'strm.WriteLine(linha)
            'strm.Close()
            'Registro0180 += 1
            'RegistroGeral += 1


            '' Registro 0990  - Encerramento do Bloco 0
            Registro0990 += 1

            linha = "|0990"
            linha &= "|" & Registro0000 + Registro0001 + Registro0007 + Registro0020 + Registro0150 + Registro0180 + Registro0990
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            RegistroGeral += 1

            '  Fim do Bloco 0

            ' Registro I001  - Abertura do Bloco I

            linha = "|I001"
            linha &= "|" & DdlIndicadorDeMovimento.SelectedIndex
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            RegistroI001 += 1
            RegistroGeral += 1

            ' Registro I010  - Identificação da Escrituração Contábil

            Sql = "  SELECT     Clientes.*, ClientesXEmpresas.*" & vbCrLf &
                  "     FROM    Clientes INNER JOIN" & vbCrLf &
                  "             ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf &
                  "             Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                  "     WHERE   Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "|I010"
                        linha &= "|" & .Item("EscrituracaoContabil")
                        If Format(CDate(txtDataFinal.Text), "yyyy") = 2014 Then
                            linha &= "|3.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2015 Then
                            linha &= "|4.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2016 Then
                            linha &= "|5.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2017 Then
                            linha &= "|6.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2018 Then
                            linha &= "|7.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2019 Then
                            linha &= "|8.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2020 Then
                            linha &= "|9.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2021 Then
                            linha &= "|9.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2022 Then
                            linha &= "|9.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2023 Then
                            linha &= "|9.00"
                        ElseIf Format(CDate(txtDataFinal.Text), "yyyy") = 2024 Then
                            linha &= "|9.00"
                        End If
                        linha &= "|"
                    End With

                    'strm.WriteLine(linha)
                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroI010 += 1
                    RegistroGeral += 1
                Next
            End If

            'Registro I030 - Termo De Abertura Do Livro
            Sql = "  SELECT     Clientes.*, ClientesXEmpresas.*" & vbCrLf &
                  "     FROM    Clientes INNER JOIN" & vbCrLf &
                  "             ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf &
                  "             Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                  "     WHERE   Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "|I030"
                        linha &= "|TERMO DE ABERTURA"
                        linha &= "|" & txtNumeroDoLivro.Text
                        linha &= "|" & Mid(DdlNaturezaDoLivro.Text, 7, 50)
                        linha &= "|QTD-LIN"
                        linha &= "|" & RTrim(.Item("Nome"))
                        linha &= "|" & Funcoes.AlinharDireita(.Item("RegistroNire"), 11, "0")
                        linha &= "|" & Left(.Item("Cliente_Id"), 14)
                        linha &= "|" & Format(CDate(.Item("DataRegistroNire")), "ddMMyyyy")
                        linha &= "|" & Format(CDate(.Item("DataRegistroNire")), "ddMMyyyy")
                        linha &= "|" & .Item("Cidade")
                        linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")
                        linha &= "|"
                    End With

                    'strm.WriteLine(linha)
                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroI030 += 1
                    RegistroGeral += 1
                Next
            End If


            'Registro I050 - Plano de Contas
            'If rbPlanoAntigo.Checked Then
            '    Sql = " SELECT     Razao.Conta_Id + Razao.Cliente_Id +  convert(nvarchar,Razao.EndCliente_Id) AS Conta_Id, isnull(Clientes.Nome, 'Cliente Nao Encontrado') AS Titulo" & vbCrLf & _
            '          "     FROM    Razao LEFT OUTER JOIN" & vbCrLf & _
            '          "             Clientes ON Razao.Cliente_Id = Clientes.Cliente_Id" & vbCrLf & _
            '          "     WHERE   Left(Razao.Empresa_Id, 8) like ('" & Left(Empresa(0), 8) & "%') And Razao.Cliente_ID <> ''" & vbCrLf & _
            '          " Group by Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id, Clientes.Nome" & vbCrLf & _
            '          "             " & vbCrLf & _
            '          " Union" & vbCrLf

            'Sql &= " SELECT Conta_Id, Titulo "
            'Sql &= " FROM   PlanoDeContas"
            'Sql &= " WHERE  PlanoDeContas.Empresa_Id = '" & Empresa(0) & "'" & " And PlanoDeContas.EndEmpresa_Id = 0"
            'Sql &= " Order by Conta_Id"

            'Sql &= " SELECT     Conta_Id, Titulo            " & vbCrLf & _
            '       "    FROM    PlanoDeContas               " & vbCrLf & _
            '       "    WHERE   Len(Conta_Id) < 9     " & vbCrLf & _
            '       "                                    " & vbCrLf & _
            '       " Union                              " & vbCrLf & _
            '       "                                    " & vbCrLf & _
            '       " SELECT  PlanoDeContas.Conta_Id, PlanoDeContas.Titulo  " & vbCrLf & _
            '       "    From Razao " & vbCrLf & _
            '       "        LEFT OUTER JOIN PlanoDeContas" & vbCrLf & _
            '       "            ON PlanoDeContas.Conta_Id = Razao.Conta_Id " & vbCrLf & _
            '       "    WHERE    (LEN(PlanoDeContas.Conta_Id) = 9) And Razao.Empresa_Id Like '" & Left(Empresa(0), 8) & "%' " & vbCrLf & _
            '       " Group by PlanoDeContas.Conta_Id, PlanoDeContas.Titulo " & vbCrLf & _
            '       " Order by Conta_Id" & vbCrLf

            'ElseIf rbPlanoNovo.Checked Then
            Sql = " SELECT    Razao.Conta_Id + Razao.Cliente_Id +  convert(nvarchar,Razao.EndCliente_Id) AS Conta_Id, isnull(Clientes.Nome, 'Cliente Nao Encontrado') AS Titulo" & vbCrLf &
                  "     FROM       Razao LEFT OUTER JOIN" & vbCrLf &
                  "            Clientes ON Razao.Cliente_Id = Clientes.Cliente_Id" & vbCrLf &
                  "                    And Razao.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                  "     WHERE      (Razao.Empresa_Id Like '" & Left(Empresa(0), 8) & "%') And Razao.Cliente_ID <> '' " & vbCrLf &
                  "            And year(Razao.Movimento_Id) < = " & Format(CDate(txtDataFinal.Text), "yyyy") & vbCrLf &
                  " Group by Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id, Clientes.Nome" & vbCrLf &
                  "                                                            " & vbCrLf &
                  " Union" & vbCrLf &
                  "                                                             " & vbCrLf &
                  " SELECT Conta_Id, Titulo " & vbCrLf &
                  "     FROM   PlanoDeContas" & vbCrLf &
                  "     WHERE     Len(Conta_Id) < 9" & vbCrLf &
                  "                                                     " & vbCrLf &
                  " Union" & vbCrLf &
                  "                                                     " & vbCrLf &
                  " SELECT  PlanoDeContas.Conta_Id, PlanoDeContas.Titulo  " & vbCrLf &
                  "     From Razao " & vbCrLf &
                  "         LEFT OUTER JOIN PlanoDeContas" & vbCrLf &
                  "             ON PlanoDeContas.Conta_Id = Razao.Conta_Id " & vbCrLf &
                  "     WHERE    (LEN(PlanoDeContas.Conta_Id) = 9) And Razao.Empresa_Id Like '" & Left(Empresa(0), 8) & "%' " & vbCrLf &
                  " Group by PlanoDeContas.Conta_Id, PlanoDeContas.Titulo " & vbCrLf &
                  " Order by Conta_Id" & vbCrLf
            'End If

            ds = Banco.ConsultaDataSet(Sql, "PlanoDeContas")

            Dim Nivel1 As String = ""
            Dim Nivel2 As String = ""
            Dim Conta As String = ""

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    If Conta = dr("Conta_Id") Then
                    Else
                        With dr
                            linha = "|I050"
                            linha &= "|0101" & Year(CDate(txtDataInicial.Text))


                            '****Plano Novo
                            Nivel1 = Left(.Item("Conta_Id"), 1)
                            Nivel2 = Left(.Item("Conta_Id"), 3)
                            If Nivel2 = "204" Then
                                linha &= "|03"
                            ElseIf Nivel2 = "105" Then
                                '    linha &= "|05"
                                'ElseIf conta = "1050101" Then
                                linha &= "|01"
                            ElseIf Nivel2 = "205" Then
                                linha &= "|05"
                            Else
                                If Nivel1 = "1" Then
                                    linha &= "|01"
                                ElseIf Nivel1 = "2" Then
                                    linha &= "|02"
                                ElseIf Nivel1 = "3" Then
                                    linha &= "|04"
                                ElseIf Nivel1 = "4" Then
                                    linha &= "|04"
                                ElseIf Nivel1 = "5" Then
                                    'Contas de Resultado - Furlan - 03/07/2020
                                    If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                                        linha &= "|09"
                                    Else
                                        linha &= "|04"
                                    End If
                                End If
                            End If

                            If Len(.Item("Conta_Id")) > 7 Then
                                linha &= "|A"
                            Else
                                linha &= "|S"
                            End If

                            If Len(.Item("Conta_Id")) > 7 Then
                                linha &= "|5"
                            ElseIf Len(.Item("Conta_Id")) = 7 Then
                                linha &= "|4"
                            ElseIf Len(.Item("Conta_Id")) = 5 Then
                                linha &= "|3"
                            ElseIf Len(.Item("Conta_Id")) = 3 Then
                                linha &= "|2"
                            ElseIf Len(.Item("Conta_Id")) = 1 Then
                                linha &= "|1"
                            End If

                            linha &= "|" & .Item("Conta_Id")
                            Conta = .Item("Conta_Id")

                            If Len(.Item("Conta_Id")) > 7 Then
                                linha &= "|" & Left(.Item("Conta_Id"), 7)
                            ElseIf Len(.Item("Conta_Id")) = 7 Then
                                linha &= "|" & Left(.Item("Conta_Id"), 5)
                            ElseIf Len(.Item("Conta_Id")) = 5 Then
                                linha &= "|" & Left(.Item("Conta_Id"), 3)
                            ElseIf Len(.Item("Conta_Id")) = 3 Then
                                linha &= "|" & Left(.Item("Conta_Id"), 1)
                            ElseIf Len(.Item("Conta_Id")) = 1 Then
                                linha &= "|"
                            End If


                            linha &= "|" & RTrim(.Item("Titulo"))
                            linha &= "|"


                            'strm.WriteLine(linha)
                            ConteudoDoArquivo.AppendLine(linha)
                            RegistroI050 += 1
                            RegistroGeral += 1


                            '-------Plano de Contas Referencial-------

                            If Len(.Item("Conta_Id")) > 7 Then
                                If Len(.Item("Conta_Id")) > 7 Then
                                    Conta = Microsoft.VisualBasic.Left(.Item("Conta_Id"), 9)
                                End If

                                Sql = "  SELECT   Referencial, CodigoDeCustoECD "
                                Sql &= " FROM     PlanoDeContasXReferencialBacen"
                                Sql &= " WHERE    (Conta = '" & Conta & "' OR Conta = '" & Microsoft.VisualBasic.Left(Conta, 7) & "') Group by Referencial, CodigoDeCustoECD"

                                Dsr = Banco.ConsultaDataSet(Sql, "Consulta")

                                If Dsr.Tables(0).Rows.Count > 0 Then
                                    For Each drr In Dsr.Tables(0).Rows
                                        With drr
                                            linha = "|I051"

                                            'À Partir de 2019 não informa mais esse campo - Furlan - 11/05/2020
                                            If Format(CDate(txtDataFinal.Text), "yyyy") < 2019 Then
                                                linha &= "|1"
                                            End If

                                            linha &= "|" & .Item("CodigoDeCustoECD")
                                            linha &= "|" & .Item("Referencial")
                                            linha &= "|"
                                        End With

                                        ConteudoDoArquivo.AppendLine(linha)
                                        RegistroI051 += 1
                                        RegistroGeral += 1
                                    Next
                                End If

                            End If

                            '---------------------------------------------

                            If Len(.Item("Conta_Id")) > 7 Then
                                linha = "|I052"
                                linha &= "|"
                                linha &= "|" & Left(.Item("Conta_Id"), 7)
                                linha &= "|"

                                'strm.WriteLine(linha)
                                ConteudoDoArquivo.AppendLine(linha)
                                RegistroI052 += 1
                                RegistroGeral += 1
                            End If
                        End With

                    End If

                Next
            End If

            'Registro I100 - Centro De Custos
            'Sql = " SELECT  * FROM CentrosDeCustos"
            'Sql &= " WHERE  CentroDeCusto_Id <> 0"

            Sql = "Select Custo_Id AS CentroDeCusto_Id, Descricao FROM PlanoDeCustosECD"
            ds = Banco.ConsultaDataSet(Sql, "CentrosDeCustos")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "|I100"

                        If (Left(Empresa(0), 8) = "40938762" AndAlso Year(CDate(txtDataInicial.Text)) = 2021) Then
                            linha &= "|2202" & Year(CDate(txtDataInicial.Text))
                        Else
                            linha &= "|0101" & Year(CDate(txtDataInicial.Text))
                        End If

                        linha &= "|" & .Item("CentroDeCusto_Id")
                        linha &= "|" & RTrim(.Item("Descricao"))
                        linha &= "|"
                    End With

                    'strm.WriteLine(linha)
                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroI100 += 1
                    RegistroGeral += 1
                Next
            End If


            'Registro I150 - Saldos Periódicos - Identificação do Período

            'linha = "|I150"
            'linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy")                                                     'Data Inicial
            'linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")                                                       'Data Final
            'linha &= "|"

            'strm = New StreamWriter(arquivo, True)
            'strm.WriteLine(linha)
            'strm.Close()
            'RegistroI150 += 1
            'RegistroGeral += 1

            Dim ano As Integer = Year(txtDataInicial.Text)
            Dim Dia As Integer = 0
            Dim Mes As Integer = Month(txtDataFinal.Text)

            Dim dias As String = ""
            Dim DataInicial As String = ""
            Dim DataFinal As String = ""
            Dim Saldo As Decimal = 0

            For i As Integer = 1 To Mes

                If i = 12 Then
                    Dia = 31
                Else
                    Dia = CDate("01-" & i + 1 & "-" & ano & "").AddDays(-1).Day
                End If

                'TRATANTO OXFIBER - FURLAN - 28-07-2020
                If (Left(Empresa(0), 8) = "38388314" AndAlso ano = 2020 AndAlso i < 9) Then
                    Continue For
                End If

                'TRATANTO BAXI - FURLAN - 27-07-2022
                If (Left(Empresa(0), 8) = "40938762" AndAlso ano = 2021 AndAlso i < 2) Then
                    Continue For
                End If

                'TRATANTO NUTRI - FURLAN - 26-08-2022
                If (Left(Empresa(0), 8) = "05366261" AndAlso ano = 2020 AndAlso i < 8) Then
                    Continue For
                End If

                If (Left(Empresa(0), 8) = "24450490" AndAlso ano = 2024 AndAlso i < 7) Then
                    Continue For
                End If

                If (Left(Empresa(0), 8) = "62747840" AndAlso ano = 2024 AndAlso i < 7) Then
                    Continue For
                End If

                linha = "|I150"

                'TRATANTO BAXI - FURLAN - 27-07-2022
                If (Left(Empresa(0), 8) = "40938762" AndAlso ano = 2021 AndAlso i = 2) Then
                    linha &= "|2202" & Year(CDate(txtDataInicial.Text))
                Else
                    linha &= "|01" & String.Format("{0:D2}", i) & ano
                End If

                linha &= "|" & String.Format("{0:D2}", Dia) & String.Format("{0:D2}", i) & ano
                linha &= "|"

                'TRATANTO BAXI - FURLAN - 27-07-2022
                If (Left(Empresa(0), 8) = "40938762" AndAlso ano = 2021 AndAlso i = 2) Then
                    DataInicial = "22/" & String.Format("{0:D2}", i) & "/" & ano
                Else
                    DataInicial = "01/" & String.Format("{0:D2}", i) & "/" & ano
                End If

                DataFinal = String.Format("{0:D2}", Dia) & "/" & String.Format("{0:D2}", i) & "/" & ano


                'TRATANTO BAXI - FURLAN - 27-07-2022
                If (Left(Empresa(0), 8) = "40938762" AndAlso ano = 2021 AndAlso i = 2) Then
                    DataInicial = "22/" & String.Format("{0:D2}", i) & "/" & ano
                End If


                'strm.WriteLine(linha)
                ConteudoDoArquivo.AppendLine(linha)
                RegistroI150 += 1
                RegistroGeral += 1

                'Registro I155 - Detalhe dos Saldos Periodicos

                Sql = "Select Conta, Cliente, EndCliente_Id, Sum(Inicial) as Inicial, Sum(Debitos) as Debitos, Sum(Creditos) as Creditos from ( " & vbCrLf &
                      "  SELECT Conta_Id As Conta, Cliente_Id as Cliente, Case when Cliente_Id = '' then 0 else EndCliente_Id end as EndCliente_Id, " & vbCrLf &
                      "          ISNULL((SELECT     SUM(DebitoOficial - CreditoOficial) AS Inicial" & vbCrLf &
                      "                     FROM    Razao AS RID" & vbCrLf &
                      "                     WHERE  Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'  And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Razao.Conta_Id = Conta_Id) And (Razao.Cliente_Id = Cliente_Id) And (Razao.EndCliente_Id = EndCliente_Id)), 0) AS Inicial, " & vbCrLf &
                      "          ISNULL((SELECT     SUM(DebitoOficial) AS Inicial" & vbCrLf &
                      "                     FROM         Razao AS RMD" & vbCrLf &
                      "                     WHERE   Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'  And (Movimento_ID BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Razao.Conta_Id = Conta_Id) And (Razao.Cliente_Id = Cliente_Id) And (Razao.EndCliente_Id = EndCliente_Id)), 0) AS Debitos, " & vbCrLf &
                      "          ISNULL((SELECT     SUM(CreditoOficial) AS Inicial" & vbCrLf &
                      "                     FROM         Razao AS RMC" & vbCrLf &
                      "                     WHERE  Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'  And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "')  AND (Razao.Conta_Id = Conta_Id) And (Razao.Cliente_Id = Cliente_Id) And (Razao.EndCliente_Id = EndCliente_Id)), 0) AS Creditos" & vbCrLf &
                      "     FROM Razao" & vbCrLf &
                      "     WHERE  Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%' " & vbCrLf &
                      "     GROUP BY Conta_Id, Cliente_Id, EndCliente_Id" & vbCrLf &
                      "     ) as Consulta " & vbCrLf &
                      "     where Inicial <> 0 Or Debitos <> 0 Or Creditos <> 0" & vbCrLf &
                      "     Group By Conta, Cliente, EndCliente_Id" & vbCrLf &
                      "     ORDER BY Conta, Cliente, EndCliente_Id" & vbCrLf


                'Sql = " SELECT   Conta,   Cliente, Endereco, SUM(Inicial) AS Inicial, Sum(Debitos) as Debitos, Sum(Creditos) AS Creditos From ("

                'Sql &= " SELECT Conta_Id AS Conta, Cliente_Id AS Cliente, EndCliente_Id as Endereco, SUM(DebitoOficial - CreditoOficial) AS Inicial, 0 as Debitos, 0 AS Creditos"
                'Sql &= "  FROM Razao"
                'Sql &= " WHERE Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'  And (Movimento_Id < '" & Format(CDate(DataInicial), "yyyy/MM/dd") & "') And Left(Conta_Id, 1) <  3 "
                'Sql &= " GROUP BY Conta_Id, Cliente_Id, EndCliente_Id"

                'Sql &= " Union"
                'Sql &= " SELECT Conta_Id AS Conta, Cliente_Id AS Cliente,   EndCliente_Id as Endereco, 0 AS Inicial, SUM(DebitoOficial) AS Debitos, SUM(CreditoOficial) AS Creditos"
                'Sql &= " FROM Razao"
                'Sql &= " WHERE (LEFT(Conta_Id, 1) < 3) And Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'  And (Movimento_Id BETWEEN '" & Format(CDate(DataInicial), "yyyy/MM/dd") & "' AND '" & Format(CDate(DataFinal), "yyyy/MM/dd") & "')"
                'Sql &= " GROUP BY Conta_Id, Cliente_Id, EndCliente_Id"


                'Sql &= " Union"
                'Sql &= " SELECT Conta_Id AS Conta, '' AS Cliente, 0 AS Endereco, SUM(DebitoOficial - CreditoOficial) AS Inicial, 0 as Debitos, 0 AS Creditos"
                'Sql &= " FROM Razao"
                'Sql &= " WHERE Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'  And (Movimento_Id < '" & Format(CDate(DataInicial), "yyyy/MM/dd") & "') And Left(Conta_Id, 1) >  2 "
                'Sql &= " GROUP BY Conta_Id, Cliente_Id, EndCliente_Id"

                'Sql &= " Union"
                'Sql &= " SELECT Conta_Id AS Conta, '' AS Cliente, 0 AS Endereco, 0 AS Inicial, SUM(DebitoOficial) AS Debitos, SUM(CreditoOficial) AS Creditos"
                'Sql &= " FROM Razao"
                'Sql &= " WHERE Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%'  And (Movimento_Id BETWEEN '" & Format(CDate(DataInicial), "yyyy/MM/dd") & "' AND '" & Format(CDate(DataFinal), "yyyy/MM/dd") & "')"

                'Sql &= " GROUP BY Conta_Id, Cliente_Id, EndCliente_Id"

                'Sql &= " ) as Consulta"
                'Sql &= " Group By Conta, Cliente, Endereco"
                'Sql &= " Order By Conta, Cliente, Endereco"


                ds = Banco.ConsultaDataSet(Sql, "SaldosPeriodicos")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        If dr("Inicial") <> 0 Or dr("Debitos") <> 0 Or dr("Creditos") <> 0 Then
                            With dr
                                linha = "|I155"

                                If .Item("Cliente") <> "" Then
                                    linha &= "|" & .Item("Conta") & .Item("Cliente") & .Item("EndCliente_Id")
                                Else
                                    linha &= "|" & .Item("Conta")
                                End If

                                linha &= "|"

                                If .Item("Inicial") < 0 Then
                                    linha &= "|" & .Item("Inicial") * -1
                                    linha &= "|C"
                                Else
                                    linha &= "|" & .Item("Inicial")
                                    linha &= "|D"
                                End If

                                linha &= "|" & .Item("Debitos")
                                linha &= "|" & .Item("Creditos")

                                Saldo = .Item("Inicial") + (.Item("Debitos") - .Item("Creditos"))

                                If Saldo < 0 Then
                                    linha &= "|" & Saldo * -1
                                    linha &= "|C"
                                Else
                                    linha &= "|" & Saldo
                                    linha &= "|D"
                                End If

                                linha &= "|"
                            End With

                            'strm.WriteLine(linha)
                            ConteudoDoArquivo.AppendLine(linha)
                            RegistroI155 += 1
                            RegistroGeral += 1
                        End If
                    Next

                End If

            Next i


            'Registro I200 - Lançamento Contábil
            Dim Ordem As Integer = 0

            Sql = "  SELECT  Movimento_Id, Lote_Id, 0 as Sequencia_Id, '' as Conta_Id, '' as Cliente_Id, 0 as EndCliente_Id,  " & vbCrLf &
                  "           0 as Custo, Sum(DebitoOficial) AS DebitoOficial, Sum(CreditoOficial) as CreditoOficial, '' as Historico " & vbCrLf &
                  " FROM Razao" & vbCrLf &
                  " WHERE Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%' And Movimento_ID between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf
            'Sql &= " And Lote_Id <> 7500"

            Sql &= " Group by Movimento_Id, Lote_Id" & vbCrLf &
                   "                                                            " & vbCrLf &
                   " Union All" & vbCrLf &
                   "                                                            " & vbCrLf &
                   " SELECT  Movimento_Id, Lote_Id, Sequencia_Id, Conta_Id, Cliente_Id, EndCliente_Id,  " & vbCrLf &
                   " Custo, DebitoOficial, CreditoOficial, ISNULL(Historico, '')" & vbCrLf &
                   " FROM Razao" & vbCrLf &
                   " WHERE Left(Empresa_Id, 8) like '" & Left(Empresa(0), 8) & "%' And Movimento_ID between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf
            'Sql &= " And Lote_Id <> 7500"

            Sql &= " ORDER BY Movimento_Id, Lote_Id, Sequencia_Id"

            ds = Banco.ConsultaDataSet(Sql, "LancamentoContabil")

            Dim Lote As Integer = 0

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    With dr
                        If .Item("Sequencia_Id") = 0 Then
                            linha = "|I200"
                            'Lote += 1
                            'linha &= "|" & Lote
                            'Lote += 1
                            linha &= "|" & Format(CDate(.Item("Movimento_Id")), "ddMMyyyy").Replace("/", "") & "-" & .Item("Lote_Id")
                            linha &= "|" & Format(CDate(.Item("Movimento_Id")), "ddMMyyyy")  'Data Inicial
                            linha &= "|" & .Item("DebitoOficial")
                            If .item("Lote_Id") = "7500" Then
                                'If cbTrimestral.Checked = True Then
                                linha &= "|E"
                            Else
                                linha &= "|N"
                            End If
                            'Else
                            'linha &= "|N"
                            'End If

                            linha &= "|"
                            linha &= "|"

                            'strm.WriteLine(linha)
                            ConteudoDoArquivo.AppendLine(linha)
                            RegistroI200 += 1
                            RegistroGeral += 1
                        Else
                            'Registro I250 - Partidas Do Lancamento
                            linha = "|I250"

                            If Trim(.Item("Cliente_Id")) <> "" Then
                                linha &= "|" & .Item("Conta_Id") & .Item("Cliente_Id") & .Item("EndCliente_Id")
                            Else
                                linha &= "|" & .Item("Conta_Id")
                            End If

                            'If .Item("CentroDeCustos") <> 0 Then
                            '    linha &= "|" & .Item("CentroDeCustos")
                            'Else
                            linha &= "|"
                            'End If

                            If .Item("DebitoOficial") > 0 Then
                                linha &= "|" & .Item("DebitoOficial")
                                linha &= "|D"
                            Else
                                linha &= "|" & .Item("CreditoOficial")
                                linha &= "|C"
                            End If

                            linha &= "|" '& Lote

                            'linha &= "|" & Format(CDate(.Item("Movimento_Id")), "MMdd") & String.Format("{0:D4}", .Item("Lote_Id")) & String.Format("{0:D5}", .Item("Sequencia_Id"))
                            linha &= "|"

                            linha &= "|" & Funcoes.RemoveAllEnterKey(RTrim(.Item("Historico"))).Replace("|", "").Trim
                            linha &= "|"
                            linha &= "|"

                            'strm.WriteLine(linha)
                            ConteudoDoArquivo.AppendLine(linha)
                            RegistroI250 += 1
                            RegistroGeral += 1
                        End If
                    End With
                Next
            End If

            '--- Bloco de Apuracao Trimestral

            If cbTrimestral.Checked = True Then


                Ciclo = 4

                For i As Integer = 1 To Ciclo

                    '' AJUSTE PARA EXECULTAR APENAS NO TRIMESTRE DO PERIDO SELECIONADO
                    If Ciclo = 4 And i = 1 Then
                        DataInicial = "01/01" & " / " & ano
                        DataFinal = "31/03/" & ano
                    ElseIf Ciclo = 4 And i = 2 Then
                        DataInicial = "01/04" & " / " & ano
                        DataFinal = "30/06/" & ano
                    ElseIf Ciclo = 4 And i = 3 Then
                        DataInicial = "01/07" & " / " & ano
                        DataFinal = "30/09/" & ano
                    ElseIf Ciclo = 4 And i = 4 Then
                        DataInicial = "01/10" & " / " & ano
                        DataFinal = "31/12/" & ano
                    End If

                    ''TRATANTO OXFIBER - FURLAN - 28-07-2020
                    'If (Left(Empresa(0), 8) = "38388314" AndAlso ano = 2021 AndAlso i > 3) Then
                    '    Continue For
                    'End If

                    ''TRATANTO NUTRI - FURLAN - 26-08-2022
                    'If (Left(Empresa(0), 8) = "05366261" AndAlso ano = 2020 AndAlso i < 8) Then
                    '    Continue For
                    'End If

                    '''TRATANTO BAXI - FURLAN - 27-07-2022
                    ''If (Left(Empresa(0), 8) = "40938762" AndAlso ano = 2021 AndAlso i > 1) Then
                    ''    Continue For
                    ''End If

                    If (Left(Empresa(0), 8) = "24450490" AndAlso ano = 2024 AndAlso i < 3) Then
                        Continue For
                    End If

                    linha = "|I350"
                    linha &= "|" & Format(CDate(DataFinal), "ddMMyyyy")
                    linha &= "|"
                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroI350 += 1
                    RegistroGeral += 1


                    'Registro I355 - Detalhe dos Saldos Periodicos

                    'Sql = " Select Conta, Sum(Saldo) as Saldo, Referencial from "
                    'Sql &= " ( "
                    'Sql &= " SELECT  Consulta.Conta,   Consulta.Saldo, Consulta.Referencial "
                    'Sql &= " FROM"
                    'Sql &= " ("
                    'Sql &= " SELECT  Consulta.Conta,    Consulta.Saldo, PlanoDeContasXPlanoReferencial.Referencial FROM "
                    'Sql &= " (SELECT LEFT(Empresa_Id, 8) AS Empresa,    Left(Conta_Id, 7)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo   FROM Razao    "
                    'Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) And  Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%')   "
                    'Sql &= " And (Movimento_Id BETWEEN '" & Format(CDate(DataInicial), "yyyy-MM-dd") & "' AND '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') AND (Lote_Id < 7500)     "
                    'Sql &= " GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)          ) AS Consulta LEFT OUTER JOIN   "
                    'Sql &= " PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id"
                    'Sql &= " Where (Referencial Is Not null)"
                    'Sql &= " Union"
                    'Sql &= " SELECT  Consulta.Conta,   Consulta.Saldo, PlanoDeContasXPlanoReferencial.Referencial FROM        "
                    'Sql &= " (SELECT  LEFT(Empresa_Id, 8) AS Empresa,   Conta_Id  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo FROM Razao          "
                    'Sql &= " WHERE   (LEFT(Conta_Id, 1) IN ('3', '4')) And  Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%')   "
                    'Sql &= " And (Movimento_Id BETWEEN '" & Format(CDate(DataInicial), "yyyy-MM-dd") & "' AND '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') AND (Lote_Id < 7500) "
                    'Sql &= " Group by LEFT(Empresa_Id, 8), Conta_Id) AS Consulta LEFT OUTER JOIN  "
                    'Sql &= " PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 9) = PlanoDeContasXPlanoReferencial.Conta_Id"
                    'Sql &= " Where (Referencial Is Not null) ) as Consulta "
                    'Sql &= " Group by Conta, Saldo, Referencial) as Consulta"
                    'Sql &= " Group by Conta, Saldo, Referencial"

                    Sql = "  SELECT LEFT(Empresa_Id, 8) AS Empresa, Conta_Id AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo"
                    Sql &= " FROM Razao"

                    'Tirar conta de resultado - Furlan - 03/07/2020
                    If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                        Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) "
                    Else
                        Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) "
                    End If

                    Sql &= "   And  Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%')   "
                    Sql &= "   And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') "
                    Sql &= "   And (Lote_Id <> 7500)   "
                    'Sql &= "   And (Lote_Id <>  3)     "

                    Sql &= " GROUP BY LEFT(Empresa_Id, 8), Conta_Id"
                    Sql &= " Order By  Conta_Id"

                    ds = Banco.ConsultaDataSet(Sql, "SaldosPeriodicos")

                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each dr In ds.Tables(0).Rows
                            'If dr("Saldo") <> 0 Then
                            With dr
                                linha = "|I355"
                                linha &= "|" & .Item("Conta")
                                linha &= "|"

                                If .Item("Saldo") < 0 Then
                                    linha &= "|" & .Item("Saldo") * -1
                                    linha &= "|C"
                                Else
                                    linha &= "|" & .Item("Saldo")
                                    linha &= "|D"
                                End If
                                linha &= "|"
                            End With

                            'strm.WriteLine(linha)
                            ConteudoDoArquivo.AppendLine(linha)
                            RegistroI355 += 1
                            RegistroGeral += 1
                            ' End If
                        Next
                    Else

                        If (Left(Empresa(0), 8) = "38388314" AndAlso ano = 2021 AndAlso i = 3) Then
                            linha = "|I355|402020105||0,00|D|"

                            ConteudoDoArquivo.AppendLine(linha)
                            RegistroI355 += 1
                            RegistroGeral += 1
                        End If

                    End If

                Next i
            End If

            '-----Apuração Anual --------------------------------------------------------

            If cbTrimestral.Checked = False Then


                DataInicial = "01/01" & " / " & ano
                DataFinal = "31/12/" & ano


                linha = "|I350"
                linha &= "|" & Format(CDate(DataFinal), "ddMMyyyy")
                linha &= "|"
                ConteudoDoArquivo.AppendLine(linha)
                RegistroI350 += 1
                RegistroGeral += 1

                'Registro I355 - Apuração Anual

                'Sql = " Select Conta, Sum(Saldo) as Saldo, Referencial from "
                'Sql &= " ( "
                'Sql &= " SELECT  Consulta.Conta,   Consulta.Saldo, Consulta.Referencial "
                'Sql &= " FROM"
                'Sql &= " ("
                'Sql &= " SELECT  Consulta.Conta,    Consulta.Saldo, PlanoDeContasXPlanoReferencial.Referencial FROM "
                'Sql &= " (SELECT LEFT(Empresa_Id, 8) AS Empresa,    Left(Conta_Id, 7)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo   FROM Razao    "
                'Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) And  Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%')   "
                'Sql &= " And (Movimento_Id BETWEEN '" & Format(CDate(DataInicial), "yyyy-MM-dd") & "' AND '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') AND (Lote_Id < 7500)     "
                'Sql &= " GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)          ) AS Consulta LEFT OUTER JOIN   "
                'Sql &= " PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id"
                'Sql &= " Where (Referencial Is Not null)"
                'Sql &= " Union"
                'Sql &= " SELECT  Consulta.Conta,   Consulta.Saldo, PlanoDeContasXPlanoReferencial.Referencial FROM        "
                'Sql &= " (SELECT  LEFT(Empresa_Id, 8) AS Empresa,   Conta_Id  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo FROM Razao          "
                'Sql &= " WHERE   (LEFT(Conta_Id, 1) IN ('3', '4')) And  Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%')   "
                'Sql &= " And (Movimento_Id BETWEEN '" & Format(CDate(DataInicial), "yyyy-MM-dd") & "' AND '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') AND (Lote_Id < 7500) "
                'Sql &= " Group by LEFT(Empresa_Id, 8), Conta_Id) AS Consulta LEFT OUTER JOIN  "
                'Sql &= " PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 9) = PlanoDeContasXPlanoReferencial.Conta_Id"
                'Sql &= " Where (Referencial Is Not null) ) as Consulta "
                'Sql &= " Group by Conta, Saldo, Referencial) as Consulta"
                'Sql &= " Group by Conta, Saldo, Referencial"

                Sql = "  SELECT LEFT(Empresa_Id, 8) AS Empresa, Conta_Id AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo"
                Sql &= " FROM Razao"

                'Tirar conta de resultado - Furlan - 03/07/2020
                If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                    Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) "
                Else
                    Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) "
                End If

                Sql &= "   And  Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%')   "
                Sql &= "   And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') "
                Sql &= "   And (Lote_Id <> 7500)   "
                'Sql &= "   And (Lote_Id <>  3)     "

                Sql &= " GROUP BY LEFT(Empresa_Id, 8), Conta_Id"
                Sql &= " Order By  Conta_Id"



                ds = Banco.ConsultaDataSet(Sql, "SaldosPeriodicos")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        'If dr("Saldo") <> 0 Then
                        With dr
                            linha = "|I355"
                            linha &= "|" & .Item("Conta")
                            linha &= "|"

                            If .Item("Saldo") < 0 Then
                                linha &= "|" & .Item("Saldo") * -1
                                linha &= "|C"
                            Else
                                linha &= "|" & .Item("Saldo")
                                linha &= "|D"
                            End If
                            linha &= "|"
                        End With

                        ConteudoDoArquivo.AppendLine(linha)
                        RegistroI355 += 1
                        RegistroGeral += 1

                    Next
                End If
            End If

            'Else
            'DataInicial = "01/01" & " / " & ano
            'DataFinal = "31/12/" & ano

            'linha = "|I350"
            'linha &= "|" & Format(CDate(DataFinal), "ddMMyyyy")
            'linha &= "|"
            'strm.WriteLine(linha)
            'RegistroI350 += 1
            'RegistroGeral += 1


            'Sql = "  SELECT LEFT(Empresa_Id, 8) AS Empresa, Conta_Id AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo"
            'Sql &= " FROM Razao"
            'Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) "
            'Sql &= "   And  Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%')   "
            'Sql &= "   And (Movimento_Id BETWEEN '" & Format(CDate(DataInicial), "yyyy-MM-dd") & "' AND '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') "
            'Sql &= "   And (Lote_Id < 7500)     "
            'Sql &= " GROUP BY LEFT(Empresa_Id, 8), Conta_Id"

            'ds = Banco.ConsultaDataSet(Sql, "SaldosPeriodicos")

            'If ds.Tables(0).Rows.Count > 0 Then
            '    For Each dr In ds.Tables(0).Rows
            '        'If dr("Saldo") <> 0 Then
            '        With dr
            '            linha = "|I355"
            '            linha &= "|" & .Item("Conta")
            '            linha &= "|"

            '            If .Item("Saldo") < 0 Then
            '                linha &= "|" & .Item("Saldo") * -1
            '                linha &= "|C"
            '            Else
            '                linha &= "|" & .Item("Saldo")
            '                linha &= "|D"
            '            End If
            '            linha &= "|"
            '        End With

            '        strm.WriteLine(linha)
            '        RegistroI355 += 1
            '        RegistroGeral += 1
            '        'End If
            '    Next
            'End If


            'End If


            ' Registro I990  - Encerramento do Bloco I
            RegistroI990 += 1

            linha = "|I990"
            linha &= "|" & RegistroI001 + RegistroI010 + RegistroI030 + RegistroI050 + RegistroI051 + RegistroI052 + RegistroI100 + RegistroI150 + RegistroI155 + RegistroI200 + RegistroI250 + RegistroI350 + RegistroI355 + RegistroI990
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            RegistroGeral += 1

            ' Registro J001  - Abertura do Bloco J

            linha = "|J001"
            linha &= "|0"
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            RegistroJ001 += 1
            RegistroGeral += 1

            ' Registro J005  - Demonstrações Contábeis

            If cbTrimestral.Checked = True Then
                Ciclo = 4
            Else
                Ciclo = 1
            End If

            For i As Integer = 1 To Ciclo

                If i = 1 And Ciclo = 1 Then
                    DataInicial = "01/01" & " / " & ano
                    DataFinal = "31/12/" & ano
                ElseIf Ciclo = 4 And i = 1 Then
                    DataInicial = "01/01" & " / " & ano
                    DataFinal = "31/03/" & ano
                ElseIf Ciclo = 4 And i = 2 Then
                    DataInicial = "01/04" & " / " & ano
                    DataFinal = "30/06/" & ano
                ElseIf Ciclo = 4 And i = 3 Then
                    DataInicial = "01/07" & " / " & ano
                    DataFinal = "30/09/" & ano
                ElseIf Ciclo = 4 And i = 4 Then
                    DataInicial = "01/10" & " / " & ano
                    DataFinal = "31/12/" & ano
                End If

                'TRATANTO OXFIBER - FURLAN - 28-07-2020
                If (Left(Empresa(0), 8) = "38388314" AndAlso ano = 2021 AndAlso i = 4) Then
                    Continue For
                End If

                linha = "|J005"
                linha &= "|" & Format(CDate(DataInicial), "ddMMyyyy")                'Data Inicial
                linha &= "|" & Format(CDate(DataFinal), "ddMMyyyy")               'Data Final
                linha &= "|1"
                linha &= "|"
                linha &= "|"

                'strm.WriteLine(linha)
                ConteudoDoArquivo.AppendLine(linha)
                RegistroJ005 += 1
                RegistroGeral += 1

                'Registro J100 - Balanço Patrimonial

                'TRATANTO OXFIBER - FURLAN - 28-07-2020
                If Left(Empresa(0), 8) = "38388314" AndAlso i = 3 AndAlso ano = 2021 Then
                    linha = "|J100|1|T|1||A|A T I V O |0,00|D|0,00|D||"

                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroJ100 += 1
                    RegistroGeral += 1

                    linha = "|J100|101|T|2|1|A|ATIVO CIRCULANTE |0,00|D|0,00|D||"

                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroJ100 += 1
                    RegistroGeral += 1

                    linha = "|J100|10101|T|3|101|A|CAIXA E EQUIVALENTES DE CAIXA|0,00|D|0,00|D||"

                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroJ100 += 1
                    RegistroGeral += 1

                    linha = "|J100|1010104|D|4|10101|A|CAIXA|0,00|D|0,00|D||"

                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroJ100 += 1
                    RegistroGeral += 1

                End If

                Sql = " Select Consulta.Empresa, Consulta.Conta, PlanoDeContas.Titulo, Sum(Consulta.Saldo) as Saldo, Sum(Consulta.SaldoInicial) as SaldoInicial " & vbCrLf &
                      "     From" & vbCrLf &
                      "             (" & vbCrLf &
                      "                 SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 1)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo, 0 as SaldoInicial " & vbCrLf &
                      "                     FROM Razao" & vbCrLf &
                      "                     WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') And (Movimento_Id <= '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 9000) " & vbCrLf &
                      "                         And (LEFT(Conta_Id, 3) not in (205))" & vbCrLf &
                      "                 GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 1)" & vbCrLf &
                      "                                                                             " & vbCrLf &
                      "                 Union" & vbCrLf &
                      "                                                                             " & vbCrLf &
                      "                 SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 3)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo, 0 as SaldoInicial" & vbCrLf &
                      "                     FROM Razao" & vbCrLf &
                      "                     WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') And (Movimento_Id <= '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 9000) " & vbCrLf &
                      "                         And (LEFT(Conta_Id, 3) not in (205))" & vbCrLf &
                      "                 GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 3)" & vbCrLf &
                      "                                                                         " & vbCrLf &
                      "                 Union" & vbCrLf &
                      "                                                             " & vbCrLf &
                      "                 SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 5)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo, 0 as SaldoInicial" & vbCrLf &
                      "                     FROM Razao" & vbCrLf &
                      "                     WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') And (Movimento_Id <= '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 9000) " & vbCrLf &
                      "                         And (LEFT(Conta_Id, 3) not in (205))" & vbCrLf &
                      "                 GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 5)" & vbCrLf &
                      "                                                                             " & vbCrLf &
                      "                 Union " & vbCrLf &
                      "                                                                 " & vbCrLf &
                      "                 SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 7)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo, 0 as SaldoInicial" & vbCrLf &
                      "                     FROM Razao" & vbCrLf &
                      "                     WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') And (Movimento_Id <= '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 9000) " & vbCrLf &
                      "                         And (LEFT(Conta_Id, 3) not in (205))" & vbCrLf &
                      "                 GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)" & vbCrLf &
                      "                " & vbCrLf &
                      "                 Union " & vbCrLf &
                      "                 SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 1)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial " & vbCrLf &
                      "                     FROM Razao" & vbCrLf &
                      "                     WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000) " & vbCrLf &
                      "                         And (LEFT(Conta_Id, 3) not in (205))" & vbCrLf &
                      "                 GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 1)" & vbCrLf &
                      "                                                                             " & vbCrLf &
                      "                 Union" & vbCrLf &
                      "                                                                             " & vbCrLf &
                      "                 SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 3)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                      "                     FROM Razao" & vbCrLf &
                      "                     WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000) " & vbCrLf &
                      "                         And (LEFT(Conta_Id, 3) not in (205))" & vbCrLf &
                      "                 GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 3)" & vbCrLf &
                      "                                                                         " & vbCrLf &
                      "                 Union" & vbCrLf &
                      "                                                             " & vbCrLf &
                      "                 SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 5)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                      "                     FROM Razao" & vbCrLf &
                      "                     WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000) " & vbCrLf &
                      "                         And (LEFT(Conta_Id, 3) not in (205))" & vbCrLf &
                      "                 GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 5)" & vbCrLf &
                      "                                                                             " & vbCrLf &
                      "                 Union " & vbCrLf &
                      "                                                                 " & vbCrLf &
                      "                 SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 7)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                      "                     FROM Razao" & vbCrLf &
                      "                     WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000) " & vbCrLf &
                      "                         And (LEFT(Conta_Id, 3) not in ( 205))" & vbCrLf &
                      "                 GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)" & vbCrLf &
                      "         ) AS Consulta" & vbCrLf &
                      "         INNER JOIN PlanoDeContas" & vbCrLf &
                      "              ON Consulta.Conta = PlanoDeContas.Conta_Id " & vbCrLf &
                      "         Group By Consulta.Empresa, Consulta.Conta, PlanoDeContas.Titulo" & vbCrLf &
                      "         Order By Consulta.Conta" & vbCrLf


                ds = Banco.ConsultaDataSet(Sql, "BalancoPatrimonial")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        'If (dr("Saldo") <> 0 Or dr("SaldoInicial") <> 0 Or dr("Conta").ToString().Length <= 5) OrElse (Left(Empresa(0), 8) = "38388314" AndAlso i = 3 AndAlso ano = 2021) Then
                        'End If
                        '' 
                        If (dr("Saldo") <> 0 Or dr("SaldoInicial") <> 0 Or dr("Conta").ToString().Length <= 5) OrElse (Left(Empresa(0), 8) = "38388314" AndAlso i = 3 AndAlso ano = 2021) Then
                            With dr
                                'Campo 1 - Tipo de Registro J100
                                linha = "|J100"

                                'Campo 2 - Código de Aglutinação das Contas
                                linha &= "|" & .Item("Conta")

                                'Campo 3 - Indicado do Tipo de Código de Aplutinação 
                                If Len(.Item("Conta")) > 5 Then
                                    linha &= "|D"
                                ElseIf Len(.Item("Conta")) = 5 Then
                                    linha &= "|T"
                                ElseIf Len(.Item("Conta")) = 3 Then
                                    linha &= "|T"
                                ElseIf Len(.Item("Conta")) = 2 Then
                                    linha &= "|T"
                                ElseIf Len(.Item("Conta")) = 1 Then
                                    linha &= "|T"
                                End If

                                'Campo 4 - Nível de Código de Aglutinação
                                If Len(.Item("Conta")) > 7 Then
                                    linha &= "|5"
                                ElseIf Len(.Item("Conta")) = 7 Then
                                    linha &= "|4"
                                ElseIf Len(.Item("Conta")) = 5 Then
                                    linha &= "|3"
                                ElseIf Len(.Item("Conta")) = 3 Then
                                    linha &= "|2"
                                ElseIf Len(.Item("Conta")) = 1 Then
                                    linha &= "|1"
                                End If

                                'Campo 5 - Código de Aglutinação de nível superior 
                                If Len(.Item("Conta")) > 7 Then
                                    linha &= "|" & Left(.Item("Conta"), 7)
                                ElseIf Len(.Item("Conta")) = 7 Then
                                    linha &= "|" & Left(.Item("Conta"), 5)
                                ElseIf Len(.Item("Conta")) = 5 Then
                                    linha &= "|" & Left(.Item("Conta"), 3)
                                ElseIf Len(.Item("Conta")) = 3 Then
                                    linha &= "|" & Left(.Item("Conta"), 1)
                                ElseIf Len(.Item("Conta")) = 1 Then
                                    linha &= "|"
                                End If

                                'Campo 6 - Indicador do Grupo do Balanço
                                'linha &= "|" & Left(.Item("Conta"), 1)
                                If Left(.Item("Conta"), 1) = 1 Then
                                    linha &= "|A"
                                Else
                                    linha &= "|P"
                                End If

                                'Campo 7 - Descrição do código de Aglutinação
                                linha &= "|" & .Item("Titulo")

                                'Campo 8 - Valor Inicial do Código de Aglutinação - Campo 9 - Indicador da Situação do Saldo informado no campo anterior
                                If .Item("SaldoInicial") < 0 Then
                                    linha &= "|" & .Item("SaldoInicial") * -1
                                    linha &= "|C"
                                Else
                                    linha &= "|" & .Item("SaldoInicial")
                                    linha &= "|D"
                                End If

                                'Campo 10 - Valor Final do Código de Aglutinação - Campo 11 - Indicador da Situação do Saldo Informado
                                If .Item("Saldo") < 0 Then
                                    linha &= "|" & .Item("Saldo") * -1
                                    linha &= "|C"
                                Else
                                    linha &= "|" & .Item("Saldo")
                                    linha &= "|D"
                                End If

                                'Campo 12 - Referencia a Numeração das notas Explicativas
                                If Format(CDate(txtDataFinal.Text), "yyyy") > 2016 Then linha &= "|" 'Acrescentado em 21/05/2018 - Furlan

                                'Campo 13
                                linha &= "|"

                            End With

                            'strm.WriteLine(linha)
                            ConteudoDoArquivo.AppendLine(linha)
                            RegistroJ100 += 1
                            RegistroGeral += 1
                        End If
                    Next
                End If

                'Registro J150 - DRE Demonstrativo de Resultado

                If Format(CDate(txtDataFinal.Text), "yyyy") > 2018 Then

                    'A partir de 2019 deve vir apenas a conta a de Resultando nível 1 - Furlan - 03/07/2020
                    Sql = " Select Distinct Consulta.Empresa, Consulta.Conta, PlanoDeContas.Titulo, Sum(Consulta.Saldo) as Saldo, Sum(Consulta.SaldoInicial) as SaldoInicial " & vbCrLf &
                          "     From" & vbCrLf &
                          "         (SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 1)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo, 0 as SaldoInicial" & vbCrLf &
                          "                    FROM Razao" & vbCrLf &
                          "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                          "                    And (Movimento_Id between '" & DataInicial.ToSqlDate() & "' and '" & DataFinal.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                          "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 1)" & vbCrLf &
                          "             Union" & vbCrLf &
                          "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 3)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo, 0 as SaldoInicial" & vbCrLf &
                          "                FROM Razao" & vbCrLf &
                          "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                          "                    And (Movimento_Id between '" & DataInicial.ToSqlDate() & "' and '" & DataFinal.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                          "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 3)" & vbCrLf &
                          "             Union" & vbCrLf &
                          "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 5)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo, 0 as SaldoInicial" & vbCrLf &
                          "                  FROM Razao" & vbCrLf &
                          "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                          "                    And (Movimento_Id between '" & DataInicial.ToSqlDate() & "' and '" & DataFinal.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                          "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 5)" & vbCrLf &
                          "             Union " & vbCrLf &
                          "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 7)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo, 0 as SaldoInicial" & vbCrLf &
                          "                 FROM Razao" & vbCrLf &
                          "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                          "                    And (Movimento_Id between '" & DataInicial.ToSqlDate() & "' and '" & DataFinal.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                          "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)" & vbCrLf

                    ' Saldo Exercicio Anterior
                    If Format(CDate(txtDataFinal.Text), "yyyy") > 2019 Then
                        Dim anoAntIni As String
                        Dim anoAntFim As String

                        If i = 1 And Ciclo = 1 Then
                            anoAntIni = "01/01" & " / " & (ano - 1)
                            anoAntFim = "31/12" & " / " & (ano - 1)
                        ElseIf Ciclo = 4 And i = 1 Then
                            anoAntIni = "01/01" & " / " & (ano - 1)
                            anoAntFim = "31/12" & " / " & (ano - 1)
                        ElseIf Ciclo = 4 And i = 2 Then
                            anoAntIni = "01/01" & " / " & ano
                            anoAntFim = "31/03/" & ano
                        ElseIf Ciclo = 4 And i = 3 Then
                            anoAntIni = "01/04" & " / " & ano
                            anoAntFim = "30/06/" & ano
                        ElseIf Ciclo = 4 And i = 4 Then
                            anoAntIni = "01/07" & " / " & ano
                            anoAntFim = "30/09/" & ano
                        End If

                        Sql &= "             Union " & vbCrLf &
                              "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 1)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                              "                    FROM Razao" & vbCrLf &
                              "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                              "                    And (Movimento_Id between '" & anoAntIni.ToSqlDate() & "' and '" & anoAntFim.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                              "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 1)" & vbCrLf &
                              "             Union" & vbCrLf &
                              "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 3)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                              "                FROM Razao" & vbCrLf &
                              "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                              "                    And (Movimento_Id between '" & anoAntIni.ToSqlDate() & "' and '" & anoAntFim.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                              "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 3)" & vbCrLf &
                              "             Union" & vbCrLf &
                              "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 5)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                              "                  FROM Razao" & vbCrLf &
                              "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                              "                    And (Movimento_Id between '" & anoAntIni.ToSqlDate() & "' and '" & anoAntFim.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                              "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 5)" & vbCrLf &
                              "             Union " & vbCrLf &
                              "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 7)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                              "                 FROM Razao" & vbCrLf &
                              "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                              "                    And (Movimento_Id between '" & anoAntIni.ToSqlDate() & "' and '" & anoAntFim.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                              "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)" & vbCrLf &
                              "         ) AS Consulta" & vbCrLf
                    Else
                        Sql &= "             Union " & vbCrLf &
                              "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 1)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                              "                    FROM Razao" & vbCrLf &
                              "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                              "                    And (Movimento_Id < '" & DataInicial.ToSqlDate() & "')" & vbCrLf &
                              "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 1)" & vbCrLf &
                              "             Union" & vbCrLf &
                              "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 3)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                              "                FROM Razao" & vbCrLf &
                              "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                              "                    And (Movimento_Id < '" & DataInicial.ToSqlDate() & "')" & vbCrLf &
                              "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 3)" & vbCrLf &
                              "             Union" & vbCrLf &
                              "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 5)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                              "                  FROM Razao" & vbCrLf &
                              "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                              "                    And (Movimento_Id < '" & DataInicial.ToSqlDate() & "')" & vbCrLf &
                              "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 5)" & vbCrLf &
                              "             Union " & vbCrLf &
                              "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 7)  AS Conta, 0 AS Saldo, SUM(DebitoOficial - CreditoOficial) as SaldoInicial" & vbCrLf &
                              "                 FROM Razao" & vbCrLf &
                              "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                              "                    And (Movimento_Id < '" & DataInicial.ToSqlDate() & "')" & vbCrLf &
                              "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)" & vbCrLf &
                              "         ) AS Consulta" & vbCrLf

                    End If
                Else
                    Sql = " Select Distinct Consulta.Empresa, Consulta.Conta, PlanoDeContas.Titulo, Sum(Consulta.Saldo) as Saldo " & vbCrLf &
                          "     From" & vbCrLf &
                          "         (SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 1)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo " & vbCrLf &
                          "                    FROM Razao" & vbCrLf &
                          "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                          "                    And (Movimento_Id between '" & DataInicial.ToSqlDate() & "' and '" & DataFinal.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                          "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 1)" & vbCrLf &
                          "             Union" & vbCrLf &
                          "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 3)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo" & vbCrLf &
                          "                FROM Razao" & vbCrLf &
                          "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                          "                    And (Movimento_Id between '" & DataInicial.ToSqlDate() & "' and '" & DataFinal.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                          "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 3)" & vbCrLf &
                          "             Union" & vbCrLf &
                          "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 5)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo" & vbCrLf &
                          "                  FROM Razao" & vbCrLf &
                          "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                          "                    And (Movimento_Id between '" & DataInicial.ToSqlDate() & "' and '" & DataFinal.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                          "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 5)" & vbCrLf &
                          "             Union " & vbCrLf &
                          "             SELECT  LEFT(Empresa_Id, 8) AS Empresa, Left(Conta_Id, 7)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo" & vbCrLf &
                          "                 FROM Razao" & vbCrLf &
                          "             WHERE  (LEFT(Conta_Id, 1) IN ('3', '4', '5')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') " & vbCrLf &
                          "                    And (Movimento_Id between '" & DataInicial.ToSqlDate() & "' and '" & DataFinal.ToSqlDate() & "') And Lote_Id  < 7500" & vbCrLf &
                          "             GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)" & vbCrLf &
                          "         ) AS Consulta" & vbCrLf
                End If

                'Sql &= " Union"

                'Sql &= " SELECT  LEFT(Empresa_Id, 8) AS Empresa, '5' AS Conta, SUM(CreditoOficial - DebitoOficial) AS Saldo "
                'Sql &= "        FROM Razao"
                'Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') "
                'Sql &= "        And (Movimento_Id between '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "' and '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') And Lote_Id  < 7500"
                'Sql &= " GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 1)"

                'Sql &= " Union"

                'Sql &= " SELECT  LEFT(Empresa_Id, 8) AS Empresa,'501' AS Conta, SUM(CreditoOficial - DebitoOficial) AS Saldo"
                'Sql &= "    FROM Razao"
                'Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') "
                'Sql &= "        And (Movimento_Id between '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "' and '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') And Lote_Id  < 7500"
                'Sql &= " GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 3)"

                'Sql &= " Union"

                'Sql &= " SELECT  LEFT(Empresa_Id, 8) AS Empresa, '50101' AS Conta, SUM(CreditoOficial - DebitoOficial) AS Saldo"
                'Sql &= "      FROM Razao"
                'Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') "
                'Sql &= "        And (Movimento_Id between '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "' and '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') And Lote_Id  < 7500"
                'Sql &= " GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 5)"

                'Sql &= " Union "

                'Sql &= " SELECT  LEFT(Empresa_Id, 8) AS Empresa, '5010101' AS Conta, SUM(CreditoOficial - DebitoOficial) AS Saldo"
                'Sql &= "     FROM Razao"
                'Sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) AND Empresa_Id  Like ('" & Left(Empresa(0), 8) & "%') "
                'Sql &= "        And (Movimento_Id between '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "' and '" & Format(CDate(DataFinal), "yyyy-MM-dd") & "') And Lote_Id  < 7500"
                'Sql &= " GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)"


                Sql &= "  LEFT OUTER JOIN PlanoDeContas ON Consulta.Conta = left(PlanoDeContas.Conta_Id, 9)  " & vbCrLf &
                       "    Group by Consulta.Empresa, Consulta.Conta, PlanoDeContas.Titulo  " & vbCrLf &
                       "    Order By Consulta.Conta" & vbCrLf
                '"    Having Sum(Consulta.Saldo) <> 0" & vbCrLf & _
                '"    Order By Consulta.Conta" & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "BalancoPatrimonial")
                Dim Resultado As String

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows

                        If CDate(DataFinal).Year = 2018 AndAlso Left(dr("Conta"), 7) = "5010101" Then
                            dr("Saldo") = 0
                        ElseIf CDate(DataFinal).Year >= 2019 AndAlso Left(dr("Conta"), 1) = "5" Then 'Acrescentei >= para à partir de 2019
                            'dr("Saldo") = 0
                            'dr("SaldoInicial") = 0
                        End If

                        If dr("Saldo") <> 0 _
                            OrElse dr("SaldoInicial") <> 0 _
                            OrElse (CDate(DataFinal).Year = 2018 And (((Left(Empresa(0), 8) = "03189063" And dr("Conta") = "40102")) Or (Left(Empresa(0), 8) = "07577827" And (dr("Conta") = "403") Or dr("Conta") = "40301"))) _
                            OrElse (CDate(DataFinal).Year >= 2019 And dr("Conta") = "40102") Then 'Acrescentei >= para à partir de 2019
                            With dr
                                linha = "|J150" 'Campo 1 - Texto fixo contendo J150

                                Ordem += 1

                                If Format(CDate(txtDataFinal.Text), "yyyy") > 2018 Then
                                    linha &= "|" & Ordem 'Campo 2 - Número de ordem da linha na visualização da demonstração. - Furlan - 12/05/2020
                                End If

                                linha &= "|" & .Item("Conta") 'Campo 3 - Código de aglutinação das linhas, atribuído pela pessoa jurídica

                                'Campo 4 - Indicador do tipo de código de aglutinação das linhas
                                If Len(.Item("Conta")) > 5 Then
                                    linha &= "|D"
                                ElseIf Len(.Item("Conta")) = 5 Then
                                    linha &= "|T"
                                ElseIf Len(.Item("Conta")) = 3 Then
                                    linha &= "|T"
                                ElseIf Len(.Item("Conta")) = 2 Then
                                    linha &= "|T"
                                ElseIf Len(.Item("Conta")) = 1 Then
                                    linha &= "|T"
                                End If

                                'Campo 5 - Nível do Código de aglutinação (mesmo conceito do plano de contas – Registro I050).
                                If Len(.Item("Conta")) > 7 Then
                                    'A partir de 2019 contas mudam o nível devido a de Resultando deve ser nível 1 - Furlan - 03/07/2020
                                    If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                                        linha &= "|6"
                                    Else
                                        linha &= "|5"
                                    End If
                                ElseIf Len(.Item("Conta")) = 7 Then
                                    'A partir de 2019 contas mudam o nível devido a de Resultando deve ser nível 1 - Furlan - 03/07/2020
                                    If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                                        linha &= "|5"
                                    Else
                                        linha &= "|4"
                                    End If
                                ElseIf Len(.Item("Conta")) = 5 Then
                                    'A partir de 2019 contas mudam o nível devido a de Resultando deve ser nível 1 - Furlan - 03/07/2020
                                    If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                                        linha &= "|4"
                                    Else
                                        linha &= "|3"
                                    End If
                                ElseIf Len(.Item("Conta")) = 3 Then
                                    'A partir de 2019 contas mudam o nível devido a de Resultando deve ser nível 1 - Furlan - 03/07/2020
                                    If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                                        linha &= "|3"
                                    Else
                                        linha &= "|2"
                                    End If
                                ElseIf Len(.Item("Conta")) = 1 Then
                                    'A partir de 2019 conta de Resultando deve ser nível 1 - Furlan - 03/07/2020
                                    If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                                        If Left(.Item("Conta"), 1) = 5 Then
                                            linha &= "|1"
                                        Else
                                            linha &= "|2"
                                        End If
                                    Else
                                        linha &= "|1"
                                    End If
                                End If


                                'Campo 6 - Código de aglutinação sintético/grupo de código de aglutinação de nível superior.
                                If Len(.Item("Conta")) > 7 Then
                                    linha &= "|" & Left(.Item("Conta"), 7)
                                ElseIf Len(.Item("Conta")) = 7 Then
                                    linha &= "|" & Left(.Item("Conta"), 5)
                                ElseIf Len(.Item("Conta")) = 5 Then
                                    linha &= "|" & Left(.Item("Conta"), 3)
                                ElseIf Len(.Item("Conta")) = 3 Then
                                    linha &= "|" & Left(.Item("Conta"), 1)
                                ElseIf Len(.Item("Conta")) = 1 Then
                                    'A partir de 2019 conta de receitas  e despesas devem amarrar na Conta de Resultado - Furlan - 03/07/2020
                                    If Format(CDate(txtDataFinal.Text), "yyyy") >= 2019 Then
                                        If Left(.Item("Conta"), 1) = 3 OrElse Left(.Item("Conta"), 1) = 4 Then
                                            linha &= "|5"
                                        Else
                                            linha &= "|"
                                        End If
                                    Else
                                        linha &= "|"
                                    End If
                                End If

                                'Campo 7 - Descrição do Código de aglutinação
                                linha &= "|" & .Item("Titulo")

                                If Format(CDate(txtDataFinal.Text), "yyyy") > 2018 Then
                                    'Campo 8 - Descrição do Código de aglutinação
                                    'Campo 9 - Descrição do Código de aglutinação
                                    If .Item("SaldoInicial") < 0 Then
                                        linha &= "|" & .Item("SaldoInicial") * -1
                                        If .Item("Conta") = "5" Or .Item("Conta") = "501" Or .Item("Conta") = "50101" Or .Item("Conta") = "5010101" Or .Item("Conta") = "501010101" Then
                                            linha &= "|D"
                                        Else
                                            linha &= "|C"
                                        End If
                                        'linha &= "|P"
                                    ElseIf .Item("SaldoInicial") > 0 Then
                                        linha &= "|" & .Item("SaldoInicial")
                                        If .Item("Conta") = "5" Or .Item("Conta") = "501" Or .Item("Conta") = "50101" Or .Item("Conta") = "5010101" Or .Item("Conta") = "501010101" Then
                                            linha &= "|C"
                                        Else
                                            linha &= "|D"
                                        End If
                                        'linha &= "|N"
                                    Else
                                        linha &= "|" & .Item("SaldoInicial")
                                        linha &= "|C"
                                    End If
                                End If


                                If .Item("Saldo") < 0 Then
                                    linha &= "|" & .Item("Saldo") * -1
                                    If .Item("Conta") = "5" Or .Item("Conta") = "501" Or .Item("Conta") = "50101" Or .Item("Conta") = "5010101" Or .Item("Conta") = "501010101" Then
                                        linha &= "|D"
                                    Else
                                        linha &= "|C"
                                    End If
                                    'linha &= "|P"
                                ElseIf .Item("Saldo") > 0 Then
                                    linha &= "|" & .Item("Saldo")
                                    If .Item("Conta") = "5" Or .Item("Conta") = "501" Or .Item("Conta") = "50101" Or .Item("Conta") = "5010101" Or .Item("Conta") = "501010101" Then
                                        linha &= "|C"
                                    Else
                                        linha &= "|D"
                                    End If
                                    'linha &= "|N"
                                Else
                                    linha &= "|" & .Item("Saldo")
                                    linha &= "|C"
                                End If

                                'linha &= "|"

                                If Left(.Item("Conta"), 1) = 3 Then
                                    linha &= "|R"
                                Else
                                    linha &= "|D"
                                End If

                                If Format(CDate(txtDataFinal.Text), "yyyy") > 2014 Then
                                    linha &= "|"        'Campo 7 - Valor se refere ao último arquivo enviado do ECD 
                                    linha &= "|"        'Campo 8 - Indicador da Natureza contábil da conta (R(-Receita), D(-Despesa), P - Subtotal Positivo, N - Subtotal Negativo
                                End If

                                'If Format(CDate(txtDataFinal.Text), "yyyy") > 2016 Then linha &= "|" 'Acrescentado em 21/05/2018 - Furlan

                                Resultado = linha
                                'strm.WriteLine(Resultado)
                                ConteudoDoArquivo.AppendLine(Resultado)
                                RegistroJ150 += 1
                                RegistroGeral += 1

                            End With

                        End If
                    Next
                Else

                    'TRATANTO OXFIBER - FURLAN - 28-07-2020
                    If (Left(Empresa(0), 8) = "38388314" AndAlso ano = 2020 AndAlso i = 3) Then

                        'linha = "|J150|1|3|T|2|5|RECEITAS|0,00|C|0,00|C|R||"

                        'Resultado = linha
                        ''strm.WriteLine(Resultado)
                        'ConteudoDoArquivo.AppendLine(Resultado)
                        'RegistroJ150 += 1
                        'RegistroGeral += 1

                        linha = "|J150|1|4|T|2|5|CUSTOS E DESPESAS|0,00|C|0,00|D|D||"

                        Resultado = linha
                        'strm.WriteLine(Resultado)
                        ConteudoDoArquivo.AppendLine(Resultado)
                        RegistroJ150 += 1
                        RegistroGeral += 1

                        linha = "|J150|2|402|T|3|4|DESPESAS|0,00|C|0,00|D|D||"

                        Resultado = linha
                        'strm.WriteLine(Resultado)
                        ConteudoDoArquivo.AppendLine(Resultado)
                        RegistroJ150 += 1
                        RegistroGeral += 1

                        linha = "|J150|3|40202|T|4|402|RESULTADO FINANCEIRO  LIQUIDO|0,00|C|0,00|D|D||"

                        Resultado = linha
                        'strm.WriteLine(Resultado)
                        ConteudoDoArquivo.AppendLine(Resultado)
                        RegistroJ150 += 1
                        RegistroGeral += 1

                        linha = "|J150|4|4020201|D|5|40202|DESPESAS FINANCEIRAS |0,00|C|0,00|D|D||"

                        Resultado = linha
                        'strm.WriteLine(Resultado)
                        ConteudoDoArquivo.AppendLine(Resultado)
                        RegistroJ150 += 1
                        RegistroGeral += 1

                        linha = "|J150|5|5|T|1||RESULTADO DO EXERCICIO|0,00|C|0,00|D|D||"

                        Resultado = linha
                        'strm.WriteLine(Resultado)
                        ConteudoDoArquivo.AppendLine(Resultado)
                        RegistroJ150 += 1
                        RegistroGeral += 1
                    End If

                End If

            Next i 'Fim do Ciclo Trimestral/Anual

            ' Registro J800  - Outras Informaçoes
            'txtArquivoAuxiliar.Text = "c:\sped\ecd\2013\03189063.rtf"

            If txtArquivoAuxiliar.Text <> "" Then
                If File.Exists(txtArquivoAuxiliar.Text) = True Then
                    linha = "|J800"
                    linha &= "|099"
                    linha &= "|"
                    linha &= "|"
                    Dim strm1 As StreamReader = New StreamReader(txtArquivoAuxiliar.Text, True)
                    linha &= "|" & strm1.ReadToEnd
                    strm1.Close()
                    linha &= "|J800FIM|"

                    'strm.WriteLine(linha)
                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroJ800 += 1
                    RegistroGeral += 1
                End If
            End If


            ' Registro J900  - Termo De Encerramento 
            Sql = "  SELECT    Clientes.*, ClientesXEmpresas.*" & vbCrLf &
                  "         FROM    Clientes INNER JOIN" & vbCrLf &
                  "                 ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf &
                  "                 Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                  "     WHERE   Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            For Each dr In ds.Tables(0).Rows
                With dr
                    linha = "|J900"
                    linha &= "|TERMO DE ENCERRAMENTO"
                    linha &= "|" & txtNumeroDoLivro.Text
                    linha &= "|" & Mid(DdlNaturezaDoLivro.Text, 7, 50)
                    linha &= "|" & RTrim(.Item("Nome"))
                    linha &= "|QTD-J900" '& RegistroGeral
                    linha &= "|" & Format(CDate(txtDataInicial.Text), "ddMMyyyy")                                                     'Data Inicial
                    linha &= "|" & Format(CDate(txtDataFinal.Text), "ddMMyyyy")                                                       'Data Final
                    linha &= "|"
                End With
            Next

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            RegistroJ900 += 1
            RegistroGeral += 1

            ' Registro J930  - Termo De Encerramento 
            Sql = "  SELECT    Clientes.*, ClientesXEmpresas.*, " & vbCrLf &
                  "            Contador.Telefone AS TelefoneDoContador, Contador.Estado AS EstadoDoCRC, Contador.Email AS EmailDoContador" & vbCrLf &
                  "     FROM      Clientes INNER JOIN" & vbCrLf &
                  "               ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf &
                  "               Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                  "               LEFT OUTER JOIN Clientes AS Contador ON ClientesXEmpresas.CPFContador = Contador.Cliente_Id" & vbCrLf &
                  "     WHERE     Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            For Each dr In ds.Tables(0).Rows
                With dr
                    '1 - Registro Fixo
                    linha = "|J930"
                    '2 - Nome Contador
                    linha &= "|" & RTrim(.Item("NomeDoContador"))
                    '3 - CPF do Contador
                    linha &= "|" & .Item("CPFContador")
                    '4 - Qualificação assinante
                    linha &= "|" & .Item("QualificacaoContador")
                    '5 - Código de Qualificação
                    linha &= "|" & Left(.Item("CodigoQualificacaoContador"), 3)



                    '6 - Inscrição do Contador
                    linha &= "|" & Left(.Item("CRCContador"), 11)
                    '7 - E-mail
                    linha &= "|" & .Item("EmailDoContador")



                    linha &= "|" & .Item("TelefoneDoContador")
                    linha &= "|" & .Item("EstadoDoCRC")
                    linha &= "|"
                    linha &= "|"
                    linha &= "|"
                    linha &= "N|"


                    'strm.WriteLine(linha)
                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroJ930 += 1
                    RegistroGeral += 1

                    linha = "|J930"
                    linha &= "|" & Funcoes.EliminarCaracteresEspeciais(RTrim(.Item("NomeDoTitular")))
                    linha &= "|" & .Item("CPFTitular")
                    linha &= "|" & Funcoes.EliminarCaracteresEspeciais(RTrim(.Item("QualificacaoTitular")))
                    linha &= "|" & Left(.Item("CodigoQualificacaoTitular"), 3)
                    linha &= "|"
                    linha &= "|"
                    linha &= "|"
                    linha &= "|"
                    linha &= "|"
                    linha &= "|"
                    linha &= "|"
                    linha &= "S|"


                    'strm.WriteLine(linha)
                    ConteudoDoArquivo.AppendLine(linha)
                    RegistroJ930 += 1
                    RegistroGeral += 1
                End With
            Next

            ' Registro J990  - Encerramento do Bloco J
            RegistroJ990 += 1

            linha = "|J990"
            linha &= "|" & RegistroJ001 + RegistroJ005 + RegistroJ100 + RegistroJ150 + RegistroJ800 + RegistroJ900 + RegistroJ930 + RegistroJ990
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            RegistroGeral += 1

            ' Registro 9001  - Abertura do Bloco 9
            linha = "|9001"
            linha &= "|0"
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            Registro9001 += 1
            RegistroGeral += 1

            ' Registro 9900  - Registros Do Arquivo

            For i As Integer = 1 To 33

                linha = "|9900"

                Select Case i
                    Case 1
                        linha &= "|0000"
                        linha &= "|" & Registro0000
                    Case 2
                        linha &= "|0001"
                        linha &= "|" & Registro0001
                    Case 3
                        linha &= "|0007"
                        linha &= "|" & Registro0007
                    Case 4
                        linha &= "|0020"
                        linha &= "|" & Registro0020
                    Case 5
                        linha &= "|0150"
                        linha &= "|" & Registro0150
                    Case 6
                        linha &= "|0180"
                        linha &= "|" & Registro0180
                    Case 7
                        linha &= "|0990"
                        linha &= "|" & Registro0990

                    Case 8
                        linha &= "|I001"
                        linha &= "|" & RegistroI001
                    Case 9
                        linha &= "|I010"
                        linha &= "|" & RegistroI010
                    Case 10
                        linha &= "|I030"
                        linha &= "|" & RegistroI030
                    Case 11
                        linha &= "|I050"
                        linha &= "|" & RegistroI050
                    Case 12
                        linha &= "|I051"
                        linha &= "|" & RegistroI051

                    Case 13
                        linha &= "|I052"
                        linha &= "|" & RegistroI052
                    Case 14
                        linha &= "|I100"
                        linha &= "|" & RegistroI100
                    Case 15
                        linha &= "|I150"
                        linha &= "|" & RegistroI150
                    Case 16
                        linha &= "|I155"
                        linha &= "|" & RegistroI155
                    Case 17
                        linha &= "|I200"
                        linha &= "|" & RegistroI200
                    Case 18
                        linha &= "|I250"
                        linha &= "|" & RegistroI250
                    Case 19
                        linha &= "|I350"
                        linha &= "|" & RegistroI350
                    Case 20
                        linha &= "|I355"
                        linha &= "|" & RegistroI355
                    Case 21
                        linha &= "|I990"
                        linha &= "|" & RegistroI990
                    Case 22
                        linha &= "|J001"
                        linha &= "|" & RegistroJ001

                    Case 23
                        linha &= "|J005"
                        linha &= "|" & RegistroJ005
                    Case 24
                        linha &= "|J100"
                        linha &= "|" & RegistroJ100
                    Case 25
                        linha &= "|J150"
                        linha &= "|" & RegistroJ150


                    Case 26
                        linha &= "|J800"
                        linha &= "|" & RegistroJ800
                    Case 27
                        linha &= "|J900"
                        linha &= "|" & RegistroJ900
                    Case 28
                        linha &= "|J930"
                        linha &= "|" & RegistroJ930
                    Case 29
                        linha &= "|J990"
                        linha &= "|" & RegistroJ990


                    Case 30
                        linha &= "|9001"
                        linha &= "|" & Registro9001
                    Case 31
                        linha &= "|9900"
                        linha &= "|" & Registro9900 + 3
                    Case 32
                        linha &= "|9990"
                        linha &= "|1"
                    Case 33
                        linha &= "|9999"
                        linha &= "|1"
                End Select

                linha &= "|"

                'strm.WriteLine(linha)
                ConteudoDoArquivo.AppendLine(linha)
                Registro9900 += 1
                RegistroGeral += 1
            Next

            ' Registro 9990  - Encerramento do Bloco 9
            Registro9990 += 1

            linha = "|9990"
            linha &= "|" & Registro9001 + Registro9900 + Registro9990 + 1
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)
            RegistroGeral += 1

            ' Registro 9999  - Encerramento do Arquivo Digital
            Registro9999 += 1
            RegistroGeral += 1

            linha = "|9999"
            linha &= "|" & RegistroGeral
            linha &= "|"

            'strm.WriteLine(linha)
            ConteudoDoArquivo.AppendLine(linha)

            'strm.Flush()
            'strm.Close()

            'Dim LeituraDoArquivo As New StringBuilder()
            'Dim LinhaTemp As String

            ConteudoDoArquivo.Replace("QTD-LIN", RegistroGeral)
            ConteudoDoArquivo.Replace("QTD-J900", RegistroGeral)
            ConteudoDoArquivo.Remove(ConteudoDoArquivo.Length - 1, 1)

            strmSecundario = New StreamWriter(arquivo, True)
            strmSecundario.Write(ConteudoDoArquivo)
            strmSecundario.Flush()
            strmSecundario.Close()

            txtArquivoDeSaida.Text = "Sped\Ecd\" & Today().Year() - 1 & "\" & Left(Empresa(0), 8) & ".txt"
            txtArquivoAuxiliar.Text = "Sped\Ecd\" & Today().Year() - 1 & "\" & Left(Empresa(0), 8) & ".rtf"

            cmdArquivoDeSaida.Visible = True
            cmdArquivoAuxDeSaida.Visible = True
            '  Fim do Bloco I
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Protected Sub DdlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If DdlEmpresa.SelectedIndex > 0 Then
    '        Empresa = DdlEmpresa.SelectedValue.Split("-")
    '        txtArquivoDeSaida.Text = "Sped\Ecd\" & Today().Year() - 1 & "\" & Left(Empresa(0), 8) & ".txt"
    '        txtArquivoAuxiliar.Text = "Sped\Ecd\" & Today().Year() - 1 & "\" & Left(Empresa(0), 8) & ".rtf"
    '    End If
    'End Sub

#Region "Sqls Comentados"

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    ''' 

    Private Sub Comentados()

    End Sub

#End Region

    Private Sub Download(ByVal arquivo As String, ByVal extensao As String)
        Try
            Response.ContentType = "text/plain"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & arquivo.Replace(".txt", extensao))
            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(arquivo, FileMode.Open, FileAccess.Read)
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
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmdArquivoDeSaida_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdArquivoDeSaida.Click
        Try
            Download(Server.MapPath(String.Format("~/{0}", txtArquivoDeSaida.Text)), ".txt")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdArquivoAuxDeSaida_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdArquivoAuxDeSaida.Click
        Try
            Download(Server.MapPath(String.Format("~/{0}", txtArquivoDeSaida.Text)), ".rtf")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkProcessar_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                Processar()
            Else
                MsgBox(Me.Page, "Informe a Empresa.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SpedContabil")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class