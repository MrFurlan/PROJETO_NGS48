Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioControleDeCarregamento
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioControleDeCarregamento", "ACESSAR") Then 'criar processo
                    BuscaEmpresa()
                    BuscarOperacoes()
                    txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaClientes.SetarHID(HID.Value)
                    hIntacta.Value = False
                    ucSelecaoProduto.WhereProduto = "Agrupar = 'N'"
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteCliOri" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteCliOri" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClienteOrigem.Text = itemCliente.Text
            txtCodigoCliOrigem.Value = itemCliente.Value
            Session.Remove("objClienteCliOri" & HID.Value)
        ElseIf Session("objClienteCliDest" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteCliDest" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClienteDestino.Text = itemCliente.Text
            txtCodigoCliDestino.Value = itemCliente.Value
            Session.Remove("objClienteCliDest" & HID.Value)
        End If
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
        Funcoes.VerificaEmpresa(cmbEmpresa)
    End Sub

    Private Sub BuscarOperacoes()
        ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
    End Sub

    Private Sub BuscarSubOperacoes()
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id=" & cmbOperacao.SelectedValue & " ", True)
    End Sub

    Private Function ValidarCampos() As Boolean
        If cmbEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf txtDataInicial.Text.Length = 0 Or IsDate(txtDataInicial.Text) = False Then
            MsgBox(Me.Page, "Informe uma data inicial válida.")
            txtDataInicial.Focus()
            Return False
        ElseIf txtDataFinal.Text.Length = 0 Or IsDate(txtDataFinal.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida.")
            txtDataFinal.Focus()
            Return False
        ElseIf CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final.")
            txtDataInicial.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbEmpresa.Enabled = False
        End If
    End Sub

    Public Sub BuscarRegistros()
        Dim sqlProd As String = ""
        Dim DescricaoProduto As String = ""
        Dim dsProduto As New DataSet
        Dim ds As New DataSet
        Dim Sql As String
        Dim Cliente As String = ""
        Dim strEmpresa() As String = cmbEmpresa.SelectedValue.Split("-")
        Dim strClienteOrigem() As String = txtCodigoCliOrigem.Value.Split("-")
        Dim strClienteDestino() As String = txtCodigoCliDestino.Value.Split("-")

        Dim x, y As String

        If rdOrigemDestino.Checked Then
            x = "A"
            y = "B"
        Else
            x = "B"
            y = "A"
        End If

        Sql = "SELECT NFORIGEM.DataDaNota        as Data" & x & vbCrLf &
              "      ,CLIORI.Cliente_Id +'-'+ convert(nvarchar,CLIORI.endereco_id) as Cliente" & x & vbCrLf &
              "      ,CLIORI.Nome                as Nome" & x & vbCrLf &
              "      ,CLIORI.Complemento         as Complemento" & x & vbCrLf &
              "      ,NFORIGEM.pedido            as Pedido" & x & vbCrLf &
              "      ,PRDORI.Nome                as Produto" & x & vbCrLf &
              "      ,NFORIGEM.Nota_Id           as Nota" & x & vbCrLf &
              "      ,NFORIGEM.EntradaSaida_Id   as Es" & x & vbCrLf &
              "      ,OPEORI.Classe              as Classe" & x & vbCrLf &
              "      ,NFIORIGEM.QuantidadeFisica as PesoBruto" & x & vbCrLf &
              "      ,NFIORIGEM.QuantidadeFiscal as PesoLiquido" & x & vbCrLf &
              "      ,NFIORIGEM.Unitario         as Unitario" & x & vbCrLf &
              "      ,NFIORIGEM.Valor            as Valor" & x & vbCrLf &
              "      ,NFDESTINO.DataDaNota        as Data" & y & vbCrLf &
              "      ,CLIDES.Cliente_Id +'-'+ convert(nvarchar,CLIDES.endereco_id) as Cliente" & y & vbCrLf &
              "      ,CLIDES.Nome                 as Nome" & y & vbCrLf &
              "      ,CLIDES.Complemento          as Complemento" & y & vbCrLf &
              "      ,NFDESTINO.Pedido            as Pedido" & y & vbCrLf &
              "      ,PRDDST.nome                 as Produto" & y & vbCrLf &
              "      ,NFDESTINO.Nota_Id           as Nota" & y & vbCrLf &
              "      ,NFDESTINO.EntradaSaida_Id   as Es" & y & vbCrLf &
              "      ,OPEDES.classe               as Classe" & y & vbCrLf &
              "      ,NFIDESTINO.QuantidadeFisica as PesoBruto" & y & vbCrLf &
              "      ,NFIDESTINO.QuantidadeFiscal as PesoLiquido" & y & vbCrLf &
              "      ,NFIDESTINO.Unitario         as Unitario" & y & vbCrLf &
              "      ,NFIDESTINO.Valor            as Valor" & y & vbCrLf &
              "  FROM NotasXNotas NxN" & vbCrLf &
              " INNER JOIN NotasFiscais AS NFORIGEM" & vbCrLf &
              "    ON NFORIGEM.Empresa_Id       = NxN.OrigemEmpresa_Id" & vbCrLf &
              "   AND NFORIGEM.EndEmpresa_Id    = NxN.OrigemEndEmpresa_Id" & vbCrLf &
              "   AND NFORIGEM.Cliente_Id       = NxN.OrigemCliente_Id" & vbCrLf &
              "   AND NFORIGEM.EndCliente_Id    = NxN.OrigemEndCliente_Id" & vbCrLf &
              "   AND NFORIGEM.EntradaSaida_Id  = NxN.OrigemEntradaSaida_Id" & vbCrLf &
              "   AND NFORIGEM.Serie_Id         = NxN.OrigemSerie_Id" & vbCrLf &
              "   AND NFORIGEM.Nota_Id          = NxN.OrigemNota_Id" & vbCrLf &
              "   AND NFORIGEM.TipoDeDocumento  = 1" & vbCrLf &
              "   AND NFORIGEM.Situacao         = 1" & vbCrLf &
              "  Left JOIN NotasFiscais AS NFDESTINO" & vbCrLf &
              "    ON NFDESTINO.Empresa_Id       = NxN.Empresa_Id" & vbCrLf &
              "   AND NFDESTINO.EndEmpresa_Id    = NxN.EndEmpresa_Id" & vbCrLf &
              "   AND NFDESTINO.Cliente_Id       = NxN.Cliente_Id" & vbCrLf &
              "   AND NFDESTINO.EndCliente_Id    = NxN.EndCliente_Id" & vbCrLf &
              "   AND NFDESTINO.EntradaSaida_Id  = NxN.EntradaSaida_Id" & vbCrLf &
              "   AND NFDESTINO.Serie_Id         = NxN.Serie_Id" & vbCrLf &
              "   AND NFDESTINO.Nota_Id          = NxN.Nota_Id" & vbCrLf &
              "   AND NFDESTINO.TipoDeDocumento  = 1" & vbCrLf &
              "   AND NFDESTINO.Situacao         = 1" & vbCrLf &
              " INNER JOIN NotasFiscaisXItens AS NFIDESTINO" & vbCrLf &
              "    ON NFDESTINO.Empresa_Id      = NFIDESTINO.Empresa_Id" & vbCrLf &
              "   AND NFDESTINO.EndEmpresa_Id   = NFIDESTINO.EndEmpresa_Id" & vbCrLf &
              "   AND NFDESTINO.Cliente_Id      = NFIDESTINO.Cliente_Id" & vbCrLf &
              "   AND NFDESTINO.EndCliente_Id   = NFIDESTINO.EndCliente_Id" & vbCrLf &
              "   AND NFDESTINO.EntradaSaida_Id = NFIDESTINO.EntradaSaida_Id" & vbCrLf &
              "   AND NFDESTINO.Serie_Id        = NFIDESTINO.Serie_Id" & vbCrLf &
              "   AND NFDESTINO.Nota_Id         = NFIDESTINO.Nota_Id" & vbCrLf

        If hIntacta.Value Then
            If Not rdITodos.Checked Then
                Sql &= " INNER JOIN NotasFiscaisXRomaneios AS NxR" & vbCrLf &
                       "    ON NxR.Empresa_Id       = NFORIGEM.Empresa_Id" & vbCrLf &
                       "   AND NxR.EndEmpresa_Id    = NFORIGEM.EndEmpresa_Id" & vbCrLf &
                       "   AND NxR.Cliente_Id       = NFORIGEM.Cliente_Id" & vbCrLf &
                       "   AND NxR.EndCliente_Id    = NFORIGEM.EndCliente_Id" & vbCrLf &
                       "   AND NxR.EntradaSaida_Id  = NFORIGEM.EntradaSaida_Id" & vbCrLf &
                       "   AND NxR.Serie_Id         = NFORIGEM.Serie_Id" & vbCrLf &
                       "   AND NxR.Nota_Id          = NFORIGEM.Nota_Id" & vbCrLf &
                       " INNER JOIN Romaneios AS RO" & vbCrLf &
                       "    ON RO.Empresa_Id       = NxR.Empresa_Id" & vbCrLf &
                       "   AND RO.EndEmpresa_Id    = NxR.EndEmpresa_Id" & vbCrLf &
                       "   AND RO.Romaneio_Id      = NxR.Romaneio_Id" & vbCrLf &
                       " INNER JOIN RomaneiosXDescontos AS ROxD" & vbCrLf &
                       "    ON ROxD.Empresa_Id       = RO.Empresa_Id" & vbCrLf &
                       "   AND ROxD.EndEmpresa_Id    = RO.EndEmpresa_Id" & vbCrLf &
                       "   AND ROxD.Romaneio_Id      = RO.Romaneio_Id" & vbCrLf &
                       "   AND ROxD.Analise_Id       = 12" & vbCrLf
            End If

            If rdISim.Checked Then
                Sql &= "AND ROxD.Percentual       = 1" & vbCrLf
            ElseIf rdPositivo.Checked Then
                Sql &= "AND ROxD.Percentual       = 2" & vbCrLf
            ElseIf rdPSim.Checked Then
                Sql &= "AND ROxD.Percentual IN(1,2)" & vbCrLf
            ElseIf rdINao.Checked Then
                Sql &= "AND ROxD.Percentual       = 0" & vbCrLf
            End If
        End If

        Sql &= " INNER JOIN Clientes AS CLIDES" & vbCrLf &
              "    ON NFDESTINO.Cliente_Id    = CLIDES.Cliente_Id" & vbCrLf &
              "   AND NFDESTINO.EndCliente_Id = CLIDES.Endereco_Id" & vbCrLf &
              " INNER JOIN Produtos AS PRDDES" & vbCrLf &
              "    ON NFIDESTINO.Produto_Id    = PRDDES.Produto_Id " & vbCrLf &
              " INNER JOIN SubOperacoes AS OPEDES " & vbCrLf &
              "    ON NFDESTINO.Operacao    = OPEDES.Operacao_Id " & vbCrLf &
              "   AND NFDESTINO.SubOperacao = OPEDES.SubOperacoes_Id " & vbCrLf &
              " INNER JOIN NotasFiscaisXItens AS NFIORIGEM " & vbCrLf &
              "    ON NFORIGEM.Empresa_Id      = NFIORIGEM.Empresa_Id " & vbCrLf &
              "   AND NFORIGEM.EndEmpresa_Id   = NFIORIGEM.EndEmpresa_Id " & vbCrLf &
              "   AND NFORIGEM.Cliente_Id      = NFIORIGEM.Cliente_Id " & vbCrLf &
              "   AND NFORIGEM.EndCliente_Id   = NFIORIGEM.EndCliente_Id " & vbCrLf &
              "   AND NFORIGEM.EntradaSaida_Id = NFIORIGEM.EntradaSaida_Id " & vbCrLf &
              "   AND NFORIGEM.Serie_Id        = NFIORIGEM.Serie_Id " & vbCrLf &
              "   AND NFORIGEM.Nota_Id         = NFIORIGEM.Nota_Id " & vbCrLf &
              " INNER JOIN Clientes AS CLIORI " & vbCrLf &
              "    ON NFORIGEM.Cliente_Id    = CLIORI.Cliente_Id " & vbCrLf &
              "   AND NFORIGEM.EndCliente_Id = CLIORI.Endereco_Id " & vbCrLf &
              " INNER JOIN Produtos AS PRDORI " & vbCrLf &
              "    ON NFIORIGEM.Produto_Id    = PRDORI.Produto_Id " & vbCrLf &
              " INNER JOIN Produtos AS PRDDST" & vbCrLf &
              "    ON NFIDESTINO.Produto_Id    = PRDDST.Produto_Id" & vbCrLf &
              " INNER JOIN SubOperacoes AS OPEORI" & vbCrLf &
              "    ON NFORIGEM.Operacao    = OPEORI.Operacao_Id " & vbCrLf &
              "   AND NFORIGEM.SubOperacao = OPEORI.SubOperacoes_Id " & vbCrLf &
              " WHERE (isnull(NFORIGEM.troca,0) = 1 or (NFDESTINO.Empresa_id is not null and OPEORI.Devolucao = 'N'))" & vbCrLf &
              "   AND NFORIGEM.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

        If strEmpresa(0).Length > 0 Then
            Sql &= " AND NFDESTINO.Empresa_Id    = '" & strEmpresa(0) & "'" & vbCrLf
            Sql &= " AND NFDESTINO.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
        End If

        If strClienteOrigem(0).Length > 0 Then
            If chkConsolidarClienteOrigem.Checked Then
                Sql &= " AND left(NFORIGEM.Cliente_Id,8) = '" & strClienteOrigem(0).Substring(0, 8) & "'" & vbCrLf
            Else
                Sql &= " AND NFORIGEM.Cliente_Id    = '" & strClienteOrigem(0) & "'" & vbCrLf
                Sql &= " AND NFORIGEM.EndCliente_Id = " & strClienteOrigem(1) & " " & vbCrLf
            End If
        End If

        If strClienteDestino(0).Length > 0 Then
            If chkConsolidarClienteDestino.Checked Then
                Sql &= " AND left(NFDESTINO.Cliente_Id,8)    = '" & strClienteDestino(0).Substring(0, 8) & "'" & vbCrLf
            Else
                Sql &= " AND NFDESTINO.Cliente_Id    = '" & strClienteDestino(0) & "'" & vbCrLf
                Sql &= " AND NFDESTINO.EndCliente_Id = " & strClienteDestino(1) & " " & vbCrLf
            End If
        End If

        If strClienteOrigem(0).Length > 0 And strClienteDestino(0).Length > 0 OrElse strClienteOrigem(0).Length = 0 And strClienteDestino(0).Length = 0 OrElse strClienteOrigem(0).Length = 0 And strClienteDestino(0).Length > 0 Then

            If ucSelecaoProduto.TemSelecionado Then
                Sql &= "AND " & ucSelecaoProduto.GetSqlEParametrosRelatorio("PRDDES.Grupo", "NFIDESTINO.Produto_Id", "", True)(0)
            End If

            If cmbOperacao.Text.Trim <> "" Then
                Sql &= " AND OPEDES.Operacao_Id = " & cmbOperacao.SelectedValue & " " & vbCrLf
            End If

            If cmbSubOperacao.Text.Trim <> "" Then
                Sql &= " AND OPEDES.SubOperacoes_Id = " & cmbSubOperacao.SelectedValue & " " & vbCrLf
            End If

        ElseIf strClienteOrigem(0).Length > 0 And strClienteDestino(0).Length = 0 Then

            If ucSelecaoProduto.TemSelecionado Then
                Sql &= "AND  " & ucSelecaoProduto.GetSqlEParametrosRelatorio("PRDORI.Grupo", "NFIORIGEM.Produto_Id", "", True)(0)
            End If

            If cmbOperacao.Text.Trim <> "" Then
                Sql &= " AND OPEORI.Operacao_Id = " & cmbOperacao.SelectedValue & " " & vbCrLf
            End If

            If cmbSubOperacao.Text.Trim <> "" Then
                Sql &= " AND OPEORI.SubOperacoes_Id = " & cmbSubOperacao.SelectedValue & " " & vbCrLf
            End If

        End If

        If rdOrigemDestino.Checked Then
            Sql &= "  ORDER BY CLIORI.Cliente_Id, CLIORI.Endereco_id, NFORIGEM.DataDaNota, NFORIGEM.Nota_Id "
        Else
            Sql &= "  ORDER BY CLIDES.Cliente_Id, CLIDES.Endereco_id, NFDESTINO.DataDaNota, NFDESTINO.Nota_Id "
        End If


        ds = Banco.ConsultaDataSet(Sql, "RelatorioTrocaDeNota")
        AlimentaCrptRelatorios(ds, "Cr_RelatorioTrocaDeNota")
    End Sub

    Public Sub AlimentaCrptRelatorios(ByVal Ds As DataSet, ByVal NomeDoArquivo As String)
        Dim crptRelatorio As New ReportDocument()

        Try
            Dim parameters As New Dictionary(Of String, Object)
            parameters.Add("Nome", HttpContext.Current.Session("ssNomeEmpresa"))
            parameters.Add("Cidade", HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa"))
            parameters.Add("Origem", IIf(rdOrigemDestino.Checked, "NOTA DE ORIGEM", "NOTA DE DESTINO"))
            parameters.Add("Destino", IIf(rdOrigemDestino.Checked, "NOTA DE DESTINO", "NOTA DE ORIGEM"))

            Funcoes.BindReport(Me.Page, Ds, NomeDoArquivo, eExportType.PDF, parameters)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crptRelatorio.Close()
            crptRelatorio.Dispose()
        End Try
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            cmbSubOperacao.SelectedIndex = 0
            If cmbOperacao.SelectedIndex <> 0 Then BuscarSubOperacoes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliOrigem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaCliOrigem.Click
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCliOri" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliDestino_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaCliDestino.Click
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            Popup.ConsultaDeClientes(Me.Page, "objClienteCliDest" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        If Funcoes.VerificaPermissao("RelatorioControleDeCarregamento", "RELATORIO") Then
            If ValidarCampos() Then
                BuscarRegistros()
            End If
        Else
            MsgBox(Me.Page, "Usuario sem permissao para emitir relatório.")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            chkConsolidarClienteOrigem.Checked = False
            chkConsolidarClienteDestino.Checked = False
            txtClienteOrigem.Text = ""
            txtClienteDestino.Text = ""
            txtCodigoCliOrigem.Value = ""
            txtCodigoCliDestino.Value = ""
            hIntacta.Value = False
            ucSelecaoProduto.Limpar()
            cmbOperacao.SelectedIndex = 0
            cmbSubOperacao.SelectedIndex = 0
            txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
            txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
            cmbEmpresa.SelectedIndex = 0
            HID.Value = Guid.NewGuid().ToString()
            ucConsultaClientes.SetarHID(HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioControleDeCarregamento")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class