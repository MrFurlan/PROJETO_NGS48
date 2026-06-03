Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ControleDeLote
    Inherits BasePage

    Dim objLote As [Lib].Negocio.Lote

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ControleDeLote", "ACESSAR") Then
                ddl.Carregar(ddlSafraConsulta, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
                carregaTiposDeLotes(ddlConsultaTipoDeLote)
                carregaTiposDeLotes(ddlTipoDeLote)
                objLote = New [Lib].Negocio.Lote
                SalvaObjSessao()
                CarregarClasseNoFormulario()

                txtFornecedorCad.Text = ""
                txtCodigoFornecedorCad.Value = 0
                HID.Value = Guid.NewGuid().ToString()
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaProduto.SetarHID(HID.Value)
                tcControleDeLote.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx", eTitulo.Info)
                Exit Sub
            End If
        End If
    End Sub

    Public Sub RecuperaObjSessao()
        objLote = Session("ObjLoteXClassificacao")
    End Sub

    Public Sub SalvaObjSessao()
        Session("ObjLoteXClassificacao") = objLote
    End Sub

    Public Function ValidaClassificacao() As Boolean
        If String.IsNullOrWhiteSpace(txtClassificacao.Text) Then
            MsgBox(Me.Page, "Informe a Classificação", eTitulo.Info)
            Return False
        End If

        If Not IsNumeric(txtPesoSaco.Text) OrElse CDec(txtPesoSaco.Text) <= 0 Then
            MsgBox(Me.Page, "Informe o Peso do Saco", eTitulo.Info)
            Return False
        End If
        Return True
    End Function

    Private Function FormatarCliente(ByVal ObjetoCliente As [Lib].Negocio.Cliente) As String
        Dim strDescricao As String = Funcoes.AlinharEsquerda(ObjetoCliente.Nome, 28, ".")
        strDescricao &= " - " & Funcoes.AlinharEsquerda(ObjetoCliente.Cidade, 20, ".") & " " & ObjetoCliente.CodigoEstado
        strDescricao &= " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(ObjetoCliente.Codigo), 18, ".")
        strDescricao &= "-" & ObjetoCliente.CodigoEndereco.ToString() & "-" & ObjetoCliente.Reduzido
        Return strDescricao
    End Function

    Public Function ValidarCampos() As Boolean
        RecuperaObjSessao()
        If String.IsNullOrWhiteSpace(txtLote.Text) Then
            MsgBox(Me.Page, "Informe o Lote")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodProdutoCadastro.Text) Then
            MsgBox(Me.Page, "Informe o Produto")
            Return False
        ElseIf ddlSafra.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione a Safra")
            Return False
        ElseIf txtCodigoFornecedorCad.Value.Length = 0 Then
            MsgBox(Me.Page, "Informe o Fabricante/fornecedor do lote.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataLancamento.Text) Or Not IsDate(txtDataLancamento.Text) Then
            MsgBox(Me.Page, "Informe uma Data de Lançamento válida.")
            txtDataLancamento.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataValidade.Text) Or Not IsDate(txtDataValidade.Text) Then
            MsgBox(Me.Page, "Informe uma Data de Validade válida.")
            txtDataValidade.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtNomeProdutoCad.Text) Then
            MsgBox(Me.Page, "Informe o Fabricante do Lote")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtRenasem.Text) And ddlTipoDeLote.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe o Renasem")
            txtRenasem.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtTermConformidade.Text) And ddlTipoDeLote.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe o Termo de Conformidade")
            txtTermConformidade.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtBoletim.Text) And ddlTipoDeLote.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe o Boletim")
            txtBoletim.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtGerminacao.Text) And ddlTipoDeLote.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Germinação")
            txtGerminacao.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtPureza.Text) And ddlTipoDeLote.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a Pureza")
            txtPureza.Focus()
            Return False
        ElseIf objLote.Classificacoes.Count = 0 And ddlTipoDeLote.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe ao menos uma Classificação")
            Return False
        End If
        Return True
    End Function

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteCrtL" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteCrtL" & HID.Value), [Lib].Negocio.Cliente))
            txtNomeFornecedor.Text = itemCliente.Text
            txtCodigoFornecedor.Value = itemCliente.Value
            Session.Remove("objClienteCrtL" & HID.Value)
        ElseIf Session("objClienteCLC" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteCLC" & HID.Value), [Lib].Negocio.Cliente))
            txtFornecedorCad.Text = itemCliente.Text
            txtCodigoFornecedorCad.Value = itemCliente.Value
            Session.Remove("objClienteCLC" & HID.Value)
        ElseIf Session("objProdutoCL" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoCL" & HID.Value)
            txtNomeProdutoCad.Text = objProduto.Descricao
            txtCodProdutoCadastro.Text = objProduto.Codigo
            Session.Remove("objProdutoCL" & HID.Value)
        End If
    End Sub

    Protected Sub btnFornecedor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFornecedor.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCrtL" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnProdutoCad_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProdutoCad.Click
        ucConsultaProduto.Limpar()
        Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeProduto(Me.Page, "objProdutoCL" & HID.Value, txtNome.ClientID, True)
    End Sub

    Protected Sub gridClassificacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtClassificacao.Text = gridClassificacao.SelectedRow.Cells(1).Text
        txtPesoSaco.Text = CDec(gridClassificacao.SelectedRow.Cells(2).Text).ToString("N4")

    End Sub

    Protected Sub gridLote_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridLote.SelectedIndexChanged
        Dim sql As String = ""
        sql &= "	SELECT 'Nota' as Doc, " & vbCrLf & _
               "		   NF.EntradaSaida_Id, " & vbCrLf & _
               "		   NF.Movimento," & vbCrLf & _
               "		   NF.Empresa_Id, " & vbCrLf & _
               "		   NF.EndEmpresa_Id, " & vbCrLf & _
               "		   Empresa.Nome AS NomeEmpresa," & vbCrLf & _
               "		   Empresa.Cidade AS CidadeEmpresa," & vbCrLf & _
               "		   Empresa.Estado AS EstadoEmpresa," & vbCrLf & _
               "		   NF.Cliente_Id, " & vbCrLf & _
               "		   NF.EndCliente_Id, " & vbCrLf & _
               "		   Clientes.Nome, " & vbCrLf & _
               "		   Clientes.Cidade, " & vbCrLf & _
               "		   Clientes.Estado, " & vbCrLf & _
               "		   NF.Nota_Id," & vbCrLf & _
               "		   NF.Serie_Id, " & vbCrLf & _
               "		   NFxI.Lote, " & vbCrLf & _
               "		   NFxI.Classificacao, " & vbCrLf & _
               "		   NFxI.QuantidadeFiscal " & vbCrLf & _
               "	  FROM NotasFiscais NF " & vbCrLf & _
               "	 INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
               "		ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
               "	   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
               "	   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
               "	   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
               "	   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
               "	   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
               "	   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
               "	 INNER JOIN Clientes " & vbCrLf & _
               "		ON NF.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf & _
               "	   AND NF.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf & _
               "	 INNER JOIN Clientes AS Empresa " & vbCrLf & _
               "		ON NF.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf & _
               "	   AND NF.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
               "	 where nfxi.lote = '" & gridLote.SelectedRow.Cells(3).Text & "'" & vbCrLf & _
               IIf(gridLote.SelectedRow.Cells(4).Text <> "&nbsp;", " And nfxi.Classificacao = '" & gridLote.SelectedRow.Cells(4).Text & "'", "") & vbCrLf & _
               "	 union" & vbCrLf & _
               "	select 'Prod'," & vbCrLf & _
               "		   case" & vbCrLf & _
               "			 when P.saidas > 0" & vbCrLf & _
               "			   Then 'S'" & vbCrLf & _
               "			   Else 'E'" & vbCrLf & _
               "		   end EntradaSaida_Id," & vbCrLf & _
               "		   P.Movimento_Id," & vbCrLf & _
               "		   P.Empresa_id," & vbCrLf & _
               "		   P.EndEmpresa_id," & vbCrLf & _
               "		   Empresa.Nome," & vbCrLf & _
               "		   Empresa.Cidade," & vbCrLf & _
               "		   Empresa.Estado, " & vbCrLf & _
               "		   P.Deposito_id," & vbCrLf & _
               "		   P.EndDeposito_id," & vbCrLf & _
               "		   Deposito.Nome," & vbCrLf & _
               "		   Deposito.Cidade," & vbCrLf & _
               "		   Deposito.Estado, " & vbCrLf & _
               "		   0," & vbCrLf & _
               "		   'P'," & vbCrLf & _
               "		   P.Lote_id," & vbCrLf & _
               "		   P.Classificacao_id, " & vbCrLf & _
               "		   P.Entradas + P.Saidas " & vbCrLf & _
               "  	  from producao P " & vbCrLf & _
               " 	 Inner Join Clientes Empresa" & vbCrLf & _
               "		ON P.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf & _
               "	   AND P.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
               "	 Inner Join Clientes Deposito" & vbCrLf & _
               "		ON P.Deposito_id    = Deposito.Cliente_Id " & vbCrLf & _
               "	   AND P.EndDeposito_Id = Deposito.Endereco_Id" & vbCrLf & _
               "	 where p.lote_id = '" & gridLote.SelectedRow.Cells(3).Text & "'" & vbCrLf & _
               IIf(gridLote.SelectedRow.Cells(4).Text <> "&nbsp;", " And P.Classificacao_id = '" & gridLote.SelectedRow.Cells(4).Text & "'", "") & vbCrLf & _
               " order by Movimento"


        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Mov")
        gridMovimentacao.DataSource = ds
        gridMovimentacao.DataBind()

        objLote = New [Lib].Negocio.Lote(gridLote.SelectedRow.Cells(3).Text, gridLote.SelectedRow.Cells(1).Text)
        txtFornecedorCad.Text = objLote.CodigoFornecedor & ";" & objLote.CodigoFornecedorEndereco
        SalvaObjSessao()
        CarregarClasseNoFormulario()
        tcControleDeLote.ActiveTabIndex = 2
    End Sub

    Protected Sub btnFornecedorCad_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFornecedorCad.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCLC" & HID.Value, "txtNome")
    End Sub

    Public Sub CarregarClasseNoFormulario()
        RecuperaObjSessao()

        If objLote.Lote = "" Then
            objLote.IUD = "I"
            lnkNovo.Enabled = True
            lnkAtualizar.Enabled = False
            lnkExcluir.Enabled = False
            txtLote.Enabled = True
            btnProdutoCad.Enabled = True
        Else
            objLote.IUD = "U"
            lnkNovo.Enabled = False
            lnkAtualizar.Enabled = True
            lnkExcluir.Enabled = True
            txtLote.Enabled = False
            txtCodProdutoCadastro.Enabled = False
            txtNomeProdutoCad.Enabled = False
            btnProdutoCad.Enabled = False
            ddlTipoDeLote.Enabled = False
        End If

        txtLote.Text = objLote.Lote
        txtDataLancamento.Text = objLote.DataLancamento.ToString("dd/MM/yyyy")
        txtDataValidade.Text = objLote.DataValidade.ToString("dd/MM/yyyy")
        txtRenasem.Text = objLote.Renasem
        txtCodProdutoCadastro.Text = objLote.CodigoProduto
        txtNomeProdutoCad.Text = objLote.Produto.Nome
        txtFornecedorCad.Text = FormatarCliente(objLote.Fornecedor)
        txtCodigoFornecedorCad.Value = objLote.CodigoFornecedor & "-" & objLote.CodigoFornecedorEndereco
        txtTermConformidade.Text = objLote.TermoConformidade
        txtBoletim.Text = objLote.Boletim
        txtGerminacao.Text = objLote.Germinacao
        txtPureza.Text = objLote.Pureza
        ddlSafra.SelectedValue = objLote.CodigoSafra
        txtClassificacao.Text = String.Empty
        txtPesoSaco.Text = String.Empty
        ddlTipoDeLote.SelectedIndex = objLote.Tipo - 1
        ddlTipoDeLote_SelectedIndexChanged(New Object, New EventArgs)
        gridClassificacao.DataSource = objLote.Classificacoes.ToArray
        gridClassificacao.DataBind()
    End Sub


    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("ControleDeLote", "GRAVAR") Then
            If ValidarCampos() Then
                RecuperaObjSessao()

                objLote.Lote = txtLote.Text
                objLote.Tipo = Left(ddlTipoDeLote.SelectedValue, 1)
                objLote.CodigoProduto = txtCodProdutoCadastro.Text
                objLote.CodigoSafra = ddlSafra.SelectedValue
                objLote.DataLancamento = CDate(txtDataLancamento.Text)
                objLote.DataValidade = CDate(txtDataValidade.Text)
                Dim StrFornecedor() As String = txtCodigoFornecedorCad.Value.Split("-")
                objLote.CodigoFornecedor = StrFornecedor(0)
                objLote.CodigoFornecedorEndereco = StrFornecedor(1)

                If ddlTipoDeLote.SelectedIndex > 0 Then
                    objLote.Renasem = 0
                    objLote.TermoConformidade = 0
                    objLote.Boletim = 0
                    objLote.Germinacao = 0
                    objLote.Pureza = 0
                    objLote.Classificacoes.Clear()

                    Dim Classificacao As New [Lib].Negocio.LoteXClassificacao(objLote)
                    Classificacao.IUD = "I"
                    Classificacao.Classificacao = 1
                    Classificacao.PesoSaco = CDec(0)
                    objLote.Classificacoes.Add(Classificacao)
                Else
                    objLote.Renasem = txtRenasem.Text
                    objLote.TermoConformidade = txtTermConformidade.Text
                    objLote.Boletim = txtBoletim.Text
                    objLote.Germinacao = txtGerminacao.Text
                    objLote.Pureza = txtPureza.Text
                End If

                objLote.IUD = "I"

                If objLote.Salvar Then
                    limpar()
                    MsgBox(Me.Page, "Lote Salvo com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, "ERRO")
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If

    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        RecuperaObjSessao()
        If Not ValidarCampos() Then Exit Sub
        objLote.DataLancamento = CDate(txtDataLancamento.Text)
        objLote.DataValidade = CDate(txtDataValidade.Text)

        Dim StrFornecedor() As String = txtCodigoFornecedorCad.Value.Split("-")
        objLote.CodigoFornecedor = StrFornecedor(0)
        objLote.CodigoFornecedorEndereco = StrFornecedor(1)
        objLote.CodigoSafra = ddlSafra.SelectedValue

        objLote.Renasem = txtRenasem.Text
        objLote.TermoConformidade = txtTermConformidade.Text
        objLote.Boletim = txtBoletim.Text
        objLote.Germinacao = txtGerminacao.Text
        objLote.Pureza = txtPureza.Text
        If objLote.Salvar Then
            limpar()
            MsgBox(Me.Page, "Lote Alterado com Sucesso.", eTitulo.Sucess)
            objLote = New [Lib].Negocio.Lote()
            SalvaObjSessao()
            CarregarClasseNoFormulario()
        Else
            MsgBox(Me.Page, "Erro ao Altera o Lote tente novamente.")
        End If
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("ControleDeLote", "EXCLUIR") Then
            RecuperaObjSessao()
            If objLote.JaExisteMovimentacao Then
                MsgBox(Me.Page, objLote.MensagemDeControle)
                Exit Sub
            End If

            objLote.IUD = "D"
            If objLote.Salvar Then
                limpar()
                MsgBox(Me.Page, "Lote Excluido com Sucesso.", eTitulo.Sucess)
                objLote = New [Lib].Negocio.Lote()
                SalvaObjSessao()
                CarregarClasseNoFormulario()
            Else
                MsgBox(Me.Page, "Erro ao Excluir o Lote tente novamente.")
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir Lote. ")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        limpar()
    End Sub

    Private Sub limpar()
        objLote = New [Lib].Negocio.Lote()
        SalvaObjSessao()
        CarregarClasseNoFormulario()
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
        ddlTipoDeLote.Enabled = True
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Dim sql As String
        sql = "	select Lote.Produto_Id, P.Nome, Lote.Lote_Id as Lote, " & IIf(RdLote.Checked, "'' as Classificacao,", "sb.Classificacao,") & vbCrLf & _
              "        isnull(fabricante.cliente_id,'') as cliente_id," & vbCrLf & _
              "        isnull(fabricante.endereco_id,0) as endereco_id," & vbCrLf & _
              "        isnull(fabricante.nome,'') as NomeFabricante," & vbCrLf & _
              "        isnull(fabricante.cidade,'') as cidade," & vbCrLf & _
              "        isnull(fabricante.estado,'') as estado," & vbCrLf & _
              "        sum(isnull(sb.Qtde,0)) Saldo," & vbCrLf & _
              "        lote.safra as CodigoSafra" & vbCrLf & _
              "	  from Lote " & vbCrLf & _
              "   Left join " & vbCrLf & _
              "	   (" & vbCrLf & _
              "		SELECT NotasFiscaisXItens.Produto_Id," & vbCrLf & _
              "			   NotasFiscaisXItens.Lote," & vbCrLf & _
              "			   NotasFiscaisXItens.Classificacao," & vbCrLf & _
              "			   sum(case" & vbCrLf & _
              "					 when Suboperacoes.EntradaSaida = 'S'" & vbCrLf & _
              "					  then NotasFiscaisXItens.QuantidadeFiscal * - 1" & vbCrLf & _
              "					  else  NotasFiscaisXItens.QuantidadeFiscal " & vbCrLf & _
              "				   end) as Qtde" & vbCrLf & _
              "		  FROM NotasFiscais" & vbCrLf & _
              "		 INNER JOIN NotasFiscaisXItens " & vbCrLf & _
              "			ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
              "		   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
              "		   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
              "		   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
              "		   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
              "		   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
              "		   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
              "		 INNER JOIN SubOperacoes " & vbCrLf & _
              "			ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id " & vbCrLf & _
              "		   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf

        If txtPosicaoDia.Text.Length > 0 Then
            sql &= " Where NotasFiscais.Movimento <= '" & txtPosicaoDia.Text.ToSqlDate() & "'" & vbCrLf
        End If




        sql &= "		 Group by  NotasFiscaisXItens.Produto_Id," & vbCrLf & _
               "				   NotasFiscaisXItens.Lote," & vbCrLf & _
               "				   NotasFiscaisXItens.Classificacao" & vbCrLf & _
               "		 union" & vbCrLf & _
               "		 SELECT P.Produto_Id, P.Lote_Id, P.Classificacao_Id," & vbCrLf & _
               "				sum(Entradas) - sum(Saidas) as Saldo" & vbCrLf & _
               "		   FROM Producao P" & vbCrLf & _
               "          Where FisicoFiscal_Id = 2" & vbCrLf

        If txtPosicaoDia.Text.Length > 0 Then
            sql &= "   And P.Movimento_id <= '" & txtPosicaoDia.Text.ToSqlDate() & "'" & vbCrLf
        End If

        sql &= "		  GROUP BY P.Produto_Id, P.Lote_Id, P.Classificacao_Id" & vbCrLf & _
              "	   ) Sb" & vbCrLf & _
              "		ON Sb.Produto_id = Lote.Produto_id " & vbCrLf & _
              "	   AND Sb.Lote       = Lote.Lote_id " & vbCrLf & _
              "	  left Join Clientes Fabricante" & vbCrLf & _
              "		on Fabricante.Cliente_id  = Lote.Fornecedor" & vbCrLf & _
              "	   and Fabricante.Endereco_id = Lote.EndFornecedor" & vbCrLf & _
              "	 Inner Join Produtos P" & vbCrLf & _
              "		on P.Produto_Id    = Lote.Produto_Id" & vbCrLf & _
              "	   and P.ControlarLote = 'S'" & vbCrLf & _
              "  Where 1 = 1"

        If txtCodigoFornecedor.Value.Length > 0 Then
            Dim fornecedor() As String = txtCodigoFornecedor.Value.Split("-")
            sql &= "	   And Fabricante.Cliente_id  ='" & fornecedor(0) & "'" & vbCrLf & _
                   "       And fabricante.endereco_id = " & fornecedor(1) & vbCrLf
        End If

        If chkLancados.Checked And IsDate(txtDataLancamentoInicial.Text) And IsDate(txtDataLancamentoFinal.Text) Then
            sql &= "	   And Lote.DataLancamento between '" & txtDataLancamentoInicial.Text.ToSqlDate() & "' and '" & txtDataLancamentoFinal.Text.ToSqlDate() & "'" & vbCrLf
        End If

        If chkValidade.Checked And IsDate(txtDataVencimentoInicial.Text) And IsDate(txtDataVencimentoFinal.Text) Then
            sql &= "	   And Lote.DataValidade between '" & txtDataVencimentoInicial.Text.ToSqlDate() & "' and '" & txtDataVencimentoFinal.Text.ToSqlDate() & "'" & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Dim Prd As ArrayList
            Prd = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "P.Produto_Id")
            sql &= "	   And " & Prd(0) & vbCrLf
        End If

        If ddlSafraConsulta.Text.Length > 0 Then
            sql &= "	   And Lote.Safra ='" & ddlSafraConsulta.SelectedValue & "'" & vbCrLf
        End If

        If txtLoteConsulta.Text.Trim.Length > 0 Then
            sql &= "	   And Lote.Lote_Id like '%" & txtLoteConsulta.Text & "%'" & vbCrLf
        End If

        If RdAbertos.Checked Then
            If txtPosicaoDia.Text.Length > 0 Then
                sql &= "	   And Lote.DataValidade >= '" & txtPosicaoDia.Text.ToSqlDate() & "'" & vbCrLf
            Else
                sql &= "	   And Lote.DataValidade >= '" & Date.Now.ToSqlDate() & "'" & vbCrLf
            End If
        ElseIf RdVencidos.Checked Then
            If txtPosicaoDia.Text.Length > 0 Then
                sql &= "	   And Lote.DataValidade <= '" & txtPosicaoDia.Text.ToSqlDate() & "'" & vbCrLf
            Else
                sql &= "	   And Lote.DataValidade <= '" & Date.Now.ToSqlDate() & "'" & vbCrLf
            End If
        End If

        sql &= " AND Lote.Tipo = " & Left(ddlConsultaTipoDeLote.SelectedValue, 1) & vbCrLf

        sql &= "	group by Lote.Produto_Id, P.Nome, Lote.Lote_Id, " & IIf(RdLote.Checked, "", "sb.Classificacao,") & " fabricante.cliente_id, fabricante.endereco_id, fabricante.nome, fabricante.cidade , fabricante.estado, lote.safra" & vbCrLf
        If RdComSaldo.Checked Then
            sql &= " having sum(sb.Qtde) > 0"
        ElseIf RdSemSaldo.Checked Then
            sql &= " having sum(sb.Qtde) = 0"
        End If

        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "Lote")
        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            gridLote.DataSource = ds
            gridLote.DataBind()
        Else
            gridLote.DataSource = Nothing
            gridLote.DataBind()
            MsgBox(Me.Page, "Nenhum resultado referente aos dados de pesquisa.")
        End If
    End Sub

    Protected Sub lnkLimparConsulta_Click(sender As Object, e As EventArgs) Handles lnkLimparConsulta.Click
        txtNomeFornecedor.Text = ""
        txtCodigoFornecedor.Value = ""

        ddlSafraConsulta.SelectedIndex = 0
        txtDataLancamentoInicial.Text = ""
        txtDataLancamentoFinal.Text = ""

        txtDataVencimentoInicial.Text = ""
        txtDataVencimentoFinal.Text = ""

        txtPosicaoDia.Text = ""
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucSelecaoProduto.Limpar()
        ddlConsultaTipoDeLote.SelectedIndex = 0
        gridLote.DataSource = Nothing
        gridLote.DataBind()
        gridMovimentacao.DataSource = Nothing
        gridMovimentacao.DataBind()
    End Sub

    Protected Sub chkLancados_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkLancados.CheckedChanged
        PnlLancados.Visible = chkLancados.Checked
    End Sub

    Protected Sub chkValidade_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkValidade.CheckedChanged
        pnlValidade.Visible = chkValidade.Checked
    End Sub

    Private Sub carregaTiposDeLotes(ddlTipoLote As DropDownList)
        ddlTipoLote.Items.Clear()
        ddlTipoLote.Items.Add("1 - Sementes")
        ddlTipoLote.Items.Add("2 - Defensivos")
    End Sub

    Protected Sub ddlTipoDeLote_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTipoDeLote.SelectedIndexChanged
        If ddlTipoDeLote.SelectedIndex = 0 Then
            pnlClassificacao.Visible = True
            pnlCamposLoteSemente.Visible = True
        ElseIf ddlTipoDeLote.SelectedIndex = 1 Then
            pnlClassificacao.Visible = False
            pnlCamposLoteSemente.Visible = False
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ControleDeLote")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkAdicionarClassificacao_Click(sender As Object, e As EventArgs) Handles lnkAdicionarClassificacao.Click
        If Not ValidaClassificacao() Then Exit Sub

        RecuperaObjSessao()

        If objLote.Classificacoes.JaExisteClassificacao(txtClassificacao.Text.Trim) Then
            If objLote.Classificacoes.ClassificacaoExistente.IUD = "I" Then
                objLote.Classificacoes.ClassificacaoExistente.PesoSaco = CDec(txtPesoSaco.Text)
            Else
                objLote.Classificacoes.ClassificacaoExistente.IUD = "U"
                objLote.Classificacoes.ClassificacaoExistente.PesoSaco = CDec(txtPesoSaco.Text)
                If objLote.Classificacoes.ClassificacaoExistente.Salvar Then
                    objLote.Classificacoes = Nothing
                End If
            End If
        Else
            Dim Classificacao As New [Lib].Negocio.LoteXClassificacao(objLote)
            Classificacao.IUD = "I"
            Classificacao.Classificacao = txtClassificacao.Text.Trim
            Classificacao.PesoSaco = CDec(txtPesoSaco.Text)
            objLote.Classificacoes.Add(Classificacao)
        End If
        gridClassificacao.DataSource = objLote.Classificacoes.ToArray
        gridClassificacao.DataBind()
        SalvaObjSessao()
    End Sub

    Protected Sub lnkRemoverClassificacao_Click(sender As Object, e As EventArgs) Handles lnkRemoverClassificacao.Click
        If Not ValidaClassificacao() Then Exit Sub
        Dim i As Integer
        i = gridClassificacao.SelectedIndex
        RecuperaObjSessao()

        If objLote.Classificacoes.Count = 1 Then
            MsgBox(Me.Page, "O Lote dever ter no minimo uma Classificacao")
            Exit Sub
        End If

        If Not objLote.Classificacoes.Remover(txtClassificacao.Text) Then
            MsgBox(Me.Page, objLote.Classificacoes.MensagemControle)
            Exit Sub
        End If

        SalvaObjSessao()
        gridClassificacao.DataSource = objLote.Classificacoes.ToArray
        gridClassificacao.DataBind()

        txtClassificacao.Text = gridClassificacao.Rows(IIf(gridClassificacao.Rows.Count = i, i - 1, 0)).Cells(1).Text
        gridClassificacao.SelectedIndex = IIf(i > 0, i - 1, 0)
    End Sub
End Class