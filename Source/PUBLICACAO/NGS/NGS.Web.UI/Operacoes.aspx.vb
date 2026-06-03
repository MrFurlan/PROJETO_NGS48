Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Operacoes
    Inherits BasePage

    Dim Sql As String
    Dim ObjListOperacao As ListOperacao
    Dim ObjOperacao As Operacao
    Dim ObjSubOperacao As SubOperacao

#Region "Session"
    Private Sub SessaosSalvaListOperacao()
        Session("ObjListOperacao" & HID.Value) = ObjListOperacao
    End Sub

    Private Sub SessaoRecuperaListOperacao()
        If Session("ObjListOperacao" & HID.Value) Is Nothing Then
            ObjListOperacao = New [Lib].Negocio.ListOperacao(True)
        Else
            ObjListOperacao = CType(Session("ObjListOperacao" & HID.Value), [Lib].Negocio.ListOperacao)
        End If
    End Sub

    Private Sub SessaosSalvaOperacao()
        Session("ObjOperacao" & HID.Value) = ObjOperacao
    End Sub

    Private Sub SessaoRecuperaOperacao()
        ObjOperacao = CType(Session("ObjOperacao" & HID.Value), [Lib].Negocio.Operacao)
    End Sub

    Private Sub SessaosSalvaSubOperacao()
        Session("ObjSubOperacao" & HID.Value) = ObjSubOperacao
    End Sub

    Private Sub SessaoRecuperaSubOperacao()
        ObjSubOperacao = CType(Session("ObjSubOperacao" & HID.Value), [Lib].Negocio.SubOperacao)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Gestao.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("Operacoes", "ACESSAR") Then
                    HID.Value = Guid.NewGuid().ToString
                    ddl.Carregar(CmbClasse, CarregarDDL.Tabela.ClasseDeOperacao)
                    ddl.Carregar(ddlClasse, CarregarDDL.Tabela.ClasseDeOperacao)

                    SessaoRecuperaListOperacao()
                    gridOperacao.DataSource = ObjListOperacao.ToArray()
                    gridOperacao.DataBind()

                    ddl.Carregar(DdlGrupoDeContas, CarregarDDL.Tabela.PlanoDeContas, " (Cliente = 'S' and len(conta_id) = 7) or (len(conta_id) = 9 and left(conta_id, 1) in (1,2,3,4)) order by len(conta_id), conta_id", True)

                    ddl.Carregar(ddlContraPartidaCusto, CarregarDDL.Tabela.PlanoDeCustos, "Codigo_id <> 0", True)
                    ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "Situacao_Id <= 5")
                    ddl.Carregar(ddlOperacaoDestino, CarregarDDL.Tabela.Operacao, "")
                    ddl.Carregar(DdlTabelaDeApuracaoDeCustos, CarregarDDL.Tabela.PlanoDeCustos)

                    LimparOperacoes()
                    LimparSubOperacoes(False)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#Region "Botoes Operacao"
    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Operacoes", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim SqlArray As New ArrayList
                    Sql = "Insert into Operacoes (Operacao_Id,Descricao,Producao, Classe, UFDepositoDestino, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf & _
                          "values (" & TxtOperacaoCodigo.Text & ",'" & Funcoes.EliminarCaracteresEspeciais(TxtOperacaoDescricao.Text) & "'," & vbCrLf & _
                          "'" & IIf(ChkOperacaoProducao.Checked, "S", "N") & "','" & ddlClasse.SelectedValue & "','" & ChkUnitarioPedido.Checked.ToString & "','" & UsuarioServidor.NomeUsuario & "', '" & DateTime.Now.ToString("yyyy-MM-dd") & "')"
                    SqlArray.Add(Sql)
                    If Banco.GravaBanco(SqlArray) = False Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        ObjListOperacao = Nothing
                        SessaosSalvaListOperacao()

                        gridSubOperacao.DataSource = Nothing
                        gridSubOperacao.DataBind()

                        SessaoRecuperaListOperacao()
                        gridOperacao.DataSource = ObjListOperacao.ToArray()
                        gridOperacao.DataBind()

                        LimparOperacoes()
                        LimparSubOperacoes(False)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Operacoes", "ALTERAR") Then
                If ValidarCampos() Then
                    Dim SqlArray As New ArrayList
                    Sql = "UPDATE Operacoes set " & vbCrLf & _
                          "    Descricao            ='" & Funcoes.EliminarCaracteresEspeciais(TxtOperacaoDescricao.Text) & "'" & vbCrLf & _
                          "   ,Producao             ='" & IIf(ChkOperacaoProducao.Checked, "S", "N") & "'" & vbCrLf & _
                          "   ,UFDepositoDestino    ='" & chkEstadoDestino.Checked.ToString & "'" & vbCrLf & _
                          "   ,Classe               ='" & ddlClasse.SelectedValue & "'" & vbCrLf & _
                          "   ,UsuarioAlteracao     ='" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                          "   ,UsuarioAlteracaoData ='" & DateTime.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                          " WHERE Operacao_id = " & TxtOperacaoCodigo.Text & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) = False Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        ObjListOperacao = Nothing
                        SessaosSalvaListOperacao()

                        gridSubOperacao.DataSource = Nothing
                        gridSubOperacao.DataBind()

                        SessaoRecuperaListOperacao()
                        gridOperacao.DataSource = ObjListOperacao.ToArray()
                        gridOperacao.DataBind()

                        LimparOperacoes()
                        LimparSubOperacoes(False)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Operacoes", "EXCLUIR") Then
                Dim op As Integer = CInt(TxtOperacaoCodigo.Text)
                Dim so As New [Lib].Negocio.ListSubOperacao(op)

                If so.Count > 0 Then
                    MsgBox(Me.Page, "Operação com SubOperações não pode ser excluída.")
                Else
                    Dim SqlArray As New ArrayList
                    Sql = "Delete from Operacoes Where Operacao_Id = " & TxtOperacaoCodigo.Text
                    SqlArray.Add(Sql)
                    If Banco.GravaBanco(SqlArray) = False Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        ObjListOperacao = Nothing
                        SessaosSalvaListOperacao()

                        gridSubOperacao.DataSource = Nothing
                        gridSubOperacao.DataBind()

                        SessaoRecuperaListOperacao()
                        gridOperacao.DataSource = ObjListOperacao.ToArray()
                        gridOperacao.DataBind()

                        LimparOperacoes()
                        LimparSubOperacoes(False)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparOperacoes()
            LimparSubOperacoes(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            RelatorioOperacoes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Botoes SubOperacao"
    Protected Sub lnksubNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnksubNovo.Click
        Try
            If Funcoes.VerificaPermissao("Operacoes", "GRAVAR") Then
                If SubValidarCampos() Then
                    Dim SqlArray As New ArrayList

                    Sql = " Insert into SubOperacoes (operacao_id , SubOperacoes_id , Descricao ," & vbCrLf &
                          " EntradaSaida , Devolucao , PrecoFixo ,Laudo , EstoqueInicial , EstoqueFisico ," & vbCrLf &
                          " EstoqueFiscal , QuantidadeFisico , QuantidadeFiscal ,QuantidadePedido ," & vbCrLf &
                          " UnitarioPedido , GrupoDeContas , Financeiro , Contabil , " & vbCrLf &
                          " ApuracaoDeCustos , ApuracaodeCustosContraPartida, Classe, Deposito," & vbCrLf &
                          " situacao, ControlarPecas, OperacaoDestino, SubOperacaoDestino," & vbCrLf &
                          " consignacao, AmostraGratis, Memorando, Pedido, CobraServico, Liminar, ProdutoDeTerceiro," & vbCrLf &
                          " ParaFinsDeExportacao, FinalidadeDaNota, ContaDeAdiantamento, Representante, UsuarioInclusao, UsuarioInclusaoData,ControlarNumeroDoLote,ProprietarioDaMercadoria, NotaDebito, NotaCredito)" & vbCrLf &
                          " values (" & vbCrLf &
                          txtCodigoOperacaoSuboperacao.Text & "," & txtCodigoSuboperacao.Text & ",'" & TxtDescricaoSubOperacao.Text & "'" & vbCrLf &
                          ",'" & IIf(RbtEntrada.Checked, "E", "S") & "','" & IIf(ChkDevolucao.Checked, "S", "N") & "','" & IIf(ChkPrecoFixo.Checked, "S", "N") & "'" & vbCrLf &
                          ",'" & IIf(ChkLaudo.Checked, "S", "N") & "','" & IIf(ChkEstoqueInicial.Checked, "S", "N") & "','" & IIf(ChkEstoqueFisico.Checked, "S", "N") & "'" & vbCrLf &
                          ",'" & IIf(ChkEstoqueFiscal.Checked, "S", "N") & "','" & IIf(ChkQuantidadeFisica.Checked, "S", "N") & "','" & IIf(ChkQuantidadeFiscal.Checked, "S", "N") & "'" & vbCrLf &
                          ",'" & IIf(ChkQuantidadePedido.Checked, "S", "N") & "','" & IIf(ChkUnitarioPedido.Checked, "S", "N") & "','" & DdlGrupoDeContas.SelectedValue & "'" & vbCrLf &
                          ",'" & IIf(ChkFinanceiro.Checked, "S", "N") & "','" & IIf(ChkContabil.Checked, "S", "N") & "'" & vbCrLf

                    If DdlTabelaDeApuracaoDeCustos.SelectedIndex > 0 Then
                        Sql &= "," & DdlTabelaDeApuracaoDeCustos.SelectedValue & vbCrLf
                    Else
                        Sql &= ",0" & vbCrLf
                    End If

                    If ddlContraPartidaCusto.SelectedIndex > 0 Then
                        Sql &= "," & ddlContraPartidaCusto.SelectedValue & vbCrLf
                    Else
                        Sql &= ",0" & vbCrLf
                    End If

                    Dim sbop As String = 0
                    If ddlSubOperacaoDestino.SelectedIndex > 0 Then
                        sbop = ddlSubOperacaoDestino.SelectedValue.Split("-")(1)
                    End If

                    Sql &= ",'" & CmbClasse.SelectedValue & "', '" & IIf(ChkDeposito.Checked, "S", "N") & "'" & vbCrLf &
                           "," & ddlSituacao.SelectedIndex & "," & IIf(chkPecas.Checked, 1, 0) & "" & vbCrLf &
                           "," & IIf(ddlOperacaoDestino.SelectedIndex = 0, 0, ddlOperacaoDestino.SelectedValue) & ", " & sbop & ", " & IIf(chkConsignacao.Checked, 1, 0) & ", " & IIf(chkAmostraGratis.Checked, 1, 0) & vbCrLf &
                           "," & IIf(chkMemorando.Checked, 1, 0) & "," & IIf(chkPedido.Checked, 1, 0) & ",'" & IIf(chkCobraServico.Checked, "S", "N") & "'" & vbCrLf &
                           "," & IIf(chkLiminar.Checked, 1, 0) & "," & IIf(chkProdutoDeTerceiros.Checked, 1, 0) & vbCrLf &
                           "," & IIf(chkParaFinsDeExportacao.Checked, 1, 0) & "," & ddlFinalidadeDaNota.SelectedValue & vbCrLf &
                           ",'" & ddlContaAdiantamento.SelectedValue & "'," & IIf(chkRepresentante.Checked, 1, 0) & vbCrLf &
                           ",'" & UsuarioServidor.NomeUsuario & "', '" & DateTime.Now.ToString("yyyy-MM-dd") & "'," & IIf(chkNumeroDoLote.Checked, 1, 0) & "," & IIf(chkProprietarioDaMercadoria.Checked, 1, 0) & vbCrLf &
                           ",'" & IIf(chkNotaDebito.Checked, "S", "N") & "','" & IIf(chkNotaCredito.Checked, "S", "N") & "')" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) = False Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        LimparSubOperacoes(True)
                        ObjOperacao.SubOperacoes = Nothing
                        gridSubOperacao.DataSource = ObjOperacao.SubOperacoes
                        gridSubOperacao.DataBind()
                        SessaosSalvaOperacao()
                        TabContainer1.ActiveTabIndex = 0
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnksubAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnksubAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Operacoes", "ALTERAR") Then
                If ValidarCampos() Then
                    Dim SqlArray As New ArrayList

                    Sql = "UPDATE SubOperacoes set Descricao = '" & TxtDescricaoSubOperacao.Text & "', EntradaSaida = '" & IIf(RbtEntrada.Checked, "E", "S") & "'" & vbCrLf &
                          ", Devolucao = '" & IIf(ChkDevolucao.Checked, "S", "N") & "', PrecoFixo = '" & IIf(ChkPrecoFixo.Checked, "S", "N") & "'" & vbCrLf &
                          ", Laudo = '" & IIf(ChkLaudo.Checked, "S", "N") & "', EstoqueInicial = '" & IIf(ChkEstoqueInicial.Checked, "S", "N") & "'" & vbCrLf &
                          ", EstoqueFisico = '" & IIf(ChkEstoqueFisico.Checked, "S", "N") & "', EstoqueFiscal = '" & IIf(ChkEstoqueFiscal.Checked, "S", "N") & "'" & vbCrLf &
                          ", QuantidadeFisico = '" & IIf(ChkQuantidadeFisica.Checked, "S", "N") & "', QuantidadeFiscal = '" & IIf(ChkQuantidadeFiscal.Checked, "S", "N") & "'" & vbCrLf &
                          ", QuantidadePedido = '" & IIf(ChkQuantidadePedido.Checked, "S", "N") & "', UnitarioPedido = '" & IIf(ChkUnitarioPedido.Checked, "S", "N") & "'" & vbCrLf &
                          ", GrupoDeContas = '" & DdlGrupoDeContas.SelectedValue & "', Financeiro = '" & IIf(ChkFinanceiro.Checked, "S", "N") & "', Classe = '" & CmbClasse.SelectedValue & "'" & vbCrLf &
                          ", Contabil = '" & IIf(ChkContabil.Checked, "S", "N") & "', Deposito = '" & IIf(ChkDeposito.Checked, "S", "N") & "'" & vbCrLf &
                          ", NotaDebito = '" & IIf(chkNotaDebito.Checked, "S", "N") & "', NotaCredito = '" & IIf(chkNotaCredito.Checked, "S", "N") & "'" & vbCrLf

                    If DdlTabelaDeApuracaoDeCustos.SelectedIndex > 0 Then
                        Sql &= ", ApuracaoDeCustos = " & DdlTabelaDeApuracaoDeCustos.SelectedValue & vbCrLf
                    Else
                        Sql &= ", ApuracaoDeCustos = 0 " & vbCrLf
                    End If
                    If ddlContraPartidaCusto.SelectedIndex > 0 Then
                        Sql &= ", ApuracaodeCustosContraPartida = " & ddlContraPartidaCusto.SelectedValue & vbCrLf
                    Else
                        Sql &= ", ApuracaodeCustosContraPartida = 0 " & vbCrLf
                    End If

                    Dim sbop As String = 0
                    If ddlSubOperacaoDestino.SelectedIndex > 0 Then
                        sbop = ddlSubOperacaoDestino.SelectedValue.Split("-")(1)
                    End If

                    Sql &= " ,Situacao                 = " & ddlSituacao.SelectedIndex & vbCrLf &
                           " ,ControlarPecas           = " & IIf(chkPecas.Checked, 1, 0) & vbCrLf &
                           " ,Consignacao              = " & IIf(chkConsignacao.Checked, 1, 0) & vbCrLf &
                           " ,AmostraGratis            = " & IIf(chkAmostraGratis.Checked, 1, 0) & vbCrLf &
                           " ,Memorando                = " & IIf(chkMemorando.Checked, 1, 0) & vbCrLf &
                           " ,Pedido                   = " & IIf(chkPedido.Checked, 1, 0) & vbCrLf &
                           " ,CobraServico             ='" & IIf(chkCobraServico.Checked, "S", "N") & "'" & vbCrLf &
                           " ,Liminar                  = " & IIf(chkLiminar.Checked, 1, 0) & vbCrLf &
                           " ,ProdutoDeTerceiro        = " & IIf(chkProdutoDeTerceiros.Checked, 1, 0) & vbCrLf &
                           " ,ParaFinsDeExportacao     = " & IIf(chkParaFinsDeExportacao.Checked, 1, 0) & vbCrLf &
                           " ,OperacaoDestino          = " & IIf(ddlOperacaoDestino.SelectedIndex = 0, 0, ddlOperacaoDestino.SelectedValue) & vbCrLf &
                           " ,SubOperacaoDestino       = " & sbop & vbCrLf &
                           " ,FinalidadeDaNota         = " & ddlFinalidadeDaNota.SelectedValue & vbCrLf &
                           " ,ContaDeAdiantamento      = '" & ddlContaAdiantamento.SelectedValue & "'" & vbCrLf &
                           " ,Representante            = " & IIf(chkRepresentante.Checked, 1, 0) & vbCrLf &
                           " ,UsuarioAlteracao         = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf &
                           " ,UsuarioAlteracaoData     = '" & DateTime.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                           " ,ControlarNumeroDoLote    = " & IIf(chkNumeroDoLote.Checked, 1, 0) & vbCrLf &
                           " ,ProprietarioDaMercadoria = " & IIf(chkProprietarioDaMercadoria.Checked, 1, 0) & vbCrLf &
                           "  WHERE (Operacao_id   = " & txtCodigoOperacaoSuboperacao.Text & " and SubOperacoes_id = " & txtCodigoSuboperacao.Text & ")" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) = False Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        LimparSubOperacoes(True)
                        ObjOperacao.SubOperacoes = Nothing
                        gridSubOperacao.DataSource = ObjOperacao.SubOperacoes
                        gridSubOperacao.DataBind()
                        SessaosSalvaOperacao()
                        TabContainer1.ActiveTabIndex = 0
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnksubExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnksubExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Operacoes", "EXCLUIR") Then
                Sql = "Select Top 10 Operacao, SubOperacao From NotasFiscais " & vbCrLf &
                      "Where Operacao = " & TxtOperacaoCodigo.Text & " and SubOperacao = " & txtCodigoSuboperacao.Text

                Dim dsNf As New DataSet
                dsNf = Banco.ConsultaDataSet(Sql, "ConsultaNF")

                If Not dsNf Is Nothing AndAlso dsNf.Tables(0).Rows.Count > 0 Then
                    MsgBox(Me.Page, "Operação com lancamento de notas fiscais não pode ser excluída.")
                    Exit Sub
                End If

                Dim ped As New [Lib].Negocio.Pedidos()
                ped.Selecionar(10, 0, "", "", "", 0, "", 0, eSituacao.Normal, "", TxtOperacaoCodigo.Text, txtCodigoSuboperacao.Text)

                If ped.Count > 0 Then
                    MsgBox(Me.Page, "Operação com lançamento de pedidos não pode ser excluída.")
                    Exit Sub
                End If

                Sql = "Select Operacao_Id, SubOperacao_Id From Producao " & vbCrLf &
                      "Where Operacao_Id = " & TxtOperacaoCodigo.Text & " and SubOperacao_Id = " & txtCodigoSuboperacao.Text

                Dim dsPro As New DataSet
                dsPro = Banco.ConsultaDataSet(Sql, "ConsultaPro")

                If Not dsPro Is Nothing AndAlso dsPro.Tables(0).Rows.Count > 0 Then
                    MsgBox(Me.Page, "Operação com lançamento na produção não pode ser excluída.")
                    Exit Sub
                End If

                Dim SqlArray As New ArrayList

                Sql = "Delete from SubOperacoes Where (Operacao_id = " & txtCodigoOperacaoSuboperacao.Text & " and SubOperacoes_id = " & txtCodigoSuboperacao.Text & ")"
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    LimparSubOperacoes(True)
                    ObjOperacao.SubOperacoes = Nothing
                    gridSubOperacao.DataSource = ObjOperacao.SubOperacoes
                    gridSubOperacao.DataBind()
                    SessaosSalvaOperacao()
                    TabContainer1.ActiveTabIndex = 0
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnksubLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnksubLimpar.Click
        Try
            LimparSubOperacoes(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnksubRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnksubRelatorio.Click
        Try
            RelatorioSubOperacoes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Relatorios"
    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(TxtOperacaoCodigo.Text) Then
            param &= "Código: " & TxtOperacaoCodigo.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(TxtOperacaoDescricao.Text) Then
            param &= "Descrição: " & TxtOperacaoDescricao.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlClasse.SelectedValue) Then
            param &= "Classe: " & ddlClasse.SelectedValue
        End If

        Return param
    End Function

    Private Function getSubParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtCodigoOperacaoSuboperacao.Text) Then
            param &= "Operação: " & txtCodigoOperacaoSuboperacao.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtCodigoSuboperacao.Text) Then
            param &= "Sub Operação: " & txtCodigoSuboperacao.Text
        End If
        If Not String.IsNullOrWhiteSpace(TxtDescricaoSubOperacao.Text) Then
            param &= "Descrição: " & TxtDescricaoSubOperacao.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(CmbClasse.SelectedValue) Then
            param &= "Classe: " & CmbClasse.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlFinalidadeDaNota.SelectedValue) Then
            param &= "Finalidade da Nota: " & ddlFinalidadeDaNota.SelectedValue
        End If

        Return param
    End Function

    Private Function getDataSet() As DataSet
        Sql = "SELECT Operacao_Id as Codigo, Descricao, Producao, isnull(classe,'') as CodigoClasse " & vbCrLf &
                               "  FROM Operacoes " & vbCrLf &
                               " ORDER BY Operacao_Id" & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Operacoes")
    End Function

    Private Function getSubDataSet() As DataSet
        Sql = "SELECT   Operacoes.Operacao_Id as operacao, Operacoes.Descricao as OperacaoDescricao, Suboperacoes.SubOperacoes_Id as Suboperacao, " & vbCrLf &
                            "           SubOperacoes.Descricao AS SuboperacaoDescricao from Operacoes INNER JOIN SubOperacoes ON Operacoes.Operacao_Id = Suboperacoes.Operacao_Id"
        Return Banco.ConsultaDataSet(Sql, "Suboperacoes")


    End Function

    Private Sub RelatorioOperacoes()
        Try
            If Funcoes.VerificaPermissao("Operacoes", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()
                Dim parameters = New Dictionary(Of String, Object)()

                parameters.Add("ConsultaParametros", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_Operacoes", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir o relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub RelatorioSubOperacoes()
        Try
            If Funcoes.VerificaPermissao("Operacoes", "RELATORIO") Then
                Dim Ds As DataSet = getSubDataSet()
                Dim parameters = New Dictionary(Of String, Object)()

                parameters.Add("Titulo", "Relatório De Operações / SubOperações")
                parameters.Add("ConsultaParametros", getSubParam())

                Funcoes.BindReport(Me.Page, Ds, "Cr_Suboperacoes", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Limpar"
    Private Sub LimparOperacoes()
        TxtOperacaoCodigo.Enabled = True
        TxtOperacaoCodigo.Text = ""

        TxtOperacaoDescricao.Enabled = True
        TxtOperacaoDescricao.Text = ""

        chkEstadoDestino.Checked = False
        ChkOperacaoProducao.Checked = False
        ddlClasse.SelectedIndex = 0

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        gridSubOperacao.SelectedIndex = -1
        gridSubOperacao.DataBind()
        txtCodigoOperacaoSuboperacao.Text = String.Empty
        TxtDescricaoSubOperacao.Text = String.Empty
        CmbClasse.SelectedValue = String.Empty
        ddlUsuarios.Items.Clear()
        ddlSubUsuarios.Items.Clear()

        HID.Value = Guid.NewGuid().ToString
    End Sub

    Private Sub LimparSubOperacoes(ByVal Carregar As Boolean)
        pnlchks.Enabled = True

        ChkEstoqueFisico.Checked = False
        ChkQuantidadeFisica.Checked = False
        ChkQuantidadePedido.Checked = False
        RbtEntrada.Checked = False
        RbtSaida.Checked = False
        ChkEstoqueFiscal.Checked = False
        ChkQuantidadeFiscal.Checked = False
        ChkUnitarioPedido.Checked = False
        ChkDevolucao.Checked = False
        ChkEstoqueInicial.Checked = False
        ChkPrecoFixo.Checked = False
        ChkLaudo.Checked = False
        ChkFinanceiro.Checked = False
        ChkContabil.Checked = False
        ChkDeposito.Checked = False
        chkPecas.Checked = False
        chkMemorando.Checked = False
        chkConsignacao.Checked = False
        chkAmostraGratis.Checked = False
        chkCobraServico.Checked = False
        chkLiminar.Checked = False
        chkPedido.Checked = False
        chkRepresentante.Checked = False
        chkNumeroDoLote.Checked = False
        chkProprietarioDaMercadoria.Checked = False
        chkNotaDebito.Checked = False
        chkNotaCredito.Checked = False

        txtCodigoSuboperacao.Text = ""
        ddlSituacao.SelectedIndex = 0
        DdlGrupoDeContas.SelectedIndex = 0
        DdlTabelaDeApuracaoDeCustos.SelectedIndex = 0
        ddlContraPartidaCusto.SelectedIndex = 0
        ddlOperacaoDestino.SelectedIndex = 0

        txtCodigoSuboperacao.Enabled = True

        lnksubNovo.Parent.Visible = True
        lnksubAtualizar.Parent.Visible = False
        lnksubExcluir.Parent.Visible = False
        chkProdutoDeTerceiros.Checked = False
        ddlFinalidadeDaNota.SelectedValue = 1
        ddlContaAdiantamento.ClearSelection()
        'TabContainer1.ActiveTabIndex = 0
        ddlSubUsuarios.Items.Clear()

        If Carregar Then
            SessaoRecuperaOperacao()
            gridSubOperacao.DataSource = ObjOperacao.SubOperacoes.ToArray()
            gridSubOperacao.DataBind()
            gridSubOperacao.SelectedIndex = -1
        End If
    End Sub
#End Region

#Region "Validar"
    Function ValidarCampos() As Boolean
        If TxtOperacaoCodigo.Text.Length = 0 Then
            MsgBox(Me.Page, "Código da Operação não foi informado.")
            Return False
        ElseIf TxtOperacaoDescricao.Text.Length = 0 Then
            MsgBox(Me.Page, "Descrição da Operação não foi informada.")
            Return False
        ElseIf ddlClasse.SelectedIndex = -1 Then
            MsgBox(Me.Page, "Classe da Operação não foi informada.")
            Return False
        End If
        Return True
    End Function

    Function SubValidarCampos() As Boolean
        If txtCodigoOperacaoSuboperacao.Text.Length = 0 Then
            MsgBox(Me.Page, "Código da Operação não foi selecionado.")
            Return False
        ElseIf txtCodigoSuboperacao.Text.Length = 0 Then
            MsgBox(Me.Page, "Código da SubOperação não foi informada.")
            Return False
        ElseIf TxtDescricaoSubOperacao.Text.Length = 0 Then
            MsgBox(Me.Page, "Descrição da SubOperação não foi informada.")
            Return False
        ElseIf CmbClasse.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Classe da Operação não foi selecionanda.")
            Return False
        ElseIf RbtEntrada.Checked = False AndAlso RbtSaida.Checked = False Then
            MsgBox(Me.Page, "Selecione Entrada ou Saída.")
            Return False
        ElseIf ddlSituacao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Situação não foi selecionanda.")
            Return False
        ElseIf CmbClasse.SelectedValue.Equals(eClassesOperacoes.COMPRASAORDEM.ToString) OrElse CmbClasse.SelectedValue.Equals(eClassesOperacoes.CONTAEORDEM.ToString) OrElse CmbClasse.SelectedValue.Equals(eClassesOperacoes.VENDASAORDEM.ToString) Then
            If Not ChkPrecoFixo.Checked Then
                MsgBox(Me.Page, "Preço fixo é obrigatório estar selecionado para a classe: '" & CmbClasse.SelectedValue & "' selecionada.")
                Return False
            End If
        ElseIf (ChkUnitarioPedido.Checked OrElse ChkQuantidadePedido.Checked) AndAlso Not chkPedido.Checked Then
            MsgBox(Me.Page, "Para marcar: 'Unitário Pedido' ou  'quantidade pedido', a 'Op. de Pedido' também deve ser marcado.")
            Return False
        ElseIf (ChkEstoqueInicial.Checked OrElse ChkEstoqueFisico.Checked OrElse ChkEstoqueFiscal.Checked) AndAlso Not ChkQuantidadeFiscal.Checked Then
            MsgBox(Me.Page, "Para marcar Estoque: 'Inicial', 'Fisico' ou 'Fiscal', a 'Quantidade Fiscal' também deve ser marcado.")
            Return False
        ElseIf ChkQuantidadePedido.Checked AndAlso Not ChkQuantidadeFiscal.Checked Then
            MsgBox(Me.Page, "Se marcar Quantidade Pedido, a quantidade Fiscal também deve ser selecionado.")
            Return False
        End If
        Return True
    End Function
#End Region

#Region "Grid Operacao/Suboperacao"
    Protected Sub gridOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridOperacao.SelectedIndexChanged
        Try
            SessaoRecuperaListOperacao()
            ObjOperacao = ObjListOperacao(gridOperacao.SelectedIndex)
            TxtOperacaoCodigo.Text = ObjOperacao.Codigo
            TxtOperacaoDescricao.Text = ObjOperacao.Descricao

            chkEstadoDestino.Checked = ObjOperacao.UFDepositoDestino
            ChkOperacaoProducao.Checked = ObjOperacao.Producao
            ddlClasse.SelectedValue = ObjOperacao.CodigoClasse.Trim()

            SessaosSalvaOperacao()

            LimparSubOperacoes(True)

            txtCodigoOperacaoSuboperacao.Text = ObjOperacao.Codigo
            TxtDescricaoSubOperacao.Text = ObjOperacao.Descricao
            CmbClasse.SelectedValue = ObjOperacao.CodigoClasse.Trim()

            ddlUsuarios.Items.Clear()
            If Not String.IsNullOrWhiteSpace(ObjOperacao.UsuarioAlteracao) Then
                ddlUsuarios.Items.Add("Alt.- " & ObjOperacao.UsuarioAlteracao)
            End If
            ddlUsuarios.Items.Add("Inc.- " & ObjOperacao.UsuarioInclusao)

            TxtOperacaoCodigo.Enabled = False
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            ddlSubUsuarios.Items.Clear()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridSubOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridSubOperacao.SelectedIndexChanged
        Try
            LimparSubOperacoes(False)

            SessaoRecuperaOperacao()
            ObjSubOperacao = ObjOperacao.SubOperacoes.Where(Function(s) s.Codigo = gridSubOperacao.SelectedRow.Cells(2).Text()).FirstOrDefault
            SessaosSalvaSubOperacao()

            If ObjSubOperacao.Descricao.Length = 0 Then Exit Sub

            txtCodigoOperacaoSuboperacao.Text = ObjSubOperacao.CodigoOperacao
            txtCodigoSuboperacao.Text = ObjSubOperacao.Codigo
            TxtDescricaoSubOperacao.Text = ObjSubOperacao.Descricao

            RbtSaida.Checked = IIf(ObjSubOperacao.EntradaSaida = eEntradaSaida.Saida, True, False)
            RbtEntrada.Checked = IIf(ObjSubOperacao.EntradaSaida = eEntradaSaida.Entrada, True, False)
            CarregarContaAdiantamento(RbtEntrada.Checked)
            ChkDevolucao.Checked = ObjSubOperacao.Devolucao
            ChkPrecoFixo.Checked = ObjSubOperacao.PrecoFixo
            ChkLaudo.Checked = ObjSubOperacao.Laudo
            ChkEstoqueInicial.Checked = ObjSubOperacao.EstoqueInicial
            ChkEstoqueFisico.Checked = ObjSubOperacao.EstoqueFisico
            ChkEstoqueFiscal.Checked = ObjSubOperacao.EstoqueFiscal
            ChkQuantidadeFisica.Checked = ObjSubOperacao.QuantidadeFisico
            ChkQuantidadeFiscal.Checked = ObjSubOperacao.QuantidadeFiscal
            ChkQuantidadePedido.Checked = ObjSubOperacao.QuantidadePedido
            ChkUnitarioPedido.Checked = ObjSubOperacao.UnitarioPedido
            ChkFinanceiro.Checked = ObjSubOperacao.Financeiro
            ChkContabil.Checked = ObjSubOperacao.Contabil
            ChkDeposito.Checked = ObjSubOperacao.Deposito

            chkPecas.Checked = ObjSubOperacao.ControlarPecas
            chkPedido.Checked = ObjSubOperacao.Pedido
            chkCobraServico.Checked = ObjSubOperacao.CobraServico

            chkLiminar.Checked = ObjSubOperacao.Liminar
            chkProdutoDeTerceiros.Checked = ObjSubOperacao.ProdutoDeTerceiro
            chkParaFinsDeExportacao.Checked = ObjSubOperacao.ParaFinsDeExportacao

            chkNumeroDoLote.Checked = ObjSubOperacao.ControlarNumeroDoLote

            DdlTabelaDeApuracaoDeCustos.SelectedIndex = DdlTabelaDeApuracaoDeCustos.Items.IndexOf(DdlTabelaDeApuracaoDeCustos.Items.FindByValue(ObjSubOperacao.ApuracaoCustos))
            ddlContraPartidaCusto.SelectedIndex = ddlContraPartidaCusto.Items.IndexOf(ddlContraPartidaCusto.Items.FindByValue(ObjSubOperacao.ApuracaoCustosContraPartida))
            DdlGrupoDeContas.SelectedIndex = DdlGrupoDeContas.Items.IndexOf(DdlGrupoDeContas.Items.FindByValue(ObjSubOperacao.CodigoGrupoContas))
            CmbClasse.SelectedIndex = CmbClasse.Items.IndexOf(CmbClasse.Items.FindByValue(ObjSubOperacao.Classe.ToString))

            ddlSituacao.SelectedIndex = ObjSubOperacao.CodigoSituacao
            ddlContaAdiantamento.SelectedValue = ObjSubOperacao.CodigoContaAdiantamento

            chkMemorando.Checked = ObjSubOperacao.Memorando
            chkConsignacao.Checked = ObjSubOperacao.Consignacao
            chkAmostraGratis.Checked = ObjSubOperacao.AmostraGratis

            If ObjSubOperacao.CodigoOperacaoDestino > 0 Then
                ddlOperacaoDestino.SelectedIndex = ddlOperacaoDestino.Items.IndexOf(ddlOperacaoDestino.Items.FindByValue(ObjSubOperacao.CodigoOperacaoDestino))
                ddlOperacaoDestino_SelectedIndexChanged(Nothing, Nothing)
                ddlSubOperacaoDestino.SelectedIndex = ddlSubOperacaoDestino.Items.IndexOf(ddlSubOperacaoDestino.Items.FindByValue(ObjSubOperacao.CodigoOperacaoDestino & "-" & ObjSubOperacao.CodigoSuboperacaoDestino))
            End If
            ddlFinalidadeDaNota.SelectedValue = ObjSubOperacao.FinalidadeDaNota

            chkRepresentante.Checked = ObjSubOperacao.Representante
            chkRepresentante.Enabled = IIf(ObjSubOperacao.Pedido, True, False)

            chkProprietarioDaMercadoria.Checked = ObjSubOperacao.ProprietarioDaMercadoria
            chkProprietarioDaMercadoria.Enabled = IIf(ObjSubOperacao.Pedido, True, False)
            chkNotaDebito.Checked = ObjSubOperacao.NotaDebito
            chkNotaCredito.Checked = ObjSubOperacao.NotaCredito

            ddlSubUsuarios.Items.Clear()
            If Not String.IsNullOrWhiteSpace(ObjSubOperacao.UsuarioAlteracao) Then
                ddlSubUsuarios.Items.Add("Alt.- " & ObjSubOperacao.UsuarioAlteracao)
            End If
            ddlSubUsuarios.Items.Add("Inc.- " & ObjSubOperacao.UsuarioInclusao)

            txtCodigoSuboperacao.Enabled = False
            lnksubNovo.Parent.Visible = False
            lnksubAtualizar.Parent.Visible = True
            lnksubExcluir.Parent.Visible = True
            pnlchks.Enabled = False
            TabContainer1.ActiveTabIndex = 1
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

    Private Sub CarregarContaAdiantamento(ByVal Entrada As Boolean)
        Dim Where As String = String.Empty
        Where = " Adiantamento = 1 AND Cliente = 'S' AND LEFT(conta_id, 1) = " & IIf(Entrada, "1", "2")
        ddl.Carregar(ddlContaAdiantamento, CarregarDDL.Tabela.PlanoDeContas, Where)
    End Sub

    Protected Sub ddlOperacaoDestino_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlSubOperacaoDestino, CarregarDDL.Tabela.OperacaoSubOperacao, " So.Operacao_id =" & ddlOperacaoDestino.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ChkDevolucao_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ChkDevolucao.CheckedChanged
        Try
            ChkQuantidadePedido.Enabled = Not ChkDevolucao.Checked
            ChkUnitarioPedido.Enabled = Not ChkDevolucao.Checked
            chkPedido.Enabled = Not ChkDevolucao.Checked

            If CType(sender, CheckBox).Checked Then
                ChkQuantidadePedido.Checked = False
                ChkUnitarioPedido.Checked = False
                chkPedido.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkPedido_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkPedido.CheckedChanged
        Try
            ChkDevolucao.Enabled = Not chkPedido.Checked
            If CType(sender, CheckBox).Checked Then
                ChkDevolucao.Checked = False
            End If

            chkRepresentante.Enabled = chkPedido.Checked
            If Not CType(sender, CheckBox).Checked Then
                chkRepresentante.Checked = False
                chkProprietarioDaMercadoria.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RbtEntrada_CheckedChanged(sender As Object, e As EventArgs) Handles RbtEntrada.CheckedChanged
        Try
            CarregarContaAdiantamento(RbtEntrada.Checked)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RbtSaida_CheckedChanged(sender As Object, e As EventArgs) Handles RbtSaida.CheckedChanged
        Try
            CarregarContaAdiantamento(RbtEntrada.Checked)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub CmbClasse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbClasse.SelectedIndexChanged
        Try
            ValidarChecados(CmbClasse.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ValidarChecados(ByVal classe As String)
        If Not String.IsNullOrWhiteSpace(classe) Then
            SessaoRecuperaOperacao()

            If gridSubOperacao.SelectedIndex <> -1 Then
                ObjSubOperacao = ObjOperacao.SubOperacoes.Where(Function(s) s.Codigo = gridSubOperacao.SelectedRow.Cells(2).Text()).FirstOrDefault
                SessaosSalvaSubOperacao()
                If CmbClasse.SelectedValue = eClassesOperacoes.AFIXAR.ToString Then
                    chkPedido.Checked = ObjSubOperacao.Pedido
                    ChkPrecoFixo.Checked = False
                ElseIf CmbClasse.SelectedValue = eClassesOperacoes.COMPLEMENTACOES.ToString Then
                    ChkPrecoFixo.Checked = False
                    chkPedido.Checked = False
                Else
                    ChkPrecoFixo.Checked = ObjSubOperacao.PrecoFixo
                    chkPedido.Checked = ObjSubOperacao.Pedido
                End If

            Else
                If CmbClasse.SelectedValue = eClassesOperacoes.AFIXAR.ToString Then
                    'chkPedido.Checked = True
                    'ChkPrecoFixo.Checked = False
                ElseIf CmbClasse.SelectedValue = eClassesOperacoes.COMPLEMENTACOES.ToString Then
                    ChkPrecoFixo.Checked = False
                    chkPedido.Checked = False
                Else
                    ChkPrecoFixo.Checked = False
                    chkPedido.Checked = False
                End If
            End If

            'If classe = eClassesOperacoes.AFIXAR.ToString OrElse classe = "COMPLEMENTACOES" Then
            '    chkPedido.Enabled = False
            '    ChkPrecoFixo.Enabled = False
            'Else
            '    chkPedido.Enabled = True
            '    ChkPrecoFixo.Enabled = True
            'End If


        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Operacoes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class

