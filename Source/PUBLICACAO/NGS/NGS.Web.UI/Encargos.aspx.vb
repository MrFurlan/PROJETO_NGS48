Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Encargos
    Inherits BasePage

    Private ObjListEncargos As ListEncargo
    Private ObjEncargo As Encargo

#Region "Session"
    Protected Sub SessaoRecuperarListaEncargo()
        ObjListEncargos = Session("ObjListEncargos")
    End Sub

    Protected Sub SessaoSalvarListaEncargo()
        Session("ObjListEncargos") = ObjListEncargos
    End Sub

    Protected Sub SessaoRecuperarEncargo()
        ObjEncargo = Session("ObjEncargo")
    End Sub

    Protected Sub SessaoSalvarEncargo()
        Session("ObjEncargo") = ObjEncargo
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Encargos", "ACESSAR") Then
                    ddl.Carregar(DdlDebito, CarregarDDL.Tabela.PlanoDeContas, "len(conta_id) >= 7")
                    ddl.Carregar(DdlCredito, CarregarDDL.Tabela.PlanoDeContas, "len(conta_id) >= 7")

                    ObjListEncargos = New ListEncargo(True)
                    GridEncargos.DataSource = ObjListEncargos.ToArray
                    GridEncargos.DataBind()
                    SessaoSalvarListaEncargo()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCampos() As Boolean

        If chkOperacaoXEncargos.Checked And ddlOperador.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Preencha o campo operador so com % ou *.")
            Return False
        ElseIf chkOperacaoXEncargos.Checked And ddlValorOuPeso.SelectedValue = -1 Then
            MsgBox(Me.Page, "Selecione se o Encargo será calculado pelo Peso ou Valor.")
            Return False
        ElseIf chkOperacaoXEncargos.Checked And chkRetencao.Checked And ddlPessoa.SelectedValue = 0 Then
            MsgBox(Me.Page, "Selecione um Tipo de Pessoa.")
            Return False
        ElseIf TxtEncargo.Text = "" Then
            MsgBox(Me.Page, "Encargo é obrigatório.")
            Return False
        ElseIf DdlCredito.SelectedIndex = 0 And DdlDebito.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Obrigatório conta crédito ou conta débito.")
            Return False
        ElseIf (Not String.IsNullOrWhiteSpace(txtBaseCalculo.Text) AndAlso (CDec(txtBaseCalculo.Text) < 0)) Then
            MsgBox(Me.Page, "O campo 'base de cálculo' não pode ser menor que zero.")
            Return False
        ElseIf (Not String.IsNullOrWhiteSpace(txtBaseCalculo.Text) AndAlso (CDec(txtBaseCalculo.Text) > 100)) Then
            MsgBox(Me.Page, "O campo 'base de cálculo' não pode ser maior que 100.")
            Return False
        ElseIf (Not String.IsNullOrWhiteSpace(txtAliquota.Text) AndAlso (CDec(txtAliquota.Text) < 0)) Then
            MsgBox(Me.Page, "O campo 'alíquota' não pode ser menor que zero.")
            Return False
        ElseIf (Not String.IsNullOrWhiteSpace(txtAliquota.Text) AndAlso (CDec(txtAliquota.Text) > 100)) Then
            MsgBox(Me.Page, "O campo 'alíquota' não pode ser maior que 100.")
            Return False
        ElseIf chkOperacaoXEncargos.Checked AndAlso ddlPessoa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "O Encargo de Pessoa deve ser definido para Física, Jurídica ou Ambas.")
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        TxtEncargo.Text = ""
        TxtEncargo.Enabled = True
        TxtNomeEncargo.Text = ""
        txtBaseCalculo.Text = "0,00"
        txtAliquota.Text = "0,00"
        ddlOperador.SelectedIndex = 0
        chkGravaCentroDeCusto.Checked = False
        chkAtualizar.Checked = False
        chkImprimirNFE.Checked = False
        chkOperacaoXEncargos.Checked = False
        chkRetencao.Checked = False
        ddlPessoa.SelectedValue = 0
        chkObgEmpresa.Checked = False
        ddlPessoaRetencao.SelectedValue = 0
        ddlEtapa.SelectedIndex = 0
        DdlDebito.SelectedIndex = 0
        DdlCredito.SelectedIndex = 0
        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False

        ddlValorOuPeso.SelectedIndex = 0
        ddlValorOuPeso.Enabled = False
        chkRetencao.Enabled = False
        ddlOperador.Enabled = False
        ddlPessoa.Enabled = False
        chkObgEmpresa.Enabled = False
        ddlPessoaRetencao.Enabled = False
        ddlEncargoAgrupador.ClearSelection()

        ObjEncargo = New Encargo
        ObjEncargo.IUD = "I"
        SessaoSalvarEncargo()
    End Sub

    Private Sub ValidarConta()
        Try
            Dim sql As String = " select Tributo as Encargo from ContasAPagar " & vbCrLf &
                                " where Tributo = '" & TxtEncargo.Text & "'" & vbCrLf &
                                " select Tributo as Encargo from ContasAReceber " & vbCrLf &
                                " where Tributo = '" & TxtEncargo.Text & "'" & vbCrLf &
                                " select Encargo_id as Encargo from OperacaoXEstadoXEncargo " & vbCrLf &
                                " where Encargo_Id = '" & TxtEncargo.Text & "'"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "OperacaoXEstadoXEncargo")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                lnkNovo.Parent.Visible = False
                lnkExcluir.Parent.Visible = False

                MsgBox(Me.Page, "Encargo não pode ser alterado/excluído pois foi contabilizado.")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub GridEncargos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridEncargos.SelectedIndexChanged
        Try
            Limpar()
            SessaoRecuperarListaEncargo()
            ObjEncargo = ObjListEncargos(GridEncargos.SelectedIndex)
            ObjEncargo.IUD = "U"
            SessaoSalvarEncargo()

            TxtEncargo.Text = ObjEncargo.Codigo
            TxtEncargo.Enabled = False
            TxtNomeEncargo.Text = ObjEncargo.Descricao

            ddlPessoa.SelectedValue = ObjEncargo.TipoPessoa
            chkObgEmpresa.Checked = ObjEncargo.VerificaEmpresa

            DdlDebito.SelectedValue = ObjEncargo.ContaDebito
            DdlCredito.SelectedValue = ObjEncargo.ContaCredito

            txtBaseCalculo.Text = ObjEncargo.BaseCalculo
            txtAliquota.Text = ObjEncargo.Aliquota
            chkOperacaoXEncargos.Checked = ObjEncargo.OperacaoXEncargo
            chkRetencao.Checked = ObjEncargo.PodeSofreRetencao
            ddlPessoaRetencao.SelectedValue = ObjEncargo.TipoPessoaRetencao
            chkAtualizar.Checked = ObjEncargo.Atualizacao
            chkGravaCentroDeCusto.Checked = ObjEncargo.GravaCentroDeCusto

            ddlEtapa.SelectedValue = ObjEncargo.Etapa
            chkImprimirNFE.Checked = ObjEncargo.ImprimirNFE
            ddlEncargoAgrupador.SelectedValue = ObjEncargo.EncargoAgrupador
            ddlOperador.SelectedValue = ObjEncargo.Operador
            ddlValorOuPeso.SelectedValue = ObjEncargo.ValorOuPeso

            ddlOperador.Enabled = ObjEncargo.OperacaoXEncargo
            ddlValorOuPeso.Enabled = ObjEncargo.OperacaoXEncargo
            chkRetencao.Enabled = ObjEncargo.OperacaoXEncargo
            ddlPessoa.Enabled = ObjEncargo.OperacaoXEncargo
            chkObgEmpresa.Enabled = ObjEncargo.OperacaoXEncargo
            ddlPessoaRetencao.Enabled = ObjEncargo.OperacaoXEncargo

            lnkNovo.Parent.Visible = True
            lnkExcluir.Parent.Visible = True

            ValidarConta()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Encargos", "GRAVAR") Then
                If ValidaCampos() Then
                    SessaoRecuperarEncargo()

                    If ObjEncargo.IUD = "I" Then ObjEncargo.Codigo = Funcoes.EliminarCaracteresEspeciais(TxtEncargo.Text)

                    ObjEncargo.Descricao = TxtNomeEncargo.Text

                    ObjEncargo.TipoPessoa = ddlPessoa.SelectedValue
                    ObjEncargo.VerificaEmpresa = chkObgEmpresa.Checked
                    ObjEncargo.ContaDebito = DdlDebito.SelectedValue
                    ObjEncargo.ContaCredito = DdlCredito.SelectedValue

                    ObjEncargo.BaseCalculo = txtBaseCalculo.Text
                    ObjEncargo.Aliquota = txtAliquota.Text

                    ObjEncargo.OperacaoXEncargo = chkOperacaoXEncargos.Checked
                    ObjEncargo.PodeSofreRetencao = chkRetencao.Checked
                    ObjEncargo.TipoPessoaRetencao = ddlPessoaRetencao.SelectedValue

                    ObjEncargo.GravaCentroDeCusto = chkGravaCentroDeCusto.Checked
                    ObjEncargo.Atualizacao = chkAtualizar.Checked

                    ObjEncargo.Etapa = ddlEtapa.SelectedValue
                    ObjEncargo.ImprimirNFE = chkImprimirNFE.Checked
                    ObjEncargo.EncargoAgrupador = ddlEncargoAgrupador.SelectedValue
                    ObjEncargo.Operador = ddlOperador.SelectedValue
                    ObjEncargo.ValorOuPeso = ddlValorOuPeso.SelectedValue


                    If ObjEncargo.Salvar Then
                        SessaoRecuperarListaEncargo()
                        If ObjEncargo.IUD = "I" Then
                            ObjListEncargos.Add(ObjEncargo)
                        End If
                        ObjEncargo.IUD = ""
                        MsgBox(Me.Page, "Registro Gravado com Sucesso.", eTitulo.Sucess)
                        SessaoSalvarListaEncargo()

                        GridEncargos.DataSource = ObjListEncargos.ToArray
                        GridEncargos.DataBind()
                        Limpar()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar o registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Encargos", "EXCLUIR") Then
                SessaoRecuperarEncargo()
                If ObjEncargo.IUD <> "U" Then
                    Limpar()
                    Exit Sub
                End If

                Dim sql As String
                If Not String.IsNullOrWhiteSpace(TxtEncargo.Text) AndAlso TxtEncargo.Enabled = False Then
                    sql = "select sum(sb.x)" & vbCrLf & _
                          "  from (" & vbCrLf & _
                          "		select Count(*) x" & vbCrLf & _
                          "		  from pedidosxencargos" & vbCrLf & _
                          "		 where encargo_id = '" & ObjEncargo.Codigo & "'" & vbCrLf & _
                          "		 union all" & vbCrLf & _
                          "		select Count(*)" & vbCrLf & _
                          "		  from notasfiscaisxencargos" & vbCrLf & _
                          "		 where encargo_id = '" & ObjEncargo.Codigo & "'" & vbCrLf & _
                          "		 union all" & vbCrLf & _
                          "		select count(*)" & vbCrLf & _
                          "		  from OperacaoXEstadoxEncargo" & vbCrLf & _
                          "		 where Encargo_Id = '" & ObjEncargo.Codigo & "'" & vbCrLf
                    If FinanceiroNovo Then
                        sql &= "       ) sb" & vbCrLf
                    Else
                        sql &= "		 union all" & vbCrLf & _
                               "		select count(*)" & vbCrLf & _
                               "		  from ContasaReceber" & vbCrLf & _
                               "		 where tributo = '" & ObjEncargo.Codigo & "'" & vbCrLf & _
                               "		 union all" & vbCrLf & _
                               "		select count(*)" & vbCrLf & _
                               "		  from Contasapagar" & vbCrLf & _
                               "		 where tributo = '" & ObjEncargo.Codigo & "'" & vbCrLf & _
                               "       ) sb" & vbCrLf
                    End If

                    Dim ds As DataSet
                    ds = Banco.ConsultaDataSet(sql, "ExisteMovimentoComEsseEncargo")
                    If CInt(ds.Tables(0).Rows(0)(0)) > 0 Then
                        MsgBox(Me.Page, "Este encargo já esta sendo usado e não pode ser deletado, entre em contato com o setor de desenvolvimento, caso realmente tenha que deletar este encargo.")
                        Exit Sub
                    End If

                    ObjEncargo.IUD = "D"

                    SessaoRecuperarListaEncargo()
                    ObjListEncargos.Remove(ObjListEncargos.Where(Function(s) s.Codigo = ObjEncargo.Codigo).FirstOrDefault)
                    SessaoSalvarListaEncargo()

                    If ObjEncargo.Salvar Then
                        SessaoRecuperarListaEncargo()
                        ObjListEncargos.Remove(ObjListEncargos.Where(Function(s) s.Codigo = ObjEncargo.Codigo).FirstOrDefault)
                        SessaoSalvarListaEncargo()

                        GridEncargos.DataSource = ObjListEncargos.ToArray
                        GridEncargos.DataBind()

                        MsgBox(Me.Page, "Encargo excluído com sucesso.", eTitulo.Info, False)
                        Limpar()
                    End If
                Else
                    MsgBox(Me.Page, "Selecione o encargo para efetuar a exclusão.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir o registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("Encargos", "RELATORIO") Then
                Dim sql As String
                sql = "SELECT Encargo_id, Descricao, ContaDebito, ContaCredito, BaseCalculo, Aliquota, " & vbCrLf & _
                      "       case when isnull(GravaCentroDeCusto,0) = 0 then 'NÃO' else 'SIM' end as GravaCentroDeCusto, " & vbCrLf & _
                      "       case when isnull(OperacaoXEncargo, 'N') = 'N' then 'NÃO' else 'SIM' end as OperacaoXEncargo, " & vbCrLf & _
                      "       case when Atualizacao = 0 then 'NÃO' else 'SIM' end as Atualizacao, " & vbCrLf & _
                      "       case Etapa  when  0 then 'Normal' when 1 then 'Adiantamento' when 2 then 'Amortização' when  3 then 'Liquidação' end as Etapa, " & vbCrLf & _
                      "       case when ISNULL(ImprimirNFE,0)=0 then 'NÃO' else 'SIM' end as ImprimirNFE, " & vbCrLf & _
                      "       Operador" & vbCrLf & _
                      "  FROM Encargos" & vbCrLf & _
                      " ORDER BY Encargo_id" & vbCrLf
                Dim dsEncargo As DataSet = Banco.ConsultaDataSet(sql, "Encargo")

                Funcoes.BindReport(Me.Page, dsEncargo, "Cr_Encargos", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Encargos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub chkOperacaoXEncargos_CheckedChanged(sender As Object, e As EventArgs) Handles chkOperacaoXEncargos.CheckedChanged

        SessaoRecuperarEncargo()
        ObjEncargo.OperacaoXEncargo = chkOperacaoXEncargos.Checked
        SessaoSalvarEncargo()

        ddlOperador.Enabled = chkOperacaoXEncargos.Checked
        ddlValorOuPeso.Enabled = chkOperacaoXEncargos.Checked
        ddlPessoa.Enabled = chkOperacaoXEncargos.Checked
        chkObgEmpresa.Enabled = chkOperacaoXEncargos.Checked
        chkRetencao.Enabled = chkOperacaoXEncargos.Checked
        ddlPessoaRetencao.Enabled = chkOperacaoXEncargos.Checked

        If Not chkOperacaoXEncargos.Checked Then
            ddlOperador.SelectedIndex = 0
            ddlValorOuPeso.SelectedValue = 0
            ddlPessoa.SelectedValue = 0
            chkObgEmpresa.Checked = False
            chkRetencao.Checked = False
            ddlPessoaRetencao.SelectedValue = 0
        End If
    End Sub

    Protected Sub GridEncargos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridEncargos.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                CType(e.Row.FindControl("lblCC"), Label).Text = IIf(CBool(CType(e.Row.FindControl("lblCC"), Label).Text), "Sim", "Não")
                CType(e.Row.FindControl("lblOpEnc"), Label).Text = IIf(CBool(CType(e.Row.FindControl("lblOpEnc"), Label).Text), "Sim", "Não")
                CType(e.Row.FindControl("lblAtua"), Label).Text = IIf(CBool(CType(e.Row.FindControl("lblAtua"), Label).Text), "Sim", "Não")
                CType(e.Row.FindControl("lblNfe"), Label).Text = IIf(CBool(CType(e.Row.FindControl("lblNfe"), Label).Text), "Sim", "Não")
                CType(e.Row.FindControl("lblVerEmp"), Label).Text = IIf(CBool(CType(e.Row.FindControl("lblVerEmp"), Label).Text), "Sim", "Não")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub ddlPessoa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlPessoa.SelectedIndexChanged
        If ddlPessoa.SelectedValue = 1 Or ddlPessoa.SelectedValue = 2 Then
            If chkRetencao.Checked Then
                ddlPessoaRetencao.SelectedValue = ddlPessoa.SelectedValue
                ddlPessoaRetencao.Enabled = False
            Else
                ddlPessoaRetencao.SelectedValue = 0
            End If
        Else
            ddlPessoaRetencao.Enabled = chkRetencao.Checked
        End If
    End Sub

    Protected Sub chkRetencao_CheckedChanged(sender As Object, e As EventArgs) Handles chkRetencao.CheckedChanged
        If Not chkRetencao.Checked Then ddlPessoaRetencao.SelectedValue = 0
    End Sub
End Class