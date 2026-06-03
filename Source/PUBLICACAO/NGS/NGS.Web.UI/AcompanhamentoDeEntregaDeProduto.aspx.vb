Imports NGS.Lib.Negocio

Public Class AcompanhamentoDeEntregaDeProduto
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AcompanhamentoDeEntregaDeProduto", "ACESSAR") Then
                    HID.Value = Guid.NewGuid.ToString()
                    ucConsultaProduto.SetarHID(HID.Value)
                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
                    CarregaClasse()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProduto_Click(sender As Object, e As EventArgs) Handles btnProduto.Click
        Try
            Session("Where" & HID.Value) = " Situacao = 1 "
            ucConsultaProduto.Limpar()
            ucConsultaProduto.BuscarProduto()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProduto" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("AcompanhamentoDeEntregaDeProduto", "RELATORIO") Then
                If validaCampos() Then
                    Dim ds As DataSet = getDataSet()
                    ds.Tables(1).TableName = "EntregasTotais"

                    If Pdf Then
                        Dim parameters As Dictionary(Of String, Object) = getParameters()
                        parameters.Add("Empresa", Funcoes.FormatarCpfCnpj(UsuarioServidor.CodigoEmpresa) & "-" & UsuarioServidor.NomeEmpresa & "-" & UsuarioServidor.CidadeEmpresa & "/" & UsuarioServidor.EstadoEmpresa)
                        parameters.Add("QtdeFiscUltDias", "Qtde Física Ult. " & txtMovimentoDia.Text & " Dias")
                        parameters.Add("QtdeFiscUltDiasMaior", "Qtde Física Ult. " & txtMovimentoDia.Text & " Dias")
                        parameters.Add("QtdeFiscUltDiasMenor", "Qtde Física Ult. " & txtUltimoDiaMovimentoNota.Text & " Dias")
                        Funcoes.BindReport(Me.Page, ds, "Cr_AcompanhamentoDeEntregaDeProduto", eExportType.PDF, parameters)
                    Else
                        Dim param As New Dictionary(Of String, eTipoCampo)
                        param.Add("UltLaudo", eTipoCampo.Data)
                        param.Add("UltimaNota", eTipoCampo.Data)
                        param.Add("QtdeFisica", eTipoCampo.Numerico)
                        Funcoes.BindExcelOffice(Me.Page, ds, "Retenção Clientes", param)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If validaCampos() Then
                Dim ds As DataSet = getDataSet()
                grdRetencao.DataSource = ds
                grdRetencao.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AcompanhamentoDeEntregaDeProduto")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Session("objProduto" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = Session("objProduto" & HID.Value)
            txtNomeProduto.Text = objProduto.Descricao
            txtCodigoProduto.Value = objProduto.Codigo
            Session.Remove("objProduto" & HID.Value)
        End If
    End Sub

    Private Sub CarregaClasse()
        Try
            Dim sql As String = " select Classe_Id from ClassesDeOperacoes Where Operacao = 1"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Classes")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                lstClasses.Items.Clear()
                For Each row As DataRow In ds.Tables(0).Rows
                    lstClasses.Items.Add(New ListItem(row("Classe_Id"), row("Classe_Id")))
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function getParameters() As Dictionary(Of String, Object)
        Dim param As New Dictionary(Of String, Object)
        param.Add("ParametersConsulta", "Empresa" & IIf(chkConsolidarEmpresa.Checked, " Consolidada : ", ": ") & ddlEmpresa.SelectedItem.Text & vbCrLf)
        param("ParametersConsulta") &= "Classes de Operações: " & String.Join(", ", getListClasses()) & vbCrLf
        param("ParametersConsulta") &= "Produto: " & txtCodigoProduto.Value & "-" & txtNomeProduto.Text & vbCrLf
        param("ParametersConsulta") &= "Com Movimento nos últimos: " & txtMovimentoDia.Text & " dias. / "
        param("ParametersConsulta") &= "Última nota há: " & txtUltimoDiaMovimentoNota.Text & " dias."
        Return param
    End Function

    Private Function getDataSet() As DataSet
        Dim Classes As List(Of String) = getListClasses()

        Dim sql As String = "  SELECT Pe.Empresa_Id,                                                                                   " & vbCrLf & _
                    "          Pe.EndEmpresa_Id,                                                                                       " & vbCrLf & _
                    "          Emp.Nome as NomeEmpresa,                                                                                " & vbCrLf & _
                    "          Emp.Cidade as CidadeEmpresa,                                                                            " & vbCrLf & _
                    "          Emp.Estado as EstadoEmpresa,                                                                            " & vbCrLf & _
                    "          pe.Cliente,                                                                                             " & vbCrLf & _
                    "          pe.EndCliente,                                                                                          " & vbCrLf & _
                    "          C.Nome,                                                                                                 " & vbCrLf & _
                    "          C.Cidade,                                                                                               " & vbCrLf & _
                    "          C.Estado,                                                                                               " & vbCrLf & _
                    "          C.Complemento,                                                                                          " & vbCrLf & _
                    "          c.Telefone,                                                                                             " & vbCrLf & _
                    "          MAX(pe.EntradaBalanca) as UltimaCarga,                                                                  " & vbCrLf & _
                    "          SUM(pe.Liquido) as QtdeFisicaPeriodoTotal,                                                              " & vbCrLf & _
                    "          Sum(case                                                                                                " & vbCrLf & _
                    "                when (pe.EntradaBalanca) >=  GETDATE() - " & txtUltimoDiaMovimentoNota.Text & "                   " & vbCrLf & _
                    "                 then pe.Liquido                                                                                  " & vbCrLf & _
                    "                 else 0                                                                                           " & vbCrLf & _
                    "              end) QtdeFisicaPeriodoSemEmissao                                                                    " & vbCrLf & _
                    "     into #Base                                                                                                   " & vbCrLf & _
                    "     FROM Pesagem AS Pe                                                                                           " & vbCrLf & _
                    "   Inner Join Clientes C                                                                                          " & vbCrLf & _
                    "      on Pe.Cliente    = C.Cliente_Id                                                                             " & vbCrLf & _
                    "     and Pe.EndCliente = C.Endereco_id                                                                            " & vbCrLf & _
                    "   Inner Join Clientes emp                                                                                        " & vbCrLf & _
                    "      on Pe.Empresa_Id    = emp.Cliente_Id                                                                        " & vbCrLf & _
                    "     and Pe.EndEmpresa_Id = emp.Endereco_id                                                                       " & vbCrLf & _
                    "   Inner Join Operacoes OP                                                                                        " & vbCrLf & _
                    "      on pe.Operacao =op.Operacao_Id                                                                              " & vbCrLf & _
                    "   Where pe.Produto     = '" & txtCodigoProduto.Value & "'                                                        " & vbCrLf & _
                    "     and pe.EntradaBalanca  >= GETDATE() - " & txtMovimentoDia.Text & "                                           " & vbCrLf & _
                    "     and Op.Classe     in ('" & String.Join("','", Classes) & "')                                                 " & vbCrLf

        If chkConsolidarEmpresa.Checked Then
            sql &= "   and Left(pe.Empresa_id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf
        Else
            sql &= "   and pe.Empresa_id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'                  " & vbCrLf
        End If

        sql &= "     and pe.sequencia_id = 0                                                                                      " & vbCrLf & _
          "    group by  pe.Empresa_Id,                                                                                      " & vbCrLf & _
          "             pe.EndEmpresa_Id,                                                                                    " & vbCrLf & _
          "             Emp.Nome,                                                                                            " & vbCrLf & _
          "             Emp.Cidade,                                                                                          " & vbCrLf & _
          "             Emp.Estado,                                                                                          " & vbCrLf & _
          "             pe.Cliente,                                                                                          " & vbCrLf & _
          "             pe.EndCliente,                                                                                       " & vbCrLf & _
          "             C.Nome,                                                                                              " & vbCrLf & _
          "             C.Cidade,                                                                                            " & vbCrLf & _
          "             C.Estado,                                                                                            " & vbCrLf & _
          "             C.Complemento,                                                                                       " & vbCrLf & _
          "             c.Telefone                                                                                           " & vbCrLf & _
          "    order by c.Nome                                                                                               " & vbCrLf & _
          "                                                                                                                  " & vbCrLf & _
          "   select sum(QtdeFisicaPeriodoTotal)      as PeriodoTotal,                                                       " & vbCrLf & _
          "          sum(QtdeFisicaPeriodoSemEmissao) as PeriodoSemEmissao                                                   " & vbCrLf & _
          "     into #Totais                                                                                                 " & vbCrLf & _
          "     from #Base                                                                                                   " & vbCrLf & _
          "                                                                                                                  " & vbCrLf & _
          "   Select dbo.FormatarCpfCnpj(Empresa_Id) AS Empresa_Id, EndEmpresa_Id, NomeEmpresa, CidadeEmpresa, EstadoEmpresa, dbo.FormatarCpfCnpj(Cliente) as Cliente,  " & vbCrLf & _
          "     EndCliente, Nome, Cidade, Estado, Complemento, Telefone, UltimaCarga, QtdeFisicaPeriodoTotal, QtdeFisicaPeriodoSemEmissao" & vbCrLf & _
          "     from #base                                                                                                   " & vbCrLf & _
          "    where UltimaCarga <  GETDATE() - " & txtUltimoDiaMovimentoNota.Text & "                                       " & vbCrLf & _
          "                                                                                                                  " & vbCrLf & _
          "   select dbo.FormatarCpfCnpj(b.Empresa_id) AS Empresa_id,                                                        " & vbCrLf & _
          "          b.EndEmpresa_id,                                                                                        " & vbCrLf & _
          "          b.NomeEmpresa,                                                                                          " & vbCrLf & _
          "          left(b.cliente,8) as Cliente,                                                                           " & vbCrLf & _
          "          (select top 1 c.Nome From clientes c Where left(c.Cliente_id ,8) = left(b.cliente,8)) as NomeCliente,   " & vbCrLf & _
          "          sum(QtdeFisicaPeriodoTotal)      as QtdePeriodoTotal,                                                   " & vbCrLf & _
          "          sum(QtdeFisicaPeriodoTotal)      * 100 / case when #Totais.PeriodoTotal = 0 then 1 else #Totais.PeriodoTotal end [PeriodoTotalPerc],                       " & vbCrLf & _
          "          sum(QtdeFisicaPeriodoSemEmissao) as QtdePeriodoMenor,                                                   " & vbCrLf & _
          "          sum(QtdeFisicaPeriodoSemEmissao) * 100 / case when #Totais.PeriodoSemEmissao = 0 then 1 else #Totais.PeriodoSemEmissao end [PeriodoMenorPerc]                   " & vbCrLf & _
          "     from #base b                                                                                                 " & vbCrLf & _
          "    inner join #totais                                                                                            " & vbCrLf & _
          "       on 1 = 1                                                                                                   " & vbCrLf & _
          "    group by dbo.FormatarCpfCnpj(b.Empresa_id),                                                                   " & vbCrLf & _
          "             b.EndEmpresa_id,                                                                                     " & vbCrLf & _
          "             b.NomeEmpresa,                                                                                       " & vbCrLf & _
          "             left(b.cliente,8),                                                                                   " & vbCrLf & _
          "             #Totais.PeriodoTotal,                                                                                " & vbCrLf & _
          "             #Totais.PeriodoSemEmissao                                                                            " & vbCrLf & _
          "       order by 5                                                                                                 " & vbCrLf

        Return Banco.ConsultaDataSet(sql, "EntregasPorPeriodo")
    End Function

    Private Function getListClasses() As List(Of String)
        Try
            Dim lst As New List(Of String)

            For Each item As String In lstClasses.GetSelectedValues()
                lst.Add(item)
            Next
            Return lst
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Function validaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf lstClasses.GetSelectedValues().Count = 0 Then
            MsgBox(Me.Page, "Informe a classe de Operação")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtNomeProduto.Text) Then
            MsgBox(Me.Page, "Informe produto.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtMovimentoDia.Text) Then
            MsgBox(Me.Page, "Informe a qtde de dias para averiguação de movimentos.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtUltimoDiaMovimentoNota.Text) Then
            MsgBox(Me.Page, "Informe a quantidade de dias para aveiguação de ultimas notas entregues.")
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
        chkConsolidarEmpresa.Checked = False
        txtNomeProduto.Text = String.Empty
        txtCodigoProduto.Value = String.Empty
        txtMovimentoDia.Text = String.Empty
        txtUltimoDiaMovimentoNota.Text = String.Empty
        grdRetencao.DataBind()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

#End Region


End Class