Imports System.Data
Imports System.IO
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class PrecoMedio
    Inherits BasePage

    Dim ParametrosDaConsulta As New StringBuilder

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Comercial)
                If Funcoes.VerificaPermissao("PrecoMedio", "ACESSAR") Then
                    HID.Value = Guid.NewGuid.ToString
                    txtData.Text = Now.Date.ToString("dd/MM/yyyy")
                    ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "")
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
                    ddl.Carregar(cmbSafra, CarregarDDL.Tabela.Safra, "")
                    ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "")
                    ddlMoeda.Items.Item(0).Value = 0
                    ddlMoeda.Items.Item(0).Text = "..::Todas::.."

                    ucSelecaoProduto.WhereProduto = "Agrupar = 'N'"
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
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

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Dim crptRelatorio As New ReportDocument()

        Try
            If Funcoes.VerificaPermissao("PrecoMedio", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim dsPosicoes As New DataSet
                    Dim Texto As String = ""
                    Dim strSQL As String

                    strSQL = SqlPrecoMedio()

                    dsPosicoes = Banco.ConsultaDataSet(strSQL, "Consulta")
                    dsPosicoes.Tables(0).TableName = "ResumoPedido"
                    dsPosicoes.Tables(1).TableName = "DetalhamentoPedido"
                    dsPosicoes.Tables(2).TableName = "ResumoGeral"
                    dsPosicoes.Tables(3).TableName = "CompraAnalitica"

                    Dim Caminho As String = String.Empty
                    Dim NomeArquivo As String = String.Empty
                    Dim Arquivo As String = String.Empty
                    Dim CaminhoImagem As String = String.Empty

                    If rdSintetico.Checked Then
                        Caminho = HttpContext.Current.Server.MapPath("~/Reports/Cr_PrecoMedio.rpt")
                    Else
                        Caminho = HttpContext.Current.Server.MapPath("~/Reports/Cr_PrecoMedioAnalitico.rpt")
                    End If

                    crptRelatorio.Load(Caminho)

                    NomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & IIf(Pdf, ".PDF", ".XLS")
                    CaminhoImagem = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
                    dsPosicoes.Tables(0).Columns.Add("Imagem", GetType(System.Byte()))

                    For Each row As DataRow In dsPosicoes.Tables(0).Rows
                        row("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(CaminhoImagem)
                    Next

                    crptRelatorio.SetDataSource(dsPosicoes)

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("MostrarDetalhamento", rdAnalitico.Checked)

                    Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")
                    Dim Empresa As Cliente = New Cliente(Emp(0), Emp(1))

                    parameters.Add("NomeEmpresa", Empresa.Nome & " (" & Funcoes.FormatarCpfCnpj(Empresa.Codigo) & ")")
                    parameters.Add("CidadeEstadoEmpresa", Empresa.Cidade & "/" & Empresa.Estado.Codigo)
                    parameters.Add("ParametrosDaConsulta", ParametrosDaConsulta.ToString)

                    Funcoes.BindParameters(crptRelatorio, parameters)

                    Arquivo = HttpContext.Current.Server.MapPath(NomeArquivo)

                    If Dir(Arquivo).Length > 0 Then Kill(Arquivo)

                    If Pdf Then
                        crptRelatorio.ExportToDisk(ExportFormatType.PortableDocFormat, Arquivo)
                    Else
                        crptRelatorio.ExportToDisk(ExportFormatType.Excel, Arquivo)
                    End If

                    If IO.File.Exists(Arquivo) Then
                        If Pdf = True Then
                            ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo & "');", True)
                        Else
                            ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString, "window.location = '" & NomeArquivo & "';", True)
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crptRelatorio.Close()
            crptRelatorio.Dispose()
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidadeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaClientes.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objClientexPrecoMedio" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PrecoMedio")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub Limpar()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "")
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
        ChkTrocaFrete.Checked = False
        chkTroca.Checked = False
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        ucSelecaoProduto.CarregarNivel(5)
        cmbSafra.SelectedIndex = 0
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClientexPrecoMedio" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objClientexPrecoMedio" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtCliente.Text = pCliente.Nome
            txtCodigoCliente.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objClientexPrecoMedio" & HID.Value.ToString)
        End If
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtData.Text) OrElse Not IsDate(txtData.Text) Then
            MsgBox(Me.Page, "Informe uma data válida.")
            txtData.Focus()
            Return False
        ElseIf Not ucSelecaoProduto.TemSelecionado Then
            MsgBox(Me.Page, "Selecione ao menos 1 produto.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(cmbSafra.SelectedValue) Then
            MsgBox(Me.Page, "Selecione uma safra.")
            cmbSafra.Focus()
            Return False
        End If
        If Not ChkNormal.Checked And Not chkAntecipada.Checked And Not chkRecompra.Checked And Not chkTroca.Checked Then
            MsgBox(Me.Page, "Selecione ao menos 1 tipo de Pedido.")
            Return False
        End If
        Return True
    End Function

    Public Function SqlPrecoMedio() As String
        'O sql foi centralizado na classe NGS.Lib.Negocio\Movimentacao\PosicaoDePedido.vb
        Dim Cliente As Cliente = Nothing
        Dim par As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id")
        Dim DataEntregaInicial As String = ""
        Dim DataEntregaFinal As String = ""
        Dim DataLancamentoPedidoInicial As String = ""
        Dim DataLancamentoPedidoFinal As String = ""
        Dim SituacaoPedido As String = ""

        Dim listEmpresa As New List(Of String)

        ParametrosDaConsulta.Append("Parâmetros da Consulta").AppendLine()
        If Not String.IsNullOrWhiteSpace(ddlUnidadeNegocio.SelectedValue) Then
            ParametrosDaConsulta.Append("Unidade: " & ddlUnidadeNegocio.SelectedItem.Text).AppendLine()
        End If

        If ddlEmpresa.SelectedValue.Length > 0 Then
            listEmpresa.Add(ddlEmpresa.SelectedValue.Replace("-", ""))

            If chkConsolidarEmpresa.Checked Then
                ParametrosDaConsulta.Append("Consolidar Empresa: SIM").AppendLine()
            End If
        End If

        If txtCodigoCliente.Value.Length > 0 Then
            Cliente = New Cliente(txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1))
            If chkConsolidarCliente.Checked Then
                ParametrosDaConsulta.Append("Consolidar Cliente: SIM").AppendLine()
            End If
            ParametrosDaConsulta.Append("Cliente: " & Cliente.Nome).AppendLine()
        End If

        If Not String.IsNullOrWhiteSpace(cmbSafra.SelectedValue) Then
            ParametrosDaConsulta.Append("Safra: " & cmbSafra.SelectedValue).AppendLine()
        End If

        ParametrosDaConsulta.Append("Pedido de Normal: " & IIf(ChkNormal.Checked, "SIM", "NAO")).AppendLine()
        ParametrosDaConsulta.Append("Pedidos de Troca: " & IIf(chkTroca.Checked, "SIM", "NAO")).AppendLine()
        ParametrosDaConsulta.Append("Pedido de Compra/Venda Antecipada: " & IIf(chkAntecipada.Checked, "SIM", "NAO")).AppendLine()
        ParametrosDaConsulta.Append("Pedido de Recompra: " & IIf(chkRecompra.Checked, "SIM", "NAO")).AppendLine()

        ParametrosDaConsulta.Append("Posição no Dia: " & CDate(txtData.Text)).AppendLine()

        If rdAnalitico.Checked Then
            ParametrosDaConsulta.Append("Tipo de Relatório: Analítico").AppendLine()
        ElseIf rdSintetico.Checked Then
            ParametrosDaConsulta.Append("Tipo de Relatório: Sintético").AppendLine()
        End If

        ParametrosDaConsulta.Append("Tipo de Moeda: " & ddlMoeda.SelectedItem.Text).AppendLine()
        ParametrosDaConsulta.Append("Tipo do Frete: " & ddlFrete.SelectedItem.Text).AppendLine()
        ParametrosDaConsulta.Append(par(1)).AppendLine()

        Dim Sql As String = ""
        Dim pos As New ListPosicaoDePedido

        Sql = pos.SqlPosicao(CDate(txtData.Text), listEmpresa, chkConsolidarEmpresa.Checked, Cliente, chkConsolidarCliente.Checked, "#PM", "", cmbSafra.SelectedValue, cmbSafra.SelectedValue, par(0), DataEntregaInicial, DataEntregaFinal, DataLancamentoPedidoInicial, DataLancamentoPedidoFinal, SituacaoPedido, 0, "", "", 0, 0, 0, ddlMoeda.SelectedItem.Value)

        Sql &= " Select Case " & vbCrLf & _
               "         When isnull(SO.PrecoFixo,'N') = 'S'" & vbCrLf & _
               "           then 'SIM'" & vbCrLf & _
               "           else 'NAO'" & vbCrLf & _
               "       end PrecoFixo," & vbCrLf & _
               "       OP.Classe," & vbCrLf & _
               "       case" & vbCrLf & _
               "         when isnull(P.Antecipada,0) = 1" & vbCrLf & _
               "           then 'SIM'" & vbCrLf & _
               "           else 'NAO'" & vbCrLf & _
               "       end as Antecipado," & vbCrLf & _
               "       Case" & vbCrLf & _
               "         when isnull(P.Recompra,0) = 1" & vbCrLf & _
               "           then 'x'" & vbCrLf & _
               "           else '-'" & vbCrLf & _
               "       end as Recompra," & vbCrLf & _
               "       case" & vbCrLf & _
               "         when isnull(p.troca,0) = 1    then 'Troca'" & vbCrLf & _
               "         when isnull(pag.AVista,0) = 0 then 'A Prazo'" & vbCrLf & _
               "         else 'A Vista'" & vbCrLf & _
               "       end Pagamento," & vbCrLf & _
               "       #pm.NomeMoeda," & vbCrLf & _
               "       case" & vbCrLf & _
               "         when len(#pm.FreteCIFFOBrel) > 0" & vbCrLf & _
               "           then case when len(isnull(p.CidadeEntrega,'')) = 0 then 'Nao Informado' else p.CidadeEntrega end" & vbCrLf & _
               "           else ''" & vbCrLf & _
               "       end CidadeEntrega," & vbCrLf & _
               "       case" & vbCrLf & _
               "         when len(#pm.FreteCIFFOBrel) > 0" & vbCrLf & _
               "           then case when len(isnull(p.EstadoEntrega,'')) = 0 then 'NI' else p.EstadoEntrega end" & vbCrLf & _
               "           else ''" & vbCrLf & _
               "       end EstadoEntrega," & vbCrLf & _
               "       #pm.DataEntregaRel," & vbCrLf & _
               "       #pm.Pedido," & vbCrLf & _
               "       #pm.NomeOperacao," & vbCrLf & _
               "       #pm.Cliente," & vbCrLf & _
               "       #pm.EndCliente," & vbCrLf & _
               "       #pm.NomeCliente as nome," & vbCrLf & _
               "       #pm.cidadecliente as cidade," & vbCrLf & _
               "       #pm.Estadocliente as Estado," & vbCrLf & _
               "       isnull(pa.produto_id,#pm.Produto) as Produto," & vbCrLf & _
               "       prd.Nome as NomeProduto," & vbCrLf & _
               "       Case " & vbCrLf & _
               "         When isnull(SO.PrecoFixo,'N') = 'S'" & vbCrLf & _
               "           then #pm.Contratado" & vbCrLf & _
               "           else #pm.Fixada" & vbCrLf & _
               "       end Quantidade," & vbCrLf & _
               "       Case" & vbCrLf & _
               "		 When isnull(SO.PrecoFixo,'N') = 'S'" & vbCrLf & _
               "		   then case" & vbCrLf & _
               "				  when #pm.Moeda = 1" & vbCrLf & _
               "					then #pm.ValorTotalContratado" & vbCrLf & _
               "					else #pm.TotalMoeda" & vbCrLf & _
               "				end" & vbCrLf & _
               "		   else case" & vbCrLf & _
               "				  when #pm.Moeda = 1" & vbCrLf & _
               "					then #pm.ValorFixadoOficial" & vbCrLf & _
               "					else #pm.ValorFixadoMoeda" & vbCrLf & _
               "				end" & vbCrLf & _
               "	   end Valor," & vbCrLf & _
               "	   ((Case" & vbCrLf & _
               "			 When isnull(SO.PrecoFixo,'N') = 'S'" & vbCrLf & _
               "			   then case" & vbCrLf & _
               "					  when #pm.Moeda = 1" & vbCrLf & _
               "						then #pm.ValorTotalContratado" & vbCrLf & _
               "						else #pm.TotalMoeda" & vbCrLf & _
               "					end" & vbCrLf & _
               "			   else case" & vbCrLf & _
               "					  when #pm.Moeda = 1" & vbCrLf & _
               "						then #pm.ValorFixadoOficial" & vbCrLf & _
               "						else #pm.ValorFixadoMoeda" & vbCrLf & _
               "					end" & vbCrLf & _
               "		   end) * isnull(sbComissoes.PercComissao,0) / 100) as ValorComissao," & vbCrLf & _
               "		(Case" & vbCrLf & _
               "			 When isnull(SO.PrecoFixo,'N') = 'S'" & vbCrLf & _
               "			   then #pm.Contratado" & vbCrLf & _
               "			   else #pm.Fixada" & vbCrLf & _
               "		   end) * isnull(#pm.UnitarioMedioDoFrete,0) as ValorFrete" & vbCrLf & _
               "  into #Detalhado" & vbCrLf & _
               "  from #pm" & vbCrLf & _
               " inner join Pedidos P" & vbCrLf & _
               "    on #pm.Empresa    = P.Empresa_id" & vbCrLf & _
               "   and #pm.EndEmpresa = P.EndEmpresa_id" & vbCrLf & _
               "   and #pm.pedido     = P.Pedido_Id" & vbCrLf & _
               " inner join Operacoes OP" & vbCrLf & _
               "    on OP.Operacao_Id     = P.Operacao" & vbCrLf & _
               " inner join SubOperacoes SO" & vbCrLf & _
               "    on SO.Operacao_Id     = P.Operacao" & vbCrLf & _
               "   and so.SubOperacoes_Id = p.SubOperacao" & vbCrLf & _
               " left join Pagamentos Pag" & vbCrLf & _
               "    on Pag.Pagamento_Id = p.CondicaoPagamento" & vbCrLf & _
               " left join(Select Empresa_Id," & vbCrLf & _
               "                  EndEmpresa_Id," & vbCrLf & _
               "                  Pedido_Id," & vbCrLf & _
               "                  Sum(Percentual) as PercComissao" & vbCrLf & _
               "             from comissoes" & vbCrLf & _
               "            group by Empresa_Id," & vbCrLf & _
               "                     EndEmpresa_Id," & vbCrLf & _
               "                     Pedido_Id" & vbCrLf & _
               "           ) as sbComissoes" & vbCrLf & _
               "    on P.Empresa_Id    = sbComissoes.Empresa_Id" & vbCrLf & _
               "   and P.EndEmpresa_Id = sbComissoes.EndEmpresa_Id" & vbCrLf & _
               "   and P.Pedido_Id     = sbComissoes.Pedido_Id" & vbCrLf & _
               "  left join ProdutosAgrupados PA" & vbCrLf & _
               "    on PA.ProdutoAgrupado_Id = #pm.produto" & vbCrLf & _
               " Inner join Produtos Prd" & vbCrLf & _
               "    on Prd.Produto_Id = ISNULL(PA.Produto_Id,#pm.Produto)" & vbCrLf & _
               "  Where OP.Classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "')" & vbCrLf & _
               "    and SO.Classe <> '" & eClassesOperacoes.CONTAEORDEM.ToString & "'" & vbCrLf & _
               IIf(ddlFrete.SelectedIndex > 0, "    And isnull(P.FreteCIFFOB,'NEN') = '" & ddlFrete.SelectedValue & "'", "")


        If ChkNormal.Checked Then
            Sql &= IIf(Not chkAntecipada.Checked, "   and isnull(P.Antecipada,0) = 0", "") & vbCrLf & _
                   IIf(Not chkRecompra.Checked, "   and isnull(P.Recompra,0) = 0", "") & vbCrLf & _
                   IIf(Not chkTroca.Checked, "   and isnull(p.troca,0) = 0;", ";")
        End If

        If Not ChkNormal.Checked Then
            Dim orsql As String
            orsql = IIf(chkAntecipada.Checked, "isnull(P.Antecipada,0) = 1", "")
            orsql &= IIf(chkRecompra.Checked, IIf(orsql.Length > 0, " or ", "") & "isnull(P.Recompra,0) = 1", "")
            orsql &= IIf(chkTroca.Checked, IIf(orsql.Length > 0, " or ", "") & "isnull(p.troca,0) = 1", "")
            Sql &= " AND (" & orsql & ");"
        End If

        Sql &= "Select PrecoFixo," & vbCrLf & _
               "       Classe," & vbCrLf & _
               "       Antecipado," & vbCrLf & _
               "       Recompra," & vbCrLf & _
               "       Pagamento," & vbCrLf & _
               "       NomeMoeda," & vbCrLf & _
               "       CidadeEntrega," & vbCrLf & _
               "       EstadoEntrega," & vbCrLf & _
               "       Produto," & vbCrLf & _
               "       NomeProduto," & vbCrLf & _
               "       sum(Quantidade) as Quantidade," & vbCrLf & _
               "       sum(Valor) as Valor," & vbCrLf & _
               "       sum(ValorComissao) as ValorComissao," & vbCrLf & _
               "       sum(ValorFrete) as ValorFrete," & vbCrLf & _
               "       sum(valor) / sum(Quantidade) as UnitarioBruto," & vbCrLf & _
               "	   Case" & vbCrLf & _
               "		 when classe = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
               "		   then (sum(valor) - sum(valorComissao) - sum(valorfrete)) / sum(Quantidade)" & vbCrLf & _
               "		   else (sum(valor) + sum(valorComissao) + sum(valorfrete)) / sum(Quantidade)" & vbCrLf & _
               "	   end as UnitarioLiquido" & vbCrLf & _
               "  into #Resumo" & vbCrLf & _
               "  from #Detalhado" & vbCrLf & _
               "  Where quantidade > 0" & vbCrLf & _
               "  group by PrecoFixo," & vbCrLf & _
               "		   Classe," & vbCrLf & _
               "		   Antecipado," & vbCrLf & _
               "           Recompra," & vbCrLf & _
               "		   Pagamento," & vbCrLf & _
               "		   NomeMoeda," & vbCrLf & _
               "		   CidadeEntrega," & vbCrLf & _
               "		   EstadoEntrega," & vbCrLf & _
               "           Produto," & vbCrLf & _
               "           NomeProduto;" & vbCrLf

        Sql &= "Select PrecoFixo, Classe, Antecipado, Recompra, Pagamento, NomeMoeda, CidadeEntrega, EstadoEntrega, cast(Produto as numeric(15)) as Produto, NomeProduto," & vbCrLf & _
               "       Quantidade," & vbCrLf & _
               "       Valor," & vbCrLf & _
               "       ValorComissao," & vbCrLf & _
               "       ValorFrete," & vbCrLf & _
               "       UnitarioBruto   * 60 as UnitarioBruto," & vbCrLf & _
               "       UnitarioLiquido * 60 as UnitarioLiquido" & vbCrLf & _
               "  from #Resumo" & vbCrLf & _
               " order by 2,1" & vbCrLf


        Sql &= "Select PrecoFixo," & vbCrLf & _
               "       Classe," & vbCrLf & _
               "       Antecipado," & vbCrLf & _
               "       Recompra," & vbCrLf & _
               "       Pagamento," & vbCrLf & _
               "       NomeMoeda," & vbCrLf & _
               "       CidadeEntrega," & vbCrLf & _
               "       EstadoEntrega," & vbCrLf & _
               "       DataEntregaRel," & vbCrLf & _
               "       Pedido," & vbCrLf & _
               "	   NomeOperacao," & vbCrLf & _
               "	   Cliente," & vbCrLf & _
               "	   EndCliente," & vbCrLf & _
               "	   Nome," & vbCrLf & _
               "	   cidade," & vbCrLf & _
               "	   Estado," & vbCrLf & _
               "	   cast(Produto as numeric(15)) as Produto," & vbCrLf & _
               "	   NomeProduto," & vbCrLf & _
               "	   Quantidade," & vbCrLf & _
               "	   Valor," & vbCrLf & _
               "	   ValorComissao," & vbCrLf & _
               "	   ValorFrete," & vbCrLf & _
               "		case" & vbCrLf & _
               "		  when Quantidade = 0 " & vbCrLf & _
               "		    then 0" & vbCrLf & _
               "		    else (valor / Quantidade) * 60" & vbCrLf & _
               "		end as UnitarioBruto," & vbCrLf & _
               "		case" & vbCrLf & _
               "		  when Quantidade = 0 " & vbCrLf & _
               "		    then 0" & vbCrLf & _
               "		    else Case" & vbCrLf & _
               "				   when classe = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
               "				     then (((valor) - (valorComissao) - (valorfrete)) / (Quantidade)) * 60" & vbCrLf & _
               "				     else (((valor) + (valorComissao) + (valorfrete)) / (Quantidade)) * 60" & vbCrLf & _
               "				 end " & vbCrLf & _
               "	   end as UnitarioLiquido" & vbCrLf & _
               "  from #Detalhado" & vbCrLf & _
               " Where Quantidade > 0;" & vbCrLf

        Sql &= "Select Classe," & vbCrLf & _
               "       NomeMoeda," & vbCrLf & _
               "       cast(Produto as numeric(15)) as Produto," & vbCrLf & _
               "       NomeProduto," & vbCrLf & _
               "       sum(quantidade) as Quantidade," & vbCrLf & _
               "       sum(quantidade * unitarioLiquido) as Valor," & vbCrLf & _
               "       ((sum(quantidade * unitarioLiquido)) / sum(quantidade)) * 60 as UnitarioMedio" & vbCrLf & _
               "  from #Resumo" & vbCrLf & _
               " Group by Classe," & vbCrLf & _
               "          NomeMoeda," & vbCrLf & _
               "          cast(Produto as numeric(15))," & vbCrLf & _
               "          NomeProduto;" & vbCrLf

        Sql &= "Select op.Classe," & vbCrLf & _
               IIf(ChkTrocaFrete.Checked, "       '' as FreteCIFFOB,", "       p.FreteCIFFOB,") & vbCrLf & _
               "	   P.Moeda," & vbCrLf & _
               "	   prd.Produto_Id," & vbCrLf & _
               "	   sum(pe.TotalOficial) as Valor," & vbCrLf & _
               "	   sum(pe.TotalOficial / T.unitarioOficial /60) as Sacos" & vbCrLf & _
               "  Into #BaseTroca" & vbCrLf & _
               "  from #pm t" & vbCrLf & _
               " inner join Pedidos p" & vbCrLf & _
               "    on t.empresa    = p.Empresa_Id" & vbCrLf & _
               "   and t.endempresa = p.EndEmpresa_Id" & vbCrLf & _
               "   and t.pedido    = p.Pedido_Id" & vbCrLf & _
               " Inner join VW_PedidosXItensXFixacoes pe " & vbCrLf & _
               "    on p.empresa_id    = pe.Empresa_Id" & vbCrLf & _
               "   and p.endempresa_id = pe.EndEmpresa_Id" & vbCrLf & _
               "   and p.pedido_id     = pe.Pedido_Id" & vbCrLf & _
               " Inner join Operacoes Op" & vbCrLf & _
               "    on op.Operacao_Id = p.Operacao" & vbCrLf & _
               "  Left Join ProdutosAgrupados PA" & vbCrLf & _
               "    on PA.ProdutoAgrupado_Id = pe.Produto_Id" & vbCrLf & _
               " inner join Produtos prd" & vbCrLf & _
               "    on prd.Produto_Id = isnull(Pa.produto_id,pe.produto_Id)" & vbCrLf & _
               " Where isnull(P.Troca,0) = 1" & vbCrLf
        If Not chkTroca.Checked Then Sql &= "  and 1=2" & vbCrLf

        Sql &= " group by op.Classe," & vbCrLf & _
               IIf(ChkTrocaFrete.Checked, "", "	      p.FreteCIFFOB,") & vbCrLf & _
               "	      P.Moeda," & vbCrLf & _
               "	      prd.Produto_Id" & vbCrLf

        Sql &= "Select op.Classe," & vbCrLf & _
               IIf(ChkTrocaFrete.Checked, "	   '' as FreteCIFFOB,", "	   p.FreteCIFFOB,") & vbCrLf & _
               "	   P.Moeda," & vbCrLf & _
               "	   prd.produto_id," & vbCrLf & _
               "	   SB.Grupo_Id," & vbCrLf & _
               "	   SB.Descricao," & vbCrLf & _
               "	   sum(SB.TotalOficial) as Total," & vbCrLf & _
               "	   sum(t.contratado * sb.perc / 100) / 60 as Sacos," & vbCrLf & _
               "       sum(t.ValorTotalContratado * sb.perc / 100) VlrTroca" & vbCrLf & _
               "  Into #troca" & vbCrLf & _
               "  from #pm t" & vbCrLf & _
               " inner join Pedidos p" & vbCrLf & _
               "    on t.empresa    = p.Empresa_Id" & vbCrLf & _
               "   and t.endempresa = p.EndEmpresa_Id" & vbCrLf & _
               "   and t.pedido     = p.Pedido_Id" & vbCrLf & _
               " Inner join (Select vw.Empresa_id," & vbCrLf & _
               "                    vw.EndEmpresa_id," & vbCrLf & _
               "					vw.Pedido_id," & vbCrLf & _
               "					GE.Grupo_Id," & vbCrLf & _
               "					ge.Descricao," & vbCrLf & _
               "					P.Moeda," & vbCrLf & _
               "					sum(vw.TotalOficial) as TotalOficial," & vbCrLf & _
               "					sum(vw.TotalOficial) * 100 / vwT.TotalOficial as Perc" & vbCrLf & _
               "               from VW_PedidosXItensXFixacoes vw" & vbCrLf & _
               "              Inner join (Select Empresa_id," & vbCrLf & _
               "							 	 EndEmpresa_id," & vbCrLf & _
               "								 Pedido_id," & vbCrLf & _
               "								 sum(TotalOficial) as TotalOficial" & vbCrLf & _
               "					        from VW_PedidosXItensXFixacoes" & vbCrLf & _
               "						  Group by Empresa_id, EndEmpresa_id, Pedido_id" & vbCrLf & _
               "                          ) as vwT" & vbCrLf & _
               "                 on vwT.Empresa_Id    = vw.Empresa_Id" & vbCrLf & _
               "				and vwT.EndEmpresa_Id = vw.EndEmpresa_Id" & vbCrLf & _
               "				and vwT.Pedido_Id     = vw.Pedido_Id" & vbCrLf & _
               "              Inner join Pedidos P" & vbCrLf & _
               "			     on P.Empresa_Id    = vw.Empresa_Id" & vbCrLf & _
               "				and P.EndEmpresa_Id = vw.EndEmpresa_Id" & vbCrLf & _
               "				and P.Pedido_id     = vw.Pedido_id" & vbCrLf & _
               "              inner join produtos prd" & vbCrLf & _
               "			     on prd.Produto_Id = vw.Produto_Id" & vbCrLf & _
               "              inner join GruposDeEstoques GE" & vbCrLf & _
               "			     on GE.Grupo_Id = left(prd.grupo,2)" & vbCrLf & _
               "              Group by vw.Empresa_id," & vbCrLf & _
               "                       vw.EndEmpresa_id," & vbCrLf & _
               "					   vw.Pedido_id," & vbCrLf & _
               "					   GE.Grupo_Id," & vbCrLf & _
               "					   ge.Descricao," & vbCrLf & _
               "					   P.Moeda," & vbCrLf & _
               "                       vwT.TotalOficial" & vbCrLf & _
               "             ) Sb" & vbCrLf & _
               "    on P.EmpresaTroca    = Sb.Empresa_Id" & vbCrLf & _
               "   and P.EndEmpresaTroca = Sb.EndEmpresa_Id" & vbCrLf & _
               "   and P.PedidoTroca     = Sb.Pedido_Id" & vbCrLf & _
               " Inner join Operacoes Op" & vbCrLf & _
               "    on op.Operacao_Id = p.Operacao" & vbCrLf & _
               "  Left Join ProdutosAgrupados PA" & vbCrLf & _
               "    on PA.ProdutoAgrupado_Id = t.Produto" & vbCrLf & _
               " inner join Produtos prd" & vbCrLf & _
               "    on prd.Produto_Id = isnull(Pa.produto_id,t.produto)" & vbCrLf & _
               " Where isnull(P.Troca,0) = 1" & vbCrLf

        If Not chkTroca.Checked Then Sql &= "and 1=2" & vbCrLf

        Sql &= " group by op.Classe," & vbCrLf & _
               IIf(ChkTrocaFrete.Checked, "", "          p.FreteCIFFOB,") & vbCrLf & _
               "	      P.Moeda," & vbCrLf & _
               "	      SB.Grupo_Id," & vbCrLf & _
               "	      SB.Descricao," & vbCrLf & _
               "		  prd.produto_id" & vbCrLf

        Sql &= "Select B.Classe," & vbCrLf & _
               "       B.FreteCIFFOB," & vbCrLf & _
               "	   B.Moeda," & vbCrLf & _
               "	   cast(B.Produto_Id as numeric(15)) as Produto_Id," & vbCrLf & _
               "	   B.Valor," & vbCrLf & _
               "	   B.Sacos," & vbCrLf & _
               "	   T.Sacos * 100 / b.Sacos Perc," & vbCrLf & _
               "	   T.Grupo_Id," & vbCrLf & _
               "	   T.Descricao," & vbCrLf & _
               "	   T.Total," & vbCrLf & _
               "       T.VlrTroca," & vbCrLf & _
               "	   T.Sacos ScsTroca," & vbCrLf & _
               "	   T.VlrTroca / T.Sacos unitario" & vbCrLf & _
               "  from #BaseTroca B" & vbCrLf & _
               " inner join #troca T" & vbCrLf & _
               "    on B.Classe      = T.Classe" & vbCrLf & _
               "   and B.FreteCIFFOB = T.FreteCIFFOB" & vbCrLf & _
               "   and B.Moeda       = T.Moeda" & vbCrLf & _
               "   and B.Produto_Id  = T.Produto_id" & vbCrLf & _
               "  left join ProdutosAgrupados PA" & vbCrLf & _
               "    on PA.ProdutoAgrupado_Id = B.produto_Id" & vbCrLf
              
        Return Sql
    End Function
#End Region

End Class