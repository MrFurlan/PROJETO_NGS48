Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class NovoRelatorioDeTitulos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim Codigo As String
    Dim Descricao As String
    Dim Cliente As String
    Dim Endereco As String
    Dim index As Integer
    Dim Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioDeTitulos", "ACESSAR") Then
                HttpContext.Current.Session("ssRegistros") = ""
                HttpContext.Current.Session("ssObservacoes") = ""
                txtPeriodoInicialConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
                txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
                ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, "", True)
                ddl.Carregar(ddlCarteiraDoTitulo, CarregarDDL.Tabela.CarteiraDoTitulo, "", True)
                ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "")
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(ddlRepresentante, CarregarDDL.Tabela.ClientesXTipos, "6")
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
        AtribuirValoresCampos()
    End Sub

    Private Sub AtribuirValoresCampos()
        If Not Session("objClienteRelTiT") Is Nothing Then
            Dim objCliente = CType(Session("objClienteRelTiT"), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtNomeCliente.Text = itemCliente.Text
            hdfCodigoCliente.Value = itemCliente.Value

            Session.Remove("objClienteRelTiT")
        End If
    End Sub

    Private Sub Limpar()
        DdlUnidadeConsultaTitulos.SelectedIndex = 0
        DdlEmpresaConsultaTitulos.Items.Clear()
        txtNomeCliente.Text = ""
        hdfCodigoCliente.Value = ""
        txtPeriodoInicialConsultaTitulos.Text = ""
        txtPeriodoFinalConsultaTitulos.Text = ""

        chkUsarPeriodo.Checked = True
        pnlData.Visible = True

        txtPeriodoInicialConsultaTitulos.Text = "01/" & Now.Month & "/" & Now.Year
        txtPeriodoInicialConsultaTitulos.Text = Date.DaysInMonth(Now.Year, Now.Month) & "/" & Now.Month & "/" & Now.Year
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlUnidadeConsultaTitulos.SelectedIndexChanged
        Try
            ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, DdlUnidadeConsultaTitulos.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteTitulo" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlSafra_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkUsarPeriodo.Checked = False
            pnlData.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub radData_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles radData.CheckedChanged
        Try
            chkListarProdutos.Visible = False
            chkListarProdutos.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub radClienteData_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkListarProdutos.Visible = True
            chkListarProdutos.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RdClientePedido_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            chkListarProdutos.Visible = True
            chkListarProdutos.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkUsarPeriodo_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkUsarPeriodo.CheckedChanged
        Try
            If CType(sender, CheckBox).Checked Then
                pnlData.Visible = True
            Else
                pnlData.Visible = False
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
        Dim crpt As New ReportDocument()

        Try
            If Funcoes.VerificaPermissao("RelatorioDeTitulos", "RELATORIO") Then
                Dim DataInicial As String = ""
                Dim DataFinal As String = ""

                If chkUsarPeriodo.Checked Then
                    DataInicial = Format(CDate(txtPeriodoInicialConsultaTitulos.Text), "yyyy/MM/dd")
                    DataFinal = Format(CDate(txtPeriodoFinalConsultaTitulos.Text), "yyyy/MM/dd")

                    If Not IsDate(DataInicial) Or Not IsDate(DataFinal) Then
                        MsgBox(Me.Page, "Informe um Periodo Inicial / Final Valido")
                        Exit Sub
                    End If
                End If

                Dim DS As New DataSet
                Dim Campo As String()
                Dim Parametros As String = ""

                Dim UnidadeDeNegocio As String = ""
                Dim NomeUnidadeDeNegocio As String = ""
                Dim Empresa As String = ""
                Dim EndEmpresa As String = ""
                Dim NomeEmpresa As String = ""
                Dim CidadeEmpresa As String = ""
                Dim EstadoEmpresa As String = ""
                Dim GrupoEmpresa As String = ""

                Dim Cliente As String = ""
                Dim EndCliente As String = ""


                If DdlUnidadeConsultaTitulos.Text <> "" Then
                    UnidadeDeNegocio = DdlUnidadeConsultaTitulos.SelectedValue
                    NomeUnidadeDeNegocio = DdlUnidadeConsultaTitulos.SelectedItem.Text
                Else
                    UnidadeDeNegocio = ""
                    NomeUnidadeDeNegocio = ""
                End If
                If DdlEmpresaConsultaTitulos.Text <> "" Then
                    Campo = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")
                    Dim EmpresaCs As New Cliente(Campo(0), Campo(1))
                    Empresa = EmpresaCs.Codigo
                    EndEmpresa = EmpresaCs.CodigoEndereco
                    NomeEmpresa = EmpresaCs.Nome
                    CidadeEmpresa = EmpresaCs.Cidade
                    EstadoEmpresa = EmpresaCs.CodigoEstado
                    Parametros &= "Empresa: " & DdlEmpresaConsultaTitulos.SelectedItem.Text & vbCrLf
                    If chkConsolidarEmpresa.Checked Then Parametros &= "Empresa Consolidado" & vbCrLf
                End If
                If hdfCodigoCliente.Value.Length > 0 Then
                    Campo = hdfCodigoCliente.Value.ToString.Split(";")
                    Cliente = Campo(0)
                    EndCliente = Campo(1)
                    Parametros &= "Cliente: " & txtNomeCliente.Text & vbCrLf
                    If chkConsolidarCliente.Checked Then Parametros &= "Cliente Consolidado" & vbCrLf
                Else
                    Cliente = ""
                    EndCliente = 0
                End If

                Sql = "SELECT Empresas.Reduzido AS ReduzidoEmpresa, " & vbCrLf & _
                      "       T.Empresa, " & vbCrLf & _
                      "       T.EndEmpresa," & vbCrLf & _
                      "       Empresas.Nome AS NomeEmpresa, " & vbCrLf & _
                      "       Empresas.Cidade AS CidadeEmpresa, " & vbCrLf & _
                      "       Empresas.Estado AS EstadoEmpresa, " & vbCrLf & _
                      "       convert(nvarchar,T.Titulo_Id) + ' ' + M.Cifrao AS Registro," & vbCrLf & _
                      "       T.Pedido," & vbCrLf & _
                      "       T.CliFor as Cliente, Clientes.Nome AS NomeCliente," & vbCrLf & _
                      "       T.Movimento, T.Reprogramacao AS Vencimento, T.DataBaixa as Baixa," & vbCrLf & _
                      "       T.Provisao, " & vbCrLf & _
                      "       T.CarteiraDoTitulo, " & vbCrLf & _
                      "       Carteira.Descricao AS NomeCarteira," & vbCrLf & _
                      "       T.Historico," & vbCrLf & _
                      "       case" & vbCrLf & _
                      "           when T.Provisao = 1" & vbCrLf & _
                      "			 then isnull(TC_LIQ.ValorOficial ,0)" & vbCrLf & _
                      "           else" & vbCrLf & _
                      "             case when M.Classificacao <> 'O'" & vbCrLf & _
                      "                then convert(numeric(18,2), isnull(TC_PROD.ValorMoeda,0) * Cotacoes.indice)" & vbCrLf & _
                      "             else isnull(TC_PROD.ValorOficial,0)" & vbCrLf & _
                      "           end" & vbCrLf & _
                      "        end ValorLiquido," & vbCrLf & _
                      "        case" & vbCrLf & _
                      "           when T.Provisao = 1" & vbCrLf & _
                      "			 then isnull(TC_LIQ.ValorMoeda,0)" & vbCrLf & _
                      "           else" & vbCrLf & _
                      "             case when M.Classificacao = 'O'" & vbCrLf & _
                      "                then convert(numeric(18,2),isnull(TC_PROD.ValorOficial,0) / Cotacoes.indice)" & vbCrLf & _
                      "             else isnull(TC_PROD.ValorMoeda,0)" & vbCrLf & _
                      "           end " & vbCrLf & _
                      "        end MoedaValorLiquido" & vbCrLf

                If chkListarProdutos.Checked Then
                    Sql &= " into #temp" & vbCrLf
                End If

                Sql &= "  FROM Titulos T" & vbCrLf & _
                       "  Left Join TitulosxContaContabil TC_LIQ" & vbCrLf & _
                       "    on TC_LIQ.Titulo_Id = T.Titulo_Id" & vbCrLf & _
                       "   and TC_LIQ.Conta_Id  = T.ContaContabilRecPag" & vbCrLf & _
                       "   and TC_LIQ.DC_Id     = case " & vbCrLf & _
                       "                            when T.RecPag in ('P','C') " & vbCrLf & _
                       "                              then 'C'" & vbCrLf & _
                       "                              else 'D'" & vbCrLf & _
                       "                          end" & vbCrLf & _
                       "  Left Join TitulosxContaContabil TC_PROD" & vbCrLf & _
                       "    on TC_PROD.Titulo_Id = T.Titulo_Id" & vbCrLf & _
                       "   and TC_PROD.Conta_Id  = T.ContaContabilCliFor" & vbCrLf & _
                       "   and TC_PROD.DC_Id     = case " & vbCrLf & _
                       "							 when T.RecPag in ('P','C')" & vbCrLf & _
                       "							   then 'D'" & vbCrLf & _
                       "							   else 'C'" & vbCrLf & _
                       "						   end" & vbCrLf & _
                       "  Inner Join cotacoes" & vbCrLf & _
                       "     on Cotacoes.Data_id      = T.Reprogramacao " & vbCrLf & _
                       "    and Cotacoes.Indexador_Id = T.Indexador " & vbCrLf & _
                       "  INNER JOIN Clientes AS Empresas " & vbCrLf & _
                       "     ON T.Empresa    = Empresas.Cliente_Id " & vbCrLf & _
                       "    AND T.EndEmpresa = Empresas.Endereco_Id " & vbCrLf & _
                       "  INNER JOIN Clientes " & vbCrLf & _
                       "     ON T.Clifor         = Clientes.Cliente_Id " & vbCrLf & _
                       "    AND T.EnderecoCliFor = Clientes.Endereco_Id " & vbCrLf & _
                       "  INNER JOIN Moedas M" & vbCrLf & _
                       "     ON M.Moeda_Id = T.Moeda  " & vbCrLf & _
                       "   LEFT JOIN Carteira " & vbCrLf & _
                       "     ON T.CarteiraDoTitulo = Carteira.Carteira_Id " & vbCrLf

                If ddlSafra.SelectedIndex > 0 Or ddlRepresentante.SelectedIndex > 0 Then
                    Sql &= "   Inner Join Pedidos AS P " & vbCrLf & _
                           "     ON P.Empresa_id    = T.Empresa" & vbCrLf & _
                           "    And P.EndEmpresa_Id = T.EndEmpresa " & vbCrLf & _
                           "    And P.Pedido_id     = T.Pedido " & vbCrLf
                End If

                    Sql &= "  Where T.Situacao = 1" & vbCrLf & _
                           "    AND (T.ContaContabilCliFor NOT LIKE '1010101%') " & vbCrLf & _
                           "    AND (T.ContaContabilCliFor NOT LIKE '1010102%') " & vbCrLf

                    If RdReceber.Checked Then
                        Sql &= "   And T.RecPag = 'R'" & vbCrLf
                        Parametros &= "Contas a Receber" & vbCrLf
                    ElseIf RdPagar.Checked Then
                        Sql &= "   And T.RecPag = 'P'" & vbCrLf
                        Parametros &= "Contas a Pagar" & vbCrLf
                    ElseIf rdContabil.Checked Then
                        Sql &= "   And T.RecPag = 'C'" & vbCrLf
                        Parametros &= "Lancamentos Contabeis" & vbCrLf
                    End If

                    If rdMestre.Checked Then
                        Sql &= "    AND (isnull(T.RegistroMestre,0) = 0 or T.titulo_id = T.RegistroMestre) " & vbCrLf
                    Else
                        Sql &= "    AND (isnull(T.RegistroMestre,0) = 0 or T.titulo_id <> T.RegistroMestre) " & vbCrLf
                    End If

                    If ddlSafra.SelectedIndex > 0 Then
                        Sql &= "  Where P.Safra         = '" & ddlSafra.SelectedValue & "' "
                        Parametros &= "safra: " & ddlSafra.SelectedValue & vbCrLf
                    End If

                    If ddlRepresentante.SelectedIndex > 0 Then
                        Dim rep As String() = ddlRepresentante.SelectedValue.Split("-")
                        Sql &= "  And exists (Select 1" & vbCrLf & _
                               "                from Comissoes C" & vbCrLf & _
                               "               where C.Empresa_id          = P.Empresa_id" & vbCrLf & _
                               "                 And C.EndEmpresa_Id       = P.EndEmpresa_Id " & vbCrLf & _
                               "                 And C.Pedido_id           = P.Pedido_id " & vbCrLf & _
                               "                 And P.Sequencia_id        = 0" & vbCrLf & _
                               "                 And C.Representante_id    ='" & rep(0) & "'" & vbCrLf & _
                               "                 And C.EndRepresentante_Id = " & rep(1) & ")"

                        Parametros &= "Representante: " & ddlRepresentante.SelectedItem.Text & vbCrLf
                    End If

                    Dim Provisao As String = ""
                    Provisao = IIf(chkBaixados.Checked, "1", "")
                    Provisao &= IIf(chkProvisao.Checked, IIf(Provisao.Length > 0, ",2", "2"), "")
                    Provisao &= IIf(chkPrevisao.Checked, IIf(Provisao.Length > 0, ",3", "3"), "")
                    Sql &= "   and T.Provisao in (" & Provisao & ")"
                    Parametros &= "Titulos : " & IIf(chkBaixados.Checked, "Baixados ", "") & IIf(chkProvisao.Checked, "Provisionado ", "") & IIf(chkPrevisao.Checked, "Previsionado ", "") & vbCrLf


                    If ddlMoeda.SelectedIndex > 0 Then
                        Sql &= " and T.Moeda = " & ddlMoeda.SelectedValue
                        Parametros &= "Moeda: " & ddlMoeda.SelectedItem.Text & vbCrLf
                    End If

                    If UnidadeDeNegocio <> "" Then
                        Sql &= " and T.UnidadeDeNegocio = '" & UnidadeDeNegocio & "' " & vbCrLf
                    End If

                    If Empresa <> "" Then
                        If chkConsolidarEmpresa.Checked Then
                            Sql &= " and left(T.Empresa,8) = '" & Empresa.Substring(0, 8) & "'" & vbCrLf
                        Else
                            Sql &= " and T.Empresa    ='" & Empresa & "'" & vbCrLf & _
                                   " and T.EndEmpresa = " & EndEmpresa & vbCrLf
                        End If
                    End If

                    If Cliente <> "" Then
                        If chkConsolidarCliente.Checked Then
                            Sql &= " and left(T.Cliente,8) = '" & Cliente.Substring(0, 8) & "'" & vbCrLf
                        Else
                            Sql &= " and T.Cliente    ='" & Cliente & "'" & vbCrLf & _
                                   " and T.EndCliente = " & EndCliente & vbCrLf
                        End If
                    End If

                    If ddlCarteiraDoTitulo.SelectedIndex > 0 Then
                        Sql &= " and T.CarteiraDoTitulo = '" & ddlCarteiraDoTitulo.SelectedValue & "'" & vbCrLf
                        Parametros &= "Carteira: " & ddlCarteiraDoTitulo.SelectedItem.Text & vbCrLf
                    End If



                    If chkUsarPeriodo.Checked Then
                        Sql &= " AND T.Reprogramacao BETWEEN '" & DataInicial & "' and '" & DataFinal & "'" & vbCrLf
                    Parametros &= "Vecimentos entre " & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & " a " & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & vbCrLf
                    End If

                    If chkListarProdutos.Checked Then
                        Sql &= " Select * from #Temp;" & vbCrLf
                    End If

                    If Not radData.Checked Then
                    Sql &= " Select T.Empresa," & vbCrLf & _
                       "       T.EndEmpresa," & vbCrLf & _
                       "       convert(nvarchar,T.Titulo_Id)  + ' ' + M.Cifrao AS Registro," & vbCrLf & _
                       "       PxI.CodigoProduto," & vbCrLf & _
                       "       PxI.NomeProduto," & vbCrLf & _
                       "       PxI.Quantidade," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when  PxI.Quantidade = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else PxI.Valor / PxI.Quantidade" & vbCrLf & _
                       "       end as Unitario," & vbCrLf & _
                       "       PxI.Valor," & vbCrLf & _
                       "       PxI.UnidadeDeMedida" & vbCrLf & _
                       " From Titulos T" & vbCrLf & _
                       " Inner Join Moedas M" & vbCrLf & _
                       "    on T.Moeda = M.Moeda_Id" & vbCrLf & _
                       " Inner Join (" & vbCrLf & _
                       "				Select Pxi.Empresa_id," & vbCrLf & _
                       "					   Pxi.EndEmpresa_id," & vbCrLf & _
                       "					   Pxi.Pedido_Id," & vbCrLf & _
                       "					   Pxi.Produto_Id as CodigoProduto," & vbCrLf & _
                       "					   Prd.Nome as NomeProduto," & vbCrLf & _
                       "					   Sum(case" & vbCrLf & _
                       "						     When Pxi.TipoDeLancamento = 'E'" & vbCrLf & _
                       "						  	   then Pxi.Quantidade * - 1" & vbCrLf & _
                       "							   else Pxi.Quantidade" & vbCrLf & _
                       "					       end) as Quantidade," & vbCrLf & _
                       "					   Sum(case" & vbCrLf & _
                       "						     When P.Moeda = 1 " & vbCrLf & _
                       "						  	   then case " & vbCrLf & _
                       "                                   When Pxi.TipoDeLancamento = 'E'" & vbCrLf & _
                       "                                     then Pxi.TotalOficial * -1 " & vbCrLf & _
                       "                                     else Pxi.TotalOficial " & vbCrLf & _
                       "                                 end" & vbCrLf & _
                       "							   else case " & vbCrLf & _
                       "                                   When Pxi.TipoDeLancamento = 'E'" & vbCrLf & _
                       "                                     then Pxi.TotalMoeda * -1 " & vbCrLf & _
                       "                                     else Pxi.TotalMoeda " & vbCrLf & _
                       "                                 end" & vbCrLf & _
                       "					       end) as Valor," & vbCrLf & _
                       "					   Prd.Unidade as UnidadeDeMedida" & vbCrLf & _
                       "				 from Pedidos P" & vbCrLf & _
                       "                INNER JOIN PedidoXItemxLancamento Pxi" & vbCrLf & _
                       "                   ON P.Empresa_Id    = Pxi.Empresa_Id " & vbCrLf & _
                       "                  AND P.EndEmpresa_Id = Pxi.EndEmpresa_Id " & vbCrLf & _
                       "                  AND P.Pedido_Id     = Pxi.Pedido_Id " & vbCrLf & _
                       "				Inner join Produtos Prd" & vbCrLf & _
                       "				   on Pxi.Produto_Id = Prd.Produto_Id" & vbCrLf & _
                       "             group by Pxi.Empresa_id, Pxi.EndEmpresa_id, Pxi.Pedido_Id, Pxi.Produto_Id, Prd.Nome, Prd.Unidade" & vbCrLf & _
                       "             ) PxI" & vbCrLf & _
                       "   on PxI.Empresa_id    = T.Empresa" & vbCrLf & _
                       "  and PxI.EndEmpresa_id = T.EndEmpresa" & vbCrLf & _
                       "  and PxI.Pedido_id     = T.Pedido" & vbCrLf & _
                       " where T.Titulo_Id in (" & IIf(chkListarProdutos.Checked, "Select Titulo_id from #Temp", "-1") & ");" & vbCrLf

                    End If

                    DS = Banco.ConsultaDataSet(Sql, "Titulos")

                    If Not radData.Checked Then DS.Tables("Titulos1").TableName = "PedidoProduto"

                    Dim rpt As String = ""

                    If radData.Checked Then
                        If DdlEmpresaConsultaTitulos.SelectedIndex = 0 Then
                            rpt = "CrTitulos"
                        Else
                            rpt = "CrTitulosPorEmpresa"
                        End If
                    End If

                    If radClienteData.Checked Then rpt = "Cr_TitulosPorCliente"
                    If RdClientePedido.Checked Then rpt = "Cr_TitulosPorPedido"

                    Dim param As New Dictionary(Of String, Object)
                    param.Add("Relatorio", "Relação de Titulos A " & IIf(RdPagar.Checked, "Pagar", "Receber"))
                    param.Add("Nome", NomeEmpresa)
                    param.Add("Cidade", CidadeEmpresa)
                    param.Add("Estado", EstadoEmpresa)
                    param.Add("UnidadeDeNegocio", Parametros)
                    param.Add("VisualizarProdutos", chkListarProdutos.Checked)
                    param.Add("DataInicial", "01/01/2013")
                    param.Add("DataFinal", "01/01/2013")
                    param.Add("TipoDaCarteira", "Carteira")

                    Funcoes.BindReport(Me.Page, DS, rpt, IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), param)
                Else
                    MsgBox(Me.Page, "Usuário sem permissao para emitir relatório.")
                End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
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