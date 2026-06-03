Imports System.Data
Imports System.IO
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class PosicaoDePedidos
    Inherits BasePage

    Dim Codigo As String
    Dim Descricao As String

#Region "Load"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PosicaoDePedidos", "ACESSAR") Then
                    HID.Value = Guid.NewGuid.ToString
                    txtData.Text = Now.Date.ToString("dd/MM/yyyy")

                    ddl.Carregar(cmbUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "")
                    Funcoes.VerificaUnidadeDeNegocio(cmbUnidadeNegocio, cmbEmpresa)

                    ddl.Carregar(cmbClasse, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1")
                    ddl.Carregar(ddlClasseSubOperacao, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(suboperacao,0) = 1")

                    ddl.Carregar(cmbSafra, CarregarDDL.Tabela.Safra, "")
                    'ddl.Carregar(DdlOperacoes, CarregarDDL.Tabela.Operacao, "")
                    ucSelecaoProduto.WhereProduto = "Agrupar = 'N'"
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuario sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Eventos"

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

    Protected Sub lnkExcelDados_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcelDados.Click
        Try
            EmitirRelatorio(False, True)
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

    Protected Sub cmbUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbUnidadeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, cmbUnidadeNegocio.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaClientes.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objClientexPosicao" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaRepresentante_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaRepresentante.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("6")
            Popup.ConsultaDeClientes(Me, "objRepresentantexPosicao" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClientexPosicao" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objClientexPosicao" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtCliente.Text = pCliente.Nome
            txtCodigoCliente.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objClientexPosicao" & HID.Value.ToString)
        ElseIf Not Session("objRepresentantexPosicao" & HID.Value) Is Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objRepresentantexPosicao" & HID.Value.ToString), [Lib].Negocio.Cliente)
            txtRepresentante.Text = pCliente.Nome
            txtCodigoRepresentante.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            Session.Remove("objRepresentantexPosicao" & HID.Value.ToString)
        End If
    End Sub

    'Protected Sub DdlOperacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        ddl.Carregar(DdlSubOperacoes, CarregarDDL.Tabela.OperacaoSubOperacao, " So.Operacao_id = " & DdlOperacoes.SelectedValue)
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    Protected Sub ckDataDeAbertura_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ckDataDeAbertura.CheckedChanged
        Try
            pnlDataAberturaPedido.Visible = ckDataDeAbertura.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkDataEntregaPedido_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkDataEntregaPedido.CheckedChanged
        Try
            pnlDataEntregaPedido.Visible = chkDataEntregaPedido.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PosicaoDePedidos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Validação"

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtData.Text) Or Not IsDate(txtData.Text) Then
            MsgBox(Me.Page, "Informe uma data válida.", eTitulo.Info)
            txtData.Focus()
            Return False
        ElseIf cmbEmpresa.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Selecione a empresa.")
            Return False
        ElseIf Not ucSelecaoProduto.TemSelecionado Then
            MsgBox(Me.Page, "Selecione ao menos 1 produto.")
            Return False
        End If
        Return True
    End Function

#End Region

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean, Optional ByVal ExcelDados As Boolean = False)
        Dim rpt As New ReportDocument()
        Try
            If ValidarCampos() Then
                Dim dsPosicoes As New DataSet
                Dim Texto As String = ""

                'Dim StrProcuracoes As String = SqlProc()
                'dsPosicoes = Banco.ConsultaDataSet(StrProcuracoes, "Procuracoes")

                dsPosicoes = (Banco.ConsultaDataSet(SqlPosicao(), "PP"))
                dsPosicoes.Tables(0).TableName = "PosicaoDePedidos"
                dsPosicoes.Tables(1).TableName = "Procuracoes"

                Dim strCaminho As String = String.Empty

                If ExcelDados Then
                    strCaminho &= HttpContext.Current.Server.MapPath("~/Reports/Cr_PosicaoDePedidosexcell2.rpt")
                Else
                    If Left(cmbEmpresa.SelectedValue.Split("-")(0), 8) = "62780383" Then
                        strCaminho &= HttpContext.Current.Server.MapPath("~/Reports/Cr_PosicaoDePedidosHorus.rpt")
                    Else
                        strCaminho &= HttpContext.Current.Server.MapPath("~/Reports/Cr_PosicaoDePedidos.rpt")
                    End If
                End If

                rpt.Load(strCaminho)

                Dim strNomeArquivo As String
                If Pdf Then
                    strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Else
                    strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".XLS"
                End If

                Dim strArquivo As String = HttpContext.Current.Server.MapPath(strNomeArquivo)
                rpt.SetDataSource(dsPosicoes)

                If Not ExcelDados Then
                    Dim crParameterValues As CrystalDecisions.Shared.ParameterValues
                    Dim crParameterDiscreteValue As CrystalDecisions.Shared.ParameterDiscreteValue
                    Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                    Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

                    crParameterFieldDefinitions = rpt.DataDefinition.ParameterFields()

                    Dim Parametros As String = String.Empty
                    If cmbUnidadeNegocio.SelectedIndex > 0 Then Parametros &= "Unidade De Negocio: " & cmbUnidadeNegocio.SelectedItem.Text & vbCrLf
                    If cmbEmpresa.SelectedIndex > 0 Then Parametros &= "Empresa: " & cmbEmpresa.SelectedItem.Text & vbCrLf
                    If txtCliente.Text.Length > 0 Then Parametros &= "Cliente: " & txtCliente.Text & vbCrLf
                    If chkConsolidarCliente.Checked Then Parametros &= "Clientes Consolidados" & vbCrLf
                    If cmbClasse.SelectedIndex > 0 Then Parametros &= "Classe: " & cmbClasse.SelectedItem.Text & vbCrLf
                    If rdTrocaSim.Checked Then Parametros &= "Pedidos De Troca" & vbCrLf
                    If rdTrocaNao.Checked Then Parametros &= "SEM Pedidos De Troca" & vbCrLf

                    If rdAntecipadaSim.Checked Then Parametros &= "Pedidos Com Operação Antecipada" & vbCrLf
                    If rdAntecipadaNao.Checked Then Parametros &= "SEM Pedidos Com Operação Antecipada" & vbCrLf

                    Select Case ddlFretes.SelectedIndex
                        Case 1 : Parametros &= "Fretes CIF" & vbCrLf
                        Case 2 : Parametros &= "Fretes FOB" & vbCrLf
                    End Select

                    If IsDate(txtDataInicial.Text) And IsDate(txtDataFinal.Text) Then
                        Parametros &= " Data de Entrega de '" & txtDataInicial.Text & " a " & txtDataFinal.Text & vbCrLf
                    End If

                    'If DdlOperacoes.SelectedIndex > 0 Or DdlSubOperacoes.SelectedIndex > 0 Then Parametros &= vbCrLf

                    Dim par As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("s", "s", "", True)
                    Parametros &= par(1)

                    Dim OP As ArrayList = ucSelecaoOperacoes.GetSqlEParametrosRelatorioOperacoes("s", "s", "")
                    Parametros &= OP(1)

                    If cmbSafra.SelectedIndex > 0 Then Parametros &= "Safra: " & cmbSafra.SelectedItem.Text


                    crParameterFieldDefinition = crParameterFieldDefinitions.Item("Parametros")
                    crParameterValues = crParameterFieldDefinition.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = Parametros
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                    crParameterFieldDefinition = crParameterFieldDefinitions.Item("Posicao")
                    crParameterValues = crParameterFieldDefinition.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue

                    If rdComSaldo.Checked Then Texto = "Posição de Pedidos Com Saldo"
                    If rdLiquidados.Checked Then Texto = "Posição de Pedidos Liquidados"
                    If rdTodos.Checked Then Texto = "Posição de Pedidos Geral"

                    If RadCessionarioSim.Checked Then
                        Texto &= " - Cessionário SIM"
                    Else
                        Texto &= " - Cessionário NÃO"
                    End If

                    crParameterDiscreteValue.Value = Texto
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                    crParameterFieldDefinition = crParameterFieldDefinitions.Item("NomeEmpresa")
                    crParameterValues = crParameterFieldDefinition.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = HttpContext.Current.Session("ssNomeEmpresa")
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                    crParameterFieldDefinition = crParameterFieldDefinitions.Item("Periodo")
                    crParameterValues = crParameterFieldDefinition.CurrentValues
                    crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                    crParameterDiscreteValue.Value = "Posicao do dia " & txtData.Text
                    crParameterValues.Add(crParameterDiscreteValue)
                    crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
                End If

                If Dir(strArquivo).Length > 0 Then Kill(strArquivo)

                If Pdf Then
                    rpt.ExportToDisk(ExportFormatType.PortableDocFormat, strArquivo)
                ElseIf ExcelDados Then
                    rpt.ExportToDisk(ExportFormatType.ExcelRecord, strArquivo)
                Else
                    rpt.ExportToDisk(ExportFormatType.Excel, strArquivo)
                End If

                If IO.File.Exists(strArquivo) Then
                    If Pdf Then
                        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & strNomeArquivo & "');", True)
                    Else
                        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.location = '" & strNomeArquivo & "';", True)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            rpt.Close()
            rpt.Dispose()
        End Try
    End Sub

    Private Sub Limpar()
        'DdlOperacoes.SelectedIndex = 0
        'DdlSubOperacoes.Items.Clear()

        txtCliente.Text = ""
        txtRepresentante.Text = ""
        txtCodigoCliente.Value = ""
        txtCodigoRepresentante.Value = ""
        ucSelecaoProduto.CarregarNivel(5)
        cmbSafra.SelectedIndex = 0

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbUnidadeNegocio.Enabled = False
            cmbEmpresa.Enabled = False
        End If
    End Sub

#Region "Sqls"

    Public Function SqlPosicao() As String
        'O sql foi centralizado na classe NGS.Lib.Negocio\Movimentacao\PosicaoDePedido.vb
        Dim Cliente As Cliente = Nothing
        Dim Represenante As Cliente = Nothing
        Dim par As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id")
        Dim OP As ArrayList = ucSelecaoOperacoes.GetSqlEParametrosRelatorioOperacoes("P.Operacao", "P.SubOperacao")
        Dim DataEntregaInicial As String = ""
        Dim DataEntregaFinal As String = ""
        Dim DataLancamentoPedidoInicial As String = ""
        Dim DataLancamentoPedidoFinal As String = ""
        Dim SituacaoPedido As String = ""
        Dim CodigoOperacao As Integer = 0
        Dim CodigoSuboperacao As Integer = 0
        Dim Saldo As Integer = 0
        Dim Troca As Integer = 0
        Dim Antecipada As Integer = 0
        Dim Recompra As Integer = 0
        Dim Cessionario As Integer = 0
        Dim lstEmpresa As New List(Of String)

        If cmbEmpresa.SelectedValue.Length > 0 Then lstEmpresa.Add(cmbEmpresa.SelectedValue.Replace("-", ""))

        If txtCodigoCliente.Value.Length > 0 Then Cliente = New Cliente(txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1))

        If txtCodigoRepresentante.Value.Length > 0 Then Represenante = New Cliente(txtCodigoRepresentante.Value.Split("-")(0), txtCodigoRepresentante.Value.Split("-")(1))

        If chkDataEntregaPedido.Checked AndAlso IsDate(txtDataInicial.Text) AndAlso IsDate(txtDataFinal.Text) Then
            DataEntregaInicial = CDate(txtDataInicial.Text).ToString("yyyy-MM-dd")
            DataEntregaFinal = CDate(txtDataFinal.Text).ToString("yyyy-MM-dd")
        End If

        If ckDataDeAbertura.Checked AndAlso IsDate(txtDataInicialAbertura.Text) AndAlso IsDate(txtDataFinalAbertura.Text) Then
            DataLancamentoPedidoInicial = CDate(txtDataInicialAbertura.Text).ToString("yyyy-MM-dd")
            DataLancamentoPedidoFinal = CDate(txtDataFinalAbertura.Text).ToString("yyyy-MM-dd")
        End If

        If rdAberto.Checked Then
            SituacaoPedido = "A"
        ElseIf rdFechado.Checked Then
            SituacaoPedido = "F"
        End If

        If rdComSaldo.Checked Then
            Saldo = 1
        ElseIf rdLiquidados.Checked Then
            Saldo = 2
        End If

        If rdTrocaSim.Checked Then
            Troca = 1
        ElseIf rdTrocaNao.Checked Then
            Troca = 2
        End If

        If rdAntecipadaSim.Checked Then
            Antecipada = 1
        ElseIf rdAntecipadaNao.Checked Then
            Antecipada = 2
        End If

        If rdReSim.Checked Then
            Recompra = 1
        ElseIf rdReNao.Checked Then
            Recompra = 2
        End If

        'If DdlOperacoes.SelectedIndex > 0 Then CodigoOperacao = DdlOperacoes.SelectedValue
        'If DdlSubOperacoes.SelectedIndex > 0 Then CodigoSuboperacao = DdlSubOperacoes.SelectedValue.Split("-")(1)

        Dim pos As New ListPosicaoDePedido
        Dim t = pos.SqlPosicao(CDate(txtData.Text), lstEmpresa, chkConsolidarEmpresa.Checked, Cliente, chkConsolidarCliente.Checked, "", "", cmbSafra.SelectedValue, cmbSafra.SelectedValue, par(0), DataEntregaInicial, DataEntregaFinal, DataLancamentoPedidoInicial, DataLancamentoPedidoFinal, SituacaoPedido, ddlFretes.SelectedIndex, cmbClasse.SelectedValue, ddlClasseSubOperacao.SelectedValue, 0, 0, Saldo, 0, True, Troca, Antecipada, Recompra, RadCessionarioSim.Checked, Represenante, OP(0))
        Return t
    End Function

    Public Function SqlProc() As String
        Dim Data As String = "'" & CDate(txtData.Text).ToString("yyyy-MM-dd") & "'"

        Dim par As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id")

        Dim sql As String
        sql = "	Select P.Empresa_Id as Empresa, P.EndEmpresa_Id as EndEmpresa, P.Pedido_Id," & vbCrLf & _
              "		   case" & vbCrLf & _
              " 		 when P.Pedido_Id = PR.PedidoCedente  " & vbCrLf & _
              " 		   then CES.Reduzido" & vbCrLf & _
              " 		   else CED.Reduzido" & vbCrLf & _
              " 	   end as Reduzido,  " & vbCrLf & _
              "		   case" & vbCrLf & _
              " 		 when P.Pedido_Id = PR.PedidoCedente  " & vbCrLf & _
              " 		   then CES.Cliente_Id" & vbCrLf & _
              " 		   else CED.Cliente_Id" & vbCrLf & _
              " 		 end as Cliente,  " & vbCrLf & _
              "		   case" & vbCrLf & _
              " 		 when P.Pedido_Id = PR.PedidoCedente  " & vbCrLf & _
              " 		   then CES.Endereco_Id" & vbCrLf & _
              " 		   else CED.Endereco_Id" & vbCrLf & _
              " 	   end as Endereco,  " & vbCrLf & _
              "		   case  " & vbCrLf & _
              " 		 when P.Pedido_Id = PR.PedidoCedente  " & vbCrLf & _
              " 		   then CES.Nome + ' - Proc.' + convert(varchar,PR.Procuracao_ID)" & vbCrLf & _
              " 		   else CED.Nome + ' - Proc.' + convert(varchar,PR.Procuracao_ID)" & vbCrLf & _
              " 	   end as Nome," & vbCrLf & _
              " 	     0 as Cedente," & vbCrLf & _
              " 	     0 as Cessionario," & vbCrLf & _
              "          0 as nota, 0 as Fixado" & vbCrLf & _
              "     from Pedidos P" & vbCrLf & _
              "	   inner join Procuracoes PR" & vbCrLf & _
              "	  	  on P.Empresa_id    = PR.Empresa_id" & vbCrLf & _
              "      and P.EndEmpresa_id = PR.EndEmpresa_Id " & vbCrLf & _
              "      and P.Pedido_Id     = PR.PedidoCedente " & vbCrLf & _
              "	   inner JOIN Clientes CES" & vbCrLf & _
              "   	  ON PR.Cessionario    = CES.Cliente_Id" & vbCrLf & _
              "	     AND PR.EndCessionario = CES.Endereco_Id" & vbCrLf & _
              "	   inner JOIN Clientes CED" & vbCrLf & _
              "  	  ON PR.Cedente    = CED.Cliente_Id" & vbCrLf & _
              "	     AND PR.EndCedente = CED.Endereco_Id" & vbCrLf & _
              "     LEFT JOIN (Select sb.Empresa_id," & vbCrLf & _
              "                       sb.EndEmpresa_id," & vbCrLf & _
              "                       sb.Procuracao," & vbCrLf & _
              "                       sum(sb.Quantidade) as Quantidade" & vbCrLf & _
              "                  from (" & vbCrLf & _
              "                        SELECT NF.Empresa_Id," & vbCrLf & _
              "                               NF.EndEmpresa_Id," & vbCrLf & _
              "                               NF.Procuracao," & vbCrLf & _
              "                               SUM(nfi.QuantidadeFisica) AS Quantidade" & vbCrLf & _
              "                          FROM NotasFiscais NF" & vbCrLf & _
              "                         INNER JOIN NotasFiscaisXItens nfi" & vbCrLf & _
              "                            ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
              "                           AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
              "                           AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
              "                           AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
              "                           AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
              "                           AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
              "                           AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
              "                         Inner join SubOperacoes SO" & vbCrLf & _
              "                            on so.Operacao_Id     = NF.Operacao" & vbCrLf & _
              "                           and so.SubOperacoes_Id = NF.SubOperacao" & vbCrLf & _
              "                         Where so.Classe <> '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf & _
              "                           and NF.Situacao in (1,4,7)" & vbCrLf & _
              "                         GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Procuracao" & vbCrLf & _
              "                        Union All" & vbCrLf & _
              "                       Select P.Empresa_Id," & vbCrLf & _
              "                              P.EndEmpresa_Id," & vbCrLf & _
              "                              PIF.Procuracao," & vbCrLf & _
              "                              sum(pif.Quantidade)" & vbCrLf & _
              "                         from Pedidos P" & vbCrLf & _
              "                        Inner join PedidosXItensXFixacoes PIF" & vbCrLf & _
              "                           on P.Empresa_Id    = PIF.Empresa_Id" & vbCrLf & _
              "                          and P.EndEmpresa_Id = PIF.EndEmpresa_Id" & vbCrLf & _
              "                          and p.Pedido_Id     = PIF.Pedido_Id" & vbCrLf & _
              "                        Where P.Situacao = 1" & vbCrLf & _
              "                        Group by P.Empresa_Id, P.EndEmpresa_Id, PIF.Procuracao" & vbCrLf & _
              "                      ) as sb" & vbCrLf & _
              "                group by sb.Empresa_id, sb.EndEmpresa_id, sb.Procuracao" & vbCrLf & _
              "              ) AS sb_Real" & vbCrLf & _
              "       on pr.Empresa_Id    = sb_Real.Empresa_Id" & vbCrLf & _
              "      and pr.EndEmpresa_Id = sb_Real.EndEmpresa_Id" & vbCrLf & _
              "      and pr.Procuracao_ID = sb_Real.Procuracao" & vbCrLf & _
              "    WHERE PR.Situacao    = 17 " & vbCrLf & _
              "      And P.DataPedido  <=" & Data & vbCrLf & _
              "      And PR.Movimento  <=" & Data & vbCrLf

        If Not String.IsNullOrWhiteSpace(cmbEmpresa.SelectedValue) Then
            If chkConsolidarEmpresa.Checked Then
                sql &= "And Left(P.Empresa_Id ,8) = " & Left(cmbEmpresa.SelectedValue.Split("-")(0), 8) & ""
            Else
                sql &= " AND P.Empresa_Id    ='" & cmbEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                       " AND P.EndEmpresa_Id = " & cmbEmpresa.SelectedValue.Split("-")(1) & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
            If chkConsolidarCliente.Checked Then
                sql &= " AND left(P.Cliente,8)    ='" & strCliente(0).Substring(0, 8) & "'" & vbCrLf
            Else
                sql &= " AND P.Cliente    ='" & strCliente(0) & "'" & vbCrLf & _
                       " AND P.EndCliente = " & strCliente(1) & vbCrLf
            End If
        End If


        Dim OP As ArrayList = ucSelecaoOperacoes.GetSqlEParametrosRelatorioOperacoes("P.Operacao", "P.SubOperacao")

        'If DdlOperacoes.SelectedValue.Length > 0 Then
        '    sql &= " AND P.Operacao = " & DdlOperacoes.SelectedValue & " " & vbCrLf
        'End If

        'If DdlSubOperacoes.SelectedValue.Length > 0 Then
        '    sql &= " AND P.SubOperacao = " & DdlSubOperacoes.SelectedValue.Split("-")(1) & vbCrLf
        'End If

        'Produtos do user Control selecao de Produtos
        'sql &= " AND " & par(0) & vbCrLf

        sql &= " Group by P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_Id, PR.Procuracao_Id, PR.PedidoCedente," & vbCrLf & _
               "       case when P.Pedido_Id = PR.PedidoCedente then CES.Reduzido    else CED.Reduzido    end," & vbCrLf & _
               "       case when P.Pedido_Id = PR.PedidoCedente then CES.Cliente_Id  else CED.Cliente_Id  end," & vbCrLf & _
               "       case when P.Pedido_Id = PR.PedidoCedente then CES.Endereco_Id else CED.Endereco_Id end," & vbCrLf & _
               "       case when P.Pedido_Id = PR.PedidoCedente then CES.Nome + ' - Proc.' + convert(varchar,PR.Procuracao_ID) else CED.Nome + ' - Proc.' + convert(varchar,PR.Procuracao_ID) end" & vbCrLf

        sql &= "having sum(case" & vbCrLf & _
               "  			 when P.Pedido_Id = isnull(PR.PedidoCedente,0)" & vbCrLf & _
               "			   then PR.Quantidade" & vbCrLf & _
               "			   else 0" & vbCrLf & _
               "		   end - case" & vbCrLf & _
               "  					when P.Pedido_Id = PR.PedidoCedente" & vbCrLf & _
               "						  then isnull(sb_Real.Quantidade,0)" & vbCrLf & _
               "					  else 0" & vbCrLf & _
               "				  end) <> 0" & vbCrLf


        Return sql
    End Function

#End Region


End Class