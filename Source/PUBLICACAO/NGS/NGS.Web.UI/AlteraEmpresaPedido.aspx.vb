Imports NGS.Lib.Negocio
Imports System.IO

Public Class AlteraEmpresaPedido
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ALTERAEMPRESAPEDIDO", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, [Lib].Negocio.CarregarDDL.Tabela.UnidadeDeNegocio)
                    ddl.Carregar(ddlUnidadeDestino, [Lib].Negocio.CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
                    ddl.Carregar(ddlOperacao, CarregarDDL.Tabela.Operacao)
                    ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            Else
                If Not Request("__EVENTARGUMENT") Is Nothing Then
                    If Request("__EVENTARGUMENT") = "DownloadArquivo" Then Download_Arquivo()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            If String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
                ddlEmpresa.Items.Clear()
            Else
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlUnidadeDestino_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDestino.SelectedIndexChanged
        Try
            If String.IsNullOrWhiteSpace(ddlUnidadeDestino.SelectedValue) Then
                ddlEmpresaDestino.Items.Clear()
            Else
                ddl.Carregar(ddlEmpresaDestino, CarregarDDL.Tabela.Empresas, ddlUnidadeDestino.SelectedValue)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnConsultaCliente_Click(sender As Object, e As EventArgs) Handles btnConsultaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AlteraEmpresaPedido")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("ALTERAEMPRESAPEDIDO", "LEITURA") Then
                If ValidaConsulta() Then
                    Dim ds As DataSet = getDataSetPedidos()
                    If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                        gridPedidos.DataSource = ds
                        gridPedidos.DataBind()
                        lnkNovo.Parent.Visible = True
                        ddlUnidadeDestino.Enabled = True
                        ddlEmpresaDestino.Enabled = True
                    Else
                        ddlUnidadeDestino.Enabled = False
                        ddlEmpresaDestino.Enabled = False
                        lnkNovo.Parent.Visible = False
                        MsgBox(Me.Page, "Nenhum resultado encontrado.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ALTERAEMPRESAPEDIDO", "GRAVAR") Then
                If validaCampos() Then
                    Dim SqlArray As New ArrayList
                    Dim sql As String = String.Empty

                    Dim empOrigem() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                    Dim empDestino() As String = ddlEmpresaDestino.SelectedValue.ToString.Split("-")

                    Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/pedidosTransferidos.txt"))

                    For i = 0 To gridPedidos.Rows.Count - 1
                        If CType(gridPedidos.Rows(i).FindControl("chkSelecionado"), CheckBox).Checked Then
                            SqlArray.Clear()

                            Dim ped As New [Lib].Negocio.Pedido(empOrigem(0), empOrigem(1), gridPedidos.Rows(i).Cells(1).Text)

                            Dim liberado As Boolean = True

                            If ped.TemPesagem Then
                                liberado = False
                                strm.WriteLine("O Pedido " & gridPedidos.Rows(i).Cells(1).Text & " não pode ser trocado porque está com Pesagem")
                            End If

                            If ped.TemFaturamento Then
                                liberado = False
                                strm.WriteLine("O Pedido " & gridPedidos.Rows(i).Cells(1).Text & " não pode ser trocado porque está com Faturamento")
                            End If

                            If ped.TemFinanceiro(1) Then
                                liberado = False
                                strm.WriteLine("O Pedido " & gridPedidos.Rows(i).Cells(1).Text & " não pode ser trocado porque está com Financeiro Baixado")
                            End If

                            If CType(gridPedidos.Rows(i).FindControl("chktemAut"), CheckBox).Checked Then
                                liberado = False
                                strm.WriteLine("O Pedido " & gridPedidos.Rows(i).Cells(1).Text & " não pode ser trocado porque está com Autorização de Retirada")
                            End If

                            If CType(gridPedidos.Rows(i).FindControl("chktemCom"), CheckBox).Checked Then
                                liberado = False
                                strm.WriteLine("O Pedido " & gridPedidos.Rows(i).Cells(1).Text & " não pode ser trocado porque está com Comissão")
                            End If

                            If CType(gridPedidos.Rows(i).FindControl("chktemTro"), CheckBox).Checked Then
                                If String.IsNullOrWhiteSpace(ped.PedidoTroca.CodigoEmpresa) Then
                                    liberado = False
                                    strm.WriteLine("O Pedido " & gridPedidos.Rows(i).Cells(1).Text & " não pode ser trocado porque está com Troca - " & ped.PedidoTroca.CodigoEmpresa & "-" & ped.PedidoTroca.Codigo)
                                End If
                            End If

                            'ESSE FOR É PARA CARREGAR OS ITENS E ENCARGOS, SE NÃO USAR NÃO CARREGA.
                            For Each itemPed In ped.Itens
                                If itemPed.Encargos.Count = 0 Then
                                    liberado = False
                                    strm.WriteLine("Pedido " & gridPedidos.Rows(i).Cells(1).Text & " - O Produto " & itemPed.Produto.Nome & ", não tem encargos cadastrados na Operação:" & ped.CodigoOperacao & "-" & ped.CodigoSubOperacao)
                                End If
                            Next

                            If liberado Then
                                ped.IUD = "C"

                                ped.SalvarSql(SqlArray)
                                'FINAL DO CANCELAMENTO

                                'INÍCIO DA TROCA PARA NOVA EMPRESA
                                ped.IUD = "I"

                                ped.CodigoUnidadeNegocio = ddlUnidadeDestino.SelectedValue
                                ped.CodigoEmpresa = empDestino(0)

                                Dim SqlN As String = "exec sp_Numerador '" & ped.CodigoEmpresa & "'," & ped.EnderecoEmpresa & "," & [Lib].Negocio.eTiposNumerador.Pedido & ""
                                Dim dsN As New DataSet
                                dsN = Banco.ConsultaDataSet(SqlN, "Numerador")
                                Dim codigoNumerador As Integer = dsN.Tables(0).Rows(0).Item(0)
                                If Not codigoNumerador > 0 Then
                                    strm.WriteLine("Numerador não cadastrado para Empresa " & ped.CodigoEmpresa & " - PEDIDO " & gridPedidos.Rows(i).Cells(1).Text & " NÃO SERÁ TRANSFERIDO")
                                    Exit For
                                End If

                                ped.Codigo = codigoNumerador

                                For Each dep In ped.Depositos
                                    If dep.Codigo = empOrigem(0) Then
                                        dep.Codigo = empDestino(0)
                                    End If
                                Next

                                For Each contratos As PedidoXContrato In ped.Contratos
                                    contratos.Codigo = empDestino(0)
                                Next

                                If String.IsNullOrWhiteSpace(gridPedidos.Rows(i).Cells(3).Text) Then
                                    ped.Observacoes = "GERADO POR TROCA DA EMPRESA " & empOrigem(0) & " PEDIDO " & gridPedidos.Rows(i).Cells(1).Text & "'"
                                Else
                                    ped.Observacoes = gridPedidos.Rows(i).Cells(3).Text & " - GERADO POR TROCA DA EMPRESA " & empOrigem(0) & " PEDIDO " & gridPedidos.Rows(i).Cells(1).Text
                                End If

                                For Each tit In ped.Vencimentos
                                    tit.UnidadeNegocio = ddlUnidadeDestino.SelectedValue
                                    tit.CodigoEmpresa = empDestino(0)
                                    tit.CodigoEmpresaPedido = empDestino(0)
                                Next

                                ped.SalvarSql(SqlArray)
                                'FINAL DA TROCA PARA NOVA EMPRESA

                                'GRAVA A EMPRESA E NUMERO DO PEDIDO NOVO REF. A TROCA NO PEDIDO DE ORIEM
                                sql = "UPDATE Pedidos SET" & vbCrLf
                                If String.IsNullOrWhiteSpace(gridPedidos.Rows(i).Cells(3).Text) Then
                                    sql &= "Observacoes = 'CANCELADO POR TROCA PARA EMPRESA " & ped.CodigoEmpresa & " PEDIDO " & ped.Codigo & "'" & vbCrLf
                                Else
                                    sql &= "Observacoes = '" & gridPedidos.Rows(i).Cells(3).Text & " - CANCELADO POR TROCA PARA EMPRESA " & ped.CodigoEmpresa & " PEDIDO " & ped.Codigo & "'" & vbCrLf
                                End If
                                sql &= " WHERE Empresa_Id    ='" & empOrigem(0) & "'" & vbCrLf & _
                                      "   AND EndEmpresa_Id = " & empOrigem(1) & vbCrLf & _
                                      "   AND Pedido_Id     = " & gridPedidos.Rows(i).Cells(1).Text & vbCrLf
                                SqlArray.Add(sql)

                                'ATUALIZA PEDIDO DA TROCA COM A NOVA NUMERAÇÃO CASO TENHA
                                If CType(gridPedidos.Rows(i).FindControl("chktemTro"), CheckBox).Checked Then
                                    sql = "UPDATE Pedidos SET" & vbCrLf & _
                                          "EmpresaTroca     = '" & ped.CodigoEmpresa & "'" & vbCrLf & _
                                          ",EndEmpresaTroca = " & ped.EnderecoEmpresa & vbCrLf & _
                                          ",PedidoTroca     = " & ped.Codigo & vbCrLf & _
                                          "WHERE Empresa_Id    ='" & ped.PedidoTroca.CodigoEmpresa & "'" & vbCrLf & _
                                          "  AND EndEmpresa_Id = " & ped.PedidoTroca.EnderecoEmpresa & vbCrLf & _
                                          "  AND Pedido_Id     = " & ped.PedidoTroca.Codigo & vbCrLf
                                    SqlArray.Add(sql)
                                End If

                                If SqlArray.Count > 0 Then
                                    If Not Banco.GravaBanco(SqlArray) Then
                                        strm.WriteLine("Erro transferindo Pedido " & empOrigem(0) & " - " & HttpContext.Current.Session("ssMessage"))
                                        Exit For
                                    End If
                                End If

                                strm.WriteLine("Pedido anterior " & empOrigem(0) & "-" & gridPedidos.Rows(i).Cells(1).Text & "  -  Pedido novo " & ped.CodigoEmpresa & "-" & ped.Codigo)
                            End If
                        End If
                    Next

                    strm.Close()

                    LimparCampos()

                    ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "downloadArquivo();", True)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub Download_Arquivo()
        Try
            Response.ContentType = "text/plain"
            Response.AppendHeader("Content-Disposition", "attachment; filename=pedidosTransferidos.txt")
            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(Server.MapPath("~/Files/pedidosTransferidos.txt"), FileMode.Open, FileAccess.Read)
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

    Protected Sub ddlOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOperacao.SelectedIndexChanged
        Try
            If String.IsNullOrWhiteSpace(ddlOperacao.SelectedValue) Then
                ddlSubOperacao.Items.Clear()
            Else
                ddl.Carregar(ddlSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & ddlOperacao.SelectedValue & " ", True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Function validaCampos() As Boolean
        If gridPedidos.Rows.Count = 0 Then
            MsgBox(Me.Page, "Selecione o pedido para alteração.")
            Return False
        End If

        Dim selected As Boolean
        For Each row As GridViewRow In gridPedidos.Rows
            If CType(row.FindControl("chkSelecionado"), CheckBox).Checked Then
                selected = True
                Exit For
            End If
        Next

        If Not selected Then
            MsgBox(Me.Page, "Nenhum pedido selecionado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEmpresaDestino.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa de destino.")
            Return False
        ElseIf ddlEmpresa.SelectedValue.Equals(ddlEmpresaDestino.SelectedValue) Then
            MsgBox(Me.Page, "Informe uma empresa diferente a que se encontra o pedido.")
            Return False
        End If

        Return True
    End Function

    Private Function getDataSetPedidos()
        Dim sql As String = String.Empty

        sql = "select p.Pedido_Id as Pedido, c.Nome + ' (' + c.Cidade + '/' + c.Estado + ')' as Nome, isnull(p.Observacoes,'') as Observacoes, " & vbCrLf & _
                "    isnull(verAut.Pedido_Id,0) as temAut, isnull(verCom.Pedido_Id,0) as temCom, " & vbCrLf & _
                "	case " & vbCrLf & _
                "		when isnull(t.Pedido_Id,0) > 0" & vbCrLf & _
                "			then 1" & vbCrLf & _
                "			else 0" & vbCrLf & _
                "	end as temTro" & vbCrLf & _
                "from Pedidos p  " & vbCrLf & _
                "  Inner Join Clientes c " & vbCrLf & _
                "     on c.Cliente_Id  = p.Cliente " & vbCrLf & _
                "    and c.Endereco_id = p.EndCliente " & vbCrLf & _
                "	left join Pedidos t  " & vbCrLf & _
                "			on t.EmpresaTroca     = p.Empresa_id " & vbCrLf & _
                "			and t.EndEmpresaTroca = p.EndEmpresa_id " & vbCrLf & _
                "			and t.PedidoTroca     = p.Pedido_Id " & vbCrLf & _
                "	left join (select aut.Empresa_id, aut.EndEmpresa_id, aut.Pedido_Id" & vbCrLf & _
                "				from AutorizacaoDeRetirada aut  " & vbCrLf & _
                "				where aut.Empresa_id     = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                "				  and aut.EndEmpresa_id  = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            sql &= "				  and aut.Pedido_id  in(" & txtPedido.Text & ")" & vbCrLf
        End If
        sql &= "				group by aut.Empresa_id, aut.EndEmpresa_id, aut.Pedido_Id) verAut" & vbCrLf & _
                "		on verAut.Empresa_id     = p.Empresa_id " & vbCrLf & _
                "		and verAut.EndEmpresa_id = p.EndEmpresa_id " & vbCrLf & _
                "		and verAut.Pedido_id     = p.Pedido_Id " & vbCrLf & _
                "	left join (select c.Empresa_id, c.EndEmpresa_id, c.Pedido_Id" & vbCrLf & _
                "				from Comissoes c " & vbCrLf & _
                "				where c.Empresa_id     = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                "				  and c.EndEmpresa_id  = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            sql &= "				  and c.Pedido_id  in(" & txtPedido.Text & ")" & vbCrLf
        End If
        sql &= "				group by c.Empresa_id, c.EndEmpresa_id, c.Pedido_Id) verCom" & vbCrLf & _
                "		on verCom.Empresa_id     = p.Empresa_id " & vbCrLf & _
                "		and verCom.EndEmpresa_id = p.EndEmpresa_id " & vbCrLf & _
                "		and verCom.Pedido_id     = p.Pedido_Id 	" & vbCrLf & _
                "where p.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf & _
                "and p.EndEmpresa_id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                "and p.Situacao = 1 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            sql &= "    and p.Cliente = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                   "    and p.endCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            sql &= "and p.Pedido_Id in(" & txtPedido.Text & ") " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtData1.Text) AndAlso Not String.IsNullOrWhiteSpace(txtData2.Text) Then
            sql &= "    and p.DataPedido between '" & txtData1.Text.ToSqlDate & "' and '" & txtData2.Text.ToSqlDate & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlOperacao.SelectedValue) Then
            sql &= "    and p.operacao = " & ddlOperacao.SelectedValue & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlSubOperacao.SelectedValue) Then
            sql &= "    and p.suboperacao = " & ddlSubOperacao.SelectedValue.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            sql &= "    and p.safra = " & ddlSafra.SelectedValue & vbCrLf
        End If

        sql &= "group by p.Pedido_Id, c.Nome, c.Cidade, c.Estado, p.Observacoes, " & vbCrLf & _
                "		isnull(verAut.Pedido_Id,0), isnull(verCom.Pedido_Id,0), isnull(t.Pedido_Id,0)"

        Return Banco.ConsultaDataSet(sql, "Pedidos")
    End Function

    Private Function ValidaConsulta() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa para efetuar a consulta.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso (String.IsNullOrWhiteSpace(txtData1.Text) OrElse String.IsNullOrWhiteSpace(txtData2.Text)) Then
            MsgBox(Me.Page, "Informe o(s) pedido(s) ou um período para efetuar a consulta.")
            Return False
        ElseIf (Not String.IsNullOrWhiteSpace(txtData1.Text) AndAlso Not IsDate(txtData1.Text)) OrElse (Not String.IsNullOrWhiteSpace(txtData2.Text) AndAlso Not IsDate(txtData2.Text)) Then
            MsgBox(Me.Page, "Período Informado contém formato de data inválido.")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtData1.Text) AndAlso Not String.IsNullOrWhiteSpace(txtData2.Text) AndAlso CDate(txtData1.Text) > CDate(txtData2.Text) Then
            MsgBox(Me.Page, "Período inicial informado, não pode ser maior que período final.")
            Return False
        End If
        Return True
    End Function

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objCliente" & HID.Value.ToString) Is Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objCliente" & HID.Value.ToString)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value.ToString)
        End If
    End Sub

    Private Sub LimparCampos()
        lnkNovo.Parent.Visible = False
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtPedido.Text = String.Empty
        txtData1.Text = String.Empty
        txtData2.Text = String.Empty
        ddlUnidadeDestino.SelectedValue = String.Empty
        ddlOperacao.SelectedValue = String.Empty
        ddlSubOperacao.Items.Clear()
        ddlUnidadeDestino.Enabled = False
        ddlEmpresaDestino.Enabled = False
        ddlEmpresaDestino.Items.Clear()
        gridPedidos.DataBind()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

End Class