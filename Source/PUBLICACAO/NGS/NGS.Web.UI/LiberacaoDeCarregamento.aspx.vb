Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class LiberacaoDeCarregamento
    Inherits BasePage

    Private objPedido As [Lib].Negocio.Pedido
    Private sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Comercial)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("LiberacaoDeCarregamento", "ACESSAR") Then
                CargaUnidadeDeNegocio()

                Limpar(True)

                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objClienteLib" & HID.Value.ToString) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteLib" & HID.Value), [Lib].Negocio.Cliente))
                txtCodigoCliente.Value = itemCliente.Value
                txtCliente.Text = itemCliente.Text
                Session.Remove("objClienteLib" & HID.Value)

            ElseIf Not Session("objRepresentanteLib" & HID.Value.ToString) Is Nothing Then
                Dim itemRepresentante As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objRepresentanteLib" & HID.Value), [Lib].Negocio.Cliente))
                txtCodigoRepresentante.Value = itemRepresentante.Value
                txtRepresentante.Text = itemRepresentante.Text
                Session.Remove("objRepresentanteLib" & HID.Value)

            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlUnidadeDeNegocio.SelectedIndex > 0 Then
                CargaEmpresa()
            Else
                ddlEmpresa.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("4")
            Popup.ConsultaDeClientes(Me, "objClienteLib" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRepresentante_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("6")
            Popup.ConsultaDeClientes(Me, "objRepresentanteLib" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("LiberacaoDeCarregamento", "LEITURA") Then
                ConsultaRegistros()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnklnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não selecionada.")
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("LiberacaoDeCarregamento", "LIBERAR") Then
                Dim nPedido As String = String.Empty
                Dim nAnd As String = String.Empty

                Dim i As Integer = 0
                While i < gridConsulta.Rows.Count
                    If CType(gridConsulta.Rows(i).FindControl("chkLiberar"), CheckBox).Checked Then
                        nPedido += nAnd & gridConsulta.Rows(i).Cells(4).Text
                        nAnd = ","
                    End If

                    i += 1
                End While

                Dim objBanco As New AcessaBanco()

                sql = "Update Pedidos" & vbCrLf &
                        "		set LiberaCarregamento = 1 " & vbCrLf &
                        "Where Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                        "  and EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                        "  and Pedido_id in (" & nPedido & ")"

                If objBanco.GravaBanco(sql) Then
                    MsgBox(Me.Page, "Pedido(s) liberado(s) com sucesso.", eTitulo.Info)

                    Limpar(False)
                    lnkConsultar_Click(lnkConsultar, Nothing)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para liberar carregamento.")
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LiberacaoDeCarregamento")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Sub Limpar(ByVal Geral As Boolean)

        lnkConsultar.Parent.Visible = True
        lnkConfirmar.Parent.Visible = False

        If Geral Then
            txtCodigoCliente.Value = String.Empty
            txtCliente.Text = String.Empty

            txtCodigoRepresentante.Value = String.Empty
            txtRepresentante.Text = String.Empty

            gridConsulta.DataSource = Nothing
            gridConsulta.DataBind()
        End If

        Session.Remove("objClienteLib" & HID.Value)
        Session.Remove("objRepresentanteLib" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        LiberaEmpresa()
    End Sub

    Protected Sub chkLiberar_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkPedido As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkPedido.NamingContainer, GridViewRow)

        Dim temPedido As Boolean = False

        Dim i As Integer = 0
        While i < gridConsulta.Rows.Count
            If CType(gridConsulta.Rows(i).FindControl("chkLiberar"), CheckBox).Checked Then
                temPedido = True
            End If

            i += 1
        End While

        If temPedido Then
            lnkConfirmar.Parent.Visible = True
            lnkConsultar.Parent.Visible = False
        Else
            lnkConfirmar.Parent.Visible = False
            lnkConsultar.Parent.Visible = True
        End If
  
    End Sub


    Private Sub ConsultaRegistros()
        If Funcoes.VerificaPermissao("LiberacaoDeCarregamento", "LEITURA") Then
            Dim Quimica As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "05272759"
            Dim Agricola As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "04854422"
            Dim Fex As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "15204808"
            Dim Nutri As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "05366261"
            Dim NutriLog As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "38198213"
            Dim Baxi As Boolean = (Left(UsuarioServidor.CodigoEmpresa, 8) = "40938762" OrElse Left(UsuarioServidor.CodigoEmpresa, 8) = "49673784")
            Dim Verde As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "44979506"
            Dim RTGraos As Boolean = Left(UsuarioServidor.CodigoEmpresa, 8) = "24450490"

            sql = "select p.Empresa_id, p.EndEmpresa_Id, cliEmp.Nome, cliEmp.Cidade, p.Pedido_Id, P.PedidoEfetivo, p.Cliente, p.EndCliente, cli.Nome as NomeCliente, " & vbCrLf & _
                  "		p.DataPedido, p.Moeda, cope.Classe_id, " & IIf(Quimica, "prd.Grupo as Produto_Id, gp.Descricao as NomeProduto," _
                                                                       , "pXi.Produto_Id, prd.Nome as NomeProduto,") & "isnull(p.PedidoBloqueado,0) AS PedidoBloqueado, " & vbCrLf & _
                  IIf(Quimica, "Sum(pXi.Quantidade) AS Quantidade,", "pXi.Quantidade,") & vbCrLf & _
                  IIf(Quimica, "Sum(", "") & "	case " & vbCrLf & _
                  "			                      when p.Moeda = 1 " & vbCrLf & _
                  "				                   then pXi.UnitarioOficial " & vbCrLf & _
                  "				                   else pXi.UnitarioMoeda " & vbCrLf & _
                  "		                         end" & IIf(Quimica, ")", "") & " as Unitario, " & vbCrLf & _
                  IIf(Quimica, "Sum(", "") & "		case " & vbCrLf & _
                  "			                          when p.Moeda = 1 " & vbCrLf & _
                  "				                       then pXi.TotalOficial " & vbCrLf & _
                  "				                       else pXi.TotalMoeda " & vbCrLf & _
                  "		                            end" & IIf(Quimica, ")", "") & " as Total, " & vbCrLf & _
                  "		                            P.UsuarioLiberacao," & vbCrLf & _
                  "					CASE " & vbCrLf & _
                  "					   WHEN P.UsuarioLiberacao <> '' " & vbCrLf & _
                  "					     THEN P.UsuarioLiberacaoData" & vbCrLf & _
                  "                      Else NULL" & vbCrLf & _
                  "					END as UsuarioLiberacaoData " & vbCrLf & _
                  " from Pedidos p " & vbCrLf & _
                  "     inner join (Select pedI.Empresa_id, pedI.EndEmpresa_Id, pedI.Pedido_Id, pedI.Produto_Id, " & vbCrLf & _
                  "                     sum(Case  " & vbCrLf & _
                  "                           when pedI.TipoDeLancamento = 'N'  " & vbCrLf & _
                  "                            then pedI.UnitarioOficial  " & vbCrLf & _
                  "                            else 0  " & vbCrLf & _
                  "                         end) UnitarioOficial,  " & vbCrLf & _
                  "                     sum(Case  " & vbCrLf & _
                  "                           when pedI.TipoDeLancamento = 'N'  " & vbCrLf & _
                  "                            then pedI.UnitarioMoeda  " & vbCrLf & _
                  "                            else 0  " & vbCrLf & _
                  "                         end) UnitarioMoeda,  " & vbCrLf & _
                  "                     sum(case   " & vbCrLf & _
                  " 						  when pedI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  " 							then pedI.Quantidade * -1  " & vbCrLf & _
                  " 							else pedI.Quantidade  " & vbCrLf & _
                  " 						end) as Quantidade,  " & vbCrLf & _
                  " 				    sum(case   " & vbCrLf & _
                  " 						  when pedI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  " 							then pedI.TotalOficial * -1   " & vbCrLf & _
                  " 							else pedI.TotalOficial  " & vbCrLf & _
                  " 						end) as TotalOficial,  " & vbCrLf & _
                  " 				    sum(case   " & vbCrLf & _
                  " 						  when pedI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  " 							then pedI.TotalMoeda * -1   " & vbCrLf & _
                  " 							else pedI.TotalMoeda  " & vbCrLf & _
                  " 						end) as TotalMoeda  " & vbCrLf & _
                  "				from PedidoXItemxLancamento pedI " & vbCrLf & _
                  "					inner join Pedidos ped " & vbCrLf & _
                  "							on ped.Empresa_Id     = pedI.Empresa_Id " & vbCrLf & _
                  "							and ped.EndEmpresa_Id = pedI.EndEmpresa_Id " & vbCrLf & _
                  "							and ped.Pedido_Id     = pedI.Pedido_Id " & vbCrLf & _
                  "				where ped.Situacao = 1 " & vbCrLf & _
                  "				group by pedI.Empresa_Id, pedI.EndEmpresa_Id, pedI.Pedido_Id, pedI.Produto_Id) pXi " & vbCrLf & _
                  "			on pXi.Empresa_Id = p.Empresa_Id " & vbCrLf & _
                  "			and pXi.EndEmpresa_Id = p.EndEmpresa_Id " & vbCrLf & _
                  "			and pXi.Pedido_Id = p.Pedido_Id " & vbCrLf & _
                  "	    inner join Produtos prd " & vbCrLf & _
                  "		    	 on prd.Produto_id = pXi.Produto_id " & vbCrLf

            sql &= "inner Join GruposDeEstoques gp" & vbCrLf & _
               "    on gp.Grupo_Id = prd.grupo" & vbCrLf

            sql &= "	    inner join SubOperacoes so " & vbCrLf & _
                  "		    	 on so.Operacao_id     = p.Operacao " & vbCrLf & _
                  "			    and so.SubOperacoes_id = p.Suboperacao " & vbCrLf & _
                  "	    inner join ClassesDeOperacoes cope " & vbCrLf & _
                  "		    	on cope.Classe_id = so.Classe " & vbCrLf & _
                  "	    inner join Clientes cliEmp " & vbCrLf & _
                  "		    	 on cliEmp.Cliente_id  = p.Empresa_Id " & vbCrLf & _
                  "		    	and cliEmp.Endereco_id = p.EndEmpresa_Id " & vbCrLf & _
                  "	    inner join Clientes cli " & vbCrLf & _
                  "		    	 on cli.Cliente_id  = p.Cliente " & vbCrLf & _
                  "		    	and cli.Endereco_id = p.EndCliente " & vbCrLf & _
                  "	    left join Comissoes repre " & vbCrLf & _
                  "		    	 on repre.Empresa_Id    = p.Empresa_Id " & vbCrLf & _
                  "		    	and repre.EndEmpresa_Id = p.EndEmpresa_Id " & vbCrLf & _
                  "		    	and repre.Pedido_Id     = p.Pedido_id " & vbCrLf & _
                  " where p.Situacao  = 1 " & vbCrLf & _
                  "  and p.Empresa_id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf & _
                  "  and isnull(p.PedidoBloqueado,0) = 0 " & vbCrLf & _
                  "  and isnull(p.LiberaCarregamento,0) = 0" & vbCrLf & _
                  "  and so.EntradaSaida = 'S'" & vbCrLf & _
                  "  and so.Devolucao    = 'N'" & vbCrLf

            If rdCIF.Checked Then
                sql &= "  and p.FreteCIFFOB    = 'CIF'"
            Else
                sql &= "  and p.FreteCIFFOB    = 'FOB'"
            End If

            If txtCliente.Text.Length > 0 Then
                Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")

                If chkConsolidarCliente.Checked Then
                    sql &= " AND left(P.Cliente,8)    ='" & Cliente(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "  and p.Cliente    = '" & Cliente(0) & "'" & vbCrLf & _
                            "  and p.EndCliente = " & Cliente(1) & vbCrLf
                End If
            End If

            If txtRepresentante.Text.Length > 0 Then
                Dim Representante() As String = txtCodigoRepresentante.Value.ToString.Split("-")

                sql &= "  and repre.Representante_Id = '" & Representante(0) & "'" & vbCrLf & _
                    "  and repre.EndRepresentante_Id = " & Representante(1) & vbCrLf
            End If

            If txtConsultaPedido.Text.Length > 0 Then
                sql &= " and p.Pedido_id = '" & Trim(txtConsultaPedido.Text) & "' " & vbCrLf
            ElseIf chkPeriodo.Checked Then
                sql &= " and p.Datapedido between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf
            End If

            sql &= "  and not  exists(select 1" & vbCrLf & _
                   "                    from notasfiscais nf " & vbCrLf & _
                   "                   where nf.empresa_id = p.empresa_id" & vbCrLf & _
                   "                     and nf.endempresa_id = p.endempresa_id" & vbCrLf & _
                   "                     and nf.pedido        = p.pedido_id" & vbCrLf & _
                   "                     and isnull(nf.nfg,0) = 1)" & vbCrLf

            sql &= "order by p.Datapedido " & vbCrLf

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(sql, "ConsultaPedidos")

            gridConsulta.DataSource = Nothing
            gridConsulta.DataBind()

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum Registro Encontrado")
                Exit Sub
            End If

            Dim dtPedidos As New DataTable("listPedidos")
            dtPedidos.Columns.Add("Empresa", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Cliente", Type.GetType("System.String"))
            dtPedidos.Columns.Add("NomeCliente", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Pedido", Type.GetType("System.String"))
            dtPedidos.Columns.Add("PedidoEfetivo", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Data", Type.GetType("System.DateTime"))
            dtPedidos.Columns.Add("Classe", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Moeda", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Produto", Type.GetType("System.String"))
            dtPedidos.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("Unitario", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("Valor", Type.GetType("System.Decimal"))
            dtPedidos.Columns.Add("UsuarioLiberacao", Type.GetType("System.String"))
            dtPedidos.Columns.Add("UsuarioLiberacaoData", Type.GetType("System.DateTime"))

            For Each dr As DataRow In ds.Tables(0).Rows
                Dim drPedido As DataRow = dtPedidos.NewRow()
                drPedido("Empresa") = dr("Empresa_Id") & "-" & Left(dr("Nome"), 20) & "/" & dr("Cidade")
                drPedido("Cliente") = dr("Cliente") & "-" & dr("EndCliente")
                drPedido("NomeCliente") = dr("NomeCliente")
                drPedido("Pedido") = dr("Pedido_Id")
                drPedido("PedidoEfetivo") = dr("PedidoEfetivo")
                drPedido("Data") = dr("DataPedido")
                drPedido("Classe") = dr("Classe_id")
                If dr("Moeda") = 1 Then
                    drPedido("Moeda") = "R$"
                Else
                    drPedido("Moeda") = "U$"
                End If
                drPedido("Produto") = dr("Produto_Id") & "-" & dr("NomeProduto")
                drPedido("Quantidade") = dr("Quantidade")
                drPedido("Unitario") = dr("Unitario")
                drPedido("Valor") = dr("Total")
                drPedido("UsuarioLiberacao") = dr("UsuarioLiberacao")
                drPedido("UsuarioLiberacaoData") = dr("UsuarioLiberacaoData")
                dtPedidos.Rows.Add(drPedido)
            Next

            gridConsulta.DataSource = dtPedidos
            gridConsulta.DataBind()

        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub
End Class