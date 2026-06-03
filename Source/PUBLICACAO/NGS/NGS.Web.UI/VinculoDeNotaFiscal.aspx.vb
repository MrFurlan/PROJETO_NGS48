Imports NGS.Lib.Negocio

Public Class VinculoDeNotaFiscal
    Inherits BasePage

    Protected objCliente As New Cliente

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("VinculoDeNotaFiscal", "ACESSAR") Then
                    CargaEmpresas()
                    ddl.Carregar(ddlTipoDeDocumento, CarregarDDL.Tabela.TipoDeDocumento, " Codigo_Id <> 4", True)
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objNotaVinculo" & HID.Value.ToString) Is Nothing Then
            objCliente = CType(Session("objNotaVinculo" & HID.Value.ToString), [Lib].Negocio.Cliente)
            If (hdnControlePopup.Value.Equals("ClienteNota")) Then
                txtCliente.Text = Funcoes.AlinharEsquerda(objCliente.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(objCliente.Cidade, 20, ".") & " " & objCliente.CodigoEstado & " " & Funcoes.AlinharEsquerda(objCliente.CodigoFormatado, 18, ".") & "-" & CStr(objCliente.CodigoEndereco) & "-" & objCliente.Reduzido
                txtCodigoCliente.Value = objCliente.Codigo
                hdnEnderecoCliente.Value = objCliente.CodigoEndereco
            Else
                txtClientesVinc.Text = Funcoes.AlinharEsquerda(objCliente.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(objCliente.Cidade, 20, ".") & " " & objCliente.CodigoEstado & " " & Funcoes.AlinharEsquerda(objCliente.CodigoFormatado, 18, ".") & "-" & CStr(objCliente.CodigoEndereco) & "-" & objCliente.Reduzido
                txtCodigoClienteVinc.Value = objCliente.Codigo
                hdnEnderecoClienteVinc.Value = objCliente.CodigoEndereco
            End If
        ElseIf Not Session("objLiberaInutilizacao" & HID.Value) Is Nothing Then
            hdLiberar.Value = CType(Session("objLiberaInutilizacao" & HID.Value), Boolean)
        End If
    End Sub

    Private Sub Limpar()
        Session.Remove("objClienteVinc" & HID.Value.ToString)
        Session.Remove("objListaDeNotas")
        Session.Remove("objTrocaDeNotaVinculadas" & HID.Value.ToString)
        Session.Remove("objLiberaInutilizacao" & HID.Value)


        ddlEmpresa.SelectedValue = String.Format("{0}-{1}", UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        txtDataInicial.Text = Now().ToShortDateString()
        txtDataFinal.Text = Now().ToShortDateString()
        txtDataInicialVinc.Text = Now().ToShortDateString()
        txtDataFinalVinc.Text = Now().ToShortDateString()
        txtClientesVinc.Text = ""
        txtCodigoClienteVinc.Value = ""
        txtNotaVinc.Text = ""
        rdEntrada.Checked = True
        rdEntradaVinc.Checked = True
        txtNotaFiscal.Text = ""
        pnlVinculoNota.Parent.Visible = False
        lnkLiberar.Parent.Visible = False
        hdLiberar.Value = False

        LiberaEmpresa()

        gridVinculoDeNota.DataBind()
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
        ddl.Carregar(ddlEmpresaVinc, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Function ValidarConsulta() As Boolean
        If String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) _
            OrElse Not IsDate(txtDataInicial.Text) OrElse Not IsDate(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Informe um período válido.")
            Return False
        End If
        Return True
    End Function


    Private Sub ConsultarRegistro()
        Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")

        Dim cliente As String = String.Empty
        Dim endCliente As Integer = 0

        If txtCodigoCliente.Value.Length > 0 Then
            cliente = txtCodigoCliente.Value.Split("-")(0)
            endCliente = txtCodigoCliente.Value.Split("-")(1)
        End If

        Dim numNota As Integer = 0
        Dim iTipoDeDocumento As Integer = 0

        If ddlTipoDeDocumento.SelectedValue = "" Then
            iTipoDeDocumento = 0
        Else
            iTipoDeDocumento = ddlTipoDeDocumento.SelectedValue
        End If

        If txtNotaFiscal.Text.Length > 0 Then numNota = CInt(txtNotaFiscal.Text)


        If numNota = 0 Then
            MsgBox(Me.Page, "Digite o numero da nota para pesquisa.")
            Exit Sub
        End If

        Dim ListaDeNotas As New [Lib].Negocio.ListNotasFiscais()

        If chkMovimento.Checked Then
            ListaDeNotas = New [Lib].Negocio.ListNotasFiscais(Empresa(0), Empresa(1), txtDataInicial.Text, txtDataFinal.Text, cliente, endCliente, IIf(rdEntrada.Checked, "E", "S"), "", iTipoDeDocumento, numNota)
        Else
            ListaDeNotas = New [Lib].Negocio.ListNotasFiscais(Empresa(0), Empresa(1), "", "", cliente, endCliente, IIf(rdEntrada.Checked, "E", "S"), "", iTipoDeDocumento, numNota)
        End If

        Session("objListaDeNotas") = ListaDeNotas

        If ListaDeNotas.Count > 0 Then
            Dim ds As New DataSet
            Dim tbTroca As New DataTable("Clientes")
            tbTroca.Columns.Add("Empresa", Type.GetType("System.String"))
            'tbTroca.Columns.Add("EndEmpresa", Type.GetType("System.String"))
            tbTroca.Columns.Add("Cliente", Type.GetType("System.String"))
            'tbTroca.Columns.Add("EndCliente", Type.GetType("System.String"))
            'tbTroca.Columns.Add("NomeCliente", Type.GetType("System.String"))
            tbTroca.Columns.Add("EntradaSaida_Id", Type.GetType("System.String"))
            tbTroca.Columns.Add("Serie_Id", Type.GetType("System.String"))
            tbTroca.Columns.Add("Nota_Id", Type.GetType("System.String"))
            tbTroca.Columns.Add("Movimento", Type.GetType("System.DateTime"))
            tbTroca.Columns.Add("Descricao", Type.GetType("System.String"))
            tbTroca.Columns.Add("OrigemEmpresa", Type.GetType("System.String"))
            'tbTroca.Columns.Add("OEndEmpresa", Type.GetType("System.String"))
            tbTroca.Columns.Add("OrigemCliente", Type.GetType("System.String"))
            'tbTroca.Columns.Add("OEndCliente", Type.GetType("System.String"))
            'tbTroca.Columns.Add("ONomeCliente", Type.GetType("System.String"))
            tbTroca.Columns.Add("OrigemEntradaSaida_Id", Type.GetType("System.String"))
            tbTroca.Columns.Add("OrigemSerie_Id", Type.GetType("System.String"))
            tbTroca.Columns.Add("OrigemNota_Id", Type.GetType("System.String"))
            tbTroca.Columns.Add("OrigemMovimento", Type.GetType("System.DateTime"))
            tbTroca.Columns.Add("OrigemDescricao", Type.GetType("System.String"))
            ds.Tables.Add(tbTroca)

            Dim i As Integer = 0

            For Each nf As [Lib].Negocio.NotaFiscal In ListaDeNotas
                Dim drRow As DataRow = ds.Tables(0).NewRow()

                drRow("Empresa") = nf.CodigoEmpresa & "-" & nf.EnderecoEmpresa
                'drRow("EndEmpresa") = nf.EnderecoEmpresa
                drRow("Cliente") = nf.CodigoCliente & "-" & nf.EnderecoCliente & " - " & nf.Cliente.Nome
                'drRow("EndCliente") = nf.EnderecoCliente
                'drRow("NomeCliente") = nf.Cliente.Nome
                drRow("EntradaSaida_Id") = IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                drRow("Serie_Id") = nf.Serie
                drRow("Nota_Id") = nf.Codigo
                drRow("Movimento") = nf.Movimento
                drRow("Descricao") = nf.TipoDeDocumento.Descricao

                If Not nf.NotasXNotas Is Nothing AndAlso Not nf.NotasXNotas.EmpresaCnpj Is Nothing AndAlso Not nf.NotasTrocaOrigem Is Nothing Then
                    For Each ori In nf.NotasTrocaOrigem

                        'A pedido da BAXI o tipo de produtor foi necessario habilitar, devido a ajustes no vinculo
                        If ori.CodigoTipoDeDocumento = 1 Or ori.CodigoTipoDeDocumento = 15 Then
                            drRow("OrigemEmpresa") = nf.NotasXNotas.OrigemEmpresaCnpj & "-" & nf.NotasXNotas.OrigemEndEmpresa
                            'drRow("OEndEmpresa") = nf.NotasXNotas.OrigemEndEmpresa
                            Dim objCliente As New [Lib].Negocio.Cliente(nf.NotasXNotas.OrigemClienteCnpj, nf.NotasXNotas.OrigemEndCliente)
                            drRow("OrigemCliente") = nf.NotasXNotas.OrigemClienteCnpj & "-" & nf.NotasXNotas.OrigemEndCliente & " - " & objCliente.Nome
                            'drRow("OEndCliente") = nf.NotasXNotas.OrigemEndCliente
                            'drRow("ONomeCliente") = objCliente.Nome
                            drRow("OrigemEntradaSaida_Id") = nf.NotasXNotas.OrigemEntradaSaida
                            drRow("OrigemSerie_Id") = nf.NotasXNotas.OrigemSerie
                            drRow("OrigemNota_id") = nf.NotasXNotas.OrigemNota
                            drRow("OrigemMovimento") = ori.Movimento
                            drRow("OrigemDescricao") = ori.TipoDeDocumento.Descricao
                            Exit For
                        Else
                            drRow("OrigemEmpresa") = "Sem Vinculo"
                        End If
                    Next
                Else
                    'MsgBox(Me.Page, "Vinculo não encontrado.")
                    drRow("OrigemEmpresa") = "Sem Vinculo"
                End If

                ds.Tables(0).Rows.Add(drRow)
            Next

            gridVinculoDeNota.DataSource = ds
            gridVinculoDeNota.DataBind()

            While i < gridVinculoDeNota.Rows.Count

                If gridVinculoDeNota.Rows(i).Cells(9).Text = "Sem Vinculo" Or gridVinculoDeNota.Rows(i).Cells(10).Text = "&nbsp;" Then
                    CType(gridVinculoDeNota.Rows(i).FindControl("imgExcluir"), ImageButton).Visible = False
                    CType(gridVinculoDeNota.Rows(i).FindControl("imgConsultar"), ImageButton).Visible = True
                Else
                    CType(gridVinculoDeNota.Rows(i).FindControl("imgExcluir"), ImageButton).Visible = True
                    CType(gridVinculoDeNota.Rows(i).FindControl("imgConsultar"), ImageButton).Visible = False
                End If

                i += 1
            End While

        Else
            MsgBox(Me.Page, "Lista não encontrada no período determinado.")
        End If
    End Sub

    Private Function ConsultarNotas() As Boolean
        If String.IsNullOrWhiteSpace(txtNotaFiscal.Text) Then
            MsgBox(Me.Page, "Informe o número da Nota Fiscal.")
            Return False
        Else
            Dim ds As New DataSet
            Dim Sql = "SELECT (NxN.Empresa_Id + '-' + convert(varchar,NxN.EndEmpresa_Id)) AS Empresa, (NxN.Cliente_Id + '-' + " & vbCrLf &
                    "		convert(varchar,NxN.EndCliente_Id) + ' - ' + CNF.Nome) AS Cliente, " & vbCrLf &
                    "		NxN.EntradaSaida_Id, NxN.Nota_Id, NxN.Serie_Id, " & vbCrLf &
                    "		NF.Movimento, TdNF.Descricao, " & vbCrLf &
                    "		(NxN.OrigemEmpresa_Id + '-' + convert(varchar,NxN.OrigemEndEmpresa_Id)) AS OrigemEmpresa, " & vbCrLf &
                    "		(NxN.OrigemCliente_Id + '-' + convert(varchar,NxN.OrigemEndCliente_Id) + ' - ' + CNFO.Nome) AS OrigemCliente, " & vbCrLf &
                    "		NxN.OrigemEntradaSaida_Id, NxN.OrigemNota_Id, NxN.OrigemSerie_Id, " & vbCrLf &
                    "		NFO.Movimento AS OrigemMovimento, TdNFO.Descricao AS OrigemDescricao " & vbCrLf &
                    "FROM  NotasXNotas AS NxN " & vbCrLf &
                    "	INNER JOIN NotasFiscais NF " & vbCrLf &
                    "			 ON NF.Empresa_Id      = NxN.Empresa_Id  " & vbCrLf &
                    "			AND NF.EndEmpresa_Id   = NxN.EndEmpresa_Id  " & vbCrLf &
                    "			AND NF.Cliente_Id      = NxN.Cliente_Id  " & vbCrLf &
                    "			AND NF.EndCliente_Id   = NxN.EndCliente_Id  " & vbCrLf &
                    "			AND NF.EntradaSaida_Id = NxN.EntradaSaida_Id  " & vbCrLf &
                    "			AND NF.Serie_Id        = NxN.Serie_Id  " & vbCrLf &
                    "			AND NF.Nota_Id         = NxN.Nota_Id  " & vbCrLf &
                    "	INNER JOIN TipoDeDocumento TdNF " & vbCrLf &
                    "			 ON TdNF.Codigo_Id     = NF.TipoDeDocumento  " & vbCrLf &
                    "	INNER JOIN Clientes CNF " & vbCrLf &
                    "			 ON CNF.Cliente_Id      = NxN.Cliente_Id  " & vbCrLf &
                    "			AND CNF.Endereco_Id     = NxN.EndCliente_Id  " & vbCrLf &
                    "	INNER JOIN NotasFiscais NFO " & vbCrLf &
                    "			 ON NFO.Empresa_Id      = NxN.OrigemEmpresa_Id  " & vbCrLf &
                    "			AND NFO.EndEmpresa_Id   = NxN.OrigemEndEmpresa_Id  " & vbCrLf &
                    "			AND NFO.Cliente_Id      = NxN.OrigemCliente_Id  " & vbCrLf &
                    "			AND NFO.EndCliente_Id   = NxN.OrigemEndCliente_Id  " & vbCrLf &
                    "			AND NFO.EntradaSaida_Id = NxN.OrigemEntradaSaida_Id  " & vbCrLf &
                    "			AND NFO.Serie_Id        = NxN.OrigemSerie_Id  " & vbCrLf &
                    "			AND NFO.Nota_Id         = NxN.OrigemNota_Id  " & vbCrLf &
                    "	INNER JOIN TipoDeDocumento TdNFO " & vbCrLf &
                    "			 ON TdNFO.Codigo_Id     = NFO.TipoDeDocumento  " & vbCrLf &
                    "	INNER JOIN Clientes CNFO " & vbCrLf &
                    "			 ON CNFO.Cliente_Id     = NxN.OrigemCliente_Id  " & vbCrLf &
                    "			AND CNFO.Endereco_Id    = NxN.OrigemEndCliente_Id  " & vbCrLf &
                    " WHERE (NxN.Nota_Id = " & txtNotaFiscal.Text & " OR NxN.OrigemNota_Id = " & txtNotaFiscal.Text & ")" & vbCrLf &
                    " ORDER BY NF.Movimento DESC " & vbCrLf
            ds = Banco.ConsultaDataSet(Sql, "NotasXNotas")

            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Não foram encontrados vínculos com a nota " & txtNotaFiscal.Text)
                Return False
            Else
                gridVinculoDeNota.DataSource = ds
                gridVinculoDeNota.DataBind()

                Dim i As Integer = 0

                While i < gridVinculoDeNota.Rows.Count

                    CType(gridVinculoDeNota.Rows(i).FindControl("imgExcluir"), ImageButton).Visible = True
                    CType(gridVinculoDeNota.Rows(i).FindControl("imgConsultar"), ImageButton).Visible = False

                    i += 1
                End While

                Return True
            End If
        End If
    End Function

    Private Sub ConsultarVinculo()
        If Funcoes.VerificaPermissao("NotaFiscalXItens", "LEITURA") Then
            Dim strJavaScript As String = ""
            Dim Empresa() As String = ddlEmpresaVinc.SelectedValue.ToString.Split("-")
            Dim Cliente() As String = txtCodigoClienteVinc.Value.Split("-")

            Session("ssCampo") = "NotaXItens"

            Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal
            objNotaFiscal = CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(HgridRowIndexVinculo.Value)
            If objNotaFiscal.Itens.Count = 0 Then
                MsgBox(Me.Page, "A Nota deve conter um Produto para o vínculo de nota ser realizado")
                Exit Sub
            End If

            If objNotaFiscal.Itens(0).Produto.Agrupar.Equals("S") Then
                MsgBox(Me.Page, "Não é permitido fazer vinculo de nota para este produto")
                Exit Sub
            End If
            If objNotaFiscal.Pedido.SubOperacao.PrecoFixo = False And objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) = "E" Then
                MsgBox(Me.Page, "A Devolução para Compra só é permitida em operações de preço Fixo")
                Exit Sub
            End If

            Dim objNotaVinculo As New [Lib].Negocio.NotaFiscal
            objNotaVinculo.Codigo = txtNotaVinc.Text
            objNotaVinculo.CodigoEmpresa = Empresa(0)
            objNotaVinculo.EnderecoEmpresa = Empresa(1)
            objNotaVinculo.CodigoCliente = txtCodigoClienteVinc.Value
            objNotaVinculo.EnderecoCliente = hdnEnderecoClienteVinc.Value

            If rdEntradaVinc.Checked Then
                objNotaVinculo.EntradaSaida = 0
            Else
                objNotaVinculo.EntradaSaida = 1
            End If

            objNotaVinculo.DataInclusao = txtDataInicialVinc.Text
            objNotaVinculo.DataTermino = txtDataFinalVinc.Text
            objNotaVinculo.Itens = objNotaFiscal.Itens
            objNotaVinculo.VincularNotas = True

            Session("objNotaFiscal" & HID.Value.ToString) = objNotaFiscal

            ucConsultaNotaTroca.SetarHID(HID.Value)
            ucConsultaNotaTroca.InicializarFormulario(False, objNotaVinculo)
            Popup.ConsultaDeNotaTroca(Me, "objNotaVinculo" & Guid.NewGuid().ToString)

        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar Registro")
        End If
    End Sub

    Protected Sub lnkLiberar_Click(sender As Object, e As EventArgs) Handles lnkLiberar.Click
        Try
            hdLiberar.Value = False
            ucLiberacao.Limpar()
            Popup.ConsultaLiberacao(Me.Page, "objLiberaInutilizacao" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaClienteVinc_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            hdnControlePopup.Value = "ClienteTrocaNota"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTituloDIV("Consulta de Clientes")
            Popup.ConsultaDeClientes(Me.Page, "objNotaVinculo" & HID.Value.ToString)
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgConsultar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("VinculoDeNotaFiscal", "LEITURA") Then
                Dim imgNotaFiscal As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgNotaFiscal.NamingContainer, GridViewRow)

                HgridRowIndexVinculo.Value = row.RowIndex

                lblNotaFiscal.Text = CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(HgridRowIndexVinculo.Value).Codigo

                txtClientesVinc.Text = ""
                txtCodigoClienteVinc.Value = ""
                txtNotaVinc.Text = ""
                If rdEntrada.Checked = True Then
                    rdSaidaVinc.Checked = True
                End If
                If rdSaida.Checked = True Then
                    rdEntradaVinc.Checked = True
                End If

                ddlEmpresaVinc.SelectedValue = String.Format("{0}-{1}", UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)

                pnlVinculoNota.Parent.Visible = True
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("VinculoDeNotaFiscal", "EXCLUIR") Then
                Dim Sqls As New ArrayList
                Dim strSQL As String = ""

                Dim imgNotaFiscal As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgNotaFiscal.NamingContainer, GridViewRow)

                'If CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).TemNotaTroca Then

                '    CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).IUD = "D"
                '    CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).SalvarNotasXNotas(Sqls)

                '    Dim obs As String = CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).ObservacoesControleInterno
                '    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                '        obs = obs & ". Removido Vinculo de Nota Fiscal em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                '    Else
                '        obs = "Removido Vinculo de Nota Fiscal em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                '    End If

                '    strSQL &= " Update NotasFiscais set UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                '                 "                      ,UsuarioAlteracaoData       = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                '                 "                      ,Troca                      = 0" & vbCrLf &
                '                 "                      ,ObservacoesControleInterno = '" & obs & "'" & vbCrLf &
                '              " Where Empresa_Id      ='" & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).CodigoEmpresa & "'" & vbCrLf &
                '              "   and EndEmpresa_Id   = " & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).EnderecoEmpresa & vbCrLf &
                '              "   and Cliente_Id      ='" & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).CodigoCliente & "'" & vbCrLf &
                '              "   and EndCliente_Id   = " & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).EnderecoCliente & vbCrLf &
                '              "   and EntradaSaida_Id ='" & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                '              "   and Serie_Id        ='" & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).Serie & "'" & vbCrLf &
                '              "   and Nota_Id         = " & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).Codigo

                '    Sqls.Add(strSQL)

                '    If Not CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).NotasXNotas Is Nothing Then
                '        strSQL &= " Update NotasFiscais set UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                '                 "                      ,UsuarioAlteracaoData       = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                '                 "                      ,Troca                      = 0" & vbCrLf &
                '                 "                      ,ObservacoesControleInterno = '" & obs & "'" & vbCrLf &
                '              " Where Empresa_Id      ='" & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).NotasXNotas.OrigemEmpresaCnpj & "'" & vbCrLf &
                '              "   and EndEmpresa_Id   = " & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).NotasXNotas.OrigemEndEmpresa & vbCrLf &
                '              "   and Cliente_Id      ='" & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).NotasXNotas.OrigemClienteCnpj & "'" & vbCrLf &
                '              "   and EndCliente_Id   = " & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).NotasXNotas.OrigemEndCliente & vbCrLf &
                '              "   and EntradaSaida_Id ='" & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).NotasXNotas.OrigemEntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                '              "   and Serie_Id        ='" & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).NotasXNotas.OrigemSerie & "'" & vbCrLf &
                '              "   and Nota_Id         = " & CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).NotasXNotas.OrigemNota
                '        Sqls.Add(strSQL)

                '    End If


                '    If Banco.GravaBanco(Sqls) Then
                '        Limpar()
                '        MsgBox(Me.Page, "Registro incluído com Sucesso.", eTitulo.Sucess)
                '    Else
                '        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                '    End If
                'End If

                Dim Empresa As String() = gridVinculoDeNota.Rows(row.RowIndex).Cells(1).Text.Split("-")
                Dim Cliente As String() = gridVinculoDeNota.Rows(row.RowIndex).Cells(2).Text.Split("-")
                Dim OrigemCliente As String() = gridVinculoDeNota.Rows(row.RowIndex).Cells(10).Text.Split("-")

                strSQL = "Delete NotasXNotas" & vbCrLf &
                      " Where Empresa_Id      ='" & Empresa(0) & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & Empresa(1) & vbCrLf &
                      "   and Cliente_Id      ='" & Cliente(0) & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & Cliente(1) & vbCrLf &
                      "   and EntradaSaida_Id ='" & gridVinculoDeNota.Rows(row.RowIndex).Cells(3).Text & "'" & vbCrLf &
                      "   and Serie_Id        ='" & gridVinculoDeNota.Rows(row.RowIndex).Cells(4).Text & "'" & vbCrLf &
                      "   and Nota_Id         = " & gridVinculoDeNota.Rows(row.RowIndex).Cells(5).Text & vbCrLf &
                      "   and OrigemEmpresa_Id      ='" & Empresa(0) & "'" & vbCrLf &
                      "   and OrigemEndEmpresa_Id   = " & Empresa(1) & vbCrLf &
                      "   and OrigemCliente_Id      ='" & OrigemCliente(0) & "'" & vbCrLf &
                      "   and OrigemEndCliente_Id   = " & OrigemCliente(1) & vbCrLf &
                      "   and OrigemEntradaSaida_Id ='" & gridVinculoDeNota.Rows(row.RowIndex).Cells(11).Text & "'" & vbCrLf &
                      "   and OrigemSerie_Id        ='" & gridVinculoDeNota.Rows(row.RowIndex).Cells(12).Text & "'" & vbCrLf &
                      "   and OrigemNota_Id         = " & gridVinculoDeNota.Rows(row.RowIndex).Cells(13).Text & vbCrLf

                Sqls.Add(strSQL)

                Dim nf As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                nf.CodigoEmpresa = Empresa(0)
                nf.EnderecoEmpresa = Empresa(1)
                nf.CodigoCliente = Cliente(0)
                nf.EnderecoCliente = Cliente(1)
                nf.EntradaSaida = IIf(gridVinculoDeNota.Rows(row.RowIndex).Cells(3).Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                nf.Codigo = gridVinculoDeNota.Rows(row.RowIndex).Cells(5).Text
                nf.Serie = gridVinculoDeNota.Rows(row.RowIndex).Cells(4).Text
                nf = New [Lib].Negocio.NotaFiscal(nf)

                If nf.CodigoTipoDeDocumento = 57 AndAlso hdLiberar.Value = False Then
                    lnkLiberar.Parent.Visible = True
                    MsgBox(Me.Page, "Vinculo com Conhecimento de Transporte não pode ser removido.", eTitulo.Info)
                    Exit Sub
                End If

                Dim obs As String = nf.ObservacoesControleInterno

                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Removido Vinculo com Nota Fiscal " & gridVinculoDeNota.Rows(row.RowIndex).Cells(13).Text & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                Else
                    obs = "Removido Vinculo de Nota Fiscal " & gridVinculoDeNota.Rows(row.RowIndex).Cells(13).Text & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                End If

                strSQL = " Update NotasFiscais set UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                             "                      ,UsuarioAlteracaoData       = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                             "                      ,Troca                      = 0" & vbCrLf &
                             "                      ,ObservacoesControleInterno = '" & obs & "'" & vbCrLf &
                          " Where Empresa_Id      ='" & nf.CodigoEmpresa & "'" & vbCrLf &
                          "   and EndEmpresa_Id   = " & nf.EnderecoEmpresa & vbCrLf &
                          "   and Cliente_Id      ='" & nf.CodigoCliente & "'" & vbCrLf &
                          "   and EndCliente_Id   = " & nf.EnderecoCliente & vbCrLf &
                          "   and EntradaSaida_Id ='" & nf.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "   and Serie_Id        ='" & nf.Serie & "'" & vbCrLf &
                          "   and Nota_Id         = " & nf.Codigo

                Sqls.Add(strSQL)

                Dim nfOrigem As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
                nfOrigem.CodigoEmpresa = Empresa(0)
                nfOrigem.EnderecoEmpresa = Empresa(1)
                nfOrigem.CodigoCliente = OrigemCliente(0)
                nfOrigem.EnderecoCliente = OrigemCliente(1)
                nfOrigem.EntradaSaida = IIf(gridVinculoDeNota.Rows(row.RowIndex).Cells(11).Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                nfOrigem.Codigo = gridVinculoDeNota.Rows(row.RowIndex).Cells(13).Text
                nfOrigem.Serie = gridVinculoDeNota.Rows(row.RowIndex).Cells(12).Text
                nfOrigem = New [Lib].Negocio.NotaFiscal(nfOrigem)

                obs = nfOrigem.ObservacoesControleInterno

                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Removido Vinculo de Nota Fiscal " & gridVinculoDeNota.Rows(row.RowIndex).Cells(5).Text & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                Else
                    obs = "Removido Vinculo de Nota Fiscal " & gridVinculoDeNota.Rows(row.RowIndex).Cells(5).Text & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                End If

                strSQL = " Update NotasFiscais set UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                             "                      ,UsuarioAlteracaoData       = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                             "                      ,Troca                      = 0" & vbCrLf &
                             "                      ,ObservacoesControleInterno = '" & obs & "'" & vbCrLf &
                          " Where Empresa_Id      ='" & nfOrigem.CodigoEmpresa & "'" & vbCrLf &
                          "   and EndEmpresa_Id   = " & nfOrigem.EnderecoEmpresa & vbCrLf &
                          "   and Cliente_Id      ='" & nfOrigem.CodigoCliente & "'" & vbCrLf &
                          "   and EndCliente_Id   = " & nfOrigem.EnderecoCliente & vbCrLf &
                          "   and EntradaSaida_Id ='" & nfOrigem.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "   and Serie_Id        ='" & nfOrigem.Serie & "'" & vbCrLf &
                          "   and Nota_Id         = " & nfOrigem.Codigo

                Sqls.Add(strSQL)

                If Banco.GravaBanco(Sqls) Then
                    Limpar()
                    MsgBox(Me.Page, "Vinculo removido com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If

            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("VinculoDeNotaFiscal", "LEITURA") Then
                If ValidarConsulta() Then
                    'ConsultarRegistro()
                    If Not ConsultarNotas() Then ConsultarRegistro()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objClienteVinc" & HID.Value.ToString)
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

    Protected Sub lnkVincular_Click(sender As Object, e As EventArgs) Handles lnkVincular.Click
        Try
            If Funcoes.VerificaPermissao("VinculoDeNotaFiscal", "GRAVAR") Then
                Dim Sqls As New ArrayList
                Dim strSQL As String = ""

                Dim objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
                objNotaFiscal.NotaTrocaOrigem = CType(Session("objTrocaDeNotaVinculadas" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)

                If objNotaFiscal.NotaTrocaOrigem.Codigo = 0 Then
                    MsgBox(Me.Page, "Nota para vincular não encontrada.")
                    Exit Sub
                End If

                objNotaFiscal.IUD = "I"
                objNotaFiscal.SalvarNotasXNotas(Sqls)

                Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Vinculo de Nota Fiscal em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                Else
                    obs = "Vinculo de Nota Fiscal em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                End If

                strSQL &= " Update NotasFiscais set UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                             "                      ,UsuarioAlteracaoData       = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                             "                      ,ObservacoesControleInterno = '" & obs & "'" & vbCrLf &
                          " Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                          "   and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                          "   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                          "   and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                          "   and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "   and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                          "   and Nota_Id         = " & objNotaFiscal.Codigo

                Sqls.Add(strSQL)

                strSQL &= " Update NotasFiscais set UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                             "                      ,UsuarioAlteracaoData       = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                             "                      ,ObservacoesControleInterno = '" & obs & "'" & vbCrLf &
                          " Where Empresa_Id      ='" & objNotaFiscal.NotaTrocaOrigem.CodigoEmpresa & "'" & vbCrLf &
                          "   and EndEmpresa_Id   = " & objNotaFiscal.NotaTrocaOrigem.EnderecoEmpresa & vbCrLf &
                          "   and Cliente_Id      ='" & objNotaFiscal.NotaTrocaOrigem.CodigoCliente & "'" & vbCrLf &
                          "   and EndCliente_Id   = " & objNotaFiscal.NotaTrocaOrigem.EnderecoCliente & vbCrLf &
                          "   and EntradaSaida_Id ='" & objNotaFiscal.NotaTrocaOrigem.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "   and Serie_Id        ='" & objNotaFiscal.NotaTrocaOrigem.Serie & "'" & vbCrLf &
                          "   and Nota_Id         = " & objNotaFiscal.NotaTrocaOrigem.Codigo

                Sqls.Add(strSQL)

                If Banco.GravaBanco(Sqls) Then
                    Limpar()
                    MsgBox(Me.Page, "Registro incluído com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarVinc_Click(sender As Object, e As EventArgs) Handles lnkConsultarVinc.Click
        Try
            If ddlEmpresaVinc.SelectedIndex = 0 Then
                MsgBox(Me.Page, "A Empresa da Nota Fiscal para Vínculo não foi selecionada.")
                ddlEmpresaVinc.Focus()
            ElseIf String.IsNullOrWhiteSpace(txtCodigoClienteVinc.Value) Then
                MsgBox(Me.Page, "O Cliente da Nota Fiscal para Vinculo não foi selecionado.")
                btnConsultaClienteVinc.Focus()
            Else
                ConsultarVinculo()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class