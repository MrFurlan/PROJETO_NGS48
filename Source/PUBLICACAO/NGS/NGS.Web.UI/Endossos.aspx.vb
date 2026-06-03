Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports System.Drawing

Public Class Endossos
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Dim objEndosso As Endosso

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Endossos", "ACESSAR") Then
                BuncarUnidadeDeNegocio()

                Limpar(True)

                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

                txtBuscaDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtBuscaDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        ElseIf Session("objPedEndosso" & HID.Value) IsNot Nothing Then
            Dim p As [Lib].Negocio.Pedido = CType(Session("objPedEndosso" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = p.Codigo
            Session.Remove("objPedEndosso" & HID.Value)
        ElseIf Not Session("objClienteEndosso" & HID.Value) Is Nothing Then
            Dim itemClienteEndosso As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteEndosso" & HID.Value), [Lib].Negocio.Cliente))
            txtClienteEndosso.Text = itemClienteEndosso.Text
            txtCodigoClienteEndosso.Value = itemClienteEndosso.Value
            Session.Remove("objClienteEndosso" & HID.Value)
        End If
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlUnidadeNegocio.SelectedIndex > 0 Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
            Else
                ddlEmpresa.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuscaPedido.Click
        Try
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Or String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido.")
                Exit Sub
            End If

            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
            parameters("situacao") = 1
            Popup.ConsultaDePedidos(Me.Page, "objPedEndosso" & HID.Value)
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClienteEndosso_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClienteEndosso.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteEndosso" & HID.Value.ToString, "txtClienteEndosso")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkAllTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If gridConsultaTitulos.Rows.Count > 0 Then
                Dim chkAllTitulos As CheckBox = CType(sender, CheckBox)
                Dim passed As Boolean = False
                For Each rowgrid As GridViewRow In gridConsultaTitulos.Rows
                    Dim chkTitulo As CheckBox = CType(rowgrid.FindControl("chkGridTitulos"), CheckBox)

                    If Not chkTitulo.Enabled Then
                        chkAllTitulos.Checked = False
                        Exit Sub
                    End If
                Next
                TotalizadorTitulos()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkGridTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkTitulo As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkTitulo.NamingContainer, GridViewRow)

        TotalizadorTitulos()
    End Sub

    Protected Sub gridEndossos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlEmpresa.SelectedValue.ToString.Length = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
            Else
                objEndosso = New Endosso(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), gridEndossos.SelectedRow.Cells(1).Text())

                If objEndosso.Codigo > 0 Then

                    Limpar(False)

                    ddlUnidadeNegocio.Enabled = False
                    ddlEmpresa.Enabled = False

                    txtSequencia.Text = objEndosso.Codigo

                    'Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objEndosso.Cliente)
                    'txtCliente.Text = itemCliente.Text
                    'txtCodigoCliente.Value = itemCliente.Value

                    Dim itemClienteEndosso As ListItem = Funcoes.FormatarListItemCliente(objEndosso.ClienteEndosso)
                    txtClienteEndosso.Text = itemClienteEndosso.Text
                    txtCodigoClienteEndosso.Value = itemClienteEndosso.Value

                    btnCliente.Enabled = False
                    txtPedido.Enabled = False
                    btnBuscaPedido.Enabled = True
                    txtNumNota.Enabled = False
                    txtDataInicial.Enabled = False
                    txtDataFinal.Enabled = False
                    btnClienteEndosso.Enabled = False
                    txtNumEndosso.Enabled = False
                    txtVencimentoEndosso.Enabled = False
                    vlrEndosso.Enabled = False

                    txtNumEndosso.Text = objEndosso.NumeroEndosso
                    txtVencimentoEndosso.Text = objEndosso.Vencimento
                    vlrEndosso.Text = objEndosso.Valor
                    txtObservacao.Text = objEndosso.Observacoes
                    txtObservacaoInterna.Text = objEndosso.ObservacoesInterna


                    linhaGrid.Visible = True

                    Dim dtEndosso As New DataTable("Endosso")
                    dtEndosso.Columns.Add("Registro", Type.GetType("System.String"))
                    dtEndosso.Columns.Add("Vencimento", Type.GetType("System.String"))
                    dtEndosso.Columns.Add("Cliente", Type.GetType("System.String"))
                    dtEndosso.Columns.Add("Historico", Type.GetType("System.String"))
                    dtEndosso.Columns.Add("Pedido", Type.GetType("System.String"))
                    dtEndosso.Columns.Add("Grupado", Type.GetType("System.String"))
                    dtEndosso.Columns.Add("ValorLiquido", Type.GetType("System.Decimal"))

                    Dim Valor As Decimal = 0

                    For Each t In objEndosso.TitulosXEndosso
                        Dim drItem As DataRow = dtEndosso.NewRow()

                        drItem("Registro") = t.Titulo.Codigo
                        drItem("Vencimento") = t.Titulo.Prorrogacao.ToString("dd/MM/yyyy")
                        drItem("Cliente") = t.Titulo.Cliente.Nome
                        drItem("Historico") = t.Titulo.Historico
                        drItem("Pedido") = t.Titulo.CodigoPedido

                        drItem("Grupado") = Left(t.Titulo.Agrupado.ToString(), 1)

                        drItem("ValorLiquido") = t.Titulo.ValorLiquido

                        dtEndosso.Rows.Add(drItem)

                        Valor += t.Titulo.ValorDoDocumento

                    Next

                    rowValor.Visible = True
                    lblTotalRegistro.Text = "Título(s) selecionado(s) no valor total de: " & String.Format("{0:N2}", Valor)

                    If objEndosso.TitulosXEndosso.Count > 0 Then
                        gridConsultaTitulos.DataSource = dtEndosso
                        gridConsultaTitulos.DataBind()

                        CType(gridConsultaTitulos.HeaderRow.FindControl("chkAllTitulos"), CheckBox).Visible = False

                        For Each row As GridViewRow In gridConsultaTitulos.Rows
                            CType(row.FindControl("chkGridTitulos"), CheckBox).Visible = False
                        Next
                    End If

                    LimparBusca()

                    lnkConsultar.Parent.Visible = False

                    If objEndosso.CodigoSituacao = eSituacao.Normal Then lnkExcluir.Parent.Visible = True
                    If objEndosso.CodigoSituacao = eSituacao.Normal Then lnkFinalizar.Parent.Visible = True

                    TabEndossos.ActiveTabIndex = 0
                Else
                    MsgBox(Me.Page, "Endosso não foi encontrado.", eTitulo.Info)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Public Sub TotalizadorTitulos()
        Dim Valor As Decimal = 0
        For Each row As GridViewRow In gridConsultaTitulos.Rows
            Dim chkTitulo As CheckBox = CType(row.FindControl("chkGridTitulos"), CheckBox)
            If (chkTitulo.Checked) Then
                Valor = Valor + CDec(row.Cells(5).Text)
            End If
        Next

        If Valor > 0 Then
            rowValor.Visible = True
            lblTotalRegistro.Text = "Título(s) selecionado(s) no valor total de: " & String.Format("{0:N2}", Valor)
        Else
            rowValor.Visible = False
        End If
    End Sub

    Private Sub Limpar(ByVal limparConsulta As Boolean)
        lnkConsultar.Parent.Visible = True
        lnkNovo.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkFinalizar.Parent.Visible = False

        ddlUnidadeNegocio.Enabled = True
        ddlEmpresa.Enabled = True
        btnCliente.Enabled = True
        txtPedido.Enabled = True
        btnBuscaPedido.Enabled = True
        txtNumNota.Enabled = True
        txtDataInicial.Enabled = True
        txtDataFinal.Enabled = True
        btnClienteEndosso.Enabled = True
        txtNumEndosso.Enabled = True
        txtVencimentoEndosso.Enabled = True
        vlrEndosso.Enabled = True

        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        Session.Remove("objCliente" & HID.Value)

        txtPedido.Text = String.Empty
        Session.Remove("objPedEndosso" & HID.Value)

        txtSequencia.Text = String.Empty
        txtNumNota.Text = String.Empty

        txtClienteEndosso.Text = String.Empty
        txtCodigoClienteEndosso.Value = String.Empty
        Session.Remove("objClienteEndosso" & HID.Value)

        txtNumEndosso.Text = String.Empty
        txtVencimentoEndosso.Text = String.Empty
        vlrEndosso.Text = String.Empty
        txtObservacao.Text = String.Empty
        txtObservacaoInterna.Text = String.Empty

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidos.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)

        gridConsultaTitulos.DataSource = Nothing
        gridConsultaTitulos.DataBind()

        linhaGrid.Visible = False
        rowValor.Visible = False

        If limparConsulta Then LimparBusca()

        TabEndossos.ActiveTabIndex = 0

        verNumerador()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub LimparBusca()
        txtSequenciaConsulta.Text = String.Empty
        txtNumEndossoConsulta.Text = String.Empty
        txtVencimentoEndossoConsulta.Text = String.Empty

        gridEndossos.DataSource = Nothing
        gridEndossos.DataBind()
    End Sub


    Private Sub BuscarEndossos()
        If txtSequenciaConsulta.Text.Length > 0 Then
            Sql = "Codigo_id = " & txtSequenciaConsulta.Text
        ElseIf txtNumEndossoConsulta.Text.Length > 0 Then
            Sql = "NumeroEndosso = " & txtNumEndossoConsulta.Text
        ElseIf txtVencimentoEndossoConsulta.Text.Length > 0 Then
            Sql = "Vencimento = '" & CDate(txtVencimentoEndossoConsulta.Text).ToString("yyyy/MM/dd") & "'"
        Else
            Sql = "Movimento Between '" & CDate(txtBuscaDataInicial.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtBuscaDataFinal.Text).ToString("yyyy/MM/dd") & "'"
        End If

        If rbNormal.Checked Then
            Sql &= " AND Situacao = 1" & vbCrLf
        ElseIf rbFinalizado.Checked Then
            Sql &= " AND Situacao = 2" & vbCrLf
        ElseIf rbExcluido.Checked Then
            Sql &= " AND Situacao = 3" & vbCrLf
        End If

        Dim endossos As New ListEndosso(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), Sql)

        If endossos.Count = 0 Then
            MsgBox(Me.Page, "Nenhum registro econtrado.")
        Else
            For Each eD In endossos
                'If eD.Cliente.CodigoFormatado.Length = 0 Then Dim erro = "erro - só para instância da propriedade ClienteDescricao READ ONLY"

                If eD.ClienteEndosso.CodigoFormatado.Length = 0 Then Dim erro = "erro - só para instância da propriedade ClienteEndossoDescricao READ ONLY"
            Next

            gridEndossos.DataSource = endossos.ToArray()
            gridEndossos.DataBind()

            Dim i As Integer = 0
            While i < gridEndossos.Rows.Count

                For Each e In endossos
                    If gridEndossos.Rows(i).Cells(1).Text = e.Codigo Then
                        If e.CodigoSituacao = eSituacao.Normal Then
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).ImageUrl = "~/images/certo.jpg"
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).ToolTip = "Normal"
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"
                        ElseIf e.CodigoSituacao = eSituacao.Cancelado Then
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).ImageUrl = "~/images/Cancelar.jpg"
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).ToolTip = "Finalizado"
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"
                        Else
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).ImageUrl = "~/images/Cancelar.jpg"
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).ToolTip = "Cancelado"
                            CType(gridEndossos.Rows(i).FindControl("imgSituacao"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"
                        End If
                    End If
                Next

                i += 1
            End While
        End If
    End Sub


    Private Sub gerarExcel()
        Try
            If ddlEmpresa.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
                Exit Sub
            End If

            Dim objBanco As New AcessaBanco()

            Sql = "SELECT e.Codigo_Id AS Codigo, cR.Cliente, cR.EndCliente AS EndCli, c.Nome AS NomeCliente, e.Movimento," & vbCrLf &
                    "		e.ClienteEndosso_Id AS ClienteEndosso, e.EndClienteEndosso_Id AS EndCliEndosso, cE.Nome AS NomeClienteEndosso, e.NumeroEndosso, e.Vencimento, e.Valor, isnull(eXt.Titulo_Id,0) AS Titulo, isnull(cR.ValorDoDocumento,0) AS ValorDoTitulo" & vbCrLf &
                    "FROM Endosso e" & vbCrLf &
                    "	INNER JOIN Clientes cE" & vbCrLf &
                    "			ON cE.Cliente_Id     = e.ClienteEndosso_Id" & vbCrLf &
                    "			AND cE.Endereco_Id   = e.EndClienteEndosso_Id" & vbCrLf &
                    "	LEFT JOIN EndossoXTitulo eXt" & vbCrLf &
                    "			ON eXt.Empresa_Id     = e.Empresa_Id" & vbCrLf &
                    "			AND eXt.EndEmpresa_Id = e.EndEmpresa_Id" & vbCrLf &
                    "			AND eXt.Codigo_Id     = e.Codigo_Id" & vbCrLf &
                    "	LEFT JOIN ContasAReceber cR" & vbCrLf &
                    "			ON cR.Empresa        = eXt.Empresa_Id" & vbCrLf &
                    "			AND cR.EndEmpresa    = eXt.EndEmpresa_Id" & vbCrLf &
                    "			AND cR.Registro_Id   = eXt.Titulo_Id" & vbCrLf &
                    "	INNER JOIN Clientes c" & vbCrLf &
                    "			ON c.Cliente_Id     = cR.Cliente" & vbCrLf &
                    "			AND c.Endereco_Id   = cR.EndCliente" & vbCrLf &
                    "WHERE e.Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf &
                    "  AND e.EndEmpresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(1) & "' " & vbCrLf

            If txtSequenciaConsulta.Text.Length > 0 Then
                Sql &= "  AND e.Codigo_id = " & txtSequenciaConsulta.Text & vbCrLf
            ElseIf txtNumEndossoConsulta.Text.Length > 0 Then
                Sql &= "  AND e.NumeroEndosso = '" & txtNumEndossoConsulta.Text & vbCrLf & "'" & vbCrLf
            ElseIf txtVencimentoEndossoConsulta.Text.Length > 0 Then
                Sql &= "  AND e.Vencimento = '" & CDate(txtVencimentoEndossoConsulta.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
            Else
                Sql &= "  AND e.Movimento Between '" & CDate(txtBuscaDataInicial.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtBuscaDataFinal.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf
            End If

            If rbNormal.Checked Then
                Sql &= "  AND e.Situacao = 1" & vbCrLf
            ElseIf rbFinalizado.Checked Then
                Sql &= "  AND e.Situacao = 2" & vbCrLf
            ElseIf rbExcluido.Checked Then
                Sql &= "  AND e.Situacao = 3" & vbCrLf
            End If

            Dim ds As DataSet = objBanco.ConsultaDataSet(Sql, "Endosso")

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Sem registros para essa seleção.", eTitulo.Info)
                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Dim objEmpresa As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("ENDOSSO DE TÍTULOS")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                    If rbNormal.Checked Then
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "ENDOSSO DE TÍTULOS - NORMAL")
                    ElseIf rbFinalizado.Checked Then
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "ENDOSSO DE TÍTULOS - FINALIZADO")
                    ElseIf rbExcluido.Checked Then
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "ENDOSSO DE TÍTULOS - EXCLUÍDO")
                    End If

                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtBuscaDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtBuscaDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1


                    Dim primeiro As Boolean = True
                    Dim codigoEndosso As Integer = 0
                    Dim cliente As String = String.Empty
                    Dim valorTotal As Decimal

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1

                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        ''formatando células datas
                        worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        ''formatando células numéricas
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'worksheet.Cells(rowIndex, 12).Value = "TOTAL"
                    'worksheet.Cells(rowIndex, 13).Value = valorTotal
                    'worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"


                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Formula = String.Format("=SUM(M6:M{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Private Sub verNumerador()
        Sql = "Select Sequencia" & vbCrLf &
                "from Numerador" & vbCrLf &
                "where Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                "  and EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                "  and Numerador_Id  = 200"

        ds = Banco.ConsultaDataSet(Sql, "Numerador")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            lnkConsultar.Parent.Visible = False
            MsgBox(Me.Page, "Numerador não encontrado, entre em Contato com o Suporte.")
        Else
            txtSequencia.Text = (ds.Tables(0).Rows(0)("Sequencia") + 1).ToString

            lnkConsultar.Parent.Visible = True
        End If

    End Sub

    Private Sub TitulosConsulta()
        Sql = "  SELECT CP.Registro_Id AS Registro, convert(varchar(10),CP.Prorrogacao,103) as Vencimento, " & vbCrLf &
              "         Cli.Nome AS Cliente, Historico, CP.ValorLiquido AS ValorLiquido, " & vbCrLf &
              "         CP.Pedido as Pedido, " & vbCrLf &
              "         isnull(CP.Grupado,'N') as Grupado" & vbCrLf &
              "    FROM ContasAReceber CP " & vbCrLf &
              "    LEFT JOIN NotaFiscalXTitulo NFXT" & vbCrLf &
              "      ON CP.Empresa     = NFXT.Empresa_Id" & vbCrLf &
              "     AND CP.EndEmpresa  = NFXT.EndEmpresa_Id" & vbCrLf &
              "     AND CP.Registro_Id = NFXT.Titulo_Id" & vbCrLf &
              "    LEFT JOIN FaturaDeFreteXTitulo FFxT" & vbCrLf &
              "      ON CP.Registro_Id = FFxT.Titulo_Id" & vbCrLf &
              "    LEFT JOIN FaturasDeFretesXItens FFxI" & vbCrLf &
              "      ON FFxT.Empresa_Id       = FFxI.EmpresaPagadora_Id " & vbCrLf &
              "     AND FFxT.EndEmpresa_Id    = FFxI.EndEmpresaPagadora_Id" & vbCrLf &
              "     AND FFxT.Conveniado_Id    = FFxI.Conveniado_Id" & vbCrLf &
              "     AND FFxT.EndConveniado_Id = FFxI.EndConveniado_Id" & vbCrLf &
              "     AND FFxT.Fatura_Id        = FFxI.Fatura_Id" & vbCrLf &
              "   INNER JOIN Clientes Cli" & vbCrLf &
              "      ON CP.Cliente = Cli.Cliente_Id " & vbCrLf &
              "     AND CP.EndCliente = Cli.Endereco_Id" & vbCrLf &
              "   WHERE 1=1 " & vbCrLf &
              "     AND CP.Situacao in(1) " & vbCrLf &
              "     AND Provisao = 2 " & vbCrLf &
              "     AND Grupado in('N','M') " & vbCrLf

        Sql &= " AND CP.Empresa = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf  'Empresa
        Sql &= " AND CP.EndEmpresa = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf      'Endereco da Empresa

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            Sql &= " AND CP.Cliente = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf  'Cliente
            Sql &= " AND CP.EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf    'Cliente da Empresa
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            Sql &= " AND CP.Pedido = '" & txtPedido.Text & "'" & vbCrLf
        ElseIf Not String.IsNullOrWhiteSpace(txtNumNota.Text) Then
            Sql &= " AND (ISNULL(NFXT.Nota_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Fatura_Id,0) in(" & txtNumNota.Text & ") OR ISNULL(FFxI.Nota_Id,0) in(" & txtNumNota.Text & "))"
        ElseIf txtDataInicial.Text <> "" And txtDataFinal.Text <> "" Then
            Sql &= " AND Prorrogacao between '" & CDate(txtDataInicial.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy/MM/dd") & "'"
        End If

        Sql &= " ORDER BY CP.Prorrogacao, Cli.Nome"

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            linhaGrid.Visible = False
            gridConsultaTitulos.DataSource = Nothing
            gridConsultaTitulos.DataBind()
            MsgBox(Me.Page, "Nenhum Registro encontrado...", eTitulo.Info)
        Else
            linhaGrid.Visible = True

            gridConsultaTitulos.DataSource = ds
            gridConsultaTitulos.DataBind()

            lnkConsultar.Parent.Visible = False
            lnkNovo.Parent.Visible = True
        End If
    End Sub


    Function ValidarRegistro() As Boolean

        Dim libera As Boolean = False

        If ddlEmpresa.SelectedValue.Length = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
            'ElseIf txtCodigoCliente.Value.ToString.Length = 0 Then
            '    MsgBox(Me.Page, "Cliente não foi selecionado.", eTitulo.Info)
        ElseIf txtCodigoClienteEndosso.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Cliente do Endosso não foi selecionado.", eTitulo.Info)
        ElseIf txtNumEndosso.Text.Length = 0 Then
            MsgBox(Me.Page, "Numero do Endosso não foi informado.", eTitulo.Info)
        ElseIf txtVencimentoEndosso.Text.Length = 0 Then
            MsgBox(Me.Page, "Vencimento não foi informado.", eTitulo.Info)
        ElseIf CDate(txtVencimentoEndosso.Text) < Today() Then
            MsgBox(Me.Page, "Vencimento não pode ser menor que a data atual.", eTitulo.Info)
        ElseIf gridConsultaTitulos.Rows.Count = 0 Then
            MsgBox(Me.Page, "Lista de está vazia.", eTitulo.Info)
        Else
            For Each row As GridViewRow In gridConsultaTitulos.Rows
                Dim chkTitulo As CheckBox = CType(row.FindControl("chkGridTitulos"), CheckBox)
                If (chkTitulo.Checked) Then libera = True
            Next
        End If

        Return libera

    End Function


    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Endossos", "GRAVAR") Then
                If ValidarRegistro() Then
                    objEndosso = New Endosso()

                    objEndosso.IUD = "I"

                    objEndosso.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    objEndosso.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                    objEndosso.Codigo = CInt(txtSequencia.Text)
                    'Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
                    'objEndosso.CodigoCliente = strCliente(0)
                    'objEndosso.EnderecoCliente = strCliente(1)

                    objEndosso.Movimento = Today.ToString("yyyy/MM/dd")

                    Dim strClienteEndosso As String() = txtCodigoClienteEndosso.Value.Split("-")
                    objEndosso.CodigoClienteEndosso = strClienteEndosso(0)
                    objEndosso.EnderecoClienteEndosso = strClienteEndosso(1)

                    objEndosso.NumeroEndosso = Trim(txtNumEndosso.Text)

                    objEndosso.Vencimento = CDate(txtVencimentoEndosso.Text)
                    objEndosso.Valor = CDec(vlrEndosso.Text)
                    objEndosso.Observacoes = Trim(txtObservacao.Text)
                    objEndosso.CodigoSituacao = eSituacao.Normal

                    objEndosso.ObservacoesInterna = "Incluído em " & Now.ToString("dd/MM/yyyy HH:mm:ss") & " por " & HttpContext.Current.Session("ssNomeUsuario")

                    For Each row As GridViewRow In gridConsultaTitulos.Rows
                        Dim chkTitulo As CheckBox = CType(row.FindControl("chkGridTitulos"), CheckBox)
                        If (chkTitulo.Checked) Then
                            Dim titEndosso = New EndossoXTitulo(objEndosso)
                            titEndosso.CodigoTitulo = CInt(gridConsultaTitulos.Rows(row.RowIndex).Cells(1).Text)
                            titEndosso.IUD = "I"
                            objEndosso.TitulosXEndosso.Add(titEndosso)
                        End If
                    Next

                    If objEndosso.Salvar Then
                        MsgBox(Me.Page, "Endosso " & objEndosso.Codigo & " incluído com Sucesso.", eTitulo.Sucess)

                        Limpar(True)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkFinalizar_Click(sender As Object, e As EventArgs) Handles lnkFinalizar.Click
        Try
            If Funcoes.VerificaPermissao("Endossos", "ALTERAR") Then

                objEndosso = New Endosso(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), txtSequencia.Text)

                If objEndosso.Codigo > 0 Then

                    objEndosso.IUD = "F"

                    objEndosso.Observacoes = Trim(txtObservacao.Text)
                    objEndosso.CodigoSituacao = eSituacao.Cancelado

                    Dim obs As String = objEndosso.ObservacoesInterna
                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Endosso finalizado em " & Now.ToString("dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    Else
                        obs = "Endosso finalizado em " & Now.ToString("dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    End If

                    objEndosso.ObservacoesInterna = obs

                    If objEndosso.Salvar Then
                        MsgBox(Me.Page, "Endosso " & objEndosso.Codigo & " finalizado com Sucesso.", eTitulo.Sucess)

                        Limpar(True)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                    End If
                Else
                    MsgBox(Me.Page, "Endosso não foi encontrado.", eTitulo.Info)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Endossos", "EXCLUIR") Then

                objEndosso = New Endosso(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), txtSequencia.Text)

                If objEndosso.Codigo > 0 Then

                    objEndosso.IUD = "D"

                    objEndosso.Observacoes = Trim(txtObservacao.Text)
                    objEndosso.CodigoSituacao = eSituacao.Excluido

                    Dim obs As String = objEndosso.ObservacoesInterna
                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Endosso excluído em " & Now.ToString("dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    Else
                        obs = "Endosso excluído em " & Now.ToString("dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    End If

                    objEndosso.ObservacoesInterna = obs

                    If objEndosso.Salvar Then
                        MsgBox(Me.Page, "Endosso " & objEndosso.Codigo & " finalizado com Sucesso.", eTitulo.Sucess)

                        Limpar(True)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                    End If
                Else
                    MsgBox(Me.Page, "Endosso não foi encontrado.", eTitulo.Info)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If ddlEmpresa.SelectedValue.ToString.Length = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
                'ElseIf txtCodigoCliente.Value.Length = 0 Then
                '    MsgBox(Me.Page, "Cliente não foi selecionado.", eTitulo.Info)
            Else
                TitulosConsulta()
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


    Protected Sub lnkBusca_Click(sender As Object, e As EventArgs) Handles lnkBusca.Click
        Try
            If ddlEmpresa.SelectedValue.ToString.Length = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
            Else
                BuscarEndossos()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            gerarExcel()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkLimparBusca_Click(sender As Object, e As EventArgs) Handles lnkLimparBusca.Click
        Try
            LimparBusca()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class