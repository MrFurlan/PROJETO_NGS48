Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class DuplicataNotaFiscal
    Inherits BasePage

    Private mensagemErro As String
    Private empresa() As String
    Private cliente() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("DuplicataNotaFiscal", "ACESSAR") Then
                BuncarUnidadeDeNegocio()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = Now.ToString("dd/MM/yyyy")
        txtDataFinal.Text = Now.ToString("dd/MM/yyyy")
        txtNotaFiscal.Text = ""
        gridNotaFiscal.DataSource = Nothing
        gridNotaFiscal.DataBind()
        gridVencimentos.DataSource = Nothing
        gridVencimentos.DataBind()
        Session.Remove("objNotaFiscalDPL")
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Function ValidarSelecao() As Boolean
        If ddlUnidadeNegocio.SelectedIndex = 0 Then
            mensagemErro = "Unidade de Negócio não foi selecionada."
            Return False
        ElseIf ddlEmpresa.SelectedIndex = 0 Then
            mensagemErro = "Empresa não foi selecionada."
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteDupNF" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Session("objClienteDupNF" & HID.Value) IsNot Nothing Then
                Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteDupNF" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
                txtClientes.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteDupNF" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkNotaFiscal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim chkNotaFiscal As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(chkNotaFiscal.NamingContainer, GridViewRow)
            hidNotaFiscal.Value = row.RowIndex

            gridVencimentos.DataSource = Nothing
            gridVencimentos.DataBind()

            If chkNotaFiscal.Checked Then
                Dim i As Integer

                For i = 0 To gridNotaFiscal.Rows.Count - 1
                    If Not row.RowIndex = i Then
                        CType(gridNotaFiscal.Rows(i).FindControl("chkNotaFiscal"), CheckBox).Checked = False
                    End If
                Next

                empresa = gridNotaFiscal.Rows(row.RowIndex).Cells(1).Text().Split("-")
                cliente = gridNotaFiscal.Rows(row.RowIndex).Cells(2).Text().Split("-")

                Dim NfConsulta As New [Lib].Negocio.NotaFiscal
                NfConsulta.CodigoEmpresa = empresa(0)
                NfConsulta.EnderecoEmpresa = empresa(1)
                NfConsulta.CodigoCliente = cliente(0)
                NfConsulta.EnderecoCliente = cliente(1)
                NfConsulta.EntradaSaida = IIf(gridNotaFiscal.Rows(row.RowIndex).Cells(4).Text() = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
                NfConsulta.Serie = gridNotaFiscal.Rows(row.RowIndex).Cells(5).Text()
                NfConsulta.Codigo = gridNotaFiscal.Rows(row.RowIndex).Cells(6).Text()

                NfConsulta = New [Lib].Negocio.NotaFiscal(NfConsulta)

                Session("objNotaFiscalDP") = NfConsulta

                If NfConsulta.CodigoPedido > 0 Then
                    If NfConsulta.VencimentosNota.Count > 0 Then
                        gridVencimentos.DataSource = NfConsulta.VencimentosNota.ToArray
                        gridVencimentos.DataBind()

                        Dim j As Integer
                        For j = 0 To gridVencimentos.Rows.Count - 1
                            CType(gridVencimentos.Rows(j).FindControl("chkVencimento"), CheckBox).Checked = True
                        Next

                        TabContainer1.ActiveTabIndex = 2
                    Else
                        MsgBox(Me.Page, "Vencimento(s) não encontrado(s).")
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgConfirmar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgConfirmar.Click
        Try
            Dim temVcto As Boolean = False
            For j = 0 To gridVencimentos.Rows.Count - 1
                If CType(gridVencimentos.Rows(j).FindControl("chkVencimento"), CheckBox).Checked = True Then
                    temVcto = True
                End If
            Next

            If temVcto Then
                Dim ds As New DataSet
                Dim tbDuplicata As New DataTable("Cabecalho")
                tbDuplicata.Columns.Add("Empresa", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Cidade", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Endereco", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Bairro", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Fone", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Cep", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Uf", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Cnpj", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Insc", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Valtot", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Numnot", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Numord", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Seqord", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Vencto", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Cliente", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("EndCli", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("CnpjCli", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Praca", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("UfPraca", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Ext001", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Ext002", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Ext003", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Valtot1", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Inscli", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Via", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Tipo", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Valnot", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Emissa", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("BaiCli", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Cepcli", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Cidcli", Type.GetType("System.String"))
                tbDuplicata.Columns.Add("Ufcli", Type.GetType("System.String"))
                ds.Tables.Add(tbDuplicata)

                Dim tbLogotipo As New DataTable("Logotipo")
                tbLogotipo.Columns.Add("path", GetType(String))
                tbLogotipo.Columns.Add("Imagem", GetType(System.Byte()))
                Dim drImagem As DataRow = tbLogotipo.NewRow()
                Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
                drImagem("path") = strCaminhoImagem
                drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                tbLogotipo.Rows.Add(drImagem)
                ds.Tables.Add(tbLogotipo)

                For j = 0 To gridVencimentos.Rows.Count - 1
                    If CType(gridVencimentos.Rows(j).FindControl("chkVencimento"), CheckBox).Checked = True Then
                        For numVia = 1 To 2
                            Dim dr As DataRow = ds.Tables(0).NewRow()
                            dr("Empresa") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.Nome
                            dr("Cidade") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.Cidade
                            dr("Endereco") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.Endereco
                            dr("Bairro") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.Bairro
                            dr("Fone") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.Telefone
                            dr("Cep") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.CEP
                            dr("UF") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.CodigoEstado
                            dr("Cnpj") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.CodigoFormatado
                            dr("Insc") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Empresa.InscricaoEstadual
                            dr("ValTot1") = gridVencimentos.Rows(j).Cells(3).Text()
                            dr("NumNot") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Codigo
                            dr("NumOrd") = j + 1
                            dr("SeqOrd") = j + 1
                            dr("Vencto") = gridVencimentos.Rows(j).Cells(2).Text()
                            dr("Cliente") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Nome
                            dr("EndCli") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Endereco
                            dr("CepCli") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.CEP
                            dr("CidCli") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Cidade
                            dr("UFCli") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.CodigoEstado
                            dr("CnpjCli") = Funcoes.FormatarCpfCnpj(CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Codigo)
                            dr("Praca") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Endereco & IIf(CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Numero.ToString.Length > 0, "," & CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Numero, "") & " - " & CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Cidade & " - " & CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.CodigoEstado
                            dr("Ext001") = Funcoes.Extenso(gridVencimentos.Rows(j).Cells(3).Text(), "Real", "Reais").ToUpper

                            'If Not CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).VencimentosNota Is Nothing AndAlso CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).VencimentosNota.Count > 0 Then
                            '    dr("Ext002") = IIf(gridNotaFiscal.Rows(hidNotaFiscal.Value).Cells(4).Text() = "E", "Titulo a Pagar:", "Titulo a Receber:") & " " & gridNotaFiscal.Rows(hidNotaFiscal.Value).Cells(6).Text()
                            'Else
                            '    dr("Ext002") = IIf(CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).VencimentosNota(j).ReceberPagar = "P", "Titulo a Pagar:", "Titulo a Receber:") & " " & CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).VencimentosNota(j).Codigo
                            'End If

                            dr("Ext002") = IIf(gridNotaFiscal.Rows(hidNotaFiscal.Value).Cells(4).Text() = "E", "Titulo a Pagar:", "Titulo a Receber:") & " " & gridVencimentos.Rows(j).Cells(1).Text()

                            dr("Ext003") = "" 'drAux("Ext003")
                            dr("Inscli") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.InscricaoEstadual
                            dr("Via") = numVia
                            dr("Tipo") = "" 'drAux("Tipo")
                            dr("Valnot") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).TotalNota
                            dr("Emissa") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Movimento.ToString("dd/MM/yyyy")
                            dr("BaiCli") = CType(Session("objNotaFiscalDP"), [Lib].Negocio.NotaFiscal).Cliente.Bairro

                            ds.Tables(0).Rows.Add(dr)
                        Next
                    End If
                Next

                Relatorio(ds)
            Else
                MsgBox(Me.Page, "Vencimento(s) não selecionado(s)")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Relatorio(ByVal ds As DataSet)
        Dim crptRelatorio As New ReportDocument()
        crptRelatorio.FileName = Server.MapPath("~/Reports/Cr_Duplicata.rpt")
        crptRelatorio.Load(crptRelatorio.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
        Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
        Dim arquivo As String = NomeArquivo

        crptRelatorio.SetDataSource(ds)

        Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions

        crParameterFieldDefinitions = crptRelatorio.DataDefinition.ParameterFields()

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)
        Try
            crptRelatorio.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString)
        Finally
            crptRelatorio.Close()
            crptRelatorio.Dispose()
            If IO.File.Exists(arquivo) Then
                Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
            End If
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If ValidarSelecao() Then
                gridNotaFiscal.DataSource = Nothing
                gridNotaFiscal.DataBind()
                gridVencimentos.DataSource = Nothing
                gridVencimentos.DataBind()
                Session.Remove("objNotaFiscalDPL")

                empresa = ddlEmpresa.SelectedValue.ToString.Split("-")
                If txtCodigoCliente.Value.ToString.Length = 0 Then
                    txtCodigoCliente.Value = "" & "-" & 0
                End If
                cliente = txtCodigoCliente.Value.ToString.Split("-")
                Dim listNotasFiscais As New [Lib].Negocio.ListNotasFiscais(empresa(0), empresa(1), txtDataInicial.Text, txtDataFinal.Text, cliente(0), cliente(1), "S")

                If listNotasFiscais.Count > 0 Then
                    Dim ds As New DataSet
                    Dim tbNotas As New DataTable("Notas")
                    tbNotas.Columns.Add("Empresa", Type.GetType("System.String"))
                    tbNotas.Columns.Add("Cliente", Type.GetType("System.String"))
                    tbNotas.Columns.Add("Nome", Type.GetType("System.String"))
                    tbNotas.Columns.Add("E/S", Type.GetType("System.String"))
                    tbNotas.Columns.Add("Serie", Type.GetType("System.String"))
                    tbNotas.Columns.Add("Nota", Type.GetType("System.String"))
                    ds.Tables.Add(tbNotas)

                    For Each nf As [Lib].Negocio.NotaFiscal In listNotasFiscais
                        If txtNotaFiscal.Text.Length > 0 Then
                            If nf.Codigo = CInt(txtNotaFiscal.Text) Then
                                Dim drRow As DataRow = ds.Tables(0).NewRow()
                                drRow("Empresa") = nf.CodigoEmpresa & "-" & nf.EnderecoEmpresa
                                drRow("Cliente") = nf.CodigoCliente & "-" & nf.EnderecoCliente
                                drRow("Nome") = nf.Cliente.Nome
                                drRow("E/S") = IIf(nf.EntradaSaida = eEntradaSaida.Saida, "S", "E")
                                drRow("Serie") = nf.Serie
                                drRow("Nota") = nf.Codigo

                                ds.Tables(0).Rows.Add(drRow)
                            End If
                        Else
                            Dim drRow As DataRow = ds.Tables(0).NewRow()
                            drRow("Empresa") = nf.CodigoEmpresa & "-" & nf.EnderecoEmpresa
                            drRow("Cliente") = nf.CodigoCliente & "-" & nf.EnderecoCliente
                            drRow("Nome") = nf.Cliente.Nome
                            drRow("E/S") = IIf(nf.EntradaSaida = eEntradaSaida.Saida, "S", "E")
                            drRow("Serie") = nf.Serie
                            drRow("Nota") = nf.Codigo

                            ds.Tables(0).Rows.Add(drRow)
                        End If
                    Next

                    gridNotaFiscal.DataSource = ds
                    gridNotaFiscal.DataBind()

                    TabContainer1.ActiveTabIndex = 1
                Else
                    MsgBox(Me.Page, "Registro(s) não encontrado(s) para a seleção.")
                End If
            Else
                MsgBox(Me.Page, mensagemErro)
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
            Funcoes.Ajuda(Me.Page, "DuplicataNotaFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class