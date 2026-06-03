Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class EmissaoDeDuplicatas
    Inherits BasePage

    Private objPedido As [Lib].Negocio.Pedido
    Private strJavaScript As String
    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("EmissaoDeDuplicatas", "ACESSAR") Then
                BuscaEmpresa()
                CarregarSafras()
                Limpar()

                Sql = "Select Descricao FROM RotinasDeImpressao WHERE Rotina_Id = 'FINANCEIRO'"
                Dim dsImpressora As New DataSet
                dsImpressora = Banco.ConsultaDataSet(Sql, "Impressora")
                HttpContext.Current.Session("printerFinanceiro") = ""

                If dsImpressora Is Nothing Then
                ElseIf dsImpressora.Tables(0).Rows.Count = 0 Then
                Else
                    HttpContext.Current.Session("printerFinanceiro") = dsImpressora.Tables(0).Rows(0).Item("Descricao")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteDpRev" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteDpRev" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteDpRev" & HID.Value)
        ElseIf Not Session("objPedidoDpRev" & HID.Value) Is Nothing Then
            objPedido = CType(Session("objPedidoDpRev" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = objPedido.Codigo
            If objPedido.Vencimentos.Count > 0 Then
                GridCondicoes.DataSource = objPedido.Vencimentos.ToDataTablePedidos()
                GridCondicoes.DataBind()
                lnkRelatorio.Enabled = True
            End If
            SalvarVariavelPedido()
            Session.Remove("objPedidoDpRev" & HID.Value)
        End If
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CarregarSafras()
        ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
    End Sub

    Private Sub CarregarVariavelPedido()
        objPedido = CType(Session("objPedido"), [Lib].Negocio.Pedido)
    End Sub

    Private Sub SalvarVariavelPedido()
        Session("objPedido") = objPedido
    End Sub

    Private Sub Limpar()
        lnkRelatorio.Enabled = False
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        ddlSafra.SelectedIndex = 0
        txtPedido.Text = ""
        GridCondicoes.DataSource = Nothing
        GridCondicoes.DataBind()
        Funcoes.VerificaEmpresa(ddlEmpresa)
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub impressaoMatricial()
        CarregarVariavelPedido()

        Dim i As Integer
        Dim j As Integer
        Dim extenso As String
        'Dim f As System.IO.File
        Dim Arquivo As System.IO.StreamWriter

        'Arquivo = System.IO.File.CreateText("Destinatario.txt")
        Arquivo = System.IO.File.CreateText(Server.MapPath("~/Files/Destinatario.txt"))

        Try

            If objPedido.Vencimentos.Count > 0 Then
                Dim total As Decimal = 0
                For i = 0 To objPedido.Vencimentos.Count - 1
                    total += objPedido.Vencimentos(i).ValorLiquidoOficial
                Next

                ' Envie seu texto para o arquivo
                For i = 0 To objPedido.Vencimentos.Count - 1
                    Arquivo.WriteLine(Chr(15) & Space(102) & Funcoes.FormatarCpfCnpj(objPedido.CodigoEmpresa) & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(102) & objPedido.Empresa.InscricaoEstadual & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(42) & objPedido.Empresa.Endereco & "," & objPedido.Empresa.Numero & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(42) & objPedido.Empresa.Bairro & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(42) & objPedido.Empresa.Cidade & "/" & objPedido.Empresa.CodigoEstado & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(42) & objPedido.Empresa.CEP & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(42) & "(" & Left(objPedido.Empresa.Telefone, 2) & ")" & Mid(objPedido.Empresa.Telefone, 2, 4) & "-" & Mid(objPedido.Empresa.Telefone, 6, 4) & Space(38) & Format(Now.Date, "dd/MM/yyyy") & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(20) & "PED " & IIf(Trim(objPedido.PedidoEfetivo).Length > 0, objPedido.PedidoEfetivo, objPedido.Codigo) & Space(13) & Format(total, "#,###,###,##0.00").ToString() & Space(15) & i + 1 & Space(13) & Format(objPedido.Vencimentos(i).ValorLiquidoOficial, "#,###,###,##0.00").ToString() & Space(13) & Format(objPedido.Vencimentos(i).DataProrrogacao, "dd/MM/yyyy") & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(43) & objPedido.Cliente.Nome & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(43) & IIf(objPedido.Cliente.Numero.ToString.Length > 0, objPedido.Cliente.Endereco & ", " & objPedido.Cliente.Numero, objPedido.Cliente.Endereco) & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(43) & objPedido.Cliente.Cidade & Space(33) & Left(objPedido.Cliente.CEP, 5) & "-" & Mid(objPedido.Cliente.CEP, 5, 3) & Space(27) & objPedido.Cliente.CodigoEstado & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(43) & objPedido.Cliente.Cidade & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(43) & Funcoes.FormatarCpfCnpj(objPedido.CodigoCliente) & Space(40) & objPedido.Cliente.InscricaoEstadual & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(43) & Chr(18))
                    extenso = UCase(Funcoes.Extenso(objPedido.Vencimentos(i).ValorLiquidoOficial.ToString("N2"), "Real", "Reais"))
                    For j = 1 To (170 - Len(extenso))
                        extenso &= "*"
                    Next
                    Arquivo.WriteLine(Chr(15) & Space(43) & Left(extenso, 85) & Chr(18))
                    Arquivo.WriteLine(Chr(15) & Space(43) & Mid(extenso, 85, 85) & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                    Arquivo.WriteLine(Chr(15) & "" & Chr(18))
                Next
            End If

            'Feche o arquivo
            Arquivo.Close()

            'copie o arquivo para a porta da impressora
            'f.Copy("Destinatario.txt", "\\REVENDA\Fx890")
            'f.Copy("Destinatario.txt", HttpContext.Current.Session("printerFinanceiro"))
            System.IO.File.Copy(Server.MapPath("~/Files/Destinatario.txt"), HttpContext.Current.Session("printerFinanceiro"))

            MsgBox(Me.Page, "Impressão concluída.")
        Catch ex As Exception
            Arquivo.Close()
            MsgBox(Me.Page, ex.Message.ToString())
        End Try
    End Sub

    Private Sub impressaoGrafica()
        CarregarVariavelPedido()

        Dim j As Integer
        Dim h As Integer

        If objPedido.Vencimentos.Count > 0 Then
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
            'Dim tbLogotipo As DataTable = ds.Tables.Add("Logotipo")
            tbLogotipo.Columns.Add("path", GetType(String))
            tbLogotipo.Columns.Add("Imagem", GetType(System.Byte()))
            Dim drImagem As DataRow = tbLogotipo.NewRow()
            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
            drImagem("path") = strCaminhoImagem
            drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
            tbLogotipo.Rows.Add(drImagem)
            ds.Tables.Add(tbLogotipo)

            For j = 0 To objPedido.Vencimentos.Count - 1
                For h = 1 To 2
                    Dim dr As DataRow = ds.Tables(0).NewRow()
                    dr("Empresa") = objPedido.Empresa.Nome
                    dr("Cidade") = objPedido.Empresa.Cidade
                    dr("Endereco") = objPedido.Empresa.Endereco
                    dr("Bairro") = objPedido.Empresa.Bairro
                    dr("Fone") = objPedido.Empresa.Telefone
                    dr("Cep") = objPedido.Empresa.CEP
                    dr("UF") = objPedido.Empresa.CodigoEstado
                    dr("Cnpj") = objPedido.Empresa.CodigoFormatado
                    dr("Insc") = objPedido.Empresa.InscricaoEstadual
                    dr("ValTot1") = objPedido.Vencimentos(j).ValorLiquidoOficial.ToString("N2")
                    dr("NumNot") = IIf(objPedido.PedidoEfetivo.Length > 0, "PED " & objPedido.PedidoEfetivo, "PED " & objPedido.Codigo)
                    dr("NumOrd") = j + 1
                    dr("SeqOrd") = j + 1
                    dr("Vencto") = objPedido.Vencimentos(j).DataProrrogacao.ToString("dd/MM/yyyy")
                    dr("Cliente") = objPedido.Cliente.Nome
                    dr("EndCli") = objPedido.Cliente.Endereco
                    dr("CepCli") = objPedido.Cliente.CEP
                    dr("CidCli") = objPedido.Cliente.Cidade
                    dr("UFCli") = objPedido.Cliente.CodigoEstado
                    dr("CnpjCli") = Funcoes.FormatarCpfCnpj(objPedido.Cliente.Codigo)
                    dr("Praca") = objPedido.Cliente.Endereco & IIf(objPedido.Cliente.Numero.ToString.Length > 0, "," & objPedido.Cliente.Numero, "") & " - " & objPedido.Cliente.Cidade & " - " & objPedido.Cliente.CodigoEstado
                    dr("Ext001") = Funcoes.Extenso(objPedido.Vencimentos(j).ValorLiquidoOficial, "Real", "Reais").ToUpper

                    dr("Ext002") = IIf(objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "Titulo a Pagar:", "Titulo a Receber:") & " " & IIf(objPedido.PedidoEfetivo.Length > 0, objPedido.PedidoEfetivo, objPedido.Codigo)

                    dr("Ext003") = "" 'drAux("Ext003")
                    dr("Inscli") = objPedido.Cliente.InscricaoEstadual
                    dr("Via") = h
                    dr("Tipo") = "" 'drAux("Tipo")
                    dr("Valnot") = objPedido.Vencimentos(j).ValorLiquidoOficial
                    dr("Emissa") = Format(Now.Date, "dd/MM/yyyy")
                    dr("BaiCli") = objPedido.Cliente.Bairro

                    ds.Tables(0).Rows.Add(dr)
                Next
            Next

            Relatorio(ds)
        Else
            MsgBox(Me.Page, "Pedido sem Vencimento(s).")
        End If
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

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo") = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteDpRev" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada!")
        ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            MsgBox(Me.Page, "Cliente não foi selecionado!")
        Else
            HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
            Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
            Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
            parameters("safra") = ddlSafra.SelectedValue
            Popup.ConsultaDePedidos(Me.Page, "objPedidoDpRev" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        End If
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("EmissaoDeDuplicatas", "RELATORIO") Then
                If rdGrafica.Checked Then
                    impressaoGrafica()
                Else
                    impressaoMatricial()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "EmissaoDeDuplicatas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class