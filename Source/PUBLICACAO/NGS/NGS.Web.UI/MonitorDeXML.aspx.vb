Imports NGS.Lib.Negocio
Imports System.IO
Imports System.Xml
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class MonitorDeXML
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)

        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("MonitorDeXML", "ACESSAR") Then
                BuncarUnidadeDeNegocio()
                Limpar(True)
                hidLoad.Value = "S"
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        Else
            hidLoad.Value = "N"
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
            Limpar(False)
            LerXML()
        Catch ex As Exception
            'Dim erro = teste

            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LerXML()

        If Not chkNFe.Checked AndAlso Not chkCTe.Checked Then
            MsgBox(Me.Page, "Selecione Nfe importadas ou Nfe pendentes ou ambos para Consulta.")
            Exit Sub
        End If

        If rbPendente.Checked = False And (txtDataInicial.Text.Length = 0 Or txtDataFinal.Text.Length = 0) Then
            MsgBox(Me.Page, "É necessario informar um período para pesquisa.")
            Exit Sub
        End If

        Dim fileEmpresa As String = "C:\NGS\xmlCopy\" & ddlEmpresa.SelectedValue.Split("-")(0) & ".xml"

        If Not File.Exists(fileEmpresa) Then
            MsgBox(Me.Page, "Arquivo de Configuracão da Empresa não foi encontrado.")
            Exit Sub
        End If

        Dim xmlConf As New XmlDocument
        xmlConf.Load(fileEmpresa)

        Dim dataDoUltimoArquivo = New DirectoryInfo(xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("eventoDFe").InnerText).GetFiles().OrderByDescending(Function(p) p.LastWriteTime).FirstOrDefault().LastWriteTime

        Dim proximaAtualizacao = dataDoUltimoArquivo.AddMinutes(70)

        lblAtualizado.Text = dataDoUltimoArquivo.ToString("dd/MM/yyyy") & " às " & dataDoUltimoArquivo.ToString("HH:mm:ss") & "."

        lblProximaAtualizacao.Text = proximaAtualizacao.ToString("dd/MM/yyyy") & " às " & proximaAtualizacao.ToString("HH:mm:ss") & "."

        Dim strSQL As String
        Dim codigo As String = String.Empty
        Dim cidade As String = String.Empty
        Dim arquivoEvento As String = String.Empty
        'Dim temCPF As Boolean

        strSQL = "WITH FilteredDocumentoXML AS (" & vbCrLf &
                    "    SELECT " & vbCrLf &
                    "        dXML.Empresa_Id," & vbCrLf &
                    "        dXML.Cliente_Id," & vbCrLf &
                    "        dXML.ClienteNome," & vbCrLf &
                    "        (dXML.ClienteCidade + '/' + dXML.ClienteUF) AS ClienteCidade," & vbCrLf &
                    "        dXML.Chave_Id," & vbCrLf &
                    "        dXML.Tipo," & vbCrLf &
                    "        dXML.Serie_Id," & vbCrLf &
                    "        dXML.Numero_Id," & vbCrLf &
                    "        dXML.Emissao," & vbCrLf &
                    "        dXML.Situacao," & vbCrLf &
                    "        dXML.CFOP," & vbCrLf &
                    "        dXML.ValorDoProduto," & vbCrLf &
                    "        dXML.ValorTotalDaNota" & vbCrLf &
                    "    FROM " & vbCrLf &
                    "        DocumentoXML dXML" & vbCrLf &
                    "    WHERE " & vbCrLf &
                    "        dXML.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                    "        AND LEN(ISNULL(dXML.Chave_Id, '')) > 0" & vbCrLf &
                    ")" & vbCrLf &
                    "" & vbCrLf &
                    "SELECT " & vbCrLf &
                    "    nFE.Nota_Id," & vbCrLf &
                    "    fXML.Cliente_Id AS Cliente," & vbCrLf &
                    "    fXML.ClienteNome," & vbCrLf &
                    "    fXML.ClienteCidade," & vbCrLf &
                    "    fXML.Chave_Id AS Chave," & vbCrLf &
                    "    fXML.Tipo," & vbCrLf &
                    "    fXML.Serie_Id AS Serie," & vbCrLf &
                    "    fXML.Numero_Id AS nNF," & vbCrLf &
                    "    fXML.Emissao," & vbCrLf &
                    "    ISNULL(nFE.Usuario,'') AS Usuario," & vbCrLf &
                    "    ISNULL(nFE.Hora,'') AS DiaLancamento," & vbCrLf &
                    "    CASE" & vbCrLf &
                    "        WHEN LEN(ISNULL(nFE.ChaveNfe, '')) > 0 THEN 'SIM' ELSE ''" & vbCrLf &
                    "    END AS Lancada," & vbCrLf &
                    "    CASE" & vbCrLf &
                    "        WHEN UPPER(ISNULL(doc.Descricao,'')) = UPPER(fXML.Chave_Id+'-'+fXML.Tipo+'.xml') THEN 'SIM' ELSE ''" & vbCrLf &
                    "    END AS TemAnexo," & vbCrLf &
                    "    fXML.Situacao," & vbCrLf &
                    "    ISNULL(n.Situacao,0) AS SituacaoNFE," & vbCrLf &
                    "    CASE " & vbCrLf &
                    "        WHEN LEN(ISNULL(nFE.ChaveNfe, '')) > 0 OR ISNULL(n.Situacao,0) = 2 THEN" & vbCrLf &
                    "            'Nota Fiscal Cancelada pelo Emitente, verifique que a mesma está lançada no Sistema.'" & vbCrLf &
                    "        ELSE " & vbCrLf &
                    "            ''" & vbCrLf &
                    "    END AS ToolTip," & vbCrLf &
                    "    CASE " & vbCrLf &
                    "        WHEN LEN(ISNULL(nFE.ChaveNfe, '')) > 0 OR ISNULL(n.Situacao,0) = 2 THEN " & vbCrLf &
                    "            'True'" & vbCrLf &
                    "        ELSE " & vbCrLf &
                    "            'False'" & vbCrLf &
                    "    END AS visibleImageButton," & vbCrLf &
                    "    fXML.CFOP," & vbCrLf &
                    "    fXML.ValorDoProduto," & vbCrLf &
                    "    fXML.ValorTotalDaNota ," & vbCrLf &
                    "    n.*" & vbCrLf &
                    "FROM " & vbCrLf &
                    "    FilteredDocumentoXML fXML" & vbCrLf &
                    "LEFT JOIN " & vbCrLf &
                    "    NFERealizadas nFE" & vbCrLf &
                    "    ON fXML.Chave_Id = nFE.ChaveNfe" & vbCrLf &
                    "    AND nFE.EntradaSaida_Id = 'E' " & vbCrLf &
                    "LEFT JOIN " & vbCrLf &
                    "    NotasFiscais n" & vbCrLf &
                    "    ON n.Empresa_Id = nFE.Empresa_Id" & vbCrLf &
                    "    AND n.EndEmpresa_iD = nFE.EndEmpresa_Id" & vbCrLf &
                    "    AND n.Cliente_Id = nFE.Cliente_Id" & vbCrLf &
                    "    AND n.EndCliente_Id = nFE.EndCliente_Id" & vbCrLf &
                    "    AND n.EntradaSaida_Id = nFE.EntradaSaida_Id" & vbCrLf &
                    "    AND n.Serie_Id = nFE.Serie_Id" & vbCrLf &
                    "    AND n.Nota_Id = nFE.Nota_Id" & vbCrLf &
                    "LEFT JOIN " & vbCrLf &
                    "    Documentos.dbo.Arquivo doc" & vbCrLf &
                    "    ON doc.Empresa_Id = nFE.Empresa_Id" & vbCrLf &
                    "    AND doc.EndEmpresa_iD = nFE.EndEmpresa_Id" & vbCrLf &
                    "    AND doc.Cliente_Id = nFE.Cliente_Id" & vbCrLf &
                    "    AND doc.EndCliente_Id = nFE.EndCliente_Id" & vbCrLf &
                    "    AND doc.Serie_Id = nFE.Serie_Id" & vbCrLf &
                    "    AND doc.Nota_Id = nFE.Nota_Id" & vbCrLf &
                    "WHERE 1=1" & vbCrLf &
                    "    AND fXML.Situacao = 1 " & vbCrLf &
                    "    AND fXML.Emissao > '2018-01-01'" & vbCrLf

        If chkNFe.Checked And chkCTe.Checked Then
        ElseIf chkNFe.Checked Then
            strSQL &= " AND fXML.Tipo = 'NFe' " & vbCrLf
        ElseIf chkCTe.Checked Then
            strSQL &= " AND fXML.Tipo = 'CTe' " & vbCrLf
        End If

        If rbPendente.Checked Then
            strSQL &= " AND nFE.ChaveNfe IS NULL " & vbCrLf
        End If

        If txtDataInicial.Text.Length > 0 Or txtDataFinal.Text.Length > 0 Then
            strSQL &= " AND fXML.Emissao BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' " & vbCrLf &
                      " AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        'Data de Leitura para o Cliente Verde - Furlan - 12/12/2023
        If ddlEmpresa.SelectedValue.Split("-")(0) = "44979506000240" Then
            strSQL &= "  And fXML.Emissao > '2023-10-31'" & vbCrLf
        ElseIf Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) = "24450490" Then 'Data de Leitura para o Cliente RT Grãos - Felipe - 30/07/2024
            strSQL &= "  AND fXML.Emissao > '2024-06-30'" & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "DocumentoXML")

        If ds.Tables(0).Rows.Count > 0 Then

            For Each row As DataRow In ds.Tables(0).Rows

                If row("Situacao") = 2 AndAlso row("SituacaoNFE") = 0 Then
                    'NÃO FAZ NADA
                    Continue For
                Else
                    Dim drXML As DataRow = CType(Session("objXML" & HID.Value), DataTable).NewRow()

                    drXML("Cliente") = row("Cliente")
                    drXML("ClienteNome") = row("ClienteNome")
                    drXML("ClienteCidade") = row("ClienteCidade")
                    drXML("Chave") = row("Chave")
                    drXML("Codigo") = row("Chave") & "-" & row("Tipo") & ".xml"
                    drXML("Tipo") = row("Tipo")
                    drXML("Emissao") = row("Emissao")
                    drXML("Serie") = row("Serie")
                    drXML("nNF") = row("nNF")
                    drXML("Lancada") = row("Lancada")
                    drXML("Usuario") = row("Usuario")
                    If row("Lancada") = "SIM" Then
                        drXML("DiaLancamento") = row("DiaLancamento")
                    Else
                        drXML("DiaLancamento") = ""
                    End If
                    drXML("TemAnexo") = row("TemAnexo")

                    If row("Situacao") = 2 Then
                        drXML("Atencao") = "S"
                    Else
                        drXML("Atencao") = "N"
                    End If

                    drXML("ToolTip") = row("ToolTip")
                    drXML("visibleImageButton") = row("visibleImageButton")

                    drXML("CFOP") = row("CFOP")
                    drXML("ValorDoProduto") = row("ValorDoProduto")
                    drXML("ValorTotalDaNota") = row("ValorTotalDaNota")

                    CType(Session("objXML" & HID.Value), DataTable).Rows.Add(drXML)

                End If
            Next
        End If

        If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
            Dim ds1 As DataView

            If rbPendente.Checked Then
                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = ''").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = '' or Atencao = 'S'").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para visualização.")
                    Exit Sub
                End If
            ElseIf rbLancados.Checked Then
                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM'").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM'").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para visualização.")
                    Exit Sub
                End If
            ElseIf rbSemAnexo.Checked Then

                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM' and TemAnexo = ''").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM' and TemAnexo = ''").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para visualização.")
                    Exit Sub
                End If
            Else
                If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable))
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para visualização.")
                    Exit Sub
                End If
            End If

            ds1.Sort = "Emissao DESC"
            'ds1.Sort = "ClienteNome"

            lblRegistros.Text = "Total de Registros: " & ds1.Table.Rows.Count

            gridXML.DataSource = ds1
            gridXML.DataBind()

            For x As Integer = 0 To gridXML.Rows.Count - 1

                CType(gridXML.Rows(x).FindControl("imgInfXML"), ImageButton).Visible = CType(gridXML.Rows(x).FindControl("visibleImageButton"), TextBox).Text = "True"
                CType(gridXML.Rows(x).FindControl("imgInfXML"), ImageButton).ToolTip = CType(gridXML.Rows(x).FindControl("ToolTip"), TextBox).Text

            Next

        Else
            MsgBox(Me.Page, "Sem registros para visualização.")
        End If

    End Sub

    Private Sub LerXML_backup()

        If Not chkNFe.Checked AndAlso Not chkCTe.Checked Then
            MsgBox(Me.Page, "Selecione Nfe importadas ou Nfe pendentes ou ambos para Consulta.")
            Exit Sub
        End If

        Dim fileEmpresa As String = "C:\NGS\xmlCopy\" & ddlEmpresa.SelectedValue.Split("-")(0) & ".xml"

        If Not File.Exists(fileEmpresa) Then
            MsgBox(Me.Page, "Arquivo de Configuracão da Empresa não foi encontrado.")
            Exit Sub
        End If

        Dim xmlConf As New XmlDocument
        xmlConf.Load(fileEmpresa)

        Dim dataDoUltimoArquivo = New DirectoryInfo(xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("eventoDFe").InnerText).GetFiles().OrderByDescending(Function(p) p.LastWriteTime).FirstOrDefault().LastWriteTime

        Dim proximaAtualizacao = dataDoUltimoArquivo.AddMinutes(70)

        lblAtualizado.Text = dataDoUltimoArquivo.ToString("dd/MM/yyyy") & " às " & dataDoUltimoArquivo.ToString("HH:mm:ss") & "."

        lblProximaAtualizacao.Text = proximaAtualizacao.ToString("dd/MM/yyyy") & " às " & proximaAtualizacao.ToString("HH:mm:ss") & "."

        Dim strSQL As String
        Dim codigo As String = String.Empty
        Dim cidade As String = String.Empty
        Dim arquivoEvento As String = String.Empty
        'Dim temCPF As Boolean

        If chkNFe.Checked Then
            strSQL = "select dXML.Cliente_Id AS Cliente, dXML.ClienteNome, (dXML.ClienteCidade + '/' + dXML.ClienteUF) AS ClienteCidade," & vbCrLf &
                        "		dXML.Chave_Id AS Chave, dXML.Tipo, dXML.Serie_Id AS Serie, dXML.Numero_Id AS nNF, dXML.Emissao, isnull(nFE.Usuario,'') AS Usuario, isnull(nFE.Hora,'') AS DiaLancamento," & vbCrLf &
                        "		case" & vbCrLf &
                        "			when len(isnull(nFE.ChaveNfe, '')) > 0" & vbCrLf &
                        "				then 'SIM'" & vbCrLf &
                        "				else ''" & vbCrLf &
                        "			end AS Lancada," & vbCrLf &
                        "		case" & vbCrLf &
                        "			when UPPER(isnull(doc.Descricao,'')) = UPPER(dXML.Chave_Id+'-'+dXML.Tipo+'.xml')" & vbCrLf &
                        "				then 'SIM'" & vbCrLf &
                        "				else ''" & vbCrLf &
                        "			end AS TemAnexo," & vbCrLf &
                        "		dXML.Situacao, isnull(n.Situacao,0) AS SituacaoNFE, " & vbCrLf &
                        "		CASE " & vbCrLf &
                        "           WHEN LEN(ISNULL(nFE.ChaveNfe, '')) > 0 OR ISNULL(n.Situacao,0) = 2 THEN " & vbCrLf &
                        "               'Nota Fiscal Cancelada pelo Emitente, verifique que a mesma está lançada no Sistema.' " & vbCrLf &
                        "           ELSE " & vbCrLf &
                        "               '' " & vbCrLf &
                        "       END AS ToolTip, " & vbCrLf &
                        "		CASE " & vbCrLf &
                        "           WHEN LEN(ISNULL(nFE.ChaveNfe, '')) > 0 OR ISNULL(n.Situacao,0) = 2 THEN " & vbCrLf &
                        "               'True' " & vbCrLf &
                        "           ELSE " & vbCrLf &
                        "               'False' " & vbCrLf &
                        "       END AS visibleImageButton, " & vbCrLf &
                        "       dXML.CFOP AS CFOP, " & vbCrLf &
                        "       dXML.ValorDoProduto AS ValorDoProduto, " & vbCrLf &
                        "       dXML.ValorTotalDaNota AS ValorTotalDaNota  " & vbCrLf &
                        "FROM DocumentoXML dXML" & vbCrLf &
                        "Left join NFERealizadas nFE  " & vbCrLf &
                        "   ON dXML.Chave_Id           = nFE.ChaveNfe  " & vbCrLf &
                        "	left join NotasFiscais n" & vbCrLf &
                        "		ON n.Empresa_Id         = nFE.Empresa_Id" & vbCrLf &
                        "		and n.EndEmpresa_iD     = nFE.EndEmpresa_Id" & vbCrLf &
                        "		and n.Cliente_Id        = nFE.Cliente_Id" & vbCrLf &
                        "		and n.EndCliente_Id     = nFE.EndCliente_Id" & vbCrLf &
                        "		and n.EntradaSaida_Id   = nFE.EntradaSaida_Id" & vbCrLf &
                        "		and n.Serie_Id          = nFE.Serie_Id" & vbCrLf &
                        "		and n.Nota_Id           = nFE.Nota_Id" & vbCrLf &
                        "	left join Documentos.dbo.Arquivo doc" & vbCrLf &
                        "		ON doc.Empresa_Id       = nFE.Empresa_Id" & vbCrLf &
                        "		and doc.EndEmpresa_iD   = nFE.EndEmpresa_Id" & vbCrLf &
                        "		and doc.Cliente_Id      = nFE.Cliente_Id" & vbCrLf &
                        "		and doc.EndCliente_Id   = nFE.EndCliente_Id" & vbCrLf &
                        "		and doc.Serie_Id        = nFE.Serie_Id" & vbCrLf &
                        "		and doc.Nota_Id         = nFE.Nota_Id" & vbCrLf &
                        "where dXML.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                        "  and dXML.Tipo = 'NFe'" & vbCrLf
            '"  and not dXML.Chave_Id in (select ChaveNfe from ManifestoNFE where ChaveNfe = dXML.Chave_Id and CodigoEvento = 210240)"

            'Data de Leitura para o Cliente Verde - Furlan - 12/12/2023
            If ddlEmpresa.SelectedValue.Split("-")(0) = "44979506000240" Then
                strSQL &= "  and dXML.Emissao > '2023-10-31'" & vbCrLf
            ElseIf Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) = "24450490" Then 'Data de Leitura para o Cliente RT Grãos - Felipe - 30/07/2024
                strSQL &= "  and dXML.Emissao > '2024-06-30'" & vbCrLf
            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "DocumentoXML")

            If ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    If row("Situacao") = 2 AndAlso row("SituacaoNFE") = 0 Then
                        'NÃO FAZ NADA
                        Continue For
                    Else
                        Dim drXML As DataRow = CType(Session("objXML" & HID.Value), DataTable).NewRow()

                        drXML("Cliente") = row("Cliente")
                        drXML("ClienteNome") = row("ClienteNome")
                        drXML("ClienteCidade") = row("ClienteCidade")
                        drXML("Chave") = row("Chave")
                        drXML("Codigo") = row("Chave") & "-" & row("Tipo") & ".xml"
                        drXML("Tipo") = row("Tipo")
                        drXML("Emissao") = row("Emissao")
                        drXML("Serie") = row("Serie")
                        drXML("nNF") = row("nNF")
                        drXML("Lancada") = row("Lancada")
                        drXML("Usuario") = row("Usuario")
                        If row("Lancada") = "SIM" Then
                            drXML("DiaLancamento") = row("DiaLancamento")
                        Else
                            drXML("DiaLancamento") = ""
                        End If
                        drXML("TemAnexo") = row("TemAnexo")

                        If row("Situacao") = 2 Then
                            drXML("Atencao") = "S"
                        Else
                            drXML("Atencao") = "N"
                        End If

                        drXML("ToolTip") = row("ToolTip")
                        drXML("visibleImageButton") = row("visibleImageButton")

                        drXML("CFOP") = row("CFOP")
                        drXML("ValorDoProduto") = row("ValorDoProduto")
                        drXML("ValorTotalDaNota") = row("ValorTotalDaNota")

                        CType(Session("objXML" & HID.Value), DataTable).Rows.Add(drXML)

                    End If
                Next
            End If
        End If

        'If chkNFe.Checked Then
        '    'Buscar XML's NFe
        '    Dim DirDiretorioNFe As DirectoryInfo = New DirectoryInfo(xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("origemNFe").InnerText)

        '    Dim oFileInfoCollectionNFe() As FileInfo

        '    Dim oFileInfoNFe As FileInfo

        '    Dim i As Integer

        '    oFileInfoCollectionNFe = DirDiretorioNFe.GetFiles("*-nfe.xml")

        '    For i = 0 To oFileInfoCollectionNFe.Length() - 1
        '        oFileInfoNFe = oFileInfoCollectionNFe.GetValue(i)

        '        Dim xmlDoc As New XmlDocument
        '        xmlDoc.Load(xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("origemNFe").InnerText & "\" & oFileInfoNFe.Name)

        '        'teste = oFileInfoNFe.Name

        '        'If teste = "41220605366261000143550010000122801801548903-nfe.xml" Or teste = "14220605383036000549550010000023941790520575-nfe.xml" Then
        '        '    Dim erro = teste
        '        'End If

        '        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then 'Se existe o Protocolo da Sefaz com autorização da NFe

        '            strSQL = "Select ChaveNfe" & vbCrLf & _
        '                        "  from ManifestoNFE" & vbCrLf & _
        '                        "where ChaveNFE = '" & oFileInfoNFe.Name.Split("-")(0) & "'" & vbCrLf & _
        '                        "and CodigoEvento = 210240"

        '            Dim dsM As DataSet = Banco.ConsultaDataSet(strSQL, "ManifestoNFE")

        '            If dsM.Tables(0).Rows.Count = 0 Then

        '                arquivoEvento = xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("eventoNFe").InnerText & "\" & oFileInfoNFe.Name.Split("-")(0) & "-proc-evcanc.xml"

        '                If Not File.Exists(arquivoEvento) Then

        '                    Dim cpfDest = xmlDoc.GetElementsByTagName("dest")

        '                    If cpfDest IsNot Nothing AndAlso cpfDest.Count() > 0 Then

        '                        For Each item As XmlNode In cpfDest

        '                            If item("CPF") IsNot Nothing Then
        '                                temCPF = True
        '                            Else
        '                                temCPF = False
        '                            End If
        '                        Next
        '                    End If

        '                    If temCPF Then
        '                        codigo = xmlDoc.GetElementsByTagName("dest").GetNode("dest").ChildNodes.GetNode("CPF").InnerText
        '                    Else
        '                        codigo = xmlDoc.GetElementsByTagName("dest").GetNode("dest").ChildNodes.GetNode("CNPJ").InnerText
        '                    End If

        '                    cidade = xmlDoc.GetElementsByTagName("dest").GetNode("dest").ChildNodes.GetNode("enderDest").ChildNodes.GetNode("xMun").InnerText

        '                    If codigo = ddlEmpresa.SelectedValue.Split("-")(0) Then

        '                        Dim drXML As DataRow = CType(Session("objXML" & HID.Value), DataTable).NewRow()

        '                        drXML("Nome") = oFileInfoNFe.Name.Split("-")(0)
        '                        drXML("Tipo") = oFileInfoNFe.Name.Split("-")(1).Split(".")(0)
        '                        drXML("Dia") = oFileInfoNFe.LastWriteTime.ToString("dd-MM-yyyy HH:mm:ss")

        '                        drXML("nNF") = xmlDoc.GetElementsByTagName("ide").GetNode("ide").ChildNodes.GetNode("nNF").InnerText

        '                        drXML("Codigo") = oFileInfoNFe.Name

        '                        If Not xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CPF") Is Nothing AndAlso xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CPF").InnerText.Count > 0 Then
        '                            drXML("Cliente") = Funcoes.FormatarCpfCnpj(xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CPF").InnerText) & " " & _
        '                                                xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("xNome").InnerText & " - " & _
        '                                                xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("enderEmit").ChildNodes.GetNode("xMun").InnerText & "/" & _
        '                                                xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("enderEmit").ChildNodes.GetNode("UF").InnerText
        '                        Else
        '                            drXML("Cliente") = Funcoes.FormatarCpfCnpj(xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CNPJ").InnerText) & " " & _
        '                                                xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("xNome").InnerText & " - " & _
        '                                                xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("enderEmit").ChildNodes.GetNode("xMun").InnerText & "/" & _
        '                                                xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("enderEmit").ChildNodes.GetNode("UF").InnerText
        '                        End If

        '                        strSQL = "SELECT nfe.Usuario, nfe.Hora, nfe.ChaveNFE, isnull(doc.Descricao,'') AS Descricao" & vbCrLf & _
        '                                    "FROM NFeRealizadas nfe" & vbCrLf & _
        '                                    "	inner join Clientes c" & vbCrLf & _
        '                                    "			on c.Cliente_id   = nfe.Cliente_id" & vbCrLf & _
        '                                    "			and c.Endereco_Id = nfe.EndCliente_Id" & vbCrLf & _
        '                                    "	left join Documentos.dbo.Arquivo doc" & vbCrLf & _
        '                                    "			on doc.Empresa_Id     = nfe.Empresa_Id" & vbCrLf & _
        '                                    "			and doc.EndEmpresa_Id = nfe.EndEmpresa_Id" & vbCrLf & _
        '                                    "			and doc.Cliente_Id    = nfe.Cliente_Id" & vbCrLf & _
        '                                    "			and doc.EndCliente_Id = nfe.EndCliente_Id" & vbCrLf & _
        '                                    "			and doc.Nota_Id       = nfe.Nota_Id" & vbCrLf & _
        '                                    "			and doc.Serie_Id      = nfe.Serie_Id" & vbCrLf & _
        '                                    "			and (doc.Descricao like '%-nfe%' or doc.Descricao like '%-CTe%')" & vbCrLf & _
        '                                    "Where nfe.ChaveNFE = '" & drXML("Nome") & "'"

        '                        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "NFeRealizadas")

        '                        If ds.Tables(0).Rows.Count > 0 Then

        '                            drXML("Lancada") = "SIM"
        '                            drXML("Usuario") = ds.Tables(0).Rows(0).Item("Usuario")
        '                            drXML("DiaLancamento") = ds.Tables(0).Rows(0).Item("Hora")

        '                            If ds.Tables(0).Rows(0).Item("Descricao").ToUpper = oFileInfoNFe.Name.ToUpper Then
        '                                drXML("TemAnexo") = "SIM"
        '                            Else
        '                                drXML("TemAnexo") = ""
        '                            End If
        '                        Else
        '                            drXML("Lancada") = ""
        '                            drXML("Usuario") = ""
        '                            drXML("DiaLancamento") = ""
        '                            drXML("TemAnexo") = ""
        '                        End If

        '                        CType(Session("objXML" & HID.Value), DataTable).Rows.Add(drXML)
        '                    End If
        '                End If
        '            End If
        '        End If
        '    Next
        'End If

        If chkCTe.Checked Then
            strSQL = "select dXML.Cliente_Id AS Cliente, dXML.ClienteNome, (dXML.ClienteCidade + '/' + dXML.ClienteUF) AS ClienteCidade," & vbCrLf &
                        "		dXML.Chave_Id AS Chave, dXML.Tipo, dXML.Serie_Id AS Serie, dXML.Numero_Id AS nNF, dXML.Emissao, isnull(nFE.Usuario,'') AS Usuario, isnull(nFE.Hora,'') AS DiaLancamento," & vbCrLf &
                        "		case" & vbCrLf &
                        "			when len(isnull(nFE.ChaveNfe, '')) > 0" & vbCrLf &
                        "				then 'SIM'" & vbCrLf &
                        "				else ''" & vbCrLf &
                        "			end AS Lancada," & vbCrLf &
                        "		case" & vbCrLf &
                        "			when UPPER(isnull(doc.Descricao,'')) = UPPER(dXML.Chave_Id+'-'+dXML.Tipo+'.xml')" & vbCrLf &
                        "				then 'SIM'" & vbCrLf &
                        "				else ''" & vbCrLf &
                        "			end AS TemAnexo," & vbCrLf &
                        "		dXML.Situacao, isnull(n.Situacao,0) AS SituacaoNFE, " & vbCrLf &
                        "		CASE " & vbCrLf &
                        "           WHEN LEN(ISNULL(nFE.ChaveNfe, '')) > 0 OR ISNULL(n.Situacao,0) = 2 THEN " & vbCrLf &
                        "               'Nota Fiscal Cancelada pelo Emitente, verifique que a mesma está lançada no Sistema.' " & vbCrLf &
                        "           ELSE " & vbCrLf &
                        "               '' " & vbCrLf &
                        "       END AS ToolTip, " & vbCrLf &
                        "		CASE " & vbCrLf &
                        "           WHEN LEN(ISNULL(nFE.ChaveNfe, '')) > 0 OR ISNULL(n.Situacao,0) = 2 THEN " & vbCrLf &
                        "               'True' " & vbCrLf &
                        "           ELSE " & vbCrLf &
                        "               'False' " & vbCrLf &
                        "       END AS visibleImageButton, " & vbCrLf &
                        "       dXML.CFOP AS CFOP, " & vbCrLf &
                        "       dXML.ValorDoProduto AS ValorDoProduto, " & vbCrLf &
                        "       dXML.ValorTotalDaNota AS ValorTotalDaNota  " & vbCrLf &
                        "from DocumentoXML dXML" & vbCrLf &
                        "Left join NFERealizadas nFE  " & vbCrLf &
                        "   ON dXML.Chave_Id           = nFE.ChaveNfe  " & vbCrLf &
                        "	left join NotasFiscais n" & vbCrLf &
                        "		ON n.Empresa_Id         = nFE.Empresa_Id" & vbCrLf &
                        "		and n.EndEmpresa_iD     = nFE.EndEmpresa_Id" & vbCrLf &
                        "		and n.Cliente_Id        = nFE.Cliente_Id" & vbCrLf &
                        "		and n.EndCliente_Id     = nFE.EndCliente_Id" & vbCrLf &
                        "		and n.EntradaSaida_Id   = nFE.EntradaSaida_Id" & vbCrLf &
                        "		and n.Serie_Id          = nFE.Serie_Id" & vbCrLf &
                        "		and n.Nota_Id           = nFE.Nota_Id" & vbCrLf &
                        "	left join Documentos.dbo.Arquivo doc" & vbCrLf &
                        "		ON doc.Empresa_Id       = nFE.Empresa_Id" & vbCrLf &
                        "		and doc.EndEmpresa_iD   = nFE.EndEmpresa_Id" & vbCrLf &
                        "		and doc.Cliente_Id      = nFE.Cliente_Id" & vbCrLf &
                        "		and doc.EndCliente_Id   = nFE.EndCliente_Id" & vbCrLf &
                        "		and doc.Serie_Id        = nFE.Serie_Id" & vbCrLf &
                        "		and doc.Nota_Id         = nFE.Nota_Id" & vbCrLf &
                        "where dXML.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                        "  and dXML.Tipo = 'CTe'"
            '"  and not dXML.Chave_Id in (select ChaveNfe from ManifestoNFE where ChaveNfe = dXML.Chave_Id and CodigoEvento = 610110)"

            'Data de Leitura para o Cliente Verde - Furlan - 12/12/2023
            If Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) = "44979506" Then
                strSQL &= "  and dXML.Emissao > '2023-12-31'" & vbCrLf
            ElseIf Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) = "24450490" Then 'Data de Leitura para o Cliente RT Grãos - Felipe - 30/07/2024
                strSQL &= "  and dXML.Emissao > '2024-06-30'" & vbCrLf
            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "DocumentoXML")

            If ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    If row("Situacao") = 2 AndAlso row("SituacaoNFE") = 0 Then
                        'NÃO FAZ NADA
                        Continue For
                    Else
                        Dim drXML As DataRow = CType(Session("objXML" & HID.Value), DataTable).NewRow()

                        drXML("Cliente") = row("Cliente")
                        drXML("ClienteNome") = row("ClienteNome")
                        drXML("ClienteCidade") = row("ClienteCidade")
                        drXML("Chave") = row("Chave")
                        drXML("Codigo") = row("Chave") & "-" & row("Tipo") & ".xml"
                        drXML("Tipo") = row("Tipo")
                        drXML("Emissao") = row("Emissao")
                        drXML("Serie") = row("Serie")
                        drXML("nNF") = row("nNF")
                        drXML("Lancada") = row("Lancada")
                        drXML("Usuario") = row("Usuario")
                        If row("Lancada") = "SIM" Then
                            drXML("DiaLancamento") = row("DiaLancamento")
                        Else
                            drXML("DiaLancamento") = ""
                        End If
                        drXML("TemAnexo") = row("TemAnexo")

                        If row("Situacao") = 2 Then
                            drXML("Atencao") = "S"
                        Else
                            drXML("Atencao") = "N"
                        End If

                        drXML("ToolTip") = row("ToolTip")
                        drXML("visibleImageButton") = row("visibleImageButton")

                        drXML("CFOP") = row("CFOP")
                        drXML("ValorDoProduto") = row("ValorDoProduto")
                        drXML("ValorTotalDaNota") = row("ValorTotalDaNota")

                        CType(Session("objXML" & HID.Value), DataTable).Rows.Add(drXML)

                    End If
                Next

            End If
        End If

        'If chkCTe.Checked Then
        '    'Buscar XML's CTe
        '    Dim DirDiretorioCTe As DirectoryInfo = New DirectoryInfo(xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("origemCTe").InnerText)

        '    Dim oFileInfoCollectionCTe() As FileInfo

        '    Dim oFileInfoCTe As FileInfo

        '    Dim j As Integer

        '    oFileInfoCollectionCTe = DirDiretorioCTe.GetFiles("*-CTe.xml")

        '    For j = 0 To oFileInfoCollectionCTe.Length() - 1
        '        oFileInfoCTe = oFileInfoCollectionCTe.GetValue(j)

        '        Dim xmlDoc As New XmlDocument
        '        xmlDoc.Load(xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("origemCTe").InnerText & "\" & oFileInfoCTe.Name)

        '        'teste = oFileInfoCTe.Name

        '        If xmlDoc.GetElementsByTagName("protCTe").Count = 1 Then 'Se existe o Protocolo da Sefaz com autorização da CTe

        '            strSQL = "Select ChaveNfe" & vbCrLf & _
        '                        "  from ManifestoNFE" & vbCrLf & _
        '                        "where ChaveNFE = '" & oFileInfoCTe.Name.Split("-")(0) & "'" & vbCrLf & _
        '                        "and CodigoEvento = 610110"

        '            Dim dsM As DataSet = Banco.ConsultaDataSet(strSQL, "ManifestoNFE")

        '            If dsM.Tables(0).Rows.Count = 0 Then

        '                arquivoEvento = xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("eventoCTe").InnerText & "\" & oFileInfoCTe.Name.Split("-")(0) & "-proc-evcanc.xml"

        '                If Not File.Exists(arquivoEvento) Then
        '                    Dim tomador As Integer

        '                    codigo = String.Empty

        '                    If xmlDoc.GetElementsByTagName("toma3").Count = 1 Then 'Se existe o Tomador de Serviço
        '                        '0 – Remetente  1 – Expedidor  2 – Recebedor  3 – Destinatário
        '                        tomador = CInt(xmlDoc.GetElementsByTagName("toma3").GetNode("toma3").ChildNodes.GetNode("toma").InnerText)

        '                        If tomador = 0 Then
        '                            codigo = xmlDoc.GetElementsByTagName("rem").GetNode("rem").ChildNodes.GetNode("CNPJ").InnerText
        '                            cidade = xmlDoc.GetElementsByTagName("rem").GetNode("rem").ChildNodes.GetNode("enderReme").ChildNodes.GetNode("xMun").InnerText
        '                        ElseIf tomador = 1 Then
        '                            codigo = xmlDoc.GetElementsByTagName("exped").GetNode("exped").ChildNodes.GetNode("CNPJ").InnerText
        '                            cidade = xmlDoc.GetElementsByTagName("exped").GetNode("exped").ChildNodes.GetNode("enderExped").ChildNodes.GetNode("xMun").InnerText
        '                        ElseIf tomador = 2 Then
        '                            codigo = xmlDoc.GetElementsByTagName("receb").GetNode("receb").ChildNodes.GetNode("CNPJ").InnerText
        '                            cidade = xmlDoc.GetElementsByTagName("receb").GetNode("receb").ChildNodes.GetNode("enderReceb").ChildNodes.GetNode("xMun").InnerText
        '                        ElseIf tomador = 3 Then
        '                            codigo = xmlDoc.GetElementsByTagName("dest").GetNode("dest").ChildNodes.GetNode("CNPJ").InnerText
        '                            cidade = xmlDoc.GetElementsByTagName("dest").GetNode("dest").ChildNodes.GetNode("enderDest").ChildNodes.GetNode("xMun").InnerText
        '                        End If
        '                    End If

        '                    If xmlDoc.GetElementsByTagName("toma4").Count = 1 Then
        '                        codigo = xmlDoc.GetElementsByTagName("toma4").GetNode("toma4").ChildNodes.GetNode("CNPJ").InnerText
        '                        cidade = xmlDoc.GetElementsByTagName("toma4").GetNode("toma4").ChildNodes.GetNode("enderToma").ChildNodes.GetNode("xMun").InnerText
        '                    End If

        '                    If codigo = ddlEmpresa.SelectedValue.Split("-")(0) Then

        '                        Dim drXML As DataRow = CType(Session("objXML" & HID.Value), DataTable).NewRow()

        '                        drXML("Nome") = oFileInfoCTe.Name.Split("-")(0)
        '                        drXML("Tipo") = oFileInfoCTe.Name.Split("-")(1).Split(".")(0)
        '                        drXML("Dia") = oFileInfoCTe.LastWriteTime.ToString("dd-MM-yyyy HH:mm:ss")

        '                        drXML("nNF") = xmlDoc.GetElementsByTagName("ide").GetNode("ide").ChildNodes.GetNode("nCT").InnerText

        '                        drXML("Codigo") = oFileInfoCTe.Name

        '                        drXML("Cliente") = Funcoes.FormatarCpfCnpj(xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CNPJ").InnerText) & " " & _
        '                                            xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("xNome").InnerText & " - " & _
        '                                            xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("enderEmit").ChildNodes.GetNode("xMun").InnerText & "/" & _
        '                                            xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("enderEmit").ChildNodes.GetNode("UF").InnerText

        '                        strSQL = "SELECT nfe.Usuario, nfe.Hora, nfe.ChaveNFE, isnull(doc.Descricao,'') AS Descricao" & vbCrLf & _
        '                                    "FROM NFeRealizadas nfe" & vbCrLf & _
        '                                    "	inner join Clientes c" & vbCrLf & _
        '                                    "			on c.Cliente_id   = nfe.Cliente_id" & vbCrLf & _
        '                                    "			and c.Endereco_Id = nfe.EndCliente_Id" & vbCrLf & _
        '                                    "	left join Documentos.dbo.Arquivo doc" & vbCrLf & _
        '                                    "			on doc.Empresa_Id     = nfe.Empresa_Id" & vbCrLf & _
        '                                    "			and doc.EndEmpresa_Id = nfe.EndEmpresa_Id" & vbCrLf & _
        '                                    "			and doc.Cliente_Id    = nfe.Cliente_Id" & vbCrLf & _
        '                                    "			and doc.EndCliente_Id = nfe.EndCliente_Id" & vbCrLf & _
        '                                    "			and doc.Nota_Id       = nfe.Nota_Id" & vbCrLf & _
        '                                    "			and doc.Serie_Id      = nfe.Serie_Id" & vbCrLf & _
        '                                    "			and (doc.Descricao like '%-nfe%' or doc.Descricao like '%-CTe%')" & vbCrLf & _
        '                                    "Where nfe.ChaveNFE = '" & drXML("Nome") & "'"

        '                        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "NFeRealizadas")

        '                        If ds.Tables(0).Rows.Count > 0 Then
        '                            drXML("Lancada") = "SIM"
        '                            drXML("Usuario") = ds.Tables(0).Rows(0).Item("Usuario")
        '                            drXML("DiaLancamento") = ds.Tables(0).Rows(0).Item("Hora")

        '                            If ds.Tables(0).Rows(0).Item("Descricao").ToUpper = oFileInfoCTe.Name.ToUpper Then
        '                                drXML("TemAnexo") = "SIM"
        '                            Else
        '                                drXML("TemAnexo") = ""
        '                            End If
        '                        Else
        '                            drXML("Lancada") = ""
        '                            drXML("Usuario") = ""
        '                            drXML("DiaLancamento") = ""
        '                            drXML("TemAnexo") = ""
        '                        End If

        '                        CType(Session("objXML" & HID.Value), DataTable).Rows.Add(drXML)
        '                    End If
        '                End If
        '            End If
        '        End If
        '    Next
        'End If

        If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
            Dim ds1 As DataView

            If rbPendente.Checked Then
                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = ''").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = '' or Atencao = 'S'").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para visualização.")
                    Exit Sub
                End If
            ElseIf rbLancados.Checked Then
                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM'").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM'").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para visualização.")
                    Exit Sub
                End If
            ElseIf rbSemAnexo.Checked Then

                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM' and TemAnexo = ''").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM' and TemAnexo = ''").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para visualização.")
                    Exit Sub
                End If
            Else
                If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable))
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para visualização.")
                    Exit Sub
                End If
            End If

            ds1.Sort = "Emissao DESC"
            'ds1.Sort = "ClienteNome"

            lblRegistros.Text = "Total de Registros: " & ds1.Table.Rows.Count

            gridXML.DataSource = ds1
            gridXML.DataBind()

            'For j As Integer = gridXML.Rows.Count - 1 To 0 Step -1
            '    For Each n In CType(Session("objXML" & HID.Value), DataTable).Rows
            '        If gridXML.Rows(j).Cells(3).Text = n("Chave") AndAlso n("Atencao") = "S" Then

            '            CType(gridXML.Rows(j).FindControl("imgInfXML"), ImageButton).Visible = True

            '            If n("Tipo") = "NFe" Then
            '                CType(gridXML.Rows(j).FindControl("imgInfXML"), ImageButton).ToolTip = "Nota Fiscal Cancelada pelo Emitente, verifique que a mesma está lançada no Sistema."
            '            Else
            '                CType(gridXML.Rows(j).FindControl("imgInfXML"), ImageButton).ToolTip = "Conhecimento de Transporte Cancelado pelo Emitente, verifique que o mesmo está lançado no Sistema."
            '            End If

            '        Else
            '            CType(gridXML.Rows(j).FindControl("imgInfXML"), ImageButton).Visible = False
            '        End If
            '    Next
            'Next j


            'For x As Integer = 0 To gridXML.Rows.Count - 1
            '    For Each n In CType(Session("objXML" & HID.Value), DataTable).Rows
            '        If gridXML.Rows(x).Cells(3).Text = n("Chave") AndAlso n("Atencao") = "S" Then

            '            CType(gridXML.Rows(x).FindControl("imgInfXML"), ImageButton).Visible = True

            '            If n("Tipo") = "NFe" Then
            '                CType(gridXML.Rows(x).FindControl("imgInfXML"), ImageButton).ToolTip = "Nota Fiscal Cancelada pelo Emitente, verifique que a mesma está lançada no Sistema."
            '            Else
            '                CType(gridXML.Rows(x).FindControl("imgInfXML"), ImageButton).ToolTip = "Conhecimento de Transporte Cancelado pelo Emitente, verifique que o mesmo está lançado no Sistema."
            '            End If

            '        End If
            '    Next
            'Next

            For x As Integer = 0 To gridXML.Rows.Count - 1

                CType(gridXML.Rows(x).FindControl("imgInfXML"), ImageButton).Visible = CType(gridXML.Rows(x).FindControl("visibleImageButton"), TextBox).Text = "True"
                CType(gridXML.Rows(x).FindControl("imgInfXML"), ImageButton).ToolTip = CType(gridXML.Rows(x).FindControl("ToolTip"), TextBox).Text

            Next

        Else
            MsgBox(Me.Page, "Sem registros para visualização.")
        End If

    End Sub

    Protected Sub imgDownload_Click(ByVal sender As Object, ByVal e As CommandEventArgs)
        Try
            'Dim arquivoDeSaida As String

            'If e.CommandName.Contains("nfe") Then
            '    arquivoDeSaida = "C:\Alfasig\LeituraNFe\-xmldestinatario\-nfe\" & e.CommandName
            'Else
            '    arquivoDeSaida = "C:\Alfasig\LeituraNFe\-xmldestinatario\-cte\" & e.CommandName
            'End If

            'Response.Clear()
            'Response.ContentType = "text/xml"
            'Response.AppendHeader("Content-Disposition", "attachment; filename=" & e.CommandName)
            'Response.TransmitFile(arquivoDeSaida)
            'Response.Flush()
            'Response.End()

            Dim arquivo As String = String.Empty

            'If e.CommandName.Contains("NFe") Then
            '    arquivo = ("c:/Alfasig/LeituraNFe/-xmldestinatario/-nfe/") + e.CommandName
            'Else
            '    arquivo = ("c:/Alfasig/LeituraNFe/-xmldestinatario/-cte/") + e.CommandName
            'End If
            arquivo = ("c:/Alfasig/LeituraNFe/-download/") + e.CommandName

            Dim info As FileInfo = New FileInfo(arquivo)
            If info.Exists Then
                Response.Clear()
                Response.AddHeader("Content-Disposition", "attachment; filename=""" & e.CommandName & """")
                Response.Charset = "utf8"
                Response.Cache.SetCacheability(HttpCacheability.NoCache)
                Response.ContentType = "application/octet-stream"
                Response.Flush()
                Response.WriteFile(info.FullName)
                Response.Flush()
                Response.Close()
            Else
                MsgBox(Me.Page, "Este arquivo não está disponivel no momento.')", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgXmlNFAutomatico_Click(ByVal sender As Object, ByVal e As CommandEventArgs)
        Try
            Dim imgXmlNFAutomatico As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgXmlNFAutomatico.NamingContainer, GridViewRow)
            Session("chaveXMLautomacao") = row.Cells(3).Text

            Dim objEmpresa As New [Lib].Negocio.ClienteXEmpresa(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            Session("empresaXMLautomacao") = objEmpresa

            ucEnviarXMLEmissao.limpar()

            Popup.EnviarXMLEmissao(Me.Page, "objEnviarXMLEmissao" & HID.Value)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgPdfNFView_Click(ByVal sender As Object, ByVal e As CommandEventArgs)
        Try

            Dim imgPdfNFView As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgPdfNFView.NamingContainer, GridViewRow)
            Dim tipo = row.Cells(7).Text.ToLower()
            Dim chaveXML = row.Cells(3).Text
            Dim arquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", chaveXML))

            Dim fm As New FilesManager()
            If fm.IsConnect() Then

                If File.Exists(arquivo) Then

                    Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", chaveXML))

                Else

                    Dim msgNFE As String = String.Empty
                    Dim objNotaFiscal As NotaFiscal = New NotaFiscal()
                    Dim Sql As String
                    Dim Sqls As New ArrayList

                    objNotaFiscal.Codigo = CInt(Mid(chaveXML, 26, 9))
                    objNotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    objNotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                    objNotaFiscal.ChaveNFE = chaveXML
                    Dim ModeloNFe As String = Mid(objNotaFiscal.ChaveNFE, 21, 2)

                    If ModeloNFe.Equals("55") Then

                        If DocumentoEletronico.ImprimirNFeDanfe(objNotaFiscal, 1, msgNFE) Then
                            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.Nota)
                            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                            Else
                                MsgBox(Me.Page, "Arquivo não encontrado!")
                            End If
                        Else
                            MsgBox(Me.Page, msgNFE)
                        End If

                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If

                    ElseIf ModeloNFe.Equals("57") Then

                        If DocumentoEletronico.ImprimirCTeDacte(objNotaFiscal, 1, msgNFE) Then
                            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.CT_E)
                            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                            Else
                                MsgBox(Me.Page, "Arquivo não encontrado!")
                            End If
                        Else
                            MsgBox(Me.Page, msgNFE)
                        End If

                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-dacte{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If

                    End If

                End If

            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub Limpar(ByVal tudo As Boolean)
        Session.Remove("objEstoqueMinimo" & HID.Value)
        Session.Remove("chaveXMLautomacao" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        ucEnviarXMLEmissao.SetarHID(HID.Value)

        Dim dtXML As New DataTable("ItensXML")
        dtXML.Columns.Add("Cliente", Type.GetType("System.String"))
        dtXML.Columns.Add("ClienteNome", Type.GetType("System.String"))
        dtXML.Columns.Add("ClienteCidade", Type.GetType("System.String"))
        dtXML.Columns.Add("Chave", Type.GetType("System.String"))
        dtXML.Columns.Add("Codigo", Type.GetType("System.String"))
        dtXML.Columns.Add("Tipo", Type.GetType("System.String"))
        dtXML.Columns.Add("Emissao", Type.GetType("System.DateTime"))
        dtXML.Columns.Add("Serie", Type.GetType("System.String"))
        dtXML.Columns.Add("nNF", Type.GetType("System.String"))
        dtXML.Columns.Add("Lancada", Type.GetType("System.String"))
        dtXML.Columns.Add("Usuario", Type.GetType("System.String"))
        dtXML.Columns.Add("DiaLancamento", Type.GetType("System.String"))
        dtXML.Columns.Add("TemAnexo", Type.GetType("System.String"))
        dtXML.Columns.Add("Atencao", Type.GetType("System.String"))
        dtXML.Columns.Add("ToolTip", Type.GetType("System.String"))
        dtXML.Columns.Add("visibleImageButton", Type.GetType("System.String"))
        dtXML.Columns.Add("CFOP", Type.GetType("System.String"))
        dtXML.Columns.Add("ValorDoProduto", Type.GetType("System.String"))
        dtXML.Columns.Add("ValorTotalDaNota", Type.GetType("System.String"))
        Session("objXML" & HID.Value) = dtXML

        gridXML.DataSource = Nothing
        gridXML.DataBind()

        lblRegistros.Text = "Total de Registros: 0"

        If tudo Then

            chkNFe.Checked = True
            chkCTe.Checked = True
            rbPendente.Checked = True
            rbLancados.Checked = False
            rbSemAnexo.Checked = False
            rbTodos.Checked = False

            txtDataInicial.Text = String.Empty
            txtDataFinal.Text = String.Empty

            If ddlEmpresa.SelectedValue.Length > 0 Then
                Dim fileEmpresa As String = "C:\NGS\xmlCopy\" & ddlEmpresa.SelectedValue.Split("-")(0) & ".xml"

                If File.Exists(fileEmpresa) Then
                    Dim xmlConf As New XmlDocument
                    xmlConf.Load(fileEmpresa)

                    If Not Directory.Exists(xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("eventoDFe").InnerText) Then
                        MsgBox(Me.Page, "Configuracão do DFe da Empresa não foi encontrado. Entre em contato com o Suporte.", "~/Expedicao.aspx")
                        Exit Sub
                    End If

                    Dim dataDoUltimoArquivo = New DirectoryInfo(xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("eventoDFe").InnerText).GetFiles().OrderByDescending(Function(p) p.LastWriteTime).FirstOrDefault().LastWriteTime

                    Dim proximaAtualizacao = dataDoUltimoArquivo.AddMinutes(70)

                    lblAtualizado.Text = dataDoUltimoArquivo.ToString("dd/MM/yyyy") & " às " & dataDoUltimoArquivo.ToString("HH:mm:ss") & "."

                    lblProximaAtualizacao.Text = proximaAtualizacao.ToString("dd/MM/yyyy") & " às " & proximaAtualizacao.ToString("HH:mm:ss") & "."
                Else
                    MsgBox(Me.Page, "Arquivo de Configuracão da Empresa não foi encontrado.")
                End If
            End If
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub


    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try

            If gridXML.Rows.Count > 0 Then
                EmitirRelatorio()
            Else
                MsgBox(Me.Page, "Nenhum registro para impressão.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If ddlEmpresa.SelectedValue.Length > 0 Then
                Limpar(False)
                LerXML()
            Else
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            End If
        Catch ex As Exception
            'Dim erro = teste

            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio()
        Try
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))

            Dim ds1 As DataView

            Limpar(False)
            LerXML()

            If rbPendente.Checked Then
                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = ''").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = ''").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para impressão.")
                    Exit Sub
                End If
            ElseIf rbLancados.Checked Then
                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM'").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM'").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para impressão.")
                    Exit Sub
                End If
            ElseIf rbSemAnexo.Checked Then

                If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM' and TemAnexo = ''").Length > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM' and TemAnexo = ''").CopyToDataTable())
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para impressão.")
                    Exit Sub
                End If
            Else
                If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
                    ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable))
                Else
                    LimparGrid()
                    MsgBox(Me.Page, "Sem registros para impressão.")
                    Exit Sub
                End If
            End If

            ds1.Sort = "Emissao DESC"
            'ds1.Sort = "ClienteNome"

            Dim ds2 As DataTable = ds1.ToTable()
            Dim ds As DataSet = New DataSet()
            ds.Tables.Add(ds2)

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("RELAÇÃO DE XML(S) DE NFe/CTe")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RELAÇÃO DE XML(S) DE NFe/CTe")
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "DATA : " & String.Format("{0:dd/MM/yyyy}", Now()))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:J" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            If col.ColumnName = "Emissao" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            End If

                            columnIndex += 1
                        Next

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    ''criando colunas de totalizadores
                    'worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Bold = True
                    'worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    'worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    'worksheet.Cells(String.Format("L{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    'worksheet.Cells(String.Format("L{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    ''criando colunas de totalizadores
                    'worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Font.Bold = True
                    'worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    'worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    'worksheet.Cells(String.Format("M{0}", rowIndex)).Formula = String.Format("=SUM(M6:M{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    'worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Bold = True
                    'worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    'worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    'worksheet.Cells(String.Format("N{0}", rowIndex)).Value = String.Empty

                    ''criando colunas de totalizadores
                    'worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Bold = True
                    'worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    'worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    'worksheet.Cells(String.Format("O{0}", rowIndex)).Formula = String.Format("=SUM(O6:O{0})", rowIndex - 1)

                    ''formatando células numéricas
                    'worksheet.Cells(String.Format("M6:M{0}", rowIndex)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                    'worksheet.Cells(String.Format("N6:N{0}", rowIndex)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                    'worksheet.Cells(String.Format("O6:O{0}", rowIndex)).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

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

    'Protected Sub rbPendente_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try

    '        If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
    '            Dim ds1 As DataView

    '            If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = ''").Length > 0 Then
    '                ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = ''").CopyToDataTable())

    '            Else
    '                LimparGrid()
    '                MsgBox(Me.Page, "Sem registros para visualização.")
    '                Exit Sub
    '            End If

    '            ds1.Sort = "Emissao DESC"
    '            'ds1.Sort = "ClienteNome"

    '            lblRegistros.Text = "Total de Registros: " & ds1.Table.Rows.Count

    '            gridXML.DataSource = ds1
    '            gridXML.DataBind()

    '        End If

    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    'Protected Sub rbLancados_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
    '            Dim ds1 As DataView

    '            If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM'").Length > 0 Then
    '                ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM'").CopyToDataTable())
    '            Else
    '                LimparGrid()
    '                MsgBox(Me.Page, "Sem registros para visualização.")
    '                Exit Sub
    '            End If

    '            ds1.Sort = "Emissao DESC"
    '            'ds1.Sort = "ClienteNome"

    '            lblRegistros.Text = "Total de Registros: " & ds1.Table.Rows.Count

    '            gridXML.DataSource = ds1
    '            gridXML.DataBind()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    'Protected Sub rbSemAnexo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
    '            Dim ds1 As DataView

    '            If CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM' and TemAnexo = ''").Length > 0 Then
    '                ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable).Select("Lancada = 'SIM' and TemAnexo = ''").CopyToDataTable())
    '            Else
    '                LimparGrid()
    '                MsgBox(Me.Page, "Sem registros para visualização.")
    '                Exit Sub
    '            End If

    '            ds1.Sort = "Emissao DESC"
    '            'ds1.Sort = "ClienteNome"

    '            lblRegistros.Text = "Total de Registros: " & ds1.Table.Rows.Count

    '            gridXML.DataSource = ds1
    '            gridXML.DataBind()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    'Protected Sub rbTodos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
    '            Dim ds1 As DataView

    '            If CType(Session("objXML" & HID.Value), DataTable).Rows.Count > 0 Then
    '                ds1 = New DataView(CType(Session("objXML" & HID.Value), DataTable))
    '            Else
    '                LimparGrid()
    '                MsgBox(Me.Page, "Sem registros para visualização.")
    '                Exit Sub
    '            End If

    '            ds1.Sort = "Emissao DESC"
    '            'ds1.Sort = "ClienteNome"

    '            lblRegistros.Text = "Total de Registros: " & ds1.Table.Rows.Count

    '            gridXML.DataSource = ds1
    '            gridXML.DataBind()
    '        End If
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    Private Sub LimparGrid()
        gridXML.DataSource = Nothing
        gridXML.DataBind()
    End Sub
End Class