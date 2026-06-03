Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class OperacoesDeFretes
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("OperacoesDeFretes", "ACESSAR") Then
                pageLoad()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub grd_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grd.SelectedIndexChanged
        If grd.SelectedRow IsNot Nothing Then
            Limpar(False)
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            ddlOperacao.SelectedValue = grd.SelectedRow.Cells(1).Text().Trim()
            BuscarSubOperacoes()
            ddlSubOperacao.SelectedValue = grd.SelectedRow.Cells(2).Text().Trim()
            rdoFisica.Checked = grd.SelectedRow.Cells(4).Text() = "F"
            rdoJuridica.Checked = grd.SelectedRow.Cells(4).Text() = "J"

            If Not String.IsNullOrWhiteSpace(grd.SelectedRow.Cells(6).Text()) AndAlso grd.SelectedRow.Cells(6).Text() <> "&nbsp;" Then
                ddlOpDestino.SelectedValue = grd.SelectedRow.Cells(6).Text().Trim()
            End If

            If Not String.IsNullOrWhiteSpace(grd.SelectedRow.Cells(7).Text()) AndAlso grd.SelectedRow.Cells(7).Text() <> "&nbsp;" Then
                BuscarSubOpDestino()
                ddlSubOpDestino.SelectedValue = grd.SelectedRow.Cells(7).Text().Trim()
            End If

            'If Not String.IsNullOrWhiteSpace(grd.SelectedRow.Cells(8).Text()) AndAlso grd.SelectedRow.Cells(8).Text() <> "&nbsp;" Then
            '    ddlOpDestino.SelectedValue = grd.SelectedRow.Cells(8).Text().Trim()
            'End If

            If Not String.IsNullOrWhiteSpace(grd.SelectedRow.Cells(8).Text()) AndAlso grd.SelectedRow.Cells(8).Text() <> "&nbsp;" Then
                ddlOpAnulacao.SelectedValue = grd.SelectedRow.Cells(8).Text().Trim()
                BuscarSubOpAnulacao()
                ddlSubOpAnulacao.SelectedValue = grd.SelectedRow.Cells(9).Text().Trim()
            End If

            If Not String.IsNullOrWhiteSpace(grd.SelectedRow.Cells(10).Text()) AndAlso grd.SelectedRow.Cells(10).Text() <> "&nbsp;" Then
                ddlOpContrapartida.SelectedValue = grd.SelectedRow.Cells(10).Text().Trim()
                BuscarSubOpContrapartida()
                ddlSubOpContrapartida.SelectedValue = grd.SelectedRow.Cells(11).Text().Trim()
            End If

            Dim obj As New [Lib].Negocio.OperacoesDeFretes(grd.SelectedRow.Cells(1).Text(), grd.SelectedRow.Cells(2).Text())
            If obj IsNot Nothing Then
                rdoCirculacao.Checked = obj.TipoOperacao = eTipoOperacao.Circulacao
                rdoPrestacaoServico.Checked = obj.TipoOperacao = eTipoOperacao.PrestacaoServico
                rdoEstadia.Checked = obj.TipoOperacao = eTipoOperacao.Estadia
                If (obj.TipoOperacao = eTipoOperacao.PrestacaoServico) Then
                    divTipoServico.Visible = obj.PrestacaoDeServico IsNot Nothing
                    rdoPropria.Checked = obj.PrestacaoDeServico IsNot Nothing AndAlso obj.PrestacaoDeServico = ePrestacaoServico.Propria
                    rdoTerceiros.Checked = obj.PrestacaoDeServico IsNot Nothing AndAlso obj.PrestacaoDeServico = ePrestacaoServico.Terceiros
                End If
            End If
        End If
    End Sub

    Protected Sub ddlOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlOperacao.SelectedIndexChanged
        BuscarSubOperacoes()
    End Sub

    Protected Sub ddlOpDestino_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlOpDestino.SelectedIndexChanged
        BuscarSubOpDestino()
    End Sub

    Protected Sub ddlOpAnulacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOpAnulacao.SelectedIndexChanged
        BuscarSubOpAnulacao()
    End Sub

    Protected Sub ddlOpContrapartida_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOpContrapartida.SelectedIndexChanged
        BuscarSubOpContrapartida()
    End Sub

    Private Sub pageLoad()
        Limpar()
        BuscarOperacao()
        BuscarSubOperacoes()
        BuscarOpDestino()
        BuscarSubOpDestino()
        BuscarOpAnulacao()
        BuscarSubOpAnulacao()
        BuscarOpContrapartida()
        BuscarSubOpContrapartida()
        BindGridView()
    End Sub

    Private Sub BindGridView()
        grd.DataSource = New [Lib].Negocio.ListOperacoesDeFretes()
        grd.DataBind()
    End Sub

    Private Sub BuscarOperacao()
        Dim objOperacoes As New [Lib].Negocio.ListOperacao()
        If objOperacoes.Selecionar() Then
            For Each objOperacao As [Lib].Negocio.Operacao In objOperacoes
                ddlOperacao.Items.Add(New ListItem(objOperacao.Codigo.ToString("00") & " - " & objOperacao.Descricao, _
                                                   objOperacao.Codigo))
            Next
            If ddlOperacao IsNot Nothing Then
                For Each li As ListItem In ddlOperacao.Items
                    li.Attributes("title") = li.Text
                Next
            End If
            Funcoes.InserirLinhaEmBranco(ddlOperacao)
        End If
    End Sub

    Private Sub BuscarSubOperacoes()
        If ddlOperacao.SelectedIndex > 0 Then
            ddlSubOperacao.Items.Clear()
            Dim obj As New [Lib].Negocio.ListSubOperacao()
            If obj.Selecionar(ddlOperacao.SelectedValue) Then
                For Each objSubOperacao As [Lib].Negocio.SubOperacao In obj
                    ddlSubOperacao.Items.Add(New ListItem(objSubOperacao.Codigo.ToString("00") & " - " & objSubOperacao.Descricao, _
                                                          objSubOperacao.Codigo))
                Next
                If obj IsNot Nothing Then
                    For Each li As ListItem In ddlSubOperacao.Items
                        li.Attributes("title") = li.Text
                    Next
                End If
                Funcoes.InserirLinhaEmBranco(ddlSubOperacao)
            Else
                MsgBox(Me.Page, obj.Erro.Message)
            End If
        Else
            ddlSubOperacao.Items.Clear()
        End If
    End Sub

    Private Sub BuscarOpDestino()
        Dim objOperacoes As New [Lib].Negocio.ListOperacao()
        If objOperacoes.Selecionar() Then
            For Each objOperacao As [Lib].Negocio.Operacao In objOperacoes
                ddlOpDestino.Items.Add(New ListItem(objOperacao.Codigo.ToString("00") & " - " & objOperacao.Descricao, _
                                                   objOperacao.Codigo))
            Next
            If ddlOpDestino IsNot Nothing Then
                For Each li As ListItem In ddlOpDestino.Items
                    li.Attributes("title") = li.Text
                Next
            End If
            Funcoes.InserirLinhaEmBranco(ddlOpDestino)
        End If
    End Sub

    Private Sub BuscarSubOpDestino()
        If ddlOpDestino.SelectedIndex > 0 Then
            ddlSubOpDestino.Items.Clear()
            Dim lst As New [Lib].Negocio.ListSubOperacao()
            If lst.Selecionar(ddlOpDestino.SelectedValue) Then
                For Each obj As [Lib].Negocio.SubOperacao In lst
                    ddlSubOpDestino.Items.Add(New ListItem(obj.Codigo.ToString("00") & " - " & obj.Descricao, obj.Codigo))
                Next
                If ddlSubOpDestino IsNot Nothing Then
                    For Each li As ListItem In ddlSubOpDestino.Items
                        li.Attributes("title") = li.Text
                    Next
                End If
                Funcoes.InserirLinhaEmBranco(ddlSubOpDestino)
            Else
                MsgBox(Me.Page, lst.Erro.Message)
            End If
        Else
            ddlSubOpDestino.Items.Clear()
        End If
    End Sub

    Private Sub BuscarOpAnulacao()
        Dim objOperacoes As New [Lib].Negocio.ListOperacao()
        If objOperacoes.Selecionar() Then
            For Each objOperacao As [Lib].Negocio.Operacao In objOperacoes
                ddlOpAnulacao.Items.Add(New ListItem(objOperacao.Codigo.ToString("00") & " - " & objOperacao.Descricao, objOperacao.Codigo))
            Next
            If ddlOpAnulacao IsNot Nothing Then
                For Each li As ListItem In ddlOpAnulacao.Items
                    li.Attributes("title") = li.Text
                Next
            End If
            Funcoes.InserirLinhaEmBranco(ddlOpAnulacao)
        End If
    End Sub

    Private Sub BuscarSubOpAnulacao()
        If ddlOpAnulacao.SelectedIndex > 0 Then
            ddlSubOpAnulacao.Items.Clear()
            Dim lst As New [Lib].Negocio.ListSubOperacao()
            If lst.Selecionar(ddlOpAnulacao.SelectedValue) Then
                For Each obj As [Lib].Negocio.SubOperacao In lst
                    ddlSubOpAnulacao.Items.Add(New ListItem(obj.Codigo.ToString("00") & " - " & obj.Descricao, obj.Codigo))
                Next
                If ddlSubOpAnulacao IsNot Nothing Then
                    For Each li As ListItem In ddlSubOpAnulacao.Items
                        li.Attributes("title") = li.Text
                    Next
                End If
                Funcoes.InserirLinhaEmBranco(ddlSubOpAnulacao)
            Else
                MsgBox(Me.Page, lst.Erro.Message)
            End If
        Else
            ddlSubOpAnulacao.Items.Clear()
        End If
    End Sub

    Private Sub BuscarOpContrapartida()
        Dim objOperacoes As New [Lib].Negocio.ListOperacao()
        If objOperacoes.Selecionar() Then
            For Each objOperacao As [Lib].Negocio.Operacao In objOperacoes
                ddlOpContrapartida.Items.Add(New ListItem(objOperacao.Codigo.ToString("00") & " - " & objOperacao.Descricao, objOperacao.Codigo))
            Next
            If ddlOpContrapartida IsNot Nothing Then
                For Each li As ListItem In ddlOpContrapartida.Items
                    li.Attributes("title") = li.Text
                Next
            End If
            Funcoes.InserirLinhaEmBranco(ddlOpContrapartida)
        End If
    End Sub

    Private Sub BuscarSubOpContrapartida()
        If ddlOpContrapartida.SelectedIndex > 0 Then
            ddlSubOpContrapartida.Items.Clear()
            Dim lst As New [Lib].Negocio.ListSubOperacao()
            If lst.Selecionar(ddlOpContrapartida.SelectedValue) Then
                For Each obj As [Lib].Negocio.SubOperacao In lst
                    ddlSubOpContrapartida.Items.Add(New ListItem(obj.Codigo.ToString("00") & " - " & obj.Descricao, obj.Codigo))
                Next
                If ddlSubOpContrapartida IsNot Nothing Then
                    For Each li As ListItem In ddlSubOpContrapartida.Items
                        li.Attributes("title") = li.Text
                    Next
                End If
                Funcoes.InserirLinhaEmBranco(ddlSubOpContrapartida)
            Else
                MsgBox(Me.Page, lst.Erro.Message)
            End If
        Else
            ddlSubOpContrapartida.Items.Clear()
        End If
    End Sub

    Private Sub Limpar(Optional ByVal LimparGrid As Boolean = True)
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        ddlOperacao.SelectedIndex = 0
        ddlSubOperacao.Items.Clear()
        ddlOpDestino.SelectedIndex = 0
        ddlSubOpDestino.Items.Clear()
        ddlOpAnulacao.SelectedIndex = 0
        ddlSubOpAnulacao.Items.Clear()
        ddlOpContrapartida.SelectedIndex = 0
        ddlSubOpContrapartida.Items.Clear()
        rdoFisica.Checked = True
        rdoJuridica.Checked = False
        rdoCirculacao.Checked = True
        rdoPrestacaoServico.Checked = False
        rdoEstadia.Checked = False
        If LimparGrid Then grd.SelectedIndex = -1
    End Sub

    Private Function Validar() As Boolean
        If (String.IsNullOrWhiteSpace(ddlOperacao.SelectedValue)) Then
            MsgBox(Me.Page, "Escolha a Operação.")
            Return False
        End If
        If (String.IsNullOrWhiteSpace(ddlSubOperacao.SelectedValue)) Then
            MsgBox(Me.Page, "Escolha a SubOperação.")
            Return False
        End If
        Return True
    End Function

    Private Sub Inserir()
        Try
            Dim obj As New [Lib].Negocio.OperacoesDeFretes()
            obj.IUD = "I"
            obj.OperacaoId = CInt(ddlOperacao.SelectedValue)
            obj.SubOperacaoId = CInt(ddlSubOperacao.SelectedValue)
            If String.IsNullOrWhiteSpace(ddlOpDestino.SelectedValue) Then
                obj.OpDestinoId = New Nullable(Of Int32)
            Else
                obj.OpDestinoId = CInt(ddlOpDestino.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlSubOpDestino.SelectedValue) Then
                obj.SubOpDestinoId = New Nullable(Of Int32)
            Else
                obj.SubOpDestinoId = CInt(ddlSubOpDestino.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlOpAnulacao.SelectedValue) Then
                obj.OpAnulacao = New Nullable(Of Int32)
            Else
                obj.OpAnulacao = CInt(ddlOpAnulacao.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlSubOpAnulacao.SelectedValue) Then
                obj.SubOpAnulacao = New Nullable(Of Int32)
            Else
                obj.SubOpAnulacao = CInt(ddlSubOpAnulacao.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlOpContrapartida.SelectedValue) Then
                obj.OpContrapartida = New Nullable(Of Int32)
            Else
                obj.OpContrapartida = CInt(ddlOpContrapartida.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlSubOpContrapartida.SelectedValue) Then
                obj.SubOpContrapartida = New Nullable(Of Int32)
            Else
                obj.SubOpContrapartida = CInt(ddlSubOpContrapartida.SelectedValue)
            End If
            obj.TipoPessoaId = IIf(rdoFisica.Checked, "F", "J")
            obj.TipoOperacao = getTipoOpFrete()
            If divTipoServico.Visible AndAlso (rdoPropria.Checked OrElse rdoTerceiros.Checked) Then
                obj.PrestacaoDeServico = IIf(rdoPropria.Checked, ePrestacaoServico.Propria, ePrestacaoServico.Terceiros)
            End If

            If obj.Salvar() Then
                MsgBox(Me.Page, "Operação de frete salva com Sucesso.", eTitulo.Sucess)
                BindGridView()
                Limpar()
            Else
                MsgBox(Me.Page, Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getTipoOpFrete() As eTipoOperacao
        If rdoPrestacaoServico.Checked Then
            Return eTipoOperacao.PrestacaoServico
        ElseIf rdoEstadia.Checked Then
            Return eTipoOperacao.Estadia
        Else
            Return eTipoOperacao.Circulacao
        End If
    End Function

    Private Sub Atualizar()
        Try
            Dim obj As New [Lib].Negocio.OperacoesDeFretes()
            obj.IUD = "U"
            obj.OperacaoId = CInt(ddlOperacao.SelectedValue)
            obj.SubOperacaoId = CInt(ddlSubOperacao.SelectedValue)
            If String.IsNullOrWhiteSpace(ddlOpDestino.SelectedValue) Then
                obj.OpDestinoId = New Nullable(Of Int32)
            Else
                obj.OpDestinoId = CInt(ddlOpDestino.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlSubOpDestino.SelectedValue) Then
                obj.SubOpDestinoId = New Nullable(Of Int32)
            Else
                obj.SubOpDestinoId = CInt(ddlSubOpDestino.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlOpAnulacao.SelectedValue) Then
                obj.OpAnulacao = New Nullable(Of Int32)
            Else
                obj.OpAnulacao = CInt(ddlOpAnulacao.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlSubOpAnulacao.SelectedValue) Then
                obj.SubOpAnulacao = New Nullable(Of Int32)
            Else
                obj.SubOpAnulacao = CInt(ddlSubOpAnulacao.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlOpContrapartida.SelectedValue) Then
                obj.OpContrapartida = New Nullable(Of Int32)
            Else
                obj.OpContrapartida = CInt(ddlOpContrapartida.SelectedValue)
            End If
            If String.IsNullOrWhiteSpace(ddlSubOpContrapartida.SelectedValue) Then
                obj.SubOpContrapartida = New Nullable(Of Int32)
            Else
                obj.SubOpContrapartida = CInt(ddlSubOpContrapartida.SelectedValue)
            End If
            obj.TipoPessoaId = IIf(rdoFisica.Checked, "F", "J")
            obj.TipoOperacao = getTipoOpFrete()
            If divTipoServico.Visible AndAlso (rdoPropria.Checked OrElse rdoTerceiros.Checked) Then
                obj.PrestacaoDeServico = IIf(rdoPropria.Checked, ePrestacaoServico.Propria, ePrestacaoServico.Terceiros)
            End If

            If obj.Salvar() Then
                MsgBox(Me.Page, "Operação de frete atualizada com Sucesso.", eTitulo.Sucess)
                BindGridView()
                Limpar()
            Else
                MsgBox(Me.Page, Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Excluir()
        Try
            Dim obj As New [Lib].Negocio.OperacoesDeFretes()
            obj.IUD = "D"
            obj.OperacaoId = CInt(ddlOperacao.SelectedValue)
            obj.SubOperacaoId = CInt(ddlSubOperacao.SelectedValue)
            obj.TipoPessoaId = IIf(rdoFisica.Checked, "F", "J")
            If obj.Salvar() Then
                MsgBox(Me.Page, "Registro removido com Sucesso.", eTitulo.Sucess)
                BindGridView()
                Limpar()
            Else
                MsgBox(Me.Page, Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdoPrestacaoServico_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdoPrestacaoServico.CheckedChanged
        divTipoServico.Visible = rdoPrestacaoServico.Checked
    End Sub

    Protected Sub rdoCirculacao_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdoCirculacao.CheckedChanged
        divTipoServico.Visible = rdoPrestacaoServico.Checked
    End Sub

    Protected Sub rdoEstadia_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rdoEstadia.CheckedChanged
        divTipoServico.Visible = rdoPrestacaoServico.Checked
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            'If Funcoes.VerificaPermissao("OperacoesDeFretes", "GRAVAR") Then
            If Validar() Then Inserir()
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            'If Funcoes.VerificaPermissao("OperacoesDeFretes", "ALTERAR") Then
            If Validar() Then Atualizar()
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para atualizar registro.")
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            'If Funcoes.VerificaPermissao("OperacoesDeFretes", "EXCLUIR") Then
            Excluir()
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "OperacoesDeFretes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            'If Funcoes.VerificaPermissao("OperacoesDeFretes", "RELATORIO") Then
            Dim ds As New DataSet
            ds = getDataSet()

            If ds.Tables(0) IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                Dim dt As DataTable = ds.Tables.Add("Logo")
                dt.Columns.Add("img", GetType(System.Byte()))

                Dim drImagem As DataRow = dt.NewRow()
                Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
                drImagem("img") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                dt.Rows.Add(drImagem)

                Dim rpt As New ReportDocument()

                Try
                    rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_OperacoesDeFretes.rpt")
                    rpt.Load(rpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                    Dim Empresa As String = HttpContext.Current.Session("ssNomeEmpresa") & vbCrLf & _
                    HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")

                    Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                    Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)
                    rpt.SetDataSource(ds)

                    Dim objEmpresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Empresa", String.Format("{0} - {1} ({2})", objEmpresa.Reduzido, objEmpresa.Nome, objEmpresa.Codigo))
                    parameters.Add("Endereco", String.Format("{0}, {1}", objEmpresa.Endereco, objEmpresa.Numero))
                    parameters.Add("Cidade", String.Format("{0} - {1}/{2} - {3}", objEmpresa.CEP, objEmpresa.Cidade, objEmpresa.CodigoEstado, objEmpresa.Complemento))

                    Funcoes.BindParameters(rpt, parameters)

                    If Dir(NomeArquivo).Length > 0 Then
                        Kill(NomeArquivo)
                    End If

                    rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)

                    If System.IO.File.Exists(NomeArquivo) Then
                        Funcoes.AbrirArquivo(Page, UrlArquivo)
                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                Finally
                    rpt.Close()
                    rpt.Dispose()
                End Try
            Else
                MsgBox(Me.Page, "Não encontrado dados")
            End If
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim ds As New DataSet
        Dim sql As String = ""
        sql &= " SELECT CAST(odf.Operacao_Id as  nvarchar) + '-' + CAST(odf.SubOperacoes_Id as NVARCHAR) + ' - ' + sub.Descricao as Operacao," & vbCrLf & _
               "	    odf.TipoPessoa_Id as Tipopessoa, isnull(cast(odf.OpDestino_Id as  nvarchar) + '-' + cast(odf.SubOpDestino_Id as  nvarchar) + ' - ' + subd.Descricao, '') as OpDestino, " & vbCrLf & _
               "	    ISNULL(CAST(odf.OpAnulacao as  nvarchar) + '-' + CAST(odf.SubOperacoes_Id as NVARCHAR) + ' - ' + suba.Descricao, '') as OpAnulacao   " & vbCrLf & _
               "   FROM OperacoesDeFretes odf" & vbCrLf & _
               "  INNER JOIN Suboperacoes sub" & vbCrLf & _
               "	 ON odf.Operacao_Id = sub.Operacao_Id" & vbCrLf & _
               "	AND odf.SubOperacoes_Id = sub.SubOperacoes_Id" & vbCrLf & _
               "   LEFT JOIN SubOperacoes subd" & vbCrLf & _
               "	 ON odf.Operacao_Id = subd.Operacao_Id" & vbCrLf & _
               "	AND odf.SubOperacoes_Id = subd.SubOperacoes_Id" & vbCrLf & _
               "   LEFT JOIN SubOperacoes suba" & vbCrLf & _
               "	 ON odf.OpAnulacao = suba.Operacao_Id" & vbCrLf & _
               "	AND odf.SubOpAnulacao = suba.SubOperacoes_Id"

        ds = Banco.ConsultaDataSet(sql, "OperacoesDeFretes")
        Return ds
    End Function
End Class