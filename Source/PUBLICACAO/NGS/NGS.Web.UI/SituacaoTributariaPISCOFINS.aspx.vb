Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports NGS.Lib
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class SituacaoTributariaPISCOFINS
    Inherits BasePage

    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINS", "ACESSAR") Then
                LimparPISCOFINS()
                LimparPISCOFINSObs()
                CarregarSTPISCOFINS()
                tcPISCOFINS.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarSTPISCOFINS()
        If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINS", "LEITURA") Then
            Dim objSituacaoTributariaPISCOFINS As New [Lib].Negocio.ListSituacaoTributariaPISCOFINS(True)
            GridSTPISCOFINS.DataSource = objSituacaoTributariaPISCOFINS.ToArray
            GridSTPISCOFINS.DataBind()
        End If
    End Sub

    Private Sub LimparPISCOFINS()
        Me.txtCodigo.Text = ""
        Me.txtDescricao.Text = ""
        Me.txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código da Situação Tributária do PIS COFINS não foi informado!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição da Situação Tributária do PISCOFINS não foi informada!")
            Return False
        Else
            Return True
        End If
    End Function

    Function ValidarCamposObs() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigoObs.Text) Then
            MsgBox(Me.Page, "Código da Observação da Situação Tributária do PIS COFINS não foi informado!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricaoObs.Text) Then
            MsgBox(Me.Page, "Observação da Situação Tributária do PISCOFINS não foi informada!")
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub GridSTPISCOFINS_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles GridSTPISCOFINS.SelectedIndexChanged
        Try
            LimparPISCOFINS()
            txtCodigo.Text = GridSTPISCOFINS.SelectedRow.Cells(1).Text().Trim()
            lblPISCOFINS.Text = GridSTPISCOFINS.SelectedRow.Cells(1).Text().Trim() & " - " & GridSTPISCOFINS.SelectedRow.Cells(2).Text().Trim()
            Dim obj As New [Lib].Negocio.SituacaoTributariaPISCOFINS(txtCodigo.Text.Trim())
            If (obj IsNot Nothing) Then
                txtCodigo.Enabled = False
                txtDescricao.Text = obj.Descricao
                txtCodigo.Focus()
                CarregarSTPISCOFINSObs()
                tcPISCOFINS.ActiveTabIndex = 1
                lnkNovo.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINS", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim objSituacaoTributariaPISCOFINS As New [Negocio].SituacaoTributariaPISCOFINS()
                    objSituacaoTributariaPISCOFINS.Codigo = txtCodigo.Text.Trim
                    objSituacaoTributariaPISCOFINS.Descricao = txtDescricao.Text.Trim
                    objSituacaoTributariaPISCOFINS.IUD = "I"
                    If objSituacaoTributariaPISCOFINS.Salvar Then
                        LimparPISCOFINS()
                        CarregarSTPISCOFINS()
                        MsgBox(Me.Page, "Sucesso ao gravar registro.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINS", "ALTERAR") Then
                If ValidarCampos() Then
                    Dim objSituacaoTributariaPISCOFINS As New [Negocio].SituacaoTributariaPISCOFINS()
                    objSituacaoTributariaPISCOFINS.Codigo = txtCodigo.Text.Trim
                    objSituacaoTributariaPISCOFINS.Descricao = txtDescricao.Text.Trim
                    objSituacaoTributariaPISCOFINS.IUD = "U"
                    If objSituacaoTributariaPISCOFINS.Salvar Then
                        LimparPISCOFINS()
                        CarregarSTPISCOFINS()
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINS", "EXCLUIR") Then
                If ValidarCampos() Then

                    Dim objSituacaoTributariaPISCOFINS As New [Negocio].SituacaoTributariaPISCOFINS(txtCodigo.Text)

                    objSituacaoTributariaPISCOFINS.Descricao = txtDescricao.Text.Trim

                    objSituacaoTributariaPISCOFINS.IUD = "D"
                    If objSituacaoTributariaPISCOFINS.Salvar Then
                        LimparPISCOFINS()
                        CarregarSTPISCOFINS()
                        MsgBox(Me.Page, "Registro excluído com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, "Registro não pode ser excluído. Obs: Registro em uso")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para excluir registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparPISCOFINS()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    '************************************ PISCOFINSObs **************************************************'
    Private Sub LimparPISCOFINSObs()
        Me.txtCodigoObs.Text = ""
        Me.txtDescricaoObs.Text = ""
        Me.txtCodigo.Enabled = True
        lnkAtualizarObs.Parent.Visible = False
        lnkNovoObs.Parent.Visible = True
        lnkExcluirObs.Parent.Visible = False
    End Sub

    Private Sub CarregarSTPISCOFINSObs()
        If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINSObs", "LEITURA") Then
            Dim p As New [Lib].Negocio.SituacaoTributariaPISCOFINS(txtCodigo.Text.Trim)
            Dim objSituacaoTributariaPISCOFINSObs As New [Lib].Negocio.ListSituacaoTributariaPISCOFINSObs(p)
            GridSTPISCOFINSObs.DataSource = objSituacaoTributariaPISCOFINSObs.ToArray
            GridSTPISCOFINSObs.DataBind()
        End If
    End Sub

    Protected Sub lnkNovoObs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoObs.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINSObs", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim p As New Negocio.SituacaoTributariaPISCOFINS(txtCodigo.Text)

                    Dim objSituacaoTributariaPISCOFINSObs As New [Negocio].SituacaoTributariaPISCOFINSObs(p)
                    objSituacaoTributariaPISCOFINSObs.CodigoObservacao = txtCodigoObs.Text.Trim
                    objSituacaoTributariaPISCOFINSObs.Descricao = txtDescricaoObs.Text.Trim
                    objSituacaoTributariaPISCOFINSObs.IUD = "I"
                    If objSituacaoTributariaPISCOFINSObs.Salvar Then
                        LimparPISCOFINSObs()
                        CarregarSTPISCOFINSObs()
                        MsgBox(Me.Page, "Registro gravado com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para gravar Registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizarObs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizarObs.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINSObs", "ALTERAR") Then
                If ValidarCamposObs() Then
                    Dim p As New Negocio.SituacaoTributariaPISCOFINS(txtCodigo.Text.Trim)

                    Dim objSituacaoTributariaPISCOFINSObs As New [Lib].Negocio.SituacaoTributariaPISCOFINSObs(p)
                    objSituacaoTributariaPISCOFINSObs.CodigoObservacao = txtCodigoObs.Text.Trim
                    objSituacaoTributariaPISCOFINSObs.Descricao = txtDescricaoObs.Text.Trim

                    objSituacaoTributariaPISCOFINSObs.IUD = "U"
                    If objSituacaoTributariaPISCOFINSObs.Salvar Then
                        LimparPISCOFINSObs()
                        CarregarSTPISCOFINSObs()
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirObs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirObs.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINSObs", "EXCLUIR") Then
                If ValidarCamposObs() Then

                    Dim p As New Negocio.SituacaoTributariaPISCOFINS(txtCodigo.Text)
                    Dim objSituacaoTributariaPISCOFINSObs As New [Negocio].SituacaoTributariaPISCOFINSObs(p)

                    objSituacaoTributariaPISCOFINSObs.Descricao = txtDescricao.Text.Trim
                    objSituacaoTributariaPISCOFINSObs.CodigoObservacao = txtCodigoObs.Text.Trim

                    objSituacaoTributariaPISCOFINSObs.IUD = "D"
                    If objSituacaoTributariaPISCOFINSObs.Salvar Then
                        LimparPISCOFINSObs()
                        CarregarSTPISCOFINSObs()
                        MsgBox(Me.Page, "Registro excluído com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, "Registro não pode ser excluído. Obs: Registro em uso.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para excluir Registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridSTPISCOFINSObs_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        Try
            LimparPISCOFINSObs()
            txtCodigoObs.Text = GridSTPISCOFINSObs.SelectedRow.Cells(1).Text().Trim()
            Dim obj As New [Lib].Negocio.SituacaoTributariaPISCOFINSObs(txtCodigoObs.Text.Trim(), txtCodigo.Text.Trim())
            If (obj IsNot Nothing) Then
                txtDescricaoObs.Text = obj.Descricao
                txtCodigoObs.Focus()
                lnkAtualizarObs.Parent.Visible = True
                lnkNovoObs.Parent.Visible = False
                lnkExcluirObs.Parent.Visible = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparObs_Click1(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparObs.Click
        Try
            LimparPISCOFINSObs()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            CreateReport(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorioObs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorioObs.Click
        Try
            CreateReport(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    '*******************************************************Listagens*************************************************'

    Private Sub CreateReport(ByVal ObsVisivel As Boolean)
        Dim rpt As New ReportDocument()

        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaPISCOFINS", "RELATORIO") Then

                Dim ds As DataSet
                Dim sql As String
                sql = "SELECT SituacaoTributariaPISCOFINS_Id, Descricao FROM SituacaoTributariaPISCOFINS" & vbCrLf & _
                      "SELECT SituacaoTributariaPISCOFINS_Id, observacao_id, Descricao FROM SituacaoTributariaPISCOFINSobs"

                ds = Banco.ConsultaDataSet(sql, "Consulta")
                ds.Tables(0).TableName = "SituacaoTributariaPISCOFINS"
                ds.Tables(1).TableName = "SituacaoTributariaPISCOFINSobs"

                Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
                ds.Tables(0).Columns.Add("Imagem", GetType(System.Byte()))
                ds.Tables(0).Columns.Add("NomeEmp", GetType(String))
                ds.Tables(0).Columns.Add("CidadeEstado", GetType(String))


                For Each row As DataRow In ds.Tables(0).Rows
                    row("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                    row("NomeEmp") = HttpContext.Current.Session("ssNomeEmpresa") & " (" & Funcoes.FormatarCpfCnpj(HttpContext.Current.Session("ssEmpresa")) & ")"
                    row("CidadeEstado") = HttpContext.Current.Session("ssCidadeEmpresa") & "/" & HttpContext.Current.Session("ssEstadoEmpresa")
                Next

                rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_SituacaoTributariaPISCOFINS.rpt")
                rpt.Load(rpt.FileName, OpenReportMethod.OpenReportByDefault)


                Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)

                rpt.SetDataSource(ds)

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("ObsVisivel", ObsVisivel)

                Funcoes.BindParameters(rpt, parameters)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)

                If System.IO.File.Exists(NomeArquivo) Then
                    Funcoes.AbrirArquivo(Page, UrlArquivo)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            rpt.Close()
            rpt.Dispose()
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SituacaoTributariaPISCOFINS")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class