Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class ObservacoesTributarias
    Inherits BasePage

    Private Sql As String
    Private objObsTrib As ObservacaoTributariaGeral

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ObservacoesTributarias", "ACESSAR") Then
                    ddl.Carregar(ddlEstado, CarregarDDL.Tabela.Estados, "", True)
                    ddl.Carregar(ddlEncargos, CarregarDDL.Tabela.Encargos, "operacaoxencargo = 'S' and Encargo_id NOT in('LIQUIDO','PIS')", True)
                    carregarGrid()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SessaoSalvaObsTrib()
        Session("objObsTrib") = objObsTrib
    End Sub

    Private Sub SessaoRecuperaObsTrib()
        objObsTrib = CType(Session("objObsTrib"), ObservacaoTributariaGeral)
    End Sub

    Private Sub carregarGrid()
        Dim lst As New ListObservacaoTributariaGeral(ddlEstado.SelectedValue, ddlEncargos.SelectedValue)
        gridObsTrib.DataSource = lst
        gridObsTrib.DataBind()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        ddlEstado.SelectedIndex = 0
        txtdescricao.Text = ""
        ddlEncargos.SelectedIndex = 0
        objObsTrib = New ObservacaoTributariaGeral()
        objObsTrib.buscarSequencia()
        txtCodigo.Text = objObsTrib.Codigo
        SessaoSalvaObsTrib()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        ddlEstado.Enabled = True
        ddlEncargos.Enabled = True
    End Sub

    Private Function validaRegistro() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código obrigatório.")
            txtCodigo.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEstado.SelectedValue) Then
            MsgBox(Me.Page, "Estado obrigatório")
            ddlEstado.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtdescricao.Text) Then
            MsgBox(Me.Page, "Descrição obrigatória.")
            txtdescricao.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEncargos.SelectedValue) Then
            MsgBox(Me.Page, "Encargo obrigatório.")
            ddlEncargos.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub gridObsTrib_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridObsTrib.SelectedIndexChanged
        Try
            Dim ObsTrib As New ObservacaoTributariaGeral(gridObsTrib.SelectedRow.Cells(1).Text)
            txtCodigo.Text = ObsTrib.Codigo
            ddlEstado.SelectedValue = ObsTrib.CodigoEstado
            ddlEncargos.SelectedValue = ObsTrib.CodigoEncargo
            txtdescricao.Text = ObsTrib.Descricao

            ddlEstado.Enabled = False
            ddlEncargos.Enabled = False

            objObsTrib = ObsTrib

            SessaoSalvaObsTrib()

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ObservacoesTributarias", "GRAVAR") Then
                If validaRegistro() Then
                    SessaoRecuperaObsTrib()
                    objObsTrib.IUD = "I"
                    objObsTrib.Codigo = txtCodigo.Text

                    Dim texto As String = Funcoes.SubstituirCaracteresEspeciais(txtdescricao.Text.Trim)

                    objObsTrib.Descricao = texto.ToUpper()
                    objObsTrib.CodigoEstado = ddlEstado.SelectedValue
                    objObsTrib.CodigoEncargo = ddlEncargos.SelectedValue
                    If objObsTrib.Salvar Then
                        MsgBox(Me.Page, "Observação tributária salva com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        carregarGrid()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("ObservacoesTributarias", "ALTERAR") Then
                SessaoRecuperaObsTrib()
                If validaRegistro() Then
                    objObsTrib.IUD = "U"
                    objObsTrib.Codigo = txtCodigo.Text

                    Dim texto As String = Funcoes.SubstituirCaracteresEspeciais(txtdescricao.Text.Trim)

                    objObsTrib.Descricao = texto.ToUpper()

                    objObsTrib.CodigoEstado = ddlEstado.SelectedValue
                    objObsTrib.CodigoEncargo = ddlEncargos.SelectedValue
                    If objObsTrib.Salvar Then
                        MsgBox(Me.Page, "Atualizado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        carregarGrid()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ObservacoesTributarias", "EXCLUIR") Then
                SessaoRecuperaObsTrib()
                If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                    objObsTrib.IUD = "D"
                    If objObsTrib.Salvar Then
                        MsgBox(Me.Page, "Observação tributária excluída com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        carregarGrid()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
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

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ObservacoesTributarias", "RELATORIO") Then
                Dim ds As DataSet

                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Codigo_Id AS varchar))) + CAST(Codigo_Id AS varchar) as Codigo," & vbCrLf & _
                      "(cast(Estado as varchar)+ replicate(' - ',1) + cast(Descricao as varchar(200)) )as Descricao" & vbCrLf & _
                      "From ObservacoesTributarias Order by Codigo_Id" & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim Campo As String() = HttpContext.Current.Server.MapPath("").Split("\")
                Dim RootSite As String = ""
                For i As Integer = 0 To Campo.GetUpperBound(0) - 1
                    RootSite &= Campo(i) & "\"
                Next

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Observações Tributárias")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
                Funcoes.BindReport(Me.Page, ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuario sem permissão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ObservacoesTributarias")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEstado_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEstado.SelectedIndexChanged
        Try
            carregarGrid()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlEncargos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEncargos.SelectedIndexChanged
        Try
            carregarGrid()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub
End Class