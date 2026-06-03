Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Reflection

Public Class Regioes
    Inherits BasePage

    Dim Sql As String
    Dim DS As DataSet
    Dim ListRegioes As [Lib].Negocio.ListRegioes
    Dim objRegiao As [Lib].Negocio.Regioes
    Dim ListMicroRegioes As [Lib].Negocio.ListMicroRegiao
    Dim objMicroRegiao As [Lib].Negocio.MicroRegiao

#Region "Session"
    Private Sub SessaoSalvaListaRegiao()
        Session("ListaRegioes" + HID.Value.ToString) = ListRegioes
    End Sub

    Private Sub SessaoRecuperaListaRegiao()
        ListRegioes = Session("ListaRegioes" + HID.Value.ToString)
    End Sub

    Private Sub SessaoSalvaListaMicroRegiao()
        Session("ListaMicroRegioes" + HID.Value.ToString) = ListMicroRegioes
    End Sub

    Private Sub SessaoRecuperaListaMicroRegiao()
        ListMicroRegioes = Session("ListaMicroRegioes" + HID.Value.ToString)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Regioes", "ACESSAR") Then
                    tcRegioes.ActiveTabIndex = 0
                    HID.Value = Guid.NewGuid.ToString
                    ListRegioes = New ListRegioes(True)
                    SessaoSalvaListaRegiao()
                    CarregaRegioesNoGrid()
                    Limpar()
                    LimparMicro()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregaRegioesNoGrid()
        SessaoRecuperaListaRegiao()
        GridRegioes.DataSource = ListRegioes.ToArray
        GridRegioes.DataBind()
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridRegioes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridRegioes.SelectedIndexChanged
        Try
            SessaoRecuperaListaRegiao()
            ListMicroRegioes = ListRegioes.Find(Function(s) s.Codigo = GridRegioes.SelectedRow.Cells(1).Text()).MicroRegioes
            objRegiao = ListRegioes.Find(Function(s) s.Codigo = GridRegioes.SelectedRow.Cells(1).Text())
            txtCodigo.Text = objRegiao.Codigo
            txtDescricao.Text = objRegiao.Descricao
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lblRegiao.Text = objRegiao.Descricao
            CarregaMicroRegiaoNoGrid()
            SessaoSalvaListaMicroRegiao()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregaMicroRegiaoNoGrid()
        SessaoRecuperaListaRegiao()
        ListMicroRegioes = ListRegioes.Find(Function(s) s.Codigo = GridRegioes.SelectedRow.Cells(1).Text()).MicroRegioes
        If Not ListMicroRegioes Is Nothing Then
            grdMicroRegiao.DataSource = ListMicroRegioes.ToArray
            grdMicroRegiao.DataBind()
        End If
        SessaoSalvaListaMicroRegiao()
    End Sub

    Protected Sub grdMicroRegiao_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles grdMicroRegiao.SelectedIndexChanged
        Try
            SessaoRecuperaListaMicroRegiao()
            objMicroRegiao = ListMicroRegioes.Find(Function(s) s.Codigo = grdMicroRegiao.SelectedRow.Cells(1).Text)
            txtCodigoMicroRegiao.Text = objMicroRegiao.Codigo
            txtDescricaoMicroRegiao.Text = objMicroRegiao.Descricao
            lnkNovoMicroRegiao.Parent.Visible = False
            lnkAtualizarMicroRegiao.Parent.Visible = True
            lnkExcluirMicroRegiao.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LimparMicro()
        txtCodigoMicroRegiao.Text = ""
        txtDescricaoMicroRegiao.Text = ""
        lnkNovoMicroRegiao.Parent.Visible = True
        lnkAtualizarMicroRegiao.Parent.Visible = False
        lnkExcluirMicroRegiao.Parent.Visible = False
    End Sub

    'REGIAO
    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Regioes", "GRAVAR") Then
                SessaoRecuperaListaRegiao()
                Dim Cod As Integer = ListRegioes.Last().Codigo + 1
                objRegiao = New [Lib].Negocio.Regioes()
                objRegiao.IUD = "I"
                objRegiao.Codigo = Cod
                objRegiao.Descricao = txtDescricao.Text
                objRegiao.Salvar()
                ListRegioes.Add(objRegiao)
                CarregaRegioesNoGrid()
                CarregaMicroRegiaoNoGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Regioes", "ALTERAR") Then
                SessaoRecuperaListaRegiao()
                objRegiao = ListRegioes.Find(Function(s) s.Codigo = GridRegioes.SelectedRow.Cells(1).Text())
                objRegiao.IUD = "U"
                objRegiao.Descricao = txtDescricao.Text
                objRegiao.Salvar()
                CarregaRegioesNoGrid()
                SessaoSalvaListaRegiao()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Regioes", "EXCLUIR") Then
                SessaoRecuperaListaRegiao()
                objRegiao = ListRegioes.Find(Function(s) s.Codigo = GridRegioes.SelectedRow.Cells(1).Text())
                objRegiao.IUD = "D"
                ListRegioes.Remove(objRegiao)
                CarregaRegioesNoGrid()
                If Not objRegiao.Salvar() Then
                    MsgBox(Me.Page, "Se existirem micro regiões, exclua-as antes de excluir esta região.")
                End If
                Limpar()
                SessaoSalvaListaRegiao()
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
            If Funcoes.VerificaPermissao("Regioes", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Regiao_Id AS varchar))) + CAST(Regiao_Id AS varchar) as Codigo, Descricao From Regioes Order by Regiao_Id"
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório de regiões.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")

                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Regioes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'MICRO REGIAO
    Protected Sub lnkNovoMicroRegiao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoMicroRegiao.Click
        Try
            If Funcoes.VerificaPermissao("Regioes", "GRAVAR") Then
                If Not String.IsNullOrWhiteSpace(txtDescricaoMicroRegiao.Text) Then
                    SessaoRecuperaListaRegiao()
                    SessaoRecuperaListaMicroRegiao()
                    Dim Cod As Integer
                    If ListRegioes.Find(Function(s) s.Codigo = GridRegioes.SelectedRow.Cells(1).Text()).MicroRegioes.Count > 0 Then
                        Cod = ListMicroRegioes.Last().Codigo + 1
                        ListMicroRegioes = ListRegioes.Find(Function(s) s.Codigo = GridRegioes.SelectedRow.Cells(1).Text()).MicroRegioes
                    Else
                        Cod = 1
                    End If

                    objMicroRegiao = New MicroRegiao(ListRegioes.Find(Function(s) s.Codigo = GridRegioes.SelectedRow.Cells(1).Text()))
                    objMicroRegiao.IUD = "I"
                    objMicroRegiao.Codigo = Cod
                    objMicroRegiao.Descricao = txtDescricaoMicroRegiao.Text
                    objMicroRegiao.Salvar()
                    ListMicroRegioes.Add(objMicroRegiao)
                    CarregaMicroRegiaoNoGrid()
                    LimparMicro()
                    SessaoSalvaListaMicroRegiao()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizarMicroRegiao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizarMicroRegiao.Click
        Try
            If Funcoes.VerificaPermissao("Regioes", "ALTERAR") Then
                If Not String.IsNullOrWhiteSpace(txtDescricaoMicroRegiao.Text) Then
                    SessaoRecuperaListaMicroRegiao()
                    objMicroRegiao = ListMicroRegioes.Find(Function(s) s.Codigo = grdMicroRegiao.SelectedRow.Cells(1).Text)
                    objMicroRegiao.IUD = "U"
                    objMicroRegiao.Descricao = txtDescricaoMicroRegiao.Text
                    objMicroRegiao.Salvar()
                    SessaoSalvaListaMicroRegiao()
                    CarregaMicroRegiaoNoGrid()
                    LimparMicro()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirMicroRegiao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirMicroRegiao.Click
        Try
            If Funcoes.VerificaPermissao("Regioes", "EXCLUIR") Then
                If Not String.IsNullOrWhiteSpace(txtCodigoMicroRegiao.Text) Then
                    SessaoRecuperaListaMicroRegiao()
                    objMicroRegiao = ListMicroRegioes.Find(Function(s) s.Codigo = grdMicroRegiao.SelectedRow.Cells(1).Text)
                    objMicroRegiao.IUD = "D"
                    objMicroRegiao.Salvar()
                    ListMicroRegioes.Remove(objMicroRegiao)
                    CarregaMicroRegiaoNoGrid()
                    SessaoSalvaListaMicroRegiao()
                    LimparMicro()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparMicroRegiao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparMicroRegiao.Click
        Try
            LimparMicro()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class