Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Cfop
    Inherits BasePage

    Dim ds As New DataSet
    Dim Sql As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Cfop", "ACESSAR") Then
                CarregarAbaTitulo()
                Limpar()
                LimparCFOP()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub grdAbaTitulo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAbaTitulo.SelectedIndexChanged
        Dim pCodigo As String = grdAbaTitulo.SelectedRow.Cells(1).Text()
        Dim obj As [Lib].Negocio.GrupoCFOP = New [Lib].Negocio.GrupoCFOP(pCodigo)
        If (obj IsNot Nothing) Then
            txtCodTitulo.Text = obj.Codigo
            txtDesTitulo.Text = obj.Descricao
            lblTitulo.Text = String.Format("Título: {0} - {1}", obj.Codigo, obj.Descricao)
            grdAbaFiscal.DataSource = CarregaCfop(obj.Codigo)
            grdAbaFiscal.DataBind()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        End If
    End Sub

    Protected Sub grdAbaFiscal_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAbaFiscal.SelectedIndexChanged
        Dim pCodigo As String = grdAbaFiscal.SelectedRow.Cells(1).Text()
        Dim obj As [Lib].Negocio.CFOP = New [Lib].Negocio.CFOP(pCodigo)
        If (obj IsNot Nothing) Then
            txtReduzidoCfop.Text = obj.Codigo
            txtDescricaoCfop.Text = obj.Descricao
            lnkNovoCfop.Parent.Visible = False
            lnkAtualizarCfop.Parent.Visible = True
            lnkExcluirCfop.Parent.Visible = True
        End If
    End Sub

    Function CarregaDados() As DataSet
        Dim sql As String
        sql = "Select Convert(VarChar,GrupoCfop_Id) as Codigo,Descricao from CfopTitulo"
        ds = Banco.ConsultaDataSet(sql, "CfopTitulo")
        Return ds
    End Function

    Function InserirCfop(ByVal GrupoCfop_Id As String, ByVal Cfop_Id As String, ByVal Descricao As String)
        Dim sql As String
        Dim arr As New ArrayList
        'Dim AuxProducao As String

        sql = "Insert into Cfop (Cfop_Id,GrupoCfop_Id,Descricao) values (" & Cfop_Id & "," & GrupoCfop_Id & ",'" & Descricao & "')"
        arr.Add(sql)
        Banco.GravaBanco(arr)
        Return sql
    End Function

    Function AlteracaoCfop(ByVal GrupoCfop_Id As String, ByVal Cfop_Id As String, ByVal Descricao As String)
        Dim sql As String
        Dim arr As New ArrayList
        'Dim AuxProducao As String

        sql = "Update Cfop Set Descricao = '" & Descricao & "'  Where GrupoCfop_Id = '" & GrupoCfop_Id & "' and Cfop_Id =" & Cfop_Id & ""
        arr.Add(sql)
        Banco.GravaBanco(arr)
        Return sql
    End Function

    Function ExcluirCfop(ByVal GrupoCfop_Id As String, ByVal Cfop_Id As String)
        Dim sql As String
        Dim arr As New ArrayList
        sql = "Delete from Cfop Where GrupoCfop_Id = '" & GrupoCfop_Id & "' and Cfop_Id=" & Cfop_Id & ""
        arr.Add(sql)
        Banco.GravaBanco(arr)
        Return sql
    End Function

    Function CarregaCfop(ByVal GrupoCfop_Id As String) As DataSet
        Dim sql As String
        sql = "Select Convert(VarChar,Cfop_Id) as Codigo, Descricao from Cfop"
        If GrupoCfop_Id <> Nothing Then
            sql &= " Where GrupoCfop_Id=" & GrupoCfop_Id & ""
        End If
        ds = Banco.ConsultaDataSet(sql, "Cfop")
        Return ds

    End Function

    Function InserirTitulo(ByVal GrupoCfop_Id As String, ByVal Descricao As String)
        Dim sql As String
        Dim arr As New ArrayList
        'Dim AuxProducao As String

        sql = "Insert into CfopTitulo (GrupoCfop_Id,Descricao) values (" & GrupoCfop_Id & ",'" & Descricao & "')"
        arr.Add(sql)
        Banco.GravaBanco(arr)
        Return sql
    End Function

    Function AlteracaoTitulo(ByVal GrupoCfop_Id As String, ByVal Descricao As String)
        Dim sql As String
        Dim arr As New ArrayList
        'Dim AuxProducao As String

        sql = "Update CfopTitulo Set Descricao = '" & Descricao & "'  Where GrupoCfop_Id = '" & GrupoCfop_Id & "' "
        arr.Add(sql)
        Banco.GravaBanco(arr)
        Return sql
    End Function

    Function ExcluiTitulo(ByVal Operacao_Id As String) As DataSet
        Dim sql As String
        Dim arr As New ArrayList
        sql = "Delete from CfopTitulo Where GrupoCfop_Id = '" & Operacao_Id & "' "
        arr.Add(sql)
        Banco.GravaBanco(arr)
        Return Nothing
    End Function

    Private Sub RelatorioCodigoFiscal()
        Dim Empresa As String = HttpContext.Current.Session("ssEmpresa")
        Dim NomeEmpresa As String = String.Empty
        Dim CidadeEmpresa As String = String.Empty
        Dim EstadoEmpresa As String = String.Empty
        Dim GrupoCfop_Id As String = String.Empty

        Dim ds As New DataSet
        Dim sql As String
        Dim Dr As DataRow

        sql = "  SELECT Clientes.Cliente_Id AS Cliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado"
        sql &= " FROM Clientes "
        sql &= " WHERE Clientes.Cliente_Id = '" & Empresa & "'"

        ds = Banco.ConsultaDataSet(sql, "Clientes")

        If ds.Tables(0).Rows.Count > 0 Then
            For Each Dr In ds.Tables(0).Rows
                NomeEmpresa = Dr("Nome")
                CidadeEmpresa = Dr("Cidade")
                EstadoEmpresa = Dr("Estado")
                Exit For
            Next
        End If

        CidadeEmpresa &= " - " & EstadoEmpresa

        sql = " SELECT Cfop.Cfop_Id, CfopTitulo.GrupoCfop_Id, CfopTitulo.Descricao AS DescTitulo, Cfop.Descricao AS DescCodigo"
        sql &= " FROM CfopTitulo INNER JOIN"
        sql &= " Cfop ON CfopTitulo.GrupoCfop_Id = Cfop.GrupoCfop_Id"
        If GrupoCfop_Id <> Nothing Then
            sql &= " where CfopTitulo.GrupoCfop_Id=" & GrupoCfop_Id & ""
        End If

        ds = Banco.ConsultaDataSet(sql, "CodigoFiscal")

        Dim parameters = New Dictionary(Of String, Object)()
        parameters.Add("Titulo", "Relatório De Cdigo Fiscal.")

        Funcoes.BindReport(Me.Page, ds, "Cr_CodigoFiscal", eExportType.PDF, parameters)
    End Sub

    Protected Sub btnSairTitulo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Response.Redirect("~/Gestao.aspx")
    End Sub

    Protected Sub txtCodTitulo_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        CarregarAbaTitulo()
    End Sub

    Protected Sub btnSairCfopTitulo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Response.Redirect("~/Gestao.aspx")
    End Sub

#Region "Fiscal"

    Private Sub CarregarAbaFiscal()
        grdAbaFiscal.DataSource = CarregaCfop(txtCodTitulo.Text.Trim())
        grdAbaFiscal.DataBind()
    End Sub

    Sub NovoCodigoFiscal()

        Dim n As New [Lib].Negocio.CFOP()
        n.IUD = "I"
        n.CodigoGrupo = Convert.ToInt32(txtCodTitulo.Text)
        n.Codigo = Convert.ToInt32(txtReduzidoCfop.Text)
        n.Descricao = txtDescricaoCfop.Text
        n.Salvar()

    End Sub

    Sub AtualizaCodigoFiscal()
        Dim n As New [Lib].Negocio.CFOP()
        n.IUD = "U"
        n.CodigoGrupo = Convert.ToInt32(txtCodTitulo.Text)
        n.Codigo = Convert.ToInt32(txtReduzidoCfop.Text)
        n.Descricao = txtDescricaoCfop.Text
        n.Salvar()
    End Sub

    Sub ExcluirCodigoFiscal()
        Dim n As New [Lib].Negocio.CFOP()
        n.IUD = "D"
        n.Codigo = Convert.ToInt32(txtReduzidoCfop.Text)
        n.Salvar()
    End Sub

    Function ValidarCampoCodigoFiscal() As Boolean
        If (String.IsNullOrWhiteSpace(txtReduzidoCfop.Text)) Then
            MsgBox(Me.Page, "Informe o codigo.")
            Return False
        End If

        If (String.IsNullOrWhiteSpace(txtDescricaoCfop.Text)) Then
            MsgBox(Me.Page, "Informe a descrição.")
            Return False
        End If

        Return True
    End Function

    Sub LimpaCamposCodigoFiscal()
        txtReduzidoCfop.Text = String.Empty
        txtDescricaoCfop.Text = String.Empty
    End Sub

#End Region

#Region "Título"

    Private Sub CarregarAbaTitulo()
        grdAbaTitulo.DataSource = CarregaDados()
        grdAbaTitulo.DataBind()
    End Sub

    Private Sub NovoTitulo()

        Dim t As New [Lib].Negocio.GrupoCFOP()
        t.IUD = "I"
        t.Codigo = Convert.ToInt32(txtCodTitulo.Text)
        t.Descricao = txtDesTitulo.Text
        t.Salvar()

    End Sub

    Private Sub AtualizaTitulo()
        Dim t As New [Lib].Negocio.GrupoCFOP()
        t.IUD = "U"
        t.Codigo = Convert.ToInt32(txtCodTitulo.Text)
        t.Descricao = txtDesTitulo.Text
        t.Salvar()
    End Sub

    Private Sub ExcluirTitulo()
        Dim t As New [Lib].Negocio.GrupoCFOP()
        t.IUD = "D"
        t.Codigo = Convert.ToInt32(txtCodTitulo.Text)
        t.Salvar()
    End Sub

    Function ValidarCampoTitulo() As Boolean
        If (String.IsNullOrWhiteSpace(txtCodTitulo.Text)) Then
            MsgBox(Me.Page, "Informe o codigo.")
            Return False
        End If

        If (String.IsNullOrWhiteSpace(txtDesTitulo.Text)) Then
            MsgBox(Me.Page, "Informe a descrição.")
            Return False
        End If

        Return True
    End Function

    Private Sub LimpaCamposTitulo()
        txtCodTitulo.Text = String.Empty
        txtDesTitulo.Text = String.Empty
    End Sub

    Private Sub RelatorioTitulo()
        Try
            If Funcoes.VerificaPermissao("Cfop", "RELATORIO") Then
                Dim Ds As New DataSet
                Sql = "SELECT GrupoCfop_Id As Codigo, Descricao FROM CfopTitulo order by GrupoCfop_Id"
                Ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Códigos Fiscais.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")

                Funcoes.BindReport(Me.Page, Ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Private Sub Limpar()
        txtCodTitulo.Text = ""
        txtDesTitulo.Text = ""
        txtReduzidoCfop.Text = ""
        txtDescricaoCfop.Text = ""
        lblTitulo.Text = ""
        grdAbaFiscal.DataSource = New List(Of Object)
        grdAbaFiscal.DataBind()
        CarregarAbaTitulo()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Sub LimparCFOP()
        txtReduzidoCfop.Text = ""
        txtDescricaoCfop.Text = ""
        lblTitulo.Text = ""
        grdAbaFiscal.DataSource = New List(Of Object)
        grdAbaFiscal.DataBind()
        lnkNovoCfop.Parent.Visible = True
        lnkAtualizarCfop.Parent.Visible = False
        lnkExcluirCfop.Parent.Visible = False
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("GrupoCfop", "GRAVAR") Then
                If ValidarCampoTitulo() Then
                    NovoTitulo()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("GrupoCfop", "ALTERAR") Then
                If grdAbaTitulo.SelectedIndex <> -1 Then   'Verifica se tem registro selecionado.'
                    If ValidarCampoTitulo() Then
                        AtualizaTitulo()
                        MsgBox(Me.Page, "Registro atualizado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                    End If
                Else
                    MsgBox(Me.Page, "Selecione um registro para realizar a alteração.")
                    Exit Sub
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("GrupoCfop", "EXCLUIR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                If grdAbaTitulo.SelectedIndex <> -1 Then
                    ExcluirTitulo()
                    MsgBox(Me.Page, "Registro excluido com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, "Selecione um registro para realizar a alteração.")
                    Exit Sub
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        RelatorioTitulo()
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Cfop")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    'CFOP
    Protected Sub lnkNovoCfop_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoCfop.Click
        Try
            If Funcoes.VerificaPermissao("Cfop", "GRAVAR") Then
                If ValidarCampoCodigoFiscal() Then
                    NovoCodigoFiscal()
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para gravar registro.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizarCfop_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizarCfop.Click
        Try
            If Funcoes.VerificaPermissao("Cfop", "ALTERAR") Then
                If grdAbaFiscal.SelectedIndex <> -1 Then
                    If ValidarCampoCodigoFiscal() Then
                        AtualizaCodigoFiscal()
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        LimparCFOP()
                    End If
                Else
                    MsgBox(Me.Page, "Selecione um registro para realizar a alteração.")
                    Exit Sub
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para alterar registro.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirCfop_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirCfop.Click
        Try
            If Funcoes.VerificaPermissao("Cfop", "EXCLUIR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                If grdAbaFiscal.SelectedIndex <> -1 Then
                    ExcluirCodigoFiscal()
                    MsgBox(Me.Page, "Registro excluido com Sucesso.", eTitulo.Sucess)
                    LimparCFOP()
                Else
                    MsgBox(Me.Page, "Selecione um registro para realizar a alteração.")
                    Exit Sub
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para excluir resgistro.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparCfop_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparCfop.Click
        LimparCFOP()
    End Sub

    Protected Sub lnkRelatorioCfop_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorioCfop.Click
        Try
            If Funcoes.VerificaPermissao("Cfop", "RELATORIO") Then
                RelatorioCodigoFiscal()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class