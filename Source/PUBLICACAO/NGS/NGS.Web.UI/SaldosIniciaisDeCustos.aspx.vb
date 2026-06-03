Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class SaldosIniciaisDeCustos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Custos)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("SaldosIniciaisDeCustos", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                CargaDepositos()
                CargaGrupos()
                CarregarCusto()
                Limpar()
                DdlAno.SelectedValue = Now.Year()
                ddlUnidade.Focus()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Custos.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Private Sub CargaDepositos()
        ddl.Carregar(DdlDeposito, CarregarDDL.Tabela.ClientesXTipos, "3", True)
    End Sub

    Private Sub CargaGrupos()
        ddl.Carregar(DdlGrupo, CarregarDDL.Tabela.NivelGrupoProduto, "(LEN(Grupo_Id) > 2) AND Custo = 'true' order by Descricao ", True)
    End Sub

    Protected Sub DdlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProduto()
        ddl.Carregar(DdlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & DdlGrupo.SelectedValue & "'", True)
    End Sub

    Private Sub CarregarCusto()
        ddl.Carregar(DdlCusto, CarregarDDL.Tabela.PlanoDeCustos, "", True)
        DdlCusto.SelectedValue = 920
        DdlCusto.Enabled = False
    End Sub

    Private Sub Limpar()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        DdlDeposito.Enabled = True
        DdlGrupo.Enabled = True
        DdlProduto.Enabled = True
        DdlMes.Enabled = True
        DdlAno.Enabled = True

        DdlGrupo.SelectedIndex = 0
        DdlProduto.Items.Clear()
        txtQuantidade.Text = ""
        txtValorDaMercadoria.Text = ""
        txtValorDoFrete.Text = ""
        SqlArray.Clear()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Function Validar_Campos() As Boolean
        If txtQuantidade.Text.Length = 0 Then txtQuantidade.Text = 0
        If txtValorDaMercadoria.Text.Length = 0 Then txtValorDaMercadoria.Text = 0
        If txtValorDoFrete.Text.Length = 0 Then txtValorDoFrete.Text = 0

        If ddlEmpresa.SelectedIndex = 0 Then
            Mensagem = "Empreda não foi selecionada"
            Return False
        ElseIf DdlDeposito.SelectedIndex = 0 Then
            Mensagem = "Depósito não foi selecionado"
            Return False
        ElseIf DdlGrupo.SelectedIndex = 0 Then
            Mensagem = "Grupo do Produto não foi selecionado"
            Return False
        ElseIf DdlProduto.SelectedIndex = 0 Then
            Mensagem = "Produto não foi selecionado"
            Return False
        ElseIf DdlAno.SelectedIndex = 0 Then
            Mensagem = "Ano não foi selecionado"
            Return False
        ElseIf CDbl(txtQuantidade.Text.Replace(",", ".")) = 0 Then
            Mensagem = "Quantidade não foi informada"
            Return False
        ElseIf CDbl(txtValorDaMercadoria.Text.Replace(",", ".")) = 0 Then
            Mensagem = "Valor da Mercadoria não foi informada"
            Return False
        Else : Return True
        End If
    End Function

    Private Sub CargaCustos()
        If Funcoes.VerificaPermissao("SaldosIniciaisDeCustos", "LEITURA") Then
            Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")

            Sql = " SELECT     ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id, ApuracaoDeCustos.Deposito_Id, ApuracaoDeCustos.EndDeposito_Id, " & vbCrLf & _
                  "            ApuracaoDeCustos.Produto_Id, ApuracaoDeCustos.CodigoDeCusto_Id, ApuracaoDeCustos.Mes_Id, ApuracaoDeCustos.Ano_Id, " & vbCrLf & _
                  "            ApuracaoDeCustos.Quantidade, ApuracaoDeCustos.ValorDoProduto, ApuracaoDeCustos.ValorDoFrete, Produtos.Grupo" & vbCrLf & _
                  " FROM       ApuracaoDeCustos INNER JOIN" & vbCrLf & _
                  "            Produtos ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                  " WHERE      Empresa_id = '" & strEmpresa(0) & "'" & vbCrLf & _
                  "            And EndEmpresa_Id = " & strEmpresa(1) & vbCrLf & _
                  "            And Ano_Id = " & DdlAno.SelectedValue & vbCrLf & _
                  "            And CodigoDeCusto_Id = 920 " & vbCrLf & _
                  " ORDER BY   ApuracaoDeCustos.Empresa_Id, ApuracaoDeCustos.EndEmpresa_Id, ApuracaoDeCustos.Deposito_Id, ApuracaoDeCustos.EndDeposito_Id, " & vbCrLf & _
                  "            ApuracaoDeCustos.Produto_Id, ApuracaoDeCustos.CodigoDeCusto_Id, ApuracaoDeCustos.Mes_Id, ApuracaoDeCustos.Ano_Id" & vbCrLf

            GridCustos.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridCustos.DataBind()
        End If
    End Sub

    Protected Sub GridCustos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()

            DdlDeposito.SelectedValue = GridCustos.SelectedRow.Cells(3).Text() & "-" & GridCustos.SelectedRow.Cells(4).Text()
            DdlGrupo.SelectedValue = GridCustos.SelectedRow.Cells(5).Text()
            CarregarProduto()
            DdlProduto.SelectedValue = GridCustos.SelectedRow.Cells(6).Text()
            DdlCusto.SelectedValue = GridCustos.SelectedRow.Cells(7).Text()
            DdlMes.SelectedValue = GridCustos.SelectedRow.Cells(8).Text()
            DdlAno.SelectedValue = GridCustos.SelectedRow.Cells(9).Text()

            txtQuantidade.Text = GridCustos.SelectedRow.Cells(10).Text()
            txtValorDaMercadoria.Text = GridCustos.SelectedRow.Cells(11).Text()
            txtValorDoFrete.Text = GridCustos.SelectedRow.Cells(12).Text()

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True

            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
            DdlDeposito.Enabled = False
            DdlGrupo.Enabled = False
            DdlProduto.Enabled = False
            DdlMes.Enabled = False
            DdlAno.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("SaldosIniciaisDeCustos", "GRAVAR") Then
                If Validar_Campos() Then
                    Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
                    Dim strDeposito() As String = DdlDeposito.SelectedValue.Split("-")

                    Sql = "   INSERT ApuracaoDeCustos (" & vbCrLf & _
                          "  Empresa_id" & vbCrLf & _
                          ", EndEmpresa_Id" & vbCrLf & _
                          ", Deposito_Id" & vbCrLf & _
                          ", EndDeposito_Id" & vbCrLf & _
                          ", Ano_Id" & vbCrLf & _
                          ", Mes_Id" & vbCrLf & _
                          ", Produto_Id" & vbCrLf & _
                          ", CodigoDeCusto_Id" & vbCrLf & _
                          ", EmpresaDestino_Id" & vbCrLf & _
                          ", EndEmpresaDestino_Id" & vbCrLf & _
                          ", DepositoDestino_Id" & vbCrLf & _
                          ", EndDepositoDestino_Id" & vbCrLf & _
                          ", ProdutoDerivado_Id" & vbCrLf & _
                          ", Quantidade" & vbCrLf & _
                          ", ValorDoProduto" & vbCrLf & _
                          ", ValorDoFrete)" & vbCrLf & _
                          " Values(" & vbCrLf & _
                          "'" & strEmpresa(0) & "'" & vbCrLf & _
                          ",  " & strEmpresa(1) & vbCrLf & _
                          ", '" & strDeposito(0) & "'" & vbCrLf & _
                          ",  " & strDeposito(1) & vbCrLf & _
                          ",  " & DdlAno.SelectedValue & vbCrLf & _
                          ",  " & DdlMes.SelectedValue & vbCrLf & _
                          ", '" & DdlProduto.SelectedValue & "'" & vbCrLf & _
                          ",  " & DdlCusto.SelectedValue & vbCrLf & _
                          ", ''" & vbCrLf & _
                          ", 0" & vbCrLf & _
                          ", ''" & vbCrLf & _
                          ", 0" & vbCrLf & _
                          ", ''" & vbCrLf & _
                          ", " & txtQuantidade.Text.Replace(".", "").Replace(",", ".") & vbCrLf & _
                          ", " & txtValorDaMercadoria.Text.Replace(".", "").Replace(",", ".") & vbCrLf & _
                          ", " & txtValorDoFrete.Text.Replace(".", "").Replace(",", ".") & vbCrLf & _
                          ")" & vbCrLf

                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) Then
                        Limpar()
                        CargaCustos()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, Mensagem)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("SaldosIniciaisDeCustos", "ALTERAR") Then
                If Validar_Campos() Then
                    Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
                    Dim strDeposito() As String = DdlDeposito.SelectedValue.Split("-")

                    Sql = "UPDATE ApuracaoDeCustos" & vbCrLf & _
                          " Set Quantidade  = " & txtQuantidade.Text.Replace(".", "").Replace(",", ".") & vbCrLf & _
                          ", ValorDoProduto = " & txtValorDaMercadoria.Text.Replace(".", "").Replace(",", ".") & vbCrLf & _
                          ", ValorDoFrete   = " & txtValorDoFrete.Text.Replace(".", "").Replace(",", ".") & vbCrLf & _
                          " WHERE  Empresa_id = '" & strEmpresa(0) & "'" & vbCrLf & _
                          "        And EndEmpresa_Id = " & strEmpresa(1) & vbCrLf & _
                          "        And Deposito_Id = '" & strDeposito(0) & "'" & vbCrLf & _
                          "        And EndDeposito_Id = " & strDeposito(1) & vbCrLf & _
                          "        And Ano_Id = " & DdlAno.SelectedValue & vbCrLf & _
                          "        And Mes_Id = " & DdlMes.SelectedValue & vbCrLf & _
                          "        And Produto_Id = '" & DdlProduto.SelectedValue & "'" & vbCrLf & _
                          "        And CodigoDeCusto_Id = " & DdlCusto.SelectedValue & vbCrLf & _
                          "        And EmpresaDestino_Id = ''" & vbCrLf & _
                          "        And EndEmpresaDestino_Id = 0" & vbCrLf & _
                          "        And DepositoDestino_Id = ''" & vbCrLf & _
                          "        And EndDepositoDestino_Id = 0" & vbCrLf & _
                          "        And ProdutoDerivado_Id = ''" & vbCrLf

                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) Then
                        Limpar()
                        CargaCustos()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, Mensagem)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("SaldosIniciaisDeCustos", "EXCLUIR") Then
                Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
                Dim strDeposito() As String = DdlDeposito.SelectedValue.Split("-")

                Sql = "DELETE FROM ApuracaoDeCustos" & vbCrLf & _
                      " WHERE  Empresa_id = '" & strEmpresa(0) & "'" & vbCrLf & _
                      "        And EndEmpresa_Id = " & strEmpresa(1) & vbCrLf & _
                      "        And Deposito_Id = '" & strDeposito(0) & "'" & vbCrLf & _
                      "        And EndDeposito_Id = " & strDeposito(1) & vbCrLf & _
                      "        And Ano_Id = " & DdlAno.SelectedValue & vbCrLf & _
                      "        And Mes_Id = " & DdlMes.SelectedValue & vbCrLf & _
                      "        And Produto_Id = '" & DdlProduto.SelectedValue & "'" & vbCrLf & _
                      "        And CodigoDeCusto_Id = " & DdlCusto.SelectedValue & vbCrLf & _
                      "        And EmpresaDestino_Id = ''" & vbCrLf & _
                      "        And EndEmpresaDestino_Id = 0" & vbCrLf & _
                      "        And DepositoDestino_Id = ''" & vbCrLf & _
                      "        And EndDepositoDestino_Id = 0" & vbCrLf & _
                      "        And ProdutoDerivado_Id = ''" & vbCrLf

                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    Limpar()
                    CargaCustos()
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
            If ddlEmpresa.SelectedIndex = 0 Then
                Mensagem = "Empreda não foi selecionada"
            Else
                CargaCustos()
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
            Funcoes.Ajuda(Me.Page, "SaldosIniciaisDeCustos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class