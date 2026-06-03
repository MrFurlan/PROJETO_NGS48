Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PlanoDeCustos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim dsOrigem As DataSet
    Dim ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Custos.aspx", eTitulo.Info)
                    Exit Sub
                End If

                Me.setMenu(eModulo.Custos)
                If Funcoes.VerificaPermissao("PlanoDeCustos", "ACESSAR") Then
                    CarregarPlanoDeCustos()
                    CarregarTotalizador()
                    CarregarPlanoDeContas()
                    CarregarHistoricos()
                    CarregarProdutos()
                    Sinal()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Custos.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarPlanoDeCustos()
        If Funcoes.VerificaPermissao("PlanoDeCustos", "LEITURA") Then
            Sql = "SELECT Codigo_Id as Codigo, Descricao, Totalizador, SinalPeso, SinalValor, " & vbCrLf & _
                    " DebitoMercadoria, CreditoMercadoria, HistoricoMercadoria, " & vbCrLf & _
                    " DebitoFrete, CreditoFrete, HistoricoFrete, Desdobramento, SaldoInicial, case when isnull(Rateio,0) = 0 then 'False' else 'True' end Rateio " & vbCrLf & _
                    " FROM PlanoDeCustos " & vbCrLf & _
                    " ORDER BY Codigo_Id" & vbCrLf

            GridPlanoDeCustos.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridPlanoDeCustos.DataBind()
        End If
    End Sub

    Private Sub CarregarTotalizador()
        ddl.Carregar(ddlTotalizador, CarregarDDL.Tabela.PlanoDeCustos, "", True)
    End Sub

    Private Sub Sinal()
        DdlSinalPeso.Items.Add(New ListItem("+ Mais ", "+"))
        DdlSinalPeso.Items.Add(New ListItem("- Menos ", "-"))

        DdlSinalValor.Items.Add(New ListItem("+ Mais ", "+"))
        DdlSinalValor.Items.Add(New ListItem("- Menos ", "-"))

        DdlSinalPeso.Items.Insert(0, "")
        DdlSinalPeso.SelectedIndex = 0
        DdlSinalValor.Items.Insert(0, "")
        DdlSinalValor.SelectedIndex = 0
    End Sub

    Private Sub CarregarProdutos()
        Sql = "SELECT Produto_Id As Codigo, Nome as Descricao FROM Produtos order by Produto_Id"
        DdlProdutoOrigem.Items.Clear()
        DdlProdutoOrigem.Items.Insert(0, "")
        For Each row As DataRow In Banco.ConsultaDataSet(Sql, "Produtos").Tables(0).Rows
            DdlProdutoOrigem.Items.Add(New ListItem(row("Codigo") & " - " & row("Descricao"), row("Codigo")))
        Next
    End Sub

    Private Sub CarregarPlanoDeContas()
        Sql = "   select Conta_ID as Codigo, Titulo as Descricao from planodeContas" & _
              "    Where Len(conta_id) >= 5"
        ' "   where Left(Conta_Id, 5) in ('10104', '10301', '20101', '30201', '30301')"
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlContaOrigem.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo") & "|" & Funcoes.SubstituirCaracteresEspeciais(Dr("Descricao"))))

            DdlDebitaMercadoria.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
            DdlCreditaMercadoria.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))

            DdlDebitaFrete.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
            DdlCreditaFrete.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next

        DdlContaOrigem.Items.Insert(0, "")
        DdlContaOrigem.SelectedIndex = 0

        DdlDebitaMercadoria.Items.Insert(0, "")
        DdlDebitaMercadoria.SelectedIndex = 0
        DdlCreditaMercadoria.Items.Insert(0, "")
        DdlCreditaMercadoria.SelectedIndex = 0

        DdlDebitaFrete.Items.Insert(0, "")
        DdlDebitaFrete.SelectedIndex = 0
        DdlCreditaFrete.Items.Insert(0, "")
        DdlCreditaFrete.SelectedIndex = 0
    End Sub

    Private Sub CarregarHistoricos()
        Sql = "SELECT Historico_Id as Codigo, Descricao " & vbCrLf & _
              "FROM Historicos " & vbCrLf & _
              "ORDER BY Historico_Id" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlHisMercadoria.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
            DdlHisFretes.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next

        DdlHisMercadoria.Items.Insert(0, "")
        DdlHisMercadoria.SelectedIndex = 0

        DdlHisFretes.Items.Insert(0, "")
        DdlHisFretes.SelectedIndex = 0

    End Sub

    Public Sub SqlOrigem(ByVal Sqls As ArrayList)
        For i As Integer = 0 To gridOrigem.Rows.Count - 1
            Dim sqlExcluir = "DELETE PlanoDeCustosXOrigem " & vbCrLf & _
                  " WHERE Codigo_Id = " & txtCodigo.Text.Trim() & vbCrLf & _
                  " AND Conta_Id = '" & gridOrigem.Rows(i).Cells(1).Text.Trim() & "'" & vbCrLf
            Sqls.Add(sqlExcluir)

            Sql = "INSERT INTO PlanoDeCustosXOrigem (" & vbCrLf & _
                  "  Codigo_Id " & vbCrLf & _
                  ", Conta_Id " & vbCrLf & _
                  ", Produto) " & vbCrLf & _
                  "  VALUES ('" & txtCodigo.Text.Trim() & "'" & vbCrLf & _
                  ", '" & gridOrigem.Rows(i).Cells(1).Text.Trim() & "'" & vbCrLf & _
                  ", " & IIf(String.IsNullOrWhiteSpace(Server.HtmlDecode(gridOrigem.Rows(i).Cells(3).Text)), "null", ("'" & Server.HtmlDecode(gridOrigem.Rows(i).Cells(3).Text).Trim() & "'")) & ")" & vbCrLf
            Sqls.Add(Sql)
        Next
    End Sub

    Function Valida() As Boolean
        If txtCodigo.Text = "" Then
            MsgBox(Me.Page, "Codigo é obrigatório.")
            txtCodigo.Focus()
            Return False
        End If
        If txtDescricao.Text = "" Then
            MsgBox(Me.Page, "Descrição é obrigatório.")
            txtDescricao.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True

        ddlTotalizador.SelectedIndex = 0
        DdlSinalPeso.SelectedIndex = 0
        DdlSinalValor.SelectedIndex = 0

        DdlContaOrigem.SelectedIndex = 0
        DdlProdutoOrigem.SelectedIndex = 0

        DdlDebitaMercadoria.SelectedIndex = 0
        DdlCreditaMercadoria.SelectedIndex = 0

        DdlDebitaFrete.SelectedIndex = 0
        DdlCreditaFrete.SelectedIndex = 0

        DdlHisMercadoria.SelectedIndex = 0
        DdlHisFretes.SelectedIndex = 0

        gridOrigem.DataSource = Nothing
        gridOrigem.DataBind()

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        TabContainer1.ActiveTabIndex = 0
        Session.Remove("Origem")
    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridPlanoDeCustos.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridPlanoDeCustos.SelectedRow.Cells(2).Text()

            Sql = "SELECT *" & vbCrLf & _
                  "FROM  PlanoDeCustos " & vbCrLf & _
                  "WHERE Codigo_Id = " & txtCodigo.Text & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows

                If Not IsDBNull(Dr("Totalizador")) And Dr("Totalizador") > 0 Then
                    ddlTotalizador.SelectedValue = Dr("Totalizador")
                End If

                If Not IsDBNull(Dr("SinalPeso")) And Dr("SinalPeso") <> "" Then
                    DdlSinalPeso.SelectedValue = Dr("SinalPeso")
                End If

                If Not IsDBNull(Dr("SinalValor")) And Dr("SinalValor") <> "" Then
                    DdlSinalValor.SelectedValue = Dr("SinalValor")
                End If

                If Not IsDBNull(Dr("DebitoMercadoria")) And Dr("DebitoMercadoria") <> "" Then
                    If DdlDebitaMercadoria.Items.FindByValue(Dr("DebitoMercadoria").ToString()) IsNot Nothing Then
                        DdlDebitaMercadoria.SelectedValue = Dr("DebitoMercadoria").ToString()
                    End If
                End If

                If Not IsDBNull(Dr("CreditoMercadoria")) And Dr("CreditoMercadoria") <> "" Then
                    If DdlCreditaMercadoria.Items.FindByValue(Dr("CreditoMercadoria").ToString()) IsNot Nothing Then
                        DdlCreditaMercadoria.SelectedValue = Dr("CreditoMercadoria").ToString()
                    End If
                End If

                If Not IsDBNull(Dr("DebitoFrete")) And Dr("DebitoFrete") <> "" Then
                    If DdlDebitaFrete.Items.FindByValue(Dr("DebitoFrete").ToString()) IsNot Nothing Then
                        DdlDebitaFrete.SelectedValue = Dr("DebitoFrete").ToString()
                    End If
                End If

                If Not IsDBNull(Dr("CreditoFrete")) And Dr("CreditoFrete") <> "" Then
                    If DdlCreditaFrete.Items.FindByValue(Dr("CreditoFrete").ToString()) IsNot Nothing Then
                        DdlCreditaFrete.SelectedValue = Dr("CreditoFrete").ToString()
                    End If
                End If

                If Not IsDBNull(Dr("HistoricoMercadoria")) And Dr("HistoricoMercadoria") > 0 Then
                    If DdlHisMercadoria.Items.FindByValue(Dr("HistoricoMercadoria").ToString()) IsNot Nothing Then
                        DdlHisMercadoria.SelectedValue = Dr("HistoricoMercadoria").ToString()
                    End If
                End If

                If Not IsDBNull(Dr("HistoricoFrete")) And Dr("HistoricoFrete") > 0 Then
                    If DdlHisFretes.Items.FindByValue(Dr("HistoricoFrete").ToString()) IsNot Nothing Then
                        DdlHisFretes.SelectedValue = Dr("HistoricoFrete").ToString()
                    End If
                End If

                If Not IsDBNull(Dr("SaldoInicial")) AndAlso CBool(Dr("SaldoInicial")) = True Then
                    ChkSaldoInicial.Checked = True
                Else
                    ChkSaldoInicial.Checked = False
                End If

                If Not IsDBNull(Dr("Desdobramento")) AndAlso CBool(Dr("Desdobramento")) = True Then
                    ChkDesdobramento.Checked = True
                Else
                    ChkDesdobramento.Checked = False
                End If

                chkRateio.Checked = Dr("Rateio")

            Next

            ListarOrigem()

            TabContainer1.ActiveTabIndex = 0

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True

            txtCodigo.Enabled = False
            txtDescricao.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Public Sub ListarOrigem()
        Sql = "SELECT PlanoDeCustosXOrigem.Conta_Id, PlanoDeContas.Titulo, " & vbCrLf & _
              "  ISNULL(PlanoDeCustosXOrigem.Produto,0) as Produto_Id, ISNULL(Produtos.Nome,'') as Produto " & vbCrLf & _
              "  FROM PlanoDeCustosXOrigem" & vbCrLf & _
              " INNER JOIN PlanoDeContas" & vbCrLf & _
              "    ON PlanoDeCustosXOrigem.Conta_Id  = PlanoDeContas.Conta_Id" & vbCrLf & _
              " LEFT JOIN Produtos" & vbCrLf & _
              "    ON PlanoDeCustosXOrigem.Produto  = Produtos.Produto_Id" & vbCrLf & _
              " Where PlanoDeCustosXOrigem.Codigo_Id =" & txtCodigo.Text
        dsOrigem = Banco.ConsultaDataSet(Sql, "Origem")

        For Each row As DataRow In dsOrigem.Tables(0).Rows
            If Not IsDBNull(row("Titulo")) Then
                row("Titulo") = Funcoes.SubstituirCaracteresEspeciais(row("Titulo"))
            End If
            If Not IsDBNull(row("Produto")) Then
                row("Produto") = Funcoes.SubstituirCaracteresEspeciais(row("Produto"))
            End If
        Next

        dsOrigem.AcceptChanges()
        gridOrigem.DataSource = dsOrigem
        gridOrigem.DataBind()
        SalvaOrigem()
    End Sub

    Protected Sub BtnAdiciona_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles BtnAdiciona.Click
        Try
            If (String.IsNullOrWhiteSpace(txtCodigo.Text)) Then
                MsgBox(Me.Page, "É necessário preencher o campo código.")
                Exit Sub
            End If

            If DdlContaOrigem.SelectedIndex = 0 Then
                MsgBox(Me.Page, "É necessário selecionar a conta origem.")
                Exit Sub
            End If

            RecuperaOrigem()

            If dsOrigem Is Nothing Then
                ListarOrigem()
            End If

            Dim objConta As String() = DdlContaOrigem.SelectedValue.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim CodProduto As String = ""
            Dim DescProd As String = ""
            If DdlProdutoOrigem.SelectedIndex > 0 Then
                CodProduto = DdlProdutoOrigem.SelectedValue
                DescProd = DdlProdutoOrigem.SelectedItem.Text.Split("-")(1).Trim()
            End If
            If txtCodigo.Enabled Then
                Dim row As DataRow
                row = dsOrigem.Tables(0).NewRow()
                row("Conta_Id") = objConta(0)
                row("Titulo") = objConta(1)
                row("Produto_Id") = CodProduto
                row("Produto") = DescProd
                dsOrigem.Tables(0).Rows.Add(row)
                gridOrigem.DataSource = dsOrigem
                gridOrigem.DataBind()
                SalvaOrigem()
            Else
                Dim sql As String
                sql = "INSERT INTO PlanoDeCustosXOrigem (Codigo_Id, Conta_Id, Produto)" & vbCrLf & _
                      " VALUES (" & txtCodigo.Text & ",'" & objConta(0) & "', " & IIf(String.IsNullOrWhiteSpace(CodProduto), "null", ("'" & CodProduto & "'")) & ")"
                Dim sqls As New ArrayList
                sqls.Add(sql)
                If Banco.GravaBanco(sqls) Then
                    Dim row As DataRow
                    row = dsOrigem.Tables(0).NewRow()
                    row("Conta_Id") = objConta(0)
                    row("Titulo") = objConta(1)
                    row("Produto_Id") = CodProduto
                    row("Produto") = DescProd
                    dsOrigem.Tables(0).Rows.Add(row)
                    gridOrigem.DataSource = dsOrigem
                    gridOrigem.DataBind()
                    SalvaOrigem()
                End If
            End If

            DdlContaOrigem.SelectedIndex = 0
            DdlProdutoOrigem.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRemove_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If DdlContaOrigem.SelectedIndex = 0 Then Exit Sub
            RecuperaOrigem()
            Dim objConta As String() = DdlContaOrigem.SelectedValue.Split("|")
            If txtCodigo.Enabled Then
                dsOrigem.Tables(0).Rows.Remove(dsOrigem.Tables(0).Rows(gridOrigem.SelectedIndex))
                gridOrigem.DataSource = dsOrigem
                gridOrigem.DataBind()
                SalvaOrigem()
            Else
                Dim sql As String
                sql = "DELETE PlanoDeCustosXOrigem " & vbCrLf & _
                      " WHERE Codigo_Id = " & txtCodigo.Text & vbCrLf & _
                      " AND Conta_Id  ='" & objConta(0) & "'" & vbCrLf

                Dim sqls As New ArrayList
                sqls.Add(sql)
                If Banco.GravaBanco(sqls) Then
                    If objConta(0) = dsOrigem.Tables(0).Rows(gridOrigem.SelectedIndex)("Conta_Id") Then
                        dsOrigem.Tables(0).Rows.Remove(dsOrigem.Tables(0).Rows(gridOrigem.SelectedIndex))
                    End If
                    gridOrigem.DataSource = dsOrigem
                    gridOrigem.DataBind()
                    SalvaOrigem()
                    DdlContaOrigem.SelectedIndex = 0
                    DdlProdutoOrigem.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub SalvaOrigem()
        Session("Origem") = dsOrigem
    End Sub

    Public Sub RecuperaOrigem()
        dsOrigem = Session("Origem")
    End Sub

    Protected Sub gridOrigem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridOrigem.SelectedIndexChanged
        Try
            If gridOrigem.Rows.Count = 0 Then Exit Sub

            If (Not String.IsNullOrWhiteSpace(gridOrigem.SelectedRow.Cells(1).Text) AndAlso Not String.IsNullOrWhiteSpace(gridOrigem.SelectedRow.Cells(2).Text)) Then
                DdlContaOrigem.SelectedValue = gridOrigem.SelectedRow.Cells(1).Text + "|" + gridOrigem.SelectedRow.Cells(2).Text
            End If

            If CInt(gridOrigem.SelectedRow.Cells(3).Text) > 0 Then
                DdlProdutoOrigem.SelectedValue = gridOrigem.SelectedRow.Cells(3).Text.Trim()
            Else
                DdlProdutoOrigem.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeCustos", "GRAVAR") Then
                If Valida() Then
                    Sql = "INSERT Into PlanoDeCustos (" & vbCrLf & _
                          "  Codigo_id" & vbCrLf & _
                          ", Descricao" & vbCrLf & _
                          ", Totalizador" & vbCrLf & _
                          ", SinalPeso" & vbCrLf & _
                          ", SinalValor" & vbCrLf & _
                          ", DebitoMercadoria" & vbCrLf & _
                          ", CreditoMercadoria" & vbCrLf & _
                          ", HistoricoMercadoria" & vbCrLf & _
                          ", DebitoFrete" & vbCrLf & _
                          ", CreditoFrete" & vbCrLf & _
                          ", HistoricoFrete" & vbCrLf & _
                          ", SaldoInicial" & vbCrLf & _
                          ", Desdobramento" & vbCrLf & _
                          ", Rateio" & vbCrLf & _
                          ", FaseDoTotalizador" & vbCrLf & _
                          ")" & vbCrLf

                    Sql &= " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                           ",'" & UCase(txtDescricao.Text) & "'" & vbCrLf & _
                    IIf(ddlTotalizador.Text <> "", ", " & ddlTotalizador.SelectedValue, ", ''") & vbCrLf & _
                    IIf(DdlSinalPeso.Text <> "", ", '" & DdlSinalPeso.SelectedValue & "'", ", ''") & vbCrLf & _
                    IIf(DdlSinalValor.Text <> "", ", '" & DdlSinalValor.SelectedValue & "'", ", ''") & vbCrLf & _
                    IIf(DdlDebitaMercadoria.Text <> "", ", " & DdlDebitaMercadoria.SelectedValue, ", ''") & vbCrLf & _
                    IIf(DdlCreditaMercadoria.Text <> "", ", " & DdlCreditaMercadoria.SelectedValue, ", ''") & vbCrLf & _
                    IIf(DdlHisMercadoria.Text <> "", ", " & DdlHisMercadoria.SelectedValue, ", ''") & vbCrLf & _
                    IIf(DdlDebitaFrete.Text <> "", ", " & DdlDebitaFrete.SelectedValue, ", ''") & vbCrLf & _
                    IIf(DdlCreditaFrete.Text <> "", ", " & DdlCreditaFrete.SelectedValue, ", ''") & vbCrLf & _
                    IIf(DdlHisFretes.Text <> "", ", = " & DdlHisFretes.SelectedValue, ", ''") & vbCrLf & _
                    IIf(ChkSaldoInicial.Checked, ",'True'", ",'False'") & vbCrLf & _
                    IIf(ChkDesdobramento.Checked, ",'True'", ",'False'") & vbCrLf & _
                    IIf(chkRateio.Checked, ",1", ",0") & vbCrLf & _
                           ",0)" & vbCrLf
                    SqlArray.Add(Sql)
                    SqlOrigem(SqlArray)

                    If Banco.GravaBanco(SqlArray) Then
                        Limpar()
                        CarregarPlanoDeCustos()
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
            If Funcoes.VerificaPermissao("PlanoDeCustos", "ALTERAR") Then
                Sql = "UPDATE PlanoDeCustos " & vbCrLf & _
                      " SET Descricao = '" & txtDescricao.Text & "'" & vbCrLf & _
                IIf(ddlTotalizador.Text <> "", ", Totalizador = " & ddlTotalizador.SelectedValue, ", Totalizador = ''") & vbCrLf & _
                IIf(DdlSinalPeso.Text <> "", ", SinalPeso = '" & DdlSinalPeso.SelectedValue & "'", ", SinalPeso = ''") & vbCrLf & _
                IIf(DdlSinalValor.Text <> "", ", SinalValor = '" & DdlSinalValor.SelectedValue & "'", ", SinalValor = ''") & vbCrLf & _
                IIf(DdlDebitaMercadoria.Text <> "", ", DebitoMercadoria = " & DdlDebitaMercadoria.SelectedValue, ", DebitoMercadoria = ''") & vbCrLf & _
                IIf(DdlCreditaMercadoria.Text <> "", ", CreditoMercadoria = " & DdlCreditaMercadoria.SelectedValue, ", CreditoMercadoria = ''") & vbCrLf & _
                IIf(DdlHisMercadoria.Text <> "", ", HistoricoMercadoria = " & DdlHisMercadoria.SelectedValue, ", HistoricoMercadoria = ''") & vbCrLf & _
                IIf(DdlDebitaFrete.Text <> "", ", DebitoFrete = " & DdlDebitaFrete.SelectedValue, ", DebitoFrete = ''") & vbCrLf & _
                IIf(DdlCreditaFrete.Text <> "", ", CreditoFrete = " & DdlCreditaFrete.SelectedValue, ", CreditoFrete = ''") & vbCrLf & _
                IIf(DdlHisFretes.Text <> "", ", HistoricoFrete = " & DdlHisFretes.SelectedValue, ", HistoricoFrete = ''") & vbCrLf & _
                IIf(ChkSaldoInicial.Checked = True, ", SaldoInicial = 'True'", ", SaldoInicial = 'False'") & vbCrLf & _
                IIf(ChkDesdobramento.Checked = True, ", Desdobramento = 'True'", ", Desdobramento = 'False'") & vbCrLf & _
                IIf(chkRateio.Checked, ", Rateio = 1", ", Rateio = 0") & vbCrLf
                Sql &= " WHERE Codigo_Id = " & txtCodigo.Text

                SqlArray.Add(Sql)
                SqlOrigem(SqlArray)

                If Banco.GravaBanco(SqlArray) Then
                    Limpar()
                    CarregarPlanoDeCustos()
                    TabContainer1.ActiveTabIndex = 1
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("PlanoDeCustos", "EXCLUIR") Then
                If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                    Dim objConta As String() = DdlContaOrigem.SelectedValue.Split("|")
                    Sql = "DELETE PlanoDeCustosXOrigem " & vbCrLf & _
                          " WHERE Codigo_Id = " & txtCodigo.Text & vbCrLf & _
                          " AND Conta_Id = '" & objConta(0) & "'" & vbCrLf
                    SqlArray.Add(Sql)

                    Sql = "DELETE FROM PlanoDeCustos" & vbCrLf & _
                          " WHERE Codigo_Id = " & txtCodigo.Text & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) = False Then
                        Limpar()
                        CarregarPlanoDeCustos()
                        TabContainer1.ActiveTabIndex = 1
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, "Selecione um PlanoDeCusto na lista, para ser excluído.")
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
            If Funcoes.VerificaPermissao("PlanoDeCustos", "RELATORIO") Then
                Sql = "  SELECT * " & vbCrLf & _
                      " FROM PlanoDeCustos " & vbCrLf & _
                      " ORDER BY Codigo_Id" & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "PlanoDeCustos")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Nome", HttpContext.Current.Session("ssNomeEmpresa"))
                parameters.Add("Cidade", HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa"))

                Funcoes.BindReport(Me.Page, ds, "Cr_PlanoDeCustos", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PlanoDeCustos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class