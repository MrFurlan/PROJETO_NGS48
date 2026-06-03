Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CulturaPragaFito
    Inherits BasePage

    Private Sql As String = ""
    Private mensagemErro As String = ""
    Private arraySql As ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CulturaPragaFito", "ACESSAR") Then
                carregarCultura()
                carregarPraga()
                carregarFito()
                carregarSolo()
                Limpar()
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub carregarCultura()
        Dim Lista As New [Lib].Negocio.ListCultura(True, "Descricao")
        For Each dr As Cultura In Lista
            ddlCultura.Items.Add(New ListItem(Funcoes.AlinharEsquerda(dr.Descricao, 45, ".") & Funcoes.AlinharEsquerda(dr.NomeCientifico, 45, ".") & " - " & Format(dr.Codigo, "000"), dr.Codigo))
        Next
        Funcoes.InserirLinhaEmBranco(ddlCultura)
    End Sub

    Private Sub carregarPraga()
        Dim Lista As New [Lib].Negocio.ListPraga(True, "NomeComum")
        For Each objPraga As [Lib].Negocio.Praga In Lista
            ddlPraga.Items.Add(New ListItem(Funcoes.AlinharEsquerda(objPraga.NomeComum, 50, ".") & "-" & Format(objPraga.Codigo, "0000"), objPraga.Codigo))
        Next
        Funcoes.InserirLinhaEmBranco(ddlPraga)
    End Sub

    Private Sub carregarFito()
        Dim Lista As New [Lib].Negocio.ListaFito("", 0, "", "NomeComercial")
        For Each objFito As [Lib].Negocio.Fito In Lista
            ddlFito.Items.Add(New ListItem(Funcoes.AlinharEsquerda(objFito.NomeComercial, 50, ".") & "-" & Format(objFito.CodigoFito, "0000"), objFito.CodigoFito))
        Next
        Funcoes.InserirLinhaEmBranco(ddlFito)
    End Sub

    Private Sub carregarFormaDeAplicacao()
        Dim Lista As New [Lib].Negocio.ListFormaDeAplicacao(True, "Descricao")

        gridFormaDeAplicacao.DataSource = Lista.ToArray()
        gridFormaDeAplicacao.DataBind()

        For i = 0 To gridFormaDeAplicacao.Rows.Count - 1
            CType(gridFormaDeAplicacao.Rows(i).FindControl("txtModoDeAplicacao"), TextBox).Attributes.Add("onKeyPress", "javascript:if(ValidarTecla(event) == 13){ValidarObservacao(event)}")
        Next
    End Sub

    Private Sub carregarSolo()
        Dim Lista As New [Lib].Negocio.ListSolo(True, "Descricao")

        For Each dr As [Lib].Negocio.Solo In Lista
            ddlSolo.Items.Add(New ListItem(Funcoes.AlinharEsquerda(dr.Descricao, 30, ".") & "-" & Format(dr.Codigo, "00"), dr.Codigo))
        Next

        Funcoes.InserirLinhaEmBranco(ddlSolo)
    End Sub

    Private Sub carregarGrid(ByVal Where As String)
        Dim ds As New DataSet
        Sql = "SELECT     CulturaXPragaXFito.CulturaPragaFito_Id, CONVERT(varchar, CulturaXPragaXFito.Cultura) + '-' + Cultura.Descricao AS Cultura, " & vbCrLf & _
              "CONVERT(varchar, CulturaXPragaXFito.Fito) + '-' + Fito.NomeComercial AS Fito, CONVERT(varchar, CulturaXPragaXFito.Praga) + '-' + Praga.NomeComum AS Praga, " & vbCrLf & _
              "                      CulturaXPragaXFito.status " & vbCrLf & _
              "FROM         CulturaXPragaXFito INNER JOIN " & vbCrLf & _
              "                      Cultura ON CulturaXPragaXFito.Cultura = Cultura.Cultura_Id INNER JOIN " & vbCrLf & _
              "                      Fito ON CulturaXPragaXFito.Fito = Fito.Fito_Id INNER JOIN " & vbCrLf & _
              "                      Praga ON CulturaXPragaXFito.Praga = Praga.Praga_Id " & vbCrLf & _
              Where

        ds = Banco.ConsultaDataSet(Sql, "Codigo")

        gridCulturaPragaFito.DataSource = ds
        gridCulturaPragaFito.DataBind()

        If ds.Tables(0).Rows.Count > 0 Then
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCulturaPragaFito.Enabled = False
            TabContainer1.ActiveTabIndex = 2
        End If
    End Sub

    Private Sub Limpar()
        txtCulturaPragaFito.Enabled = True
        radAtivo.Enabled = True
        radBaixado.Enabled = True
        ddlCultura.Enabled = True
        ddlPraga.Enabled = True
        ddlFito.Enabled = True
        ddlSolo.Enabled = True
        txtMinimo.Enabled = True
        txtRecomendada.Enabled = True
        txtMaximo.Enabled = True
        btnAdicionarDosagem.Enabled = True
        txtVazaoTerrestre.Enabled = True
        txtVazaoAerea.Enabled = True
        txtIntervaloDeSeguranca.Enabled = True
        txtEpocaDeAplicacao.Enabled = True

        txtCulturaPragaFito.Text = ""
        radAtivo.Checked = True
        radBaixado.Checked = False
        ddlCultura.SelectedIndex = 0
        ddlPraga.SelectedIndex = 0
        ddlFito.SelectedIndex = 0
        ddlSolo.SelectedIndex = 0

        txtMinimo.Text = ""
        txtRecomendada.Text = ""
        txtMaximo.Text = ""
        txtVazaoTerrestre.Text = ""
        txtVazaoAerea.Text = ""
        txtUnidadeDeMedida.Text = ""
        txtIntervaloDeSeguranca.Text = ""
        txtEpocaDeAplicacao.Text = ""

        Session.Remove("ssDosagem")
        Dim dtDosagem As New DataTable("Dosagem")
        dtDosagem.Columns.Add("Indice", GetType(String))
        dtDosagem.Columns.Add("Solo", Type.GetType("System.String"))
        dtDosagem.Columns.Add("Minima", Type.GetType("System.Double"))
        dtDosagem.Columns.Add("Recomendada", Type.GetType("System.Double"))
        dtDosagem.Columns.Add("Maxima", Type.GetType("System.Double"))
        dtDosagem.Columns.Add("UnidadeDeMedida", Type.GetType("System.String"))
        dtDosagem.Columns.Add("VazaoTerrestre", Type.GetType("System.String"))
        dtDosagem.Columns.Add("VazaoAerea", Type.GetType("System.String"))
        dtDosagem.Columns.Add("IntervaloDeSeguranca", Type.GetType("System.String"))

        Session("ssDosagem") = dtDosagem

        gridDosagem.DataBind()
        gridCulturaPragaFito.DataBind()

        carregarFormaDeAplicacao()

        TabContainer1.ActiveTabIndex = 0

        txtCulturaPragaFito.Focus()

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkLimpar.Parent.Visible = True
    End Sub

    Private Function ValidarCampos() As Boolean
        Dim chk As Boolean = False
        Dim texto As Boolean = False

        For i = 0 To gridFormaDeAplicacao.Rows.Count - 1
            If CType(gridFormaDeAplicacao.Rows(i).FindControl("chkCodigo"), CheckBox).Checked Then
                chk = True
                If CType(gridFormaDeAplicacao.Rows(i).FindControl("txtModoDeAplicacao"), TextBox).Text.ToString.Length > 0 Then
                    texto = True
                Else
                    texto = False
                    Exit For
                End If
            End If
        Next

        If ddlCultura.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Cultura não foi informada.")
            Return False
        ElseIf ddlPraga.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Praga não foi informada.")
            Return False
        ElseIf ddlFito.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Fito não foi informada.")
            Return False
        ElseIf txtEpocaDeAplicacao.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Época de Aplicação não foi informada.")
            Return False
        ElseIf gridDosagem.Rows.Count = 0 Then
            MsgBox(Me.Page, "Dosagem não foi informada.")
            Return False
        ElseIf txtEpocaDeAplicacao.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Época de Aplicação não foi informada.")
            Return False
        ElseIf chk = False Then
            MsgBox(Me.Page, "Forma de Aplicação não foi informada.")
            Return False
        ElseIf chk = True And texto = False Then
            MsgBox(Me.Page, "Modo de Aplicação não foi informado.")
            Return False
        ElseIf txtVazaoTerrestre.Text.Length = 0 Then
            MsgBox(Me.Page, "Em vazão terrestre informe N/D que significa não disponível.")
            Return False
        ElseIf txtVazaoAerea.Text.Length = 0 Then
            MsgBox(Me.Page, "Em vazão aérea informe N/D que significa não disponível.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Function ValidarCulturaPragaFito() As Boolean
        Dim ds As New DataSet
        Dim strAndWhere As String = " WHERE "

        Sql = "Select CulturaPragaFito_Id as Codigo from culturaXpragaXfito "

        If ddlCultura.SelectedIndex > 0 Then
            Sql &= strAndWhere & " Cultura = " & ddlCultura.SelectedValue
            strAndWhere = " AND "
        End If

        If ddlPraga.SelectedIndex > 0 Then
            Sql &= strAndWhere & " Praga = " & ddlPraga.SelectedValue
            strAndWhere = " AND "
        End If

        If ddlFito.SelectedIndex > 0 Then
            Sql &= strAndWhere & " Fito = " & ddlFito.SelectedValue
        End If

        ds = Banco.ConsultaDataSet(Sql, "CulturaPragaFito")

        If ds.Tables(0).Rows Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Return True
        Else
            mensagemErro = "A seleção já existe, veja na Lista do Grid. Faça a consulta e verifique."
            Return False
        End If
    End Function

    Private Sub IniciarCPF(ByVal Tipo As String)
        If ValidarCampos() Then
            Dim arraySql As New ArrayList

            Select Case Tipo
                Case "I"
                    Sql = "Select CulturaPragaFito_Id as Codigo from culturaXpragaXfito where Cultura = " & ddlCultura.SelectedValue.ToString & " AND Fito = " & ddlFito.SelectedValue.ToString & " AND Praga = " & ddlPraga.SelectedValue.ToString & ""

                    Dim ds As New DataSet

                    ds = Banco.ConsultaDataSet(Sql, "Codigo")

                    If ValidarCulturaPragaFito() Then
                        Sql = "INSERT INTO  CulturaXPragaXFito " & vbCrLf & _
                              "			(CulturaPragaFito_Id, Cultura, Fito, Praga, EpocaAplicacao, status) " & vbCrLf & _
                              "VALUES " & vbCrLf & _
                              "			(" & txtCulturaPragaFito.Text.ToString & "," & ddlCultura.SelectedValue.ToString & "," & ddlFito.SelectedValue.ToString & "," & ddlPraga.SelectedValue.ToString & ",'" & txtEpocaDeAplicacao.Text.ToString.Replace("'", "") & "','" & IIf(radAtivo.Checked, "A", "I") & "')"

                        arraySql.Add(Sql)

                        For i = 0 To gridDosagem.Rows.Count - 1
                            Dim solo() As String = gridDosagem.Rows(i).Cells(1).Text.ToString.Split("-")

                            Sql = "INSERT INTO  CulturaxPragaxFitoxDosagem " & vbCrLf & _
                                  "			(CulturaPragaFito_Id, Dosagem_Id, Solo, DosagemMinima, DosagemMaxima, UnidadeDeMedida, VazaoTerrestre, VazaoAerea, IntervaloDeSeguranca, DosagemRecomendada) " & vbCrLf & _
                                  "			(Select " & txtCulturaPragaFito.Text.ToString & ", (select max(dosagem_id) + 1 as Dosagem from culturaXpragaXfitoXdosagem) as dosagem, " & solo(1).ToString & "," & gridDosagem.Rows(i).Cells(2).Text.Replace(".", "").Replace(",", ".") & "," & gridDosagem.Rows(i).Cells(4).Text.Replace(".", "").Replace(",", ".") & ",'" & gridDosagem.Rows(i).Cells(5).Text & "','" & gridDosagem.Rows(i).Cells(6).Text & "', '" & gridDosagem.Rows(i).Cells(7).Text & "', '" & gridDosagem.Rows(i).Cells(8).Text & "', " & gridDosagem.Rows(i).Cells(3).Text.Replace(".", "").Replace(",", ".") & ")"

                            arraySql.Add(Sql)
                        Next

                        For i = 0 To gridFormaDeAplicacao.Rows.Count - 1
                            If CType(gridFormaDeAplicacao.Rows(i).FindControl("chkCodigo"), CheckBox).Checked = True Then
                                Sql = "INSERT INTO  ModoDeAplicacao " & vbCrLf & _
                                      "			(CulturaPragaFito_Id, FormaDeAplicacao_Id, Descricao) " & vbCrLf & _
                                      "VALUES " & vbCrLf & _
                                      "			(" & txtCulturaPragaFito.Text.ToString & "," & gridFormaDeAplicacao.Rows(i).Cells(1).Text & ",'" & Trim(CType(gridFormaDeAplicacao.Rows(i).FindControl("txtModoDeAplicacao"), TextBox).Text.ToString.Replace("'", "")) & "')"

                                arraySql.Add(Sql)
                            End If
                        Next

                        If Banco.GravaBanco(arraySql) Then
                            MsgBox(Me.Page, "Registro incluido com Sucesso.", eTitulo.Sucess)
                            Limpar()
                        Else
                            MsgBox(Me.Page, "Erro ao incluir: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        End If
                    Else
                        MsgBox(Me.Page, mensagemErro)
                        ConsultarFito()
                    End If
                Case "U"
                    Sql = "UPDATE  CulturaXPragaXFito " & vbCrLf & _
                          "	  SET  EpocaAplicacao = '" & txtEpocaDeAplicacao.Text.ToString.Replace("'", "") & "', status = '" & IIf(radAtivo.Checked, "A", "I") & "' " & vbCrLf & _
                          " WHERE  CulturaPragaFito_Id = " & txtCulturaPragaFito.Text.ToString & " AND Cultura = " & ddlCultura.SelectedValue.ToString & " AND Fito = " & ddlFito.SelectedValue.ToString & " AND Praga = " & ddlPraga.SelectedValue.ToString

                    arraySql.Add(Sql)

                    Sql = "DELETE FROM  CulturaxPragaxFitoxDosagem " & vbCrLf & _
                          "WHERE        CulturaPragaFito_Id = " & txtCulturaPragaFito.Text.ToString

                    arraySql.Add(Sql)

                    For i = 0 To gridDosagem.Rows.Count - 1
                        Dim solo() As String = gridDosagem.Rows(i).Cells(1).Text.ToString.Split("-")

                        Sql = "INSERT INTO  CulturaxPragaxFitoxDosagem " & vbCrLf & _
                              "			(CulturaPragaFito_Id, Dosagem_Id, Solo, DosagemMinima, DosagemMaxima, UnidadeDeMedida, VazaoTerrestre, VazaoAerea, IntervaloDeSeguranca, DosagemRecomendada) " & vbCrLf & _
                              "			(Select " & txtCulturaPragaFito.Text.ToString & ", (select max(dosagem_id) + 1 as Dosagem from culturaXpragaXfitoXdosagem) as dosagem, " & solo(1).ToString & "," & gridDosagem.Rows(i).Cells(2).Text.Replace(".", "").Replace(",", ".") & "," & gridDosagem.Rows(i).Cells(4).Text.Replace(".", "").Replace(",", ".") & ",'" & gridDosagem.Rows(i).Cells(5).Text & "','" & gridDosagem.Rows(i).Cells(6).Text & "', '" & gridDosagem.Rows(i).Cells(7).Text & "', '" & gridDosagem.Rows(i).Cells(8).Text & "', " & gridDosagem.Rows(i).Cells(3).Text.Replace(".", "").Replace(",", ".") & ")"

                        arraySql.Add(Sql)
                    Next

                    Sql = "DELETE FROM  ModoDeAplicacao " & vbCrLf & _
                          "WHERE        CulturaPragaFito_Id = " & txtCulturaPragaFito.Text.ToString

                    arraySql.Add(Sql)

                    For i = 0 To gridFormaDeAplicacao.Rows.Count - 1
                        If CType(gridFormaDeAplicacao.Rows(i).FindControl("chkCodigo"), CheckBox).Checked = True Then
                            Sql = "INSERT INTO  ModoDeAplicacao " & vbCrLf & _
                                  "			(CulturaPragaFito_Id, FormaDeAplicacao_Id, Descricao) " & vbCrLf & _
                                  "VALUES " & vbCrLf & _
                                  "			(" & txtCulturaPragaFito.Text.ToString & "," & gridFormaDeAplicacao.Rows(i).Cells(1).Text & ",'" & Trim(CType(gridFormaDeAplicacao.Rows(i).FindControl("txtModoDeAplicacao"), TextBox).Text.ToString.Replace("'", "")) & "')"

                            arraySql.Add(Sql)
                        End If
                    Next

                    If Banco.GravaBanco(arraySql) Then
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                    Else
                        MsgBox(Me.Page, "Erro ao alterar: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End If
                Case "D"
                    MsgBox(Me.Page, "Em Desenvolvimento.")
            End Select
        Else
            Select Case Tipo
                Case "I"
                    MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
                Case "U"
                    MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
                Case "D"
                    MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End Select
        End If
    End Sub

    Private Sub ConsultarFito()
        Dim strAndWhere As String = " WHERE "
        Sql = ""

        If txtCulturaPragaFito.Text.ToString.Length > 0 Then
            Sql &= strAndWhere & " CulturaPragaFito_Id = " & txtCulturaPragaFito.Text
            strAndWhere = " AND "
        End If
        If ddlCultura.SelectedIndex > 0 Then
            Sql &= strAndWhere & " Cultura = " & ddlCultura.SelectedValue
            strAndWhere = " AND "
        End If
        If ddlPraga.SelectedIndex > 0 Then
            Sql &= strAndWhere & " Praga = " & ddlPraga.SelectedValue
            strAndWhere = " AND "
        End If
        If ddlFito.SelectedIndex > 0 Then
            Sql &= strAndWhere & " Fito = " & ddlFito.SelectedValue
        End If

        carregarGrid(Sql)
    End Sub

    Protected Sub ddlFito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ValidarCulturaPragaFito() Then
            Dim ds As New DataSet

            Sql = "select max(CulturaPragaFito_Id) + 1 as Codigo from culturaXpragaXfito"

            ds = Banco.ConsultaDataSet(Sql, "Codigo")

            If ds.Tables(0).Rows Is Nothing Or ds.Tables(0).Rows.Count = 0 Then
                Limpar()
                MsgBox(Me.Page, "Numerador CulturaXPragaXFito não pode ser iniciado. Entre em contato com o suporte")
            Else
                txtCulturaPragaFito.Text = ds.Tables(0).Rows(0).Item("Codigo").ToString
                txtCulturaPragaFito.Enabled = False
            End If
        Else
            MsgBox(Me.Page, mensagemErro)
            ConsultarFito()
        End If
    End Sub

    Protected Sub gridDosagem_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs)
        If e.CommandName = "Select" Then
            Dim strArrayItem As String() = e.CommandArgument.ToString().Split(";")

            For i = 0 To CType(Session("ssDosagem"), DataTable).Rows.Count - 1
                If Convert.ToInt32(CType(Session("ssDosagem"), DataTable).Rows(i).Item("Solo").ToString().Split("-").ToArray(1)) = strArrayItem(0) And CType(Session("ssDosagem"), DataTable).Rows(i).Item("Minima") = CDec(strArrayItem(1)) And CType(Session("ssDosagem"), DataTable).Rows(i).Item("Recomendada") = CDec(strArrayItem(2)) And CType(Session("ssDosagem"), DataTable).Rows(i).Item("Maxima") = CDec(strArrayItem(3)) Then
                    CType(Session("ssDosagem"), DataTable).Rows(i).Delete()
                    Exit For
                End If
            Next

            gridDosagem.DataSource = CType(Session("ssDosagem"), DataTable)
            gridDosagem.DataBind()
        End If
    End Sub

    Protected Sub gridCulturaPragaFito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ds As New DataSet

        Sql = "SELECT     CulturaPragaFito_Id, Cultura, Fito, Praga, EpocaAplicacao, status " & vbCrLf & _
              "FROM       CulturaXPragaXFito " & vbCrLf & _
              "Where      CulturaPragaFito_Id = " & gridCulturaPragaFito.SelectedRow.Cells(1).Text()

        ds = Banco.ConsultaDataSet(Sql, "CulturaXPragaXFito")

        If ds.Tables(0).Rows.Count > 0 Then
            txtCulturaPragaFito.Text = ds.Tables(0).Rows(0).Item("CulturaPragaFito_Id").ToString
            ddlCultura.SelectedValue = ds.Tables(0).Rows(0).Item("Cultura").ToString
            ddlPraga.SelectedValue = ds.Tables(0).Rows(0).Item("Praga").ToString
            ddlFito.SelectedValue = ds.Tables(0).Rows(0).Item("Fito").ToString
            txtEpocaDeAplicacao.Text = ds.Tables(0).Rows(0).Item("EpocaAplicacao").ToString
            If ds.Tables(0).Rows(0).Item("status").ToString = "A" Then
                radAtivo.Checked = True
            Else
                radBaixado.Checked = True
            End If

            Sql = "SELECT     Dosagem_Id, Solo, DosagemMinima, DosagemMaxima, UnidadeDeMedida, " & vbCrLf & _
                  "           VazaoTerrestre, VazaoAerea, IntervaloDeSeguranca, DosagemRecomendada " & vbCrLf & _
                  "FROM       CulturaxPragaxFitoxDosagem " & vbCrLf & _
                  "Where      CulturaPragaFito_Id = " & gridCulturaPragaFito.SelectedRow.Cells(1).Text()

            ds = Banco.ConsultaDataSet(Sql, "CulturaXPragaXFitoXDosagem")

            For Each dr As DataRow In ds.Tables(0).Rows
                Dim drRow As DataRow = CType(Session("ssDosagem"), DataTable).NewRow()
                drRow("Indice") = dr("Solo") & ";" & dr("DosagemMinima") & ";" & dr("DosagemRecomendada") & ";" & dr("DosagemMaxima")
                ddlSolo.SelectedValue = dr("Solo")
                drRow("Solo") = ddlSolo.SelectedItem.Text
                drRow("Minima") = dr("DosagemMinima")
                drRow("Recomendada") = dr("DosagemRecomendada")
                drRow("Maxima") = dr("DosagemMaxima")
                drRow("UnidadeDeMedida") = dr("UnidadeDeMedida")
                drRow("VazaoTerrestre") = dr("VazaoTerrestre")
                drRow("VazaoAerea") = dr("VazaoAerea")
                drRow("IntervaloDeSeguranca") = dr("IntervaloDeSeguranca")
                CType(Session("ssDosagem"), DataTable).Rows.Add(drRow)
            Next

            gridDosagem.DataSource = CType(Session("ssDosagem"), DataTable)
            gridDosagem.DataBind()

            Sql = "SELECT     FormaDeAplicacao_Id, Descricao " & vbCrLf & _
                  "FROM       ModoDeAplicacao " & vbCrLf & _
                  "Where      CulturaPragaFito_Id = " & gridCulturaPragaFito.SelectedRow.Cells(1).Text()

            ds = Banco.ConsultaDataSet(Sql, "ModoDeAplicacao")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    For i = 0 To gridFormaDeAplicacao.Rows.Count - 1
                        If dr("FormaDeAplicacao_Id") = gridFormaDeAplicacao.Rows(i).Cells(1).Text Then
                            CType(gridFormaDeAplicacao.Rows(i).FindControl("chkCodigo"), CheckBox).Checked = True
                            CType(gridFormaDeAplicacao.Rows(i).FindControl("txtModoDeAplicacao"), TextBox).Text = dr("Descricao")
                        End If
                    Next
                Next
            End If

            ddlCultura.Enabled = False
            ddlPraga.Enabled = False
            ddlFito.Enabled = False

            TabContainer1.ActiveTabIndex = 0
        Else
            MsgBox(Me.Page, "CulturaXPragaXFito não encontrado ou cadastro com problema. Entre em contato com o Suporte.")
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CulturaPragaFito", "GRAVAR") Then
                IniciarCPF("I")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("", "") Then
                IniciarCPF("U")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("", "") Then
                IniciarCPF("D")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para ")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If txtCulturaPragaFito.Text.ToString.Length > 0 OrElse ddlCultura.SelectedIndex > 0 OrElse ddlPraga.SelectedIndex > 0 OrElse ddlFito.SelectedIndex > 0 Then
            ConsultarFito()
        Else
            MsgBox(Me.Page, "Informe o Código ou selecione Cultura, Praga ou Fito para consulta")
        End If
    End Sub

    Protected Sub btnAdicionarDosagem_Click(sender As Object, e As EventArgs) Handles btnAdicionarDosagem.Click
        If ddlSolo.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Solo não foi informado.")
        ElseIf txtMinimo.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Dosagem mínima não foi informada.")
        ElseIf txtRecomendada.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Dosagem recomendada não foi informada.")
        ElseIf txtMaximo.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Dosagem máxima não foi informada.")
        ElseIf txtUnidadeDeMedida.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Unidade de Medida não foi informada.")
        ElseIf txtIntervaloDeSeguranca.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Intervalo de Segurança não foi informado.")
        Else
            Dim tem As Boolean = False
            For i = 0 To CType(Session("ssDosagem"), DataTable).Rows.Count - 1
                If CType(Session("ssDosagem"), DataTable).Rows(i).Item("Solo") = ddlSolo.SelectedItem.Text And CType(Session("ssDosagem"), DataTable).Rows(i).Item("Minima") = txtMinimo.Text And CType(Session("ssDosagem"), DataTable).Rows(i).Item("Recomendada") = txtRecomendada.Text And CType(Session("ssDosagem"), DataTable).Rows(i).Item("Maxima") = txtMaximo.Text Then
                    tem = True
                End If
            Next

            If tem Then
                MsgBox(Me.Page, "Dosagem já existente.")
            Else
                Dim drRow As DataRow = CType(Session("ssDosagem"), DataTable).NewRow()
                drRow("Indice") = ddlSolo.SelectedItem.Text & ";" & txtMinimo.Text & ";" & txtRecomendada.Text & ";" & txtMaximo.Text
                drRow("Solo") = ddlSolo.SelectedItem.Text
                drRow("Minima") = txtMinimo.Text
                drRow("Recomendada") = txtRecomendada.Text
                drRow("Maxima") = txtMaximo.Text
                drRow("UnidadeDeMedida") = txtUnidadeDeMedida.Text
                drRow("VazaoTerrestre") = txtVazaoTerrestre.Text
                drRow("VazaoAerea") = txtVazaoAerea.Text
                drRow("IntervaloDeSeguranca") = txtIntervaloDeSeguranca.Text
                CType(Session("ssDosagem"), DataTable).Rows.Add(drRow)

                gridDosagem.DataSource = CType(Session("ssDosagem"), DataTable)
                gridDosagem.DataBind()

                ddlSolo.SelectedIndex = 0
                txtMinimo.Text = ""
                txtRecomendada.Text = ""
                txtMaximo.Text = ""
            End If
        End If
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CulturaPragaFito")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class