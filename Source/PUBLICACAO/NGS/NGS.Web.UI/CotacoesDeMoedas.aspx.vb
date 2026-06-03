Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CotacoesDeMoedas
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CotacoesDeMoedas", "ACESSAR") Then
                ddl.Carregar(ddlIndexador, CarregarDDL.Tabela.Indexador, "")
                ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, "2015;15;C", False)
                ddlAno.SelectedValue = Now.Year
                'ddlAno.SelectedValue = Today.Year
                ddlMes.SelectedValue = Now.Month
                CarregarDias(Now.Month, Now.Year)
                CarregarGridCotacoes(ddlMes.SelectedValue, ddlAno.SelectedValue)
                ddlDia.SelectedValue = Now.Day
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarDias(ByVal Mes As Integer, ByVal Ano As Integer)
        Dim Xdias As Integer = Date.DaysInMonth(Ano, Mes)
        Dim i As Integer

        For i = 1 To Xdias
            Dim x As New ListItem
            x.Value = i
            x.Text = i
            ddlDia.Items.Add(x)
        Next
        For i = 31 To 28 Step -1
            If Xdias < i Then
                gridCotacoesDeMoedas.Columns(i + 1).Visible = False
            Else
                gridCotacoesDeMoedas.Columns(i + 1).Visible = True
            End If
        Next
        If Mes = Date.Now.Month And Ano = Date.Now.Year Then ddlDia.SelectedValue = Date.Now.Day
    End Sub

    Private Sub CarregarGridCotacoes(ByVal Mes As Integer, ByVal Ano As Integer)
        ddlDia.Items.Clear()

        CarregarDias(Mes, Ano)

        Sql = "select I.Indexador_Id," & vbCrLf & _
              "       I.Descricao," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 01 then C.indice else 0 end) dia01," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 02 then C.indice else 0 end) dia02," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 03 then C.indice else 0 end) dia03," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 04 then C.indice else 0 end) dia04," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 05 then C.indice else 0 end) dia05," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 06 then C.indice else 0 end) dia06," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 07 then C.indice else 0 end) dia07," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 08 then C.indice else 0 end) dia08," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 09 then C.indice else 0 end) dia09," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 10 then C.indice else 0 end) dia10," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 11 then C.indice else 0 end) dia11," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 12 then C.indice else 0 end) dia12," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 13 then C.indice else 0 end) dia13," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 14 then C.indice else 0 end) dia14," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 15 then C.indice else 0 end) dia15," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 16 then C.indice else 0 end) dia16," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 17 then C.indice else 0 end) dia17," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 18 then C.indice else 0 end) dia18," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 19 then C.indice else 0 end) dia19," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 20 then C.indice else 0 end) dia20," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 21 then C.indice else 0 end) dia21," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 22 then C.indice else 0 end) dia22," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 23 then C.indice else 0 end) dia23," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 24 then C.indice else 0 end) dia24," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 25 then C.indice else 0 end) dia25," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 26 then C.indice else 0 end) dia26," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 27 then C.indice else 0 end) dia27," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 28 then C.indice else 0 end) dia28," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 29 then C.indice else 0 end) dia29," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 30 then C.indice else 0 end) dia30," & vbCrLf & _
              "       sum(case when day(C.data_Id) = 31 then C.indice else 0 end) dia31" & vbCrLf & _
              "  from indexadores I" & vbCrLf & _
              "  left join cotacoes C" & vbCrLf & _
              "    on I.Indexador_id = C.Indexador_id" & vbCrLf & _
              " where Year(C.Data_id)  = " & Ano & vbCrLf & _
              "   and Month(C.Data_id) = " & Mes & vbCrLf & _
              " group by I.Indexador_Id," & vbCrLf & _
              "          I.Descricao" & vbCrLf
        gridCotacoesDeMoedas.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        gridCotacoesDeMoedas.DataBind()

        CarregarCotacao()
    End Sub

    Private Sub CarregarCotacao()
        If Not ValidaCotacao() Then
            txtIndice.Text = ""
            LblLiberado.Text = ""
            LblAlterado.Text = ""
            Exit Sub
        End If

        Dim Cot As New [Lib].Negocio.Cotacao(ddlIndexador.SelectedValue, CDate(ddlAno.SelectedValue & "-" & ddlMes.SelectedValue & "-" & ddlDia.SelectedValue))
        txtIndice.Text = Cot.Indice

        If Cot.Realizado Then
            rdRealizado.Checked = True
            rdPrevisto.Checked = False
            rdRealizado.Enabled = False
            rdPrevisto.Enabled = False
        Else
            rdRealizado.Checked = False
            rdPrevisto.Checked = True
            rdRealizado.Enabled = True
            rdPrevisto.Enabled = True
        End If
        If Not Cot.UsuarioLiberacao Is Nothing Then
            LblLiberado.Text = Cot.UsuarioLiberacao & " : " & Cot.UsuarioLiberacaoData.ToString("dd-MM-yyyy HH:mm")
        Else
            LblLiberado.Text = ""
        End If
        If Not Cot.UsuarioAlteracao Is Nothing Then
            LblAlterado.Text = Cot.UsuarioAlteracao & " : " & Cot.UsuarioAlteracaoData.ToString("dd-MM-yyyy HH:mm")
        Else
            LblAlterado.Text = ""
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:"
        If Not String.IsNullOrWhiteSpace(ddlDia.SelectedValue) Then
            param &= "Data dia: " & ddlDia.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlMes.SelectedValue) Then
            param &= "Mês: " & ddlMes.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlAno.SelectedValue) Then
            param &= "Ano: " & ddlAno.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlIndexador.SelectedValue) Then
            param &= "Indexador: " & ddlIndexador.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtIndice.Text) Then
            param &= "Valor: " & txtIndice.Text & " - "
        End If
        param &= IIf(rdPrevisto.Checked, "Previsto: Sim", "Realizado: Sim")

        Return param
    End Function

    Function ValidaCotacao(Optional ByVal ExibeMensagem As Boolean = False) As Boolean
        Dim msg As String = ""

        If ddlAno.SelectedIndex = 0 Then
            msg &= "Informe o ano da cotação." & vbCrLf
        End If
        If ddlIndexador.SelectedIndex = 0 Then
            msg &= "Informe o indexador da cotação." & vbCrLf
        End If
        If msg.Length > 0 Then
            If ExibeMensagem Then MsgBox(Me.Page, msg)
            Return False
        End If

        Return True
    End Function

    Protected Sub btnMesAnterior_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlMes.SelectedValue = 1 Then
                ddlAno.SelectedValue -= 1
                ddlMes.SelectedValue = 12
            Else
                ddlMes.SelectedValue -= 1
            End If
            CarregarGridCotacoes(ddlMes.SelectedValue, ddlAno.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProxMes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlMes.SelectedValue = 12 Then
                ddlAno.SelectedValue += 1
                ddlMes.SelectedValue = 1
            Else
                ddlMes.SelectedValue += 1
            End If
            CarregarGridCotacoes(ddlMes.SelectedValue, ddlAno.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlMes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarGridCotacoes(ddlMes.SelectedValue, ddlAno.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlAno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarGridCotacoes(ddlMes.SelectedValue, ddlAno.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlIndexador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarCotacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlDia_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarCotacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        If Not ValidaCotacao(True) Then
            Exit Sub
        End If

        Dim Cot As New [Lib].Negocio.Cotacao(ddlIndexador.SelectedValue, CDate(ddlAno.SelectedValue & "-" & ddlMes.SelectedValue & "-" & ddlDia.SelectedValue))
        Dim indiceAnterior As Decimal = Cot.Indice
        Cot.Indice = CDec(txtIndice.Text)

        If Cot.Realizado Then
            If Cot.chkLancamentos Then
                MsgBox(Me.Page, "Cotação realizada com lançamento(s) não pode ser alterada.")
                txtIndice.Text = indiceAnterior
                Exit Sub
            End If
        End If
        If Cot.UsuarioLiberacao Is Nothing Then
            If Not Funcoes.VerificaPermissao("CotacoesDeMoedas", "GRAVAR") Then
                MsgBox(Me.Page, "Usuario sem permissão para liberar cotação.")
                Exit Sub
            End If
            Cot.IUD = "I"
            Cot.UsuarioLiberacao = HttpContext.Current.Session("ssNomeUsuario")
            Cot.UsuarioLiberacaoData = DateTime.Now
        Else
            If Not Funcoes.VerificaPermissao("CotacoesDeMoedas", "ALTERAR") Then
                MsgBox(Me.Page, "Usuario sem permissão para alterar cotação.")
                Exit Sub
            End If
            Cot.IUD = "U"
            Cot.UsuarioAlteracao = HttpContext.Current.Session("ssNomeUsuario")
            Cot.UsuarioAlteracaoData = DateTime.Now
        End If

        Dim sqls As New ArrayList

        If Cot.Data.ToString("yyyy-MM-dd") = Date.Today.ToString("yyyy-MM-dd") Then
            Cot.Realizado = True
            Cot.SalvarSql(sqls)
            Sql = "Update Cotacoes set " & vbCrLf & _
                  " Indice = " & Str(Cot.Indice) & vbCrLf & _
                  " Where Data_Id      > '" & Cot.Data.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                  "   and Indexador_id =  " & Cot.CodigoIndexador
            sqls.Add(Sql)
        ElseIf Cot.Data.ToString("yyyy-MM-dd") < Date.Today.ToString("yyyy-MM-dd") Then
            Cot.Realizado = True
            Cot.SalvarSql(sqls)
        Else
            If rdRealizado.Checked Then
                Cot.Realizado = True
            Else
                Cot.Realizado = False
            End If
            Cot.SalvarSql(sqls)
        End If

        If Cot.Data.DayOfWeek = DayOfWeek.Monday Then
            Sql = "Update Cotacoes set " & vbCrLf & _
                  "    Indice               = " & Str(Cot.Indice) & vbCrLf & _
                  "   ,Realizado            ='" & IIf(Cot.Realizado, "R", "P") & "'" & vbCrLf & _
                  "   ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                  "   ,UsuarioAlteracaoData ='" & Date.Today.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                  " Where Data_Id      = '" & Cot.Data.AddDays(-1).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                  "   and Indexador_id =  " & Cot.CodigoIndexador
            sqls.Add(Sql)

            Sql = "Update Cotacoes set " & vbCrLf & _
                  "    Indice               = " & Str(Cot.Indice) & vbCrLf & _
                  "   ,Realizado            ='" & IIf(Cot.Realizado, "R", "P") & "'" & vbCrLf & _
                  "   ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                  "   ,UsuarioAlteracaoData ='" & Date.Today.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                  " Where Data_Id      = '" & Cot.Data.AddDays(-2).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                  "   and Indexador_id =  " & Cot.CodigoIndexador
            sqls.Add(Sql)
        End If

        If Banco.GravaBanco(sqls) Then
            MsgBox(Me.Page, "Cotação salva com Sucesso.", eTitulo.Sucess)
            CarregarGridCotacoes(ddlMes.SelectedValue, ddlAno.SelectedValue)
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("CotacoesDeMoedas", "RELATORIO") Then
                Sql = "  Select Data_Id as Data, isnull " & vbCrLf & _
                      "           ((SELECT  Indice FROM Cotacoes AS T1" & vbCrLf & _
                      "            WHERE (Indexador_Id = 1) AND (Cotacoes.Data_Id = T1.Data_Id)), 0) AS Compra, isnull" & vbCrLf & _
                      "           ((SELECT  Indice FROM Cotacoes AS T2" & vbCrLf & _
                      "            WHERE (Indexador_Id = 2) AND (Cotacoes.Data_Id = T2.Data_Id)), 0) AS Ptax, isnull" & vbCrLf & _
                      "           ((SELECT  Indice FROM Cotacoes AS T3" & vbCrLf & _
                      "             WHERE (Indexador_Id = 3) AND (Cotacoes.Data_Id = T3.Data_Id)), 0) AS Venda, isnull" & vbCrLf & _
                      "           ((SELECT  Indice FROM Cotacoes AS T4" & vbCrLf & _
                      "            WHERE (Indexador_Id = 111) AND (Cotacoes.Data_Id = T4.Data_Id)), 0) AS Real" & vbCrLf & _
                      "            FROM Cotacoes" & vbCrLf & _
                      "            WHERE Month(Data_Id) = " & ddlMes.SelectedValue & " and year(Data_Id) = " & ddlAno.SelectedValue & vbCrLf & _
                      "            GROUP BY Data_Id Order by Data_Id" & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "CotacoesDeMoedas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("ConsultaParametros", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_CotacoesDeMoedas", eExportType.PDF, parameters)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CotacoesDeMoedas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub imgCotacao_Click(sender As Object, e As ImageClickEventArgs) Handles imgCotacao.Click
        Try
            Dim dt As DateTime = CDate(ddlAno.SelectedValue & "-" & ddlMes.SelectedValue & "-" & ddlDia.SelectedValue)

            If dt.DayOfWeek = DayOfWeek.Monday Then
                dt = dt.AddDays(-3)
            Else
                dt = dt.AddDays(-1)
            End If

            If dt < DateTime.Now.AddDays(-20) Then
                MsgBox(Me.Page, "Não é possível consultar cotações inferiores à 20 dias.")
                Exit Sub
            End If
            If ddlIndexador.SelectedValue = "1" Then
                txtIndice.Text = String.Format("{0:N8}", BACEN.getVlrCompra(dt))
            ElseIf ddlIndexador.SelectedValue = "2" Then
                txtIndice.Text = String.Format("{0:N8}", BACEN.getVlrMedio(dt))
            ElseIf ddlIndexador.SelectedValue = "3" Then
                txtIndice.Text = String.Format("{0:N8}", BACEN.getVlrVenda(dt))
            Else
                MsgBox(Me.Page, "Valor de cotação não encontrado para o indexador selecionado.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class