Imports System.Data
Imports System
Imports System.Console
Imports System.Globalization
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis


Public Class GradeSeguranca
    Inherits BasePage

    Dim objListBalanco As ListBalancoAuditadoMes

#Region "Session"
    Private Sub SessaoSalvaBalanco()
        Session("objListBalanco" & HID.Value) = objListBalanco
    End Sub

    Private Sub SessaoRecuperaBalanco()
        objListBalanco = CType(Session("objListBalanco" & HID.Value), [Lib].Negocio.ListBalancoAuditadoMes)
    End Sub
#End Region

#Region "Methods"

    Private Function PreencheMovimento(ByVal Empresa As String, ByVal EndEmpresa As String) As Boolean
        Dim sql As String = "Select Movimento_Id, Situacao" & vbCrLf &
                            "  from AcessosXMovimento" & vbCrLf &
                            " where Empresa_Id          ='" & Empresa & "'" & vbCrLf &
                            "   and EndEmpresa_Id       = " & EndEmpresa & vbCrLf &
                            "   and Year(Movimento_Id)  = " & CDate(txtDataMovimento.Text).Year & vbCrLf &
                            "   and Month(Movimento_Id) = " & CDate(txtDataMovimento.Text).Month & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "AcessosXMovimento")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            GridMovimento.DataSource = ds
            GridMovimento.DataBind()
            Return True
        Else
            GridMovimento.DataBind()
            Return False
        End If

    End Function

    Private Function getUltimoAnoCadastrado() As Integer
        Dim sql As String = "select Year(MAX(Movimento_Id)) as UltimoAno from AcessosXMovimento" & vbCrLf &
                            " where Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                            "   And EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1)
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "UltimoAno")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 AndAlso Not IsDBNull(ds.Tables(0).Rows(0)("UltimoAno")) Then
            Return CInt(ds.Tables(0).Rows(0)(0).ToString()) + 1
        Else
            Return DateTime.Now.Year
        End If
    End Function

    Public Overrides Sub Carregar(ByVal str As String)
        Try
            If Session("objGradeSeguranca" & HID.Value) IsNot Nothing Then
                If Not String.IsNullOrWhiteSpace(str) Then
                    InserirDatas(str.Split("-")(0), str.Split("-")(1), chkTodas.Checked)
                    PreencheMovimento(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
                    lnkConfirmar.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub InserirDatas(ByVal AnoInicial As Integer, ByVal AnoFinal As Integer, ByVal TodasAsFiliais As Boolean)
        If Funcoes.VerificaPermissao("GradeSeguranca", "GRAVAR") Then
            Dim arrSql As New ArrayList
            Dim sql As String = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf &
                                "    FROM Clientes " & vbCrLf &
                                "   INNER JOIN ClientesXEmpresas " & vbCrLf &
                                "      ON Clientes.Cliente_Id  = ClientesXEmpresas.Empresa_Id " & vbCrLf &
                                "     AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf
            If Not TodasAsFiliais Then
                sql &= "     AND Clientes.Cliente_Id= '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                       "     AND Clientes.Endereco_Id = " & ddlEmpresa.SelectedValue.Split("-")(1)
            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Clientes")

            For Each drEmpresa As DataRow In ds.Tables(0).Rows
                Dim myCal As Calendar = CultureInfo.InvariantCulture.Calendar

                For Ano As Integer = AnoInicial To AnoFinal
                    Dim myDT As New DateTime(Ano, 1, 1, New GregorianCalendar)

                    For Dia As Integer = 1 To myCal.GetDaysInYear(Ano)
                        arrSql.Add("INSERT INTO AcessosXMovimento (Empresa_Id, EndEmpresa_Id, Movimento_Id, Situacao) VALUES ('" & drEmpresa(0) & "', " & drEmpresa(1) & " ,'" & CDate(myDT).ToSqlDate() & "' ,'BLOQUEADO' )")
                        myDT = myCal.AddDays(myDT, 1)
                    Next Dia
                Next Ano
            Next

            If arrSql.Count > 0 Then
                If Banco.GravaBanco(arrSql) Then
                    MsgBox(Me.Page, "Registros gravados com Sucesso.", eTitulo.Sucess)
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub BuscarMovimentos(ByVal Movimento_Id As String, ByVal SituacaoMovimento As Button)
        Try
            Dim btnSitMovimento As Button = SituacaoMovimento
            Dim arrEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
            Dim Sqlarr As New ArrayList
            Dim sql As String = ""
            Dim dsEmpresas As New DataSet

            Try
                If btnSitMovimento.Text = "BLOQUEADO" Then
                    SessaoRecuperaBalanco()
                    If chkTodas.Checked Then
                        If objListBalanco.Auditado(arrEmpresa(0), arrEmpresa(1), CDate(Movimento_Id).Month, True) Then
                            MsgBox(Me.Page, "Alguma das Empresas da Selecao já esta com o mês auditado, abra a data empresa por empresa.")
                            Exit Sub
                        End If
                    Else
                        If objListBalanco.Auditado(arrEmpresa(0), arrEmpresa(1), CDate(Movimento_Id).Month) Then
                            MsgBox(Me.Page, "A Empresa esta com o mês auditado, Para abrir esta data solicite ao responsavel contabil da Empresa.")
                            Exit Sub
                        End If
                    End If
                End If

                Dim situacao As String = IIf(btnSitMovimento.Text = "BLOQUEADO", "LIBERADO", "BLOQUEADO")
                btnSitMovimento.Text = situacao

                sql = " Update AcessosXMovimento " & vbCrLf &
                       "           Set Situacao      = '" & situacao & "' " & vbCrLf &
                       "       WHERE   Movimento_Id  = '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf &
                       "           AND Empresa_Id    = '" & arrEmpresa(0) & "'" & vbCrLf &
                       "           AND EndEmpresa_Id = " & arrEmpresa(1) & vbCrLf
                Sqlarr.Add(sql)

                sql = " Delete From AcessosXProcessos " & vbCrLf &
                        "       Where   Movimento_Id  = '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf &
                        "           AND EndEmpresa_Id =  " & arrEmpresa(1) & vbCrLf &
                        "           AND Empresa_Id    = '" & arrEmpresa(0) & "'" & vbCrLf
                Sqlarr.Add(sql)

                sql = " Delete From AcessosXUsuarios " & vbCrLf &
                    "       WHERE   Movimento_Id  = '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf &
                    "           AND EndEmpresa_Id =  " & arrEmpresa(1) & vbCrLf &
                    "           AND Empresa_Id    = '" & arrEmpresa(0) & "'" & vbCrLf
                Sqlarr.Add(sql)

                If chkTodas.Checked = True Then
                    sql = "  SELECT Clientes.Cliente_Id , Clientes.Endereco_Id " & vbCrLf &
                          "     FROM   Clientes " & vbCrLf &
                          "         INNER JOIN ClientesXEmpresas " & vbCrLf &
                          "              ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id" & vbCrLf &
                          "             AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                          "     where  Clientes.Cliente_Id <> '" & arrEmpresa(0) & "' or Clientes.Endereco_Id <> " & arrEmpresa(1) & vbCrLf

                    dsEmpresas = Banco.ConsultaDataSet(sql, "Clientes")

                    For Each drEmpresa As DataRow In dsEmpresas.Tables(0).Rows

                        sql = " Update AcessosXMovimento " & vbCrLf &
                              "         Set Situacao = '" & situacao & "' " & vbCrLf &
                              "     WHERE   Movimento_Id  = '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf &
                              "         AND Empresa_Id    = '" & drEmpresa("Cliente_Id") & "'" & vbCrLf &
                              "         AND EndEmpresa_Id =  " & drEmpresa("Endereco_Id") & vbCrLf
                        Sqlarr.Add(sql)

                        sql = " Delete From AcessosXProcessos" & vbCrLf &
                            "       Where   Movimento_Id  = '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf &
                            "           AND EndEmpresa_Id =  " & drEmpresa("Endereco_Id") & vbCrLf &
                            "           AND Empresa_Id    = '" & drEmpresa("Cliente_Id") & "'" & vbCrLf
                        Sqlarr.Add(sql)

                        sql = " Delete From AcessosXUsuarios " & vbCrLf &
                            "       WHERE   Movimento_Id = '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf &
                            "           AND EndEmpresa_Id = " & drEmpresa("Endereco_Id") & vbCrLf &
                            "           AND Empresa_Id = '" & drEmpresa("Cliente_Id") & "'" & vbCrLf
                        Sqlarr.Add(sql)
                    Next
                End If

                If situacao = "LIBERADO" Then
                    btnSitMovimento.BackColor = Drawing.Color.Green
                    btnSitMovimento.BorderColor = Drawing.Color.Green
                    btnSitMovimento.ForeColor = Drawing.Color.White
                Else
                    btnSitMovimento.BackColor = Drawing.Color.Red
                    btnSitMovimento.BorderColor = Drawing.Color.Red
                    btnSitMovimento.ForeColor = Drawing.Color.Yellow
                End If

                If Sqlarr.Count > 0 Then
                    If Banco.GravaBanco(Sqlarr) Then
                        BuscarProcessos(Movimento_Id.ToSqlDate(), btnSitMovimento.Text)
                        GridMovimento.SelectedIndex = GridMovimento.SelectedIndex

                        Dim dsaux As New DataSet
                        GridUsuarios.DataSource = Nothing
                        GridUsuarios.DataBind()

                        btnSitMovimento.Focus()
                    End If
                End If
            Catch ex As Exception
                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString))
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscarProcessos(ByVal Movimento_Id As String, ByVal SituacaoMovimento As String)
        Dim sql As String
        Dim dv As DataView
        Dim ds As New DataSet
        Dim dr As DataRow
        Dim drProcessos As DataRow
        Dim dsAux As New DataSet
        Dim dtProcessos As New DataTable("Processos")

        Dim arrEmpresa() As String

        arrEmpresa = ddlEmpresa.SelectedValue.Split("-")

        dtProcessos.Columns.Add("Processo_Id", GetType(String))
        dtProcessos.Columns.Add("Situacao", GetType(String))
        dtProcessos.Columns.Add("Movimento_Id", GetType(String))

        sql = "select Processo_Id,'' as Situacao from ProcessosXAcessos "

        ds = Banco.ConsultaDataSet(sql, "ProcessosXAcessos")

        For Each dr In ds.Tables(0).Rows
            drProcessos = dtProcessos.NewRow()
            drProcessos(0) = dr("Processo_Id")
            sql = "select * from AcessosXProcessos WHERE (Movimento_Id = '" & Format(CDate(Movimento_Id), "yyyy-MM-dd") & "') AND (Processo_Id='" & dr("Processo_Id") & "') AND (EndEmpresa_Id = " & arrEmpresa(1) & ") AND (Empresa_Id = '" & arrEmpresa(0) & "')"
            dsAux = Banco.ConsultaDataSet(sql, "AcessosXProcessos")

            If SituacaoMovimento = "LIBERADO" Then
                If dsAux.Tables(0).Rows.Count > 0 Then
                    drProcessos(1) = "BLOQUEADO"
                Else
                    drProcessos(1) = "LIBERADO"
                End If
            Else
                If dsAux.Tables(0).Rows.Count > 0 Then
                    drProcessos(1) = "LIBERADO"
                Else
                    drProcessos(1) = "BLOQUEADO"
                End If
            End If
            drProcessos(2) = Movimento_Id

            dtProcessos.Rows.Add(drProcessos)

            BuscarUsuarios(Movimento_Id, dr("Processo_Id"), dr("Processo_Id"))
        Next

        dv = New DataView(dtProcessos)

        GridProcessos.DataSource = dv
        GridProcessos.DataBind()
    End Sub

    Private Sub BuscarUsuarios(ByVal Movimento_Id As String, ByVal Processo As String, ByVal SituacaoProcesso As String)
        Dim dtUsuarios As New DataTable("Usuarios")
        Dim dv As DataView

        dtUsuarios.Columns.Add("NomeCompleto", GetType(String))
        dtUsuarios.Columns.Add("Situacao", GetType(String))
        dtUsuarios.Columns.Add("Usuario_Id", GetType(String))
        dtUsuarios.Columns.Add("Movimento", GetType(String))
        dtUsuarios.Columns.Add("Processo", GetType(String))

        Dim ds As DataSet = New DataSet()
        ds.Tables.Add(dtUsuarios)

        Dim sql As String = "select Usuario_Id, NomeCompleto, '' as Situacao from Usuarios "

        ds = Banco.ConsultaDataSet(sql, "Usuarios")

        For Each dr As DataRow In ds.Tables(0).Rows
            Dim drUsuarios As DataRow = dtUsuarios.NewRow()
            drUsuarios(0) = dr("NomeCompleto")

            sql = "select Empresa_Id, EndEmpresa_Id, Movimento_Id, Processo_Id, Usuario_ID from AcessosXUsuarios " & vbCrLf &
                "       Where   Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                "           AND EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                "           And Movimento_Id  = '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf &
                "           AND Processo_Id   = '" & Processo & "'" & vbCrLf &
                "           And Usuario_ID    = '" & dr("Usuario_Id") & "'" & vbCrLf

            Dim dsAux As DataSet = Banco.ConsultaDataSet(sql, "AcessosXProcessos")

            If SituacaoProcesso = "BLOQUEADO" Then
                If dsAux.Tables(0).Rows.Count > 0 Then
                    drUsuarios(1) = "LIBERADO"
                Else
                    drUsuarios(1) = "BLOQUEADO"
                End If
            Else
                If dsAux.Tables(0).Rows.Count > 0 Then
                    drUsuarios(1) = "BLOQUEADO"
                Else
                    drUsuarios(1) = "LIBERADO"
                End If
            End If

            drUsuarios(2) = dr("Usuario_Id")
            drUsuarios(3) = Movimento_Id
            drUsuarios(4) = Processo
            dtUsuarios.Rows.Add(drUsuarios)
        Next


        dv = New DataView(dtUsuarios)

        GridUsuarios.DataSource = dv
        GridUsuarios.DataBind()
    End Sub

    Private Sub limparGridViews()
        GridUsuarios.DataSource = Nothing
        GridUsuarios.DataBind()

        GridProcessos.DataSource = Nothing
        GridProcessos.DataBind()
    End Sub

#End Region

#Region "Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("GradeSeguranca", "ACESSAR") Then
                    HID.Value = Guid.NewGuid.ToString()
                    ucInputDate.SetarHID(HID.Value)

                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
                    ddlEmpresa.SelectedValue = UsuarioServidor.CodigoEmpresa & "-" & UsuarioServidor.EnderecoEmpresa

                    txtDataMovimento.Text = Format(Today, "dd/MM/yyyy")
                    HAno.Value = Today.Year

                    objListBalanco = New ListBalancoAuditadoMes
                    objListBalanco.Carregar(HAno.Value)
                    SessaoSalvaBalanco()

                    PreencheMovimento(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    ' ** GridMovimentos
    Protected Sub GridMovimento_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridMovimento.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim btn As Button = e.Row.FindControl("btnSituacao")

                If btn.Text = UCase("Bloqueado") Then
                    btn.BackColor = Drawing.Color.Red
                    btn.BorderColor = Drawing.Color.Red
                    btn.ForeColor = Drawing.Color.Yellow
                Else
                    btn.BackColor = Drawing.Color.Green
                    btn.ForeColor = Drawing.Color.White
                    btn.BorderColor = Drawing.Color.Green
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridMovimento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("GradeSeguranca", "ALTERAR") Then
                limparGridViews()

                Dim btnSituacao As Button = GridMovimento.SelectedRow.FindControl("btnSituacao")

                BuscarMovimentos(GridMovimento.SelectedRow.Cells(1).Text(), btnSituacao)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridMovimento_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridMovimento.RowDeleting
        Try
            Dim b As Button = GridMovimento.Rows(e.RowIndex).FindControl("btnSituacao")
            BuscarProcessos(GridMovimento.Rows(e.RowIndex).Cells(1).Text(), b.Text)
            GridMovimento.SelectedIndex = e.RowIndex

            GridUsuarios.DataSource = Nothing
            GridUsuarios.DataBind()
            b.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    ' ** GridProcessos
    Protected Sub GridProcessos_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridProcessos.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim btn As Button = e.Row.FindControl("btnSituacao")

                If btn.Text = UCase("Bloqueado") Then
                    btn.BackColor = Drawing.Color.Red
                    btn.BorderColor = Drawing.Color.Red
                    btn.ForeColor = Drawing.Color.Yellow
                Else
                    btn.BackColor = Drawing.Color.Green
                    btn.ForeColor = Drawing.Color.White
                    btn.BorderColor = Drawing.Color.Green
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridProcessos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("GradeSeguranca", "ALTERAR") Then
                Dim sql As String = ""
                Dim btnSitMovimento As Button = GridMovimento.SelectedRow.FindControl("btnSituacao")
                Dim btnSitProcessos As Button = GridProcessos.SelectedRow.FindControl("btnSituacao")
                Dim movimeto As Label = GridProcessos.SelectedRow.FindControl("Movimento_Id")
                Dim arrProcessos As New ArrayList
                Dim dsEmpresas As New DataSet
                Dim drEmpresa As DataRow

                Dim arrEmpresa() As String
                arrEmpresa = ddlEmpresa.SelectedValue.Split("-")

                If btnSitProcessos.Text = "LIBERADO" Then
                    btnSitProcessos.Text = UCase("Bloqueado")
                    btnSitProcessos.BackColor = Drawing.Color.Red
                    btnSitProcessos.BorderColor = Drawing.Color.Red
                    btnSitProcessos.ForeColor = Drawing.Color.Yellow
                    sql = " INSERT INTO AcessosXProcessos" & vbCrLf &
                          " (Empresa_Id, EndEmpresa_Id, Movimento_Id, Processo_Id)" & vbCrLf &
                          " VALUES('" & arrEmpresa(0) & "', " & arrEmpresa(1) & " ,'" & movimeto.Text.ToSqlDate() & "' ,'" & GridProcessos.SelectedRow.Cells(1).Text() & "') " & vbCrLf
                    arrProcessos.Add(sql)
                    sql = " Delete From AcessosXUsuarios WHERE (Movimento_Id = '" & movimeto.Text.ToSqlDate() & "') AND (Processo_Id='" & GridProcessos.SelectedRow.Cells(1).Text() & "') AND (EndEmpresa_Id = " & arrEmpresa(1) & ") AND (Empresa_Id = '" & arrEmpresa(0) & "')"
                    arrProcessos.Add(sql)
                    If chkTodas.Checked = True Then
                        sql = "  SELECT Clientes.Cliente_Id , Clientes.Endereco_Id " & vbCrLf &
                              " FROM   Clientes " & vbCrLf &
                              " INNER JOIN ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                              " where  ((Clientes.Cliente_Id <> '" & arrEmpresa(0) & "' ) or (Clientes.Endereco_Id <> " & arrEmpresa(1) & "))" & vbCrLf
                        dsEmpresas = Banco.ConsultaDataSet(sql, "Clientes")
                        For Each drEmpresa In dsEmpresas.Tables(0).Rows
                            sql = " INSERT INTO AcessosXProcessos" & vbCrLf &
                                  " (Empresa_Id, EndEmpresa_Id, Movimento_Id, Processo_Id)" & vbCrLf &
                                  " VALUES('" & drEmpresa("Cliente_Id") & "', " & drEmpresa("Endereco_Id") & " ,'" & movimeto.Text.ToSqlDate() & "' ,'" & GridProcessos.SelectedRow.Cells(1).Text() & "') " & vbCrLf
                            arrProcessos.Add(sql)
                            sql = " Delete From AcessosXUsuarios WHERE (Movimento_Id = '" & movimeto.Text.ToSqlDate() & "') AND (Processo_Id='" & GridProcessos.SelectedRow.Cells(1).Text() & "') AND (EndEmpresa_Id = " & drEmpresa("Endereco_Id") & ") AND (Empresa_Id = '" & drEmpresa("Cliente_Id") & "')"
                            arrProcessos.Add(sql)
                        Next
                    End If
                ElseIf btnSitProcessos.Text = "BLOQUEADO" Then
                    btnSitProcessos.Text = UCase("Liberado")
                    btnSitProcessos.BackColor = Drawing.Color.Green
                    btnSitProcessos.BorderColor = Drawing.Color.Green
                    btnSitProcessos.ForeColor = Drawing.Color.White
                    sql &= "Delete From AcessosXProcessos WHERE (Movimento_Id = '" & movimeto.Text.ToSqlDate() & "') AND (Processo_Id='" & GridProcessos.SelectedRow.Cells(1).Text() & "') AND (EndEmpresa_Id = " & arrEmpresa(1) & ") AND (Empresa_Id = '" & arrEmpresa(0) & "')"
                    arrProcessos.Add(sql)
                    sql = " Delete From AcessosXUsuarios WHERE (Movimento_Id = '" & movimeto.Text.ToSqlDate() & "') AND (Processo_Id='" & GridProcessos.SelectedRow.Cells(1).Text() & "') AND (EndEmpresa_Id = " & arrEmpresa(1) & ") AND (Empresa_Id = '" & arrEmpresa(0) & "')"
                    arrProcessos.Add(sql)

                    If chkTodas.Checked = True Then
                        sql = "  SELECT Clientes.Cliente_Id , Clientes.Endereco_Id " & vbCrLf &
                              " FROM   Clientes " & vbCrLf &
                              " INNER JOIN ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                              " where  ((Clientes.Cliente_Id <> '" & arrEmpresa(0) & "' ) or (Clientes.Endereco_Id <> " & arrEmpresa(1) & "))" & vbCrLf
                        dsEmpresas = Banco.ConsultaDataSet(sql, "Clientes")
                        For Each drEmpresa In dsEmpresas.Tables(0).Rows
                            sql &= "Delete From AcessosXProcessos WHERE (Movimento_Id = '" & movimeto.Text.ToSqlDate() & "') AND (Processo_Id='" & GridProcessos.SelectedRow.Cells(1).Text() & "') AND (EndEmpresa_Id = " & drEmpresa("Endereco_Id") & ") AND (Empresa_Id = '" & drEmpresa("Cliente_Id") & "')"
                            arrProcessos.Add(sql)
                            sql = " Delete From AcessosXUsuarios WHERE (Movimento_Id = '" & movimeto.Text.ToSqlDate() & "') AND (Processo_Id='" & GridProcessos.SelectedRow.Cells(1).Text() & "') AND (EndEmpresa_Id = " & drEmpresa("Endereco_Id") & ") AND (Empresa_Id = '" & drEmpresa("Cliente_Id") & "')"
                            arrProcessos.Add(sql)
                        Next
                    End If
                End If

                If arrProcessos.Count > 0 Then
                    If Banco.GravaBanco(arrProcessos) Then
                        BuscarUsuarios(movimeto.Text, GridProcessos.SelectedRow.Cells(1).Text(), btnSitProcessos.Text)
                        GridProcessos.SelectedIndex = GridProcessos.SelectedIndex
                        btnSitProcessos.Focus()
                        btnSitMovimento.Focus()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        Exit Sub
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridProcessos_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles GridProcessos.RowDeleting
        Try
            Dim movimento As Label = GridProcessos.Rows(e.RowIndex).FindControl("Movimento_Id")
            Dim btnMov As Button = GridMovimento.SelectedRow.FindControl("btnSituacao")
            Dim b As Button = GridProcessos.Rows(e.RowIndex).FindControl("btnSituacao")
            BuscarUsuarios(movimento.Text, GridProcessos.Rows(e.RowIndex).Cells(1).Text(), b.Text)
            GridProcessos.SelectedIndex = e.RowIndex

            b.Focus()
            btnMov.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    ' ** GridUsuarios
    Protected Sub GridUsuarios_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridUsuarios.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim btn As Button = e.Row.FindControl("btnSituacao")

                If btn.Text = UCase("Bloqueado") Then
                    btn.BackColor = Drawing.Color.Red
                    btn.BorderColor = Drawing.Color.Red
                    btn.ForeColor = Drawing.Color.Yellow
                Else
                    btn.BackColor = Drawing.Color.Green
                    btn.ForeColor = Drawing.Color.White
                    btn.BorderColor = Drawing.Color.Green
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridUsuarios_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridUsuarios.SelectedIndexChanged
        Try
            If Funcoes.VerificaPermissao("GradeSeguranca", "ALTERAR") Then
                Dim sql As String = ""
                Dim btnSitMovimento As Button = GridMovimento.SelectedRow.FindControl("btnSituacao")
                Dim btnSitProcesso As Button = GridProcessos.SelectedRow.FindControl("btnSituacao")
                Dim btnSitUsuarios As Button = GridUsuarios.SelectedRow.FindControl("btnSituacao")
                Dim movimeto As Label = GridUsuarios.SelectedRow.FindControl("Movimento")
                Dim usuario As Label = GridUsuarios.SelectedRow.FindControl("Usuario_Id")
                Dim processo As Label = GridUsuarios.SelectedRow.FindControl("Processo")

                Dim arrEmpresa() As String
                arrEmpresa = ddlEmpresa.SelectedValue.Split("-")

                Dim arrProcessos As New ArrayList

                If btnSitUsuarios.Text = "LIBERADO" Then
                    Dim dsEmpresas As New DataSet
                    Dim drEmpresa As DataRow
                    btnSitUsuarios.Text = "BLOQUEADO"
                    btnSitUsuarios.BackColor = Drawing.Color.Red
                    btnSitUsuarios.BorderColor = Drawing.Color.Red
                    btnSitUsuarios.ForeColor = Drawing.Color.Yellow
                    If btnSitProcesso.Text = "LIBERADO" Then
                        sql = " INSERT INTO AcessosXUsuarios" & vbCrLf &
                              " (Empresa_Id, EndEmpresa_Id, Movimento_Id, Processo_Id,Usuario_Id)" & vbCrLf &
                              " VALUES('" & arrEmpresa(0) & "', " & arrEmpresa(1) & " ,'" & Format(CDate(movimeto.Text), "yyyy-MM-dd") & "' ,'" & processo.Text & "','" & usuario.Text & "') " & vbCrLf
                        arrProcessos.Add(sql)
                        If chkTodas.Checked = True Then

                            sql = "  SELECT Clientes.Cliente_Id , Clientes.Endereco_Id " & vbCrLf &
                                  " FROM   Clientes " & vbCrLf &
                                  " INNER JOIN ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                                  " where  ((Clientes.Cliente_Id <> '" & arrEmpresa(0) & "' ) or (Clientes.Endereco_Id <> " & arrEmpresa(1) & "))" & vbCrLf
                            dsEmpresas = Banco.ConsultaDataSet(sql, "Clientes")
                            For Each drEmpresa In dsEmpresas.Tables(0).Rows
                                sql = " INSERT INTO AcessosXUsuarios" & vbCrLf &
                                      " (Empresa_Id, EndEmpresa_Id, Movimento_Id, Processo_Id,Usuario_Id)" & vbCrLf &
                                      " VALUES('" & drEmpresa("Cliente_Id") & "', " & drEmpresa("Endereco_Id") & " ,'" & Format(CDate(movimeto.Text), "yyyy-MM-dd") & "' ,'" & processo.Text & "','" & usuario.Text & "') " & vbCrLf
                                arrProcessos.Add(sql)
                            Next
                        End If
                    Else
                        btnSitUsuarios.Text = "BLOQUEADO"
                        btnSitUsuarios.BackColor = Drawing.Color.Red
                        btnSitUsuarios.BorderColor = Drawing.Color.Red
                        btnSitUsuarios.ForeColor = Drawing.Color.Yellow
                        sql &= "Delete From AcessosXUsuarios WHERE (Usuario_Id='" & usuario.Text & "') and (Movimento_Id = '" & Format(CDate(movimeto.Text), "yyyy-MM-dd") & "') AND (Processo_Id='" & processo.Text & "') AND (EndEmpresa_Id = " & arrEmpresa(1) & ") AND (Empresa_Id = '" & arrEmpresa(0) & "')"
                        arrProcessos.Add(sql)
                        If chkTodas.Checked = True Then

                            sql = "  SELECT Clientes.Cliente_Id , Clientes.Endereco_Id " & vbCrLf &
                                  " FROM   Clientes " & vbCrLf &
                                  " INNER JOIN ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                                  " where  ((Clientes.Cliente_Id <> '" & arrEmpresa(0) & "' ) or (Clientes.Endereco_Id <> " & arrEmpresa(1) & "))" & vbCrLf
                            dsEmpresas = Banco.ConsultaDataSet(sql, "Clientes")
                            For Each drEmpresa In dsEmpresas.Tables(0).Rows
                                sql &= "Delete From AcessosXUsuarios WHERE (Usuario_Id='" & usuario.Text & "') and (Movimento_Id = '" & Format(CDate(movimeto.Text), "yyyy-MM-dd") & "') AND (Processo_Id='" & processo.Text & "') AND (EndEmpresa_Id = " & drEmpresa("Endereco_Id") & ") AND (Empresa_Id = '" & drEmpresa("Cliente_Id") & "')"
                                arrProcessos.Add(sql)
                            Next
                        End If
                    End If
                ElseIf btnSitUsuarios.Text = "BLOQUEADO" Then
                    Dim dsEmpresas As New DataSet
                    Dim drEmpresa As DataRow
                    btnSitUsuarios.Text = "LIBERADO"
                    btnSitUsuarios.BackColor = Drawing.Color.Green
                    btnSitUsuarios.BorderColor = Drawing.Color.Green
                    btnSitUsuarios.ForeColor = Drawing.Color.White
                    If btnSitProcesso.Text = "LIBERADO" Then
                        sql &= "Delete From AcessosXUsuarios WHERE (Usuario_Id='" & usuario.Text & "') and (Movimento_Id = '" & Format(CDate(movimeto.Text), "yyyy-MM-dd") & "') AND (Processo_Id='" & processo.Text & "') AND (EndEmpresa_Id = " & arrEmpresa(1) & ") AND (Empresa_Id = '" & arrEmpresa(0) & "')"
                        arrProcessos.Add(sql)
                        If chkTodas.Checked = True Then
                            sql = "  SELECT Clientes.Cliente_Id , Clientes.Endereco_Id " & vbCrLf &
                                  " FROM   Clientes " & vbCrLf &
                                  " INNER JOIN ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                                  " where  ((Clientes.Cliente_Id <> '" & arrEmpresa(0) & "' ) or (Clientes.Endereco_Id <> " & arrEmpresa(1) & "))" & vbCrLf
                            dsEmpresas = Banco.ConsultaDataSet(sql, "Clientes")
                            For Each drEmpresa In dsEmpresas.Tables(0).Rows
                                sql &= "Delete From AcessosXUsuarios WHERE (Usuario_Id='" & usuario.Text & "') and (Movimento_Id = '" & Format(CDate(movimeto.Text), "yyyy-MM-dd") & "') AND (Processo_Id='" & processo.Text & "') AND (EndEmpresa_Id = " & drEmpresa("Endereco_Id") & ") AND (Empresa_Id = '" & drEmpresa("Cliente_Id") & "')"
                                arrProcessos.Add(sql)
                            Next
                        End If
                    Else
                        sql = " INSERT INTO AcessosXUsuarios" & vbCrLf &
                              " (Empresa_Id, EndEmpresa_Id, Movimento_Id, Processo_Id,Usuario_Id)" & vbCrLf &
                              " VALUES('" & arrEmpresa(0) & "', " & arrEmpresa(1) & " ,'" & Format(CDate(movimeto.Text), "yyyy-MM-dd") & "' ,'" & processo.Text & "','" & usuario.Text & "') " & vbCrLf
                        arrProcessos.Add(sql)
                        If chkTodas.Checked = True Then
                            sql = "  SELECT Clientes.Cliente_Id , Clientes.Endereco_Id " & vbCrLf &
                                  " FROM   Clientes " & vbCrLf &
                                  " INNER JOIN ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf &
                                  " where  ((Clientes.Cliente_Id <> '" & arrEmpresa(0) & "' ) or (Clientes.Endereco_Id <> " & arrEmpresa(1) & "))" & vbCrLf
                            dsEmpresas = Banco.ConsultaDataSet(sql, "Clientes")
                            For Each drEmpresa In dsEmpresas.Tables(0).Rows
                                sql = " INSERT INTO AcessosXUsuarios" & vbCrLf &
                                      " (Empresa_Id, EndEmpresa_Id, Movimento_Id, Processo_Id,Usuario_Id)" & vbCrLf &
                                      " VALUES('" & drEmpresa("Cliente_Id") & "', " & drEmpresa("Endereco_Id") & " ,'" & Format(CDate(movimeto.Text), "yyyy-MM-dd") & "' ,'" & processo.Text & "','" & usuario.Text & "') " & vbCrLf
                                arrProcessos.Add(sql)
                            Next
                        End If
                    End If
                End If
                If arrProcessos.Count > 0 Then
                    If Banco.GravaBanco(arrProcessos) Then
                        GridUsuarios.SelectedIndex = GridUsuarios.SelectedIndex
                        btnSitMovimento.Focus()
                        btnSitProcesso.Focus()
                        btnSitUsuarios.Focus()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        Exit Sub
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataMovimento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataMovimento.TextChanged
        Try
            If IsDate(txtDataMovimento.Text) Then
                Dim arrEmpresa() As String
                arrEmpresa = ddlEmpresa.SelectedValue.Split("-")

                If PreencheMovimento(arrEmpresa(0), arrEmpresa(1)) Then
                    GridUsuarios.DataSource = Nothing
                    GridUsuarios.DataBind()

                    GridProcessos.DataSource = Nothing
                    GridProcessos.DataBind()

                    If HAno.Value <> CDate(txtDataMovimento.Text).Year Then
                        HAno.Value = CDate(txtDataMovimento.Text).Year
                        objListBalanco = New ListBalancoAuditadoMes()
                        objListBalanco.Carregar(HAno.Value)
                        SessaoSalvaBalanco()
                    End If
                Else
                    MsgBox(Me.Page, "Datas não cadastradas.")
                    'ucInputDate.Limpar()
                    'ucInputDate.SetarAnoInicial(getUltimoAnoCadastrado())
                    'Popup.InputOfDate(Me.Page, "objGradeSeguranca" & HID.Value, "txtAno2")
                    lnkConfirmar.Enabled = True

                End If
            Else
                MsgBox(Me.Page, "Data inválida.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lstEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        Try
            Dim arrEmpresa() As String
            arrEmpresa = ddlEmpresa.SelectedValue.Split("-")
            PreencheMovimento(arrEmpresa(0), arrEmpresa(1))
            lnkConfirmar.Enabled = True

            GridUsuarios.DataSource = Nothing
            GridUsuarios.DataBind()

            GridProcessos.DataSource = Nothing
            GridProcessos.DataBind()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        ucInputDate.Limpar()
        ucInputDate.SetarAnoInicial(getUltimoAnoCadastrado())
        Popup.InputOfDate(Me.Page, "objGradeSeguranca" & HID.Value, "txtAno2")
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "GradeSeguranca")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub btnLiberar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("GradeSeguranca", "ALTERAR") Then
                limparGridViews()

                Dim btnSituacao As Button

                For Each rowMovimento As GridViewRow In GridMovimento.Rows
                    btnSituacao = rowMovimento.FindControl("btnSituacao")

                    If btnSituacao.Text = "BLOQUEADO" Then
                        BuscarMovimentos(rowMovimento.Cells(1).Text, btnSituacao)
                    End If
                Next
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBloqueado_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Funcoes.VerificaPermissao("GradeSeguranca", "ALTERAR") Then
                limparGridViews()

                Dim btnSituacao As Button

                For Each rowMovimento As GridViewRow In GridMovimento.Rows
                    btnSituacao = rowMovimento.FindControl("btnSituacao")

                    If btnSituacao.Text = "LIBERADO" Then
                        BuscarMovimentos(rowMovimento.Cells(1).Text, btnSituacao)
                    End If
                Next
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class