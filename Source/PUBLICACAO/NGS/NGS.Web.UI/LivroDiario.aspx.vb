Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class LivroDiario
    Inherits BasePage

#Region "Atributos / Propriedades"

    Dim objEmpresa As New [Lib].Negocio.Cliente

#End Region

#Region "Eventos"

#Region "Page_load"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("LivroDiario", "ACESSAR") Then
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Contabil.aspx")
                Exit Sub
            End If
        End If
    End Sub

#End Region

#Region "Button"

    Protected Sub btnNome_Click(sender As Object, e As EventArgs) Handles btnNome.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me, "objLivroDiario" & HID.Value.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProcesso_Click(sender As Object, e As EventArgs) Handles btnProcesso.Click
        Try
            CarregarPopup()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Try
            FecharPopup()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnApuracao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnApuracao.Click
        Dim Erro As String = VallidarCampos()
        Try
            If String.IsNullOrEmpty(Erro) Then
                Try
                    If RelatorioLivroDiario(txtProcesso.Text, txtCNPJ.Text, txtDataIni.Text, txtDataFim.Text, txtLivro.Text, txtFolha.Text) Then
                        imbApuracao.ImageUrl = "images/certo.jpg"
                        MsgBox(Me.Page, "Livro diário contabil realizado com Sucesso.", eTitulo.Sucess)
                    Else
                        imbApuracao.ImageUrl = "images/erro2.jpg"
                    End If
                Catch ex As Exception
                    imbApuracao.ImageUrl = "images/erro2.jpg"
                End Try
            Else
                MsgBox(Me.Page, Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnTermo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnTermo.Click
        Dim Erro As String = VallidarCampos()
        Try
            If String.IsNullOrEmpty(Erro) Then
                Try
                    If Termos(txtProcesso.Text, txtLivro.Text, txtFolha.Text, txtDataIni.Text, txtDataFim.Text, txtCNPJ.Text) Then
                        imbTermo.ImageUrl = "images/certo.jpg"
                        MsgBox(Me.Page, "Termos de Abertura e Encerramento realizados com Sucesso.", eTitulo.Sucess)
                    Else
                        imbTermo.ImageUrl = "images/erro2.jpg"
                    End If
                Catch ex As Exception
                    imbTermo.ImageUrl = "images/erro2.jpg"
                End Try
            Else
                MsgBox(Me.Page, Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LivroDiario")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "GridView"

    Protected Sub gdvProcesso_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gdvProcesso.SelectedIndexChanged
        Try
            txtProcesso.Text = gdvProcesso.SelectedRow.Cells(1).Text
            txtDataIni.Text = Format(DateValue(gdvProcesso.SelectedRow.Cells(2).Text), "dd/MM/yyyy")
            txtDataFim.Text = Format(DateValue(gdvProcesso.SelectedRow.Cells(3).Text), "dd/MM/yyyy")
            txtLivro.Text = gdvProcesso.SelectedRow.Cells(4).Text
            txtFolha.Text = gdvProcesso.SelectedRow.Cells(5).Text
            FecharPopup()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#End Region

#Region "Métodos"

    Public Sub Limpar()
        Try
            Session.Remove("objLivroDiario" & HID.Value.ToString)
            HID.Value = Guid.NewGuid().ToString
            ucConsultaEmpresas.SetarHID(HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            objEmpresa = New [Lib].Negocio.Cliente
            objEmpresa = CType(Session("objLivroDiario" & HID.Value.ToString), [Lib].Negocio.Cliente)
            '1 = Empresa
            If objEmpresa IsNot Nothing Then
                With objEmpresa
                    txtCNPJ.Text = .Codigo.Substring(0, 8)
                    txtNome.Text = .Nome
                End With
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub CarregarPopup()
        Try
            SelecionaProcesso(txtCNPJ.Text)
            Dim strJavaScript As String = String.Empty
            strJavaScript &= IIf(True, "window.setTimeout(function() {", "")
            strJavaScript &= " " & Popup.ShowDialog("divProcesso", 750, "") & " "
            strJavaScript &= IIf(True, "}, " & 200 & ");", "")
            ScriptManager.RegisterClientScriptBlock(Page, Page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub FecharPopup()
        Try
            Popup.CloseDialog(Me, "divProcesso")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function VallidarCampos()
        Dim Erro As String = String.Empty

        If (String.IsNullOrEmpty(txtCNPJ.Text)) Then
            MsgBox(Me.Page, "O campo CNPJ é obrigatório.")
        End If
        If (String.IsNullOrEmpty(txtNome.Text)) Then
            MsgBox(Me.Page, "O campo nome é obrigatório.")
        End If
        If (String.IsNullOrEmpty(txtProcesso.Text)) Then
            MsgBox(Me.Page, "O campo processo é obrigatório.")
        End If
        If (String.IsNullOrEmpty(txtLivro.Text)) Then
            MsgBox(Me.Page, "O campo livro é obrigatório.")
        End If
        If (String.IsNullOrEmpty(txtFolha.Text)) Then
            MsgBox(Me.Page, "O campo follha é obrigatório.")
        End If
        If (String.IsNullOrEmpty(txtDataIni.Text)) Then
            MsgBox(Me.Page, "O campo período inicial é obrigatório.")
        End If
        If (String.IsNullOrEmpty(txtDataFim.Text)) Then
            MsgBox(Me.Page, "O campo período Final é obrigatório.")
        End If

        Return Erro
    End Function

    Public Function RelatorioLivroDiario(ByVal Processo As String, ByVal CnpjEmpresa As String, ByVal DataInicial As String, ByVal DataFinal As String, ByVal Livro As String, ByVal Pagina As String) As Boolean
        Dim cabecalho As String() = SelecaoEmpresaMatriz(CnpjEmpresa).Split("|")

        Dim fanalizado As Boolean = True

        Dim crpt As New ReportDocument()

        Try
            If Not GravarProcesso(Processo, Livro, Pagina, DataInicial, DataFinal, CnpjEmpresa) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If

            Dim sql As String
            Dim ds_livro As New DataSet

            sql = " SELECT Empresa.Reduzido AS Empresa,case When len(Clientes.Cliente_Id) = 14 then  Razao.Conta_Id + ' ' + ISNULL(Left (Clientes.Cliente_Id, 2) + '.' + Substring(Clientes.Cliente_Id,3 ,3 ) + '.' + Substring(Clientes.Cliente_Id,6 ,3 ) + '/' + Substring(Clientes.Cliente_Id,9 ,4 ) + '-' +  Substring(Clientes.Cliente_Id,13 ,2 ), '') else Razao.Conta_Id + ' ' + ISNULL(Substring(Clientes.Cliente_Id,3 ,3 ) + '.' + Substring(Clientes.Cliente_Id,6 ,3 ) + '.' + Substring(Clientes.Cliente_Id,9 ,3 ) + '-' +  Substring(Clientes.Cliente_Id,10 ,2 ), '') end  AS Conta, ISNULL(Clientes.Reduzido, '') as Cliente, Razao.Movimento_Id AS Movimento, REPLICATE('0', " & vbCrLf & _
                  " 4 - LEN(CAST(Razao.Lote_Id AS varchar))) + CAST(Razao.Lote_Id AS varchar) AS Lote, REPLICATE('0', 4 - LEN(CAST(Razao.Sequencia_Id AS varchar))) " & vbCrLf & _
                  " + CAST(Razao.Sequencia_Id AS varchar) AS Sequencia, " & vbCrLf & _
                  " Razao.DebitoOficial as Debito, Razao.CreditoOficial as Credito, " & vbCrLf & _
                  " Razao.Historico " & vbCrLf & _
                  " FROM Razao LEFT OUTER JOIN" & vbCrLf & _
                  " Clientes ON Razao.Cliente_Id = Clientes.Cliente_Id AND " & vbCrLf & _
                  " Razao.EndCliente_Id = Clientes.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                  " Clientes AS Empresa ON Razao.Empresa_Id = Empresa.Cliente_Id AND " & vbCrLf & _
                  " Razao.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
                  " WHERE  Razao.Empresa_Id LIKE '" & Microsoft.VisualBasic.Left(CnpjEmpresa, 8) & "%'" & vbCrLf & _
                  " And Razao.Movimento_Id between '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "'" & vbCrLf & _
                  " ORDER BY Razao.Empresa_Id, Movimento, Razao.Lote_Id, Razao.Sequencia_Id" & vbCrLf
            ds_livro = Banco.ConsultaDataSet(sql, "Razao")

            Dim CampoRoot As String() = HttpContext.Current.Server.MapPath("").Split("\")
            Dim RootSite As String = ""

            crpt.FileName = Server.MapPath("~/Reports/Cr_LivroDiario.rpt")
            crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo As String = Funcoes.GeraNomeArquivo & ".PDF"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

            crpt.SetDataSource(ds_livro)

            'crpt.SetParameterValue("NomeEmpresa", cabecalho(1))
            'crpt.SetParameterValue("Cidade", cabecalho(2))
            'crpt.SetParameterValue("Estado", cabecalho(3))
            'crpt.SetParameterValue("Titulo", "Livro Diário Nº: " & Format(CInt(Livro), "000"))
            'crpt.SetParameterValue("Periodo", "Período : " & Format(CDate(DataInicial), "dd/MM/yyyy") & " A " & Format(CDate(DataFinal), "dd/MM/yyyy"))
            'crpt.SetParameterValue("Pagina", Pagina)
            'crpt.SetParameterValue("CNPJ", cabecalho(4))

            Dim crparametervalues As ParameterValues
            Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

            crparameterfielddefinition = crparameterfielddefinitions.Item("NomeEmpresa")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = cabecalho(1)
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = cabecalho(2)
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Estado")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = cabecalho(3)
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Titulo")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Livro Diário Nº: " & Format(CInt(Livro), "000")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Periodo")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Período : " & Format(CDate(DataInicial), "dd/MM/yyyy") & " A " & Format(CDate(DataFinal), "dd/MM/yyyy")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Pagina")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = Pagina
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("CNPJ")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = cabecalho(4)
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

            If IO.File.Exists(arquivo) Then
                Funcoes.AbrirArquivo(Me.Page, NomeArquivo)
            End If
        Catch ex As Exception
            fanalizado = False
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try

        Return fanalizado
    End Function

    Private Function GravarProcesso(ByVal Processo As String, ByVal Livro As String, ByVal Folha As String, ByVal PeriodoInicial As String, ByVal PeriodoFinal As String, ByVal Empresa As String) As Boolean
        Dim strSQL As String
        strSQL = "SELECT * FROM ProcessoLivroDiario WHERE Empresa_Id = '" & Empresa & "' And Processo_Id = " & Processo

        Dim dsProcesso As DataSet = Banco.ConsultaDataSet(strSQL, "ProcessoLivroDiario")

        If dsProcesso.Tables(0).Rows.Count = 0 Then
            strSQL = "INSERT INTO ProcessoLivroDiario " & vbCrLf & _
                     "(Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado) " & vbCrLf & _
                     "VALUES ('" & Empresa & "', " & Processo & ", '" & PeriodoInicial.ToSqlDate() & "'," & vbCrLf & _
                     "'" & PeriodoFinal.ToSqlDate() & "', " & Livro & ", " & Folha & ", 'N')" & vbCrLf
        Else
            strSQL = "UPDATE ProcessoLivroDiario " & vbCrLf & _
                     "SET PeriodoInicial = '" & PeriodoInicial.ToSqlDate() & "'," & vbCrLf & _
                     "PeriodoFinal = '" & PeriodoFinal.ToSqlDate() & "'," & vbCrLf & _
                     "Livro = " & Livro & ", " & vbCrLf & _
                     "PaginaInicial = " & Folha & " " & vbCrLf & _
                     "WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                     "AND Processo_Id = " & Processo & vbCrLf
        End If

        Dim alSQL As New ArrayList
        alSQL.Add(strSQL)

        Return Banco.GravaBanco(alSQL)
    End Function

    Public Sub SelecionaProcesso(ByVal Empresa As String, Optional ByVal Processo As String = "")
        Dim strSQL As String

        Try
            strSQL = "SELECT Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado " & vbCrLf & _
                             "FROM ProcessoLivroDiario " & vbCrLf & _
                             "WHERE Empresa_Id = '" & Empresa & "' " & vbCrLf

            Dim dsProcesso As DataSet = Banco.ConsultaDataSet(strSQL, "ProcessoLivroDiario")
            gdvProcesso.DataSource = dsProcesso
            gdvProcesso.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function SelecaoEmpresaMatriz(ByVal CNPJ As String) As String
        Dim Texto As String = ""
        Dim Sql As String = ""

        Try
            If CNPJ <> "" Then
                Sql = "  SELECT SUBSTRING(Clientes.Cliente_Id, 1, 8) AS Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado,Clientes.Cliente_Id as Cnpj" & vbCrLf & _
                      " FROM Clientes INNER JOIN" & vbCrLf & _
                      " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                      " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf & _
                      " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id" & vbCrLf & _
                      " WHERE (ClientesXTipos.Tipo_Id = 1) AND (ClientesXEmpresas.Matriz = 'S') And (Clientes.Cliente_Id Like '" & CNPJ & "%')" & vbCrLf
            End If


            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                Texto = Dr("Cliente_Id") & "|" & Dr("Nome") & "|" & Dr("Cidade") & "|" & Dr("Estado") & "|" & Dr("Cnpj")
            Next
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

        Return Texto
    End Function

    Public Function Termos(ByVal Processo As String, ByVal Livro As String, ByVal Folha As String, ByVal PeriodoInicial As String, ByVal PeriodoFinal As String, ByVal Empresa As String, Optional ByVal Opcao As String = "01") As Boolean
        Dim Sql As String

        Dim finalizado As Boolean = True

        Dim crpt As New ReportDocument()

        Try
            If Not GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If

            Dim Ds As New DataSet
            Dim DIAINI As String
            Dim MESINI As String
            Dim ANOINI As String
            Dim DIAFIM As String
            Dim MESFIM As String
            Dim ANOFIM As String
            Dim LIVRO01 As String

            MESINI = LTrim(RTrim(Str(Month(PeriodoInicial))))
            If Val(MESINI) < 10 Then
                MESINI = "0" + MESINI
            End If
            MESINI = Mes(MESINI)
            DIAINI = LTrim(RTrim(Str(Day(PeriodoInicial))))
            If Val(DIAINI) < 10 Then
                DIAINI = "0" + DIAINI
            End If
            ANOINI = Str(Year(PeriodoInicial))

            MESFIM = LTrim(RTrim(Str(Month(PeriodoFinal))))
            If Val(MESFIM) < 10 Then
                MESFIM = "0" + MESFIM
            End If
            MESFIM = Mes(MESFIM)
            DIAFIM = LTrim(RTrim(Str(Day(PeriodoFinal))))
            If Val(DIAFIM) < 10 Then
                DIAFIM = "0" + DIAFIM
            End If
            ANOFIM = Str(Year(PeriodoFinal))
            LIVRO01 = Format(Val(Livro), "000")

            Dim mystr As String
            mystr = Format(Val(Folha), "000")
            Folha = mystr

            Sql = "  SELECT Clientes.Cliente_Id as CNPJ, Clientes.Nome as RazaoSocial, Clientes.Endereco, Clientes.Cidade, Clientes.Estado, Clientes.Inscricao, ClientesXEmpresas.RegistroNaJunta, " & vbCrLf & _
                  " ClientesXEmpresas.EstadodaJunta, ClientesXEmpresas.DatadeFundacao, ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, " & vbCrLf & _
                  " ClientesXEmpresas.CRCContador, ClientesXEmpresas.NomeDoTitular, ClientesXEmpresas.CPFTitular, ClientesXEmpresas.CondicaoTitular, " & vbCrLf & _
                  " Estados.Descricao as NomeDoEstado, 1 as Ordem" & vbCrLf & _
                  " FROM Clientes INNER JOIN" & vbCrLf & _
                  " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                  " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf & _
                  " Estados ON ClientesXEmpresas.EstadodaJunta = Estados.Estado_Id" & vbCrLf & _
                  " WHERE Clientes.Cliente_Id Like '" & Empresa & "%' And ClientesXEmpresas.Matriz = 'S'" & vbCrLf & _
                  " Union" & vbCrLf & _
                  " SELECT Clientes.Cliente_Id as CNPJ, Clientes.Nome as RazaoSocial, Clientes.Endereco, Clientes.Cidade, Clientes.Estado, Clientes.Inscricao, ClientesXEmpresas.RegistroNaJunta, " & vbCrLf & _
                  " ClientesXEmpresas.EstadodaJunta, ClientesXEmpresas.DatadeFundacao, ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, " & vbCrLf & _
                  " ClientesXEmpresas.CRCContador, ClientesXEmpresas.NomeDoTitular, ClientesXEmpresas.CPFTitular, ClientesXEmpresas.CondicaoTitular, " & vbCrLf & _
                  " Estados.Descricao as NomeDoEstado, 2 as Ordem" & vbCrLf & _
                  " FROM Clientes INNER JOIN" & vbCrLf & _
                  " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                  " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf & _
                  " Estados ON ClientesXEmpresas.EstadodaJunta = Estados.Estado_Id" & vbCrLf & _
                  " WHERE Clientes.Cliente_Id Like '" & Empresa & "%' And ClientesXEmpresas.Matriz = 'S'" & vbCrLf
            Ds = Banco.ConsultaDataSet(Sql, "Termos")

            Dim CampoRoot As String() = HttpContext.Current.Server.MapPath("").Split("\")
            Dim RootSite As String = ""

            crpt.FileName = Server.MapPath("~/Reports/Cr_TermosLivroDiario.rpt")
            crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo As String = Funcoes.GeraNomeArquivo & ".PDF"
            Dim arquivo As String = Server.MapPath(NomeArquivo)

            crpt.SetDataSource(Ds)

            Dim crparametervalues As ParameterValues
            Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

            crparameterfielddefinition = crparameterfielddefinitions.Item("Folha")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = Folha
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("DIAINI")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = DIAINI
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("MESINI")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = MESINI
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("ANOINI")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = ANOINI
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("DIAFIM")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = DIAFIM
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("MESFIM")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = MESFIM
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("ANOFIM")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = ANOFIM
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("LIVRO")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = LIVRO01
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

            If IO.File.Exists(arquivo) Then
                Funcoes.AbrirArquivo(Me.Page, NomeArquivo)
            End If
        Catch ex As Exception
            finalizado = False
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try

        Return finalizado
    End Function

    Function Mes(ByVal strMes As String) As String
        Dim linha As String = ""

        Try
            Select Case strMes
                Case "01"
                    linha = "janeiro"
                Case "02"
                    linha = "fevereiro"
                Case "03"
                    linha = "marco"
                Case "04"
                    linha = "abril"
                Case "05"
                    linha = "maio"
                Case "06"
                    linha = "junho"
                Case "07"
                    linha = "julho"
                Case "08"
                    linha = "agosto"
                Case "09"
                    linha = "setembro"
                Case "10"
                    linha = "outubro"
                Case "11"
                    linha = "novembro"
                Case "12"
                    linha = "dezembro"
            End Select
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

        Return linha
    End Function

#End Region

End Class