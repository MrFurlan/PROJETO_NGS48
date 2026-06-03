Imports System.Data
Imports System.Collections.Generic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucDarDiferencialDeAliquota
    Inherits BaseUserControl

    Dim ListaDarDif As [Lib].Negocio.ListDarDiferencialDeAliquota
    Dim objDarDifAliquota As [Lib].Negocio.DarDiferencialDeAliquota
    Dim Mes, Ano As Integer
    Dim ObjEmpresa As [Lib].Negocio.Cliente

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Limpar()
        End If
    End Sub

    Private Sub SessaoSalvaDar()
        Session("objDarDifAliquota" & HID.Value) = objDarDifAliquota
    End Sub

    Private Sub SessaoSalvaMesAno()
        Session("Mes" & HID.Value) = Mes
        Session("Ano" & HID.Value) = Ano
    End Sub

    Private Sub SessaoSalvaEmpresa()
        Session("objEmpresa" & HID.Value) = ObjEmpresa
    End Sub

    Private Sub SessaoSalvaListaDarDif()
        Session("ListaDarDifAliquota" & HID.Value) = ListaDarDif
    End Sub

    Private Sub SessaoRecuperaMesAno()
        Mes = Session("Mes" & HID.Value)
        Ano = Session("Ano" & HID.Value)
    End Sub

    Private Sub SessaoRecuperaDar()
        objDarDifAliquota = Session("objDarDifAliquota" & HID.Value)
    End Sub

    Private Sub SessaoRecuperaListaDarDif()
        ListaDarDif = Session("ListaDarDifAliquota" & HID.Value)
    End Sub

    Private Sub SessaoRecuperaEmpresa()
        ObjEmpresa = Session("objEmpresa" & HID.Value)
    End Sub

    Private Sub RemoveSessoes()
        Session.Remove("objEmpresa" & HID.Value)
        Session.Remove("ListaDarDifAliquota" & HID.Value)
        Session.Remove("ObjDarDifAliquota" & HID.Value)
        Session.Remove("Mes" & HID.Value)
        Session.Remove("Ano" & HID.Value)

    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteDarDifAliquota" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteDarDifAliquota" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value.Replace("-", ";")
            Session.Remove("objClienteDarDifAliquota" & HID.Value)
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub BindGridView(ByVal parameters As Dictionary(Of String, Object))
        SessaoRecuperaListaDarDif()

        If ListaDarDif Is Nothing Then
            ListaDarDif = New ListDarDiferencialDeAliquota(True, parameters)
        End If

        If parameters.ContainsKey("Ano") AndAlso Not String.IsNullOrEmpty(parameters("Ano")) Then
            Ano = parameters("Ano")
        End If

        If parameters.ContainsKey("Mes") AndAlso Not String.IsNullOrEmpty(parameters("Ano")) Then
            Mes = parameters("Mes")
        End If

        If parameters.ContainsKey("Empresa") AndAlso Not String.IsNullOrEmpty(parameters("Empresa")) Then
            Dim s As String = parameters("Empresa")
            ObjEmpresa = New [Lib].Negocio.Cliente(s.Split("-")(0), s.Split("-")(1))
            lblEmpresa.Text = Funcoes.FormatarListItemCliente(ObjEmpresa).Text
            CarregaSpedInformacoesAdicionais(ObjEmpresa.CodigoEstado)
            SessaoSalvaEmpresa()
        End If

        SessaoSalvaMesAno()
        SessaoSalvaListaDarDif()
        SessaoRecuperaListaDarDif()

        GrdDarDiferencialDeAliquota.DataSource = ListaDarDif.ListaDeDarDiferencialDeAliquota("")
        GrdDarDiferencialDeAliquota.DataBind()
    End Sub

    Private Sub CarregaSpedInformacoesAdicionais(ByVal Estado As String)
        Dim Sql As String = ""

        Sql = " SELECT Codigo_Id, Descricao, DataInicial, DataFinal" & vbCrLf & _
              "   FROM SpedInfAdicionaisDeApuracao" & vbCrLf & _
              "  WHERE LEFT(Codigo_Id, 2) = '" & Estado & "'"

        ddlInformacoesAdicionais.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "SpedInfAdicionais").Tables(0).Rows
            ddlInformacoesAdicionais.Items.Add(New ListItem(Dr("Codigo_Id") & "-" & Dr("Descricao"), Dr("Codigo_Id")))
        Next

        ddlInformacoesAdicionais.Items.Insert(0, "")

    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        RemoveSessoes()
        Popup.CloseDialog(Me.Page, "divDarDiferencialDeAliquota")
    End Sub

    Private Function VerificaPeriodo(ByVal pAno As Integer, pMes As Integer) As Boolean
        If pAno <> Ano Or pMes <> Mes Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Function ValidaCampos() As Boolean
        SessaoRecuperaMesAno()
        If String.IsNullOrWhiteSpace(txtDar.Text) Then
            MsgBox(Me.Page, "Digite o Número do DAR!")
            txtDar.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtData.Text) Then
            MsgBox(Me.Page, "Indique a Data de Emissao!")
            txtData.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataRef.Text) Then
            MsgBox(Me.Page, "Indique a Data de Referencia!")
            txtDataRef.Focus()
            Return False
        ElseIf CDate(txtDataRef.Text).Year <> Ano Then
            MsgBox(Me.Page, "Ano de Referencia escolhido deve ser igual ao ano do Processo!")
            Return False
        ElseIf CDate(txtDataRef.Text).Month <> Mes Then
            MsgBox(Me.Page, "Mês de Referencia escolhido deve ser igual ao mês do Processo!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            MsgBox(Me.Page, "Indique o Cliente!")
            btnCliente.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodigoReceita.Text) Then
            MsgBox(Me.Page, "Digite o Código da Receita!")
            txtCodigoReceita.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtNota.Text) Then
            MsgBox(Me.Page, "Digite a Nota Fiscal!")
            txtNota.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtSerie.Text) Then
            MsgBox(Me.Page, "Digite a Série da Nota Fiscal!")
            txtSerie.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtValor.Text) Then
            MsgBox(Me.Page, "Digite o Valor do DAR!")
            txtValor.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If ValidaCampos() Then
                SessaoRecuperaListaDarDif()
                SessaoRecuperaEmpresa()

                objDarDifAliquota = New DarDiferencialDeAliquota()
                objDarDifAliquota.IUD = "I"
                objDarDifAliquota.Dar = txtDar.Text
                objDarDifAliquota.CodigoEmpresa = ObjEmpresa.Codigo
                objDarDifAliquota.EndEmpresa = ObjEmpresa.CodigoEndereco
                objDarDifAliquota.Data = txtData.Text
                objDarDifAliquota.DataReferencia = txtDataRef.Text
                objDarDifAliquota.CodigoCliente = txtCodigoCliente.Value.Split(";")(0)
                objDarDifAliquota.EndCliente = txtCodigoCliente.Value.Split(";")(1)
                objDarDifAliquota.CodigoReceita = txtCodigoReceita.Text
                objDarDifAliquota.Nota = txtNota.Text
                objDarDifAliquota.Serie = txtSerie.Text
                objDarDifAliquota.Valor = txtValor.Text
                objDarDifAliquota.CodigoSpedInfAdicionaisDeApuracao = ddlInformacoesAdicionais.SelectedValue

                If objDarDifAliquota.Salvar() Then
                    MsgBox(Me.Page, "Registro Gravado com Sucesso.", eTitulo.Sucess)
                    ListaDarDif.Add(objDarDifAliquota)
                    Limpar()
                    Me.BindGridView(New Dictionary(Of String, Object))
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            SessaoRecuperaDar()
            SessaoRecuperaListaDarDif()
            SessaoRecuperaEmpresa()

            If ValidaCampos() Then
                objDarDifAliquota.IUD = "U"
                objDarDifAliquota.Dar = txtDar.Text
                objDarDifAliquota.CodigoEmpresa = ObjEmpresa.Codigo
                objDarDifAliquota.EndEmpresa = ObjEmpresa.CodigoEndereco
                objDarDifAliquota.Data = txtData.Text
                objDarDifAliquota.DataReferencia = txtDataRef.Text
                objDarDifAliquota.CodigoCliente = txtCodigoCliente.Value.Split(";")(0)
                objDarDifAliquota.EndCliente = txtCodigoCliente.Value.Split(";")(1)
                objDarDifAliquota.Nota = txtNota.Text
                objDarDifAliquota.Serie = txtSerie.Text
                objDarDifAliquota.CodigoReceita = txtCodigoReceita.Text
                objDarDifAliquota.Valor = txtValor.Text
                objDarDifAliquota.CodigoSpedInfAdicionaisDeApuracao = ddlInformacoesAdicionais.SelectedValue
                If objDarDifAliquota.Salvar() Then
                    MsgBox(Me.Page, "Registro Alterado com Sucesso.", eTitulo.Sucess)
                    Limpar()
                    Me.BindGridView(New Dictionary(Of String, Object))
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            SessaoRecuperaDar()
            SessaoRecuperaListaDarDif()

            If Not objDarDifAliquota Is Nothing Then
                objDarDifAliquota.IUD = "D"
                If objDarDifAliquota.Salvar() Then
                    MsgBox(Me.Page, "Registro Excluído com Sucesso.", eTitulo.Sucess)
                    ListaDarDif.Remove(objDarDifAliquota)
                    Limpar()
                    SessaoSalvaListaDarDif()
                    Me.BindGridView(New Dictionary(Of String, Object))
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        SessaoRecuperaListaDarDif()
        'Dim ds As DataSet

        Dim crpt As New ReportDocument()

        Try
            crpt.FileName = Server.MapPath("~/Reports/Cr_DarDiferencialDeAliquota.rpt")
            crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
            Dim arquivo As String = NomeArquivo

            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))

            crpt.SetDataSource(ListaDarDif.ListaDeDarDiferencialDeAliquota(strCaminhoImagem))

            Dim parameters = New Dictionary(Of String, Object)()
            Dim objEmpresa As Cliente = New Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
            parameters.Add("Empresa", objEmpresa.Nome.Trim())
            parameters.Add("Cidade", objEmpresa.Cidade.Trim() & "/" & objEmpresa.Estado.Codigo.Trim())
            parameters.Add("Cnpj", objEmpresa.Codigo.Trim())

            Funcoes.BindParameters(crpt, parameters)

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Protected Sub btnCliente_Click(sender As Object, e As EventArgs) Handles btnCliente.Click
        Session("ssCampo") = "LivreClasse"
        Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            Me.MainUserControl = Me
            Popup.ConsultaDeClientes(Me.Page, "objClienteDarDifAliquota" & HID.Value, "txtCliente")
        End If
    End Sub

    Public Overrides Sub Limpar()
        txtDar.Text = String.Empty
        txtData.Text = String.Empty
        txtDataRef.Text = String.Empty
        txtClientes.Text = String.Empty
        txtCodigoCliente.Value = ""
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtCodigoReceita.Text = String.Empty
        txtValor.Text = String.Empty
        VisibleGravar()
        ddlInformacoesAdicionais.SelectedIndex = 0
    End Sub

    Protected Sub GrdDarDiferencialDeAliquota_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GrdDarDiferencialDeAliquota.SelectedIndexChanged
        SessaoRecuperaListaDarDif()
        objDarDifAliquota = ListaDarDif.Find(Function(S) S.Dar = GrdDarDiferencialDeAliquota.SelectedRow.Cells(1).Text)
        txtDar.Text = objDarDifAliquota.Dar
        txtData.Text = objDarDifAliquota.Data
        txtDataRef.Text = objDarDifAliquota.DataReferencia
        txtCodigoCliente.Value = objDarDifAliquota.CodigoCliente & ";" & objDarDifAliquota.EndCliente
        Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objDarDifAliquota.Cliente)
        txtClientes.Text = itemCliente.Text
        txtCodigoReceita.Text = objDarDifAliquota.CodigoReceita
        txtNota.Text = objDarDifAliquota.Nota
        txtSerie.Text = objDarDifAliquota.Serie
        txtValor.Text = Str(objDarDifAliquota.Valor)
        ddlInformacoesAdicionais.SelectedValue = objDarDifAliquota.CodigoSpedInfAdicionaisDeApuracao

        VisibleAtualizar()

        SessaoSalvaListaDarDif()
        SessaoSalvaDar()
    End Sub

    Private Sub VisibleGravar()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Sub VisibleAtualizar()
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True
    End Sub

End Class