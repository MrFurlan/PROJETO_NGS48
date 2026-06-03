Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ReciboAvulso
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Dim Valor As Decimal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ReciboAvulso", "ACESSAR") Then
                    txtData.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    BuncarUnidadeDeNegocio()
                    AtribuirDados()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "Financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtribuirDados()
        If Not Session("objClienteRECAVU") Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteRECAVU"), [Lib].Negocio.Cliente))
            Dim strIdCliente() As String = itemCliente.Value.Split("-")
            TxtCpfCnpj.Text = strIdCliente(0)
        End If
    End Sub

    Public Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidarCampos() As Boolean
        If txtData.Text.Length = 0 Or IsDate(txtData.Text) = False Then
            MsgBox(Me.Page, "Informe uma data de vencimento válida.")
            txtData.Focus()
            Return False
        ElseIf ddlEmpresa.Text = "" Then
            MsgBox(Me.Page, "Informe a empresa.")
            ddlEmpresa.Focus()
            Return False
        ElseIf txtClientes.Text = "" Then
            MsgBox(Me.Page, "Informe o cliente.")
            txtClientes.Focus()
            Return False
        ElseIf txtTotalRecibo.Text = "" Then
            MsgBox(Me.Page, "Informe total Da nota.")
            txtTotalRecibo.Focus()
            Return False
        ElseIf TxtCpfCnpj.Text = "" Then
            MsgBox(Me.Page, "Informe um CPF ou CNPJ.")
            TxtCpfCnpj.Focus()
            Return False
        ElseIf TxtReferente.Text = "" Then
            MsgBox(Me.Page, "Informe uma referencia a este recibo.")
            TxtReferente.Focus()
            Return False
        End If

        Return True
    End Function

    Private Sub Limpar()
        txtClientes.Text = ""
        txtData.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtTotalRecibo.Text = ""
        txtCodigoCliente.Value = ""
        TxtCpfCnpj.Text = ""
        TxtEndereco.Text = ""
        TxtEstado.Text = ""
        TxtCidade.Text = ""
        TxtBairro.Text = ""
        TxtCEP.Text = ""
        TxtTelefone.Text = ""
        TxtComplemento.Text = ""
        TxtNumero.Text = ""
        txtNumeroRecibo.Text = ""
        TxtReferente.Text = ""
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteRECAVU" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteRECAVU" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteRECAVU" & HID.Value)
        End If
    End Sub

    Protected Sub ddlBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlBuscaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteRECAVU" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgCep_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If TxtCEP.Text.Length = 0 OrElse TxtCEP.Text.Length < 8 Then
                MsgBox(Me.Page, "CEP deve ser informado com 8 dígitos.")
            Else
                If Funcoes.VerificaConexaoInternet Then
                    Dim ds As New DataSet
                    ds = Funcoes.BuscaCep(TxtCEP.Text)

                    If Not IsNothing(ds) AndAlso (ds.Tables(0).Rows.Count > 0) Then
                        Select Case ds.Tables(0).Rows(0).Item("resultado")
                            Case "1"
                                TxtEstado.Text = ds.Tables(0).Rows(0).Item("uf").ToString().Trim().ToUpper()
                                TxtCidade.Text = ds.Tables(0).Rows(0).Item("cidade").ToString().Trim().ToUpper()
                                TxtBairro.Text = ds.Tables(0).Rows(0).Item("bairro").ToString().Trim().ToUpper()
                                If ds.Tables(0).Rows(0).Item("tipo_logradouro").ToString.Length > 0 Then
                                    TxtEndereco.Text = ds.Tables(0).Rows(0).Item("tipo_logradouro").ToString().Trim().ToUpper() & " " & ds.Tables(0).Rows(0).Item("logradouro").ToString().Trim().ToUpper()
                                Else
                                    TxtEndereco.Text = ds.Tables(0).Rows(0).Item("logradouro").ToString().Trim().ToUpper()
                                End If

                            Case "2"
                                TxtEstado.Text = ds.Tables(0).Rows(0).Item("uf").ToString().Trim().ToUpper()
                                TxtCidade.Text = ds.Tables(0).Rows(0).Item("cidade").ToString().Trim().ToUpper()

                            Case 0
                                MsgBox(Me.Page, ds.Tables(0).Rows(0).Item("resultado_txt").ToString())
                        End Select
                    Else
                        MsgBox(Me.Page, "Ocorreu um erro inesperado.")
                    End If
                Else
                    MsgBox(Me.Page, "No momento não é possível executar esse processo. Verifique sua conexão com a internet.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkEmitir.Click
        Try
            If Funcoes.VerificaPermissao("ReciboAvulso", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim xextenso As String = ""
                    Dim yextenso As String = ""
                    Dim dsEmitir As New DataSet
                    Dim row As DataRow
                    Dim j As Integer
                    Dim mes As Integer
                    Dim MESX As String = ""
                    Dim Historico As String = ""
                    Dim SqlArray As New ArrayList
                    Dim Empresa As String = ""
                    Dim EndEmpresa As String = ""
                    Dim ENome As String = ""
                    Dim EEndereco As String = ""
                    Dim ECep As String = ""
                    Dim ECidade As String = ""
                    Dim EEstado As String = ""
                    Dim ECnpj As String = ""
                    Dim EInscricao As String = ""
                    Dim Efone As String = ""
                    Dim EBairro As String = ""
                    Dim EComplemento As String = ""
                    Dim ENumero As Integer
                    Dim Cliente As String = ""
                    Dim EndCliente As String = ""
                    Dim CNome As String = ""
                    Dim CEndereco As String = ""
                    Dim CCep As String = ""
                    Dim CCidade As String = ""
                    Dim CEstado As String = ""
                    Dim CCnpj As String = ""
                    Dim CInscricao As String = ""
                    Dim Cfone As String = ""
                    Dim CBairro As String = ""
                    Dim CComplemento As String = ""
                    Dim CNumero As Integer
                    Dim CBancoCliente As String = ""
                    Dim CAgenciaCliente As String = ""
                    Dim CDigitoAgenciaCliente As String = ""
                    Dim CCcontaCliente As String = ""
                    Dim CDigitoContaCliente As String = ""
                    Dim dsRecibo As New DataSet
                    Dim dtRecibo As DataTable
                    Dim TnumeroDoCheque As Integer

                    dtRecibo = New DataTable("ReciboAPagar")

                    dtRecibo.Columns.Add("ENome", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("EEndereco", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("ECep", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("ECidade", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("EEstado", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("ECnpj", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("EInscricao", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("EFone", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("EBairro", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("EComplemento", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("ENumero", Type.GetType("System.Int32")).DefaultValue = 0
                    ' ''campos do cliente
                    dtRecibo.Columns.Add("CNome", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CEndereco", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CCep", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CCidade", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CEstado", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CCnpj", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CInscricao", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CFone", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CBairro", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CComplemento", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("CNumero", Type.GetType("System.Int32")).DefaultValue = 0
                    '' campos to titulo 
                    dtRecibo.Columns.Add("TNumtit", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TValor", Type.GetType("System.Decimal"))
                    dtRecibo.Columns.Add("TExtenso", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("THistorico", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TDia", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TMes", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TAno", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TFormaPagto", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TBanco", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TAgencia", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TConta", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TVencimento", Type.GetType("System.DateTime"))
                    dtRecibo.Columns.Add("TBaixa", Type.GetType("System.DateTime"))
                    dtRecibo.Columns.Add("TRecibo", Type.GetType("System.Int32")).DefaultValue = 0
                    dtRecibo.Columns.Add("TRegistro", Type.GetType("System.Int32")).DefaultValue = 0
                    dtRecibo.Columns.Add("TNumeroDoCheque", Type.GetType("System.Int32")).DefaultValue = 0
                    '' Campos novos do cheque 05/05/2010
                    dtRecibo.Columns.Add("TValorDoDocumento", Type.GetType("System.Decimal")).DefaultValue = 0
                    dtRecibo.Columns.Add("TDescontos", Type.GetType("System.Decimal")).DefaultValue = 0
                    dtRecibo.Columns.Add("TDeducoes", Type.GetType("System.Decimal")).DefaultValue = 0
                    dtRecibo.Columns.Add("TJuros", Type.GetType("System.Decimal")).DefaultValue = 0
                    dtRecibo.Columns.Add("TAcrescimos", Type.GetType("System.Decimal")).DefaultValue = 0
                    dtRecibo.Columns.Add("TDigito", Type.GetType("System.String"))
                    dtRecibo.Columns.Add("TDigitoAgencia", Type.GetType("System.String"))

                    '' Campos novos do cheque 05/05/2010
                    dsRecibo.Tables.Add(dtRecibo)
                    Dim ValorCobrado As Decimal
                    Dim ValorDoDocumento As Decimal
                    Dim Juros As Decimal
                    Dim Acrescimos As Decimal
                    Dim Descontos As Decimal
                    Dim deducoes As Decimal
                    Dim TipoPagto As Integer
                    Dim FormaDePagamento As String
                    '' Campos novos do cheque 05/05/2010
                    Dim TvalorDoDocumento As Decimal
                    Dim Tdescontos As Decimal
                    Dim Tdeducoes As Decimal
                    Dim TJuros As Decimal
                    Dim TAcrescimos As Decimal
                    Dim Tdigito As String
                    '' Campos novos do cheque 05/05/2010
                    ''* Campos do Data Set + titulos fim.
                    Dim emitirRecibo As Boolean = False

                    emitirRecibo = True
                    Dim DataBaixa As String
                    Dim dataVencimento As String

                    DataBaixa = txtData.Text
                    dataVencimento = txtData.Text

                    Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
                    Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
                    'Dim dr As DataRow

                    Empresa = strEmpresa(0)
                    EndEmpresa = strEmpresa(1)


                    Cliente = txtClientes.Text
                    EndCliente = "0"
                    ValorCobrado = 0
                    ValorDoDocumento = txtTotalRecibo.Text
                    Juros = 0.0
                    Acrescimos = 0.0
                    Descontos = 0.0
                    deducoes = 0.0
                    Historico = TxtReferente.Text

                    CBancoCliente = ""
                    CAgenciaCliente = ""
                    CDigitoAgenciaCliente = ""
                    CCcontaCliente = ""
                    CDigitoContaCliente = ""
                    ValorCobrado = ValorDoDocumento + Juros + Acrescimos - Descontos - deducoes
                    TipoPagto = 1
                    FormaDePagamento = ""
                    TnumeroDoCheque = 0
                    Tdescontos = 0.0
                    Tdeducoes = 0.0
                    TJuros = 0.0
                    TAcrescimos = 0.0
                    TvalorDoDocumento = txtTotalRecibo.Text
                    Tdigito = ""

                    Dim drEmp As DataRow
                    '' Dados da empresa - fim 
                    '' Consultado empresa 
                    Sql = "  SELECT Clientes.Cliente_Id ," & vbCrLf & _
                          " Clientes.Nome, Clientes.Cidade," & vbCrLf & _
                          " Clientes.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                          " Clientes.Endereco , Clientes.Cep," & vbCrLf & _
                          " Clientes.Inscricao, Clientes.Telefone," & vbCrLf & _
                          " Clientes.Bairro, Clientes.Complemento," & vbCrLf & _
                          " Clientes.Numero " & vbCrLf & _
                          " FROM Clientes " & vbCrLf & _
                          " WHERE Clientes.Cliente_Id = '" & Empresa & "'" & vbCrLf & _
                          " AND Clientes.Endereco_id = '" & EndEmpresa & "'" & vbCrLf
                    ds = Banco.ConsultaDataSet(Sql, "Clientes")

                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each drEmp In ds.Tables(0).Rows
                            ENome = drEmp("Nome")
                            EEndereco = drEmp("Endereco")
                            ECep = drEmp("Cep")
                            ECidade = drEmp("Cidade")
                            EEstado = drEmp("Estado")
                            ECnpj = drEmp("Cliente_id")
                            EInscricao = drEmp("Inscricao")
                            Efone = drEmp("Telefone")
                            EBairro = drEmp("Bairro")
                            EComplemento = drEmp("Complemento")
                            ENumero = drEmp("Numero")
                            Exit For
                        Next
                    End If


                    CNome = Cliente
                    CEndereco = TxtEndereco.Text
                    CCep = TxtCEP.Text
                    CCidade = TxtCidade.Text
                    CEstado = TxtEstado.Text
                    CCnpj = TxtCpfCnpj.Text
                    CInscricao = ""
                    Cfone = TxtTelefone.Text
                    CBairro = TxtBairro.Text
                    CComplemento = TxtComplemento.Text
                    CNumero = IIf(TxtNumero.Text.Trim = "", 0, TxtNumero.Text)

                    ' Cria Data Sete que vai ser utilizado no relatorio
                    ' Move campos para o Data Set.
                    ' Move campos da Empresa
                    row = dtRecibo.NewRow()
                    row("ENome") = ENome
                    row("EEndereco") = EEndereco
                    row("ECep") = ECep
                    row("ECidade") = ECidade
                    row("EEstado") = EEstado
                    row("ECnpj") = Funcoes.FormatarCpfCnpj(ECnpj)
                    row("EInscricao") = EInscricao
                    row("EFone") = Efone
                    row("EBairro") = EBairro
                    row("EComplemento") = EComplemento
                    row("ENumero") = ENumero
                    '' Move campos do Cliente / fornecedor
                    row("CNome") = CNome
                    row("CEndereco") = CEndereco
                    row("CCep") = CCep
                    row("CCidade") = CCidade
                    row("CEstado") = CEstado
                    row("CCnpj") = Funcoes.FormatarCpfCnpj(CCnpj)
                    row("CInscricao") = CInscricao
                    row("CFone") = Cfone
                    row("CBairro") = CBairro
                    row("CComplemento") = CComplemento
                    row("CNumero") = CNumero
                    '' Move campos do Titulo
                    row("Tnumtit") = "Avulso"
                    '' feito
                    row("TValor") = ValorCobrado
                    row("TNumeroDoCheque") = TnumeroDoCheque
                    '' campos novos 05/05/10
                    row("TValorDoDocumento") = TvalorDoDocumento
                    row("TDescontos") = Tdescontos
                    row("TDeducoes") = Tdeducoes
                    row("TJuros") = TJuros
                    row("TAcrescimos") = TAcrescimos
                    row("TDigito") = Tdigito
                    row("TDigitoAgencia") = CDigitoAgenciaCliente

                    '' campos novos 05/05/10
                    Dim valcobradostr As String
                    valcobradostr = CStr(ValorCobrado)
                    txtTotalRecibo.Text = ValorCobrado

                    Valor = Replace(txtTotalRecibo.Text, ".", "")

                    ''* Rotina de extenso inicio
                    yextenso = "("
                    yextenso &= UCase(Funcoes.Extenso(txtTotalRecibo.Text(), "Real", "Reais"))
                    yextenso &= " *"
                    xextenso = yextenso
                    For j = 1 To (120 - Len(xextenso))
                        xextenso &= " *"
                    Next
                    xextenso &= ")"
                    row("TExtenso") = xextenso
                    ''* Rotina de extenso fim 
                    row("THistorico") = Historico
                    row("TDia") = Day(DataBaixa)
                    row("TAno") = Year(DataBaixa)

                    row("TMes") = Month(DataBaixa)
                    ''* Rotina do Mes inicio
                    mes = Month(DataBaixa)
                    If mes = 1 Then
                        MESX = "JANEIRO"
                    End If
                    If mes = 2 Then
                        MESX = "FEVEREIRO"
                    End If
                    If mes = 3 Then
                        MESX = "MARCO"
                    End If
                    If mes = 4 Then
                        MESX = "ABRIL"
                    End If

                    If mes = 5 Then
                        MESX = "MAIO"
                    End If
                    If mes = 6 Then
                        MESX = "JUNHO"
                    End If
                    If mes = 7 Then
                        MESX = "JULHO"
                    End If
                    If mes = 8 Then
                        MESX = "AGOSTO"
                    End If
                    If mes = 9 Then
                        MESX = "SETEMBRO"
                    End If
                    If mes = 10 Then
                        MESX = "OUTUBRO"
                    End If
                    If mes = 11 Then
                        MESX = "NOVEMBRO"
                    End If
                    If mes = 12 Then
                        MESX = "DEZEMBRO"
                    End If

                    row("TMes") = MESX
                    row("TFormaPagto") = FormaDePagamento
                    row("TBanco") = CBancoCliente
                    row("TAgencia") = CAgenciaCliente
                    row("TConta") = CCcontaCliente
                    row("TVencimento") = dataVencimento
                    row("TBaixa") = DataBaixa

                    If String.IsNullOrWhiteSpace(txtNumeroRecibo.Text) Then
                        row("Trecibo") = 0
                    Else
                        row("Trecibo") = txtNumeroRecibo.Text
                    End If

                    row("TRegistro") = 0

                    dtRecibo.Rows.Add(row)

                    If emitirRecibo Then

                        Dim param As New Dictionary(Of String, Object)
                        param.Add("XNome", ENome)

                        Funcoes.BindReport(Me.Page, dsRecibo, "Cr_ReciboPagar", eExportType.PDF, param)
                    End If
                Else
                    MsgBox(Me.Page, "Titulo não foi selecionado para emissão.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ReciboAvulso")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class