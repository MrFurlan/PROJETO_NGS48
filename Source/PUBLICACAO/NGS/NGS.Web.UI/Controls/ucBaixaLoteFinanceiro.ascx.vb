Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucBaixaLoteFinanceiro
    Inherits BaseUserControl
    Dim ClienteLote As String
    Dim campoLote() As String
    Dim Descricao As String
    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            ddl.Carregar(DdlEmpresaPagadoraLote, CarregarDDL.Tabela.ClientesXEmpresas, "", True)

            Dim Parametros As New Hashtable
            Parametros.Clear()
            Parametros.Add("listarTudo", "N")

            ddl.Carregar(DdlTiposDePagamentosLote, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)

            CarteiraDoTitulo()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Private Sub CarteiraDoTitulo()
        Dim objCarteiraDoTitulo As New [Lib].Negocio.ListCarteiraDoTitulo()

        ddlCarteiraDoTituloLote.DataValueField = "Codigo"
        ddlCarteiraDoTituloLote.DataTextField = "Descricao"
        ddlCarteiraDoTituloLote.DataSource = objCarteiraDoTitulo.ToArray()
        ddlCarteiraDoTituloLote.DataBind()
        ddlCarteiraDoTituloLote.SelectedIndex = 0
    End Sub

    Protected Sub DdlEmpresaPagadoraLote_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        DdlTiposDePagamentosLote.SelectedIndex = 0
        CargaBancoPagadorLote()
    End Sub

    Private Sub CargaBancoPagadorLote()
        DdlBancoPagadorLote.Items.Clear()
        DdlContaPagadoraLote.Items.Clear()

        If DdlEmpresaPagadoraLote.Text <> "" Then

            ClienteLote = DdlEmpresaPagadoraLote.SelectedValue
            campoLote = ClienteLote.Split("-")

            Sql = " SELECT DISTINCT BxC.Banco_Id,  B.Descricao, BxC.ContaContabil, p.Titulo" & vbCrLf & _
                  "   FROM BancosXContas BxC " & vbCrLf & _
                  "  INNER JOIN Bancos B  " & vbCrLf & _
                  "          ON BxC.Banco_Id = B.Banco_Id" & vbCrLf & _
                  "  LEFT JOIN PlanoDeContas p" & vbCrLf & _
                  "         on p.Conta_Id = BxC.ContaContabil" & vbCrLf & _
                  "  WHERE BxC.Empresa_Id     ='" & campoLote(0) & "'" & vbCrLf & _
                  "    AND BxC.EndEmpresa_Id  = " & campoLote(1)

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Bancos").Tables(0).Rows
                Descricao = Format(Dr("Banco_Id"), "0000") & " - " & Dr("Descricao") & " - (" & Dr("ContaContabil") & "-" & Dr("Titulo") & ")"
                DdlBancoPagadorLote.Items.Add(New ListItem(Descricao, Dr("Banco_Id")))
            Next

            DdlBancoPagadorLote.Items.Insert(0, "")
            DdlBancoPagadorLote.SelectedIndex = 0
        End If
    End Sub


    Protected Sub DdlBancoPagadorLote_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BancoPagadorLote()
    End Sub

    Private Sub BancoPagadorLote()
        Dim Cliente As String
        Dim Campo() As String
        Dim Conta As String

        Cliente = DdlEmpresaPagadoraLote.SelectedValue
        Campo = Cliente.Split("-")
        DdlContaPagadoraLote.Items.Clear()

        If Campo(0) <> "" Then
            Sql = " SELECT bxc.Agencia_Id, bxc.DigitoAgencia_Id, bxc.Conta_Id, bxc.DigitoConta_Id, isnull(bxc.TipoConta,'C') as TipoConta, bxc.ContaContabil, bxc.Observacoes, Pc.Titulo" & vbCrLf & _
                  "   FROM BancosXContas bxc" & vbCrLf & _
                  "  Inner Join PlanoDeContas PC" & vbCrLf & _
                  "     on Bxc.ContaContabil = Pc.Conta_Id" & vbCrLf & _
                  "  WHERE bxc.Empresa_Id    = '" & Campo(0) & "'" & vbCrLf & _
                  "    AND bxc.EndEmpresa_Id = " & Campo(1) & vbCrLf & _
                  "    AND bxc.Banco_Id      =  " & DdlBancoPagadorLote.SelectedValue & vbCrLf & _
                  "    AND (isnull(pc.Adiantamento,0) = 0 or (pc.Adiantamento = 1 and left(pc.Conta_Id,1) = '1'))" & vbCrLf & _
                  "    AND bxc.Ativo = 1 " & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
                Conta = Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "-" & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & "-" & Dr("TipoConta") & "-" & Dr("ContaContabil")
                Descricao = "AG " & Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & " " & Dr("TipoConta") & " " & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & " - " & Dr("Observacoes") & " - (" & Dr("ContaContabil") & "-" & Dr("Titulo") & ")"
                DdlContaPagadoraLote.Items.Add(New ListItem(Descricao, Conta))
            Next

            DdlContaPagadoraLote.Items.Insert(0, "")
            DdlContaPagadoraLote.SelectedIndex = 0
        End If
    End Sub

    Public Overrides Sub Limpar()
        DdlEmpresaPagadoraLote.SelectedIndex = 0
        DdlTiposDePagamentosLote.SelectedIndex = 0
        DdlBancoPagadorLote.Items.Clear()
        DdlContaPagadoraLote.Items.Clear()
        ddlCarteiraDoTituloLote.SelectedIndex = 0
        txtMovimentoLote.Text = CDate(Today).ToString("dd/MM/yyyy")
        txtProrrogacaoLote.Text = CDate(Today).ToString("dd/MM/yyyy")
    End Sub

    Protected Sub lnkFecharBaixaLoteFinanceiro_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFecharBaixaLoteFinanceiro.Click
        Try
            Popup.CloseDialog(Me.Page, "divBaixaLoteFinanceiro")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBaixaLoteFinanceiro_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkBaixaLoteFinanceiro.Click
        Try
            If String.IsNullOrEmpty(DdlEmpresaPagadoraLote.SelectedValue) Then
                MsgBox(Me.Page, "Empresa Pagadora é Obrigatória.")
            ElseIf String.IsNullOrEmpty(DdlTiposDePagamentosLote.SelectedValue) Then
                MsgBox(Me.Page, "O Tipo de Pagamento é Obrigatório.")
            ElseIf String.IsNullOrEmpty(DdlBancoPagadorLote.SelectedValue) Then
                MsgBox(Me.Page, "O Banco Pagador é Obrigatório.")
            ElseIf String.IsNullOrEmpty(DdlContaPagadoraLote.SelectedValue) Then
                MsgBox(Me.Page, "A Conta Pagadora é Obrigatória.")
            ElseIf String.IsNullOrEmpty(ddlCarteiraDoTituloLote.SelectedValue) Then
                MsgBox(Me.Page, "Carteira do Título é Obrigatória.")
            ElseIf String.IsNullOrWhiteSpace(txtMovimentoLote.Text) Then
                MsgBox(Me.Page, "Data do movimento para contabilização é obrigatório.")
            ElseIf String.IsNullOrWhiteSpace(txtProrrogacaoLote.Text) Then
                MsgBox(Me.Page, "Data de Vencimento é obrigatório.")
            ElseIf Not IsDate(txtMovimentoLote.Text) Then
                MsgBox(Me.Page, "Data do movimento para contabilização é obrigatório.")
            ElseIf Not IsDate(txtProrrogacaoLote.Text) Then
                MsgBox(Me.Page, "Data de Vencimento é obrigatório.")
            ElseIf CDate(txtProrrogacaoLote.Text) > CDate(txtMovimentoLote.Text) Then
                MsgBox(Me.Page, "Data de Vencimento não pode ser maior que a data do Movimento.")
            Else
                Dim ObjTitulo As New Titulo()

                ObjTitulo.Codigo = 999999999
                ObjTitulo.UsuarioLiberacao = "BaixaLoteFinanceiro"

                Dim dadosEmpresaPagadora() As String = DdlEmpresaPagadoraLote.SelectedValue.Split("-")
                ObjTitulo.CodigoEmpresaPagadora = dadosEmpresaPagadora(0)
                ObjTitulo.EndEmpresaPagadora = dadosEmpresaPagadora(1)

                ObjTitulo.CodigoTipoPgto = DdlTiposDePagamentosLote.SelectedValue

                ObjTitulo.CodigoBancoPagador = DdlBancoPagadorLote.SelectedValue

                Dim dadosBanc As String() = DdlContaPagadoraLote.SelectedValue.Split("-")
                ObjTitulo.CodigoAgenciaPagadora = dadosBanc(0)
                ObjTitulo.DigitoAgenciaPagadora = dadosBanc(1)
                ObjTitulo.ContaPagadora = dadosBanc(2)
                ObjTitulo.DigitoContaPagadora = dadosBanc(3)
                ObjTitulo.TipoContaPagadora = dadosBanc(4)
                ObjTitulo.ContaContabilPagadora = dadosBanc(5)

                ObjTitulo.CodigoCarteiraDoTitulo = ddlCarteiraDoTituloLote.SelectedValue

                ObjTitulo.Prorrogacao = CDate(txtProrrogacaoLote.Text)
                ObjTitulo.Baixa = CDate(txtMovimentoLote.Text)

                Session(Session("ssTipoRetorno")) = ObjTitulo

                If Session("ssTipoRetorno") IsNot Nothing Then
                    If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                        Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                        CType(uc, IBaseUserControl).Carregar(ObjTitulo)
                    Else
                        CType(Me.Page, IBasePage).Carregar(ObjTitulo)
                    End If

                    Popup.CloseDialog(Me.Page, "divBaixaLoteFinanceiro")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class