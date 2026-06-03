Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class LancamentoEstoqueViaAnalisePeloLaudo
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.ProducaoEstoque)
            If Not IsPostBack And IsConnect Then
                'If Funcoes.VerificaPermissao("LANCAMENTOESTOQUEVIAANALISEPELOLAUDO", "ACESSAR") Then
                limpar()
                'Else
                '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub limpar()

        Session.Remove("objConsultarEstoqueXLaudo")
        HID.Value = Guid.NewGuid().ToString
        lnkProcessar.Parent.Visible = True
        txtAnalisesInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtAnalisesFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")

    End Sub

    Protected Sub lnkProcessar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkProcessar.Click
        Try
            'If Funcoes.VerificaPermissao("LANCAMENTOESTOQUEVIAANALISEPELOLAUDO", "GRAVAR") Then


            'Dim ObjNavioxinvoice As New NavioXInvoice()
            'ObjNavioxinvoice.DataDeChegada = txtDataDeChegada.Text
            'ObjNavioxinvoice.Observacao = txtObservacao.Text
            'ObjNavioxinvoice.Navio_Id = txtNavio.Text
            'ObjNavioxinvoice.Descricao = RTrim(txtDescricao.Text)
            'ObjNavioxinvoice.Ativo = chkAtivo.Checked
            'ObjNavioxinvoice.IUD = "I"

            'If ObjNavioxinvoice.Salvar Then
            '    MsgBox(Me.Page, "Processo realizado com sucesso.", eTitulo.Sucess)
            '    limpar()
            'End If


            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para gravar registro.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sernder As Object, ByVal e As System.EventArgs) Handles lnkLimpar.Click
        limpar()
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LancamentoEstoqueViaAnalisePeloLaudo")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class