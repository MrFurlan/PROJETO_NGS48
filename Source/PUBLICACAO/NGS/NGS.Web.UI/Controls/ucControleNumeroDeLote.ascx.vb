Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucControleNumeroDeLote
    Inherits BaseUserControl

    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Private Property index As Integer
        Get
            Return CInt(Session("index" & HID.Value))
        End Get
        Set(value As Integer)
            Session("index" & HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub SetarLotes()
        RecuperaNotaFiscal()

        For i = 0 To objNotaFiscal.Itens.Count - 1
            Session("ObjItemLotes" & i & HID.Value) = objNotaFiscal.Itens(i).Lotes
        Next
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub sessionLotes()
        Session("ObjItemLotes" & index & HID.Value) = objNotaFiscal.Itens(index).Lotes
    End Sub

    Public Sub carregarLotes(pIndiceDoProduto As Integer)
        If pIndiceDoProduto >= 0 Then
            Limpar()
            index = pIndiceDoProduto

            If Session("ObjItemLotes" & index & HID.Value) Is Nothing Then
                RecuperaNotaFiscal()
                sessionLotes()
            End If

            If Session("ObjItemLotes" & index & HID.Value).Count > 0 Then
                gridInfLote.DataSource = CType(Session("ObjItemLotes" & index & HID.Value), ListNotaFiscalXLote).Where(Function(s) s.IUD <> "D")
            Else
                gridInfLote.DataSource = Nothing
            End If
            gridInfLote.DataBind()
        End If
    End Sub

    Private Function validarCampos()
        If txtNumeroDoLote.Text.Trim() = "" Or txtNumeroDoLote.Text = String.Empty Then
            MsgBox(Me.Page, "Número do Lote invalido.")
            Return False
        ElseIf txtDataFabricadoLote.Text.Trim() = "" Or txtDataFabricadoLote.Text.Trim() = String.Empty Then
            MsgBox(Me.Page, "Data de Fabricação invalido.")
            Return False
        ElseIf txtDataValidadeLote.Text.Trim() = "" Or txtDataValidadeLote.Text.Trim() = String.Empty Then
            MsgBox(Me.Page, "Data de Validade invalido.")
            Return False
        ElseIf txtQuantidadeConsumo.Text.Trim() = "" Or txtQuantidadeConsumo.Text.Trim() = String.Empty Then
            MsgBox(Me.Page, "Quantidade de consumo invalido.")
            Return False
        ElseIf txtQuantidadeLote.Text.Trim() = "" Or txtQuantidadeLote.Text.Trim() = String.Empty Then
            MsgBox(Me.Page, "Quantidade de Lote invalido.")
            Return False
        ElseIf CDate(txtDataFabricadoLote.Text) > CDate(txtDataValidadeLote.Text) Then
            MsgBox(Me.Page, "Validade não pode ser menor que a data de fabricação.")
            Return False
        End If

        RecuperaNotaFiscal()

        If objNotaFiscal.NossaEmissao Then
            MsgBox(Me.Page, "A nota não pode ser nossa emissão.")
            Return False
        End If

        Return True
    End Function

    Public Overrides Sub Limpar()
        seq.Visible = False
        seq.Text = String.Empty
        txtDataFabricadoLote.Text = String.Empty
        txtDataValidadeLote.Text = String.Empty
        txtNumeroDoLote.Text = String.Empty
        txtQuantidadeLote.Text = String.Empty
        txtQuantidadeConsumo.Text = String.Empty

        lnkAdicionar.Enabled = True
        lnkAdicionar.Visible = True
        lnkConfirmar.Enabled = False
        lnkConfirmar.Visible = False

        Session.Remove("index" & HID.Value)

        gridInfLote.DataSource = Nothing
        gridInfLote.DataBind()
    End Sub

    Protected Sub lnkSelecionarLote_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim imgLote As LinkButton = CType(sender, LinkButton)
        Dim row As GridViewRow = CType(imgLote.NamingContainer, GridViewRow)

        seq.Text = row.RowIndex
        txtNumeroDoLote.Text = row.Cells(1).Text
        txtDataFabricadoLote.Text = row.Cells(2).Text
        txtDataValidadeLote.Text = row.Cells(3).Text
        txtQuantidadeLote.Text = row.Cells(4).Text
        txtQuantidadeConsumo.Text = row.Cells(5).Text

        lnkAdicionar.Enabled = False
        lnkAdicionar.Visible = False
        lnkConfirmar.Enabled = True
        lnkConfirmar.Visible = True
    End Sub

    Protected Sub imgRemoverLote_Click(sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgLote As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(imgLote.NamingContainer, GridViewRow)

        If ValidarExcluirLote(row) Then

            RecuperaNotaFiscal()

            objNotaFiscal.Itens(index).Lotes.Item(row.RowIndex).IUD = "D"
            sessionLotes()
            carregarLotes(index)

        End If

    End Sub

    Private Function ValidarExcluirLote(ByVal row As GridViewRow) As Boolean

        RecuperaNotaFiscal()

        Dim strSQL As String

        strSQL = "  SELECT 1 FROM OrdemDeProducaoXConsumoXLote  " & vbCrLf &
                              "           WHERE Empresa_Id        = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "             AND EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "             AND Produto_Id      = '" & objNotaFiscal.Itens(row.RowIndex).Produto.Codigo & "'" & vbCrLf &
                              "             AND Lote_Id         = '" & row.Cells(1).Text & "'" & vbCrLf &
                              "             AND Validade        = '" & CDate(row.Cells(3).Text).ToString("yyyy-MM-dd") & "';"

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "OrdemDeProducaoXConsumoXLote")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count() > 0 Then
            MsgBox(Me.Page, "Existe relacionamento(s) na tabela consumo, não é possivel excluir!", eTitulo.Info, False)
            Return False
        End If

        strSQL = "  SELECT 1 FROM Producao  " & vbCrLf &
                 "  WHERE Empresa_Id        = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                 "    AND EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                 "    AND Produto_Id      = '" & objNotaFiscal.Itens(row.RowIndex).Produto.Codigo & "'" & vbCrLf &
                 "    AND Lote_Id         = '" & row.Cells(1).Text & "'" & vbCrLf &
                 "    AND Validade        = '" & CDate(row.Cells(3).Text).ToString("yyyy-MM-dd") & "';"

        ds = Banco.ConsultaDataSet(strSQL, "Producao")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count() > 0 Then
            MsgBox(Me.Page, "Existe relacionamento(s) na tabela produção, não é possivel excluir!", eTitulo.Info, False)
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkAdicionar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAdicionar.Click
        RecuperaNotaFiscal()

        If validarCampos() Then
            Dim nLote As NotaFiscalXLote = New NotaFiscalXLote

            nLote.IUD = "I"
            nLote.Lote = txtNumeroDoLote.Text
            nLote.Fabricado = txtDataFabricadoLote.Text
            nLote.Validade = txtDataValidadeLote.Text
            nLote.Quantidade = txtQuantidadeLote.Text
            nLote.QuantidadeDeConsumo = txtQuantidadeConsumo.Text

            objNotaFiscal.Itens(index).Lotes.Add(nLote)

            sessionLotes()
            carregarLotes(index)
        End If
    End Sub

    Protected Sub lnkConfirmar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConfirmar.Click
        RecuperaNotaFiscal()

        If validarCampos() Then
            Dim rowIndex As Integer = CInt(seq.Text)

            objNotaFiscal.Itens(index).Lotes.Item(rowIndex).IUD = "U"
            objNotaFiscal.Itens(index).Lotes.Item(rowIndex).Lote = txtNumeroDoLote.Text
            objNotaFiscal.Itens(index).Lotes.Item(rowIndex).Fabricado = txtDataFabricadoLote.Text
            objNotaFiscal.Itens(index).Lotes.Item(rowIndex).Validade = txtDataValidadeLote.Text
            objNotaFiscal.Itens(index).Lotes.Item(rowIndex).Quantidade = txtQuantidadeLote.Text
            objNotaFiscal.Itens(index).Lotes.Item(rowIndex).QuantidadeDeConsumo = txtQuantidadeConsumo.Text

            sessionLotes()
            carregarLotes(index)
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divControleNumeroDeLote")
    End Sub
End Class