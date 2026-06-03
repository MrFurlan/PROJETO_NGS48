Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

Public Class ucNFProdutor
    Inherits BaseUserControl

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Me.Limpar()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), NotaFiscal)
    End Sub

    Public Overrides Sub Limpar()
        gridConsultaNFProdutor.DataSource = Nothing
        gridConsultaNFProdutor.DataBind()
        Session.Remove("_MainUserControl")
    End Sub

    Public Sub BindGridView()
        SessaoRecuperaNotaFiscal()

        Dim ds As DataSet
        Dim sql As String

        sql = "SELECT n.Movimento, n.Operacao, n.SubOperacao, n.Serie_Id AS Serie, n.Nota_Id AS Nota, ni.QuantidadeFiscal AS Quantidade, ni.Valor" & vbCrLf & _
                "FROM NotasFiscais n" & vbCrLf & _
                "	inner join NotasFiscaisXItens ni" & vbCrLf & _
                "			on ni.Empresa_Id       = n.Empresa_Id" & vbCrLf & _
                "			and ni.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf & _
                "			and ni.Cliente_Id      = n.Cliente_Id" & vbCrLf & _
                "			and ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf & _
                "			and ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf & _
                "			and ni.Serie_Id        = n.Serie_Id" & vbCrLf & _
                "			and ni.Nota_Id         = n.Nota_Id" & vbCrLf & _
                "	left join NotasXNotas nXt" & vbCrLf & _
                "			on nXt.OrigemEmpresa_Id       = n.Empresa_Id" & vbCrLf & _
                "			and nXt.OrigemEndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf & _
                "			and nXt.OrigemCliente_Id      = n.Cliente_Id" & vbCrLf & _
                "			and nXt.OrigemEndCliente_Id   = n.EndCliente_Id" & vbCrLf & _
                "			and nXt.OrigemEntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf & _
                "			and nXt.OrigemSerie_Id        = n.Serie_Id" & vbCrLf & _
                "			and nXt.OrigemNota_Id         = n.Nota_Id" & vbCrLf & _
                "WHERE n.Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                "  and n.EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                "  and n.Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                "  and n.EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                "  and n.Operacao        = " & objNotaFiscal.CodigoOperacao & vbCrLf & _
                "  and n.EntradaSaida_Id = 'E'" & vbCrLf & _
                "  and (n.TipoDeDocumento = 15 or n.Finalidade = 30)" & vbCrLf & _
                "  and ni.Produto_id     = '" & objNotaFiscal.Itens(0).CodigoProduto & "'" & vbCrLf & _
                "  and nXt.Nota_Id is null"

        ds = Banco.ConsultaDataSet(sql, "NotasDeProdutor")
        gridConsultaNFProdutor.DataSource = ds
        gridConsultaNFProdutor.DataBind()

    End Sub

    Protected Sub gridConsultaNFProdutor_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Selecionar()
    End Sub

    Protected Overrides Sub Selecionar()
        SessaoRecuperaNotaFiscal()

        Try
            Dim nfp As New NotaFiscal
            nfp.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
            nfp.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
            nfp.CodigoCliente = objNotaFiscal.CodigoCliente
            nfp.EnderecoCliente = objNotaFiscal.EnderecoCliente
            nfp.EntradaSaida = objNotaFiscal.EntradaSaida
            nfp.Serie = gridConsultaNFProdutor.SelectedRow.Cells(2).Text()
            nfp.Codigo = gridConsultaNFProdutor.SelectedRow.Cells(3).Text()

            nfp = New NotaFiscal(nfp)

            If nfp.Itens Is Nothing OrElse nfp.Itens(0).ValorTotal = 0 Then
                MsgBox(Me.Page, "Erro ao selecionar Nota Fiscal, verifique!", eTitulo.Info)
                Exit Sub
            End If

            objNotaFiscal.NotaTrocaOrigem = New NotaFiscal(nfp)

            SessaoSalvaNotaFiscal()

            Session(Session("ssTipoRetorno")) = nfp

            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(nfp)
                Else
                    CType(Me.Page, IBasePage).Carregar(nfp)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaNFProdutor")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFecharUCProdutor_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFecharUCProdutor.Click
        Popup.CloseDialog(Me.Page, "divConsultaNFProdutor")
    End Sub

End Class