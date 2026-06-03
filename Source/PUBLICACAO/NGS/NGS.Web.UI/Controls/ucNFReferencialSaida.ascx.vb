Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

Public Class ucNFReferencialSaida
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
        gridConsultaNFReferencialSaida.DataSource = Nothing
        gridConsultaNFReferencialSaida.DataBind()
        Session.Remove("_MainUserControl")
    End Sub

    Public Sub BindGridView()
        SessaoRecuperaNotaFiscal()

        Dim ds As DataSet
        Dim sql As String

        sql = "SELECT n.Movimento, n.Operacao, n.SubOperacao, n.EntradaSaida_Id as EntradaSaida, n.Serie_Id AS Serie, n.Nota_Id AS Nota, ni.CFOP_Id, ni.Produto_Id, ni.Sequencia_Id, ni.QuantidadeFiscal AS Quantidade, ni.Valor" & vbCrLf & _
                "FROM NotasFiscais n" & vbCrLf & _
                "	inner join NotasFiscaisXItens ni" & vbCrLf & _
                "			on ni.Empresa_Id       = n.Empresa_Id" & vbCrLf & _
                "			and ni.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf & _
                "			and ni.Cliente_Id      = n.Cliente_Id" & vbCrLf & _
                "			and ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf & _
                "			and ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf & _
                "			and ni.Serie_Id        = n.Serie_Id" & vbCrLf & _
                "			and ni.Nota_Id         = n.Nota_Id" & vbCrLf & _
                "WHERE n.Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                "  and n.EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                "  and n.Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                "  and n.EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf

        If Left(objNotaFiscal.CodigoEmpresa, 8) = "05272759" AndAlso objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES Then
            sql &= "  and n.Situacao = 1" & vbCrLf
        Else
            sql &= "  and n.Pedido        = " & objNotaFiscal.CodigoPedido & vbCrLf & _
                    "  and n.Situacao = 1" & vbCrLf
        End If

        sql &= "  and ni.QuantidadeFiscal > 0"

        ds = Banco.ConsultaDataSet(sql, "NotasDeSaida")
        gridConsultaNFReferencialSaida.DataSource = ds
        gridConsultaNFReferencialSaida.DataBind()

    End Sub


    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Selecionar()
    End Sub


    Protected Overrides Sub Selecionar()
        SessaoRecuperaNotaFiscal()

        Try
            Dim lstNotasFiscaisReferenciais As New ListNotaFiscalReferencial()
            Dim objNotaFiscalReferencial As NotaFiscalReferencial

            Dim i As Integer = 0
            While i < gridConsultaNFReferencialSaida.Rows.Count

                If CType(gridConsultaNFReferencialSaida.Rows(i).FindControl("rdSelecionado"), CheckBox).Checked Then

                    objNotaFiscalReferencial = New NotaFiscalReferencial()
                    objNotaFiscalReferencial.IUD = "I"

                    objNotaFiscalReferencial.Empresa_Id = objNotaFiscal.CodigoEmpresa
                    objNotaFiscalReferencial.EndEmpresa_Id = objNotaFiscal.EnderecoEmpresa
                    objNotaFiscalReferencial.Cliente_Id = objNotaFiscal.CodigoCliente
                    objNotaFiscalReferencial.EndCliente_Id = objNotaFiscal.EnderecoCliente
                    objNotaFiscalReferencial.EntradaSaida_Id = IIf(gridConsultaNFReferencialSaida.Rows(i).Cells(2).Text() = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    objNotaFiscalReferencial.Serie_Id = gridConsultaNFReferencialSaida.Rows(i).Cells(3).Text()
                    objNotaFiscalReferencial.Nota_Id = gridConsultaNFReferencialSaida.Rows(i).Cells(4).Text()
                    objNotaFiscalReferencial.CFOP_Id = gridConsultaNFReferencialSaida.Rows(i).Cells(5).Text()
                    objNotaFiscalReferencial.Produto_Id = gridConsultaNFReferencialSaida.Rows(i).Cells(6).Text()
                    objNotaFiscalReferencial.Sequencia_Id = gridConsultaNFReferencialSaida.Rows(i).Cells(7).Text()
                    objNotaFiscalReferencial.Quantidade = gridConsultaNFReferencialSaida.Rows(i).Cells(8).Text()  'quantidade da nota
                    objNotaFiscalReferencial.Valor = gridConsultaNFReferencialSaida.Rows(i).Cells(9).Text()  'valor da nota
                    objNotaFiscalReferencial.TipoReferencial_Id = eTipoReferencial.NFC
                    lstNotasFiscaisReferenciais.Add(objNotaFiscalReferencial)
                End If

                i += 1
            End While

            'SessaoSalvaNotaFiscal()

            Session(Session("ssTipoRetorno")) = lstNotasFiscaisReferenciais

            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(lstNotasFiscaisReferenciais)
                Else
                    CType(Me.Page, IBasePage).Carregar(lstNotasFiscaisReferenciais)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaNFReferencialSaida")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaNFReferencialSaida")
    End Sub

End Class