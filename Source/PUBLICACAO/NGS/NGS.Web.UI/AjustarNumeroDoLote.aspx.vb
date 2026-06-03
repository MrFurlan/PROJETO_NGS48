Imports NGS.Lib.Negocio
Imports System.Threading
Imports System.IO

Public Class AjustarNumeroDoLote
    Inherits BasePage
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

#Region "Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If IsConnect AndAlso Not IsPostBack Then
                If Funcoes.VerificaPermissao("AjustarNumeroDoLote", "ACESSAR") Then
                    CargaUnidadeDeNegocio()
                    limpar()
                    Dim fm As New FilesManager()
                    If Not fm.IsConnect() Then
                        MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            CargaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, txtCliente.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("AjustarNumeroDoLote", "LEITURA") Then
                RecuperaNotaFiscal()
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    MsgBox(Me.Page, "É necessário selecionar a empresa!")
                    Exit Sub
                End If

                objNotaFiscal.CodigoEmpresa = Empresa(0)
                objNotaFiscal.EnderecoEmpresa = Empresa(1)
                objNotaFiscal.CarregandoNota = True

                If txtCodigoCliente.Value.ToString.Length > 0 Then
                    Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                    objNotaFiscal.CodigoCliente = Cliente(0)
                    objNotaFiscal.EnderecoCliente = Cliente(1)
                End If

                If txtES.Text.Length > 0 Then objNotaFiscal.EntradaSaida = IIf(txtES.Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)

                If txtNota.Text.Length > 0 Then
                    objNotaFiscal.Codigo = txtNota.Text
                Else
                    objNotaFiscal.Codigo = 0
                End If

                If txtSerie.Text.Length > 0 Then
                    objNotaFiscal.Serie = txtSerie.Text
                Else
                    objNotaFiscal.Serie = ""
                End If

                Session("ssCampo" & HID.Value) = "NotaXItens"

                ucConsultaPedidosXNotas.BindGridView()
                Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaCCe" & HID.Value.ToString)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLoteProduto_Click(sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            RecuperaNotaFiscal()
            If Not objNotaFiscal Is Nothing Then
                Dim Imgproduto As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(Imgproduto.NamingContainer, GridViewRow)

                ucControleNumeroDeLote.carregarLotes(row.RowIndex)
                Popup.ControleNumeroDeLote(Me.Page, "ObjLotes" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            RecuperaNotaFiscal()

            atualizarLotes()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        limpar()
    End Sub

#End Region

#Region "Methods"

    Private Sub limpar()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objNFConsultaCCe" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("ObjLotes" & HID.Value)

        RecuperaNotaFiscal()

        If objNotaFiscal IsNot Nothing Then
            RecuperaNotaFiscal()

            For i = 0 To objNotaFiscal.Itens.Count - 1
                Session.Remove("ObjItemLotes" & i & HID.Value)
            Next
        End If

        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtCliente.Text = ""
        txtES.Text = ""
        txtNota.Text = ""
        txtSerie.Text = ""

        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

        grdCceProdutos.DataSource = Nothing
        grdCceProdutos.DataBind()

        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        LiberaEmpresa()

        HID.Value = Guid.NewGuid().ToString

        objNotaFiscal = New [Lib].Negocio.NotaFiscal()
        SalvaNotaFiscal()

        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
        ucControleNumeroDeLote.SetarHID(HID.Value)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub atualizarXML()
        'Dim sql As String = "INSERT INTO nfe.fil (nomearquivoped, texto)" & vbCrLf & _
        '                    "VALUES ('carta000009459#03189063000126.txt'," & vbCrLf & _
        '                    "'chavenfe = 41150703189063000126550010000094591725919599"

        'For Each item In objNotaFiscal.Itens
        '    For Each Lote In item.Lotes
        '        For Each lote As NotaFiscalXLote In CType(Session("ObjItemLotes" & i & HID.Value), ListNotaFiscalXLote)

        '        Next
        '    Next
        'Next

    End Sub

    Private Sub atualizarLotes()
        objNotaFiscal = CType(Session("objNFConsultaCCe" & HID.Value), [Lib].Negocio.NotaFiscal)

        Dim Sqls As New ArrayList
        Dim sql As String = String.Empty
        Dim index As Integer = 0
        Dim obs As String = String.Empty
        Dim Observacao As String = String.Empty

        For i = 0 To objNotaFiscal.Itens.Count - 1

            For Each lote As NotaFiscalXLote In CType(Session("ObjItemLotes" & i & HID.Value), ListNotaFiscalXLote)

                If lote.IUD = "D" Then

                    sql = "     IF EXISTS(SELECT 1 FROM OrdemDeProducaoXConsumoXLote  " & vbCrLf &
                              "           WHERE Empresa_Id        = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "             AND EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "             AND Produto_Id      = '" & objNotaFiscal.Itens(i).Produto.Codigo & "'" & vbCrLf &
                              "             AND Lote_Id         = '" & lote.LoteOld & "'" & vbCrLf &
                              "             AND Validade        = '" & CDate(lote.ValidadeOld).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                              "           ) " & vbCrLf &
                              " BEGIN " & vbCrLf &
                              "     RAISERROR(N'Existe relacionamento(s) na tabela de consumo para esse lote!', -- Message text. " & vbCrLf &
                              "               16, -- Severity, " & vbCrLf &
                              "               1, -- State, " & vbCrLf &
                              "               N'number', -- First argument. " & vbCrLf &
                              "               5); -- Second argument. " & vbCrLf &
                              " END " & vbCrLf &
                              "  " & vbCrLf &
                              "  " & vbCrLf &
                              " IF EXISTS(SELECT 1 FROM Producao  " & vbCrLf &
                              "           WHERE Empresa_Id        = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "             AND EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "             AND Produto_Id      = '" & objNotaFiscal.Itens(i).Produto.Codigo & "'" & vbCrLf &
                              "             AND Lote_Id         = '" & lote.LoteOld & "'" & vbCrLf &
                              "             AND Validade        = '" & CDate(lote.ValidadeOld).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                              "           ) " & vbCrLf &
                              " BEGIN " & vbCrLf &
                              "     RAISERROR(N'Existe relacionamento(s) na tabela de produção para esse lote!', -- Message text. " & vbCrLf &
                              "               16, -- Severity, " & vbCrLf &
                              "               1, -- State, " & vbCrLf &
                              "               N'number', -- First argument. " & vbCrLf &
                              "               5); -- Second argument. " & vbCrLf &
                              " END " & vbCrLf &
                              "  " & vbCrLf &
                              "  " & vbCrLf &
                              " DELETE FROM NotaFiscalXLote " & vbCrLf &
                              " WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "     AND EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "     AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                              "     AND EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                              "     AND Serie_Id = '" & objNotaFiscal.Serie & "'" & vbCrLf &
                              "     AND EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                              "     AND Nota_Id = " & objNotaFiscal.Codigo & vbCrLf &
                              "     AND Lote_Id = '" & lote.LoteOld & "'" & vbCrLf &
                              "     AND Sequencia_Id = " & objNotaFiscal.Itens(i).Sequencia & ";"

                    Sqls.Add(sql)

                    obs = objNotaFiscal.ObservacoesControleInterno
                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Removido lote " & lote.LoteOld & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    Else
                        obs = "Removido lote " & lote.LoteOld & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    End If

                    objNotaFiscal.ObservacoesControleInterno = obs

                ElseIf lote.IUD = "U" Then

                    sql = " UPDATE NotaFiscalXLote SET " & vbCrLf &
                              " Lote_Id = '" & lote.Lote & "'," & vbCrLf &
                              " Validade = '" & CDate(lote.Validade).ToString("yyyy-MM-dd") & "'," & vbCrLf &
                              " Fabricado = '" & CDate(lote.Fabricado).ToString("yyyy-MM-dd") & "'," & vbCrLf &
                              " Quantidade = " & Replace(lote.Quantidade, ",", ".") & "," & vbCrLf &
                              " QuantidadeDeConsumo = " & Replace(lote.QuantidadeDeConsumo, ",", ".") & vbCrLf &
                              " WHERE Empresa_Id = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "     AND EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "     AND Cliente_Id = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                              "     AND EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                              "     AND Serie_Id = '" & objNotaFiscal.Serie & "'" & vbCrLf &
                              "     AND EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                              "     AND Nota_Id = " & objNotaFiscal.Codigo & vbCrLf &
                              "     AND Lote_Id = '" & lote.LoteOld & "'" & vbCrLf &
                              "     AND Sequencia_Id = " & objNotaFiscal.Itens(i).Sequencia & ";"

                    Sqls.Add(sql)

                    sql = " UPDATE OrdemDeProducaoXConsumoXLote SET " & vbCrLf &
                              " Lote_Id                 = '" & lote.Lote & "'," & vbCrLf &
                              " Validade                = '" & CDate(lote.Validade).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                              " WHERE Empresa_Id        = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "     AND EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "     AND Produto_Id      = '" & objNotaFiscal.Itens(i).Produto.Codigo & "'" & vbCrLf &
                              "     AND Lote_Id         = '" & lote.LoteOld & "'" & vbCrLf &
                              "     AND Validade        = '" & CDate(lote.ValidadeOld).ToString("yyyy-MM-dd") & "'" & ";"

                    Sqls.Add(sql)

                    sql = " UPDATE Producao SET " & vbCrLf &
                              " Lote_Id                 = '" & lote.Lote & "'," & vbCrLf &
                              " Fabricado = '" & CDate(lote.Fabricado).ToString("yyyy-MM-dd") & "'," & vbCrLf &
                              " Validade                = '" & CDate(lote.Validade).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                              " WHERE Empresa_Id        = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "     AND EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "     AND Produto_Id      = '" & objNotaFiscal.Itens(i).Produto.Codigo & "'" & vbCrLf &
                              "     AND Lote_Id         = '" & lote.LoteOld & "'" & vbCrLf &
                              "     AND Validade        = '" & CDate(lote.ValidadeOld).ToString("yyyy-MM-dd") & "'" & ";"

                    Sqls.Add(sql)

                    obs = objNotaFiscal.ObservacoesControleInterno
                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Alterado lote " & lote.Lote & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    Else
                        obs = "Alterado lote " & lote.Lote & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    End If

                    objNotaFiscal.ObservacoesControleInterno = obs

                ElseIf lote.IUD = "I" Then

                    sql = " INSERT INTO NotaFiscalXLote (" & vbCrLf &
                          " Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                          " Cliente_Id, EndCliente_Id, " & vbCrLf &
                          " EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf &
                          " Produto_Id, CFOP_Id, " & vbCrLf &
                          " Sequencia_Id, Lote_Id, " & vbCrLf &
                          " Quantidade, QuantidadeDeConsumo, " & vbCrLf &
                          " Validade, Fabricado) " & vbCrLf &
                          " VALUES ( " & vbCrLf &
                          " '" & objNotaFiscal.CodigoEmpresa & "'," & vbCrLf &
                          " " & objNotaFiscal.EnderecoEmpresa & "," & vbCrLf &
                          " '" & objNotaFiscal.CodigoCliente & "'," & vbCrLf &
                          " " & objNotaFiscal.EnderecoCliente & "," & vbCrLf &
                          " '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'," & vbCrLf &
                          " '" & objNotaFiscal.Serie & "'," & vbCrLf &
                                 objNotaFiscal.Codigo & "," & vbCrLf &
                                 objNotaFiscal.Itens(i).Produto.Codigo & "," & vbCrLf &
                                 objNotaFiscal.CFOP.Codigo & "," & vbCrLf &
                                 objNotaFiscal.Itens(i).Sequencia & "," & vbCrLf &
                          " '" & lote.Lote & "'," & vbCrLf &
                                 Replace(lote.Quantidade, ",", ".") & "," & vbCrLf &
                                 Replace(lote.QuantidadeDeConsumo, ",", ".") & "," & vbCrLf &
                          " '" & CDate(lote.Validade).ToString("yyyy-MM-dd") & "'," & vbCrLf &
                          " '" & CDate(lote.Fabricado).ToString("yyyy-MM-dd") & "')"

                    Sqls.Add(sql)

                    obs = objNotaFiscal.ObservacoesControleInterno
                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Incluido lote " & lote.Lote & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    Else
                        obs = "Incluido lote " & lote.Lote & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    End If

                    objNotaFiscal.ObservacoesControleInterno = obs

                End If

            Next

            index += 1

        Next

        If objNotaFiscal.ObservacoesControleInterno.Length > 0 Then

            sql = " Update NotasFiscais set " & vbCrLf &
                 "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                 "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf &
                 "      ObservacoesControleInterno  = '" & objNotaFiscal.ObservacoesControleInterno & "'" & vbCrLf &
                 "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                 "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                 "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                 "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                 "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                 "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                 "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"

            Sqls.Add(sql)

        End If

        If Sqls.Count > 0 Then

            If Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, "Ajustado com Sucesso.", eTitulo.Sucess)
                limpar()
            Else
                Observacao = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                MsgBox(Me.Page, Observacao)
                Exit Sub
            End If

        End If

    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsultaCCe" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaCCe" & HID.Value), NotaFiscal))

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                txtCliente.Text = itemCliente.Text
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie

                txtES.Enabled = False
                txtNota.Enabled = False
                txtSerie.Enabled = False

                grdCceProdutos.DataSource = objNotaFiscal.Itens
                grdCceProdutos.DataBind()

                SalvaNotaFiscal()

                lnkConsultar.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True

                ucControleNumeroDeLote.SetarLotes()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AjustarNumeroDoLote")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class