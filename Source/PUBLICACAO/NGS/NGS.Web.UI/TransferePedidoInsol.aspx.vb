Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TransferePedidoInsol
    Inherits BasePage

    Private sConnStringDestino As String = "Provider=SqlOleDb;" & _
                      "Data Source=SRVCWB;" & _
                      "Initial Catalog=Insoja;" & _
                      "User Id=sa;" & _
                      "Password=pwd_in$ol@ctba"

    Private objPedido As [Lib].Negocio.Pedido

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TransferePedidoInsol", "ACESSAR") Then
                CargaUnidadeDeNegocio()
                limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Function GravaBancoDestino(ByVal sql As ArrayList, ByVal strConexao As String) As Boolean
        Dim SqlTrans As OleDb.OleDbTransaction
        Dim Sqlcommand As New OleDb.OleDbCommand
        Dim SQlConnectionDestino As New OleDb.OleDbConnection

        Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/gravabanco.txt"))

        Try
            If (SQlConnectionDestino.State = ConnectionState.Closed) Then
                SQlConnectionDestino = New OleDb.OleDbConnection(strConexao)
                SQlConnectionDestino.Open()
            End If
            SqlTrans = SQlConnectionDestino.BeginTransaction(IsolationLevel.ReadCommitted)
            Sqlcommand.Connection = SQlConnectionDestino
            Sqlcommand.Transaction = SqlTrans
            For index = 0 To sql.Count - 1
                Sqlcommand.CommandText = CStr(sql.Item(index))
                strm.WriteLine(CStr(sql.Item(index)))
                Sqlcommand.ExecuteNonQuery()
            Next
            strm.Close()
            SqlTrans.Commit()
            SQlConnectionDestino.Close()
            Return True
        Catch SqlException As OleDb.OleDbException
            strm.Close()

            Select Case SqlException.ErrorCode
                Case 2627
                    Throw New Exception("Registro já cadastrado!")
                Case 647
                    Throw New Exception("Registro esta sendo utilizado em outras tabelas!")
                Case Else
                    Throw New Exception(SqlException.Message)
            End Select

            SqlTrans.Rollback()
            SQlConnectionDestino.Close()
            Return False
        End Try
    End Function

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)

        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Private Sub limpar()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)

        lnkAtualizar.Enabled = False

        txtCliente.Text = String.Empty
        txtPedido.Text = String.Empty

        GridTransferePedidoInsol.DataBind()

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
    End Sub

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCliente.Text) AndAlso String.IsNullOrWhiteSpace(txtPedido.Text) Then
            MsgBox(Me.Page, "Informe o cliente ou o numero do pedido", eTitulo.Info)
            Return False
        End If

        Return True
    End Function

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
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("TransferePedidoInsol", "LEITURA") Then
                If ValidaCampos() Then
                    Dim lst As New [Lib].Negocio.Pedidos(0, IIf(txtPedido.Text.Length = 0, 0, txtPedido.Text), "", ddlUnidadeDeNegocio.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(0), 0, txtCodigoCliente.Value.Split("-")(0))

                    GridTransferePedidoInsol.DataSource = lst
                    GridTransferePedidoInsol.DataBind()

                    If lst.Count = 0 Then
                        lnkAtualizar.Enabled = False
                        MsgBox(Me.Page, "nehum resultado encontrado!", eTitulo.Info)
                    Else
                        lnkAtualizar.Enabled = True
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click

        Dim strmPed As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/pedidos.txt"))

        Try
            If Funcoes.VerificaPermissao("TransferePedidoInsol", "ALTERAR") Then
                Dim Sqls As New ArrayList
                Dim Sql As String = String.Empty

                Dim i As Integer = 0
                While i < GridTransferePedidoInsol.Rows.Count


                    If CType(GridTransferePedidoInsol.Rows(i).FindControl("chkPedido"), CheckBox).Checked Then
                        strmPed.WriteLine("Transferindo Pedido " & GridTransferePedidoInsol.Rows(i).Cells(1).Text)

                        'VERIFICA SE PEDIDO JÁ EXISTE NA BASE
                        Sql = " SELECT Empresa_id," & vbCrLf & _
                              "	       EndEmpresa_id," & vbCrLf & _
                              "	       Pedido_Id" & vbCrLf & _
                              " FROM Insoja.dbo.Pedidos" & vbCrLf & _
                              " WHERE Empresa_id       = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                              "   AND EndEmpresa_id    = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                              "   And Pedido_Id        = " & GridTransferePedidoInsol.Rows(i).Cells(1).Text & vbCrLf

                        Dim dsPed As New DataSet
                        dsPed = Banco.ConsultaDataSet(Sql, "Pedidos")

                        If dsPed IsNot Nothing AndAlso dsPed.Tables(0).Rows.Count > 0 Then
                            strmPed.Close()
                            MsgBox(Me.Page, "Pedido " & GridTransferePedidoInsol.Rows(i).Cells(1).Text & " já exite no Destino")
                            Exit Sub
                        End If

                        objPedido = New [Lib].Negocio.Pedido(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), GridTransferePedidoInsol.Rows(i).Cells(1).Text)

                        objPedido.IUD = "U"

                        If objPedido.Itens.Count = 0 Then
                            strmPed.Close()
                            MsgBox(Me.Page, "Pedido sem itens " & objPedido.Codigo)
                            Exit Sub
                        End If

                        For Each item In objPedido.Itens
                            If item.Encargos.Count = 0 Then
                                strmPed.Close()
                                MsgBox(Me.Page, "Pedido sem encargos" & objPedido.Codigo)
                                Exit Sub
                            End If
                        Next

                        objPedido.IUD = "I"

                        Sql = "Merge Pedidos as Dest" & vbCrLf & _
                              " USING (Select '" & objPedido.CodigoEmpresa & "' as Empresa_Id," & objPedido.EnderecoEmpresa & " as EndEmpresa_Id," & objPedido.Codigo & " as Pedido_Id) AS Ori" & vbCrLf & _
                              "    ON Dest.Empresa_Id    = Ori.Empresa_Id" & vbCrLf & _
                              "   and Dest.EndEmpresa_Id = Ori.EndEmpresa_Id" & vbCrLf & _
                              "   and Dest.Pedido_Id     = Ori.Pedido_Id" & vbCrLf & _
                              "  WHEN NOT MATCHED" & vbCrLf & _
                              "    THEN Insert (Empresa_Id, EndEmpresa_Id," & vbCrLf & _
                              "        Pedido_Id, UnidadeDeNegocio, Cliente, EndCliente, " & vbCrLf & _
                              "        Praca, EndPraca, PedidoEfetivo, Safra, Moeda, TemVariacao, Indexador, Operacao, SubOperacao, Situacao, " & vbCrLf & _
                              "        DataPedido, DataInicioEntrega, DataEntrega, PedidoOrigem, Solicitacao, UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf & _
                              "        UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, " & vbCrLf & _
                              "        UsuarioLiberacao, UsuarioLiberacaoData, Observacoes, CondicaoPagamento, BancoCliente, " & vbCrLf & _
                              "        AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, FreteCIFFOB, OrigemDestino, Comercializacao, " & vbCrLf & _
                              "        Finalidade, Contrato, Carteira, Taxa, VencimentoPedido,LocalEmbarque,EndLocalEmbarque, MomentoFinanceiro, " & vbCrLf & _
                              "        AgruparFinanceiro, IndiceFixado, EstadoEntrega, CidadeEntrega, FiscalAberto, FinanceiroAberto, " & vbCrLf & _
                              "        EmpresaTroca, EndEmpresaTroca, PedidoTroca, ContaAdiantamentoTroca, CondicaoPagamentoEntrega, QuotaEntrega, PeriodicidadeEntrega, PedidoBloqueado," & vbCrLf & _
                              "        VersaoPedido, VersaoUsuario, VersaoHorarioBloqueio, IndexadorFixo, Troca, Antecipada, Recompra)" & vbCrLf & _
                              " VALUES ('" & objPedido.CodigoEmpresa & "', " & objPedido.EnderecoEmpresa.ToString() & "," & vbCrLf & _
                              objPedido.Codigo & ",'" & objPedido.UnidadeNegocio.Codigo & "'," & vbCrLf & _
                              "'" & objPedido.CodigoCliente & "'," & objPedido.EnderecoCliente & "," & vbCrLf & _
                              "'" & objPedido.CodigoPraca & "'," & objPedido.EnderecoPraca & "," & vbCrLf & _
                              "'" & objPedido.PedidoEfetivo & "','" & objPedido.CodigoSafra & "'," & vbCrLf & _
                              objPedido.CodigoMoeda & "," & IIf(objPedido.TemVariacao, 1, 0) & "," & objPedido.CodigoIndexador & "," & vbCrLf & _
                              objPedido.CodigoOperacao & "," & objPedido.CodigoSubOperacao & "," & vbCrLf & _
                              objPedido.CodigoSituacao & ",'" & objPedido.DataPedido.ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                              "'" & objPedido.DataEntregaInicial.ToString("yyyy-MM-dd") & "','" & objPedido.DataEntregaFinal.ToString("yyyy-MM-dd") & "', '0', 0," & vbCrLf & _
                              "'" & objPedido.UsuarioInclusao & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112)," & vbCrLf & _
                              "'" & objPedido.UsuarioAlteracao & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112)," & vbCrLf & _
                              "'" & objPedido.UsuarioCancelamento & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112)," & vbCrLf & _
                              "'" & objPedido.UsuarioLiberacao & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112)," & vbCrLf & _
                              "'" & objPedido.Observacoes & "'," & vbCrLf

                        If Not objPedido.CondicaoPagamento Is Nothing Then Sql &= objPedido.CondicaoPagamento.Codigo & ", " Else Sql &= "0, "

                        'BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente
                        If Not objPedido.ContaBancariaSelecionada Is Nothing Then
                            Sql &= objPedido.ContaBancariaSelecionada.CodigoBanco & "," & vbCrLf & _
                                   "'" & objPedido.ContaBancariaSelecionada.CodigoAgencia & "','" & objPedido.ContaBancariaSelecionada.DigitoAgencia & "'," & vbCrLf & _
                                   "'" & objPedido.ContaBancariaSelecionada.ContaCorrente & "','" & objPedido.ContaBancariaSelecionada.DigitoConta & "'," & vbCrLf
                        Else
                            Sql &= "0, '', '', '', '', "
                        End If

                        'FreteCIFFOB, OrigemDestino, Comercializacao
                        Sql &= "'" & objPedido.FreteCIFFOB.ToString() & "','" & objPedido.OrigemDestino & "', '" & CInt(objPedido.Comercializacao) & "', "
                        'Finalidade
                        Sql &= objPedido.CodigoFinalidade

                        Sql &= ",'" & objPedido.Contrato.ToString() & "','" & objPedido.CodigoCarteira & "'" & vbCrLf & _
                               "," & Str(objPedido.Taxa) & ",'" & objPedido.DataVencimentoPedido.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                               "," & IIf(objPedido.CodigoLocalEmbarque.Length > 0, "'" & objPedido.CodigoLocalEmbarque & "'", "NULL") & vbCrLf & _
                               "," & IIf(objPedido.CodigoLocalEmbarque.Length > 0, objPedido.EndLocalEmbarque, "NULL") & "," & objPedido.MomentoFinanceiro & "," & vbCrLf & _
                               IIf(objPedido.AgruparFinanceiro, 1, 0) & "," & Str(objPedido.IndiceFixado) & ",'" & objPedido.EstadoEntrega & "','" & objPedido.CidadeEntrega & "'," & vbCrLf & _
                               IIf(objPedido.FiscalAberto, 1, 0) & "," & IIf(objPedido.FinanceiroAberto, 1, 0) & "," & vbCrLf & _
                               "'" & objPedido.CodigoEmpresaTroca & "'," & vbCrLf & _
                               objPedido.EnderecoEmpresaTroca & "," & vbCrLf & _
                               objPedido.CodigoPedidoTroca & "," & vbCrLf & _
                               "'" & objPedido.ContaAdiantamentoTroca & "'," & vbCrLf

                        Sql &= objPedido.CodigoCondicaoPagamentoDaEntrega & ", "
                        Sql &= Str(objPedido.QuotaEntrega) & "," & objPedido.PeriodicidadeEntrega & "," & IIf(objPedido.PedidoBloqueado, 1, 0)
                        Sql &= "," & IIf(objPedido.IUD = "I", 1, objPedido.VersaoPedido + 1) & ",'" & UsuarioServidor.NomeUsuario & "', NULL,"
                        Sql &= IIf(objPedido.IndexadorFixo, 1, 0) & "," & IIf(objPedido.Troca, 1, 0) & "," & IIf(objPedido.Antecipada, 1, 0) & "," & IIf(objPedido.Recompra, 1, 0) & ")"
                        Sql &= " WHEN MATCHED " & vbCrLf & _
                               "   THEN Update set " & vbCrLf & _
                               "    Cliente                ='" & objPedido.CodigoCliente & "'" & vbCrLf & _
                               "   ,EndCliente             = " & objPedido.EnderecoCliente & vbCrLf & _
                               "   ,Praca                  ='" & objPedido.CodigoPraca & "'" & vbCrLf & _
                               "   ,EndPraca               = " & objPedido.EnderecoPraca & vbCrLf & _
                               "   ,PedidoEfetivo          ='" & objPedido.PedidoEfetivo & "'" & vbCrLf & _
                               "   ,Safra                  ='" & objPedido.CodigoSafra & "'" & vbCrLf & _
                               "   ,Moeda                  = " & objPedido.CodigoMoeda & vbCrLf & _
                               "   ,Indexador              = " & objPedido.CodigoIndexador & vbCrLf & _
                               "   ,Operacao               = " & objPedido.CodigoOperacao & vbCrLf & _
                               "   ,SubOperacao            = " & objPedido.CodigoSubOperacao & vbCrLf & _
                               "   ,Situacao               = " & objPedido.CodigoSituacao & vbCrLf & _
                               "   ,DataPedido             ='" & objPedido.DataPedido.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                               "   ,DataInicioEntrega      ='" & objPedido.DataEntregaInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                               "   ,DataEntrega            ='" & objPedido.DataEntregaFinal.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                               "   ,CondicaoPagamento      ='" & objPedido.CodigoCondicaoPagamento & "'" & vbCrLf & _
                               "   ,Contrato               ='" & objPedido.Contrato & "'" & vbCrLf & _
                               "   ,Carteira               ='" & objPedido.CodigoCarteira & "'" & vbCrLf & _
                               "   ,Taxa                   = " & Str(objPedido.Taxa) & vbCrLf & _
                               "   ,VencimentoPedido       ='" & objPedido.DataVencimentoPedido.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                               "   ,LocalEmbarque          = " & IIf(objPedido.CodigoLocalEmbarque.Length > 0, "'" & objPedido.CodigoLocalEmbarque & "'", "NULL") & vbCrLf & _
                               "   ,EndLocalEmbarque       = " & IIf(objPedido.CodigoLocalEmbarque.Length > 0, "'" & objPedido.EndLocalEmbarque & "'", "NULL") & vbCrLf & _
                               "   ,MomentoFinanceiro      = " & objPedido.MomentoFinanceiro & vbCrLf & _
                               "   ,AgruparFinanceiro      = " & IIf(objPedido.AgruparFinanceiro, 1, 0) & vbCrLf & _
                               "   ,IndiceFixado           = " & Str(objPedido.IndiceFixado) & vbCrLf

                        If Not objPedido.ContaBancariaSelecionada Is Nothing Then
                            Sql &= "   ,BancoCliente         = " & objPedido.ContaBancariaSelecionada.CodigoBanco & vbCrLf & _
                                   "   ,AgenciaCliente       ='" & objPedido.ContaBancariaSelecionada.CodigoAgencia & "'" & vbCrLf & _
                                   "   ,DigitoAgenciaCliente ='" & objPedido.ContaBancariaSelecionada.DigitoAgencia & "'" & vbCrLf & _
                                   "   ,ContaCliente         ='" & objPedido.ContaBancariaSelecionada.ContaCorrente & "'" & vbCrLf & _
                                   "   ,DigitoContaCliente   ='" & objPedido.ContaBancariaSelecionada.DigitoConta & "'" & vbCrLf
                        Else
                            Sql &= "   ,BancoCliente         = 0" & vbCrLf & _
                                   "   ,AgenciaCliente       =''" & vbCrLf & _
                                   "   ,DigitoAgenciaCliente =''" & vbCrLf & _
                                   "   ,ContaCliente         =''" & vbCrLf & _
                                   "   ,DigitoContaCliente   =''" & vbCrLf
                        End If

                        Sql &= "   ,Observacoes          ='" & objPedido.Observacoes & "'" & vbCrLf & _
                               "   ,FreteCIFFOB          ='" & objPedido.FreteCIFFOB.ToString() & "'" & vbCrLf & _
                               "   ,OrigemDestino        ='" & objPedido.OrigemDestino & "'" & vbCrLf & _
                               "   ,Comercializacao      ='" & CInt(objPedido.Comercializacao) & "'" & vbCrLf & _
                               "   ,UsuarioAlteracao     ='" & objPedido.UsuarioAlteracao & "'" & vbCrLf & _
                               "   ,UsuarioAlteracaoData = CURRENT_TIMESTAMP" & vbCrLf

                        If objPedido.CodigoSituacao = eSituacao.Cancelado Then
                            Sql &= "   ,UsuarioCancelamento     = '" & objPedido.UsuarioCancelamento & "'" & vbCrLf & _
                                   "   ,UsuarioCancelamentoData = CURRENT_TIMESTAMP" & vbCrLf
                        End If

                        Sql &= "   ,Finalidade = " & objPedido.CodigoFinalidade & vbCrLf & _
                               "   ,EstadoEntrega            ='" & objPedido.EstadoEntrega & "'" & vbCrLf & _
                               "   ,CidadeEntrega            ='" & objPedido.CidadeEntrega & "'" & vbCrLf & _
                               "   ,FiscalAberto             = " & IIf(objPedido.FiscalAberto, 1, 0) & vbCrLf & _
                               "   ,FinanceiroAberto         = " & IIf(objPedido.FinanceiroAberto, 1, 0) & vbCrLf & _
                               "   ,EmpresaTroca             ='" & objPedido.CodigoEmpresaTroca & "'" & vbCrLf & _
                               "   ,EndEmpresaTroca          = " & objPedido.EnderecoEmpresaTroca & vbCrLf & _
                               "   ,PedidoTroca              = " & objPedido.CodigoPedidoTroca & vbCrLf & _
                               "   ,CondicaoPagamentoEntrega = " & objPedido.CodigoCondicaoPagamentoDaEntrega & vbCrLf & _
                               "   ,QuotaEntrega             = " & Str(objPedido.QuotaEntrega) & vbCrLf & _
                               "   ,PeriodicidadeEntrega     = " & objPedido.PeriodicidadeEntrega & vbCrLf & _
                               "   ,PedidoBloqueado          = " & IIf(objPedido.PedidoBloqueado, 1, 0) & vbCrLf & _
                               "   ,VersaoPedido             = " & (objPedido.VersaoPedido + 1) & vbCrLf & _
                               "   ,VersaoUsuario            = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                               "   ,IndexadorFixo            = " & IIf(objPedido.IndexadorFixo, 1, 0) & vbCrLf & _
                               "   ,Troca                    = " & IIf(objPedido.Troca, 1, 0) & vbCrLf & _
                               "   ,Antecipada               = " & IIf(objPedido.Antecipada, 1, 0) & vbCrLf & _
                               "   ,Recompra                 = " & IIf(objPedido.Recompra, 1, 0) & ";" & vbCrLf
                        Sqls.Add(Sql)

                        strmPed.WriteLine("Transferiu Pedido")

                        objPedido.Itens.SalvarSql(Sqls)

                        strmPed.WriteLine("Transferiu Itens")

                        objPedido.Depositos.SalvarSql(Sqls)

                        strmPed.WriteLine("Transferiu Depositos")

                        If objPedido.Roteiros.Count > 0 Then
                            objPedido.Roteiros.SalvarSql(Sqls)
                            strmPed.WriteLine("Transferiu Roteiros")
                        End If

                        If objPedido.Transportadores.Count > 0 Then
                            objPedido.Transportadores.SalvarSql(Sqls)
                            strmPed.WriteLine("Transferiu Transportadores")
                        End If

                        If objPedido.Representantes.Count > 0 Then
                            objPedido.Representantes.SalvarSql(Sqls)
                            strmPed.WriteLine("Transferiu Representantes")
                        End If

                        If Not objPedido.PedidoTroca Is Nothing AndAlso objPedido.CodigoPedidoTroca > 0 AndAlso objPedido.Troca AndAlso objPedido.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
                            Sql = "UPDATE Pedidos SET" & _
                                  "   EmpresaTroca    ='" & objPedido.CodigoEmpresa & "'" & _
                                  "  ,EndEmpresaTroca = " & objPedido.EnderecoEmpresa & _
                                  "  ,PedidoTroca     = " & objPedido.Codigo & _
                                  " WHERE Empresa_Id    ='" & objPedido.CodigoEmpresaTroca & "'" & _
                                  "   AND EndEmpresa_Id = " & objPedido.EnderecoEmpresaTroca & _
                                  "   AND Pedido_Id     = " & objPedido.CodigoPedidoTroca

                            Sqls.Add(Sql)
                            strmPed.WriteLine("Transferiu Pedido de Troca")
                        End If

                        If objPedido.Vencimentos.Count > 0 Then
                            For Each tit In objPedido.Vencimentos
                                Sql = tit.Incluir(UsuarioServidor.NomeServidor.ToUpper(), tit.Codigo)
                                Sqls.Add(Sql)
                                strmPed.WriteLine("Transferiu Titulo A " & IIf(tit.EntradaSaida = eEntradaSaida.Saida, "Receber", "Pagar") & " - " & tit.Codigo)
                            Next
                        End If


                        'SALVAR LAUDO DE PESAGEM
                        Sql = " SELECT Empresa_id," & vbCrLf & _
                              "	       EndEmpresa_id," & vbCrLf & _
                              "	       Pesagem_Id" & vbCrLf & _
                              " FROM Pesagem" & vbCrLf & _
                              " WHERE Situacao = 1 " & vbCrLf & _
                              "   AND Empresa_id       = '" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                              "   AND EndEmpresa_id    = " & objPedido.EnderecoEmpresa & vbCrLf & _
                              "   And Pedido           = " & objPedido.Codigo & vbCrLf

                        Dim dsP As New DataSet
                        dsP = Banco.ConsultaDataSet(Sql, "Pesagem")

                        If dsP IsNot Nothing AndAlso dsP.Tables(0).Rows.Count > 0 Then
                            For Each drP As DataRow In dsP.Tables(0).Rows
                                Dim objLaudo As [Lib].Negocio.Pesagem = New [Lib].Negocio.Pesagem(drP("Empresa_id"), drP("EndEmpresa_id"), drP("Pesagem_Id"))

                                objLaudo.IUD = "U"

                                If objLaudo.CodigoRomaneio > 0 AndAlso objLaudo.Romaneios.Count = 0 Then
                                    strmPed.Close()
                                    MsgBox(Me.Page, "Laudo " & objLaudo.Codigo & " sem romaneio")
                                    Exit Sub
                                End If

                                objLaudo.IUD = "I"

                                objLaudo.SalvarSql(Sqls)

                                strmPed.WriteLine("Transferiu Laudo " & objLaudo.Codigo)

                                For Each ro In objLaudo.Romaneios
                                    ro.IUD = "I"
                                    ro.SalvarSql(Sqls, False)
                                    strmPed.WriteLine("Transferiu Romaneio " & ro.Codigo)
                                Next
                            Next
                        End If


                        'SALVAR ROMANEIOS DE NOTA FISCAL
                        Sql = " SELECT Empresa_id," & vbCrLf & _
                              "	       EndEmpresa_id," & vbCrLf & _
                              "	       Romaneio_Id" & vbCrLf & _
                              " FROM Romaneios" & vbCrLf & _
                              " WHERE Processo         = 'NOTA FISCAL'" & vbCrLf & _
                              "   AND Empresa_id       = '" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                              "   AND EndEmpresa_id    = " & objPedido.EnderecoEmpresa & vbCrLf & _
                              "   And Pedido           = " & objPedido.Codigo & vbCrLf

                        Dim dsR As New DataSet
                        dsR = Banco.ConsultaDataSet(Sql, "Romaneios")

                        If dsR IsNot Nothing AndAlso dsR.Tables(0).Rows.Count > 0 Then
                            For Each drR As DataRow In dsR.Tables(0).Rows
                                Dim objRomaneio As [Lib].Negocio.Romaneio = New [Lib].Negocio.Romaneio(drR("Empresa_id"), drR("EndEmpresa_id"), drR("Romaneio_Id"))

                                objRomaneio.IUD = "I"

                                objRomaneio.SalvarSql(Sqls, False)

                                strmPed.WriteLine("Transferiu Romaneios " & objRomaneio.Codigo)
                            Next
                        End If


                        'SALVAR NOTA FISCAL
                        Sql = " SELECT Empresa_id," & vbCrLf & _
                              "	       EndEmpresa_id," & vbCrLf & _
                              "	       Cliente_Id," & vbCrLf & _
                              "	       EndCliente_Id," & vbCrLf & _
                              "	       EntradaSaida_Id," & vbCrLf & _
                              "	       Serie_Id," & vbCrLf & _
                              "	       Nota_Id" & vbCrLf & _
                              " FROM NotasFiscais" & vbCrLf & _
                              " WHERE Situacao = 1 " & vbCrLf & _
                              "   AND Empresa_id       = '" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                              "   AND EndEmpresa_id    = " & objPedido.EnderecoEmpresa & vbCrLf & _
                              "   And Pedido = " & objPedido.Codigo & vbCrLf

                        Dim dsN As New DataSet
                        dsN = Banco.ConsultaDataSet(Sql, "NotasFiscais")

                        If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                            For Each drN As DataRow In dsN.Tables(0).Rows
                                Dim objNnota As New [Lib].Negocio.NotaFiscal()
                                objNnota.CodigoEmpresa = drN("Empresa_id")
                                objNnota.EnderecoEmpresa = drN("EndEmpresa_id")
                                objNnota.CodigoCliente = drN("Cliente_Id")
                                objNnota.EnderecoCliente = drN("EndCliente_Id")
                                objNnota.EntradaSaida = IIf(drN("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                                objNnota.Serie = drN("Serie_Id")
                                objNnota.Codigo = drN("Nota_Id")
                                objNnota = New [Lib].Negocio.NotaFiscal(objNnota)

                                objNnota.IUD = "U"

                                If objNnota.Itens.Count > 0 Then
                                    For Each itemNota In objNnota.Itens
                                        If itemNota.Encargos.Count = 0 Then
                                            strmPed.Close()
                                            MsgBox(Me.Page, "Notas Fiscais sem encargos: Nota " & objNnota.Codigo & " Pedido " & objNnota.CodigoPedido)
                                            Exit Sub
                                        End If
                                    Next
                                Else
                                    strmPed.Close()
                                    MsgBox(Me.Page, "Notas Fiscais sem Itens: Nota " & objNnota.Codigo & " Pedido " & objNnota.CodigoPedido)
                                    Exit Sub
                                End If

                                objNnota.IUD = "I"

                                Sql = " Insert Into NotasFiscais " & vbCrLf & _
                                          " (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
                                          "  Cliente_Id, EndCliente_Id, " & vbCrLf & _
                                          "  EntradaSaida_Id, Serie_Id, Nota_Id, TipoDeDocumento, " & vbCrLf & _
                                          "  Formulario, Pedido, Procuracao, " & vbCrLf & _
                                          "  Operacao, SubOperacao, Finalidade, " & vbCrLf & _
                                          "  Movimento, DataDaNota, NossaEmissao, " & vbCrLf & _
                                          "  SerieNotadoProdutor, NumeroNotadoProdutor, " & vbCrLf & _
                                          "  Deposito, EndDeposito, " & vbCrLf & _
                                          "  Destino, EndDestino, " & vbCrLf & _
                                          "  Transbordo, EndTransbordo, " & vbCrLf & _
                                          "  Agenciador, EndAgenciador, " & vbCrLf & _
                                          "  CIFFOB, Observacoes, Eletronica, " & vbCrLf & _
                                          "  Autorizacao," & vbCrLf & _
                                          "  UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf

                                If objNnota.UsuarioAlteracao.Length > 0 Then Sql &= "UsuarioAlteracao, UsuarioAlteracaoData, " & vbCrLf

                                Sql &= "  Situacao,ObservacoesDoProduto,ObservacoesDeEmbarque,ObservacoesControleInterno,EstadoDoCliente,LocalEmbarque,EndLocalEmbarque, " & vbCrLf & _
                                          "  ContratoANTT, ProtocoloANTT, CartaoPgtoFrete, DataTermino, idPamcard, TipoDeDocumentoFrete, Favorecido, EndFavorecido, NFG, Conferencia, UsuarioConferencia, UsuarioConferenciaData,Troca) " & vbCrLf & _
                                          "Values ('" & objNnota.CodigoEmpresa & "'," & objNnota.EnderecoEmpresa & ", " & vbCrLf & _
                                                                     "'" & objNnota.CodigoCliente & "'," & objNnota.EnderecoCliente & ", " & vbCrLf & _
                                                                     "'" & objNnota.EntradaSaida.ToString.Substring(0, 1) & "','" & objNnota.Serie & "'," & objNnota.Codigo & "," & objNnota.CodigoTipoDeDocumento & ", " & vbCrLf & _
                                                                     " " & objNnota.Formulario & ", " & objNnota.CodigoPedido.ToSqlNULL & ", " & objNnota.CodigoProcuracao.ToSqlNULL & ", " & vbCrLf & _
                                                                     " " & objNnota.CodigoOperacao & ", " & objNnota.CodigoSubOperacao & ", " & objNnota.CodigoFinalidade & ", " & vbCrLf & _
                                                                     "'" & objNnota.Movimento.ToString("yyyy-MM-dd") & "','" & objNnota.DataNota.ToString("yyyy-MM-dd") & "','" & IIf(objNnota.NossaEmissao, "S", "N") & "' , " & vbCrLf & _
                                                                     "'" & objNnota.SerieNotaProdutor & "', " & IIf(objNnota.NotaProdutor = 0, "Null", objNnota.NotaProdutor) & ", " & vbCrLf & _
                                                                     "'" & objNnota.CodigoDeposito & "', " & objNnota.EnderecoDeposito & ", " & vbCrLf & _
                                                                     "'" & objNnota.CodigoDestino & "', " & objNnota.EnderecoDestino & ", " & vbCrLf & _
                                                                     "'" & objNnota.CodigoTransbordo & "', " & objNnota.EnderecoTransbordo & ", " & vbCrLf & _
                                                                     "'" & objNnota.CodigoAgenciador & "', " & objNnota.EnderecoAgenciador & ", " & vbCrLf & _
                                                                     "'" & objNnota.CIFFOB.ToString & "', '" & objNnota.Observacoes & "', '" & IIf(objNnota.Eletronica, "S", "N") & "', " & vbCrLf & _
                                                                     " " & objNnota.CodigoAutorizacao & "," & vbCrLf & _
                                                                     "'" & objNnota.UsuarioInclusao & "','" & objNnota.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf

                                If objNnota.UsuarioAlteracao.Length > 0 Then Sql &= "'" & objNnota.UsuarioAlteracao & "','" & objNnota.DataAlteracao.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf

                                Sql &= "'" & objNnota.CodigoSituacao & "','" & objNnota.ObservacoesDoProduto & "','" & objNnota.ObservacoesDeEmbarque & "','" & objNnota.ObservacoesControleInterno & "'" & vbCrLf & _
                                          ",'" & objNnota.Cliente.CodigoEstado & "'" & vbCrLf & _
                                          "," & IIf(objNnota.CodigoLocalEmbarque.Length > 0, "'" & objNnota.CodigoLocalEmbarque & "'", "NULL") & vbCrLf & _
                                          "," & IIf(objNnota.CodigoLocalEmbarque.Length > 0, objNnota.EndLocalEmbarque, "NULL") & vbCrLf & _
                                          ",'" & objNnota.ContratoANTT & "','" & objNnota.ProtocoloANTT & "','" & objNnota.CartaoPgtoFrete & "','" & objNnota.DataTermino.ToString("yyyy-MM-dd") & "','" & objNnota.idPamcard & "', "

                                If objNnota.TipoDeDocumentoFrete IsNot Nothing Then Sql &= "'" & objNnota.TipoDeDocumentoFrete & "'," Else Sql &= "null,"
                                If Not String.IsNullOrWhiteSpace(objNnota.CodigoFavorecido) Then Sql &= "'" & objNnota.CodigoFavorecido & "'," & objNnota.EnderecoFavorecido & "," Else Sql &= "null,null,"

                                Sql &= IIf(objNnota.NFG, 1, 0) & "," & vbCrLf & _
                                          IIf(objNnota.Conferencia, 1, IIf(objNnota.Empresa.Empresa.ConferenciaNFE, 0, "NULL")) & "," & vbCrLf & _
                                          IIf(objNnota.Conferencia AndAlso Not String.IsNullOrWhiteSpace(objNnota.UsuarioConferencia), objNnota.UsuarioConferencia & "," & objNnota.UsuarioConferenciaData.ToString("yyyy-MM-dd HH:mm:ss"), "NULL,NULL") & vbCrLf & _
                                          "," & IIf(objNnota.Troca, 1, 0) & ")"

                                Sqls.Add(Sql)

                                strmPed.WriteLine("Transferiu Nota - " & objNnota.Codigo)

                                objNnota.Itens.SalvarSql(Sqls)

                                strmPed.WriteLine("Transferiu Itens")

                                Sql = "Insert Into NotasXEmbalagens (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Quantidade, Especie, Marca, Numero, PesoBruto, PesoLiquido)" & vbCrLf & _
                                      "Values('" & objNnota.CodigoEmpresa & "'," & objNnota.EnderecoEmpresa & ",'" & objNnota.CodigoCliente & "'," & objNnota.EnderecoCliente & ",'" & objNnota.EntradaSaida.ToString.Substring(0, 1) & "','" & objNnota.Serie & "'," & objNnota.Codigo & "," & Str(objNnota.Quantidade) & ",'" & objNnota.Especie & "','" & objNnota.Marca & "','" & objNnota.Numero & "'," & Str(objNnota.PesoBruto) & "," & Str(objNnota.PesoLiquido) & ")"
                                Sqls.Add(Sql)

                                strmPed.WriteLine("Transferiu Embalagens")

                                If Not objNnota.Transportador Is Nothing AndAlso objNnota.Transportador.Codigo.Length > 0 Then
                                    Sql = "Insert Into NotasFiscaisXTransportadores(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Proprietario, EndProprietario, Motorista, EndMotorista, Placa)" & vbCrLf & _
                                           "Values('" & objNnota.CodigoEmpresa & "'," & objNnota.EnderecoEmpresa & ",'" & objNnota.CodigoCliente & "'," & objNnota.EnderecoCliente & ",'" & objNnota.EntradaSaida.ToString.Substring(0, 1) & "','" & objNnota.Serie & "'," & objNnota.Codigo & ",'" & objNnota.CodigoTransportador & "'," & objNnota.EnderecoTransportador & ","
                                    If objNnota.PlacaDetalhes Is Nothing Then
                                        Sql &= "'',0,'');"
                                    Else
                                        Sql &= "'" & objNnota.PlacaDetalhes.Motorista.Codigo & "'," & objNnota.PlacaDetalhes.Motorista.CodigoEndereco & ",'" & objNnota.PlacaDetalhes.Placa01 & "');"
                                    End If
                                    Sqls.Add(Sql)
                                    strmPed.WriteLine("Transferiu Transportador")
                                End If

                                If objNnota.ChaveNFE.Length > 0 Then
                                    Sql = " Insert into NFERealizadas(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, ChaveNfe, ObservacoesFiscais, DadosAdicionais)" & vbCrLf & _
                                          " Values('" & objNnota.CodigoEmpresa & "'," & objNnota.EnderecoEmpresa & ",'" & objNnota.CodigoCliente & "'," & objNnota.EnderecoCliente & ",'" & objNnota.EntradaSaida.ToString.Substring(0, 1) & "','" & objNnota.Serie & "'," & objNnota.Codigo & ",'" & objNnota.Movimento.ToString("yyyy-MM-dd") & "','" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "','" & HttpContext.Current.Session("ssNomeUsuario") & "','" & objNnota.ChaveNFE & "','" & objNnota.Observacoes & " " & objNnota.ObservacoesDeEmbarque & "','" & objNnota.ObservacoesDoProduto & "');"
                                    Sqls.Add(Sql)
                                    strmPed.WriteLine("Transferiu NFERealizadas")
                                End If

                                If objNnota.ComissoesXBaixas IsNot Nothing AndAlso objNnota.ComissoesXBaixas.Count > 0 Then
                                    objNnota.ComissoesXBaixas.SalvarSql(Sqls)
                                    strmPed.WriteLine("Transferiu ComissoesXBaixas")
                                End If



                                If objNnota.TipoDeDocumentoFrete Is Nothing OrElse (objNnota.TipoDeDocumentoFrete IsNot Nothing AndAlso objNnota.TipoDeDocumentoFrete <> eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro) _
                                               OrElse (objNnota.TipoDeDocumentoFrete IsNot Nothing AndAlso objNnota.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro AndAlso objNnota.NotasTrocaOrigem IsNot Nothing AndAlso objNnota.NotasTrocaOrigem.Any(Function(s) s.TipoDeDocumentoFrete IsNot Nothing AndAlso s.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao)) Then
                                    If objNnota.NotasTrocaOrigem IsNot Nothing AndAlso objNnota.NotasTrocaOrigem.Count > 0 Then
                                        For Each NotaTrocaOrigem In objNnota.NotasTrocaOrigem
                                            NotaTrocaOrigem.SalvarNotasXNotas(Sqls)
                                            strmPed.WriteLine("Transferiu SalvarNotasXNotas")
                                        Next
                                    Else
                                        If objNnota.TemNotaTroca Then
                                            objNnota.SalvarNotasXNotas(Sqls)
                                            strmPed.WriteLine("Transferiu SalvarNotasXNotas")
                                        End If
                                    End If
                                End If

                                If objNnota.CodigoRomaneio > 0 AndAlso objNnota.Romaneio IsNot Nothing AndAlso objNnota.Romaneio.Codigo > 0 Then
                                    Sql = " Insert into NotasFiscaisXRomaneios(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Romaneio_Id)" & vbCrLf & _
                                          " Values('" & objNnota.CodigoEmpresa & "'," & objNnota.EnderecoEmpresa & ",'" & objNnota.CodigoCliente & "'," & objNnota.EnderecoCliente & ",'" & objNnota.EntradaSaida.ToString.Substring(0, 1) & "','" & objNnota.Serie & "'," & objNnota.Codigo & "," & objNnota.CodigoRomaneio & ");"
                                    Sqls.Add(Sql)
                                    strmPed.WriteLine("Transferiu NotasFiscaisXRomaneios")
                                End If
                            Next
                        End If


                        'SALVAR NOTA FISCAL X TITULO
                        Sql = "SELECT nXt.Empresa_Id, nXt.EndEmpresa_Id, nXt.Cliente_Id, nXt.EndCliente_Id, " & vbCrLf & _
                              "		nXt.EntradaSaida_Id, nXt.Serie_Id, nXt.Nota_Id, nXt.Titulo_Id " & vbCrLf & _
                              "FROM   NotaFiscalXTitulo nXt " & vbCrLf & _
                              "		inner join NotasFiscais n " & vbCrLf & _
                              "				on n.Empresa_id       = nXt.Empresa_id " & vbCrLf & _
                              "				and n.EndEmpresa_id   = nXt.EndEmpresa_id " & vbCrLf & _
                              "				and n.Cliente_id      = nXt.Cliente_id " & vbCrLf & _
                              "				and n.EndCliente_id   = nXt.EndCliente_id " & vbCrLf & _
                              "				and n.EntradaSaida_id = nXt.EntradaSaida_id " & vbCrLf & _
                              "				and n.Serie_id        = nXt.Serie_id " & vbCrLf & _
                              "				and n.Nota_id         = nXt.Nota_id " & vbCrLf & _
                              "WHERE n.Empresa_id    = '" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                              "and   n.EndEmpresa_id = " & objPedido.EnderecoEmpresa & vbCrLf & _
                              "and   n.Pedido        = " & objPedido.Codigo & vbCrLf

                        Dim dsNfxT As New DataSet
                        dsNfxT = Banco.ConsultaDataSet(Sql, "NotaFiscalXTitulo")

                        If dsNfxT IsNot Nothing AndAlso dsNfxT.Tables(0).Rows.Count > 0 Then
                            For Each drNfxT As DataRow In dsNfxT.Tables(0).Rows

                                Sql = "Insert into NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf & _
                                      " values('" & drNfxT("Empresa_id") & "'," & drNfxT("EndEmpresa_id") & ", " & vbCrLf & _
                                      "'" & drNfxT("Cliente_id") & "'," & drNfxT("EndCliente_id") & "," & vbCrLf & _
                                      "'" & drNfxT("EntradaSaida_id") & "','" & drNfxT("Serie_id") & "'," & drNfxT("Nota_id") & "," & drNfxT("Titulo_Id") & ")"
                                Sqls.Add(Sql)
                                strmPed.WriteLine("Transferiu NotaFiscalXTitulo - " & drNfxT("Nota_id") & " x " & drNfxT("Titulo_Id"))
                            Next
                        End If
                    End If

                    i += 1
                End While

                strmPed.Close()

                If Sqls.Count > 0 AndAlso GravaBancoDestino(Sqls, sConnStringDestino) Then
                    MsgBox(Me.Page, "Pedido transferido com Sucesso.", eTitulo.Sucess)
                    limpar()
                End If
            Else
                strmPed.Close()
                MsgBox(Me.Page, "Usuario sem permissão para transferir registro(s).")
            End If
        Catch ex As Exception
            strmPed.Close()
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "TransferePedidoInsol")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkAtualizaNotas_Click(sender As Object, e As EventArgs) Handles lnkAtualizaNotas.Click
        Try
            Dim Sqls As New ArrayList
            Dim Sql As String = String.Empty

            'SALVAR NOTA FISCAL
            Sql = " SELECT Empresa_id," & vbCrLf & _
                  "	       EndEmpresa_id," & vbCrLf & _
                  "	       Cliente_Id," & vbCrLf & _
                  "	       EndCliente_Id," & vbCrLf & _
                  "	       EntradaSaida_Id," & vbCrLf & _
                  "	       Serie_Id," & vbCrLf & _
                  "	       Nota_Id" & vbCrLf & _
                  " FROM NotasFiscais" & vbCrLf & _
                  " WHERE Situacao = 1 " & vbCrLf & _
                  "   AND Empresa_id      = '04440724001855'" & vbCrLf & _
                  "   And Cliente_Id      = '04997612000151'" & vbCrLf & _
                  "   And EntradaSaida_Id = 'E'" & vbCrLf & _
                  "   And Operacao        = 55 " & vbCrLf & _
                  "   And Nota_id         in(51801,51803,51804,51806,51843,51844,51849,51861,51881,51882,51886,51900,51901,51910,51911,51913,51939,51949,51965,51967,51968,51963,51964,52016,52017,52018,52019,52020,52021,52022,52031,52039,52040,52044,52047,52049,52051,52054,52055,52056,52088,52091,52096,52097,52098,52121,52124) " & vbCrLf & _
                  "   And Movimento Between '2016-05-01' and '2016-05-31'" & vbCrLf

            Dim dsN As New DataSet
            dsN = Banco.ConsultaDataSet(Sql, "NotasFiscais")

            If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                For Each drN As DataRow In dsN.Tables(0).Rows
                    Dim objNnota As New [Lib].Negocio.NotaFiscal()
                    objNnota.CodigoEmpresa = drN("Empresa_id")
                    objNnota.EnderecoEmpresa = drN("EndEmpresa_id")
                    objNnota.CodigoCliente = drN("Cliente_Id")
                    objNnota.EnderecoCliente = drN("EndCliente_Id")
                    objNnota.EntradaSaida = IIf(drN("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    objNnota.Serie = drN("Serie_Id")
                    objNnota.Codigo = drN("Nota_Id")
                    objNnota = New [Lib].Negocio.NotaFiscal(objNnota)

                    objNnota.IUD = "U"

                    ucFile.Clear()

                    Dim pNomeArquivo As String = String.Format("{0}-nfe", objNnota.ChaveNFE)

                    If System.IO.File.Exists(pNomeArquivo) Then System.IO.File.Delete(pNomeArquivo)

                    If objNnota.Arquivos IsNot Nothing AndAlso objNnota.Arquivos.Count > 0 Then
                        ucFile.Bind(objNnota.Arquivos)

                        ucFile.Baixar(objNnota.Arquivos(0).Codigo, objNnota.Arquivos(0).Descricao)

                        'Dim tempoLimite As DateTime
                        'tempoLimite = Now.AddSeconds(30)

                        'If Not System.IO.File.Exists(pNomeArquivo) AndAlso Now < tempoLimite Then
                        '    System.Threading.Thread.Sleep(3000)
                        'End If

                        'Dim DsXml As New DataSet
                        'DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}.xml", pNomeArquivo)))

                        'Dim teste As String = String.Empty

                        'If DsXml.Tables.Count > 0 Then
                        '    Dim obsEmb As String = objNnota.ObservacoesDeEmbarque & ". "

                        '    If DsXml.Tables("infAdic").Columns.Contains("infAdFisco") Then
                        '        obsEmb &= DsXml.Tables("infAdic").Rows(0)("infAdFisco") & ". "
                        '    End If

                        '    If DsXml.Tables("infAdic").Columns.Contains("infCpl") Then
                        '        obsEmb &= DsXml.Tables("infAdic").Rows(0)("infCpl") & "."
                        '    End If

                        '    txtObservacoesDeEmbarque.Text = obsEmb
                        'End If

                        System.Threading.Thread.Sleep(5000)
                    End If
                Next

                MsgBox(Me.Page, "Processo concluido.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    'Protected Sub GridTransferePedidoInsol_SelectedIndexChanged(sender As Object, e As EventArgs)
    '    Try
    '        txtPedido.Text = Server.HtmlDecode(GridTransferePedidoInsol.SelectedRow.Cells(1).Text())
    '        txtCliente.Text = Server.HtmlDecode(GridTransferePedidoInsol.SelectedRow.Cells(3).Text())

    '        lnkAtualizar.Parent.Visible = True
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message)
    '    End Try
    'End Sub
End Class