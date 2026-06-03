Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatoriosDeFretes
    Inherits BasePage

    Private objCliente As [Lib].Negocio.Cliente
    Private objPedido As [Lib].Negocio.Pedido

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RELATORIODEFRETES", "ACESSAR") Then
                BuscaEmpresa()
                BuscarGruposProdutos()
                txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
                txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Sub CarregarCliente()
        If Not Session("objClienteFRT" & HID.Value) Is Nothing Then
            objCliente = CType(Session("objClienteFRT" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteFRT" & HID.Value)
        End If

        If Not Session("objClienteFRT2" & HID.Value) Is Nothing Then
            objCliente = CType(Session("objClienteFRT2" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            TxtDeposito.Text = itemCliente.Text
            TxtCodigoDeposito.Value = itemCliente.Value
            Session.Remove("objClienteFRT2" & HID.Value)
        End If

        If Not Session("objClienteFRT3" & HID.Value) Is Nothing Then
            objCliente = CType(Session("objClienteFRT3" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtTransportador.Text = itemCliente.Text
            txtCodigoTransportador.Value = itemCliente.Value
            Session.Remove("objClienteFRT3" & HID.Value)
        End If
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, "", True)
        Funcoes.VerificaEmpresa(cmbEmpresa)
    End Sub

    Private Sub BuscarGruposProdutos()
        ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub BuscarProduto()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & ddlGrupoProduto.SelectedValue & "'", True)
    End Sub

    Protected Sub cmdConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteFRT" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnDeposito_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteFRT2" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnTransportador_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteFRT3" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddlProduto.Items.Clear()
            BuscarProduto()
            ddlProduto.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub CkFretesAPagar_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If CkFretesAPagar.Checked = True Then
                CkPosicaoDeFretes.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub CkPosicaoDeFretes_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If CkPosicaoDeFretes.Checked = True Then
                CkFretesAPagar.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function ValidarCampos() As Boolean
        If txtDataInicial.Text.Length = 0 Or IsDate(txtDataInicial.Text) = False Then
            MsgBox(Me.Page, "Informe uma data inicial válida")
            txtDataInicial.Focus()
            Return False
        End If

        If txtDataFinal.Text.Length = 0 Or IsDate(txtDataFinal.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida")
            txtDataFinal.Focus()
            Return False
        End If

        If CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final")
            txtDataInicial.Focus()
            Return False
        End If

        Return True
    End Function

    Public Sub BuscarRegistros()
        Dim sqlProd As String = ""
        Dim DescricaoProduto As String = ""
        Dim dsProduto As New DataSet

        Dim ds As New DataSet
        Dim SQL As String
        Dim Cliente As String = ""
        Dim strEmpresa() As String = cmbEmpresa.SelectedValue.Split("-")
        Dim strDeposito() As String = TxtCodigoDeposito.Value.Split("-")
        Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
        Dim strTransportador() As String = txtCodigoTransportador.Value.Split("-")
        Dim strProduto() As String = ddlProduto.SelectedValue.ToString.Split("_")

        If ddlProduto.SelectedIndex > 0 Then
            sqlProd &= "SELECT * FROM PRODUTOS WHERE PRODUTO_ID = '" & strProduto(0) & "'"
            dsProduto = Banco.ConsultaDataSet(sqlProd, "Produto")
            DescricaoProduto = " - PRODUTO: " & Trim(dsProduto.Tables(0).Rows(0).Item("Nome")).ToString
        ElseIf ddlGrupoProduto.SelectedIndex > 0 Then
            sqlProd &= "SELECT * FROM GRUPOSDEESTOQUES WHERE GRUPO_ID = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "'"
            dsProduto = Banco.ConsultaDataSet(sqlProd, "Produto")
            DescricaoProduto = " - GRUPO: " & Trim(dsProduto.Tables(0).Rows(0).Item("Descricao")).ToString
        End If


        '*** SQL LIVRE -----------------------------------------------------------------------------------------------------------------------------
        If CkFretesAPagar.Checked = False And CkPosicaoDeFretes.Checked = False Then
            SQL = " SELECT NotasFiscais.nota_Id as Recibo_Id, " & vbCrLf & _
                    " NotasFiscais.Movimento,  " & vbCrLf & _
                    " CONVERT(VarChar, NotasFiscais.Nota_Id) + '-' + SUBSTRING(CONVERT(Varchar,  NotasFiscais.Serie_Id), 1, 1) AS Romaneio, " & vbCrLf & _
                    " SQ1.Produto_Id,  " & vbCrLf & _
                    " Produtos.Nome,  " & vbCrLf & _
                    " CONVERT(Varchar, Deposito.Reduzido)  + ' - ' + CONVERT(Varchar, Deposito.Fantasia) + ' - ' + CONVERT(Varchar, Deposito.Endereco) AS Deposito,  " & vbCrLf & _
                    " CONVERT(Varchar, Destino.Reduzido)  + ' - ' + CONVERT(Varchar, Destino.Fantasia) + ' - ' + CONVERT(Varchar, Destino.Endereco) AS Destino,  " & vbCrLf & _
                    " NotasFiscaisXTransportadores.Placa,  " & vbCrLf & _
                    " NotasFiscaisXTransportadores.Motorista,  " & vbCrLf & _
                    " SQ1.Peso,   " & vbCrLf & _
                    " SQ1.Valor as PrecoOficial,  " & vbCrLf & _
                    " SQ1.Valor as ValorOficial,  " & vbCrLf & _
                    " CONVERT(Varchar,  Transportador.Reduzido) + ' - ' + CONVERT(Varchar, Transportador.Fantasia) AS Transportador, " & vbCrLf & _
                    " isnull(SQ1.Valor,0) AS Adiantamento  " & vbCrLf & _
            " FROM  " & vbCrLf & _
            "   (SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,  " & vbCrLf & _
                        " NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, " & vbCrLf & _
                        " (SELECT top 1 NotasFiscaisXItens.produto_Id from NotasFiscaisXItens where " & vbCrLf & _
                        "         NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id   " & vbCrLf & _
                                " AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf & _
                                " AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                            " ) as Produto_id, " & vbCrLf & _
                        " (SELECT top 1 NotasFiscaisXItens.deposito from NotasFiscaisXItens  " & vbCrLf & _
                                " where NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                                " AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id   " & vbCrLf & _
                                " AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf & _
                                " AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                            " ) as deposito, " & vbCrLf & _
                        " (SELECT top 1 NotasFiscaisXItens.enddeposito from NotasFiscaisXItens  " & vbCrLf & _
                                " where NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND  " & vbCrLf & _
                                " NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                            " ) as enddeposito, " & vbCrLf & _
                        " (SELECT top 1 NotasFiscaisXItens.depositoterceiro from NotasFiscaisXItens where " & vbCrLf & _
                                " NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND  " & vbCrLf & _
                                " NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                            " ) as depositoterceiro, " & vbCrLf & _
                        " (SELECT top 1 NotasFiscaisXItens.enddepositoterceiro from NotasFiscaisXItens  " & vbCrLf & _
                                " where NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND  " & vbCrLf & _
                                " NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                            " ) as enddepositoterceiro ," & vbCrLf & _
                        " (SELECT Sum(NotasFiscaisXItens.PesoFiscal) from NotasFiscaisXItens where " & vbCrLf & _
                                " NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id   " & vbCrLf & _
                                " AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf & _
                                " AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                            " ) as Peso," & vbCrLf & _
                        " (SELECT Sum(NotasFiscaisXItens.Valor) from NotasFiscaisXItens where " & vbCrLf & _
                                " NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND  " & vbCrLf & _
                                " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id   " & vbCrLf & _
                                " AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id  " & vbCrLf & _
                                " AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                            " ) as Valor" & vbCrLf & _
               " FROM NotasFiscais  " & vbCrLf & _
               " GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,  " & vbCrLf & _
                        " NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id " & vbCrLf & _
             " ) SQ1  " & vbCrLf & _
             " INNER JOIN NotasFiscais  " & vbCrLf & _
                " ON  NotasFiscais.Empresa_Id       = SQ1.Empresa_Id  " & vbCrLf & _
                " AND NotasFiscais.EndEmpresa_Id    = SQ1.EndEmpresa_Id  " & vbCrLf & _
                " AND NotasFiscais.Cliente_Id       = SQ1.Cliente_Id  " & vbCrLf & _
                " AND NotasFiscais.EndCliente_Id    = SQ1.EndCliente_Id  " & vbCrLf & _
                " AND NotasFiscais.EntradaSaida_Id  = SQ1.EntradaSaida_Id  " & vbCrLf & _
                " AND NotasFiscais.Serie_Id         = SQ1.Serie_Id  " & vbCrLf & _
                " AND NotasFiscais.Nota_Id          = SQ1.Nota_Id  " & vbCrLf & _
            " INNER JOIN NotasFiscaisXTransportadores" & vbCrLf & _
                " ON  NotasFiscais.Empresa_Id  = NotasFiscaisXTransportadores.Empresa_Id       " & vbCrLf & _
                " AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXTransportadores.EndEmpresa_Id     " & vbCrLf & _
                " AND NotasFiscais.Cliente_Id = NotasFiscaisXTransportadores.Cliente_Id         " & vbCrLf & _
                " AND NotasFiscais.EndCliente_Id = NotasFiscaisXTransportadores.EndCliente_Id      " & vbCrLf & _
                " AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXTransportadores.EntradaSaida_Id    " & vbCrLf & _
                " AND NotasFiscais.Serie_Id  = NotasFiscaisXTransportadores.Serie_Id         " & vbCrLf & _
                " AND NotasFiscais.Nota_Id = NotasFiscaisXTransportadores.Nota_Id          " & vbCrLf & _
             " INNER JOIN Clientes AS Destino  " & vbCrLf & _
                " ON  NotasFiscais.Cliente_Id     = Destino.Cliente_Id  " & vbCrLf & _
                " AND NotasFiscais.EndCliente_Id  = Destino.Endereco_Id  " & vbCrLf & _
             " INNER JOIN Clientes AS Transportador  " & vbCrLf & _
                " ON  NotasFiscaisXTransportadores.Proprietario    = Transportador.Cliente_Id  " & vbCrLf & _
                " AND NotasFiscaisXTransportadores.EndProprietario = Transportador.Endereco_Id  " & vbCrLf & _
" left JOIN  NotasFiscaisXEncargos  " & vbCrLf & _
" ON NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id  " & vbCrLf & _
" AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id " & vbCrLf & _
" AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id  " & vbCrLf & _
" AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id " & vbCrLf & _
" AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id  " & vbCrLf & _
" AND NotasFiscais.Serie_Id = NotasFiscaisXEncargos.Serie_Id " & vbCrLf & _
" And NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id " & vbCrLf & _
                " AND NotasFiscaisXEncargos.Produto_Id = '502020001'  " & vbCrLf & _
                " AND NotasFiscaisXEncargos.Encargo_Id = 'ADIANTAMENTO'  " & vbCrLf & _
             " INNER JOIN Produtos  " & vbCrLf & _
                " ON SQ1.Produto_Id = Produtos.Produto_Id  " & vbCrLf & _
             " Left JOIN Clientes AS Deposito  " & vbCrLf & _
             "    ON  Deposito.Cliente_Id  = ISNULL(SQ1.Deposito,SQ1.Depositoterceiro)  " & vbCrLf & _
             "  AND Deposito.Endereco_Id = ISNULL(SQ1.EndDeposito,SQ1.EndDepositoterceiro)  " & vbCrLf & _
             " Where NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf
            '" AND NotasFiscais.Serie_Id in ('REC','UN')" & vbCrLf

            If strEmpresa(0).Length > 0 Then
                SQL &= " AND NotasFiscais.Empresa_Id= '" & strEmpresa(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscais.EndEmpresa_Id= " & strEmpresa(1) & " " & vbCrLf
            End If

            If strDeposito(0).Length > 0 Then
                SQL &= " AND SQ1.Deposito = '" & strDeposito(0) & "'" & vbCrLf
                SQL &= " AND SQ1.EndDeposito = " & strDeposito(1) & " " & vbCrLf
            End If

            If strCliente(0).Length > 0 Then
                SQL &= " AND NotasFiscais.Cliente_Nf= '" & strCliente(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscais.EndCliente_Nf= " & strCliente(1) & " " & vbCrLf
            End If

            If strTransportador(0).Length > 0 Then
                SQL &= " AND NotasFiscaisXTransportadores.Proprietario= '" & strTransportador(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscaisXTransportadores.EndProprietario= " & strTransportador(1) & " " & vbCrLf
            End If

            If ddlGrupoProduto.SelectedIndex > 0 Then SQL &= " AND (left(SQ1.Produto_Id,5) = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "') " & vbCrLf
            If ddlProduto.SelectedIndex > 0 Then SQL &= " AND And SQ1.Produto_Id = '" & strProduto(0) & "' " & vbCrLf

            SQL &= " Order by NotasFiscais.Movimento,Deposito.Reduzido,NotasFiscais.Formulario " & vbCrLf

            If CkDolar.Checked Then
                SQL = Replace(SQL, "#1#", "Moeda")
            Else
                SQL = Replace(SQL, "#1#", "Oficial")
            End If

            ds = Banco.ConsultaDataSet(SQL, "RecibosDeFretes")
            AlimentaCrptRelatorios(ds, "~/Reports/CrRelatorioDeFretes2", CkSaldoApagar.Checked)

            '*** SQL FRETES A PAGAR --------------------------------------------------------------------------------------------------------------------
        ElseIf CkFretesAPagar.Checked = True And CkPosicaoDeFretes.Checked = False Then
            '            SQL = "SELECT NotasFiscais.Movimento,   " & vbCrLf & _
            '"        NotasFiscais.Nota_Id AS Recibo_Id,   " & vbCrLf & _
            '"        CONVERT(VarChar, NotasFiscais.Nota_Id)  + '-' + SUBSTRING(CONVERT(Varchar, NotasFiscais.Serie_Id), 1, 1) AS Romaneio,   " & vbCrLf & _
            '"        SQ1.Produto_Id,   " & vbCrLf & _
            '"        Produtos.Nome,   " & vbCrLf & _
            '"        CONVERT(Varchar, Deposito.Reduzido)  AS Deposito,   " & vbCrLf & _
            '"        CONVERT(Varchar, Destino.Reduzido)  AS Destino,   " & vbCrLf & _
            '"        NotasFiscaisXTransportadores.Placa,   " & vbCrLf & _
            '"        CONVERT(Varchar, Transportador.Reduzido) + ' - ' + CONVERT(Varchar, Transportador.Fantasia) AS Transportador,   " & vbCrLf & _
            '"        NotasFiscaisXTransportadores.Motorista,   " & vbCrLf & _
            '"        SQ1.Peso AS PesosSaida,   " & vbCrLf & _
            '"        isnull(FaturasdefretesxItens.PesoDeChegada,0) AS PesosChegada,   " & vbCrLf & _
            '"        SQ1.Valor as VlrOrigem,   " & vbCrLf & _
            '"        ISNULL((SELECT SUM(rfxe.Valor) AS Valor   " & vbCrLf & _
            '"				 FROM NotasFiscaisxEncargos AS rfxe   " & vbCrLf & _
            '"               INNER JOIN NotasFiscais AS txrf   " & vbCrLf & _
            '"                    ON rfxe.Empresa_Id    = txrf.Empresa_Id   " & vbCrLf & _
            '"                 AND rfxe.EndEmpresa_Id   = txrf.EndEmpresa_Id   " & vbCrLf & _
            '"                 AND rfxe.Cliente_Id      = txrf.Cliente_Id    " & vbCrLf & _
            '"                 AND rfxe.EndCliente_Id   = txrf.EndCliente_Id  " & vbCrLf & _
            '"                 AND rfxe.EntradaSaida_Id = txrf.EntradaSaida_Id   " & vbCrLf & _
            '"                 AND rfxe.Serie_Id        = txrf.Serie_Id  " & vbCrLf & _
            '"                 AND rfxe.Nota_Id         = txrf.Nota_Id   " & vbCrLf & _
            '"                 WHERE rfxe.Encargo_Id    = 'LIQUIDOAPAGAR'   " & vbCrLf & _
            '"                  AND txrf.Movimento    <= '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "') , 0) AS VlrChegada, " & vbCrLf & _
            '"       ISNULL((SELECT SUM(rfxe.Valor) AS Valor   " & vbCrLf & _
            '"				 FROM NotasFiscaisxEncargos AS rfxe   " & vbCrLf & _
            '"                INNER JOIN NotasFiscais AS txrf   " & vbCrLf & _
            '"                    ON rfxe.Empresa_Id    = txrf.Empresa_Id   " & vbCrLf & _
            '"                 AND rfxe.EndEmpresa_Id   = txrf.EndEmpresa_Id   " & vbCrLf & _
            '"                 AND rfxe.Cliente_Id      = txrf.Cliente_Id    " & vbCrLf & _
            '"                 AND rfxe.EndCliente_Id   = txrf.EndCliente_Id  " & vbCrLf & _
            '"                 AND rfxe.EntradaSaida_Id = txrf.EntradaSaida_Id   " & vbCrLf & _
            '"                 AND rfxe.Serie_Id        = txrf.Serie_Id  " & vbCrLf & _
            '"                 AND rfxe.Nota_Id         = txrf.Nota_Id   " & vbCrLf & _
            '"                 WHERE rfxe.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id   " & vbCrLf & _
            '"                  AND rfxe.Nota_Id     = NotasFiscais.Nota_Id   " & vbCrLf & _
            '"                  AND rfxe.Encargo_Id    = 'ADIANTAMENTO'   " & vbCrLf & _
            '"                  AND txrf.Movimento    <= '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "') , 0) AS Adiantamento, " & vbCrLf & _
            '"         ISNULL(( SELECT SUM(rfxe.Valor) AS Valor   " & vbCrLf & _
            '"                  FROM NotasFiscaisxEncargos AS rfxe   " & vbCrLf & _
            '"                 INNER JOIN NotasFiscais AS txrf   " & vbCrLf & _
            '"                    ON rfxe.Empresa_Id    = txrf.Empresa_Id   " & vbCrLf & _
            '"                 AND rfxe.EndEmpresa_Id   = txrf.EndEmpresa_Id   " & vbCrLf & _
            '"                 AND rfxe.Cliente_Id      = txrf.Cliente_Id    " & vbCrLf & _
            '"                 AND rfxe.EndCliente_Id   = txrf.EndCliente_Id  " & vbCrLf & _
            '"                 AND rfxe.EntradaSaida_Id = txrf.EntradaSaida_Id   " & vbCrLf & _
            '"                 AND rfxe.Serie_Id        = txrf.Serie_Id  " & vbCrLf & _
            '"                 AND rfxe.Nota_Id         = txrf.Nota_Id   " & vbCrLf & _
            '"                 WHERE rfxe.Encargo_Id    = 'BAIXAADTO'   " & vbCrLf & _
            '"                   AND txrf.Movimento     <= '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "') , 0)   " & vbCrLf & _
            '"        +  " & vbCrLf & _
            '"        ISNULL((SELECT SUM(rfxe.Valor) AS Valor   " & vbCrLf & _
            '"                  FROM NotasFiscaisxEncargos AS rfxe   " & vbCrLf & _
            '"                 INNER JOIN NotasFiscais AS txrf   " & vbCrLf & _
            '"                    ON rfxe.Empresa_Id    = txrf.Empresa_Id   " & vbCrLf & _
            '"                 AND rfxe.EndEmpresa_Id   = txrf.EndEmpresa_Id   " & vbCrLf & _
            '"                 AND rfxe.Cliente_Id      = txrf.Cliente_Id    " & vbCrLf & _
            '"                 AND rfxe.EndCliente_Id   = txrf.EndCliente_Id  " & vbCrLf & _
            '"                 AND rfxe.EntradaSaida_Id = txrf.EntradaSaida_Id   " & vbCrLf & _
            '"                 AND rfxe.Serie_Id        = txrf.Serie_Id  " & vbCrLf & _
            '"                 AND rfxe.Nota_Id         = txrf.Nota_Id   " & vbCrLf & _
            '"                 WHERE  rfxe.Encargo_Id    = 'AMORTIZAADTO'   " & vbCrLf & _
            '"                   AND txrf.Movimento   <= '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'), 0)  AS AdiantamentoProgPgo " & vbCrLf & _
            '"   FROM   " & vbCrLf & _
            '"(SELECT  NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,   " & vbCrLf & _
            '"         NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id,   " & vbCrLf & _
            '"		(SELECT top 1 NotasFiscaisXItens.produto_Id  " & vbCrLf & _
            '"           from NotasFiscaisXItens   " & vbCrLf & _
            '"		  where	NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id   " & vbCrLf & _
            '"            AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id   " & vbCrLf & _
            '"            AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id    " & vbCrLf & _
            '"            AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf & _
            '"            AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id   " & vbCrLf & _
            '"            AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf & _
            '"            AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id   " & vbCrLf & _
            '"		 ) as Produto_id,   " & vbCrLf & _
            '"		(SELECT top 1 NotasFiscaisXItens.deposito  " & vbCrLf & _
            '"           from NotasFiscaisXItens   " & vbCrLf & _
            '"		  where NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id  " & vbCrLf & _
            '"            AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf & _
            '"            AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id    " & vbCrLf & _
            '"            AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id   " & vbCrLf & _
            '"            AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id   " & vbCrLf & _
            '"            AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf & _
            '"            AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id   " & vbCrLf & _
            '"		 ) as deposito,   " & vbCrLf & _
            '"		(SELECT top 1 NotasFiscaisXItens.enddeposito  " & vbCrLf & _
            '"           from NotasFiscaisXItens  " & vbCrLf & _
            '"		  where	NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id   " & vbCrLf & _
            '"            AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf & _
            '"            AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id   " & vbCrLf & _
            '"            AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id   " & vbCrLf & _
            '"            AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id   " & vbCrLf & _
            '"            AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id   " & vbCrLf & _
            '"            AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id   " & vbCrLf & _
            '"		) as enddeposito,   " & vbCrLf & _
            '"		(SELECT top 1 NotasFiscaisXItens.depositoterceiro  " & vbCrLf & _
            '"           from NotasFiscaisXItens   " & vbCrLf & _
            '"		  where	NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id  " & vbCrLf & _
            '"            AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf & _
            '"            AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id   " & vbCrLf & _
            '"            AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf & _
            '"            AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id   " & vbCrLf & _
            '"            AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id  " & vbCrLf & _
            '"            AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id   " & vbCrLf & _
            '"		) as depositoterceiro,   " & vbCrLf & _
            '"		(SELECT top 1 NotasFiscaisXItens.enddepositoterceiro  " & vbCrLf & _
            '"           from NotasFiscaisXItens   " & vbCrLf & _
            '"		  where NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id  " & vbCrLf & _
            '"            AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf & _
            '"            AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id    " & vbCrLf & _
            '"            AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id   " & vbCrLf & _
            '"            AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id   " & vbCrLf & _
            '"            AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id   " & vbCrLf & _
            '"            AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id   " & vbCrLf & _
            '"		) as enddepositoterceiro,  " & vbCrLf & _
            '"         (SELECT Sum(NotasFiscaisXItens.PesoFiscal) from NotasFiscaisXItens    " & vbCrLf & _
            '"           where NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND    " & vbCrLf & _
            '"                    NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id     " & vbCrLf & _
            '"                    AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id    " & vbCrLf & _
            '"                    AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id   " & vbCrLf & _
            '"                ) as Peso,  " & vbCrLf & _
            '"          (SELECT Sum(NotasFiscaisXItens.Valor) from NotasFiscaisXItens    " & vbCrLf & _
            '"            where NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND    " & vbCrLf & _
            '"                    NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id     " & vbCrLf & _
            '"                    AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id    " & vbCrLf & _
            '"                    AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id And NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id   " & vbCrLf & _
            '"                ) as Valor   " & vbCrLf & _
            '"  FROM NotasFiscais   " & vbCrLf & _
            '" GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,   " & vbCrLf & _
            '"          NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id   " & vbCrLf & _
            '")SQ1   " & vbCrLf & _
            '" INNER JOIN NotasFiscais   " & vbCrLf & _
            '"   ON NotasFiscais.Empresa_Id       = SQ1.Empresa_Id   " & vbCrLf & _
            '"  AND NotasFiscais.EndEmpresa_Id    = SQ1.EndEmpresa_Id   " & vbCrLf & _
            '"  AND NotasFiscais.Cliente_Id       = SQ1.Cliente_Id   " & vbCrLf & _
            '"  AND NotasFiscais.EndCliente_Id    = SQ1.EndCliente_Id   " & vbCrLf & _
            '"  AND NotasFiscais.EntradaSaida_Id  = SQ1.EntradaSaida_Id   " & vbCrLf & _
            '"  AND NotasFiscais.Serie_Id         = SQ1.Serie_Id   " & vbCrLf & _
            '"  AND NotasFiscais.Nota_Id          = SQ1.Nota_Id   " & vbCrLf & _
            '"LEFT JOIN FaturasdefretesxItens   " & vbCrLf & _
            '"   ON FaturasdefretesxItens.Empresa_Id       = NotasFiscais.Empresa_Id   " & vbCrLf & _
            '"  AND FaturasdefretesxItens.EndEmpresa_Id    = NotasFiscais.EndEmpresa_Id   " & vbCrLf & _
            '"  AND FaturasdefretesxItens.Cliente_Id       = NotasFiscais.Cliente_Id   " & vbCrLf & _
            '"  AND FaturasdefretesxItens.EndCliente_Id    = NotasFiscais.EndCliente_Id   " & vbCrLf & _
            '"  AND FaturasdefretesxItens.EntradaSaida_Id  = NotasFiscais.EntradaSaida_Id   " & vbCrLf & _
            '"  AND FaturasdefretesxItens.Serie_Id         = NotasFiscais.Serie_Id   " & vbCrLf & _
            '"  AND FaturasdefretesxItens.Nota_Id          = NotasFiscais.Nota_Id   " & vbCrLf & _
            '"  AND FaturasdefretesxItens.Encargo_Id          = 'LIQUIDOAPAGAR'  " & vbCrLf & _
            '" INNER JOIN Produtos   " & vbCrLf & _
            '"   ON SQ1.Produto_Id = Produtos.Produto_Id   " & vbCrLf & _
            '" INNER JOIN Clientes AS Deposito   " & vbCrLf & _
            '"   ON Deposito.Cliente_Id  = ISNULL(SQ1.Deposito,SQ1.Depositoterceiro)   " & vbCrLf & _
            '"  AND Deposito.Endereco_Id = ISNULL(SQ1.EndDeposito,SQ1.EndDepositoterceiro)   " & vbCrLf & _
            '" INNER JOIN Clientes AS Destino   " & vbCrLf & _
            '"   ON NotasFiscais.Cliente_Id     = Destino.Cliente_Id   " & vbCrLf & _
            '"  AND NotasFiscais.EndEmpresa_Id  = Destino.Endereco_Id   " & vbCrLf & _
            '"INNER JOIN NotasFiscaisXTransportadores  " & vbCrLf & _
            '"    ON  NotasFiscais.Empresa_Id  = NotasFiscaisXTransportadores.Empresa_Id         " & vbCrLf & _
            '"    AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXTransportadores.EndEmpresa_Id       " & vbCrLf & _
            '"    AND NotasFiscais.Cliente_Id = NotasFiscaisXTransportadores.Cliente_Id           " & vbCrLf & _
            '"    AND NotasFiscais.EndCliente_Id = NotasFiscaisXTransportadores.EndCliente_Id        " & vbCrLf & _
            '"    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXTransportadores.EntradaSaida_Id      " & vbCrLf & _
            '"    AND NotasFiscais.Serie_Id  = NotasFiscaisXTransportadores.Serie_Id           " & vbCrLf & _
            '"    AND NotasFiscais.Nota_Id = NotasFiscaisXTransportadores.Nota_Id   " & vbCrLf & _
            '" INNER JOIN Clientes AS Transportador    " & vbCrLf & _
            '"    ON  NotasFiscaisXTransportadores.Proprietario    = Transportador.Cliente_Id    " & vbCrLf & _
            '"    AND NotasFiscaisXTransportadores.EndProprietario = Transportador.Endereco_Id   " & vbCrLf & _
            '" LEFT OUTER JOIN NotasFiscaisXEncargos   " & vbCrLf & _
            '" ON NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id  " & vbCrLf & _
            '" AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id " & vbCrLf & _
            '" AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id  " & vbCrLf & _
            '" AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id " & vbCrLf & _
            '" AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id  " & vbCrLf & _
            '" AND NotasFiscais.Serie_Id = NotasFiscaisXEncargos.Serie_Id " & vbCrLf & _
            '" And NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id " & vbCrLf & _
            '"  AND NotasFiscaisXEncargos.Produto_Id = '502020001'   " & vbCrLf & _
            '"  AND NotasFiscaisXEncargos.Encargo_Id = 'ADIANTAMENTO'   " & vbCrLf & _
            '" Where NotasFiscais.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "' " & vbCrLf

            'If strEmpresa(0).Length > 0 Then
            '    SQL &= " AND NotasFiscais.Empresa_Id= '" & strEmpresa(0) & "'" & vbCrLf
            '    SQL &= " AND NotasFiscais.EndEmpresa_Id= " & strEmpresa(1) & " " & vbCrLf
            'End If

            'If strDeposito(0).Length > 0 Then
            '    SQL &= " AND SQ1.Deposito = '" & strDeposito(0) & "'" & vbCrLf
            '    SQL &= " AND SQ1.EndDeposito = " & strDeposito(1) & " " & vbCrLf
            'End If

            'If strCliente(0).Length > 0 Then
            '    SQL &= " AND NotasFiscais.Cliente_Id= '" & strCliente(0) & "'" & vbCrLf
            '    SQL &= " AND NotasFiscais.EndCliente_Id= " & strCliente(1) & " " & vbCrLf
            'End If

            'If strTransportador(0).Length > 0 Then
            '    SQL &= " AND NotasFiscaisXTransportadores.Proprietario= '" & strTransportador(0) & "'" & vbCrLf
            '    SQL &= " AND NotasFiscaisXTransportadores.EndProprietario= " & strTransportador(1) & " " & vbCrLf
            'End If

            'If ddlGrupoProduto.SelectedIndex > 0 Then SQL &= " AND (left(SQ1.Produto_Id,5) = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "') " & vbCrLf
            'If ddlProduto.SelectedIndex > 0 Then SQL &= " AND And SQ1.Produto_Id = '" & strProduto(0) & "' " & vbCrLf


            'SQL &= " Order by NotasFiscais.Movimento,Deposito.Reduzido,NotasFiscais.Formulario  " & vbCrLf

            'If CkDolar.Checked Then
            '    SQL = Replace(SQL, "#1#", "Moeda") & vbCrLf
            'Else
            '    SQL = Replace(SQL, "#1#", "Oficial") & vbCrLf
            'End If

            SQL = "SELECT  NotasFiscais.Movimento,     " & vbCrLf & _
         " NotasFiscais.Nota_Id AS Recibo_Id,   " & vbCrLf & _
         " isnull(Romaneios.Romaneio_Id,'') AS Romaneio,     " & vbCrLf & _
         " NotasFiscaisxItens.Produto_Id, " & vbCrLf & _
         " Produtos.Nome, " & vbCrLf & _
         " Deposito.Reduzido  AS Deposito,     " & vbCrLf & _
         " Destino.Reduzido  AS Destino,     " & vbCrLf & _
         " NotasFiscaisXTransportadores.Placa, " & vbCrLf & _
         " Transportador.Reduzido + ' - ' + Transportador.Fantasia AS Transportador,    " & vbCrLf & _
         " NotasFiscaisXTransportadores.Motorista, " & vbCrLf & _
         " NotasFiscaisxItens.PesoFiscal AS PesosSaida,  " & vbCrLf & _
         " Isnull(case when  NotasFiscais.EntradaSaida_Id='E' then   " & vbCrLf & _
             " isnull(FaturasdefretesxItens.PesoDeChegada, Romaneios.PesoBruto)  " & vbCrLf & _
         " Else  " & vbCrLf & _
             " isnull(FaturasdefretesxItens.PesoDeChegada, Romaneios.PesoBruto)  " & vbCrLf & _
         " end,0) AS PesosChegada,   " & vbCrLf & _
         " isnull(sum(case when NotasFiscaisXEncargos.Encargo_Id='PRODUTO' then   " & vbCrLf & _
         " NotasFiscaisXEncargos.valor " & vbCrLf & _
         " end),0) as VlrOrigem,  " & vbCrLf & _
        " isnull(sum(case when NotasFiscaisXEncargos.Encargo_Id='LIQUIDOAPAGAR' then  " & vbCrLf & _
             " NotasFiscaisXEncargos.valor " & vbCrLf & _
        " end),0) as VlrChegada,  " & vbCrLf & _
        " isnull(sum(case when NotasFiscaisXEncargos.Encargo_Id='ADIANTAMENTO' then  " & vbCrLf & _
             " NotasFiscaisXEncargos.valor " & vbCrLf & _
        " end),0) as Adiantamento,  " & vbCrLf & _
        " isnull(sum(case when NotasFiscaisXEncargos.Encargo_Id='BAIXAADTO' then  " & vbCrLf & _
             " NotasFiscaisXEncargos.valor " & vbCrLf & _
        " end),0)  " & vbCrLf & _
      " +  " & vbCrLf & _
        " isnull(sum(case when NotasFiscaisXEncargos.Encargo_Id='AMORTIZAADTO' then  " & vbCrLf & _
             " NotasFiscaisXEncargos.valor " & vbCrLf & _
        " end),0) AS AdiantamentoProgPgo   " & vbCrLf & _
             " FROM NotasFiscais  " & vbCrLf & _
 " LEFT JOIN NotasXNotas   " & vbCrLf & _
 " ON NotasFiscais.Empresa_Id = NotasXNotas.Empresa_Id   " & vbCrLf & _
 " AND NotasFiscais.EndEmpresa_Id = NotasXNotas.EndEmpresa_Id   " & vbCrLf & _
 " AND NotasFiscais.Cliente_Id = NotasXNotas.Cliente_Id   " & vbCrLf & _
 " AND NotasFiscais.EndCliente_Id = NotasXNotas.EndCliente_Id   " & vbCrLf & _
 " AND NotasFiscais.EntradaSaida_Id = NotasXNotas.EntradaSaida_Id   " & vbCrLf & _
 " AND NotasFiscais.Serie_Id = NotasXNotas.Serie_Id   " & vbCrLf & _
 " AND NotasFiscais.Nota_Id = NotasXNotas.Nota_Id   " & vbCrLf & _
 " LEFT JOIN  NotasFiscaisXRomaneios nfxr  " & vbCrLf & _
 " ON NotasXNotas.OrigemEmpresa_Id = nfxr.Empresa_Id   " & vbCrLf & _
 " AND NotasXNotas.OrigemEndEmpresa_Id = nfxr.EndEmpresa_Id   " & vbCrLf & _
 " AND NotasXNotas.OrigemCliente_Id = nfxr.Cliente_Id   " & vbCrLf & _
 " AND NotasXNotas.OrigemEndCliente_Id = nfxr.EndCliente_Id   " & vbCrLf & _
 " AND NotasXNotas.OrigemEntradaSaida_Id = nfxr.EntradaSaida_Id   " & vbCrLf & _
 " AND NotasXNotas.OrigemSerie_Id = nfxr.Serie_Id   " & vbCrLf & _
 " AND NotasXNotas.OrigemNota_Id = nfxr.Nota_Id   " & vbCrLf & _
 " left JOIN Romaneios   " & vbCrLf & _
 " ON nfxr.Empresa_Id = Romaneios.Empresa_Id   " & vbCrLf & _
 " AND nfxr.EndEmpresa_Id = Romaneios.EndEmpresa_Id   " & vbCrLf & _
 " AND nfxr.Romaneio_Id = Romaneios.Romaneio_Id  " & vbCrLf & _
 " LEFT OUTER JOIN NotasFiscaisXEncargos     " & vbCrLf & _
  " ON NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id    " & vbCrLf & _
  " AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
  " AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id    " & vbCrLf & _
  " AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id   " & vbCrLf & _
  " AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id   " & vbCrLf & _
  " AND NotasFiscais.Serie_Id = NotasFiscaisXEncargos.Serie_Id   " & vbCrLf & _
  " And NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id  " & vbCrLf & _
  " LEFT  JOIN NotasFiscaisXItens    " & vbCrLf & _
              " ON	NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id    " & vbCrLf & _
             " AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf & _
             " AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id     " & vbCrLf & _
             " AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id   " & vbCrLf & _
             " AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id     " & vbCrLf & _
             " AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id    " & vbCrLf & _
             " AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id    " & vbCrLf & _
 " LEFT JOIN  Produtos   " & vbCrLf & _
 " ON Produtos.Produto_Id = NotasFiscaisXItens.Produto_Id    " & vbCrLf & _
  " LEFT  JOIN NotasFiscaisXItens AS NXIOrigem   " & vbCrLf & _
              " ON	NotasXNotas.OrigemEmpresa_Id      = NXIOrigem.Empresa_Id    " & vbCrLf & _
             " AND NotasXNotas.OrigemEndEmpresa_Id   = NXIOrigem.EndEmpresa_Id  " & vbCrLf & _
             " AND NotasXNotas.OrigemCliente_Id      = NXIOrigem.Cliente_Id     " & vbCrLf & _
             " AND NotasXNotas.OrigemEndCliente_Id   = NXIOrigem.EndCliente_Id   " & vbCrLf & _
             " AND NotasXNotas.OrigemEntradaSaida_Id = NXIOrigem.EntradaSaida_Id     " & vbCrLf & _
             " AND NotasXNotas.OrigemSerie_Id        = NXIOrigem.Serie_Id    " & vbCrLf & _
             " AND NotasXNotas.OrigemNota_Id         = NXIOrigem.Nota_Id    " & vbCrLf & _
 " LEFT JOIN  Produtos  AS ProdOrigem " & vbCrLf & _
 " ON ProdOrigem.Produto_Id = NXIOrigem.Produto_Id    " & vbCrLf & _
 " LEFT JOIN NotasFiscaisXTransportadores    " & vbCrLf & _
     " ON  NotasFiscais.Empresa_Id  = NotasFiscaisXTransportadores.Empresa_Id   " & vbCrLf & _
     " AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXTransportadores.EndEmpresa_Id    " & vbCrLf & _
     " AND NotasFiscais.Cliente_Id = NotasFiscaisXTransportadores.Cliente_Id     " & vbCrLf & _
     " AND NotasFiscais.EndCliente_Id = NotasFiscaisXTransportadores.EndCliente_Id     " & vbCrLf & _
     " AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXTransportadores.EntradaSaida_Id    " & vbCrLf & _
     " AND NotasFiscais.Serie_Id  = NotasFiscaisXTransportadores.Serie_Id      " & vbCrLf & _
     " AND NotasFiscais.Nota_Id = NotasFiscaisXTransportadores.Nota_Id     " & vbCrLf & _
  " LEFT JOIN Clientes AS Transportador      " & vbCrLf & _
 "     ON  NotasFiscaisXTransportadores.Proprietario    = Transportador.Cliente_Id   " & vbCrLf & _
     " AND NotasFiscaisXTransportadores.EndProprietario = Transportador.Endereco_Id   " & vbCrLf & _
  " LEFT JOIN Clientes AS Destino     " & vbCrLf & _
    " ON NotasFiscais.Cliente_Id     = Destino.Cliente_Id   " & vbCrLf & _
   " AND NotasFiscais.EndEmpresa_Id  = Destino.Endereco_Id   " & vbCrLf & _
 " LEFT JOIN Clientes AS Deposito     " & vbCrLf & _
 "    ON  ISNULL(NotasFiscaisXItens.Deposito,NotasFiscaisXItens.DepositoTerceiro) = Deposito.Cliente_Id     " & vbCrLf & _
   " AND ISNULL(NotasFiscaisXItens.EndDeposito,NotasFiscaisXItens.EndDepositoTerceiro)  = Deposito.Endereco_Id   " & vbCrLf & _
 " LEFT JOIN FaturasdefretesxItens     " & vbCrLf & _
 "    ON FaturasdefretesxItens.Empresa_Id       = NotasFiscais.Empresa_Id     " & vbCrLf & _
   " AND FaturasdefretesxItens.EndEmpresa_Id    = NotasFiscais.EndEmpresa_Id   " & vbCrLf & _
   " AND FaturasdefretesxItens.Cliente_Id       = NotasFiscais.Cliente_Id     " & vbCrLf & _
   " AND FaturasdefretesxItens.EndCliente_Id    = NotasFiscais.EndCliente_Id    " & vbCrLf & _
   " AND FaturasdefretesxItens.EntradaSaida_Id  = NotasFiscais.EntradaSaida_Id    " & vbCrLf & _
   " AND FaturasdefretesxItens.Serie_Id         = NotasFiscais.Serie_Id     " & vbCrLf & _
   " AND FaturasdefretesxItens.Nota_Id          = NotasFiscais.Nota_Id   " & vbCrLf & _
   " where NotasFiscaisXitens.Produto_Id = '502020001'   " & vbCrLf & _
   " AND  NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf
            If strEmpresa(0).Length > 0 Then
                SQL &= " AND NotasFiscais.Empresa_Id= '" & strEmpresa(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscais.EndEmpresa_Id= " & strEmpresa(1) & " " & vbCrLf
            End If

            If strDeposito(0).Length > 0 Then
                SQL &= " AND NotasFiscaisXItens.Deposito = '" & strDeposito(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscaisXItens.EndDeposito = " & strDeposito(1) & " " & vbCrLf
            End If

            If strCliente(0).Length > 0 Then
                SQL &= " AND NotasFiscais.Cliente_Id= '" & strCliente(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscais.EndCliente_Id= " & strCliente(1) & " " & vbCrLf
            End If

            If strTransportador(0).Length > 0 Then
                SQL &= " AND NotasFiscaisXTransportadores.Proprietario= '" & strTransportador(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscaisXTransportadores.EndProprietario= " & strTransportador(1) & " " & vbCrLf
            End If

            If ddlGrupoProduto.SelectedIndex > 0 Then SQL &= " AND (left(NXIOrigem.Produto_Id,5) = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "') " & vbCrLf
            If ddlProduto.SelectedIndex > 0 Then SQL &= " AND And NXIOrigem.Produto_Id = '" & strProduto(0) & "' " & vbCrLf



            SQL &= " Group By NotasFiscais.Movimento, NotasFiscais.Nota_Id,Romaneios.Romaneio_Id,NotasFiscaisxItens.Produto_Id, Produtos.Nome,    " & vbCrLf & _
                  " Deposito.Reduzido, Destino.Reduzido, NotasFiscaisXTransportadores.Placa, Transportador.Reduzido, Transportador.Fantasia, " & vbCrLf & _
                  " NotasFiscaisXTransportadores.Motorista ,NotasFiscaisxItens.PesoFiscal,FaturasdefretesxItens.PesoDeChegada, Romaneios.PesoBruto,NotasFiscais.EntradaSaida_Id " & vbCrLf

            SQL &= " Order by NotasFiscais.Movimento,Deposito.Reduzido,NotasFiscais.Nota_Id   " & vbCrLf

            ds = Banco.ConsultaDataSet(SQL, "RecibosDeFretes")
            ds.Merge(Banco.ConsultaDataSet(BuscaSQlRazaoFrete, "RazaoFretes")) 'precisa de uma conta no plano de contas para adiantamento de fretes
            AlimentaCrptRelatorios(ds, "~/Reports/CrRelatorioDeFretesPago", CkSaldoApagar.Checked)


            '*** SQL POSIÇÃO DE FRETES ------------------------------------------------------------------------------------------------
        ElseIf CkPosicaoDeFretes.Checked = True Then
            SQL = "Select   " & vbCrLf & _
            " NotasFiscais.Empresa_Id,  " & vbCrLf & _
            " NotasFiscais.EndEmpresa_Id,  " & vbCrLf & _
            " NotasFiscais.Nota_Id AS Recibo_Id,  " & vbCrLf & _
            " NotasFiscaisXTransportadores.Proprietario as Transportador,   " & vbCrLf & _
            " NotasFiscaisXTransportadores.EndProprietario as EndTransportador,   " & vbCrLf & _
            "  (Select Sum(NotasFiscaisXEncargos.Valor) as RTRC from NotasFiscaisXEncargos    " & vbCrLf & _
            "  where  " & vbCrLf & _
            "  NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id    " & vbCrLf & _
            "  AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
            "  AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id           " & vbCrLf & _
            "  AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id        " & vbCrLf & _
            "  AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id      " & vbCrLf & _
            "  AND NotasFiscais.Serie_Id  = NotasFiscaisXEncargos.Serie_Id         " & vbCrLf & _
            "  AND NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id   " & vbCrLf & _
            "  AND NotasFiscaisXEncargos.Encargo_Id = 'RTRC ') as RTRC,   " & vbCrLf & _
            "  (Select Sum(NotasFiscaisXEncargos.Valor) as ADIANTAMENTO from NotasFiscaisXEncargos   " & vbCrLf & _
            "  Where   " & vbCrLf & _
            "   NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id    " & vbCrLf & _
            "  AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
            "  AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id           " & vbCrLf & _
            "  AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id        " & vbCrLf & _
            "  AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id      " & vbCrLf & _
            "  AND NotasFiscais.Serie_Id  = NotasFiscaisXEncargos.Serie_Id         " & vbCrLf & _
            "  AND NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id   " & vbCrLf & _
            "  AND NotasFiscaisXEncargos.Encargo_Id = 'ADIANTAMENTO ') as ADIANTAMENTO ,   " & vbCrLf & _
            "  isnull((SELECT  Sum(NotasFiscaisXEncargos.Valor) as  BAIXAADIANTAMENTO  " & vbCrLf & _
            "  FROM NotasFiscaisXEncargos   " & vbCrLf & _
            " INNER JOIN NotasFiscais   " & vbCrLf & _
            " ON NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id   " & vbCrLf & _
            " AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
            " AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id   " & vbCrLf & _
            " AND NotasFiscais.Serie_Id = NotasFiscaisXEncargos.Serie_Id   " & vbCrLf & _
            " AND NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id   " & vbCrLf & _
            " INNER JOIN FaturasDeFretesXItens   " & vbCrLf & _
            " ON NotasFiscais.Empresa_Id = FaturasDeFretesXItens.Empresa_Id   " & vbCrLf & _
            " AND NotasFiscais.EndEmpresa_Id = FaturasDeFretesXItens.EndEmpresa_Id   " & vbCrLf & _
            " AND NotasFiscais.Cliente_Id = FaturasDeFretesXItens.Cliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EndCliente_Id = FaturasDeFretesXItens.EndCliente_Id   " & vbCrLf & _
            " AND NotasFiscais.Serie_Id = FaturasDeFretesXItens.Serie_Id   " & vbCrLf & _
            " AND NotasFiscais.EntradaSaida_Id = FaturasDeFretesXItens.EntradaSaida_Id   " & vbCrLf & _
            " AND NotasFiscais.Nota_Id = FaturasDeFretesXItens.Nota_Id  " & vbCrLf & _
            " INNER JOIN  FaturasDeFretes   " & vbCrLf & _
            " ON FaturasDeFretesXItens.EmpresaPagadora_Id = FaturasDeFretes.Empresa_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.EndEmpresaPagadora_Id = FaturasDeFretes.EndEmpresa_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.Conveniado_Id = FaturasDeFretes.Conveniado_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.EndConveniado_Id = FaturasDeFretes.EndConveniado_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.Fatura_Id = FaturasDeFretes.Fatura_Id  " & vbCrLf & _
            " INNER JOIN ContasAPagar ON FaturasDeFretes.Titulo = ContasAPagar.Registro_Id  " & vbCrLf & _
            "  Where  " & vbCrLf & _
            "   NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id    " & vbCrLf & _
            "  AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
            "  AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id           " & vbCrLf & _
            "  AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id        " & vbCrLf & _
            "  AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id      " & vbCrLf & _
            "  AND NotasFiscais.Serie_Id  = NotasFiscaisXEncargos.Serie_Id         " & vbCrLf & _
            "  AND NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id   " & vbCrLf & _
            "  AND NotasFiscais.Movimento  <= '" & txtDataFinal.Text.ToSqlDate() & "'    " & vbCrLf & _
            "  AND NotasFiscaisXEncargos.Encargo_Id = 'BAIXAADTO'),0) as BAIXAADIANTAMENTO ,  " & vbCrLf & _
            "  isnull((SELECT  Sum(NotasFiscaisXEncargos.Valor) as  AMOADIANTAMENTO  " & vbCrLf & _
            "  FROM NotasFiscaisXEncargos   " & vbCrLf & _
            " INNER JOIN NotasFiscais   " & vbCrLf & _
            " ON NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id   " & vbCrLf & _
            " AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
            " AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id   " & vbCrLf & _
            " AND NotasFiscais.Serie_Id = NotasFiscaisXEncargos.Serie_Id   " & vbCrLf & _
            " AND NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id   " & vbCrLf & _
            " INNER JOIN FaturasDeFretesXItens   " & vbCrLf & _
            " ON NotasFiscais.Empresa_Id = FaturasDeFretesXItens.Empresa_Id   " & vbCrLf & _
            " AND NotasFiscais.EndEmpresa_Id = FaturasDeFretesXItens.EndEmpresa_Id   " & vbCrLf & _
            " AND NotasFiscais.Cliente_Id = FaturasDeFretesXItens.Cliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EndCliente_Id = FaturasDeFretesXItens.EndCliente_Id   " & vbCrLf & _
            " AND NotasFiscais.Serie_Id = FaturasDeFretesXItens.Serie_Id   " & vbCrLf & _
            " AND NotasFiscais.EntradaSaida_Id = FaturasDeFretesXItens.EntradaSaida_Id   " & vbCrLf & _
            " AND NotasFiscais.Nota_Id = FaturasDeFretesXItens.Nota_Id  " & vbCrLf & _
            " INNER JOIN  FaturasDeFretes   " & vbCrLf & _
            " ON FaturasDeFretesXItens.EmpresaPagadora_Id = FaturasDeFretes.Empresa_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.EndEmpresaPagadora_Id = FaturasDeFretes.EndEmpresa_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.Conveniado_Id = FaturasDeFretes.Conveniado_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.EndConveniado_Id = FaturasDeFretes.EndConveniado_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.Fatura_Id = FaturasDeFretes.Fatura_Id  " & vbCrLf & _
            " INNER JOIN ContasAPagar ON FaturasDeFretes.Titulo = ContasAPagar.Registro_Id  " & vbCrLf & _
            "  Where  " & vbCrLf & _
            "   NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id    " & vbCrLf & _
            "  AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
            "  AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id           " & vbCrLf & _
            "  AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id        " & vbCrLf & _
            "  AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id      " & vbCrLf & _
            "  AND NotasFiscais.Serie_Id  = NotasFiscaisXEncargos.Serie_Id         " & vbCrLf & _
            "  AND NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id   " & vbCrLf & _
            "  AND NotasFiscais.Movimento  <= '" & txtDataFinal.Text.ToSqlDate() & "'    " & vbCrLf & _
            "  AND NotasFiscaisXEncargos.Encargo_Id = 'AMORTIZAADTO'),0) as AMOADIANTAMENTO ,  " & vbCrLf & _
            "  isnull((SELECT  Sum(NotasFiscaisXEncargos.Valor) as  LIQUIDOAPAGAR  " & vbCrLf & _
            "  FROM NotasFiscaisXEncargos   " & vbCrLf & _
            " INNER JOIN NotasFiscais   " & vbCrLf & _
            " ON NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id   " & vbCrLf & _
            " AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
            " AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id   " & vbCrLf & _
            " AND NotasFiscais.Serie_Id = NotasFiscaisXEncargos.Serie_Id   " & vbCrLf & _
            " AND NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id   " & vbCrLf & _
            " LEFT JOIN FaturasDeFretesXItens   " & vbCrLf & _
            " ON NotasFiscais.Empresa_Id = FaturasDeFretesXItens.Empresa_Id   " & vbCrLf & _
            " AND NotasFiscais.EndEmpresa_Id = FaturasDeFretesXItens.EndEmpresa_Id   " & vbCrLf & _
            " AND NotasFiscais.Cliente_Id = FaturasDeFretesXItens.Cliente_Id   " & vbCrLf & _
            " AND NotasFiscais.EndCliente_Id = FaturasDeFretesXItens.EndCliente_Id   " & vbCrLf & _
            " AND NotasFiscais.Serie_Id = FaturasDeFretesXItens.Serie_Id   " & vbCrLf & _
            " AND NotasFiscais.EntradaSaida_Id = FaturasDeFretesXItens.EntradaSaida_Id   " & vbCrLf & _
            " AND NotasFiscais.Nota_Id = FaturasDeFretesXItens.Nota_Id  " & vbCrLf & _
            " LEFT JOIN  FaturasDeFretes   " & vbCrLf & _
            " ON FaturasDeFretesXItens.EmpresaPagadora_Id = FaturasDeFretes.Empresa_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.EndEmpresaPagadora_Id = FaturasDeFretes.EndEmpresa_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.Conveniado_Id = FaturasDeFretes.Conveniado_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.EndConveniado_Id = FaturasDeFretes.EndConveniado_Id   " & vbCrLf & _
            " AND FaturasDeFretesXItens.Fatura_Id = FaturasDeFretes.Fatura_Id  " & vbCrLf & _
            " INNER JOIN ContasAPagar ON FaturasDeFretes.Titulo = ContasAPagar.Registro_Id  " & vbCrLf & _
            "  Where  " & vbCrLf & _
            "   NotasFiscais.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id    " & vbCrLf & _
            "  AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id   " & vbCrLf & _
            "  AND NotasFiscais.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id           " & vbCrLf & _
            "  AND NotasFiscais.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id        " & vbCrLf & _
            "  AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id      " & vbCrLf & _
            "  AND NotasFiscais.Serie_Id  = NotasFiscaisXEncargos.Serie_Id         " & vbCrLf & _
            "  AND NotasFiscais.Nota_Id = NotasFiscaisXEncargos.Nota_Id   " & vbCrLf & _
            "  AND NotasFiscais.Movimento  <= '" & txtDataFinal.Text.ToSqlDate() & "'   " & vbCrLf & _
            "  AND NotasFiscaisXEncargos.Encargo_Id = 'LIQUIDOAPAGAR'),0) as LIQUIDOAPAGAR   " & vbCrLf & _
            " Into #RecibosDeFretes  " & vbCrLf & _
            "  from NotasFiscais   " & vbCrLf & _
            "  INNER JOIN NotasFiscaisXItens   " & vbCrLf & _
            "     ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id   " & vbCrLf & _
            "     AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id    " & vbCrLf & _
            "     AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id   " & vbCrLf & _
            "     AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id    " & vbCrLf & _
            "     AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id   " & vbCrLf & _
            "     AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id    " & vbCrLf & _
            "     AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id   " & vbCrLf & _
            " INNER JOIN NotasFiscaisXTransportadores  " & vbCrLf & _
            "     ON  NotasFiscais.Empresa_Id  = NotasFiscaisXTransportadores.Empresa_Id         " & vbCrLf & _
            "     AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXTransportadores.EndEmpresa_Id       " & vbCrLf & _
            "     AND NotasFiscais.Cliente_Id = NotasFiscaisXTransportadores.Cliente_Id           " & vbCrLf & _
            "     AND NotasFiscais.EndCliente_Id = NotasFiscaisXTransportadores.EndCliente_Id        " & vbCrLf & _
            "     AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXTransportadores.EntradaSaida_Id      " & vbCrLf & _
            "     AND NotasFiscais.Serie_Id  = NotasFiscaisXTransportadores.Serie_Id           " & vbCrLf & _
            "     AND NotasFiscais.Nota_Id = NotasFiscaisXTransportadores.Nota_Id     " & vbCrLf & _
            "  Where NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
            "  AND NotasFiscais.Serie_Id in ('REC','UN') " & vbCrLf

            If strEmpresa(0).Length > 0 Then
                SQL &= " AND NotasFiscais.Empresa_Id= '" & strEmpresa(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscais.EndEmpresa_Id= " & strEmpresa(1) & " " & vbCrLf
            End If

            If strDeposito(0).Length > 0 Then
                SQL &= " AND NotasFiscaisXItens.Deposito = '" & strDeposito(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscaisXItens.EndDeposito = " & strDeposito(1) & " " & vbCrLf
            End If

            If strCliente(0).Length > 0 Then
                SQL &= " AND NotasFiscais.Cliente_Id= '" & strCliente(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscais.EndCliente_Id= " & strCliente(1) & " " & vbCrLf
            End If

            If strTransportador(0).Length > 0 Then
                SQL &= " AND NotasFiscaisXTransportadores.Proprietario= '" & strTransportador(0) & "'" & vbCrLf
                SQL &= " AND NotasFiscaisXTransportadores.EndProprietario= " & strTransportador(1) & " " & vbCrLf
            End If

            If ddlGrupoProduto.SelectedIndex > 0 Then SQL &= " AND (left(NotasFiscaisXItens.Produto_Id,5) = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "') " & vbCrLf
            If ddlProduto.SelectedIndex > 0 Then SQL &= " AND And NotasFiscaisXItens.Produto_Id = '" & strProduto(0) & "' " & vbCrLf


            SQL &= "; SELECT Empresa_Id,EndEmpresa_Id,Recibo_Id,Transportador,EndTransportador,RTRC,ADIANTAMENTO,BAIXAADIANTAMENTO,AMOADIANTAMENTO,LIQUIDOAPAGAR," & vbCrLf & _
             " SALDOADIANTAMENTO = " & vbCrLf & _
             " Case BAIXAADIANTAMENTO + AMOADIANTAMENTO" & vbCrLf & _
             " when 0 then ADIANTAMENTO" & vbCrLf & _
             " ELSE" & vbCrLf & _
             " 0.00" & vbCrLf & _
             " end," & vbCrLf & _
             " SALDOCARTAFRETE =" & vbCrLf & _
             " Case LIQUIDOAPAGAR" & vbCrLf & _
             " when 0 then RTRC - ADIANTAMENTO" & vbCrLf & _
             " ELSE" & vbCrLf & _
             " 0.00" & vbCrLf & _
             " End" & vbCrLf & _
             " into FretesAux#" & vbCrLf & _
             " from #RecibosDeFretes " & vbCrLf & _
             "; select Transportador,EndTransportador, sum(RTRC) as RTRC,Sum(ADIANTAMENTO) as ADIANTAMENTO , Sum (BAIXAADIANTAMENTO) as BAIXAADIANTAMENTO,Sum(AMOADIANTAMENTO) as AMOADIANTAMENTO, Sum(LIQUIDOAPAGAR) as LIQUIDOAPAGAR,Sum(SALDOCARTAFRETE) as SALDOCARTAFRETE ,Sum(SALDOADIANTAMENTO) as SALDOADIANTAMENTO " & vbCrLf & _
             " into #RecibosDeFretesAux" & vbCrLf & _
             " from FretesAux# group by Transportador,EndTransportador;" & vbCrLf & _
             " drop table FretesAux#;" & vbCrLf & _
             " SELECT     Transportador, EndTransportador, Trasnp.Nome, Trasnp.Reduzido, SUM(RTRC) AS RTRC, SUM(ADIANTAMENTO) AS ADIANTAMENTO, " & vbCrLf & _
             " SUM(BAIXAADIANTAMENTO) AS BAIXAADIANTAMENTO, SUM(AMOADIANTAMENTO) AS AMOADIANTAMENTO, SUM(LIQUIDOAPAGAR) " & vbCrLf & _
             " AS LIQUIDOAPAGAR, SUM(SALDOCARTAFRETE) AS SALDOCARTAFRETE, SUM(SALDOADIANTAMENTO) AS SALDOADIANTAMENTO, SUM(0.00) " & vbCrLf & _
             " AS Credito, SUM(0.00) AS Debito" & vbCrLf & _
             " into #Teste" & vbCrLf & _
             " FROM  [#RecibosDeFretesAux] INNER JOIN" & vbCrLf & _
             " Clientes AS Trasnp ON Transportador = Trasnp.Cliente_Id AND EndTransportador = Trasnp.Endereco_Id" & vbCrLf & _
             " GROUP BY Transportador, EndTransportador, Trasnp.Nome, Trasnp.Reduzido;" & vbCrLf & _
             " insert Into #Teste SELECT  Razao.Cliente_Id as Transportador, Razao.EndCliente_Id as EndTransportador, Clientes.Nome,Clientes.Reduzido,0.00 as RTRC,0.00 as ADIANTAMENTO,0.00 as BAIXAADIANTAMENTO,0.00 as AMOADIANTAMENTO,0.00 as LIQUIDOAPAGAR,0.00 as SALDOCARTADRETE,0.00 as SALDOADIANTAMENTO, SUM(Razao.Credito#1#) AS Credito, SUM(Razao.Debito#1#) AS Debito" & vbCrLf & _
             " FROM Razao INNER JOIN" & vbCrLf & _
             " Clientes ON Razao.Cliente_Id = Clientes.Cliente_Id AND Razao.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
             " WHERE Razao.Conta_Id = '11304'" & vbCrLf & _
             " AND razao.movimento_id <= '" & txtDataFinal.Text.ToSqlDate() & "'  " & vbCrLf 'Adicionado por Edson 21/07/08

            If strTransportador(0).Length > 0 Then
                SQL &= " AND razao.Cliente_Id= '" & strTransportador(0) & "'" & vbCrLf
                SQL &= " AND razao.EndCliente_Id= " & strTransportador(1) & " " & vbCrLf
            End If

            SQL &= " GROUP BY Razao.Cliente_Id, Razao.EndCliente_Id, Clientes.Nome, Clientes.Reduzido;  " & vbCrLf & _
             " select Transportador,EndTransportador,Nome,Reduzido,Sum(RTRC) as RTRC ,Sum(ADIANTAMENTO) as ADIANTAMENTO,Sum(BAIXAADIANTAMENTO) as BAIXAADIANTAMENTO,Sum(AMOADIANTAMENTO) as AMOADIANTAMENTO ,Sum(LIQUIDOAPAGAR) as LIQUIDOAPAGAR,sum(SALDOCARTAFRETE) as SALDOCARTAFRETE ,Sum(SALDOADIANTAMENTO) as SALDOADIANTAMENTO,Sum(Credito) as Credito,Sum(Debito) as Debito from #Teste " & vbCrLf & _
             " group by Transportador,EndTransportador,Nome,Reduzido order by Nome" & vbCrLf

            If CkDolar.Checked Then
                SQL = Replace(SQL, "#1#", "Moeda")
            Else
                SQL = Replace(SQL, "#1#", "Oficial")
            End If

            ds = Banco.ConsultaDataSet(SQL, "RecibosDeFretesAux")
            AlimentaCrptRelatorios(ds, "~/Reports/CrPosicaoDeFretes", CkSaldoApagar.Checked)
        End If
    End Sub

    Function BuscaSQlRazaoFrete() As String
        Dim sql2 As String = ""
        Dim Cliente As String = ""
        Dim strEmpresa() As String = cmbEmpresa.SelectedValue.Split("-")
        Dim strDeposito() As String = TxtCodigoDeposito.Value.Split("-")
        Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
        Dim strTransportador() As String = txtCodigoTransportador.Value.Split("-")


        sql2 = " SELECT ISNULL(SUM(Debito#1# - Credito#1#), 0) AS SaldoRazao " & vbCrLf & _
         "  FROM Razao " & vbCrLf & _
         " WHERE  Movimento_Id <= '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
         "        AND Conta_Id      = 11304 " & vbCrLf

        If strTransportador(0).Length > 0 Then
            sql2 &= " AND Cliente_Id= '" & strTransportador(0) & "'" & vbCrLf
            sql2 &= " AND EndCliente_Id= " & strTransportador(1) & " " & vbCrLf
        End If

        If CkDolar.Checked Then
            sql2 = Replace(sql2, "#1#", "Moeda")
        Else
            sql2 = Replace(sql2, "#1#", "Oficial")
        End If

        Return sql2
    End Function

    Public Sub AlimentaCrptRelatorios(ByVal Ds As DataSet, ByVal Caminho As String, ByVal SaldoZerado As Boolean)
        Dim crptRelatorio As New ReportDocument()

        Try
            crptRelatorio.FileName = Server.MapPath(Caminho & ".rpt")
            crptRelatorio.Load(crptRelatorio.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
            Dim arquivo As String = NomeArquivo

            crptRelatorio.SetDataSource(Ds)

            Dim crParameterValues As CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As CrystalDecisions.Shared.ParameterDiscreteValue
            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

            crParameterFieldDefinitions = crptRelatorio.DataDefinition.ParameterFields()
            If Caminho = "~/Reports/CrRelatorioDeFretes2" Then

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Nome")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = HttpContext.Current.Session("ssNomeEmpresa")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Cidade")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Zerado")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = SaldoZerado
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Titulo")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Relatório de Fretes - Periodo " & txtDataInicial.Text & " a " & txtDataFinal.Text & " "
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Pagina")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Página"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Data")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Emissão"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Numero")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Numero"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Romaneio")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Romaneio"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Produto")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Produto"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Origem")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Origem"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Destino")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Destino"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Placa")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Placa"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Transportador")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Transportador"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Motorista")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Motorista"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Quantidade")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Quantidade"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Valor/Ton")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Valor p/Ton."
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("ValorFrete")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Valor Frete"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Adiantamento")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Adiantamento"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Saldo")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Saldo"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            ElseIf Caminho = "~/Reports/CrRelatorioDeFretesPago" Then

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("NomeEmpresa")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = HttpContext.Current.Session("ssNomeEmpresa")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Cidade")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Zerado")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = SaldoZerado
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Titulo")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Relatório de Fretes a Pagar - Periodo  " & txtDataInicial.Text & " a " & txtDataFinal.Text & ""
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Pagina")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Página"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Data")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Emissão"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Recibo")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Numero"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Romaneio")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Romaneio"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Nome")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Produto"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Deposito")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Origem"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Destino")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Destino"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Placa")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Placa"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Transportador")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Transportador"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Motorista")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Motorista"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("PesosChegada")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Chegada"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("PesosSaida")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Saida"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("VlrOrigem")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Origem"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("VlrChegada")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Chegada"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Adiantamento")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Adiantamento"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("SaldoLiquido")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Saldo.Liquido"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("LiquidoProgramado")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Liq.Prog/Pgo"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("AdiantamentoProgramado")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Adto.Prog./Pgo"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("SaldoAPagar")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Saldo a Pagar"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            ElseIf Caminho = "~/Reports/CrPosicaoDeFretes" Then

                Dim cotacao As Decimal
                Dim ds_cotacao As DataSet
                Dim sql As String
                sql = "select indice from Cotacoes where Indexador_Id = 3 And Data_Id ='" & Format(Now, "yyyy/MM/dd") & "'"
                ds_cotacao = Banco.ConsultaDataSet(sql, "Cotacoes")
                cotacao = ds_cotacao.Tables(0).Rows(0)(0)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Nome")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = HttpContext.Current.Session("ssNomeEmpresa")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Cidade")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Zerado")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = SaldoZerado
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Titulo")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "POSIÇÃO DE FRETES - Periodo " & txtDataInicial.Text & " a " & txtDataFinal.Text & ""
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Cotacao")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = cotacao
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


                crParameterFieldDefinition = crParameterFieldDefinitions.Item("Transportador")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Transportador"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("CartaFrete")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Carta Frete"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("ViaAdiantamento")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Via Adto"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("FreteTotal")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Frete Total"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("AdtoEmpresa")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Adto Empresa"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                crParameterFieldDefinition = crParameterFieldDefinitions.Item("SaldoAPagar")
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
                crParameterDiscreteValue.Value = "Saldo A Pagar"
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
            End If

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crptRelatorio.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

            If IO.File.Exists(arquivo) Then
                ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crptRelatorio.Close()
            crptRelatorio.Dispose()
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If ValidarCampos() Then
                BuscarRegistros()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatoriosDeFretes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class