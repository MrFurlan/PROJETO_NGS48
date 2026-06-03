Imports System.Data
Imports System.IO
Imports System.Web
Imports System.Net
Imports System.Net.Mail
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Runtime.CompilerServices
Imports System.Globalization

Public Class CarteiraDeCobranca
    Inherits BasePage

    Private Sql As String
    Private Empresa() As String
    Private strJavaScript As String
    Private strSQL As String

    Dim dsCarteira As DataSet
    Dim dvCarteira As DataView
    Dim arr As New ArrayList
    Dim Instrucao01Bradesco As BoletoNet.Instrucao_Bradesco
    Dim Instrucao02Bradesco As BoletoNet.Instrucao_Bradesco
    Dim HashOcorrenciaRetorno As New Hashtable
    Dim MoraD As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CarteiraDeCobranca", "ACESSAR") Then
                CargaEmpresas()
                Funcoes.VerificaEmpresa(ddlEmpresa)
                CargaCarteiras()
                CargaSituacao()
                Limpar_Campos()
                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                CarregarTitulosEnviadosAoBanco()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If

        SalvarSessaoCarregaObjetos()

        If IsPostBack And fup.HasFile Then
            Dim strCaminho As String = Server.MapPath("Files/") & Path.GetFileName(fup.FileName)
            If Dir(strCaminho).Length > 0 Then Kill(strCaminho)

            'Salvamos o mesmo 
            lblArquivoRetorno.Text = fup.FileName
            divRetorno.Visible = True

            fup.PostedFile.SaveAs(strCaminho)
            CarregarRetornoRemessa()
            TabContainer1.ActiveTabIndex = 1
            fup.Visible = False
        End If

    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteCARCOB" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteCARCOB" & HID.Value), Cliente))
            txtNomeCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteCARCOB" & HID.Value)
        End If
    End Sub

    Private Sub SalvarSessaoCarregaObjetos()
        If Not Session("objCarteiraDoTituloCCobranca") Is Nothing Then
            Dim objCarteira As [Lib].Negocio.CarteiraDoTitulo = CType(Session("objCarteiraDoTituloCCobranca"), [Lib].Negocio.CarteiraDoTitulo)

            strSQL = "UPDATE ContasAReceber "
            strSQL &= "SET CarteiraDoTitulo = " & objCarteira.Codigo & " "
            strSQL &= " WHERE Registro_Id = " & gridCarteiraDeCobranca.Rows(HfCarteiraDeCobrancaIndex.Value).Cells(2).Text
            strSQL &= " AND Sequencia_Id = " & gridCarteiraDeCobranca.Rows(HfCarteiraDeCobrancaIndex.Value).Cells(3).Text

            Dim alSQL As New ArrayList
            alSQL.Add(strSQL)

            dvCarteira = New DataView(CType(Session("objTitulos"), DataSet).Tables("Titulos"))
            dvCarteira.Sort = "Registro_Id"

            If Banco.GravaBanco(alSQL) = True Then
                With dvCarteira

                    .AllowEdit = True
                    .Item(HfCarteiraDeCobrancaIndex.Value).BeginEdit()
                    .Item(HfCarteiraDeCobrancaIndex.Value)("Carteira") = Format(objCarteira.Codigo, "000") & " - " & objCarteira.Descricao
                    .Item(HfCarteiraDeCobrancaIndex.Value)("Carteira_Id") = objCarteira.Codigo
                    .Item(HfCarteiraDeCobrancaIndex.Value).EndEdit()
                End With
                gridCarteiraDeCobranca.DataSource = dvCarteira
                gridCarteiraDeCobranca.DataBind()
            End If

            HfCarteiraDeCobrancaIndex.Value = ""
            Session.Remove("objCarteiraDoTituloCCobranca")
        End If
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub ListarClientes()
        Session("ssCampo") = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCARCOB" & HID.Value, "txtNome")
    End Sub

    Private Sub CargaCarteiras()
        Dim objCarteira As New ListCarteiraDoTitulo
        ddlCarteira.Items.Clear()
        For Each row As [Lib].Negocio.CarteiraDoTitulo In objCarteira
            ddlCarteira.Items.Add(New ListItem(Funcoes.AlinharDireita(row.Codigo, 3, "0") & " - " & row.Descricao, row.Codigo))
        Next
        Funcoes.InserirLinhaEmBranco(ddlCarteira)
    End Sub

    Private Sub CarregarReceber()
        Empresa = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim objReceber As New ListBancosXContas(True, Empresa(0))
        Dim agencia As String = ""
        Dim conta As String = ""

        ddlReceber.Items.Clear()

        For Each row As [Lib].Negocio.BancosXContas In objReceber
            agencia = IIf(row.DigitoAgencia.Length > 0, Funcoes.AlinharDireita(row.Agencia, 4, " ") & "-" & row.DigitoAgencia, Funcoes.AlinharDireita(row.Agencia, 6, "0"))
            conta = IIf(row.DigitoConta.Length > 0, Funcoes.AlinharDireita(row.Conta, 10, " ") & "-" & row.DigitoConta, Funcoes.AlinharDireita(row.Conta, 10, " "))
            ddlReceber.Items.Add(New ListItem(Funcoes.AlinharEsquerda(row.NomeBanco, 20, ".") & " - AG: " & agencia & " - CTA: " & conta,
                                 row.CodigoBanco & ";" & row.Agencia & ";" & row.DigitoAgencia & ";" & row.Conta & ";" & row.DigitoConta))
        Next

        Funcoes.InserirLinhaEmBranco(ddlReceber)
    End Sub

    Private Sub CargaSituacao()
        Dim objSituacao As New ListSituacao(True)
        ddlSituacao.Items.Clear()
        For Each row As Situacao In objSituacao
            ddlSituacao.Items.Add(New ListItem(Funcoes.AlinharDireita(row.Codigo, 3, "0") & " - " & row.Descricao, row.Codigo))
        Next
        Funcoes.InserirLinhaEmBranco(ddlSituacao)
    End Sub

    Private Sub Limpar_Campos()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = ""

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param &= "Parametros:" & vbCrLf & "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Data: " & txtDataInicial.Text & "à:" & txtDataFinal.Text
        End If

        Return param
    End Function

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlEmpresa.SelectedIndex > 0 Then
            CarregarReceber()
        End If
    End Sub

    Public Function Relatorio(ByVal RegistrosCarteira As DataSet, ByVal Periodo As String, ByVal CNPJ As String, ByVal Endereco As String)
        Try
            If ValidarCampos() Then
                Dim Ds_Carteira As New DataSet
                Ds_Carteira = RegistrosCarteira

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", IIf(CkRelSimples.Checked, "Relatório De Títulos Em Carteira Não Agrupada.", "Relatório De Títulos Em Carteira."))
                parameters.Add("Parametros", getParam())

                Funcoes.BindReport(Me.Page, Ds_Carteira, IIf(CkRelSimples.Checked, "Cr_CarteiraDeCobrancaNaoAgrupada", "Cr_CarteiraDeCobranca"), eExportType.PDF, parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return String.Empty
    End Function

    Private Function ValidarCampos() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada")
            Return False
        End If
        Return True
    End Function

#Region "Ordem e Acesso ao DataGrid"
    ' Carrega o DataGrid com as informacoes do banco de dados
    Private Sub CarregarDataGrid()
        Dim drCarteira As DataRow
        If ValidarCampos() Then
            dsCarteira = Banco.ConsultaDataSet(GetSqlPreencheDataGrid(), "Titulos")
            If ddlCarteira.SelectedValue.ToString = "4" Then 'And Funcoes.VerificaPermissoes("BANCOSIPAL") Then
                'TemSipal = True
                dsCarteira.Tables(0).Columns.Add("Check", GetType(Boolean))
                dsCarteira.Tables(0).Columns("Check").ReadOnly = False
                dsCarteira.AcceptChanges()
                For Each drCarteira In dsCarteira.Tables(0).Rows
                    If (drCarteira("Situacao") = 51 Or drCarteira("Situacao") = 52) Then
                        drCarteira("Check") = True
                    End If
                Next
                dsCarteira.AcceptChanges()
            Else
                'TemSipal = False
            End If
            'dvCarteira = New DataView(dsCarteira.Tables("Titulos"))
            'dvCarteira.Sort = "Registro_Id"
            'gridCarteiraDeCobranca.DataSource = dvCarteira
            Session("objTitulos") = dsCarteira
            gridCarteiraDeCobranca.DataSource = Session("objTitulos")
            gridCarteiraDeCobranca.DataBind()

            Ordena()
        Else : gridCarteiraDeCobranca.DataSource = Nothing
        End If
    End Sub

    Private Sub Ordena()
        'lbOrdenacao
        'Verifica a Ordem no listbox e concatena as colunas
        'com o Order by na instrucao Select

        'Dim coluna As New Hashtable
        'For i As Integer = 0 To lbOrdenacao.Items.Count - 1
        '    Select Case lbOrdenacao.Items(i).ToString()
        '        Case "Registro"
        '            coluna.Add(i, clnRegistro.DataPropertyName)
        '        Case "Sequencia"
        '            coluna.Add(i, clnSequencia.DataPropertyName)
        '        Case "Vencimento"
        '            coluna.Add(i, clnVencimento.DataPropertyName)
        '        Case "Cliente"
        '            coluna.Add(i, clnNomeSacado.DataPropertyName)
        '            'Case "Representante"
        '            '    coluna.Add(i, clnRepresentante.DataPropertyName)
        '        Case "Valor"
        '            coluna.Add(i, clnValor.DataPropertyName)
        '    End Select
        'Next
        'dvCarteira.Sort = coluna(0) & "," & coluna(1) & "," & coluna(2) & "," & coluna(3) & "," & coluna(4)
    End Sub

#End Region

#Region "SQL"


    ' Monta o Select para preencher o DataGrid
    Private Function GetSqlPreencheDataGrid() As String
        Dim strSQL As String

        strSQL = "  SELECT "
        strSQL &= " 'R' as ReceberPagar,"
        strSQL &= " Case When CR.DataEnvio is null Then 'TRUE' Else 'FALSE' end as Enviar,      "
        strSQL &= " Case         When CR.Situacao = 101           Then 'RM'         "
        strSQL &= " When CR.Motivo is Not null and CR.Ocorrencia = 2           Then 'RM-OK'         "
        strSQL &= " When CR.Motivo is Not null and CR.Ocorrencia <> 2           Then 'RM-NOK'         "
        strSQL &= " When CR.DataEnvio is null           Then '  '         Else '  '       "
        strSQL &= " end as TAG,       "
        strSQL &= " " & IIf(CkRelSimples.Checked, "convert(nvarchar,NFXT.Nota_Id) ", "'NF: ' + convert(nvarchar,NFXT.Nota_Id)  +' - ' +  CR.Historico") & "  as Nota_Id,         "
        strSQL &= " CR.Registro_Id as Registro_Id, "
        strSQL &= " CR.Sequencia_Id,       "
        strSQL &= " CR.Vencimento,       "
        strSQL &= " Isnull(CR.Prorrogacao,'') as Prorrogacao,       "
        strSQL &= " CO.Reduzido AS Origem,        "
        strSQL &= " CD.Reduzido AS Destino,       "
        strSQL &= " Sacado.Nome AS Sacado,       "
        strSQL &= " " & IIf(CkRelSimples.Checked, "Sacado.Reduzido +'-' + Sacado.Nome", "Sacado.Nome") & "  AS NomeSacado, "
        strSQL &= " CR.ValorDoDocumento AS Receber, "
        strSQL &= " 0.00 as Pagar,       "
        strSQL &= " S.Situacao_Id as Situacao,       "
        strSQL &= " S.Descricao as DescSituacao,       "
        strSQL &= " Isnull(REPLACE(STR(Ca.Carteira_Id, 3), ' ', '0') + ' - ' + Ca.Descricao,'') AS Carteira,       "
        strSQL &= " Isnull(Ca.Carteira_Id,0) as Carteira_Id,       "
        strSQL &= " Repre.Reduzido +'-' + Repre.Nome as NomeRepre,       "
        strSQL &= " 0 as NumDias,       "
        strSQL &= " Sacado.Cliente_Id,       "
        strSQL &= " Sacado.Nome,       "
        strSQL &= " case when EndCobranca.Endereco is null then Sacado.Endereco else EndCobranca.Endereco end Endereco,        "
        strSQL &= " case when EndCobranca.Endereco is null then Sacado.cidade else EndCobranca.cidade end cidade,        "
        strSQL &= " case when EndCobranca.Endereco is null then isnull(Sacado.numero,0) else isnull(EndCobranca.numero,0) end numero,        "
        strSQL &= " case when EndCobranca.Endereco is null then Sacado.complemento else EndCobranca.complemento end complemento,        "
        strSQL &= " case when EndCobranca.Endereco is null then Sacado.CEP else EndCobranca.CEP end CEP,        "
        strSQL &= " case when EndCobranca.Endereco is null then Sacado.Bairro else EndCobranca.Bairro end Bairro,        "
        strSQL &= " case when EndCobranca.Endereco is null then Sacado.Estado else EndCobranca.Estado end Estado_Id,        "
        strSQL &= " Isnull(CR.DataEnvio,'') as DataEnvio,        "
        strSQL &= " Isnull(CR.Ocorrencia,'') as Ocorrencia,        "
        strSQL &= " isnull(CR.Motivo,'          ') as Motivo,        "
        strSQL &= " CR.bancoPagador,       "
        strSQL &= " Isnull(CR.NossoNumero,0) as NossoNumero, "
        strSQL &= " Isnull(CR.DigitoNossoNumero,'') as  DigitoNossoNumero "
        If CkRelSimples.Checked Then strSQL &= ", NotasFiscais.DataDaNota "
        strSQL &= " FROM ContasAReceber CR   "
        strSQL &= " Left Join Pedidos P    on  CR.Pedido  = P.Pedido_id  "
        strSQL &= " Left join Clientes EndCobranca    ON EndCobranca.Cliente_Id  = P.Praca   "
        strSQL &= " AND EndCobranca.Endereco_Id = P.EndPraca "
        strSQL &= " LEFT JOIN Clientes CO     ON CO.Cliente_Id  = CR.EmpresaPagadora  "
        strSQL &= " AND CO.Endereco_Id = CR.EndEmpresaPagadora "
        strSQL &= " INNER JOIN Clientes CD     ON CD.Cliente_Id  = CR.Cliente  "
        strSQL &= " AND CD.Endereco_Id = CR.EndCliente  "
        strSQL &= " LEFT JOIN Clientes Sacado    ON Sacado.Cliente_Id  = CR.Destinatario   "
        strSQL &= " AND Sacado.Endereco_Id = CR.EndDestinatario  "
        strSQL &= " LEFT JOIN Carteira Ca    ON Ca.Carteira_Id = Isnull(CR.CarteiraDoTitulo,0) "
        strSQL &= " Left outer Join Comissoes     On Comissoes.Pedido_Id = CR.Pedido    and Principal           = 'S'  "
        strSQL &= " Left outer Join Clientes as Repre     on Comissoes.Representante_Id    = Repre.Cliente_Id    "
        strSQL &= " and Comissoes.EndRepresentante_Id = Repre.Endereco_Id  "
        strSQL &= " Left join situacoes s    on CR.situacao = s.situacao_Id "
        strSQL &= " Inner join NotaFiscalXTitulo as NFXT ON CR.Registro_Id = NFXT.Titulo_Id "
        If CkRelSimples.Checked Then
            strSQL &= " LEFT JOIN  NotasFiscais "
            strSQL &= " ON NFXT.Empresa_Id = NotasFiscais.Empresa_Id "
            strSQL &= " AND NFXT.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id "
            strSQL &= " AND NFXT.Cliente_Id = NotasFiscais.Cliente_Id "
            strSQL &= " AND NFXT.EndCliente_Id = NotasFiscais.EndCliente_Id "
            strSQL &= " AND NFXT.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id "
            strSQL &= " AND NFXT.Serie_Id = NotasFiscais.Serie_Id "
            strSQL &= " AND NFXT.Nota_Id = NotasFiscais.Nota_Id "
        End If
        strSQL &= " WHERE  Sacado.Nome is not null "
        If rdEmissao.Checked Then
            strSQL &= " AND CR.Movimento "
        Else
            strSQL &= " AND CR.Vencimento "
        End If
        strSQL &= " BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' "
        strSQL &= "     AND CR.Provisao = 2"
        strSQL &= "     AND CR.Situacao not in (2,3,4,5)"

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            strSQL &= " AND CR.Empresa    ='" & strEmpresa(0) & "'" & vbCrLf &
                      " AND CR.EndEmpresa = " & strEmpresa(1) & vbCrLf
        End If

        If txtCodigoCliente.Value.Length > 0 Then
            Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
            If chkConsolidarCliente.Checked Then
                Sql &= " and left(CR.Destinatario ,8) = '" & strCliente(0).Substring(0, 8) & "'"      'Cliente

            Else
                strSQL &= " AND CR.Destinatario    ='" & strCliente(0) & "'" & vbCrLf &
                          " AND CR.EndDestinatario = " & strCliente(1) & vbCrLf
            End If
        End If

        If ddlCarteira.SelectedIndex > 0 Then
            strSQL &= " AND CR.CarteiradoTitulo= '" & ddlCarteira.SelectedValue & "' "
        End If

        strSQL &= "Order By " & IIf(rbOrdRegistro.Checked, "CR.Registro_Id", IIf(rbOrdSequencia.Checked, "CR.Sequencia_Id", IIf(rbOrdVencimento.Checked, IIf(rdEmissao.Checked, "CR.Movimento", "CR.Vencimento"), IIf(rbOrdCliente.Checked, "Sacado.Nome", "CR.ValorDoDocumento"))))

        'If ddlSituacao.SelectedIndex <> 0 Then
        '    If ddlSituacao.SelectedIndex = 1 Then
        '        strSQL &= " AND CR.Situacao in (101,102,103,104,105,106)"
        '    ElseIf ddlSituacao.SelectedIndex = 2 Then
        '        strSQL &= " AND CR.Situacao not in (101,102,103,104,105,106)"
        '    Else
        '        Dim strSituacao As String() = ddlSituacao.SelectedValue.Split("-")
        '        strSQL &= " AND CR.Situacao = " & strSituacao(0)
        '    End If
        'End If

        Return strSQL
    End Function


#End Region

#Region "Gerar Dados Bradesco"

    Public Sub GeraDadosBradesco(ByVal pBanco As String)
        Dim boletos As New BoletoNet.Boletos()
        'Dim Titulos As String
        Dim sql As String
        Dim EnviarNovamente As Boolean = False
        Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim strConta As String() = ddlReceber.SelectedValue.Split(";")
        Dim conta As New Conta(strConta(1), strConta(2), strConta(3), strConta(4))
        arr.Clear()
        Dim c As BoletoNet.Cedente = BuscaCedente("", strEmpresa(0), strEmpresa(1), conta)

        If c Is Nothing Then
            MsgBox(Me.Page, "Verifique os Parametros da carteira de cobrança para esta Empresa e Banco")
            Exit Sub
        End If

        For Each row As DataRow In CType(Session("objTitulos"), DataSet).Tables(0).Rows

            If CInt(row.Item("Carteira_Id")) = CInt(ddlCarteira.SelectedValue) And row.Item("Enviar").ToString.ToUpper = "TRUE" And row.Item("TAG") <> "PE" Then
                If (row.Item("TAG") = "RM" Or row.Item("TAG") = "RM-OK" Or row.Item("TAG") = "RM-NOK") And Not EnviarNovamente Then
                    EnviarNovamente = True
                End If

                'Data de Vencimento
                Dim vencimento As DateTime = DirectCast(row.Item("Vencimento"), DateTime)
                'new DateTime(2007, 9, 10);

                Dim s As New BoletoNet.Sacado(row.Item("Cliente_Id"), row("Nome"))
                s.Endereco.End = row.Item("Endereco")
                If IsDBNull(row.Item("Numero")) = True Then
                    s.Endereco.Numero = 0
                Else
                    s.Endereco.Numero = row.Item("Numero")
                End If

                s.Endereco.Bairro = row.Item("Bairro")
                s.Endereco.Cidade = row.Item("Cidade")
                s.Endereco.CEP = row.Item("CEP")
                s.Endereco.UF = row.Item("Estado_Id")

                ' Cria o Boleto
                Dim b As New BoletoNet.Boleto(vencimento, row.Item("Receber"), "09", New String("0", 11), c)
                b.NumeroDocumento = row.Item("Nota")
                'b.NumeroControleParticipante = row.Item("Registro_Id")
                b.JurosMora = CDbl(MoraD)
                b.DataDocumento = Date.Now

                b.Sacado = s

                ' Instrucoes
                If Not Instrucao01Bradesco Is Nothing Then
                    b.Instrucoes.Add(Instrucao01Bradesco)
                Else
                    b.Instrucoes.Add(New BoletoNet.Instrucao_Bradesco(0, 0))
                End If

                If Not Instrucao02Bradesco Is Nothing Then
                    b.Instrucoes.Add(Instrucao02Bradesco)
                Else
                    b.Instrucoes.Add(New BoletoNet.Instrucao_Bradesco(0, 0))
                End If

                ' Banco
                b.Banco = New BoletoNet.Banco(pBanco)
                'b.EspecieDocumento = New BoletoNet.EspecieDocumento_BancoBradesco(1)
                boletos.Add(b)


                sql = "  Update ContasAReceber set"
                sql &= "     DataEnvio       = '" & Date.Now.ToString("yyyy-MM-dd") & "'"
                sql &= "    ,BancoPagador           =  " & strConta(0)
                sql &= "    ,AgenciaPagadora         =  " & c.ContaBancaria.Agencia
                sql &= "    ,DigitoAgenciaPagadora   = '" & c.ContaBancaria.DigitoAgencia & "'"
                sql &= "    ,ContaPagadora   = '" & c.ContaBancaria.Conta & "'"
                sql &= "    ,DigitoContaPagadora     = '" & c.ContaBancaria.DigitoConta & "'"
                sql &= "    ,Destinacao      = ''"
                sql &= "    ,EmpresaPagadora   = '" & strEmpresa(0) & "'"
                sql &= "    ,EndEmpresaPagadora  =  " & strEmpresa(1)
                'sql &= "    ,ContaContabilPagadora     =  " & c.ContaContabil.ToString
                sql &= "    ,Motivo          =  NULL"
                sql &= "    ,Situacao     =  101"
                sql &= "  where Registro_Id = " & row.Item("Registro_Id").ToString
                arr.Add(sql)
            Else
                row.Item("Enviar") = "FALSE"
            End If
        Next


        If boletos.Count = 0 Then
            MsgBox(Me.Page, "Não há titulos a serem enviados")
            Exit Sub
        End If

        If GeraArquivoCNAB400(New BoletoNet.Banco(pBanco), boletos(0).Cedente, boletos) Then


            sql = "  Update ParametrosDaCarteiraDeCobranca set "
            sql &= "     SequenciaDeRemessa = SequenciaDeRemessa +1 "
            sql &= "  Where ParametrosDaCarteiraDeCobranca.Empresa_id         ='" & strEmpresa(0) & "'"
            sql &= "    and ParametrosDaCarteiraDeCobranca.EndEmpresa_id      = " & strEmpresa(1)
            sql &= "    and ParametrosDaCarteiraDeCobranca.Agencia_id         = " & c.ContaBancaria.Agencia
            sql &= "    and ParametrosDaCarteiraDeCobranca.DigitoDaAgencia_Id ='" & c.ContaBancaria.DigitoAgencia & "'"
            sql &= "    and ParametrosDaCarteiraDeCobranca.Conta_Id           = " & c.ContaBancaria.Conta
            sql &= "    and ParametrosDaCarteiraDeCobranca.DigitoConta_Id     ='" & c.ContaBancaria.DigitoConta & "'"
            arr.Add(sql)

            'If Banco.GravaBanco(arr) = True Then
            If False Then
                MsgBox(Me.Page, "Arquivo gerado com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "Arquivo gerado com sucesso, porem os registro não foram atualizados em nossa base de dados, entre em contato com os desenvolvedores!")
            End If
            lnkConsulta_Click(Nothing, Nothing)
        Else
            MsgBox(Me.Page, "Erro Durante a geração do arquivo!")
        End If
    End Sub

    Public Sub GerarRemessaItau()
        Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim strConta As String() = ddlReceber.SelectedValue.Split(";")
        Dim conta As New Conta(strConta(1), strConta(2), strConta(3), strConta(4))

        Dim ds = BuscaParametrosDaCarteiraDeCobranca(strEmpresa(0), strEmpresa(1), conta)

        Dim arquivoGerado = GeraArquivoCNAB400(ds, CType(Session("objTitulos"), DataSet))

        If (arquivoGerado) Then
            'TODO: Verificar necessidade desse update
            'Dim Sql = "  UPDATE ParametrosDaCarteiraDeCobranca SET " & vbCrLf &
            '          "        SequenciaDeRemessa = SequenciaDeRemessa + 1 " & vbCrLf &
            '          "  WHERE Empresa_id         = '" & strEmpresa(0) & "'" & vbCrLf &
            '          "    AND EndEmpresa_id      = " & strEmpresa(1) & vbCrLf &
            '          "    AND Agencia_id         = " & conta.Agencia & vbCrLf &
            '          "    AND DigitoDaAgencia_Id = '" & conta.DigitoAgencia & "'" & vbCrLf &
            '          "    AND Conta_Id           = " & conta.Conta_Id & vbCrLf &
            '          "    AND DigitoConta_Id     = '" & conta.DigitoConta_Id & "'" & vbCrLf
            'arr.Add(Sql)

            If Banco.GravaBanco(arr) = True Then
                CarregarDataGrid()
                MsgBox(Me.Page, "Arquivo gerado com Sucesso.", eTitulo.Sucess)
                btnDownload.Visible = True
                btnDownload.Text = "Download do arquivo de remessa do dia " & Date.Now.ToString("dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Arquivo gerado com sucesso, porem os registro não foram atualizados em nossa base de dados, entre em contato com o suporte!")
            End If
            'lnkConsulta_Click(Nothing, Nothing)
        Else
            MsgBox(Me.Page, "Erro Durante a geração do arquivo!")
        End If

    End Sub

    Public Function BuscaCedente(ByVal AgCt As String, Optional ByVal cliente_id As String = "", Optional ByVal endereco As String = "", Optional ByVal Conta As Conta = Nothing) As BoletoNet.Cedente
        MoraD = ""

        Dim sql As String
        sql = "  SELECT Clientes.Cliente_Id," & vbCrLf &
              "       Clientes.Endereco_Id," & vbCrLf &
              "       Clientes.Nome," & vbCrLf &
              "       PCC.Agencia_Id," & vbCrLf &
              "       PCC.DigitoDaAgencia_Id," & vbCrLf &
              "       PCC.Conta_Id," & vbCrLf &
              "       PCC.DigitoConta_Id," & vbCrLf &
              "       PCC.SubConta_Id," & vbCrLf &
              "       PCC.Instrucao01," & vbCrLf &
              "       PCC.Instrucao02," & vbCrLf &
              "       PCC.MensagemDeInstrucao," & vbCrLf &
              "       isnull(PCC.DiasParaProtesto,0) as DiasParaProtesto," & vbCrLf &
              "       isnull(PCC.DiasBaixaDecursoPrazo,0) as DiasBaixaDecursoPrazo," & vbCrLf &
              "       PCC.JurosMes," & vbCrLf &
              "       PCC.MoraDiaria," & vbCrLf &
              "       PCC.SequenciaDeRemessa," & vbCrLf &
              "       PCC.CodigoDaEmpresaNoBanco," & vbCrLf &
              "       PCC.ContaContabil," & vbCrLf &
              "       PCC.NomeDaEmpresaNoBanco" & vbCrLf &
              "  FROM Clientes" & vbCrLf &
              " INNER JOIN ParametrosDaCarteiraDeCobranca PCC" & vbCrLf &
              "    ON Clientes.Cliente_Id  = PCC.Empresa_Id " & vbCrLf &
              "   AND Clientes.Endereco_Id = PCC.EndEmpresa_Id" & vbCrLf &
              " Where 1 = 1" & vbCrLf
        If AgCt <> "" Then
            sql &= " and convert(varchar,PCC.Agencia_id) + convert(varchar,PCC.Conta_Id) = '" & AgCt & "'" & vbCrLf
        Else
            sql &= " and Clientes.Cliente_id    ='" & cliente_id & "'" & vbCrLf &
                   " and Clientes.Endereco_id   = " & endereco & vbCrLf &
                   " and PCC.Agencia_id         = " & Conta.Agencia & vbCrLf &
                   " and PCC.DigitoDaAgencia_Id ='" & Conta.DigitoAgencia & "'" & vbCrLf &
                   " and PCC.Conta_Id           = " & Conta.Conta_Id & vbCrLf &
                   " and PCC.DigitoConta_Id     ='" & Conta.DigitoConta_Id & "'" & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Cedente")
        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            For Each dr As DataRow In ds.Tables(0).Rows
                Instrucao01Bradesco = Nothing
                Instrucao02Bradesco = Nothing
                If Not IsDBNull(dr("Instrucao01")) Then
                    Select Case CInt(dr("Instrucao01"))
                        Case 6 : Instrucao01Bradesco = New BoletoNet.Instrucao_Bradesco(6, CInt(dr("DiasParaProtesto")))
                        Case 18 : Instrucao01Bradesco = New BoletoNet.Instrucao_Bradesco(18, CInt(dr("DiasBaixaDecursoPrazo")))
                        Case Else : Instrucao01Bradesco = New BoletoNet.Instrucao_Bradesco(CInt(dr("Instrucao01")), 0)
                    End Select
                End If
                If Not IsDBNull(dr("Instrucao02")) Then
                    Select Case CInt(dr("Instrucao02"))
                        Case 6 : Instrucao02Bradesco = New BoletoNet.Instrucao_Bradesco(6, CInt(dr("DiasParaProtesto")))
                        Case 18 : Instrucao02Bradesco = New BoletoNet.Instrucao_Bradesco(18, CInt(dr("DiasBaixaDecursoPrazo")))
                        Case Else : Instrucao02Bradesco = New BoletoNet.Instrucao_Bradesco(CInt(dr("Instrucao02")), 0)
                    End Select
                End If

                Dim c As New BoletoNet.Cedente(dr("Cliente_Id"), dr("Nome"), CInt(dr("Agencia_Id")).ToString("0000"), dr("DigitoDaAgencia_Id"), dr("Conta_Id"), dr("DigitoConta_Id"))
                'c.Codigo = Convert.ToInt64(dr("CodigoDaEmpresaNoBanco"))
                'c.NrArqRemessa = CInt(dr("SequenciaDeRemessa"))
                'c.ContaContabil = CInt(dr("ContaContabil"))
                'c.NomeDaEmpresaNoBanco = dr("NomeDaEmpresaNoBanco")
                'c.EndCnpj = CInt(dr("Endereco_Id"))

                MoraD = IIf(IsDBNull(dr("MoraDiaria")), "0", dr("MoraDiaria").ToString)
                Return c
            Next
        End If
        Return Nothing
    End Function

    Public Function BuscaParametrosDaCarteiraDeCobranca(ByVal cliente_id As String, Optional ByVal endereco As String = "", Optional ByVal Conta As Conta = Nothing) As DataSet

        Dim sql As String
        sql = "  SELECT Clientes.Cliente_Id," & vbCrLf &
              "       Clientes.Endereco_Id," & vbCrLf &
              "       Clientes.Nome," & vbCrLf &
              "       PCC.Agencia_Id," & vbCrLf &
              "       PCC.DigitoDaAgencia_Id," & vbCrLf &
              "       PCC.Conta_Id," & vbCrLf &
              "       PCC.DigitoConta_Id," & vbCrLf &
              "       PCC.SubConta_Id," & vbCrLf &
              "       PCC.Instrucao01," & vbCrLf &
              "       PCC.Instrucao02," & vbCrLf &
              "       PCC.MensagemDeInstrucao," & vbCrLf &
              "       isnull(PCC.DiasParaProtesto,0) as DiasParaProtesto," & vbCrLf &
              "       isnull(PCC.DiasBaixaDecursoPrazo,0) as DiasBaixaDecursoPrazo," & vbCrLf &
              "       PCC.JurosMes," & vbCrLf &
              "       PCC.MoraDiaria," & vbCrLf &
              "       PCC.SequenciaDeRemessa," & vbCrLf &
              "       PCC.CodigoDaEmpresaNoBanco," & vbCrLf &
              "       PCC.ContaContabil," & vbCrLf &
              "       PCC.NomeDaEmpresaNoBanco" & vbCrLf &
              "  FROM Clientes" & vbCrLf &
              " INNER JOIN ParametrosDaCarteiraDeCobranca PCC" & vbCrLf &
              "    ON Clientes.Cliente_Id  = PCC.Empresa_Id " & vbCrLf &
              "   AND Clientes.Endereco_Id = PCC.EndEmpresa_Id" & vbCrLf &
              " Where 1 = 1" & vbCrLf &
              " and Clientes.Cliente_id    ='" & cliente_id & "'" & vbCrLf &
                   " and Clientes.Endereco_id   = " & endereco & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Cedente")
        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        End If
        Return ds
    End Function

    Public Function GeraArquivoCNAB400(ByVal banco As BoletoNet.IBanco, ByVal cedente As BoletoNet.Cedente, ByVal boletos As BoletoNet.Boletos) As Boolean
        Try
            Dim mem = New MemoryStream()
            Dim arquivo As New BoletoNet.ArquivoRemessa(BoletoNet.TipoArquivo.CNAB400)
            arquivo.GerarArquivoRemessa("0", banco, cedente, boletos, mem, 1)

            Dim NomeArquivo2 As String = "RemessaBancaria/" & "remessa" & Date.Now.Day.ToString("00") & Date.Now.Month.ToString("00") & Date.Now.Minute.ToString("00") & ".rem"
            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            Dim bytes = mem.ToArray()
            File.WriteAllBytes(NomeArquivo, bytes)

            Return True

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Function GeraArquivoCNAB400(ByVal dsCedente As DataSet, ByVal dsTitulo As DataSet) As Boolean
        Try

            Dim NomeArquivo2 As String = "RemessaBancaria/" & "remessa.rem"

            Session("NomeArquivoDownload") = NomeArquivo2

            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            Dim remessa = New RemessaBancaria()

            Dim i As Integer = 0

            Dim _detalhe As New StringBuilder

            For Each dr As DataRow In dsCedente.Tables(0).Rows
                'Header
                i = i + 1
                remessa.Header.AppendRemessa(1, 1, True, "0")
                remessa.Header.AppendRemessa(2, 2, True, "1")
                remessa.Header.AppendRemessa(3, 9, False, "REMESSA")
                remessa.Header.AppendRemessa(10, 11, True, "04")
                remessa.Header.AppendRemessa(12, 26, False, "EMPRESTIMO")
                remessa.Header.AppendRemessa(27, 38, True, dr.Item("CodigoDaEmpresaNoBanco"))
                remessa.Header.AppendRemessa(39, 46, False, " ")
                remessa.Header.AppendRemessa(47, 76, False, dr.Item("Nome"))
                remessa.Header.AppendRemessa(77, 79, True, "341")
                remessa.Header.AppendRemessa(80, 94, False, "BANCO ITAU SA")
                remessa.Header.AppendRemessa(95, 100, True, DateTime.Now.ToString("ddMMyy"))
                remessa.Header.AppendRemessa(101, 394, False, " ")
                remessa.Header.AppendRemessa(395, 400, True, Convert.ToString(i))

                For Each drTitulo As DataRow In dsTitulo.Tables(0).Rows

                    Dim AdicionaAoArquivo = CInt(drTitulo.Item("Carteira_Id")) = CInt(ddlCarteira.SelectedValue) And drTitulo.Item("Enviar").ToString.ToUpper = "TRUE" And drTitulo.Item("TAG") <> "PE"

                    If Not AdicionaAoArquivo Then
                        Continue For
                    End If

                    'Detalhe
                    i = i + 1
                    _detalhe = New StringBuilder
                    _detalhe.AppendRemessa(1, 1, True, "1")
                    _detalhe.AppendRemessa(2, 3, True, "02")
                    _detalhe.AppendRemessa(4, 17, True, dr.Item("Cliente_Id"))
                    _detalhe.AppendRemessa(18, 29, True, dr.Item("CodigoDaEmpresaNoBanco"))
                    _detalhe.AppendRemessa(30, 37, False, " ")
                    _detalhe.AppendRemessa(38, 62, False, drTitulo.Item("Nota_Id"))
                    _detalhe.AppendRemessa(63, 70, True, "0")
                    _detalhe.AppendRemessa(71, 83, False, " ")
                    _detalhe.AppendRemessa(84, 86, True, "458")
                    _detalhe.AppendRemessa(87, 107, False, " ")
                    _detalhe.AppendRemessa(108, 108, False, "D")
                    _detalhe.AppendRemessa(109, 110, True, "01")
                    _detalhe.AppendRemessa(111, 120, False, drTitulo.Item("Registro_Id"))
                    _detalhe.AppendRemessa(121, 126, True, Convert.ToDateTime(drTitulo.Item("Prorrogacao")).ToString("ddMMyy"))
                    _detalhe.AppendRemessa(127, 139, True, drTitulo.Item("Receber").ToString().Replace(",", ""))
                    _detalhe.AppendRemessa(140, 142, True, "341")
                    _detalhe.AppendRemessa(143, 146, True, "0")
                    _detalhe.AppendRemessa(147, 147, True, "0")
                    _detalhe.AppendRemessa(148, 149, False, "01")
                    _detalhe.AppendRemessa(150, 150, False, "A")
                    _detalhe.AppendRemessa(151, 156, True, Convert.ToDateTime(drTitulo.Item("Prorrogacao")).ToString("ddMMyy"))
                    _detalhe.AppendRemessa(157, 158, False, "09")
                    _detalhe.AppendRemessa(159, 160, False, " ")
                    _detalhe.AppendRemessa(161, 166, True, DateTime.Now.ToString("ddMMyy"))
                    _detalhe.AppendRemessa(167, 218, False, " ")
                    _detalhe.AppendRemessa(219, 220, True, "02")
                    _detalhe.AppendRemessa(221, 234, True, drTitulo.Item("Cliente_Id"))
                    _detalhe.AppendRemessa(235, 264, False, drTitulo.Item("NomeSacado"))
                    _detalhe.AppendRemessa(265, 274, False, " ")
                    _detalhe.AppendRemessa(275, 314, False, drTitulo.Item("Endereco"))
                    _detalhe.AppendRemessa(315, 326, False, drTitulo.Item("Bairro"))
                    _detalhe.AppendRemessa(327, 334, False, drTitulo.Item("CEP"))
                    _detalhe.AppendRemessa(335, 349, False, drTitulo.Item("Cidade"))
                    _detalhe.AppendRemessa(350, 351, False, drTitulo.Item("Estado_Id"))
                    _detalhe.AppendRemessa(352, 381, False, dr.Item("Nome"))
                    _detalhe.AppendRemessa(382, 391, False, " ")
                    _detalhe.AppendRemessa(392, 393, True, "05")
                    _detalhe.AppendRemessa(394, 394, False, " ")
                    _detalhe.AppendRemessa(395, 400, True, Convert.ToString(i))

                    remessa.AddDetalhe(_detalhe)

                    Dim sql = "UPDATE ContasAReceber " & vbCrLf &
                              "      SET DataEnvio = " & DateTime.Now.ToString("dd/MM/yyyy") & vbCrLf &
                              "     ,Situacao = " & eSituacao.RemessaBancaria & vbCrLf &
                              "     ,SituacaoRemessaBancaria = " & eSituacaoBancaria.AguardandoArquivoDeRetorno & vbCrLf &
                              "     ,HistoricoRemessa = 'ENVIADO AO ARQUIVO DE REMESSA DO ITAÚ(CONVÊNIO 458), NA DATA DE: " + DateTime.Now.ToString("dd/MM/yyyy") & "'" & vbCrLf &
                              " WHERE Registro_Id = " & drTitulo.Item("Registro_Id")
                    arr.Add(sql)

                Next
            Next

            'Trailler
            i = i + 1
            Dim _trailer As New StringBuilder
            remessa.Trailler.AppendRemessa(1, 1, True, "9")
            remessa.Trailler.AppendRemessa(2, 394, False, " ")
            remessa.Trailler.AppendRemessa(395, 400, True, Convert.ToString(i))

            remessa.GerarArquivoDeRemessa(NomeArquivo)
            Return True

        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Public Sub GeraArquivoCNAB240(ByVal banco As BoletoNet.IBanco, ByVal cedente As BoletoNet.Cedente, ByVal boletos As BoletoNet.Boletos)
        'saveFileDialog.Filter = "Arquivos de Retorno (*.rem)|*.rem|Todos Arquivos (*.*)|*.*"
        'If saveFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
        '    Dim arquivo As New BoletoNet.ArquivoRemessa(BoletoNet.TipoArquivo.CNAB240)
        '    arquivo.GerarArquivoRemessa("1200303001417053", banco, cedente, boletos, saveFileDialog.OpenFile(), 1)

        '    MsgBox(Me.Page, "Arquivo gerado com Sucesso.", eTitulo.Sucess)
        'End If
    End Sub

    Public Sub CarregarRetornoRemessa()
        Dim arquivo As String = Server.MapPath("~/Files/" & fup.FileName)
        Dim objArquivo As New StreamReader(arquivo)
        Dim strLinha As String
        'Dim campos() As String
        Dim intLinha As Integer

        Dim ListTitulo As New [Lib].Negocio.ListTitulo
        Try
            Do While objArquivo.Peek >= 0
                intLinha += 1
                strLinha = objArquivo.ReadLine()
                If intLinha <> 1 Then
                    If Mid(strLinha, 1, 1).ToString().Equals("1") Then
                        Dim codArqRetorno = Mid(strLinha, 117, 10).Trim()
                        Dim DataPgto = Mid(strLinha, 111, 2) & "/" & Mid(strLinha, 113, 2) & "/" & Mid(strLinha, 115, 2)
                        DataPgto = DateTime.ParseExact(DataPgto, "dd/MM/yy", CultureInfo.InvariantCulture).ToString("dd/MM/yyyy")
                        Dim codRetorno = Mid(strLinha, 109, 2)

                        Dim descricaoRetorno = VerificaSituacoesBancarias(Mid(strLinha, 109, 2))

                        For i = 0 To GridRetornoTitulos.Rows.Count - 1
                            Dim codigoNaGrid = CType(GridRetornoTitulos.Rows(i).FindControl("hpTitulo"), HyperLink).Text
                            If codigoNaGrid = codArqRetorno Then
                                Dim enumerador As eRetornoRemessaItau = CType(CInt(codRetorno), eRetornoRemessaItau)
                                GridRetornoTitulos.Rows(i).Cells(3).Text = CInt(enumerador).ToString("00") & " - " & Textos.GetEnumDescription(enumerador)
                                GridRetornoTitulos.Rows(i).Cells(4).Text = DataPgto

                                CType(GridRetornoTitulos.Rows(i).FindControl("chkVencimento"), CheckBox).Checked = True
                            End If
                        Next

                    End If
                End If
            Loop

        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao Processar Arquivo de Retorno.", eTitulo.Erro)
        Finally
            fup.PostedFile.InputStream.Dispose()
            'fup.Dispose()
            objArquivo.Dispose()
        End Try
    End Sub

    Private Function VerificaSituacoesBancarias(ByVal CodRetorno As String)
        Dim Sql As String
        Dim ds As New DataSet
        Dim Resultado As String

        Sql = " SELECT     Situacao_Id, Descricao, Provisao " & vbCrLf &
              "   FROM SituacoesBancariasRetorno " & vbCrLf &
              "   WHERE Situacao_Id = '" & CodRetorno & "'"

        ds = Banco.ConsultaDataSet(Sql, "SituacoesBancariasRetorno")

        If ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                Resultado = Dr("Descricao")
                Return Resultado
            Next
        End If
        Return "Não Encontrada"
    End Function

    Private Sub Download()
        Response.ContentType = "text/plain"
        Response.AppendHeader("Content-Disposition", "attachment; filename=" & Session("NomeArquivoDownload").ToString.Split("/").ToArray(1))
        Const bufferLength As Integer = 10000
        Dim buffer As Byte() = New Byte(bufferLength - 1) {}
        Dim length As Integer = 0
        Dim download As Stream = Nothing
        Try
            download = New FileStream(Server.MapPath("~/" & Session("NomeArquivoDownload")), FileMode.Open, FileAccess.Read)
            Do
                If Response.IsClientConnected Then
                    length = download.Read(buffer, 0, bufferLength)
                    Response.OutputStream.Write(buffer, 0, length)
                    buffer = New Byte(bufferLength - 1) {}
                Else
                    length = -1
                End If
            Loop While length > 0
            Response.Flush()
            Response.End()
        Finally
            If download IsNot Nothing Then
                download.Close()
            End If
        End Try
    End Sub

    Public Sub ProcessarRetornoRemessa()

        Dim i = 0
        For Each row As GridViewRow In GridRetornoTitulos.Rows
            Dim chk As CheckBox = row.Cells(0).FindControl("chkVencimento")
            If chk.Checked Then
                i = i + 1
            End If
        Next
        If i = 0 Then
            MsgBox(Me.Page, "Nenhum título processado neste arquivo de retorno", eTitulo.Info)
            Exit Sub
        End If

        Dim resultado = False

        Try
            For Each row As GridViewRow In GridRetornoTitulos.Rows

                Dim chk As CheckBox = row.Cells(0).FindControl("chkVencimento")
                Dim _registroId As String = CType(row.FindControl("hpTitulo"), HyperLink).Text

                If chk.Checked Then

                    Dim tit As New Titulo(_registroId)
                    If tit.CodigoProvisao <> eProvisao.Baixa Then

                        Dim retorno As String() = TryCast(row.Cells(3), DataControlFieldCell).Text.Split("-")
                        Dim carteria As New ParametrosDaCarteiraDeCobranca(458) 'TODO: Codigo do convenio bancario
                        tit.IUD = "U"
                        Dim enumerador As eRetornoRemessaItau = CType(CInt(retorno(0)), eRetornoRemessaItau)

                        Select Case CInt(retorno(0))
                            Case eRetornoRemessaItau.LiquidacaoNormal
                                tit.CodigoProvisao = eProvisao.Baixa
                                tit.Baixa = CDate(row.Cells(4).Text)
                                tit.CodigoSituacao = eSituacao.Normal
                                tit.ContaCliente = carteria.ContaContabil
                                tit.CodigoBancoPagador = carteria.CodigoBanco
                                tit.CodigoAgenciaPagadora = carteria.CodigoAgencia
                                tit.DigitoAgenciaPagadora = carteria.DigitoAgencia
                                tit.ContaPagadora = carteria.Conta.ToString()
                                tit.DigitoContaPagadora = carteria.DigitoConta
                                tit.ContaContabilPagadora = carteria.BaixarPor
                                tit.UsuarioBaixa = Session("ssNomeUsuario")
                                tit.UsuarioBaixaData = Today()
                            Case eRetornoRemessaItau.DescontoRecusado
                                tit.CodigoSituacao = eSituacao.Normal
                        End Select

                        tit.SituacaoRemessaBancaria = enumerador
                        tit.HistoricoRemessa = tit.HistoricoRemessa + "\n ARQUIVO DE RETORNO DO ITAÚ, DATA: " & DateTime.Now.ToString("dd/MM/yyyy") & " " & retorno(0) & " - " & retorno(1)
                        resultado = tit.Salvar()

                        If resultado = False Then
                            Exit For
                        End If
                    End If
                End If
            Next

            If resultado = True Then
                MsgBox(Me.Page, "Arquivo de retorno processado com sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "Erro ao processar arquivo de retorno.", eTitulo.Erro)
            End If

        Catch ex As Exception
            Dim e = ex.Message
        Finally

        End Try
    End Sub

#End Region

    Protected Sub chkGridTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim chkTitulo As CheckBox = CType(sender, CheckBox)
        Dim row As GridViewRow = CType(chkTitulo.NamingContainer, GridViewRow)

        If chkTitulo.Checked Then
            CType(Session("objTitulos"), DataSet).Tables(0).Rows(row.RowIndex).Item("Enviar") = "TRUE"
        Else
            CType(Session("objTitulos"), DataSet).Tables(0).Rows(row.RowIndex).Item("Enviar") = "FALSE"
        End If
    End Sub

    Protected Sub gridCarteiraDeCobranca_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        If gridCarteiraDeCobranca.SelectedRow.Cells(2).Text = "RM" Or gridCarteiraDeCobranca.SelectedRow.Cells(2).Text = "RM-OK" Then
            MsgBox(Me.Page, "A alteração de carteira e proibida para titulos já enviados para cobrança")
            Exit Sub
        End If

        Dim alCols As New ArrayList
        Dim alLabels As New ArrayList
        Dim alResultado As New ArrayList

        strJavaScript = "var x = (screen.height / 2) - 250; "
        strJavaScript += "var y = (screen.width / 2) - 400; "
        strJavaScript += "window.open(""ConsultaCarteirasDeTitulo.aspx?tipo=CCobranca"", """", ""resizable=no, menubar=no, scrollbars=yes, width=430, height=430, top="" + x + "", left="" + y + """");"
        ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), "ConsultaCarteirasDeTitulo", strJavaScript, True)
        HfCarteiraDeCobrancaIndex.Value = gridCarteiraDeCobranca.SelectedIndex
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo") = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCARCOB" & HID.Value, "txtNome")
    End Sub

    Protected Sub lnkConsulta_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        btnDownload.Visible = False
        CarregarDataGrid()
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("CarteiraDeCobranca", "RELATORIO") Then
                Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
                CarregarDataGrid()
                Relatorio(dsCarteira, txtDataInicial.Text & " à " & txtDataFinal.Text, strEmpresa(0), strEmpresa(1))
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRemessa_Click(sender As Object, e As EventArgs) Handles lnkRemessa.Click
        If Session("objTitulos") Is Nothing OrElse CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Não há registros carregados")
            Exit Sub
        End If

        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione uma empresa")
            Exit Sub
        End If

        If ddlCarteira.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione uma Carteira")
            Exit Sub
        End If

        If ddlReceber.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione uma Agencia/Conta")
            Exit Sub
        End If
        Dim strBanco As String() = ddlReceber.SelectedValue.Split(";")
        Select Case CInt(strBanco(0))
            Case 237
                GeraDadosBradesco(237)
            Case 341
            Case 479
                GerarRemessaItau()
        End Select

    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar_Campos()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CarteiraDeCobranca")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkProcessarRetorno_Click(sender As Object, e As EventArgs) Handles lnkProcessarRetorno.Click
        Try
            ProcessarRetornoRemessa()
            CarregarTitulosEnviadosAoBanco()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível processar o retorno da remessa bancária.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub btnProcessarRetorno_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If fup.HasFile Then
            'Pegamos as informacoes do arquivo postado 
            'Definimos onde ele será salvo 

        End If
    End Sub

    Protected Sub btnDownload_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDownload.Click
        Try
            Download()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Sub CarregarTitulosEnviadosAoBanco()
        fup.Visible = True
        divRetorno.Visible = False
        lblArquivoRetorno.Text = ""

        Dim where = "Situacao = " & CInt(eSituacao.RemessaBancaria) & " AND Provisao <> " & CInt(eProvisao.Baixa)
        Dim titulos As New ListTitulo(where)

        If titulos.Any() Then

            For Each tit In titulos
                If tit.SituacaoBancaria = CInt(eSituacaoBancaria.AguardandoArquivoDeRetorno) Then
                    tit.Observacoes = CInt(eSituacaoBancaria.AguardandoArquivoDeRetorno).ToString("00") & " - " & Textos.GetEnumDescription(eSituacaoBancaria.AguardandoArquivoDeRetorno)
                End If
            Next

            GridRetornoTitulos.DataSource = From tit In titulos
                                            Select tit.Codigo, tit.Vencimento, tit.Observacoes, tit.Historico, tit.ValorLiquido, CodigoCifrado = Funcoes.Cifrar("ContasAReceber-" & tit.Codigo)
            GridRetornoTitulos.DataBind()
        End If
    End Sub

    Protected Sub imbExcluirRetorno_Click(sender As Object, e As ImageClickEventArgs)
        Try
            CarregarTitulosEnviadosAoBanco()
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao excluir arquivo de retorno")
        Finally
            'Kill(Server.MapPath("~/Boletos/" & e.CommandName & ".pdf"))
        End Try
    End Sub

End Class

#Region "Class Conta"

Public Class Conta
    Private _Agencia As String
    Private _DigitoAgencia As String
    Private _Conta_Id As String
    Private _DigitoConta_Id As String
    Private _SequenciaDePagamento As Integer
    Private _SequenciaDeRegistroDePagamento As Integer

    Public Sub New(ByVal a As String, ByVal da As String, ByVal c As String, ByVal dc As String)
        _Agencia = a
        _DigitoAgencia = da
        _Conta_Id = c
        _DigitoConta_Id = dc
    End Sub

    Public Sub New(ByVal a As String, ByVal da As String, ByVal c As String, ByVal dc As String, ByVal s As Integer, ByVal sr As Integer)
        _Agencia = a
        _DigitoAgencia = da
        _Conta_Id = c
        _DigitoConta_Id = dc
        _SequenciaDePagamento = s
        _SequenciaDeRegistroDePagamento = sr
    End Sub

    Public Property Agencia() As String
        Get
            Return _Agencia
        End Get
        Set(ByVal value As String)
            _Agencia = value
        End Set
    End Property

    Public Property DigitoAgencia() As String
        Get
            Return _DigitoAgencia
        End Get
        Set(ByVal value As String)
            _DigitoAgencia = value
        End Set
    End Property

    Public Property Conta_Id() As String
        Get
            Return _Conta_Id
        End Get
        Set(ByVal value As String)
            _Conta_Id = value
        End Set
    End Property

    Public Property DigitoConta_Id() As String
        Get
            Return _DigitoConta_Id
        End Get
        Set(ByVal value As String)
            _DigitoConta_Id = value
        End Set
    End Property

    Public Property SequenciaDePagamento() As Integer
        Get
            Return _SequenciaDePagamento
        End Get
        Set(ByVal value As Integer)
            _SequenciaDePagamento = value
        End Set
    End Property

    Public Property SequenciaDeRegistroDePagamento() As Integer
        Get
            Return _SequenciaDeRegistroDePagamento
        End Get
        Set(ByVal value As Integer)
            _SequenciaDeRegistroDePagamento = value
        End Set
    End Property

    Public ReadOnly Property Descricao()
        Get
            If Agencia.Length = 0 Then
                Return "Selecione uma Agencia/Conta"
            Else
                Return "Agencia - " & Agencia & "-" & DigitoAgencia & "  Conta - " & Conta_Id & "-" & DigitoConta_Id
            End If

        End Get
    End Property

End Class

Public Class OcorrenciaRetorno
    Private _codigo As Integer
    Private _descricao As String
    Private _baixaTitulo As Boolean

    Public Property Codigo() As Integer
        Get
            Return _codigo
        End Get
        Set(ByVal value As Integer)
            _codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _descricao
        End Get
        Set(ByVal value As String)
            _descricao = value
        End Set
    End Property

    Public Property BaixaTitulo() As Boolean
        Get
            Return _baixaTitulo
        End Get
        Set(ByVal value As Boolean)
            _baixaTitulo = value
        End Set
    End Property

End Class

Module RemessaBancariaExtension
    <Extension()>
    Public Sub AppendRemessa(ByVal sb As StringBuilder,
        ByVal BeginPosition As Integer,
        ByVal EndPosition As Integer,
        ByVal IsNumber As Boolean,
        ByVal Value As String)

        sb.Append(FitStringLength(
                 Value,
                 (EndPosition + 1) - BeginPosition,
                 (EndPosition + 1) - BeginPosition,
                  IIf(IsNumber, "0", " "),
                  0,
                  True,
                  True,
                  IsNumber).ToString().ToUpper())
    End Sub

    Private Function FitStringLength(ByVal SringToBeFit As String, ByVal maxLength As Integer, ByVal minLength As Integer,
                                           ByVal FitChar As Char, ByVal maxStartPosition As Integer, ByVal maxTest As Boolean,
                                           ByVal minTest As Boolean, ByVal isNumber As Boolean) As String
        Try
            Dim result As String = ""

            If maxTest = True Then
                ' max
                If SringToBeFit.Length > maxLength Then
                    result += SringToBeFit.Substring(maxStartPosition, maxLength)
                End If
            End If

            If minTest = True Then
                ' min
                If SringToBeFit.Length <= minLength Then
                    If isNumber = True Then
                        result += DirectCast((New String(FitChar, (minLength - SringToBeFit.Length)) + SringToBeFit), String)
                    Else
                        result += DirectCast((SringToBeFit + New String(FitChar, (minLength - SringToBeFit.Length))), String)
                    End If
                End If
            End If
            Return result
        Catch ex As Exception
            Dim tmpEx As New Exception("Problemas ao Formatar a string. String = " & SringToBeFit, ex)
            Throw tmpEx
        End Try
    End Function
End Module

#End Region