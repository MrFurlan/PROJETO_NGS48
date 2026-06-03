Imports System.Data
Imports System.Data.SqlClient
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CalculaJurosAdiantamentos
    Inherits BasePage

    Dim sql As String
    Dim ds As New DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Adiantamento)
        If Not IsPostBack AndAlso IsConnect Then
            If Funcoes.VerificaPermissao("CalculaJurosAdiantamentos", "ACESSAR") Then
                txtMovimento.Text = Format(Today, "dd/MM/yyyy")
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaEmpresas.SetarHID(HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Adiantamento.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub ImgEmpresa_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        HttpContext.Current.Session("ssCampo") = "Adiantamento"
        ucConsultaEmpresas.Limpar()
        Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaCalcJuros" & HID.Value)
    End Sub

    Protected Sub ImgClientes_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        HttpContext.Current.Session("ssCampo") = "Adiantamento"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCalcJuros" & HID.Value)
    End Sub

    Protected Sub btnNovo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If Funcoes.VerificaPermissao("CalculaJurosAdiantamentos", "GRAVAR") Then
            If CalculaJurosAdiantamento(txtMovimento.Text, txtCnpjDaEmpresa.Text.Replace(".", "").Replace("/", "").Replace("-", ""), txtEndEmpresa.Text, txtCnpjDoCliente.Text.Replace(".", "").Replace("/", "").Replace("-", ""), txtEndDoCliente.Text, "24") = True Then
                ScriptManager.RegisterClientScriptBlock(Me, Me.btnCalcular.GetType(), Guid.NewGuid().ToString(), "alert('Processo concluido !!!');", True)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
        End If
    End Sub

    Private Function CalculaJurosAdiantamento(ByVal DataAtualizacao As String, ByVal CnpjEmpresa As String, ByVal EndEmpresa As String, ByVal CnpjCliente As String, ByVal EndCliente As String, ByVal Lote As String) As Boolean
        Try
            Dim sql As String
            Dim ArraySql As New ArrayList
            Dim ds_Adiantamentos As New DataSet
            Dim ds_Cotacoes As New DataSet
            Dim drCotacao As DataRow
            Dim drVariacao As DataRow
            Dim NumDias As Integer
            Dim juro As Double
            Dim cotacao As Double
            Dim ds_sequencia As New DataSet
            Dim dr_sequencia As DataRow
            Dim sequencia As Integer

            If DataAtualizacao = Nothing Then
                MsgBox(Me.Page, "Informe a data de atualização!")
                Return False
            End If

            Dim dt As New DateTime
            dt = DateTime.Now()
            dt = dt.AddSeconds(1)

            'Busca Adiantamentos menores que a data de atualização informada
            sql = "SELECT Adiantamentos.Empresa_ID,Adiantamentos.EndEmpresa_ID,Empresa.Reduzido as ReduzidoEmpresa, Adiantamentos.Cliente_ID, Adiantamentos.EndCliente_ID,Adiantamentos.Adiantamento_Id, Adiantamentos.Movimento, Adiantamentos.Vencimento, Adiantamentos.UltimaAtualizacao, "
            sql &= " Adiantamentos.ValorMoeda AS CapitalAdiantamento,Adiantamentos.ValorOficial AS CapitalAdiantamentoOficial, SUM(ISNULL(AdiantamentosXBaixas.ValorMoeda, 0)) AS ValorBaixa, "
            sql &= " SUM(ISNULL(AdiantamentosXBaixas.VariacaoMoeda, 0)) AS ValorJuros, Adiantamentos.Taxa, convert(Decimal,0) as Juro , Adiantamentos.ValorMoeda + SUM(ISNULL(AdiantamentosXBaixas.VariacaoMoeda, 0)) - SUM(ISNULL(AdiantamentosXBaixas.ValorMoeda, 0)) AS Saldo, Clientes.Nome, SUM(ISNULL(AdiantamentosXBaixas.VariacaoOficial, 0)) AS ValorVariacao, SUM(ISNULL(AdiantamentosXBaixas.ValorOficial, 0)) AS ValorBaixaOficial"
            sql &= " FROM Adiantamentos INNER JOIN"
            sql &= " Clientes AS Empresa ON  Adiantamentos.Empresa_ID = Empresa.Cliente_Id AND "
            sql &= " Adiantamentos.EndEmpresa_ID = Empresa.Endereco_Id INNER JOIN Clientes ON Adiantamentos.Cliente_ID = Clientes.Cliente_Id AND "
            sql &= " Adiantamentos.EndCliente_ID = Clientes.Endereco_Id LEFT OUTER JOIN"
            sql &= " AdiantamentosXBaixas ON "
            sql &= " Adiantamentos.Empresa_ID = AdiantamentosXBaixas.Empresa_Id AND Adiantamentos.EndEmpresa_ID = AdiantamentosXBaixas.EndEmpresa_Id AND "
            sql &= " Adiantamentos.Cliente_ID = AdiantamentosXBaixas.Cliente_Id AND Adiantamentos.EndCliente_ID = AdiantamentosXBaixas.EndCliente_Id AND "
            sql &= " Adiantamentos.Adiantamento_Id = AdiantamentosXBaixas.Adiantamento_Id"
            sql &= " WHERE Adiantamentos.Movimento < '" & Format(CDate(DataAtualizacao), "yyyy/MM/dd") & "'"
            If CnpjEmpresa = Nothing And EndEmpresa = Nothing And CnpjCliente <> Nothing And EndCliente <> Nothing Then
                MsgBox(Me.Page, "Informe a empresa para realizar os cálculos!")
                Return False
                Exit Function
            End If
            If CDate(DataAtualizacao) > Now Then
                MsgBox(Me.Page, "A data de atualização não pode ser maior que a data atual!")
                Return False
                Exit Function
            End If
            If CnpjEmpresa <> Nothing And EndEmpresa <> Nothing Then
                sql &= "And Adiantamentos.Empresa_ID='" & CnpjEmpresa & "' And Adiantamentos.EndEmpresa_ID=" & EndEmpresa & ""
            End If
            If CnpjCliente <> Nothing And EndCliente <> Nothing Then
                sql &= "And Adiantamentos.Cliente_ID='" & CnpjCliente & "' And Adiantamentos.EndCliente_ID=" & EndCliente & ""
            End If
            sql &= " GROUP BY Adiantamentos.Adiantamento_Id, Adiantamentos.Movimento, Adiantamentos.Vencimento, Adiantamentos.UltimaAtualizacao, "
            sql &= " Adiantamentos.ValorMoeda,Adiantamentos.ValorOficial, Adiantamentos.Taxa, Adiantamentos.Empresa_ID, Adiantamentos.EndEmpresa_ID,Empresa.Reduzido, Adiantamentos.Cliente_ID, "
            sql &= " Adiantamentos.EndCliente_ID,Clientes.Nome"
            sql &= " having(Adiantamentos.ValorMoeda + SUM(ISNULL(AdiantamentosXBaixas.VariacaoMoeda, 0)) - SUM(ISNULL(AdiantamentosXBaixas.ValorMoeda, 0)) > 0)"

            ds_Adiantamentos = banco.ConsultaDataSet(sql, "Adiantamentos")

            'Busca Valor da Contação da Moeda
            sql = "Select Data_Id,Realizado,Indice From Cotacoes "
            sql &= " WHERE (Indexador_Id =" & 3 & " And Data_Id='" & Format(CDate(DataAtualizacao), "yyyy/MM/dd") & "')"
            ds_Cotacoes = banco.ConsultaDataSet(sql, "Cotacoes")
            For Each drCotacao In ds_Cotacoes.Tables(0).Rows
                cotacao = drCotacao("Indice")
            Next

            sql = "Select isnull(max(Sequencia_Id),0) as Sequencia_Id from AdiantamentosXBaixas where Empresa_ID='" & CnpjEmpresa & "' and EndEmpresa_ID=" & EndEmpresa & " "
            ds_sequencia = banco.ConsultaDataSet(sql, "AdiantamentosXBaixas")
            For Each dr_sequencia In ds_sequencia.Tables(0).Rows
                sequencia = dr_sequencia("Sequencia_Id")
            Next

            Dim drSeqRazao As DataRow
            Dim dsSeqRazao As New DataSet
            Dim seqLoteRazao As Integer
            '''''''''''''''''''''''''
            'GERA SEQUENCIA DE LOTES RAZAO 
            '''''''''''''''''''''''''
            Try

                sql = "Select isnull(max(Sequencia_Id),0) as Sequencia_Id from Razao"
                sql &= " WHERE (Empresa_Id = '" & CnpjEmpresa & "' And EndEmpresa_Id=" & EndEmpresa & " and Movimento_Id ='" & Format(CDate(DataAtualizacao), "yyyy/MM/dd") & "' and Lote_Id=" & Lote & ")"
                dsSeqRazao = banco.ConsultaDataSet(sql, "AdiantamentosXBaixas")
                For Each drSeqRazao In dsSeqRazao.Tables(0).Rows
                    seqLoteRazao = drSeqRazao("Sequencia_Id")
                Next
            Catch ex As SqlException
                MsgBox(Me.Page, ex.Message)
                Return False
                Exit Function
            End Try


            '''''''''''''''''''''
            ''Calcula a variação 
            '''''''''''''''''''''
            For Each drVariacao In ds_Adiantamentos.Tables(0).Rows


                '    If IsDBNull(drVariacao("UltimaAtualizacao")) = False Then
                '        'Busca Valor da Contação da Moeda
                '        sql = "Select Data_Id,Realizado,Cotacao From Cotacoes "
                '        sql &= " WHERE (Indexador_Id =" & 3 & " And Data_Id='" & Format(CDate(drVariacao("UltimaAtualizacao")), "yyyy/MM/dd") & "')"
                '        ds_Cotacoes = banco.ConsultaDataSet(sql, "Cotacoes")
                '        For Each drCotacao In ds_Cotacoes.Tables(0).Rows
                '            ultimacotacao = drCotacao("Cotacao")
                '        Next
                '    Else
                '        'Busca Valor da Contação da Moeda
                '        sql = "Select Data_Id,Realizado,Cotacao From Cotacoes "
                '        sql &= " WHERE (Indexador_Id =" & 3 & " And Data_Id='" & Format(CDate(drVariacao("Movimento")), "yyyy/MM/dd") & "')"
                '        ds_Cotacoes = banco.ConsultaDataSet(sql, "Cotacoes")
                '        For Each drCotacao In ds_Cotacoes.Tables(0).Rows
                '            ultimacotacao = drCotacao("Cotacao")
                '        Next
                '    End If
                '    If IsDBNull(drVariacao("UltimaAtualizacao")) = True Then
                '        NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("Movimento"), "yyyy/MM/dd")), CDate(DataAtualizacao))
                '    Else
                '        NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("UltimaAtualizacao"), "yyyy/MM/dd")), CDate(DataAtualizacao))
                '    End If
                '    If NumDias > 0 Then
                '        'ultimavaricao = ((drVariacao("CapitalAdiantamento") + drVariacao("ValorJuros")) * ultimacotacao)
                '        ultimavaricao = ((drVariacao("CapitalAdiantamentoOficial") + drVariacao("ValorVariacao")) - drVariacao("ValorBaixaOficial"))
                '        variacao = (((drVariacao("CapitalAdiantamento") + drVariacao("ValorJuros") - drVariacao("ValorBaixa")) * cotacao) - ultimavaricao)
                '    End If
                '    If variacao <> 0 Then
                '        Registro = ""
                '        Registro = Format(dt, "yy")
                '        Registro &= Format(dt.Month(), "00")
                '        Registro &= Format(dt.Day(), "00")
                '        Registro &= Format(dt.Hour, "00")
                '        Registro &= Format(dt.Minute(), "00")
                '        Registro &= Format(dt.Second(), "00")

                '        sql = "INSERT INTO AdiantamentosXBaixas"
                '        sql &= "(Pais_Id, Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id,Sequencia_Id,RegistroPedido,Titulo,VariacaoMoeda,VariacaoOficial, DataBaixa)"
                '        sql &= " VALUES('" & Utilitarios.Global.GPais & "','" & drVariacao("Empresa_ID") & "' ," & drVariacao("EndEmpresa_ID") & ",'" & drVariacao("Cliente_ID") & "'," & drVariacao("EndCliente_ID") & ",'" & drVariacao("Adiantamento_Id") & "','" & drVariacao("ReduzidoEmpresa") + Registro & "'," & "0" & " ," & 0 & "," & Replace(0, ",", ".") & "," & Replace(variacao, ",", ".") & " ,'" & Format(CDate(DataAtualizacao), "yyyy/MM/dd") & "')"
                '        ArraySql.Add(sql)
                '        'Variação Positiva
                '        If variacao > 0 Then
                '            'Lançamento de Debito
                '            ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "11301", drVariacao("Cliente_Id"), drVariacao("EndCliente_ID"), "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), variacao, 0, 0, 0, "Var. Cambiaria sobre adelanto N. " & drVariacao("Adiantamento_Id"), "4", "R"))
                '            'Lançamento de Credito
                '            ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "422020003", "", 0, "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), 0, variacao, 0, 0, "Var. Cambiaria sobre adelanto N. " & drVariacao("Adiantamento_Id") & " - " & drVariacao("Nome"), "4", "R"))
                '            'Variação Negativa
                '        ElseIf variacao < 0 Then
                '            'Lançamento de Debito
                '            ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "422010003", "", 0, "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), variacao * (-1), 0, 0, 0, "Var. Cambiaria sobre adelanto N. " & drVariacao("Adiantamento_Id") & " - " & drVariacao("Nome"), "4", "R"))
                '            'Lançamento de Credito
                '            ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "11301", drVariacao("Cliente_Id"), drVariacao("EndCliente_ID"), "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), 0, variacao * (-1), 0, 0, "Var. Cambiaria sobre adelanto N. " & drVariacao("Adiantamento_Id"), "4", "R"))
                '        End If
                '        dt = dt.AddSeconds(1)
                '    End If

                '''''''''''''''
                'Calcula Juros
                '''''''''''''''
                If CDate(drVariacao("Vencimento")) < CDate(DataAtualizacao) Then
                    Dim TotCapital As Decimal
                    TotCapital = drVariacao("CapitalAdiantamento") + drVariacao("ValorJuros") - drVariacao("ValorBaixa")

                    If IsDBNull(drVariacao("UltimaAtualizacao")) = True Then
                        NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("Vencimento"), "yyyy/MM/dd")), CDate(DataAtualizacao))
                    Else
                        If CDate(Format(drVariacao("Vencimento"), "yyyy/MM/dd")) > CDate(Format(drVariacao("UltimaAtualizacao"), "yyyy/MM/dd")) Then
                            NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("Vencimento"), "yyyy/MM/dd")), CDate(DataAtualizacao))
                        Else
                            NumDias = DateDiff(DateInterval.Day, CDate(Format(drVariacao("UltimaAtualizacao"), "yyyy/MM/dd")), CDate(DataAtualizacao))
                        End If
                    End If
                    If NumDias > 0 Then
                        juro = ((((TotCapital * CDec(drVariacao("Taxa")) / 100) / 30) * NumDias))
                    End If
                End If

                If NumDias > 0 Then

                    sequencia = sequencia + 1

                    seqLoteRazao = seqLoteRazao + 1

                    If juro > 0 Then

                        sql = "INSERT INTO AdiantamentosXBaixas"
                        sql &= "(Empresa_ID, EndEmpresa_ID, Cliente_ID, EndCliente_ID, Adiantamento_Id,Sequencia_Id,RegistroPedido,Titulo,VariacaoMoeda,VariacaoOficial, DataBaixa)"
                        sql &= " VALUES('" & drVariacao("Empresa_ID") & "' ," & drVariacao("EndEmpresa_ID") & ",'" & drVariacao("Cliente_ID") & "'," & drVariacao("EndCliente_ID") & ",'" & drVariacao("Adiantamento_Id") & "'," & sequencia & "," & "0" & " ," & 0 & "," & Replace(CDec(FormatNumber(juro, 2)), ",", ".") & " ," & Replace(CDec(FormatNumber(juro, 2)) * cotacao, ",", ".") & " ,'" & Format(CDate(DataAtualizacao), "yyyy/MM/dd") & "')"
                        ArraySql.Add(sql)

                        'Lançamento de Debito
                        ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "1010905", drVariacao("Cliente_Id"), drVariacao("EndCliente_ID"), "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), CDec(FormatNumber(juro, 2)) * cotacao, 0, CDec(FormatNumber(juro, 2)), 0, "Juros sobre adiantamento N. " & drVariacao("Adiantamento_Id"), "0", "R", seqLoteRazao))

                        seqLoteRazao = seqLoteRazao + 1
                        'Lançamento de Credito
                        ArraySql.Add(GravaRazao(drVariacao("Empresa_ID"), drVariacao("EndEmpresa_ID"), Format(CDate(DataAtualizacao), "dd/MM/yyyy"), Lote, "303030301 ", "", 0, "0", "", "", "1", "", "3", Format(CDate(DataAtualizacao), "yyyy/MM/dd"), 0, CDec(FormatNumber(juro, 2)) * cotacao, 0, CDec(FormatNumber(juro, 2)), "Juros sobre adiantamento N. " & drVariacao("Adiantamento_Id") & " - " & drVariacao("Nome"), "0", "R", seqLoteRazao))

                    End If

                    dt = dt.AddSeconds(1)
                End If


                If NumDias > 0 Then
                    'Atualiza a data da ultima atualização de jurus realizadas 
                    sql = "Update Adiantamentos"
                    sql &= " SET UltimaAtualizacao ='" & Format(CDate(DataAtualizacao), "yyyy/MM/dd") & "'"
                    sql &= " WHERE (Empresa_ID = '" & drVariacao("Empresa_ID") & "') AND (EndEmpresa_ID = " & drVariacao("EndEmpresa_ID") & ") AND (Cliente_ID = '" & drVariacao("Cliente_ID") & "') AND (EndCliente_ID = " & drVariacao("EndCliente_ID") & ") AND (Adiantamento_Id = '" & drVariacao("Adiantamento_Id") & "')"
                    ArraySql.Add(sql)
                End If
            Next


            If banco.GravaBanco(ArraySql) = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function GravaRazao(ByVal CnpjEmpresa As String, ByVal EndEmpresa As Integer, ByVal DataMovimento As String, ByVal Lote As String, ByVal Conta As String, ByVal Cliente_Id As String, ByVal EndCliente_Id As Integer, ByVal Produto As String, ByVal Titulo As String, ByVal Pedido As String, ByVal TipoLancto As String, ByVal CentroDeCustos As String, ByVal Indexador As String, ByVal DataMoeda As String, ByVal DebitoOficial As Double, ByVal CreditoOficial As Double, ByVal DebitoMoeda As Double, ByVal CreditoMoeda As Double, ByVal Historico As String, ByVal Moeda As String, ByVal PrevistoRealizado As String, ByVal Sequencia As Integer) As String
        Dim sql As String

        '''''''''''''''''''''''''''''''''''  
        'INSERE NO RAZAO
        '''''''''''''''''''''''''''''''''''
        Try
            sql = "INSERT INTO Razao(Empresa_Id,EndEmpresa_Id,Conta_Id,Cliente_Id,EndCliente_Id,Produto,"
            sql &= "Movimento_Id,Lote_Id,Sequencia_Id,Titulo,Pedido,Custo,Indexador,DataMoeda,"
            sql &= "DebitoOficial,CreditoOficial,DebitoMoeda,CreditoMoeda,Historico,PrevistoRealizado) Values ('"
            sql &= CnpjEmpresa & "',"
            sql &= EndEmpresa & ",'"
            sql &= Conta & "','"
            sql &= Cliente_Id & "',"
            sql &= CInt(EndCliente_Id) & ",'"
            sql &= Produto & "','"
            sql &= Format(CDate(DataMovimento), "yyyy/MM/dd") & "',"
            sql &= CInt(Lote) & ","
            sql &= CInt(Sequencia) & ",'"
            sql &= Titulo & "',"
            sql &= 0 & ","
            If CentroDeCustos <> "" Then
                sql &= CInt(CentroDeCustos) & ","
            Else
                sql &= 0 & ","
            End If
            sql &= 3 & ",'"
            sql &= Format(CDate(DataMoeda), "yyyy/MM/dd") & "',"

            sql &= Replace(DebitoOficial, ",", ".") & ","
            sql &= Replace(CreditoOficial, ",", ".") & ","
            sql &= Replace(DebitoMoeda, ",", ".") & ","
            sql &= Replace(CreditoMoeda, ",", ".") & ",'"
            sql &= UCase(RTrim(Historico)) & "','"
            sql &= PrevistoRealizado & "')"

            Return sql
        Catch ex1 As Exception
            MsgBox(Me.Page, ex1.Message)
            Return False
        End Try
    End Function

    Public Overrides Sub Carregar(obj As [Lib].Negocio.IBaseEntity)
        If Session("objEmpresaCalcJuros" & HID.Value) IsNot Nothing Then
            Dim objEmpresa As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtCnpjDaEmpresa.Text = objEmpresa.Codigo
            txtEndEmpresa.Text = objEmpresa.CodigoEndereco
            txtNomeDaEmpresa.Text = objEmpresa.Nome
            Session.Remove("objEmpresaCalcJuros" & HID.Value)
        ElseIf Session("objClienteCalcJuros" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtCnpjDoCliente.Text = objCliente.Codigo
            txtEndDoCliente.Text = objCliente.CodigoEndereco
            txtNomeDoCliente.Text = objCliente.Nome
            Session.Remove("objClienteCalcJuros" & HID.Value)
        End If
    End Sub

End Class