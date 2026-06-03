Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PosicaoDeEncargos
    Inherits BasePage

#Region "Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Expedicao)
                If Funcoes.VerificaPermissao("PosicaoDeEncargos", "ACESSAR") Then
                    CarregarUnidade()
                    CarregarEncargos()
                    Limpar()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Protected Sub DdlUnidade_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(lstEmpresa, CarregarDDL.Tabela.Empresas, "", False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Protected Sub BtnCliente_Click(sender As Object, e As EventArgs) Handles BtnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Protected Sub lnkEmiteTodosEncargos_Click(sender As Object, e As EventArgs) Handles lnkEmiteTodosEncargos.Click
        Try
            If isValidField() Then
                If chkResumoEncargo.Checked Then
                    RelatorioResumo(True)
                Else
                    If ChkAgrupar.Checked Then
                        RelatorioPorClientes(True)
                    Else
                        RelatorioResumo(True)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Protected Sub GridOpcoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridOpcoes.SelectedIndexChanged
        Try
            If isValidField() Then
                If chkResumoEncargo.Checked Then
                    RelatorioResumo(False)
                Else
                    If ChkAgrupar.Checked Then
                        RelatorioPorClientes(False)
                    Else
                        RelatorioResumo(False)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PosicaoDeEncargos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
#End Region

#Region "Methods"
    Private Sub CarregarUnidade()
        ddl.Carregar(DdlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, lstEmpresa, False)
    End Sub
    Private Sub CarregarEncargos()
        Dim sql As String = "SELECT Encargo_id AS Codigo, Descricao" & vbCrLf & _
                            "  FROM Encargos" & vbCrLf & _
                            " WHERE OperacaoXEncargo = 'S'" & vbCrLf

        GridOpcoes.DataSource = Banco.ConsultaDataSet(sql, "Consulta")
        GridOpcoes.DataBind()
    End Sub
    Private Sub Limpar()
        txtDataInicial.Text = Now().ToString("01/MM/yyyy")
        txtDataFinal.Text = Funcoes.UltimoDiaDoMes
        chkResumoEncargo.Checked = False
        chkTodosEncargos.Checked = False
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            lstEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        Try
            If Session("objCliente" & HID.Value) IsNot Nothing Then
                Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objCliente" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Function isValidField()
        If DdlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a unidade de negócio.")
            Return False
        ElseIf lstEmpresa.GetSelectedValues.Count = 0 Then
            MsgBox(Me.Page, "Informe a Empresa.")
            Return False
        ElseIf txtDataInicial.Text = "" Then
            MsgBox(Me.Page, "Informe o período inicial.")
            Return False
        ElseIf txtDataFinal.Text = "" Then
            MsgBox(Me.Page, "Informe o período final.")
            Return False
        End If

        Return True
    End Function
    Private Function getParametrosConsulta() As String
        Dim parametros As String = String.Empty

        If lstEmpresa.GetSelectedValues.Count > 0 Then parametros &= "Empresa(s)" & String.Join("','", lstEmpresa.GetSelecteds()) & vbCrLf
        If txtCliente.Text.Length > 0 Then parametros &= "Cliente: " & txtCliente.Text & vbCrLf
        If IsDate(txtDataInicial.Text) AndAlso IsDate(txtDataFinal.Text) Then
            parametros &= "Período de: " & txtDataInicial.Text & " até " & txtDataFinal.Text & vbCrLf
        End If
        'parametros &= EntradaSaidaDescricao & vbCrLf

        parametros &= "Notas: " & IIf(RadEntradas.Checked, "Entradas", "Saídas") & vbCrLf
        If ChkAgrupar.Checked AndAlso chkResumoEncargo.Checked Then
            parametros &= "Resumo dos Encargos Agrupados Por Cliente" & vbCrLf
        ElseIf ChkAgrupar.Checked Then
            parametros &= "Agrupado Por Cliente" & vbCrLf
        End If

        Return parametros
    End Function
    Private Function getParameters() As Dictionary(Of String, Object)
        Dim Parameters As New Dictionary(Of String, Object)
        Dim objEmp As New Cliente(lstEmpresa.GetSelecteds(0).Split("-")(0), lstEmpresa.GetSelecteds(0).Split("-")(1))
        Parameters.Add("EmpresaCodigo", Funcoes.FormatarCpfCnpj(lstEmpresa.GetSelecteds(0).Split("-")(0)))
        Parameters.Add("EmpresaNome", objEmp.Nome)
        Parameters.Add("EmpresaCidade", objEmp.Cidade & " - " & objEmp.Estado.Descricao)
        Parameters.Add("EmpresaInscricao", "IE: " & objEmp.InscricaoEstadual)
        Parameters.Add("Parametros", getParametrosConsulta())
        Return Parameters
    End Function
    Private Sub RelatorioResumo(ByVal todosEncargo As Boolean)
        Try
            Dim strEncargo As String = ""
            If Not todosEncargo Then strEncargo = Server.HtmlDecode(GridOpcoes.SelectedRow.Cells(1).Text())

            Dim sql As String = "	SELECT Clientes.Cliente_Id as CodigoCliente, Clientes.Nome, NotasFiscaisXEncargos.Encargo_id as Encargo," & vbCrLf

            If Not ChkAgrupar.Checked Then
                sql &= "                  P.Nome as Produto,NotasFiscais.Nota_Id as Nota, NotasFiscais.Serie_Id as Serie, NotasFiscais.Movimento as Data," & vbCrLf
            End If

            sql &= "	   SUM(case" & vbCrLf & _
                  "				 when SO.EntradaSaida = 'E'" & vbCrLf & _
                  "				   then NotasFiscaisXItens.PesoFiscal" & vbCrLf & _
                  "				   else 0" & vbCrLf & _
                  "			   end" & vbCrLf & _
                  "			   ) AS QuantidadeNormal, " & vbCrLf & _
                  "		   SUM(case" & vbCrLf & _
                  "				 When SO.EntradaSaida = 'E'" & vbCrLf & _
                  "				   then NotasFiscaisXItens.Valor" & vbCrLf & _
                  "				   else 0" & vbCrLf & _
                  "				end" & vbCrLf & _
                  "			   ) AS ValorDaNotaNormal, " & vbCrLf & _
                  "		   SUM(case " & vbCrLf & _
                  "				 When SO.EntradaSaida = 'E'" & vbCrLf & _
                  "				   then NotasFiscaisXEncargos.Base" & vbCrLf & _
                  "				   else 0" & vbCrLf & _
                  "				end " & vbCrLf & _
                  "			   ) AS BaseDeCalculoNormal, " & vbCrLf & _
                  "		   SUM(case" & vbCrLf & _
                  "				 when SO.EntradaSaida = 'E'" & vbCrLf & _
                  "				   then NotasFiscaisXEncargos.Valor" & vbCrLf & _
                  "				   else 0" & vbCrLf & _
                  "				end" & vbCrLf & _
                  "				) AS ValorDoEncargoNormal," & vbCrLf & _
                  "		   SUM(case" & vbCrLf & _
                  "				 when SO.EntradaSaida = 'S'" & vbCrLf & _
                  "				   then NotasFiscaisXItens.PesoFiscal" & vbCrLf & _
                  "				   else 0" & vbCrLf & _
                  "				end" & vbCrLf & _
                  "			   ) AS QuantidadeDevolucao, " & vbCrLf & _
                  "		   SUM(case" & vbCrLf & _
                  "				 When SO.EntradaSaida = 'S'" & vbCrLf & _
                  "				   then NotasFiscaisXItens.Valor" & vbCrLf & _
                  "				   else 0" & vbCrLf & _
                  "				end" & vbCrLf & _
                  "			   ) AS ValorDaNotaDevolucao," & vbCrLf & _
                  "		   SUM(case" & vbCrLf & _
                  "				 When SO.EntradaSaida = 'S'" & vbCrLf & _
                  "				   then NotasFiscaisXEncargos.Base" & vbCrLf & _
                  "				   else 0" & vbCrLf & _
                  "				end" & vbCrLf & _
                  "			   ) AS BaseDeCalculoDevolucao, " & vbCrLf & _
                  "		   SUM(case" & vbCrLf & _
                  "				 when SO.EntradaSaida = 'S'" & vbCrLf & _
                  "				   then NotasFiscaisXEncargos.Valor" & vbCrLf & _
                  "				   else 0" & vbCrLf & _
                  "				end" & vbCrLf & _
                  "				) AS ValorDoEncargoDevolucao, " & vbCrLf & _
                  "		   SUM(case" & vbCrLf & _
                  "				 when SO.EntradaSaida = 'E'" & vbCrLf & _
                  "				   then NotasFiscaisXEncargos.Valor" & vbCrLf & _
                  "				   else NotasFiscaisXEncargos.Valor * - 1" & vbCrLf & _
                  "				end" & vbCrLf & _
                  "				) AS ValorDoEncargoTotal" & vbCrLf & _
                  "   into #Encargo" & vbCrLf & _
                  "	  FROM NotasFiscais " & vbCrLf & _
                  "	 INNER JOIN Clientes " & vbCrLf & _
                  "		ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf & _
                  "	   AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf & _
                  "  INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                  "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                  "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                  "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                  "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                  "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                  "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                  "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                  "  INNER JOIN NotasFiscaisXEncargos" & vbCrLf & _
                  "     ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf & _
                  "    AND NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf & _
                  "	 INNER JOIN SubOperacoes SO" & vbCrLf & _
                  "	    ON SO.Operacao_Id     = NotasFiscais.Operacao" & vbCrLf & _
                  "	   AND SO.SubOperacoes_Id = NotasFiscais.SubOperacao" & vbCrLf & _
                  "  INNER JOIN Produtos P" & vbCrLf & _
                  "     ON P.Produto_Id = NotasFiscaisXItens.Produto_Id" & vbCrLf & _
                  "	 Where NotasFiscais.Situacao   = 1" & vbCrLf

            sql &= "    AND NotasFiscais.Empresa_Id + '-' + cast(NotasFiscais.EndEmpresa_Id as varchar) in ('" & String.Join("','", lstEmpresa.GetSelecteds()) & "')" & vbCrLf
           
            sql &= "    AND NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                  "    AND NotasFiscaisXEncargos.Encargo_ID " & IIf(todosEncargo, "not in ('LIQUIDO', 'PRODUTO')", " = '" & strEncargo & "'") & vbCrLf

            If Not String.IsNullOrWhiteSpace(txtPercentual.Text) AndAlso CDec(txtPercentual.Text) > 0 Then
                sql &= "    AND NotasFiscaisXEncargos.Percentual = " & txtPercentual.Text.ToString.Replace(",", ".") & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                sql &= "AND NotasFiscais.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                       "AND NotasFiscais.EndCliente_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
            End If

            sql &= "	GROUP BY Clientes.Cliente_Id, Clientes.Nome," & vbCrLf

            If Not ChkAgrupar.Checked Then
                sql &= "         P.Nome, NotasFiscais.Movimento, NotasFiscais.Nota_Id, NotasFiscais.Serie_Id, NotasFiscaisXEncargos.Encargo_id" & vbCrLf
            Else
                sql &= "         NotasFiscaisXEncargos.Encargo_id" & vbCrLf
            End If

            sql &= "	ORDER BY Clientes.Nome" & IIf(Not ChkAgrupar.Checked, ",P.Nome, NotasFiscais.Movimento", "") & vbCrLf & _
                   " select CodigoCliente, Nome, Encargo, QuantidadeNormal, ValorDaNotaNormal, BaseDeCalculoNormal, ValorDoEncargoNormal,    " & vbCrLf

            If Not ChkAgrupar.Checked Then
                sql &= "    Produto, Nota, Serie, Data," & vbCrLf
            End If

            sql &= "        QuantidadeDevolucao, ValorDaNotaDevolucao, BaseDeCalculoDevolucao, ValorDoEncargoDevolucao, ValorDoEncargoTotal  " & vbCrLf & _
                   "   from #Encargo order by nome;                                                                                                           " & vbCrLf & _
                   "                                                                                                                         " & vbCrLf & _
                   " select " & IIf(Not ChkAgrupar.Checked, "CodigoCliente, Nome, ", "") & " Encargo, SUM(QuantidadeNormal) as QuantidadeNormal, SUM(ValorDaNotaNormal) as ValorDaNotaNormal, SUM(BaseDeCalculoNormal) as BaseDeCalculoNormal,                 " & vbCrLf & _
                   "        SUM(ValorDoEncargoNormal) as ValorDoEncargoNormal, SUM(QuantidadeDevolucao) as QuantidadeDevolucao, SUM(ValorDaNotaDevolucao) as ValorDaNotaDevolucao,            " & vbCrLf & _
                   "    	SUM(BaseDeCalculoDevolucao) as BaseDeCalculoDevolucao, SUM(ValorDoEncargoDevolucao) as ValorDoEncargoDevolucao, SUM(ValorDoEncargoTotal) as ValorDoEncargoTotal   " & vbCrLf & _
                   "   from #Encargo" & vbCrLf & _
                   "  group by " & IIf(Not ChkAgrupar.Checked, "CodigoCliente, Nome, ", "") & " Encargo" & vbCrLf & _
                   "  order by " & IIf(Not ChkAgrupar.Checked, "Nome, ", "") & " Encargo" & vbCrLf

            If Not ChkAgrupar.Checked Then
                sql &= " select Encargo, SUM(QuantidadeNormal) as QuantidadeNormal, SUM(ValorDaNotaNormal) as ValorDaNotaNormal, SUM(BaseDeCalculoNormal) as BaseDeCalculoNormal,                 " & vbCrLf & _
                   "        SUM(ValorDoEncargoNormal) as ValorDoEncargoNormal, SUM(QuantidadeDevolucao) as QuantidadeDevolucao, SUM(ValorDaNotaDevolucao) as ValorDaNotaDevolucao,            " & vbCrLf & _
                   "    	SUM(BaseDeCalculoDevolucao) as BaseDeCalculoDevolucao, SUM(ValorDoEncargoDevolucao) as ValorDoEncargoDevolucao, SUM(ValorDoEncargoTotal) as ValorDoEncargoTotal   " & vbCrLf & _
                   "   from #Encargo" & vbCrLf & _
                   "  group by Encargo" & vbCrLf & _
                   "  order by Encargo" & vbCrLf
            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "PosicaoDeEncargos")
            ds.Tables(0).TableName = "PosicaoDeEncargos"
            ds.Tables(1).TableName = "PosicaoDeEncargosResumo"
            If Not ChkAgrupar.Checked Then ds.Tables(2).TableName = "PosicaoDeEncargosResumoGeral"

            Dim Parameters As Dictionary(Of String, Object) = getParameters()
            Parameters.Add("Titulo", IIf(todosEncargo, "Resumo de Todos os Encargos.", "Resumo do Encargo : " & strEncargo))
            Parameters.Add("Todos", todosEncargo)

            Funcoes.BindReport(Me.Page, ds, IIf(ChkAgrupar.Checked, "Cr_PosicaoDeEncargos", "Cr_PosicaoDeEncargosResumoNota"), IIf(RadPDF.Checked, eExportType.PDF, eExportType.ExcelCrystal), Parameters, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Private Sub RelatorioPorClientes(ByVal todosEncargo As Boolean)
        Try
            Dim strEncargo As String = ""
            If Not todosEncargo Then strEncargo = Server.HtmlDecode(GridOpcoes.SelectedRow.Cells(1).Text())
            Dim sql As String = "SELECT CLientes.Cliente_Id as CodigoCliente, Clientes.Nome, NotasFiscaisXEncargos.Encargo_Id as Encargo," & vbCrLf & _
                  "       SUM(NotasFiscaisXItens.PesoFiscal) AS Quantidade," & vbCrLf & _
                  "       SUM(NotasFiscaisXItens.Valor) AS ValorDaNota," & vbCrLf & _
                  "       SUM(NotasFiscaisXEncargos.Base) AS BaseDeCalculo," & vbCrLf & _
                  "       SUM(NotasFiscaisXEncargos.Valor) AS ValorDoEncargo" & vbCrLf & _
                  "  into #Encargos" & vbCrLf & _
                  "  FROM NotasFiscais" & vbCrLf & _
                  " INNER JOIN Clientes" & vbCrLf & _
                  "    ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id" & vbCrLf & _
                  "   AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                  " INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                  "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                  "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                  "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                  "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                  "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                  "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                  "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                  " INNER JOIN NotasFiscaisXEncargos" & vbCrLf & _
                  "    ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf & _
                  "   AND NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf & _
                  " WHERE NotasFiscais.Situacao = 1" & vbCrLf


            sql &= "    AND NotasFiscais.Empresa_Id + '-' + cast(NotasFiscais.EndEmpresa_Id as varchar) in ('" & String.Join("','", lstEmpresa.GetSelecteds()) & "')" & vbCrLf
           
            sql &= "   AND NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                  "   AND NotasFiscaisXEncargos.Encargo_Id " & IIf(todosEncargo, "not in ('LIQUIDO', 'PRODUTO')", " = '" & strEncargo & "'") & vbCrLf & _
                  "   AND NotasFiscais.EntradaSaida_Id = '" & IIf(RadEntradas.Checked, "E", "S") & "'" & vbCrLf

            If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                sql &= "AND NotasFiscais.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                       "AND NotasFiscais.EndCliente_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
            End If

            sql &= " GROUP BY Clientes.Cliente_Id, Clientes.Nome, NotasFiscaisXEncargos.Encargo_Id" & vbCrLf & _
                   " ORDER BY Clientes.Nome, NotasFiscaisXEncargos.Encargo_Id" & vbCrLf & _
                   " select CodigoCliente, Nome, Encargo, Quantidade, ValorDaNota, BaseDeCalculo, ValorDoEncargo    " & vbCrLf & _
                   "   from #Encargos                                                                               " & vbCrLf & _
                   " select Encargo, Sum(Quantidade) as Quantidade, sum(ValorDaNota) as ValorDaNota,                " & vbCrLf & _
                   "        sum(BaseDeCalculo) as BaseDeCalculo, sum(ValorDoEncargo) as ValorDoEncargo              " & vbCrLf & _
                   "   from #Encargos                                                                               " & vbCrLf & _
                   "  group by Encargo                                                                              " & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "PosicaoDeEncargosPorClientes")
            ds.Tables(0).TableName = "PosicaoDeEncargosPorClientes"
            ds.Tables(1).TableName = "PosicaoDeEncargosResumo"

            Dim Parameters As Dictionary(Of String, Object) = getParameters()
            Parameters.Add("Titulo", "Posição de " & IIf(todosEncargo, "Encargos", strEncargo) & " Por Cliente")
            Parameters.Add("TodosEncargo", todosEncargo)

            Funcoes.BindReport(Me.Page, ds, "Cr_PosicaoDeEncargosPorClientes", IIf(RadPDF.Checked, eExportType.PDF, eExportType.ExcelCrystal), Parameters, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Removido pois ficou redundante, ja existe um relatorio nesta tela que tras o resumo da nota,
    'Foi comentado para verificação.
    '
    'Private Sub RelatorioPorNotas()
    '    Try
    '        Dim sql As String = "SELECT NotasFiscais.Movimento AS Data, NotasFiscais.Nota_Id AS Nota," & vbCrLf & _
    '              "       NotasFiscais.Serie_Id AS Serie, Clientes.Nome, Clientes.Cidade," & vbCrLf & _
    '              "       SUM(NotasFiscaisXItens.PesoFiscal) AS Quantidade," & vbCrLf & _
    '              "       SUM(NotasFiscaisXItens.Valor) AS ValorDaNota," & vbCrLf & _
    '              "       SUM(NotasFiscaisXEncargos.Base) AS BaseDeCalculo," & vbCrLf & _
    '              "       SUM(NotasFiscaisXEncargos.Valor) AS ValorDoEncargo," & vbCrLf & _
    '              "       NotasFiscais.Operacao, NotasFiscais.SubOperacao, NotasFiscaisXItens.Produto_Id AS Produto," & vbCrLf & _
    '              "       Produtos.Nome AS NomeDoProduto, NotasFiscais.Pedido, Clientes.Estado" & vbCrLf & _
    '              "  FROM NotasFiscais" & vbCrLf & _
    '              " INNER JOIN Clientes" & vbCrLf & _
    '              "    ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf & _
    '              "   AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
    '              " INNER JOIN NotasFiscaisXItens" & vbCrLf & _
    '              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
    '              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
    '              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
    '              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
    '              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
    '              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
    '              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
    '              " INNER JOIN Produtos" & vbCrLf & _
    '              "    ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
    '              " INNER JOIN NotasFiscaisXEncargos" & vbCrLf & _
    '              "    ON NotasFiscaisXItens.Empresa_Id      = NotasFiscaisXEncargos.Empresa_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.EndEmpresa_Id   = NotasFiscaisXEncargos.EndEmpresa_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.Cliente_Id      = NotasFiscaisXEncargos.Cliente_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.EndCliente_Id   = NotasFiscaisXEncargos.EndCliente_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.Serie_Id        = NotasFiscaisXEncargos.Serie_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.Nota_Id         = NotasFiscaisXEncargos.Nota_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.Produto_Id      = NotasFiscaisXEncargos.Produto_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.CFOP_Id         = NotasFiscaisXEncargos.CFOP_Id" & vbCrLf & _
    '              "   AND NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id" & vbCrLf & _
    '              " WHERE NotasFiscais.Situacao = 1 " & vbCrLf & _
    '              "   AND NotasFiscais.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
    '              "   AND NotasFiscais.Movimento BETWEEN '" & Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd") & "' AND '" & Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd") & "'" & vbCrLf & _
    '              "   AND NotasFiscaisXEncargos.Encargo_Id = '" & Server.HtmlDecode(Server.HtmlDecode(GridOpcoes.SelectedRow.Cells(1).Text())) & "'" & vbCrLf & _
    '              "   AND NotasFiscais.EntradaSaida_Id = '" & IIf(RadEntradas.Checked, "E", "S") & "'" & vbCrLf

    '        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
    '            sql &= "AND NotasFiscais.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
    '                   "AND NotasFiscais.EndCliente_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
    '        End If

    '        sql &= " GROUP BY NotasFiscais.Movimento, NotasFiscais.Nota_Id, NotasFiscais.Serie_Id, Clientes.Nome, Clientes.Cidade, NotasFiscais.Operacao, NotasFiscais.SubOperacao," & vbCrLf & _
    '               "          NotasFiscaisXItens.Produto_Id, Produtos.Nome, NotasFiscais.Pedido, Clientes.Estado" & vbCrLf & _
    '               " ORDER BY Data, Clientes.Nome, Nota"
    '        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "PosicaoDeEncargosPorNotas")

    '        Dim Parameters As Dictionary(Of String, Object) = getParameters()
    '        Parameters.Add("Titulo", "Posição de " & Server.HtmlDecode(GridOpcoes.SelectedRow.Cells(1).Text()) & " Por Nota")

    '        Funcoes.BindReport(Me.Page, ds, "Cr_PosicaoDeEncargosPorNotas", IIf(RadPDF.Checked, eExportType.PDF, eExportType.ExcelCrystal), Parameters, True)
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub
#End Region

End Class