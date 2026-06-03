Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports System.Xml
Imports System.Security
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Pkcs
Imports System.Security.Cryptography.Xml
Imports System.Security.Cryptography.Xml.SignedXml
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports System.Drawing

Partial Class SpedReinf
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("SpedFiscal", "ACESSAR") Then
                    If Not Directory.Exists(Server.MapPath("~/SpedReinf")) Then Directory.CreateDirectory(Server.MapPath("~/SpedReinf"))
                    CargaUnidade()
                    VerificaUnidade()

                    Limpar()

                    If Today.Month = 1 Then
                        txtDataInicial.Text = "01/" & Today.AddMonths(-1).ToString("MM") & "/" & Today.AddYears(-1).ToString("yyyy")
                        txtDataFinal.Text = Date.DaysInMonth(Today.Year, Today.AddMonths(-1).Month) & "/" & Today.AddMonths(-1).ToString("MM") & "/" & Today.AddYears(-1).ToString("yyyy")
                    Else
                        txtDataInicial.Text = "01/" & Today.AddMonths(-1).ToString("MM") & "/" & Today.ToString("yyyy")
                        txtDataFinal.Text = Date.DaysInMonth(Today.Year, Today.AddMonths(-1).Month) & "/" & Today.AddMonths(-1).ToString("MM") & "/" & Today.ToString("yyyy")
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Fiscal.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa, True)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Download(ByVal arquivo As String)
        Try
            Response.ContentType = "text/plain"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & arquivo)
            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(Server.MapPath("~/SpedReinf/" & arquivo), FileMode.Open, FileAccess.Read)
                'download = New FileStream(arquivo, FileMode.Open, FileAccess.Read)
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
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()

        If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "KAYNÃ" Then
            btn1000.Visible = True
            btn2010.Visible = True
            btn2040.Visible = True
            btn2055.Visible = True
            btn2099.Visible = True
            btn2098.Visible = True
            btn4010.Visible = True
            btn4020.Visible = True
            btn4099.Visible = True
            btnReat2010.Visible = True
            btnReat2040.Visible = True
            btnReat2055.Visible = True
            btnReat4020.Visible = True
            radReab4099.Visible = True
        Else
            btn1000.Visible = False
            btn2010.Visible = False
            btn2040.Visible = False
            btn2055.Visible = False
            btn2099.Visible = False
            btn2098.Visible = False
            btn4010.Visible = False
            btn4020.Visible = False
            btn4099.Visible = False
            btnReat2010.Visible = False
            btnReat2040.Visible = False
            btnReat2055.Visible = False
            btnReat4020.Visible = False
            radReab4099.Visible = False
        End If

        radReab4099.Checked = False

        btnArquivo1000.Visible = True
        txtArquivo1000.Text = String.Empty

        btnArquivo2010.Visible = False
        txtRecib2010.Visible = False
        txtseq2010.Visible = False
        txtArquivo2010.Text = String.Empty

        btnArquivo2040.Visible = False
        txtRecib2040.Visible = False
        txtseq2040.Visible = False
        txtArquivo2040.Text = String.Empty

        btnArquivo2055.Visible = False
        txtRecib2055.Visible = False
        txtseq2055.Visible = False
        txtArquivo2055.Text = String.Empty

        btnArquivo2099.Visible = False
        txtArquivo2099.Text = String.Empty

        btnArquivo4010.Visible = False
        txtArquivo4010.Text = String.Empty

        btnArquivo4020.Visible = False
        txtRecib4020.Visible = False
        txtseq4020.Visible = False
        txtArquivo4020.Text = String.Empty

        btnArquivo4099.Visible = False
        txtArquivo4099.Text = String.Empty

        txtRecib2055.Visible = False

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            MsgBox(Me.Page, "Informe a unidade de negocio.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Informe o periodo inicial.")
            Return False
        ElseIf Not IsDate(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Valor informado no período inicial não é uma data válida.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período final.")
            Return False
        ElseIf Not IsDate(txtDataInicial.Text) Then
            MsgBox(Me.Page, "Valor informado no período final não é uma data válida.")
            Return False
        ElseIf CDate(txtDataInicial.Text).Year <> CDate(txtDataInicial.Text).Year Then
            MsgBox(Me.Page, "Ano não pode ser diferente.")
            Return False
        ElseIf CDate(txtDataInicial.Text).Month <> CDate(txtDataInicial.Text).Month Then
            MsgBox(Me.Page, "Mês não pode ser diferente.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Function Relatorio2055(ByVal codEmpresa As String) As DataSet
        Dim ds2055 As DataSet
        Dim Sql As String = String.Empty

        Sql = "SELECT nXOri.Empresa_Id AS Empresa, nXOri.Cliente_Id AS Cliente, c.Nome, nXOri.EntradaSaida_Id AS ES, nXOri.Nota_Id AS Nota, nXOri.Movimento," & vbCrLf &
        "		niXOri.Produto_Id AS Produto, prd.Nome as NomeProduto, neXOri.Valor, neXOriF.Valor as VlrFunrural," & vbCrLf &
        "		isnull(neXOriS.Valor,0) AS ValorSenar, " & vbCrLf &
        "		nDev.Nota_Id AS NotaDevolucao, nDev.Movimento AS DataDevolucao, nDev.Valor AS ValorDevolucao, " & vbCrLf &
        "		round(((nDev.PercentualFunrural * nDev.Valor)/100),2) as ValorDevolucaoFunrural, round(((nDev.PercenturalSenar * nDev.Valor)/100),2) as ValorDevolucaoSenar" & vbCrLf &
        "into #Notas" & vbCrLf &
        "FROM   NotasFiscais n" & vbCrLf &
        "	inner join NotasXNotas nXn" & vbCrLf &
        "			ON nXn.Empresa_Id             = n.Empresa_Id" & vbCrLf &
        "		 	and nXn.EndEmpresa_iD         = n.EndEmpresa_Id" & vbCrLf &
        "			and nXn.Cliente_Id            = n.Cliente_Id" & vbCrLf &
        "			and nXn.EndCliente_Id         = n.EndCliente_Id" & vbCrLf &
        "			and nXn.OrigemEntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
        "			and nXn.OrigemSerie_Id        = n.Serie_Id" & vbCrLf &
        "			and nXn.OrigemNota_Id         = n.Nota_Id" & vbCrLf &
        "	inner join NotasFiscais nXOri" & vbCrLf &
        "			ON nXOri.Empresa_Id       = nXn.Empresa_Id" & vbCrLf &
        "		 	and nXOri.EndEmpresa_iD   = nXn.EndEmpresa_Id" & vbCrLf &
        "			and nXOri.Cliente_Id      = nXn.Cliente_Id" & vbCrLf &
        "			and nXOri.EndCliente_Id   = nXn.EndCliente_Id" & vbCrLf &
        "			and nXOri.EntradaSaida_Id = nXn.EntradaSaida_Id" & vbCrLf &
        "			and nXOri.Serie_Id        = nXn.Serie_Id" & vbCrLf &
        "			and nXOri.Nota_Id         = nXn.Nota_Id" & vbCrLf &
        "	inner join NotasFiscaisXItens niXOri" & vbCrLf &
        "			ON niXOri.Empresa_Id       = nXOri.Empresa_Id" & vbCrLf &
        "		 	and niXOri.EndEmpresa_iD   = nXOri.EndEmpresa_Id" & vbCrLf &
        "			and niXOri.Cliente_Id      = nXOri.Cliente_Id" & vbCrLf &
        "			and niXOri.EndCliente_Id   = nXOri.EndCliente_Id" & vbCrLf &
        "			and niXOri.EntradaSaida_Id = nXOri.EntradaSaida_Id" & vbCrLf &
        "			and niXOri.Serie_Id        = nXOri.Serie_Id" & vbCrLf &
        "			and niXOri.Nota_Id         = nXOri.Nota_Id" & vbCrLf &
        "	inner join NotasFiscaisXEncargos neXOri" & vbCrLf &
        "			ON neXOri.Empresa_Id       = niXOri.Empresa_Id" & vbCrLf &
        "		 	and neXOri.EndEmpresa_iD   = niXOri.EndEmpresa_Id" & vbCrLf &
        "			and neXOri.Cliente_Id      = niXOri.Cliente_Id" & vbCrLf &
        "			and neXOri.EndCliente_Id   = niXOri.EndCliente_Id" & vbCrLf &
        "			and neXOri.EntradaSaida_Id = niXOri.EntradaSaida_Id" & vbCrLf &
        "			and neXOri.Serie_Id        = niXOri.Serie_Id" & vbCrLf &
        "			and neXOri.Nota_Id         = niXOri.Nota_Id" & vbCrLf &
        "			and neXOri.CFOP_Id         = niXOri.CFOP_Id" & vbCrLf &
        "			and neXOri.Produto_Id      = niXOri.Produto_Id" & vbCrLf &
        "			and neXOri.Sequencia_Id    = niXOri.Sequencia_Id" & vbCrLf &
        "			and neXOri.Encargo_Id      = 'PRODUTO'" & vbCrLf &
        "	inner join SubOperacoes soOri" & vbCrLf &
        "			ON soOri.Operacao_id       = nXOri.Operacao" & vbCrLf &
        "		 	AND soOri.SubOperacoes_Id  = nXOri.SubOperacao" & vbCrLf &
        "	left join NotasFiscaisXEncargos neXOriF" & vbCrLf &
        "			ON neXOriF.Empresa_Id       = niXOri.Empresa_Id" & vbCrLf &
        "		 	and neXOriF.EndEmpresa_iD   = niXOri.EndEmpresa_Id" & vbCrLf &
        "			and neXOriF.Cliente_Id      = niXOri.Cliente_Id" & vbCrLf &
        "			and neXOriF.EndCliente_Id   = niXOri.EndCliente_Id" & vbCrLf &
        "			and neXOriF.EntradaSaida_Id = niXOri.EntradaSaida_Id" & vbCrLf &
        "			and neXOriF.Serie_Id        = niXOri.Serie_Id" & vbCrLf &
        "			and neXOriF.Nota_Id         = niXOri.Nota_Id" & vbCrLf &
        "			and neXOriF.CFOP_Id         = niXOri.CFOP_Id" & vbCrLf &
        "			and neXOriF.Produto_Id      = niXOri.Produto_Id" & vbCrLf &
        "			and neXOriF.Sequencia_Id    = niXOri.Sequencia_Id" & vbCrLf &
        "			and neXOriF.Encargo_Id      = 'FUNRURAL'" & vbCrLf &
        "	left join NotasFiscaisXEncargos neXOriS" & vbCrLf &
        "			ON neXOriS.Empresa_Id       = niXOri.Empresa_Id" & vbCrLf &
        "		 	and neXOriS.EndEmpresa_iD   = niXOri.EndEmpresa_Id" & vbCrLf &
        "			and neXOriS.Cliente_Id      = niXOri.Cliente_Id" & vbCrLf &
        "			and neXOriS.EndCliente_Id   = niXOri.EndCliente_Id" & vbCrLf &
        "			and neXOriS.EntradaSaida_Id = niXOri.EntradaSaida_Id" & vbCrLf &
        "			and neXOriS.Serie_Id        = niXOri.Serie_Id" & vbCrLf &
        "			and neXOriS.Nota_Id         = niXOri.Nota_Id" & vbCrLf &
        "			and neXOriS.CFOP_Id         = niXOri.CFOP_Id" & vbCrLf &
        "			and neXOriS.Produto_Id      = niXOri.Produto_Id" & vbCrLf &
        "			and neXOriS.Sequencia_Id    = niXOri.Sequencia_Id" & vbCrLf &
        "			and neXOriS.Encargo_Id      = 'SENAR'" & vbCrLf &
        "	left join (SELECT DEV.EmpresaDevolucao_Id, DEV.EndEmpresaDevolucao_Id, DEV.ClienteDevolucao_Id, DEV.EndClienteDevolucao_Id, " & vbCrLf &
        "						DEV.EntradaSaida_Id, DEV.Serie_Id, DEV.Nota_Id, DEV.CFOP_Id, DEV.Produto_Id, nDevNota.Movimento," & vbCrLf &
        "						DEV.Valor, nDevF.Percentual AS PercentualFunrural, nDevS.Percentual AS PercenturalSenar" & vbCrLf &
        "				FROM NotaFiscalDevolucaoXNotaFiscal DEV" & vbCrLf &
        "						inner join NotasFiscais nDevNota" & vbCrLf &
        "								ON nDevNota.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
        "		 						and nDevNota.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
        "								and nDevNota.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
        "								and nDevNota.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
        "								and nDevNota.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
        "								and nDevNota.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
        "								and nDevNota.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
        "						inner join NotasFiscaisXEncargos nDevP" & vbCrLf &
        "								ON nDevP.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
        "		 						and nDevP.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
        "								and nDevP.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
        "								and nDevP.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
        "								and nDevP.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
        "								and nDevP.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
        "								and nDevP.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
        "								and nDevP.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
        "								and nDevP.Produto_Id      = DEV.Produto_Id" & vbCrLf &
        "								and nDevP.Encargo_Id      = 'PRODUTO'" & vbCrLf &
        "						left join NotasFiscaisXEncargos nDevF" & vbCrLf &
        "								ON nDevF.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
        "		 						and nDevF.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
        "								and nDevF.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
        "								and nDevF.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
        "								and nDevF.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
        "								and nDevF.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
        "								and nDevF.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
        "								and nDevF.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
        "								and nDevF.Produto_Id      = DEV.Produto_Id" & vbCrLf &
        "								and nDevF.Encargo_Id      = 'FUNRURAL'" & vbCrLf &
        "						left join NotasFiscaisXEncargos nDevS" & vbCrLf &
        "								ON nDevS.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
        "		 						and nDevS.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
        "								and nDevS.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
        "								and nDevS.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
        "								and nDevS.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
        "								and nDevS.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
        "								and nDevS.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
        "								and nDevS.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
        "								and nDevS.Produto_Id      = DEV.Produto_Id" & vbCrLf &
        "								and nDevS.Encargo_Id      = 'SENAR'" & vbCrLf &
        "						WHERE nDevNota.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "') as nDev" & vbCrLf &
        "			ON nDev.EmpresaDevolucao_Id     = niXOri.Empresa_Id" & vbCrLf &
        "		 	and nDev.EndEmpresaDevolucao_Id = niXOri.EndEmpresa_Id" & vbCrLf &
        "			and nDev.ClienteDevolucao_Id    = niXOri.Cliente_Id" & vbCrLf &
        "			and nDev.EndClienteDevolucao_Id = niXOri.EndCliente_Id" & vbCrLf &
        "			and nDev.EntradaSaida_Id        = niXOri.EntradaSaida_Id" & vbCrLf &
        "			and nDev.Serie_Id               = niXOri.Serie_Id" & vbCrLf &
        "			and nDev.Nota_Id                = niXOri.Nota_Id" & vbCrLf &
        "			and nDev.CFOP_Id                = niXOri.CFOP_Id" & vbCrLf &
        "			and nDev.Produto_Id             = niXOri.Produto_Id" & vbCrLf &
        "	inner join Clientes c" & vbCrLf &
        "			on c.Cliente_Id   = nXOri.Cliente_Id" & vbCrLf &
        "			and c.Endereco_Id = nXOri.EndCliente_Id" & vbCrLf &
        "	inner join Produtos prd" & vbCrLf &
        "			on prd.Produto_Id = niXOri.Produto_Id" & vbCrLf &
        "WHERE nXOri.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'" & vbCrLf &
        "AND   LEFT(nXOri.Empresa_id,8) = '" & Left(codEmpresa, 8) & "'" & vbCrLf &
        "AND   nXOri.TipoDeDocumento = 1" & vbCrLf &
        "AND   n.TipoDeDocumento = 15" & vbCrLf &
        "AND   n.NFG = 1" & vbCrLf &
        "AND   nXOri.EntradaSaida_Id = 'E'" & vbCrLf &
        "AND   soOri.Classe in('COMPRAS', 'REAJUSTES')" & vbCrLf

        Sql &= "" & vbCrLf &
        "" & vbCrLf &
        "" & vbCrLf

        Sql &= "insert into #Notas" & vbCrLf &
                "SELECT n.Empresa_Id AS Empresa, n.Cliente_Id AS Cliente, c.Nome, n.EntradaSaida_Id AS ES, n.Nota_Id AS Nota, n.Movimento," & vbCrLf &
                "		ni.Produto_Id AS Produto, prd.Nome as NomeProduto, ne.Valor, neF.Valor as VlrFunrural," & vbCrLf &
                "		isnull(neS.Valor,0) AS ValorSenar," & vbCrLf &
                "		nDev.Nota_Id AS NotaDevolucao, nDev.Movimento AS DataDevolucao, nDev.Valor AS ValorDevolucao," & vbCrLf &
                "		round(((nDev.PercentualFunrural * nDev.Valor)/100),2) as ValorDevolucaoFunrural, round(((nDev.PercenturalSenar * nDev.Valor)/100),2) as ValorDevolucaoSenar" & vbCrLf &
                "FROM   NotasFiscais n" & vbCrLf &
                "	inner join NotasFiscaisXItens ni" & vbCrLf &
                "			ON ni.Empresa_Id       = n.Empresa_Id" & vbCrLf &
                "		 	and ni.EndEmpresa_iD   = n.EndEmpresa_Id" & vbCrLf &
                "			and ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                "			and ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                "			and ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                "			and ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                "			and ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                "	inner join NotasFiscaisXEncargos ne" & vbCrLf &
                "			ON ne.Empresa_Id       = ni.Empresa_Id" & vbCrLf &
                "		 	and ne.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf &
                "			and ne.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                "			and ne.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                "			and ne.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                "			and ne.Serie_Id        = ni.Serie_Id" & vbCrLf &
                "			and ne.Nota_Id         = ni.Nota_Id" & vbCrLf &
                "			and ne.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                "			and ne.Produto_Id      = ni.Produto_Id" & vbCrLf &
                "			and ne.Sequencia_Id    = ni.Sequencia_Id" & vbCrLf &
                "			and ne.Encargo_Id      = 'PRODUTO'" & vbCrLf &
                "	left join NotasFiscaisXEncargos neF" & vbCrLf &
                "			ON neF.Empresa_Id       = ni.Empresa_Id" & vbCrLf &
                "		 	and neF.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf &
                "			and neF.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                "			and neF.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                "			and neF.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                "			and neF.Serie_Id        = ni.Serie_Id" & vbCrLf &
                "			and neF.Nota_Id         = ni.Nota_Id" & vbCrLf &
                "			and neF.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                "			and neF.Produto_Id      = ni.Produto_Id" & vbCrLf &
                "			and neF.Sequencia_Id    = ni.Sequencia_Id" & vbCrLf &
                "			and neF.Encargo_Id      = 'FUNRURAL'" & vbCrLf &
                "	left join NotasFiscaisXEncargos neS" & vbCrLf &
                "			ON neS.Empresa_Id       = ni.Empresa_Id" & vbCrLf &
                "		 	and neS.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf &
                "			and neS.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                "			and neS.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                "			and neS.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                "			and neS.Serie_Id        = ni.Serie_Id" & vbCrLf &
                "			and neS.Nota_Id         = ni.Nota_Id" & vbCrLf &
                "			and neS.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                "			and neS.Produto_Id      = ni.Produto_Id" & vbCrLf &
                "			and neS.Sequencia_Id    = ni.Sequencia_Id" & vbCrLf &
                "			and neS.Encargo_Id      = 'SENAR'" & vbCrLf &
                "	inner join SubOperacoes so" & vbCrLf &
                "			ON so.Operacao_id       = n.Operacao" & vbCrLf &
                "		 	AND so.SubOperacoes_Id  = n.SubOperacao" & vbCrLf &
                "	left join (SELECT DEV.EmpresaDevolucao_Id, DEV.EndEmpresaDevolucao_Id, DEV.ClienteDevolucao_Id, DEV.EndClienteDevolucao_Id, " & vbCrLf &
                "						DEV.EntradaSaida_Id, DEV.Serie_Id, DEV.Nota_Id, DEV.CFOP_Id, DEV.Produto_Id, nDevNota.Movimento," & vbCrLf &
                "						DEV.Valor, nDevF.Percentual AS PercentualFunrural, nDevS.Percentual AS PercenturalSenar" & vbCrLf &
                "				FROM NotaFiscalDevolucaoXNotaFiscal DEV" & vbCrLf &
                "						inner join NotasFiscais nDevNota" & vbCrLf &
                "								ON nDevNota.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevNota.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevNota.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevNota.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevNota.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevNota.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevNota.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "						inner join NotasFiscaisXEncargos nDevP" & vbCrLf &
                "								ON nDevP.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevP.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevP.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevP.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevP.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevP.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevP.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevP.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevP.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevP.Encargo_Id      = 'PRODUTO'" & vbCrLf &
                "						left join NotasFiscaisXEncargos nDevF" & vbCrLf &
                "								ON nDevF.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevF.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevF.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevF.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevF.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevF.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevF.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevF.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevF.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevF.Encargo_Id      = 'FUNRURAL'" & vbCrLf &
                "						left join NotasFiscaisXEncargos nDevS" & vbCrLf &
                "								ON nDevS.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevS.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevS.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevS.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevS.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevS.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevS.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevS.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevS.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevS.Encargo_Id      = 'SENAR'" & vbCrLf &
                "						WHERE nDevNota.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "') as nDev" & vbCrLf &
                "			ON nDev.EmpresaDevolucao_Id     = ni.Empresa_Id" & vbCrLf &
                "		 	and nDev.EndEmpresaDevolucao_Id = ni.EndEmpresa_Id" & vbCrLf &
                "			and nDev.ClienteDevolucao_Id    = ni.Cliente_Id" & vbCrLf &
                "			and nDev.EndClienteDevolucao_Id = ni.EndCliente_Id" & vbCrLf &
                "			and nDev.EntradaSaida_Id        = ni.EntradaSaida_Id" & vbCrLf &
                "			and nDev.Serie_Id               = ni.Serie_Id" & vbCrLf &
                "			and nDev.Nota_Id                = ni.Nota_Id" & vbCrLf &
                "			and nDev.CFOP_Id                = ni.CFOP_Id" & vbCrLf &
                "			and nDev.Produto_Id             = ni.Produto_Id" & vbCrLf &
                "	inner join Clientes c" & vbCrLf &
                "			on c.Cliente_Id   = n.Cliente_Id" & vbCrLf &
                "			and c.Endereco_Id = n.EndCliente_Id" & vbCrLf &
                "	inner join Produtos prd" & vbCrLf &
                "			on prd.Produto_Id = ni.Produto_Id" & vbCrLf &
                "WHERE n.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'" & vbCrLf &
                "AND   LEFT(n.Empresa_id,8) = '" & Left(codEmpresa, 8) & "'" & vbCrLf &
                "AND   n.TipoDeDocumento = 1" & vbCrLf &
                "AND   n.NFG = 0" & vbCrLf &
                "AND   n.EntradaSaida_Id = 'E'" & vbCrLf &
                "AND   n.Eletronica = 'S'" & vbCrLf &
                "AND   n.NossaEmissao = 'N'" & vbCrLf &
                "AND   so.Classe in('COMPRAS', 'REAJUSTES')" & vbCrLf &
                "AND   so.Devolucao = 'N'" & vbCrLf &
                "AND   c.Categoria in(1,2,3)" & vbCrLf

        Sql &= "select Empresa, Cliente, Nome, ES, Nota, Movimento, Produto, NomeProduto, Valor, VlrFunrural, ValorSenar," & vbCrLf &
                "NotaDevolucao, DataDevolucao, ValorDevolucao, ValorDevolucaoFunrural, ValorDevolucaoSenar" & vbCrLf &
                "from #Notas" & vbCrLf &
                "where isnull(VlrFunrural,0) > 0 or isnull(ValorSenar,0) > 0"

        ds2055 = Banco.ConsultaDataSet(Sql, "Consulta")

        Return ds2055

    End Function


    Protected Sub lnkExcel2010_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel2010.Click
        Try
            EmitirRelatorio2010()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel2040_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel2040.Click
        Try
            MsgBox(Me.Page, "EM DESENVOLVIMENTO", eTitulo.Erro)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel2055_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel2055.Click
        Try
            EmitirRelatorio2055()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub llnkExcel4010_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel4010.Click
        Try
            EmitirRelatorio4000("F")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel4020_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel4020.Click
        Try
            EmitirRelatorio4000("J")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Private Sub EmitirRelatorio4000(ByVal tipoPessoa As String)
        'Dim Periodo As String = "PERÍODO: " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text))

        Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))

        Try
            Dim ds As DataSet = consulta4000(objEmpresa.Codigo, True, tipoPessoa)

            If ds IsNot Nothing AndAlso Not ds.Tables(0).Rows.Count > 0 Then
                MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

                Exit Sub
            End If

            If ds IsNot Nothing AndAlso Not ds.Tables(1).Rows.Count > 0 Then
                MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    'criando planilha 
                    Dim worksheet As ExcelWorksheet

                    'criando título da aba 
                    If tipoPessoa = "F" Then
                        worksheet = package.Workbook.Worksheets.Add("NOTAS R4010")
                    Else
                        worksheet = package.Workbook.Worksheets.Add("NOTAS R4020")
                    End If

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

                    If tipoPessoa = "F" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "NOTAS FISCAIS R4010")
                    Else
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "NOTAS FISCAIS R4020")
                    End If
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    For Each col As DataColumn In ds.Tables(0).Columns
                        If col.ColumnName <> "Row" Then
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        End If
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'formatando células valores
                        worksheet.Cells(String.Format("I{0}:K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    rowIndex += 1

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    ''criando colunas de totalizadores "BASE"
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Formula = String.Format("=SUM(I6:I{0})", rowIndex - 1)

                    'criando colunas de totalizadores "VALOR IR"
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Formula = String.Format("=SUM(J6:J{0})", rowIndex - 1)

                    ''criando colunas de totalizadores "VALOR AGREGADO"
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Formula = String.Format("=SUM(K6:K{0})", rowIndex - 1)


                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)


                    '******************************************
                    '*** 2 - 'Criando Série 4000
                    '******************************************

                    'criando título da aba 
                    If tipoPessoa = "F" Then
                        worksheet = package.Workbook.Worksheets.Add("R4010")
                    Else
                        worksheet = package.Workbook.Worksheets.Add("R4020")
                    End If
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape

                    'criando linha com o cabeçalho da planilha
                    rowIndex = 1
                    columnIndex = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    If tipoPessoa = "F" Then
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "R4010")
                    Else
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "R4020")
                    End If
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables(1).Columns
                        If col.ColumnName = "Empresa_Id" Then
                            worksheet.Cells(rowIndex, 1).Value = "EMPRESA"
                            worksheet.Cells(rowIndex, 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 1).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Cliente_Id" Then
                            worksheet.Cells(rowIndex, 2).Value = "CLIENTE"
                            worksheet.Cells(rowIndex, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 2).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Nome" Then
                            worksheet.Cells(rowIndex, 3).Value = "NOME"
                            worksheet.Cells(rowIndex, 3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 3).Style.Font.Size = 12
                        ElseIf col.ColumnName = "CodigoNaturezaDeRendimento" Then
                            worksheet.Cells(rowIndex, 4).Value = "NATUREZA"
                            worksheet.Cells(rowIndex, 4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 4).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Movimento" Then
                            worksheet.Cells(rowIndex, 5).Value = "MOVIMENTO"
                            worksheet.Cells(rowIndex, 5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 5).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Base" Then
                            worksheet.Cells(rowIndex, 6).Value = "Base"
                            worksheet.Cells(rowIndex, 6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 6).Style.Font.Size = 12
                        ElseIf col.ColumnName = "ValorIR" Then
                            worksheet.Cells(rowIndex, 7).Value = "ValorIR"
                            worksheet.Cells(rowIndex, 7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 7).Style.Font.Size = 12
                        ElseIf col.ColumnName = "ValorAgregado" Then
                            worksheet.Cells(rowIndex, 8).Value = "ValorAgregado"
                            worksheet.Cells(rowIndex, 8).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 8).Style.Font.Size = 12
                        End If
                    Next


                    'criando auto filtro na planilha
                    worksheet.Cells("A5:C" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(1).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(1).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(1).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'formatando células valores
                        worksheet.Cells(String.Format("F{0}:H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(1).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    rowIndex += 1

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores "BASE"
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Formula = String.Format("=SUM(F6:F{0})", rowIndex - 1)

                    ''criando colunas de totalizadores "VALOR IR"
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Formula = String.Format("=SUM(G6:G{0})", rowIndex - 1)

                    ''criando colunas de totalizadores "VALOR AGREGADO"
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Formula = String.Format("=SUM(H6:H{0})", rowIndex - 1)

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function consulta4000(ByVal codEmpresa As String, ByVal comNome As Boolean, ByVal tipoPessoa As String) As DataSet
        Dim ds4000 As DataSet
        Dim Sql As String = String.Empty

        Sql = " SELECT NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id, C.Nome, " & vbCrLf &
              " NF.CodigoNaturezaDeRendimento," & vbCrLf &
              " ISNULL( NFxI.cfop_Id, 0) AS Cfop, " & vbCrLf &
              " NF.Serie_Id, NF.Nota_Id, NF.Movimento, " & vbCrLf &
              " SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('PRODUTO') THEN NFxE.Base       END, 0))           AS Base, " & vbCrLf &
              " SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('IRRF PF','IRRF PJ') THEN NFxE.Valor      END, 0)) AS ValorIR, " & vbCrLf &
              " SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('CSRF') THEN NFxE.Valor      END, 0))              AS ValorAgregado " & vbCrLf &
              " INTO #Notas" & vbCrLf &
              " FROM NotasFiscais AS NF " & vbCrLf &
             " INNER Join NotasFiscaisXItens AS NFxI " & vbCrLf &
                " ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf &
               " AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
               " AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
               " AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
               " AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
               " AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
               " AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
             " INNER Join NotasFiscaisXEncargos AS NFxE " & vbCrLf &
                " ON NFxI.Empresa_Id      = NFxE.Empresa_Id " & vbCrLf &
               " AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id " & vbCrLf &
               " AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf &
               " And NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf &
               " AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id " & vbCrLf &
               " And NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
               " AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf &
               " AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
               " AND NFxI.CFOP_Id         = NFxE.CFOP_Id " & vbCrLf &
               " AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id " & vbCrLf &
             " INNER JOIN OperacaoXEstado OE " & vbCrLf &
                " ON OE.Codigo_id = NFxI.OperacaoxEstado " & vbCrLf &
              "INNER JOIN Clientes AS C" & vbCrLf &
                 "ON C.Cliente_Id = NF.Cliente_Id" & vbCrLf & " WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
               " AND LEFT(Nf.Empresa_id,8)  = '" & Left(codEmpresa, 8) & "'" & vbCrLf &
               " AND (NF.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "')" & vbCrLf

        If tipoPessoa = "F" Then
            Sql &= "AND LEN(NF.Cliente_Id) = 11" & vbCrLf
        Else
            Sql &= "AND LEN(NF.Cliente_Id) > 11" & vbCrLf
        End If

        Sql &= " AND NF.CodigoNaturezaDeRendimento > 0" & vbCrLf &
             " GROUP BY NF.EntradaSaida_Id, NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id, C.Nome, NF.CodigoNaturezaDeRendimento, NF.Serie_Id, NF.Nota_Id, NF.Movimento, NF.DataDaNota, NFxI.cfop_Id, ISNULL(OE.STPISCOFINS, 0), NF.Empresa_Id " & vbCrLf &
            " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('IRRF PF','IRRF PJ','CSRF') THEN NFxE.Valor END, 0)) > 0 OR NF.CodigoNaturezaDeRendimento = 12001)" & vbCrLf &
             " ORDER BY Nf.Empresa_id, NF.DataDaNota" & vbCrLf

        Sql &= "" & vbCrLf &
                "" & vbCrLf

        Sql &= "SELECT Empresa_Id AS Empresa, Cliente_Id AS CodigoCliente, Nome, CodigoNaturezaDeRendimento AS CodigoNatureza, Cfop, Serie_Id AS Serie, Nota_Id AS Nota, Movimento, Base, ValorIR, ValorAgregado " & vbCrLf &
                "FROM  #Notas" & vbCrLf &
                "" & vbCrLf

        Sql &= "" & vbCrLf &
        "" & vbCrLf

        Sql &= " SELECT Empresa_Id, Cliente_Id, Nome, CodigoNaturezaDeRendimento, Movimento, sum(Base) as Base, sum(ValorIR) AS ValorIR, sum(ValorAgregado) AS ValorAgregado " & vbCrLf &
               "   FROM #Notas" & vbCrLf &
               "  GROUP BY Empresa_Id, Cliente_Id, Nome, CodigoNaturezaDeRendimento, Movimento" & vbCrLf

        Sql &= "" & vbCrLf &
                "" & vbCrLf

        Sql &= "SELECT Empresa_Id, Cliente_Id, Nome, CodigoNaturezaDeRendimento" & vbCrLf &
                "FROM #Notas" & vbCrLf &
                "GROUP BY Empresa_Id, Cliente_Id, Nome, CodigoNaturezaDeRendimento"

        ds4000 = Banco.ConsultaDataSet(Sql, "Consulta")

        Return ds4000
    End Function

    Private Sub EmitirRelatorio2010()
        Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))

        Try
            Dim ds As DataSet = consulta2010(objEmpresa.Codigo)

            If ds IsNot Nothing AndAlso Not ds.Tables(0).Rows.Count > 0 Then
                MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("R2010")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "NOTAS FISCAIS DE SERVIÇO")
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    For Each col As DataColumn In ds.Tables(0).Columns
                        If col.ColumnName = "Empresa_Id" Then
                            worksheet.Cells(rowIndex, 1).Value = "EMPRESA"
                            worksheet.Cells(rowIndex, 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 1).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Cliente_Id" Then
                            worksheet.Cells(rowIndex, 2).Value = "CLIENTE"
                            worksheet.Cells(rowIndex, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 2).Style.Font.Size = 12
                        ElseIf col.ColumnName = "EndCliente_Id" Then
                            worksheet.Cells(rowIndex, 3).Value = "END"
                            worksheet.Cells(rowIndex, 3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 3).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Cfop" Then
                            worksheet.Cells(rowIndex, 4).Value = "CFOP"
                            worksheet.Cells(rowIndex, 4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 4).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Serie_Id" Then
                            worksheet.Cells(rowIndex, 5).Value = "SERIE"
                            worksheet.Cells(rowIndex, 5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 5).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Nota_Id" Then
                            worksheet.Cells(rowIndex, 6).Value = "NOTA"
                            worksheet.Cells(rowIndex, 6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 6).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Movimento" Then
                            worksheet.Cells(rowIndex, 7).Value = "MOVIMENTO"
                            worksheet.Cells(rowIndex, 7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 7).Style.Font.Size = 12
                        ElseIf col.ColumnName = "DataDaNota" Then
                            worksheet.Cells(rowIndex, 8).Value = "DATA DA NOTA"
                            worksheet.Cells(rowIndex, 8).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 8).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Vl_Doc" Then
                            worksheet.Cells(rowIndex, 9).Value = "VALOR NOTA"
                            worksheet.Cells(rowIndex, 9).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 9).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Base_Inss" Then
                            worksheet.Cells(rowIndex, 10).Value = "BASE"
                            worksheet.Cells(rowIndex, 10).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 10).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Per_Inss" Then
                            worksheet.Cells(rowIndex, 11).Value = "%"
                            worksheet.Cells(rowIndex, 11).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 11).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Vl_Inss" Then
                            worksheet.Cells(rowIndex, 12).Value = "VALOR INSS"
                            worksheet.Cells(rowIndex, 12).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 12).Style.Font.Size = 12
                        End If
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("I{0}:L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    rowIndex += 1
                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Formula = String.Format("=SUM(I6:I{0})", rowIndex - 1)

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Formula = String.Format("=SUM(J6:J{0})", rowIndex - 1)

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("L{0}", rowIndex)).Formula = String.Format("=SUM(L6:L{0})", rowIndex - 1)

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using

            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub EmitirRelatorio2040()

    End Sub

    Private Sub EmitirRelatorio2055()
        'Dim Periodo As String = "PERÍODO: " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text))

        Dim objEmpresa As New [Lib].Negocio.Cliente(DdlEmpresa.SelectedValue.ToString.Split("-")(0), DdlEmpresa.SelectedValue.ToString.Split("-")(1))

        Try
            Dim ds As DataSet = Relatorio2055(objEmpresa.Codigo)

            If ds IsNot Nothing AndAlso Not ds.Tables(0).Rows.Count > 0 Then
                MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

                Exit Sub
            End If

            Dim ds2 As DataSet = consulta2055(objEmpresa.Codigo, True)

            If ds2 IsNot Nothing AndAlso Not ds2.Tables(0).Rows.Count > 0 Then
                MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("NOTAS DE PRODUTOR")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "NOTAS FISCAIS DE PRODUTOR")
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    For Each col As DataColumn In ds.Tables(0).Columns
                        If col.ColumnName <> "Row" Then
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        End If
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("I{0}:K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                        worksheet.Cells(String.Format("N{0}:Q{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    rowIndex += 1

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("I{0}", rowIndex)).Formula = String.Format("=SUM(I6:I{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("J{0}", rowIndex)).Formula = String.Format("=SUM(J6:J{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("K{0}", rowIndex)).Formula = String.Format("=SUM(K6:K{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("N{0}", rowIndex)).Formula = String.Format("=SUM(N6:N{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("O{0}", rowIndex)).Formula = String.Format("=SUM(O6:O{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("P{0}", rowIndex)).Formula = String.Format("=SUM(P6:P{0})", rowIndex - 1)

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)


                    '******************************************
                    '*** 2 - 'Criando 2055
                    '******************************************
                    worksheet = package.Workbook.Worksheets.Add("R2055")
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape

                    'criando linha com o cabeçalho da planilha
                    rowIndex = 1
                    columnIndex = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "R2055 - NOTAS DE PRODUTOR")
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds2.Tables(0).Columns
                        If col.ColumnName = "Empresa_Id" Then
                            worksheet.Cells(rowIndex, 1).Value = "EMPRESA"
                            worksheet.Cells(rowIndex, 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 1).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Cliente_Id" Then
                            worksheet.Cells(rowIndex, 2).Value = "CLIENTE"
                            worksheet.Cells(rowIndex, 2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 2).Style.Font.Size = 12
                        ElseIf col.ColumnName = "Nome" Then
                            worksheet.Cells(rowIndex, 3).Value = "NOME"
                            worksheet.Cells(rowIndex, 3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                            worksheet.Cells(rowIndex, 3).Style.Font.Size = 12
                        ElseIf col.ColumnName = "SaldoValor" Then
                            worksheet.Cells(rowIndex, 4).Value = "VALOR"
                            worksheet.Cells(rowIndex, 4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 4).Style.Font.Size = 12
                        ElseIf col.ColumnName = "SaldoFunrural" Then
                            worksheet.Cells(rowIndex, 5).Value = "FUNRURAL"
                            worksheet.Cells(rowIndex, 5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 5).Style.Font.Size = 12
                        ElseIf col.ColumnName = "SaldoRat" Then
                            worksheet.Cells(rowIndex, 6).Value = "RAT"
                            worksheet.Cells(rowIndex, 6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 6).Style.Font.Size = 12
                        ElseIf col.ColumnName = "SaldoSenar" Then
                            worksheet.Cells(rowIndex, 7).Value = "SENAR"
                            worksheet.Cells(rowIndex, 7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                            worksheet.Cells(rowIndex, 7).Style.Font.Size = 12
                        End If
                    Next


                    'criando auto filtro na planilha
                    worksheet.Cells("A5:C" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds2.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds2.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds2.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("D{0}:G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds2.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    rowIndex += 1

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("C{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("C{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("C{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("D{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("D{0}", rowIndex)).Formula = String.Format("=SUM(D6:D{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("E{0}", rowIndex)).Formula = String.Format("=SUM(E6:E{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("F{0}", rowIndex)).Formula = String.Format("=SUM(F6:F{0})", rowIndex - 1)

                    ''criando colunas de totalizadores
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Formula = String.Format("=SUM(G6:G{0})", rowIndex - 1)

                    worksheet.Cells(String.Format("D{0}:G{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btn4099_Click(sender As Object, e As EventArgs) Handles btn4099.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R4099()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnArquivo4099_Click(sender As Object, e As EventArgs) Handles btnArquivo4099.Click
        Try
            Download(txtArquivo4099.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btn4020_Click(sender As Object, e As EventArgs) Handles btn4020.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R4020()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnArquivo4020_Click(sender As Object, e As EventArgs) Handles btnArquivo4020.Click
        Try
            Download(txtArquivo4020.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnReat4020_Click(sender As Object, e As EventArgs) Handles btnReat4020.Click
        txtRecib4020.Enabled = True
        txtseq4020.Visible = True
        txtRecib4020.Visible = True
        txtRecib4020.Parent.Visible = True
        txtRecib4020.Parent.Visible = True
    End Sub

    Protected Sub btn4010_Click(sender As Object, e As EventArgs) Handles btn4010.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R4010()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnArquivo4010_Click(sender As Object, e As EventArgs) Handles btnArquivo4010.Click
        Try
            Download(txtArquivo4010.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub btn2098_Click(sender As Object, e As EventArgs) Handles btn2098.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R2098()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub btn2099_Click(sender As Object, e As EventArgs) Handles btn2099.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R2099()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnArquivo2099_Click(sender As Object, e As EventArgs) Handles btnArquivo2099.Click
        Try
            Download(txtArquivo2099.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btn1000_Click(sender As Object, e As EventArgs) Handles btn1000.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R1000()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnArquivo1000_Click(sender As Object, e As EventArgs) Handles btnArquivo1000.Click
        Try
            Download(txtArquivo1000.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btn2010_Click(sender As Object, e As EventArgs) Handles btn2010.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R2010()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnArquivo2010_Click(sender As Object, e As EventArgs) Handles btnArquivo2010.Click
        Try
            Download(txtArquivo2010.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnReat2010_Click(sender As Object, e As EventArgs) Handles btnReat2010.Click
        txtRecib2010.Enabled = True
        txtseq2010.Visible = True
        txtRecib2010.Visible = True
        txtRecib2010.Parent.Visible = True
        txtseq2010.Parent.Visible = True
    End Sub

    Protected Sub btn2040_Click(sender As Object, e As EventArgs) Handles btn2040.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R2040()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnArquivo2040_Click(sender As Object, e As EventArgs) Handles btnArquivo2040.Click
        Try
            Download(txtArquivo2040.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnReat2040_Click(sender As Object, e As EventArgs) Handles btnReat2040.Click
        txtRecib2040.Enabled = True
        txtRecib2040.Visible = True
        txtseq2040.Visible = True
        txtRecib2040.Parent.Visible = True
        txtseq2040.Parent.Visible = True
    End Sub

    Protected Sub btn2055_Click(sender As Object, e As EventArgs) Handles btn2055.Click
        Try
            If Funcoes.VerificaPermissao("SpedFiscal", "LEITURA") Then
                If Validar() Then
                    R2055()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para Leitura")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnArquivo2055_Click(sender As Object, e As EventArgs) Handles btnArquivo2055.Click
        Try
            Download(txtArquivo2055.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub bntReat2055_Click(sender As Object, e As EventArgs) Handles btnReat2055.Click
        txtRecib2055.Enabled = True
        txtRecib2055.Visible = True
        txtseq2055.Visible = True
        txtRecib2055.Parent.Visible = True
        txtseq2055.Parent.Visible = True
    End Sub

    Private Sub R1000()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))
        Dim sequencia As Integer = 1

        txtArquivo1000.Text = "R1000-" & Empresa.Codigo & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

        If File.Exists(Server.MapPath("~/SpedReinf/" & txtArquivo1000.Text)) Then
            MsgBox(Me.Page, "Arquivo de abertura já existe.")
            Exit Sub
        End If

        Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo1000.Text)

        Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

        xtw.Formatting = Formatting.Indented

        xtw.WriteStartDocument()
        xtw.WriteStartElement("Reinf")

        xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evtInfoContribuinte/v2_01_02")


        xtw.WriteStartElement("evtInfoContri")

        Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

        xtw.WriteAttributeString("id", idEvento)

        xtw.WriteStartElement("ideEvento")
        xtw.WriteElementString("tpAmb", "1")   ' 1 - Produção    2 - Homologação
        xtw.WriteElementString("procEmi", "1")
        xtw.WriteElementString("verProc", "2.01.02")
        xtw.WriteEndElement()

        xtw.WriteStartElement("ideContri")
        xtw.WriteElementString("tpInsc", "1")
        xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
        xtw.WriteEndElement()

        xtw.WriteStartElement("infoContri")
        xtw.WriteStartElement("inclusao")
        xtw.WriteStartElement("idePeriodo")
        xtw.WriteElementString("iniValid", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
        xtw.WriteElementString("fimValid", Format(CDate(txtDataFinal.Text), "yyyy") & "-" & Format(CDate(txtDataFinal.Text), "MM"))
        xtw.WriteEndElement()

        xtw.WriteStartElement("infoCadastro")
        'Preencher com o código correspondente à classificação tributária do contribuinte, conforme tabela 8. 
        'Validação: Deve ser um dos códigos existentes na tabela 08. Os códigos [21] e [22] somente podem ser utilizados se {tpInsc} for igual a [2]. 
        'Para os demais códigos, {tpInsc} deve ser igual a [1].
        xtw.WriteElementString("classTrib", "99")
        '0 - Empresa Não obrigada à ECD; 1 - Empresa obrigada à ECD. Valores Válidos: 0, 1.
        xtw.WriteElementString("indEscrituracao", "1")
        '0 - Não Aplicável; 1 - Empresa enquadrada nos termos da Lei 12.546/2011 e alterações. Valores Válidos: 0, 1.
        xtw.WriteElementString("indDesoneracao", "0")
        'Indicativo da existência de acordo internacional para isenção de multa: 0 - Sem acordo; 1 - Com acordo. Validação: Só pode ser igual a [1] se {classTrib} for igual a [60]. Valores Válidos: 0, 1.
        xtw.WriteElementString("indAcordoIsenMulta", "0")
        'Indicativo da Situação da Pessoa Jurídica 0-Normal, 1-Extinção, 2-Fusão, 3-Cisão, 4-Incorporação
        xtw.WriteElementString("indSitPJ", "0")

        xtw.WriteStartElement("contato")

        If Left(Empresa.Codigo, 8) = "05366261" OrElse Left(Empresa.Codigo, 8) = "38198213" OrElse Left(Empresa.Codigo, 8) = "40938762" Then
            xtw.WriteElementString("nmCtt", "LUIZ HENRIQUE MULINARI TONETT")
            xtw.WriteElementString("cpfCtt", "04121606175")
        Else
            xtw.WriteElementString("nmCtt", Empresa.Empresa.NomeDoContador)
            xtw.WriteElementString("cpfCtt", Empresa.Empresa.CPFContador)
        End If

        xtw.WriteElementString("foneFixo", Empresa.Empresa.TelefoneContador)
        xtw.WriteEndElement()

        xtw.WriteEndElement()
        xtw.WriteEndElement()
        xtw.WriteEndElement()
        xtw.WriteEndElement()

        xtw.WriteEndElement()
        xtw.WriteEndDocument()
        'libera o XmlTextWriter
        xtw.Flush()
        'fechar o XmlTextWriter
        xtw.Close()
        'Termina aqui

        'btnArquivo1000.Visible = True

    End Sub

    Private Sub R1070()
        'Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        'Dim Empresa As New Cliente(emp(0), emp(1))


        'Response.Clear()
        'Response.Charset = "utf-8"
        'Dim strFilePath As String = Server.MapPath("~/SpedReinf/R1070-" & Empresa.Codigo & "-" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & "-" & "00001" & ".xml")
        'Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, Encoding.UTF8)
        ''A linha abaixo vai identar o c�digo, se n�o usar isso tudo ficar� em uma linha.
        'xtw.Formatting = Formatting.Indented
        ''Escreve a declara��o do documento <?xml version="1.0" encoding="utf-8"?>
        'xtw.WriteStartDocument()
        'xtw.WriteStartElement("Reinf")

        ''xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evtInfoContribuinte/v2_01_02")

        'xtw.WriteStartElement("evtTabProcesso")

        'Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & "00001"

        'xtw.WriteAttributeString("id", idEvento)

        'xtw.WriteStartElement("ideEvento")
        'xtw.WriteElementString("tpAmb", "2")
        'xtw.WriteElementString("procEmi", "1")
        'xtw.WriteElementString("verProc", "1.0.0")
        'xtw.WriteEndElement()

        'xtw.WriteStartElement("ideContri")
        'xtw.WriteElementString("tpInsc", "1")
        'xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
        'xtw.WriteEndElement()

        'xtw.WriteStartElement("infoProcesso")
        'xtw.WriteStartElement("inclusao")
        'xtw.WriteStartElement("ideProcesso")
        'xtw.WriteElementString("tpProc", "1")
        'xtw.WriteEndElement()
        'xtw.WriteEndElement()
        'xtw.WriteEndElement()

        'xtw.WriteEndElement()
        'xtw.WriteEndDocument()
        ''libera o XmlTextWriter
        'xtw.Flush()
        ''fechar o XmlTextWriter
        'xtw.Close()
        ''Termina aqui

    End Sub

    Private Sub R2010()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))

        Dim dsNotas As DataSet = consulta2010(Empresa.Codigo)
        Dim dsPrestador As DataSet = consulta2010PorCliente(Empresa.Codigo)

        If dsPrestador IsNot Nothing AndAlso Not dsPrestador.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

            Exit Sub
        End If

        Dim sequencia As Integer = 1

        For Each row In dsPrestador.Tables(0).Rows

            txtArquivo2010.Text = "R2010-" & Empresa.Codigo & "-" & row("Cliente_Id") & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

            If File.Exists(Server.MapPath("~/SpedReinf/" & txtArquivo2010.Text)) Then
                MsgBox(Me.Page, "Arquivo 2010 já existe.")
                Exit Sub
            End If

            Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo2010.Text)

            Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

            xtw.Formatting = Formatting.Indented

            xtw.WriteStartDocument()
            xtw.WriteStartElement("Reinf")

            xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evtTomadorServicos/v2_01_02")

            xtw.WriteStartElement("evtServTom")

            Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

            xtw.WriteAttributeString("id", idEvento)

            xtw.WriteStartElement("ideEvento")

            If Not String.IsNullOrWhiteSpace(txtRecib2010.Text) AndAlso (sequencia = txtseq2010.Text) Then 'Informe [1] para arquivo original ou [2] para arquivo de retificação.
                xtw.WriteElementString("indRetif", "2")
                xtw.WriteElementString("nrRecibo", txtRecib2010.Text)  'Preencher com o número do recibo "nrRecArqBase" do arquivo a ser retificado.
            Else
                xtw.WriteElementString("indRetif", "1")
            End If

            'Validação: O preenchimento é obrigatório se {indRetif} = [2]. Deve ser um recibo de entrega válido, correspondente ao arquivo objeto da retificação.
            xtw.WriteElementString("perApur", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
            xtw.WriteElementString("tpAmb", "1")  ' 1 - Produção    2 - Homologação
            xtw.WriteElementString("procEmi", "1")
            xtw.WriteElementString("verProc", "2.01.02")
            xtw.WriteEndElement()

            xtw.WriteStartElement("ideContri")
            xtw.WriteElementString("tpInsc", "1")
            xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
            xtw.WriteEndElement()

            xtw.WriteStartElement("infoServTom")

            xtw.WriteStartElement("ideEstabObra")
            xtw.WriteElementString("tpInscEstab", "1")
            xtw.WriteElementString("nrInscEstab", row("Empresa_Id"))
            'Indicativo de Prestação de Serviços em Obra de Construção Civil: 
            '0 - Não é obra de construção civil ou não está sujeita a matrícula de obra; 
            '1 - Obra de Construção Civil - Empreitada Total; 
            '2 - Obra de Construção Civil - Empreitada Parcial. 
            'Validação: Em arquivo gerado por Pessoa Física, deve ser igual a [1,2]. Valores Válidos: 0, 1, 2
            xtw.WriteElementString("indObra", "0")

            xtw.WriteStartElement("idePrestServ")
            xtw.WriteElementString("cnpjPrestador", row("Cliente_Id"))
            xtw.WriteElementString("vlrTotalBruto", row("Vl_Doc"))
            xtw.WriteElementString("vlrTotalBaseRet", row("Base_Inss"))
            xtw.WriteElementString("vlrTotalRetPrinc", row("Vl_Inss"))

            'Ajustado conforme conversar com Maurício Commodity - Furlan - 17/01/2025
            If CDec(row("Per_Inss")) = 11 Then
                xtw.WriteElementString("indCPRB", "0")  'Inss 11%
            Else
                xtw.WriteElementString("indCPRB", "1")  'Inss 3.5%
            End If

            For Each rowNotas In dsNotas.Tables(0).Rows
                If row("Empresa_Id") = rowNotas("Empresa_Id") AndAlso row("Cliente_Id") = rowNotas("Cliente_Id") Then
                    xtw.WriteStartElement("nfs")
                    xtw.WriteElementString("serie", rowNotas("Serie_Id"))
                    xtw.WriteElementString("numDocto", rowNotas("Nota_Id"))
                    xtw.WriteElementString("dtEmissaoNF", Format(rowNotas("DataDaNota"), "yyyy-MM-dd"))
                    xtw.WriteElementString("vlrBruto", rowNotas("Vl_Doc"))

                    xtw.WriteStartElement("infoTpServ")
                    'Informar o tipo de serviço, conforme tabela 6. Validação: O código informado deve existir na tabela 6.
                    xtw.WriteElementString("tpServico", "100000020")
                    xtw.WriteElementString("vlrBaseRet", rowNotas("Base_Inss"))
                    xtw.WriteElementString("vlrRetencao", rowNotas("Vl_Inss"))

                    xtw.WriteEndElement()
                    xtw.WriteEndElement()
                End If
            Next

            xtw.WriteEndElement()
            xtw.WriteEndElement()

            xtw.WriteEndElement()
            xtw.WriteEndElement()

            xtw.WriteEndElement()
            'xtw.WriteEndElement()
            'xtw.WriteEndElement()
            xtw.WriteEndDocument()
            'libera o XmlTextWriter
            xtw.Flush()
            'fechar o XmlTextWriter
            xtw.Close()
            'Termina aqui

            sequencia += 1

        Next

        'btnArquivo2010.Visible = True

    End Sub

    Public Function consulta2010(ByVal codEmpresa As String) As DataSet
        Dim ds2010 As DataSet
        Dim Sql As String = String.Empty

        'ACRESCENTADO CFOP 1933 COM INSS RETIDO SOLICITADO POR JÉSSICA NUTRI - FURLAN 02/06/2022

        Sql = " SELECT NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf &
            "        ISNULL( NFxI.cfop_Id, 0) AS Cfop, " & vbCrLf &
            "        NF.Serie_Id, NF.Nota_Id, NF.Movimento, NF.DataDaNota, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id in ('PRODUTO') THEN NFxE.Valor      END, 0)) AS Vl_Doc, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Base       END, 0)) AS Base_Inss, " & vbCrLf &
            "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Percentual END, 0)) AS Per_Inss, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Valor      END, 0)) AS Vl_Inss " & vbCrLf &
            "   FROM NotasFiscais AS NF " & vbCrLf &
            "  INNER JOIN NotasFiscaisXItens AS NFxI " & vbCrLf &
            "     ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf &
            "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
            "    AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
            "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
            "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
            "    AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
            "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
            "  INNER JOIN NotasFiscaisXEncargos AS NFxE " & vbCrLf &
            "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id " & vbCrLf &
            "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id " & vbCrLf &
            "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf &
            "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf &
            "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id " & vbCrLf &
            "    AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
            "    AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf &
            "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
            "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id " & vbCrLf &
            "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id " & vbCrLf &
            "   inner join OperacaoXEstado OE " & vbCrLf &
            "      on OE.Codigo_id = NFxI.OperacaoxEstado " & vbCrLf &
            "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
            "    AND LEFT(Nf.Empresa_id,8) = '" & Left(codEmpresa, 8) & "'" & vbCrLf &
            "    AND (NF.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "')" & vbCrLf &
            "    AND (NF.Serie_Id not in('D','F')) " & vbCrLf &
            "    AND LEN(NF.Cliente_Id) > 11 " & vbCrLf &
            "    AND ((NFxI.cfop_Id in (1556, 2556, 3556)) OR (NFxI.cfop_Id in (1933, 2933) AND NFxE.Encargo_Id IN ('INSS RETIDO','PRODUTO'))) " & vbCrLf &
            "  GROUP BY NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, NF.Movimento, NF.DataDaNota, NFxI.cfop_Id,  ISNULL(OE.STPISCOFINS, 0), NF.Empresa_Id " & vbCrLf &
            " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Valor END, 0)) > 0)" & vbCrLf &
            " ORDER BY Nf.Empresa_id,  NF.DataDaNota"

        ds2010 = Banco.ConsultaDataSet(Sql, "Consulta")

        Return ds2010
    End Function

    Public Function consulta2010PorCliente(ByVal codEmpresa As String) As DataSet
        Dim ds2010 As DataSet
        Dim Sql As String = String.Empty

        'ACRESCENTADO CFOP 1933 COM INSS RETIDO SOLICITADO POR JÉSSICA NUTRI - FURLAN 02/06/2022

        Sql = " SELECT NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id in ('PRODUTO') THEN NFxE.Valor      END, 0)) AS Vl_Doc, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Base       END, 0)) AS Base_Inss, " & vbCrLf &
            "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Percentual END, 0)) AS Per_Inss, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Valor      END, 0)) AS Vl_Inss " & vbCrLf &
            "   FROM NotasFiscais AS NF " & vbCrLf &
            "  INNER JOIN NotasFiscaisXItens AS NFxI " & vbCrLf &
            "     ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf &
            "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
            "    AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
            "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
            "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
            "    AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
            "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
            "  INNER JOIN NotasFiscaisXEncargos AS NFxE " & vbCrLf &
            "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id " & vbCrLf &
            "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id " & vbCrLf &
            "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf &
            "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf &
            "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id " & vbCrLf &
            "    AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
            "    AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf &
            "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
            "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id " & vbCrLf &
            "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id " & vbCrLf &
            "   inner join OperacaoXEstado OE " & vbCrLf &
            "      on OE.Codigo_id = NFxI.OperacaoxEstado " & vbCrLf &
            "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
            "    AND LEFT(Nf.Empresa_id,8) = '" & Left(codEmpresa, 8) & "'" & vbCrLf &
            "    AND (NF.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "')" & vbCrLf &
            "    AND (NF.Serie_Id not in('D','F')) " & vbCrLf &
            "    AND LEN(NF.Cliente_Id) > 11 " & vbCrLf &
            "    AND ((NFxI.cfop_Id in (1556, 2556, 3556)) OR (NFxI.cfop_Id in (1933, 2933) AND NFxE.Encargo_Id IN ('INSS RETIDO','PRODUTO'))) " & vbCrLf &
            "  GROUP BY NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id " & vbCrLf &
            " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Valor END, 0)) > 0)" & vbCrLf &
            " ORDER BY Nf.Empresa_id"

        ds2010 = Banco.ConsultaDataSet(Sql, "Consulta")

        Return ds2010
    End Function

    Private Sub R2040()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))
        'Dim Sql As String = String.Empty
        Dim LerEmp As String = String.Empty
        Dim LerCli As String = String.Empty
        Dim primeiraEmp As Boolean = True
        Dim primeiraCli As Boolean = True
        Dim ds As DataSet

        ds = consulta2040(Empresa.Codigo)

        If ds IsNot Nothing AndAlso Not ds.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

            Exit Sub
        End If

        Dim sequencia As Integer = 1

        For Each row In ds.Tables(0).Rows

            txtArquivo2040.Text = "R2040-" & row("Empresa_Id") & "-" & row("Cliente_Id") & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & ".xml"

            Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo2040.Text)

            Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

            xtw.Formatting = Formatting.Indented

            xtw.WriteStartDocument()
            xtw.WriteStartElement("Reinf")


            xtw.WriteStartElement("evtAssocDespRep")

            Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

            xtw.WriteAttributeString("id", idEvento)

            xtw.WriteStartElement("ideEvento")

            If Not String.IsNullOrWhiteSpace(txtRecib2040.Text) AndAlso (sequencia = txtseq2040.Text) Then 'Informe [1] para arquivo original ou [2] para arquivo de retificação.
                xtw.WriteElementString("indRetif", "2")
                xtw.WriteElementString("nrRecibo", txtRecib2040.Text)  'Preencher com o número do recibo "nrRecArqBase" do arquivo a ser retificado.
            Else
                xtw.WriteElementString("indRetif", "1")
            End If

            'Validação: O preenchimento é obrigatório se {indRetif} = [2]. Deve ser um recibo de entrega válido, correspondente ao arquivo objeto da retificação.
            xtw.WriteElementString("perApur", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
            xtw.WriteElementString("tpAmb", "1")   ' 1 - Produção    2 - Homologação
            xtw.WriteElementString("procEmi", "1")
            xtw.WriteElementString("verProc", "2.01.02")
            xtw.WriteEndElement()

            xtw.WriteStartElement("ideContri")
            xtw.WriteElementString("tpInsc", "1")
            xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))

            xtw.WriteStartElement("ideEstab")
            xtw.WriteElementString("tpInscEstab", "1")
            xtw.WriteElementString("nrInscEstab", row("Empresa_Id"))

            xtw.WriteStartElement("recursosRep")
            xtw.WriteElementString("cnpjAssocDesp", row("Cliente_Id"))

            xtw.WriteElementString("vlrTotalRep", row("Base_Inss"))
            xtw.WriteElementString("vlrTotalRet", row("Vl_Inss"))

            xtw.WriteStartElement("infoRecurso")
            xtw.WriteElementString("tpRepasse", 1)
            xtw.WriteElementString("descRecurso", "PATROCÍNIO")

            xtw.WriteElementString("vlrBruto", row("Base_Inss"))
            xtw.WriteElementString("vlrRetApur", row("Vl_Inss"))
            xtw.WriteEndElement()

            xtw.WriteEndElement()

            xtw.WriteEndElement()

            xtw.WriteEndElement()

            xtw.WriteEndElement()

            xtw.WriteEndDocument()
            'libera o XmlTextWriter
            xtw.Flush()
            'fechar o XmlTextWriter
            xtw.Close()
            'Termina aqui

            sequencia += 1
        Next

    End Sub

    Public Function consulta2040(ByVal codEmpresa As String) As DataSet
        Dim ds2040 As DataSet
        Dim Sql As String = String.Empty

        'ACRESCENTADO CFOP 1933 COM INSS RETIDO SOLICITADO POR JÉSSICA NUTRI - FURLAN 02/06/2022

        Sql = " SELECT NF.Empresa_Id, NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf &
            "        ISNULL( NFxI.cfop_Id, 0) AS Cfop, " & vbCrLf &
            "        NF.Serie_Id, NF.Nota_Id, NF.Movimento, NF.DataDaNota, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id in ('PRODUTO') THEN NFxE.Valor      END, 0)) AS Vl_Doc, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Base       END, 0)) AS Base_Inss, " & vbCrLf &
            "        MAX(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Percentual END, 0)) AS Per_Inss, " & vbCrLf &
            "        SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Valor      END, 0)) AS Vl_Inss " & vbCrLf &
            "   FROM NotasFiscais AS NF " & vbCrLf &
            "  INNER JOIN NotasFiscaisXItens AS NFxI " & vbCrLf &
            "     ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf &
            "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
            "    AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
            "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
            "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
            "    AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
            "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
            "  INNER JOIN NotasFiscaisXEncargos AS NFxE " & vbCrLf &
            "     ON NFxI.Empresa_Id      = NFxE.Empresa_Id " & vbCrLf &
            "    AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id " & vbCrLf &
            "    AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf &
            "    AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf &
            "    AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id " & vbCrLf &
            "    AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
            "    AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf &
            "    AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
            "    AND NFxI.CFOP_Id         = NFxE.CFOP_Id " & vbCrLf &
            "    AND NFxI.Sequencia_Id    = NFxE.Sequencia_Id " & vbCrLf &
            "   inner join OperacaoXEstado OE " & vbCrLf &
            "      on OE.Codigo_id = NFxI.OperacaoxEstado " & vbCrLf &
            "  WHERE (ISNULL(NF.Situacao, 1) = 1) " & vbCrLf &
            "    AND LEFT(Nf.Empresa_id,8) = '" & Left(codEmpresa, 8) & "'" & vbCrLf &
            "    AND (NF.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "')" & vbCrLf &
            "    AND LEN(NF.Cliente_Id) > 11 " & vbCrLf &
            "    AND ((NFxI.cfop_Id in (1949) AND NFxE.Encargo_Id IN ('INSS RETIDO','PRODUTO'))) " & vbCrLf &
            "  GROUP BY NF.EntradaSaida_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.Serie_Id, NF.Nota_Id, NF.Movimento, NF.DataDaNota, NFxI.cfop_Id,  ISNULL(OE.STPISCOFINS, 0), NF.Empresa_Id " & vbCrLf &
            " HAVING (SUM(ISNULL(CASE WHEN NFxE.Encargo_Id IN ('INSS','INSS RETIDO') THEN NFxE.Valor END, 0)) > 0)" & vbCrLf &
            " ORDER BY Nf.Empresa_id,  NF.DataDaNota"

        ds2040 = Banco.ConsultaDataSet(Sql, "Consulta")

        Return ds2040
    End Function

    Private Sub R2050()

    End Sub

    Private Sub R2055()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))
        'Dim Sql As String = String.Empty
        Dim LerEmp As String = String.Empty
        Dim LerCli As String = String.Empty
        Dim primeiraEmp As Boolean = True
        Dim primeiraCli As Boolean = True
        Dim ds As DataSet

        ds = consulta2055(Empresa.Codigo, False)

        If ds IsNot Nothing AndAlso Not ds.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

            Exit Sub
        End If

        Dim sequencia As Integer = 1

        For Each row In ds.Tables(0).Rows

            txtArquivo2055.Text = "R2055-" & row("Empresa_Id") & "-" & row("Cliente_Id") & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & ".xml"

            Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo2055.Text)

            Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

            xtw.Formatting = Formatting.Indented

            xtw.WriteStartDocument()
            xtw.WriteStartElement("Reinf")

            xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evt2055AquisicaoProdRural/v2_01_02")

            xtw.WriteStartElement("evtAqProd")

            Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

            xtw.WriteAttributeString("id", idEvento)

            xtw.WriteStartElement("ideEvento")

            If Not String.IsNullOrWhiteSpace(txtRecib2055.Text) AndAlso (sequencia = txtseq2055.Text) Then 'Informe [1] para arquivo original ou [2] para arquivo de retificação.
                xtw.WriteElementString("indRetif", "2")
                xtw.WriteElementString("nrRecibo", txtRecib2055.Text)  'Preencher com o número do recibo "nrRecArqBase" do arquivo a ser retificado.
            Else
                xtw.WriteElementString("indRetif", "1")
            End If

            'Validação: O preenchimento é obrigatório se {indRetif} = [2]. Deve ser um recibo de entrega válido, correspondente ao arquivo objeto da retificação.
            xtw.WriteElementString("perApur", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
            xtw.WriteElementString("tpAmb", "1")   ' 1 - Produção    2 - Homologação
            xtw.WriteElementString("procEmi", "1")
            xtw.WriteElementString("verProc", "2.01.02")
            xtw.WriteEndElement()

            xtw.WriteStartElement("ideContri")
            xtw.WriteElementString("tpInsc", "1")
            xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
            xtw.WriteEndElement()

            xtw.WriteStartElement("infoAquisProd")

            xtw.WriteStartElement("ideEstabAdquir")
            xtw.WriteElementString("tpInscAdq", "1")
            xtw.WriteElementString("nrInscAdq", row("Empresa_Id"))

            xtw.WriteStartElement("ideProdutor")
            If row("Cliente_Id").ToString.Length = 11 Then
                xtw.WriteElementString("tpInscProd", "2")
            ElseIf row("Cliente_Id").ToString.Length = 14 Then
                xtw.WriteElementString("tpInscProd", "1")
            Else
                MsgBox(Me.Page, "Verificar Cadastro do Cliente " & row("Cliente_Id").ToString & " - " & row("Nome"), eTitulo.Info)

                Exit Sub
            End If

            xtw.WriteElementString("nrInscProd", row("Cliente_Id"))

            If row("SaldoFunrural") = 0 Then xtw.WriteElementString("indOpcCP", "S") ' Remover pois se estiver S não declara o Funrural

            xtw.WriteStartElement("detAquis")
            xtw.WriteElementString("indAquis", "1")
            xtw.WriteElementString("vlrBruto", row("SaldoValor"))

            If row("SaldoFunrural") > 0 Then
                xtw.WriteElementString("vlrCPDescPR", Math.Round(row("SaldoFunrural"), 2, MidpointRounding.AwayFromZero))
                xtw.WriteElementString("vlrRatDescPR", Math.Round(row("SaldoRat"), 2, MidpointRounding.AwayFromZero))
            Else
                xtw.WriteElementString("vlrCPDescPR", "0,00")
                xtw.WriteElementString("vlrRatDescPR", "0,00")
            End If

            xtw.WriteElementString("vlrSenarDesc", Math.Round(row("SaldoSenar"), 2, MidpointRounding.AwayFromZero))
            xtw.WriteEndElement()

            xtw.WriteEndElement()

            xtw.WriteEndElement()

            xtw.WriteEndElement()

            xtw.WriteEndDocument()
            'libera o XmlTextWriter
            xtw.Flush()
            'fechar o XmlTextWriter
            xtw.Close()
            'Termina aqui

            sequencia += 1

        Next

    End Sub

    Public Function consulta2055(ByVal codEmpresa As String, ByVal comNome As Boolean) As DataSet
        Dim ds2055 As DataSet
        Dim Sql As String = String.Empty

        Sql = "SELECT nXOri.Empresa_Id, nXOri.EndEmpresa_Id, nXOri.Cliente_Id, nXOri.EndCliente_Id, c.Nome, nXOri.EntradaSaida_Id, nXOri.Nota_Id, nXOri.Movimento," & vbCrLf &
                "		niXOri.Produto_Id, neXOri.Valor, neXOriF.Valor as VlrFunruralOriginal," & vbCrLf &
                "		case" & vbCrLf &
                "			when (neXOriF.Percentual = 1.3 or neXOriF.Percentual = 1.5)" & vbCrLf &
                "				then 1.2" & vbCrLf &
                "				else neXOriF.Percentual" & vbCrLf &
                "			end as Percentual," & vbCrLf &
                "		isnull(neXOriS.Valor,0) AS ValorSenar, " & vbCrLf &
                "		neXOriS.Percentual AS PercenturalSenar," & vbCrLf & _
                "		nDev.Valor AS ValorDevolucao, 0 as ValorDevolucaoFunrural, 0 as ValorDevolucaoSenar" & vbCrLf &
                "into #Notas" & vbCrLf &
                "FROM   NotasFiscais n" & vbCrLf &
                "	inner join NotasXNotas nXn" & vbCrLf &
                "			ON nXn.Empresa_Id             = n.Empresa_Id" & vbCrLf &
                "		 	and nXn.EndEmpresa_iD         = n.EndEmpresa_Id" & vbCrLf &
                "			and nXn.Cliente_Id            = n.Cliente_Id" & vbCrLf &
                "			and nXn.EndCliente_Id         = n.EndCliente_Id" & vbCrLf &
                "			and nXn.OrigemEntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                "			and nXn.OrigemSerie_Id        = n.Serie_Id" & vbCrLf &
                "			and nXn.OrigemNota_Id         = n.Nota_Id" & vbCrLf &
                "	inner join NotasFiscais nXOri" & vbCrLf &
                "			ON nXOri.Empresa_Id       = nXn.Empresa_Id" & vbCrLf &
                "		 	and nXOri.EndEmpresa_iD   = nXn.EndEmpresa_Id" & vbCrLf &
                "			and nXOri.Cliente_Id      = nXn.Cliente_Id" & vbCrLf &
                "			and nXOri.EndCliente_Id   = nXn.EndCliente_Id" & vbCrLf &
                "			and nXOri.EntradaSaida_Id = nXn.EntradaSaida_Id" & vbCrLf &
                "			and nXOri.Serie_Id        = nXn.Serie_Id" & vbCrLf &
                "			and nXOri.Nota_Id         = nXn.Nota_Id" & vbCrLf &
                "	inner join NotasFiscaisXItens niXOri" & vbCrLf &
                "			ON niXOri.Empresa_Id       = nXOri.Empresa_Id" & vbCrLf &
                "		 	and niXOri.EndEmpresa_iD   = nXOri.EndEmpresa_Id" & vbCrLf &
                "			and niXOri.Cliente_Id      = nXOri.Cliente_Id" & vbCrLf &
                "			and niXOri.EndCliente_Id   = nXOri.EndCliente_Id" & vbCrLf &
                "			and niXOri.EntradaSaida_Id = nXOri.EntradaSaida_Id" & vbCrLf &
                "			and niXOri.Serie_Id        = nXOri.Serie_Id" & vbCrLf &
                "			and niXOri.Nota_Id         = nXOri.Nota_Id" & vbCrLf &
                "	inner join NotasFiscaisXEncargos neXOri" & vbCrLf &
                "			ON neXOri.Empresa_Id       = niXOri.Empresa_Id" & vbCrLf &
                "		 	and neXOri.EndEmpresa_iD   = niXOri.EndEmpresa_Id" & vbCrLf &
                "			and neXOri.Cliente_Id      = niXOri.Cliente_Id" & vbCrLf &
                "			and neXOri.EndCliente_Id   = niXOri.EndCliente_Id" & vbCrLf &
                "			and neXOri.EntradaSaida_Id = niXOri.EntradaSaida_Id" & vbCrLf &
                "			and neXOri.Serie_Id        = niXOri.Serie_Id" & vbCrLf &
                "			and neXOri.Nota_Id         = niXOri.Nota_Id" & vbCrLf &
                "			and neXOri.CFOP_Id         = niXOri.CFOP_Id" & vbCrLf &
                "			and neXOri.Produto_Id      = niXOri.Produto_Id" & vbCrLf &
                "			and neXOri.Sequencia_Id    = niXOri.Sequencia_Id" & vbCrLf &
                "			and neXOri.Encargo_Id      = 'PRODUTO'" & vbCrLf &
                "	inner join SubOperacoes soOri" & vbCrLf &
                "			ON soOri.Operacao_id       = nXOri.Operacao" & vbCrLf &
                "		 	AND soOri.SubOperacoes_Id  = nXOri.SubOperacao" & vbCrLf &
                "	left join NotasFiscaisXEncargos neXOriF" & vbCrLf &
                "			ON neXOriF.Empresa_Id       = niXOri.Empresa_Id" & vbCrLf &
                "		 	and neXOriF.EndEmpresa_iD   = niXOri.EndEmpresa_Id" & vbCrLf &
                "			and neXOriF.Cliente_Id      = niXOri.Cliente_Id" & vbCrLf &
                "			and neXOriF.EndCliente_Id   = niXOri.EndCliente_Id" & vbCrLf &
                "			and neXOriF.EntradaSaida_Id = niXOri.EntradaSaida_Id" & vbCrLf &
                "			and neXOriF.Serie_Id        = niXOri.Serie_Id" & vbCrLf &
                "			and neXOriF.Nota_Id         = niXOri.Nota_Id" & vbCrLf &
                "			and neXOriF.CFOP_Id         = niXOri.CFOP_Id" & vbCrLf &
                "			and neXOriF.Produto_Id      = niXOri.Produto_Id" & vbCrLf &
                "			and neXOriF.Sequencia_Id    = niXOri.Sequencia_Id" & vbCrLf &
                "			and neXOriF.Encargo_Id      = 'FUNRURAL'" & vbCrLf &
                "	left join NotasFiscaisXEncargos neXOriS" & vbCrLf &
                "			ON neXOriS.Empresa_Id       = niXOri.Empresa_Id" & vbCrLf &
                "		 	and neXOriS.EndEmpresa_iD   = niXOri.EndEmpresa_Id" & vbCrLf &
                "			and neXOriS.Cliente_Id      = niXOri.Cliente_Id" & vbCrLf &
                "			and neXOriS.EndCliente_Id   = niXOri.EndCliente_Id" & vbCrLf &
                "			and neXOriS.EntradaSaida_Id = niXOri.EntradaSaida_Id" & vbCrLf &
                "			and neXOriS.Serie_Id        = niXOri.Serie_Id" & vbCrLf &
                "			and neXOriS.Nota_Id         = niXOri.Nota_Id" & vbCrLf &
                "			and neXOriS.CFOP_Id         = niXOri.CFOP_Id" & vbCrLf &
                "			and neXOriS.Produto_Id      = niXOri.Produto_Id" & vbCrLf &
                "			and neXOriS.Sequencia_Id    = niXOri.Sequencia_Id" & vbCrLf &
                "			and neXOriS.Encargo_Id      = 'SENAR'" & vbCrLf &
                "	left join (SELECT DEV.EmpresaDevolucao_Id, DEV.EndEmpresaDevolucao_Id, DEV.ClienteDevolucao_Id, DEV.EndClienteDevolucao_Id, " & vbCrLf &
                "						DEV.EntradaSaida_Id, DEV.Serie_Id, DEV.Nota_Id, sum(DEV.Valor) AS Valor" & vbCrLf &
                "				FROM NotaFiscalDevolucaoXNotaFiscal DEV" & vbCrLf &
                "						inner join NotasFiscais nDevNota" & vbCrLf &
                "								ON nDevNota.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevNota.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevNota.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevNota.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevNota.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevNota.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevNota.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "						inner join NotasFiscaisXEncargos nDevP" & vbCrLf &
                "								ON nDevP.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevP.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevP.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevP.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevP.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevP.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevP.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevP.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevP.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevP.Encargo_Id      = 'PRODUTO'" & vbCrLf &
                "						left join NotasFiscaisXEncargos nDevF" & vbCrLf &
                "								ON nDevF.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevF.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevF.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevF.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevF.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevF.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevF.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevF.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevF.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevF.Encargo_Id      = 'FUNRURAL'" & vbCrLf &
                "						left join NotasFiscaisXEncargos nDevS" & vbCrLf &
                "								ON nDevS.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevS.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevS.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevS.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevS.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevS.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevS.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevS.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevS.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevS.Encargo_Id      = 'SENAR'" & vbCrLf &
                "						WHERE nDevNota.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'" & vbCrLf &
                "						  AND DEV.EntradaSaidaDevolucao_Id = 'S'" & vbCrLf &
                "						GROUP BY DEV.EmpresaDevolucao_Id, DEV.EndEmpresaDevolucao_Id, DEV.ClienteDevolucao_Id, DEV.EndClienteDevolucao_Id,DEV.EntradaSaida_Id, DEV.Serie_Id, DEV.Nota_Id) as nDev" & vbCrLf &
                "			ON nDev.EmpresaDevolucao_Id     = niXOri.Empresa_Id" & vbCrLf &
                "		 	and nDev.EndEmpresaDevolucao_Id = niXOri.EndEmpresa_Id" & vbCrLf &
                "			and nDev.ClienteDevolucao_Id    = niXOri.Cliente_Id" & vbCrLf &
                "			and nDev.EndClienteDevolucao_Id = niXOri.EndCliente_Id" & vbCrLf &
                "			and nDev.EntradaSaida_Id        = niXOri.EntradaSaida_Id" & vbCrLf &
                "			and nDev.Serie_Id               = niXOri.Serie_Id" & vbCrLf &
                "			and nDev.Nota_Id                = niXOri.Nota_Id" & vbCrLf &
                "	inner join Clientes c" & vbCrLf &
                "			on c.Cliente_Id   = nXOri.Cliente_Id" & vbCrLf &
                "			and c.Endereco_Id = nXOri.EndCliente_Id" & vbCrLf &
                "WHERE nXOri.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'" & vbCrLf &
                "AND   LEFT(nXOri.Empresa_id,8) = '" & Left(codEmpresa, 8) & "'" & vbCrLf &
                "AND   nXOri.TipoDeDocumento = 1" & vbCrLf &
                "AND   n.TipoDeDocumento = 15" & vbCrLf &
                "AND   n.NFG = 1" & vbCrLf &
                "AND   nXOri.EntradaSaida_Id = 'E'" & vbCrLf &
                "AND   soOri.Classe in('COMPRAS', 'REAJUSTES')" & vbCrLf

        Sql &= "" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf

        Sql &= "insert into #Notas" & vbCrLf &
                "SELECT n.Empresa_Id, n.EndEmpresa_Id, n.Cliente_Id, n.EndCliente_Id, c.Nome, n.EntradaSaida_Id, n.Nota_Id, n.Movimento, " & vbCrLf &
                "		ni.Produto_Id, ne.Valor, neF.Valor as VlrFunruralOriginal, " & vbCrLf &
                "		case " & vbCrLf &
                "			when (neF.Percentual = 1.3 or neF.Percentual = 1.5) " & vbCrLf &
                "				then 1.2" & vbCrLf &
                "				else neF.Percentual" & vbCrLf &
                "			end as Percentual," & vbCrLf &
                "		isnull(neS.Valor,0) AS ValorSenar, " & vbCrLf &
                "		neS.Percentual AS PercenturalSenar," & vbCrLf &
                "		nDev.Valor AS ValorDevolucao, 0 as ValorDevolucaoFunrural, 0 as ValorDevolucaoSenar" & vbCrLf &
                "FROM   NotasFiscais n" & vbCrLf &
                "	inner join NotasFiscaisXItens ni" & vbCrLf &
                "			ON ni.Empresa_Id       = n.Empresa_Id" & vbCrLf &
                "		 	and ni.EndEmpresa_iD   = n.EndEmpresa_Id" & vbCrLf &
                "			and ni.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                "			and ni.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                "			and ni.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                "			and ni.Serie_Id        = n.Serie_Id" & vbCrLf &
                "			and ni.Nota_Id         = n.Nota_Id" & vbCrLf &
                "	inner join NotasFiscaisXEncargos ne" & vbCrLf &
                "			ON ne.Empresa_Id       = ni.Empresa_Id" & vbCrLf &
                "		 	and ne.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf &
                "			and ne.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                "			and ne.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                "			and ne.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                "			and ne.Serie_Id        = ni.Serie_Id" & vbCrLf &
                "			and ne.Nota_Id         = ni.Nota_Id" & vbCrLf &
                "			and ne.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                "			and ne.Produto_Id      = ni.Produto_Id" & vbCrLf &
                "			and ne.Sequencia_Id    = ni.Sequencia_Id" & vbCrLf &
                "			and ne.Encargo_Id      = 'PRODUTO'" & vbCrLf &
                "	left join NotasFiscaisXEncargos neF" & vbCrLf &
                "			ON neF.Empresa_Id       = ni.Empresa_Id" & vbCrLf &
                "		 	and neF.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf &
                "			and neF.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                "			and neF.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                "			and neF.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                "			and neF.Serie_Id        = ni.Serie_Id" & vbCrLf &
                "			and neF.Nota_Id         = ni.Nota_Id" & vbCrLf &
                "			and neF.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                "			and neF.Produto_Id      = ni.Produto_Id" & vbCrLf &
                "			and neF.Sequencia_Id    = ni.Sequencia_Id" & vbCrLf &
                "			and neF.Encargo_Id      = 'FUNRURAL'" & vbCrLf &
                "	left join NotasFiscaisXEncargos neS" & vbCrLf &
                "			ON neS.Empresa_Id       = ni.Empresa_Id" & vbCrLf &
                "		 	and neS.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf &
                "			and neS.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                "			and neS.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                "			and neS.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                "			and neS.Serie_Id        = ni.Serie_Id" & vbCrLf &
                "			and neS.Nota_Id         = ni.Nota_Id" & vbCrLf &
                "			and neS.CFOP_Id         = ni.CFOP_Id" & vbCrLf &
                "			and neS.Produto_Id      = ni.Produto_Id" & vbCrLf &
                "			and neS.Sequencia_Id    = ni.Sequencia_Id" & vbCrLf &
                "			and neS.Encargo_Id      = 'SENAR'" & vbCrLf &
                "	inner join SubOperacoes so" & vbCrLf &
                "			ON so.Operacao_id       = n.Operacao" & vbCrLf &
                "		 	AND so.SubOperacoes_Id  = n.SubOperacao" & vbCrLf &
                "	left join (SELECT DEV.EmpresaDevolucao_Id, DEV.EndEmpresaDevolucao_Id, DEV.ClienteDevolucao_Id, DEV.EndClienteDevolucao_Id, " & vbCrLf &
                "						DEV.EntradaSaida_Id, DEV.Serie_Id, DEV.Nota_Id, sum(DEV.Valor) AS Valor" & vbCrLf &
                "				FROM NotaFiscalDevolucaoXNotaFiscal DEV" & vbCrLf &
                "						inner join NotasFiscais nDevNota" & vbCrLf &
                "								ON nDevNota.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevNota.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevNota.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevNota.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevNota.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevNota.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevNota.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "						inner join NotasFiscaisXEncargos nDevP" & vbCrLf &
                "								ON nDevP.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevP.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevP.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevP.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevP.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevP.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevP.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevP.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevP.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevP.Encargo_Id      = 'PRODUTO'" & vbCrLf &
                "						left join NotasFiscaisXEncargos nDevF" & vbCrLf &
                "								ON nDevF.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevF.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevF.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevF.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevF.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevF.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevF.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevF.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevF.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevF.Encargo_Id      = 'FUNRURAL'" & vbCrLf &
                "						left join NotasFiscaisXEncargos nDevS" & vbCrLf &
                "								ON nDevS.Empresa_Id       = DEV.EmpresaDevolucao_Id" & vbCrLf &
                "		 						and nDevS.EndEmpresa_Id   = DEV.EndEmpresaDevolucao_Id" & vbCrLf &
                "								and nDevS.Cliente_Id      = DEV.ClienteDevolucao_Id" & vbCrLf &
                "								and nDevS.EndCliente_Id   = DEV.EndClienteDevolucao_Id" & vbCrLf &
                "								and nDevS.EntradaSaida_Id = DEV.EntradaSaidaDevolucao_Id" & vbCrLf &
                "								and nDevS.Serie_Id        = DEV.SerieDevolucao_Id" & vbCrLf &
                "								and nDevS.Nota_Id         = DEV.NotaDevolucao_Id" & vbCrLf &
                "								and nDevS.CFOP_Id         = DEV.CFOPDevolucao_Id" & vbCrLf &
                "								and nDevS.Produto_Id      = DEV.Produto_Id" & vbCrLf &
                "								and nDevS.Encargo_Id      = 'SENAR'" & vbCrLf &
                "						WHERE nDevNota.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'" & vbCrLf &
                "						  AND DEV.EntradaSaidaDevolucao_Id = 'S'" & vbCrLf &
                "						GROUP BY DEV.EmpresaDevolucao_Id, DEV.EndEmpresaDevolucao_Id, DEV.ClienteDevolucao_Id, DEV.EndClienteDevolucao_Id,DEV.EntradaSaida_Id, DEV.Serie_Id, DEV.Nota_Id) as nDev" & vbCrLf &
                "			ON nDev.EmpresaDevolucao_Id     = ni.Empresa_Id" & vbCrLf &
                "		 	and nDev.EndEmpresaDevolucao_Id = ni.EndEmpresa_Id" & vbCrLf &
                "			and nDev.ClienteDevolucao_Id    = ni.Cliente_Id" & vbCrLf &
                "			and nDev.EndClienteDevolucao_Id = ni.EndCliente_Id" & vbCrLf &
                "			and nDev.EntradaSaida_Id        = ni.EntradaSaida_Id" & vbCrLf &
                "			and nDev.Serie_Id               = ni.Serie_Id" & vbCrLf &
                "			and nDev.Nota_Id                = ni.Nota_Id" & vbCrLf &
                "	inner join Clientes c" & vbCrLf &
                "			on c.Cliente_Id   = n.Cliente_Id" & vbCrLf &
                "			and c.Endereco_Id = n.EndCliente_Id" & vbCrLf &
                "WHERE n.Movimento BETWEEN '" & Format(CDate(txtDataInicial.Text), "yyyy-MM-dd") & "' AND '" & Format(CDate(txtDataFinal.Text), "yyyy-MM-dd") & "'" & vbCrLf &
                "AND   LEFT(n.Empresa_id,8) = '" & Left(codEmpresa, 8) & "'" & vbCrLf &
                "AND   n.TipoDeDocumento = 1" & vbCrLf &
                "AND   n.NFG = 0" & vbCrLf &
                "AND   n.EntradaSaida_Id = 'E'" & vbCrLf &
                "AND   n.Eletronica = 'S'" & vbCrLf &
                "AND   n.NossaEmissao = 'N'" & vbCrLf &
                "AND   so.Classe in('COMPRAS', 'REAJUSTES')" & vbCrLf &
                "AND   so.Devolucao = 'N'" & vbCrLf &
                "AND   c.Categoria in(1,2,3)" & vbCrLf

        Sql &= "" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf

        Sql &= "select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Nome, EntradaSaida_Id, Nota_Id, Movimento," & vbCrLf &
                "		Produto_Id, Valor, " & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(Percentual,0) > 0" & vbCrLf &
                "				then round((Valor * Percentual) / 100, 2)" & vbCrLf &
                "				else 0" & vbCrLf &
                "			end as ValorFunrural," & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(Percentual,0) > 0" & vbCrLf &
                "				then round((Valor * 0.1) / 100, 2)" & vbCrLf &
                "				else 0" & vbCrLf &
                "			end as ValorRat," & vbCrLf &
                "			ValorSenar, " & vbCrLf &
                "			ValorDevolucao, " & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(ValorDevolucao,0) > 0" & vbCrLf &
                "				then round((ValorDevolucao * Percentual) / 100, 2)" & vbCrLf &
                "				else 0" & vbCrLf &
                "			end as ValorDevolucaoFunrural," & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(ValorDevolucao,0) > 0" & vbCrLf &
                "				then round((ValorDevolucao * 0.1) / 100, 2)" & vbCrLf &
                "				else 0" & vbCrLf &
                "			end as ValorDevolucaoRat," & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(ValorSenar,0) > 0" & vbCrLf &
                "				then round((ValorDevolucao * PercenturalSenar) / 100, 2)" & vbCrLf &
                "				else 0" & vbCrLf &
                "			end as ValorDevolucaoSenar" & vbCrLf &
                "into #NotasDev" & vbCrLf &
                "from #Notas" & vbCrLf

        Sql &= "" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf

        Sql &= "select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Nome, EntradaSaida_Id, Nota_Id, Movimento," & vbCrLf &
                "		Produto_Id, Valor, ValorFunrural, ValorRat, ValorSenar, " & vbCrLf &
                "		ValorDevolucao, ValorDevolucaoFunrural, ValorDevolucaoRat, ValorDevolucaoSenar," & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(ValorDevolucao,0) > 0" & vbCrLf &
                "				then (Valor - ValorDevolucao)" & vbCrLf &
                "				else Valor" & vbCrLf &
                "			end as SaldoValor, " & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(ValorDevolucaoFunrural,0) > 0" & vbCrLf &
                "				then (ValorFunrural - ValorDevolucaoFunrural)" & vbCrLf &
                "				else ValorFunrural" & vbCrLf &
                "			end as SaldoFunrural, " & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(ValorDevolucaoFunrural,0) > 0 and isnull(ValorDevolucaoRat,0) > 0" & vbCrLf &
                "				then (ValorRat - ValorDevolucaoRat)" & vbCrLf &
                "				else ValorRat" & vbCrLf &
                "			end as SaldoRat, " & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(ValorDevolucaoSenar,0) > 0" & vbCrLf &
                "				then (ValorSenar - ValorDevolucaoSenar)" & vbCrLf &
                "				else ValorSenar" & vbCrLf &
                "			end as SaldoSenar" & vbCrLf &
                "into #NotasReinf" & vbCrLf &
                "from #NotasDev" & vbCrLf

        Sql &= "" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf


        If comNome Then
            Sql &= "select Empresa_Id, Cliente_Id, Nome, sum(SaldoValor) as SaldoValor, sum(SaldoFunrural) as SaldoFunrural, sum(SaldoRat) as SaldoRat, sum(SaldoSenar) as SaldoSenar" & vbCrLf &
                    "from #NotasReinf" & vbCrLf &
                    "group by Empresa_Id, Cliente_Id, Nome" & vbCrLf &
                    "HAVING isnull(sum(SaldoFunrural),0) > 0 or isnull(sum(SaldoRat),0) > 0 or isnull(sum(SaldoSenar),0) > 0"
        Else
            Sql &= "select Empresa_Id, Cliente_Id, sum(SaldoValor) as SaldoValor, sum(SaldoFunrural) as SaldoFunrural, sum(SaldoRat) as SaldoRat, sum(SaldoSenar) as SaldoSenar" & vbCrLf &
                    "from #NotasReinf" & vbCrLf &
                    "group by Empresa_Id, Cliente_Id" & vbCrLf &
                    "HAVING isnull(sum(SaldoFunrural),0) > 0 or isnull(sum(SaldoRat),0) > 0 or isnull(sum(SaldoSenar),0) > 0"
        End If

        ds2055 = Banco.ConsultaDataSet(Sql, "Consulta")

        Return ds2055
    End Function

    Private Sub R4099()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))
        Dim sequencia As Integer = 1

        txtArquivo4099.Text = "R4099-" & Empresa.Codigo & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

        Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo4099.Text)

        Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

        xtw.Formatting = Formatting.Indented

        xtw.WriteStartDocument()
        xtw.WriteStartElement("Reinf")

        xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evt4099FechamentoDirf/v2_01_02")

        xtw.WriteStartElement("evtFech")

        Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

        xtw.WriteAttributeString("id", idEvento)

        xtw.WriteStartElement("ideEvento")
        xtw.WriteElementString("perApur", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
        xtw.WriteElementString("tpAmb", "1")  ' 1 - Produção    2 - Homologação
        xtw.WriteElementString("procEmi", "1")
        xtw.WriteElementString("verProc", "v2_01_02")
        xtw.WriteEndElement()

        xtw.WriteStartElement("ideContri")
        xtw.WriteElementString("tpInsc", "1")
        xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
        xtw.WriteEndElement()

        xtw.WriteStartElement("ideRespInf")
        If Left(Empresa.Codigo, 8) = "05366261" OrElse Left(Empresa.Codigo, 8) = "38198213" OrElse Left(Empresa.Codigo, 8) = "40938762" Then
            xtw.WriteElementString("nmResp", "LUIZ HENRIQUE MULINARI TONETT")
            xtw.WriteElementString("cpfResp", "04121606175")
        Else
            xtw.WriteElementString("nmResp", Empresa.Empresa.NomeDoContador)
            xtw.WriteElementString("cpfResp", Empresa.Empresa.CPFContador)
        End If

        xtw.WriteElementString("telefone", Empresa.Empresa.TelefoneContador)
        xtw.WriteElementString("email", Empresa.Empresa.EmailContador)
        xtw.WriteEndElement()

        xtw.WriteStartElement("infoFech")

        If radReab4099.Checked Then ' 0 - Fechamento    1 - Reabertura
            xtw.WriteElementString("fechRet", "1")
        Else
            xtw.WriteElementString("fechRet", "0")
        End If

        xtw.WriteEndElement()

        xtw.WriteEndElement()
        xtw.WriteEndElement()
        xtw.WriteEndDocument()
        'libera o XmlTextWriter
        xtw.Flush()
        'fechar o XmlTextWriter
        xtw.Close()
        'Termina aqui

        'btnArquivo4099.Visible = True
    End Sub

    Private Sub R4020()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))
        Dim ds As DataSet

        ds = consulta4000(Empresa.Codigo, False, "J")

        If ds IsNot Nothing AndAlso Not ds.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

            Exit Sub
        End If

        Dim sequencia As Integer = 1

        For Each row In ds.Tables(2).Rows
            If Not String.IsNullOrWhiteSpace(txtseq4020.Text) Then
                If Not sequencia = txtseq4020.Text Then
                    sequencia += 1
                    Continue For
                End If
            End If

            txtArquivo4020.Text = "R4020-" & row("Empresa_Id") & "-" & row("Cliente_Id") & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

            Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo4020.Text)

            Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

            xtw.Formatting = Formatting.Indented

            xtw.WriteStartDocument()
            xtw.WriteStartElement("Reinf")

            xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evt4020PagtoBeneficiarioPJ/v2_01_02")

            xtw.WriteStartElement("evtRetPJ")

            Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

            xtw.WriteAttributeString("id", idEvento)

            xtw.WriteStartElement("ideEvento")

            If Not String.IsNullOrWhiteSpace(txtRecib4020.Text) AndAlso (sequencia = txtseq4020.Text) Then 'Informe [1] para arquivo original ou [2] para arquivo de retificação.
                xtw.WriteElementString("indRetif", "2")
                xtw.WriteElementString("nrRecibo", txtRecib4020.Text)  'Preencher com o número do recibo "nrRecArqBase" do arquivo a ser retificado.
            Else
                xtw.WriteElementString("indRetif", "1")
            End If

            xtw.WriteElementString("perApur", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
            xtw.WriteElementString("tpAmb", "1")  ' 1 - Produção    2 - Homologação
            xtw.WriteElementString("procEmi", "1")
            xtw.WriteElementString("verProc", "v2_01_02")
            xtw.WriteEndElement()

            xtw.WriteStartElement("ideContri")
            xtw.WriteElementString("tpInsc", "1")
            xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
            xtw.WriteEndElement()

            xtw.WriteStartElement("ideEstab")
            xtw.WriteElementString("tpInscEstab", "1")
            xtw.WriteElementString("nrInscEstab", row("Empresa_Id"))

            xtw.WriteStartElement("ideBenef")
            xtw.WriteElementString("cnpjBenef", row("Cliente_Id"))

            xtw.WriteStartElement("idePgto")
            xtw.WriteElementString("natRend", row("CodigoNaturezaDeRendimento"))

            For Each rowPgto In ds.Tables(1).Rows
                If rowPgto("Empresa_Id") = row("Empresa_Id") AndAlso rowPgto("Cliente_Id") = row("Cliente_Id") AndAlso rowPgto("CodigoNaturezaDeRendimento") = row("CodigoNaturezaDeRendimento") Then
                    xtw.WriteStartElement("infoPgto")
                    xtw.WriteElementString("dtFG", Format(rowPgto("Movimento"), "yyyy-MM-dd"))
                    xtw.WriteElementString("vlrBruto", rowPgto("Base"))

                    xtw.WriteStartElement("retencoes")

                    If rowPgto("ValorIR") > 0 Then
                        xtw.WriteElementString("vlrBaseIR", rowPgto("Base"))
                        xtw.WriteElementString("vlrIR", rowPgto("ValorIR"))
                    End If

                    If rowPgto("ValorAgregado") > 0 Then
                        xtw.WriteElementString("vlrBaseAgreg", rowPgto("Base"))
                        xtw.WriteElementString("vlrAgreg", rowPgto("ValorAgregado"))
                    End If

                    xtw.WriteEndElement()

                    'infoPgto
                    xtw.WriteEndElement()
                End If
            Next

            'idePgto
            xtw.WriteEndElement()
            'ideBenef
            xtw.WriteEndElement()
            'ideEstab
            xtw.WriteEndElement()

            xtw.WriteEndElement()
            xtw.WriteEndElement()
            xtw.WriteEndDocument()
            'libera o XmlTextWriter
            xtw.Flush()
            'fechar o XmlTextWriter
            xtw.Close()
            'Termina aqui

            sequencia += 1

        Next
        'btnArquivo4020.Visible = True
    End Sub

    Private Sub R4010()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))
        Dim ds As DataSet

        ds = consulta4000(Empresa.Codigo, False, "F")

        If ds IsNot Nothing AndAlso Not ds.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, "Período sem movimento", eTitulo.Info)

            Exit Sub
        End If

        Dim sequencia As Integer = 1

        For Each row In ds.Tables(2).Rows

            txtArquivo4010.Text = "R4010-" & Empresa.Codigo & "-" & row("Cliente_Id") & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

            Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo4010.Text)

            Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

            xtw.Formatting = Formatting.Indented

            xtw.WriteStartDocument()
            xtw.WriteStartElement("Reinf")

            xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evt4010PagtoBeneficiarioPF/v2_01_02")

            xtw.WriteStartElement("evtRetPF")

            Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

            xtw.WriteAttributeString("id", idEvento)

            xtw.WriteStartElement("ideEvento")

            Dim txtReci4010 As String = ""
            Dim txtSeq4010 As String = "1"

            If Not String.IsNullOrWhiteSpace(txtReci4010) AndAlso (sequencia = txtSeq4010) Then 'Informe [1] para arquivo original ou [2] para arquivo de retificação.
                xtw.WriteElementString("indRetif", "2")
                xtw.WriteElementString("nrRecibo", txtReci4010)  'Preencher com o número do recibo "nrRecArqBase" do arquivo a ser retificado.
            Else
                xtw.WriteElementString("indRetif", "1")
            End If

            xtw.WriteElementString("perApur", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
            xtw.WriteElementString("tpAmb", "1")  ' 1 - Produção    2 - Homologação
            xtw.WriteElementString("procEmi", "1")
            xtw.WriteElementString("verProc", "v2_01_02")
            xtw.WriteEndElement()

            xtw.WriteStartElement("ideContri")
            xtw.WriteElementString("tpInsc", "1")
            xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
            xtw.WriteEndElement()

            xtw.WriteStartElement("ideEstab")
            xtw.WriteElementString("tpInscEstab", "1")
            xtw.WriteElementString("nrInscEstab", row("Empresa_Id"))

            xtw.WriteStartElement("ideBenef")
            xtw.WriteElementString("cpfBenef", row("Cliente_Id"))

            xtw.WriteStartElement("idePgto")
            xtw.WriteElementString("natRend", row("CodigoNaturezaDeRendimento"))


            For Each rowPgto In ds.Tables(1).Rows
                If rowPgto("Cliente_Id") = row("Cliente_Id") AndAlso rowPgto("CodigoNaturezaDeRendimento") = row("CodigoNaturezaDeRendimento") Then

                    xtw.WriteStartElement("infoPgto")
                    xtw.WriteElementString("dtFG", Format(rowPgto("Movimento"), "yyyy-MM-dd"))
                    xtw.WriteElementString("vlrRendBruto", rowPgto("Base"))
                    If row("CodigoNaturezaDeRendimento") = "12001" Then
                        'NÃO DEVE INCLUIR A TAG "vlrRendTrib e vlrIR" PARA ESSE CÓDIGO 12001 DE lucros e dividendos - FURLAN - 09/01/2024
                    Else
                        xtw.WriteElementString("vlrRendTrib", rowPgto("Base"))
                        xtw.WriteElementString("vlrIR", rowPgto("ValorIR"))
                    End If
                    xtw.WriteEndElement()
                End If
            Next

            'idePgto
            xtw.WriteEndElement()
            'ideBenef
            xtw.WriteEndElement()
            'ideEstab
            xtw.WriteEndElement()

            xtw.WriteEndElement()
            xtw.WriteEndElement()
            xtw.WriteEndDocument()
            'libera o XmlTextWriter
            xtw.Flush()
            'fechar o XmlTextWriter
            xtw.Close()
            'Termina aqui


            sequencia += 1

        Next

        'btnArquivo4010.Visible = True
    End Sub

    Private Sub R2098()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))
        Dim sequencia As Integer = 1

        'Dim arquivoAbertura As String = "R2098-" & Empresa.Codigo & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

        txtArquivo2098.Text = "R2098-" & Empresa.Codigo & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

        If File.Exists(Server.MapPath("~/SpedReinf/" & txtArquivo2098.Text)) Then
            MsgBox(Me.Page, "Arquivo de reabertura já existe.")
            Exit Sub
        End If

        Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo2098.Text)

        Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

        xtw.Formatting = Formatting.Indented

        xtw.WriteStartDocument()
        xtw.WriteStartElement("Reinf")

        xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evtReabreEvPer/v2_01_02")

        xtw.WriteStartElement("evtReabreEvPer")

        Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

        xtw.WriteAttributeString("id", idEvento)

        xtw.WriteStartElement("ideEvento")
        xtw.WriteElementString("perApur", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
        xtw.WriteElementString("tpAmb", "1")  ' 1 - Produção    2 - Homologação
        xtw.WriteElementString("procEmi", "1")
        xtw.WriteElementString("verProc", "2.01.02")
        xtw.WriteEndElement()

        xtw.WriteStartElement("ideContri")
        xtw.WriteElementString("tpInsc", "1")
        xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
        xtw.WriteEndElement()

        xtw.WriteEndElement()
        xtw.WriteEndElement()
        xtw.WriteEndDocument()
        'libera o XmlTextWriter
        xtw.Flush()
        'fechar o XmlTextWriter
        xtw.Close()
        'Termina aqui

    End Sub

    Private Sub R2099()
        Dim emp() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim Empresa As New Cliente(emp(0), emp(1))
        Dim sequencia As Integer = 1

        Dim tem2010 As Boolean = False
        Dim ds2010 As DataSet

        Dim tem2040 As Boolean = False
        Dim ds2040 As DataSet

        Dim tem2055 As Boolean = False
        Dim ds2055 As DataSet

        Dim arquivoAbertura As String = "R1000-" & Empresa.Codigo & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

        txtArquivo2099.Text = "R2099-" & Empresa.Codigo & "-" & Format(CDate(txtDataInicial.Text), "yyyyMM") & "-" & sequencia.ToString("00000") & ".xml"

        ds2010 = consulta2010(Empresa.Codigo)

        If ds2010 IsNot Nothing AndAlso ds2010.Tables(0).Rows.Count > 0 Then tem2010 = True

        ds2040 = consulta2040(Empresa.Codigo)

        If ds2040 IsNot Nothing AndAlso ds2040.Tables(0).Rows.Count > 0 Then tem2040 = True

        ds2055 = consulta2055(Empresa.Codigo, False)

        If ds2055 IsNot Nothing AndAlso ds2055.Tables(0).Rows.Count > 0 Then tem2055 = True

        If Not File.Exists(Server.MapPath("~/SpedReinf/" & arquivoAbertura)) Then
            MsgBox(Me.Page, "Arquivo de abertura não foi encontrado.")
            Exit Sub
        ElseIf File.Exists(Server.MapPath("~/SpedReinf/" & txtArquivo2099.Text)) Then
            MsgBox(Me.Page, "Arquivo de fechamento já existe.")
            Exit Sub
        End If

        Dim strFilePath As String = Server.MapPath("~/SpedReinf/" & txtArquivo2099.Text)

        Dim xtw As XmlTextWriter = New XmlTextWriter(strFilePath, New UTF8Encoding(False))

        xtw.Formatting = Formatting.Indented

        xtw.WriteStartDocument()
        xtw.WriteStartElement("Reinf")

        xtw.WriteAttributeString("xmlns", "http://www.reinf.esocial.gov.br/schemas/evtFechamento/v2_01_02")

        xtw.WriteStartElement("evtFechaEvPer")

        Dim idEvento As String = "ID1" & Left(Empresa.Codigo, 8) & "000000" & Date.Now.ToString("yyyyMMdd") & Format(Date.Now, "HHmmss") & sequencia.ToString("00000")

        xtw.WriteAttributeString("id", idEvento)

        xtw.WriteStartElement("ideEvento")
        xtw.WriteElementString("perApur", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
        xtw.WriteElementString("tpAmb", "1")  ' 1 - Produção    2 - Homologação
        xtw.WriteElementString("procEmi", "1")
        xtw.WriteElementString("verProc", "2.01.02")
        xtw.WriteEndElement()

        xtw.WriteStartElement("ideContri")
        xtw.WriteElementString("tpInsc", "1")
        xtw.WriteElementString("nrInsc", Left(Empresa.Codigo, 8))
        xtw.WriteEndElement()

        xtw.WriteStartElement("ideRespInf")

        If Left(Empresa.Codigo, 8) = "05366261" OrElse Left(Empresa.Codigo, 8) = "38198213" OrElse Left(Empresa.Codigo, 8) = "40938762" Then
            xtw.WriteElementString("nmResp", "LUIZ HENRIQUE MULINARI TONETT")
            xtw.WriteElementString("cpfResp", "04121606175")
        Else
            xtw.WriteElementString("nmResp", Empresa.Empresa.NomeDoContador)
            xtw.WriteElementString("cpfResp", Empresa.Empresa.CPFContador)
        End If

        xtw.WriteElementString("telefone", Empresa.Empresa.TelefoneContador)
        xtw.WriteElementString("email", Empresa.Empresa.EmailContador)
        xtw.WriteEndElement()

        xtw.WriteStartElement("infoFech")
        'Contratou serviços sujeitos à retenção de contribuição previdenciária? Validação: Se for igual a [S], deve existir evento R-2010 para o período de apuração. Caso contrário, não deve existir o referido evento Valores Válidos: S, N.
        If tem2010 Then
            xtw.WriteElementString("evtServTm", "S")
        Else
            xtw.WriteElementString("evtServTm", "N")
        End If

        'Prestou serviços sujeitos à retenção de contribuição previdenciária? Validação: Se for igual a [S], deve existir evento R-2020 para o mesmo período de apuração. Caso contrário, não deve existir o referido evento. Valores Válidos: S, N.
        xtw.WriteElementString("evtServPr", "N")

        'A associação desportiva que mantém equipe de futebol profissional, possui informações sobre recursos recebidos? Validação: Se for igual a [S], deve existir o evento R-2030 para o mesmo
        xtw.WriteElementString("evtAssDespRec", "N")

        'Possui informações sobre repasses efetuados à associação desportiva que mantém equipe de futebol profissional? Validação: Se for igual a [S], deve existir o evento R-2040 para o mesmo período de apuração. Caso contrário, não deve existir o referido evento. Valores Válidos: S, N.
        If tem2040 Then
            xtw.WriteElementString("evtAssDespRep", "S")
        Else
            xtw.WriteElementString("evtAssDespRep", "N")
        End If

        'O produtor rural PJ/Agroindústria possui informações de comercialização de produção? Validação: Se for igual a [S], deve existir o evento R-2050 no período de apuração. Caso contrário, não deve existir o referido evento. Valores Válidos: S, N.
        xtw.WriteElementString("evtComProd", "N")

        'Possui informações sobre a apuração da Contribuição Previdenciária sobre a Receita Bruta? Validação: Se for igual a [S], deve existir o evento R-2060 no período de apuração. Caso contrário, não deve existir o evento Valores Válidos: S, N.
        xtw.WriteElementString("evtCPRB", "N")

        'Possui Notas Produtor Rural Bloco 2055
        If tem2055 Then
            xtw.WriteElementString("evtAquis", "S")
        Else
            xtw.WriteElementString("evtAquis", "N")
        End If

        'Possui informações de pagamentos diversos no período de apuração? Se for igual a [S], deve existir evento R-2070 válido para o período de apuração. Caso contrário, não deve existir o referido evento. Valores Válidos: S, N.
        'xtw.WriteElementString("evtPgtos", "N") - mes 11/2018 para frente não informa

        'Informar a primeira competência a partir da qual não houve movimento, cuja situação perdura até a competência atual. Validação: Preenchimento obrigatório se todos os campos a seguir mencionados forem preenchidos com [N]: {evtServTm}, {evtServPr}, {evtAssDespRec}, {evtAssDespRep}, {evtComProd}, {evtCPRB}, {evtPgtos}.

        'If Not tem2010 Then      mes 11/2018 para frente não informa
        '    xtw.WriteElementString("compSemMovto", Format(CDate(txtDataInicial.Text), "yyyy") & "-" & Format(CDate(txtDataInicial.Text), "MM"))
        'End If

        xtw.WriteEndElement()
        xtw.WriteEndElement()
        xtw.WriteEndDocument()
        'libera o XmlTextWriter
        xtw.Flush()
        'fechar o XmlTextWriter
        xtw.Close()
        'Termina aqui

        'btnArquivo2099.Visible = True
    End Sub

    Public Sub testeDeAssinatura()
        'AssinarDocumentoXML(Server.MapPath("~/SpedReinf/teste.xml"), "Reinf", "id")

        'Try
        '    ' Create a new CspParameters object to specify
        '    ' a key container.
        '    Dim cspParams As CspParameters = New CspParameters
        '    cspParams.KeyContainerName = "XML_DSIG_RSA_KEY"
        '    ' Create a new RSA signing key and save it in the container. 
        '    Dim rsaKey As RSACryptoServiceProvider = New RSACryptoServiceProvider(cspParams)
        '    ' Create a new XML document.
        '    Dim xmlDoc As XmlDocument = New XmlDocument()
        '    ' Load an XML file into the XmlDocument object.
        '    xmlDoc.PreserveWhitespace = True
        '    xmlDoc.Load(Server.MapPath("~/SpedReinf/teste.xml"))
        '    ' Sign the XML document. 

        '    SignXml(xmlDoc, rsaKey)

        '    Console.WriteLine("XML file signed.")
        '    ' Save the document.
        '    xmlDoc.Save(Server.MapPath("~/SpedReinf/testeassinado.xml"))
        'Catch e As Exception
        '    Console.WriteLine(e.Message)
        'End Try
    End Sub

    Public Function SelecionarCertificado() As String
        'Try

        '    'Representa um certificado X509
        '    Dim objCertificadoX509 As New X509Certificate2

        '    'Representa o local onde os certificados X509 são armazenados
        '    'Seleciona os certificados locais e do usuário atual
        '    Dim getCertificadosX509 As New X509Store("MY", StoreLocation.CurrentUser)
        '    getCertificadosX509.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)

        '    'Representa uma coleção de objetos X509Certificate2
        '    Dim objColecaoCertificadosX509 As New X509Certificate2Collection

        '    'Abre a caixa de diálogo com os certificados diponiveis
        '    objColecaoCertificadosX509 = X509Certificate2UI.SelectFromCollection(getCertificadosX509.Certificates,
        '    "Certificado(s) Digital(is) disponível(is)", "Selecione o certificado digital para uso no aplicativo",
        '    X509SelectionFlag.SingleSelection)

        '    'Verifica se existe algum certificado selecionado
        '    If objColecaoCertificadosX509.Count > 0 Then
        '        'Mostra o assunto do certficado selecionado
        '        'lblCertificado.Text = objColecaoCertificadosX509.Item(0).Subject.ToString
        '        'Adiciona o número do serial na variável
        '        Return objColecaoCertificadosX509.Item(0).SerialNumber.ToString
        '    End If
        'Catch ex As Exception
        '    Return String.Empty
        'End Try
    End Function

    Private Sub AssinarDocumentoXML(ByVal ArqXMLAssinar As String, ByVal TagXML As String, ByVal IdEvento2 As String)
        'Dim p_NumeroSerialCertificado As String = SelecionarCertificado()

        'If String.IsNullOrEmpty(p_NumeroSerialCertificado) Then
        '    MsgBox(Me.Page, "Certificado não foi selecionado")

        '    Exit Sub
        'End If

        'Try

        '    'Pega o conteúdo XML que será assinado
        '    Dim srdDocXml As StreamReader
        '    Dim strXML As String
        '    Dim strTagXml As String = TagXML

        '    srdDocXml = File.OpenText(ArqXMLAssinar)
        '    strXML = srdDocXml.ReadToEnd()
        '    srdDocXml.Close()

        '    'Representa uma coleção de objetos X509Certificate2
        '    Dim objColecaoCertificadosX509 As X509Certificate2Collection = Nothing

        '    'Instância que representa um certificado X509
        '    Dim objCertificadoX509 As New X509Certificate2()

        '    'Representa o local onde os certificados X509 são armazenados e
        '    'seleciona os certificados locais e do usuário atual
        '    Dim getCertificadosX509 As New X509Store("MY", StoreLocation.CurrentUser)

        '    'Abrir em modo de leitura e apenas os certificados existentes
        '    getCertificadosX509.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)

        '    'Procure por certificados usando o número serial como parâmetro, o último parâmetro indica para a
        '    'função retornar somente certificados válidos
        '    objColecaoCertificadosX509 = getCertificadosX509.Certificates.Find(X509FindType.FindBySerialNumber, p_NumeroSerialCertificado, True)

        '    'Verifica se existe algum certificado selecionado
        '    If objColecaoCertificadosX509.Count > 0 Then

        '        'Selecionar o certificado e armazena na váriavel
        '        objCertificadoX509 = objColecaoCertificadosX509.Item(0)
        '        Dim docXML = New XmlDocument() 'Iniciamos um novo documento XML

        '        docXML.PreserveWhitespace = False 'Não deixar espaços no documento
        '        docXML.Load(ArqXMLAssinar) 'Carrega o XML

        '        'Retorna a quantidade de tags de assinatura, deve existir apenas uma.
        '        If docXML.GetElementsByTagName(strTagXml).Count = 1 Then

        '            'Classe usada para assinar o documento XML
        '            Dim signedXml As New Xml.SignedXml(docXML)
        '            'Configura a chave de assinatura
        '            signedXml.SigningKey = objCertificadoX509.PrivateKey


        '            'Representa o elemento de referência da assinatura digital do XML
        '            Dim Referencia As New Xml.Reference()
        '            ' Representar a idendificação do elemento de assinatura
        '            Referencia.Uri = "#ID00000000002"  'VER COMO RESOLVER.. ESTÁ O MESMO NUMERO DA ID

        '            'Essa classe remove os elementos de assintura <Signature> antes de
        '            'criar um digest do documento XML.
        '            Dim env As New XmlDsigEnvelopedSignatureTransform

        '            'Adicionamos o elemento na referência
        '            Referencia.AddTransform(env)

        '            'Representa a canonização C14N XML
        '            Dim c14 As New XmlDsigC14NTransform

        '            'Adicionamos o elemento na referência
        '            Referencia.AddTransform(c14)

        '            'Adiciona as referências na assinatura
        '            signedXml.AddReference(Referencia)


        '            'A Classe KeyInfo representa o elemento <KeyInfo> da assinatura digital
        '            Dim KeyInfo As New KeyInfo

        '            'Carrega o certificado usando o objetoKeyInfoX509Data e
        '            'adiciona ao elemento <KeyInfo>
        '            KeyInfo.AddClause(New RSAKeyValue(CType(objColecaoCertificadosX509.Item(0).PrivateKey, RSA)))
        '            KeyInfo.AddClause(New KeyInfoX509Data(objCertificadoX509))

        '            'Adiciona o objeto KeyInfo na assinatura
        '            signedXml.KeyInfo = KeyInfo
        '            'Computa a assinatura digital
        '            signedXml.ComputeSignature()

        '            'Cria um novo elemento XML
        '            Dim xmlDigitalSignature As XmlElement
        '            'Pega a assinatura digital e adiciona ao elemento
        '            xmlDigitalSignature = signedXml.GetXml()

        '            ' Gravar o elemento no final do documento XML
        '            docXML.DocumentElement.AppendChild(docXML.ImportNode(xmlDigitalSignature, True))

        '            'Salva o documento XML já assinado
        '            Dim EscreverXML As New StreamWriter(Server.MapPath("~/SpedReinf/testeassinado.xml"))
        '            EscreverXML.Write(docXML.OuterXml)

        '            EscreverXML.Close()
        '        End If
        '    End If
        'Catch ex As Exception
        '    MsgBox(Me.Page, ex.Message)
        'End Try

    End Sub

    '' Sign an XML file. 
    '' This document cannot be verified unless the verifying 
    '' code has the key with which it was signed.
    'Private Sub SignXml(ByVal xmlDoc As XmlDocument, ByVal key As RSA)
    '    ' Check arguments.
    '    If (xmlDoc Is Nothing) Then
    '        Throw New ArgumentException("xmlDoc")
    '    End If

    '    If (key Is Nothing) Then
    '        Throw New ArgumentException("key")
    '    End If

    '    Dim signedXml As SignedXml = New SignedXml(xmlDoc)
    '    ' Add the key to the SignedXml document.
    '    signedXml.SigningKey = key
    '    ' Create a reference to be signed.
    '    Dim reference As Reference = New Reference
    '    reference.Uri = ""
    '    ' Add an enveloped transformation to the reference.
    '    Dim env As XmlDsigEnvelopedSignatureTransform = New XmlDsigEnvelopedSignatureTransform
    '    reference.AddTransform(env)
    '    ' Add the reference to the SignedXml object.
    '    signedXml.AddReference(reference)
    '    ' Compute the signature.
    '    signedXml.ComputeSignature()
    '    ' Get the XML representation of the signature and save
    '    ' it to an XmlElement object.
    '    Dim xmlDigitalSignature As XmlElement = signedXml.GetXml
    '    ' Append the element to the XML document.
    '    xmlDoc.DocumentElement.AppendChild(xmlDoc.ImportNode(xmlDigitalSignature, True))
    'End Sub

    'Public Overloads Shared Function SignFile(ByVal certs As X509Certificate2Collection, ByVal data() As Byte) As Byte()
    '    Try
    '        Dim content As ContentInfo = New ContentInfo(data)
    '        Dim signedCms As SignedCms = New SignedCms(content, False)
    '        If VerifySign(data) Then
    '            signedCms.Decode(data)
    '            ' Se o arquivo est� assinado, coloca o mesmo no objeto signedCms para ser co-assinado (mant�m as assinaturas atuais e assina com os certificados do par�metro certs.
    '        End If

    '        For Each cert As X509Certificate2 In certs
    '            Dim signer As CmsSigner = New CmsSigner(cert)
    '            signer.IncludeOption = X509IncludeOption.WholeChain
    '            signedCms.ComputeSignature(signer)
    '        Next
    '        Return signedCms.Encode
    '    Catch ex As Exception
    '        Throw New Exception(("Erro ao assinar arquivo. A mensagem retornada foi: " + ex.Message))
    '    End Try

    'End Function

    'Public Overloads Shared Function SignFile(ByVal CertFile As String, ByVal CertPass As String, ByVal data() As Byte) As Byte()
    '    Dim fs As FileStream = New FileStream(CertFile, FileMode.Open)
    '    Dim buffer() As Byte = New Byte((fs.Length) - 1) {}
    '    fs.Read(buffer, 0, buffer.Length)
    '    Dim cert As X509Certificate2 = New X509Certificate2(buffer, CertPass)
    '    fs.Close()
    '    fs.Dispose()
    '    Return SignFile(cert, data)
    'End Function

    'Public Overloads Shared Function SignFile(ByVal cert As X509Certificate2, ByVal data() As Byte) As Byte()
    '    Dim certs As X509Certificate2Collection = New X509Certificate2Collection(cert)
    '    Return SignFile(certs, data)
    'End Function

    'Public Overloads Shared Function SignFile(ByVal cert As X509Certificate2, ByVal FileName As String) As Byte()
    '    Dim fs As FileStream = New FileStream(FileName, FileMode.Open)
    '    Dim buffer() As Byte = New Byte((fs.Length) - 1) {}
    '    fs.Read(buffer, 0, buffer.Length)
    '    fs.Close()
    '    fs.Dispose()
    '    Return SignFile(cert, buffer)
    'End Function

    'Public Overloads Shared Function SignFile(ByVal CertFile As String, ByVal CertPass As String, ByVal FileName As String) As Byte()
    '    Dim fsArq As FileStream = New FileStream(FileName, FileMode.Open)
    '    Dim bufArq() As Byte = New Byte((fsArq.Length) - 1) {}
    '    fsArq.Read(bufArq, 0, bufArq.Length)
    '    fsArq.Close()
    '    fsArq.Dispose()
    '    Return SignFile(CertFile, CertPass, bufArq)
    'End Function

    'Public Shared Function VerifySign(ByVal data() As Byte) As Boolean
    '    Try
    '        Dim signed As SignedCms = New SignedCms
    '        signed.Decode(data)
    '    Catch ex As Exception
    '        Return False
    '        ' Arquivo n�o assinado
    '    End Try

    '    Return True
    'End Function

    'Private Shared Function FindCertOnStore(ByVal cert As X509Certificate2) As X509Certificate2
    '    Dim st As X509Store = New X509Store(StoreLocation.CurrentUser)
    '    st.Open(OpenFlags.ReadOnly)
    '    Dim ret As X509Certificate2 = Nothing
    '    For Each c As X509Certificate2 In st.Certificates
    '        If c.Subject.Equals(cert.Subject) Then
    '            ret = c
    '            Exit For
    '        End If

    '    Next
    '    Return ret
    'End Function
End Class