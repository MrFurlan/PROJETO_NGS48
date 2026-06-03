Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ComparativoFinanceiroXContabil
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim objEmpresa As New Cliente
    Dim strJavaScript As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ComparativoFinanceiroXContabil", "ACESSAR") Then
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        Try
            Session("ssCampo") = "Livre"
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaCFXC" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCFXC" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtCodigoEmpresa.Value = ""
        txtEmpresa.Text = ""
        txtCodigoCliente.Value = ""
        txtCliente.Text = ""
        ChkIsola.Checked = True
        RbPagar.Checked = False
        RbReceber.Checked = False
        txtPeriodoInicialConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
        SelecionarEmpresa(Session("ssEmpresa"), Session("ssEndEmpresa"))
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaEmpresas.SetarHID(HID.Value)
    End Sub

    Private Function getParam() As String
        Dim param As String = ""

        If Not String.IsNullOrWhiteSpace(txtEmpresa.Text) Then
            param &= "Parametros:" & vbCrLf & "Empresa: " & txtEmpresa.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtPeriodoInicialConsultaTitulos.Text) Then
            param &= "Data Limite: " & txtPeriodoInicialConsultaTitulos.Text
        End If

        Return param
    End Function

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaCFXC" & HID.Value) Is Nothing Then
            SelecionarEmpresa(CType(Session("objEmpresaCFXC" & HID.Value), Cliente).Codigo, CType(Session("objEmpresaCFXC" & HID.Value), Cliente).CodigoEndereco)
            Session.Remove("objEmpresaCFXC" & HID.Value)
        ElseIf Not Session("objClienteCFXC" & HID.Value) Is Nothing Then
            With CType(Session("objClienteCFXC" & HID.Value), Cliente)
                txtCodigoCliente.Value = .Codigo & "-" & .CodigoEndereco
                txtCliente.Text = .Nome & " - " & .Cidade & "/" & .CodigoEstado & " - " & .CodigoFormatado
            End With
            Session.Remove("objClienteCFXC" & HID.Value)
        End If
    End Sub

    Private Sub SelecionarEmpresa(ByVal Empresa As String, ByVal EndEmpresa As String)
        objEmpresa = New Cliente(Empresa, EndEmpresa)
        With objEmpresa
            txtCodigoEmpresa.Value = .Codigo & "-" & .CodigoEndereco
            txtEmpresa.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
        End With
    End Sub

    Private Function ValidarCampos() As Boolean
        If RbReceber.Checked = False And RbPagar.Checked = False Then
            MsgBox(Me.Page, "Tipo do relatório não foi selecionado.")
            Return False
        ElseIf txtPeriodoInicialConsultaTitulos.Text.Length = 0 Then
            MsgBox(Me.Page, "Período inicial não foi selecionado.")
            Return False
        End If
        Return True
    End Function

    Protected Sub RbReceber_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Chk_1010201.Checked = False
            Chk_1010202.Checked = False
            Chk_1010301.Checked = True
            Chk_1010302.Checked = False
            Chk_2010101.Checked = False
            Chk_2010102.Checked = False
            Chk_2010103.Checked = False
            Chk_2010104.Checked = False
            Chk_2010105.Checked = False
            Chk_2010108.Checked = False
            Chk_2010109.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RbPagar_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Chk_1010201.Checked = False
            Chk_1010202.Checked = False
            Chk_1010301.Checked = False
            Chk_1010302.Checked = False
            Chk_2010101.Checked = True
            Chk_2010102.Checked = False
            Chk_2010103.Checked = False
            Chk_2010104.Checked = False
            Chk_2010105.Checked = False
            Chk_2010108.Checked = False
            Chk_2010109.Checked = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ComparativoFinanceiroXContabil", "RELATORIO") Then
                If radPDF.Checked Then
                    RelatorioPDF()
                Else
                    RelatorioHTML()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
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

    Protected Sub RelatorioPDF()
        Try
            Dim Empresa() As String = txtCodigoEmpresa.Value.Split("-")
            Dim Cliente() As String = txtCodigoCliente.Value.Split("-")
            Dim NomeRelatorio As String = ""
            Dim NomeUnidadeDeNegocio As String = "Consolidado"
            Dim Empresas As String = ""
            Empresas = Empresa(0)
            Dim EndEmpresa As String = ""
            EndEmpresa = Empresa(1)

            Dim NomeEmpresa As String = ""
            Dim CidadeEmpresa As String = ""
            Dim EstadoEmpresa As String = ""
            Dim Clientes As String = ""
            Dim EndCliente As String = ""

            If Cliente(0) <> "" Then
                Clientes = Cliente(0)
                EndCliente = Cliente(1)
            End If

            Dim DataInicial As String = txtPeriodoInicialConsultaTitulos.Text.ToSqlDate()
            Dim sql As String
            Dim ds As New DataSet
            Dim dr As DataRow

            sql = "  SELECT Clientes.Cliente_Id AS Cliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, 'CONSOLIDADO' as Descricao" & vbCrLf & _
                  "   FROM Clientes " & vbCrLf & _
                  "  WHERE Clientes.Cliente_Id  ='" & Empresa(0) & "'" & vbCrLf & _
                  "    and Clientes.Endereco_id = " & Empresa(1) & vbCrLf

            ds = Banco.ConsultaDataSet(sql, "Clientes")
            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    NomeEmpresa = dr("Nome")
                    CidadeEmpresa = dr("Cidade")
                    EstadoEmpresa = dr("Estado")
                    'UnidadeDeNegocio = dr("Descricao")
                    Exit For
                Next
            End If


            If RbPagar.Checked = True Then
                NomeRelatorio = "Comparativo Financeiro x Contabil (Contas A Pagar) " & vbCrLf

                sql = "SELECT "

                If chkConsolidarCliente.Checked Then
                    sql &= " left(Consulta.Cliente,8) as Cliente, 0 as EndCliente, (select top 1 nome from clientes where left(cliente_id,8) = left(Consulta.Cliente,8)) as Nome,  '' as Cidade, '' as Estado,"
                Else
                    sql &= " Consulta.Cliente, Consulta.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf
                End If


                sql &= "       sum(Consulta.SdoRazao) as SdoRazao," & vbCrLf & _
                       "       Sum(Consulta.SdoTitulos) as SdoTitulos, " & vbCrLf & _
                       "       Sum(Consulta.SdoRazao - Consulta.SdoTitulos) as Diferenca " & vbCrLf & _
                       "  FROM (SELECT Cliente_Id AS Cliente, EndCliente_Id AS EndCliente, " & vbCrLf & _
                       "               SUM(DebitoOficial) - SUM(CreditoOficial) AS SdoRazao, 0 AS SdoTitulos" & vbCrLf & _
                       "          FROM Razao" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= "         WHERE left(Empresa_Id,8)    ='" & Empresa(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "         WHERE Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                          "           AND EndEmpresa_ID = " & Empresa(1) & vbCrLf
                End If

                If Cliente(0) <> "" Then
                    If chkConsolidarCliente.Checked Then
                        sql &= " and left(Cliente_id,8) ='" & Cliente(0).Substring(0, 8) & "'" & vbCrLf
                    Else
                        sql &= " and Cliente_id    ='" & Cliente(0) & "'" & vbCrLf & _
                               " and EndCliente_id = " & Cliente(1) & vbCrLf
                    End If
                End If

                sql &= " AND Conta_Id IN (''"

                If Chk_2010101.Checked Then sql &= ", '2010101'"
                If Chk_2010102.Checked Then sql &= ", '2010102'"
                If Chk_2010103.Checked Then sql &= ", '2010103'"
                If Chk_2010104.Checked Then sql &= ", '2010104'"
                If Chk_2010105.Checked Then sql &= ", '2010105'"
                If Chk_2010108.Checked Then sql &= ", '2010108'"
                If Chk_2010109.Checked Then sql &= ", '2010109'"

                sql &= ")" & vbCrLf & _
                       "   and razao.Movimento_Id <= '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf & _
                       " GROUP BY Cliente_Id, EndCliente_Id" & vbCrLf & _
                       " UNION " & vbCrLf & _
                       " SELECT Cliente, EndCliente, 0 AS SdoRAzao, SUM(ValorDoDocumento) AS SdoTitulos" & vbCrLf & _
                       "   FROM ContasAPagar" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= "         WHERE left(Empresa,8)    ='" & Empresa(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "         WHERE Empresa    ='" & Empresa(0) & "'" & vbCrLf & _
                          "           AND EndEmpresa = " & Empresa(1) & vbCrLf
                End If


                If Cliente(0) <> "" Then
                    If chkConsolidarCliente.Checked Then
                        sql &= " and left(Cliente,8) ='" & Cliente(0).Substring(0, 8) & "'" & vbCrLf
                    Else
                        sql &= " and Cliente    ='" & Cliente(0) & "'" & vbCrLf & _
                               " and EndCliente = " & Cliente(1) & vbCrLf
                    End If
                End If

                sql &= " and ContaContabilCliente IN (''"
                If Chk_2010101.Checked Then sql &= ", '2010101'"
                If Chk_2010102.Checked Then sql &= ", '2010102'"
                If Chk_2010103.Checked Then sql &= ", '2010103'"
                If Chk_2010104.Checked Then sql &= ", '2010104'"
                If Chk_2010105.Checked Then sql &= ", '2010105'"
                If Chk_2010108.Checked Then sql &= ", '2010108'"
                If Chk_2010109.Checked Then sql &= ", '2010109'"

                sql &= "        )" & vbCrLf & _
                       "        AND Provisao = 2 " & vbCrLf & _
                       "        AND Situacao = 1" & vbCrLf & _
                       "      GROUP BY Cliente, EndCliente" & vbCrLf & _
                       "      ) AS Consulta " & vbCrLf & _
                       " INNER JOIN Clientes " & vbCrLf & _
                       "    ON Consulta.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
                       "   AND Consulta.EndCliente = Clientes.Endereco_Id" & vbCrLf

                If ChkIsola.Checked = True Then
                    sql &= " where (Consulta.SdoRazao <> Consulta.SdoTitulos) "
                End If

                If chkConsolidarCliente.Checked Then
                    sql &= " Group by  left(Consulta.Cliente,8)"
                Else
                    sql &= " Group by Consulta.Cliente, Consulta.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf
                End If

                sql &= " Order by Nome" & vbCrLf
            Else
                NomeRelatorio = "Comparativo Financeiro x Contabil (Contas A Receber) "

                sql = "Select "
                If chkConsolidarCliente.Checked Then
                    sql &= " left(Consulta.Cliente,8) as Cliente, 0 as EndCliente, (select top 1 nome from clientes where left(cliente_id,8) = left(Consulta.Cliente,8)) as Nome,  '' as Cidade, '' as Estado,"
                Else
                    sql &= " Consulta.Cliente, Consulta.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf
                End If

                sql &= "       sum(Consulta.SdoRazao) as SdoRazao, Sum(Consulta.SdoTitulos) as SdoTitulos, " & vbCrLf & _
                       "       Sum(Consulta.SdoRazao - Consulta.SdoTitulos) as Diferenca " & vbCrLf & _
                       "  FROM (SELECT Cliente_Id AS Cliente, EndCliente_Id AS EndCliente, " & vbCrLf & _
                       "               SUM(DebitoOficial) - SUM (CreditoOficial) AS SdoRazao, " & vbCrLf & _
                       "               0 AS SdoTitulos" & vbCrLf & _
                       "          FROM Razao" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= "         WHERE left(Empresa_Id,8)    ='" & Empresa(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "         WHERE Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                          "           AND EndEmpresa_ID = " & Empresa(1) & vbCrLf
                End If

                If Cliente(0) <> "" Then
                    If chkConsolidarCliente.Checked Then
                        sql &= " and left(Cliente_id,8) ='" & Cliente(0).Substring(0, 8) & "'" & vbCrLf
                    Else
                        sql &= " and Cliente_id    ='" & Cliente(0) & "'" & vbCrLf & _
                               " and EndCliente_id = " & Cliente(1) & vbCrLf
                    End If
                End If

                sql &= " AND Conta_Id IN (''"

                If Chk_1010201.Checked Then sql &= ", '1010201'"
                If Chk_1010202.Checked Then sql &= ", '1010202'"
                If Chk_1010301.Checked Then sql &= ", '1010301'"
                If Chk_1010302.Checked Then sql &= ", '1010302'"

                sql &= "    )" & vbCrLf & _
                       "    and razao.Movimento_Id <= '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf & _
                       "  GROUP BY Cliente_Id, EndCliente_Id" & vbCrLf & _
                       "  UNION " & vbCrLf & _
                       " SELECT Cliente, EndCliente, 0 AS SdoRAzao, SUM(ValorDoDocumento) AS SdoTitulos" & vbCrLf & _
                       "   FROM ContasAReceber" & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= "         WHERE left(Empresa,8)    ='" & Empresa(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "         WHERE Empresa    ='" & Empresa(0) & "'" & vbCrLf & _
                          "           AND EndEmpresa = " & Empresa(1) & vbCrLf
                End If

                If Cliente(0) <> "" Then
                    If chkConsolidarCliente.Checked Then
                        sql &= " and left(Cliente,8) ='" & Cliente(0).Substring(0, 8) & "'" & vbCrLf
                    Else
                        sql &= " and Cliente    ='" & Cliente(0) & "'" & vbCrLf & _
                               " and EndCliente = " & Cliente(1) & vbCrLf
                    End If
                End If

                sql &= " and ContaContabilCliente IN (''"
                If Chk_1010201.Checked Then sql &= ", '1010201'"
                If Chk_1010202.Checked Then sql &= ", '1010202'"
                If Chk_1010301.Checked Then sql &= ", '1010301'"
                If Chk_1010302.Checked Then sql &= ", '1010302'"

                sql &= "           )" & vbCrLf & _
                       "         AND (Provisao = 2)" & vbCrLf & _
                       "         AND (Situacao = 1)" & vbCrLf & _
                       "       GROUP BY Cliente, EndCliente" & vbCrLf & _
                       "      ) AS Consulta " & vbCrLf & _
                       " INNER JOIN Clientes " & vbCrLf & _
                       "    ON Consulta.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
                       "   AND Consulta.EndCliente = Clientes.Endereco_Id" & vbCrLf

                If ChkIsola.Checked = True Then
                    sql &= " where (Consulta.SdoRazao <> Consulta.SdoTitulos) "
                End If

                If chkConsolidarCliente.Checked Then
                    sql &= " Group by left(Consulta.Cliente,8)"
                Else
                    sql &= " Group by Consulta.Cliente, Consulta.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf
                End If
                sql &= " Order by Nome"
            End If

            ds = Banco.ConsultaDataSet(sql, "Comparativo")

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Parametros", getParam())

            Funcoes.BindReport(Me.Page, ds, "cr_Comparativo", eExportType.PDF, parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub RelatorioHTML()
        Try
            Dim Empresa() As String = txtCodigoEmpresa.Value.Split("-")
            Dim Cliente() As String = txtCodigoCliente.Value.Split("-")
            Dim sql As String

            If RbPagar.Checked = True Then
                sql = "SELECT" & vbCrLf & _
                      " Consulta.Cliente, Consulta.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf & _
                      "sum(Consulta.SdoRazao) as SdoRazao, Sum(Consulta.SdoTitulos) as SdoTitulos, " & vbCrLf & _
                      "SUm(Consulta.SdoRazao - Consulta.SdoTitulos) as Diferenca " & vbCrLf & _
                      "FROM (SELECT     Cliente_Id AS Cliente, EndCliente_Id AS EndCliente, " & vbCrLf & _
                      "SUM  (CreditoOficial) - SUM(DebitoOficial) AS SdoRazao, 0 AS SdoTitulos" & vbCrLf & _
                      " FROM Razao" & vbCrLf & _
                      " WHERE (Empresa_Id = '" & Empresa(0) & "')" & vbCrLf & _
                      " AND (EndEmpresa_ID = " & Empresa(1) & ") " & vbCrLf
                If Cliente(0) <> "" Then
                    sql &= " and (Cliente_id = '" & Cliente(0) & "') " & vbCrLf & _
                           " and (EndCliente_id = " & Cliente(1) & ")" & vbCrLf
                End If
                sql &= " AND (Conta_Id IN (''"
                If Chk_2010101.Checked = True Then
                    sql &= ", '2010101'"
                End If
                If Chk_2010102.Checked = True Then
                    sql &= ", '2010102'"
                End If
                If Chk_2010103.Checked = True Then
                    sql &= ", '2010103'"
                End If
                If Chk_2010104.Checked = True Then
                    sql &= ", '2010104'"
                End If
                If Chk_2010105.Checked = True Then
                    sql &= ", '2010105'"
                End If
                If Chk_2010108.Checked = True Then
                    sql &= ", '2010108'"
                End If
                If Chk_2010109.Checked = True Then
                    sql &= ", '2010109'"
                End If
                sql &= ")) " & vbCrLf & _
                       " and razao.Movimento_Id <= '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf & _
                       " GROUP BY Cliente_Id, EndCliente_Id" & vbCrLf

                sql &= " UNION "

                sql &= " SELECT     Cliente, EndCliente, 0 AS SdoRAzao, SUM(ValorDoDocumento) AS SdoTitulos" & vbCrLf & _
                       " FROM ContasAPagar" & vbCrLf & _
                       " WHERE (Empresa = '" & Empresa(0) & "')" & vbCrLf & _
                       " AND (EndEmpresa = " & Empresa(1) & ") " & vbCrLf
                If Cliente(0) <> "" Then
                    sql &= " and (Cliente = '" & Cliente(0) & "') "   'Cliente
                    sql &= " and (EndCliente = " & Cliente(1) & ") "      'Cliente da Empresa
                End If
                sql &= " and (ContaContabilCliente IN (''"
                If Chk_2010101.Checked = True Then
                    sql &= ", '2010101'"
                End If
                If Chk_2010102.Checked = True Then
                    sql &= ", '2010102'"
                End If
                If Chk_2010103.Checked = True Then
                    sql &= ", '2010103'"
                End If
                If Chk_2010104.Checked = True Then
                    sql &= ", '2010104'"
                End If
                If Chk_2010105.Checked = True Then
                    sql &= ", '2010105'"
                End If
                If Chk_2010108.Checked = True Then
                    sql &= ", '2010108'"
                End If
                If Chk_2010109.Checked = True Then
                    sql &= ", '2010109'"
                End If
                sql &= ")) AND (Provisao = 2) AND (Situacao = 1)" & vbCrLf & _
                       " GROUP BY Cliente, EndCliente) AS Consulta INNER JOIN" & vbCrLf & _
                       " Clientes ON Consulta.Cliente = Clientes.Cliente_Id AND Consulta.EndCliente = Clientes.Endereco_Id" & vbCrLf
                If ChkIsola.Checked = True Then
                    sql &= " where (Consulta.SdoRazao <> Consulta.SdoTitulos) "
                End If
                sql &= " Group by Consulta.Cliente, Consulta.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf & _
                       " Order by Nome" & vbCrLf
            Else
                sql = "SELECT" & vbCrLf & _
                      " Consulta.Cliente, Consulta.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf & _
                      "sum(Consulta.SdoRazao) as SdoRazao, Sum(Consulta.SdoTitulos) as SdoTitulos, " & vbCrLf & _
                      "SUm(Consulta.SdoRazao - Consulta.SdoTitulos) as Diferenca " & vbCrLf & _
                      "FROM (SELECT     Cliente_Id AS Cliente, EndCliente_Id AS EndCliente, " & vbCrLf & _
                      "SUM  (CreditoOficial) - SUM(DebitoOficial) AS SdoRazao, 0 AS SdoTitulos" & vbCrLf & _
                      " FROM Razao" & vbCrLf & _
                      " WHERE (Empresa_Id = '" & Empresa(0) & "')" & vbCrLf & _
                      " AND (EndEmpresa_ID = " & Empresa(1) & ") " & vbCrLf
                If Cliente(0) <> "" Then
                    sql &= " and (Cliente_id = '" & Cliente(0) & "') "   'Cliente
                    sql &= " and (EndCliente_id = " & Cliente(1) & ")"      'Cliente da Empresa
                End If
                sql &= " AND (Conta_Id IN (''"
                If Chk_1010201.Checked Then sql &= ", '1010201'"
                If Chk_1010202.Checked Then sql &= ", '1010202'"
                If Chk_1010301.Checked = True Then
                    sql &= ", '1010301'"
                End If
                If Chk_1010302.Checked = True Then
                    sql &= ", '1010302'"
                End If

                sql &= ")) " & vbCrLf & _
                       " and razao.Movimento_Id <= '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'" & vbCrLf & _
                       " GROUP BY Cliente_Id, EndCliente_Id" & vbCrLf

                sql &= " UNION "

                sql &= " SELECT     Cliente, EndCliente, 0 AS SdoRAzao, SUM(ValorDoDocumento) AS SdoTitulos" & vbCrLf & _
                       " FROM ContasAReceber" & vbCrLf & _
                       " WHERE (Empresa = '" & Empresa(0) & "')" & vbCrLf & _
                       " AND (EndEmpresa = " & Empresa(1) & ") " & vbCrLf
                If Cliente(0) <> "" Then
                    sql &= " and (Cliente = '" & Cliente(0) & "') " & vbCrLf & _
                           " and (EndCliente = " & Cliente(1) & ") " & vbCrLf
                End If
                sql &= " and (ContaContabilCliente IN (''"
                If Chk_1010201.Checked Then sql &= ", '1010201'"
                If Chk_1010202.Checked Then sql &= ", '1010202'"
                If Chk_1010301.Checked = True Then
                    sql &= ", '1010301'"
                End If
                If Chk_1010302.Checked = True Then
                    sql &= ", '1010302'"
                End If

                sql &= " )) AND (Provisao = 2) AND (Situacao = 1)" & vbCrLf & _
                       " GROUP BY Cliente, EndCliente) AS Consulta INNER JOIN" & vbCrLf & _
                       " Clientes ON Consulta.Cliente = Clientes.Cliente_Id AND Consulta.EndCliente = Clientes.Endereco_Id" & vbCrLf
                If ChkIsola.Checked = True Then
                    sql &= " where (Consulta.SdoRazao <> Consulta.SdoTitulos) "
                End If
                sql &= " Group by Consulta.Cliente, Consulta.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf & _
                       " Order by Nome" & vbCrLf
            End If

            Dim linha As String
            Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
            Dim ds As New DataSet
            Dim dr As DataRow

            Dim strm As StreamWriter = Nothing
            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            ds = Banco.ConsultaDataSet(sql, "Clientes")

            linha = "<HTML>" & vbCrLf
            '<HEAD>
            linha &= "<HEAD>" & vbCrLf & _
                     "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf & _
                     "<TITLE>ComparativoFinanceiroXContabil</TITLE>" & vbCrLf & _
                     "</HEAD>" & vbCrLf

            '<BODY>
            linha &= "<BODY>" & vbCrLf

            '-----------------
            'Cabeçalho Padrao
            '-----------------
            linha &= "<table width= '4000' cellpadding='0' cellspacing='0' Border=1>"

            linha &= "<TR>" & vbCrLf & _
                     "<TD>Cliente</TD>" & vbCrLf & _
                     "<TD>EndCliente</TD>" & vbCrLf & _
                     "<TD>Nome</TD>" & vbCrLf & _
                     "<TD>Cidade</TD>" & vbCrLf & _
                     "<TD>Estado</TD>" & vbCrLf & _
                     "<TD>Valor Contabil</TD>" & vbCrLf & _
                     "<TD>Valor Financeiro</TD>" & vbCrLf & _
                     "<TD>Saldo (Contabil / Financeiro)</TD>" & vbCrLf & _
                     "</TR>" & vbCrLf


            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    linha &= "<TR><TD>" & Funcoes.FormatarCpfCnpj(dr("Cliente")) & "</TD>" & vbCrLf & _
                             "<TD>" & dr("EndCliente") & "</TD>" & vbCrLf & _
                             "<TD>" & dr("Nome") & "</TD>" & vbCrLf & _
                             "<TD>" & dr("Cidade") & "</TD>" & vbCrLf & _
                             "<TD>" & dr("Estado") & "</TD>" & vbCrLf & _
                             "<TD>" & dr("SdoRazao") & "</TD>" & vbCrLf & _
                             "<TD>" & dr("SdoTitulos") & "</TD>" & vbCrLf & _
                             "<TD>" & dr("Diferenca") & "</TD>" & vbCrLf & _
                             "</TR>" & vbCrLf
                Next
            End If
            Try
                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
                ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
            Finally
                strm.Close()
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ComparativoFinanceiroXContabil")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class