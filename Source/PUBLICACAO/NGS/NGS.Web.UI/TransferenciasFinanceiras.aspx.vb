Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TransferenciasFinanceiras
    Inherits BasePage

    Dim Sql As String

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TransferenciasFinanceiras", "ACESSAR") Then
                CarregarContas()
                BindGridView()
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaEmpresas.SetarHID(HID.Value)
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub BindGridView()
        Sql = "SELECT TransferenciasFinanceiras.EmpresaDebito, TransferenciasFinanceiras.EnderecoDebito, " & vbCrLf & _
              "TransferenciasFinanceiras.EmpresaCredito, TransferenciasFinanceiras.EnderecoCredito, " & vbCrLf & _
              "TransferenciasFinanceiras.EmpresaContabil, TransferenciasFinanceiras.EnderecoContabil, " & vbCrLf & _
              "TransferenciasFinanceiras.ContaContabil, ISNULL(TransferenciasFinanceiras.ClienteContabil, '') AS " & vbCrLf & _
              "ClienteContabil, ISNULL(TransferenciasFinanceiras.EndClienteContabil, '') AS EndClienteContabil, " & vbCrLf & _
              "TransferenciasFinanceiras.DebitoCredito, TransferenciasFinanceiras.EmiteAviso, " & vbCrLf & _
              "EmpresaDebito.Nome AS NomeEmpDebito, EmpresaDebito.Reduzido AS ReduzidoEmpDebito, " & vbCrLf & _
              "EmpresaCredito.Nome AS NomeEmpCredito, EmpresaCredito.Reduzido AS ReduzidoEmpCredito, " & vbCrLf & _
              "EmpresaContabil.Nome AS NomeEmpContabil, EmpresaContabil.Reduzido AS ReduzidoEmpContabil, " & vbCrLf & _
              "PlanoDeContas.Titulo, ISNULL(ClienteContabil.Nome, '') AS NomeCliContabil, " & vbCrLf & _
              "ISNULL(ClienteContabil.Reduzido, '') AS ReduzidoCliContabil " & vbCrLf & _
              "FROM TransferenciasFinanceiras INNER JOIN " & vbCrLf & _
              "Clientes AS EmpresaDebito ON TransferenciasFinanceiras.EmpresaDebito = EmpresaDebito.Cliente_Id AND " & vbCrLf & _
              "TransferenciasFinanceiras.EnderecoDebito = EmpresaDebito.Endereco_Id INNER JOIN " & vbCrLf & _
              "Clientes AS EmpresaCredito ON TransferenciasFinanceiras.EmpresaCredito = EmpresaCredito.Cliente_Id AND " & vbCrLf & _
              "TransferenciasFinanceiras.EnderecoCredito = EmpresaCredito.Endereco_Id INNER JOIN " & vbCrLf & _
              "Clientes AS EmpresaContabil ON TransferenciasFinanceiras.EmpresaContabil = EmpresaContabil.Cliente_Id AND " & vbCrLf & _
              "TransferenciasFinanceiras.EnderecoContabil = EmpresaContabil.Endereco_Id INNER JOIN " & vbCrLf & _
              "PlanoDeContas ON TransferenciasFinanceiras.ContaContabil = PlanoDeContas.Conta_Id LEFT OUTER JOIN " & vbCrLf & _
              "Clientes AS ClienteContabil ON TransferenciasFinanceiras.ClienteContabil = ClienteContabil.Cliente_Id AND " & vbCrLf & _
              "TransferenciasFinanceiras.EndClienteContabil = ClienteContabil.Endereco_Id " & vbCrLf & _
              "ORDER BY TransferenciasFinanceiras.EmpresaDebito" & vbCrLf
        grd.DataSource = Banco.ConsultaDataSet(Sql, "TransferenciasFinanceiras")
        grd.DataBind()
    End Sub

    Private Sub Relatorio()
        Try
            If Funcoes.VerificaPermissao("TransferenciasFinanceiras", "RELATORIO") Then
                Sql = "  SELECT     TransferenciasFinanceiras.EmpresaDebito, TransferenciasFinanceiras.EnderecoDebito, TransferenciasFinanceiras.EmpresaCredito, " & vbCrLf & _
                      "  TransferenciasFinanceiras.EnderecoCredito, TransferenciasFinanceiras.EmpresaContabil, TransferenciasFinanceiras.EnderecoContabil, " & vbCrLf & _
                      "  TransferenciasFinanceiras.ContaContabil, TransferenciasFinanceiras.ClienteContabil, TransferenciasFinanceiras.EndClienteContabil, " & vbCrLf & _
                      "  TransferenciasFinanceiras.DebitoCredito, TransferenciasFinanceiras.EmiteAviso, EmpresaDebito.Estado AS EstadoDebito, EmpresaDebito.Nome AS NomeDebito, " & vbCrLf & _
                      "  EmpresaDebito.Cidade AS CidadeDebito, EmpresaDebito.Reduzido AS ReduzidoDebito, EmpresaCredito.Estado AS EstadoCredito, " & vbCrLf & _
                      "  EmpresaCredito.Nome AS NomeCredito, EmpresaCredito.Cidade AS CidadeCredito, EmpresaCredito.Reduzido AS ReduzidoCredito, " & vbCrLf & _
                      "  EmpresaContabil.Estado AS EstadoContabil, EmpresaContabil.Nome AS NomeContabil, EmpresaContabil.Cidade AS CidadeContabil, " & vbCrLf & _
                      "  EmpresaContabil.Reduzido AS ReduzidoContabil, PlanoDeContas.Titulo AS NomeContaContabil, ISNULL(ClienteContabil.Estado, '') AS EstadoCliente, " & vbCrLf & _
                      "  ISNULL(ClienteContabil.Nome, '') AS NomeCliente, ISNULL(ClienteContabil.Cidade, '') AS CidadeCliente, ISNULL(ClienteContabil.Reduzido, '') AS ReduzidoCliente " & vbCrLf & _
                      "  FROM         TransferenciasFinanceiras INNER JOIN " & vbCrLf & _
                      "  Clientes AS EmpresaDebito ON TransferenciasFinanceiras.EmpresaDebito = EmpresaDebito.Cliente_Id AND " & vbCrLf & _
                      "  TransferenciasFinanceiras.EnderecoDebito = EmpresaDebito.Endereco_Id INNER JOIN " & vbCrLf & _
                      "  Clientes AS EmpresaCredito ON TransferenciasFinanceiras.EmpresaCredito = EmpresaCredito.Cliente_Id AND " & vbCrLf & _
                      "  TransferenciasFinanceiras.EnderecoCredito = EmpresaCredito.Endereco_Id INNER JOIN " & vbCrLf & _
                      "  Clientes AS EmpresaContabil ON TransferenciasFinanceiras.EmpresaContabil = EmpresaContabil.Cliente_Id AND " & vbCrLf & _
                      "  TransferenciasFinanceiras.EnderecoContabil = EmpresaContabil.Endereco_Id LEFT OUTER JOIN " & vbCrLf & _
                      "  PlanoDeContas ON TransferenciasFinanceiras.ContaContabil = PlanoDeContas.Conta_Id LEFT OUTER JOIN " & vbCrLf & _
                      "  Clientes AS ClienteContabil ON TransferenciasFinanceiras.ClienteContabil = ClienteContabil.Cliente_Id AND " & vbCrLf & _
                      "  TransferenciasFinanceiras.EndClienteContabil = ClienteContabil.Endereco_Id " & vbCrLf & _
                      "  ORDER BY TransferenciasFinanceiras.EmpresaDebito" & vbCrLf

                Dim ds As New DataSet
                ds = Banco.ConsultaDataSet(Sql, "TransferenciasFinanceiras")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Transferências Financeiras")
                parameters.Add("ConsultaParam", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_TransferenciasFinanceiras", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParam() As String
        Dim Texto As String = ""
        Dim obj = New Cliente()

        If Not String.IsNullOrWhiteSpace(txtEmpresaDebito.Text) Then
            obj = New Cliente(hdfEmpresaDebito.Value.Split("-")(0), hdfEmpresaDebito.Value.Split("-")(1))
            Texto &= "Empresa Débito: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & "/" & obj.Estado.Codigo & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtEmpresaCredito.Text) Then
            obj = New Cliente(hdfEmpresaCredito.Value.Split("-")(0), hdfEmpresaCredito.Value.Split("-")(1))
            Texto &= "Empresa Crédito: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & "/" & obj.Estado.Codigo & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtEmpresaContabil.Text) Then
            obj = New Cliente(hdfEmpresaContabil.Value.Split("-")(0), hdfEmpresaContabil.Value.Split("-")(1))
            Texto &= "Empresa Contábil: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & "/" & obj.Estado.Codigo & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtClienteContabil.Text) Then
            obj = New Cliente(hdfClienteContabil.Value.Split("-")(0), hdfClienteContabil.Value.Split("-")(1))
            Texto &= "Cliente Contábil: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & "/" & obj.Estado.Codigo & vbCrLf
        End If
        Texto &= "Tipo de Operação: " & IIf(rdoDebito.Checked, "Débito", "Crédito")
        Texto &= " - Opção: " & IIf(rdoEmiteAvisoA.Checked, "Emite Aviso", "Não Emite Aviso")

        Return Texto
    End Function

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaTFED" & HID.Value) Is Nothing Then
            Dim objTfed As [Lib].Negocio.Cliente = CType(Session("objEmpresaTFED" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objTfed)
            hdfEmpresaDebito.Value = itemCliente.Value
            txtEmpresaDebito.Text = itemCliente.Text
            Session.Remove("objEmpresaTFED" & HID.Value)
        ElseIf Not Session("objEmpresaTFEC" & HID.Value) Is Nothing Then
            Dim objTfec As [Lib].Negocio.Cliente = CType(Session("objEmpresaTFEC" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objTfec)
            hdfEmpresaCredito.Value = itemCliente.Value
            txtEmpresaCredito.Text = itemCliente.Text
            Session.Remove("objEmpresaTFEC" & HID.Value)
        ElseIf Not Session("objEmpresaTFC" & HID.Value) Is Nothing Then
            Dim objTfc As [Lib].Negocio.Cliente = CType(Session("objEmpresaTFC" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objTfc)
            hdfEmpresaContabil.Value = itemCliente.Value
            txtEmpresaContabil.Text = itemCliente.Text
            Session.Remove("objEmpresaTFC" & HID.Value)
        ElseIf Not Session("objClienteTFC" & HID.Value) Is Nothing Then
            Dim objClienteTFC As [Lib].Negocio.Cliente = CType(Session("objClienteTFC" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objClienteTFC)
            hdfClienteContabil.Value = itemCliente.Value
            txtClienteContabil.Text = itemCliente.Text
            Session.Remove("objClienteTFC" & HID.Value)
        End If
    End Sub

    Private Sub CarregarContas()
        ddl.Carregar(ddlContaContabil, CarregarDDL.Tabela.PlanoDeContas, "", True)
    End Sub

    Private Sub Limpar()
        hdfEmpresaDebito.Value = ""
        txtEmpresaDebito.Text = ""
        hdfEmpresaCredito.Value = ""
        txtEmpresaCredito.Text = ""
        hdfEmpresaContabil.Value = ""
        txtEmpresaContabil.Text = ""
        hdfClienteContabil.Value = ""
        txtClienteContabil.Text = ""
        ddlContaContabil.SelectedIndex = 0
        BindGridView()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Public Sub Incluir(ByVal EmpresaDebito As String, ByVal EnderecoDebito As String, ByVal EmpresaCredito As String, ByVal EnderecoCredito As String, ByVal EmpresaContabil As String, ByVal EnderecoContabil As String, ByVal ContaContabil As String, ByVal ClienteContabil As String, ByVal EndClienteContabil As String, ByVal DebitoCredito As String, ByVal EmiteAviso As String)
        Try
            Dim SqlArray As New ArrayList

            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, ".", "")
            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, "/", "")
            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, "-", "")

            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, ".", "")
            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, "/", "")
            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, "-", "")

            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, ".", "")
            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, "/", "")
            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, "-", "")

            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, ".", "")
            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, "/", "")
            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, "-", "")

            If EndClienteContabil = "" Then
                EndClienteContabil = "0"
            End If

            Sql = "Insert Into TransferenciasFinanceiras (EmpresaDebito, EnderecoDebito, EmpresaCredito, EnderecoCredito ," & vbCrLf & _
                  "EmpresaContabil, EnderecoContabil, ContaContabil, ClienteContabil, EndClienteContabil, DebitoCredito, EmiteAviso) " & vbCrLf & _
                  "Values('" & EmpresaDebito & "'," & EnderecoDebito & ",'" & EmpresaCredito & "'," & EnderecoCredito & "," & vbCrLf & _
                  "'" & EmpresaContabil & "'," & EnderecoContabil & ",'" & ContaContabil & "','" & ClienteContabil & "'," & vbCrLf & _
                  "" & EndClienteContabil & ",'" & DebitoCredito & "','" & EmiteAviso & "')" & vbCrLf
            SqlArray.Add(Sql)

            If Not Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub Alterar(ByVal EmpresaDebito As String, ByVal EnderecoDebito As String, ByVal EmpresaCredito As String, ByVal EnderecoCredito As String, ByVal EmpresaContabil As String, ByVal EnderecoContabil As String, ByVal ContaContabil As String, ByVal ClienteContabil As String, ByVal EndClienteContabil As String, ByVal DebitoCredito As String, ByVal EmiteAviso As String)
        Try
            Dim SqlArray As New ArrayList

            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, ".", "")
            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, "/", "")
            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, "-", "")

            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, ".", "")
            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, "/", "")
            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, "-", "")

            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, ".", "")
            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, "/", "")
            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, "-", "")

            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, ".", "")
            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, "/", "")
            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, "-", "")

            Sql = "Update TransferenciasFinanceiras Set DebitoCredito = '" & DebitoCredito & "', EmiteAviso = '" & EmiteAviso & "' " & vbCrLf & _
                  "WHERE EmpresaDebito = '" & EmpresaDebito & "' AND EnderecoDebito = " & EnderecoDebito & " AND " & vbCrLf & _
                  "EmpresaCredito = '" & EmpresaCredito & "' AND EnderecoCredito = " & EnderecoCredito & " AND " & vbCrLf & _
                  "EmpresaContabil = '" & EmpresaContabil & "' AND EnderecoContabil = " & EnderecoContabil & " AND " & vbCrLf & _
                  "ContaContabil = '" & ContaContabil & "' AND ClienteContabil = '" & ClienteContabil & "' AND " & vbCrLf & _
                  "EndClienteContabil = " & EndClienteContabil & "" & vbCrLf
            SqlArray.Add(Sql)

            If Not Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub Excluir(ByVal EmpresaDebito As String, ByVal EnderecoDebito As String, ByVal EmpresaCredito As String, ByVal EnderecoCredito As String, ByVal EmpresaContabil As String, ByVal EnderecoContabil As String, ByVal ContaContabil As String, ByVal ClienteContabil As String, ByVal EndClienteContabil As String, ByVal DebitoCredito As String, ByVal EmiteAviso As String)
        Try
            Dim SqlArray As New ArrayList

            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, ".", "")
            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, "/", "")
            EmpresaDebito = Microsoft.VisualBasic.Replace(EmpresaDebito, "-", "")

            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, ".", "")
            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, "/", "")
            EmpresaCredito = Microsoft.VisualBasic.Replace(EmpresaCredito, "-", "")

            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, ".", "")
            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, "/", "")
            EmpresaContabil = Microsoft.VisualBasic.Replace(EmpresaContabil, "-", "")

            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, ".", "")
            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, "/", "")
            ClienteContabil = Microsoft.VisualBasic.Replace(ClienteContabil, "-", "")

            Sql = "Delete From TransferenciasFinanceiras WHERE EmpresaDebito = '" & EmpresaDebito & "' AND " & vbCrLf & _
                  "EnderecoDebito = " & EnderecoDebito & " And EmpresaCredito = '" & EmpresaCredito & "' AND " & vbCrLf & _
                  "EnderecoCredito = " & EnderecoCredito & " And EmpresaContabil = '" & EmpresaContabil & "' AND " & vbCrLf & _
                  "EnderecoContabil = " & EnderecoContabil & " And ContaContabil = '" & ContaContabil & "' AND " & vbCrLf & _
                  "ClienteContabil = '" & ClienteContabil & "' AND EndClienteContabil = " & EndClienteContabil & "" & vbCrLf
            SqlArray.Add(Sql)

            If Not Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grd_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grd.SelectedIndexChanged
        Try
            Dim empDebito As Cliente = New Cliente(grd.SelectedRow.Cells(1).Text(), grd.SelectedRow.Cells(2).Text())
            Dim empCredito As Cliente = New Cliente(grd.SelectedRow.Cells(3).Text(), grd.SelectedRow.Cells(4).Text())
            Dim empContabil As Cliente = New Cliente(grd.SelectedRow.Cells(5).Text(), grd.SelectedRow.Cells(6).Text())
            Dim contaContabil As String = grd.SelectedRow.Cells(7).Text()
            Dim cliContabil As Cliente = New Cliente(grd.SelectedRow.Cells(8).Text(), grd.SelectedRow.Cells(9).Text())
            Dim debCredito As String = grd.SelectedRow.Cells(10).Text()
            Dim aviso As String = grd.SelectedRow.Cells(11).Text()
            hdfEmpresaDebito.Value = Funcoes.FormatarListItemCliente(empDebito).Value
            txtEmpresaDebito.Text = Funcoes.FormatarListItemCliente(empDebito).Text
            hdfEmpresaCredito.Value = Funcoes.FormatarListItemCliente(empCredito).Value
            txtEmpresaCredito.Text = Funcoes.FormatarListItemCliente(empCredito).Text
            hdfEmpresaContabil.Value = Funcoes.FormatarListItemCliente(empContabil).Value
            txtEmpresaContabil.Text = Funcoes.FormatarListItemCliente(empContabil).Text
            hdfClienteContabil.Value = Funcoes.FormatarListItemCliente(cliContabil).Value
            txtClienteContabil.Text = Funcoes.FormatarListItemCliente(cliContabil).Text
            ddlContaContabil.SelectedValue = contaContabil
            rdoCredito.Checked = debCredito = "C"
            rdoDebito.Checked = debCredito = "D"
            rdoEmiteAvisoA.Checked = aviso = "S"
            rdoEmiteAvisoB.Checked = aviso = "N"
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnMostraEmpresaDebito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnMostraEmpresaDebito.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaTFED" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnLimpaEmpresaDebito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLimpaEmpresaDebito.Click
        Try
            hdfEmpresaDebito.Value = ""
            txtEmpresaDebito.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnMostraEmpresaCredito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnMostraEmpresaCredito.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaTFEC" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnLimpaEmpresaCredito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLimpaEmpresaCredito.Click
        Try
            hdfEmpresaCredito.Value = ""
            txtEmpresaCredito.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnMostraEmpresaContabil_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnMostraEmpresaContabil.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaTFC" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnLimpaEmpresaContabil_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLimpaEmpresaContabil.Click
        Try
            hdfEmpresaContabil.Value = ""
            txtEmpresaContabil.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnMostraClienteContabil_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnMostraClienteContabil.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteTFC" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnLimpaClienteContabil_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLimpaClienteContabil.Click
        Try
            hdfClienteContabil.Value = ""
            txtClienteContabil.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("TransferenciasFinanceiras", "GRAVAR") Then
                Dim empDebito() As String = hdfEmpresaDebito.Value.Split("-")
                Dim empCredito() As String = hdfEmpresaCredito.Value.Split("-")
                Dim empContabil() As String = hdfEmpresaContabil.Value.Split("-")
                Dim contaContabil As String = ddlContaContabil.SelectedValue
                Dim cliContabil() As String = hdfClienteContabil.Value.Split("-")
                Dim debCredito As String = IIf(rdoCredito.Checked, "C", "D")
                Dim aviso As String = IIf(rdoEmiteAvisoA.Checked, "S", "N")
                Incluir(empDebito(0), empDebito(1), empCredito(0), empCredito(1), empContabil(0), empContabil(1), contaContabil, cliContabil(0), cliContabil(1), debCredito, aviso)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("TransferenciasFinanceiras", "ALTYERAR") Then
                Dim empDebito() As String = hdfEmpresaDebito.Value.Split("-")
                Dim empCredito() As String = hdfEmpresaCredito.Value.Split("-")
                Dim empContabil() As String = hdfEmpresaContabil.Value.Split("-")
                Dim contaContabil As String = ddlContaContabil.SelectedValue
                Dim cliContabil() As String = hdfClienteContabil.Value.Split("-")
                Dim debCredito As String = IIf(rdoCredito.Checked, "C", "D")
                Dim aviso As String = IIf(rdoEmiteAvisoA.Checked, "S", "N")
                Alterar(empDebito(0), empDebito(1), empCredito(0), empCredito(1), empContabil(0), empContabil(1), contaContabil, cliContabil(0), cliContabil(1), debCredito, aviso)
            Else
                MsgBox(Me.Page, "Usuário sem permissã para alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("TransferenciasFinanceiras", "EXCLUIR") Then
                Dim empDebito() As String = hdfEmpresaDebito.Value.Split("-")
                Dim empCredito() As String = hdfEmpresaCredito.Value.Split("-")
                Dim empContabil() As String = hdfEmpresaContabil.Value.Split("-")
                Dim contaContabil As String = ddlContaContabil.SelectedValue
                Dim cliContabil() As String = hdfClienteContabil.Value.Split("-")
                Dim debCredito As String = IIf(rdoCredito.Checked, "C", "D")
                Dim aviso As String = IIf(rdoEmiteAvisoA.Checked, "S", "N")
                Excluir(empDebito(0), empDebito(1), empCredito(0), empCredito(1), empContabil(0), empContabil(1), contaContabil, cliContabil(0), cliContabil(1), debCredito, aviso)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            Relatorio()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "TransferenciasFinanceiras")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class