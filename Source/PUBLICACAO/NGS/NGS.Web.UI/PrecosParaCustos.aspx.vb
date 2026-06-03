Imports System.Data
Imports System.Linq
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PrecosParaCustos
    Inherits BasePage

#Region "Váriavies locais / Auxiliares"

    Dim Sql As String
    Private objEmpresa As [Lib].Negocio.Cliente

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Custos)
                If Funcoes.VerificaPermissao("PrecosParaCustos", "Acessar") Then
                    Limpar("", "")
                    CarregarGrupo()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPesquisarEmpresa_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPesquisarEmpresa.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me, "objPrecosParaCustos" & HID.Value.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPesquisarDeposito_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnPesquisarDeposito.Click
        Try
            ucConsultaEmpresas.Limpar()
            Session("Deposito" & HID.Value.ToString) = True
            Popup.ConsultaDeEmpresas(Me, "objPrecosParaCustos" & HID.Value.ToString)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAjuda.Click
        Try
            Dim NomeArquivo As String = "Manual/PrecosParaCustos.mht"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
            ScriptManager.RegisterClientScriptBlock(Me, Me.lnkAjuda.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlGrupo.SelectedIndexChanged
        Try
            CarregarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Public Sub Limpar(ByVal evento As String, ByVal Numero As String)
        Session.Remove("objPrecosParaCustos" & HID.Value.ToString)
        Session.Remove("Deposito" & HID.Value.ToString)
        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        objEmpresa = New [Lib].Negocio.Cliente
        objEmpresa = CType(Session("objPrecosParaCustos" & HID.Value.ToString), [Lib].Negocio.Cliente)

        '1 = Empresa
        If objEmpresa.Tipos.Where(Function(s) s.TipoDeCliente IsNot Nothing).Select(Function(s) s.TipoDeCliente.CodigoTipo).Contains(1) And Session("Deposito" & HID.Value.ToString) Is Nothing Then
            With objEmpresa
                txtReduzidoEmp.Text = .Reduzido
                txtCNPJEmp.Text = .CodigoFormatado.ToString
                txtCodigoEnderecoEmp.Text = .CodigoEndereco
                txtNomeEmp.Text = .Nome
            End With
        Else
            CarregarDeposito() 'DEPOSITOS
        End If
    End Sub

    Public Sub CarregarDeposito()
        objEmpresa = New [Lib].Negocio.Cliente
        objEmpresa = CType(Session("objPrecosParaCustos" & HID.Value.ToString), [Lib].Negocio.Cliente)
        '3 = DESPOSITOS
        If objEmpresa.Tipos.Where(Function(s) s.TipoDeCliente IsNot Nothing).Select(Function(s) s.TipoDeCliente.CodigoTipo).Contains(3) Then
            With objEmpresa
                txtReduzidoDep.Text = .Reduzido
                txtCNPJDep.Text = .CodigoFormatado.ToString
                txtCodigoEnderecoDep.Text = .CodigoEndereco
                txtNomeDep.Text = .Nome
            End With
        Else
            MsgBox(Me.Page, "A Empresa não é do TIPO: DEPOSITOS! Selecione outra empresa.")
            Popup.ConsultaDeEmpresas(Me, "objPrecosParaCustos" & HID.Value.ToString)
        End If
    End Sub

    Private Sub CargaProduto()
        Dim Produtos As New [Lib].Negocio.ListProduto("")
        Dim strProduto As String

        ddlProduto.Items.Clear()

        For Each dr As [Lib].Negocio.Produto In Produtos
            'strProduto = Funcoes.AlinharEsquerda(dr.Nome, 50, ".") & " - " & dr.Codigo
            strProduto = dr.Codigo & " - " & dr.Nome
            ddlProduto.Items.Add(New ListItem(strProduto, dr.Codigo))
        Next
        Funcoes.InserirLinhaEmBranco(ddlProduto)
    End Sub

    Private Sub CarregarGrupo()
        ddl.Carregar(ddlGrupo, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub CarregarProduto()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, " GRUPO = '" & ddlGrupo.SelectedValue & "'", True)
    End Sub

    Protected Sub GridEstados_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar("", "")
            ddlGrupo.SelectedIndex = grdAbaTitulo.SelectedRow.Cells(1).Text()
            ddlProduto.SelectedIndex = grdAbaTitulo.SelectedRow.Cells(2).Text()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    '<Ajax.AjaxMethod(Ajax.HttpSessionStateRequirement.ReadWrite)> _
    Public Function Saida()
        Return HttpContext.Current.Session("ssMenuDeAcesso")
    End Function

    '<Ajax.AjaxMethod(Ajax.HttpSessionStateRequirement.ReadWrite)> _
    Public Function Atualiza_Html()
        Dim Texto As String = ""
        'Dim Valor As String

        Sql = "SELECT PrecosParaCustos.Empresa_Id, PrecosParaCustos.EndEmpresa_Id, Empresa.Nome AS NomeEmpresa, " & vbCrLf & _
              "Empresa.Reduzido AS ReduzidoEmpresa, PrecosParaCustos.Deposito_Id, PrecosParaCustos.EndDeposito_Id, " & vbCrLf & _
              "Deposito.Nome AS NomeDeposito, Deposito.Reduzido AS ReduzidoDeposito, PrecosParaCustos.Produto_Id, " & vbCrLf & _
              "Produtos.Nome AS NomeProduto, PrecosParaCustos.Data_Id, PrecosParaCustos.Fator, PrecosParaCustos.Preco " & vbCrLf & _
              "FROM PrecosParaCustos INNER JOIN " & vbCrLf & _
              "Clientes AS Empresa ON PrecosParaCustos.Empresa_Id = Empresa.Cliente_Id AND " & vbCrLf & _
              "PrecosParaCustos.EndEmpresa_Id = Empresa.Endereco_Id INNER JOIN " & vbCrLf & _
              "Clientes AS Deposito ON PrecosParaCustos.Deposito_Id = Deposito.Cliente_Id AND " & vbCrLf & _
              "PrecosParaCustos.EndDeposito_Id = Deposito.Endereco_Id INNER JOIN " & vbCrLf & _
              "Produtos ON PrecosParaCustos.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
              "ORDER BY PrecosParaCustos.Empresa_Id, PrecosParaCustos.EndEmpresa_Id, PrecosParaCustos.Deposito_Id, " & vbCrLf & _
              "PrecosParaCustos.EndDeposito_Id, PrecosParaCustos.Produto_Id, PrecosParaCustos.Data_Id" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "PrecosParaCustos").Tables(0).Rows
            Texto &= Dr("Empresa_Id") & StrDup(18 - Len(Dr("Empresa_Id")), " ") & "|"
            Texto &= Dr("EndEmpresa_Id") & "|"
            Texto &= Dr("NomeEmpresa") & "|"
            Texto &= Dr("ReduzidoEmpresa") & "|"
            Texto &= Dr("Deposito_Id") & StrDup(18 - Len(Dr("Deposito_Id")), " ") & "|"
            Texto &= Dr("EndDeposito_Id") & "|"
            Texto &= Dr("NomeDeposito") & "|"
            Texto &= Dr("ReduzidoDeposito") & "|"
            Texto &= Dr("Produto_Id") & StrDup(18 - Len(Dr("Produto_Id")), " ") & "|"
            Texto &= Dr("NomeProduto") & "|"
            Texto &= Format(Dr("Data_Id"), "dd/MM/yyyy") & "|"
            Texto &= CStr(Dr("Fator")) & StrDup(10 - Len(CStr(Dr("Fator"))), " ") & "|"
            Texto &= StrDup(15 - Len(CStr(Format(Dr("Preco"), "###,###,##0.00"))), " ") & CStr(Format(Dr("Preco"), "###,###,##0.00")) & "]"
        Next

        Return Texto

    End Function

    '<Ajax.AjaxMethod(Ajax.HttpSessionStateRequirement.ReadWrite)> _
    Public Function ExecutarSQL(ByVal Registro As String)
        Dim SqlArray As New ArrayList
        Dim Campo As String() = Registro.Split("|")

        If Campo(0).Length > 0 Then

            Dim strSQL As String = ""

            Dim Opcao As String = Campo(0)
            Dim Empresa As String = RTrim(Campo(1))
            Empresa = Replace(Empresa, ".", "")
            Empresa = Replace(Empresa, "/", "")
            Empresa = Replace(Empresa, "-", "")
            Dim EndEmpresa As String = Campo(2)
            Dim Deposito As String = RTrim(Campo(3))
            Deposito = Replace(Deposito, ".", "")
            Deposito = Replace(Deposito, "/", "")
            Deposito = Replace(Deposito, "-", "")
            Dim EndDeposito As String = Campo(4)
            Dim Produto As String = RTrim(Campo(5))
            Dim Data As String = Format(CDate(Campo(6)), "yyyy/MM/dd")
            Dim Preco As String = Campo(7)
            Preco = Replace(Preco, ".", "")
            Preco = Replace(Preco, ",", ".")
            Dim Fator As String = Campo(8)

            Select Case Opcao
                Case "Incluir"
                    strSQL = "INSERT INTO PrecosParaCustos " & vbCrLf & _
                             "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Data_Id, Preco, Fator) VALUES " & vbCrLf & _
                             "('" & Empresa & "', " & vbCrLf & _
                             EndEmpresa & ", " & vbCrLf & _
                             "'" & Deposito & "', " & vbCrLf & _
                             EndDeposito & ", " & vbCrLf & _
                             "'" & Produto & "', " & vbCrLf & _
                             "'" & Data & "', " & vbCrLf & _
                             Preco & ", " & vbCrLf & _
                             Fator & ")" & vbCrLf

                Case "Alterar"
                    strSQL = "UPDATE PrecosParaCustos SET " & vbCrLf & _
                             "Preco = " & Preco & ", " & vbCrLf & _
                             "Fator = " & Fator & vbCrLf & _
                             " WHERE (" & vbCrLf & _
                             "Empresa_Id = '" & Empresa & "' AND " & vbCrLf & _
                             "EndEmpresa_Id = " & EndEmpresa & " AND " & vbCrLf & _
                             "Deposito_Id = '" & Deposito & "' AND " & vbCrLf & _
                             "EndDeposito_Id = " & EndDeposito & " AND " & vbCrLf & _
                             "Produto_Id = '" & Produto & "' AND " & vbCrLf & _
                             "Data_Id = '" & Data & "') " & vbCrLf

                Case "Excluir"
                    strSQL = "DELETE FROM PrecosParaCustos " & vbCrLf & _
                             " WHERE (" & vbCrLf & _
                             "Empresa_Id = '" & Empresa & "' AND " & vbCrLf & _
                             "EndEmpresa_Id = " & EndEmpresa & " AND " & vbCrLf & _
                             "Deposito_Id = '" & Deposito & "' AND " & vbCrLf & _
                             "EndDeposito_Id = " & EndDeposito & " AND " & vbCrLf & _
                             "Produto_Id = '" & Produto & "' AND " & vbCrLf & _
                             "Data_Id = '" & Data & "') " & vbCrLf
            End Select

            SqlArray.Add(strSQL)
            Return Banco.GravaBanco(SqlArray)
        Else
            Return "ERRO"
        End If
    End Function

    Public Function Relatorio()
        Dim Ds_PrecosParaCustos As New DataSet
        Dim mensagem As String = ""

        Sql = "SELECT PrecosParaCustos.Empresa_Id, PrecosParaCustos.EndEmpresa_Id, " & vbCrLf & _
              "PrecosParaCustos.Deposito_Id, PrecosParaCustos.EndDeposito_Id, " & vbCrLf & _
              "PrecosParaCustos.Produto_Id, PrecosParaCustos.Data_Id, " & vbCrLf & _
              "PrecosParaCustos.Preco, PrecosParaCustos.Fator, Produtos.Nome AS NomeProd, " & vbCrLf & _
              "Clientes.Nome AS NomeEmp, Clientes.Cidade As CidadeEmp, Clientes.Estado AS EstadoEmp, " & vbCrLf & _
              "Deposito.Nome AS NomeDep,Deposito.Cidade AS CidadeDep,Deposito.Estado AS EstadoDep " & vbCrLf & _
              "FROM  PrecosParaCustos INNER JOIN " & vbCrLf & _
              "Produtos ON PrecosParaCustos.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
              "INNER JOIN " & vbCrLf & _
              "Clientes ON PrecosParaCustos.Empresa_Id = Clientes.Cliente_Id AND " & vbCrLf & _
              "PrecosParaCustos.EndEmpresa_Id = Clientes.Endereco_Id " & vbCrLf & _
              "INNER JOIN " & vbCrLf & _
              " Clientes as Deposito ON PrecosParaCustos.Deposito_Id = " & vbCrLf & _
              "Deposito.Cliente_Id AND " & vbCrLf & _
              "PrecosParaCustos.EndDeposito_Id = Deposito.Endereco_Id" & vbCrLf
        Ds_PrecosParaCustos = Banco.ConsultaDataSet(Sql, "PrecosParaCustos")

        Dim CampoRoot As String() = HttpContext.Current.Server.MapPath("").Split("\")
        Dim RootSite As String = ""
        For j As Integer = 0 To CampoRoot.GetUpperBound(0) - 1
            RootSite &= CampoRoot(j) & "\"
        Next
        RootSite = RootSite & "Cr_PrecosParaCustos.rpt"

        Dim crpt As New ReportDocument()
        crpt.Load(RootSite)

        Dim NomeArquivo As String = Funcoes.GeraNomeArquivo & ".PDF"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

        Try
            crpt.SetDataSource(Ds_PrecosParaCustos)

            Dim crparametervalues As ParameterValues
            Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As ParameterFieldDefinitions
            Dim crparameterfielddefinition As ParameterFieldDefinition

            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

            crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssNomeEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

            mensagem = NomeArquivo

        Catch ex As Exception
            mensagem = ex.Message
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try

        Return mensagem
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PrecosParaCustos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class