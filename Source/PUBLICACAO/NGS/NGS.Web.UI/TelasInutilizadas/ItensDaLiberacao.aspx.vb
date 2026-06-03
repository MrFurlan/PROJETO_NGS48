Imports System.Data
Imports System.IO
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ItensDaLiberacao
    Inherits BasePage

    Dim sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Compras)
        If Not IsPostBack AndAlso IsConnect Then
            If Funcoes.VerificaPermissao("ItensDaLiberacao", "ACESSAR") Then
                sql = " SELECT  Distinct   CotacoesDeComprasXItens.Ordem_Id as Ordem, Clientes.Cliente_Id as Cliente, Clientes.Endereco_Id as Endereco, Clientes.Nome, CotacoesDeComprasXItens.Checado, "
                sql &= "           CotacoesDeComprasXItens.Quantidade, CotacoesDeComprasXItens.Unidade, CotacoesDeComprasXItens.Descricao, "
                sql &= " CotacoesDeComprasXItens.Unitario, CotacoesDeComprasXItens.Total, CotacoesDeComprasXItens.IPI, CotacoesDeComprasXItens.Entrega"
                sql &= " FROM         CotacoesDeCompras INNER JOIN"
                sql &= "              CotacoesDeComprasXItens ON CotacoesDeCompras.Registro_ID = CotacoesDeComprasXItens.Registro_Id LEFT OUTER JOIN"
                sql &= "              Clientes ON CotacoesDeComprasXItens.Fornecedor_Id = Clientes.Cliente_Id AND "
                sql &= "   CotacoesDeComprasXItens.EndFornecedor_Id = Clientes.Endereco_Id"
                sql &= " WHERE     (CotacoesDeComprasXItens.Registro_Id = " & HttpContext.Current.Session("ssRegistro") & "and CotacoesDeComprasXItens.Unitario <> 0)"
                sql &= " Order by CotacoesDeComprasXItens.Ordem_Id"

                Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ConsultaItensLiberacao")

                Dim parameters = New Dictionary(Of String, Object)()

                Funcoes.BindReport(Me.Page, ds, "", eExportType.PDF, parameters)

                'Dim crpt As New ReportDocument()
                'crpt.FileName = Server.MapPath("~/Reports/.rpt")
                'crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)
                'crpt.SetDataSource(ds)

                'Dim crparametervalues As ParameterValues
                'Dim crparameterdiscretevalue As ParameterDiscreteValue
                'Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                'Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
                'crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                'crparameterfielddefinition = crparameterfielddefinitions.Item("Registro")
                'crparametervalues = crparameterfielddefinition.CurrentValues
                'crparameterdiscretevalue = New ParameterDiscreteValue
                'crparameterdiscretevalue.Value = HttpContext.Current.Session("ssRegistro")
                'crparametervalues.Add(crparameterdiscretevalue)
                'crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                'crparameterfielddefinition = crparameterfielddefinitions.Item("Empresa")
                'crparametervalues = crparameterfielddefinition.CurrentValues
                'crparameterdiscretevalue = New ParameterDiscreteValue
                'crparameterdiscretevalue.Value = HttpContext.Current.Session("ssNomeEmpresa")
                'crparametervalues.Add(crparameterdiscretevalue)
                'crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                'crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
                'crparametervalues = crparameterfielddefinition.CurrentValues
                'crparameterdiscretevalue = New ParameterDiscreteValue
                'crparameterdiscretevalue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
                'crparametervalues.Add(crparameterdiscretevalue)
                'crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                'Crv_Liberacoes.ReportSource = crpt
                '' Set the Crystal Report Viewer's property to the oRpt Report object that we 
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx", eTitulo.Info)
                Exit Sub
            End If
        End If
    End Sub

End Class