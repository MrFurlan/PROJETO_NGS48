Imports System.Data
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucRegistrosIpiAjustaResumo
    Inherits BaseUserControl

    Dim Sql As String
    Dim Sqla As String

    Dim SqlArray As New ArrayList
    Dim Empresa() As String
    Dim Cliente() As String
    Dim SequenciaLote As Integer = 0
    Dim ds As New DataSet

    Dim Condicao As String = ""
    Dim Codigo As String = ""
    Dim Endereco As Integer = 0
    Dim Mensagem As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            txtDescricao.Attributes.Add("maxlength", "598")
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Sub BindGridView()
        CarregarDescricaoIcms()
        CarregarOrigemDocAjusteIPI()
        CarregarResumo()
        HttpContext.Current.Session("CodigoIpi") = 0
        HttpContext.Current.Session("DescricaoIpi") = ""
    End Sub

    Private Sub CarregarDescricaoIcms()
        Dim Descricao As String
        Sql = "SELECT * FROM DescricaoRAIPI WHERE Calculo = 'I'"
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = Format(Dr("Codigo_Id"), "000") & " - " & Dr("Descricao")
            DdlDescricaoRAIPI.Items.Add(New ListItem(Descricao, Dr("Codigo_ID")))
        Next
        DdlDescricaoRAIPI.Items.Insert(0, "")
        DdlDescricaoRAIPI.SelectedIndex = 0
    End Sub

    Private Sub CarregarDescricaoAjusteIPISpedFiscal()
        Dim Descricao As String
        DdlDescricaoAjusteIPISpedFiscal.Items.Clear()
        Sql = " SELECT     AjustesDaApuracaoIPI.Codigo_Id AS Codigo_Id, AjustesDaApuracaoIPI.Descricao AS Descricao "
        Sql &= " FROM         DescricaoRAIPI "
        Sql &= " INNER JOIN      AjustesDaApuracaoIPI ON DescricaoRAIPI.Codigo_Id = AjustesDaApuracaoIPI.DescricaoRAIPI_Id "
        Sql &= " WHERE     (DescricaoRAIPI.Calculo = 'I') and DescricaoRAIPI.codigo_id = " & DdlDescricaoRAIPI.SelectedValue & ""
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = Dr("Codigo_Id") & " - " & Dr("Descricao")
            DdlDescricaoAjusteIPISpedFiscal.Items.Add(New ListItem(Descricao, Dr("Codigo_ID")))
        Next
        DdlDescricaoAjusteIPISpedFiscal.Items.Insert(0, "")
        DdlDescricaoAjusteIPISpedFiscal.SelectedIndex = 0
    End Sub

    Private Sub CarregarOrigemDocAjusteIPI()
        Dim Descricao As String
        Sql = " SELECT * FROM OrigemDocumentoAjusteIPI "
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = Format(Dr("Codigo_Id"), "000") & " - " & Dr("Descricao")
            ddlOrgDocVinc.Items.Add(New ListItem(Descricao, Dr("Codigo_Id")))
        Next
        ddlOrgDocVinc.Items.Insert(0, "")
        ddlOrgDocVinc.SelectedIndex = 0
    End Sub

    Private Sub CarregarResumo()
        Sql = "Select Codigo_Id as Codigo, Descricao, Valor, AjustesApuracaoIPI_Id, OrigemDocumentoAjusteIpi_Id from ResumoItensRAIPI " & _
              " Where Empresa_ID = '" & HttpContext.Current.Session("EmpresaIpi") & "' And" & _
              " Processo_ID = " & HttpContext.Current.Session("ProcessoIpi")
        ds = Banco.ConsultaDataSet(Sql, "Resumo")
        GridResumo.DataSource = ds
        GridResumo.DataBind()
    End Sub

    Protected Sub GridResumo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        DdlDescricaoRAIPI.SelectedValue = GridResumo.SelectedRow.Cells(1).Text()
        CarregarDescricaoAjusteIPISpedFiscal()
        txtDescricao.Text = GridResumo.SelectedRow.Cells(2).Text()
        txtValor.Text = GridResumo.SelectedRow.Cells(3).Text()
        HttpContext.Current.Session("CodigoIpi") = GridResumo.SelectedRow.Cells(1).Text()
        HttpContext.Current.Session("DescricaoIpi") = GridResumo.SelectedRow.Cells(2).Text()
        DdlDescricaoAjusteIPISpedFiscal.SelectedIndex = GridResumo.SelectedRow.Cells(4).Text()
        Session("AjustesApuracaoIPI_Id") = GridResumo.SelectedRow.Cells(4).Text()
        ddlOrgDocVinc.SelectedValue = GridResumo.SelectedRow.Cells(5).Text()
        Session("OrigemDocumentoAjusteIpi_Id") = GridResumo.SelectedRow.Cells(5).Text()
    End Sub

    Private Overloads Sub Limpar()
        DdlDescricaoRAIPI.SelectedIndex = 0
        ddlOrgDocVinc.SelectedIndex = 0
        DdlDescricaoAjusteIPISpedFiscal.SelectedIndex = 0
        txtDescricao.Text = ""
        txtValor.Text = ""
        HttpContext.Current.Session("CodigoIpi") = 0
        HttpContext.Current.Session("DescricaoIpi") = ""
    End Sub

    Protected Sub DdlDescricaoRAIPI_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CarregarDescricaoAjusteIPISpedFiscal()
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            Sql = "Delete from ResumoItensRAIPI " & _
                          " Where Empresa_ID = '" & HttpContext.Current.Session("EmpresaIpi") & "' And" & _
                          " Processo_ID = " & HttpContext.Current.Session("ProcessoIpi") & " And " & _
                          " Codigo_Id = " & HttpContext.Current.Session("CodigoIpi") & " And " & _
                          " Descricao = '" & HttpContext.Current.Session("DescricaoIpi") & "' And " & _
                          " AjustesApuracaoIPI_Id = '" & HttpContext.Current.Session("AjustesApuracaoIPI_Id") & "' And " & _
                          " OrigemDocumentoAjusteIpi_Id = '" & HttpContext.Current.Session("OrigemDocumentoAjusteIpi_Id") & "'"
            SqlArray.Add(Sql)

            Sql = "INSERT INTO ResumoItensRAIPI " & _
                    "(Empresa_Id, Processo_Id, Codigo_Id, Descricao, Valor, AjustesApuracaoIPI_Id, OrigemDocumentoAjusteIpi_Id) " & _
                    "VALUES ('" & HttpContext.Current.Session("EmpresaIpi") & _
                    "', " & HttpContext.Current.Session("ProcessoIpi") & ", " & DdlDescricaoRAIPI.SelectedValue & _
                    ", '" & txtDescricao.Text & "', " & Replace(Replace(txtValor.Text, ".", ""), ",", ".") & _
                    ", " & DdlDescricaoAjusteIPISpedFiscal.SelectedValue & ", " & ddlOrgDocVinc.SelectedValue & ")"

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, "Erro de Gravação na Base de Dados...")
            Else
                Limpar()
                CarregarResumo()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            Sql = "Delete from ResumoItensRAIPI " & _
                                 " Where Empresa_ID = '" & HttpContext.Current.Session("EmpresaIpi") & "' And" & _
                                 " Processo_ID = " & HttpContext.Current.Session("ProcessoIpi") & " And " & _
                                 " Codigo_Id = " & HttpContext.Current.Session("CodigoIpi") & " And " & _
                                 " Descricao = '" & HttpContext.Current.Session("DescricaoIpi") & "' And " & _
                                 " AjustesApuracaoIPI_Id = '" & HttpContext.Current.Session("AjustesApuracaoIPI_Id") & "' And " & _
                                 " OrigemDocumentoAjusteIpi_Id = '" & HttpContext.Current.Session("OrigemDocumentoAjusteIpi_Id") & "'"

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, "Sucesso ao deletar.", eTitulo.Info)
            Else
                Limpar()
                CarregarResumo()
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

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Try
            Popup.CloseDialog(Me.Page, "divRegistrosIpiAjustaResumo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class