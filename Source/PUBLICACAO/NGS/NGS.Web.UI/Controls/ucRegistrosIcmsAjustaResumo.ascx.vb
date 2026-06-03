Imports System.Data
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucRegistrosIcmsAjustaResumo
    Inherits BaseUserControl

    Dim Sql As String

    Dim SqlArray As New ArrayList
    Dim ds As New DataSet

    Private ListaDeDescricoesCompletas As New Dictionary(Of String, Object)

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Try
                Limpar()
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Property ListaDeDescricoesCompletasSessao() As Dictionary(Of String, Object)
        Get
            ListaDeDescricoesCompletas = Session("ListaDeDescricoesCompletas")
            Return ListaDeDescricoesCompletas
        End Get
        Set(value As Dictionary(Of String, Object))
            Session("ListaDeDescricoesCompletas") = value
        End Set
    End Property

    Public Sub BindGridView()
        CarregarResumo()
        CarregarDescricaoIcms()
        CarregarDescricaoAjusteIcmsSpedFiscal()

        HttpContext.Current.Session("CodigoIcms") = 0
        HttpContext.Current.Session("DescricaoIcms") = ""
        HttpContext.Current.Session("AjustesApuracaoIcms_Id") = ""
    End Sub

    Private Sub CarregarDescricaoIcms()
        Dim Descricao As String
        Sql = "SELECT * FROM DescricaoRAICMS WHERE Calculo = 'I'"
        ddlCodigo.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = Format(Dr("Codigo_Id"), "000") & " - " & Dr("Descricao")
            ddlCodigo.Items.Add(New ListItem(Descricao, Dr("Codigo_ID")))
        Next
        ddlCodigo.Items.Insert(0, "")
        ddlCodigo.SelectedIndex = 0
    End Sub

    Private Sub CarregarDescricaoAjusteIcmsSpedFiscal()
        Dim Descricao As String
        ddlCodigoAjusteICMSSpedFiscal.Items.Clear()

        Sql = "SELECT A.Codigo_Id AS Codigo_Id, A.Descricao AS Descricao, isnull(A.DescricaoDetalhadaObservacoes, '') as DescricaoDetalhadaObservacoes " & vbCrLf & _
              "  FROM AjustesDaApuracaoIcms A" & vbCrLf & _
              "  LEFT JOIN DescricaoRAICMS Dra " & vbCrLf & _
              "    ON Dra.Codigo_Id = A.DescricaoRAICMS_Id      " & vbCrLf & _
              " INNER JOIN Clientes ON Clientes.Cliente_Id = '" & HttpContext.Current.Session("EmpresaIcms") & "'" & vbCrLf & _
              "   AND Clientes.Endereco_Id = " & HttpContext.Current.Session("EndEmpresaIcms") & vbCrLf & _
              "   AND Clientes.Estado = substring(A.Codigo_Id, 0, 3) " & vbCrLf & _
              " WHERE ISNULL(Dra.Calculo, 'I') = 'I'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = Dr("Codigo_Id") & " - " & Dr("Descricao")
            ddlCodigoAjusteICMSSpedFiscal.Items.Add(New ListItem(Descricao, Dr("Codigo_Id")))

            If Not ListaDeDescricoesCompletas.ContainsKey(Dr("Codigo_id")) Then
                ListaDeDescricoesCompletas.Add(Dr("Codigo_id"), Dr("DescricaoDetalhadaObservacoes"))
            End If

        Next

        ListaDeDescricoesCompletasSessao = ListaDeDescricoesCompletas

        ddlCodigoAjusteICMSSpedFiscal.Items.Insert(0, "")
        ddlCodigoAjusteICMSSpedFiscal.SelectedIndex = 0
    End Sub

    Private Sub CarregarResumo()
        Sql = "Select Codigo_Id as Codigo, Descricao, Valor,AjustesApuracaoIcms_Id from ResumoItensRAICMS " & _
              " Where Empresa_ID = '" & HttpContext.Current.Session("EmpresaIcms") & "' And" & _
              " Processo_ID = " & HttpContext.Current.Session("ProcessoIcms")

        ds = Banco.ConsultaDataSet(Sql, "Resumo")
        GridResumo.DataSource = ds
        GridResumo.DataBind()
    End Sub

    Protected Sub GridResumo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddlCodigo.SelectedValue = GridResumo.SelectedRow.Cells(1).Text()
            CarregarDescricaoAjusteIcmsSpedFiscal()
            txtDescricao.Text = Server.HtmlDecode(Textos.RetirarAcentos(GridResumo.SelectedRow.Cells(2).Text()))
            txtValor.Text = Server.HtmlDecode(GridResumo.SelectedRow.Cells(3).Text()).Trim
            Session("CodigoIcms") = GridResumo.SelectedRow.Cells(1).Text()
            Session("DescricaoIcms") = Textos.RetirarAcentos(GridResumo.SelectedRow.Cells(2).Text())
            ddlCodigoAjusteICMSSpedFiscal.SelectedValue = GridResumo.SelectedRow.Cells(4).Text()
            Session("AjustesApuracaoIcms_Id") = GridResumo.SelectedRow.Cells(4).Text()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Overloads Sub Limpar()
        ddlCodigo.SelectedIndex = 0
        txtDescricao.Text = String.Empty
        txtValor.Text = String.Empty
        HttpContext.Current.Session("CodigoIcms") = 0
        HttpContext.Current.Session("DescricaoIcms") = String.Empty
        ddlCodigoAjusteICMSSpedFiscal.Items.Clear()
        Session("AjustesApuracaoIcms_Id") = String.Empty
        Session.Remove("ListaDeDescricoesCompletas")
        txtDescAjusteICMS.Text = String.Empty
    End Sub

    Protected Sub ddlCodigo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCodigo.SelectedIndexChanged
        Try
            CarregarDescricaoAjusteIcmsSpedFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(ddlCodigo.SelectedValue) Then
            MsgBox(Me.Page, "Campo Código é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlCodigoAjusteICMSSpedFiscal.SelectedValue) Then
            MsgBox(Me.Page, "Campo Cod.Ajust.Icms é Obrigatório!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtValor.Text) Then
            MsgBox(Me.Page, "Campo Valor é Obrigatório!")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Validar() Then
                Sql = "Delete ResumoItensRAICMS " & vbCrLf & _
                      " Where Empresa_ID             = '" & HttpContext.Current.Session("EmpresaIcms") & "'" & vbCrLf & _
                      "   And Processo_ID            =  " & HttpContext.Current.Session("ProcessoIcms") & vbCrLf & _
                      "   And Codigo_Id              =  " & HttpContext.Current.Session("CodigoIcms") & vbCrLf & _
                      "   And Descricao              = '" & Server.HtmlDecode(Textos.RetirarAcentos(HttpContext.Current.Session("DescricaoIcms")).Trim) & "'" & vbCrLf & _
                      "   And AjustesApuracaoIcms_Id = '" & HttpContext.Current.Session("AjustesApuracaoIcms_Id") & "'"
                SqlArray.Add(Sql)

                Sql = "INSERT INTO ResumoItensRAICMS " & vbCrLf & _
                      "            (Empresa_Id, Processo_Id, Codigo_Id, Descricao, Valor,AjustesApuracaoIcms_Id) " & vbCrLf & _
                      "     VALUES ('" & HttpContext.Current.Session("EmpresaIcms") & "', " & vbCrLf & _
                      "              " & HttpContext.Current.Session("ProcessoIcms") & ", " & ddlCodigo.SelectedValue & ", " & vbCrLf & _
                      "             '" & Textos.RetirarAcentos((Server.HtmlDecode(txtDescricao.Text).Trim)) & "', " & Replace(Replace(txtValor.Text, ".", ""), ",", ".") & ", '" & ddlCodigoAjusteICMSSpedFiscal.SelectedValue & "')"

                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "Registro gravado com Sucesso.", eTitulo.Sucess)
                    Limpar()
                    CarregarResumo()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            Sql = "Delete ResumoItensRAICMS " & _
                  " Where Empresa_ID             = '" & HttpContext.Current.Session("EmpresaIcms") & "'" & vbCrLf & _
                  "   And Processo_ID            =  " & HttpContext.Current.Session("ProcessoIcms") & vbCrLf & _
                  "   And Codigo_Id              =  " & HttpContext.Current.Session("CodigoIcms") & vbCrLf & _
                  "   And Descricao              = '" & Server.HtmlDecode(Textos.RetirarAcentos(HttpContext.Current.Session("DescricaoIcms"))).Trim & "'" & vbCrLf & _
                  "   And AjustesApuracaoIcms_Id = '" & HttpContext.Current.Session("AjustesApuracaoIcms_Id") & "'"

            If Banco.GravaBanco(Sql) Then
                MsgBox(Me.Page, "Registro Excluído com Sucesso.", eTitulo.Sucess)
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
            Popup.CloseDialog(Me.Page, "divRegistrosIcmsAjustaResumo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCodigoAjusteICMSSpedFiscal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCodigoAjusteICMSSpedFiscal.SelectedIndexChanged
        Try
            txtDescAjusteICMS.Text = ListaDeDescricoesCompletasSessao.Item(ddlCodigoAjusteICMSSpedFiscal.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class