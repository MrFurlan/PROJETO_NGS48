Imports NGS.Lib.Negocio

Public Class RecompraDeTitulos
    Inherits BasePage
    Dim consulta As Boolean = False

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Not Funcoes.VerificaPermissao("RecompraDeTitulos", "ACESSAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If ValidaDados() Then
                If PreencheTitulosReceber() Then
                    If Not consulta Then
                        lnkConsultar.Parent.Visible = False
                        lnkExcluir.Parent.Visible = False
                        lnkNovo.Parent.Visible = True
                    Else
                        lnkExcluir.Parent.Visible = True
                    End If
                    txtTitulo.Enabled = False
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("RecompraDeTitulos", "GRAVAR") Then
                Dim sqls As New ArrayList
                Dim sql As String

                If grdTitulos.Rows.Count = 0 Then
                    sql = "Update ContasAPagar " & vbCrLf & _
                            "       set ContratoBancario = '" & txtContrato.Text & "'" & vbCrLf & _
                            "   Where Registro_Id = '" & txtTitulo.Text & "'"

                    sqls.Add(sql)
                Else
                    If ValidaValores() Then
                        For Each row As GridViewRow In grdTitulos.Rows
                            Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                            If chk.Checked Then
                                sql = " insert into TituloXRecompraDeDuplicatas " & vbCrLf & _
                                      "		    (TituloRecompra, TituloRelacionado, Valor, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                      "     values('" & txtTitulo.Text & "', '" & row.Cells(1).Text & "', " & Str(CType(row.FindControl("txtValor"), TextBox).Text) & ", " & vbCrLf & _
                                                    "'" & UsuarioServidor.NomeUsuario & "', '" & DateTime.Now.ToString("yyyy/MM/dd") & "')" & vbCrLf
                                sqls.Add(sql)
                            End If
                        Next
                    Else
                        For Each row As GridViewRow In grdTitulos.Rows
                            Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                            If chk.Checked Then
                                Dim txt As TextBox = CType(row.FindControl("txtValor"), TextBox)
                                txt.Enabled = True
                            End If
                        Next
                    End If
                End If

                If sqls.Count > 0 Then
                    If Banco.GravaBanco(sqls) Then
                        MsgBox(Me.Page, "Títulos Atualizados com Sucesso.", eTitulo.Sucess)
                        LimparCampos()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissoes para Gravar.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            Dim sql As String = "Delete TituloXRecompraDeDuplicatas where TituloRecompra = '" & txtTitulo.Text & "'"
            Banco.GravaBanco(sql)
            MsgBox(Me.Page, "sucesso na exclusão.", eTitulo.Sucess)
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RecompraDeTitulos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Function ValidaValores() As Boolean

        For Each row As GridViewRow In grdTitulos.Rows
            Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)

            If chk.Checked Then
                Dim txtvlrdoc As TextBox = CType(row.FindControl("txtValorDocumento"), TextBox)
                Dim txtvlrrec As TextBox = CType(row.FindControl("txtValorRecomprado"), TextBox)
                Dim txt As TextBox = CType(row.FindControl("txtValor"), TextBox)
                Dim saldo As Decimal = CDec(txtvlrdoc.Text) - CDec(txtvlrrec.Text)

                If String.IsNullOrWhiteSpace(txt.Text) OrElse CDec(txt.Text) = 0 Then
                    MsgBox(Me.Page, "Todos os valores de recompra devem ser informados e não podem ser 0.00.")
                    Return False
                ElseIf CDec(txt.Text) > saldo Then
                    MsgBox(Me.Page, "Atenção o valor informado é maior do que o valor de saldo.")
                    Return False
                End If
            End If
        Next
        If CDec(txtTotal.Text.Replace(".", ",")) <> CDec(txtValorTP.Text) Then
            MsgBox(Me.Page, "A soma da recompra tem que coincidir com o valor do título de recompra.")
            Return False
        End If
        Return True
    End Function

    Private Function getContratoBancario() As String
        Dim contrato As String = ""

        Dim sql As String = "select ContratoBancario, ValorDoDocumento from ContasAPagar where ContratoBancario is not null and ContratoBancario <> '' and Registro_Id = '" & txtTitulo.Text & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "TituloPagar")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            txtContrato.Text = ds.Tables(0).Rows(0)(0)
            txtValorTP.Text = ds.Tables(0).Rows(0)(1)
            Return txtContrato.Text
        End If

        txtContrato.Text = ""
        txtValorTP.Text = ""

        Return ""
    End Function

    Private Function ExisteTituloPagar() As Boolean
        Dim sql = "select Registro_Id from ContasAPagar where Registro_Id = '" & txtTitulo.Text & "'"
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ContasAPagar")

        If ds IsNot Nothing AndAlso ds.Tables.Count AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            MsgBox(Me.Page, "Tiulo não encontrado.")
            Return False
        End If
    End Function

    Private Function ValidaDados() As Boolean
        If String.IsNullOrWhiteSpace(txtTitulo.Text) Then
            MsgBox(Me.Page, "Informe o Titulo")
            Return False
        ElseIf String.IsNullOrWhiteSpace(getContratoBancario()) AndAlso ExisteTituloPagar() Then
            MsgBox(Me.Page, "Título a Pagar não contém contratos, informe o Contrato, Grave o Titulo e faça a Consulta Novamente")
            txtContrato.Enabled = True
            lnkConsultar.Parent.Visible = False
            lnkNovo.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtTitulo.Enabled = False
            Return False
        End If
        Return True
    End Function

    Private Function PreencheTitulosReceber() As Boolean
        Dim sql As String = "Select tr.TituloRelacionado as TituloRecomprar,                                                                       " & vbCrLf & _
                            "	   cr.Empresa + '-' + emp.Nome as Empresa, cr.Cliente + '-' + cli.Nome as Cliente,                                 " & vbCrLf & _
                            "cr.ValorDoDocumento, SUM(tr.Valor) as ValorRecomprado, cr.ValorDoDocumento - ISNULL(sum(tr.valor), 0) as SaldoRecompra" & vbCrLf & _
                            "	From TituloXRecompraDeDuplicatas tr                                                                                " & vbCrLf & _
                            "	Inner Join ContasAReceber cr	                                                                                   " & vbCrLf & _
                            "		On cr.Registro_Id = tr.TituloRelacionado                                                                       " & vbCrLf & _
                            "	Inner Join Clientes cli                                                                                            " & vbCrLf & _
                            "		ON cli.Cliente_Id = cr.Cliente                                                                                 " & vbCrLf & _
                            "		And cli.Endereco_Id = cr.EndCliente                                                                            " & vbCrLf & _
                            "	Inner JOin Clientes as Emp                                                                                         " & vbCrLf & _
                            "       	 On Emp.Cliente_Id  = cr.Empresa                                                                           " & vbCrLf & _
                            "       	and emp.Endereco_Id = cr.EndEmpresa                                                                        " & vbCrLf & _
                            "	where tr.TituloRecompra = '" & txtTitulo.Text & "'                                                                 " & vbCrLf & _
                            "group by tr.TituloRelacionado, cr.Empresa + '-' + emp.Nome, cr.Cliente + '-' + cli.Nome, cr.ValorDoDocumento          " & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Recompra")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            grdTitulos.DataSource = ds
            grdTitulos.DataBind()

            For Each row As GridViewRow In grdTitulos.Rows
                Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                chk.Enabled = False
            Next
            txtTotal.Text = ds.Tables(0).Compute("Sum(ValorRecomprado)", "")
            consulta = True
        Else
            sql = "select cr.Registro_Id as TituloRecomprar, isnull(tr.TituloRelacionado, 0) as TituloRelacionado," & vbCrLf & _
                  "       cr.Empresa + '-' + emp.Nome as Empresa, cr.Cliente + '-' + cli.Nome as Cliente, " & vbCrLf & _
                  "	    cr.ValorDoDocumento, isnull(tr.valor,0) as ValorRecomprado, cr.ValorDoDocumento - ISNULL(tr.valor, 0) as SaldoRecompra " & vbCrLf & _
                  "	from ContasAReceber cr                                                                                                     " & vbCrLf & _
                  "	    Inner JOin Clientes as Emp                                                                                             " & vbCrLf & _
                  "       	 On Emp.Cliente_Id  = cr.Empresa                                                                                   " & vbCrLf & _
                  "       	and emp.Endereco_Id = cr.EndEmpresa                                                                                " & vbCrLf & _
                  "       Inner Join Clientes as Cli                                                                                        " & vbCrLf & _
                  "       	 On Cli.Cliente_Id  = cr.Cliente                                                                                   " & vbCrLf & _
                  "       	and Cli.Endereco_Id = cr.EndCliente                                                                                " & vbCrLf & _
                  "		Left Join ( Select TituloRelacionado,  sum(Valor) as Valor                                             " & vbCrLf & _
                  "				   		From TituloXRecompraDeDuplicatas                                                                       " & vbCrLf & _
                  "				   	 group by TituloRelacionado                                                                " & vbCrLf & _
                  "				  ) as tr                                                                                                      " & vbCrLf & _
                  "			ON tr.TituloRelacionado = cr.Registro_Id                                                                           " & vbCrLf & _
                  "	Where cr.ContratoBancario = '" & txtContrato.Text & "'" & vbCrLf

            ds = Banco.ConsultaDataSet(sql, "TituloReceber")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                grdTitulos.DataSource = ds
                grdTitulos.DataBind()
            Else
                MsgBox(Me.Page, "Nenhum resultado encontrado")
            End If
        End If

        Return True

    End Function

    Private Sub LimparCampos()
        lnkConsultar.Parent.Visible = True
        lnkNovo.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtTitulo.Enabled = True
        txtContrato.Enabled = False
        txtContrato.Text = String.Empty
        txtValorTP.Text = String.Empty
        txtTitulo.Text = String.Empty
        txtTotal.Text = "0.00"
        grdTitulos.DataBind()
        txtTitulo.Focus()
    End Sub

#End Region
   
End Class