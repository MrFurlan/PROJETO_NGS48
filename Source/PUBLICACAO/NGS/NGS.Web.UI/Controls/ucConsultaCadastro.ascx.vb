Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaCadastro
    Inherits BaseUserControl

    Public Property Origem() As eTipoDeDocumento
        Get
            Return ViewState("ObjOrigem" + HID.Value)
        End Get
        Set(ByVal value As eTipoDeDocumento)
            ViewState("ObjOrigem" + HID.Value) = value
        End Set
    End Property

    Public Property objConhecimento() As [Lib].Negocio.NotaFiscal
        Get
            Return CType(Session("ssConhecimentoDeTransporte" & HID.Value), NotaFiscal)
        End Get
        Set(ByVal value As [Lib].Negocio.NotaFiscal)
            Session("ssConhecimentoDeTransporte" & HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack Then
                ddl.Carregar(ddlUfConsulta, CarregarDDL.Tabela.EstadosUF)
            End If
        Catch ex As Exception
            Dim sLine As String = ex.StackTrace.Substring(ex.StackTrace.IndexOf("linha"))
            sLine = sLine.Substring(0, sLine.IndexOf("em"))

            MsgBox(Me.Page, ex.StackTrace)
            MsgBox(Me.Page, ex.TargetSite.ReflectedType.Name + " + " + ex.TargetSite.Name + " + " + sLine)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If txtCpfCnpj.Text.Trim.Length <> 11 AndAlso txtCpfCnpj.Text.Trim.Length <> 14 Then
                MsgBox(Me.Page, "CPF ou CNPJ Inválido.")
            ElseIf String.IsNullOrWhiteSpace(ddlUfConsulta.SelectedValue) Then
                MsgBox(Me.Page, "Informe o estado onde foi registrado o " & IIf(txtCpfCnpj.Text.Trim.Length = 11, "CPF", "CNPJ") & ".")
            ElseIf txtCpfCnpj.Text.Trim.Length = 11 AndAlso Not Funcoes.ValidaCPF(txtCpfCnpj.Text.Trim) Then
                MsgBox(Me.Page, "CPF Inválido.")
            ElseIf txtCpfCnpj.Text.Trim.Length = 14 AndAlso Not Funcoes.ValidaCNPJ(txtCpfCnpj.Text.Trim) Then
                MsgBox(Me.Page, "CNPJ inválido.")
            Else
                Consultar()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.StackTrace.Substring(ex.StackTrace.IndexOf("line")))
        End Try
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Try
            Popup.CloseDialog(Me.Page, "divConsultaCadastro")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtCpfCnpj.Text = String.Empty
        ddlUfConsulta.SelectedValue = String.Empty
        gridCadastro.DataBind()
    End Sub

    Private Sub Consultar()
        Try
            If Origem = eTipoDeDocumento.CTRC OrElse Origem = eTipoDeDocumento.CT_E Then
                Dim Sqls As New ArrayList
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("cadastrocte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                obj.Texto = getTextoCadastro()
                obj.SalvarSql(Sqls)

                If Banco.GravaBanco(Sqls) Then
                    'AGUARDAR RESPOSTA SEFAZ
                    Dim resp As [Lib].Negocio.Resp = Nothing
                    Dim fileName As String = String.Format("resp-cadastrocte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)

                    While resp Is Nothing
                        resp = GetResp(fileName)
                    End While

                    If resp IsNot Nothing Then MsgBox(Me.Page, resp.Texto.Trim(), eTitulo.Info)
                End If

            ElseIf Origem = eTipoDeDocumento.Nota Then

                Dim MsgDeErro As String = String.Empty


                Dim lstMsg As String() = DocumentoEletronico.ConsultaCadastro(txtCpfCnpj.Text, ddlUfConsulta.SelectedValue, MsgDeErro)

                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Dim Codigo As String = String.Empty
                If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                    Codigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strMsg As String = String.Empty
                If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                    strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                If Not String.IsNullOrWhiteSpace(Codigo) AndAlso Codigo <> "111" Then
                    MsgBox(Me.Page, String.Format("{0} - {1}", Codigo, strMsg))
                    Exit Sub
                Else
                    MsgBox(Me.Page, strMsg)
                End If

                Dim ds As DataSet = geraDataSet(lstMsg)

                gridCadastro.DataSource = ds
                gridCadastro.DataBind()

                '********* Captura dos Dados ********************
                'Dim Ie As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("IE")).FirstOrDefault()
                'Dim Situacao As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("SITUACAO")).FirstOrDefault()
                'Dim SituacaoNfe As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("SITUACAONFE")).FirstOrDefault()
                'Dim SituacaoCte As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("SITUACAOCTE")).FirstOrDefault()
                'Dim Nome As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NOME")).FirstOrDefault()
                'Dim Fantasia As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("FANTASIA")).FirstOrDefault()
                'Dim Logradouro As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("LOGRADOURO")).FirstOrDefault()
                'Dim Complemento As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("COMPLEMENTO")).FirstOrDefault()
                'Dim Numero As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NUMERO")).FirstOrDefault()
                'Dim Bairro As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("BAIRRO")).FirstOrDefault()
                'Dim Municipio As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MUNICIPIO")).FirstOrDefault()
                'Dim Cep As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("CEP")).FirstOrDefault()
                'Dim InicioAtividade As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("INICIOATIVIDADE")).FirstOrDefault()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function geraDataSet(lst As String()) As DataSet
        Dim ds As New DataSet
        ds.Tables.Add("Result")

        Dim index As Integer = 0

        For Each pos As String In lst
            If pos.Trim.Equals(String.Empty) Then
                index += 1
            Else
                If index = 0 Then ds.Tables(0).Columns.Add(pos.Split("=")(0).Trim)
                If ds.Tables(0).Rows.Count = index Then ds.Tables(0).Rows.Add()

                ds.Tables(0).Rows(index)(pos.Split("=")(0).Trim) = pos.Split("=")(1).Trim
            End If


            'ds.Tables(0).Rows(pos.Split("=")(0).ToString) = ""
        Next
        Return ds
    End Function

    Private Function getTextoCadastro() As String
        Dim sb As New StringBuilder()
        sb.Append(IIf(txtCpfCnpj.Text.Trim.Length = 14, "CNPJ=", "CPF=") & txtCpfCnpj.Text.OnlyNumbers().Trim() & ControlChars.CrLf)
        sb.Append("UF=MT" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function
End Class