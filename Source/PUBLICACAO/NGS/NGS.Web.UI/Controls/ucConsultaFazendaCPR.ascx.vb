Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaFazendaCPR
    Inherits BaseUserControl

    Private strSQL As String = String.Empty
    Private objCPRxMatricula As CPRxMatricula
    Private ListCPRxMatriculas As ListCPRxMatricula
    Private CodigoDoCliente As String = ""
    Private objCPRxFaz As [Lib].Negocio.CPRxFazenda

#Region "CPR"
    Private ObjCpr As [Lib].Negocio.CPR

    Private Sub CarregarCPR()
        ObjCpr = CType(Session("objCpr"), [Lib].Negocio.CPR)
        If ObjCpr Is Nothing Then ObjCpr = New [Lib].Negocio.CPR
    End Sub

    Private Sub SalvarCPR()
        Session("objCpr") = ObjCpr
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Function Consultar() As Integer
        CarregarCPR()
        Dim itemFazenda As ListItem = Funcoes.FormatarListItemCliente(ObjCpr.Cliente)
        lblCliente.Text = itemFazenda.Value & "-" & ObjCpr.Cliente.Nome

        Dim sql As String
        sql = " SELECT CxM.Matricula_Id, CxM.Area, C.Cliente_id CodigoCliente, C.Endereco_Id EndCliente, C.Complemento Fazenda, " & vbCrLf & _
              "        (select CPRxM.area " & vbCrLf & _
              "           from CPRxMatricula CPRxM " & vbCrLf & _
              "          where CPRxM.cartorio_id = '" & ObjCpr.CodigoCartorio & "' " & vbCrLf & _
              "            and CPRxM.EndCartorio_id = " & ObjCpr.EndCartorio & " " & vbCrLf & _
              "            and CPRxM.CPR_id = '" & ObjCpr.CodigoCPR & "'" & vbCrLf & _
              "            and CPRxM.Matricula_id = CxM.Matricula_Id ) as AreaCPR" & vbCrLf & _
              "   FROM ClienteXMatricula CxM" & vbCrLf & _
              "   JOIN Clientes C" & vbCrLf & _
              "     ON CxM.Cliente_Id = C.Cliente_Id " & vbCrLf & _
              "    AND CxM.EndCliente_Id = C.Endereco_Id " & vbCrLf & _
              "  Where left(CxM.Cliente_Id,8)  = '" & Left(ObjCpr.CodigoCliente, 8) & "'" & vbCrLf & _
              "  union" & vbCrLf & _
              " SELECT CxA.Matricula_Id, CxA.Area, C.Cliente_id CodigoCliente, C.Endereco_Id EndCliente, C.Complemento Fazenda," & vbCrLf & _
              "        (select CPRxM.area " & vbCrLf & _
              "           from CPRxMatricula CPRxM " & vbCrLf & _
              "          where CPRxM.cartorio_id = '" & ObjCpr.CodigoCartorio & "' " & vbCrLf & _
              "            and CPRxM.EndCartorio_id = " & ObjCpr.EndCartorio & vbCrLf & _
              "            and CPRxM.CPR_id = '" & ObjCpr.CodigoCPR & "'" & vbCrLf & _
              "            and CPRxM.Matricula_id = CxA.Matricula_Id ) as AreaCPR " & vbCrLf & _
              "   FROM ClienteXArrendante CxA" & vbCrLf & _
              "   JOIN Clientes C" & vbCrLf & _
              "     ON CxA.Cliente_Id = C.Cliente_Id " & vbCrLf & _
              "    AND CxA.EndCliente_Id = C.Endereco_Id " & vbCrLf & _
              "  Where left(CxA.Cliente_Id,8)  ='" & Left(ObjCpr.CodigoCliente, 8) & "'" & vbCrLf & _
              "    and CxA.DataContrato_Id    <='" & ObjCpr.DataEmissao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "    and CxA.DataVencimento     >='" & ObjCpr.DataVencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "  ORDER BY Fazenda" & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Matricula")
        gridMatricula.DataSource = ds.Tables(0)
        gridMatricula.DataBind()
        upNewUpdatePanel.Update()
        Return ds.Tables(0).Rows.Count
    End Function

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
        upNewUpdatePanel.Update()
    End Sub

    Public Overrides Sub Limpar()
        Session.Remove("_MainUserControl")
        gridMatricula.DataSource = New List(Of Object)()
        gridMatricula.DataBind()
    End Sub

    Public Sub SetarTituloDIV(ByVal Titulo As String)
        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "$(document).ready(function () { $('#divConsultaFazendaCPR').prop('title', '" & Titulo & "'); });", True)
        upNewUpdatePanel.Update()
    End Sub

    Protected Sub btnConfirmar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Me.Selecionar()
    End Sub

    Protected Overrides Sub Selecionar()
        CarregarCPR()

        Dim area As Decimal = 0
        For i As Integer = 0 To gridMatricula.Rows.Count - 1
            Dim txt As TextBox = gridMatricula.Rows(i).FindControl("txtArea")
            If IIf(txt.Text.Length = 0, 0, CDec(txt.Text)) > CDec(gridMatricula.Rows(i).Cells(3).Text) Then
                MsgBox(Me.Page, "Área da Matrícula na CPR é maior que a área da mesma. Matrícula: " & gridMatricula.Rows(i).Cells(2).Text)
                Exit Sub
            End If
            area += IIf(txt.Text.Length = 0, 0, CDec(txt.Text))
        Next

        If area = 0 Then
            For Each row As [Lib].Negocio.CPRxFazenda In ObjCpr.Fazendas
                If row.CodigoFazenda = ObjCpr.CodigoCliente And row.EndFazenda = ObjCpr.EndCliente Then
                    If ObjCpr.IUD = "U" Then
                        row.IUD = "D"
                        If row.Salvar Then
                            ObjCpr.Fazendas.Remove(row)
                        Else
                            MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        End If
                    Else
                        ObjCpr.Fazendas.Remove(row)
                    End If
                    Exit For
                End If
            Next
            Exit Sub
        End If

        Dim EndCliente As Integer = -1
        Dim objFazenda As [Lib].Negocio.Cliente
        Dim cxf As [Lib].Negocio.CPRxFazenda
        Dim sqls As New ArrayList
        Dim ObjClienteCPR As [Lib].Negocio.CPRxCliente
        Dim ObjClienteCPRTemp As [Lib].Negocio.CPRxCliente

        For i As Integer = 0 To gridMatricula.Rows.Count - 1
            Dim txtarea As TextBox = gridMatricula.Rows(i).FindControl("txtArea")
            If gridMatricula.Rows(i).Cells(0).Text <> EndCliente And IIf(txtarea.Text.Length = 0, 0, CDec(txtarea.Text)) > 0 Then
                EndCliente = gridMatricula.Rows(i).Cells(0).Text

                '****************  INCLUI A FAZENDA NA CPR  ************************************************
                objFazenda = New [Lib].Negocio.Cliente(ObjCpr.CodigoCliente, gridMatricula.Rows(i).Cells(0).Text)
                objCPRxFaz = New [Lib].Negocio.CPRxFazenda(ObjCpr)
                objCPRxFaz.CodigoFazenda = objFazenda.Codigo
                objCPRxFaz.EndFazenda = objFazenda.CodigoEndereco
                objCPRxFaz.Fazenda = objFazenda

                'ObjCpr.Fazendas.Add(objCPRxFaz)

                ''************** INCLUI OS SOCIOS NA FAZENDA ************************************************
                ObjClienteCPRTemp = Nothing

                ObjClienteCPR = New [Lib].Negocio.CPRxCliente(objCPRxFaz)
                ObjClienteCPR.CodigoCliente = objFazenda.Codigo
                ObjClienteCPR.EndCliente = objFazenda.CodigoEndereco
                CodigoDoCliente = ObjClienteCPR.CodigoCliente
                ObjClienteCPRTemp = objCPRxFaz.Clientes.Find(AddressOf ClientePorCodigo)
                If ObjClienteCPRTemp Is Nothing Then
                    objCPRxFaz.Clientes.Add(ObjClienteCPR)
                End If

                If objFazenda.Socios.Count > 0 Then
                    For Each Socio As [Lib].Negocio.ClienteXSocio In objFazenda.Socios
                        ObjClienteCPR = New [Lib].Negocio.CPRxCliente(objCPRxFaz)
                        ObjClienteCPR.CodigoCliente = Socio.CodigoSocio
                        ObjClienteCPR.EndCliente = Socio.EndSocio

                        CodigoDoCliente = ObjClienteCPR.CodigoCliente
                        ObjClienteCPRTemp = objCPRxFaz.Clientes.Find(AddressOf ClientePorCodigo)
                        If ObjClienteCPRTemp Is Nothing Then
                            objCPRxFaz.Clientes.Add(ObjClienteCPR)
                        End If
                    Next
                End If

                '************** INCLUI AS MATRICULAS NA FAZENDA ********************************************
                objCPRxFaz.Matriculas.Clear()
                For j As Integer = 0 To gridMatricula.Rows.Count - 1
                    Dim txt As TextBox = gridMatricula.Rows(j).FindControl("txtArea")
                    If IIf(txt.Text.Length = 0, 0, CDec(txt.Text)) > 0 And gridMatricula.Rows(j).Cells(0).Text = EndCliente Then
                        Dim Matricula As New [Lib].Negocio.CPRxMatricula(objCPRxFaz)
                        Matricula.CodigoMatricula = gridMatricula.Rows(j).Cells(2).Text
                        Matricula.Area = CDec(txt.Text)
                        objCPRxFaz.Matriculas.Add(Matricula)
                    End If
                    area += IIf(txt.Text.Length = 0, 0, CDec(txt.Text))
                Next

                '************** REMOVE A(s) FAZENDA(s) CASO TENHA(M) SIDO INSERIDA(S) ANTERIORMENTE ******
                cxf = Nothing
                sqls.Clear()
                For Each row As [Lib].Negocio.CPRxFazenda In ObjCpr.Fazendas
                    If row.CodigoFazenda = ObjCpr.CodigoCliente And row.EndFazenda = EndCliente Then
                        If ObjCpr.IUD = "U" Then
                            row.IUD = "D"
                            row.SalvarSql(sqls)
                            cxf = row
                        Else
                            ObjCpr.Fazendas.Remove(row)
                        End If
                        Exit For
                    End If
                Next

                '************** Insere no BD Fazendas, Proprietarios, Matriculas ******
                If ObjCpr.IUD = "U" Then
                    objCPRxFaz.IUD = "I"
                    objCPRxFaz.SalvarSql(sqls)
                    objCPRxFaz.IUD = ""
                    If Banco.GravaBanco(sqls) Then
                        ObjCpr.Fazendas.Remove(cxf)
                        ObjCpr.Fazendas.Add(objCPRxFaz)
                    Else
                        MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End If
                Else
                    ObjCpr.Fazendas.Add(objCPRxFaz)
                End If
            End If
        Next
        SalvarCPR()
        Session("CPR") = True

        Try
            Session(Session("ssTipoRetorno")) = ObjCpr
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(ObjCpr)
                Else
                    CType(Me.Page, IBasePage).Carregar(ObjCpr)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaFazendaCPR")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Problemas no banco de dados.", eTitulo.Erro)
        End Try
    End Sub

    Private Function ClientePorCodigo(ByVal objClienteCPR As CPRxCliente) As Boolean
        If objClienteCPR.CodigoCliente = CodigoDoCliente Then
            Return True
            Exit Function
        End If
        Return False
    End Function

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaFazendaCPR")
    End Sub

End Class