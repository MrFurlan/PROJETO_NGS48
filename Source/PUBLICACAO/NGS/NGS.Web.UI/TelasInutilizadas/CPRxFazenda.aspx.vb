Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CPRxFazenda
    Inherits BasePage

    Private objCpr As [Lib].Negocio.CPR

    Private Sub CarregarCPR()
        objCpr = CType(Session("objCpr"), [Lib].Negocio.CPR)
        If objCpr Is Nothing Then objCpr = New [Lib].Negocio.CPR
    End Sub

    Private Sub SalvarCPR()
        Session("objCpr") = objCpr
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack AndAlso IsConnect Then
            HID.Value = Guid.NewGuid().ToString
            ucConsultaClientes.SetarHID(HID.Value)
        End If
    End Sub

    Public Sub CarregarCliente()
        If Not Session("ObjClienteCPRxFAZ" & HID.Value) Is Nothing Then
            Dim objFazenda As [Lib].Negocio.Cliente = CType(Session("ObjClienteCPRxFAZ" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemFazenda As ListItem = Funcoes.FormatarListItemCliente(objFazenda)
            txtFazenda.Text = itemFazenda.Text
            txtCodigoFazenda.Value = itemFazenda.Value & "-" & CType(Session("ObjClienteCPRxFAZ" & HID.Value), [Lib].Negocio.Cliente).Nome
            CarregarCPR()
            Dim sql As String
            sql = " SELECT CxM.Matricula_Id, CxM.Area," & vbCrLf & _
                  "        (select CPRxM.area from CPRxMatricula CPRxM where CPRxM.cartorio_id = '" & objCpr.CodigoCartorio & "' and CPRxM.EndCartorio_id = " & objCpr.EndCartorio & " and CPRxM.CPR_id = " & objCpr.CodigoCPR & " and CPRxM.Matricula_id = CxM.Matricula_Id ) as AreaCPR" & vbCrLf & _
                  "   FROM ClienteXMatricula CxM" & vbCrLf & _
                  "  Where CxM.Cliente_Id    = '" & objFazenda.Codigo & "'" & vbCrLf & _
                  "    and CxM.EndCliente_Id =  " & objFazenda.CodigoEndereco & vbCrLf & _
                  "  union" & vbCrLf & _
                  " SELECT CxA.Matricula_Id, CxA.Area," & vbCrLf & _
                  "        (select CPRxM.area from CPRxMatricula CPRxM where CPRxM.cartorio_id = '" & objCpr.CodigoCartorio & "' and CPRxM.EndCartorio_id = " & objCpr.EndCartorio & " and CPRxM.CPR_id = " & objCpr.CodigoCPR & " and CPRxM.Matricula_id = CxA.Matricula_Id ) as AreaCPR " & vbCrLf & _
                  "   FROM ClienteXArrendante CxA" & vbCrLf & _
                  "  Where CxA.Cliente_Id       = '" & objFazenda.Codigo & "'" & vbCrLf & _
                  "    and CxA.EndCliente_Id    =  " & objFazenda.CodigoEndereco & vbCrLf & _
                  "    and CxA.DataContrato_Id <= '" & objCpr.DataEmissao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                  "    and CxA.DataVencimento  >= '" & objCpr.DataVencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Matricula")
            gridMatricula.DataSource = ds.Tables(0)
            gridMatricula.DataBind()
            Session.Remove("ObjClienteCPRxFAZ" & HID.Value)
        End If
    End Sub

    Protected Sub btnConfirmar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        CarregarCPR()

        Dim area As Decimal = 0
        For i As Integer = 0 To gridMatricula.Rows.Count - 1
            Dim txt As TextBox = gridMatricula.Rows(i).FindControl("txtArea")
            If IIf(txt.Text.Length = 0, 0, CDec(txt.Text)) > CDec(gridMatricula.Rows(i).Cells(1).Text) Then
                MsgBox(Me.Page, "Area da matricula na CPR é maior que a area da mesma. matricula: " & gridMatricula.Rows(i).Cells(0).Text)
                Exit Sub
            End If
            area += IIf(txt.Text.Length = 0, 0, CDec(txt.Text))
        Next

        If area = 0 Then
            For Each row As [Lib].Negocio.CPRxFazenda In objCpr.Fazendas
                If row.CodigoFazenda = txtCodigoFazenda.Value.Split("-")(0) And row.EndFazenda = txtCodigoFazenda.Value.Split("-")(1) Then
                    If objCpr.IUD = "U" Then
                        row.IUD = "D"
                        If row.Salvar Then
                            objCpr.Fazendas.Remove(row)
                        Else
                            MsgBox(Me.Page, "Erro ao alterar as matriculas da CPR.")
                        End If
                    Else
                        objCpr.Fazendas.Remove(row)
                    End If
                    Exit For
                End If
            Next

            Exit Sub
        End If

        '****************  INCLUI A FAZENDA NA CPR  ************************************************
        Dim Fazenda As New [Lib].Negocio.Cliente(txtCodigoFazenda.Value.Split("-")(0), txtCodigoFazenda.Value.Split("-")(1))
        Dim CPRxFaz As New [Lib].Negocio.CPRxFazenda(objCpr)
        CPRxFaz.CodigoFazenda = Fazenda.Codigo
        CPRxFaz.EndFazenda = Fazenda.CodigoEndereco
        CPRxFaz.Fazenda = Fazenda

        '************** INCLUI OS SOCIOS NA FAZENDA ************************************************
        If Fazenda.Socios.Count = 0 Then
            Dim ClienteCPR As New [Lib].Negocio.CPRxCliente(CPRxFaz)
            ClienteCPR.CodigoCliente = Fazenda.Codigo
            ClienteCPR.EndCliente = Fazenda.CodigoEndereco
            CPRxFaz.Clientes.Add(ClienteCPR)
        Else
            For Each Socio As [Lib].Negocio.ClienteXSocio In Fazenda.Socios
                Dim ClienteCPR As New [Lib].Negocio.CPRxCliente(CPRxFaz)
                ClienteCPR.CodigoCliente = Socio.CodigoSocio
                ClienteCPR.EndCliente = Socio.EndSocio
                CPRxFaz.Clientes.Add(ClienteCPR)
            Next
        End If

        '************** INCLUI AS MATRICULAS NA FAZENDA ********************************************
        For i As Integer = 0 To gridMatricula.Rows.Count - 1
            Dim txt As TextBox = gridMatricula.Rows(i).FindControl("txtArea")
            If IIf(txt.Text.Length = 0, 0, CDec(txt.Text)) > 0 Then
                Dim Matricula As New [Lib].Negocio.CPRxMatricula(CPRxFaz)
                Matricula.CodigoMatricula = gridMatricula.Rows(i).Cells(0).Text
                Matricula.Area = CDec(txt.Text)
                CPRxFaz.Matriculas.Add(Matricula)
            End If
            area += IIf(txt.Text.Length = 0, 0, CDec(txt.Text))
        Next

        Dim cxf As [Lib].Negocio.CPRxFazenda = Nothing
        Dim sqls As New ArrayList
        For Each row As [Lib].Negocio.CPRxFazenda In objCpr.Fazendas
            If row.CodigoFazenda = txtCodigoFazenda.Value.Split("-")(0) And row.EndFazenda = txtCodigoFazenda.Value.Split("-")(1) Then
                If objCpr.IUD = "U" Then
                    row.IUD = "D"
                    row.SalvarSql(sqls)
                    cxf = row
                Else
                    objCpr.Fazendas.Remove(row)
                End If
                Exit For
            End If
        Next

        If objCpr.IUD = "U" Then
            CPRxFaz.IUD = "I"
            CPRxFaz.SalvarSql(sqls)
            CPRxFaz.IUD = ""
            If Banco.GravaBanco(sqls) Then
                objCpr.Fazendas.Remove(cxf)
                objCpr.Fazendas.Add(CPRxFaz)
            Else
                MsgBox(Me.Page, "Erro ao alterar as matriculas da CPR.")
            End If
        Else
            objCpr.Fazendas.Add(CPRxFaz)
        End If


        SalvarCPR()
        Session("CPR") = True
        Dim strJS As String = "window.opener.document.forms[0].submit();" & _
                              "window.close();"
        ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), strJS, True)
    End Sub

    Protected Sub btnSair_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim strJS As String = "window.opener.document.forms[0].submit();" & _
                      "window.close();"
        ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), strJS, True)
    End Sub

    Protected Sub btnFazenda_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo") = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCPRxFAZ" & HID.Value, "txtNome")
    End Sub

End Class