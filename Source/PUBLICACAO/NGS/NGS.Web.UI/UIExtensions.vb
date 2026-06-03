Imports System
Imports System.IO
Imports System.Linq
Imports System.Web
Imports System.Xml
Imports System.Data
Imports System.Collections
Imports System.Collections.Generic
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Runtime.Serialization

Module UIExtensions

    Public Function Concat(Of t)(ByVal lst As IEnumerable(Of t), ByVal valueFunction As Func(Of t, Object)) As String
        Return Concat(lst, valueFunction, ", ", "", "")
    End Function

    Public Function Concat(Of t)(ByVal lst As IEnumerable(Of t), ByVal valueFunction As Func(Of t, Object), ByVal separator As String) As String
        Return Concat(lst, valueFunction, ", ", "", "")
    End Function

    Public Function Concat(Of t)(ByVal lst As IEnumerable(Of t), ByVal valueFunction As Func(Of t, Object), ByVal separator As String, ByVal format As String) As String
        Return Concat(lst, valueFunction, ", ", format, "")
    End Function

    Public Function Concat(Of t)(ByVal lst As IEnumerable(Of t), ByVal valueFunction As Func(Of t, Object), ByVal separator As String, ByVal format As String, ByVal defaultValue As String) As String
        If Not (lst IsNot Nothing AndAlso lst.Count() > 0) Then
            Return defaultValue
        End If

        Dim str As New StringBuilder()
        For Each obj As t In lst
            Dim value As Object = valueFunction(obj)
            Dim valuestr As String = Convert.ToString(value)

            If String.IsNullOrWhiteSpace(valuestr) Then
                Continue For
            End If

            If String.IsNullOrWhiteSpace(format) Then
                str.Append(valuestr & separator)
            Else
                str.Append(String.Format("{0:" & format & " }", value) & separator)
            End If
        Next

        Return IIf(String.IsNullOrWhiteSpace(str.ToString()), defaultValue, str.Remove(str.Length - separator.Length, separator.Length).ToString())
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IndexOfAny(ByVal str As [String], ByVal anyOf As [String]()) As Int32
        Return IndexOfAny(str, anyOf, 0)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IndexOfAny(ByVal str As [String], ByVal anyOf As [String](), ByVal startIndex As Int32) As Int32
        If str Is Nothing OrElse [String].IsNullOrEmpty(str) Then
            Return -1
        End If

        Dim minIndex As Int32 = -1
        For Each s As [String] In anyOf
            Dim newIndex As Int32 = str.IndexOf(s, startIndex)
            If minIndex = -1 OrElse newIndex < minIndex AndAlso newIndex <> -1 Then
                minIndex = newIndex
            End If
        Next
        Return minIndex
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToNullableInt32(ByVal obj As [Object]) As System.Nullable(Of Int32)
        If obj Is Nothing Then
            Return Nothing
        End If
        Return Convert.ToInt32(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToInt32(ByVal obj As [Object]) As Int32
        Return Convert.ToInt32(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToNullableDecimal(ByVal obj As [Object]) As System.Nullable(Of [Decimal])
        If obj Is Nothing Then
            Return Nothing
        End If
        If obj.ToString() = "" Then
            Return Nothing
        End If
        Return Convert.ToDecimal(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDecimal(ByVal obj As [Object]) As [Decimal]
        Return (If([String].IsNullOrEmpty(obj.ToString()), 0D, Convert.ToDecimal(obj)))
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToNullableDouble(ByVal obj As [Object]) As System.Nullable(Of [Double])
        If obj Is Nothing Then
            Return Nothing
        End If
        Return Convert.ToDouble(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDouble(ByVal obj As [Object]) As [Double]
        Return Convert.ToDouble(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function OnlyChars(ByVal str As [String], ByVal charListToKeep As [String]) As [String]
        Dim newStr = New StringBuilder()
        str.ToCharArray().Where(Function(c) charListToKeep.Any(Function(cc) cc = c)).ToList().ForEach(Function(c) newStr.Append(c))
        Return newStr.ToString()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function OnlyNumbers(ByVal str As [String]) As [String]
        Return str.OnlyChars("0123456789")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function OnlyChars(ByVal str As [String]) As [String]
        Return str.OnlyChars("abcdefghijklmnopqrstuvxwyzABCDEFGHIJKLMNOPQRSTUVXWYZ")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function OnlyConsonants(ByVal str As [String]) As [String]
        Return str.OnlyChars("bcdfghjklmnpqrstvxwyzBCDFGHJKLMNPQRSTVXWYZ")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function WithoutAccents(ByVal str As [String]) As [String]
        Dim chars1 = "áéíóúçãõàèìòùäëïöüÁÉÍÓÚÇÃÕÀÈÌÒÙÄËÏÖÜ"
        Dim chars2 = "aeioucaoaeiouaeiouAEIOUCAOAEIOUAEIOU"
        For x As Int32 = 0 To chars1.Length - 1
            str = str.Replace(chars1(x), chars2(x))
        Next
        Return str
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetHour(ByVal time As [Decimal]) As [String]
        Dim dHour = Math.Floor(time)

        Dim sHour = Convert.ToString(dHour.ToString("00")) & "hrs"

        Return sHour
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetHourMinute(ByVal time As [Decimal]) As [String]
        Dim dHour = Math.Floor(time)
        Dim dMinute = Math.Floor((time - dHour) * 60)

        Dim sHour = Convert.ToString(dHour.ToString("00")) & "hrs "
        Dim sMinute = Convert.ToString(dMinute.ToString("00")) & "min"

        Return sHour & sMinute
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetHourMinuteSecond(ByVal time As [Decimal]) As [String]
        Dim dHour = Math.Floor(time)
        Dim dMinute = Math.Floor((time - dHour) * 60)
        Dim dSecond = Math.Floor((((time - dHour) * 60) - dMinute) * 60)

        Dim sHour = Convert.ToString(dHour.ToString("00")) & "hrs "
        Dim sMinute = Convert.ToString(dMinute.ToString("00")) & "min "
        Dim sSecond = Convert.ToString(dSecond.ToString("00")) & "seg"

        Return sHour & sMinute & sSecond
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetNode(ByVal lst As XmlNodeList, ByVal name As [String]) As XmlNode
        If lst Is Nothing OrElse lst.Count = 0 Then
            Return Nothing
        End If
        Dim list = lst.Cast(Of XmlNode)().ToList()
        If list Is Nothing OrElse list.Count = 0 Then
            Return Nothing
        End If
        Return list.FirstOrDefault(Function(n) n.Name = name)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function TryParse(Of T)(ByVal theEnum As [Enum], ByVal valueToParse As [String], ByRef returnValue As T) As [Boolean]
        returnValue = Nothing
        Dim intEnumValue As Integer
        If Int32.TryParse(valueToParse, intEnumValue) Then
            If [Enum].IsDefined(GetType(T), intEnumValue) Then
                returnValue = DirectCast(DirectCast(intEnumValue, Object), T)
                Return True
            End If
        End If
        Return False
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetSelectedDataKeys(ByVal control As ListView, ByVal chkBoxId As [String]) As List(Of DataKey)
        Dim lst = New List(Of DataKey)()
        If control IsNot Nothing AndAlso control.Items IsNot Nothing AndAlso control.Items.Count > 0 Then
            lst = control.Items.Where(Function(x) IsChecked(x, chkBoxId)).[Select](Function(x) control.DataKeys(x.DisplayIndex)).ToList()
        End If
        Return lst
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SetSelectedDataKeys(ByVal control As ListView, ByVal chkBoxId As [String], ByVal lstIDs As List(Of Int32))
        For Each item As ListViewDataItem In control.Items
            Dim checkbox = DirectCast(item.FindControl(chkBoxId), CheckBox)
            If checkbox IsNot Nothing AndAlso item IsNot Nothing Then
                Dim id = Convert.ToInt32(control.DataKeys(item.DisplayIndex).Value)
                If id > 0 Then
                    checkbox.Checked = lstIDs.Contains(id)
                End If
            End If
        Next
    End Sub

    Private Function IsChecked(ByVal item As ListViewDataItem, ByVal chkBoxId As [String]) As [Boolean]
        Dim control = TryCast(item.FindControl(chkBoxId), CheckBox)
        If control Is Nothing Then
            Return False
        End If
        Return control.Checked
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetSelectedItems(ByVal rpt As Repeater, ByVal chkBoxId As [String]) As ArrayList
        Dim selectedValues = New ArrayList()
        For i As Integer = 0 To rpt.Items.Count - 1
            Dim chkBox = TryCast(rpt.Items(i).FindControl(chkBoxId), HtmlInputCheckBox)
            If chkBox IsNot Nothing AndAlso chkBox.Checked Then
                selectedValues.Add(chkBox.Value)
            End If
        Next
        Return selectedValues
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SetSelectedItems(ByVal rpt As Repeater, ByVal chkBoxId As [String], ByVal lstIDs As List(Of Int32))
        If rpt IsNot Nothing AndAlso rpt.Items IsNot Nothing AndAlso rpt.Items.Count > 0 Then
            For i As Integer = 0 To rpt.Items.Count - 1
                Dim chkBox = TryCast(rpt.Items(i).FindControl(chkBoxId), HtmlInputCheckBox)
                If chkBox IsNot Nothing AndAlso Not [String].IsNullOrEmpty(chkBox.Value) Then
                    chkBox.Checked = lstIDs.Contains(chkBox.Value.ToInt32())
                End If
            Next
        End If
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetSelectedItems(ByVal grd As GridView, ByVal chkBoxId As [String]) As List(Of Int32)
        Dim checkedIDs As New List(Of Int32)()
        For Each row As GridViewRow In grd.Rows
            Dim chk As CheckBox = DirectCast(row.FindControl(chkBoxId), CheckBox)
            If chk IsNot Nothing AndAlso chk.Checked Then
                checkedIDs.Add(Integer.Parse(grd.DataKeys(row.RowIndex).Value.ToString()))
            End If
        Next
        Return checkedIDs
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetSelectedValues(ByVal grd As GridView, ByVal chkBoxId As [String]) As List(Of String)
        Dim checkedIDs As New List(Of String)()
        For Each row As GridViewRow In grd.Rows
            Dim chk As CheckBox = DirectCast(row.FindControl(chkBoxId), CheckBox)
            If chk IsNot Nothing AndAlso chk.Checked Then
                checkedIDs.Add(grd.DataKeys(row.RowIndex).Value.ToString())
            End If
        Next
        Return checkedIDs
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SetSelectedItems(ByVal grd As GridView, ByVal chkBoxId As [String], ByVal lstIDs As List(Of Int32))
        If lstIDs IsNot Nothing AndAlso lstIDs.Count > 0 Then
            For Each row As GridViewRow In grd.Rows
                Dim chk As CheckBox = DirectCast(row.FindControl(chkBoxId), CheckBox)
                If chk IsNot Nothing Then
                    chk.Checked = lstIDs.Contains(Integer.Parse(grd.DataKeys(row.RowIndex).Value.ToString()))
                End If
            Next
        End If
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SelectValue(ByVal ddl As DropDownList, ByVal value As [String])
        Try
            If ddl.SelectedIndex >= 0 Then
                ddl.Items(ddl.SelectedIndex).Selected = False
            End If
            ddl.Items.FindByValue(value).Selected = True
        Catch
            Throw New GenericDropDownListException("value", value)
        End Try
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SelectText(ByVal ddl As DropDownList, ByVal text As [String])
        Try
            If ddl.SelectedIndex >= 0 Then
                ddl.Items(ddl.SelectedIndex).Selected = False
            End If
            ddl.Items.FindByText(text).Selected = True
        Catch
            Throw New GenericDropDownListException("value", text)
        End Try
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetSelecteds(ByVal lbox As ListBox) As List(Of String)
        Dim selectedValues As New List(Of String)
        Dim selectedIndices As Integer() = lbox.GetSelectedIndices()
        For Each i As Integer In selectedIndices
            selectedValues.Add(lbox.Items(i).Value)
        Next
        Return selectedValues
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetSelectedValues(ByVal lbox As ListBox) As ArrayList
        Dim selectedValues As New ArrayList()
        Dim selectedIndices As Integer() = lbox.GetSelectedIndices()
        For Each i As Integer In selectedIndices
            selectedValues.Add(lbox.Items(i).Value)
        Next
        Return selectedValues
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetSelectedItems(ByVal lbox As ListBox) As ArrayList
        Dim selectedValues As New ArrayList()
        Dim selectedIndices As Integer() = lbox.GetSelectedIndices()
        For Each i As Integer In selectedIndices
            selectedValues.Add(lbox.Items(i).Text)
        Next
        Return selectedValues
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SetSelectedValues(ByVal lbox As ListBox, ByVal values As [String]())
        For Each value As String In values
            If Not [String].IsNullOrEmpty(value) Then
                Dim item As ListItem = lbox.Items.FindByValue(value)
                If item IsNot Nothing Then
                    item.Selected = True
                End If
            End If
        Next
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub CheckAll(ByVal cbox As CheckBoxList)
        For Each li As ListItem In cbox.Items
            li.Selected = True
        Next
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub UncheckAll(ByVal cbox As CheckBoxList)
        For Each li As ListItem In cbox.Items
            li.Selected = False
        Next
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub ShowMessage(ByVal lblMsg As Label, ByVal msg As [String])
        lblMsg.Text = msg
        lblMsg.Visible = True
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub HideMessage(ByVal lblMsg As Label)
        lblMsg.Text = String.Empty
        lblMsg.Visible = False
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub PrintGridView(ByVal grd As GridView, ByVal page As Page, ByVal btn As Button)
        Dim printScript As String = "function PrintGridView()" & vbCr & vbLf & "                 {" & vbCr & vbLf & "                    var gridInsideDiv = document.getElementById('gvDiv');" & vbCr & vbLf & "                    var printWindow = window.open('gview.htm','PrintWindow','letf=0,top=0,width=150,height=300,toolbar=1,scrollbars=1,status=1');" & vbCr & vbLf & "                    printWindow.document.write(gridInsideDiv.innerHTML);" & vbCr & vbLf & "                    printWindow.document.close();" & vbCr & vbLf & "                    printWindow.focus();" & vbCr & vbLf & "                    printWindow.print();" & vbCr & vbLf & "                    printWindow.close();" & vbCr & vbLf & "                }"
        page.ClientScript.RegisterStartupScript(page.[GetType](), "PrintGridView", printScript.ToString(), True)
        btn.Attributes.Add("onclick", "PrintGridView();")
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Function FindControlRecursive(ByVal root As Control, ByVal controlID As [String]) As Control
        If root.ID = controlID Then
            Return root
        End If

        For Each ctl As Control In root.Controls
            Dim ctlFound As Control = FindControlRecursive(ctl, controlID)
            If ctlFound IsNot Nothing Then
                Return ctlFound
            End If
        Next

        Return Nothing
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsValid(ByVal txt As TextBox) As [Boolean]
        If txt Is Nothing Then
            Return False
        End If
        Return Not [String].IsNullOrWhiteSpace(txt.Text)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsValid(ByVal ddl As DropDownList) As [Boolean]
        If ddl Is Nothing Then
            Return False
        End If
        Return (ddl.SelectedItem IsNot Nothing AndAlso ddl.SelectedItem.Value.ToInt32() > 0)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsValid(ByVal chk As CheckBox) As [Boolean]
        If chk Is Nothing Then
            Return False
        End If
        Return chk.Checked
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsValid(ByVal rdo As RadioButton) As [Boolean]
        If rdo Is Nothing Then
            Return False
        End If
        Return rdo.Checked
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub [Select](ByVal txt As TextBox)
        If ScriptManager.GetCurrent(txt.Page) IsNot Nothing AndAlso ScriptManager.GetCurrent(txt.Page).IsInAsyncPostBack Then
            ScriptManager.RegisterStartupScript(txt.Page, txt.Page.[GetType](), Guid.NewGuid().ToString(), [String].Format("ctrlToSelect='{0}';", txt.ClientID), True)
        Else
            txt.Page.ClientScript.RegisterStartupScript(txt.Page.[GetType](), Guid.NewGuid().ToString(), [String].Format("document.getElementById('{0}').select();", txt.ClientID), True)
        End If
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub SetText(ByVal lbl As Label, ByVal value As [String])
        If lbl Is Nothing Then
            Return
        End If
        If Not [String].IsNullOrWhiteSpace(lbl.Text) Then
            lbl.Text += " "
        End If
        lbl.Text += value
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsCnpj(ByVal cnpj As String) As Boolean
        If Not String.IsNullOrEmpty(cnpj) Then
            cnpj = cnpj.Trim()
        End If
        If Not cnpj.IsMatch("(^(\d{2,3}.\d{3}.\d{3}/\d{4}-\d{2})|(\d{14,15})$)") Then
            Return False
        End If

        cnpj = cnpj.Replace("/", "")
        cnpj = cnpj.Replace(".", "")
        cnpj = cnpj.Replace("-", "")

        Dim digits As String
        If Not (cnpj.Length >= 14 AndAlso cnpj.Length <= 15) OrElse cnpj = "0".Repeat(cnpj.Length) Then
            Return False
        End If

        digits = cnpj.Substring(cnpj.Length - 2)
        Return verifyDigit(Convert.ToInt32(digits.Substring(0, 1)), cnpj.Substring(0, cnpj.Length - 2)) AndAlso verifyDigit(Convert.ToInt32(digits.Substring(1, 1)), cnpj.Substring(0, cnpj.Length - 1))
    End Function

    Private Function verifyDigit(ByVal digit As Int32, ByVal numbers As String) As Boolean
        Dim sum As Int32 = 0, i As Int32, result As Int32
        Dim pos As Int32 = numbers.Length - 7
        For i = 0 To numbers.Length - 1
            sum += Convert.ToInt32(numbers.Substring(i, 1)) * System.Math.Max(System.Threading.Interlocked.Decrement(pos), pos + 1)
            If pos < 2 Then
                pos = 9
            End If
        Next
        result = If(sum Mod 11 < 2, 0, 11 - (sum Mod 11))
        Return result = digit
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Repeat(ByVal str As String, ByVal n As Int32) As String
        Dim rs = ""
        For x As Int32 = 0 To n - 1
            rs += str
        Next
        Return rs
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsMatch(ByVal str As String, ByVal pattern As String) As Boolean
        If String.IsNullOrEmpty(str) Then
            Return False
        End If
        Return Regex.IsMatch(str, pattern, RegexOptions.IgnoreCase)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function DistinctBy(Of TSource, TKey)(ByVal source As IEnumerable(Of TSource), ByVal keySelector As Func(Of TSource, TKey)) As IEnumerable(Of TSource)
        Dim seenKeys As New HashSet(Of TKey)()
        For Each element As TSource In source
            If seenKeys.Add(keySelector(element)) Then
                Return New List(Of TSource) From {element}
            End If
        Next
        Return Nothing
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToSqlNULL(ByVal pInt As Integer) As String
        If pInt = 0 Then
            Return "NULL"
        Else : Return pInt
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToSqlDate(ByVal pDate As Date) As String
        Return pDate.ToString("yyyy-MM-dd")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToSqlDate(ByVal pDate As String) As String
        If IsDate(pDate) Then
            Return CDate(pDate).ToString("yyyy-MM-dd")
        Else
            Return ("yyyy-MM-dd")
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function RemoveMask(ByVal pCpfCnpj As String) As String
        'Return pCpfCnpj.Replace(".", "").Replace("-", "").Replace("/", "").Replace("(", "").Replace(")", "").Replace(" ", "")
        Return String.Join("", System.Text.RegularExpressions.Regex.Split(pCpfCnpj, "[^\d]"))
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToStrDate(ByVal pDate As Date) As String
        Return pDate.ToString("dd-MM-yyyy")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToStrDate(ByVal pDate As String) As String
        If IsDate(pDate) Then
            Return CDate(pDate).ToString("dd-MM-yyyy")
        Else
            Return ("dd-MM-yyyy")
        End If
    End Function
End Module

<Serializable()> _
Public Class GenericDropDownListException
    Inherits Exception
    Implements ISerializable

    Public Sub New(ByVal type As [String], ByVal value As [String])
        MyBase.New(String.Format("Unable to set  ""{0}"" to {1}", type, value))
    End Sub

End Class