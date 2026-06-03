Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Reflection
Imports System.ComponentModel

Public Class WebHelpers
    Inherits Page

    Public Shared Function GetEnumDescription(ByVal value As [Enum]) As [String]
        Dim fi As FieldInfo = value.[GetType]().GetField(value.ToString())
        Dim attributes = DirectCast(fi.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
        Return If((attributes.Length > 0), attributes(0).Description, value.ToString())
    End Function

    Public Shared Function FindControlRecursive(ByVal root As System.Web.UI.Control, ByVal id As String) As System.Web.UI.Control
        If root.ID = id Then
            Return root
        End If
        For Each c As System.Web.UI.Control In root.Controls
            Dim t As System.Web.UI.Control = FindControlRecursive(c, id)
            If t IsNot Nothing Then
                Return t
            End If
        Next
        Return Nothing
    End Function

    Public Shared Sub BindDropDownWithEnum(ByRef drop As DropDownList, ByVal myEnum As Type)
        Dim item As New ListItem()
        drop.Items.Clear()
        Dim names = [Enum].GetNames(myEnum)
        For i As Integer = 0 To names.GetUpperBound(0)
            With item
                .Text = GetEnumDescription(DirectCast([Enum].Parse(myEnum, names(i)), [Enum]))
                .Value = Convert.ToInt32([Enum].Parse(myEnum, names(i))).ToString()
                drop.Items.Add(item)
            End With
        Next
    End Sub

    Public Shared Sub BindDropDownWithEnum(ByRef drop As DropDownList, ByVal myEnum As Type, ByVal addSelecione As Boolean, ByVal firstItemText As [String])
        drop.Items.Clear()
        Dim names As [String]() = [Enum].GetNames(myEnum)
        For i As Integer = 0 To names.GetUpperBound(0)
            Dim item = New ListItem()
            With item
                .Text = GetEnumDescription(DirectCast([Enum].Parse(myEnum, names(i)), [Enum]))
                .Value = Convert.ToInt32([Enum].Parse(myEnum, names(i))).ToString()
            End With
            drop.Items.Add(item)
        Next
        If addSelecione Then
            drop.Items.Insert(0, New ListItem(firstItemText, "0"))
        End If
    End Sub

End Class