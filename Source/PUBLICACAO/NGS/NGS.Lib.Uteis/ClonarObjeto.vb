Imports System.Collections
Imports System.Collections.Generic
Imports System.Reflection

''' <summary>
''' Generic Clone Class
''' </summary>
Public NotInheritable Class ClonarObjeto(Of T)
    ''' <summary>
    ''' Cache para os Campos do(s) Tipo(s) Utilizado(s)
    ''' </summary>
    Private ReadOnly _FieldsCache As New Dictionary(Of Type, Dictionary(Of String, FieldInfo))()
    ''' <summary>
    ''' Clona um Objeto
    ''' </summary>
    ''' <param name="target"></param>
    ''' <returns></returns>
    Public Function Clonar(ByVal target As T) As T
        ' Create Objetos
        Dim resultSet As Object
        Dim typeFields As Dictionary(Of String, FieldInfo)
        Dim typeLessConstructor As ConstructorInfo
        Dim isICloneType As Type
        Dim iClone As ICloneable
        Dim isIEnumerableType As Type
        Dim iEnum As IEnumerable
        Dim isIListType As Type
        Dim isIDicType As Type
        Dim list As IList
        Dim dic As IDictionary
        ' Retreave Default TypeLess Constructor (Public or NonPublic)
        typeLessConstructor = target.[GetType]().GetConstructor(BindingFlags.NonPublic Or BindingFlags.[Public] Or BindingFlags.Instance, Nothing, New Type() {}, Nothing)
        ' Initialize ResultSet Object
        resultSet = typeLessConstructor.Invoke(Nothing)
        ' Retreave All Type Field's
        typeFields = GetFields(target.[GetType]())
        ' Looping on Fields
        For Each info As KeyValuePair(Of String, FieldInfo) In typeFields
            'Query if the fiels support the ICloneable Interface
            isICloneType = info.Value.FieldType.GetInterface("ICloneable", False)
            ' Supports ICloneable Interface, Then, Use Direct Clone Approach
            If isICloneType IsNot Nothing Then
                ' Getting the ICloneable interface from the object.
                iClone = DirectCast(info.Value.GetValue(target), ICloneable)
                ' If Value Are NotNull
                If iClone IsNot Nothing Then
                    Dim clonedObject As Object = iClone.Clone()
                    'We use the clone method to set the new value to the field.
                    info.Value.SetValue(resultSet, clonedObject)
                End If
            Else
                ' Property Not Support ICloneable Interface
                info.Value.SetValue(resultSet, info.Value.GetValue(target))
            End If
            ' Check If Object Supports a IEnumerable interface so if it does we need to enumerate all its items and check if they support the ICloneable interface.
            isIEnumerableType = info.Value.FieldType.GetInterface("IEnumerable", False)
            ' Check Result for IEnumerable Interface Search
            If isIEnumerableType IsNot Nothing Then
                ' Get the IEnumerable interface from the field.
                iEnum = DirectCast(info.Value.GetValue(target), IEnumerable)
                ' Get The IList and the IDictionary Interfaces to iterate on collections.
                isIListType = info.Value.FieldType.GetInterface("IList", False)
                isIDicType = info.Value.FieldType.GetInterface("IDictionary", False)
                ' If Collection Supports a IList Interface
                If isIListType IsNot Nothing Then
                    ' Getting the IList interface.
                    list = DirectCast(info.Value.GetValue(resultSet), IList)
                    ' Looping on IEnumerable Itens
                    If list IsNot Nothing Then
                        For Each obj As Object In iEnum
                            ' Checking to see if the current item support the ICloneable interface.
                            isICloneType = obj.[GetType]().GetInterface("ICloneable", False)
                            ' Supports a Clone Type
                            If isICloneType IsNot Nothing Then
                                ' If it does support the ICloneable interface,
                                ' we use it to set the clone of
                                ' the object in the list.
                                Dim clone__1 As ICloneable = DirectCast(obj, ICloneable)
                                list.Add(clone__1.Clone())
                            End If
                        Next
                    End If
                ElseIf isIDicType IsNot Nothing Then
                    ' IDictionary Interface
                    ' Getting the dictionary interface.
                    dic = DirectCast(info.Value.GetValue(resultSet), IDictionary)
                    ' Looping on
                    If iEnum IsNot Nothing Then
                        For Each de As DictionaryEntry In iEnum
                            ' Checking to see if the item support the ICloneable interface.
                            isICloneType = de.Value.[GetType]().GetInterface("ICloneable", False)
                            If isICloneType IsNot Nothing Then
                                Dim clone__1 As ICloneable = DirectCast(de.Value, ICloneable)
                                dic(de.Key) = clone__1.Clone()
                            End If
                        Next
                    End If
                End If
            End If
        Next
        ' Returns a Clonned ResultSet Object
        Return DirectCast(resultSet, T)
    End Function
    ''' <summary>
    ''' Recupera os Fields de um Objeto pelo Tipo
    ''' </summary>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Private Function GetFields(ByVal type As Type) As Dictionary(Of String, FieldInfo)
        ' Verifica no Cache
        If Not _FieldsCache.ContainsKey(type) Then
            SyncLock GetType(ClonarObjeto(Of T))
                If Not _FieldsCache.ContainsKey(type) Then
                    _FieldsCache.Add(type, GetAllFields(type))
                End If
            End SyncLock
        End If
        ' Retorna do Cache
        Return _FieldsCache(type)
    End Function
    ''' <summary>
    ''' Método para Extrair os Fields de um Determinado Objeto
    ''' </summary>
    ''' <param name="type"></param>
    ''' <returns></returns>
    Private Function GetAllFields(ByVal type As Type) As Dictionary(Of String, FieldInfo)
        ' Cria Objeto
        Dim fields As New Dictionary(Of String, FieldInfo)()
        For Each info As FieldInfo In type.GetFields(BindingFlags.[Public] Or BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.FlattenHierarchy)
            If Not fields.ContainsKey(info.Name) Then
                fields.Add(info.Name, info)
            End If
        Next
        If type.BaseType IsNot Nothing Then
            Dim baseFields As Dictionary(Of String, FieldInfo) = GetAllFields(type.BaseType)
            For Each pair As KeyValuePair(Of String, FieldInfo) In baseFields
                If Not fields.ContainsKey(pair.Key) Then
                    fields.Add(pair.Key, pair.Value)
                End If
            Next
        End If
        Return fields
    End Function
End Class
