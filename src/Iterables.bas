Attribute VB_Name = "Iterables"
Option Explicit

Private Const ITERABLE_ERROR As Long = 438

Public Function Includes(ByVal IterableObject As Variant, ByVal SearchedElement As Variant) As Boolean

    Dim Found As Boolean
    Dim i As Long
    Found = False
    If IsArray(IterableObject) Then
        For i = 0 To UBound(IterableObject)
            If SearchedElement = IterableObject(i) Then
                Found = True
                Exit For
            End If
        Next i
    ElseIf IsCollection(IterableObject) Then
        For i = 1 To IterableObject.Count
            If SearchedElement = IterableObject(i) Then
                Found = True
                Exit For
            End If
        Next i
    ElseIf IsDictionary(IterableObject) Then
        Dim Elem As Variant
        For Each Elem In IterableObject
            If SearchedElement = IterableObject(Elem) Then
                Found = True
                Exit For
            End If
        Next Elem
    Else
        Err.Raise ITERABLE_ERROR
    End If

    Includes = Found
    
End Function

Public Function IsCollection(ByVal TestedArgument As Variant) As Boolean
    
    IsCollection = TypeName(TestedArgument) = "Collection"
    
End Function

Public Function IsDictionary(ByVal TestedArgument As Variant) As Boolean

    IsDictionary = TypeName(TestedArgument) = "Dictionary"

End Function

Public Function ArrayLength(ByVal TestedArray As Variant) As Long

    If Not IsArray(TestedArray) Then Err.Raise ITERABLE_ERROR
    ArrayLength = UBound(TestedArray) - LBound(TestedArray) + 1

End Function

Public Function ConcatArrays(ByVal FirstArray As Variant, ByVal SecondArray As Variant) As Variant

    Dim NewArrayLength As Long
    Dim NewArray As Variant
    
    NewArray = Null
    
    NewArrayLength = ArrayLength(FirstArray) + ArrayLength(SecondArray) - 1
    ReDim NewArray(0 To NewArrayLength)
    Dim i As Long
    Dim j As Long
    j = 0
    For i = LBound(FirstArray) To UBound(FirstArray)
        NewArray(j) = FirstArray(i)
        j = j + 1
    Next i
    For i = LBound(SecondArray) To UBound(SecondArray)
        NewArray(j) = SecondArray(i)
        j = j + 1
    Next i

    ConcatArrays = NewArray

End Function

Public Function ConcatCollections(ByVal Collection1 As Collection, Collection2 As Collection) As Collection

    Dim i As Long
    For i = 1 To Collection2.Count
        Collection1.add Collection2(i)
    Next i
    
    Set ConcatCollections = Collection1

End Function

Public Function MergeDictionaries(ByVal Dictionary1 As Object, Dictionary2 As Object) As Object

    Dim Elem As Variant
    For Each Elem In Dictionary2
        If Not Dictionary1.Exists(Elem) Then
            Dictionary1.add Elem, Dictionary2(Elem)
        End If
    Next Elem
    
    Set MergeDictionaries = Dictionary1

End Function

