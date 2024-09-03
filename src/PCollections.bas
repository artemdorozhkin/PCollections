Attribute VB_Name = "PCollections"
'@Folder "PCollectionsProject.src"
Option Explicit

Public Sub Extend(ByRef Collection1 As Collection, ByRef Collection2 As Collection)
    If Collection1 Is Nothing Then Exit Function
    If Collection2 Is Nothing Then Exit Function
    If Collection2.Count() = 0 Then Exit Sub

    Dim Item As Variant
    For Each Item In Collection2
        Collection1.Add Item
    Next
End Sub

Public Function FindAll(ByRef SourceCollection As Collection, ByVal Item As Variant) As Collection
    Dim Items As Collection: Set Items = New Collection

    Dim Check As Variant
    For Each Check In SourceCollection
        If Information.IsObject(Item) Then
            If Item Is Check Then Items.Add Check
        Else
            If Item = Check Then Items.Add Check
        End If
    Next

    Set FindAll = Items
End Function

Public Function FindIndex(ByRef SourceCollection As Collection, ByVal Item As Variant) As Long
    Dim Items As Collection: Set Items = New Collection

    Dim i As Long
    For i = 1 To SourceCollection.Count()
        Dim Found As Boolean
        If Information.IsObject(SourceCollection(i)) Then
            Found = SourceCollection(i) Is Check
        Else
            Found = SourceCollection(i) = Check
        End If

        If Found Then
            FindIndex = i
            Exit Function
        End If
    Next

    FindIndex = -1
End Function

Public Function FindLastIndex(ByRef SourceCollection As Collection, ByVal Item As Variant) As Long
    Dim Items As Collection: Set Items = New Collection

    Dim i As Long
    For i = SourceCollection.Count() To 1 Step -1
        Dim Found As Boolean
        If Information.IsObject(SourceCollection(i)) Then
            Found = SourceCollection(i) Is Check
        Else
            Found = SourceCollection(i) = Check
        End If

        If Found Then
            FindLastIndex = i
            Exit Function
        End If
    Next

    FindLastIndex = -1
End Function

Public Function FromArray(ByRef SourceArray As Variant) As Collection
    Dim Buffer As Collection: Set Buffer = New Collection
    Dim Item As Variant
    For Each Item In SourceArray
        Buffer.Add Item
    Next

    Set FromArray = Buffer
End Function

Public Function ItemExists(ByRef SourceCollection As Collection, ByVal Item As Variant) As Boolean
    Dim Check As Variant
    For Each Check In SourceCollection
        Dim IsSame As Boolean
        If Information.IsObject(Item) Then
            IsSame = Item Is Check
        Else
            IsSame = Item = Check
        End If

        If IsSame Then
            ItemExists = True
            Exit Function
        End If
    Next
End Function

Public Function Join(ByRef SourceCollection As Collection, Optional ByVal Delimiter As String = " ") As String
    Dim Items As Variant: Items = PCollections.ToArray(SourceCollection:=SourceCollection)
    Join = Strings.Join(Items, Delimiter)
End Function

Public Function KeyExists(ByRef SourceCollection As Collection, ByVal Key As Variant) As Boolean
    On Error Resume Next
    Dim Dummy As Boolean
    Dummy = Information.IsObject(SourceCollection(Key))

    KeyExists = Information.Err.Number <> 0
End Function

Public Function Max(ByRef SourceCollection As Collection) As Variant
    Dim MaxItem As Variant
    Dim Item As Variant
    For Each Item In SourceCollection
        If Information.IsObject(Item) Then
            Information.Err.Raise _
                Number:=5, _
                Source:="PCollections.Min", _
                Description:="Function Max can't compare objects"
        End If

        If Information.IsEmpty(MaxItem) Then
            MaxItem = Item
        Else
            If MaxItem < Item Then MaxItem = Item
        End If
    Next

    Max = MaxItem
End Function

Public Function Min(ByRef SourceCollection As Collection) As Variant
    Dim MinItem As Variant
    Dim Item As Variant
    For Each Item In SourceCollection
        If Information.IsObject(Item) Then
            Information.Err.Raise _
                Number:=5, _
                Source:="PCollections.Min", _
                Description:="Function Min can't compare objects"
        End If

        If Information.IsEmpty(MinItem) Then
            MinItem = Item
        Else
            If MinItem > Item Then MinItem = Item
        End If
    Next

    Min = MinItem
End Function

Public Function Pop(ByRef SourceCollection As Collection) As Variant
    If SourceCollection Is Nothing Then Exit Function
    If SourceCollection.Count() = 0 Then Exit Function

    Dim LastIndex As Long: LastIndex = SourceCollection.Count()

    If Information.IsObject(SourceCollection.Item(LastIndex)) Then
        Set Pop = SourceCollection.Item(LastIndex)
    Else
        Pop = SourceCollection.Item(LastIndex)
    End If

    SourceCollection.Remove Index:=LastIndex
End Function

Public Function Split(ByVal SourceText As String, Optional ByVal Delimiter As String = " ") As Collection
    Set Split = PCollections.FromArray(Strings.Split(SourceText, Delimiter))
End Function

Public Function Reverse(ByRef SourceCollection As Collection) As Collection
    If SourceCollection Is Nothing Then Exit Function
    If SourceCollection.Count() = 0 Then
        Set Reverse = SourceCollection
        Exit Function
    End If

    Dim Reversed As Collection: Set Reversed = New Collection
    Dim i As Long
    For i = SourceCollection.Count() To 1 Step -1
        Reversed.Add SourceCollection(i)
    Next

    Set Reverse = Reversed
End Function

Public Function ToArray(ByRef SourceCollection As Collection) As Variant
    If SourceCollection Is Nothing Then
        ToArray = Array()
        Exit Function
    End If

    If SourceCollection.Count() = 0 Then
        ToArray = Array()
        Exit Function
    End If

    Dim Items() As Variant: ReDim Items(1 To SourceCollection.Count())

    Dim i As Long
    For i = 1 To SourceCollection.Count()
        If Information.IsObject(SourceCollection(i)) Then
            Set Items(i) = SourceCollection(i)
        Else
            Items(i) = SourceCollection(i)
        End If
    Next

    ToArray = Items
End Function
