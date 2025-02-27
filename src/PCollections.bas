Attribute VB_Name = "PCollections"
'@Folder "PCollectionsProject.src"
Option Explicit

'@Description "Adds all items from the Source collection to the Destination collection."
Public Sub Extend(ByRef Destination As Collection, ByRef Source As Collection)
Attribute Extend.VB_Description = "Adds all items from the Source collection to the Destination collection."
    If Destination Is Nothing Then Exit Sub
    If Source Is Nothing Then Exit Sub
    If Source.Count() = 0 Then Exit Sub

    Dim Item As Variant
    For Each Item In Source
        Destination.Add Item
    Next
End Sub

'@Description "Finds and returns a new collection containing all occurrences of the specified item in the Source collection."
Public Function FindAll(ByRef SourceCollection As Collection, ByVal Item As Variant) As Collection
Attribute FindAll.VB_Description = "Finds and returns a new collection containing all occurrences of the specified item in the Source collection."
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

'@Description "Returns the index of the first occurrence of the specified item in the Source collection, or -1 if not found."
Public Function FindIndex(ByRef SourceCollection As Collection, ByVal Item As Variant) As Long
Attribute FindIndex.VB_Description = "Returns the index of the first occurrence of the specified item in the Source collection, or -1 if not found."
    Dim Items As Collection: Set Items = New Collection

    Dim i As Long
    For i = 1 To SourceCollection.Count()
        Dim Found As Boolean
        If Information.IsObject(SourceCollection(i)) Then
            Found = SourceCollection(i) Is Item
        Else
            Found = SourceCollection(i) = Item
        End If

        If Found Then
            FindIndex = i
            Exit Function
        End If
    Next

    FindIndex = -1
End Function

'@Description "Returns the index of the last occurrence of the specified item in the Source collection, or -1 if not found."
Public Function FindLastIndex(ByRef SourceCollection As Collection, ByVal Item As Variant) As Long
Attribute FindLastIndex.VB_Description = "Returns the index of the last occurrence of the specified item in the Source collection, or -1 if not found."
    Dim Items As Collection: Set Items = New Collection

    Dim i As Long
    For i = SourceCollection.Count() To 1 Step -1
        Dim Found As Boolean
        If Information.IsObject(SourceCollection(i)) Then
            Found = SourceCollection(i) Is Item
        Else
            Found = SourceCollection(i) = Item
        End If

        If Found Then
            FindLastIndex = i
            Exit Function
        End If
    Next

    FindLastIndex = -1
End Function

'@Description "Creates and returns a new collection from the given array."
Public Function FromArray(ByRef SourceArray As Variant) As Collection
Attribute FromArray.VB_Description = "Creates and returns a new collection from the given array."
    Dim Buffer As Collection: Set Buffer = New Collection
    Dim Item As Variant
    For Each Item In SourceArray
        Buffer.Add Item
    Next

    Set FromArray = Buffer
End Function

'@Description "Checks if the specified item exists in the Source collection and returns True or False."
Public Function ItemExists(ByRef SourceCollection As Collection, ByVal Item As Variant) As Boolean
Attribute ItemExists.VB_Description = "Checks if the specified item exists in the Source collection and returns True or False."
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

'@Description "Concatenates all items in the Source collection into a string, separated by the specified delimiter."
Public Function Join(ByRef SourceCollection As Collection, Optional ByVal Delimiter As String = " ") As String
Attribute Join.VB_Description = "Concatenates all items in the Source collection into a string, separated by the specified delimiter."
    Dim Items As Variant: Items = PCollections.ToArray(SourceCollection:=SourceCollection)
    Join = Strings.Join(Items, Delimiter)
End Function

'@Description "Checks if a specified key exists in the Source collection and returns True or False."
Public Function KeyExists(ByRef SourceCollection As Collection, ByVal Key As Variant) As Boolean
Attribute KeyExists.VB_Description = "Checks if a specified key exists in the Source collection and returns True or False."
    On Error Resume Next
    Dim Dummy As Boolean
    Dummy = Information.IsObject(SourceCollection(Key))

    KeyExists = Information.Err.Number = 0
End Function

'@Description "Returns the maximum value in the Source collection, raising an error if it contains objects."
Public Function Max(ByRef SourceCollection As Collection) As Variant
Attribute Max.VB_Description = "Returns the maximum value in the Source collection, raising an error if it contains objects."
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

'@Description "Returns the minimum value in the Source collection, raising an error if it contains objects."
Public Function Min(ByRef SourceCollection As Collection) As Variant
Attribute Min.VB_Description = "Returns the minimum value in the Source collection, raising an error if it contains objects."
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

'@Description "Removes and returns the last item from the Source collection."
Public Function Pop(ByRef SourceCollection As Collection) As Variant
Attribute Pop.VB_Description = "Removes and returns the last item from the Source collection."
    If SourceCollection Is Nothing Then Information.Err.Raise 9
    If SourceCollection.Count() = 0 Then Information.Err.Raise 9

    Dim LastIndex As Long: LastIndex = SourceCollection.Count()

    If Information.IsObject(SourceCollection.Item(LastIndex)) Then
        Set Pop = SourceCollection.Item(LastIndex)
    Else
        Pop = SourceCollection.Item(LastIndex)
    End If

    SourceCollection.Remove Index:=LastIndex
End Function

'@Description "Splits a string by the specified delimiter and returns a collection of the resulting substrings."
Public Function Split(ByVal Expression As String, Optional ByVal Delimiter As String = " ") As Collection
Attribute Split.VB_Description = "Splits a string by the specified delimiter and returns a collection of the resulting substrings."
    Set Split = PCollections.FromArray(Strings.Split(Expression, Delimiter))
End Function

'@Description "Returns a new collection with items from the Source collection in reverse order."
Public Function Reverse(ByRef SourceCollection As Collection) As Collection
Attribute Reverse.VB_Description = "Returns a new collection with items from the Source collection in reverse order."
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

'@Description "Converts the Source collection into an array and returns it."
Public Function ToArray(ByRef SourceCollection As Collection) As Variant
Attribute ToArray.VB_Description = "Converts the Source collection into an array and returns it."
    If SourceCollection Is Nothing Then
        ToArray = Array()
        Exit Function
    End If

    If SourceCollection.Count() = 0 Then
        ToArray = Array()
        Exit Function
    End If

    Dim Items() As Variant: ReDim Items(0 To SourceCollection.Count() - 1)

    Dim i As Long
    For i = 1 To SourceCollection.Count()
        If Information.IsObject(SourceCollection(i)) Then
            Set Items(i - 1) = SourceCollection(i)
        Else
            Items(i - 1) = SourceCollection(i)
        End If
    Next

    ToArray = Items
End Function
