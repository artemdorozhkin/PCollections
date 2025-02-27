Attribute VB_Name = "Examples"
'@Folder("PCollectionsProject")
Option Explicit

Public Sub ExampleItemExists()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print PCollections.ItemExists(Coll, "2b") ' True
    Debug.Print PCollections.ItemExists(Coll, "2a") ' False
End Sub

Public Sub ExampleKeyExists()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add Item:=1, Key:="1a"
    Coll.Add Item:=2, Key:="2b"
    Coll.Add Item:=3, Key:="3c"

    Debug.Print PCollections.KeyExists(Coll, "2b") ' True
    Debug.Print PCollections.KeyExists(Coll, "2a") ' False
End Sub

Public Sub ExampleJoin()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print PCollections.Join(Coll, ", ") ' 1a, 2b, 3c
End Sub

Public Sub ExampleSplit()
    Dim Coll As Collection
    Set Coll = PCollections.Split("1a, 2b, 3c", ", ")

    Debug.Print PCollections.Join(Coll, ", ") ' 1a, 2b, 3c
End Sub

Public Sub ExampleToArray()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print Strings.Join(PCollection.ToArray(Coll), ", ") ' 1a, 2b, 3c
End Sub

Public Sub ExampleFromArray()
    Dim Coll As Collection
    Set Coll = PCollections.FromArray( _
        Array("1a", "2b", "3c") _
    )

    Debug.Print Coll.Count ' 3
    Debug.Print PCollections.Join(Coll, ", ") ' 1a, 2b, 3c
End Sub

Public Sub ExampleExtend()
    Dim Coll1 As Collection
    Set Coll1 = New Collection

    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"

    Dim Coll2 As Collection
    Set Coll2 = New Collection

    Coll2.Add "4d"
    Coll2.Add "5e"
    Coll2.Add "6f"

    PCollections.Extend Coll1, Coll2
    Debug.Print PCollections.Join(Coll1, ", ") ' 1a, 2b, 3c, 4d, 5e, 6f
End Sub

Public Sub ExampleFindAll()
    Dim Coll1 As Collection
    Set Coll1 = New Collection

    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"
    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"

    Dim Coll2 As Collection
    Set Coll2 = PCollections.FindAll(Coll1, "2b")
    Debug.Print PCollections.Join(Coll2, ", ") ' 2b, 2b
End Sub

Public Sub ExampleFindIndex()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print PCollections.FindIndex(Coll, "2b") ' 2
    Debug.Print PCollections.FindIndex(Coll, "4e") ' -1
End Sub

Public Sub ExampleFindLastIndex()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"
    Coll.Add "1a"
    Coll.Add "2b"
    Coll.Add "3c"

    Debug.Print PCollections.FindLastIndex(Coll, "2b") ' 5
    Debug.Print PCollections.FindLastIndex(Coll, "4e") ' -1
End Sub

Public Sub ExampleMax()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    Debug.Print PCollections.Max(Coll) ' 482
End Sub

Public Sub ExampleMin()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    Debug.Print PCollections.Min(Coll) ' 1
End Sub

Public Sub ExamplePop()
    Dim Coll As Collection
    Set Coll = New Collection

    Coll.Add 123
    Coll.Add 3
    Coll.Add 482
    Coll.Add 69
    Coll.Add 1
    Coll.Add 4

    Dim Item As Long
    Item = PCollections.Pop(Coll)
    Debug.Print Item ' 4
End Sub

Public Sub ExampleReverse()
    Dim Coll1 As Collection
    Set Coll1 = New Collection

    Coll1.Add "1a"
    Coll1.Add "2b"
    Coll1.Add "3c"

    Dim Coll2 As Collection
    Set Coll2 = PCollections.Reverse(Coll1)
    Debug.Print PCollections.Join(Coll2, ", ") ' 3c, 2b, 1a
End Sub
