VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'                   INSTRUCTIONS
'
'To get an item from the ArrayList, you must use the Item() property.
'   i.e., you must say ArrayList1.Item(0) = foo
'   You cannot use the syntax ArrayList1(0) or ArrayList1[0]
'
'Note: the elements inside the ArrayList cannot be processed using a for-each loop

Private elementData() As Variant
Private size As Long
Private actualCap As Long

Private Sub Class_Initialize()
    '10 initial elements
    ReDim elementData(9)
    size = 0
    actualCap = 10
End Sub

Private Sub Class_Terminate()
    Erase elementData
End Sub

Public Property Get Capacity() As Long
  'all capacity in the array, including unused space
    Dim l As Long
    l = 0
    Dim e As Variant
    For Each e In elementData
        l = l + 1
    Next e
    
    Capacity = l
End Property
Public Property Let Capacity(ByVal TotalCapacity As Long)
    ReDim Preserve elementData(TotalCapacity - 1)
    actualCap = TotalCapacity
End Property

Public Function Length() As Long
'includes only used elements in the array
    Dim ct As Long
    ct = 0
    Dim e As Variant
    For Each e In elementData
        If Not VBA.isEmpty(e) Then
            ct = ct + 1
        End If
    Next e
    Length = ct
End Function

'begin standard arraylist resizing methods

Public Sub trimToSize()

    'goal: if capacity is large and length < 50% of capacity,
    'trim total capacity to:
    '   (number of used elements * 1.5)
    
    'only worth messing with capacity if the capacity is significantly large, say 100+
    If actualCap > 99 Then
        If (size < (actualCap / 2)) Then
            Dim newUBound As Long
            newUBound = Conversion.CLng(size * 1.5)
            If newUBound < 9 Then 'need at least 10 els
                newUBound = 9
            End If
            ReDim Preserve elementData(newUBound)
            actualCap = newUBound + 1
        End If
    End If
End Sub


Private Sub ensureCapacity(ByVal minCapacity As Long)
    If actualCap < minCapacity Then
        Dim newUBound As Long
        newUBound = Conversion.CLng(actualCap * 1.5)
        ReDim Preserve elementData(newUBound)
        actualCap = newUBound + 1
    End If
End Sub
'end resizing methods

Public Function isEmpty() As Boolean
    isEmpty = (size = 0)
End Function


'BEGIN STANDARD COLLECTION METHODS
Public Sub Add(Item As Variant, Optional ByVal Before As Long = -1)
'Inserts the specified element at the specified position in this
'* list. Shifts the element currently at that position (if any) and
'* any subsequent elements to the right (adds one to their indices).
    Call ensureCapacity(size + 1)
    
    'shift everything to the right of Before by 1
    If (Before > -1) Then
        checkIndex (Before)
        Dim temp() As Variant
        ReDim temp(size)
        
        Call arrayCopy(elementData(), Before, temp(), 0, size - Before)
        Call arrayCopy(temp(), 0, elementData(), Before + 1, size - Before)
        If Not IsObject(Item) Then
           elementData(Before) = Item
        Else
            Set elementData(Before) = Item
        End If
    Else ' no "Before" param
        If Not IsObject(Item) Then
            elementData(size) = Item
        Else
            Set elementData(size) = Item
        End If
    End If
    size = size + 1
End Sub

Sub removeAt(ByVal index As Long)
    checkIndex (index)
    
    If index < size - 1 Then
        Dim i As Integer
        For i = index To size - 1
            If Not IsObject(elementData(i + 1)) Then
                elementData(i) = elementData(i + 1)
            Else
                Set elementData(i) = elementData(i + 1)
            End If
        Next i
        elementData(size - 1) = Empty
    ElseIf index = size - 1 Then
        elementData(size - 1) = Empty
    End If
    
    size = size - 1
    Call trimToSize

End Sub

Public Property Get Item(ByVal index As Long) As Variant
    If IsObject(elementData(index)) Then
        Set Item = elementData(index)
        Exit Property
    Else: Item = elementData(index)
    End If
End Property
Public Property Let Item(ByVal index As Long, ByVal value As Variant)
    checkIndex (index)
    'must be nonnegative and less than size
    If IsObject(value) Then
        Set elementData(index) = value
    Else: elementData(index) = value
    End If
End Property

'END STANDARD METHODS


'BEGIN CUSTOM METHODS

Public Sub Remove(ByVal objElement As Variant)
    'removes the first occurrence of the given objElement in the VBArrayList

    Dim i As Long
    For i = 0 To size - 1
        If (elementData(i) = objElement) Then
            Call Me.removeAt(i) ' will decrement size
            Exit For
        End If
    Next i
End Sub

Public Sub RemoveAllOccurencesOf(ByVal objElement As Variant)
    Dim changes As Long
    changes = 0
    Dim i As Long
    For i = 0 To size - 1
        If (elementData(i - changes) = objElement) Then
            Call Me.removeAt(i - changes) ' will decrement size
            changes = changes + 1
        End If
    Next i
    Call trimToSize
End Sub

Public Sub RemoveRange(ByVal StartingIndex As Long, ByVal EndingIndex As Long)
'startindex= first element to remove index, endingindex=final element to remove
'TODO: what if startindex > endindex?
    checkIndex (StartingIndex)
    checkIndex (EndingIndex)

    Dim oldsize As Long
    oldsize = size
    'get all the elements to the right of the range (if there are any elements to the right)
    If EndingIndex < size - 1 Then
        Dim temp() As Variant
        temp = Me.getArrayEls(EndingIndex + 1, size - 1)
        Call arrayCopy(temp, 0, elementData, StartingIndex, UBound(temp) + 1)
    End If
    
    'update the size
    size = size - (EndingIndex - StartingIndex + 1)
    
    'set everything >= the new size =empty
    Dim i As Long
    For i = size To oldsize - 1
        elementData(i) = Empty
    Next i
    
End Sub
Public Function Contains(ByVal Element As Variant) As Boolean
    Dim result As Boolean
    result = False
    
    Dim e As Variant
    For Each e In elementData
        If e = Element Then
            result = True
            Exit For
        End If
    Next e
    
    Contains = result
End Function

Public Function indexOf(ByVal Element As Variant) As Long
'Searches for the specified Object and returns the zero-based index of
'the first occurrence within the entire ArrayList.
    Dim result As Long
    result = -1
    
    Dim index As Long
    index = 0
    
    Dim e As Variant
    For Each e In elementData
        If e = Element Then
            result = index
            Exit For
        End If
        index = index + 1
    Next e

    indexOf = result
End Function

Public Function LastIndexOf(ByVal Element As Variant) As Long
'Searches for the specified Object and returns the
'zero-based index of the last occurrence within the entire ArrayList.

    Dim result As Long
    result = -1
    
    Dim i As Long
    For i = size - 1 To 0 Step -1
        If elementData(i) = Element Then
            result = i
            Exit For
        End If
    Next i

    LastIndexOf = result
End Function

Public Sub Clear()
    ReDim elementData(9)
    actualCap = 10
    size = 0
End Sub

Private Sub checkIndex(ByVal index As Long)
    If (index >= size) Or (index < 0) Then
        Call Err.Raise(index, Description:="The index specified is out of bounds")
    End If
End Sub

Public Sub Swap(ByVal Index1 As Long, ByVal Index2 As Long)
    Dim temp As Variant
    checkIndex (Index1)
    checkIndex (Index2)
    
    If Not IsObject(elementData(Index2)) Then
       temp = elementData(Index2)
    Else: Set temp = elementData(Index2)
    End If
    
    If Not IsObject(elementData(Index1)) Then
       elementData(Index2) = elementData(Index1)
    Else
        Set elementData(Index2) = elementData(Index1)
    End If
    
    If Not IsObject(temp) Then
        elementData(Index1) = temp
    Else
        Set elementData(Index1) = temp
    End If
End Sub

Public Sub Reverse()
    If size > 1 Then
    
        Dim hiIndex As Long
        hiIndex = size - 1
        Dim loIndex As Long
        loIndex = 0
        
        Do While (hiIndex > loIndex)
            Call Swap(loIndex, hiIndex)
            hiIndex = hiIndex - 1
            loIndex = loIndex + 1
        Loop
    End If
End Sub

Public Sub Shuffle()
    'uses Fisher-Yates algo
    
    Dim i As Long
    Dim randomNbr As Long

    For i = size - 1 To 1 Step -1
        Randomize
        randomNbr = Int((i + 1) * Rnd) 'random integer with 0 <= rndnbr <= i, uniformly distributed
        Call Swap(randomNbr, i)
    Next i
End Sub

Public Function GetDistinctValues() As ArrayList
    Dim distinctVals As New ArrayList
    
    Dim e As Variant
    For Each e In elementData
        If Not distinctVals.Contains(e) Then
            distinctVals.Add e
        End If
    Next e
    
    Set GetDistinctValues = distinctVals
End Function

Public Function GetRange(ByVal StartingIndex As Long, ByVal TotalElementsToGet As Long) _
As ArrayList
'Returns an ArrayList which represents a subset of the elements in this ArrayList.
'Index: The 0-based array index at which the range starts.
'Count: The number of elements in the range to get.
    
    Dim newAL As ArrayList
    Set newAL = New ArrayList
        
    If TotalElementsToGet > 0 Then
        Dim i As Long
        If TotalElementsToGet > 9 Then
        newAL.Capacity = TotalElementsToGet
        Else: newAL.Capacity = 10
        End If
        For i = StartingIndex To (StartingIndex + TotalElementsToGet - 1)
            newAL.Add elementData(i)
        Next i
    End If
    
    Set GetRange = newAL
End Function

Public Sub arrayCopy(array1() As Variant, ByVal startingIndex1 As Long, array2() As Variant, _
  startingIndex2 As Long, ByVal TotalElements As Long)
    On Error Resume Next
    'copies from arr1, starting at stin1, to arr2, starting at stin2, totalelements.
    'both arrays must be declared using syntax: dim array1(<number>) as <datatype> or redim array1(<number>)
    
    'ensure arr2 has at least TE els
    If UBound(array2) < TotalElements - 1 Then
       ReDim Preserve array2(TotalElements - 1)
    End If
    
    Dim i As Long
    Dim j As Long
    j = startingIndex2
    For i = startingIndex1 To startingIndex1 + TotalElements - 1
        If Not IsObject(array1(i)) Then
            array2(j) = array1(i)
        Else: Set array2(j) = array1(i)
        End If
        j = j + 1
    Next i
End Sub

Public Sub Sort()
    'use quicksort algo
    
    'optionally un-comment this if your arraylist may contain both objects and primitives
    'If Me.ContainsObjects() Then
    '    MsgBox "This VBArrayList contains at least 1 object.  Quicksort only works on alphanumeric values."
     '   Exit Sub
    'Else
        Call QuickSort2
    'End If
End Sub

Private Sub QuickSort2( _
    Optional intLeft As Long = -2, _
    Optional intRight As Long = -2)
    
    Dim i As Long
    Dim j As Long
    Dim varTestVal As Variant
    Dim intMid As Long
    If intLeft = -2 Then intLeft = 0
    If intRight = -2 Then intRight = size - 1
    If intLeft < intRight Then
        intMid = (intLeft + intRight) \ 2
        varTestVal = elementData(intMid)
        i = intLeft
        j = intRight
        Do
            Do While elementData(i) < varTestVal
                i = i + 1
            Loop
            Do While elementData(j) > varTestVal
                j = j - 1
            Loop
            If i <= j Then
                Call Me.Swap(i, j)
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        
        ' To optimize the sort, always sort the
        ' smallest segment first.
        If j <= intMid Then
            Call QuickSort2(intLeft, j)
            Call QuickSort2(i, intRight)
        Else
            Call QuickSort2(i, intRight)
            Call QuickSort2(intLeft, j)
        End If
    End If
End Sub

Public Function ContainsObjects() As Boolean
'determines whether the arraylist contains any objects
    Dim result As Boolean
    result = False
    
    Dim e As Variant
    For Each e In elementData
        If IsObject(e) Then
            result = True
            Exit For
        End If
    Next e
    
    ContainsObjects = result
End Function

Public Function getArrayEls(ByVal StartingIndex As Long, ByVal EndingIndex As Long) As Variant()
    Dim els() As Variant
    ReDim els(EndingIndex - StartingIndex)
    
    Dim i As Long
    Dim j As Long
    j = 0
    
    If StartingIndex <= EndingIndex Then
        For i = StartingIndex To EndingIndex
            If Not IsObject(elementData(i)) Then
                els(j) = elementData(i)
            Else
                Set els(j) = elementData(i)
            End If
            j = j + 1
        Next i
    Else
        For i = StartingIndex To EndingIndex Step -1
            If Not IsObject(elementData(i)) Then
                els(j) = elementData(i)
            Else
                Set els(j) = elementData(i)
            End If
            j = j + 1
        Next i
    End If
    
    getArrayEls = els
End Function

Public Function ToCollection() As Collection
    Dim coll As New Collection
    Dim i As Long
    For i = 0 To size - 1
        coll.Add elementData(i)
    Next i
    Set ToCollection = coll
End Function

Public Function ToArray() As Variant()
     ToArray = elementData
End Function

Public Sub IntakeArray(yourArray() As Variant)
    'array must be a variant array
    elementData = yourArray
    actualCap = Me.Capacity
    size = Me.Length
End Sub

Public Sub IntakeCollection(ByVal yourCollection As Collection)
    'completely replaces anything in elementData with the elements of a collection
    'do not use parentheses around the argument
    
    ReDim elementData(yourCollection.Count - 1)
    
    Dim i As Long
    For i = 0 To UBound(elementData)
        If IsObject(yourCollection.Item(i + 1)) Then
            Set elementData(i) = yourCollection.Item(i + 1)
        Else: elementData(i) = yourCollection.Item(i + 1)
        End If
    Next i
    
    actualCap = Me.Capacity
    size = Me.Length
End Sub

