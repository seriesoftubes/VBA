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
'   i.e., you must say yourArrayList.Item(0) = foo
'   or foo = yourArrayList.Item(1)
'   You *cannot* use the syntax yourArrayList(0) or yourArrayList[0]
'
'The ArrayList object itself cannot be processed using a for-each loop.
'To do that, iterate over yourArrayList.Items()

Private m_elements() As Variant
Private m_size As Long
Private m_capacity As Long

Private Sub Class_Initialize()
    ReDim m_elements(9)
    m_size = 0
    m_capacity = 10
End Sub

Public Property Get Capacity() As Long
  'all capacity in the array, including unused space
    Capacity = UBound(m_elements) + 1
End Property
Public Property Let Capacity(ByVal TotalCapacity As Long)
    ReDim Preserve m_elements(TotalCapacity - 1)
    m_capacity = TotalCapacity
End Property

Public Function Length() As Long
'includes only used elements in the array
    Length = m_size
End Function

Private Sub trimToSize()
    'If capacity is large and length < 50% of capacity,
    'trim total capacity to: (number of used elements * 1.5)
    
    If m_capacity > 99 Then
        If (m_size < (m_capacity / 2)) Then
            Dim newUBound As Long
            newUBound = Conversion.CLng(m_size * 1.5)
            If newUBound < 9 Then 'need at least 10 els
                newUBound = 9
            End If
            ReDim Preserve m_elements(newUBound)
            m_capacity = newUBound + 1
        End If
    End If
End Sub

Private Sub ensureCapacity(ByVal minCapacity As Long)
    If m_capacity < minCapacity Then
        Dim newUBound As Long
        newUBound = Conversion.CLng(m_capacity * 1.5)
        ReDim Preserve m_elements(newUBound)
        m_capacity = newUBound + 1
    End If
End Sub

Public Function isEmpty() As Boolean
    isEmpty = (m_size = 0)
End Function

Public Sub Add(Item As Variant, Optional ByVal Before As Long = -1)
'Inserts the specified element at the specified position in this
'list. Shifts the element currently at that position (if any) and
'any subsequent elements to the right (adds one to their indices).
    Call ensureCapacity(m_size + 1)
    
    'shift everything to the right of Before by 1
    If (Before > -1) Then
        checkIndex (Before)
        Dim temp() As Variant
        ReDim temp(m_size)
        
        Call arrayCopy(m_elements(), Before, temp(), 0, m_size - Before)
        Call arrayCopy(temp(), 0, m_elements(), Before + 1, m_size - Before)
        If Not IsObject(Item) Then
           m_elements(Before) = Item
        Else
            Set m_elements(Before) = Item
        End If
    Else ' no "Before" param
        If Not IsObject(Item) Then
            m_elements(m_size) = Item
        Else
            Set m_elements(m_size) = Item
        End If
    End If
    m_size = m_size + 1
End Sub

Sub removeAt(ByVal index As Long)
    checkIndex (index)
    
    If index < m_size - 1 Then
        Dim i As Integer
        For i = index To m_size - 1
            If Not IsObject(m_elements(i + 1)) Then
                m_elements(i) = m_elements(i + 1)
            Else
                Set m_elements(i) = m_elements(i + 1)
            End If
        Next i
        m_elements(m_size - 1) = Empty
    ElseIf index = m_size - 1 Then
        m_elements(m_size - 1) = Empty
    End If
    
    m_size = m_size - 1
    Call trimToSize
End Sub

Public Property Get Item(ByVal index As Long) As Variant
    If IsObject(m_elements(index)) Then
        Set Item = m_elements(index)
        Exit Property
    Else
        Item = m_elements(index)
    End If
End Property
Public Property Let Item(ByVal index As Long, ByVal value As Variant)
    checkIndex (index)
    If IsObject(value) Then
        Set m_elements(index) = value
    Else
        m_elements(index) = value
    End If
End Property

Public Sub Remove(ByVal objElement As Variant)
    'Remove the first occurrence of the given objElement
    Dim i As Long
    For i = 0 To m_size - 1
        If (m_elements(i) = objElement) Then
            Call Me.removeAt(i)
            Exit For
        End If
    Next i
End Sub

Public Sub RemoveAll(ByVal objElement As Variant)
    'Remove all occurrences of objElement
    Dim changes As Long
    changes = 0
    Dim i As Long
    For i = 0 To m_size - 1
        If (m_elements(i - changes) = objElement) Then
            Call Me.removeAt(i - changes) ' will decrement m_size
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

    Dim oldm_size As Long
    oldm_size = m_size
    'get all the elements to the right of the range (if there are any elements to the right)
    If EndingIndex < m_size - 1 Then
        Dim temp() As Variant
        temp = Me.Items(EndingIndex + 1, m_size - 1)
        Call arrayCopy(temp, 0, m_elements, StartingIndex, UBound(temp) + 1)
    End If
    
    m_size = m_size - (EndingIndex - StartingIndex + 1)
    Dim i As Long
    For i = m_size To oldm_size - 1
        m_elements(i) = Empty
    Next i
End Sub

Public Function Contains(ByVal Element As Variant) As Boolean
    Dim result As Boolean
    result = False
    Dim e As Variant
    For Each e In m_elements
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
'Returns -1 if the Element was not found
    
    Dim result As Long
    result = -1
    Dim index As Long
    index = 0
    
    Dim e As Variant
    For Each e In m_elements
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
'Returns -1 if not found

    Dim result As Long
    result = -1
    
    Dim i As Long
    For i = m_size - 1 To 0 Step -1
        If m_elements(i) = Element Then
            result = i
            Exit For
        End If
    Next i
    LastIndexOf = result
End Function

Public Sub Clear()
    ReDim m_elements(9)
    m_capacity = 10
    m_size = 0
End Sub

Private Sub checkIndex(ByVal index As Long)
    If (index >= m_size) Or (index < 0) Then
        Call Err.Raise(index, Description:="The index specified is out of bounds")
    End If
End Sub

Public Sub Swap(ByVal Index1 As Long, ByVal Index2 As Long)
    Dim temp As Variant
    checkIndex (Index1)
    checkIndex (Index2)
    
    If Not IsObject(m_elements(Index2)) Then
       temp = m_elements(Index2)
    Else: Set temp = m_elements(Index2)
    End If
    
    If Not IsObject(m_elements(Index1)) Then
       m_elements(Index2) = m_elements(Index1)
    Else
        Set m_elements(Index2) = m_elements(Index1)
    End If
    
    If Not IsObject(temp) Then
        m_elements(Index1) = temp
    Else
        Set m_elements(Index1) = temp
    End If
End Sub

Public Sub Reverse()
    If m_size > 1 Then
        Dim hiIndex As Long
        hiIndex = m_size - 1
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
    For i = m_size - 1 To 1 Step -1
        Randomize
        'random integer with 0 <= rndnbr <= i, uniformly distributed
        randomNbr = Int((i + 1) * Rnd)
        Call Swap(randomNbr, i)
    Next i
End Sub

Public Function GetDistinctValues() As ArrayList
    Dim distinctVals As New ArrayList
    Dim e As Variant
    For Each e In m_elements
        If Not distinctVals.Contains(e) Then
            distinctVals.Add e
        End If
    Next e
    Set GetDistinctValues = distinctVals
End Function

Public Function GetRange(ByVal StartingIndex As Long, ByVal TotalElementsToGet As Long) _
As ArrayList
'Returns a subset of the elements in this ArrayList.
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
            newAL.Add m_elements(i)
        Next i
    End If
    Set GetRange = newAL
End Function

Public Sub arrayCopy(array1() As Variant, ByVal startingIndex1 As Long, array2() As Variant, _
  startingIndex2 As Long, ByVal TotalElements As Long)
    On Error Resume Next
    'copies from arr1, starting at stin1, to arr2, starting at stin2, TotalElements.
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
    If Me.ContainsObjects() Then
        MsgBox "This VBArrayList contains at least 1 object.  Quicksort only works on alphanumeric values."
        Exit Sub
    Else
        Call QuickSort
    End If
End Sub

Private Sub QuickSort(Optional intLeft As Long = -2, _
    Optional intRight As Long = -2)
    
    Dim i As Long
    Dim j As Long
    Dim varTestVal As Variant
    Dim intMid As Long
    If intLeft = -2 Then intLeft = 0
    If intRight = -2 Then intRight = m_size - 1
    If intLeft < intRight Then
        intMid = (intLeft + intRight) \ 2
        varTestVal = m_elements(intMid)
        i = intLeft
        j = intRight
        Do
            Do While m_elements(i) < varTestVal
                i = i + 1
            Loop
            Do While m_elements(j) > varTestVal
                j = j - 1
            Loop
            If i <= j Then
                Call Me.Swap(i, j)
                i = i + 1
                j = j - 1
            End If
        Loop Until i > j
        
        If j <= intMid Then
            Call QuickSort(intLeft, j)
            Call QuickSort(i, intRight)
        Else
            Call QuickSort(i, intRight)
            Call QuickSort(intLeft, j)
        End If
    End If
End Sub

Public Function ContainsObjects() As Boolean
    Dim result As Boolean
    result = False
    Dim e As Variant
    For Each e In m_elements
        If IsObject(e) Then
            result = True
            Exit For
        End If
    Next e
    ContainsObjects = result
End Function

Public Function Items(Optional ByVal StartingIndex As Long = 0, Optional ByVal EndingIndex As Long = -1) As Variant()
    If EndingIndex = -1 Then EndingIndex = m_size - 1
    
    Dim els() As Variant
    ReDim els(EndingIndex - StartingIndex)
    
    Dim i As Long
    Dim j As Long
    j = 0
    
    If StartingIndex <= EndingIndex Then
        For i = StartingIndex To EndingIndex
            If Not IsObject(m_elements(i)) Then
                els(j) = m_elements(i)
            Else
                Set els(j) = m_elements(i)
            End If
            j = j + 1
        Next i
    Else
        For i = StartingIndex To EndingIndex Step -1
            If Not IsObject(m_elements(i)) Then
                els(j) = m_elements(i)
            Else
                Set els(j) = m_elements(i)
            End If
            j = j + 1
        Next i
    End If
    
    Items = els
End Function

Public Function ToCollection() As Collection
    Dim coll As New Collection
    Dim i As Long
    For i = 0 To m_size - 1
        coll.Add m_elements(i)
    Next i
    Set ToCollection = coll
End Function

Public Function ToArray() As Variant()
     ToArray = m_elements
End Function

Public Sub IntakeArray(yourArray() As Variant)
    'array must be a variant array
    m_elements = yourArray
    m_capacity = Me.Capacity
    m_size = Me.Length
End Sub

Public Sub IntakeCollection(ByVal yourCollection As Collection)
    'completely replaces anything in m_elements with the elements of a collection
    'do not use parentheses around the argument
    
    ReDim m_elements(yourCollection.Count - 1)
    Dim i As Long
    For i = 0 To UBound(m_elements)
        If IsObject(yourCollection.Item(i + 1)) Then
            Set m_elements(i) = yourCollection.Item(i + 1)
        Else: m_elements(i) = yourCollection.Item(i + 1)
        End If
    Next i
    m_capacity = Me.Capacity
    m_size = Me.Length
End Sub
