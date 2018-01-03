Attribute VB_Name = "usefulFunctions"
Public Const Pi As Double = 3.14159265359

Public Function Alpha_Sort_Array(varArray As Variant)
'uses bubble sort algorithm. may change to another at some point.
Dim blnFlipped As Boolean
Dim varthing As Variant, i As Integer, temp As Variant
'first convert all to ascii text
For Each varthing In varArray
    varthing = CStr(varthing)
Next varthing
Do
    blnFlipped = False
    For i = LBound(varArray) To UBound(varArray) - 1
        If varArray(i) > varArray(i + 1) Then
            temp = varArray(i + 1)
            varArray(i + 1) = varArray(i)
            varArray(i) = temp
            blnFlipped = True
        End If
    Next i
Loop While blnFlipped

Alpha_Sort_Array = varArray
End Function

Public Function Numeric_Array_Sort(ByVal A As Variant)
'bubble sort an array of numbers

Dim blnFlipped As Boolean, sizeA As Integer
Dim varthing As Variant, i As Integer, temp As Double

sizeA = UBound(A) - LBound(A)

'force all array members to double type. If there's a non-numeric
'member, it'll throw a (Trappable) error
i = LBound(A)
For Each varthing In A
    varthing = CDbl(varthing)
    i = i + 1
Next varthing
Do
    blnFlipped = False
    For i = LBound(A) To UBound(A) - 1
        If A(i) > A(i + 1) Then
            temp = A(i + 1)
            A(i + 1) = A(i)
            A(i) = temp
            blnFlipped = True
        End If
    Next i
Loop While blnFlipped
Numeric_Array_Sort = A

        

End Function

Public Function EliminateDuplicates(varArray As Variant)
Dim i As Integer, varItem As Variant, j As Integer
Dim intNoRepeats As Integer, tempArray() As Variant
Dim min As Integer, max As Integer
i = LBound(varArray)
min = i
max = UBound(varArray)
For i = min To max
    varItem = varArray(i)
    'if the positionInArray function returns -999,
    'the current item appears later on in the list
    'so don't put the current item in the new array
    If PositionInArray(varItem, varArray, i + 1, max) = -999 Then
        ReDim Preserve tempArray(intNoRepeats)
        tempArray(intNoRepeats) = varItem
        intNoRepeats = intNoRepeats + 1
    End If
    
Next i

EliminateDuplicates = tempArray
End Function

Public Function ArrSum(varNumbers As Variant) As Double
'returns the sum of the elements of a variant array
'use by making an array of type variant (e.g. myNumbers) containing at least one
'numerical value and use syntax
'mySum = ArrSum(myNumbers)
'where mySum is a Double

Dim RunningTotal As Variant
Dim nItems As Integer, N As Integer
nItems = UBound(varNumbers)
For N = LBound(varNumbers) To nItems
'ignore non-numeric array entries
    If Not TypeName(varNumbers(N)) = "String" Or _
        TypeName(varNumbers(N)) = "Date" Or _
        TypeName(varNumbers(N)) = "Object" Then
        
        RunningTotal = RunningTotal + varNumbers(N)
    End If
Next N
ArrSum = RunningTotal
End Function

Public Function SumOf(ParamArray A() As Variant) As Double
'same as ArrSum, but takes a comma separated list of either numbers or
'numeric variables.
'will ignore non-numeric arguements
'e.g. mySum = SumOf(1, 17.256, myRandomNumber)

Dim dblSumA As Double, varItem As Variant
For Each varItem In A
    If IsNumeric(varItem) Then
        dblSumA = dblSumA + CDbl(varItem)
    End If
Next varItem
SumOf = dblSumA
End Function


Public Function Largest(ByVal varNumbers As Variant) As Double
'returns the numerically largest element in an n-dimensioned array
Dim BiggestSoFar As Double
Dim varElement As Variant

For Each varElement In varNumbers
    If varElement > BiggestSoFar Then
        BiggestSoFar = varElement
    End If
Next varElement
Largest = BiggestSoFar
End Function

Public Function SubscriptOfLargest(ByVal varNumbers As Variant) As Integer
'returns the subscript of the largest element in an array.
'if there are more than one elements sharing the same largest value, it will return the
'subscript from the first one.

Dim BiggestSoFar As Double
Dim nBiggest As Integer
Dim intItem As Integer

For intItem = LBound(varNumbers) To UBound(varNumbers)
    If varNumbers(intItem) > BiggestSoFar Then
        BiggestSoFar = varNumbers(intItem)
        nBiggest = intItem
    End If
Next intItem
SubscriptOfLargest = nBiggest
End Function

Public Function GaussianDeviate(StdDev As Double) As Double
'returns a zero-mean, StdDev standard deviation gaussian deviate.
'adapted for VB from Numerical recipies in C (1988); CUP, by Me.
Static blnIFlag As Boolean
Static gset As Double ' will hold the spare value.
Dim fac As Double, r As Double, v1 As Double, v2 As Double
If blnIFlag = False Then
    Randomize
    Do
        v1 = (Rnd * 2) - 1
        v2 = (Rnd * 2) - 1
        r = (v1 * v1) + (v2 * v2)
    Loop While r >= 1
     fac = Sqr((-2) * Log(r) / r)
     gset = StdDev * v1 * fac
     blnIFlag = True
     GaussianDeviate = StdDev * v2 * fac
Else
    blnIFlag = False
    GaussianDeviate = gset
End If
End Function

Public Function DotProduct(ByVal A As Variant, ByVal B As Variant) As Variant
'returns the dot product of two array vectors A and B
Dim tempArray() As Variant, intSize As Integer, intMember As Integer, intStart As Integer
'first check they are both arrays of the same size.
If (IsArray(A) And IsArray(B)) And (UBound(A) = UBound(B)) Then
    intSize = UBound(A)
    intStart = LBound(A)
    ReDim tempArray(intSize)
    For intMember = intStart To intSize
        tempArray(intMember) = A(intMember) * B(intMember)
    Next intMember
    DotProduct = tempArray
    Exit Function
Else
    Err.Raise Number:=1247, _
    Description:="Cannot do a dot product calculation on these"
End If
End Function


Public Function ShuffleArray(A As Variant, Optional No_Three_In_A_Row As Boolean) As Variant
'shuffle array takes an array of variants, and shuffles them into a random order
'the optional parameter No_Three_In_A_Row is used when you don't want any
'runs of three or more items in the same order as in the original array.
'e.g. myNewArray = ShuffleArray(myOldArray, True) will return a reordered array with no
'three-in-a-rows from the myOldArray array.
'myNewArray = ShuffleArray(myOldArray) just shuffles the old array with no reference to the
'order of oldArray.

Dim N As Integer, L As Integer, intRnd As Integer, i As Integer, tempA() As Variant
Dim blnTemp() As Boolean, blnThreeInARow As Boolean, blnSubscriptErr As Boolean
On Error GoTo errTrap

L = LBound(A)
N = UBound(A)
ReDim tempA(N - L)
ReDim blnTemp(L To N)

For i = 0 To N - L
    Do
        
        DoEvents
        blnThreeInARow = False
        blnSubscriptErr = False
        intRnd = Int(Rnd * (N) + L)
        If (tempA(i - 1) = A(intRnd)) And (tempA(i - 2) = A(intRnd)) Then
            If No_Three_In_A_Row = True And blnSubscriptErr = False Then
                blnThreeInARow = True
            End If
        End If
    Loop While (Not blnTemp(intRnd) = False) Or (blnThreeInARow = True)
    tempA(i) = A(intRnd)
    blnTemp(intRnd) = True
Next i
ShuffleArray = tempA
Exit Function

errTrap:
If Err.Number = 9 Then
    blnThreeInARow = False
    blnSubscriptErr = True
    Resume Next
Else
    MsgBox Err.Description, vbOKOnly, "error" & Err.Number
End If

End Function


Public Function PickNfromX(ByVal N As Integer, varThings As Variant) As Variant
'picks n items from a 1-based array of variants. size is aritrary but >= N

Dim tempList() As Variant, intRnd As Integer, i As Integer
Dim blnUsed() As Boolean

ReDim tempList(1 To N)
ReDim blnUsed(1 To UBound(varThings))

For i = 1 To N
    Do
        intRnd = Int(Rnd * UBound(varThings)) + 1
    Loop While blnUsed(intRnd) = True
    blnUsed(intRnd) = True
    tempList(i) = varThings(intRnd)
Next i
PickNfromX = tempList
End Function

Public Function ArrayMean(ByVal A As Variant) As Double
'returns the arithmetic mean of an array of numeric values
'any non-numeric variants give either a zero-entry or an error
Dim N As Integer, i As Integer

If IsArray(A) Then
    N = UBound(A) - LBound(A) + 1
    ArrayMean = ArrSum(A) / N
Else
    ArrayMean = A
End If
End Function

Public Function ArrayStdDev(ByVal A As Variant) As Double
'returns the standard deviation based on a sample of an array of numeric values, A
Dim N As Integer, i As Integer, SqrA As Variant

N = UBound(A) - LBound(A) + 1
 SqrA = ArrSqr(A)
 ArrayStdDev = Sqr((N * ArrSum(SqrA) - (ArrSum(A) * ArrSum(A))) / (N * (N - 1)))

End Function

Public Function ArrSqr(ByVal A As Variant) As Variant
'returns an array of numeric items in which each array member is squared
Dim i As Integer
Dim dblX As Double
Dim tempArray() As Double
ReDim tempArray(LBound(A) To UBound(A))

On Error GoTo NonNumericHandler

For i = LBound(A) To UBound(A)
    dblX = CDbl(A(i))
    tempArray(i) = dblX * dblX
Next i
ArrSqr = tempArray
Exit Function

NonNumericHandler:
    If Err.Number = 13 Then
        'if this was a type mismatch error then
        tempArray(i) = 0
        Resume Next
    End If
End Function

Public Function ArraysAreIdentical(ByVal A As Variant, ByVal B As Variant) As Boolean
'checks whether array A and B are identical.
' returns true if they are, and false otherwise
Dim La As Integer, Lb As Integer
Dim Na As Integer, Nb As Integer
Dim i As Integer, j As Integer

La = LBound(A)
Lb = LBound(B)
Na = UBound(A)
Nb = UBound(B)

If Na - La <> Nb - Lb Then
    ArraysAreIdentical = False
    Exit Function
End If

For i = La To Na
    If Not (A(i) = B(Lb + j)) Then
        ArraysAreIdentical = False
        Exit Function
    End If
    j = j + 1
Next i
'if it gets to here, then they are identical
ArraysAreIdentical = True
End Function

Public Function ArraysSharePosition(ByVal A As Variant, ByVal B As Variant) As Boolean
'takes two variant arrays as args. returns true if they
'share an item in the same position relative to the lowest-value array
'subscript

Dim La As Integer, Lb As Integer
Dim Na As Integer, Nb As Integer
Dim i As Integer, j As Integer
Dim intSize As Integer

La = LBound(A)
Lb = LBound(B)
Na = UBound(A)
Nb = UBound(B)

'find the size of the smaller array
intSize = SmallestN(Na - La, Nb - Lb)

' go around looking for shared positions relative to the start of the array
For i = 0 To intSize
    If A(La + i) = B(Lb + i) Then
        ArraysSharePosition = True
        Exit Function
    End If
Next i
ArraysSharePosition = False
End Function

Public Function LargestN(ParamArray Nums() As Variant) As Double
'takes a list of numerical values and returns the largest of them
'e.g. myLargest = LargestN(0.001, 0.000001, 5) will return 5

Dim dblLargestSoFar As Double
Dim varItem As Variant
For Each varItem In Nums
    If CDbl(varItem) > dblLargestSoFar Then
        dblLargestSoFar = CDbl(varItem)
    End If
Next varItem
LargestN = dblLargestSoFar
End Function

Public Function SmallestN(ParamArray Nums() As Variant) As Double
'same as LargestN, but with one obvious difference
Dim dblSmallestSoFar As Double
Dim varthing As Variant
dblSmallestSoFar = Nums(LBound(Nums))
For Each varthing In Nums
    If CDbl(varthing) < dblSmallestSoFar Then
        dblSmallestSoFar = CDbl(dblSmallestSoFar)
    End If
Next varthing
SmallestN = dblSmallestSoFar
End Function

Public Function IsEven(ByVal Number As Integer) As Boolean
'returns true if Number is even
If Number Mod 2 = 0 Then IsEven = True
End Function

Public Function AlphabetPosition(ByVal Letter As String) As Integer
'returns the position within the alphabet of a letter
Dim strAlphabet As String, i As Integer
strAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
For i = 1 To 26
    If Mid(strAlphabet, i, 1) = Letter Then
        AlphabetPosition = i
        Exit For
    End If
Next i
End Function

Public Function LetterAtPosition(ByVal Position As Byte) As String
Dim strAlphabet As String
strAlphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
LetterAtPosition = Mid(strAlphabet, Position, 1)
End Function

Public Function IsInArray(ByVal Item As Variant, A As Variant) As Boolean
Dim arrayItem As Variant
For Each arrayItem In A
    If Item = arrayItem Then
        IsInArray = True
        Exit Function
    End If
Next arrayItem
End Function

Public Function PositionInArray(ByVal Item As Variant, ByVal A As Variant, Optional ByVal min As Integer, Optional ByVal max As Integer) As Integer
Dim arrayItem As Variant, i As Integer
PositionInArray = -999
If min = 0 And max = 0 Then
    i = LBound(A)
    For Each arrayItem In A
        If Item = arrayItem Then
            PositionInArray = i
            Exit Function
        End If
        i = i + 1
    Next arrayItem
Else
    For i = min To max
        arrayItem = A(i)
        If Item = arrayItem Then
            PositionInArray = i
            Exit Function
        End If
    Next i
End If
End Function


Public Function Factorial(ByVal N As Integer) As Long
'returns the factorial of an integer.
Dim i As Integer, lngResult As Long
lngResult = 1
For i = 1 To N
    lngResult = lngResult * i
Next i
Factorial = lngResult
End Function

Public Function ArrayAContainsArrayB(ByVal A As Variant, ByVal B As Variant) As Integer
'returns the start point of an array B contained within array A. if A doesn't contain B,
' -999 is returned
'similar to InStr function but for variant arrays not string vars
Dim intPosA As Integer, intPosB As Integer
Dim intBStart As Integer, BSize As Integer
Dim chunkSize As Integer

BSize = (UBound(B) - LBound(B))
intBStart = LBound(B)
For intPosA = LBound(A) To (UBound(A) - BSize)
    chunkSize = 0
    For intPosB = 0 To BSize
        If A(intPosA + intPosB) = B(intPosB) Then
            'carry on
            chunkSize = chunkSize + 1
        Else
            
            Exit For
        End If
     
    Next intPosB
    If chunkSize = BSize + 1 Then
        ArrayAContainsArrayB = intPosA
        Exit Function
    End If
Next intPosA
ArrayAContainsArrayB = -999
End Function
