Attribute VB_Name = "modRecursivePermutations"
Option Explicit
Option Base 0

Private strString As String
Private intPositionArrayPointer As Integer
Private intPositionArray() As Integer
Private strPermutation As String
Public bolCancel As Boolean

Public Sub RecursivePermutations(strPassedString As String)
    'Set the string to permutate.
    strString = strPassedString
    
    'Initialise the position array pointer.
    intPositionArrayPointer = -1
    
    'Initialise the position array.
    ReDim intPositionArray(Len(strString) - 1)
    
    bolCancel = False

    'Calculate the possible permutations.
    Call Permutations(0)
        
    If bolCancel Then Call MsgBox("Cancelled", vbInformation, "dcaenlcle")
End Sub

Private Sub Permutations(intElement As Integer)
    'Local variables.
    Dim i As Integer

    DoEvents
    If bolCancel Then Exit Sub

    'Increase the position pointer.
    intPositionArrayPointer = intPositionArrayPointer + 1
    
    'Assign the position to the position array.
    intPositionArray(intElement) = intPositionArrayPointer
    
    'See if the position pointer is at the end of the array.
    If intPositionArrayPointer = Len(strString) Then
        'Set the permutation.
        strPermutation = ""
        For i = 0 To UBound(intPositionArray)
            strPermutation = strPermutation & _
                Mid$(strString, intPositionArray(i), 1)
        Next i
    
        frmMain.lstCombs.AddItem strPermutation
        frmMain.UpdateProgress
    Else
        For i = 0 To Len(strString) - 1
            If intPositionArray(i) = 0 Then Call Permutations(i)
        Next i
    End If
    
    'Reset the position array element.
    intPositionArrayPointer = intPositionArrayPointer - 1
    intPositionArray(intElement) = 0
End Sub
