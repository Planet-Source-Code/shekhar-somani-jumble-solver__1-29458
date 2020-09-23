Attribute VB_Name = "modLexicographicPermutations"
Option Explicit
Option Base 0

Public bolCancel As Boolean

Public Sub LexicographicPermutations(strString As String)
    'Local variables.
    Dim i As Integer    'Loop variable.
    Dim j As Integer    'Loop variable.
    Dim intTemp As Integer  'Temp swaping variable.
    Dim intPositionArray() As Integer   'Poisition array.
    Dim strPermutation As String        'Permutation variable.
    
    'Initialise the position array.
    ReDim intPositionArray(Len(strString) - 1)
    For i = 1 To Len(strString)
        intPositionArray(i - 1) = i
    Next i

    'Get each permutation.
    Do
        DoEvents
        If bolCancel Then GoTo Finish

        i = UBound(intPositionArray)
    
        'Loop through the permutation position array and reorder it.
        Do While intPositionArray(i - 1) >= intPositionArray(i)
            i = i - 1
            If i = 0 Then GoTo Finish
        Loop
    
        j = UBound(intPositionArray) + 1
    
        Do While intPositionArray(j - 1) <= intPositionArray(i - 1)
            j = j - 1
        Loop
    
        'Swap (i-1) and (j-1).
        intTemp = intPositionArray(i - 1)
        intPositionArray(i - 1) = intPositionArray(j - 1)
        intPositionArray(j - 1) = intTemp
    
        i = i + 1
        j = UBound(intPositionArray) + 1
        
        Do While i < j
            'Swap (i-1) and (j-1).
            intTemp = intPositionArray(i - 1)
            intPositionArray(i - 1) = intPositionArray(j - 1)
            intPositionArray(j - 1) = intTemp
            i = i + 1
            j = j - 1
        Loop
    
        'Set the permutation.
        strPermutation = ""
        For i = 0 To UBound(intPositionArray)
            strPermutation = strPermutation & _
                Mid$(strString, intPositionArray(i), 1)
        Next i

        frmMain.Process strPermutation
'        frmMain.UpdateProgress

    Loop
    
Finish:
    'If bolCancel Then MsgBox "Cancelled"
End Sub
