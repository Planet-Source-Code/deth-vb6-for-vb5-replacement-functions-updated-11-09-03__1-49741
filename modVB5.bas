Attribute VB_Name = "modVB5"
Option Explicit

'****************************************************
' Author: Lewis Miller (aka Deth)
'  Email: dethbomb@hotmail.com
'   Date:   11/08/03
'Purpose: Complete vb5 replacement functions for missing intrinsic vb6 functions
'****************************************************
'- Change log -
'Updated: 11/09/03 Bug fix in InstrRev() - Didnt allow for Position to be 1.
'                  Bug fix in InstrRev() - Start wasnt correctly calculated
'                  Bug fix in InstrRev() - Unused variable removed
'                  Bug fix in Split()    - didnt account for delimiters longer than 1
'                  Bug fix in Filter()   - a check for no items found added (see next item)
'                  Filter() now returns original array value if nothing is replaced
'                  Some code optimized in Replace()
'                  Lots of small tweaks here an there
'                  Added lots comments :)
'********************************************************************
'Notes:
'The only function that still uses string concatenation is Replace()
'all the others were written to allocate a buffer of size needed and then
'the values are pushed into the string using the Mid() statement (not function)
'Most generally this results in much faster string manipulation, than the old way
'However you wont notice any speed increase generally unless your using larger
'strings, which is what this module is all about. On large strings the old functions
'found in other similar modules are extremely SLOW, with this module you get consistent
'performance with small or large strings
'
'-some ways performance is increased for each function over old ways-
'in all functions string lengths are stored in long variables instead of repeatedly calling the
'len() function, variable access is much faster then calling a function
'
'Replace()
'this function is a pain to implement with a pre-allocated buffer and still uses the old way, until i
'figure out a way to do it better at least :(
'
' Split()
'by counting the number of needed dimensions in the array before actually dimming it
'you avoid the use of many many redim preserve calls, with one call to redim you have the size needed
'
'Join()
'this function counts the length of the string buffer needed before actually joining the array, this is a
'great example of a function that will work fast with large strings
'
'StrReverse()
'instead of concatenating each character to a string, this pre-allocates the buffer needed then simply
'iterates the string and pushes the characters into the new buffer, very fast compared to older methods
'
'Round()
'this function uses a little known trick to calculate its result, making it probably the fastest pure vb
'round function ever coded, no loops or complicated string functions and coercing needed
'
'InstrRev()
'this function is similar to older methods except it works better like the vb6 version than other similar coded functions
'
'Filter()
'this is probably the first publicly available coded filter() function that properly implements the vb6
'version, again it counts the number of array dims needed before allocating the array, skipping the need to
'call redim preserve repeatedly
'
'********************************************************************

'returns a string with all 'find' items in 'expression' replaced with 'replacement'
Public Function Replace(ByVal Expression As String, _
                        ByVal Find As String, _
               Optional ByVal Replacement As String, _
               Optional ByVal Start As Long = 1, _
               Optional ByVal Count As Long = -1, _
               Optional ByVal Compare As VbCompareMethod) As String

  Dim lExpLength As Long  'expression length
  Dim lPosition  As Long  'position of item
  Dim lFindLen   As Long  'find length
  Dim lRepLen    As Long  'replacement length
  Dim lCount     As Long  'count of items found

    If Find <> Replacement Then 'not equal?
        'store lengths
        lExpLength = Len(Expression)
        lFindLen = Len(Find)
        'make sure a search is needed
        If lExpLength > 0 And lFindLen > 0 Then
            'store replacement length
            lRepLen = Len(Replacement)
            'is start within bounds?
            If Start <= lExpLength Then
                'get the first position
                lPosition = InStr(Start, Expression, Find, Compare)
                'loop thru all positions found
                Do While lPosition > 0
                    'by doing a quick compare we can eliminate a string concat and a function
                    If lPosition = 1 Then
                        Expression = Replacement & Mid$(Expression, lFindLen + 1)
                    Else
                        Expression = Left$(Expression, lPosition - 1) & Replacement & Mid$(Expression, lPosition + lFindLen)
                    End If
                    'find next position
                    lPosition = InStr(lPosition + lRepLen, Expression, Find, Compare)
                    lCount = lCount + 1
                    If lCount = Count Then 'limit?
                        Exit Do
                    End If
                Loop
            End If
        End If
    End If
    Replace = Expression

End Function

'returns a zero based array in a variant, splits the string into an array with all delimiters removed
'if delimiter isnt used, it returns the entire string in array(0), you can specify limit to only return
'number of items you want
Public Function Split(ByVal Expression As String, _
             Optional ByVal Delimiter As String, _
             Optional ByVal Limit As Long = -1, _
             Optional ByVal Compare As VbCompareMethod) As Variant

  Dim lPosition  As Long   'position
  Dim lDelimLen  As Long   'length of delimiter
  Dim strArr()   As String 'temp array
  Dim lExpLen    As Long   'length of expression
  Dim lCount     As Long   'count of array items needed
    
    'store lengths
    lExpLen = Len(Expression)
    lDelimLen = Len(Delimiter)
    'check for proper values
    If (lExpLen > 0) And (lDelimLen > 0) Then
        'find first position
        lPosition = InStr(1, Expression, Delimiter, Compare)
        If lPosition > 0 Then
            'count number of array items needed
            Do While lPosition > 0
                lCount = lCount + 1
                If lCount = Limit Then
                    Exit Do
                End If
                lPosition = InStr(lPosition + lDelimLen, Expression, Delimiter, Compare)
            Loop
            ReDim strArr(lCount) As String
            lCount = 0
            're-loop thru again an actually grab items
            lPosition = InStr(1, Expression, Delimiter, Compare)
            lExpLen = 1
            Do While lPosition > 0
                strArr(lCount) = Left$(Expression, lPosition - lExpLen)
                lExpLen = lPosition + lDelimLen
                'find next item
                lPosition = InStr(lPosition + lDelimLen, Expression, Delimiter, Compare)
                lCount = lCount + 1
                If lCount = (Limit - 1) Then 'reached limit?
                    Exit Do
                End If
            Loop
            'get final item
            strArr(lCount) = Mid$(Expression, lExpLen)
            GoTo Done
        End If
    End If

    ReDim strArr(0) As String
    strArr(0) = Expression

Done:
    Split = strArr

End Function

'returns a string by joining all array items together, seperated by delimiter...
'you can shorten it to only a certain amount of items by using the count argument.
'remeber that in a zero based array count should be 1 more than UBound of array items desired
Public Function Join(SourceArray() As String, _
      Optional ByVal Delimiter As String, _
      Optional ByVal Count As Long = -1) As String

  Dim lUpperBound  As Long  'array count
  Dim lLowerBound  As Long  'array lowest dim
  Dim lPosition    As Long  'position of item in array
  Dim lDelimLen    As Long  'length of delimiter
  Dim lTotal       As Long  'total length of string needed
    
    Err.Clear
    On Error Resume Next 'just in case array is not initialized
        lUpperBound = UBound(SourceArray)
        If Err.Number = 0 Then
            lLowerBound = LBound(SourceArray)
            'set upper bound to count
            If (Count <> -1) Then
                If (Count <= lUpperBound + 1) And (Count > lLowerBound) Then
                    lUpperBound = Count - 1
                End If
            End If
            lPosition = lLowerBound 'lowest array element
            lDelimLen = Len(Delimiter)
            'count length of string needed
            Do
                lTotal = lTotal + Len(SourceArray(lPosition)) + lDelimLen
                lPosition = lPosition + 1
            Loop While lPosition < lUpperBound + 1
            Join = Space$(lTotal - lDelimLen)
            lPosition = 1
            If lLowerBound < lUpperBound Then 'sanity check
                'loop thru array and add to string
                Do While lLowerBound < lUpperBound + 1
                    lTotal = Len(SourceArray(lLowerBound))
                    Mid$(Join, lPosition, lTotal) = SourceArray(lLowerBound) 'add string
                    lPosition = lPosition + lTotal
                    Mid$(Join, lPosition, lDelimLen) = Delimiter             'add delimiter
                    lPosition = lPosition + lDelimLen
                    lLowerBound = lLowerBound + 1
                Loop
            End If
            'get last item
            Mid$(Join, lPosition, Len(SourceArray(lUpperBound))) = SourceArray(lUpperBound)
        End If

End Function

'reverses a string, (very fast) :)...returns a string
Function StrReverse(ByVal Expression As String) As String

  Dim lLength   As Long  'length of expression
  Dim X         As Long  'loop counter

    lLength = Len(Expression)
    If lLength > 0 Then
        StrReverse = Space$(lLength)
        For X = lLength To 1 Step -1
            Mid$(StrReverse, X, 1) = Mid$(Expression, (lLength + 1) - X, 1)
        Next X
    End If

End Function

'rounds off a number,can specify number of places after decimal to keep,default is 0
Public Function Round(ByVal Expression As Variant, _
             Optional ByVal NumDigitsAfterDecimal As Long) As Variant

  Dim dFactor As Double
    
    'dont ask, just know that it works :)
    dFactor = CDbl("1" & String$(NumDigitsAfterDecimal, "0"))
    Round = Int(Expression * dFactor + 0.5) / dFactor

End Function

'looks thru a string in reverse, for an item, everything after start is ignored if you use it
'returns 0 if nothing found, position of item otherwise
Public Function InStrRev(ByVal StringCheck As String, _
                         ByVal StringMatch As String, _
                Optional ByVal Start As Long = -1, _
                Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

  Dim lCheckLen  As Long   'length of stringcheck
  Dim lPosition  As Long   'current position

        lCheckLen = Len(StringCheck) 'length of stringcheck
        If (lCheckLen > 0) And (Len(StringMatch) > 0) Then
            If Start < 1 Then Start = lCheckLen
            If Start <= lCheckLen Then 'sanity check
                'loop/look backwards thru string
                Do While (Start > 0) And (lPosition = 0)
                    lPosition = InStr(Start, StringCheck, StringMatch, Compare)
                    If lPosition > Start Then lPosition = 0 'ignore positions greater than start
                    Start = Start - 1                       'de-increment position
                Loop
                InStrRev = lPosition
            End If
        End If
        
End Function

'this function looks thru an array an removes any items that match the 'Value"
'unless 'Include' is true, if it is, it removes all items that dont match 'Value'
'returns an array in a variant, or original array if criteria is incorrect
Public Function Filter(InputStrings As Variant, _
                 ByVal Value As String, _
        Optional ByVal Include As Boolean, _
        Optional ByVal Compare As VbCompareMethod) As Variant

  Dim strFinal    As Variant 'temp array
  Dim lLowerBound As Long    'lower count of array
  Dim lUpperBound As Long    'upper count of array
  Dim lInclude    As Long    'absolute value of include
  Dim lCount      As Long    'count of items to remove
  Dim X           As Long    'For...Next loop counter
    
    On Error Resume Next         ' in case array is not initialized
        strFinal = InputStrings  'this way array passed in, is returned if its invalid criteria

        lUpperBound = UBound(InputStrings) 'upper count of array (will error if unitialized)
        If Err.Number = 0 Then
            lLowerBound = LBound(InputStrings) 'low count of array
            lInclude = Abs(Include)
            'count dimensions needed
            For X = lLowerBound To lUpperBound
                If (StrComp(InputStrings(X), Value, Compare) Xor lInclude) Then
                    lCount = lCount + 1
                End If
            Next X
            If lCount > 0 Then 'make sure we have at least 1 item
                ReDim strFinal(lCount - 1)
                lCount = 0
                'now fill temp array with values
                For X = lLowerBound To lUpperBound
                    If (StrComp(InputStrings(X), Value, Compare) Xor lInclude) Then
                        strFinal(lCount) = InputStrings(X)
                        lCount = lCount + 1
                    End If
                Next X
            End If
        End If
        
        Filter = strFinal

End Function


