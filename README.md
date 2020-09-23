<div align="center">

## convert/write a number in words


</div>

### Description

Takes any numerical value (less a billion) like "203463110" and outputs "Two Hundred Three Million Four Hundred Sixty Three Thousand One Hundred Ten"
 
### More Info
 
] Sub Translate(Number As String) [

The user/app send the translate sub some numerical value from a text or input box

Purpose:

was used to automate typing numerical values in words for stock certificates but can also be used in other financial apps (like an app that prints out information on checks).

'

----

Instalation:

cut and paste this code onto a form. Then hit F5

] NumberInText$ [

The sub return a string named "NumberInText$," which is the number worded out.

No side effects


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Noel Khan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/noel-khan.md)
**Level**          |Beginner
**User Rating**    |3.3 (10 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/noel-khan-convert-write-a-number-in-words__1-31456/archive/master.zip)





### Source Code

```
'*********************
'  DECLARATIONS
'*********************
Dim X As String   'the number input
Dim q As Long    'currently parsed-digit counter
Dim i As Long    'currently parsed 3-digit set, i.e., "000######", "###000###", etc
Dim NumberInText As String 'output, this is the translation of the numerical value
Dim BeginningInterval As Long  'counter to tract which 3-digit set the program is reading
Dim EndingInterval As Long   'counter to tract which 3-digit set the program is reading
Dim Temp As Variant 'temporary parse
'===============================================
Private Sub Translate(Number As String)
'INPUT: "NUMBER" PARAMETER,i.e., some numerical value
'OUTPUT: "NumberInText$" STRING, i.e., the number spelled out in words
'ASSUMES: input must be in 9-digit format, use the format function to ensure that it is
'REQUIRES: the following two related subs
    '1)HundredsPlaceOROnesPlace
    '2)TensPlace
    'and also the above declarations
'*********************
'  INITIALIZATION
'*********************
NumberInText$ = Empty
q = Empty
i = Empty
BeginningInterval = Empty
EndingInterval = Empty
Temp = Empty
'**********************
'   TRANSLATION
'**********************
  'the program reads the input in upto 3 sets (intervals) of 3 digits
  'at a time i.e., the millions, thousands, and hundreds
  For i = 1 To 3
    'the following counters keep track of which 3-digit set
    'the program is reading from
    BeginningInterval = EndingInterval + 1
    EndingInterval = EndingInterval + 3
      'now that the program has parsed upto three digits, its reads
      'and translates one digit at a time
      For q = BeginningInterval To EndingInterval
          'i use a temp variable to hold the single digit parse
          'if the parse is a zero, then skip on over to the next digit
          Temp = Mid(X$, q, 1): If Temp = "0" Then GoTo Escape
            'the next few lines essentially determines if the
            'suffix, "hundreds," is used and also determines
            'where to send the parse for translation.
            If q = 1 Xor q = 4 Xor q = 7 Then Call HundredsPlaceOROnesPlace: NumberInText$ = NumberInText$ & "Hundred "
            If q = 2 Xor q = 5 Xor q = 8 Then Call TensPlace
            If q = 3 Xor q = 6 Xor q = 9 Then Call HundredsPlaceOROnesPlace
Escape:
      Next q
    'the next couple lines essentially determines
    'if the suffix, million or thousand
    If EndingInterval = 3 And Not X$ Like "000######" Then NumberInText$ = NumberInText$ & "Million "
    If EndingInterval = 6 And Not X$ Like "###000###" Then NumberInText$ = NumberInText$ & "Thousand "
  Next i
End Sub
'===============================================
Private Sub HundredsPlaceOROnesPlace()
  Select Case Temp
    Case Is = "1": NumberInText$ = NumberInText$ & "One "
    Case Is = "2": NumberInText$ = NumberInText$ & "Two "
    Case Is = "3": NumberInText$ = NumberInText$ & "Three "
    Case Is = "4": NumberInText$ = NumberInText$ & "Four "
    Case Is = "5": NumberInText$ = NumberInText$ & "Five "
    Case Is = "6": NumberInText$ = NumberInText$ & "Six "
    Case Is = "7": NumberInText$ = NumberInText$ & "Seven "
    Case Is = "8": NumberInText$ = NumberInText$ & "Eight "
    Case Is = "9": NumberInText$ = NumberInText$ & "Nine "
    Case Else:
  End Select
End Sub
'===============================================
Private Sub TensPlace()
If Temp = 1 Then
  Temp = Mid(X$, q, 2)
    Select Case Temp
      Case Is = "10": NumberInText$ = NumberInText$ & "Ten ": q = q + 1
      Case Is = "11": NumberInText$ = NumberInText$ & "Eleven ": q = q + 1
      Case Is = "12": NumberInText$ = NumberInText$ & "Twelve ": q = q + 1
      Case Is = "13": NumberInText$ = NumberInText$ & "Thirteen ": q = q + 1
      Case Is = "14": NumberInText$ = NumberInText$ & "Fourteen ": q = q + 1
      Case Is = "15": NumberInText$ = NumberInText$ & "Fifteen ": q = q + 1
      Case Is = "16": NumberInText$ = NumberInText$ & "Sixteen ": q = q + 1
      Case Is = "17": NumberInText$ = NumberInText$ & "Seventeen ": q = q + 1
      Case Is = "18": NumberInText$ = NumberInText$ & "Eighteen ": q = q + 1
      Case Is = "19": NumberInText$ = NumberInText$ & "Nineteen ": q = q + 1
    End Select
Else
    Select Case Temp
      Case Is = "2": NumberInText$ = NumberInText$ & "Twenty "
      Case Is = "3": NumberInText$ = NumberInText$ & "Thirty "
      Case Is = "4": NumberInText$ = NumberInText$ & "Forty "
      Case Is = "5": NumberInText$ = NumberInText$ & "Fifty "
      Case Is = "6": NumberInText$ = NumberInText$ & "Sixty "
      Case Is = "7": NumberInText$ = NumberInText$ & "Seventy "
      Case Is = "8": NumberInText$ = NumberInText$ & "Eighty "
      Case Is = "9": NumberInText$ = NumberInText$ & "Ninety "
      Case Else
    End Select
End If
End Sub
'===============================================
Private Sub Form_Load()
Again:
  X$ = InputBox("Enter any number less than a billion." & vbCrLf & vbCrLf & "Type 'exit' to exit", "Number to Translate")
  If X$ = "exit" Then
    GoTo Exiting
  Else
    X$ = Format(X$, "000000000")  'input must be in nine digit format
    Call Translate(X$)
    MsgBox Format(X$, "###,###,###") & " = " & vbCrLf & vbCrLf & NumberInText$, vbOKOnly, "Translation"
    GoTo Again:
  End If
Exiting:
Unload Me
End Sub
```

