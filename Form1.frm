VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB6 Functions For VB5"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReverseSpeed 
      Caption         =   "Reverse"
      Height          =   330
      Left            =   5490
      TabIndex        =   23
      Top             =   2835
      Width           =   915
   End
   Begin VB.CommandButton cmdReverse 
      Caption         =   "Reverse"
      Height          =   330
      Left            =   5535
      TabIndex        =   22
      Top             =   1890
      Width           =   825
   End
   Begin VB.CommandButton cmdRoundSpeed 
      Caption         =   "Round"
      Height          =   330
      Left            =   4545
      TabIndex        =   21
      Top             =   2835
      Width           =   915
   End
   Begin VB.CommandButton cmdRound 
      Caption         =   "Round"
      Height          =   330
      Left            =   5535
      TabIndex        =   20
      Top             =   1530
      Width           =   825
   End
   Begin VB.CommandButton cmdFilterSpeed 
      Caption         =   "Filter"
      Height          =   330
      Left            =   2745
      TabIndex        =   19
      Top             =   2835
      Width           =   825
   End
   Begin VB.CommandButton cmdInstrRev 
      Caption         =   "InstrRev"
      Height          =   330
      Left            =   5535
      TabIndex        =   18
      Top             =   2250
      Width           =   825
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter"
      Height          =   330
      Left            =   5520
      TabIndex        =   17
      Top             =   1170
      Width           =   825
   End
   Begin VB.CommandButton cmdInstrRevSpeed 
      Caption         =   "InstrRev"
      Height          =   330
      Left            =   3600
      TabIndex        =   16
      Top             =   2835
      Width           =   915
   End
   Begin VB.CommandButton cmdJoinSpeed 
      Caption         =   "Join"
      Height          =   330
      Left            =   1035
      TabIndex        =   15
      Top             =   2835
      Width           =   825
   End
   Begin VB.CommandButton cmdSplitSpeed 
      Caption         =   "Split"
      Height          =   330
      Left            =   1890
      TabIndex        =   14
      Top             =   2835
      Width           =   825
   End
   Begin VB.TextBox txtTime 
      Height          =   330
      Left            =   90
      TabIndex        =   10
      Top             =   3240
      Width           =   6315
   End
   Begin VB.CommandButton cmdReplaceSpeed 
      Caption         =   "Replace"
      Height          =   330
      Left            =   90
      TabIndex        =   8
      Top             =   2835
      Width           =   915
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   90
      TabIndex        =   7
      Top             =   495
      Width           =   1680
   End
   Begin VB.CommandButton cmdSplit 
      Caption         =   "Split"
      Height          =   330
      Left            =   5535
      TabIndex        =   6
      Top             =   90
      Width           =   825
   End
   Begin VB.TextBox txtJoin 
      Height          =   330
      Left            =   90
      TabIndex        =   5
      Top             =   90
      Width           =   5280
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join"
      Height          =   330
      Left            =   5535
      TabIndex        =   4
      Top             =   450
      Width           =   825
   End
   Begin VB.TextBox txtReplace 
      Height          =   330
      Left            =   2835
      TabIndex        =   3
      Text            =   "cr"
      Top             =   1305
      Width           =   2535
   End
   Begin VB.TextBox txtFind 
      Height          =   330
      Left            =   2835
      TabIndex        =   2
      Text            =   "cw"
      Top             =   900
      Width           =   2535
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Height          =   330
      Left            =   5535
      TabIndex        =   1
      Top             =   810
      Width           =   825
   End
   Begin VB.TextBox txtExpression 
      Height          =   330
      Left            =   2835
      TabIndex        =   0
      Text            =   "something's really scwewy here"
      Top             =   495
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Replace:"
      Height          =   240
      Index           =   2
      Left            =   2115
      TabIndex        =   13
      Top             =   1395
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Find:"
      Height          =   240
      Index           =   1
      Left            =   2385
      TabIndex        =   12
      Top             =   990
      Width           =   420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Expression:"
      Height          =   240
      Index           =   0
      Left            =   1980
      TabIndex        =   11
      Top             =   585
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed Comparison Tests"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   2520
      Width           =   2085
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Option Explicit ':( Line inserted by Formatter
Const pi = 3.14159265359

'split test
Private Sub cmdSplit_Click()

  Dim TestArray As Variant
  Dim X As Long

    List1.Clear

    txtJoin = "test|test|test|test|test|test|test|test|test"
    TestArray = Split(txtJoin, "|")

    'display results
    For X = 0 To UBound(TestArray)
        List1.AddItem TestArray(X)
    Next X

End Sub

'join an array to a string
Private Sub cmdJoin_Click()

  Dim TestArray(5) As String
  Dim X As Long

    List1.Clear

    'make a test array
    For X = 0 To 5
        TestArray(X) = "test"
        List1.AddItem TestArray(X)
    Next X

    'display results
    txtJoin = Join$(TestArray, "|")

End Sub

'test string replacing
Private Sub cmdReplace_Click()

    txtExpression = "something's really scwewy here"
    txtFind = "cw"
    txtReplace = "cr"

    txtJoin = Replace(txtExpression, txtFind, txtReplace)
    txtJoin = Replace(txtJoin, "something", "nothing")

End Sub

'test filtering an array
Private Sub cmdFilter_Click()

  Dim Test As Variant
  Dim X As Long

    List1.Clear

    ReDim Test(25)

    For X = 0 To 25
        If X Mod 2 = 0 Then
            Test(X) = "test"
          Else
            Test(X) = "remove"
        End If
    Next X

    Test = Filter(Test, "remove", False, vbTextCompare)

    'display result
    For X = 0 To UBound(Test)
        List1.AddItem Test(X)
    Next X

End Sub

'test rounding
Private Sub cmdRound_Click()

    txtExpression = pi
    txtFind = "2"       '2 decimal places
    txtReplace = ""

    txtJoin = modVB5.Round(txtExpression, txtFind)

End Sub

'string reverse test
Private Sub cmdReverse_Click()

    txtExpression = "sdrawkcab yllanigiro saw gnirts sihT"
    txtFind = ""
    txtReplace = ""

    txtJoin = StrReverse$(txtExpression)

End Sub

'test reverse instr()
Private Sub cmdInstrRev_Click()

    txtExpression = "000000000000001000000"
    txtFind = "010"
    txtReplace = ""
    txtJoin = modVB5.InStrRev(txtExpression, txtFind)

End Sub

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'SPEED TESTS
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdReplaceSpeed_Click()

  Dim TestString As String
  Dim sngTest1 As Single
  Dim sngTest2 As Single
  Dim X As Long

    txtTime = "preparing string"
    DoEvents

    'prepare a temp string
    TestString = Space$(10000)
    For X = 1 To 10000
        If X Mod 500 = 0 Then
            Mid$(TestString, X, 1) = "X"
          Else
            Mid$(TestString, X, 1) = "A"
        End If
    Next X

    txtTime = "replacing"
    DoEvents

    'loop and replace 1000 times
    'method 1
    sngTest1 = Timer
    For X = 0 To 1000
        Call Replace(TestString, "X", "Y")
    Next X
    sngTest1 = Timer - sngTest1

    txtTime = "Loops: 1000   VB5 Time1: " & CStr(sngTest1) '& "  Time2: " & CStr(sngTest2)

End Sub

Private Sub cmdJoinSpeed_Click()

  Dim sngTest1 As Single
  Dim sngTest2 As Single
  Dim Test(500) As String
  Dim X As Long

    For X = 0 To 500
        Test(X) = "test"
    Next X

    txtTime = ""

    sngTest1 = Timer
    For X = 0 To 1000
        Call modVB5.Join(Test, "|")
    Next X

    sngTest1 = Timer - sngTest1

    'sngTest2 = Timer
    'Call VBA.Join(Test, "|")
    'sngTest2 = Timer - sngTest2

    txtTime = "Loops: 1000  VB5 Time: " & CStr(sngTest1) '& "   VB6 Time: " & CStr(timer - sngTest2)

End Sub

Private Sub cmdSplitSpeed_Click()

  Dim sngTest1 As Single
  Dim sngTest2 As Single
  Dim Test As String
  Dim X As Long

    txtTime = ""

    For X = 1 To 500
        Test = Test & "test|"
    Next X

    Test = Test & "test"

    sngTest1 = Timer
    For X = 0 To 1000
        Call Split(Test, "|")
    Next X
    sngTest1 = Timer - sngTest1

    'sngTest2 = Timer
    'Call VBA.Split(Test, "|")
    'sngTest2 = Timer - sngTest2

    txtTime = "Loops: 1000  VB5 Time: " & CStr(sngTest1) '& "   VB6 Time: " & CStr(sngTest2)

End Sub

Private Sub cmdFilterSpeed_Click()

  Dim sngTest1 As Single
  Dim sngTest2 As Single
  Dim Test As Variant
  Dim X As Long

    ReDim Test(100)
    txtTime = ""

    'prepare an array
    For X = 0 To 100
        If X Mod 4 = 0 Then
            Test(X) = "remove"
          Else
            Test(X) = "test"
        End If
    Next X

    sngTest1 = Timer
    For X = 0 To 1000
        Call Filter(Test, "remove", False)
    Next X
    sngTest1 = Timer - sngTest1

    'sngTest2 = Timer
    'For X = 0 To 100
    'Call VBA.Filter(Test, "remove", False)
    'Next X
    'sngTest2 = Timer - sngTest2

    txtTime = "Loops: 1000  VB5 Time: " & CStr(sngTest1) '& "   VB6 Time: " & CStr(sngTest2)

End Sub

Private Sub cmdInstrRevSpeed_Click()

  Dim sngTest1 As Single
  Dim sngTest2 As Single
  Dim Test As String
  Dim X As Long

    txtTime = ""

    For X = 1 To 5000
        Test = Test & "test|"
    Next X

    Test = Test & "test"

    sngTest1 = Timer
    For X = 0 To 1000
        Call InStrRev(Test, "|")
    Next X
    sngTest1 = Timer - sngTest1

    'sngTest2 = Timer
    'Call VBA.InstrRev(Test, "|")
    'sngTest2 = Timer - sngTest2

    txtTime = "Loops: 1000  VB5 Time: " & CStr(sngTest1) '& "   VB6 Time: " & CStr(sngTest2)

End Sub

Private Sub cmdRoundSpeed_Click()

  Dim sngTest1 As Single
  Dim sngTest2 As Single
  Dim X As Long

    txtTime = ""

    sngTest1 = Timer

    For X = 0 To 100000
        Call Round(pi, 3)
    Next X
    sngTest1 = Timer - sngTest1

    'sngTest2 = Timer
    'Call VBA.InstrRev(Test, "|")
    'sngTest2 = Timer - sngTest2

    txtTime = "Loops: 100000    VB5 Time: " & CStr(sngTest1) '& "   VB6 Time: " & CStr(sngTest2)

End Sub

Private Sub cmdReverseSpeed_Click()

  Dim sngTest1 As Single
  Dim sngTest2 As Single
  Dim Test As String
  Dim X As Long

    txtTime = ""

    For X = 1 To 100
        Test = Test & "test"
    Next X

    sngTest1 = Timer
    For X = 0 To 1000
        Call StrReverse$(Test)
    Next X
    sngTest1 = Timer - sngTest1

    'sngTest2 = Timer
    'Call VBA.InstrRev(Test, "|")
    'sngTest2 = Timer - sngTest2

    txtTime = "Loops: 1000  VB5 Time: " & CStr(sngTest1) '& "   VB6 Time: " & CStr(sngTest2)

End Sub

