VERSION 5.00
Begin VB.Form FrmCalculator 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   3216
   ClientLeft      =   36
   ClientTop       =   516
   ClientWidth     =   2808
   ForeColor       =   &H8000000D&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3216
   ScaleWidth      =   2808
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCE 
      Caption         =   "CE"
      Height          =   348
      Left            =   768
      TabIndex        =   20
      Top             =   576
      Width           =   516
   End
   Begin VB.PictureBox Pic1 
      Height          =   348
      Left            =   96
      ScaleHeight     =   300
      ScaleWidth      =   468
      TabIndex        =   19
      Top             =   576
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   348
      Index           =   0
      Left            =   84
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2604
      UseMaskColor    =   -1  'True
      Width           =   516
   End
   Begin VB.CommandButton CmdEqual 
      Caption         =   "="
      Height          =   348
      Left            =   1428
      TabIndex        =   17
      Top             =   2604
      Width           =   516
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      Height          =   348
      Index           =   3
      Left            =   2184
      TabIndex        =   16
      Top             =   2604
      Width           =   516
   End
   Begin VB.CommandButton Operator 
      Caption         =   "*"
      Height          =   348
      Index           =   2
      Left            =   2184
      TabIndex        =   15
      Top             =   2100
      Width           =   516
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      Height          =   348
      Index           =   1
      Left            =   2184
      TabIndex        =   14
      Top             =   1596
      Width           =   516
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      Height          =   348
      Index           =   0
      Left            =   2184
      TabIndex        =   13
      Top             =   1092
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Index           =   10
      Left            =   756
      TabIndex        =   12
      Top             =   2604
      Width           =   516
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "Back"
      Height          =   348
      Left            =   2184
      TabIndex        =   10
      Top             =   576
      Width           =   516
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   348
      Left            =   1428
      TabIndex        =   9
      Top             =   576
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   348
      Index           =   5
      Left            =   756
      TabIndex        =   4
      Top             =   1596
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   348
      Index           =   9
      Left            =   1440
      TabIndex        =   8
      Top             =   1092
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   348
      Index           =   8
      Left            =   768
      TabIndex        =   7
      Top             =   1092
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      Height          =   348
      Index           =   7
      Left            =   84
      TabIndex        =   6
      Top             =   1092
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   348
      Index           =   6
      Left            =   1440
      TabIndex        =   5
      Top             =   1596
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   348
      Index           =   4
      Left            =   84
      TabIndex        =   3
      Top             =   1596
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   348
      Index           =   3
      Left            =   1428
      TabIndex        =   2
      Top             =   2100
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   348
      Index           =   2
      Left            =   756
      TabIndex        =   1
      Top             =   2100
      Width           =   516
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   348
      Index           =   1
      Left            =   84
      TabIndex        =   0
      Top             =   2100
      Width           =   516
   End
   Begin VB.Label Display 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   84
      TabIndex        =   11
      Top             =   84
      Width           =   2640
   End
   Begin VB.Menu mnuFileItem 
      Caption         =   "&File"
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FrmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Hi. This is my first Visual Basic program. I have used QuickBasic
'for a couple of months but that's about it. So bare with me!!
'Comments would be great. I would like to hear any ways to improve
'or whatever. Enjoy!! Please rate it if you get the time.

'Sorry for all the globals. Still learning.
Dim EqualFlag, OpFlag, ErrorFlag  As Boolean
Dim DisplayLen, MaxLen As Integer
Dim Dot, StoreIt, EntryLen As Integer
Dim Entry, LastEntry, NewEntry As String
Dim OverFlow, Operand As String
Dim Value1, Value2 As Double
Dim GetResult As Variant

Private Sub CmdBack_Click()
             
   If Entry = "" Then Exit Sub                  'If the Entry string is empty Exit Sub.
   If OpFlag = True Then Exit Sub               'If operator pushed then Exit Sub.
   If EqualFlag = True Then Exit Sub            'If user just pushed the equal button
                                                'then Exit Sub.
   NewEntry = Entry                             'Store Entry into New String.
   NewEntry = Mid(Entry, 1, (Len(Entry) - 1))   'Subtract one from the right of the NewString.
   Entry = Trim(NewEntry)                       'Trim it up to get any spaces out of it.
   
   ShowResult (Entry)                           'Finally Display the "new" Entry.
   
   If Not InStr(Entry, ".") Then Dot = 0        'If user erases the decimal point with
                                                'the BackSpace then we need to re-initialize it.
   Pic1.SetFocus
    
End Sub
 
Private Sub cmdCE_Click()
    
    Display = "0. "                                'Let's reset the display.
    Entry = ""                                     'Let's clear Entry only,
                                                   'not the whole operation.
    Dot = 0                                        'Reset the decimal flag.
    Pic1.SetFocus                                  'Give focus to pic1.

End Sub

Private Sub CmdClear_Click()
    
    Form_Load                                       'Clear all the variables.
        
End Sub

Private Sub CmdEqual_Click()
    
    EqualFlag = True                                'Flag for the BackSpace key
    If ErrorFlag = True Then
        Display = "OverFlow. Please Clear."
        Beep
        Exit Sub
    End If
    
    'Make sure last entry was a number.
    If LastEntry < "0" Or LastEntry > "9" Then Exit Sub
    If Number(10) And Entry = "" Then Exit Sub
    
    
    Value2 = Entry                                  'Get Entry and store it.
    GetResult = EqualIt(Value1, Operand, Value2)    'Pass all three variables to Function
                                                    'EqualIT and do the math.
    ShowResult (GetResult)                          'When the value gets returned
                                                    'call the sub ShowResult
End Sub                                             'to show the result.

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 46                                     'KeyCode for CE. (Delete)
            Call cmdCE_Click
    End Select
            
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Static DotPushed As Boolean
    Select Case KeyAscii
       
        Case 48 To 57                               'KeyCode for numbers 0 through 9.
            Call Number_Click(Chr$(KeyAscii))
        Case 43                                     'KeyCode for addition key.
           Call Operator_Click(0)
           DotPushed = False
        Case 45                                     'KeyCode for subtraction.
           Call Operator_Click(1)
           DotPushed = False
        Case 42                                     'KeyCode for multiplication.
           Call Operator_Click(2)
           DotPushed = False
        Case 47                                     'KeyCode for the division.
          Call Operator_Click(3)
          DotPushed = False
        Case 13                                     'KeyCode for Enter.
          Call CmdEqual_Click
          DotPushed = False
        Case 8                                      'KeyCode for Backspace.
          Call CmdBack_Click
          If Not InStr(Entry, ".") Then DotPushed = False
        Case 27                                     'KeyCode for Clear. (Escape)
          Call CmdClear_Click
          DotPushed = False
        Case 46                                     'KeyCode for Decimal Point.
          If DotPushed = True Then Exit Sub
          Call Number_Click(10)
          DotPushed = True
        
    End Select

End Sub

Private Sub Form_Load()
    
    Display = "0. "                                 'Set the Display to 0.
    Dot = 0
                                                    
    Entry = ""                                      'Clear and reset all values.
    EntryLen = 0
    LastEntry = ""
       
    Value1 = 0
    Value2 = 0
    StoreIt = 0
    Operand = ""
    
    ErrorFlag = False
    EqualFlag = False
            
    Me.Show
    Pic1.SetFocus
    
End Sub

Private Sub mnuExitItem_Click()
    
    End
    
End Sub

'Let's get some numbers.
Private Sub Number_Click(Index As Integer)
  
    If ErrorFlag = True Then Exit Sub
    
    OpFlag = False                              'Reset the Op flag which is used for BackSpace.
    MaxLen = 18                                 'Max Length of entry allowed is 15 characters.
    If EntryLen > MaxLen Then Exit Sub          'If MaxLength of string is reached then accept no more.
    
    If Display = "0. " And Number(0) Then Exit Sub 'Do not let user start out with a zero.
            
    If Number(10) Then Dot = Dot + 1            'If decimal point is pressed then set
      'Number(10) = Decimal Point.              'Dot to one.
    
    If Dot > 1 Then                             'If dot is greater than one, then
        Entry = Entry                           'accept no more decimal points.
        Dot = 1                                 'Reset dot to one to prevent
    Else                                        'any more decimal points.
        Entry = Entry & Number(Index).Caption   'Add whatever number user presses to
    End If                                      'entry string.
          
    ShowResult (Entry)                          'This sub is called to show the result
                                                'of what number was pressed.
    
    EntryLen = Len(Entry)                       'This is used to get the length of the string.
    LastEntry = Number(Index).Caption           'Store the last number pressed.
    Pic1.SetFocus
    
End Sub

'Now we need to store the numbers
'and get which operator was pressed.
Private Sub Operator_Click(Index As Integer)
    
    If ErrorFlag = True Then Exit Sub
    
    'Let's collect the last character and test to see if it's a number or not.
    If LastEntry < "0" Or LastEntry > "9" Then Exit Sub
            
    'Here we use Select Case to store the values.
    'Everytime an operator (+,-,*,=) is pressed StoreIt increments by one,
    'therefore storing the value of Entry in it's rightful container.
    'The only reason we need to store the second value is if say ..
    'the user inputs  x + y + y instead of x + y =.
    
    'We we want to do the math as we go to make it more convinient for the user.
    StoreIt = StoreIt + 1
    Select Case StoreIt
        
        Case 1
            Value1 = Entry
        Case Else
            Value2 = Entry
            GetResult = EqualIt(Value1, Operand, Value2)
            ShowResult (GetResult)
    End Select
                        
    'This line gets which operator was pressed and stores it in Operand.
    Operand = Operator(Index).Caption
        
    'Let's clear out everything so we can accept a clean input.
    Entry = ""
    EntryLen = 0
    LastEntry = ""
    Dot = 0                             'Reset the Decimal Point Flag.
    OpFlag = True                       'Set OpFlag in case user tries to press
                                        'BackSPace right after an operator.
    
    'Storeit is reset back to one so we can store a value into case else.(Optional)
    StoreIt = 1
    
    'Return the focus to pic1(upperleft corner of calculator).
    'If we don't do this then the keyboard input will not work properly.
    'Plus it gets rid of the ugly focus dealy.
    Pic1.SetFocus
    
End Sub

'This function is used to do the math of whatever is passed to it.
Function EqualIt(x, Op As String, y As Double)
'x = Value1     'Op = Operator       'y = Value2

On Error GoTo OverFlow

    Select Case Op
               
        Case "+"                                            'Add x and y.
            x = Val(x) + Val(y)
        Case "-"                                            'Subtract x and y.
            x = Val(x) - Val(y)
        Case "*"                                            'Multiply x and y.
            x = Val(x) * Val(y)
        Case "/"                                            'Divide x and y
            If Val(x) = 0 Or Val(y) = 0 Then                'If either values
                MsgBox ("Division by zero"), vbOKOnly       'equal zero
                Form_Load                                   'throw up an error
            Else                                            'and reset everything.
                x = Val(x) / Val(y)
            End If
        
    End Select

    EqualIt = x                                             'Return the answer.
    
OverFlow:
    If Err.Number = 6 Then ErrorFlag = True                 'If overflow then set flag.
    
End Function

Private Sub ShowResult(Value As String)
    
    MaxLen = 20
      
    If InStr(Value, ".") Then                  'Let's decide which Display Format to use.
        Display = Mid(Value, 1, MaxLen) & " "  'If the decimal point has been pressed,
    Else                                       'then stop displaying it at the end.
        Display = Mid(Value, 1, MaxLen) & ". " 'Else keep showing it.
    End If

End Sub
