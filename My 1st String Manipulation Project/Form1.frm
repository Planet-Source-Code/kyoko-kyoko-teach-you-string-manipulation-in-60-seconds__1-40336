VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Complete String Manipulation"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Text            =   "Text12"
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Text            =   "Text11"
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Text            =   "Text10"
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Text            =   "Text9"
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Text            =   "Text8"
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Text            =   "Text7"
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Text            =   "Text6"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Text            =   "Text5"
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   960
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Just Do It!"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "I Love You"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2760
      Y1              =   4200
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   1200
      Y1              =   4200
      Y2              =   4320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "I Love You + I Hate You"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text2.Text = Len(Text1.Text) ' shows that text1.text got how many letters(I Love You has 10 letters)

    Text3.Text = Left(Text1.Text, 2) 'get 2 letters from left

    Text4.Text = Right(Text1.Text, 3) 'get 3 letters from right
    
    Text5.Text = Mid(Text1.Text, 3, 4) 'Count from the 3rd letter of the sentence, and get 4 letters after it

    Text6.Text = UCase$(Text1.Text) 'This function turns all letters in text1 to Upper Case

    Text7.Text = LCase$(Text1.Text) 'This function turns all letters in text1 to Lower Case

    Text8.Text = StrReverse(Text1.Text) 'This function reverse the string in text1(Run this code and u will know!)

    Text9.Text = Replace(Text1.Text, "Love", "Hate") 'replace the "Love" letter in text1.text to "Hate"

    a = InStr("1", Text1.Text, "Love") ' find "Love" letter start from the 1st letter in text1.text, if none then a = 0
    If a <> 0 Then Text10.Text = "The Letter Love Found!!" 'if a word/letter found, then a > 0

    Text11.Text = Split(Label1.Caption, "+")(0)    'get the word before the +
    Text12.Text = Split(Label1.Caption, "+")(1)    'get the word after the +
'Split() Function allows you To create a one-dimensional array, by splitting a String by reconising a certain character, Then putting any text after the character on a new line in the array.
End Sub
