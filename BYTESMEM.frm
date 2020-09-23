VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Get Memory"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4920
   Icon            =   "BYTESMEM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Bytes Strings Boolean Integer Long Single Double Conversions"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton Command8 
         Height          =   615
         Left            =   3600
         Picture         =   "BYTESMEM.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Get Memory"
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Get Memory"
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Get Memory"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   840
         TabIndex        =   16
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Get Memory"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Get Memory"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Get Memory"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get Memory"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Add All:"
         Height          =   255
         Left            =   3600
         TabIndex        =   24
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Depends on amount of char. in each string"
         Height          =   615
         Left            =   3360
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4080
         Picture         =   "BYTESMEM.frx":1194
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label7 
         Caption         =   "Double Declared:"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Single Declared:"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Long Declared:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Integer Declared:"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Boolean Declared:"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Strings Declared:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Bytes Declared:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
TaleyMem Text1, 1
End Sub
Function TaleyMem(Text As TextBox, NumDataTypes As Byte)
    'First we must see if some of you people
    'like to enter things in besides number
    'This will get the users data from textbox
    'then multiply it by the amount of data types
If Text.Text < Chr$(48) Or Text.Text > Chr$(57) Or Text.Text = vbNullString Then
   MsgBox ("You have entered a charector in the textbox or have to numerical value."), vbExclamation
   GoTo Terminate
Else
    Number = (Text.Text * NumDataTypes)
    MsgBox ("Number of bytes equals " & Number & " " & " Bytes in Memory")
End If
Terminate:

End Function
Sub TempTaleyTemp()
Sum = 0
TotalNumDT = 0
    '#Temp#
    'This will get the users data from textbox
    'then multiply it by the amount of data types
If Text1.Text = vbNullString Or Text2.Text = vbNullString Or _
Text3.Text = vbNullString Or Text4.Text = vbNullString Or _
Text5.Text = vbNullString Or Text6.Text = vbNullString Or _
Text7.Text = vbNullString Then
    MsgBox ("You have not enterted in a value."), vbExclamation
Else
On Error GoTo Err
'This will count all variables
Sum = Sum + Val(Text1.Text * 1) + Val(Text2.Text * 1) + Val(Text3.Text * 2) + _
 Val(Text4.Text * 2) + Val(Text5.Text * 4) + Val(Text6.Text * 4) + Val(Text7.Text * 8)
    'This will determine all btyes according to the memory addresses
    TotalNumDT = TotalNumDT + Val(Text1.Text) + Val(Text2.Text) + Val(Text3.Text) + _
 Val(Text4.Text) + Val(Text5.Text) + Val(Text6.Text) + Val(Text7.Text)
    
    'Output
    MsgBox ("Bytes in Mem = " & TotalNumDT & vbCrLf & "Amount of Variables = " & Sum)
GoTo Sucess
End If
Err:
MsgBox ("Error calculating the data"), vbExclamation
Sucess:
'if Sucess go here
End Sub
Private Sub Command2_Click()
TaleyMem Text2, 1
End Sub

Private Sub Command3_Click()
TaleyMem Text3, 2
End Sub

Private Sub Command4_Click()
TaleyMem Text4, 2
End Sub

Private Sub Command5_Click()
TaleyMem Text5, 4
End Sub

Private Sub Command6_Click()
TaleyMem Text6, 4
End Sub

Private Sub Command7_Click()
TaleyMem Text7, 8
End Sub
Private Sub Command8_Click()
TempTaleyTemp
End Sub

