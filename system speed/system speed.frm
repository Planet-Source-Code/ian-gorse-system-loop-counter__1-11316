VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "System Speed"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Copy"
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtresults 
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4320
      Width           =   5175
   End
   Begin VB.ListBox List1 
      Columns         =   3
      Height          =   1425
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Some Information"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.TextBox txtwin 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Text            =   "98"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtRAM 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Text            =   "128"
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtProc 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "AMD K6 2 500 mhz"
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Windows"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "RAM"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Processor"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '''///Please run a 15 loop count for adding to PSC\\\\'''
    
    'Number of times to loop
    averagecnt = 15
    
    'Starts the loop
    For lop = 1 To averagecnt
    starttime = Timer
    Do
        
        cnt = cnt + 1
        If Int(Timer - starttime) = 1 Then Exit Do
    Loop
        List1.AddItem "Attempt " & lop & " > " & cnt
        Max& = Max& + cnt
        cnt = 0
    Next lop
    
    'Now find the average
    average = Max& / averagecnt
    
    
    'Update the results box
    txtresults = "Processor :" & txtProc & vbCrLf & _
                 "RAM :" & txtRAM & vbCrLf & _
                 "Windows " & txtwin & vbCrLf & _
                 "Average Loop :" & average
                 
End Sub

Private Sub Command2_Click()
    Clipboard.SetText txtresults
End Sub
