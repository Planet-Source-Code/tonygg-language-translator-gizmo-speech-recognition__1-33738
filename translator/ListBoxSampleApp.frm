VERSION 5.00
Object = "*\ASAPI51ListBox.vbp"
Begin VB.Form MainForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Translater"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8805
   Icon            =   "ListBoxSampleApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8040
      Top             =   4080
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6240
      TabIndex        =   15
      Top             =   5760
      Width           =   2415
   End
   Begin VB.TextBox TxtNewItem2 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      Caption         =   "What Reciever Hears"
      Height          =   2175
      Left            =   4440
      TabIndex        =   9
      Top             =   1560
      Width           =   4215
      Begin VB.ListBox StandardListBox 
         Height          =   1815
         ItemData        =   "ListBoxSampleApp.frx":000C
         Left            =   120
         List            =   "ListBoxSampleApp.frx":000E
         TabIndex        =   16
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "What Sender Says"
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4215
      Begin SAPI51ListBox.Sample SpeechListBox 
         Height          =   1815
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   3201
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PreCommandString=   ""
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "About this demonstration"
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      Begin VB.Label Label2 
         Caption         =   $"ListBoxSampleApp.frx":0010
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   355
      Left            =   3000
      TabIndex        =   1
      Top             =   5760
      Width           =   3135
   End
   Begin VB.CheckBox chkSpeechEnabled 
      Caption         =   "Speech &enabled"
      Height          =   255
      Left            =   6840
      TabIndex        =   0
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   355
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox txtNewItem 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "You must specify both, this demonstration is not an engine."
      Height          =   195
      Left            =   2040
      TabIndex        =   14
      Top             =   5160
      Width           =   4140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "What the reciever hears, e.g Danke"
      Height          =   195
      Left            =   4800
      TabIndex        =   13
      Top             =   4560
      Width           =   2550
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Add a new word"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Language spoken, Thankyou"
      Height          =   195
      Left            =   4920
      TabIndex        =   10
      Top             =   3960
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add a new word"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1155
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'It is not perfect it is just a demonstration.
'A better version would use dictation and automatically translate between languages
'but dictation seems still very experimental. Accuracy is really bad

'Engine One: You would need dictionaries that could translate on the fly.
'meaning sending emails in any language is immediately understood providing
'the translation engine is apparent. Although that wouldn't require speech
'recognition, although it could. But a U.N conference on the other hand would,
'benerfit more perdominantly from this technology

'Engine two: You do not neccersary have to add a translation in the standard list box
'you could add a phrase so it seems the computer answers you. That's engine two
'Technology that gives robots and computers the ability to communicate based on
'how many optional responses are possible for a given thing you may say.

Option Explicit

Dim i As Integer

'Declare the SpVoice object.
Dim Voice As SpVoice
'Note - Applications that require handling of SAPI events should declair the
'SpVoice as follows:
'Dim WithEvents Voice As SpVoice

Private Sub CmdExit_Click()
End
End Sub

Private Sub Form_Load()

'   Initialize the voice object
    Set Voice = New SpVoice
    
    If SpeechListBox.SpeechEnabled Then
        chkSpeechEnabled = 1
    Else
        chkSpeechEnabled = 0
    End If
    
    
    SpeechListBox.AddItem "Thankyou"
    StandardListBox.AddItem "You are welcome"
    
    SpeechListBox.AddItem "How are you"
    StandardListBox.AddItem "Comme Stai"

    SpeechListBox.AddItem "Hello"
    StandardListBox.AddItem "Ciao"
    
    SpeechListBox.AddItem "What is the whether"
    StandardListBox.AddItem "It is very hot."
    
    SpeechListBox.AddItem "What is your CPU temperture"
    StandardListBox.AddItem "Status: Normal. 39 degrees"
    
End Sub

Private Sub chkSpeechEnabled_Click()
    SpeechListBox.SpeechEnabled = (chkSpeechEnabled = 1)
End Sub

Private Sub cmdAdd_Click()
    ' Add the new item. Internally to SpeechListBox, this will cause a rebuild
    ' of the dynamic grammar used by speech recognition engine.
    SpeechListBox.AddItem (txtNewItem)
    txtNewItem = ""
    
    StandardListBox.AddItem (TxtNewItem2)
    TxtNewItem2 = ""
End Sub

Private Sub cmdRemove_Click()
Dim i As Integer
    ' Just remove the current selected item. Same as AddItem, removing an item
    ' causes a grammar rebuild as well.
    If SpeechListBox.ListIndex <> -1 Then
    
    'must remove what is selected in sender lsitbox first becuase
    'then i becomes -1 before corresponding reciever word is removed.

          i = SpeechListBox.ListIndex
          StandardListBox.RemoveItem i
          
          SpeechListBox.RemoveItem SpeechListBox.ListIndex
          
    End If
End Sub

Private Sub StandardListBox_Click()
    Dim i As Integer
    
'       Call the Speak method with the text from the text box. We use the
'       SVSFlagsAsync flag to speak asynchronously and return immediately
'       from this call.
        If Not StandardListBox.ListCount - 1 Then
            Voice.Speak StandardListBox.List(StandardListBox.ListIndex), SVSFlagsAsync
        End If
        
Do
DoEvents                    'DoEvents lets events happen
Loop Until Voice.WaitUntilDone(10)  'Loop until voice finishes
    
End Sub

Private Sub Timer1_Timer()

    StandardListBox.ListIndex = SpeechListBox.ListIndex

End Sub

Private Sub txtNewItem_Change()
    ' Disallow empty item.
    cmdAdd.Enabled = txtNewItem <> ""
End Sub

Private Sub txtNewItem_GotFocus()
    ' When user focuses on the new item box, make the Add button default
    ' so that return key is same as clicking on Add button.
    cmdAdd.Default = True
End Sub

Private Sub TxtNewItem2_Change()
cmdAdd.Enabled = TxtNewItem2 <> ""
End Sub
