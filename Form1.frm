VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Find CD"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "List All Drives"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
allDrives$ = Space$(64)
       
Form1.Cls   'clear form of lettering

ret& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
'trim off any trailing spaces. AllDrives$
'now contains all the drive letters.
allDrives$ = Left$(allDrives$, ret&)

Do
   'first check that there is a chr$(0) in the string
   pos% = InStr(allDrives$, Chr$(0))
   'if there's one, then...

     If pos% Then
     'extract the drive up to the chr$(0)
     JustOneDrive$ = Left$(allDrives$, pos% - 1)
               
     'and remove that from the Alldrives string,
     'so it won't be checked again
     allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
               
     'with the one drive, call the API to
     'determine the drive type
     DriveType& = GetDriveType(JustOneDrive$)
     'check if it's what we want

             If DriveType& = 5 Then 'then it is a CD Drive
                Print UCase$(JustOneDrive$) & " is a CD Drive"
             Else
                Print UCase$(JustOneDrive$) & " is NOT a CD Drive"
             End If

     End If


Loop Until allDrives$ = ""
 
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()


       Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
       
   End Sub






