VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Trial Test Program"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pBar 
      Height          =   360
      Left            =   1665
      TabIndex        =   7
      Top             =   1837
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   1
      Max             =   30
   End
   Begin VB.CommandButton cmdTry 
      Caption         =   "Try"
      Height          =   350
      Left            =   4230
      TabIndex        =   6
      Top             =   2685
      Width           =   1010
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   2835
      TabIndex        =   5
      Top             =   2685
      Width           =   1230
   End
   Begin VB.PictureBox pTop 
      Align           =   1  'Align Top
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   0
      Width           =   5385
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   585
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Label lblTrial 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trial Test Program"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Label lblTitle1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   1515
      Width           =   255
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1920
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is an example of a 30 day trial program, That you can use in your programs in Visual Basic"
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Width           =   5400
   End
End
Attribute VB_Name = "frmTrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a basic way of adding trial days to your program
'Use this code as you like
'By DreamVB
'http://www.bm-it-solutions.co.uk/

Private TrialVal As Integer
Private iVal As Integer
Private Const MaxDays As Integer = 30
Private Filename As String

Private Function FixPath(ByVal lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Private Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Sub PutByte(ByteVal As Byte)
Dim fp As Long
Dim mByte As Byte
    'Puts a single byte to a file
    fp = FreeFile
    Open Filename For Binary As #fp
        Put #fp, 1, ByteVal
    Close #fp
End Sub

Private Function GetByte() As Byte
Dim fp As Long
Dim mByte As Byte
    'Returns a single byte form a file
    fp = FreeFile
    Open Filename For Binary As #fp
        Get #fp, , mByte
    Close #fp
    
    GetByte = mByte
End Function

Private Sub WriteDateToFile()
Dim fp As Long
Dim dDate As Date
Dim mByte As Byte

    fp = FreeFile
    'Writes date to the file
    Open Filename For Binary As #fp
        dDate = (Date + 30)
        mByte = 1
        Put #fp, , mByte
        Put #fp, , dDate
    Close #fp
    
End Sub

Private Sub cmdCancel_Click()
    Call Unload(frmTrial)
End Sub

Private Sub cmdTry_Click()
    
    If (iVal > MaxDays) Then
        MsgBox "Your trail preiord has ended", vbInformation, "Trial Finsihed"
    Else
        MsgBox "Load your program here.", vbInformation, frmTrial.Caption
    End If
    
End Sub

Private Sub Form_Load()
Dim Date1 As Date
Dim Date2 As Date
Dim fp As Long
Dim mByte As Byte

    'Trial Date file, better hideing this in System folder or maybe at the end of a file
    'so no ones knows were it is
    
    Filename = FixPath(App.Path) & "trial.txt"
    
    'Write trial date to file if not found
    If Not FindFile(Filename) Then Call WriteDateToFile

    'Open Date file and see what the date is
    fp = FreeFile
    Open Filename For Binary As #fp
        Get #fp, , mByte
        Get #fp, , Date2
    Close #fp
    
    lblTitle1.Caption = "You may use this program for " & MaxDays & " days"
    
    'Current Date
    Date1 = Date
    'Find out how many days we have left
    TrialVal = (MaxDays - DateDiff("d", Date1, Date2) + 1)
    iVal = TrialVal
    
    'Check to see if user trys to switch back the date
    If (TrialVal <= 0) Then
        TrialVal = MaxDays
        iVal = (MaxDays + 1)
        'Add finish flag
        Call PutByte(0)
    End If

    'Check to see if the user trys to put date forward
    If (TrialVal > MaxDays) Then
        TrialVal = MaxDays
        iVal = (MaxDays + 1)
        'Add finish flag
        Call PutByte(0)
    End If
    
    'Check if date change flag was added
    If (GetByte = 0) Then
        TrialVal = MaxDays
        iVal = (MaxDays + 1)
    End If
    
    'Update displays
    lblCount.Caption = "Day " & TrialVal & " of " & MaxDays
    pBar.Value = TrialVal
    

End Sub

Private Sub Form_Resize()
    Line1.X2 = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmTrial = Nothing
End Sub

Private Sub pTop_Resize()
    Line1.X2 = pTop.ScaleWidth
    
End Sub
