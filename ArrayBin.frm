VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmArrayBin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Using Arrays to Access Binary Files"
   ClientHeight    =   5880
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   8880
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDL 
      Left            =   3240
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   240
      Left            =   5775
      TabIndex        =   11
      Top             =   5625
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   5550
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   582
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10504
            MinWidth        =   5080
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      Height          =   30
      Left            =   -300
      TabIndex        =   9
      Top             =   0
      Width           =   9615
   End
   Begin VB.Timer tmrInfo 
      Interval        =   5000
      Left            =   3300
      Top             =   5175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Record List (Using Field 1 as caption)"
      Height          =   5340
      Left            =   75
      TabIndex        =   8
      Top             =   150
      Width           =   3315
      Begin VB.ListBox lstPut 
         Height          =   4815
         IntegralHeight  =   0   'False
         Left            =   150
         TabIndex        =   0
         Top             =   375
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Item Details (make changes to record below)"
      Height          =   5340
      Left            =   3525
      TabIndex        =   4
      Top             =   150
      Width           =   5265
      Begin VB.PictureBox picContainer 
         BackColor       =   &H8000000C&
         Height          =   2565
         Left            =   825
         ScaleHeight     =   2505
         ScaleWidth      =   4230
         TabIndex        =   13
         Top             =   2625
         Width           =   4290
         Begin VB.CommandButton cmdTopLeft 
            Height          =   240
            Left            =   3825
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Back to top left"
            Top             =   2250
            Width           =   240
         End
         Begin VB.HScrollBar vsrX 
            Height          =   240
            LargeChange     =   50
            Left            =   0
            SmallChange     =   10
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   2250
            Width           =   1815
         End
         Begin VB.VScrollBar vsrY 
            Height          =   1590
            LargeChange     =   50
            Left            =   3825
            SmallChange     =   10
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   240
         End
         Begin VB.PictureBox picImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1890
            Left            =   0
            ScaleHeight     =   126
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   226
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   0
            Width           =   3390
         End
      End
      Begin RichTextLib.RichTextBox rtfPut 
         Height          =   1365
         Index           =   2
         Left            =   825
         TabIndex        =   3
         Top             =   1125
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   2408
         _Version        =   327680
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"ArrayBin.frx":0000
      End
      Begin RichTextLib.RichTextBox rtfPut 
         Height          =   315
         Index           =   0
         Left            =   825
         TabIndex        =   1
         Top             =   375
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   327680
         BorderStyle     =   0
         Enabled         =   -1  'True
         TextRTF         =   $"ArrayBin.frx":00C2
      End
      Begin RichTextLib.RichTextBox rtfPut 
         Height          =   315
         Index           =   1
         Left            =   825
         TabIndex        =   2
         Top             =   750
         Width           =   4290
         _ExtentX        =   7567
         _ExtentY        =   556
         _Version        =   327680
         BorderStyle     =   0
         Enabled         =   -1  'True
         TextRTF         =   $"ArrayBin.frx":0184
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Field 5"
         Height          =   285
         Left            =   75
         TabIndex        =   12
         Top             =   2625
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Field 4"
         Height          =   285
         Left            =   75
         TabIndex        =   7
         Top             =   1125
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Field 3"
         Height          =   285
         Left            =   75
         TabIndex        =   6
         Top             =   750
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Field 2"
         Height          =   285
         Left            =   75
         TabIndex        =   5
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Database"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Database"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Database"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuRecords 
      Caption         =   "&Record"
      Begin VB.Menu mnuRecordsAdd 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuRecordsSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuRecordsDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRecordsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecordsFind 
         Caption         =   "&Find..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsAdd 
         Caption         =   "&Add Records..."
      End
      Begin VB.Menu mnuToolsCount 
         Caption         =   "&Count Records"
      End
      Begin VB.Menu mnuToolsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmArrayBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  --------------------------------------  '
'  Code by:        Jim Hunt                '
'  E-mail:         jim@huntcs.com          '
'                                          '
'  Enjoy!  If you make any modifications   '
'  or improvements to this code, I would   '
'  appreciate an e-mail with the changes   '
'  - or at least honourable mention in     '
'    your software release.                '
'  - or you can you can send me money :)   '
'  --------------------------------------  '

Option Explicit

'the following type declarations define the records
'and make the setting/records/recordcount/etc. public
'to the entire application

'User Defined Type (UDT) Declarations

'The record used to store TotalRecords in the binary file
Private Type PrgSettings
    NumberOfRecords As Long
End Type

'The actual database record format
Private Type Records
    Field1 As String
    Field2 As String
    Field3 As String
    Field4 As String
    Field5() As Long 'will be used to store the image data
End Type

Private RecordArray() As Records 'Stores the database records
Private TotalRecords As Long 'Keeps track of how many records are in the DB
Private TempArray() As Records 'Used for record deletion
Private FName As String
Private Counter As Long
Private Record As Records
Private RecCount As PrgSettings


'More declarations, etc.
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'API Declarations
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Sub cmdTopLeft_Click()
    'Return picture preview to top left
    picImage.Move 0, 0
    vsrX.Value = 0
    vsrY.Value = 0
End Sub

Private Sub Form_Load()
    Me.Refresh 'Show the form right away
    
    'move the progress bar
    Progress.Move Status.Panels(2).Left + 45, Status.Top + 60, Status.Panels(2).Width - 90, Status.Height - 90
    
    picImage.ScaleMode = 3 'Switch to Pixels
    
    'Position image preview controls
    picImage.Move 0, 0, picContainer.ScaleWidth, picContainer.ScaleHeight
    vsrY.Move picContainer.ScaleWidth - vsrY.Width, 0, vsrY.Width, picContainer.ScaleHeight - vsrX.Height
    vsrX.Move 0, picContainer.ScaleHeight - vsrX.Height, picContainer.ScaleWidth - vsrY.Width
    cmdTopLeft.Move picContainer.ScaleWidth - cmdTopLeft.Width, picContainer.ScaleHeight - cmdTopLeft.Height
    
    mnuFileNew_Click 'Clear the form
    
    ReDim RecordArray(0) As Records 'Initialize array
    
    'Open the file "binarytest.dat" and if it's not there, just continue normally
    If Dir(App.Path & "\binarytest.dat") <> "" Then
        mnuFileNew_Click
    End If
    
    Status.Panels(1).Text = "Ready"
    
End Sub

Private Function RefillListBox()
    'Clears, then refills using the array data
    Dim RecordNo As Integer
    lstPut.Clear
    For RecordNo = 0 To TotalRecords - 1
        lstPut.AddItem RecordArray(RecordNo).Field1
    Next
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Free up the memory used for records
    Erase RecordArray
End Sub

Private Sub lstPut_Click()
    On Error GoTo ErrHandler
    
    'Show field contents for selected record (Field 1 is displayed in listbox)
    rtfPut(0).Text = RecordArray(lstPut.ListIndex).Field2
    rtfPut(1).Text = RecordArray(lstPut.ListIndex).Field3
    rtfPut(2).Text = RecordArray(lstPut.ListIndex).Field4
    
    'Variables to track picture preview coordinates
    Dim x As Long
    Dim y As Long
    
    'move as many control properties as possible to variables
    'since variables are much faster than control properties in loops!
    Dim LowerBoundX As Integer
    Dim UpperBoundX As Integer
    Dim LowerBoundY As Integer
    Dim UpperBoundY As Integer
    Dim CurrentRecordNo As Integer
    Dim PicHDC As Integer
    
    LowerBoundX = LBound(RecordArray(lstPut.ListIndex).Field5, 1)
    UpperBoundX = UBound(RecordArray(lstPut.ListIndex).Field5, 1)
    LowerBoundY = LBound(RecordArray(lstPut.ListIndex).Field5, 2)
    UpperBoundY = UBound(RecordArray(lstPut.ListIndex).Field5, 2)
    
    CurrentRecordNo = lstPut.ListIndex
    
    'Force the preview size to the container size
    picImage.Move 0, 0, UpperBoundX * Screen.TwipsPerPixelX, UpperBoundY * Screen.TwipsPerPixelY
    
    'Clear current image (I find loadpicture works better than cls)
    With picImage
        .Picture = LoadPicture()
        .CurrentX = 0
        .CurrentY = 0
    End With
    picImage.Print "Loading..."
    DoEvents
    
    PicHDC = picImage.hdc
    
    'Load image data
    For x = LowerBoundX To UpperBoundX
        For y = LowerBoundY To UpperBoundY
            SetPixel PicHDC, x, y, RecordArray(CurrentRecordNo).Field5(x, y)
        Next
    Next
    
    picImage.Refresh
    
    Exit Sub
    
ErrHandler:
    If Err.Number = 9 Then
        With picImage
            .Move 0, 0, picContainer.ScaleWidth, picContainer.ScaleHeight
            .Picture = LoadPicture()
            .CurrentX = 0
            .CurrentY = 0
        End With
        picImage.Print "No Picture - click to insert picture"
        Exit Sub 'There is no image data to display
    Else
        MsgBox Err.Number & vbCrLf & Err.Description 'Some other error occured
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    'Clear all text controls and reset array
    Dim Counter As Integer
    lstPut.Clear
    For Counter = 0 To 2
        rtfPut(Counter).Text = ""
    Next
    ReDim RecordArray(0) As Records
    TotalRecords = 0
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo ErrHandler
    
    FName = App.Path & "\binarytest.dat"
    
    'If the above filename is missing, exit
    If Dir(FName) = "" Then Exit Sub
    
    Me.MousePointer = vbHourglass
    
    Open FName For Binary As #1
    Get #1, , RecCount 'Retrieve number of records
    
    With Progress
        .Min = 0
        .Max = RecCount.NumberOfRecords
    End With
    
    'Update lblInfo & Timer
    tmrInfo.Enabled = False
    Status.Panels(1).Text = "Retrieving Records..."
    
    'Loop through the data and add it to the array
    For Counter = 0 To RecCount.NumberOfRecords - 1
        ReDim Preserve RecordArray(Counter) As Records 'Increase the size of the array
        Get #1, , RecordArray(Counter) 'Add record to array
        Progress.Value = Counter
    Next
    Close #1
    
    TotalRecords = RecCount.NumberOfRecords
    
    If Counter < 1 Then
        Exit Sub
    End If
    
    RefillListBox
    
    'Display message for 5 seconds
    tmrInfo.Enabled = False
    Status.Panels(1).Text = TotalRecords & " records retrieved successfully... Ready"
    tmrInfo.Enabled = True
    
    'Select the first item in the listbox
    lstPut.Selected(0) = True
    
    'Reset Progress bar
    Progress.Value = 0
    
    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Open failed!"
    Progress.Value = 0
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuFileSave_Click()
    
    'Exit if there aren't any items in the database
    If lstPut.ListCount = 0 Then Exit Sub
    
    'Save all your hard work!
    Dim FName As String
    Dim FNamebak As String
    Dim Counter As Long
    
    'Set filenames
    FName = App.Path & "\binarytest.dat"
    FNamebak = App.Path & "\binarytest.bak"
    
    'If there is a backup file already, delete it
    If Dir(FNamebak) <> "" Then
        Kill FNamebak
    End If
    
    'if the file is there, get rid of it
    'this is a good spot to add some code to backup the original file
    If Dir(FName) <> "" Then
        'Make a backup of the original file
        FileCopy FName, FNamebak
        Kill FName
    End If
    
    'Add RecordCount
    RecCount.NumberOfRecords = TotalRecords
    
    'Prepare progress bar
    If TotalRecords - 1 <> 0 Then
        Progress.Max = TotalRecords - 1
    Else
        Progress.Max = 1
    End If
    
    'Update lblInfo & Timer
    tmrInfo.Enabled = False
    Status.Panels(1).Text = "Saving Records..."
    
    'Output all data to the file
    Open FName For Binary As #1
    Put #1, , RecCount 'Write the record count
    
    'It's much easier to save the entire array directly (without looping),
    'but looping is required to implement a progress bar.
    For Counter = 0 To TotalRecords - 1
        Put #1, , RecordArray(Counter)
        Progress.Value = Counter
    Next
    Close #1
    
    Progress.Value = 0
    
    'Display message for 5 seconds
    Status.Panels(1).Text = TotalRecords & " records saved successfully... Ready"
    tmrInfo.Enabled = True
End Sub

Private Sub mnuHelpAbout_Click()
    Dim MSG As String
    MSG = MSG & "This example demonstrates how to use a" & vbCrLf
    MSG = MSG & "dynamic array to load the contents of a" & vbCrLf
    MSG = MSG & "binary file used as a simple flat-file database." & vbCrLf
    MSG = MSG & "" & vbCrLf
    MSG = MSG & "This type of data access will also reduce" & vbCrLf
    MSG = MSG & "your project distribution size by several" & vbCrLf
    MSG = MSG & "megabytes, since you won't need to include" & vbCrLf
    MSG = MSG & "DAO/ADO/RDO support files with your app!" & vbCrLf
    MSG = MSG & "" & vbCrLf
    MSG = MSG & "I often use this to save program settings" & vbCrLf
    MSG = MSG & "instead of using INI files or the registry." & vbCrLf
    MSG = MSG & "" & vbCrLf
    MSG = MSG & "If you find this project useful in any way," & vbCrLf
    MSG = MSG & "post your comment good or bad!" & vbCrLf
    MSG = MSG & "" & vbCrLf
    MSG = MSG & "You may use this however you see fit." & vbCrLf
    MSG = MSG & "" & vbCrLf
    MSG = MSG & "Created by Jim Hunt" & vbCrLf
    
    MsgBox MSG, vbInformation
End Sub

Private Sub mnuRecordsAdd_Click()
    Dim Response As String
    Dim Record As Records
    
    Response = InputBox("Enter a title for this record")
    If Response = "" Then Exit Sub
    
    'Add some data to the fields
    Record.Field1 = Response
    Record.Field2 = "Title: " & Response
    Record.Field3 = "8K of Data following..."
    
    'Open the 8Kb file from the example folder:
    'The following code from "Past Tips of the Week 1998"
    Dim Handle As Integer
    Dim TmpFile As String
    Dim FileString As String
    Handle = FreeFile
    
    'Add the contents of the following file to Field4
    TmpFile = App.Path & "\8Kb.txt"
    Open TmpFile For Binary As #Handle
    FileString = Space(FileLen(TmpFile))
    Get #Handle, , FileString
    Close #Handle
    
    Record.Field4 = FileString
    
    If TotalRecords = 0 Then
        ReDim RecordArray(0) As Records 'Simply reinitialize the array to prepare for new data
    Else
        ReDim Preserve RecordArray(TotalRecords) As Records 'Remember, TotalRecords is always 1 more than UBound(RecordArray, 1)
    End If

    TotalRecords = TotalRecords + 1 'Update number of records
    
    'Add the data from above to the array
    RecordArray(UBound(RecordArray)) = Record
    
    RefillListBox
    
    'Select the added item in the listbox
    lstPut.Selected(lstPut.NewIndex) = True
End Sub

Private Sub mnuRecordsDelete_Click()
    On Error GoTo ErrHandler
    
    'Check if anything selected in listbox
    If lstPut.ListIndex = -1 Then Exit Sub
    
    'Ask if it's okay
    Dim Response As Integer
    Response = MsgBox("You are about to delete the record: " & Chr(34) & RecordArray(lstPut.ListIndex).Field1 & Chr(34) & "." & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbYesNo + vbExclamation, "WARNING")
    If Response = 7 Then Exit Sub
    
    'Update lblInfo & Timer
    tmrInfo.Enabled = False
    Status.Panels(1).Text = "Deleting Record..."
    
    'With verification out of the way let's continue
    Dim Counter As Long
    Dim DeletedFlag As Boolean

    Me.MousePointer = vbHourglass

    'Prepare TempArray to receive data
    ReDim TempArray(TotalRecords - 1) As Records
        
    'Copy the contents of RecordArray, minus the deleted record
    For Counter = 0 To TotalRecords - 1
        If Counter = lstPut.ListIndex Then
            DeletedFlag = True 'Raise flag
        Else
            If DeletedFlag Then 'Move remaining records down by one to fill the gap
                TempArray(Counter - 1) = RecordArray(Counter)
            Else
                TempArray(Counter) = RecordArray(Counter)
            End If
        End If
    Next

    'Now initialize RecordArray and fill with TempArray values
    TotalRecords = TotalRecords - 1 'Update total records to show deletion
    If TotalRecords > 0 Then
        ReDim RecordArray(TotalRecords - 1)
    Else
        ReDim RecordArray(0)
        TotalRecords = 0
    End If
    
    'Start filling RecordArray
    For Counter = 0 To TotalRecords - 1
        RecordArray(Counter) = TempArray(Counter)
    Next Counter
    
    'Clear Field Boxes and refresh listbox
    For Counter = 0 To 2
        rtfPut(Counter).Text = ""
    Next
    
    'Remove item from listbox
    lstPut.RemoveItem (lstPut.ListIndex)
    
    ' There's no need to call RefillListBox since the listbox keeps track of what we need

    Me.MousePointer = vbDefault

    'Display message for 5 seconds
    tmrInfo.Enabled = False
    Status.Panels(1).Text = "Record deleted successfully... Ready"
    tmrInfo.Enabled = True
    Progress.Value = 0
    
    Exit Sub
    
ErrHandler:
    Me.MousePointer = vbDefault
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub mnuRecordsSave_Click()
    On Error GoTo ErrHandler
    
    'Make sure there is a selected record
     If lstPut.ListCount = 0 Then Exit Sub
    
    'update the record
    RecordArray(lstPut.ListIndex).Field2 = rtfPut(0).Text
    RecordArray(lstPut.ListIndex).Field3 = rtfPut(1).Text
    RecordArray(lstPut.ListIndex).Field4 = rtfPut(2).Text
    
    'Save image data to array
    Dim x As Long
    Dim y As Long
    
    'move as many control properties as possible to variables
    'since variables are much faster than control properties in loops!
    Dim UpperBoundX As Integer
    Dim UpperBoundY As Integer
    Dim CurrentRecordNo As Integer
    Dim PicHDC As Integer
    
    On Error GoTo ErrHandler
    
    UpperBoundX = picImage.ScaleWidth
    UpperBoundY = picImage.ScaleHeight
    
    CurrentRecordNo = lstPut.ListIndex
    
    ReDim RecordArray(lstPut.ListIndex).Field5(1 To picImage.ScaleWidth, 1 To picImage.ScaleHeight)
    
    Me.MousePointer = vbHourglass
    For x = 1 To UpperBoundX
        For y = 1 To UpperBoundY
            RecordArray(CurrentRecordNo).Field5(x, y) = GetPixel(picImage.hdc, x, y)
        Next
    Next
    Me.MousePointer = vbDefault

    Exit Sub
    
ErrHandler:
    MsgBox Err.Number & vbCrLf & Err.Description
    Me.MousePointer = vbDefault
End Sub

Private Sub mnuToolsAdd_Click()
    Dim Record As Records
    Dim Counter As Long
    Dim Response As String
    
    Dim RecordsToAdd As Long
    
    Response = InputBox("This will add records to the database" & vbCrLf & vbCrLf & "Enter a number")
    
    If Response = "" Then Exit Sub
    If CLng(Response) = 0 Then Exit Sub
    RecordsToAdd = CLng(Response)
    
    Progress.Max = RecordsToAdd
    
    If TotalRecords = 0 Then
        ReDim RecordArray(0) As Records 'Simply reinitialize the array to prepare for new data
    End If

    
    'Update lblInfo & Timer
    tmrInfo.Enabled = False
    Status.Panels(1).Text = "Working..."
    
    For Counter = 1 To RecordsToAdd
    
        'Add some data to the fields
        Record.Field1 = Counter
        Record.Field2 = "Title: " & Counter
        Record.Field3 = "8K of Data following..."
        
        'Open the 8Kb file from the example folder:
        'The following code from "Past Tips of the Week 1998"
        Dim Handle As Integer
        Dim TmpFile As String
        Dim FileString As String
        Handle = FreeFile
        TmpFile = App.Path & "\8Kb.txt"
        Open TmpFile For Binary As #Handle
        FileString = Space(FileLen(TmpFile))
        Get #Handle, , FileString
        Close #Handle
        
        Record.Field4 = FileString
        
        ReDim Preserve RecordArray((TotalRecords - 1) + Counter) As Records 'Remember, TotalRecords is always 1 more than UBound(RecordArray, 1)
    
        'Add the data from above to the array
        RecordArray(UBound(RecordArray)) = Record
    
        Progress.Value = Counter
    
    Next
    
    TotalRecords = TotalRecords + RecordsToAdd 'Update number of records
    
    If TotalRecords < 32767 Then
        RefillListBox
        'Select the added item in the listbox
        lstPut.Selected(lstPut.NewIndex) = True
    Else
        MsgBox "There are too many records to display in the listbox!"
    End If
    
    'Display message for 5 seconds
    tmrInfo.Enabled = False
    Status.Panels(1).Text = RecordsToAdd & " records added successfully... Ready"
    tmrInfo.Enabled = True
    Progress.Value = 0
End Sub

Private Sub mnuToolsCount_Click()
    'I just used this for testing
    If TotalRecords = 1 Then
        MsgBox "There is only " & TotalRecords & " record."
    Else
        MsgBox "There are " & TotalRecords & " records."
    End If
End Sub

Private Sub picContainer_Click()
    picImage_Click
End Sub

Private Sub picImage_Click()
    CDL.Filter = "Pictures(*.jpg;*.gif)|*.jpg;*.gif"
    CDL.DialogTitle = "Select Image"
    CDL.InitDir = App.Path
    CDL.ShowOpen
    If Not CDL.FileName = "" Then
        'Clear existing picture
        picImage.Picture = LoadPicture()
        'Load new picture
        picImage.Picture = LoadPicture(CDL.FileName)
    End If
End Sub

Private Sub picImage_Resize()
    'Setup scroll bars and image preview
    Dim MaxX As Integer
    Dim MaxY As Integer
    
    If picImage.Width > picContainer.ScaleWidth Then
        vsrX.Enabled = True
        vsrX.Max = (picImage.Width - picContainer.ScaleWidth + vsrX.Height) / Screen.TwipsPerPixelX
        vsrX.Value = 0
    Else
        vsrX.Enabled = False
    End If
    
    If picImage.Height > picContainer.ScaleHeight Then
        vsrY.Enabled = True
        vsrY.Max = (picImage.Height - picContainer.ScaleHeight + vsrY.Width) / Screen.TwipsPerPixelY
        vsrY.Value = 0
    Else
        vsrY.Enabled = False
    End If
    
    picImage.Move 0, 0
End Sub

Private Sub tmrInfo_Timer()
    ' This timer resets lblInfo to Ready
    Status.Panels(1).Text = "Ready"
    tmrInfo.Enabled = False
End Sub

Private Sub vsrX_Change()
    picImage.Left = -vsrX.Value * Screen.TwipsPerPixelX
End Sub

Private Sub vsrX_Scroll()
    picImage.Left = -vsrX.Value * Screen.TwipsPerPixelX
End Sub

Private Sub vsrY_Change()
    picImage.Top = -vsrY.Value * Screen.TwipsPerPixelY
End Sub

Private Sub vsrY_Scroll()
    picImage.Top = -vsrY.Value * Screen.TwipsPerPixelY
End Sub
