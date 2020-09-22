VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFastCat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "String Concatenation Bench Tester"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMethods 
      Caption         =   "Methods To Be Benchmarked"
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   7935
      Begin VB.CheckBox chkMethods 
         Caption         =   "Temporary File Method"
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   13
         Top             =   960
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "Paged Buffering Method"
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   12
         Top             =   2400
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "Byte Array Buffering Method"
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   11
         Top             =   1920
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "String Buffering Method"
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   1440
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkMethods 
         Caption         =   "Serial Concatenation (str1 = str1 && str2)"
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Value           =   1  'Checked
         Width           =   4335
      End
   End
   Begin MSComctlLib.ListView lvwResults 
      Height          =   1675
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2963
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Method"
         Object.Width           =   6774
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Execution Time"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Length (characters)"
         Object.Width           =   3546
      EndProperty
   End
   Begin VB.Frame fraConditions 
      Caption         =   "Concatenation Conditions"
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txtAppend 
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Text            =   "Paged Buffer, Faster By Design"
         Top             =   360
         Width           =   5295
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run"
         Height          =   375
         Left            =   6240
         TabIndex        =   4
         Top             =   870
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "10000"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "String to be Appended:"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   427
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "times."
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   975
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Concatenate"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   982
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   135
      Left            =   3720
      ScaleHeight     =   75
      ScaleWidth      =   3315
      TabIndex        =   14
      Top             =   6000
      Width           =   3375
   End
End
Attribute VB_Name = "frmFastCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FastCat As cFastCat
Private Timer As cPrecisionTimer
Private SpeedString As cSpeedString
Private Concat As clsConcat
Private FileConcat As cFileConcat

Private Sub cmdRun_Click()
    Dim lngCount As Long
    Dim lngTimes As Long
    Dim strAppend As String
    Dim strBuffer As String
    Dim lngResult As Long
    Dim itmListItem As ListItem
    
    
    lngTimes = Val(Text1.Text)
    strAppend = txtAppend
    lvwResults.ListItems.Clear
    
    Screen.MousePointer = vbHourglass
    
    ' Serial Concatenation
    If (chkMethods(0) = vbChecked) Then
        Timer.ResetTimer
        For lngCount = 1 To lngTimes
            strBuffer = strBuffer & strAppend
        Next
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Serial Concatenation"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(Len(strBuffer), "#,###")
    End If
    
    ' Temporary File Concatenation
    If (chkMethods(1) = vbChecked) Then
        Timer.ResetTimer
        For lngCount = 1 To lngTimes
            FileConcat.Append strAppend
        Next
        strBuffer = FileConcat.Value ' Defaults the same way as textboxes, to the value
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Temporary File Concatenation"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(Len(strBuffer), "#,###")
        
        ' Clean up
        FileConcat.Flush
    End If
    
    ' Standard Buffering Concatenation
    If (chkMethods(2) = vbChecked) Then
        Timer.ResetTimer
        For lngCount = 1 To lngTimes
            Concat.SConcat strAppend
        Next
        strBuffer = Concat.GetString
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Conventional String Buffering"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(Len(strBuffer), "#,###")

        ' Clean up
        Concat.ReInit
    End If
    
    ' Byte Array Buffering Concatenation
    If (chkMethods(3) = vbChecked) Then
        Timer.ResetTimer
        For lngCount = 1 To lngTimes
            SpeedString.Append strAppend
        Next
        strBuffer = SpeedString.Data
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Byte Array Buffering"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(Len(strBuffer), "#,###")
        
        ' Clean up
        SpeedString.Reset
    End If
    
    ' Paged Buffering Concatenation
    If (chkMethods(4) = vbChecked) Then
        Timer.ResetTimer
        For lngCount = 1 To lngTimes
            FastCat.Append strAppend
        Next
        strBuffer = FastCat ' Defaults the same way as textboxes, to the value
        Timer.StopTimer
        lngResult = Timer.Elapsed
        
        ' Display the results
        Set itmListItem = lvwResults.ListItems.Add
        itmListItem.Text = "Paged Buffering"
        itmListItem.SubItems(1) = Format$(lngResult, "#,###") & " µs"
        itmListItem.SubItems(2) = Format$(Len(strBuffer), "###,###")
        
        ' Clean up
        FastCat.Flush
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Set Timer = New cPrecisionTimer
    Set FastCat = New cFastCat
    Set SpeedString = New cSpeedString
    Set Concat = New clsConcat
    Set FileConcat = New cFileConcat
    
     ' Configure the buffer
    FastCat.BufferPageSize = 4096
    FastCat.BufferPages = 99999
    
    ' Beautify the listview
    lvwResults.FullRowSelect = True
    SetListViewLedger lvwResults, vbLedgerLightBlue, vbLedgerYellow, sizeNone
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Timer = Nothing
    Set FastCat = Nothing
    Set SpeedString = Nothing
    Set Concat = Nothing
    Set FileConcat = Nothing
End Sub

