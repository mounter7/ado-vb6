VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtSearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.CommandButton CmdMoveLast 
      Caption         =   ">|"
      Height          =   495
      Left            =   5040
      TabIndex        =   16
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdMoveNext 
      Caption         =   ">"
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdMovePrev 
      Caption         =   "<"
      Height          =   495
      Left            =   3840
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdMoveFirst 
      Caption         =   "|<"
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox TxtBirthday 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   4
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox TxtId 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthday"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================================================

Dim DB As DAO.Database
Dim RS As DAO.Recordset

Dim AppName As String

Private Sub CmdAdd_Click()
    Add
End Sub

Private Sub CmdDelete_Click()
    Delete
End Sub

Private Sub CmdMoveFirst_Click()
    MoveFirst
End Sub

Private Sub CmdMoveLast_Click()
    MoveLast
End Sub

Private Sub CmdMoveNext_Click()
    MoveNext
End Sub



' Functions =============================================================================

Public Sub Start()
    Set DB = OpenDatabase("studentDB.mdb")
    Set RS = DB.OpenRecordset("SELECT * FROM students;")
    AppName = "Student Details"
    Form1.Caption = AppName
End Sub

Private Sub CmdMovePrev_Click()
    MovePrev
End Sub

Private Sub CmdNew_Click()
    Clear
End Sub

Private Sub CmdUpdate_Click()
    Update
End Sub

Private Sub Form_Load()
    Start
    LoadData
End Sub

Public Sub LoadData()
    TxtId.Text = RS!ID
    TxtName.Text = RS!Name
    TxtAddress.Text = RS!address
    TxtBirthday.Text = RS!birthday
End Sub

Public Sub GetData()
    RS!ID = TxtId.Text
    RS!Name = TxtName.Text
    RS!address = TxtAddress.Text
    RS!birthday = TxtBirthday.Text
End Sub

Public Sub MoveLast()
    RS.MoveLast
    LoadData
End Sub

Public Sub MoveFirst()
    RS.MoveFirst
    LoadData
End Sub

Public Sub MoveNext()
    RS.MoveNext
    If RS.EOF Then
        MsgBox "This is the last record.", vbInformation, AppName
        RS.MoveLast
    Else
        LoadData
    End If
End Sub

Public Sub MovePrev()
    RS.MovePrevious
    If RS.BOF Then
        MsgBox "This is the first record.", vbInformation, AppName
        RS.MoveFirst
    Else
        LoadData
    End If
End Sub

Public Sub Clear()
    TxtId.Text = ""
    TxtName.Text = ""
    TxtAddress.Text = ""
    TxtBirthday.Text = ""
    TxtId.SetFocus
End Sub

' CRUD ===============================================================
Public Sub Add()
    Dim ID As String
    ID = TxtId.Text
    RS.FindFirst "ID = '" & ID & "'"
    
    If RS.NoMatch = False Then
        MsgBox "Record has already added.", vbInformation, AppName
    Else
        RS.AddNew
        GetData
        RS.Update
        MsgBox "Record has added successfully.", vbInformation, AppName
    End If
End Sub

Public Sub Update()
    RS.Edit
    GetData
    RS.Update
    MsgBox "Record has updated successfully.", vbInformation, AppName
End Sub

Public Sub Delete()
    Dim Msg As String
    Msg = MsgBox("Are you sure?", vbInformation + vbYesNo, AppName)
    
    If Msg = vbYes Then
        RS.Delete
        Clear
        TxtId.SetFocus
    Else
        CmdNew.SetFocus
    End If
End Sub

Private Sub TxtSearch_Change()
    Dim Str As String
    Str = TxtSearch.Text
    RS.FindFirst "ID = '" & Str & "'"
    
    If RS.NoMatch = False Then
        LoadData
    End If
End Sub
