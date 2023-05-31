VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Details"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc AdoMarks 
      Height          =   330
      Left            =   240
      Top             =   2280
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentDB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentDB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM Marks;"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   600
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc AdoStudent 
      Height          =   375
      Left            =   240
      Top             =   2760
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentDB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=studentDB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Student"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox TxtMarks 
      Appearance      =   0  'Flat
      DataField       =   "marks"
      DataSource      =   "AdoMarks"
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox TxtSubject 
      Appearance      =   0  'Flat
      DataField       =   "subject"
      DataSource      =   "AdoMarks"
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      DataField       =   "name"
      DataSource      =   "AdoStudent"
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox TxtId 
      Appearance      =   0  'Flat
      DataField       =   "student_id"
      DataSource      =   "AdoStudent"
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Marks"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =================================== Ravindu Madhushankha =============================================
' =================================== https://github.com/mounter7/ado-vb6/ =============================

Public AppName As String

Sub Form_Load()
    AppName = "Student Details"
End Sub

' Selecting a record from tables by ID
Private Sub AdoStudent_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If AdoStudent.Recordset.BOF = False And AdoStudent.Recordset.EOF = False Then
        AdoMarks.RecordSource = "SELECT * FROM Marks WHERE student_id = '" & AdoStudent.Recordset!student_id & "'"
        AdoMarks.Refresh
    End If
    
    AdoStudent.Caption = AdoStudent.Recordset.AbsolutePosition & " of " & AdoStudent.Recordset.RecordCount
End Sub

' Adding new record to the student and marks tables
Private Sub CmdAdd_Click()
    ' Form2.Show
    AddItem
End Sub

' Exit program
Private Sub CmdCancel_Click()
    End
End Sub

' Removing a record from student and marks tables
Private Sub CmdRemove_Click()
    Dim Confirm As Integer
    ' Confirmation
    Confirm = MsgBox("Are you sure?", vbYesNo, AppName)
    If Confirm = vbYes Then
        RemoveItem
    End If
End Sub

' Saving updated records
Private Sub CmdSave_Click()
    UpdateItem
End Sub



' Functions =============================================================================

' Add
Public Sub AddItem()
    AdoStudent.Recordset.AddNew
    AdoMarks.Recordset.AddNew
End Sub

' Update
Public Sub UpdateItem()
    MsgBox "Updated records have been saved.", vbOKOnly, AppName
    AdoMarks.Recordset!student_id = TxtId.Text
    AdoStudent.Recordset.Update
    AdoMarks.Recordset.Update
End Sub

' Remove
Public Sub RemoveItem()
    AdoStudent.Recordset.Delete
    AdoMarks.Recordset.Delete
    
    MsgBox "The record has been deleted.", vbOKOnly, AppName
    AdoMarks.Refresh
    AdoStudent.Refresh
End Sub
