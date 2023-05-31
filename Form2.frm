VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Details | Add a Record"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5265
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   435
      Left            =   3480
      TabIndex        =   8
      Top             =   2220
      Width           =   1455
   End
   Begin VB.TextBox TxtId 
      Appearance      =   0  'Flat
      DataField       =   "student_id"
      DataSource      =   "AdoStudent"
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      DataField       =   "name"
      DataSource      =   "AdoStudent"
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox TxtSubject 
      Appearance      =   0  'Flat
      DataField       =   "subject"
      DataSource      =   "AdoMarks"
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox TxtMarks 
      Appearance      =   0  'Flat
      DataField       =   "marks"
      DataSource      =   "AdoMarks"
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Marks"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.AddItem
    Form1.UpdateItem
End Sub

