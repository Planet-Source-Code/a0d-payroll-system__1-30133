VERSION 5.00
Begin VB.Form frmEmployees 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employees"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Employees"
   Begin VB.CommandButton cmdMenu 
      Caption         =   "Return to Menu"
      Height          =   495
      Left            =   5280
      TabIndex        =   38
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton cmdMonthlyPayDets 
      Caption         =   "Monthly Pay Details"
      Height          =   495
      Left            =   3360
      TabIndex        =   37
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   2280
      TabIndex        =   36
      Tag             =   "&Update"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   1200
      TabIndex        =   35
      Tag             =   "&Delete"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   120
      TabIndex        =   34
      Tag             =   "&Add"
      Top             =   6360
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "Pay2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Employees"
      Top             =   7425
      Width           =   7695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Notes"
      DataSource      =   "Data1"
      Height          =   675
      Index           =   16
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   5150
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "SupervisorID"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   15
      Left            =   3000
      TabIndex        =   31
      Top             =   4840
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Standard Hourly Rate"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   14
      Left            =   3000
      TabIndex        =   29
      Top             =   4520
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Date Hired"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   13
      Left            =   3000
      TabIndex        =   27
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Birthdate"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   12
      Left            =   3000
      TabIndex        =   25
      Top             =   3880
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "WorkPhone"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   11
      Left            =   3000
      MaxLength       =   30
      TabIndex        =   23
      Top             =   3560
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "HomePhone"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   10
      Left            =   3000
      MaxLength       =   30
      TabIndex        =   21
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PostalCode"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   3000
      MaxLength       =   20
      TabIndex        =   19
      Top             =   2920
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "City\Town"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2600
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address Line 2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address Line 1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1960
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "National Insurance No"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   3000
      MaxLength       =   30
      TabIndex        =   11
      Top             =   1640
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Last Name"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "First Name"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1000
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Title"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   5
      Top             =   680
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "DepartmentID"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EmployeeNumber"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   40
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   120
      TabIndex        =   32
      Tag             =   "Notes:"
      Top             =   5180
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "SupervisorID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   30
      Tag             =   "SupervisorID:"
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Hourly Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   28
      Tag             =   "Standard Hourly Rate:"
      Top             =   4545
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Hired:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   26
      Tag             =   "Date Hired:"
      Top             =   4220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthdate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Tag             =   "Birthdate:"
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "WorkPhone:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Tag             =   "WorkPhone:"
      Top             =   3580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "HomePhone:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   20
      Tag             =   "HomePhone:"
      Top             =   3260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "PostalCode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Tag             =   "PostalCode:"
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "City\Town:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Tag             =   "City\Town:"
      Top             =   2620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Tag             =   "Address Line 2:"
      Top             =   2300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Line 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Tag             =   "Address Line 1:"
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "National Insurance No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Tag             =   "National Insurance No:"
      Top             =   1665
      Width           =   2415
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Tag             =   "Last Name:"
      Top             =   1340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Tag             =   "First Name:"
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Tag             =   "Title:"
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "DepartmentID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Tag             =   "DepartmentID:"
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "EmployeeNumber:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "EmployeeNumber:"
      Top             =   60
      Width           =   2055
   End
End
Attribute VB_Name = "frmEmployees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    Data1.Recordset.AddNew
End Sub


Private Sub cmdDelete_Click()
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    With Data1.Recordset
        .Delete
        .MoveNext
        If .EOF Then .MoveLast
    End With
End Sub


Private Sub cmdMenu_Click()
frmEmployees.Hide
frmMenu.Show
End Sub

Private Sub cmdMonthlyPayDets_Click()
frmEmployees.Hide
frmMonthlyPayDetails.Show


End Sub



Private Sub cmdUpdate_Click()
    Data1.UpdateRecord
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub


Private Sub cmdGrid_Click()
    On Error GoTo cmdGrid_ClickErr


    Dim f As New frmDataGrid
    Set f.Data1.Recordset = Data1.Recordset
    f.Caption = Me.Caption & " Grid"
    f.Show


    Exit Sub
cmdGrid_ClickErr:
End Sub


Private Sub Data1_Error(DataErr As Integer, Response As Integer)
    'This is where you would put error handling code
    'If you want to ignore errors, comment out the next line
    'If you want to trap them, add code here to handle them
    MsgBox "Data error event hit err:" & Error$(DataErr)
    Response = 0  'throw away the error
End Sub


Private Sub Data1_Reposition()
    Screen.MousePointer = vbDefault
    On Error Resume Next
    'This will display the current record position
    'for dynasets and snapshots
    Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
    'for the table object you must set the index property when
    'the recordset gets created and use the following line
    'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub


Private Sub Data1_Validate(Action As Integer, Save As Integer)
    'This is where you put validation code
    'This event gets called when the following actions occur
    Select Case Action
        Case vbDataActionMoveFirst
        Case vbDataActionMovePrevious
        Case vbDataActionMoveNext
        Case vbDataActionMoveLast
        Case vbDataActionAddNew
        Case vbDataActionUpdate
        Case vbDataActionDelete
        Case vbDataActionFind
        Case vbDataActionBookmark
        Case vbDataActionClose
            Screen.MousePointer = vbDefault
    End Select
    Screen.MousePointer = vbHourglass
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub



