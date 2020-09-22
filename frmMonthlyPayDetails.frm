VERSION 5.00
Begin VB.Form frmMonthlyPayDetails 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MonthlyPayDetails"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "MonthlyPayDetails"
   Begin VB.CommandButton cndExit 
      Caption         =   "Exit Program"
      Height          =   375
      Left            =   6960
      TabIndex        =   42
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdReturnMenu 
      Caption         =   "Return to Menu"
      Height          =   375
      Left            =   5400
      TabIndex        =   41
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox txtSH 
      Height          =   375
      Left            =   7440
      TabIndex        =   38
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtOHH 
      Height          =   375
      Left            =   7440
      TabIndex        =   37
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtDT 
      Height          =   375
      Left            =   7440
      TabIndex        =   36
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate Pay and Deductions"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   2280
      Width           =   4215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3840
      TabIndex        =   31
      Tag             =   "&Update"
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   2640
      TabIndex        =   30
      Tag             =   "&Refresh"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Tag             =   "&Delete"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Tag             =   "&Add"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "I:\IT\Gnvq ICT\New Int GNVQ ICT\Programming\Pay2.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MonthlyPayDetails"
      Top             =   7020
      Width           =   8775
   End
   Begin VB.TextBox txtFields 
      DataField       =   "YTD_Netpay"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   13
      Left            =   6480
      TabIndex        =   27
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "YTD_Pension"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   12
      Left            =   6480
      TabIndex        =   25
      Top             =   4485
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "YTD_NI"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   11
      Left            =   6480
      TabIndex        =   23
      Top             =   4155
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "YTD_Tax"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   10
      Left            =   6480
      TabIndex        =   21
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "YTD_Gross Pay"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   6480
      TabIndex        =   19
      Top             =   3525
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Net Pay"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   17
      Top             =   4755
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Pension"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   15
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NI Deducted"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   13
      Top             =   4125
      Width           =   1935
   End
   Begin VB.TextBox txtTaxDeducted 
      DataField       =   "Tax Deducted"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   3795
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Gross Pay"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   9
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Hourly Rate"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Hours Worked"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Month No"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EmployeeNumber"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   40
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Time and a Half"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   39
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Double Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   35
      Top             =   480
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   4440
      X2              =   7440
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   2640
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Year To Date Pay Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   34
      Top             =   3000
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Pay Details "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "YTD_Netpay:"
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
      Left            =   4440
      TabIndex        =   26
      Tag             =   "YTD_Netpay:"
      Top             =   4815
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "YTD_Pension:"
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
      Left            =   4440
      TabIndex        =   24
      Tag             =   "YTD_Pension:"
      Top             =   4500
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "YTD_NI:"
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
      Left            =   4440
      TabIndex        =   22
      Tag             =   "YTD_NI:"
      Top             =   4185
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "YTD_Tax:"
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
      Left            =   4440
      TabIndex        =   20
      Tag             =   "YTD_Tax:"
      Top             =   3855
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "YTD_Gross Pay:"
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
      Left            =   4440
      TabIndex        =   18
      Tag             =   "YTD_Gross Pay:"
      Top             =   3540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Pay:"
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
      Left            =   240
      TabIndex        =   16
      Tag             =   "Net Pay:"
      Top             =   4785
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Pension:"
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
      Left            =   240
      TabIndex        =   14
      Tag             =   "Pension:"
      Top             =   4455
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "NI Deducted:"
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
      Left            =   240
      TabIndex        =   12
      Tag             =   "NI Deducted:"
      Top             =   4140
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Deducted:"
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
      Left            =   240
      TabIndex        =   10
      Tag             =   "Tax Deducted:"
      Top             =   3825
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Pay:"
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
      Left            =   240
      TabIndex        =   8
      Tag             =   "Gross Pay:"
      Top             =   3495
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Hourly Rate:"
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
      Left            =   2400
      TabIndex        =   6
      Tag             =   "Hourly Rate:"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Hours Worked:"
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
      Left            =   240
      TabIndex        =   4
      Tag             =   "Hours Worked:"
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Month No:"
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
      Left            =   2400
      TabIndex        =   2
      Tag             =   "Month No:"
      Top             =   120
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
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMonthlyPayDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HW, DTH, THH, SH, GP As Single

Private Sub cmdAdd_Click()
    Data1.Recordset.AddNew
End Sub


Private Sub cmdCalc_Click()
CalcGrossPay
CalcTax
CalcNI
CalcPension
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


Private Sub cmdRefresh_Click()
    'this is really only needed for multi user apps
    Data1.Refresh
End Sub


Private Sub cmdReturnMenu_Click()
frmMonthlyPayDetails.Hide
frmMenu.Show
End Sub

Private Sub cmdUpdate_Click()
    Data1.UpdateRecord
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub




Private Sub cndExit_Click()
End
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




Public Sub CalcGrossPay()

DTH = 0  'Initialise all the variables, Double Time Hours
OHH = 0 'Initialise all the variables, One and a Half Time Hours
SH = 0  'Initialise all the variables, Standard Hours
HR = 0  'Initialise all the variables, Hourly Rate
Let HW = Val(txtFields(2)) ' Load hours Worked
Let HR = Val(txtFields(3)) 'Load hourly rate
If HW > 188 Then ' All over 188 at Double Time
DTH = HW - 188
OHH = 40
SH = 148
ElseIf HW > 148 Then ' Between 148 and 188 at Time and a half
DTH = 0
OHH = HW - 148
SH = 148
Else                ' Remainder at Standard Rate
SH = HW
End If

GP = 2 * DTH * HR + 1.5 * OHH * HR + SH * HR

txtFields(4).Text = GP
txtDT = DTH
txtOHH = OHH
txtSH = SH
End Sub

Public Sub CalcTax()
Dim BandATax, BandBTax, BandCTax, Tax, TaxablePay, TaxFreePay As Single
TaxablePay = GP - TaxFreePay
If TaxablePay > 2000 Then

        BandATax = 125 * 10 / 100 'First £125 taxed at 10%

        BandBTax = 500 * 15 / 100 'Next £500 taxed at 15%

        BandCTax = (TaxablePay - 2000) * 40 / 100 'Over £2000 taxed at 40%
        
ElseIf TaxablePay > 125 Then

        BandATax = 125 * 10 / 100 'First £125 taxed at 10%
        
        BandBTax = (TaxablePay - 125) * 15 / 100 'Next £500 taxed at 15%
        
        BandCTax = 0
                
Else

        BandATax = TaxablePay * 10 / 100 ' All at 10%
        BandBTax = 0
        BandCTax = 0
End If

Tax = BandATax + BandBTax + BandCTax

txtTaxDeducted = Tax
End Sub

Public Sub CalcNI()

End Sub

Public Sub CalcPension()

End Sub
