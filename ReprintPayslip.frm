VERSION 5.00
Begin VB.Form frmReprintPayslip 
   Caption         =   "f"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   7080
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3600
      Width           =   3135
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CommandButton frmReprintPayslip 
      Caption         =   "Reprint Pay Slip"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "To Date"
      Height          =   195
      Left            =   5160
      TabIndex        =   4
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From Date"
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   360
      Width           =   165
   End
End
Attribute VB_Name = "frmReprintPayslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Call prcCallAllEmp(CStr(Combo2.Text), CStr(Combo3.Text))

Unload Me
End Sub

Private Sub Form_Load()
If typeID = "c" Then
    frmReprintPayslip.Enabled = False
    Combo1.Enabled = False
End If
prcConnection

rs.Open "select * from PaySlipGenerated", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
If typeID = "a" Then
Combo1.AddItem rs("payslipid")
ElseIf typeID = "b" Then
Combo1.AddItem rs("empno") + " " + rs("name")
End If
rs.MoveNext
Loop
rs.Close
rs.Open "Select distinct(month) from payslipgenerated", con, adOpenDynamic, adLockOptimistic
Do While Not rs.EOF
Combo2.AddItem Format(rs(0), "mmm-yyyy")
Combo3.AddItem Format(rs(0), "mmm-yyyy")
rs.MoveNext
Loop
rs.Close
prcDisableMenu
End Sub

Private Sub Form_Terminate()
prcEnabledMenu
End Sub

Private Sub frmReprintPayslip_Click()
If typeID = "a" Then
prcCallPaySlip (Combo1.Text)
ElseIf typeID = "b" Then
prcCallempPaySlip (Mid(Combo1.Text, 1, 4))
End If
typeID = ""
Unload Me
End Sub
