VERSION 5.00
Begin VB.MDIForm frmMainForm 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll System"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuEmployee 
         Caption         =   "Employee"
      End
      Begin VB.Menu mnuSetting 
         Caption         =   "Setting"
      End
   End
   Begin VB.Menu mnuPaySlip 
      Caption         =   "Payslip"
      Begin VB.Menu mnuPaySlipGenerate 
         Caption         =   "Generate Payslip"
      End
      Begin VB.Menu mnuRPrintPayslip 
         Caption         =   "Re-Print Payslip"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuName 
         Caption         =   "All Payslips of one Employee"
      End
      Begin VB.Menu mnuALl 
         Caption         =   "All Payslip Data"
      End
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuALl_Click()
typeID = "c"
Load frmReprintPayslip
frmReprintPayslip.Show
End Sub

Private Sub mnuEmployee_Click()
Load frmEmpDetails
frmEmpDetails.Show

End Sub

Private Sub mnuName_Click()
typeID = "b"
Load frmReprintPayslip
frmReprintPayslip.Show

End Sub

Private Sub mnuPaySlipGenerate_Click()
Load frmPaySlipGeneration
frmPaySlipGeneration.Show
End Sub

Private Sub mnuRPrintPayslip_Click()
typeID = "a"
Load frmReprintPayslip
frmReprintPayslip.Show
End Sub

Private Sub mnuSetting_Click()
Load frmSetting
frmSetting.Show
End Sub
