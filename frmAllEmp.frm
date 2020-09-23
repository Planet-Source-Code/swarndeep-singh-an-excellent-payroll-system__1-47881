VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmAllEmp 
   Caption         =   "All Details"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmAllEmp.frx":0000
      Height          =   1200
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   2117
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      DataMember      =   "comALL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmAllEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
prcDisableMenu
End Sub

Private Sub Form_Resize()
MSHFlexGrid1.Height = Me.Height - 400
MSHFlexGrid1.Width = Me.Width - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
prcEnabledMenu
End Sub

