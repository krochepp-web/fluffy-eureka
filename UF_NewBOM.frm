VERSION 5.00
Begin VB.UserForm UF_NewBOM
   Caption         =   "New BOM"
   ClientHeight    =   2700
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel
      Caption         =   "Cancel"
      Height          =   330
      Left            =   3480
      TabIndex        =   9
      Top             =   2250
      Width           =   960
   End
   Begin VB.CommandButton cmdCreate
      Caption         =   "Create BOM"
      Height          =   330
      Left            =   2340
      TabIndex        =   8
      Top             =   2250
      Width           =   1080
   End
   Begin VB.TextBox txtTADesc
      Height          =   390
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   3120
   End
   Begin VB.TextBox txtTARev
      Height          =   300
      Left            =   1320
      TabIndex        =   5
      Top             =   1260
      Width           =   1200
   End
   Begin VB.TextBox txtTAPN
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   900
      Width           =   3120
   End
   Begin VB.TextBox txtTAID
      Height          =   300
      Left            =   1320
      TabIndex        =   1
      Top             =   540
      Width           =   3120
   End
   Begin VB.Label lblTADesc
      Caption         =   "Description"
      Height          =   240
      Left            =   180
      TabIndex        =   6
      Top             =   1740
      Width           =   1020
   End
   Begin VB.Label lblTARev
      Caption         =   "TA Rev"
      Height          =   240
      Left            =   180
      TabIndex        =   4
      Top             =   1290
      Width           =   900
   End
   Begin VB.Label lblTAPN
      Caption         =   "TA Part Number"
      Height          =   240
      Left            =   180
      TabIndex        =   2
      Top             =   930
      Width           =   1200
   End
   Begin VB.Label lblTAID
      Caption         =   "TA ID"
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   570
      Width           =   900
   End
End
Attribute VB_Name = "UF_NewBOM"
Option Explicit

Private mWb As Workbook

Public Sub InitForm(ByVal wb As Workbook)
    Set mWb = wb
    SetControlValue "txtTAID", vbNullString
    SetControlValue "txtTAPN", vbNullString
    SetControlValue "txtTARev", vbNullString
    SetControlValue "txtTADesc", vbNullString
End Sub

Private Sub cmdCreate_Click()
    On Error GoTo EH

    M_Data_BOMs_Entry.Create_BOM_For_Assembly_FromInputs _
        GetControlValue("txtTAID"), _
        GetControlValue("txtTAPN"), _
        GetControlValue("txtTARev"), _
        GetControlValue("txtTADesc")

    Unload Me
    Exit Sub

EH:
    MsgBox "New BOM creation failed." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbExclamation, "New BOM"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function GetControlValue(ByVal controlName As String) As String
    Dim ctl As MSForms.Control
    Set ctl = GetNamedControl(controlName)
    GetControlValue = Trim$(CStr(ctl.Value))
End Function

Private Sub SetControlValue(ByVal controlName As String, ByVal value As String)
    Dim ctl As MSForms.Control
    Set ctl = GetNamedControl(controlName)
    ctl.Value = value
End Sub

Private Function GetNamedControl(ByVal controlName As String) As MSForms.Control
    Dim ctl As MSForms.Control

    On Error Resume Next
    Set ctl = Me.Controls(controlName)
    On Error GoTo 0

    If ctl Is Nothing Then
        Err.Raise vbObjectError + 9100, "UF_NewBOM.GetNamedControl", _
                  "Missing control '" & controlName & "' on form UF_NewBOM."
    End If

    Set GetNamedControl = ctl
End Function
