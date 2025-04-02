VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnApply_Click()
    Me.Hide ' Hide the form after applying settings
    InsertPageNumbers Me
End Sub

Private Sub UserForm_Initialize()
    ' Set default values
    cmbFont.Text = "Arial"
    txtFontSize.Text = "12"
    optCenter.Value = True ' Default position: Bottom-Center
    txtPrefixSuffix.Text = "" ' Default: No prefix/suffix
End Sub

Private Sub btnCancel_Click()
    Unload Me ' Close form
End Sub

