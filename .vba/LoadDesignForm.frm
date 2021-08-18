VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoadDesignForm 
   Caption         =   "Load Design"
   ClientHeight    =   5887
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   4305
   OleObjectBlob   =   "LoadDesignForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoadDesignForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

LoadDesignForm.Hide

End Sub

Private Sub Listbox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    For i = 0 To Me.ListBox1.ListCount - 1
      If Me.ListBox1.Selected(i) Then
        LoadDesignForm.Hide
      End If
    Next
End Sub
