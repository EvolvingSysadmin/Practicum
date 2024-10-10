VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} testForm 
   Caption         =   "UserForm1"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10515
   OleObjectBlob   =   "testForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "testForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub submitButton_Click()
    ' Get the input value from the organizationInput text box
    Dim organization As String
    
    organization = organizationInput.Value
    
    ' Insert the value into multiple bookmarks
    If ActiveDocument.Bookmarks.Exists("organization") Then
        ActiveDocument.Bookmarks("organization").Range.Text = organization
    End If
    
    If ActiveDocument.Bookmarks.Exists("organization2") Then
        ActiveDocument.Bookmarks("organization2").Range.Text = organization
    End If
    
    If ActiveDocument.Bookmarks.Exists("organization3") Then
        ActiveDocument.Bookmarks("organization3").Range.Text = organization
    End If
    
    ' Close the form
    Unload Me
End Sub

