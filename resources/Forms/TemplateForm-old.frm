VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TemplateForm 
   Caption         =   "Cybersecurity Policy Template Form"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11175
   OleObjectBlob   =   "TemplateForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TemplateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OrganizationAddress_Change()

End Sub

Private Sub OrganizationAddressExample_Click()

End Sub

Private Sub OrganizationAddressLabel_Click()

End Sub

Private Sub OrganizationLabelExample_Click()

End Sub

Private Sub OrganizationName_Change()

End Sub

Private Sub OrganizationNameLabel_Click()

End Sub

Private Sub SubmitButton_Click()

    ' Get the input values from the text boxes
    Dim orgName As String
    Dim orgAddress As String
    Dim authValue As String  ' Renamed from Authority to authValue to avoid conflict
    Dim ownerValue As String ' Renamed to ownerValue for clarity
    Dim docNumber As String
    
    orgName = OrganizationName.Value
    orgAddress = OrganizationAddress.Value
    authValue = Authority.Value
    ownerValue = Owner.Value
    docNumber = DocumentNumber.Value
    
    ' Insert the organization name into multiple bookmarks
    If ActiveDocument.Bookmarks.Exists("OrganizationName1") Then
        ActiveDocument.Bookmarks("OrganizationName1").Range.Text = orgName
    End If
    If ActiveDocument.Bookmarks.Exists("OrganizationName2") Then
        ActiveDocument.Bookmarks("OrganizationName2").Range.Text = orgName
    End If
    
    ' Insert the organization address into the corresponding bookmark
    If ActiveDocument.Bookmarks.Exists("OrganizationAddress") Then
        ActiveDocument.Bookmarks("OrganizationAddress").Range.Text = orgAddress
    End If
    
    ' Insert the authority value into multiple bookmarks (Authority1 and Authority2)
    If ActiveDocument.Bookmarks.Exists("Authority1") Then
        ActiveDocument.Bookmarks("Authority1").Range.Text = authValue
    End If
    If ActiveDocument.Bookmarks.Exists("Authority2") Then
        ActiveDocument.Bookmarks("Authority2").Range.Text = authValue
    End If
    If ActiveDocument.Bookmarks.Exists("Authority3") Then
        ActiveDocument.Bookmarks("Authority3").Range.Text = authValue
    End If
    If ActiveDocument.Bookmarks.Exists("Authority4") Then
        ActiveDocument.Bookmarks("Authority4").Range.Text = authValue
    End If
    If ActiveDocument.Bookmarks.Exists("Authority5") Then
        ActiveDocument.Bookmarks("Authority5").Range.Text = authValue
    End If
    ' Insert the owner value into the corresponding bookmark
    If ActiveDocument.Bookmarks.Exists("Owner") Then
        ActiveDocument.Bookmarks("Owner").Range.Text = ownerValue
    End If
    
    ' Insert the document number into the corresponding bookmark
    If ActiveDocument.Bookmarks.Exists("DocumentNumber") Then
        ActiveDocument.Bookmarks("DocumentNumber").Range.Text = docNumber
    End If
    
    ' Close the form
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

