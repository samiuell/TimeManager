VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "sleepfacultyform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    
    Me.SleepComboBox.AddItem "6 hours"
    Me.SleepComboBox.AddItem "7 hours"
    Me.SleepComboBox.AddItem "8 hours"
    Me.SleepComboBox.AddItem "9 hours"
    Me.SleepComboBox.AddItem "10 hours"
    Me.FacultyComboBox.AddItem "Arts"
    Me.FacultyComboBox.AddItem "Engineering"
    Me.FacultyComboBox.AddItem "Environment"
    Me.FacultyComboBox.AddItem "Health"
    Me.FacultyComboBox.AddItem "Mathematics"
    Me.FacultyComboBox.AddItem "Science"
    
   
End Sub
