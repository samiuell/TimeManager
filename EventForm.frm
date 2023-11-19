VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EventForm 
   Caption         =   "Enter Events"
   ClientHeight    =   5940
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4950
   OleObjectBlob   =   "EventForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EventForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OptionButton1_Click()

End Sub


Private Sub comboYear_Change()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

    ' Initialize CategoryComboBox
    CategoryComboBox.Clear
    CategoryComboBox.AddItem "School"
    CategoryComboBox.AddItem "Work"
    CategoryComboBox.AddItem "Exercise"
    CategoryComboBox.AddItem "Chores"
    CategoryComboBox.AddItem "Social/Fun"
    CategoryComboBox.AddItem "Commute"
    CategoryComboBox.AddItem ""
    CategoryComboBox.AddItem ""
    CategoryComboBox.AddItem ""
    CategoryComboBox.AddItem ""
    
    comboYear.Value = 2023
    comboMonth.Value = 1
    comboDay.Value = 1

    ' Initialize StartTimeComboBox
    StartTimeComboBox.Clear
    For i = 0 To 23
        For j = 0 To 45 Step 15
            Dim starttime As String
            starttime = Format(i, "00") & ":" & Format(j, "00")
            StartTimeComboBox.AddItem starttime
        Next j
    Next i

    ' Initialize EndTimeComboBox
    EndTimeComboBox.Clear
    For i = 0 To 23
        For j = 0 To 45 Step 15
            Dim endtime As String
            endtime = Format(i, "00") & ":" & Format(j, "00")
            EndTimeComboBox.AddItem endtime
        Next j
    Next i
    
    ' Initialize DayComboBox
    comboDay.Clear
    For d = 1 To 31
        comboDay.AddItem Format(d, "00")
    Next d

    ' Initialize MonthComboBox
    comboMonth.Clear
    For m = 1 To 12
        comboMonth.AddItem Format(m, "00")
    Next m

    ' Initialize YearComboBox
    comboYear.Clear
    Dim currentYear As Integer
    currentYear = Year(Date)
    For y = currentYear To currentYear + 10
        comboYear.AddItem y
    Next y
End Sub

Private Sub comboDay_Change()
    ' Update the days based on the selected month and year

    Dim lastDayOfMonth As Integer
    lastDayOfMonth = Day(DateSerial(selectedYear, selectedMonth + 1, 0))

    ' Clear and repopulate the DayComboBox
    comboDay.Clear
    For d = 1 To lastDayOfMonth
        comboDay.AddItem Format(d, "00")
    Next d
    
    selectedMonth = CInt(comboMonth.Value)

    selectedYear = CInt(comboYear.Value)
End Sub



