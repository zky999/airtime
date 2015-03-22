Sub CalculateAirtime()
Dim UserRange As Range
Dim row As Range
Dim total_real_time As Date
Dim total_carrier_time As Date

total_real_time = TimeValue("00:00:00")
total_carrier_time = TimeValue("00:00:00")

Dim myErr As Long
On Error Resume Next
Set UserRange = Application.InputBox(Prompt:="Please Select Range", Default:=Selection.Address(ReferenceStyle:=xlR1C1), Title:="Range Select", Type:=8)
myErr = Err
On Error GoTo 0
If myErr <> 0 Then
    End
End If

If Not UserRange.Columns.Count = 1 Then
    MsgBox "Please select only one column data!", vbCritical, "Selected too much data"
    End
End If

For Each row In UserRange.Rows
            If InStr(1, row, "小时", vbTextCompare) Then
                total_real_time = total_real_time + TimeValue(Split(row.Value, "小时", -1, vbTextCompare)(0) & ":00:00")
                If InStr(1, row, "分", vbTextCompare) Then
                    total_real_time = total_real_time + TimeValue("00:" & Split(Split(row.Value, "小时", -1, vbTextCompare)(1), "分", -1, vbTextCompare)(0) & ":00")
                    If InStr(1, row, "秒", vbTextCompare) Then
                        total_real_time = total_real_time + TimeValue("00:00:" & Split(Split(Split(row.Value, "小时", -1, vbTextCompare)(1), "分", -1, vbTextCompare)(1), "秒", -1, vbTextCompare)(0))
                    End If
                ElseIf InStr(1, row, "秒", vbTextCompare) Then
                    total_real_time = total_real_time + TimeValue("00:00:" & Split(Split(row.Value, "小时", -1, vbTextCompare)(1), "秒", -1, vbTextCompare)(0))
                End If
            ElseIf InStr(1, row, "分", vbTextCompare) Then
                total_real_time = total_real_time + TimeValue("00:" & Split(row.Value, "分", -1, vbTextCompare)(0) & ":00")
                If InStr(1, row, "秒", vbTextCompare) Then
                    total_real_time = total_real_time + TimeValue("00:00:" & Split(Split(row.Value, "分", -1, vbTextCompare)(1), "秒", -1, vbTextCompare)(0))
                End If
            ElseIf InStr(1, row, "秒", vbTextCompare) Then
                total_real_time = total_real_time + TimeValue("00:00:" & Replace(row.Value, "秒", "", 1, -1, vbTextCompare))
            End If
        Next

Dim time_array() As String
Dim hour As Integer, min As Integer, sec As Integer
hour = 0
min = 0
sec = 0
For Each row In UserRange.Rows
    If InStr(1, row, "小时", vbTextCompare) Then
        time_array = Split(row.Value, "小时", -1, vbTextCompare)
        hour = hour + time_array(0)
        
        If InStr(1, row, "分", vbTextCompare) Then
            time_array = Split(time_array(1), "分", -1, vbTextCompare)
            min = min + time_array(0)
        End If
        
        If InStr(1, row, "秒", vbTextCompare) Then
            time_array = Split(time_array(1), "秒", -1, vbTextCompare)
            sec = time_array(0)
        End If
    ElseIf InStr(1, row, "分", vbTextCompare) Then

        time_array = Split(row, "分", -1, vbTextCompare)
        min = min + time_array(0)
        
        If InStr(1, row, "秒", vbTextCompare) Then
            time_array = Split(time_array(1), "秒", -1, vbTextCompare)
            sec = time_array(0)
        End If
    ElseIf InStr(1, row, "秒", vbTextCompare) Then
        time_array = Split(row, "秒", -1, vbTextCompare)
        sec = time_array(0)
    End If
    
    If sec > 0 Then
        min = min + 1
        sec = 0
    End If
    If min >= 60 Then
     hour = hour + 1
     min = min - 60
     End If
Next

MsgBox "Real Time:" & vbTab & total_real_time & vbCrLf & "Carrier Time:" & vbTab & hour & ":" & min & ":0" & " (" & hour * 60 + min & " minutes)", vbOKOnly, "Cell Phone Airtime"
End Sub
