Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Option Explicit

Sub SEATING()
Dim StuRoll() As Long
Dim TotalClasses As Integer, TotalRooms As Integer
Dim SeatLimit As Integer, LimitCount As Integer
Dim SchoolName As String, ExamName As String
Dim i As Integer, BlockHeight As Integer
Dim j As Integer

Dim TotalStudents As Integer
Dim FirstRoll As Integer, LastRoll As Integer
Dim StudentRoll As Integer, StuIndex As Integer
   
Dim Label As Integer
Dim k As Integer, l As Integer
''' Rows and Column value provided by the teacher'''
Dim row As Integer, col As Integer
row = Worksheets("Sheet1").Range("D3").Value
col = Worksheets("Sheet1").Range("D4").Value

TotalStudents = Worksheets("Sheet1").Range("D8").Value
ReDim StuRoll(TotalStudents) As Long
TotalClasses = Worksheets("Sheet1").Range("D7").Value


TotalRooms = Worksheets("Sheet1").Range("A2").Value

'''Combining All Students Roll no in One Array '''
StuIndex = 0
For i = 1 To TotalClasses
       FirstRoll = Worksheets("Sheet1").Cells(3, 7).Offset(0, i).Value
       LastRoll = Worksheets("Sheet1").Cells(3, 7).Offset(1, i).Value
       For StudentRoll = FirstRoll To LastRoll
            StuRoll(StuIndex) = StudentRoll
            StuIndex = StuIndex + 1
            
       Next StudentRoll
   Next i

Debug.Print "THE TOTAL STUDENTS ARE: " & StuIndex

SchoolName = Worksheets("Sheet1").Range("D12").Value
ExamName = Worksheets("Sheet1").Range("D13").Value
''' LABEL is width caused by '''
''' 1. School Name '''
''' 2. Exam Name '''
''' 3. Room NAME '''
Label = 3
''' SCHOOL NAME till Students seats ''''
''' EXAM NAME '''
''' STUDENTS NAME '''
''' BLOCKHEIGHT is parchi ki height'''
BlockHeight = Label + 1 + Worksheets("Sheet1").Range("D3").Value

SeatLimit = Worksheets("Sheet1").Range("F5").Value
Debug.Print "There is a limit of " & SeatLimit & " students."
StuIndex = 0
For i = 1 To TotalRooms
    j = 1 + (i - 1) * BlockHeight
    Cells(j, 1).Offset(0, 0).Value = SchoolName
    Cells(j, 1).Offset(1, 0).Value = ExamName
    Cells(j, 1).Offset(2, 0).Value = Worksheets("Sheet1").Cells(i + 2, 1).Value
    LimitCount = 0
        For l = 1 To col
            For k = 1 To row
                Cells(j, 1).Offset(Label + k - 1, l - 1).Value = StuRoll(StuIndex)
                StuIndex = StuIndex + 1
                If StuIndex = TotalStudents Then
                    Exit For
                End If
                If SeatLimit > 0 Then
                    LimitCount = LimitCount + 1
                    Debug.Print LimitCount
                    If LimitCount >= SeatLimit Then
                        Exit For
                    End If
                End If
                
            Next k
            If StuIndex = TotalStudents Then
                    Exit For
            End If
            If SeatLimit > 0 Then
                
                If LimitCount >= SeatLimit Then
                    Exit For
                End If
            End If
            
        Next l
    If StuIndex = TotalStudents Then
                    Exit For
    End If
Next i
    
Debug.Print "-------------------------------------------"
End Sub





