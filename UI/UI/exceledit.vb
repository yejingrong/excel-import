Imports Excel = Microsoft.Office.Interop.Excel



Public Class exceledit
    Dim app As Excel.Application = Nothing
    Dim wbook As Excel.Workbook = Nothing
    Dim wsheet As Excel.Worksheet = Nothing
    'Public _RecordContent As String() '记录 getOneRecord 内容，以免多次调用 getOneRecord
    'Private _RecordGettedFlag As Boolean = False
    'Private _RecordGettedRowIndex As Integer = 0

    'Public ReadOnly Property Record(ByVal i As Integer, ByVal Rowindex As Integer)
    '    Get
    '        If _RecordGettedFlag = True Then
    '            If (Rowindex = _RecordGettedRowIndex) Then
    '                Return _RecordContent(i)
    '            Else
    '                Return worksheet(Rowindex, i)
    '            End If
    '        ElseIf _RecordGettedFlag = False Then
    '            MsgBox("please call getOneRecord first")
    '        End If
    '    End Get
    'End Property



    Public Sub open(ByVal FilePath As String, Optional ByVal Index As Object = 1)
        Try
            If app Is Nothing Then              'ExcelApp   这样读取 会造成？
                app = New Excel.Application
            End If
            wbook = app.Workbooks.Open(FilePath)
            wsheet = wbook.Worksheets(Index)
        Catch ex As Exception
            close()
            MsgBox("open Excel error:" + vbCrLf + ex.Message, MsgBoxStyle.Exclamation, "系统消息")
        End Try
    End Sub

    Public Sub close()
        Try
            If wbook IsNot Nothing Then
                wbook.Close()
            End If
            wbook = Nothing
            app.Workbooks.Close()
            app.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app)
            app = Nothing
            GC.Collect()
        Catch ex As Exception
            MsgBox("close Excel error:" + vbCrLf + ex.Message, MsgBoxStyle.Exclamation, "系统消息")
        End Try
    End Sub
    Public Sub fortest()

    End Sub
    Public Function getNumofRows(Optional ByVal col As Integer = 2) As Integer
        Try
            Dim Num As Integer = 0
            Dim row As Integer = 8
            Dim blankNum As Integer = 0 '连续的空格数量
            While blankNum < 4 '若连续空格为4，判断终止
                If worksheet(row, col) Is Nothing Then
                    blankNum = blankNum + 1
                Else
                    blankNum = 0
                    Num = Num + 1
                End If
                row = row + 1
            End While
            Num = row - 4
            Return Num
        Catch ex As Exception
            MsgBox(ex.Message)
            Return -1
        End Try
    End Function
    Public Function getOneRecord(ByVal Rowindex As Integer, ByVal Col() As Integer) As String()
        Try
            Dim mystr(Col.Length - 1) As String
            For i As Byte = 0 To Col.Length - 1
                If worksheet(Rowindex, Col(i)) Is Nothing Then
                    mystr(i) = Nothing
                Else
                    mystr(i) = worksheet(Rowindex, Col(i)).ToString
                End If
            Next
            Return mystr
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")
            Return Nothing
        End Try
    End Function

    Public Function worksheet(ByVal rowIndex As Integer, ByVal columnIndex As Integer) As Object
        Try
            Return wsheet.Cells(rowIndex, columnIndex).value
        Catch ex As Exception
            MsgBox("read Excel error:" + vbCrLf + wbook.Name + "," + wsheet.Name + ":" + vbCrLf + rowIndex.ToString + "行" + columnIndex.ToString + "列" + vbCrLf.ToString + ex.Message, MsgBoxStyle.Exclamation, "系统消息")
            End
        End Try
    End Function
End Class
