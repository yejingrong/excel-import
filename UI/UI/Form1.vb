Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        importfromExcel()
    End Sub
    Structure sample_flowday
        Public datadate As String
        Public flow_day As Decimal
    End Structure



    Public Function importfromExcel() As Boolean
        Dim e As New exceledit
        Try
            e.open("E:\****.xls")
            MsgBox(e.worksheet(8, "B").ToString)
            Dim totalnum As Integer = e.getNumofRows(1)
            Dim i As Integer = 8
            TextBox1.Text = "正在导入 " + totalnum.ToString + "个变量……" + vbCrLf
            Dim Arecord As New sample_flowday
            While i < totalnum
                Arecord.datadate = e.getOneRecord(i, {1, 8})(0)
                Arecord.flow_day = CDec(e.getOneRecord(i, {1, 8})(1))
                Dim mystr$ = "insert sample_flowday ( datadate,flow_day) values(" & Arecord.datadate & "," & Arecord.flow_day & ")"
                TextBox1.Text = TextBox1.Text & mystr
                i = i + 1
            End While

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error")

            e.close()
            e = Nothing
            Return False
        End Try
        If e IsNot Nothing Then e.close()
        Return True
    End Function


End Class
