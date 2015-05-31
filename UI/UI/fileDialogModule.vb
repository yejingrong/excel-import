Module fileDialogModule
    Enum filetype
        Excel = 0
    End Enum

    Dim filterStr() As String = _
        {"excel文件(*.xls,*.xlsx)|*.xls;*.xlsx"}


#Region "OpenfileDialog"

    Public Function getfile(ByVal filter As filetype, Optional ByVal filterIndex As Integer = 1) As String
        Dim ofd As OpenFileDialog
        ofd = New OpenFileDialog

        ofd.Filter = filterStr(filter)
        ofd.FilterIndex = filterIndex
        If ofd.ShowDialog = DialogResult.OK Then
            Return ofd.FileName
        Else
            Return Nothing
        End If
    End Function

#End Region


End Module
