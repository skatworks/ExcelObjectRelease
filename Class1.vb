Imports Microsoft.Office.Interop
Public Structure ExcelInfomation
    Dim xlApp As Excel.Application
    Dim xlWBs As Excel.Workbooks
    Dim xlWB As Excel.Workbook
End Structure

Public Class ClsParent

    public Function Parent() As Boolean

        Dim excel As ExcelInfomation
        Dim location As String = "D:\worl\Sample.xlsx"
        
        Try
            Dim cls As New ClsChild()
            For  i = 0 To 2
                cls.Child(location, excel)

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel.xlWB)
                excel.xlWB = Nothing
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel.xlWBs)
                excel.xlWBs = Nothing

                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()

                excel.xlApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel.xlApp)
                excel.xlApp = Nothing

                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
            Next
        Catch ex As Exception
        Finally
        End Try

        Return True
    End Function
End Class

Public Class ClsChild
    public Sub Child(ByRef location As String, ByRef excel As ExcelInfomation)

        Try
            excel.xlApp = New Excel.Application()
            excel.xlApp.DisplayAlerts = False
            excel.xlWBs = excel.xlApp.Workbooks
            excel.xlWB = excel.xlWBs.Open(location)

        Catch ex As Exception

        End Try

    End Sub
End Class