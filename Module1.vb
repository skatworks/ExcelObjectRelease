Module Module1

    Sub Main()
        Dim cls As New ClsParent()
        If cls.Parent() Then
            MsgBox("True")
        End If
    End Sub

End Module
