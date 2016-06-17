Attribute VB_Name = "ModDelay"
'**************结合类模块"clswaitabletimer"使用****************
Public mobjWaitTimer As clswaitabletimer
Public Sub Delay(Wtime As Long)
    Set mobjWaitTimer = New clswaitabletimer
    Do
        If mbWorkToDo Then
            'Call ProcessWork
        Else
            mobjWaitTimer.Wait (Wtime)
        End If
    Loop Until Not mbStop
    Set mobjWaitTimer = Nothing
End Sub


