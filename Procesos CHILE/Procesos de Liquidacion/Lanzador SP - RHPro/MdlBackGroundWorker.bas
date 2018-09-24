Attribute VB_Name = "MdlBackGorundWorker"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public backgroundWorker As NetFX20Wrapper.BackgroundWorkerWrapper

Public Sub BackgroundWork(ByRef argument As Variant, ByRef e As NetFX20Wrapper.RunWorkerCompletedEventArgsWrapper)
On Error GoTo eh
    Sleep argument
    e.SetResult "Background work done"
    Exit Sub
    
eh:
    e.Error.Number = Err.Number
    e.Error.Description = Err.Description
End Sub

Public Sub BackgroundWork2(ByRef argument As Variant, ByRef e As NetFX20Wrapper.RunWorkerCompletedEventArgsWrapper)
On Error GoTo eh
Progreso = 0
    For I = 1 To argument / 500
        Sleep 500
        If backgroundWorker.CancellationPending Then
            e.SetResult "Cancelled after " & I * 500 & " ms"
            e.Cancelled = True
                 
            Exit Sub
        End If
        IncPorc = (CDbl(I) / (argument / 500)) * 100
        backgroundWorker.ReportProgress (CDbl(I) / (argument / 500)) * 100
        TiempoFinalProceso = GetTickCount
        Progreso = Progreso + IncPorc
        StrSql = "UPDATE batch_proceso SET bprctiempo = " & (TiempoFinalProceso - TiempoInicialProceso) & ", bprcprogreso = " & Progreso & " WHERE bpronro = " & NroProcesoBatch
        objconnProgreso.Execute StrSql, , adExecuteNoRecords
    Next
    e.SetResult "Background work done"
    Exit Sub
    
eh:
    e.Error.Number = Err.Number
    e.Error.Description = Err.Description
End Sub

