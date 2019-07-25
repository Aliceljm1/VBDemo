Set objArgs = WScript.Arguments

Set ppApp = WSH.CreateObject("powerpoint.Application", "event_")
ppApp.Visible = True
Set ppPresentations = ppApp.Presentations
Set ppPres = ppPresentations.Open(objArgs(0))

Do
    WSH.Sleep 1000
Loop

Sub event_presentationclose(pres)
    WSH.Echo("close")
    WSH.Quit
End Sub 

Sub event_PresentationSave(pres)
    WSH.Echo("save")
End Sub 

Sub event_PresentationOpen(pres)
    WSH.Echo("open")
End Sub

Sub event_NewPresentation(pres)
    WSH.Echo("new")
End Sub