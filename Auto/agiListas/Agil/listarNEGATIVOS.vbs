If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1003"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "0001"
session.findById("wnd[0]/usr/ctxtMATKLA-LOW").text = "130131001"
session.findById("wnd[0]/usr/ctxtMATKLA-HIGH").text = "190194003"
session.findById("wnd[0]/usr/ctxtMATKLA-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtMATKLA-HIGH").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkNEGATIV").selected = true
session.findById("wnd[0]/usr/chkNEGATIV").setFocus

session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/usr/radRB_1").setFocus
session.findById("wnd[1]/usr/radRB_1").select


