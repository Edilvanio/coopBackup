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
session.findById("wnd[0]/usr/ctxtMATKLA-HIGH").text = ""
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb52"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkPA_SOND").selected = true
session.findById("wnd[0]/usr/chkNEGATIV").selected = false
session.findById("wnd[0]/usr/chkXMCHB").selected = false
session.findById("wnd[0]/usr/chkNOZERO").selected = false
session.findById("wnd[0]/usr/chkNOVALUES").selected = false
session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtCHARG-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtMATART-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtMATKLA-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtEKGRUP-LOW").text = "*"
session.findById("wnd[0]/usr/chkNOVALUES").setFocus
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/tbar[1]/btn[43]").press
session.findById("wnd[1]/usr/radRB_1").setFocus
session.findById("wnd[1]/usr/radRB_1").select
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[11]").press
