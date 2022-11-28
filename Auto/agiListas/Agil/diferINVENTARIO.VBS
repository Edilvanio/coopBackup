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
session.findById("wnd[0]/tbar[0]/okcd").text = "/NMI20"
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/ctxtIM_MATNR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtIM_WERKS-LOW").text = "1003"
session.findById("wnd[0]/usr/ctxtIM_LGORT-LOW").text = "*"

session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[3]/menu[6]/menu[0]").select
session.findById("wnd[0]/tbar[0]/btn[86]").press
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[2]/usr/lbl[41,4]").setFocus
session.findById("wnd[2]/usr/lbl[41,4]").caretPosition = 8
session.findById("wnd[2]").sendVKey 2
session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").setFocus
session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/cmbPRIPAR_DYN-PRIMM").key = "X"
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
