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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmb51"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtMATNR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "1003"
session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtBWART-LOW").text = "z43"
session.findById("wnd[0]/usr/ctxtBUDAT-LOW").text = "260822"
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").text = ""
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell 5,"MAKTX"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectedRows = "5"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectContextMenuItem "&XXL"
session.findById("wnd[1]/usr/chkCB_ALWAYS").setFocus
session.findById("wnd[1]/usr/chkCB_ALWAYS").selected = true
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "TrocasZ43"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
session.findById("wnd[1]/tbar[0]/btn[11]").press
