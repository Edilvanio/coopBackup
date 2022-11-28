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
session.findById("wnd[0]/tbar[0]/okcd").text = "/n/avalara/ztdr011"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtS_DEMI-LOW").text = "010822"
session.findById("wnd[0]/usr/ctxtS_DEMI-HIGH").text = "010130"
session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").text = "1922"
session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = "1000"
session.findById("wnd[0]/usr/ctxtS_BRANCH-LOW").text = "1003"
session.findById("wnd[0]/usr/ctxtP_STATUS").setFocus
session.findById("wnd[0]/usr/ctxtP_STATUS").caretPosition = 0
session.findById("wnd[0]").sendVKey 4
session.findById("wnd[1]/usr/lbl[11,4]").setFocus
session.findById("wnd[1]/usr/lbl[11,4]").caretPosition = 2
session.findById("wnd[1]").sendVKey 2
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = -1
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "NNF"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&SORT_ASC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "NNF",12
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "NAME1",29
