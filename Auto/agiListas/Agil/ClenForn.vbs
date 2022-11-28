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
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"STEP_DESC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "STEP_DESC"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").contextMenu
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectContextMenuItem "&FILTER"
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,7]").setFocus
session.findById("wnd[3]/usr/lbl[1,7]").caretPosition = 12
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").caretPosition = 0
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV").select
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,1]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,6]").setFocus
session.findById("wnd[3]/usr/lbl[1,6]").caretPosition = 15
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,2]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,8]").setFocus
session.findById("wnd[3]/usr/lbl[1,8]").caretPosition = 8
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").setFocus
session.findById("wnd[2]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,3]").caretPosition = 0
session.findById("wnd[2]").sendVKey 4
session.findById("wnd[3]/usr/lbl[1,4]").setFocus
session.findById("wnd[3]/usr/lbl[1,4]").caretPosition = 17
session.findById("wnd[3]").sendVKey 2
session.findById("wnd[2]/tbar[0]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
