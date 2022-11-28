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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmi01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkIKPF-SPERR").selected = true
session.findById("wnd[0]/usr/ctxtIKPF-WERKS").text = "1003"
session.findById("wnd[0]/usr/ctxtIKPF-LGORT").text = "0001"
session.findById("wnd[0]/usr/chkIKPF-SPERR").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000017775 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000017806 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000017736 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000017746 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000017738 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000018052 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000018072 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000018128 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000017925 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000017958 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[10,5]").text = "10000017995 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[11,5]").text = "10000018004 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[12,5]").text = "10000018070 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[13,5]").text = "10000018109 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[14,5]").text = "10000018089 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000018100 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000018006 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000018023 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000017791 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000018038 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000017914 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000017789 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000018063 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000013102 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000018065 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[10,5]").text = "10000018150 "
