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
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000010132 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000010153 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000010162 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000010134 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000010130 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000010128 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000010124 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000010118 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000010164 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000010141 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000010147 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000052873 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000010183 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000010204 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000010196 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000010193 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000010202 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000010187 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000010200 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000010185 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000010221 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000010223 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000010225 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000010227 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000010236 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000010244 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000010251 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000010259 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000010275 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000010295 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000010297 "
