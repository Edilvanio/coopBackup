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
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000008104 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000009141 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000009148 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000009138 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000008106 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000008250 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000008115 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000008117 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000008116 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000009149 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[10,5]").text = "10000008979 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[11,5]").text = "10000008880 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[12,5]").text = "10000008851 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[13,5]").text = "10000008885 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[14,5]").text = "10000008863 "
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[0,5]").text = "10000008862 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[1,5]").text = "10000008864 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[2,5]").text = "10000008909 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[3,5]").text = "10000008897 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[4,5]").text = "10000008944 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[5,5]").text = "10000008950 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[6,5]").text = "10000008901 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[7,5]").text = "10000008943 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[8,5]").text = "10000008940 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[9,5]").text = "10000008881 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[10,5]").text = "10000008872 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[11,5]").text = "10000008976 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[12,5]").text = "10000008978 "
session.findById("wnd[0]/usr/sub:SAPMM07I:0721/ctxtISEG-MATNR[13,5]").text = "10000009151 "
