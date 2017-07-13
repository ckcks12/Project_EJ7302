Attribute VB_Name = "mod_declare"
Option Explicit

Public value_SetAlpha_Alpha As Double
Public value_SetAlpha_AlwaysTop As Boolean

Public value_Zoom_Zoom As Integer

Public value_CB_Max As Integer
Public value_CB_Script As Boolean
Public value_CB_PictureAutoSave As Boolean


Public value_Search_Url As String







Public Function value_Load()
    value_SetAlpha_Alpha = CDbl(GetSetting(App.hInstance, App.hInstance, "value_setalpha_alpha", 0.4))
    value_SetAlpha_AlwaysTop = CBool(GetSetting(App.hInstance, App.hInstance, "value_setalpha_alwaystop", True))
    value_Zoom_Zoom = CInt(GetSetting(App.hInstance, App.hInstance, "value_zoom_zoom", 20))
    value_CB_Max = CInt(GetSetting(App.hInstance, App.hInstance, "value_cb_max", 10))
    value_CB_Script = CBool(GetSetting(App.hInstance, App.hInstance, "value_cb_script", False))
    value_CB_PictureAutoSave = CBool(GetSetting(App.hInstance, App.hInstance, "value_cb_pictureautosave", False))
    value_Search_Url = GetSetting(App.hInstance, App.hInstance, "value_search_url", "http://www.google.co.kr/?gws_rd=cr#newwindow=1&output=search&sclient=psy-ab&q=검색어&oq=검색어")
End Function

Public Function value_Save()
    SaveSetting App.hInstance, App.hInstance, "value_setalpha_alpha", value_SetAlpha_Alpha
    SaveSetting App.hInstance, App.hInstance, "value_setalpha_alwaystop", value_SetAlpha_AlwaysTop
    SaveSetting App.hInstance, App.hInstance, "value_zoom_zoom", value_Zoom_Zoom
    SaveSetting App.hInstance, App.hInstance, "value_cb_max", value_CB_Max
    SaveSetting App.hInstance, App.hInstance, "value_cb_script", value_CB_Script
    SaveSetting App.hInstance, App.hInstance, "value_cb_pictureautosave", value_CB_PictureAutoSave
    SaveSetting App.hInstance, App.hInstance, "value_search_url", value_Search_Url
End Function

