Attribute VB_Name = "QuickAndDirtyModule"
Option Explicit

' note:
' slight changes in this version on 2020-07-21
' emergency implementation - issue with order with bigger parts list - macro error!
' fckg no-pro :)

' change in lines ~138 - wrong clearing selections!

' added private sub and function (visible only in scope of this module) - previous code was not working.
' On bottom:
' findRow
' clearBreaks

'
' used in lines:
' ~ 326 for findRow
' ~ > 420 managing h page breaks...




' next sub version 2.23 - added 6th static page

Private Sub unhideAll()
    
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        sh.Visible = xlSheetVisible
    Next sh
End Sub
