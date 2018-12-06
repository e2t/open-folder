Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object

Sub Main()
    Dim currentDoc As ModelDoc2
    Dim selmgr As SelectionMgr
    Dim pathname As String
    Dim selectComp As Component2
    Dim selectView As View
    Dim refDoc As ModelDoc2
    
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If Not currentDoc Is Nothing Then
        pathname = currentDoc.GetPathName
        Set selmgr = currentDoc.SelectionManager
        If selmgr.GetSelectedObjectCount2(-1) > 0 Then
            Select Case currentDoc.GetType
                Case swDocASSEMBLY
                    Set selectComp = selmgr.GetSelectedObjectsComponent3(1, -1)
                    If Not selectComp Is Nothing Then
                        pathname = selectComp.GetPathName
                    End If
                Case swDocDRAWING
                    Set selectView = selmgr.GetSelectedObjectsDrawingView2(1, -1)
                    If Not selectView Is Nothing Then
                        Set refDoc = selectView.ReferencedDocument
                        If Not refDoc Is Nothing Then
                            pathname = refDoc.GetPathName
                        End If
                    End If
            End Select
        End If
        Shell "explorer /select,""" & pathname & """", vbNormalFocus
    End If
End Sub
