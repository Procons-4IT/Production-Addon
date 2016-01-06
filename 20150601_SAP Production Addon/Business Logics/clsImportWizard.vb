Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System

Public Class clsImportWizard
    Inherits clsBase
    Private strQuery As String
    Private oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub LoadForm()
        Try
            oForm = oApplication.Utilities.LoadForm(xml_ImpWiz, frm_ImpWiz)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(True)
            oForm.PaneLevel = 1
            Initialize(oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.ActiveForm
            Select Case pVal.MenuUID
                Case mnu_ImpWiz
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
            End Select
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_ImpWiz Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "9" Then 'Browse
                                    oApplication.Utilities.OpenFileDialogBox(oForm, "17")
                                ElseIf (pVal.ItemUID = "7") Then 'Next
                                    If CType(oForm.Items.Item("17").Specific, SAPbouiCOM.StaticText).Caption <> "" Then
                                        If oApplication.Utilities.ValidateFile(oForm, "17") Then
                                            Dim strDFNPath As String = String.Empty
                                            If oApplication.Utilities.CopyFile(oForm, "17", strDFNPath) Then
                                                If oApplication.Excel.updateExcelTemplate(strDFNPath) Then
                                                    oApplication.Utilities.strSFilePath = strDFNPath
                                                    oApplication.Utilities.ShowSaveFile(oForm, strDFNPath)
                                                    Dim intVal As Integer = oApplication.SBO_Application.MessageBox("Do You Want to Open File", 2, "Yes", "No", "")
                                                    If intVal = 1 Then
                                                        System.Diagnostics.Process.Start(oApplication.Utilities.strDFilePath)
                                                    End If
                                                Else
                                                    oApplication.Utilities.Message("Cannot Find Style,ParentItem and Description....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                End If
                                                If strDFNPath <> "" Then
                                                    If File.Exists(strDFNPath) Then
                                                        File.Delete(strDFNPath)
                                                    End If
                                                End If
                                            End If
                                        End If
                                Else
                                        oApplication.Utilities.Message("Select File to Import....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False
                                End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                        End Select
                End Select
            End If
        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Data Events"
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Right Click"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Try
            oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
            If oForm.TypeEx = frm_ImpWiz Then

            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Function"

    Private Sub Initialize(ByVal oForm As SAPbouiCOM.Form)
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

#End Region

End Class
