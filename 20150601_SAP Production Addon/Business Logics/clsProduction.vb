Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports System
Public Class clsProduction
    Inherits clsBase
    Private strQuery As String
    Private oRecordSet As SAPbobsCOM.Recordset

    Public Sub New()
        MyBase.New()
    End Sub
#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Production Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED


                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oApplication.Utilities.AddControls(oForm, "_20", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", , , "2", "Remove UnIsssed Childs", 150)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "_20" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    If oApplication.SBO_Application.MessageBox("Do you want to the unissued items?", , "Continue", "Cancel") = 2 Then
                                        BubbleEvent = False
                                        Exit Sub
                                    Else
                                        oApplication.Utilities.RemoveUnIssuedItems(oForm)
                                    End If

                                End If
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
End Class
