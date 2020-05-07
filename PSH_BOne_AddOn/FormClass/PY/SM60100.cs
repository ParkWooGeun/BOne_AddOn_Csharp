using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    internal class SM60100 : PSH_BaseClass
    {
        ////****************************************************************************
        ////  File           : SM60100.cls
        ////  Module         : 인사관리 > 사용자 정의 필드
        ////  Desc           : OHEM
        ////****************************************************************************

        public string oFormUniqueID;
        //public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Matrix oMat1;
        public SAPbouiCOM.Grid oGrid1;

        private int sFromPane;
        private int sToPane;
        private bool sInit;
        private string sTable;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        public override void LoadForm(string oFormUid)
        {
            oFormUniqueID = oFormUid;
            oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

            SM60100_CreateItems();

            SM60100_FormItemEnabled();
        }

        private void SM60100_CreateItems()
        {            
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Item oItem01 = null;
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbouiCOM.Folder oFolder = null;
            SAPbouiCOM.EditText oEdit = null;
            SAPbobsCOM.UserFieldsMD oUserField = null;

            try
            {
                oForm.Freeze(true);
                             
                oItem = null;                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("SM60100_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {                
                oForm.Freeze(false);
                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oFolder); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField); //메모리 해제
            }
        }

        private void SM60100_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("SM60100_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        ////1
                        if (pVal.Before_Action == false)
                        {
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        ////2
                        break;

                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        ////3
                        break;
                    //// 종료, 취소 막음

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        ////4
                        break;

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        ////5
                        if (pVal.Before_Action == false)
                        {
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        ////6
                        break;
                    //// 종료, 취소 막음

                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        ////7
                        break;

                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        ////8
                        break;

                    case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                        ////9
                        break;

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        ////10
                        break;

                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
                        ////11
                        break;

                    case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
                        ////12
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        ////16
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
                        ////17
                        if (pVal.BeforeAction == true)
                        {
                        }
                        else if (pVal.BeforeAction == false)
                        {
                            SubMain.Remove_Forms(oFormUniqueID);
                            oForm = null;
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                        ////18
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                        ////19
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        ////20
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
                        ////21
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
                        ////22
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
                        ////23
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        ////27
                        break;

                    case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
                        ////37
                        break;

                    case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
                        ////38
                        break;

                    case SAPbouiCOM.BoEventTypes.et_Drag:
                        ////39
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_ItemEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {               
            }        
        }

        public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if ((pVal.BeforeAction == true))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            break;
                        case "1282":
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1293":
                            break;
                        case "1281":
                            SM60100_FormItemEnabled();
                            break;
                        case "1282":
                            oForm.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            SM60100_FormItemEnabled();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }       
        }

        public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            ////36
                            break;
                    }
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            ////33
                            //UPGRADE_WARNING: oForm.Items.Item(SOTYPE).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            if (oForm.Items.Item("SOTYPE").Specific.Value == "2" & sInit == false)
                            {
                                SM60100_CreateItems();
                            }
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            ////36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {             
            }
        }

        public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            oForm = PSH_Globals.SBO_Application.Forms.Item(pVal.FormUID);

            try
            {
                if (pVal.BeforeAction)
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK:
                            break;
                            //            If Left$(pVal.ItemUID, 3) = "CMB" Or Left$(pVal.ItemUID, 3) = "EDT" Then
                            //                oForm.ActiveItem = "16"
                            //                BubbleEvent = False
                            //                Exit Sub
                            //            ElseIf pVal.ItemUID = "38" Then
                            //                If GF_Nz(GF_DLookup("SUM(U_INGQTY)", "POR1", " DOCENTRY = " & Trim$(oForm.Items("8").Specific.String))) > 0 Then
                            //                    oForm.ActiveItem = "16"
                            //                    BubbleEvent = False
                            //                    Exit Sub
                            //                End If
                            //            End If
                    }
                }
                else
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK:
                            break;
                            //            If Left$(pVal.ItemUID, 3) = "CMB" Or Left$(pVal.ItemUID, 3) = "EDT" Then
                            //                BubbleEvent = False
                            //                Exit Sub
                            //            End If
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        private void SM60100_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string sQry = null;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            try
            {
                if (dataHelpClass.Get_ReData("COUNT(CardCode)", "DocEntry", "[ORDR]", "'" + oForm.Items.Item("8").Specific.Value + "'", "") > 0)
                {
                    /// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/
                    WinTitle = "[SM60100] : ORDER SHEET";
                    ReportName = "SM60100_1.RPT";


                    //sQry = " Exec SM60100_1 " + "'" + Strings.Trim(oForm.Items.Item("8").Specific.Value) + "'";
                    //PSH_Globals.gRpt_Formula = new string[9];
                    //PSH_Globals.gRpt_Formula_Value = new string[9];
                    //PSH_Globals.gRpt_SRptSqry = new string[2];
                    //PSH_Globals.gRpt_SRptName = new string[2];

                    ///// Formula 수식필드***************************************************/

                    ///// SubReport /

                    //if (PSH_Globals.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false)
                    //{
                    //    PSH_Globals.SBO_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    //}
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("저장된 판매오더 문서가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("SM60100_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        public bool SM60100_DataValidCheck()
        {
            bool functionReturnValue = false;
            functionReturnValue = false;

            try
            {
                //
                //    '// 만약 내수일 경우에는 SKIP
                //    If Not oForm.Items("SOTYPE").Specific.Selected Is Nothing Then
                //        If Trim$(oForm.Items("SOTYPE").Specific.Selected.Value) = "2" Then
                //            If oForm.Items("SETTLE").Specific.Selected Is Nothing Then
                //                SBO_Application.SetStatusBarMessage "결제조건은 필수 입력 항목입니다.", bmt_Short, True
                //                SM60100_DataValidCheck = False
                //                Exit Function
                //            End If
                //        End If
                //    End If

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                functionReturnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText("SM60100_DataValidCheck_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
            
            return functionReturnValue;       
        }
    }
}