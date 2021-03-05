using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// OHEM AddOn
    /// </summary>
    internal class SM60100 : PSH_BaseClass
    {
        private string oFormUniqueID;

        public override void LoadForm(string oFormUid)
        {
            oFormUniqueID = oFormUid;
            oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

            SM60100_CreateItems();
            SM60100_FormItemEnabled();
            PSH_Globals.ExecuteEventFilter(typeof(SM60100));
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private bool SM60100_DataValidCheck()
        {
            bool functionReturnValue = false;
            
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }

            return functionReturnValue;
        }

        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                        break;

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                        break;

                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                        break;
                    
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                        break;

                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                        break;
                    
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                        break;

                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                        break;

                    case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                        break;

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                        break;

                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                        break;

                    case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                        Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                        break;

                    case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                        break;

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                        break;

                    case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                        break;

                    case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                        break;

                    case SAPbouiCOM.BoEventTypes.et_Drag: //39
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {               
            }        
        }

        /// <summary>
        /// FORM_UNLOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    SubMain.Remove_Forms(oFormUniqueID);
                    oForm = null;
                }
                else if (pVal.Before_Action == false)
                {

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
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
                else if (pVal.BeforeAction == false)
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }       
        }

        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {             
            }
        }

        /// <summary>
        /// FORM_DATA_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="BusinessObjectInfo">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_DATA_LOAD(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    if (oForm.Items.Item("SOTYPE").Specific.Value == "2")
                    {
                        SM60100_CreateItems();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
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
                    }
                }
                else
                {
                    switch (pVal.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK:
                            break;
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
    }
}