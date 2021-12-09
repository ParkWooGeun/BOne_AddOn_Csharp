using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn.Core
{
	/// <summary>
	/// 입고
	/// </summary>
	internal class S721 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="formUID"></param>
        public override void LoadForm(string formUID)
		{
            try
            {
                oForm = PSH_Globals.SBO_Application.Forms.Item(formUID);
                oForm.Freeze(true);

                oFormUniqueID = formUID;
                oMat01 = oForm.Items.Item("13").Specific;
                SubMain.Add_Forms(this, formUID, "S721");

                S721_CreateItems();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void S721_CreateItems()
        {
            SAPbouiCOM.Item oItem;

            try
            {
                oItem = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Top = oForm.Items.Item("1").Top - 12;
                oItem.Left = oForm.Items.Item("1").Left;
                oItem.Height = 12;
                oItem.Width = 120;
                oItem.FontSize = 10;
                oItem.Specific.Caption = "Addon running";

                oMat01.Columns.Item("U_CardName").Editable = false;
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
        /// Form Item Event
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">pVal</param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    //Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    break;
            }
        }

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "13")
                    {
                        if (pVal.ColUID == "1") //품목코드
                        {
                            if (pVal.CharPressed == 9)
                            {
                                PS_SM010 tempForm = new PS_SM010();
                                tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, oMat01.VisualRowCount);
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (pVal.ColUID == "U_CardCode") //거래처코드
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("U_CardCode").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
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

        /// <summary>
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {    
                        if (pVal.ItemUID == "13") //매트릭스
                        {
                            if (pVal.ColUID == "1")
                            {
                            }
                            else if (pVal.ColUID == "U_Qty")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value = 0; //수량
                                    oMat01.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = 1; //중량
                                }
                                else
                                {
                                    string itemCode = oMat01.Columns.Item("1").Cells.Item(pVal.Row).Specific.Value;
                                    
                                    if (dataHelpClass.GetItem_SbasUnit(itemCode) == "101") //EA자체품
                                    {
                                        oMat01.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "102") //EAUOM
                                    {
                                        if (Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode)) == 0)
                                        {
                                            oMat01.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(itemCode));
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "201") //KGSPEC
                                    {
                                        if ((Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) == 0)
                                        {
                                            oMat01.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(itemCode)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(itemCode)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(itemCode)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "202") //KG단중
                                    {
                                        if (System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 0) == 0)
                                        {
                                            oMat01.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = 1;
                                        }
                                        else
                                        {
                                            oMat01.Columns.Item("9").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("U_Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(itemCode)) / 1000, 2);
                                        }
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(itemCode) == "203") //KG입력
                                    {
                                    }
                                }
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "13") //매트릭스
                        {
                            if (pVal.ColUID == "U_CardCode")
                            {
                                string sQry = "select cardname from ocrd where cardcode = '" + oMat01.Columns.Item("U_CardCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oMat01.Columns.Item("U_CardName").Editable = true;
                                oMat01.Columns.Item("U_CardName").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
                                oMat01.Columns.Item("U_CardCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat01.Columns.Item("U_CardName").Editable = false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                }
                else if (pVal.Before_Action == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
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
    }
}
