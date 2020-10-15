using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 부자재불출대장(부서)
    /// </summary>
    internal class PS_MM921 : PSH_BaseClass
    {
        public string oFormUniqueID;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값       
        private int oLastColRow01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM921.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_MM921_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_MM921");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_MM921_CreateItems();
                PS_MM921_ComboBox_Setting();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
                oForm.Items.Item("U_ItmBsort").Specific.VALUE = "401";
                oForm.Items.Item("ItmBname").Specific.VALUE = "부자재";
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_MM921_CreateItems()
        {
            try
            {
                oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMMdd");

                oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
                oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }


        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_MM921_ComboBox_Setting()
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ComboBox oCombo = null;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                //// 사업장
                oCombo = oForm.Items.Item("BPLId").Specific;
                sQry = "SELECT U_Minor, U_CdName  From [@PS_SY001L] WHERE Code = 'C105' AND U_UseYN Like 'Y' ORDER BY U_Seq";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oCombo.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

                ////사용처
                oCombo = oForm.Items.Item("OcrCode").Specific;
                sQry = "Select PrcCode, PrcName From [OPRC] Where DimCode = '1' Order by PrcCode";
                oRecordSet01.DoQuery(sQry);
                oCombo.ValidValues.Add("", "");
                while (!(oRecordSet01.EoF))
                {
                    oCombo.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                oCombo = oForm.Items.Item("prtdiv").Specific;
                //출력구분
                oCombo.ValidValues.Add("10", "일일불출현황");
                oCombo.ValidValues.Add("20", "불출현황(부서집계)");
                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo); //메모리 해제

            }
        }

        ///// <summary>
        ///// FormItemEnabled
        ///// </summary>
        //private void PS_MM921_FormItemEnabled()
        //{
        //    try
        //    {
        //        oForm.Freeze(true);

        //        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //        {
        //        }
        //        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
        //        {
        //        }
        //        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //        {
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }
        //}

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
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.ItemUID == "1")
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    {
                    }

                    //출력버튼 클릭시
                }
                else if (pVal.ItemUID == "Btn01")
                {
                    System.Threading.Thread thread = new System.Threading.Thread(PS_MM921_Print_Report01);
                    thread.SetApartmentState(System.Threading.ApartmentState.STA);
                    thread.Start(); 
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

        /////// <summary>
        /////// KEY_DOWN 이벤트
        /////// </summary>
        /////// <param name="FormUID">Form UID</param>
        /////// <param name="pVal">ItemEvent 객체</param>
        /////// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        ////private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        ////{
        ////    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        ////    try
        ////    {
        ////        if (pVal.Before_Action == true)
        ////        {
        ////            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
        ////        }
        ////        else if (pVal.Before_Action == false)
        ////        {
        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        ////    }
        ////    finally
        ////    {
        ////    }
        ////}

        /////// <summary>
        /////// CLICK 이벤트
        /////// </summary>
        /////// <param name="FormUID">Form UID</param>
        /////// <param name="pVal">ItemEvent 객체</param>
        /////// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        ////private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        ////{
        ////    try
        ////    {
        ////        if (pVal.Before_Action == true)
        ////        {
        ////            if (pVal.ItemUID == "Grid01")
        ////            {
        ////                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        ////                {
        ////                    if (pVal.Row > 0)
        ////                    {

        ////                    }
        ////                }
        ////                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        ////                {
        ////                }
        ////                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        ////                {
        ////                }
        ////            }
        ////        }
        ////        else if (pVal.Before_Action == false)
        ////        {
        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        ////    }
        ////    finally
        ////    {
        ////    }
        ////}

        /////// <summary>
        /////// MATRIX_LOAD 이벤트
        /////// </summary>
        /////// <param name="FormUID">Form UID</param>
        /////// <param name="pVal">ItemEvent 객체</param>
        /////// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        ////private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        ////{
        ////    try
        ////    {
        ////        if (pVal.Before_Action == true)
        ////        {

        ////        }
        ////        else if (pVal.Before_Action == false)
        ////        {
        ////            PS_MM921_FormItemEnabled();
        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        ////    }
        ////    finally
        ////    {
        ////    }
        ////}

        /////// <summary>
        /////// GOT_FOCUS 이벤트
        /////// </summary>
        /////// <param name="FormUID">Form UID</param>
        /////// <param name="pVal">ItemEvent 객체</param>
        /////// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        ////private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        ////{
        ////    try
        ////    {
        ////        if (pVal.Before_Action == true)
        ////        {

        ////            if (pVal.ItemUID == "Mat01")
        ////            {
        ////                if (pVal.Row > 0)
        ////                {
        ////                    oLastItemUID01 = pVal.ItemUID;
        ////                    oLastColUID01 = pVal.ColUID;
        ////                    oLastColRow01 = pVal.Row;
        ////                }
        ////            }
        ////            else
        ////            {
        ////                oLastItemUID01 = pVal.ItemUID;
        ////                oLastColUID01 = "";
        ////                oLastColRow01 = 0;
        ////            }
        ////        }
        ////        else if (pVal.Before_Action == false)
        ////        {
        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        ////    }
        ////    finally
        ////    {
        ////    }
        ////}

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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            //취소
                            break;
                        case "1286":
                            //닫기
                            break;
                        case "1281":
                            //찾기
                            break;
                        case "1282":
                            //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
                            break;
                        case "1293":
                            //행삭제
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            //취소
                            break;
                        case "1286":
                            //닫기
                            break;
                        case "1293":
                            //행삭제
                            break;
                        ////Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281":
                            //찾기
                            break;
                        case "1282":
                            //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
                            break;
                    }
                    ////BeforeAction = False
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            //취소
                            break;
                        case "1286":
                            //닫기
                            break;
                        case "1293":
                            //행삭제
                            break;
                        ////Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281":
                            //찾기
                            break;
                        ////Call PS_MM921_FormItemEnabled '//UDO방식
                        case "1282":
                            //추가
                            break;
                        ////Call PS_MM921_FormItemEnabled '//UDO방식
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
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

        /// <summary>
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
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
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_MM921_Print_Report01()
        {
            string WinTitle = null;
            string ReportName = null;
            string BPLID = null;
            string DocDateFr = null;
            string DocDateTo = null;
            string ItmMsort = null;
            string OcrCode = null;
            string prtdiv = null;
            string ItemCode = null;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //// 조회조건문
                BPLID = oForm.Items.Item("BPLId").Specific.VALUE.ToString().Trim();
                DocDateFr = oForm.Items.Item("DocDateFr").Specific.VALUE.ToString().Trim();
                DocDateTo = oForm.Items.Item("DocDateTo").Specific.VALUE.ToString().Trim();
                ItmMsort = oForm.Items.Item("ItmMsort").Specific.VALUE.ToString().Trim();
                ItemCode = oForm.Items.Item("ItemCode").Specific.VALUE.ToString().Trim();
                OcrCode = oForm.Items.Item("OcrCode").Specific.VALUE.ToString().Trim();
                prtdiv = oForm.Items.Item("prtdiv").Specific.VALUE.ToString().Trim();

                if (string.IsNullOrEmpty(BPLID))
                    BPLID = "%";
                if (string.IsNullOrEmpty(ItmMsort))
                    ItmMsort = "%";
                if (string.IsNullOrEmpty(OcrCode))
                    OcrCode = "%";
                if (string.IsNullOrEmpty(ItemCode))
                    ItemCode = "%";

                WinTitle = "[PS_MM921] 분말부자재비용분석";
                if (prtdiv == "10")
                {
                    WinTitle = "[PS_MM921_01] 부자재 불출대장(일자별)";
                    ReportName = "PS_MM921_01.RPT";
                    dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLID));
                    dataPackParameter.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr));
                    dataPackParameter.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo));
                    dataPackParameter.Add(new PSH_DataPackClass("@ItmMsort", ItmMsort));
                    dataPackParameter.Add(new PSH_DataPackClass("@OcrCode", OcrCode)); 
                    dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
                }
                else
                {
                    WinTitle = "[PS_MM921_02] 부자재 불출대장(담당별집계)";
                    ReportName = "PS_MM921_02.RPT";
                    dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLID)); 
                    dataPackParameter.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr)); 
                    dataPackParameter.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo)); 
                    dataPackParameter.Add(new PSH_DataPackClass("@ItmMsort", ItmMsort)); 
                    dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode)); 
                }

                //Formula                
                dataPackFormula.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr.Substring(0, 4) + "-" + DocDateFr.Substring(4, 2) + "-" + DocDateFr.Substring(6, 2))); //년월
                dataPackFormula.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo.Substring(0, 4) + "-" + DocDateTo.Substring(4, 2) + "-" + DocDateTo.Substring(6, 2))); //년월
                dataPackFormula.Add(new PSH_DataPackClass("@BPLId", dataHelpClass.Get_ReData("BPLName", "BPLId", "OBPL", BPLID, ""))); //사업장

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }
    }
}
