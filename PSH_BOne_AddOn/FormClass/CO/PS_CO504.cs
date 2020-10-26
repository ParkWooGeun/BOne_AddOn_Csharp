using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 품목별원가등록
    /// </summary>
    internal class PS_CO504 : PSH_BaseClass
    {
        public string oFormUniqueID;
        //public SAPbouiCOM.Form oForm;
        public SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_USERDS01;

        private string oLast_Item_UID;        //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private string oLast_Col_UID;        //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oLast_Col_Row;
        private int oSeq;
        private string TmpCode;

        private SAPbouiCOM.BoFormMode oFormMode01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO504.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_CO504_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_CO504");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                //oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);

                oForm.EnableMenu(("1281"), false);                //// 제거
                oForm.EnableMenu(("1292"), false);                //// 행삭제

                PS_CO504_CreateItems();
                //PS_CO504_FormItemEnabled();
                //PS_CO504_FormClear();
                PS_CO504_Initial_Setting();
                //PS_CO504_AddMatrixRow(0, true);
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
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_CO504_CreateItems()
        {
            try
            {
                oDS_PS_USERDS01 = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oMat01 = oForm.Items.Item("Mat01").Specific;

                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
                oForm.DataSources.UserDataSources.Item("DocDate").Value = DateTime.Now.ToString("yyyyMMdd");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_CO504_Initial_Setting()
        {
            try
            {
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("1", "개별");
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("2", "집계");
                oForm.Items.Item("Gubun").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        ///// <summary>
        ///// 모드에 따른 아이템 설정
        ///// </summary>
        //private void PS_CO504_FormItemEnabled()
        //{
        //    try
        //    {
        //        if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
        //        {
        //            ////각모드에따른 아이템설정
        //            //        oForm.Items("DocEntry").Enabled = False
        //            oForm.Items.Item("DocDate").Enabled = true;

        //        }
        //        else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
        //        {
        //            ////각모드에따른 아이템설정
        //            //        oForm.Items("DocEntry").Enabled = True
        //            oForm.Items.Item("DocDate").Enabled = true;

        //        }
        //        else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
        //        {
        //            ////각모드에따른 아이템설정
        //            //        oForm.Items("DocEntry").Enabled = False
        //            oForm.Items.Item("DocDate").Enabled = false;
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

        ///// <summary>
        ///// 
        ///// </summary>
        ///// <param name="oRow">행 번호</param>
        ///// <param name="RowIserted">행 추가 여부</param>
        //private void PS_CO504_AddMatrixRow(int oRow, bool RowIserted = false)
        //{
        //    try
        //    {
        //        if (RowIserted == false)
        //        {
        //            oDS_PS_CO504L.InsertRecord((oRow));
        //        }

        //        oMat01.AddRow();
        //        oDS_PS_CO504L.Offset = oRow;
        //        oDS_PS_CO504L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

        //        oMat01.LoadFromDataSource();
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

        ///// <summary>
        ///// DocEntry 초기화
        ///// </summary>
        //private void PS_CO504_FormClear()
        //{
        //    string DocNum = null;
        //    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        //    try
        //    {
        //        DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO504'", "");
        //        if (Convert.ToDouble(DocNum) == 0)
        //        {
        //            oDS_PS_CO504H.SetValue("DocEntry", 0, "1");
        //        }
        //        else
        //        {
        //            oDS_PS_CO504H.SetValue("DocEntry", 0, DocNum);
        //            // 화면에 적용이 안되기 때문
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //}

        /////// <summary>
        /////// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /////// </summary>
        /////// <param name="oUID"></param>
        /////// <param name="oRow"></param>
        /////// <param name="oCol"></param>
        ////private void PS_CO504_FlushToItemValue(string oUID, int oRow, string oCol)
        ////{
        ////    string i = null;
        ////    string sQry = null;
        ////    SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        ////    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

        ////    try
        ////    {
        ////        // Matrix 필드에 질의 응답 창 띄워주기
        ////        switch (oCol)
        ////        {
        ////            case "Code":
        ////                oMat01.FlushToDataSource();
        ////                oDS_PS_CO504L.Offset = oRow - 1;

        ////                oForm.Freeze(true);
        ////                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ////                sQry = "Select t1.U_CdName From [@PS_SY001H] t Inner Join [@PS_SY001L] t1 On t.Code = t1.Code Where t.Code = 'F002' and t1.U_Minor = '" + oMat01.Columns.Item("Code").Cells.Item(oRow).Specific.VALUE.ToString().Trim() + "'";
        ////                oRecordSet01.DoQuery(sQry);
        ////                oDS_PS_CO504L.SetValue("U_Name", oRow - 1, oRecordSet01.Fields.Item("U_CdName").Value.ToString().Trim());
        ////                oForm.Freeze(false);
        ////                oMat01.LoadFromDataSource();

        ////                //--------------------------------------------------------------------------------------------
        ////                if (oRow == oMat01.RowCount & !string.IsNullOrEmpty(oDS_PS_CO504L.GetValue("U_Name", oRow - 1).ToString().Trim()))
        ////                {
        ////                    //// 다음 라인 추가
        ////                    PS_CO504_AddMatrixRow(oMat01.RowCount, false);

        ////                    oMat01.Columns.Item("Value").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        ////                }
        ////                break;

        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        ////    }
        ////    finally
        ////    {
        ////        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
        ////    }
        ////}

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
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        LoadData();
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

        ///// <summary>
        ///// GOT_FOCUS 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {
        //            oLast_Item_UID = pVal.ItemUID;
        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //            oLast_Item_UID = pVal.ItemUID;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

        /////// <summary>
        /////// VALIDATE 이벤트
        /////// </summary>
        /////// <param name="FormUID">Form UID</param>
        /////// <param name="pVal">ItemEvent 객체</param>
        /////// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        ////private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        ////{
        ////    string sQry = null;
        ////    SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        ////    try
        ////    {
        ////        if (pVal.Before_Action == true)
        ////        {
        ////        }
        ////        else if (pVal.Before_Action == false)
        ////        {

        ////            if (pVal.ItemUID == "ItmBsort" & pVal.ItemChanged == true)
        ////            {
        ////                sQry = "Select Name From [@PSH_ItmBsort] Where Code = '" + oForm.Items.Item("ItmBsort").Specific.VALUE.ToString().Trim() + "'";
        ////                oRecordSet01.DoQuery(sQry);
        ////                oForm.Items.Item("ItmBname").Specific.VALUE = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
        ////            }
        ////            if (pVal.ColUID == "Code" & pVal.ItemChanged == true)
        ////            {
        ////                PS_CO504_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
        ////            }


        ////        }
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        ////        BubbleEvent = false;
        ////    }
        ////    finally
        ////    {
        ////    }
        ////}

        ///// <summary>
        ///// MATRIX_LOAD 이벤트
        ///// </summary>
        ///// <param name="FormUID">Form UID</param>
        ///// <param name="pVal">ItemEvent 객체</param>
        ///// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        //private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //    try
        //    {
        //        if (pVal.Before_Action == true)
        //        {

        //        }
        //        else if (pVal.Before_Action == false)
        //        {
        //            PS_CO504_AddMatrixRow(1, true);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //    }
        //}

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
                        case "1281":
                            //찾기
                            break;

                        case "1282":
                            //추가
                            break;

                        case "1287":
                            //복제
                            break;

                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            //레코드이동버튼
                            break;

                        case "1293":
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

        /////// <summary>
        /////// PS_CO504_MTX01
        /////// </summary>
        /////// <param name=""></param>
        ////private void PS_CO504_MTX01()
        ////{
        ////    string sQry = null;
        ////    int i = 0;

        ////    SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        ////    SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        ////    try
        ////    {
        ////        sQry = "Select a.ItemCode, a.ItemName, a.U_ItmMsort, ItmMname = (select U_CodeName from [@PSH_ITMMSORT] Where U_Code = a.U_ItmMsort)  From [OITM] a Where a.U_ItmBsort = '" + oForm.Items.Item("ItmBsort").Specific.VALUE.ToString().Trim() + "' Order By a.U_ItmMsort, a.ItemCode";
        ////        oRecordSet02.DoQuery(sQry);

        ////        oDS_PS_CO504L.Clear();
        ////        oMat01.Clear();
        ////        oMat01.FlushToDataSource();

        ////        i = 0;
        ////        while (!(oRecordSet02.EoF))
        ////        {
        ////            oDS_PS_CO504L.InsertRecord(i);
        ////            oDS_PS_CO504L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        ////            oDS_PS_CO504L.SetValue("U_ItemCode", i, oRecordSet02.Fields.Item(0).Value.ToString().Trim());
        ////            oDS_PS_CO504L.SetValue("U_ItemName", i, oRecordSet02.Fields.Item(1).Value.ToString().Trim());
        ////            oDS_PS_CO504L.SetValue("U_ItmMsort", i, oRecordSet02.Fields.Item(2).Value.ToString().Trim());
        ////            oDS_PS_CO504L.SetValue("U_ItmMName", i, oRecordSet02.Fields.Item(3).Value.ToString().Trim());
        ////            i = i + 1;
        ////            oRecordSet02.MoveNext();
        ////        }

        ////        oMat01.LoadFromDataSource();
        ////    }
        ////    catch (Exception ex)
        ////    {
        ////        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        ////    }
        ////    finally
        ////    {
        ////    }
        ////}

        ///// <summary>
        ///// 필수입력사항 체크
        ///// </summary>
        ///// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        //private bool PS_CO504_HeaderSpaceLineDel()
        //{
        //    bool functionReturnValue = false;
        //    short ErrNum = 0;

        //    try
        //    {
        //        if (string.IsNullOrEmpty(oDS_PS_CO504H.GetValue("U_DocDate", 0)))
        //        {
        //            ErrNum = 1;
        //            throw new Exception();
        //        }
        //        functionReturnValue = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ErrNum == 1)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("일자는 필수입력 사항입니다.확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //        functionReturnValue = false;
        //    }
        //    finally
        //    {
        //    }
        //    return functionReturnValue;
        //}

        ///// <summary>
        ///// 필수입력사항 체크
        ///// </summary>
        ///// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        //private bool PS_CO504_MatrixSpaceLineDel()
        //{
        //    bool functionReturnValue = false;
        //    short ErrNum = 0;

        //    try
        //    {
        //        oForm.Freeze(true);
        //        ErrNum = 0;

        //        //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //        //// 화면상의 메트릭스에 입력된 내용을 모두 디비데이터소스로 넘긴다
        //        //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //        oMat01.FlushToDataSource();

        //        //// 라인
        //        if (oMat01.VisualRowCount <= 1)
        //        {
        //            ErrNum = 1;
        //            throw new Exception();
        //        }

        //        //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //        //// 맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
        //        //// 이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
        //        //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //        if (oMat01.VisualRowCount > 0)
        //        {
        //            if (string.IsNullOrEmpty(oDS_PS_CO504L.GetValue("U_ItmBsort", oMat01.VisualRowCount - 1)))
        //            {
        //                oDS_PS_CO504L.RemoveRecord(oMat01.VisualRowCount - 1);
        //            }
        //        }
        //        //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //        //행을 삭제하였으니 DB데이터 소스를 다시 가져온다
        //        //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
        //        oMat01.LoadFromDataSource();
        //        functionReturnValue = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (ErrNum == 1)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //        functionReturnValue = false;
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //    }

        //    return functionReturnValue;
        //}

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns>True:필수입력사항을 모두 입력, Fasle:필수입력사항 중 하나라도 입력하지 않았음</returns>
        public void LoadData()
        {
            short i = 0;
            string sQry = null;
            string DocDate = null;
            string GUBUN = null;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oMat01.Clear();
                oDS_PS_USERDS01.Clear();

                GUBUN = oForm.Items.Item("Gubun").Specific.VALUE.ToString().Trim();
                DocDate = oForm.Items.Item("DocDate").Specific.VALUE.ToString().Trim();



                sQry = "EXEC [PS_CO504_01] '" + DocDate + "', '" + GUBUN + "'";
                oRecordSet01.DoQuery(sQry);

                if (oRecordSet01.RecordCount == 0)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.",BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Warning);
                    oRecordSet01 = null;
                    oForm.Freeze(false);
                    return;
                }

                SAPbouiCOM.ProgressBar ProgBar01 = null;
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PS_USERDS01.Size)
                    {
                        oDS_PS_USERDS01.InsertRecord((i));
                    }

                    oMat01.AddRow();
                    oDS_PS_USERDS01.Offset = i;
                    oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_USERDS01.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("ItmBname").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("ItmMname").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColSum01", i, oRecordSet01.Fields.Item("SQTy").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("SWgt").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColSum02", i, oRecordSet01.Fields.Item("SAmt").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColSum03", i, oRecordSet01.Fields.Item("SmQty").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColQty02", i, oRecordSet01.Fields.Item("SmWgt").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColSum04", i, oRecordSet01.Fields.Item("SmAmt").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColSum05", i, oRecordSet01.Fields.Item("MQTy").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColQty03", i, oRecordSet01.Fields.Item("MWgt").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColSum06", i, oRecordSet01.Fields.Item("MAmt").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColSum07", i, oRecordSet01.Fields.Item("MmQty").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColQty04", i, oRecordSet01.Fields.Item("MmWgt").Value.ToString().Trim());
                    oDS_PS_USERDS01.SetValue("U_ColSum08", i, oRecordSet01.Fields.Item("MmAmt").Value.ToString().Trim());
                    //----------------------------------------------------------------------------------------------------------
                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                //            oMat01.AutoResizeColumns
                ProgBar01.Stop();
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
    }
}
