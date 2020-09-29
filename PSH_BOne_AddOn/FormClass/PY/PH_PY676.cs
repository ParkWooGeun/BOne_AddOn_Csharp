using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 근태시간내역조회
    /// </summary>
    internal class PH_PY676 : PSH_BaseClass
    {
        public string oFormUniqueID;
        public SAPbouiCOM.Matrix oMat1;

        private SAPbouiCOM.DBDataSource oDS_PH_PY676B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        public string ItemUID { get; private set; }

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY676.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY676_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY676");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY676_CreateItems();
                PH_PY676_EnableMenus();
                PH_PY676_SetDocument(oFromDocEntry01);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// <returns></returns>
        private void PH_PY676_CreateItems()
        {
            string CLTCOD = string.Empty;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY676B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oMat1 = oForm.Items.Item("Mat01").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 기준일자
                oForm.DataSources.UserDataSources.Add("FrDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("FrDate").Specific.DataBind.SetBound(true, "", "FrDate");
                oForm.DataSources.UserDataSources.Item("FrDate").Value = DateTime.Now.ToString("yyyyMM") + "01";

                oForm.DataSources.UserDataSources.Add("ToDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("ToDate").Specific.DataBind.SetBound(true, "", "ToDate");
                oForm.DataSources.UserDataSources.Item("ToDate").Value = DateTime.Now.ToString("yyyyMMdd");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 성명
                oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTNAM").Specific.DataBind.SetBound(true, "", "MSTNAM");

                // 부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                // 담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");
                oForm.Items.Item("RspCode").DisplayDesc = true;

                // 반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");
                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY676_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY676_EnableMenus()
        {
            try
            {

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY676_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY676_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);         //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true);                             ////문서찾기
                    oForm.EnableMenu("1282", false);                            ////문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);         //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", false);                            ////문서찾기
                    oForm.EnableMenu("1282", true);                             ////문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);        //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    oForm.EnableMenu("1281", true);                             ////문서찾기
                    oForm.EnableMenu("1282", true);                             ////문서추가

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY676_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY676_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY676_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY676_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY676_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if ((BusinessObjectInfo.BeforeAction == true))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
                            break;
                    }
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
            string Code = string.Empty;

            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                if ((pVal.BeforeAction == true))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            break;
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

                        case "7169":
                            //엑셀 내보내기

                            //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            PH_PY676_AddMatrixRow();
                            break;

                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY676_FormItemEnabled();
                            PH_PY676_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //            Case "1293":
                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281":
                            ////문서찾기
                            PH_PY676_FormItemEnabled();
                            PH_PY676_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY676_FormItemEnabled();
                            PH_PY676_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY676_FormItemEnabled();
                            break;
                        case "1293":
                            //// 행삭제

                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while ((i <= oDS_PH_PY676B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY676B.GetValue("U_LineNum", i)))
                                    {
                                        oDS_PH_PY676B.RemoveRecord((i));
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY676B.Size; i++)
                                {
                                    oDS_PH_PY676B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            PH_PY676_AddMatrixRow();
                            break;

                        case "7169":
                            //엑셀 내보내기

                            //엑셀 내보내기 이후 처리
                            oForm.Freeze(true);
                            oDS_PH_PY676B.RemoveRecord(oDS_PH_PY676B.Size - 1);
                            oMat1.LoadFromDataSource();
                            oForm.Freeze(false);
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //    break;

                    ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
            string CLTCOD = string.Empty;
            string YM = string.Empty;
            string Code = string.Empty;
            string MSTCOD = string.Empty;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "BtnSearch")
                    {

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY676_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            PH_PY676_MTX01();
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        //if (pVal.ColUID == "Name" & pVal.CharPressed == Convert.ToDouble("9"))
                        //{
                        //    if (string.IsNullOrEmpty(oMat1.Columns.Item("Name").Cells.Item(pVal.Row).Specific.VALUE))
                        //    {
                        //        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                        //        BubbleEvent = false;
                        //    }
                        //}
                    }
                    //else if (pVal.ItemUID == "CntcCode" & pVal.CharPressed == Convert.ToDouble("9"))
                    //{
                    //    if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.VALUE))
                    //    {
                    //        PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                    //        BubbleEvent = false;
                    //    }
                    //}
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat1.SelectRow(pVal.Row, true, false);
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID = pVal.ItemUID;
                                oLastColUID = pVal.ColUID;
                                oLastColRow = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = "";
                            oLastColRow = 0;
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            int i = 0;
            SAPbobsCOM.Recordset oRecordSet = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {

                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            //사업장이 바뀌면 부서와 담당 재설정
                            case "CLTCOD":
                                ////부서
                                ////삭제
                                if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");

                                ////담당
                                ////삭제
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                ////반
                                ////삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;

                            ////부서가 바뀌면 담당 재설정
                            case "TeamCode":
                                ////담당은 그 부서의 담당만 표시
                                ////삭제
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.VALUE + "' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                ////반
                                ////삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;

                            ////담당이 바뀌면 반 재설정
                            case "RspCode":
                                ////반
                                ////삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;


                        }
                        oMat1.AutoResizeColumns();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// Raise_EVENT_VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "MSTCOD":
                                oForm.Items.Item("MSTNAM").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.VALUE + "'", "");
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    oMat1.LoadFromDataSource();

                    PH_PY676_FormItemEnabled();
                    PH_PY676_AddMatrixRow();
                    oMat1.AutoResizeColumns();

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                }
                else if (pVal.Before_Action == false)
                {
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY676B);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// RightClickEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID = pVal.ItemUID;
                            oLastColUID = pVal.ColUID;
                            oLastColRow = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID = pVal.ItemUID;
                        oLastColUID = "";
                        oLastColRow = 0;
                        break;
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

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        public void PH_PY676_AddMatrixRow()
        {
            int oRow = 0;
            try
            {
                oForm.Freeze(true);

                ////[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY676B.GetValue("U_LineNum", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY676B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY676B.InsertRecord((oRow));
                        }
                        oDS_PH_PY676B.Offset = oRow;
                        oDS_PH_PY676B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY676B.SetValue("U_ColDt01", oRow, "");   //일자
                        oDS_PH_PY676B.SetValue("U_ColReg01", oRow, "");   //사번
                        oDS_PH_PY676B.SetValue("U_ColReg02", oRow, "");   //성명
                        oDS_PH_PY676B.SetValue("U_ColReg12", oRow, "");   //직급
                        oDS_PH_PY676B.SetValue("U_ColReg03", oRow, "");   //부서
                        oDS_PH_PY676B.SetValue("U_ColReg04", oRow, "");   //담당
                        oDS_PH_PY676B.SetValue("U_ColReg05", oRow, "");  //반
                        oDS_PH_PY676B.SetValue("U_ColReg06", oRow, "");   //근무조
                        oDS_PH_PY676B.SetValue("U_ColReg07", oRow, "");  //요일
                        oDS_PH_PY676B.SetValue("U_ColReg08", oRow, "");  //요일구분
                        oDS_PH_PY676B.SetValue("U_ColQty01", oRow, "");     //기본
                        oDS_PH_PY676B.SetValue("U_ColQty02", oRow, "");   //연장
                        oDS_PH_PY676B.SetValue("U_ColQty03", oRow, "");  //특근
                        oDS_PH_PY676B.SetValue("U_ColQty04", oRow, ""); //특연
                        oDS_PH_PY676B.SetValue("U_ColQty05", oRow, ""); //심야
                        oDS_PH_PY676B.SetValue("U_ColQty06", oRow, "");  //조출
                        oDS_PH_PY676B.SetValue("U_ColQty07", oRow, ""); //휴조
                        oDS_PH_PY676B.SetValue("U_ColReg09", oRow, "");  //근무내용
                        oDS_PH_PY676B.SetValue("U_ColReg10", oRow, ""); //근태구분
                        oDS_PH_PY676B.SetValue("U_ColReg11", oRow, "");  //비고
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY676B.Offset = oRow - 1;
                        oDS_PH_PY676B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY676B.SetValue("U_ColDt01", oRow - 1, "");   //일자
                        oDS_PH_PY676B.SetValue("U_ColReg01", oRow - 1, "");   //사번
                        oDS_PH_PY676B.SetValue("U_ColReg02", oRow - 1, "");   //성명
                        oDS_PH_PY676B.SetValue("U_ColReg12", oRow - 1, "");   //직급
                        oDS_PH_PY676B.SetValue("U_ColReg03", oRow - 1, "");   //부서
                        oDS_PH_PY676B.SetValue("U_ColReg04", oRow - 1, "");   //담당
                        oDS_PH_PY676B.SetValue("U_ColReg05", oRow - 1, "");  //반
                        oDS_PH_PY676B.SetValue("U_ColReg06", oRow - 1, "");   //근무조
                        oDS_PH_PY676B.SetValue("U_ColReg07", oRow - 1, "");  //요일
                        oDS_PH_PY676B.SetValue("U_ColReg08", oRow - 1, "");  //요일구분
                        oDS_PH_PY676B.SetValue("U_ColQty01", oRow - 1, "");     //기본
                        oDS_PH_PY676B.SetValue("U_ColQty02", oRow - 1, "");   //연장
                        oDS_PH_PY676B.SetValue("U_ColQty03", oRow - 1, "");  //특근
                        oDS_PH_PY676B.SetValue("U_ColQty04", oRow - 1, ""); //특연
                        oDS_PH_PY676B.SetValue("U_ColQty05", oRow - 1, ""); //심야
                        oDS_PH_PY676B.SetValue("U_ColQty06", oRow - 1, "");  //조출
                        oDS_PH_PY676B.SetValue("U_ColQty07", oRow - 1, ""); //휴조
                        oDS_PH_PY676B.SetValue("U_ColReg09", oRow - 1, "");  //근무내용
                        oDS_PH_PY676B.SetValue("U_ColReg10", oRow - 1, ""); //근태구분
                        oDS_PH_PY676B.SetValue("U_ColReg11", oRow - 1, "");  //비고
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY676B.Offset = oRow;
                    oDS_PH_PY676B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY676B.SetValue("U_ColDt01", oRow, "");   //일자
                    oDS_PH_PY676B.SetValue("U_ColReg01", oRow, "");   //사번
                    oDS_PH_PY676B.SetValue("U_ColReg02", oRow, "");   //성명
                    oDS_PH_PY676B.SetValue("U_ColReg12", oRow, "");   //직급
                    oDS_PH_PY676B.SetValue("U_ColReg03", oRow, "");   //부서
                    oDS_PH_PY676B.SetValue("U_ColReg04", oRow, "");   //담당
                    oDS_PH_PY676B.SetValue("U_ColReg05", oRow, "");  //반
                    oDS_PH_PY676B.SetValue("U_ColReg06", oRow, "");   //근무조
                    oDS_PH_PY676B.SetValue("U_ColReg07", oRow, "");  //요일
                    oDS_PH_PY676B.SetValue("U_ColReg08", oRow, "");  //요일구분
                    oDS_PH_PY676B.SetValue("U_ColQty01", oRow, "");     //기본
                    oDS_PH_PY676B.SetValue("U_ColQty02", oRow, "");   //연장
                    oDS_PH_PY676B.SetValue("U_ColQty03", oRow, "");  //특근
                    oDS_PH_PY676B.SetValue("U_ColQty04", oRow, ""); //특연
                    oDS_PH_PY676B.SetValue("U_ColQty05", oRow, ""); //심야
                    oDS_PH_PY676B.SetValue("U_ColQty06", oRow, "");  //조출
                    oDS_PH_PY676B.SetValue("U_ColQty07", oRow, ""); //휴조
                    oDS_PH_PY676B.SetValue("U_ColReg09", oRow, "");  //근무내용
                    oDS_PH_PY676B.SetValue("U_ColReg10", oRow, ""); //근태구분
                    oDS_PH_PY676B.SetValue("U_ColReg11", oRow, "");  //비고
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY676_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY676_DataValidCheck()
        {
            bool functionReturnValue = false;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //년도
                //if (string.IsNullOrEmpty(oForm.Items.Item("RpmtDate").Specific.VALUE.ToString().Trim()))
                //{
                //    PSH_Globals.SBO_Application.SetStatusBarMessage("상환일자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                //    oForm.Items.Item("RpmtDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                //    functionReturnValue = false;
                //    return functionReturnValue;
                //}

                oMat1.FlushToDataSource();
                //// Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY676B.Size > 1)
                    oDS_PH_PY676B.RemoveRecord((oDS_PH_PY676B.Size - 1));

                oMat1.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY676_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
            functionReturnValue = true;
            return functionReturnValue;
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY676_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY676'", "");
                if (Convert.ToInt32(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY676_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PH_PY676_MTX01()
        {
            short i = 0;
            string sQry = null;
            short ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet01 = null;
            oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string CLTCOD = string.Empty;
            string FrDate = string.Empty;
            string ToDate = string.Empty;
            string MSTCOD = string.Empty;
            string TeamCode = string.Empty;
            string RspCode = string.Empty;
            string ClsCode = string.Empty;

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim(); //사업장
            FrDate = oForm.Items.Item("FrDate").Specific.VALUE.Trim(); //기간(시작)
            ToDate = oForm.Items.Item("ToDate").Specific.VALUE.Trim(); //기간(종료)
            MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim(); //사원번호
            TeamCode = oForm.Items.Item("TeamCode").Specific.VALUE.Trim(); //부서
            RspCode = oForm.Items.Item("RspCode").Specific.VALUE.Trim(); //담당
            ClsCode = oForm.Items.Item("ClsCode").Specific.VALUE.Trim(); //반

            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

            oForm.Freeze(true);

            sQry = "EXEC PH_PY676_01 '" + CLTCOD + "', '" + FrDate + "', '" + ToDate + "', '" + MSTCOD + "', '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "'";

            try
            {
                oRecordSet01.DoQuery(sQry);

                oMat1.Clear();
                oDS_PH_PY676B.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    ErrNum = 1;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY676B.Size)
                    {
                        oDS_PH_PY676B.InsertRecord(i);
                    }

                    oMat1.AddRow();
                    oDS_PH_PY676B.Offset = i;

                    oDS_PH_PY676B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY676B.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet01.Fields.Item("PosDate").Value.ToString().Trim()).ToString("yyyyMMdd"));   //일자
                    oDS_PH_PY676B.SetValue("U_ColReg01", i, oRecordSet01.Fields.Item("MSTCOD").Value);   //사번
                    oDS_PH_PY676B.SetValue("U_ColReg02", i, oRecordSet01.Fields.Item("MSTNAM").Value);   //성명
                    oDS_PH_PY676B.SetValue("U_ColReg12", i, oRecordSet01.Fields.Item("JIGNAM").Value);   //직급
                    oDS_PH_PY676B.SetValue("U_ColReg03", i, oRecordSet01.Fields.Item("TeamName").Value); //부서
                    oDS_PH_PY676B.SetValue("U_ColReg04", i, oRecordSet01.Fields.Item("RspName").Value);  //담당
                    oDS_PH_PY676B.SetValue("U_ColReg05", i, oRecordSet01.Fields.Item("ClsName").Value);  //반
                    oDS_PH_PY676B.SetValue("U_ColReg06", i, oRecordSet01.Fields.Item("GNMUJO").Value);   //근무조
                    oDS_PH_PY676B.SetValue("U_ColReg07", i, oRecordSet01.Fields.Item("DayWeek").Value);  //요일
                    oDS_PH_PY676B.SetValue("U_ColReg08", i, oRecordSet01.Fields.Item("DayType").Value);  //요일구분
                    oDS_PH_PY676B.SetValue("U_ColQty01", i, oRecordSet01.Fields.Item("Base").Value);     //기본
                    oDS_PH_PY676B.SetValue("U_ColQty02", i, oRecordSet01.Fields.Item("Extend").Value);   //연장
                    oDS_PH_PY676B.SetValue("U_ColQty03", i, oRecordSet01.Fields.Item("Special").Value);  //특근
                    oDS_PH_PY676B.SetValue("U_ColQty04", i, oRecordSet01.Fields.Item("SpExtend").Value); //특연
                    oDS_PH_PY676B.SetValue("U_ColQty05", i, oRecordSet01.Fields.Item("MidNight").Value); //심야
                    oDS_PH_PY676B.SetValue("U_ColQty06", i, oRecordSet01.Fields.Item("EarlyTo").Value);  //조출
                    oDS_PH_PY676B.SetValue("U_ColQty07", i, oRecordSet01.Fields.Item("SEarlyTo").Value); //휴조
                    oDS_PH_PY676B.SetValue("U_ColReg09", i, oRecordSet01.Fields.Item("ActText").Value);  //근무내용
                    oDS_PH_PY676B.SetValue("U_ColReg10", i, oRecordSet01.Fields.Item("WorkType").Value); //근태구분
                    oDS_PH_PY676B.SetValue("U_ColReg11", i, oRecordSet01.Fields.Item("Comment").Value);  //비고

                    oRecordSet01.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";

                }

                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                ProgBar01.Stop();

            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY011_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY676_Validate(string ValidateType)
        {
            bool functionReturnValue;

            functionReturnValue = false;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY676A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    return functionReturnValue;
                }

                if (ValidateType == "수정")
                {

                }
                else if (ValidateType == "행삭제")
                {

                }
                else if (ValidateType == "취소")
                {

                }
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY676_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }

            return functionReturnValue;
        }
    }
}
