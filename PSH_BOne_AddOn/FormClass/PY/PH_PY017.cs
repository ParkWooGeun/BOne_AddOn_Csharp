
using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 월근태집계처리
    /// </summary>
    internal class PH_PY017 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY017A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY017B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        public string ItemUID { get; private set; }

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY017.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY017_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY017");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                //***************************************************************
                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                oForm.DataBrowser.BrowseBy = "Code";
                //***************************************************************

                oForm.Freeze(true);
                PH_PY017_CreateItems();
                PH_PY017_EnableMenus();
                PH_PY017_SetDocument(oFormDocEntry01);
                oForm.Update();
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
        private void PH_PY017_CreateItems()
        {
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oDS_PH_PY017A = oForm.DataSources.DBDataSources.Item("@PH_PY017A");
                oDS_PH_PY017B = oForm.DataSources.DBDataSources.Item("@PH_PY017B");

                oMat1 = oForm.Items.Item("Mat01").Specific;
                ////@PH_PY017B

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                //// 사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //// 접속자에 따른 사업장 선택
                CLTCOD = dataHelpClass.Get_ReData("Branch", "USER_CODE", "OUSR", "'" + PSH_Globals.oCompany.UserName + "'","");
                oDS_PH_PY017A.SetValue("U_CLTCOD", 0, CLTCOD);

                oDS_PH_PY017A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY017_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY017_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1281", true);                ////찾기
                oForm.EnableMenu("1282", true);                ////신규
                oForm.EnableMenu("1283", true);                ////제거
                oForm.EnableMenu("1284", false);                ////취소
                oForm.EnableMenu("1293", false);                ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY017_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY017_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFormDocEntry01)))
                {
                    PH_PY017_FormItemEnabled();
                    PH_PY017_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY017_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY017_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY017_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD",false);
                    ////년월
                    oDS_PH_PY017A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));

                    oMat1.Clear();
                    oForm.EnableMenu("1281", true);                    ////문서찾기
                    oForm.EnableMenu("1282", false);                    ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1281", false);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                    ////문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select( oForm,  "CLTCOD",  false);

                    oForm.EnableMenu("1281", true);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                    ////문서추가

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY017_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);

                    break;

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
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //    break;

                    ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    // case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    // break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    // break;

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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY017_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        ////해야할일 작업
                    }

                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY017_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY017_FormItemEnabled();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY017_FormItemEnabled();
                            }
                        }
                    }
                    if (pVal.ItemUID == "Btn_CREATE")
                    {
                        PH_PY017_ITEM_CREATE();
                    }

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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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

                    PH_PY017_FormItemEnabled();
                    PH_PY017_AddMatrixRow();

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
        /// Raise_EVENT_GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                oForm.Freeze(true);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY017A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY017B);
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            oForm.Freeze(true);
            try
            {
                oForm.Freeze(false);
                if (pVal.BeforeAction == true)
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
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY017_FormItemEnabled();
                            PH_PY017_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY017_FormItemEnabled();
                            PH_PY017_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY017_FormItemEnabled();
                            PH_PY017_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY017_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat1, oDS_PH_PY017B, "U_CODNBR"); 
                            PH_PY017_AddMatrixRow();
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
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            string sQry = string.Empty;

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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY017_DataValidCheck()
        {
            bool functionReturnValue = false;
            string errCode = string.Empty;

            try
            {
                string tCode = string.Empty;

                if (oMat1.VisualRowCount > 0)
                {
                }
                else
                {
                    errCode = "1";
                    throw new Exception();
                }

                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                tCode = Convert.ToString(Convert.ToDouble(oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim()) + oForm.Items.Item("YM").Specific.Value);
                oDS_PH_PY017A.SetValue("Code", 0, tCode);
                oDS_PH_PY017A.SetValue("Name", 0, tCode);

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY017_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        public void PH_PY017_AddMatrixRow()
        {
            int oRow = 0;
            
            try
            {
                oForm.Freeze(true);
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY017B.GetValue("U_MSTCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY017B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY017B.InsertRecord((oRow));
                        }
                        oDS_PH_PY017B.Offset = oRow;
                        oDS_PH_PY017B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY017B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY017B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY017B.SetValue("U_TeamCode", oRow, "");
                        oDS_PH_PY017B.SetValue("U_RspCode", oRow, "");
                        oDS_PH_PY017B.SetValue("U_StdGDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_StdPDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_StdNDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_GetDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_WoHDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_PayDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_AbsDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_Base", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_Extend", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_Midnight", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EarlyTo", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_Special", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SpExtend", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SMidnigh", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SEarlyTo", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EduTime", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_LateToC", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EarlyOfC", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_GoOutC", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_LateToT", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EarlyOfT", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_GoOutT", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_JCHDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_YCHDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_YCHHGA", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SNHDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SNHHGA", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_DNGDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_WHMDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY1", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY2", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY3", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY4", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY5", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY6", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY7", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY8", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY9", oRow, Convert.ToString(0));

                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY017B.Offset = oRow - 1;
                        oDS_PH_PY017B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY017B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY017B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY017B.SetValue("U_TeamCode", oRow, "");
                        oDS_PH_PY017B.SetValue("U_RspCode", oRow, "");
                        oDS_PH_PY017B.SetValue("U_StdGDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_StdPDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_StdNDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_GetDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_WoHDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_PayDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_AbsDay", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_Base", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_Extend", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_Midnight", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EarlyTo", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_Special", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SpExtend", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SMidnigh", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SEarlyTo", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EduTime", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_LateToC", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EarlyOfC", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_GoOutC", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_LateToT", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EarlyOfT", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_GoOutT", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_JCHDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_YCHDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_YCHHGA", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SNHDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_SNHHGA", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_DNGDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_WHMDAY", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY1", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY2", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY3", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY4", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY5", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY6", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY7", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY8", oRow, Convert.ToString(0));
                        oDS_PH_PY017B.SetValue("U_EtcDAY9", oRow, Convert.ToString(0));

                        oMat1.LoadFromDataSource();

                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY017B.Offset = oRow;
                    oDS_PH_PY017B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY017B.SetValue("U_MSTCOD", oRow, "");
                    oDS_PH_PY017B.SetValue("U_MSTNAM", oRow, "");
                    oDS_PH_PY017B.SetValue("U_TeamCode", oRow, "");
                    oDS_PH_PY017B.SetValue("U_RspCode", oRow, "");
                    oDS_PH_PY017B.SetValue("U_StdGDay", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_StdPDay", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_StdNDay", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_GetDay", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_WoHDay", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_PayDay", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_AbsDay", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_Base", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_Extend", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_Midnight", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EarlyTo", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_Special", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_SpExtend", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_SMidnigh", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_SEarlyTo", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EduTime", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_LateToC", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EarlyOfC", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_GoOutC", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_LateToT", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EarlyOfT", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_GoOutT", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_JCHDAY", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_YCHDAY", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_YCHHGA", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_SNHDAY", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_SNHHGA", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_DNGDAY", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_WHMDAY", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY1", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY2", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY3", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY4", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY5", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY6", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY7", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY8", oRow, Convert.ToString(0));
                    oDS_PH_PY017B.SetValue("U_EtcDAY9", oRow, Convert.ToString(0));

                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY017_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY017_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            functionReturnValue = true;

            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY017A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    ErrNumm = 1;
                    functionReturnValue = false;
                    throw new Exception();
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
            }
            catch (Exception ex)
            {
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }
            return functionReturnValue;
        }


        /// <summary>
        /// 행삭제(사용자 메소드로 구현)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        /// <param name="oMat">매트릭스 이름</param>
        /// <param name="DBData">DB데이터소스</param>
        /// <param name="CheckField">데이터 체크 필드명</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent, SAPbouiCOM.Matrix oMat, SAPbouiCOM.DBDataSource DBData, string CheckField)
        {
            int i = 0;

            try
            {
                if (oLastColRow > 0)
                {
                    if (pVal.BeforeAction == true)
                    {

                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oMat.RowCount != oMat.VisualRowCount)
                        {
                            oMat.FlushToDataSource();

                            while (i <= DBData.Size - 1)
                            {
                                if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i)))
                                {
                                    DBData.RemoveRecord((i));
                                    i = 0;
                                }
                                else
                                {
                                    i = i + 1;
                                }
                            }

                            for (i = 0; i <= DBData.Size; i++)
                            {
                                DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                            }

                            oMat.LoadFromDataSource();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ROW_DELETE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// PH_PY130_Create_Data
        /// </summary>
        public void PH_PY017_ITEM_CREATE()
        {
            int i = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string YM = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm = PSH_Globals.SBO_Application.Forms.ActiveForm;

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                YM = oForm.Items.Item("YM").Specific.Value;

                oForm.Freeze(true);

                sQry = "EXEC [PH_PY017_01] '" + CLTCOD + "', '" + YM + "'";
                oRecordSet.DoQuery(sQry);

                SAPbouiCOM.ProgressBar ProgressBar01 = null;
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("자료집계중!", oRecordSet.RecordCount, false);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    dataHelpClass.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.:", "W");
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY017B.Size)
                    {
                        oDS_PH_PY017B.InsertRecord((i));
                    }

                    oMat1.AddRow();
                    oDS_PH_PY017B.Offset = i;
                    oDS_PH_PY017B.SetValue("U_LineNum", i, Convert.ToString(i + 1));

                    oDS_PH_PY017B.SetValue("U_MSTCOD", i, oRecordSet.Fields.Item("MSTCOD").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_MSTNAM", i, oRecordSet.Fields.Item("FullName").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_TeamCode", i, oRecordSet.Fields.Item("TeamCode").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_RspCode", i, oRecordSet.Fields.Item("RspCode").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_ClsCode", i, oRecordSet.Fields.Item("ClsCode").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_StdGDay", i, oRecordSet.Fields.Item("StdGDay").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_StdPDay", i, oRecordSet.Fields.Item("StdPDay").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_StdNDay", i, oRecordSet.Fields.Item("StdNDay").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_GetDay", i, oRecordSet.Fields.Item("GetDay").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_WoHDay", i, oRecordSet.Fields.Item("WoHDay").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_PayDay", i, oRecordSet.Fields.Item("PayDay").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_AbsDay", i, oRecordSet.Fields.Item("AbsDay").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_Base", i, oRecordSet.Fields.Item("Base").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_Extend", i, oRecordSet.Fields.Item("Extend").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_Midnight", i, oRecordSet.Fields.Item("Midnight").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EarlyTo", i, oRecordSet.Fields.Item("EarlyTo").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_Special", i, oRecordSet.Fields.Item("Special").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_SpExtend", i, oRecordSet.Fields.Item("SpExtend").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_SMidnigh", i, oRecordSet.Fields.Item("SMidnigh").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_SEarlyTo", i, oRecordSet.Fields.Item("SEarlyTo").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EduTime", i, oRecordSet.Fields.Item("EduTime").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_LateToC", i, oRecordSet.Fields.Item("LateToC").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EarlyOfC", i, oRecordSet.Fields.Item("EarlyOfC").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_GoOutC", i, oRecordSet.Fields.Item("GoOutC").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_LateToT", i, oRecordSet.Fields.Item("LateToT").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EarlyOfT", i, oRecordSet.Fields.Item("EarlyOfT").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_GoOutT", i, oRecordSet.Fields.Item("GoOutT").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_JCHDAY", i, oRecordSet.Fields.Item("JCHDAY").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_YCHDAY", i, oRecordSet.Fields.Item("YCHDAY").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_YCHHGA", i, oRecordSet.Fields.Item("YCHHGA").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_SNHDAY", i, oRecordSet.Fields.Item("SNHDAY").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_SNHHGA", i, oRecordSet.Fields.Item("SNHHGA").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_DNGDAY", i, oRecordSet.Fields.Item("DNGDAY").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_WHMDAY", i, oRecordSet.Fields.Item("WHMDAY").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY1", i, oRecordSet.Fields.Item("EtcDAY1").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY2", i, oRecordSet.Fields.Item("EtcDAY2").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY3", i, oRecordSet.Fields.Item("EtcDAY3").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY4", i, oRecordSet.Fields.Item("EtcDAY4").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY5", i, oRecordSet.Fields.Item("EtcDAY5").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY6", i, oRecordSet.Fields.Item("EtcDAY6").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY7", i, oRecordSet.Fields.Item("EtcDAY7").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY8", i, oRecordSet.Fields.Item("EtcDAY8").Value.ToString().Trim());
                    oDS_PH_PY017B.SetValue("U_EtcDAY9", i, oRecordSet.Fields.Item("EtcDAY9").Value.ToString().Trim());

                    oRecordSet.MoveNext();

                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();

                PH_PY017_AddMatrixRow();
                ProgressBar01.Stop();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY017_ITEM_CREATE_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}
