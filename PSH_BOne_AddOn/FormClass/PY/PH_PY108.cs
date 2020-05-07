using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 상여지급률설정
    /// </summary>
    internal class PH_PY108 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY108A;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        string g_preBankSel;

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY108.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY108_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY108");

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
                PH_PY108_CreateItems();
                PH_PY108_EnableMenus();
                PH_PY108_SetDocument(oFromDocEntry01);
                oForm.Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY108_CreateItems()
        {
            string sQry = string.Empty;
            int i = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY108A = oForm.DataSources.DBDataSources.Item("@PH_PY108A");

                ////헤더

                ////사업장
                //사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                ////지급구분
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y' order by 1";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "");
                oForm.Items.Item("JOBGBN").DisplayDesc = true;

                ////상여계산방법
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P215' AND U_UseYN= 'Y' order by 1";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("BNSCAL").Specific, "");
                oForm.Items.Item("BNSCAL").DisplayDesc = true;

                //// 지급기준
                for (i = 1; i <= 8; i++)
                {
                    oForm.Items.Item("AP" + i + "GBN").Specific.ValidValues.Add("1", "개월 이상");
                    oForm.Items.Item("AP" + i + "GBN").Specific.ValidValues.Add("2", "개월 미만");
                    oForm.Items.Item("AP" + i + "GBN").Specific.ValidValues.Add("3", "일수 이상");
                    oForm.Items.Item("AP" + i + "GBN").Specific.ValidValues.Add("4", "일수 미만");
                    oForm.Items.Item("AP" + i + "GBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY108_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                ////제거
                oForm.EnableMenu("1284", false);               ////취소
                oForm.EnableMenu("1287", true);                ////복제
                oForm.EnableMenu("1293", false);               ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY108_SetDocument(string oFromDocEntry01)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY108_FormItemEnabled();
                    //dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY108_FormItemEnabled();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY108_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY108_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {

                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("JOBGBN").Enabled = true;
                    oForm.Items.Item("BNSRAT").Enabled = true;
                    oForm.Items.Item("BNSCAL").Enabled = true;
                    oForm.Items.Item("BNSMON").Enabled = true;

                    oDS_PH_PY108A.SetValue("U_BNSCAL", 0, "1");
                    oDS_PH_PY108A.SetValue("U_BNSMON", 0, Convert.ToString(0));

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD",true);

                    oForm.EnableMenu("1281", true);
                    ////문서찾기
                    oForm.EnableMenu("1282", false);
                    ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {

                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("JOBGBN").Enabled = true;
                    oForm.Items.Item("BNSRAT").Enabled = true;
                    oForm.Items.Item("BNSCAL").Enabled = true;
                    oForm.Items.Item("BNSMON").Enabled = true;

                    oForm.EnableMenu("1281", false);
                    ////문서찾기
                    oForm.EnableMenu("1282", true);
                    ////문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {

                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("JOBGBN").Enabled = false;
                    oForm.Items.Item("BNSRAT").Enabled = false;
                    oForm.Items.Item("BNSCAL").Enabled = false;
                    oForm.Items.Item("BNSMON").Enabled = false;

                    oForm.EnableMenu("1281", true);
                    ////문서찾기
                    oForm.EnableMenu("1282", true);
                    ////문서추가

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY108_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
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
                                //                            oForm.Mode = fm_OK_MODE
                                //                             Call Sbo_Application.ActivateMenuItem("1282")
                            }
                        }
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
                    case "Mat1":
                    case "Grid1":
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
                        case "Mat1":
                        case "Grid1":
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY108A);
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
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
                            dataHelpClass.AuthorityCheck( oForm,  "CLTCOD",  "@PH_PY108A",  "Code");
                            ////접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY108_FormItemEnabled();
                            PH_PY108_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //            Case "1293":
                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281":
                            ////문서찾기
                            PH_PY108_FormItemEnabled();
                            PH_PY108_AddMatrixRow();
                            break;
                        //                oForm.Items("Code").CLICK ct_Regular
                        case "1282":
                            ////문서추가
                            PH_PY108_FormItemEnabled();
                            PH_PY108_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY108_FormItemEnabled();
                            break;
                        case "1293":
                            //// 행삭제
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
                    case "Mat1":
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
        public bool PH_PY108_DataValidCheck()
        {
            bool functionReturnValue = false;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY108A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                /// 상여율
                if (Convert.ToInt32(oDS_PH_PY108A.GetValue("U_BNSRAT", 0)) == 0)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상여율은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("BNSRAT").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                /// 상여계산 방법
                if (string.IsNullOrEmpty(oDS_PH_PY108A.GetValue("U_BNSCAL", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상여계산방법은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("BNSCAL").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                /// 상여지급 주기
                if (Convert.ToInt32(oDS_PH_PY108A.GetValue("U_BNSMON", 0)) < 0 | Convert.ToInt32(oDS_PH_PY108A.GetValue("U_BNSMON", 0)) > 12)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상여지급 주기필수이며, 0 ~ 12개월 범위 이내이어야 합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("BNSMON").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                oDS_PH_PY108A.SetValue("Code", 0, oDS_PH_PY108A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY108A.GetValue("U_JOBGBN", 0).ToString().Trim());
                oDS_PH_PY108A.SetValue("Name", 0, oDS_PH_PY108A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY108A.GetValue("U_JOBGBN", 0).ToString().Trim());

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (dataHelpClass.Value_ChkYn("[@PH_PY108A]", "Code", "'" + oDS_PH_PY108A.GetValue("Code", 0).ToString().Trim() + "'","") == false)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short);
                        return functionReturnValue;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                functionReturnValue = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }
        
        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        public void PH_PY108_AddMatrixRow()
        {
            int oRow = 0;
            oForm.Freeze(true);
            try
            {
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}

//using Microsoft.VisualBasic;
//using Microsoft.VisualBasic.Compatibility;
//using System;
//using System.Collections;
//using System.Data;
//using System.Diagnostics;
//using System.Drawing;
//using System.Windows.Forms;
// // ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//    internal class PH_PY108
//    {
//        ////********************************************************************************
//        ////  File           : PH_PY108.cls
//        ////  Module         : 인사관리 > 급여관리
//        ////  Desc           : 상여지급율설정
//        ////********************************************************************************

//        public string oFormUniqueID;
//        public SAPbouiCOM.Form oForm;

//        private SAPbouiCOM.DBDataSource oDS_PH_PY108A;

//        private string oLastItemUID;
//        private string oLastColUID;
//        private int oLastColRow;

//        public void LoadForm(string oFromDocEntry01 = "")
//        {

//            int i = 0;
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            // ERROR: Not supported in C#: OnErrorStatement


//            oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY108.srf");
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//            for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//            {
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//            }
//            oFormUniqueID = "PH_PY108_" + GetTotalFormsCount();
//            SubMain.AddForms(this, oFormUniqueID, "PH_PY108");
//            PSH_Globals.SBO_Application.LoadBatchActions(out (oXmlDoc.xml));

//            oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//            oForm.SupportedModes = -1;
//            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

//            //***************************************************************
//            //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
//            oForm.DataBrowser.BrowseBy = "Code";
//            //***************************************************************

//            oForm.Freeze(true);
//            PH_PY108_CreateItems();
//            PH_PY108_EnableMenus();
//            PH_PY108_SetDocument(oFromDocEntry01);
//            //    Call PH_PY108_FormResize

//            oForm.Update();
//            oForm.Freeze(false);

//            oForm.Visible = true;
//            //UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oXmlDoc = null;
//            return;
//        LoadForm_Error:

//            oForm.Update();
//            oForm.Freeze(false);
//            //UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oXmlDoc = null;
//            //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oForm = null;
//            PSH_Globals.SBO_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private bool PH_PY108_CreateItems()
//        {
//            bool functionReturnValue = false;

//            string sQry = null;
//            int i = 0;

//            SAPbouiCOM.CheckBox oCheck = null;
//            SAPbouiCOM.EditText oEdit = null;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbouiCOM.Column oColumn = null;
//            SAPbouiCOM.Columns oColumns = null;
//            SAPbouiCOM.OptionBtn optBtn = null;

//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            oDS_PH_PY108A = oForm.DataSources.DBDataSources("@PH_PY108A");
//            ////헤더

//            ////사업장
//            oCombo = oForm.Items.Item("CLTCOD").Specific;
//            //    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//            //    Call SetReDataCombo(oForm, sQry, oCombo)
//            oForm.Items.Item("CLTCOD").DisplayDesc = true;

//            ////지급구분
//            oCombo = oForm.Items.Item("JOBGBN").Specific;
//            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y'";
//            MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//            oForm.Items.Item("JOBGBN").DisplayDesc = true;

//            ////상여계산방법
//            oCombo = oForm.Items.Item("BNSCAL").Specific;
//            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P215' AND U_UseYN= 'Y'";
//            MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//            oForm.Items.Item("BNSCAL").DisplayDesc = true;

//            //// 지급기준
//            for (i = 1; i <= 8; i++)
//            {
//                oCombo = oForm.Items.Item("AP" + i + "GBN").Specific;
//                oCombo.ValidValues.Add("1", "개월 이상");
//                oCombo.ValidValues.Add("2", "개월 미만");
//                oCombo.ValidValues.Add("3", "일수 이상");
//                oCombo.ValidValues.Add("4", "일수 미만");
//                oCombo.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
//            }

//            //UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCheck = null;
//            //UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oEdit = null;
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumn = null;
//            //UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumns = null;
//            //UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            optBtn = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            return functionReturnValue;
//        PH_PY108_CreateItems_Error:

//            //UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCheck = null;
//            //UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oEdit = null;
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumn = null;
//            //UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oColumns = null;
//            //UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            optBtn = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            oForm.Freeze(false);
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        private void PH_PY108_EnableMenus()
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.EnableMenu("1283", true);
//            ////제거
//            oForm.EnableMenu("1284", false);
//            ////취소
//            oForm.EnableMenu("1287", true);
//            ////복제
//            oForm.EnableMenu("1293", false);
//            ////행삭제

//            return;
//        PH_PY108_EnableMenus_Error:

//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY108_SetDocument(string oFromDocEntry01)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            if ((string.IsNullOrEmpty(oFromDocEntry01)))
//            {
//                PH_PY108_FormItemEnabled();

//            }
//            else
//            {
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                PH_PY108_FormItemEnabled();
//                //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("Code").Specific.Value = oFromDocEntry01;
//                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            return;
//        PH_PY108_SetDocument_Error:

//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY108_FormItemEnabled()
//        {
//            SAPbouiCOM.ComboBox oCombo = null;

//            // ERROR: Not supported in C#: OnErrorStatement



//            oForm.Freeze(true);
//            if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//            {
//                oForm.Items.Item("CLTCOD").Enabled = true;
//                oForm.Items.Item("JOBGBN").Enabled = true;
//                oForm.Items.Item("BNSRAT").Enabled = true;
//                oForm.Items.Item("BNSCAL").Enabled = true;
//                oForm.Items.Item("BNSMON").Enabled = true;

//                oDS_PH_PY108A.SetValue("U_BNSCAL", 0, "1");
//                oDS_PH_PY108A.SetValue("U_BNSMON", 0, Convert.ToString(0));

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//                oForm.EnableMenu("1281", true);
//                ////문서찾기
//                oForm.EnableMenu("1282", false);
//                ////문서추가

//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//            {
//                oForm.Items.Item("CLTCOD").Enabled = true;
//                oForm.Items.Item("JOBGBN").Enabled = true;
//                oForm.Items.Item("BNSRAT").Enabled = true;
//                oForm.Items.Item("BNSCAL").Enabled = true;
//                oForm.Items.Item("BNSMON").Enabled = true;

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//                oForm.EnableMenu("1281", false);
//                ////문서찾기
//                oForm.EnableMenu("1282", true);
//                ////문서추가

//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//            {
//                oForm.Items.Item("CLTCOD").Enabled = false;
//                oForm.Items.Item("JOBGBN").Enabled = false;
//                oForm.Items.Item("BNSRAT").Enabled = false;
//                oForm.Items.Item("BNSCAL").Enabled = false;
//                oForm.Items.Item("BNSMON").Enabled = false;

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//                oForm.EnableMenu("1281", true);
//                ////문서찾기
//                oForm.EnableMenu("1282", true);
//                ////문서추가


//            }
//            oForm.Freeze(false);
//            return;
//        PH_PY108_FormItemEnabled_Error:

//            oForm.Freeze(false);
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }


//        public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            string sQry = null;
//            int i = 0;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            switch (pVal.EventType)
//            {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//                    ////1

//                    if (pVal.BeforeAction == true)
//                    {
//                        if (pVal.ItemUID == "1")
//                        {
//                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
//                            {
//                                if (PH_PY108_DataValidCheck() == false)
//                                {
//                                    BubbleEvent = false;
//                                }
//                            }
//                        }
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemUID == "1")
//                        {
//                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                            {
//                                if (pVal.ActionSuccess == true)
//                                {
//                                    //                            oForm.Mode = fm_OK_MODE
//                                    //                             Call Sbo_Application.ActivateMenuItem("1282")
//                                }
//                            }
//                        }
//                    }
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                    ////2
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                    ////3
//                    switch (pVal.ItemUID)
//                    {
//                        case "Mat1":
//                        case "Grid1":
//                            if (pVal.Row > 0)
//                            {
//                                oLastItemUID = pVal.ItemUID;
//                                oLastColUID = pVal.ColUID;
//                                oLastColRow = pVal.Row;
//                            }
//                            break;
//                        default:
//                            oLastItemUID = pVal.ItemUID;
//                            oLastColUID = "";
//                            oLastColRow = 0;
//                            break;
//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//                    ////4
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//                    ////5
//                    break;
//                //            Call oForm.Freeze(True)
//                //            If pVal.BeforeAction = True Then
//                //
//                //            ElseIf pVal.BeforeAction = False Then
//                //                If pVal.ItemChanged = True Then
//                //
//                //                End If
//                //            End If
//                //            Call oForm.Freeze(False)
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_CLICK:
//                    ////6
//                    if (pVal.BeforeAction == true)
//                    {
//                        switch (pVal.ItemUID)
//                        {
//                            case "Mat1":
//                            case "Grid1":
//                                if (pVal.Row > 0)
//                                {
//                                    oLastItemUID = pVal.ItemUID;
//                                    oLastColUID = pVal.ColUID;
//                                    oLastColRow = pVal.Row;
//                                }
//                                break;
//                            default:
//                                oLastItemUID = pVal.ItemUID;
//                                oLastColUID = "";
//                                oLastColRow = 0;
//                                break;
//                        }
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//                    ////7
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//                    ////8
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//                    ////9
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//                    ////10
//                    break;
//                //            Call oForm.Freeze(True)
//                //            If pVal.BeforeAction = True Then
//                //
//                //            ElseIf pVal.BeforeAction = False Then
//                //                If pVal.ItemChanged = True Then
//                //
//                //                End If
//                //            End If
//                //            Call oForm.Freeze(False)
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//                    ////11
//                    break;
//                //            If pVal.BeforeAction = True Then
//                //            ElseIf pVal.BeforeAction = False Then
//                //                oMat1.LoadFromDataSource
//                //
//                //                Call PH_PY108_FormItemEnabled
//                //                Call PH_PY108_AddMatrixRow
//                //
//                //            End If
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//                    ////12
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//                    ////16
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//                    ////17
//                    if (pVal.BeforeAction == true)
//                    {
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        SubMain.RemoveForms(oFormUniqueID);
//                        //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oForm = null;
//                        //UPGRADE_NOTE: oDS_PH_PY108A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oDS_PH_PY108A = null;

//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//                    ////18
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//                    ////19
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//                    ////20
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//                    ////21
//                    break;
//                //            If pVal.BeforeAction = True Then
//                //
//                //            ElseIf pVal.BeforeAction = False Then
//                //
//                //            End If
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//                    ////22
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//                    ////23
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//                    ////27
//                    break;
//                //            If pVal.BeforeAction = True Then
//                //
//                //            ElseIf pVal.Before_Action = False Then
//                //
//                //            End If
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//                    ////37
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//                    ////38
//                    break;

//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_Drag:
//                    ////39
//                    break;

//            }

//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;

//            return;
//        Raise_FormItemEvent_Error:
//            ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//            oForm.Freeze((false));
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            PSH_Globals.SBO_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }


//        public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//        {
//            int i = 0;
//            // ERROR: Not supported in C#: OnErrorStatement

//            oForm.Freeze(true);

//            if ((pVal.BeforeAction == true))
//            {
//                switch (pVal.MenuUID)
//                {
//                    case "1283":
//                        if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
//                        {
//                            BubbleEvent = false;
//                            return;
//                        }
//                        break;
//                    case "1284":
//                        break;
//                    case "1286":
//                        break;
//                    case "1293":
//                        break;
//                    case "1281":
//                        break;
//                    case "1282":
//                        break;
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@PH_PY108A", ref "Code");
//                        ////접속자 권한에 따른 사업장 보기
//                        break;
//                }
//            }
//            else if ((pVal.BeforeAction == false))
//            {
//                switch (pVal.MenuUID)
//                {
//                    case "1283":
//                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                        PH_PY108_FormItemEnabled();
//                        PH_PY108_AddMatrixRow();
//                        break;
//                    case "1284":
//                        break;
//                    case "1286":
//                        break;
//                    //            Case "1293":
//                    //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
//                    case "1281":
//                        ////문서찾기
//                        PH_PY108_FormItemEnabled();
//                        PH_PY108_AddMatrixRow();
//                        break;
//                    //                oForm.Items("Code").CLICK ct_Regular
//                    case "1282":
//                        ////문서추가
//                        PH_PY108_FormItemEnabled();
//                        PH_PY108_AddMatrixRow();
//                        break;
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        PH_PY108_FormItemEnabled();
//                        break;
//                    case "1293":
//                        //// 행삭제
//                        break;
//                        //                '// [MAT1 용]
//                        //                 If oMat1.RowCount <> oMat1.VisualRowCount Then
//                        //                    oMat1.FlushToDataSource
//                        //
//                        //                    While (i <= oDS_PH_PY108B.Size - 1)
//                        //                        If oDS_PH_PY108B.GetValue("U_FILD01", i) = "" Then
//                        //                            oDS_PH_PY108B.RemoveRecord (i)
//                        //                            i = 0
//                        //                        Else
//                        //                            i = i + 1
//                        //                        End If
//                        //                    Wend
//                        //
//                        //                    For i = 0 To oDS_PH_PY108B.Size
//                        //                        Call oDS_PH_PY108B.setValue("U_LineNum", i, i + 1)
//                        //                    Next i
//                        //
//                        //                    oMat1.LoadFromDataSource
//                        //                End If
//                        //                Call PH_PY108_AddMatrixRow
//                }
//            }
//            oForm.Freeze(false);
//            return;
//        Raise_FormMenuEvent_Error:
//            oForm.Freeze(false);
//            PSH_Globals.SBO_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            if ((BusinessObjectInfo.BeforeAction == true))
//            {
//                switch (BusinessObjectInfo.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                        ////33
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                        ////34
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                        ////35
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                        ////36
//                        break;
//                }
//            }
//            else if ((BusinessObjectInfo.BeforeAction == false))
//            {
//                switch (BusinessObjectInfo.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                        ////33
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                        ////34
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                        ////35
//                        break;
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                        ////36
//                        break;
//                }
//            }
//            return;
//        Raise_FormDataEvent_Error:


//            PSH_Globals.SBO_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//        }

//        public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            if (pVal.BeforeAction == true)
//            {
//            }
//            else if (pVal.BeforeAction == false)
//            {
//            }
//            switch (pVal.ItemUID)
//            {
//                case "Mat1":
//                    if (pVal.Row > 0)
//                    {
//                        oLastItemUID = pVal.ItemUID;
//                        oLastColUID = pVal.ColUID;
//                        oLastColRow = pVal.Row;
//                    }
//                    break;
//                default:
//                    oLastItemUID = pVal.ItemUID;
//                    oLastColUID = "";
//                    oLastColRow = 0;
//                    break;
//            }
//            return;
//        Raise_RightClickEvent_Error:

//            PSH_Globals.SBO_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY108_AddMatrixRow()
//        {
//            int oRow = 0;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            //    '//[Mat1 용]
//            //    oMat1.FlushToDataSource
//            //    oRow = oMat1.VisualRowCount
//            //
//            //    If oMat1.VisualRowCount > 0 Then
//            //        If Trim(oDS_PH_PY108B.GetValue("U_FILD01", oRow - 1)) <> "" Then
//            //            If oDS_PH_PY108B.Size <= oMat1.VisualRowCount Then
//            //                oDS_PH_PY108B.InsertRecord (oRow)
//            //            End If
//            //            oDS_PH_PY108B.Offset = oRow
//            //            oDS_PH_PY108B.setValue "U_LineNum", oRow, oRow + 1
//            //            oDS_PH_PY108B.setValue "U_FILD01", oRow, ""
//            //            oDS_PH_PY108B.setValue "U_FILD02", oRow, ""
//            //            oDS_PH_PY108B.setValue "U_FILD03", oRow, 0
//            //            oMat1.LoadFromDataSource
//            //        Else
//            //            oDS_PH_PY108B.Offset = oRow - 1
//            //            oDS_PH_PY108B.setValue "U_LineNum", oRow - 1, oRow
//            //            oDS_PH_PY108B.setValue "U_FILD01", oRow - 1, ""
//            //            oDS_PH_PY108B.setValue "U_FILD02", oRow - 1, ""
//            //            oDS_PH_PY108B.setValue "U_FILD03", oRow - 1, 0
//            //            oMat1.LoadFromDataSource
//            //        End If
//            //    ElseIf oMat1.VisualRowCount = 0 Then
//            //        oDS_PH_PY108B.Offset = oRow
//            //        oDS_PH_PY108B.setValue "U_LineNum", oRow, oRow + 1
//            //        oDS_PH_PY108B.setValue "U_FILD01", oRow, ""
//            //        oDS_PH_PY108B.setValue "U_FILD02", oRow, ""
//            //        oDS_PH_PY108B.setValue "U_FILD03", oRow, 0
//            //        oMat1.LoadFromDataSource
//            //    End If

//            oForm.Freeze(false);
//            return;
//        PH_PY108_AddMatrixRow_Error:
//            oForm.Freeze(false);
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY108_FormClear()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            string DocEntry = null;
//            //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY108'", ref "");
//            if (Convert.ToDouble(DocEntry) == 0)
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("DocEntry").Specific.Value = 1;
//            }
//            else
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
//            }
//            return;
//        PH_PY108_FormClear_Error:
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public bool PH_PY108_DataValidCheck()
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement


//            int i = 0;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            functionReturnValue = false;

//            //    '//----------------------------------------------------------------------------------
//            //    '//필수 체크
//            //    '//----------------------------------------------------------------------------------
//            /// 지급구분
//            if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY108A.GetValue("U_CLTCOD", 0))))
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                return functionReturnValue;
//            }

//            /// 상여율
//            if (Conversion.Val(oDS_PH_PY108A.GetValue("U_BNSRAT", 0)) == 0)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("상여율은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oForm.Items.Item("BNSRAT").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                return functionReturnValue;
//            }

//            /// 상여계산 방법
//            if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY108A.GetValue("U_BNSCAL", 0))))
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("상여계산방법은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oForm.Items.Item("BNSCAL").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                return functionReturnValue;
//            }

//            /// 상여지급 주기
//            if (Conversion.Val(oDS_PH_PY108A.GetValue("U_BNSMON", 0)) < 0 | Conversion.Val(oDS_PH_PY108A.GetValue("U_BNSMON", 0)) > 12)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("상여지급 주기필수이며, 0 ~ 12개월 범위 이내이어야 합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                oForm.Items.Item("BNSMON").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                return functionReturnValue;
//            }

//            oDS_PH_PY108A.SetValue("Code", 0, Strings.Trim(oDS_PH_PY108A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY108A.GetValue("U_JOBGBN", 0)));
//            oDS_PH_PY108A.SetValue("Name", 0, Strings.Trim(oDS_PH_PY108A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY108A.GetValue("U_JOBGBN", 0)));

//            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//            {
//                if (MDC_SetMod.Value_ChkYn("[@PH_PY108A]", "Code", "'" + Strings.Trim(oDS_PH_PY108A.GetValue("Code", 0)) + "'") == false)
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                    return functionReturnValue;
//                }
//            }

//            functionReturnValue = true;
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY108_DataValidCheck_Error:


//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            functionReturnValue = false;
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }


//        public bool PH_PY108_Validate(string ValidateType)
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement

//            functionReturnValue = true;
//            object i = null;
//            int j = 0;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;
//            oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            //UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY108A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY108A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                functionReturnValue = false;
//                goto PH_PY108_Validate_Exit;
//            }
//            //
//            if (ValidateType == "수정")
//            {

//            }
//            else if (ValidateType == "행삭제")
//            {

//            }
//            else if (ValidateType == "취소")
//            {

//            }
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY108_Validate_Exit:
//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return functionReturnValue;
//        PH_PY108_Validate_Error:
//            functionReturnValue = false;
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY108_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }
//    }
//}
