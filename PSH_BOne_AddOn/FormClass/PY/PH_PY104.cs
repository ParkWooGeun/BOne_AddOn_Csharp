using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 고정수당공제금액일괄등록
    /// </summary>
    internal class PH_PY104 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.Grid oGrid2;
        public SAPbouiCOM.DataTable oDS_PH_PY104_01;
        public SAPbouiCOM.DataTable oDS_PH_PY104_02;
        ////그리드1의 체크 순번
        public int tSeqAll;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY104.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY104_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY104");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY104_CreateItems();
                PH_PY104_EnableMenus();
                PH_PY104_SetDocument(oFromDocEntry01);
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
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid1").Specific;
                oForm.DataSources.DataTables.Add("PH_PY104_01");
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("이름", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("선택", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("순서", SAPbouiCOM.BoFieldsType.ft_Float);
                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY104_01");
                oDS_PH_PY104_01 = oForm.DataSources.DataTables.Item("PH_PY104_01");

                oGrid2 = oForm.Items.Item("Grid2").Specific;
                oForm.DataSources.DataTables.Add("PH_PY104_02");
                oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("체크", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("이름", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oGrid2.DataTable = oForm.DataSources.DataTables.Item("PH_PY104_02");
                oDS_PH_PY104_02 = oForm.DataSources.DataTables.Item("PH_PY104_02");

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                // 담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
                oForm.Items.Item("RspCode").DisplayDesc = true;

                // 급여형태
                oForm.DataSources.UserDataSources.Add("PAYTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("PAYTYP").Specific.DataBind.SetBound(true, "", "PAYTYP");
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P132' AND U_UseYN= 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("PAYTYP").Specific, "");
                oForm.Items.Item("PAYTYP").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("PAYTYP").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("PAYTYP").DisplayDesc = true;

                // 직급형태From
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JIGCODF").Specific,"");
                oForm.Items.Item("JIGCODF").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("JIGCODF").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("JIGCODF").DisplayDesc = true;

                // 직급형태To
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JIGCODT").Specific,"");
                oForm.Items.Item("JIGCODT").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("JIGCODT").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("JIGCODT").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY104_EnableMenus
        /// </summary>
        public void PH_PY104_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.EnableMenu("1283", true);                // 제거
                oForm.EnableMenu("1284", false);               // 취소
                oForm.EnableMenu("1293", true);                // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY104_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY104_SetDocument
        /// </summary>
        public void PH_PY104_SetDocument(string oFromDocEntry01)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY104_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY104_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY104_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PH_PY104_FormItemEnabled()
        {
            int i = 0;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    // 기본사항 - 부서 (사업장에 따른 부서변경)
                    if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "");
                        oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    if (!string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx))
                    {
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                        sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "'";
                        sQry = sQry + " ORDER BY U_Code";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific,"");
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
                        oForm.Items.Item("TeamCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    // 담당 (사업장에 따른 담당변경)
                    if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("", "");
                        oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    if (!string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx))
                    {
                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                        sQry = sQry + " WHERE Code = '2' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "'";
                        sQry = sQry + " Order By U_Code";
                        dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific,"");
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
                        oForm.Items.Item("RspCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    tSeqAll = 0;
                    PH_PY104_DataLoad();

                    oForm.EnableMenu("1281", true);                    // 문서찾기
                    oForm.EnableMenu("1282", false);                   // 문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    // 부서
                    if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("TeamCode").Specific.ValidValues.Add("", "-");
                        oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    // 담당
                    if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                    {
                        for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                        {
                            oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                        }
                        oForm.Items.Item("RspCode").Specific.ValidValues.Add("", "-");
                        oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    tSeqAll = 0;
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    PH_PY104_DataLoad();

                    oForm.EnableMenu("1281", false);                   // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1281", true);                    // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY104_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">이벤트 </param>
        /// <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                    ////2
                //    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS://                    4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                    ////7
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:                    ////8
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:                    ////9
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:                    ////12
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                    ////16
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                    ////18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                    ////19
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:                    ////20
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:                    ////22
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:                    ////23
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:                    ////37
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT:                    ////38
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_Drag:                    ////39
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
            int i = 0;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn_Serch")
                    {
                        if (PH_PY104_DataValidCheck() == true)
                        {
                            PH_PY104_DataFind();
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn_Save")
                    {
                        if (PH_PY104_DataSave() == false)
                        {
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn_AllChk")
                    {
                        if (oGrid2.Rows.Count > 0)
                        {
                            oForm.Freeze(true);
                            for (i = 0; i <= oGrid2.Rows.Count - 1; i++)
                            {
                                oDS_PH_PY104_02.SetValue("ChkBox", i, "Y");
                            }
                            oForm.Freeze(false);
                        }
                    }
                    if (pVal.ItemUID == "Btn_Copy")
                    {
                        PH_PY104_DataCopy();
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
        /// Raise_EVENT_GOT_FOCUS
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.ItemUID)
                {
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
            }
        }

        /// <summary>
        /// Raise_EVENT_COMBO_SELECT
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            int i = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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

                        if (pVal.ItemUID == "CLTCOD")
                        {
                            // 기본사항 - 부서 (사업장에 따른 부서변경)
                            if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
                            sQry = sQry + " ORDER BY U_Code";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific,"");
                            oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
                            oForm.Items.Item("TeamCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("TeamCode").DisplayDesc = true;

                            PH_PY104_DataLoad();

                            // 담당 (사업장에 따른 담당변경)
                            if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }
                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry = sQry + " WHERE Code = '2' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
                            sQry = sQry + " Order By U_Code";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific,"");
                            oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
                            oForm.Items.Item("RspCode").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("RspCode").DisplayDesc = true;
                        }
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
        /// Raise_EVENT_CLICK
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
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
                    if (pVal.ActionSuccess == true)
                    {
                        if (pVal.ItemUID == "Grid1" & pVal.ColUID == "SLT")
                        {
                            Check_Seq(pVal.ColUID, pVal.Row);
                        }
                    }
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
        /// Raise_EVENT_MATRIX_LOAD
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    PH_PY104_FormItemEnabled();
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
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
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY104_FormItemEnabled();
                            break;

                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281":
                            ////문서찾기
                            PH_PY104_FormItemEnabled();
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY104_FormItemEnabled();
                            break;

                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY104_FormItemEnabled();
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            // 33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                             // 34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                          // 35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                          // 36
                            break;
                    }
                    ////BeforeAction = False
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            // 33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                             // 34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                          // 35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                          // 36
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY104_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            int j = 0;
            try
            {
                if (oGrid1.Rows.Count > 0)
                {
                    for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                    {
                        if (oGrid1.DataTable.GetValue("SLT", i) == "Y")
                        {
                            if (!string.IsNullOrEmpty(oGrid1.DataTable.GetValue("U_CSUCOD", i)))
                            {
                                j = j + 1;
                            }
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("수당, 공제 테이블에 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }

                if (j == 0)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("선택된 수당, 공제 데이터가 없습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_DataFind()
        {
            int i = 0;
            int iRow = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string TeamCode = string.Empty;
            string RspCode = string.Empty;
            string PAYTYP = string.Empty;
            string JIGCODF = string.Empty;
            string JIGCODT = string.Empty;
            string HOBONGF = string.Empty;
            string HOBONGT = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);

                // PH_PY104_TEMP 테이블 초기화
                sQry = "DELETE PH_PY104_TEMP";
                oRecordSet.DoQuery(sQry);
                // PH_PY104_TEMP2 테이블 초기화
                sQry = "DELETE PH_PY104_TEMP2";
                oRecordSet.DoQuery(sQry);

                // 그리드1 체크 데이터 PH_PY104_TEMP  저장
                if (oGrid1.Rows.Count > 0)
                {
                    for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                    {
                        if (oDS_PH_PY104_01.GetValue("SLT", i) == "Y")
                        {
                            sQry = "EXEC PH_PY104_Grid1 '" + oDS_PH_PY104_01.GetValue("GBN", i) + "','";
                            sQry = sQry + oDS_PH_PY104_01.GetValue("U_CSUCOD", i) + "','";
                            sQry = sQry + oDS_PH_PY104_01.GetValue("U_CSUNAM", i) + "','";
                            sQry = sQry + oDS_PH_PY104_01.GetValue("SEQ", i) + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                    }
                }

                CLTCOD   = oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.VALUE.ToString().Trim(); 
                RspCode  = oForm.Items.Item("RspCode").Specific.VALUE.ToString().Trim();
                PAYTYP   = oForm.Items.Item("PAYTYP").Specific.VALUE.ToString().Trim();
                JIGCODF  = (oForm.Items.Item("JIGCODF").Specific.VALUE.Trim() == "%" ? "00" : oForm.Items.Item("JIGCODF").Specific.VALUE.Trim());
                JIGCODT  = (oForm.Items.Item("JIGCODT").Specific.VALUE.Trim() == "%" ? "ZZ" : oForm.Items.Item("JIGCODT").Specific.VALUE.Trim());
                HOBONGF  = (string.IsNullOrEmpty(oForm.Items.Item("HOBONGF").Specific.VALUE.Trim()) ? "000" : oForm.Items.Item("HOBONGF").Specific.VALUE.Trim());
                HOBONGT  = (string.IsNullOrEmpty(oForm.Items.Item("HOBONGT").Specific.VALUE.Trim()) ? "ZZZ" : oForm.Items.Item("HOBONGT").Specific.VALUE.Trim());

                // 검색 조건 - 임시 테이블 저장 PH_PY104_TEMP2
                sQry = "Exec PH_PY104_Grid2 '" + CLTCOD + "','" + TeamCode + "',";
                sQry = sQry + "'" + RspCode + "','" + PAYTYP + "',";
                sQry = sQry + "'" + JIGCODF + "','" + JIGCODT + "',";
                sQry = sQry + "'" + HOBONGF + "','" + HOBONGT + "'";
                oRecordSet.DoQuery(sQry);

                // 그리드1 체크 데이터P PH_PY104_TEMP 불러옴
                sQry = "SELECT GUBUN, CSUCOD, CSUNAM FROM PH_PY104_TEMP ORDER BY SEQ";
                oRecordSet.DoQuery(sQry);

                // PH_PY104_TEMP2  데이터에 코드와 이름 붙여서 다시 셀렉트
                if (oRecordSet.RecordCount > 0)
                {
                    sQry = "SELECT '' AS ChkBox , T0.CODE, T0.Name ";
                    for (i = 1; i <= oRecordSet.RecordCount; i++)
                    {
                        if (Strings.Trim(oRecordSet.Fields.Item(0).Value) == "수당")
                        {
                            sQry = sQry + ", ISNULL((SELECT U_FILD03 FROM [@PH_PY001B] WHERE U_FILD02 ='" + oRecordSet.Fields.Item(2).Value + "'  AND Code = T0.Code ),0) AS N'" + oRecordSet.Fields.Item(2).Value + "'";

                        }
                        else if (Strings.Trim(oRecordSet.Fields.Item(0).Value) == "공제")
                        {
                            sQry = sQry + ", ISNULL((SELECT U_FILD03 FROM [@PH_PY001C] WHERE U_FILD02 ='" + oRecordSet.Fields.Item(2).Value + "'  AND Code = T0.Code ),0) AS N'" + oRecordSet.Fields.Item(2).Value + "'";
                        }
                        oRecordSet.MoveNext();
                    }
                    sQry = sQry + " FROM  PH_PY104_TEMP2 T0";
                    oDS_PH_PY104_02.ExecuteQuery(sQry);
                }
                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
                PH_PY104_TitleSetting_Grid2(iRow);
                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataFind_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY104_DataLoad
        /// </summary>
        /// <returns></returns>
        private object PH_PY104_DataLoad()
        {
            object functionReturnValue = string.Empty;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);

                sQry = "SELECT '수당' AS GBN, T0.U_CSUCOD, T0.U_CSUNAM, '' AS SLT, SPACE(5) AS SEQ";
                sQry = sQry + " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code";
                sQry = sQry + " WHERE U_FIXGBN = 'Y' AND T1.U_YM = (select Max(U_YM) AS U_YM from [@PH_PY102A])";
                sQry = sQry + " AND T1.U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.VALUE + "'";
                sQry = sQry + " Union All";
                sQry = sQry + " SELECT '공제' AS GBN, T0.U_CSUCOD, T0.U_CSUNAM, '' AS SLT, SPACE(5) AS SEQ";
                sQry = sQry + " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code";
                sQry = sQry + " WHERE U_FIXGBN = 'Y' AND T1.U_YM = (select Max(U_YM) AS U_YM from [@PH_PY103A])";
                sQry = sQry + " AND T1.U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.VALUE + "'";

                oDS_PH_PY104_01.ExecuteQuery(sQry);
                PH_PY104_TitleSetting_Grid1();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataLoad_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
                
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY104_DataCopy
        /// </summary>
        /// <returns></returns>
        private bool PH_PY104_DataCopy()
        {
            bool functionReturnValue = false;
            int i = 0;
            int j = 0;
            string ShiftDat = string.Empty;
            int First = 0;
            object[] FirstData = null;
            int TOTCNT = 0;
            int V_StatusCnt = 0;
            int oProValue = 0;
            int tRow = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            try
            {
                oForm.Freeze(true);

                functionReturnValue = false;

                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    ProgressBar01 = null;
                }

                First = 0;
                if (oGrid2.Rows.Count > 0)
                {
                    ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, false);

                    // 최대값 구하기
                    TOTCNT = oGrid2.Rows.Count;
                    V_StatusCnt = TOTCNT / 50;
                    oProValue = 1;
                    tRow = 1;

                    for (i = 0; i <= oGrid2.Rows.Count - 1; i++)
                    {
                        if (oDS_PH_PY104_02.GetValue("ChkBox", i) == "Y")
                        {
                            First = First + 1;
                            if (First == 1)
                            {
                                FirstData = new object[oGrid2.Columns.Count + 1];
                                for (j = 0; j <= oGrid2.Columns.Count - 1; j++)
                                {
                                    FirstData[j] = oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value;
                                }
                            }
                            else
                            {
                                for (j = 3; j <= oGrid2.Columns.Count - 1; j++)
                                {
                                    oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value = FirstData[j];
                                }
                            }
                        }

                        tRow = tRow + 1;
                        if ((TOTCNT > 50 & tRow == oProValue * V_StatusCnt) | TOTCNT <= 50)
                        {
                            ProgressBar01.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
                            oProValue = oProValue + 1;
                            ProgressBar01.Value = oProValue;
                        }
                    }
                }
                ProgressBar01.Stop();

                if (First == 0)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("선택된 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("선택된 필드의 전체 값 복사가 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    functionReturnValue = true;
                }
                ProgressBar01 = null;
               
            }
            catch (Exception ex)
            {
                if ((ProgressBar01 != null))
                {
                    ProgressBar01.Stop();
                    ProgressBar01 = null;
                }
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataCopy_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// DataSave
        /// </summary>
        /// <returns></returns>
        private bool PH_PY104_DataSave()
        {
            bool functionReturnValue = false;
            int i = 0;
            int j = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                functionReturnValue = false;
                if (oGrid2.Rows.Count > 0)
                {
                    for (i = 0; i <= oGrid2.Rows.Count - 1; i++)
                    {
                        if (oDS_PH_PY104_02.GetValue("ChkBox", i) == "Y")
                        {
                            for (j = 3; j <= oGrid2.Columns.Count - 1; j++)
                            {
                                sQry = "SELECT GUBUN FROM PH_PY104_TEMP WHERE CSUNAM = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
                                oRecordSet.DoQuery(sQry);

                                if (oRecordSet.RecordCount > 0)
                                {
                                    if (Strings.Trim(oRecordSet.Fields.Item(0).Value) == "수당")
                                    {
                                        sQry = "UPDATE [@PH_PY001B] SET U_FILD03 = '" + oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value + "'";
                                        sQry = sQry + " WHERE Code = '" + Strings.Trim(oDS_PH_PY104_02.GetValue("CODE", i)) + "'";
                                        sQry = sQry + " AND U_FILD02 = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
                                        oRecordSet.DoQuery(sQry);
                                    }
                                    else if (Strings.Trim(oRecordSet.Fields.Item(0).Value) == "공제")
                                    {
                                        sQry = "UPDATE [@PH_PY001C] SET U_FILD03 = '" + oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value + "'";
                                        sQry = sQry + " WHERE Code = '" + Strings.Trim(oDS_PH_PY104_02.GetValue("CODE", i)) + "'";
                                        sQry = sQry + " AND U_FILD02 = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
                                        oRecordSet.DoQuery(sQry);
                                    }
                                }
                            }
                            PSH_Globals.SBO_Application.SetStatusBarMessage("[" + Strings.Trim(oDS_PH_PY104_02.GetValue("CODE", i)) + "] 의 수당,공제 데이터가 갱신중입니다..", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        }
                    }
                    functionReturnValue = true;
                    PSH_Globals.SBO_Application.SetStatusBarMessage("수당,공제 데이터가 갱신 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_DataSave_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY104_TitleSetting_Grid1 그리드 타이블 변경
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_TitleSetting_Grid1()
        {
            int i = 0;
            string sQry = string.Empty;
            string[] COLNAM = new string[5];
            try
            {
                oForm.Freeze(true);

                COLNAM[0] = "구분";
                COLNAM[1] = "코드";
                COLNAM[2] = "코드명";
                COLNAM[3] = "선택";
                COLNAM[4] = "순서";

                for (i = 0; i <= Information.UBound(COLNAM); i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 & i <= 2)
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                    else if (i == 3)
                    {
                        oGrid1.Columns.Item(i).Editable = true;
                        oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                    }
                    else if (i == 4)
                    {
                        oGrid1.Columns.Item(i).Editable = true;
                    }

                }
                oGrid1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_TitleSetting_Grid1_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY104_TitleSetting_Grid2 그리드 타이블 변경
        /// </summary>
        /// <returns></returns>
        private void PH_PY104_TitleSetting_Grid2(int iRoW)
        {
            int i = 0;
            string sQry = string.Empty;
            string[] COLNAM = new string[3];
            try
            {
                oForm.Freeze(true);

                COLNAM[0] = "체크";
                COLNAM[1] = "사번";
                COLNAM[2] = "이름";


                for (i = 0; i <= oGrid2.Columns.Count - 1; i++)
                {
                    if (i == 0)
                    {
                        oGrid2.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                        oGrid2.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                        oGrid2.Columns.Item(i).Editable = true;
                    }
                    else if (i == 1 | i == 2)
                    {
                        oGrid2.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                        oGrid2.Columns.Item(i).Editable = false;
                    }
                    else
                    {
                        oGrid2.Columns.Item(i).Editable = true;
                        oGrid2.Columns.Item(i).RightJustified = true;
                    }

                }
                oGrid2.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY104_TitleSetting_Grid2_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Check_Seq 그리드 타이블 변경
        /// </summary>
        /// <returns></returns>
        private void Check_Seq(string ColUID, int Row)
        {
            int i = 0;
            int tSeq = 0;
            try
            {
                oForm.Freeze(true);

                if (oGrid1.Rows.Count > 0 & Row >= 0)
                {
                    if (tSeqAll < 0)
                        tSeqAll = 0;
                    if (oGrid1.DataTable.GetValue("SLT", Row) == "Y")
                    {
                        tSeqAll = tSeqAll + 1;
                        oGrid1.DataTable.SetValue("SEQ", Row, tSeqAll);
                    }
                    else if (oGrid1.DataTable.GetValue("SLT", Row) == "N")
                    {
                        tSeqAll = tSeqAll - 1;
                        if (string.IsNullOrEmpty(oGrid1.DataTable.GetValue("SEQ", Row)))
                        {
                            oGrid1.DataTable.SetValue("SEQ", Row, 0);
                        }
                        tSeq = oGrid1.DataTable.GetValue("SEQ", Row);
                        oGrid1.DataTable.SetValue("SEQ", Row, "");

                        for (i = 0; i <= oGrid1.Rows.Count - 1; i++)
                        {
                            if (oGrid1.DataTable.GetValue("SEQ", i) > tSeq & !string.IsNullOrEmpty(oGrid1.DataTable.GetValue("SEQ", i)))
                            {
                                oGrid1.DataTable.SetValue("SEQ", i, Convert.ToInt16(oGrid1.DataTable.GetValue("SEQ", i)) - 1);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("Check_Seq_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {

                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid2);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY104_01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY104_02);
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
//	internal class PH_PY104
//	{
//////********************************************************************************
//////  File           : PH_PY104.cls
//////  Module         : 급여관리 > 급여관리
//////  Desc           : 고정수당공제금액일괄등록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Grid oGrid1;
//		public SAPbouiCOM.Grid oGrid2;
//		public SAPbouiCOM.DataTable oDS_PH_PY104_01;
//		public SAPbouiCOM.DataTable oDS_PH_PY104_02;

//			////그리드1의 체크 순번
//		public int tSeqAll;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY104.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY104_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY104");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


//			oForm.Freeze(true);
//			PH_PY104_CreateItems();
//			PH_PY104_EnableMenus();
//			PH_PY104_SetDocument(oFromDocEntry01);

//			//    Call PH_PY104_FormResize

//			oForm.Update();
//			oForm.Freeze(false);

//			oForm.Visible = true;
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			return;
//			LoadForm_Error:

//			oForm.Update();
//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oXmlDoc 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oXmlDoc = null;
//			//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oForm = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Form_Load Error:" + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private bool PH_PY104_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;

//			SAPbouiCOM.CheckBox oCheck = null;
//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;
//			SAPbouiCOM.OptionBtn optBtn = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oGrid1 = oForm.Items.Item("Grid1").Specific;

//			oForm.DataSources.DataTables.Add("PH_PY104_01");
//			oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
//			oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
//			oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("이름", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
//			oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("선택", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
//			oForm.DataSources.DataTables.Item("PH_PY104_01").Columns.Add("순서", SAPbouiCOM.BoFieldsType.ft_Float);

//			oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY104_01");
//			oDS_PH_PY104_01 = oForm.DataSources.DataTables.Item("PH_PY104_01");

//			oGrid2 = oForm.Items.Item("Grid2").Specific;

//			oForm.DataSources.DataTables.Add("PH_PY104_02");
//			oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("체크", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
//			oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
//			oForm.DataSources.DataTables.Item("PH_PY104_02").Columns.Add("이름", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

//			oGrid2.DataTable = oForm.DataSources.DataTables.Item("PH_PY104_02");
//			oDS_PH_PY104_02 = oForm.DataSources.DataTables.Item("PH_PY104_02");


//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			////사업장
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			oCombo.DataBind.SetBound(true, "", "CLTCOD");
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			////부서
//			oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo = oForm.Items.Item("TeamCode").Specific;
//			oCombo.DataBind.SetBound(true, "", "TeamCode");
//			oForm.Items.Item("TeamCode").DisplayDesc = true;
//			//
//			//// 담당
//			oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo = oForm.Items.Item("RspCode").Specific;
//			oCombo.DataBind.SetBound(true, "", "RspCode");
//			oForm.Items.Item("RspCode").DisplayDesc = true;

//			//// 급여형태
//			oForm.DataSources.UserDataSources.Add("PAYTYP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo = oForm.Items.Item("PAYTYP").Specific;
//			oCombo.DataBind.SetBound(true, "", "PAYTYP");

//			sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P132' AND U_UseYN= 'Y' ";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oCombo.ValidValues.Add("%", "전체");
//			oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oForm.Items.Item("PAYTYP").DisplayDesc = true;

//			//// 직급형태From
//			oCombo = oForm.Items.Item("JIGCODF").Specific;
//			sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code ";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oCombo.ValidValues.Add("%", "전체");
//			oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oForm.Items.Item("JIGCODF").DisplayDesc = true;

//			//// 직급형태To
//			oCombo = oForm.Items.Item("JIGCODT").Specific;
//			sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P129' ORDER BY U_Code ";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oCombo.ValidValues.Add("%", "전체");
//			oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//			oForm.Items.Item("JIGCODT").DisplayDesc = true;

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			optBtn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY104_CreateItems_Error:

//			//UPGRADE_NOTE: oCheck 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCheck = null;
//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: optBtn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			optBtn = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY104_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", true);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", true);
//			////행삭제

//			return;
//			PH_PY104_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY104_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY104_FormItemEnabled();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY104_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY104_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY104_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;
//			int i = 0;
//			string sQry = null;
//			 // ERROR: Not supported in C#: OnErrorStatement



//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				////기본사항 - 부서 (사업장에 따른 부서변경)
//				oCombo = oForm.Items.Item("TeamCode").Specific;

//				if (oCombo.ValidValues.Count > 0) {
//					for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//						oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//					}
//					oCombo.ValidValues.Add("", "");
//					oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}

//				if (!string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx)) {
//					sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//					//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//					sQry = sQry + " ORDER BY U_Code";
//					MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//					oCombo.ValidValues.Add("%", "전체");
//					oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//				}

//				////담당 (사업장에 따른 담당변경)

//				oCombo = oForm.Items.Item("RspCode").Specific;

//				if (oCombo.ValidValues.Count > 0) {
//					for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//						oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//					}
//					oCombo.ValidValues.Add("", "");
//					oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}

//				if (!string.IsNullOrEmpty(oForm.DataSources.UserDataSources.Item("CLTCOD").ValueEx)) {
//					sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//					//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					sQry = sQry + " WHERE Code = '2' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//					sQry = sQry + " Order By U_Code";
//					MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//					oCombo.ValidValues.Add("%", "전체");
//					oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//				}

//				tSeqAll = 0;
//				PH_PY104_DataLoad();

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				////부서
//				oCombo = oForm.Items.Item("TeamCode").Specific;
//				if (oCombo.ValidValues.Count > 0) {
//					for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//						oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//					}
//					oCombo.ValidValues.Add("", "-");
//					oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}

//				////담당
//				oCombo = oForm.Items.Item("RspCode").Specific;
//				if (oCombo.ValidValues.Count > 0) {
//					for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//						oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//					}
//					oCombo.ValidValues.Add("", "-");
//					oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//				}


//				tSeqAll = 0;
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//				PH_PY104_DataLoad();

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가
//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);


//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY104_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1
//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "Btn_Serch") {
//							if (PH_PY104_DataValidCheck() == true) {
//								PH_PY104_DataFind();
//							} else {
//								BubbleEvent = false;
//							}
//						}
//						if (pval.ItemUID == "Btn_Save") {
//							if (PH_PY104_DataSave() == false) {
//								BubbleEvent = false;
//							}
//						}
//						if (pval.ItemUID == "Btn_AllChk") {
//							if (oGrid2.Rows.Count > 0) {
//								oForm.Freeze(true);
//								for (i = 0; i <= oGrid2.Rows.Count - 1; i++) {
//									oDS_PH_PY104_02.SetValue("ChkBox", i, "Y");
//								}
//								oForm.Freeze(false);
//							}
//						}
//						if (pval.ItemUID == "Btn_Copy") {
//							PH_PY104_DataCopy();
//						}

//					} else if (pval.BeforeAction == false) {

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					if (pval.BeforeAction == false) {
//						//                If pval.ItemUID = "HOBONGF" Or pval.ItemUID = "HOBONGT" Then
//						//                    If pval.CharPressed = "9" Then
//						//                        If oForm.Items(pval.ItemUID).Specific.Value = "" Then
//						//                            oForm.Items(pval.ItemUID).CLICK ct_Regular
//						//                            Sbo_Application.ActivateMenuItem ("7425")
//						//                        End If
//						//                    End If
//						//                End If
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Grid1":
//							if (pval.Row > 0) {
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = pval.ColUID;
//								oLastColRow = pval.Row;
//							}
//							break;
//						default:
//							oLastItemUID = pval.ItemUID;
//							oLastColUID = "";
//							oLastColRow = 0;
//							break;
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//					////4
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//					////5
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {
//							if (pval.ItemUID == "CLTCOD") {
//								////기본사항 - 부서 (사업장에 따른 부서변경)
//								oCombo = oForm.Items.Item("TeamCode").Specific;

//								if (oCombo.ValidValues.Count > 0) {
//									for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//										oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//									}
//								}

//								sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//								//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//								sQry = sQry + " ORDER BY U_Code";
//								MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//								oCombo.ValidValues.Add("%", "전체");
//								oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//								oForm.Items.Item("TeamCode").DisplayDesc = true;

//								PH_PY104_DataLoad();

//								////담당 (사업장에 따른 담당변경)

//								oCombo = oForm.Items.Item("RspCode").Specific;

//								if (oCombo.ValidValues.Count > 0) {
//									for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//										oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//									}
//								}

//								sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//								//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sQry = sQry + " WHERE Code = '2' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//								sQry = sQry + " Order By U_Code";
//								MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//								oCombo.ValidValues.Add("%", "전체");
//								oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//								oForm.Items.Item("RspCode").DisplayDesc = true;

//							}
//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6

//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Grid1":
//								if (pval.Row > 0) {
//									oLastItemUID = pval.ItemUID;
//									oLastColUID = pval.ColUID;
//									oLastColRow = pval.Row;
//								}
//								break;
//							default:
//								oLastItemUID = pval.ItemUID;
//								oLastColUID = "";
//								oLastColRow = 0;
//								break;
//						}
//					} else if (pval.BeforeAction == false) {
//						if (pval.ActionSuccess == true) {
//							if (pval.ItemUID == "Grid1" & pval.ColUID == "SLT") {
//								Check_Seq(ref (pval.ColUID), ref (pval.Row));
//							}
//						}
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "Grid1") {
//							oForm.Freeze((false));
//							BubbleEvent = false;
//						}
//						return;
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//					////8
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//					////9
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//					////10
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemChanged == true) {

//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						PH_PY104_FormItemEnabled();
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//					////12
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//					////16
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//					////17
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY104_01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY104_01 = null;
//						//UPGRADE_NOTE: oGrid1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oGrid1 = null;

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//					////18
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//					////19
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//					////20
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//					////21
//					if (pval.BeforeAction == true) {

//					} else if (pval.BeforeAction == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//					////22
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//					////23
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//					////27
//					if (pval.BeforeAction == true) {
//					} else if (pval.Before_Action == false) {
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//					////37
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//					////38
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_Drag:
//					////39
//					break;

//			}

//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			return;
//			Raise_FormItemEvent_Error:
//			///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//			oForm.Freeze((false));
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			oForm.Freeze(true);

//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						if (MDC_Globals.Sbo_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
//							BubbleEvent = false;
//							return;
//						}
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					case "1293":
//						break;
//					case "1281":
//						break;
//					case "1282":
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						break;
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY104_FormItemEnabled();
//						break;

//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY104_FormItemEnabled();
//						break;
//					case "1282":
//						////문서추가
//						PH_PY104_FormItemEnabled();
//						break;

//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY104_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((BusinessObjectInfo.BeforeAction == true)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						////34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						////35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						////36
//						break;
//				}
//			} else if ((BusinessObjectInfo.BeforeAction == false)) {
//				switch (BusinessObjectInfo.EventType) {
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//						////33
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//						////34
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//						////35
//						break;
//					case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//						////36
//						break;
//				}
//			}
//			return;
//			Raise_FormDataEvent_Error:


//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pval.BeforeAction == true) {
//			} else if (pval.BeforeAction == false) {
//			}
//			switch (pval.ItemUID) {
//				case "Grid1":
//					if (pval.Row > 0) {
//						oLastItemUID = pval.ItemUID;
//						oLastColUID = pval.ColUID;
//						oLastColRow = pval.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pval.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY104_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY104'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY104_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY104_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			j = 0;
//			if (oGrid1.Rows.Count > 0) {
//				for (i = 0; i <= oGrid1.Rows.Count - 1; i++) {
//					//UPGRADE_WARNING: oGrid1.DataTable.GetValue(SLT, i) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (oGrid1.DataTable.GetValue("SLT", i) == "Y") {
//						//UPGRADE_WARNING: oGrid1.DataTable.GetValue(U_CSUCOD, i) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (!string.IsNullOrEmpty(oGrid1.DataTable.GetValue("U_CSUCOD", i))) {
//							j = j + 1;
//						}
//					}
//				}
//			} else {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("수당, 공제 테이블에 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//			if (j == 0) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("선택된 수당, 공제 데이터가 없습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE)) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			functionReturnValue = true;
//			return functionReturnValue;



//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY104_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		public bool PH_PY104_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY104A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY104A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY104_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY104_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY104_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		private void PH_PY104_DataFind()
//		{
//			int i = 0;
//			int iRow = 0;
//			string sQry = null;
//			string sQry2 = null;

//			string CLTCOD = null;
//			string TeamCode = null;
//			string RspCode = null;
//			string PAYTYP = null;
//			string JIGCODF = null;
//			string JIGCODT = null;
//			string HOBONGF = null;
//			string HOBONGT = null;

//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbobsCOM.Recordset pRecordset = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			pRecordset = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			oForm.Freeze((true));

//			/// PH_PY104_TEMP 테이블 초기화
//			sQry = "DELETE PH_PY104_TEMP";
//			oRecordSet.DoQuery(sQry);
//			/// PH_PY104_TEMP2 테이블 초기화
//			sQry = "DELETE PH_PY104_TEMP2";
//			oRecordSet.DoQuery(sQry);

//			/// 그리드1 체크 데이터 PH_PY104_TEMP  저장
//			if (oGrid1.Rows.Count > 0) {
//				for (i = 0; i <= oGrid1.Rows.Count - 1; i++) {
//					//UPGRADE_WARNING: oDS_PH_PY104_01.GetValue(SLT, i) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (oDS_PH_PY104_01.GetValue("SLT", i) == "Y") {
//						//UPGRADE_WARNING: oDS_PH_PY104_01.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sQry = "EXEC PH_PY104_Grid1 '" + oDS_PH_PY104_01.GetValue("GBN", i) + "','";
//						//UPGRADE_WARNING: oDS_PH_PY104_01.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sQry = sQry + oDS_PH_PY104_01.GetValue("U_CSUCOD", i) + "','";
//						//UPGRADE_WARNING: oDS_PH_PY104_01.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sQry = sQry + oDS_PH_PY104_01.GetValue("U_CSUNAM", i) + "','";
//						//UPGRADE_WARNING: oDS_PH_PY104_01.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						sQry = sQry + oDS_PH_PY104_01.GetValue("SEQ", i) + "'";
//						oRecordSet.DoQuery(sQry);
//					}
//				}
//			}

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TeamCode = oForm.Items.Item("TeamCode").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			RspCode = oForm.Items.Item("RspCode").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			PAYTYP = oForm.Items.Item("PAYTYP").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items(JIGCODF).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JIGCODF = (oForm.Items.Item("JIGCODF").Specific.VALUE == "%" ? "00" : oForm.Items.Item("JIGCODF").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items(JIGCODT).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JIGCODT = (oForm.Items.Item("JIGCODT").Specific.VALUE == "%" ? "ZZ" : oForm.Items.Item("JIGCODT").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items(HOBONGF).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			HOBONGF = (string.IsNullOrEmpty(oForm.Items.Item("HOBONGF").Specific.VALUE) ? "000" : oForm.Items.Item("HOBONGF").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items(HOBONGT).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			HOBONGT = (string.IsNullOrEmpty(oForm.Items.Item("HOBONGT").Specific.VALUE) ? "ZZZ" : oForm.Items.Item("HOBONGT").Specific.VALUE);

//			/// 검색 조건 - 임시 테이블 저장 PH_PY104_TEMP2
//			sQry = "Exec PH_PY104_Grid2 '" + CLTCOD + "','" + TeamCode + "',";
//			sQry = sQry + "'" + RspCode + "','" + PAYTYP + "',";
//			sQry = sQry + "'" + JIGCODF + "','" + JIGCODT + "',";
//			sQry = sQry + "'" + HOBONGF + "','" + HOBONGT + "'";
//			oRecordSet.DoQuery(sQry);

//			/// 그리드1 체크 데이터P PH_PY104_TEMP 불러옴
//			sQry = "SELECT GUBUN, CSUCOD, CSUNAM FROM PH_PY104_TEMP ORDER BY SEQ";
//			oRecordSet.DoQuery(sQry);

//			/// PH_PY104_TEMP2  데이터에 코드와 이름 붙여서 다시 셀렉트
//			if (oRecordSet.RecordCount > 0) {
//				sQry = "SELECT '' AS ChkBox , T0.CODE, T0.Name ";
//				for (i = 1; i <= oRecordSet.RecordCount; i++) {
//					if (Strings.Trim(oRecordSet.Fields.Item(0).Value) == "수당") {
//						sQry = sQry + ", ISNULL((SELECT U_FILD03 FROM [@PH_PY001B] WHERE U_FILD02 ='" + oRecordSet.Fields.Item(2).Value + "'  AND Code = T0.Code ),0) AS N'" + Strings.Trim(oRecordSet.Fields.Item(2).Value) + "'";

//					} else if (Strings.Trim(oRecordSet.Fields.Item(0).Value) == "공제") {
//						sQry = sQry + ", ISNULL((SELECT U_FILD03 FROM [@PH_PY001C] WHERE U_FILD02 ='" + oRecordSet.Fields.Item(2).Value + "'  AND Code = T0.Code ),0) AS N'" + Strings.Trim(oRecordSet.Fields.Item(2).Value) + "'";
//					}
//					oRecordSet.MoveNext();
//				}
//				sQry = sQry + " FROM  PH_PY104_TEMP2 T0";
//				oDS_PH_PY104_02.ExecuteQuery(sQry);
//			}


//			iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
//			PH_PY104_TitleSetting_Grid2(ref iRow);

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			//UPGRADE_NOTE: pRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			pRecordset = null;
//			oForm.Freeze((false));
//			return;
//			PH_PY104_DataFind_Error:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			//UPGRADE_NOTE: pRecordset 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			pRecordset = null;
//			oForm.Freeze((false));
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_DataFind_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//		private object PH_PY104_DataLoad()
//		{
//			object functionReturnValue = null;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze((true));

//			sQry = "SELECT '수당' AS GBN, T0.U_CSUCOD, T0.U_CSUNAM, '' AS SLT, SPACE(5) AS SEQ";
//			sQry = sQry + " FROM [@PH_PY102B] T0 INNER JOIN [@PH_PY102A] T1 ON T0.Code = T1.Code";
//			sQry = sQry + " WHERE U_FIXGBN = 'Y' AND T1.U_YM = (select Max(U_YM) AS U_YM from [@PH_PY102A])";
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " AND T1.U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.VALUE + "'";
//			sQry = sQry + " Union All";
//			sQry = sQry + " SELECT '공제' AS GBN, T0.U_CSUCOD, T0.U_CSUNAM, '' AS SLT, SPACE(5) AS SEQ";
//			sQry = sQry + " FROM [@PH_PY103B] T0 INNER JOIN [@PH_PY103A] T1 ON T0.Code = T1.Code";
//			sQry = sQry + " WHERE U_FIXGBN = 'Y' AND T1.U_YM = (select Max(U_YM) AS U_YM from [@PH_PY103A])";
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = sQry + " AND T1.U_CLTCOD = '" + oForm.Items.Item("CLTCOD").Specific.VALUE + "'";

//			oDS_PH_PY104_01.ExecuteQuery(sQry);

//			PH_PY104_TitleSetting_Grid1();


//			oForm.Freeze((false));
//			return functionReturnValue;
//			PH_PY104_DataLoad_ERROR:

//			oForm.Freeze((false));
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_DataLoad_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private bool PH_PY104_DataCopy()
//		{
//			bool functionReturnValue = false;
//			int i = 0;
//			int j = 0;
//			string ShiftDat = null;
//			short First = 0;
//			object[] FirstData = null;
//			int TOTCNT = 0;
//			int V_StatusCnt = 0;
//			int oProValue = 0;
//			int tRow = 0;
//			////progbar

//			 // ERROR: Not supported in C#: OnErrorStatement

//			oForm.Freeze((true));

//			functionReturnValue = false;

//			if ((MDC_Globals.oProgBar != null)) {
//				MDC_Globals.oProgBar.Stop();
//				//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				MDC_Globals.oProgBar = null;
//			}

//			First = 0;
//			if (oGrid2.Rows.Count > 0) {
//				//프로그레스 바    ///////////////////////////////////////
//				MDC_Globals.oProgBar = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, false);

//				//최대값 구하기
//				TOTCNT = oGrid2.Rows.Count;

//				V_StatusCnt = System.Math.Round(TOTCNT / 50, 0);
//				oProValue = 1;
//				tRow = 1;
//				///////////////////////////////////////////////////////

//				for (i = 0; i <= oGrid2.Rows.Count - 1; i++) {
//					//UPGRADE_WARNING: oDS_PH_PY104_02.GetValue(ChkBox, i) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (oDS_PH_PY104_02.GetValue("ChkBox", i) == "Y") {
//						First = First + 1;
//						if (First == 1) {
//							FirstData = new object[oGrid2.Columns.Count + 1];
//							for (j = 0; j <= oGrid2.Columns.Count - 1; j++) {
//								//UPGRADE_WARNING: oDS_PH_PY104_02.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//UPGRADE_WARNING: FirstData(j) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								FirstData[j] = oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value;
//							}
//						} else {
//							for (j = 3; j <= oGrid2.Columns.Count - 1; j++) {
//								//UPGRADE_WARNING: FirstData(j) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value = FirstData[j];
//							}
//						}
//					}

//					tRow = tRow + 1;
//					if ((TOTCNT > 50 & tRow == oProValue * V_StatusCnt) | TOTCNT <= 50) {
//						MDC_Globals.oProgBar.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
//						oProValue = oProValue + 1;
//						MDC_Globals.oProgBar.Value = oProValue;
//					}
//				}
//			}
//			MDC_Globals.oProgBar.Stop();

//			if (First == 0) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("선택된 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//			} else {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("선택된 필드의 전체 값 복사가 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//				functionReturnValue = true;
//			}

//			//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			MDC_Globals.oProgBar = null;

//			oForm.Freeze((false));
//			return functionReturnValue;
//			PH_PY104_DataCopy_Error:

//			if ((MDC_Globals.oProgBar != null)) {
//				MDC_Globals.oProgBar.Stop();
//				//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				MDC_Globals.oProgBar = null;
//			}
//			//UPGRADE_NOTE: oProgBar 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			MDC_Globals.oProgBar = null;
//			oForm.Freeze((false));
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_DataCopy_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private bool PH_PY104_DataSave()
//		{
//			bool functionReturnValue = false;
//			int i = 0;
//			int j = 0;
//			int SuDangLineId = 0;
//			int GongjeLineId = 0;
//			string sQry = null;
//			string sQry2 = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			 // ERROR: Not supported in C#: OnErrorStatement


//			functionReturnValue = false;

//			if (oGrid2.Rows.Count > 0) {
//				for (i = 0; i <= oGrid2.Rows.Count - 1; i++) {
//					//UPGRADE_WARNING: oDS_PH_PY104_02.GetValue(ChkBox, i) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (oDS_PH_PY104_02.GetValue("ChkBox", i) == "Y") {
//						for (j = 3; j <= oGrid2.Columns.Count - 1; j++) {
//							//                    oDS_PH_PY104_02.Columns(j).Cells(i).Value
//							//                    Trim (oDS_PH_PY104_02.GetValue("CODE", i))
//							sQry = "SELECT GUBUN FROM PH_PY104_TEMP WHERE CSUNAM = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
//							oRecordSet.DoQuery(sQry);

//							if (oRecordSet.RecordCount > 0) {
//								if (Strings.Trim(oRecordSet.Fields.Item(0).Value) == "수당") {
//									sQry = "UPDATE [@PH_PY001B] SET U_FILD03 = '" + oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value + "'";
//									//UPGRADE_WARNING: oDS_PH_PY104_02.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									sQry = sQry + " WHERE Code = '" + Strings.Trim(oDS_PH_PY104_02.GetValue("CODE", i)) + "'";
//									sQry = sQry + " AND U_FILD02 = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
//									oRecordSet.DoQuery(sQry);
//								} else if (Strings.Trim(oRecordSet.Fields.Item(0).Value) == "공제") {
//									sQry = "UPDATE [@PH_PY001C] SET U_FILD03 = '" + oDS_PH_PY104_02.Columns.Item(j).Cells.Item(i).Value + "'";
//									//UPGRADE_WARNING: oDS_PH_PY104_02.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									sQry = sQry + " WHERE Code = '" + Strings.Trim(oDS_PH_PY104_02.GetValue("CODE", i)) + "'";
//									sQry = sQry + " AND U_FILD02 = '" + oDS_PH_PY104_02.Columns.Item(j).Name + "'";
//									oRecordSet.DoQuery(sQry);
//								}
//							}
//						}
//						//UPGRADE_WARNING: oDS_PH_PY104_02.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("[" + Strings.Trim(oDS_PH_PY104_02.GetValue("CODE", i)) + "] 의 수당,공제 데이터가 갱신중입니다..", SAPbouiCOM.BoMessageTime.bmt_Short, false);
//					}
//				}
//				functionReturnValue = true;
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("수당,공제 데이터가 갱신 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
//			} else {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("데이터가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY104_DataSave_Error:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_DataSave_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		private bool PH_PY104_OHEM_DI_UPDATE()
//		{
//			bool functionReturnValue = false;
//			int i = 0;
//			int j = 0;
//			int IErrCode = 0;
//			string sErrMsg = null;
//			string sQry = null;
//			string sItemName = null;
//			string EmpID = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			SAPbobsCOM.EmployeesInfo oOHEM = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			functionReturnValue = false;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			oOHEM = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oEmployeesInfo);

//			sQry = "SELECT U_Empid FROM [@PH_PY001A] WHERE Code = '" + Strings.Trim(oDS_PH_PY104_01.Columns.Item("Code").Cells.Item(i).Value) + "'";

//			oRecordSet.DoQuery(sQry);

//			//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			EmpID = oRecordSet.Fields.Item(0).Value;

//			MDC_Globals.Sbo_Application.SetStatusBarMessage(Strings.Trim(oDS_PH_PY104_01.Columns.Item("Code").Cells.Item(i).Value) + "의 근무조가 갱신중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);

//			if (string.IsNullOrEmpty(EmpID)) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("empID 로드에 실패하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
//				goto Error_Handle;
//			}
//			if (oOHEM.GetByKey(Convert.ToInt32(EmpID)) == true) {

//				var _with1 = oOHEM;
//				_with1.GetByKey(Convert.ToInt32(EmpID));

//				//UPGRADE_WARNING: oDS_PH_PY104_01.Columns.Item().Cells().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				_with1.UserFields.Fields.Item("U_GNMUJO").VALUE = oDS_PH_PY104_01.Columns.Item("GNMUJO").Cells.Item(i).Value;
//				////근무조

//				if (0 != oOHEM.Update()) {
//					MDC_Globals.oCompany.GetLastError(out out IErrCode, out out sErrMsg);
//					MDC_Globals.Sbo_Application.SetStatusBarMessage(sErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//				} else {
//					functionReturnValue = true;
//					MDC_Globals.Sbo_Application.SetStatusBarMessage(Strings.Trim(oDS_PH_PY104_01.Columns.Item("Code").Cells.Item(i).Value) + "근무조가 갱신되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);

//				}
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			//UPGRADE_NOTE: oOHEM 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oOHEM = null;
//			return functionReturnValue;
//			Error_Handle:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			//UPGRADE_NOTE: oOHEM 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oOHEM = null;
//			return functionReturnValue;

//		}

//		private void PH_PY104_TitleSetting_Grid1()
//		{
//			int i = 0;
//			int j = 0;
//			string sQry = null;

//			string[] COLNAM = new string[5];

//			SAPbouiCOM.EditTextColumn oColumn = null;
//			SAPbouiCOM.ComboBoxColumn oComboCol = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);


//			COLNAM[0] = "구분";
//			COLNAM[1] = "코드";
//			COLNAM[2] = "코드명";
//			COLNAM[3] = "선택";
//			COLNAM[4] = "순서";

//			for (i = 0; i <= Information.UBound(COLNAM); i++) {
//				oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
//				if (i >= 0 & i <= 2) {
//					oGrid1.Columns.Item(i).Editable = false;
//				} else if (i == 3) {
//					oGrid1.Columns.Item(i).Editable = true;
//					oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
//				} else if (i == 4) {
//					oGrid1.Columns.Item(i).Editable = true;
//				}

//			}

//			oGrid1.AutoResizeColumns();

//			oForm.Freeze(false);

//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;

//			return;
//			Error_Message:

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_TitleSetting_Grid1 Error : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}



//		private void PH_PY104_TitleSetting_Grid2(ref int iRow)
//		{
//			int i = 0;
//			int j = 0;
//			string sQry = null;

//			string[] COLNAM = new string[6];

//			SAPbouiCOM.EditTextColumn oColumn = null;
//			SAPbouiCOM.ComboBoxColumn oComboCol = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm.Freeze(true);

//			COLNAM[0] = "체크";
//			COLNAM[1] = "사번";
//			COLNAM[2] = "이름";


//			for (i = 0; i <= oGrid2.Columns.Count - 1; i++) {
//				if (i == 0) {
//					oGrid2.Columns.Item(i).TitleObject.Caption = COLNAM[i];
//					oGrid2.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
//					oGrid2.Columns.Item(i).Editable = true;
//				} else if (i == 1 | i == 2) {
//					oGrid2.Columns.Item(i).TitleObject.Caption = COLNAM[i];
//					oGrid2.Columns.Item(i).Editable = false;
//				} else {
//					oGrid2.Columns.Item(i).Editable = true;
//					oGrid2.Columns.Item(i).RightJustified = true;
//				}

//			}

//			oGrid2.AutoResizeColumns();

//			oForm.Freeze(false);

//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;

//			return;
//			Error_Message:

//			oForm.Freeze(false);
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY104_TitleSetting_Grid2 Error : " + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void Check_Seq(ref string ColUID, ref int Row)
//		{
//			int i = 0;
//			int tSeq = 0;

//			oForm.Freeze(true);
//			if (oGrid1.Rows.Count > 0 & Row >= 0) {
//				if (tSeqAll < 0)
//					tSeqAll = 0;
//				//UPGRADE_WARNING: oGrid1.DataTable.GetValue(SLT, Row) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (oGrid1.DataTable.GetValue("SLT", Row) == "Y") {
//					tSeqAll = tSeqAll + 1;
//					oGrid1.DataTable.SetValue("SEQ", Row, tSeqAll);
//					//UPGRADE_WARNING: oGrid1.DataTable.GetValue(SLT, Row) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				} else if (oGrid1.DataTable.GetValue("SLT", Row) == "N") {
//					tSeqAll = tSeqAll - 1;

//					//UPGRADE_WARNING: oGrid1.DataTable.GetValue(SEQ, Row) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oGrid1.DataTable.GetValue("SEQ", Row))) {
//						oGrid1.DataTable.SetValue("SEQ", Row, 0);
//					}
//					//UPGRADE_WARNING: oGrid1.DataTable.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					tSeq = oGrid1.DataTable.GetValue("SEQ", Row);
//					oGrid1.DataTable.SetValue("SEQ", Row, "");

//					for (i = 0; i <= oGrid1.Rows.Count - 1; i++) {
//						//UPGRADE_WARNING: oGrid1.DataTable.GetValue(SEQ, i) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (oGrid1.DataTable.GetValue("SEQ", i) > tSeq & !string.IsNullOrEmpty(oGrid1.DataTable.GetValue("SEQ", i))) {
//							//UPGRADE_WARNING: oGrid1.DataTable.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							oGrid1.DataTable.SetValue("SEQ", i, Convert.ToInt16(oGrid1.DataTable.GetValue("SEQ", i)) - 1);
//						}
//					}
//				}
//			}
//			oForm.Freeze(false);
//		}
//	}
//}
