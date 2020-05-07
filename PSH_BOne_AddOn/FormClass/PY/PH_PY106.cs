
using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 수당계산식설정
    /// </summary>
    internal class PH_PY106 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat01;
        public SAPbouiCOM.Matrix oMat02;
        public SAPbouiCOM.Matrix oMat03;

        private SAPbouiCOM.DBDataSource oDS_PH_PY106A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY106B;
        private SAPbouiCOM.DBDataSource oDS_PH_PY106C;
        private SAPbouiCOM.DBDataSource oDS_PH_PY106D;

        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY106.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY106_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY106");

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
                oForm.Items.Item("FLD01").Specific.Select();
                PH_PY106_CreateItems();
                PH_PY106_EnableMenus();
                PH_PY106_SetDocument(oFromDocEntry01);
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
        private void PH_PY106_CreateItems()
        {
            string sQry = string.Empty;
            int i = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                //    '//매트릭스--------------------------------------------------------------------------------------
                oDS_PH_PY106A = oForm.DataSources.DBDataSources.Item("@PH_PY106A");                ////헤더
                oDS_PH_PY106B = oForm.DataSources.DBDataSources.Item("@PH_PY106B");                ////라인
                oDS_PH_PY106C = oForm.DataSources.DBDataSources.Item("@PH_PY106C");                ////라인
                oDS_PH_PY106D = oForm.DataSources.DBDataSources.Item("@PH_PY106D");                ////라인

                oForm.DataSources.UserDataSources.Add("DISSIL", SAPbouiCOM.BoDataType.dt_LONG_TEXT);

                ////공식
                oForm.Items.Item("DISSIL").Specific.DataBind.SetBound(true, "", "DISSIL");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat03 = oForm.Items.Item("Mat03").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                ////----------------------------------------------------------------------------------------------
                //// 헤더 설정
                ////----------------------------------------------------------------------------------------------
                //// 귀속년월
                oForm.Items.Item("YM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");

                ////사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //// 1.급여형태-계산식
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P132' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("PAYTYP").Specific,"");
                oForm.Items.Item("PAYTYP").DisplayDesc = true;

                //// 1.근속년수 계산기준
                oForm.Items.Item("GNSGBN").Specific.ValidValues.Add("1", "그룹입사일");
                oForm.Items.Item("GNSGBN").Specific.ValidValues.Add("2", "입사  일자");
                if (oForm.Items.Item("GNSGBN").Specific.ValidValues.Count >= 1)
                {
                    oForm.Items.Item("GNSGBN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                //// 2.상여 계산단위 (
                oForm.Items.Item("BNSLEN").Specific.ValidValues.Add("1", "  원");
                oForm.Items.Item("BNSLEN").Specific.ValidValues.Add("10", "십원");
                oForm.Items.Item("BNSLEN").Specific.ValidValues.Add("100", "백원");
                oForm.Items.Item("BNSLEN").Specific.ValidValues.Add("1000", "천원");
                if (oForm.Items.Item("BNSLEN").Specific.ValidValues.Count >= 1)
                {
                    oForm.Items.Item("BNSLEN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                //// 3.상여 끝전처리
                oForm.Items.Item("BNSRND").Specific.ValidValues.Add("R", "반올림");
                oForm.Items.Item("BNSRND").Specific.ValidValues.Add("F", "절사");
                oForm.Items.Item("BNSRND").Specific.ValidValues.Add("C", "절상");
                if (oForm.Items.Item("BNSRND").Specific.ValidValues.Count >= 1)
                {
                    oForm.Items.Item("BNSRND").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY106_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                ////제거
                oForm.EnableMenu("1284", false);               ////취소
                oForm.EnableMenu("1293", true);                ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY106_SetDocument(string oFromDocEntry01)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY106_FormItemEnabled();
                    PH_PY106_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY106_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY106_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY106_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("PAYTYP").Enabled = true;
                    oForm.Items.Item("GNSGBN").Enabled = true;
                    oForm.Items.Item("BNSLEN").Enabled = true;
                    oForm.Items.Item("BNSRND").Enabled = true;

                    PH_PY106_Display_CsuItem();                    ////Mat02 초기값 가져옴

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    /// 귀속년월
                    //UPGRADE_WARNING: oForm.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    oForm.Items.Item("YM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");

                    //// 1.근속년수 계산기준
                    oForm.Items.Item("GNSGBN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);

                    //// 2.상여 계산단위 
                    oForm.Items.Item("BNSLEN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                    //// 3.상여 끝전처리
                    oForm.Items.Item("BNSRND").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                    oForm.EnableMenu("1293", true);                    ////행삭제
                    oForm.EnableMenu("1283", true);                    ////제거

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("PAYTYP").Enabled = true;
                    oForm.Items.Item("GNSGBN").Enabled = true;
                    oForm.Items.Item("BNSLEN").Enabled = true;
                    oForm.Items.Item("BNSRND").Enabled = true;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1293", true);                    ////행삭제
                    oForm.EnableMenu("1283", true);                    ////제거
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("PAYTYP").Enabled = false;
                    oForm.Items.Item("GNSGBN").Enabled = false;
                    oForm.Items.Item("BNSLEN").Enabled = false;
                    oForm.Items.Item("BNSRND").Enabled = false;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1293", true);                    ////행삭제
                    oForm.EnableMenu("1283", true);                    ////제거
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY106_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (MatrixSpaceLineDel() == false)
                                {
                                    BubbleEvent = false;
                                }
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "FLD01" | pVal.ItemUID == "FLD02")
                    {
                        oForm.PaneLevel = Convert.ToInt32(codeHelpClass.Right(pVal.ItemUID, 2));
                    }
                    if (pVal.ItemUID == "1" & pVal.ActionSuccess == true & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        ///문서추가
                        /// 가져오기
                    }
                    else if (pVal.ItemUID == "Btn1")
                    {
                        if (PH_PY106_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            Display_PH_PY106();
                        }
                        /// 공식 검증
                    }
                    else if (pVal.ItemUID == "Btn2")
                    {
                        PH_PY106_Display_CsuItem();
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
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                        if (pVal.ItemUID == "Mat1")
                        {
                            if (pVal.ColUID == "CSUCOD")
                            {
                                PH_PY106_AddMatrixRow();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                        if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "CSUCOD")
                            {
                                PH_PY106_AddMatrixRow();
                                oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
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
                if (pVal.BeforeAction == true & pVal.ItemUID == "YM" & pVal.CharPressed == 9 & oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    if (oMat01.RowCount > 0)
                    {
                        oMat01.Columns.Item("LINSEQ").Cells.Item(oMat01.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        BubbleEvent = false;
                    }
                }
                else if (pVal.BeforeAction == true & pVal.ColUID == "LINSEQ" & pVal.CharPressed == 9)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("순서는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
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
                    oMat01.LoadFromDataSource();
                    oMat02.LoadFromDataSource();
                    oMat03.LoadFromDataSource();
                    PH_PY106_FormItemEnabled();
                    PH_PY106_AddMatrixRow();
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
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
                if (pVal.FormUID == oForm.UniqueID & pVal.BeforeAction == true & oLastItemUID == "Mat1" & oLastColUID == "LINSEQ" & oLastColRow > 0 & (oLastItemUID != pVal.ColUID | oLastColRow != pVal.Row) & pVal.ItemUID != "1000001" & pVal.ItemUID != "2")
                {
                    if (oLastColRow > oMat01.VisualRowCount)
                    {
                        return;
                    }
                }
                else if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE & pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "Mat1" & pVal.Row > 0)
                    {
                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx = oMat01.Columns.Item("SILCUN").Cells.Item(pVal.Row).Specific.VALUE;
                    }
                    else if (pVal.ItemUID == "Mat02" & pVal.Row > 0)
                    {
                        //UPGRADE_WARNING: oMat02.Columns(FILCOD).Cells(pVal.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx = oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx + oMat02.Columns.Item("FILCOD").Cells.Item(pVal.Row).Specific.VALUE;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY106A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY106B);
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
		/// FORM_RESIZE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {

                }
                else if (pVal.BeforeAction == false)
                {
                    oForm.Freeze(true);
                    oForm.Items.Item("Mat02").Top = oForm.Items.Item("Mat1").Top;
                    oForm.Items.Item("Mat02").Left = oForm.Items.Item("Mat1").Width + 15;
                    oForm.Items.Item("Mat02").Width = Convert.ToInt32("240");
                    oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat1").Height;
                    oMat02.Columns.Item("Code").Width = Convert.ToInt32("20");
                    oMat02.Columns.Item("Name").Width = Convert.ToInt32("30");
                    oMat02.Columns.Item("FILCOD").Width = Convert.ToInt32("90");
                    oMat02.Columns.Item("REMARK").Width = Convert.ToInt32("80");
                    oForm.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY106A", "Code");
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
                            PH_PY106_FormItemEnabled();
                            PH_PY106_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //            Case "1293":
                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281":
                            ////문서찾기
                            PH_PY106_FormItemEnabled();
                            PH_PY106_AddMatrixRow();
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY106_FormItemEnabled();
                            PH_PY106_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY106_FormItemEnabled();
                            break;
                        case "1293":
                            //// 행삭제
                            //// [MAT1 용]
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                oMat01.FlushToDataSource();

                                while ((i <= oDS_PH_PY106B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_LINSEQ", i)))
                                    {
                                        oDS_PH_PY106B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY106B.Size; i++)
                                {
                                    oDS_PH_PY106B.SetValue("U_LINSEQ", i, Convert.ToString(i + 1));
                                }

                                oMat01.LoadFromDataSource();
                            }
                            PH_PY106_AddMatrixRow();
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
        public bool PH_PY106_DataValidCheck()
        {
            bool functionReturnValue = false;
            string sQry = string.Empty;
            string ExistYN = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_CLTCOD", 0).ToString().Trim())){
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_YM", 0).ToString().Trim())){
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업연월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_PAYTYP", 0).ToString().Trim())){
                    PSH_Globals.SBO_Application.StatusBar.SetText("급여형태는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("PAYTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_GNSGBN", 0).ToString().Trim())){
                    PSH_Globals.SBO_Application.StatusBar.SetText("근속일계산기준은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("GNSGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_BNSLEN", 0).ToString().Trim())){
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여 계산단위는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("BNSLEN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                if (string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_BNSRND", 0).ToString().Trim())){
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여 끝전처리는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("BNSRND").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //// Code & Name 생성
                oDS_PH_PY106A.SetValue("Code", 0, oDS_PH_PY106A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY106A.GetValue("U_YM", 0).ToString().Trim() + oDS_PH_PY106A.GetValue("U_PAYTYP", 0).ToString().Trim());
                oDS_PH_PY106A.SetValue("NAME", 0, oDS_PH_PY106A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY106A.GetValue("U_YM", 0).ToString().Trim() + oDS_PH_PY106A.GetValue("U_PAYTYP", 0).ToString().Trim());

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    ////저장된 데이터 체크
                    sQry = "SELECT Top 1 Code FROM [@PH_PY106A] ";
                    sQry = sQry + " WHERE Code = '" + oDS_PH_PY106A.GetValue("Code", 0).ToString().Trim() + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (oRecordSet01.Fields.Count > 0)
                    {
                        ExistYN = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                    }

                    if (!string.IsNullOrEmpty(ExistYN) & oDS_PH_PY106A.GetValue("Code", 0).ToString().Trim() != ExistYN)
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("Code" + "데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                functionReturnValue = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY106_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            functionReturnValue = true;

            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY106A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
                    throw new Exception();
                }
                //
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
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_Validate_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        public void PH_PY106_AddMatrixRow()
        {
            int oRow = 0;
            
            try
            {
                oForm.Freeze(true);

                ////[Mat1 용]
                oMat01.FlushToDataSource();
                oRow = oMat01.VisualRowCount;
                //
                if (oMat01.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_CSUCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY106B.Size <= oMat01.VisualRowCount)
                        {
                            oDS_PH_PY106B.InsertRecord((oRow));
                        }
                        oDS_PH_PY106B.Offset = oRow;
                        oDS_PH_PY106B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY106B.SetValue("U_LINSEQ", oRow, "");
                        oDS_PH_PY106B.SetValue("U_CSUCOD", oRow, "");
                        oDS_PH_PY106B.SetValue("U_CSUNAM", oRow, "");
                        oDS_PH_PY106B.SetValue("U_SILCUN", oRow, "");
                        oDS_PH_PY106B.SetValue("U_SILCOD", oRow, "");
                        oDS_PH_PY106B.SetValue("U_BNSBAS", oRow, "N");
                        oDS_PH_PY106B.SetValue("U_REMARK", oRow, "");
                        oMat01.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY106B.Offset = oRow - 1;
                        oDS_PH_PY106B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY106B.SetValue("U_LINSEQ", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_CSUCOD", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_CSUNAM", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_SILCUN", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_SILCOD", oRow - 1, "");
                        oDS_PH_PY106B.SetValue("U_BNSBAS", oRow - 1, "N");
                        oDS_PH_PY106B.SetValue("U_REMARK", oRow - 1, "");
                        oMat01.LoadFromDataSource();
                    }
                }
                else if (oMat01.VisualRowCount == 0)
                {
                    oDS_PH_PY106B.Offset = oRow;
                    oDS_PH_PY106B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY106B.SetValue("U_LINSEQ", oRow, "");
                    oDS_PH_PY106B.SetValue("U_CSUCOD", oRow, "");
                    oDS_PH_PY106B.SetValue("U_CSUNAM", oRow, "");
                    oDS_PH_PY106B.SetValue("U_SILCUN", oRow, "");
                    oDS_PH_PY106B.SetValue("U_SILCOD", oRow, "");
                    oDS_PH_PY106B.SetValue("U_BNSBAS", oRow, "N");
                    oDS_PH_PY106B.SetValue("U_REMARK", oRow, "");
                    oMat01.LoadFromDataSource();
                }
                //
                ////[Mat02 용]
                oMat03.FlushToDataSource();
                oRow = oMat03.VisualRowCount;
                //
                if (oMat03.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY106D.GetValue("U_CSUCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY106D.Size <= oMat03.VisualRowCount)
                        {
                            oDS_PH_PY106D.InsertRecord((oRow));
                        }
                        oDS_PH_PY106D.Offset = oRow;
                        oDS_PH_PY106D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY106D.SetValue("U_LINSEQ", oRow, "");
                        oDS_PH_PY106D.SetValue("U_Status", oRow, "");
                        oDS_PH_PY106D.SetValue("U_WorkType", oRow, "");
                        oDS_PH_PY106D.SetValue("U_Order", oRow, "");
                        oDS_PH_PY106D.SetValue("U_CSUCOD", oRow, "");
                        oDS_PH_PY106D.SetValue("U_SILCUN", oRow, "");
                        oDS_PH_PY106D.SetValue("U_REMARK", oRow, "");
                        oMat03.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY106D.Offset = oRow - 1;
                        oDS_PH_PY106D.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY106D.SetValue("U_LINSEQ", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_Status", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_WorkType", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_Order", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_CSUCOD", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_SILCUN", oRow - 1, "");
                        oDS_PH_PY106D.SetValue("U_REMARK", oRow - 1, "");
                        oMat03.LoadFromDataSource();
                    }
                }
                else if (oMat03.VisualRowCount == 0)
                {
                    oDS_PH_PY106D.Offset = oRow;
                    oDS_PH_PY106D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY106D.SetValue("U_LINSEQ", oRow, "");
                    oDS_PH_PY106D.SetValue("U_Status", oRow, "");
                    oDS_PH_PY106D.SetValue("U_WorkType", oRow, "");
                    oDS_PH_PY106D.SetValue("U_Order", oRow, "");
                    oDS_PH_PY106D.SetValue("U_CSUCOD", oRow, "");
                    oDS_PH_PY106D.SetValue("U_SILCUN", oRow, "");
                    oDS_PH_PY106D.SetValue("U_REMARK", oRow, "");
                    oMat03.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY106_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData( "AutoKey",  "ObjectCode",  "ONNM",  "'PH_PY106'",  "");
                if (Convert.ToDouble(DocEntry) == 0)
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY106_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 필수 사항 check
        /// 구현은 되어 있지만 사용하지 않음
        /// </summary>
        /// <returns></returns>
        private bool MatrixSpaceLineDel()
        {
            bool functionReturnValue = false;

            int iRow = 0;
            int kRow = 0;
            short ErrNum = 0;
            string Chk_Data = string.Empty;

            ErrNum = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.RowCount == 1)
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                for (iRow = 0; iRow <= oMat01.VisualRowCount - 2; iRow++)
                {
                    oDS_PH_PY106B.Offset = iRow;
                    if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow).ToString().Trim()))
                    {
                        ErrNum = 2;
                        oMat01.Columns.Item("CSUCOD").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_LINSEQ", iRow).ToString().Trim()) & codeHelpClass.Left(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow), 1) != "X")
                    {
                        ErrNum = 5;
                        oMat01.Columns.Item("LINSEQ").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_SILCUN", iRow).ToString().Trim()))
                    {
                        ErrNum = 4;
                        oMat01.Columns.Item("SILCUN").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (oDS_PH_PY106B.GetValue("U_BNSBAS", iRow).ToString().Trim() == "Y")
                    {
                        if (codeHelpClass.Left(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow).ToString().Trim(), 1) == "X")
                        {
                            ErrNum = 6;
                            oMat01.Columns.Item("BNSBAS").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (oDS_PH_PY106B.GetValue("U_CSUCOD", iRow).ToString().Trim() == "A04")
                        {
                            ErrNum = 7;
                            oMat01.Columns.Item("CSUCOD").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                    else
                    {
                        //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        //중복체크작업
                        //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                        Chk_Data = oDS_PH_PY106B.GetValue("U_CSUCOD", iRow).ToString().Trim();
                        for (kRow = iRow + 1; kRow <= oMat01.VisualRowCount - 2; kRow++)
                        {
                            oDS_PH_PY106B.Offset = kRow;
                            if (Chk_Data.ToString().Trim() ==oDS_PH_PY106B.GetValue("U_CSUCOD", kRow).ToString().Trim())
                            {
                                ErrNum = 3;
                                oMat01.Columns.Item("LINSEQ").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                        }
                    }
                }
                //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                ////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
                ////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
                //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                oDS_PH_PY106B.RemoveRecord(oDS_PH_PY106B.Size - 1);
                //// Mat1에 마지막라인(빈라인) 삭제

                //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                //행을 삭제하였으니 DB데이터 소스를 다시 가져온다
                //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
                oMat01.LoadFromDataSource();

            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("입력할 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("내용이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("계산식은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("순서는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("기본일급/통상일급/기본시급/통상시급은 상여지정을 할 수 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여금에는 상여지정을 할 수 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("MatrixSpaceLineDel_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                functionReturnValue = false;
                return functionReturnValue;
            }
            return functionReturnValue;
        }

        /// <summary>
        /// Display_PH_PY106
        /// </summary>
        private void Display_PH_PY106()
        {
            int i = 0;
            int cnt = 0;
            string sQry = string.Empty;
            string oCLTCOD = string.Empty;
            string oJOBYMM = string.Empty;
            string CSUCOD = string.Empty;
            string SILCUN = string.Empty;
            string SILTYP = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                /// Matrix2 초기화
                cnt = oDS_PH_PY106B.Size;
                if (cnt > 0)
                {
                    for (i = 1; i <= cnt - 1; i++)
                    {
                        oDS_PH_PY106B.RemoveRecord(oDS_PH_PY106B.Size - 1);
                    }
                }
                else
                {
                    oMat01.LoadFromDataSource();
                }
                oCLTCOD = oDS_PH_PY106A.GetValue("U_CLTCOD", 0).ToString().Trim();
                oJOBYMM = oDS_PH_PY106A.GetValue("U_YM", 0).ToString().Trim();
                i = 0;
                oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                /// 기본셋팅값 가져오기
                sQry = "SELECT Code, Name, U_FILCOD ,U_REMARK FROM [@PH_PY106C] WHERE Code BETWEEN 'X01' AND 'X05' ORDER BY CODE";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    if (i + 1 > oDS_PH_PY106B.Size)
                    {
                        oDS_PH_PY106B.InsertRecord((i));
                    }
                    oDS_PH_PY106B.Offset = i;
                    oDS_PH_PY106B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY106B.SetValue("U_LINSEQ", i, "");
                    oDS_PH_PY106B.SetValue("U_CSUCOD", i, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                    oDS_PH_PY106B.SetValue("U_CSUNAM", i, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oDS_PH_PY106B.SetValue("U_SILCOD", i, "");
                    oDS_PH_PY106B.SetValue("U_SILCUN", i, "");
                    oDS_PH_PY106B.SetValue("U_BNSBAS", i, "N");
                    oDS_PH_PY106B.SetValue("U_REMARK", i, oRecordSet01.Fields.Item(3).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                    i = i + 1;
                }
                cnt = i;
                sQry = "Exec PH_PY102 '" + oCLTCOD.ToString().Trim() + "','" + oJOBYMM.ToString().Trim() + "', '', '', '', ''";
                oRecordSet01.DoQuery(sQry);
                while (!oRecordSet01.EoF)
                {
                    CSUCOD = "";
                    SILTYP = "";
                    if (i + 1 > oDS_PH_PY106B.Size)
                    {
                        oDS_PH_PY106B.InsertRecord((i));
                    }
                    CSUCOD = oRecordSet01.Fields.Item("U_CSUCOD").Value.ToString().Trim();
                    oDS_PH_PY106B.Offset = i;
                    oDS_PH_PY106B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY106B.SetValue("U_LINSEQ", i, Convert.ToString(i - cnt + 1));
                    oDS_PH_PY106B.SetValue("U_CSUCOD", i, CSUCOD);
                    oDS_PH_PY106B.SetValue("U_CSUNAM", i, oRecordSet01.Fields.Item("U_CSUNAM").Value.ToString().Trim());
                    oDS_PH_PY106B.SetValue("U_SILCOD", i, "");
                    oDS_PH_PY106B.SetValue("U_BNSBAS", i, "N");
                    oDS_PH_PY106B.SetValue("U_REMARK", i, "");
                    if (CSUCOD == "A01")
                    {
                        SILCUN = "T1.U_STDAMT";
                    }
                    else
                    {
                        SILTYP = dataHelpClass.Get_ReData("U_FIXGBN + isnull(U_INSLIN,'')", "U_CSUCOD", "[@PH_PY102B]", "'" + CSUCOD + "'", " AND Code = '" + oCLTCOD.ToString().Trim() + oJOBYMM.ToString().Trim() + "'");
                        if (codeHelpClass.Left(SILTYP, 1) == "Y")
                        {
                            SILCUN = "T2.U_CSUD" + SILTYP.Replace("Y", "").PadLeft(2,'0');
                        }
                        else
                        {
                            SILCUN = "0";
                        }
                    }

                    oDS_PH_PY106B.SetValue("U_SILCUN", i, SILCUN);

                    i = i + 1;
                    oRecordSet01.MoveNext();
                }
                oMat01.LoadFromDataSource();
                PH_PY106_AddMatrixRow();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Display_PH_PY106_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Display_PH_PY106
        /// </summary>
        private void PH_PY106_Display_CsuItem()
        {
            string sQry = string.Empty;
            int i = 0;
            int cnt = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                /// Matrix2 초기화
                cnt = oDS_PH_PY106C.Size;
                if (cnt > 0)
                {
                    for (i = 1; i <= cnt - 1; i++)
                    {
                        oDS_PH_PY106C.RemoveRecord(oDS_PH_PY106C.Size - 1);
                    }
                }
                else
                {
                    oMat02.LoadFromDataSource();
                }

                i = 0;

                sQry = "SELECT Code, Name, U_FILCOD, U_REMARK FROM [@PH_PY106C] ORDER BY CODE";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    if (i + 1 > oDS_PH_PY106C.Size)
                    {
                        oDS_PH_PY106C.InsertRecord((i));
                    }
                    oDS_PH_PY106C.Offset = i;
                    oDS_PH_PY106C.SetValue("Code", i, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
                    oDS_PH_PY106C.SetValue("Name", i, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oDS_PH_PY106C.SetValue("U_FILCOD", i, oRecordSet01.Fields.Item(2).Value.ToString().Trim());
                    oDS_PH_PY106C.SetValue("U_REMARK", i, oRecordSet01.Fields.Item(3).Value.ToString().Trim());
                    i = i + 1;
                    oRecordSet01.MoveNext();
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY106_Display_CsuItem_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
//// ERROR: Not supported in C#: OptionDeclaration
//namespace MDC_HR_Addon
//{
//    internal class PH_PY106
//    {
//        ////********************************************************************************
//        ////  File           : PH_PY106.cls
//        ////  Module         : 인사관리 > 급여관리
//        ////  Desc           : 수당계산식설정
//        ////********************************************************************************

//        public string oFormUniqueID;
//        public SAPbouiCOM.Form oForm;


//        //'// 매트릭스 사용시
//        public SAPbouiCOM.Matrix oMat01;
//        public SAPbouiCOM.Matrix oMat02;
//        public SAPbouiCOM.Matrix oMat03;

//        private SAPbouiCOM.DBDataSource oDS_PH_PY106A;
//        private SAPbouiCOM.DBDataSource oDS_PH_PY106B;
//        private SAPbouiCOM.DBDataSource oDS_PH_PY106C;
//        private SAPbouiCOM.DBDataSource oDS_PH_PY106D;

//        private string oLastItemUID;
//        private string oLastColUID;
//        private int oLastColRow;

//        public void LoadForm(string oFromDocEntry01 = "")
//        {

//            int i = 0;
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            // ERROR: Not supported in C#: OnErrorStatement


//            oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY106.srf");
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//            oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//            for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//            {
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//            }
//            oFormUniqueID = "PH_PY106_" + GetTotalFormsCount();
//            SubMain.AddForms(this, oFormUniqueID, "PH_PY106");
//            PSH_Globals.SBO_Application.LoadBatchActions(out (oXmlDoc.xml));
//            oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);


//            oForm.SupportedModes = -1;
//            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//            oForm.DataBrowser.BrowseBy = "Code";
//            oForm.PaneLevel = 1;

//            oForm.Freeze(true);
//            PH_PY106_CreateItems();
//            PH_PY106_EnableMenus();
//            PH_PY106_SetDocument(oFromDocEntry01);
//            //    Call PH_PY106_FormResize

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

//        private bool PH_PY106_CreateItems()
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

//            SAPbobsCOM.Recordset oRecordSet01 = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            //    '//매트릭스--------------------------------------------------------------------------------------
//            oDS_PH_PY106A = oForm.DataSources.DBDataSources("@PH_PY106A");
//            ////헤더
//            oDS_PH_PY106B = oForm.DataSources.DBDataSources("@PH_PY106B");
//            ////라인
//            oDS_PH_PY106C = oForm.DataSources.DBDataSources("@PH_PY106C");
//            ////라인
//            oDS_PH_PY106D = oForm.DataSources.DBDataSources("@PH_PY106D");
//            ////라인

//            oForm.DataSources.UserDataSources.Add("DISSIL", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
//            oEdit = oForm.Items.Item("DISSIL").Specific;
//            ////공식
//            oEdit.DataBind.SetBound(true, "", "DISSIL");

//            oMat01 = oForm.Items.Item("Mat1").Specific;
//            oMat02 = oForm.Items.Item("Mat02").Specific;
//            oMat03 = oForm.Items.Item("Mat03").Specific;

//            oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//            oMat01.AutoResizeColumns();
//            oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//            oMat03.AutoResizeColumns();

//            ////----------------------------------------------------------------------------------------------
//            //// 헤더 설정
//            ////----------------------------------------------------------------------------------------------
//            /// 귀속년월
//            //UPGRADE_WARNING: oForm.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("YM").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");

//            ////사업장
//            oCombo = oForm.Items.Item("CLTCOD").Specific;
//            //    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//            //    Call SetReDataCombo(oForm, sQry, oCombo)
//            oForm.Items.Item("CLTCOD").DisplayDesc = true;

//            //// 1.급여형태-계산식
//            oCombo = oForm.Items.Item("PAYTYP").Specific;
//            sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P132' AND U_UseYN= 'Y'";
//            MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//            oForm.Items.Item("PAYTYP").DisplayDesc = true;

//            //// 1.근속년수 계산기준
//            oCombo = oForm.Items.Item("GNSGBN").Specific;
//            oCombo.ValidValues.Add("1", "그룹입사일");
//            oCombo.ValidValues.Add("2", "입사  일자");
//            if (oCombo.ValidValues.Count >= 1)
//            {
//                oCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
//            }

//            //// 2.상여 계산단위 (
//            oCombo = oForm.Items.Item("BNSLEN").Specific;
//            oCombo.ValidValues.Add("1", "  원");
//            oCombo.ValidValues.Add("10", "십원");
//            oCombo.ValidValues.Add("100", "백원");
//            oCombo.ValidValues.Add("1000", "천원");
//            if (oCombo.ValidValues.Count >= 1)
//            {
//                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//            }

//            //// 3.상여 끝전처리
//            oCombo = oForm.Items.Item("BNSRND").Specific;
//            oCombo.ValidValues.Add("R", "반올림");
//            oCombo.ValidValues.Add("F", "절사");
//            oCombo.ValidValues.Add("C", "절상");
//            if (oCombo.ValidValues.Count >= 1)
//            {
//                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            oForm.Freeze(false);
//            return functionReturnValue;
//        PH_PY106_CreateItems_Error:

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
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            oForm.Freeze(false);
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        private void PH_PY106_EnableMenus()
//        {

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.EnableMenu("1283", true);
//            ////제거
//            oForm.EnableMenu("1284", false);
//            ////취소
//            oForm.EnableMenu("1293", true);
//            ////행삭제

//            return;
//        PH_PY106_EnableMenus_Error:

//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        private void PH_PY106_SetDocument(string oFromDocEntry01)
//        {
//            // ERROR: Not supported in C#: OnErrorStatement


//            if ((string.IsNullOrEmpty(oFromDocEntry01)))
//            {
//                PH_PY106_FormItemEnabled();
//                PH_PY106_AddMatrixRow();
//            }
//            else
//            {
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                PH_PY106_FormItemEnabled();
//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//            }
//            return;
//        PH_PY106_SetDocument_Error:

//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY106_FormItemEnabled()
//        {
//            SAPbouiCOM.ComboBox oCombo = null;

//            // ERROR: Not supported in C#: OnErrorStatement



//            oForm.Freeze(true);
//            if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//            {
//                oForm.Items.Item("CLTCOD").Enabled = true;
//                oForm.Items.Item("YM").Enabled = true;
//                oForm.Items.Item("PAYTYP").Enabled = true;
//                oForm.Items.Item("GNSGBN").Enabled = true;
//                oForm.Items.Item("BNSLEN").Enabled = true;
//                oForm.Items.Item("BNSRND").Enabled = true;
//                //        oMat01.Columns("CSUCOD").Editable = False
//                //        oMat01.Columns("CSUNAM").Editable = False

//                PH_PY106_Display_CsuItem();
//                ////Mat02 초기값 가져옴

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//                /// 귀속년월
//                //UPGRADE_WARNING: oForm.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("YM").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");

//                //// 1.근속년수 계산기준
//                oCombo = oForm.Items.Item("GNSGBN").Specific;
//                oCombo.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);

//                //// 2.상여 계산단위 (
//                oCombo = oForm.Items.Item("BNSLEN").Specific;
//                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                //// 3.상여 끝전처리
//                oCombo = oForm.Items.Item("BNSRND").Specific;
//                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                oForm.EnableMenu("1293", true);
//                ////행삭제
//                oForm.EnableMenu("1283", true);
//                ////제거

//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//            {
//                oForm.Items.Item("CLTCOD").Enabled = true;
//                oForm.Items.Item("YM").Enabled = true;
//                oForm.Items.Item("PAYTYP").Enabled = true;
//                oForm.Items.Item("GNSGBN").Enabled = true;
//                oForm.Items.Item("BNSLEN").Enabled = true;
//                oForm.Items.Item("BNSRND").Enabled = true;
//                //        oMat01.Columns("CSUCOD").Editable = False
//                //        oMat01.Columns("CSUNAM").Editable = False

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//                oForm.EnableMenu("1293", true);
//                ////행삭제
//                oForm.EnableMenu("1283", true);
//                ////제거
//            }
//            else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//            {
//                oForm.Items.Item("CLTCOD").Enabled = false;
//                oForm.Items.Item("YM").Enabled = false;
//                oForm.Items.Item("PAYTYP").Enabled = false;
//                oForm.Items.Item("GNSGBN").Enabled = false;
//                oForm.Items.Item("BNSLEN").Enabled = false;
//                oForm.Items.Item("BNSRND").Enabled = false;
//                //        oMat01.Columns("CSUCOD").Editable = False
//                //        oMat01.Columns("CSUNAM").Editable = False

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//                oForm.EnableMenu("1293", true);
//                ////행삭제
//                oForm.EnableMenu("1283", true);
//                ////제거

//            }
//            oForm.Freeze(false);
//            return;
//        PH_PY106_FormItemEnabled_Error:

//            oForm.Freeze(false);
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }


//        public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            string sQry = null;
//            int i = 0;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbobsCOM.Recordset oRecordSet01 = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
//                                if (PH_PY106_DataValidCheck() == false)
//                                {
//                                    BubbleEvent = false;
//                                    return;
//                                }
//                                else
//                                {
//                                    if (MatrixSpaceLineDel() == false)
//                                    {
//                                        BubbleEvent = false;
//                                    }
//                                }
//                            }
//                        }
//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemUID == "FLD01" | pVal.ItemUID == "FLD02")
//                        {
//                            oForm.PaneLevel = Convert.ToInt32(Strings.Right(pVal.ItemUID, 2));
//                        }
//                        if (pVal.ItemUID == "1" & pVal.ActionSuccess == true & oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                        {
//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
//                            ///문서추가
//                            /// 가져오기
//                        }
//                        else if (pVal.ItemUID == "Btn1")
//                        {
//                            if (PH_PY106_DataValidCheck() == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                            else
//                            {
//                                Display_PH_PY102();
//                            }
//                            /// 공식 검증
//                        }
//                        else if (pVal.ItemUID == "Btn2")
//                        {
//                            PH_PY106_Display_CsuItem();
//                        }
//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                    ////2
//                    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//                    ////추가모드에서 코드이벤트가 코드에서 일어 났을때
//                    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//                    if (pVal.BeforeAction == true & pVal.ItemUID == "YM" & pVal.CharPressed == 9 & pVal.FormMode != SAPbouiCOM.BoFormMode.fm_FIND_MODE)
//                    {
//                        if (oMat01.RowCount > 0)
//                        {
//                            oMat01.Columns.Item("LINSEQ").Cells.Item(oMat01.VisualRowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            BubbleEvent = false;
//                        }
//                    }
//                    else if (pVal.BeforeAction == true & pVal.ColUID == "LINSEQ" & pVal.CharPressed == 9)
//                    {
//                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String)))
//                        {
//                            PSH_Globals.SBO_Application.StatusBar.SetText("순서는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                            BubbleEvent = false;
//                        }
//                    }
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                    ////3
//                    switch (pVal.ItemUID)
//                    {
//                        case "Mat1":
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
//                    oForm.Freeze(true);
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemChanged == true)
//                        {

//                        }
//                    }
//                    oForm.Freeze(false);
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_CLICK:
//                    ////6
//                    if (pVal.FormUID == oForm.UniqueID & pVal.BeforeAction == true & oLastItemUID == "Mat1" & oLastColUID == "LINSEQ" & oLastColRow > 0 & (oLastItemUID != pVal.ColUID | oLastColRow != pVal.Row) & pVal.ItemUID != "1000001" & pVal.ItemUID != "2")
//                    {
//                        if (oLastColRow > oMat01.VisualRowCount)
//                        {
//                            return;
//                        }
//                    }
//                    else if (pVal.FormMode != SAPbouiCOM.BoFormMode.fm_FIND_MODE & pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemUID == "Mat1" & pVal.Row > 0)
//                        {
//                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx = oMat01.Columns.Item("SILCUN").Cells.Item(pVal.Row).Specific.VALUE;
//                        }
//                        else if (pVal.ItemUID == "Mat02" & pVal.Row > 0)
//                        {
//                            //UPGRADE_WARNING: oMat02.Columns(FILCOD).Cells(pVal.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx = oForm.DataSources.UserDataSources.Item("DISSIL").ValueEx + oMat02.Columns.Item("FILCOD").Cells.Item(pVal.Row).Specific.VALUE;
//                        }
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
//                    oForm.Freeze(true);
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        if (pVal.ItemChanged == true)
//                        {
//                            if (pVal.ItemUID == "Mat1")
//                            {
//                                if (pVal.ColUID == "CSUCOD")
//                                {
//                                    PH_PY106_AddMatrixRow();
//                                    oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                                }
//                            }
//                            if (pVal.ItemUID == "Mat03")
//                            {
//                                if (pVal.ColUID == "CSUCOD")
//                                {
//                                    PH_PY106_AddMatrixRow();
//                                    oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                                }
//                            }
//                        }
//                    }
//                    oForm.Freeze(false);
//                    break;
//                //----------------------------------------------------------
//                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//                    ////11
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        oMat01.LoadFromDataSource();
//                        oMat02.LoadFromDataSource();
//                        oMat03.LoadFromDataSource();
//                        PH_PY106_FormItemEnabled();
//                        PH_PY106_AddMatrixRow();
//                    }
//                    break;
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
//                        //UPGRADE_NOTE: oDS_PH_PY106A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oDS_PH_PY106A = null;
//                        //UPGRADE_NOTE: oDS_PH_PY106B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oDS_PH_PY106B = null;
//                        //UPGRADE_NOTE: oDS_PH_PY106C 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oDS_PH_PY106C = null;
//                        //UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oMat01 = null;
//                        //UPGRADE_NOTE: oMat02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                        oMat02 = null;
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
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.BeforeAction == false)
//                    {
//                        oForm.Freeze(true);
//                        oForm.Items.Item("Mat02").Top = oForm.Items.Item("Mat1").Top;
//                        oForm.Items.Item("Mat02").Left = oForm.Items.Item("Mat1").Width + 15;
//                        oForm.Items.Item("Mat02").Width = Convert.ToInt32("240");
//                        oForm.Items.Item("Mat02").Height = oForm.Items.Item("Mat1").Height;
//                        oMat02.Columns.Item("Code").Width = Convert.ToInt32("20");
//                        oMat02.Columns.Item("Name").Width = Convert.ToInt32("30");
//                        oMat02.Columns.Item("FILCOD").Width = Convert.ToInt32("90");
//                        oMat02.Columns.Item("REMARK").Width = Convert.ToInt32("80");
//                        oForm.Freeze(false);
//                    }
//                    break;
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
//                    if (pVal.BeforeAction == true)
//                    {

//                    }
//                    else if (pVal.Before_Action == false)
//                    {

//                    }
//                    break;
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
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;

//            return;
//        Raise_FormItemEvent_Error:
//            ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//            oForm.Freeze((false));
//            //UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oCombo = null;
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
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
//                        MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@PH_PY106A", ref "Code");
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
//                        PH_PY106_FormItemEnabled();
//                        PH_PY106_AddMatrixRow();
//                        break;
//                    case "1284":
//                        break;
//                    case "1286":
//                        break;
//                    //            Case "1293":
//                    //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
//                    case "1281":
//                        ////문서찾기
//                        PH_PY106_FormItemEnabled();
//                        PH_PY106_AddMatrixRow();
//                        break;
//                    case "1282":
//                        ////문서추가
//                        PH_PY106_FormItemEnabled();
//                        PH_PY106_AddMatrixRow();
//                        break;
//                    case "1288":
//                    case "1289":
//                    case "1290":
//                    case "1291":
//                        PH_PY106_FormItemEnabled();
//                        break;
//                    case "1293":
//                        //// 행삭제
//                        //// [MAT1 용]
//                        if (oMat01.RowCount != oMat01.VisualRowCount)
//                        {
//                            oMat01.FlushToDataSource();

//                            while ((i <= oDS_PH_PY106B.Size - 1))
//                            {
//                                if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_LINSEQ", i)))
//                                {
//                                    oDS_PH_PY106B.RemoveRecord((i));
//                                    i = 0;
//                                }
//                                else
//                                {
//                                    i = i + 1;
//                                }
//                            }

//                            for (i = 0; i <= oDS_PH_PY106B.Size; i++)
//                            {
//                                oDS_PH_PY106B.SetValue("U_LINSEQ", i, Convert.ToString(i + 1));
//                            }

//                            oMat01.LoadFromDataSource();
//                        }
//                        //                '// [Mat02 용]
//                        //                 If oMat02.RowCount <> oMat02.VisualRowCount Then
//                        //                    oMat02.FlushToDataSource
//                        //
//                        //                    While (i <= oDS_PH_PY106C.Size - 1)
//                        //                        If oDS_PH_PY106C.GetValue("U_FILD01", i) = "" Then
//                        //                            oDS_PH_PY106C.RemoveRecord (i)
//                        //                            i = 0
//                        //                        Else
//                        //                            i = i + 1
//                        //                        End If
//                        //                    Wend
//                        //
//                        //                    For i = 0 To oDS_PH_PY106B.Size
//                        //                        Call oDS_PH_PY106B.setValue("U_LineNum", i, i + 1)
//                        //                    Next i
//                        //
//                        //                    oMat01.LoadFromDataSource
//                        //                End If
//                        PH_PY106_AddMatrixRow();
//                        break;
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

//        public void PH_PY106_AddMatrixRow()
//        {
//            int oRow = 0;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oForm.Freeze(true);

//            ////[Mat1 용]
//            oMat01.FlushToDataSource();
//            oRow = oMat01.VisualRowCount;
//            //
//            if (oMat01.VisualRowCount > 0)
//            {
//                if (!string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_CSUCOD", oRow - 1))))
//                {
//                    if (oDS_PH_PY106B.Size <= oMat01.VisualRowCount)
//                    {
//                        oDS_PH_PY106B.InsertRecord((oRow));
//                    }
//                    oDS_PH_PY106B.Offset = oRow;
//                    oDS_PH_PY106B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                    oDS_PH_PY106B.SetValue("U_LINSEQ", oRow, "");
//                    oDS_PH_PY106B.SetValue("U_CSUCOD", oRow, "");
//                    oDS_PH_PY106B.SetValue("U_CSUNAM", oRow, "");
//                    oDS_PH_PY106B.SetValue("U_SILCUN", oRow, "");
//                    oDS_PH_PY106B.SetValue("U_SILCOD", oRow, "");
//                    oDS_PH_PY106B.SetValue("U_BNSBAS", oRow, "N");
//                    oDS_PH_PY106B.SetValue("U_REMARK", oRow, "");
//                    oMat01.LoadFromDataSource();
//                }
//                else
//                {
//                    oDS_PH_PY106B.Offset = oRow - 1;
//                    oDS_PH_PY106B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//                    oDS_PH_PY106B.SetValue("U_LINSEQ", oRow - 1, "");
//                    oDS_PH_PY106B.SetValue("U_CSUCOD", oRow - 1, "");
//                    oDS_PH_PY106B.SetValue("U_CSUNAM", oRow - 1, "");
//                    oDS_PH_PY106B.SetValue("U_SILCUN", oRow - 1, "");
//                    oDS_PH_PY106B.SetValue("U_SILCOD", oRow - 1, "");
//                    oDS_PH_PY106B.SetValue("U_BNSBAS", oRow - 1, "N");
//                    oDS_PH_PY106B.SetValue("U_REMARK", oRow - 1, "");
//                    oMat01.LoadFromDataSource();
//                }
//            }
//            else if (oMat01.VisualRowCount == 0)
//            {
//                oDS_PH_PY106B.Offset = oRow;
//                oDS_PH_PY106B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                oDS_PH_PY106B.SetValue("U_LINSEQ", oRow, "");
//                oDS_PH_PY106B.SetValue("U_CSUCOD", oRow, "");
//                oDS_PH_PY106B.SetValue("U_CSUNAM", oRow, "");
//                oDS_PH_PY106B.SetValue("U_SILCUN", oRow, "");
//                oDS_PH_PY106B.SetValue("U_SILCOD", oRow, "");
//                oDS_PH_PY106B.SetValue("U_BNSBAS", oRow, "N");
//                oDS_PH_PY106B.SetValue("U_REMARK", oRow, "");
//                oMat01.LoadFromDataSource();
//            }
//            //
//            ////[Mat02 용]
//            oMat03.FlushToDataSource();
//            oRow = oMat03.VisualRowCount;
//            //
//            if (oMat03.VisualRowCount > 0)
//            {
//                if (!string.IsNullOrEmpty(oDS_PH_PY106D.GetValue("U_CSUCOD", oRow - 1))))
//                {
//                    if (oDS_PH_PY106D.Size <= oMat03.VisualRowCount)
//                    {
//                        oDS_PH_PY106D.InsertRecord((oRow));
//                    }
//                    oDS_PH_PY106D.Offset = oRow;
//                    oDS_PH_PY106D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                    oDS_PH_PY106D.SetValue("U_LINSEQ", oRow, "");
//                    oDS_PH_PY106D.SetValue("U_Status", oRow, "");
//                    oDS_PH_PY106D.SetValue("U_WorkType", oRow, "");
//                    oDS_PH_PY106D.SetValue("U_Order", oRow, "");
//                    oDS_PH_PY106D.SetValue("U_CSUCOD", oRow, "");
//                    oDS_PH_PY106D.SetValue("U_SILCUN", oRow, "");
//                    oDS_PH_PY106D.SetValue("U_REMARK", oRow, "");
//                    oMat03.LoadFromDataSource();
//                }
//                else
//                {
//                    oDS_PH_PY106D.Offset = oRow - 1;
//                    oDS_PH_PY106D.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//                    oDS_PH_PY106D.SetValue("U_LINSEQ", oRow - 1, "");
//                    oDS_PH_PY106D.SetValue("U_Status", oRow - 1, "");
//                    oDS_PH_PY106D.SetValue("U_WorkType", oRow - 1, "");
//                    oDS_PH_PY106D.SetValue("U_Order", oRow - 1, "");
//                    oDS_PH_PY106D.SetValue("U_CSUCOD", oRow - 1, "");
//                    oDS_PH_PY106D.SetValue("U_SILCUN", oRow - 1, "");
//                    oDS_PH_PY106D.SetValue("U_REMARK", oRow - 1, "");
//                    oMat03.LoadFromDataSource();
//                }
//            }
//            else if (oMat03.VisualRowCount == 0)
//            {
//                oDS_PH_PY106D.Offset = oRow;
//                oDS_PH_PY106D.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                oDS_PH_PY106D.SetValue("U_LINSEQ", oRow, "");
//                oDS_PH_PY106D.SetValue("U_Status", oRow, "");
//                oDS_PH_PY106D.SetValue("U_WorkType", oRow, "");
//                oDS_PH_PY106D.SetValue("U_Order", oRow, "");
//                oDS_PH_PY106D.SetValue("U_CSUCOD", oRow, "");
//                oDS_PH_PY106D.SetValue("U_SILCUN", oRow, "");
//                oDS_PH_PY106D.SetValue("U_REMARK", oRow, "");
//                oMat03.LoadFromDataSource();
//            }
//            //
//            oForm.Freeze(false);
//            return;
//        PH_PY106_AddMatrixRow_Error:
//            oForm.Freeze(false);
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public void PH_PY106_FormClear()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            string DocEntry = null;
//            //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY106'", ref "");
//            if (Convert.ToDouble(DocEntry) == 0)
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//            }
//            else
//            {
//                //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//            }
//            return;
//        PH_PY106_FormClear_Error:
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        }

//        public bool PH_PY106_DataValidCheck()
//        {
//            bool functionReturnValue = false;
//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            string sQry = null;
//            string ExistYN = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            functionReturnValue = false;


//            /// Check
//            switch (true)
//            {
//                case string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_CLTCOD", 0)):
//                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//                case MDC_SetMod.ChkYearMonth(ref oDS_PH_PY106A.GetValue("U_YM", 0)) == false:
//                    PSH_Globals.SBO_Application.StatusBar.SetText("작업연월은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//                case string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_PAYTYP", 0)):
//                    PSH_Globals.SBO_Application.StatusBar.SetText("급여형태는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                    oForm.Items.Item("PAYTYP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//                case string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_GNSGBN", 0)):
//                    PSH_Globals.SBO_Application.StatusBar.SetText("근속일계산기준은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                    oForm.Items.Item("GNSGBN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//                case string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_BNSLEN", 0)):
//                    PSH_Globals.SBO_Application.StatusBar.SetText("상여 계산단위는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                    oForm.Items.Item("BNSLEN").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//                case string.IsNullOrEmpty(oDS_PH_PY106A.GetValue("U_BNSRND", 0)):
//                    PSH_Globals.SBO_Application.StatusBar.SetText("상여 끝전처리는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                    oForm.Items.Item("BNSRND").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    functionReturnValue = false;
//                    return functionReturnValue;
//            }

//            //// Code & Name 생성
//            oDS_PH_PY106A.SetValue("Code", 0, oDS_PH_PY106A.GetValue("U_CLTCOD", 0)) + oDS_PH_PY106A.GetValue("U_YM", 0)) + oDS_PH_PY106A.GetValue("U_PAYTYP", 0)));
//            oDS_PH_PY106A.SetValue("NAME", 0, oDS_PH_PY106A.GetValue("U_CLTCOD", 0)) + oDS_PH_PY106A.GetValue("U_YM", 0)) + oDS_PH_PY106A.GetValue("U_PAYTYP", 0)));

//            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//            {
//                ////저장된 데이터 체크
//                sQry = "SELECT Top 1 Code FROM [@PH_PY106A] ";
//                sQry = sQry + " WHERE Code = '" + oDS_PH_PY106A.GetValue("Code", 0)) + "'";
//                oRecordSet01.DoQuery(sQry);

//                if (oRecordSet01.Fields.Count > 0)
//                {
//                    ExistYN = oRecordSet01.Fields.Item(0).Value);
//                }

//                if (!string.IsNullOrEmpty(ExistYN) & oDS_PH_PY106A.GetValue("Code", 0)) != ExistYN)
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("Code" + "데이터가 일치합니다. 저장되지 않습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//                }
//            }

//            functionReturnValue = true;
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            return functionReturnValue;
//        PH_PY106_DataValidCheck_Error:

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            functionReturnValue = false;
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }

//        private bool MatrixSpaceLineDel()
//        {
//            bool functionReturnValue = false;
//            //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//            //저장할 데이터의 유효성을 점검한다
//            //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//            // ERROR: Not supported in C#: OnErrorStatement

//            int iRow = 0;
//            int kRow = 0;
//            short ErrNum = 0;
//            string Chk_Data = null;

//            ErrNum = 0;

//            oMat01.FlushToDataSource();

//            if (oMat01.RowCount == 1)
//            {
//                ErrNum = 1;
//                goto Error_Message;
//            }

//            for (iRow = 0; iRow <= oMat01.VisualRowCount - 2; iRow++)
//            {
//                oDS_PH_PY106B.Offset = iRow;
//                if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow))))
//                {
//                    ErrNum = 2;
//                    oMat01.Columns.Item("CSUCOD").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    goto Error_Message;
//                }
//                else if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_LINSEQ", iRow))) & Strings.Left(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow), 1) != "X")
//                {
//                    ErrNum = 5;
//                    oMat01.Columns.Item("LINSEQ").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    goto Error_Message;
//                }
//                else if (string.IsNullOrEmpty(oDS_PH_PY106B.GetValue("U_SILCUN", iRow))))
//                {
//                    ErrNum = 4;
//                    oMat01.Columns.Item("SILCUN").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    goto Error_Message;
//                }
//                else if (oDS_PH_PY106B.GetValue("U_BNSBAS", iRow)) == "Y")
//                {
//                    if (Strings.Left(oDS_PH_PY106B.GetValue("U_CSUCOD", iRow)), 1) == "X")
//                    {
//                        ErrNum = 6;
//                        oMat01.Columns.Item("BNSBAS").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        goto Error_Message;
//                    }
//                    else if (oDS_PH_PY106B.GetValue("U_CSUCOD", iRow)) == "A04")
//                    {
//                        ErrNum = 7;
//                        oMat01.Columns.Item("CSUCOD").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                        goto Error_Message;

//                    }
//                }
//                else
//                {
//                    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//                    //중복체크작업
//                    //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//                    Chk_Data = oDS_PH_PY106B.GetValue("U_CSUCOD", iRow));
//                    for (kRow = iRow + 1; kRow <= oMat01.VisualRowCount - 2; kRow++)
//                    {
//                        oDS_PH_PY106B.Offset = kRow;
//                        if (Chk_Data) == oDS_PH_PY106B.GetValue("U_CSUCOD", kRow)))
//                        {
//                            ErrNum = 3;
//                            oMat01.Columns.Item("LINSEQ").Cells.Item(iRow + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            goto Error_Message;
//                        }
//                    }
//                }
//            }

//            //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//            ////맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
//            ////이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
//            //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//            oDS_PH_PY106B.RemoveRecord(oDS_PH_PY106B.Size - 1);
//            //// Mat1에 마지막라인(빈라인) 삭제

//            //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//            //행을 삭제하였으니 DB데이터 소스를 다시 가져온다
//            //ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//            oMat01.LoadFromDataSource();

//            functionReturnValue = true;
//            return functionReturnValue;
//        Error_Message:
//            ///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//            if (ErrNum == 1)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("입력할 데이터가 없습니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//            else if (ErrNum == 2)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//            else if (ErrNum == 3)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("내용이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//            else if (ErrNum == 4)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("계산식은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//            else if (ErrNum == 5)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("순서는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//            else if (ErrNum == 6)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("기본일급/통상일급/기본시급/통상시급은 상여지정을 할 수 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//            else if (ErrNum == 7)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("상여금에는 상여지정을 할 수 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//            else
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("MatrixSpaceLineDel Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//            }
//            functionReturnValue = false;
//            return functionReturnValue;
//        }

//        public bool PH_PY106_Validate(string ValidateType)
//        {
//            bool functionReturnValue = false;
//            // ERROR: Not supported in C#: OnErrorStatement

//            functionReturnValue = true;
//            object i = null;
//            int j = 0;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            //UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY106A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY106A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                functionReturnValue = false;
//                goto PH_PY106_Validate_Exit;
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
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            return functionReturnValue;
//        PH_PY106_Validate_Exit:
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            return functionReturnValue;
//        PH_PY106_Validate_Error:
//            functionReturnValue = false;
//            PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY106_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            return functionReturnValue;
//        }



//        ///수당
//        private void Display_PH_PY102()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            string sQry = null;
//            string oCLTCOD = null;
//            string oJOBYMM = null;
//            short i = 0;
//            short cnt = 0;
//            string CSUCOD = null;
//            string SILCUN = null;
//            string SILTYP = null;

//            /// Matrix2 초기화
//            cnt = oDS_PH_PY106B.Size;
//            if (cnt > 0)
//            {
//                for (i = 1; i <= cnt - 1; i++)
//                {
//                    oDS_PH_PY106B.RemoveRecord(oDS_PH_PY106B.Size - 1);
//                }
//            }
//            else
//            {
//                oMat01.LoadFromDataSource();
//            }
//            oCLTCOD = oDS_PH_PY106A.GetValue("U_CLTCOD", 0));
//            oJOBYMM = oDS_PH_PY106A.GetValue("U_YM", 0));
//            i = 0;
//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            /// 기본셋팅값 가져오기
//            sQry = "SELECT Code, Name, U_FILCOD ,U_REMARK FROM [@PH_PY106C] WHERE Code BETWEEN 'X01' AND 'X05' ORDER BY CODE";
//            oRecordSet01.DoQuery(sQry);
//            while (!(oRecordSet01.EoF))
//            {
//                if (i + 1 > oDS_PH_PY106B.Size)
//                {
//                    oDS_PH_PY106B.InsertRecord((i));
//                }
//                oDS_PH_PY106B.Offset = i;
//                oDS_PH_PY106B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                oDS_PH_PY106B.SetValue("U_LINSEQ", i, "");
//                oDS_PH_PY106B.SetValue("U_CSUCOD", i, oRecordSet01.Fields.Item(0).Value));
//                oDS_PH_PY106B.SetValue("U_CSUNAM", i, oRecordSet01.Fields.Item(1).Value));
//                oDS_PH_PY106B.SetValue("U_SILCOD", i, "");
//                oDS_PH_PY106B.SetValue("U_SILCUN", i, "");
//                oDS_PH_PY106B.SetValue("U_BNSBAS", i, "N");
//                oDS_PH_PY106B.SetValue("U_REMARK", i, oRecordSet01.Fields.Item(3).Value));
//                oRecordSet01.MoveNext();
//                i = i + 1;
//            }
//            cnt = i;
//            sQry = "Exec PH_PY102 '" + oCLTCOD) + "','" + oJOBYMM) + "', '', '', '', ''";
//            oRecordSet01.DoQuery(sQry);
//            while (!(oRecordSet01.EoF))
//            {
//                CSUCOD = "";
//                SILTYP = "";
//                if (i + 1 > oDS_PH_PY106B.Size)
//                {
//                    oDS_PH_PY106B.InsertRecord((i));
//                }
//                CSUCOD = oRecordSet01.Fields.Item("U_CSUCOD").Value);
//                oDS_PH_PY106B.Offset = i;
//                oDS_PH_PY106B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                oDS_PH_PY106B.SetValue("U_LINSEQ", i, Convert.ToString(i - cnt + 1));
//                oDS_PH_PY106B.SetValue("U_CSUCOD", i, CSUCOD);
//                oDS_PH_PY106B.SetValue("U_CSUNAM", i, oRecordSet01.Fields.Item("U_CSUNAM").Value));
//                oDS_PH_PY106B.SetValue("U_SILCOD", i, "");
//                oDS_PH_PY106B.SetValue("U_BNSBAS", i, "N");
//                oDS_PH_PY106B.SetValue("U_REMARK", i, "");
//                if (CSUCOD) == "A01")
//                {
//                    SILCUN = "T1.U_STDAMT";
//                }
//                else
//                {
//                    //UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    SILTYP = MDC_SetMod.Get_ReData(ref "U_FIXGBN + U_INSLIN", ref "U_CSUCOD", ref "[@PH_PY102B]", ref "'" + CSUCOD + "'", ref " AND Code = '" + oCLTCOD) + oJOBYMM) + "'");
//                    if (Strings.Left(SILTYP, 1) == "Y")
//                    {
//                        SILCUN = "T2.U_CSUD" + Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Strings.Mid(SILTYP, 2), "00");
//                    }
//                    else
//                    {
//                        SILCUN = "0";
//                    }
//                }

//                oDS_PH_PY106B.SetValue("U_SILCUN", i, SILCUN);

//                i = i + 1;
//                oRecordSet01.MoveNext();
//            }
//            oMat01.LoadFromDataSource();
//            PH_PY106_AddMatrixRow();

//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            return;
//        Error_Message:
//            ///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            PSH_Globals.SBO_Application.StatusBar.SetText("Display_PH_PY102 Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        }


//        private void PH_PY106_Display_CsuItem()
//        {
//            // ERROR: Not supported in C#: OnErrorStatement

//            SAPbobsCOM.Recordset oRecordSet01 = null;
//            string sQry = null;
//            short i = 0;
//            short cnt = 0;

//            oRecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            /// Matrix2 초기화
//            cnt = oDS_PH_PY106C.Size;
//            if (cnt > 0)
//            {
//                for (i = 1; i <= cnt - 1; i++)
//                {
//                    oDS_PH_PY106C.RemoveRecord(oDS_PH_PY106C.Size - 1);
//                }
//            }
//            else
//            {
//                oMat02.LoadFromDataSource();
//            }

//            i = 0;

//            sQry = "SELECT Code, Name, U_FILCOD, U_REMARK FROM [@PH_PY106C] ORDER BY CODE";
//            oRecordSet01.DoQuery(sQry);
//            while (!(oRecordSet01.EoF))
//            {
//                if (i + 1 > oDS_PH_PY106C.Size)
//                {
//                    oDS_PH_PY106C.InsertRecord((i));
//                }
//                oDS_PH_PY106C.Offset = i;
//                oDS_PH_PY106C.SetValue("Code", i, oRecordSet01.Fields.Item(0).Value));
//                oDS_PH_PY106C.SetValue("Name", i, oRecordSet01.Fields.Item(1).Value));
//                oDS_PH_PY106C.SetValue("U_FILCOD", i, oRecordSet01.Fields.Item(2).Value));
//                oDS_PH_PY106C.SetValue("U_REMARK", i, oRecordSet01.Fields.Item(3).Value));
//                i = i + 1;
//                oRecordSet01.MoveNext();
//            }
//            oMat02.LoadFromDataSource();
//            oMat02.AutoResizeColumns();
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            return;
//        Error_Message:
//            ///'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet01 = null;
//            PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY106_Display_CsuItem Error :" + Strings.Space(10) + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//        }
//    }
//}
