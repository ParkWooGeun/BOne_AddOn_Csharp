using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 성과급 차등 개인별 계산
    /// </summary>
    internal class PH_PY132 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;

        private SAPbouiCOM.DBDataSource oDS_PH_PY132A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY132B;

        private string oLastItemUID = string.Empty;
        private string oLastColUID = string.Empty;
        private int oLastColRow = 0;

        string g_preBankSel = string.Empty;
        private string oJOBTYP = string.Empty;
        private string oJOBGBN = string.Empty;
        private string oYM = string.Empty;

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY132.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY132_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY132");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY132_CreateItems();
                PH_PY132_EnableMenus();
                PH_PY132_SetDocument(oFromDocEntry01);
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
        private void PH_PY132_CreateItems()
        {
            int iCol = 0;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);

                oDS_PH_PY132A = oForm.DataSources.DBDataSources.Item("@PH_PY132A");
                oDS_PH_PY132B = oForm.DataSources.DBDataSources.Item("@PH_PY132B");

                oMat1 = oForm.Items.Item("Mat1").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                // 귀속연월
                oDS_PH_PY132A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));

                // 지급구분
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific,"");
                oForm.Items.Item("JOBGBN").DisplayDesc = true;
                oForm.Items.Item("JOBGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                // 상여
                for (iCol = 1; iCol <= 8; iCol++)
                {
                    oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Add("1", "개월 이상");
                    oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Add("2", "개월 미만");
                    oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Add("3", "일수 이상");
                    oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Add("4", "일수 미만");
                    if (oForm.Items.Item("AP" + iCol + "GBN").Specific.ValidValues.Count > 0)
                    {
                        oDS_PH_PY132A.SetValue("U_AP" + iCol + "GBN", 0, "1");
                    }
                    oForm.Items.Item("AP" + iCol + "GBN").DisplayDesc = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY132_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY132_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1284", false);               // 취소
                oForm.EnableMenu("1287", true);                // 복제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY132_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY132_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.ActiveItem = "YM";
                    oForm.Items.Item("Code").Enabled = false;
                    oForm.Items.Item("Btn1").Visible = true;

                    oForm.EnableMenu("1281", true);                  // 문서찾기
                    oForm.EnableMenu("1282", false);                 // 문서추가
                    // 귀속연월
                    oDS_PH_PY132A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));
                    // 지급구분
                    oForm.Items.Item("JOBGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.Items.Item("Code").Enabled = true;
                    oForm.Items.Item("Btn1").Visible = false;
                    oForm.ActiveItem = "Code";

                    oForm.EnableMenu("1281", false);                   // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.ActiveItem = "YM";
                    oForm.Items.Item("Code").Enabled = false;
                    oForm.Items.Item("Btn1").Visible = true;

                    oForm.EnableMenu("1281", true);                    // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY132_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private void PH_PY132_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY132_FormItemEnabled();
                    PH_PY132_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY132_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY132_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        public void PH_PY132_AddMatrixRow()
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
                    if (!string.IsNullOrEmpty(oDS_PH_PY132B.GetValue("U_CLTCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY132B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY132B.InsertRecord((oRow));
                        }
                        oDS_PH_PY132B.Offset = oRow;
                        oDS_PH_PY132B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY132B.Offset = oRow - 1;
                        oDS_PH_PY132B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY132B.Offset = oRow;
                    oDS_PH_PY132B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY132_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY132A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY132B);
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            String tDocEntry = String.Empty;
            bool CalcYN = false;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            // 유효성 검사
                            if (PH_PY132_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        tDocEntry = oForm.Items.Item("Code").Specific.VALUE;
                    }
                    else if (pVal.ItemUID == "Btn1")  // 급(상)여계산
                    {
                        CalcYN = true;
                        tDocEntry = oForm.Items.Item("Code").Specific.VALUE;
                        // 유효성 검사
                        if (PH_PY132_DataValidCheck() == true & Pay_Calc() == true)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("성과급 차등계산이 진행중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                if (oForm.Items.Item("JOBGBN").Specific.VALUE != "2")
                                {
                                    // 정상계산
                                    oRecordSet.DoQuery("EXEC PH_PY132 '" + tDocEntry + "'");
                                }
                                else
                                {
                                    // 소급계산
                                    oRecordSet.DoQuery("EXEC PH_PY132_SOGUB '" + tDocEntry + "'");
                                }
                                PSH_Globals.SBO_Application.StatusBar.SetText("성과급 차등계산이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            }
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            if (CalcYN == true)
                            {
                                CalcYN = false;
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("성과급 차등계산이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    PH_PY132_FormItemEnabled();
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                            }
                            else if (CalcYN == false)
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    PH_PY132_FormItemEnabled();
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                else
                                {
                                    PH_PY132_FormItemEnabled();
                                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
                if (pVal.BeforeAction == true & pVal.ItemUID == "YM" & pVal.CharPressed == 9 & pVal.FormMode == (int)SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.String))
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("귀속연월은 필수입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat1":
                            if (pVal.Row > 0)
                            {
                                oMat1.SelectRow(pVal.Row, true, false);
                            }
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
            string JIGBIL = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == false & pVal.ItemChanged == true)
                {
                    // 지급구분
                    if (pVal.ItemUID == "JOBGBN")
                    {
                        if ((oForm.Items.Item("JOBGBN").Specific.Selected != null))
                        {
                            oYM = oDS_PH_PY132A.GetValue("U_YM", 0).ToString().Trim();
                            JIGBIL = DateTime.Now.ToString("yyyyMMdd");
                            oForm.Items.Item("JIGBIL").Specific.VALUE = JIGBIL;
                            oJOBGBN = oDS_PH_PY132A.GetValue("U_JOBGBN", 0).ToString().Trim();
                        }
                        else
                        {
                            oJOBGBN = "";
                        }
                        if (!string.IsNullOrEmpty(oJOBGBN))
                            Display_BonussRate();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    PH_PY132_FormItemEnabled();
                    PH_PY132_AddMatrixRow();
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
                            break;

                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY132_FormItemEnabled();
                            PH_PY132_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //            Case "1293":
                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
                        case "1281":
                            ////문서찾기
                            PH_PY132_FormItemEnabled();
                            PH_PY132_AddMatrixRow();
                            break;

                        case "1282":
                            ////문서추가
                            PH_PY132_FormItemEnabled();
                            PH_PY132_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY132_FormItemEnabled();
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
        private bool PH_PY132_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            int ErrNum = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                functionReturnValue = true;

                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY132A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
                {
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
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY132_Validate_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
            return functionReturnValue;
        }

        /// <summary>
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == false & pVal.ItemChanged == true)
                {
                    // 귀속년월
                    if (pVal.ItemUID == "YM")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.VALUE.ToString().Trim()))
                        {
                            oDS_PH_PY132A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));
                        }
                        else
                        {
                            oDS_PH_PY132A.SetValue("U_YM", 0, oForm.Items.Item("YM").Specific.String);
                        }
                        oYM = oDS_PH_PY132A.GetValue("U_YM", 0).ToString().Trim();

                        if (!string.IsNullOrEmpty(oYM))
                            Display_BonussRate();
                    }

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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY132_DataValidCheck()
        {
            bool functionReturnValue = false;
            int ErrNum = 0;
            string oBNSRAT = string.Empty;
            string oJIGBIL = string.Empty;
            string sQry = string.Empty;
            string tCode = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oYM = oDS_PH_PY132A.GetValue("U_YM", 0).ToString().Trim();
                oJOBGBN = oDS_PH_PY132A.GetValue("U_JOBGBN", 0).ToString().Trim();
                oBNSRAT = oDS_PH_PY132A.GetValue("U_bBNSRAT", 0).ToString().Trim();
                oJIGBIL = oDS_PH_PY132A.GetValue("U_JIGBIL", 0).ToString().Trim();

                // Check
                if (dataHelpClass.ChkYearMonth(oYM) == false)
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oJOBGBN))
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                if (string.IsNullOrEmpty(oJIGBIL))
                {
                    ErrNum = 3;
                    throw new Exception();
                }
                // 상여계산시 체크
                if (oJOBTYP != "1")
                {
                    if (string.IsNullOrEmpty(oBNSRAT))
                    {
                        ErrNum = 4;
                        throw new Exception();
                    }
                }
                tCode = oDS_PH_PY132A.GetValue("U_YM", 0).ToString().Trim() + oDS_PH_PY132A.GetValue("U_JOBGBN", 0).ToString().Trim() + oDS_PH_PY132A.GetValue("U_Number", 0).ToString().Trim();

                // 코드생성
                // 코드 중복 체크
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    sQry = "SELECT CODE FROM [@PH_PY132A] WHERE CODE = '" + tCode + "'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount > 0)
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("코드가 존재합니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        functionReturnValue = false;
                    }
                    else
                    {
                        oDS_PH_PY132A.SetValue("Code", 0, tCode);
                        oDS_PH_PY132A.SetValue("Name", 0, tCode);
                    }
                }
                //// Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY132B.Size > 1)
                    oDS_PH_PY132B.RemoveRecord((oDS_PH_PY132B.Size - 1));

                oMat1.LoadFromDataSource();

                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("귀속 연월을 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급 구분을 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("상여율은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                { 
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY132_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
            return functionReturnValue;
        }

        /// <summary>
        /// Pay_Calc
        /// </summary>
        /// <returns></returns>
        private bool Pay_Calc()
        {
            bool functionReturnValue = false;
            int ErrNum = 0;

            string YM = string.Empty;
            string JOBGBN = string.Empty;
            string Number = string.Empty;
            string JIGBIL = string.Empty;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                YM     = oDS_PH_PY132A.GetValue("U_YM", 0).ToString().Trim();
                JOBGBN = oDS_PH_PY132A.GetValue("U_JOBGBN", 0).ToString().Trim();
                Number = oDS_PH_PY132A.GetValue("U_Number", 0).ToString().Trim();
                JIGBIL = oDS_PH_PY132A.GetValue("U_JIGBIL", 0).ToString().Trim();

                sQry = "Select Count(*) From [@PH_PY132A] ";
                sQry = sQry + " Where U_YM = '" + YM + "'";
                sQry = sQry + " AND U_Number = '" + Number + "'";
                sQry = sQry + " AND U_JIGBIL = '" + JIGBIL + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value > 0)
                {
                    if (PSH_Globals.SBO_Application.MessageBox("기존에 급여계산 결과가 있습니다. 계속 진행하시겠습니까?", 2, "Yes", "No") == 2)
                    {
                        ErrNum = 1;
                        functionReturnValue = false;
                    }
                    functionReturnValue = true;
                }
                else
                {
                    functionReturnValue = true;
                }
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("Pay_Calc_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
            return functionReturnValue;
        }

        /// <summary>
        /// Display_BonussRate
        /// </summary>
        private void Display_BonussRate()
        {
            string sQry = string.Empty;
            int iCol = 0;
            int iBNSMON = 0;
            string JOBGBN = string.Empty;
            string oCLTCOD = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordset2 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                // 소급일때 상여와 동일하게
                if (oJOBGBN == "2")
                {
                    JOBGBN = "1";
                }
                else
                {
                    JOBGBN = oJOBGBN;
                }

                oCLTCOD = "1";

                sQry = " SELECT U_BNSCAL, U_BNSMON, U_BNSRAT, U_AP1MON, U_AP2MON, U_AP3MON, U_AP4MON, U_AP5MON,";
                sQry = sQry + " U_AP6MON, U_AP7MON, U_AP8MON, U_AP1RAT, U_AP2RAT, U_AP3RAT, U_AP4RAT, U_AP5RAT, ";
                sQry = sQry + " U_AP6RAT, U_AP7RAT, U_AP8RAT, U_AP1AMT, U_AP2AMT, U_AP3AMT, U_AP4AMT, U_AP5AMT,";
                sQry = sQry + " U_AP6AMT, U_AP7AMT, U_AP8AMT, U_AP1GBN, U_AP2GBN, U_AP3GBN, U_AP4GBN, U_AP5GBN,";
                sQry = sQry + " U_AP6GBN, U_AP7GBN, U_AP8GBN  FROM [@PH_PY108A] ";
                sQry = sQry + " WHERE U_CLTCOD = '" + oCLTCOD + "'  AND U_JOBGBN = '" + JOBGBN + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.RecordCount == 0 | oJOBTYP == "1")
                {
                    oDS_PH_PY132A.SetValue("U_bBNSRAT", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bMONTH1", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bMONTH2", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bMONTH3", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bMONTH4", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bMONTH5", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bMONTH6", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bMONTH7", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bMONTH8", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPRAT1", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPRAT2", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPRAT3", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPRAT4", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPRAT5", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPRAT6", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPRAT7", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPRAT8", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPAMT1", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPAMT2", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPAMT3", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPAMT4", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPAMT5", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPAMT6", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPAMT7", 0, Convert.ToString(0));
                    oDS_PH_PY132A.SetValue("U_bAPPAMT8", 0, Convert.ToString(0));
                    for (iCol = 1; iCol <= 8; iCol++)
                    {
                        oDS_PH_PY132A.SetValue("U_AP" + iCol + "GBN", 0, "1");
                    }
                    // 2010.04.05 최동권 추가
                    oYM = oDS_PH_PY132A.GetValue("U_YM", 0).ToString().Trim();
                }
                else
                {
                   // oDS_PH_PY132A.SetValue("U_bBNSCAL", 0, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_PY132A.SetValue("U_bBNSRAT", 0, oRecordSet.Fields.Item(2).Value);
                    oDS_PH_PY132A.SetValue("U_bMONTH1", 0, oRecordSet.Fields.Item(3).Value);
                    oDS_PH_PY132A.SetValue("U_bMONTH2", 0, oRecordSet.Fields.Item(4).Value);
                    oDS_PH_PY132A.SetValue("U_bMONTH3", 0, oRecordSet.Fields.Item(5).Value);
                    oDS_PH_PY132A.SetValue("U_bMONTH4", 0, oRecordSet.Fields.Item(6).Value);
                    oDS_PH_PY132A.SetValue("U_bMONTH5", 0, oRecordSet.Fields.Item(7).Value);
                    oDS_PH_PY132A.SetValue("U_bMONTH6", 0, oRecordSet.Fields.Item(8).Value);
                    oDS_PH_PY132A.SetValue("U_bMONTH7", 0, oRecordSet.Fields.Item(9).Value);
                    oDS_PH_PY132A.SetValue("U_bMONTH8", 0, oRecordSet.Fields.Item(10).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPRAT1", 0, oRecordSet.Fields.Item(11).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPRAT2", 0, oRecordSet.Fields.Item(12).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPRAT3", 0, oRecordSet.Fields.Item(13).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPRAT4", 0, oRecordSet.Fields.Item(14).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPRAT5", 0, oRecordSet.Fields.Item(15).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPRAT6", 0, oRecordSet.Fields.Item(16).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPRAT7", 0, oRecordSet.Fields.Item(17).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPRAT8", 0, oRecordSet.Fields.Item(18).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPAMT1", 0, oRecordSet.Fields.Item(19).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPAMT2", 0, oRecordSet.Fields.Item(20).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPAMT3", 0, oRecordSet.Fields.Item(21).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPAMT4", 0, oRecordSet.Fields.Item(22).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPAMT5", 0, oRecordSet.Fields.Item(23).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPAMT6", 0, oRecordSet.Fields.Item(24).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPAMT7", 0, oRecordSet.Fields.Item(25).Value);
                    oDS_PH_PY132A.SetValue("U_bAPPAMT8", 0, oRecordSet.Fields.Item(26).Value);
                    for (iCol = 1; iCol <= 8; iCol++)
                    {
                        oDS_PH_PY132A.SetValue("U_AP" + iCol + "GBN", 0, oRecordSet.Fields.Item(26 + iCol).Value);
                    }
                    // 2010.04.05 최동권 추가
                    oYM = oDS_PH_PY132A.GetValue("U_YM", 0).ToString().Trim();
                    if (!string.IsNullOrEmpty(oYM))
                    {

                        iBNSMON = oRecordSet.Fields.Item(1).Value * -1;
                        if (iBNSMON != 0)
                            iBNSMON = iBNSMON + 1;
                        oRecordset2.DoQuery(("SELECT CONVERT(VARCHAR(6),DATEADD(MM, " + Convert.ToString(iBNSMON) + ", '" + oYM + "01'),112) FROM OADM"));
                    }
                }
                oForm.Update();

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Display_BonussRate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordset2); //메모리 해제
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
//	internal class PH_PY132
//	{
//////********************************************************************************
//////  File           : PH_PY132.cls
//////  Module         : 급여관리 > 성과급 차등 개인별 계산
//////  Desc           : 성과급 차등 개인별 계산
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY132A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY132B;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//////------------------------------------------------------
//			/// 대상인원수
//		private short G_TotCnt;
//			/// 계산인원수
//		private short G_PayCnt;
//			/// 잠금제외수
//		private short G_ChkCnt;


//			/// 귀속연월
//		private string oYM;
//			/// 지급종류
//		private string oJOBTYP;
//			/// 지급구분
//		private string oJOBGBN;
//			/// 지급일자
//		private string oJIGBIL;
//			/// 지급대상구분
//		private string oJOBTRG;
//			/// 사업장
//		private string oCLTCOD;
//			/// 부서시작코드
//		private string oSTRDPT;
//			/// 부서종료코드
//		private string oENDDPT;
//			/// 사원번호
//		private string oMSTCOD;
//			/// 연말정산포함
//		private string oJSNCHK;
//			/// 연차지급포함
//		private string oYCHCHK;

//			/// 상여계산방법
//		private string oBNSCAL;
//			/// 상여율
//		private string oBNSRAT;
//			/// 세액대상기간시작(급여)
//		private string oSTRTAX;
//			/// 세액대상기간종료(급여)
//		private string oENDTAX;
//			/// 세액대상기간시작(상여)
//		private string oSTRBNS;
//			/// 세액대상기간종료(상여)
//		private string oENDBNS;
//			/// 상여계산기준일
//		private string oGNEDAT;
//			/// 퇴사자제외일
//		private string oEXPDAT;
//			/// 상여퇴직임금에포함
//		private string oRETCHK;

//		private string StrDate;
//		private string EndDate;
//		private int MaxRow;
//		private short JSNYER;

//			/// 소득세정산
//		private bool G06_CHK;
//			/// 주민세정산
//		private bool G07_CHK;
//			/// 건강보험정산
//		private bool G08_CHK;
//			/// 국민연금정산
//		private bool G90_CHK;
//			/// 농특세정산
//		private bool G91_CHK;
//			/// 고용보험정산
//		private bool G92_CHK;
//		private string G04_BNSUSE;
//		private string PAY_001;
//		private string PAY_007;

//		private struct WG03TILR
//		{
//			[VBFixedArray(24)]
//			public string[] CSUCOD;
//			[VBFixedArray(24)]
//			public string[] CSUNAM;
//			[VBFixedArray(24)]
//				/// 월정급여
//			public string[] MPYGBN;
//			[VBFixedArray(24)]
//				/// 수당한도금액
//			public double[] CSUKUM;
//			[VBFixedArray(24)]
//				/// 과세구분
//			public string[] GWATYP;
//			[VBFixedArray(24)]
//				/// 고용보험여부
//			public string[] GBHGBN;
//			[VBFixedArray(24)]
//				/// 사사오입구분(끝전처리)
//			public string[] ROUNDT;
//			[VBFixedArray(24)]
//				/// 끝전처리자릿수
//			public short[] RODLEN;
//			[VBFixedArray(24)]
//				/// 급여수식
//			public string[] GONSIL;
//			[VBFixedArray(24)]
//				/// 상여항목
//			public string[] BNSUSE;
//			[VBFixedArray(24)]
//				/// 비과세코드
//			public string[] BTXCOD;

//			//UPGRADE_TODO: 해당 구조체의 인스턴스를 초기화하려면 "Initialize"를 호출해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
//			public void Initialize()
//			{
//				//UPGRADE_WARNING: CSUCOD 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				CSUCOD = new string[25];
//				//UPGRADE_WARNING: CSUNAM 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				CSUNAM = new string[25];
//				//UPGRADE_WARNING: MPYGBN 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				MPYGBN = new string[25];
//				//UPGRADE_WARNING: CSUKUM 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				CSUKUM = new double[25];
//				//UPGRADE_WARNING: GWATYP 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				GWATYP = new string[25];
//				//UPGRADE_WARNING: GBHGBN 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				GBHGBN = new string[25];
//				//UPGRADE_WARNING: ROUNDT 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				ROUNDT = new string[25];
//				//UPGRADE_WARNING: RODLEN 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				RODLEN = new short[25];
//				//UPGRADE_WARNING: GONSIL 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				GONSIL = new string[25];
//				//UPGRADE_WARNING: BNSUSE 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				BNSUSE = new string[25];
//				//UPGRADE_WARNING: BTXCOD 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				BTXCOD = new string[25];
//			}
//		}
////UPGRADE_WARNING: WK_C 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//		WG03TILR WK_C;

//		private struct WG04TILR
//		{
//			[VBFixedArray(18)]
//			public string[] GONCOD;
//			[VBFixedArray(18)]
//			public string[] GONNAM;
//			[VBFixedArray(18)]
//				/// 상여
//			public string[] BNSUSE;
//			[VBFixedArray(18)]
//				/// 계산식
//			public string[] GONSIL;
//			[VBFixedArray(18)]
//			public string[] ROUNDT;
//			[VBFixedArray(18)]
//			public short[] RODLEN;

//			//UPGRADE_TODO: 해당 구조체의 인스턴스를 초기화하려면 "Initialize"를 호출해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
//			public void Initialize()
//			{
//				//UPGRADE_WARNING: GONCOD 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				GONCOD = new string[19];
//				//UPGRADE_WARNING: GONNAM 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				GONNAM = new string[19];
//				//UPGRADE_WARNING: BNSUSE 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				BNSUSE = new string[19];
//				//UPGRADE_WARNING: GONSIL 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				GONSIL = new string[19];
//				//UPGRADE_WARNING: ROUNDT 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				ROUNDT = new string[19];
//				//UPGRADE_WARNING: RODLEN 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				RODLEN = new short[19];
//			}
//		}
////UPGRADE_WARNING: WK_G 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//		WG04TILR WK_G;

//		private string U_CSUCOD;
//		private string U_GONCOD;
//			//시    급
//		private double WK_TIMAMT;
//			//일    급
//		private double WK_DAYAMT;
//			//월    급
//		private double WK_STDAMT;
//			//상여기본
//		private double WK_BNSAMT;
//			//적용상여금
//		private double WK_APPBNS;
//			//급여기본등록의 평균임금
//		private double WK_AVRAMT;

//		private struct WG33PAYR
//		{
//				//문서번호
//			public short DocNum;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//			public char[] U_MSTCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(50), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 50)]
//			public char[] U_MSTNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//			public char[] U_EmpID;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//			public char[] U_MSTBRK;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//			public char[] U_CLTCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//부서
//			public char[] U_MSTDPT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//				//직책
//			public char[] U_MSTSTP;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//			public char[] U_CLTNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//			public char[] U_BRKNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//부서
//			public char[] U_DPTNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 20)]
//				//직책
//			public char[] U_STPNAM;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//			public char[] U_PAYTYP;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//급여지급일구분
//			public char[] U_JOBTRG;
//			public double U_MEDAMT;
//			public double U_KUKAMT;
//			public double U_GBHAMT;
//			public string U_PERNBR;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//			public char[] U_JIGCOD;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(10), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 10)]
//			public char[] U_HOBONG;
//				//기 본 급
//			public double U_STDAMT;
//				//통상일급
//			public double U_BASAMT;
//				//기본일급
//			public double U_DAYAMT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//입사일자
//			public char[] U_INPDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//수습만료일
//			public char[] U_INEDAT;
//			//UPGRADE_WARNING: 고정 길이 문자열 크기가 버퍼와 맞아야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
//			[VBFixedString(8), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst = 8)]
//				//퇴사일자
//			public char[] U_OUTDAT;
//			public short U_TAXCNT;
//			public short U_CHLCNT;
//			public string U_NJCGBN;
//			public double U_MEDFRG;
//			[VBFixedArray(24)]
//			public double[] U_CSUAMT;
//			public double U_GWASEE;
//			public double U_BTAX01;
//			public double U_BTAX02;
//			public double U_BTAX03;
//			public double U_BTAX04;
//			public double U_BTAX05;
//			public double U_BTAX06;
//			public double U_BTAX07;
//			public double U_BTXG01;
//			public double U_BTXH01;
//			public double U_BTXH05;
//			public double U_BTXH06;
//			public double U_BTXH07;
//			public double U_BTXH08;
//			public double U_BTXH09;
//			public double U_BTXH10;
//			public double U_BTXH11;
//			public double U_BTXH12;
//			public double U_BTXH13;
//			public double U_BTXI01;
//			public double U_BTXK01;
//			public double U_BTXM01;
//			public double U_BTXM02;
//			public double U_BTXM03;
//			public double U_BTXO01;
//			public double U_BTXQ01;
//			public double U_BTXR10;
//			public double U_BTXS01;
//			public double U_BTXT01;
//			public double U_BTXY01;
//			public double U_BTXY02;
//			public double U_BTXY03;
//			public double U_BTXY21;
//			public double U_BTXZ01;
//			public double U_BTXY22;
//			public double U_BTXX01;
//			public double U_BTXY20;
//			public double U_BTXTOT;
//			public double U_TOTPAY;
//			[VBFixedArray(18)]
//			public double[] U_GONAMT;
//			public double U_TOTGON;
//			public double U_SILJIG;
//			//// 상여금
//			public double U_AVRPAY;
//			public double U_NABTAX;
//			public double U_BNSRAT;
//			public double U_APPRAT;
//			public short U_GNSYER;
//			public short U_GNSMON;
//			public short U_TAXTRM;
//			public double U_BONUSS;

//			//UPGRADE_TODO: 해당 구조체의 인스턴스를 초기화하려면 "Initialize"를 호출해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
//			public void Initialize()
//			{
//				//UPGRADE_WARNING: U_CSUAMT 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				U_CSUAMT = new double[25];
//				//UPGRADE_WARNING: U_GONAMT 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				U_GONAMT = new double[19];
//			}
//		}
////UPGRADE_WARNING: WG03 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//		WG33PAYR WG03;

//		private struct WG01CODR
//		{
//			[VBFixedArray(10)]
//			public double[] TB1AMT;
//			[VBFixedArray(10)]
//			public double[] TB1GON;
//			[VBFixedArray(10)]
//			public double[] TB1RAT;
//			[VBFixedArray(10)]
//			public double[] TB1KUM;

//			//UPGRADE_TODO: 해당 구조체의 인스턴스를 초기화하려면 "Initialize"를 호출해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
//			public void Initialize()
//			{
//				//UPGRADE_WARNING: TB1AMT 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				TB1AMT = new double[11];
//				//UPGRADE_WARNING: TB1GON 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				TB1GON = new double[11];
//				//UPGRADE_WARNING: TB1RAT 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				TB1RAT = new double[11];
//				//UPGRADE_WARNING: TB1KUM 배열의 하한이 1에서 0(으)로 변경되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
//				TB1KUM = new double[11];
//			}
//		}
////UPGRADE_WARNING: WG01 구조체의 배열은 사용하기 전에 초기화해야 합니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
//		WG01CODR WG01;
//			/// 사용하는 국외비과세코드
//		private string TB1_BT3COD;
//			/// 사용하는 연구비과세코드
//		private string TB1_BT5COD;


//			///X01
//		private double X01_Val;
//			///X02
//		private double X02_Val;
//			///X03
//		private double X03_Val;
//			///X04
//		private double X04_Val;

//		private short X10_Val;
//		private short X11_Val;
//		private short X12_Val;
//		private short X13_Val;
//		private string X14_Val;
//		private double X15_Val;
//		private short X16_Val;
//		private double X17_Val;
//		private short X18_Val;
//		private short X19_Val;
//		private short X20_Val;

//		private string REMARK1;
//		private string REMARK2;
//		private string REMARK3;
//		private bool TermCHK;
//		private SAPbobsCOM.Recordset sRecordset;

//			////저장전 문서번호 저장
//		private int tDocEntry;
//			////계산 여부
//		private bool CalcYN;


//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY132.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY132_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY132");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			oForm.DataBrowser.BrowseBy = "Code";


//			oForm.Freeze(true);
//			PH_PY132_CreateItems();
//			PH_PY132_EnableMenus();
//			PH_PY132_SetDocument(oFromDocEntry01);
//			//    Call PH_PY132_FormResize

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

//		private bool PH_PY132_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;
//			short iCol = 0;
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

//			////----------------------------------------------------------------------------------------------
//			//// 데이터셋정의
//			////----------------------------------------------------------------------------------------------
//			oDS_PH_PY132A = oForm.DataSources.DBDataSources("@PH_PY132A");
//			////헤더
//			oDS_PH_PY132B = oForm.DataSources.DBDataSources("@PH_PY132B");
//			////라인

//			oMat1 = oForm.Items.Item("Mat1").Specific;
//			//

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//			oMat1.AutoResizeColumns();
//			////----------------------------------------------------------------------------------------------
//			//// 아이템 설정
//			////----------------------------------------------------------------------------------------------
//			/// 사업장
//			//    Set oCombo = oForm.Items("CLTCOD").Specific
//			//    oForm.Items("CLTCOD").DisplayDesc = True


//			/// 귀속연월
//			oDS_PH_PY132A.SetValue("U_YM", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM"));

//			/// 지급구분
//			oCombo = oForm.Items.Item("JOBGBN").Specific;
//			sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P212' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("JOBGBN").DisplayDesc = true;
//			oCombo.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);


//			//    Call oForm.DataSources.UserDataSources.Add("Col0", dt_SHORT_TEXT)
//			//    Call oForm.DataSources.UserDataSources.Add("Col1", dt_SHORT_TEXT)
//			//    Set oColumn = oMat1.Columns("Col0")
//			//    oColumn.DataBind.SetBound True, "", "Col0"
//			//    Set oColumn = oMat1.Columns("Col1")
//			//    oColumn.DataBind.SetBound True, "", "Col1"


//			//// 상여
//			for (iCol = 1; iCol <= 8; iCol++) {
//				oCombo = oForm.Items.Item("AP" + iCol + "GBN").Specific;
//				oCombo.ValidValues.Add("1", "개월 이상");
//				oCombo.ValidValues.Add("2", "개월 미만");
//				oCombo.ValidValues.Add("3", "일수 이상");
//				oCombo.ValidValues.Add("4", "일수 미만");
//				if (oCombo.ValidValues.Count > 0) {
//					oDS_PH_PY132A.SetValue("U_AP" + iCol + "GBN", 0, "1");
//				}
//				oForm.Items.Item("AP" + iCol + "GBN").DisplayDesc = true;
//			}

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
//			PH_PY132_CreateItems_Error:

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
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY132_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY132_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1287", true);
//			////복제

//			return;
//			PH_PY132_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY132_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY132_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY132_FormItemEnabled();
//				PH_PY132_AddMatrixRow();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY132_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY132_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY132_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY132_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;

//			 // ERROR: Not supported in C#: OnErrorStatement



//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				oForm.ActiveItem = "YM";
//				oForm.Items.Item("Code").Enabled = false;
//				oForm.Items.Item("Btn1").Visible = true;

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가


//				/// 귀속연월
//				oDS_PH_PY132A.SetValue("U_YM", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM"));
//				/// 지급구분
//				oCombo = oForm.Items.Item("JOBGBN").Specific;
//				oCombo.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);


//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				oForm.Items.Item("Code").Enabled = true;
//				oForm.Items.Item("Btn1").Visible = false;
//				oForm.ActiveItem = "Code";

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//        oForm.Mode = fm_UPDATE_MODE
//				oForm.ActiveItem = "YM";
//				oForm.Items.Item("Code").Enabled = false;
//				oForm.Items.Item("Btn1").Visible = true;

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY132_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY132_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			string JIGBIL = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;


//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1
//					if (pval.BeforeAction == true) {
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								/// 유효성 검사
//								if (PH_PY132_DataValidCheck() == false) {
//									BubbleEvent = false;
//								}
//							}
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							tDocEntry = oForm.Items.Item("Code").Specific.VALUE;
//						/// 급(상)여계산
//						} else if (pval.ItemUID == "Btn1") {
//							CalcYN = true;
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							tDocEntry = oForm.Items.Item("Code").Specific.VALUE;
//							/// 유효성 검사
//							if (PH_PY132_DataValidCheck() == true & Pay_Calc() == true) {
//								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//									oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//									MDC_Globals.Sbo_Application.StatusBar.SetText("성과급 차등계산이 진행중입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
//									//UPGRADE_WARNING: oForm.Items(JOBGBN).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									if (oForm.Items.Item("JOBGBN").Specific.VALUE != "2") {
//										////정상계산
//										oRecordSet.DoQuery("EXEC PH_PY132 '" + tDocEntry + "'");
//									} else {
//										////소급계산
//										oRecordSet.DoQuery("EXEC PH_PY132_SOGUB '" + tDocEntry + "'");
//									}
//									MDC_Globals.Sbo_Application.StatusBar.SetText("성과급 차등계산이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//								}
//							} else {
//								BubbleEvent = false;
//							}
//						}
//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemUID == "1") {
//							if (pval.ActionSuccess == true) {
//								if (CalcYN == true) {
//									CalcYN = false;
//									if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//										//                                Sbo_Application.StatusBar.SetText "성과급 차등계산이 진행중입니다.", bmt_Short, smt_Warning
//										//                                If oForm.Items("JOBGBN").Specific.VALUE <> "2" Then
//										//                                '//정상계산
//										//                                    oRecordSet.DoQuery "EXEC PH_PY132 '" & tDocEntry & "'"
//										//                                Else
//										//                                    '//소급계산
//										//                                    oRecordSet.DoQuery "EXEC PH_PY132_SOGUB '" & tDocEntry & "'"
//										//                                End If

//										MDC_Globals.Sbo_Application.StatusBar.SetText("성과급 차등계산이 완료되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//										oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//										PH_PY132_FormItemEnabled();
//										oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//									}
//								} else if (CalcYN == false) {
//									if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//										oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//										PH_PY132_FormItemEnabled();
//										oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//									} else {
//										PH_PY132_FormItemEnabled();

//									}

//								}
//							}
//						}


//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					if (pval.BeforeAction == true & pval.ItemUID == "YM" & pval.CharPressed == 9 & pval.FormMode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//						//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//						if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item(pval.ItemUID).Specific.String))) {
//							MDC_Globals.Sbo_Application.StatusBar.SetText("귀속연월은 필수입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							BubbleEvent = false;
//						}
//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat1":
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
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						////지급구분
//						if (pval.ItemUID == "JOBGBN") {
//							//UPGRADE_WARNING: oForm.Items(JOBGBN).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if ((oForm.Items.Item("JOBGBN").Specific.Selected != null)) {
//								oYM = Strings.Trim(oDS_PH_PY132A.GetValue("U_YM", 0));
//								JIGBIL = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd");
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oForm.Items.Item("JIGBIL").Specific.VALUE = JIGBIL;
//								oJOBGBN = Strings.Trim(oDS_PH_PY132A.GetValue("U_JOBGBN", 0));
//							} else {
//								oJOBGBN = "";
//							}
//							if (!string.IsNullOrEmpty(oJOBGBN))
//								Display_BonussRate();

//						}
//						////급여지급대상일


//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//					////7
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
//					if (pval.BeforeAction == false & pval.ItemChanged == true) {
//						////귀속년월
//						if (pval.ItemUID == "YM") {
//							//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("YM").Specific.VALUE))) {
//								oDS_PH_PY132A.SetValue("U_YM", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM"));
//							} else {
//								//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								oDS_PH_PY132A.SetValue("U_YM", 0, oForm.Items.Item("YM").Specific.String);
//							}
//							oYM = Strings.Trim(oDS_PH_PY132A.GetValue("U_YM", 0));

//							if (!string.IsNullOrEmpty(oYM))
//								Display_BonussRate();
//						}

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						oMat1.LoadFromDataSource();

//						PH_PY132_FormItemEnabled();
//						PH_PY132_AddMatrixRow();
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
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					//컬렉션에서 삭제및 모든 메모리 제거
//					//ㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡㅡ
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY132A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY132A = null;
//						//
//						//                Set oMat1 = Nothing
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
//					break;
//				//            If pval.BeforeAction = True Then
//				//
//				//            ElseIf pval.BeforeAction = False Then
//				//
//				//            End If
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
//					break;
//				//            If pval.BeforeAction = True Then
//				//
//				//            ElseIf pval.Before_Action = False Then
//				//
//				//            End If
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
//					//Call AuthorityCheck(oForm, "CLTCOD", "@PH_PY132A", "DocEntry")      '//접속자 권한에 따른 사업장 보기
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY132_FormItemEnabled();
//						PH_PY132_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY132_FormItemEnabled();
//						PH_PY132_AddMatrixRow();
//						break;
//					//                oForm.Items("Code").CLICK ct_Regular
//					case "1282":
//						////문서추가
//						PH_PY132_FormItemEnabled();
//						PH_PY132_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY132_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						break;
//					//                '// [MAT1 용]
//					//                 If oMat1.RowCount <> oMat1.VisualRowCount Then
//					//                    oMat1.FlushToDataSource
//					//
//					//                    While (i <= oDS_PH_PY132B.Size - 1)
//					//                        If oDS_PH_PY132B.GetValue("U_FILD01", i) = "" Then
//					//                            oDS_PH_PY132B.RemoveRecord (i)
//					//                            i = 0
//					//                        Else
//					//                            i = i + 1
//					//                        End If
//					//                    Wend
//					//
//					//                    For i = 0 To oDS_PH_PY132B.Size
//					//                        Call oDS_PH_PY132B.setValue("U_LineNum", i, i + 1)
//					//                    Next i
//					//
//					//                    oMat1.LoadFromDataSource
//					//                End If
//					//                Call PH_PY132_AddMatrixRow
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
//				case "Mat1":
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

//		public void PH_PY132_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			//    '//[Mat1 용]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY132B.GetValue("U_CLTCOD", oRow - 1)))) {
//					if (oDS_PH_PY132B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY132B.InsertRecord((oRow));
//					}
//					oDS_PH_PY132B.Offset = oRow;
//					oDS_PH_PY132B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					//            oDS_PH_PY132B.setValue "U_CLTCOD", oRow, ""
//					//            oDS_PH_PY132B.setValue "U_TeamCode", oRow, ""
//					//            oDS_PH_PY132B.setValue "U_TeamName", oRow, ""

//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY132B.Offset = oRow - 1;
//					oDS_PH_PY132B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					//            oDS_PH_PY132B.setValue "U_CLTCOD", oRow - 1, ""
//					//            oDS_PH_PY132B.setValue "U_TeamCode", oRow - 1, ""
//					//            oDS_PH_PY132B.setValue "U_FILD03", oRow - 1, 0
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY132B.Offset = oRow;
//				oDS_PH_PY132B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				//        oDS_PH_PY132B.setValue "U_CLTCOD", oRow, ""
//				//        oDS_PH_PY132B.setValue "U_TeamCode", oRow, ""
//				//        oDS_PH_PY132B.setValue "U_FILD03", oRow, 0
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY132_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY132_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY132_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY132'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY132_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY132_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY132_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			short ErrNum = 0;
//			string sQry = null;
//			string tCode = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			ErrNum = 0;

//			functionReturnValue = false;

//			///
//			oYM = Strings.Trim(oDS_PH_PY132A.GetValue("U_YM", 0));
//			oJOBGBN = Strings.Trim(oDS_PH_PY132A.GetValue("U_JOBGBN", 0));
//			oBNSRAT = oDS_PH_PY132A.GetValue("U_bBNSRAT", 0);
//			oJIGBIL = oDS_PH_PY132A.GetValue("U_JIGBIL", 0);

//			/// Check
//			switch (true) {
//				case MDC_SetMod.ChkYearMonth(ref oYM) == false:
//					MDC_Globals.Sbo_Application.StatusBar.SetText("귀속 연월을 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					return functionReturnValue;
//				case string.IsNullOrEmpty(oJOBGBN):
//					MDC_Globals.Sbo_Application.StatusBar.SetText("지급 구분을 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					return functionReturnValue;
//				case string.IsNullOrEmpty(oJIGBIL):
//					MDC_Globals.Sbo_Application.StatusBar.SetText("지급일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					return functionReturnValue;
//			}

//			/// 상여계산시 체크
//			if (oJOBTYP != "1") {
//				switch (true) {
//					case Conversion.Val(oBNSRAT) == 0:
//						MDC_Globals.Sbo_Application.StatusBar.SetText("상여율은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//						return functionReturnValue;
//				}
//			}


//			tCode = Strings.Trim(oDS_PH_PY132A.GetValue("U_YM", 0)) + Strings.Trim(oDS_PH_PY132A.GetValue("U_JOBGBN", 0)) + Strings.Trim(oDS_PH_PY132A.GetValue("U_Number", 0));

//			////코드생성
//			////코드 중복 체크
//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				sQry = "SELECT CODE FROM [@PH_PY132A] WHERE CODE = '" + tCode + "'";
//				oRecordSet.DoQuery(sQry);
//				if (oRecordSet.RecordCount > 0) {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("코드가 존재합니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//					functionReturnValue = false;
//					return functionReturnValue;
//				} else {
//					oDS_PH_PY132A.SetValue("Code", 0, tCode);
//					oDS_PH_PY132A.SetValue("Name", 0, tCode);
//				}
//			}


//			functionReturnValue = true;
//			return functionReturnValue;
//			PH_PY132_DataValidCheck_Error:

//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY132_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		public bool PH_PY132_Validate(string ValidateType)
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
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY132A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY132A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY132_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY132_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY132_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY132_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		private bool Pay_Calc()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			short ErrNum = 0;
//			string Temp1 = null;
//			string Temp2 = null;
//			int i = 0;

//			string YM = null;
//			string JOBGBN = null;
//			string Number = null;
//			string JIGBIL = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			functionReturnValue = false;


//			YM = Strings.Trim(oDS_PH_PY132A.GetValue("U_YM", 0));
//			JOBGBN = Strings.Trim(oDS_PH_PY132A.GetValue("U_JOBGBN", 0));
//			Number = Strings.Trim(oDS_PH_PY132A.GetValue("U_Number", 0));
//			JIGBIL = Strings.Trim(oDS_PH_PY132A.GetValue("U_JIGBIL", 0));

//			sQry = "Select Count(*) From [@PH_PY132A] ";
//			sQry = sQry + " Where U_YM = '" + YM + "'";
//			sQry = sQry + " AND U_Number = '" + Number + "'";
//			sQry = sQry + " AND U_JIGBIL = '" + JIGBIL + "'";

//			//sQry = "Select U_CLTCOD, U_YM, U_JOBTYP, U_JOBGBN, U_JOBTRG From [@PH_PY112A] Where U_CLTCOD = '" & Trim(oDS_PH_PY132A.GetValue("U_CLTCOD", 0)) & "' AND U_YM = '" & Trim(oDS_PH_PY132A.GetValue("U_YM", 0)) & "' AND U_JOBTYP = '" & Trim(oDS_PH_PY132A.GetValue("U_JOBTYP", 0)) & "' "
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.Fields.Item(0).Value > 0) {
//				if (MDC_Globals.Sbo_Application.MessageBox("기존에 급여계산 결과가 있습니다. 계속 진행하시겠습니까?", 2, "Yes", "No") == 2) {
//					functionReturnValue = false;
//					//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//					oRecordSet = null;
//					return functionReturnValue;
//				}
//				functionReturnValue = true;
//			} else {
//				functionReturnValue = true;
//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			Error_Message:
//			//    Temp1 = Trim(oDS_PH_PY132A.GetValue("U_CLTCOD", 0)) & Trim(oDS_PH_PY132A.GetValue("U_YM", 0)) & Trim(oDS_PH_PY132A.GetValue("U_JOBTYP", 0)) & _
//			//'            Trim(oDS_PH_PY132A.GetValue("U_JOBGBN", 0)) & Trim(oDS_PH_PY132A.GetValue("U_JOBTRG", 0))
//			//
//			//    If oRecordSet.RecordCount > 0 Then
//			//        Do Until oRecordSet.EOF
//			//            Temp2 = ""
//			//            For i = 0 To oRecordSet.Fields.Count - 1
//			//                Temp2 = Temp2 & Trim(oRecordSet.Fields(i).VALUE)
//			//            Next
//			//
//			//            If Temp1 = Temp2 Then
//			//                GoTo Continue
//			//            End If
//			//            oRecordSet.MoveNext
//			//        Loop
//			//    End If
//			//    GoTo Continue2
//			//
//			//Continue:
//			//    If Sbo_Application.MessageBox("기존에 급여계산 결과가 있습니다. 계속 진행하시겠습니까?", 2, "Yes", "No") = 2 Then
//			//        Pay_Calc = False
//			//        Exit Function
//			//    End If
//			//
//			//
//			//Continue2:
//			//    Pay_Calc = True
//			//    Set oRecordSet = Nothing
//			//


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			return functionReturnValue;
//		}

//		private void Display_BonussRate()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			SAPbobsCOM.Recordset oRecordSet = null;
//			SAPbobsCOM.Recordset oRecordset2 = null;
//			string sQry = null;
//			short ErrNum = 0;
//			short iCol = 0;
//			int iBNSMON = 0;

//			ErrNum = 0;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//			oRecordset2 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			string JOBGBN = null;

//			////소급일때 상여와 동일하게
//			if (oJOBGBN == "2") {
//				JOBGBN = "1";
//			} else {
//				JOBGBN = oJOBGBN;
//			}

//			oCLTCOD = "1";

//			sQry = " SELECT U_BNSCAL, U_BNSMON, U_BNSRAT, U_AP1MON, U_AP2MON, U_AP3MON, U_AP4MON, U_AP5MON,";
//			sQry = sQry + " U_AP6MON, U_AP7MON, U_AP8MON, U_AP1RAT, U_AP2RAT, U_AP3RAT, U_AP4RAT, U_AP5RAT, ";
//			sQry = sQry + " U_AP6RAT, U_AP7RAT, U_AP8RAT, U_AP1AMT, U_AP2AMT, U_AP3AMT, U_AP4AMT, U_AP5AMT,";
//			sQry = sQry + " U_AP6AMT, U_AP7AMT, U_AP8AMT, U_AP1GBN, U_AP2GBN, U_AP3GBN, U_AP4GBN, U_AP5GBN,";
//			sQry = sQry + " U_AP6GBN, U_AP7GBN, U_AP8GBN  FROM [@PH_PY108A] ";
//			sQry = sQry + " WHERE U_CLTCOD = '" + oCLTCOD + "'  AND U_JOBGBN = '" + JOBGBN + "'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0 | oJOBTYP == "1") {
//				oDS_PH_PY132A.SetValue("U_bBNSRAT", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bMONTH1", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bMONTH2", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bMONTH3", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bMONTH4", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bMONTH5", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bMONTH6", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bMONTH7", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bMONTH8", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPRAT1", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPRAT2", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPRAT3", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPRAT4", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPRAT5", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPRAT6", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPRAT7", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPRAT8", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPAMT1", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPAMT2", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPAMT3", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPAMT4", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPAMT5", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPAMT6", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPAMT7", 0, Convert.ToString(0));
//				oDS_PH_PY132A.SetValue("U_bAPPAMT8", 0, Convert.ToString(0));
//				for (iCol = 1; iCol <= 8; iCol++) {
//					oDS_PH_PY132A.SetValue("U_AP" + iCol + "GBN", 0, "1");
//				}

//				//// 2010.04.05 최동권 추가
//				oYM = Strings.Trim(oDS_PH_PY132A.GetValue("U_YM", 0));


//			} else {
//				oDS_PH_PY132A.SetValue("U_bBNSCAL", 0, oRecordSet.Fields.Item(0).Value);
//				oDS_PH_PY132A.SetValue("U_bBNSRAT", 0, oRecordSet.Fields.Item(2).Value);
//				oDS_PH_PY132A.SetValue("U_bMONTH1", 0, oRecordSet.Fields.Item(3).Value);
//				oDS_PH_PY132A.SetValue("U_bMONTH2", 0, oRecordSet.Fields.Item(4).Value);
//				oDS_PH_PY132A.SetValue("U_bMONTH3", 0, oRecordSet.Fields.Item(5).Value);
//				oDS_PH_PY132A.SetValue("U_bMONTH4", 0, oRecordSet.Fields.Item(6).Value);
//				oDS_PH_PY132A.SetValue("U_bMONTH5", 0, oRecordSet.Fields.Item(7).Value);
//				oDS_PH_PY132A.SetValue("U_bMONTH6", 0, oRecordSet.Fields.Item(8).Value);
//				oDS_PH_PY132A.SetValue("U_bMONTH7", 0, oRecordSet.Fields.Item(9).Value);
//				oDS_PH_PY132A.SetValue("U_bMONTH8", 0, oRecordSet.Fields.Item(10).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPRAT1", 0, oRecordSet.Fields.Item(11).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPRAT2", 0, oRecordSet.Fields.Item(12).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPRAT3", 0, oRecordSet.Fields.Item(13).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPRAT4", 0, oRecordSet.Fields.Item(14).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPRAT5", 0, oRecordSet.Fields.Item(15).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPRAT6", 0, oRecordSet.Fields.Item(16).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPRAT7", 0, oRecordSet.Fields.Item(17).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPRAT8", 0, oRecordSet.Fields.Item(18).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPAMT1", 0, oRecordSet.Fields.Item(19).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPAMT2", 0, oRecordSet.Fields.Item(20).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPAMT3", 0, oRecordSet.Fields.Item(21).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPAMT4", 0, oRecordSet.Fields.Item(22).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPAMT5", 0, oRecordSet.Fields.Item(23).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPAMT6", 0, oRecordSet.Fields.Item(24).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPAMT7", 0, oRecordSet.Fields.Item(25).Value);
//				oDS_PH_PY132A.SetValue("U_bAPPAMT8", 0, oRecordSet.Fields.Item(26).Value);
//				for (iCol = 1; iCol <= 8; iCol++) {
//					//            oForm.DataSources.UserDataSources("AP" & iCol & "GBN").ValueEx = oRecordSet.Fields(26 + iCol).Value
//					oDS_PH_PY132A.SetValue("U_AP" + iCol + "GBN", 0, oRecordSet.Fields.Item(26 + iCol).Value);
//				}

//				//// 2010.04.05 최동권 추가
//				oYM = Strings.Trim(oDS_PH_PY132A.GetValue("U_YM", 0));
//				if (!string.IsNullOrEmpty(oYM)) {

//					iBNSMON = oRecordSet.Fields.Item(1).Value * -1;
//					if (iBNSMON != 0)
//						iBNSMON = iBNSMON + 1;
//					oRecordset2.DoQuery(("SELECT CONVERT(VARCHAR(6),DATEADD(MM, " + Convert.ToString(iBNSMON) + ", '" + oYM + "01'),112) FROM OADM"));


//				}
//			}

//			oForm.Update();

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			Error_Message:
//			///////////////////////////////////////////////////////////////////////////////////////////////////
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//		}
//	}
//}
