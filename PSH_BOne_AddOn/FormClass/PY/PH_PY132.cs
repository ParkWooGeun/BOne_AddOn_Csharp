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
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
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

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY132_CreateItems();
                PH_PY132_EnableMenus();
                PH_PY132_SetDocument(oFormDocEntry01);
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
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
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
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("Code").Enabled = true;
                    oForm.Items.Item("Btn1").Visible = false;
                    oForm.ActiveItem = "Code";

                    oForm.EnableMenu("1281", false);                   // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
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
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY132_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFormDocEntry01)))
                {
                    PH_PY132_FormItemEnabled();
                    PH_PY132_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY132_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
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
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                        tDocEntry = oForm.Items.Item("Code").Specific.Value;
                    }
                    else if (pVal.ItemUID == "Btn1")  // 급(상)여계산
                    {
                        CalcYN = true;
                        tDocEntry = oForm.Items.Item("Code").Specific.Value;
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
                                if (oForm.Items.Item("JOBGBN").Specific.Value != "2")
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
                            oForm.Items.Item("JIGBIL").Specific.Value = JIGBIL;
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
            try
            {
                oForm.Freeze(true);
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
                            PH_PY132_FormItemEnabled();
                            PH_PY132_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY132_FormItemEnabled();
                            PH_PY132_AddMatrixRow();
                            break;
                        case "1282": //문서추가
                            PH_PY132_FormItemEnabled();
                            PH_PY132_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY132_FormItemEnabled();
                            break;
                        case "1293": //행삭제
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

                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY132A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
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
                        if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value.ToString().Trim()))
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

