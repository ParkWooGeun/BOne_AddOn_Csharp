using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 공정코드등록
    /// </summary>
    internal class PS_PP001 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PS_PP001H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP001L; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP001.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP001_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP001");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);

                PS_PP001_CreateItems();
                PS_PP001_EnableFormItem();
                PS_PP001_SetDocEntry();
                PS_PP001_AddMatrixRow(0, oMat01.RowCount);

                oForm.EnableMenu("1283", true); // 제거
                oForm.EnableMenu("1293", true); // 행삭제
                oForm.EnableMenu("1287", true); // 복제
                oForm.EnableMenu("1284", false);

                oForm.Items.Item("CpBCode").Click();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP001_CreateItems()
        {
            try
            {
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oDS_PS_PP001H = oForm.DataSources.DBDataSources.Item("@PS_PP001H");
                oDS_PS_PP001L = oForm.DataSources.DBDataSources.Item("@PS_PP001L");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각모드에따른 아이템설정
        /// </summary>
        private void PS_PP001_EnableFormItem()
        {
            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CpBCode").Enabled = true;
                    oForm.Items.Item("CpBName").Enabled = true;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CpBCode").Enabled = true;
                    oForm.Items.Item("CpBName").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CpBCode").Enabled = false;
                    oForm.Items.Item("CpBName").Enabled = false;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP001_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP001'", "");
                if (DocEntry == "0")
                {
                    oDS_PS_PP001H.SetValue("DocEntry", 0, "1");
                }
                else
                {
                    oDS_PS_PP001H.SetValue("DocEntry", 0, DocEntry);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 행 추가
        /// </summary>
        /// <param name="oSeq"></param>
        /// <param name="oRow"></param>
        private void PS_PP001_AddMatrixRow(short oSeq, int oRow)
        {
            try
            {
                switch (oSeq)
                {
                    case 0:
                        oDS_PS_PP001L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oMat01.LoadFromDataSource();
                        break;
                    case 1:
                        oDS_PS_PP001L.InsertRecord(oRow);
                        oDS_PS_PP001L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oMat01.LoadFromDataSource();
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 헤더 필수 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP001_CheckHeaderSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_PP001H.GetValue("U_CpBCode", 0)) || string.IsNullOrEmpty(oDS_PS_PP001H.GetValue("U_CpBName", 0)))
                {
                    errMessage = "대분류 또는 대분류명은 필수입력 사항입니다. 확인하세요.";
                    throw new Exception();
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 라인 필수 입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP001_CheckMatrixSpaceLine()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                oMat01.FlushToDataSource();

                if (oMat01.VisualRowCount <= 1)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하세요.";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount > 0)
                {
                    for (int i = 0; i <= oMat01.VisualRowCount - 2; i++)
                    {
                        oDS_PS_PP001L.Offset = i;

                        if (string.IsNullOrEmpty(oDS_PS_PP001L.GetValue("U_CpCode", i)))
                        {
                            errMessage = "공정코드 데이터는 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_PP001L.GetValue("U_CpName", i)))
                        {
                            errMessage = "공정명 데이터는 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                        else if (string.IsNullOrEmpty(oDS_PS_PP001L.GetValue("U_PartCode", i)))
                        {
                            errMessage = "소속코드는 필수입니다. 확인하세요.";
                            throw new Exception();
                        }
                    }

                    if (string.IsNullOrEmpty(oDS_PS_PP001L.GetValue("U_CpCode", oMat01.VisualRowCount - 1)))
                    {
                        oDS_PS_PP001L.RemoveRecord(oMat01.VisualRowCount - 1);
                    }
                }

                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// FlushToItemValue
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oCID"></param>
        /// <param name="oRow"></param>
        private void PS_PP001_FlushToItemValue(string oUID, string oCID, int oRow)
        {
            switch (oUID)
            {
                case "Mat01":
                    switch (oCID)
                    {
                        case "CpCode":
                            if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 2) && !string.IsNullOrEmpty(oMat01.Columns.Item("CpCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
                            {
                                oMat01.FlushToDataSource();
                                PS_PP001_AddMatrixRow(1, oMat01.RowCount);
                                oMat01.Columns.Item("CpCode").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            break;
                    }

                    oMat01.AutoResizeColumns();
                    break;
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
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP001_CheckHeaderSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (PS_PP001_CheckMatrixSpaceLine() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDS_PS_PP001H.SetValue("Code", 0, oDS_PS_PP001H.GetValue("U_CpBCode", 0).ToString().Trim());
                            oDS_PS_PP001H.SetValue("Name", 0, oDS_PS_PP001H.GetValue("U_CpBName", 0).ToString().Trim());
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == false)
                        {
                            PS_PP001_EnableFormItem();
                            PS_PP001_AddMatrixRow(1, oMat01.RowCount);
                        }
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
                    //공정코드
                    if (string.IsNullOrEmpty(oForm.Items.Item("CpBCode").Specific.Value))
                    {
                        if (pVal.ItemUID == "CpBCode" && pVal.CharPressed == 9)
                        {
                            oForm.Items.Item("CpBCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //소속코드
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "PartCode" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //작업반코드
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "WkClsCod" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //부서코드
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "DeptCode" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //단위
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "Unit" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    //대분류코드
                    if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItmBsort" && pVal.CharPressed == 9)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }

                    PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "CCCode")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CCCode");
                        }
                        else if (pVal.ColUID == "ActCode1")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ActCode1");
                        }
                        else if (pVal.ColUID == "ActCode2")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ActCode2");
                        }
                        else if (pVal.ColUID == "ActCode3")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ActCode3");
                        }
                        else if (pVal.ColUID == "ActCode4")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ActCode4");
                        }
                        else if (pVal.ColUID == "ActCode5")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ActCode5");
                        }
                        else if (pVal.ColUID == "ActCode6")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ActCode6");
                        }
                        else if (pVal.ColUID == "ActCode7")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ActCode7");
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oLastItemUID01 = pVal.ItemUID;
                                oLastColUID01 = pVal.ColUID;
                                oLastColRow01 = pVal.Row;
                            }
                            break;
                        default:
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = "";
                            oLastColRow01 = 0;
                            break;
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
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
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "CpCode")
                        {
                            PS_PP001_FlushToItemValue(pVal.ItemUID, pVal.ColUID, pVal.Row);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CpBCode" && pVal.ItemChanged == true)
                        {
                            sQry = "Select U_CdName From [@PS_SY001L] Where U_Minor = '" + oForm.Items.Item("CpBCode").Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oForm.Items.Item("CpBName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.AutoResizeColumns();
                        }

                        //팀명
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "DeptCode")
                        {
                            sQry = "  SELECT  T0.U_CodeNm";
                            sQry += " FROM    [@PS_HR200L] AS T0";
                            sQry += " WHERE   T0.Code = '1'";
                            sQry += "         AND T0.U_UseYN = 'Y'";
                            sQry += "         AND T0.U_Code = '" + oMat01.Columns.Item("DeptCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";

                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("DeptName").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //담당명
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "PartCode")
                        {
                            sQry = "  SELECT  T0.U_CodeNm";
                            sQry += " FROM    [@PS_HR200L] AS T0";
                            sQry += " WHERE   T0.Code = '2'";
                            sQry += "         AND T0.U_UseYN = 'Y'";
                            sQry += "         AND T0.U_Code = '" + oMat01.Columns.Item("PartCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            sQry += "         AND T0.U_Char1 = '" + oMat01.Columns.Item("DeptCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";

                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("PartName").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //반명
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "WkClsCod")
                        {

                            sQry = "  SELECT  T0.U_CodeNm";
                            sQry += " FROM    [@PS_HR200L] AS T0";
                            sQry += " WHERE   T0.Code = '9'";
                            sQry += "         AND T0.U_UseYN = 'Y'";
                            sQry += "         AND T0.U_Code = '" + oMat01.Columns.Item("WkClsCod").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            sQry += "         AND T0.U_Char1 = '" + oMat01.Columns.Item("PartCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            sQry += "         AND T0.U_Char2 = '" + oMat01.Columns.Item("DeptCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";

                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("WkClsNam").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //대분류명
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItmBsort")
                        {
                            sQry = "Select Name From [@PSH_ITMBSORT] Where Code = '" + oMat01.Columns.Item("ItmBsort").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ItmBname").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //코스트센터명
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "CCCode")
                        {
                            sQry = "Select PrcName From [OPRC] Where PrcCode = '" + oMat01.Columns.Item("CCCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("CCName").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //활동유형이름1
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ActCode1")
                        {
                            sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" + oMat01.Columns.Item("ActCode1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ActName1").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //활동유형이름2
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ActCode2")
                        {
                            sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" + oMat01.Columns.Item("ActCode2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ActName2").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //활동유형이름3
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ActCode3")
                        {
                            sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" + oMat01.Columns.Item("ActCode3").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ActName3").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //활동유형이름4
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ActCode4")
                        {
                            sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" + oMat01.Columns.Item("ActCode4").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ActName4").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //활동유형이름5
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ActCode5")
                        {
                            sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" + oMat01.Columns.Item("ActCode5").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ActName5").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //활동유형이름6
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ActCode6")
                        {
                            sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" + oMat01.Columns.Item("ActCode6").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ActName6").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        //활동유형이름7
                        if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ActCode7")
                        {
                            sQry = "Select U_ActName From [@PS_CO050L] Where U_ActCode = '" + oMat01.Columns.Item("ActCode7").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
                            oRecordSet01.DoQuery(sQry);
                            oMat01.Columns.Item("ActName7").Cells.Item(pVal.Row).Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_PP001_AddMatrixRow(1, oMat01.VisualRowCount);
                    oMat01.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP001H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP001L);
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
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat01.AutoResizeColumns();
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
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1293": //행삭제
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1281": //찾기
                            PS_PP001_EnableFormItem();
                            break;
                        case "1282": //추가
                            PS_PP001_EnableFormItem();
                            PS_PP001_SetDocEntry();
                            PS_PP001_AddMatrixRow(0, oMat01.RowCount);
                            oForm.Items.Item("CpBCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1287": //복제
                            oForm.Freeze(true);
                            oDS_PS_PP001H.SetValue("Code", 0, "");
                            oDS_PS_PP001H.SetValue("Name", 0, "");
                            oDS_PS_PP001H.SetValue("U_CpBCode", 0, "");
                            oDS_PS_PP001H.SetValue("U_CpBName", 0, "");

                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_PP001L.SetValue("Code", i, "");
                                oMat01.LoadFromDataSource();
                            }
                            oForm.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_PP001_EnableFormItem();
                            break;
                        case "1293": //행삭제
                            if (oMat01.RowCount != oMat01.VisualRowCount)
                            {
                                for (int i = 1; i <= oMat01.VisualRowCount; i++)
                                {
                                    oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                                }
                                oMat01.FlushToDataSource();
                                oDS_PS_PP001L.RemoveRecord(oDS_PS_PP001L.Size - 1);
                                oMat01.LoadFromDataSource();
                                if (oMat01.RowCount == 0)
                                {
                                    PS_PP001_AddMatrixRow(1, 0);
                                }
                                else
                                {
                                    if (!string.IsNullOrEmpty(oDS_PS_PP001L.GetValue("U_CpCode", oMat01.RowCount - 1).ToString().Trim()))
                                    {
                                        PS_PP001_AddMatrixRow(1, oMat01.RowCount);
                                    }
                                }
                            }
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
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
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
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
