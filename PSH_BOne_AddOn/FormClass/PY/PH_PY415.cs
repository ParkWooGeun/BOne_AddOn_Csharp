using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 소득정산계산
    /// </summary>
    internal class PH_PY415 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY415B;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY415.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY415_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY415");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
          //      oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                oForm.Visible = true;
                PH_PY415_CreateItems();
                PH_PY415_EnableMenus();
                PH_PY415_SetDocument(oFromDocEntry01);
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
        private void PH_PY415_CreateItems()
        {
            int i = 0;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY415B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                oMat1 = oForm.Items.Item("Mat01").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat1.AutoResizeColumns();

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("CLTCOD").Specific, "");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 년도
                oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY415_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY415_EnableMenus()
        {
            try
            {
                //oForm.EnableMenu("1283", true); //제거
                //oForm.EnableMenu("1284", false); //취소
                //oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY415_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY415_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY415_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY415_FormItemEnabled();

                    oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY415_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY415_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.Items.Item("Year").Specific.VALUE = Convert.ToString(DateTime.Now.Year - 1);

                    oForm.EnableMenu("1281", true); ////문서찾기
                    oForm.EnableMenu("1282", false); ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.EnableMenu("1281", false); ////문서찾기
                    oForm.EnableMenu("1282", true); ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);
                    oForm.EnableMenu("1281", true);  ////문서찾기
                    oForm.EnableMenu("1282", true); ////문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
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
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
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
                            PH_PY415_AddMatrixRow();
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY415_FormItemEnabled();
                            PH_PY415_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281":
                            // 문서찾기
                            PH_PY415_FormItemEnabled();
                            PH_PY415_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            // 문서추가
                            PH_PY415_FormItemEnabled();
                            PH_PY415_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY415_FormItemEnabled();
                            break;
                        case "1293":
                            // 행삭제
                            break;
                        case "7169":
                            //엑셀 내보내기

                            //엑셀 내보내기 이후 처리
                            oForm.Freeze(true);
                            oDS_PH_PY415B.RemoveRecord(oDS_PH_PY415B.Size - 1);
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                    //    break;
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY415_AddMatrixRow()
        {
            int oRow = 0;

            try
            {
                oForm.Freeze(true);
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY415B.GetValue("U_LineNum", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY415B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY415B.InsertRecord((oRow));
                        }
                        oDS_PH_PY415B.Offset = oRow;
                        oDS_PH_PY415B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY415B.SetValue("U_ColReg01", oRow, "");
                        oDS_PH_PY415B.SetValue("U_ColReg02", oRow, "");
                        oDS_PH_PY415B.SetValue("U_ColReg03", oRow, "");
                        oDS_PH_PY415B.SetValue("U_ColSum01", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY415B.Offset = oRow - 1;
                        oDS_PH_PY415B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY415B.SetValue("U_ColReg01", oRow - 1, "");
                        oDS_PH_PY415B.SetValue("U_ColReg02", oRow - 1, "");
                        oDS_PH_PY415B.SetValue("U_ColReg03", oRow - 1, "");
                        oDS_PH_PY415B.SetValue("U_ColSum01", oRow - 1, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY415B.Offset = oRow;
                    oDS_PH_PY415B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY415B.SetValue("U_ColReg01", oRow, "");
                    oDS_PH_PY415B.SetValue("U_ColReg02", oRow, "");
                    oDS_PH_PY415B.SetValue("U_ColReg03", oRow, "");
                    oDS_PH_PY415B.SetValue("U_ColSum01", oRow, "");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY415_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
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
                oForm.Freeze(false);
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
                            if (pVal.Row > 0)
                            {
                                oMat1.SelectRow(pVal.Row, true, false);
                            }
                            break;
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
                    PH_PY415_FormItemEnabled();
                    PH_PY415_AddMatrixRow();
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
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
                        oMat1.AutoResizeColumns();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_COMBO_SELECT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY415B);
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
                        switch (pVal.ItemUID)
                        {
                            case "MSTCOD":
                                // 사원명 찿아서 화면 표시 하기
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.VALUE.Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("FullName").Specific.String = oRecordSet.Fields.Item("U_FullName").Value.Trim();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {

                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat1" & pVal.ColUID == "ACTCOD")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY415B", "U_ACTCOD,U_ACTNAM", pVal.ItemUID, (short)pVal.Row, "", "", "");
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CHOOSE_FROM_LIST_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string yyyy = string.Empty;
            string FullName = string.Empty;
            string sQry = string.Empty;
            string Result = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY415_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }

                    }
                    if (pVal.ItemUID == "BtnCalC")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY415_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            yyyy = oForm.Items.Item("Year").Specific.VALUE.Trim();
                            FullName = oForm.Items.Item("FullName").Specific.VALUE.Trim();
                            if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE.Trim()))
                            {
                                if (PSH_Globals.SBO_Application.MessageBox(yyyy + "년 전사원 정산계산을 하시겠습니까?", 2, "Yes", "No") == 2)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                if (PSH_Globals.SBO_Application.MessageBox(yyyy + "년 (" + FullName + ") 계산을 하시겠습니까?", 2, "Yes", "No") == 2)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }

                            sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
                            oRecordSet.DoQuery(sQry);

                            Result = oRecordSet.Fields.Item(0).Value;
                            if (Result != "Y")
                            {
                                PSH_Globals.SBO_Application.MessageBox("계산불가 년도입니다. 담당자에게 문의바랍니다.");
                            }
                            if (Result == "Y")
                            {
                                PH_PY415_MTX01();
                                PSH_Globals.SBO_Application.MessageBox(PH_PY415_Calc());
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        //{
                        //    if (pVal.ActionSuccess == true)
                        //    {
                        //        PH_PY415_FormItemEnabled();
                        //        PH_PY415_AddMatrixRow();
                        //    }
                        //}
                        //else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        //{
                        //    if (pVal.ActionSuccess == true)
                        //    {
                        //        PH_PY415_FormItemEnabled();
                        //        PH_PY415_AddMatrixRow();
                        //    }
                        //}
                        //else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //{
                        //    if (pVal.ActionSuccess == true)
                        //    {
                        //        PH_PY415_FormItemEnabled();
                        //    }
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                if (pVal.CharPressed == Convert.ToDouble("9"))
                {
                    if (pVal.ItemUID == "MSTCOD")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE.Trim()))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                            BubbleEvent = false;
                        }
                    }

                }
                else
                if (pVal.Before_Action == false)
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY415_DataValidCheck()
        {
            bool functionReturnValue = false;

            try
            {
                // 사업장
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE.Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //년도
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.VALUE.Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
                functionReturnValue = true;
                return functionReturnValue;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY415_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                functionReturnValue = false;
                return functionReturnValue;
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY415_MTX01 
        /// </summary>
        private void PH_PY415_MTX01()
        {
            int i = 0;
            int ErrNum = 0;

            string sQry = string.Empty;
            string CLTCOD = string.Empty;       // 사업장
            string Year = string.Empty;         // 년도
            string MSTCOD = string.Empty;       

            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            
            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                Year = oForm.Items.Item("Year").Specific.VALUE.Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.VALUE.Trim();

                if (string.IsNullOrEmpty(MSTCOD))
                {
                    MSTCOD = "%";
                }
               
                //    If Param02 <= "2013" Then
                //        sQry = "EXEC PH_PY415_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
                //    Else
                //        '//2014년 변경
                //        sQry = "EXEC PH_PY415_2015 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
                //    End If

                // 2019귀속
                if (Convert.ToDouble(Year) >= 2019)
                {
                    sQry = "EXEC PH_PY415_2019 '" + CLTCOD + "','" + Year + "','" + MSTCOD + "'";
                }
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if ((oRecordSet.RecordCount == 0))
                {
                    oMat1.Clear();
                    ErrNum = 1;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY415B.InsertRecord((i));
                    }
                    oDS_PH_PY415B.Offset = i;
                    oDS_PH_PY415B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY415B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("MSTCOD").Value);
                    oDS_PH_PY415B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("FullName").Value);
                    oDS_PH_PY415B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("TeamCode").Value);
                    oDS_PH_PY415B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("TeamName").Value);
                    oDS_PH_PY415B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("RspCode").Value);
                    oDS_PH_PY415B.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("RspName").Value);
                    oDS_PH_PY415B.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("ClsCode").Value);
                    oDS_PH_PY415B.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("ClsName").Value);

                    oRecordSet.MoveNext();
                }

                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY415_MTX01_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 정산, 표준공제, 기부조정 계산
        /// </summary>
        /// <returns>메소트 처리 결과</returns>
        private string PH_PY415_Calc()
        {
            string returnValue = string.Empty;

            int i = 0;

            string sQry = string.Empty;
            string CLTCOD = string.Empty;       // 사업장
            string Year = string.Empty;         // 년도
            string MSTCOD = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("계산 시작!", oRecordSet.RecordCount, false);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
                Year = oForm.Items.Item("Year").Specific.VALUE.Trim();

                for (i = 1; i <= oMat1.VisualRowCount; i++)
                {
                    MSTCOD = oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.VALUE.Trim();

                    //        If Param02 = "2014" And (oMat1.Columns("FullName").Cells(i).Specific.VALUE = "박용철" Or oMat1.Columns("FullName").Cells(i).Specific.VALUE = "국요한") Then
                    //            '//2014년 재정산시 박용철, 국요한 제외
                    //        Else
                    //            If Param02 <= "2013" Then
                    //                sQry = "EXEC PH_PY415_02 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
                    //            Else
                    //                '//2015 연말정산
                    //                sQry = "EXEC PH_PY415_2015_02 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
                    //            End If
                    //            Call oRecordSet.DoQuery(sQry)
                    //        End If

                    if (Convert.ToDouble(Year) >= 2019)
                    {
                        sQry = "EXEC PH_PY415_2019_02 '" + CLTCOD + "','" + Year + "','" + MSTCOD + "'"; //정산 계산
                        oRecordSet.DoQuery(sQry);
                    }

                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = "정산 계산 " + ProgressBar01.Value + "/" + oMat1.VisualRowCount + "건 처리중...!";
                }

                ProgressBar01.Value = 0;

                for (i = 1; i <= oMat1.VisualRowCount; i++)
                {
                    MSTCOD = oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.VALUE.Trim();

                    //        If Param02 = "2014" And (oMat1.Columns("FullName").Cells(i).Specific.VALUE = "박용철" Or oMat1.Columns("FullName").Cells(i).Specific.VALUE = "국요한") Then
                    //            '//2014년 재정산시 박용철, 국요한 제외
                    //        Else
                    //            If Param02 <= "2013" Then
                    //
                    //            Else
                    //                '//2015 연말정산_표준세액공제 적용받을시 계산
                    //                sQry = "EXEC PH_PY415_2015_03 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
                    //            End If
                    //            Call oRecordSet.DoQuery(sQry)
                    //        End If

                    if (Convert.ToDouble(Year) >= 2019)
                    {
                        sQry = "EXEC PH_PY415_2019_03 '" + CLTCOD + "','" + Year + "','" + MSTCOD + "'"; //표준공제 계산
                        oRecordSet.DoQuery(sQry);
                    }

                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = "표준공제 계산 " + ProgressBar01.Value + "/" + oMat1.VisualRowCount + "건 처리중...!";
                }

                ProgressBar01.Value = 0;
                ProgressBar01.Text = "기부조정 계산 처리중...!";

                //2018년(2017귀속)부터 기부조정계산 추가 (정산계산은 사원별로 하고 기부조정은 전체로함)
                if (Convert.ToDouble(Year) >= 2019)
                {
                    sQry = "EXEC PH_PY415_2019_04 '" + CLTCOD + "','" + Year + "'"; //기부조정 계산
                    oRecordSet.DoQuery(sQry);
                }

                ProgressBar01.Value = 100;

                returnValue = "정산계산을 완료하였습니다.";
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);

                returnValue = System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message;
                //PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY415_Calc_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }

            return returnValue;
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
//	internal class PH_PY415
//	{
//////********************************************************************************
//////  File           : PH_PY415.cls
//////  Module         : 급여관리 > 정산관리
//////  Desc           : 소득정산계산
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY415A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY415B;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY415.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY415_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY415");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "DocEntry"

//			oForm.Freeze(true);
//			PH_PY415_CreateItems();
//			PH_PY415_EnableMenus();
//			PH_PY415_SetDocument(oFromDocEntry01);
//			//    Call PH_PY415_FormResize

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

//		private bool PH_PY415_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;

//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//    Set oDS_PH_PY415A = oForm.DataSources.DBDataSources("@PH_PY415A")
//			oDS_PH_PY415B = oForm.DataSources.DBDataSources("@PS_USERDS01");

//			oMat1 = oForm.Items.Item("Mat01").Specific;

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
//			oMat1.AutoResizeColumns();


//			//사업장_S
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
//			//사업장_E

//			//년도_S
//			oForm.DataSources.UserDataSources.Add("Year", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("Year").Specific.DataBind.SetBound(true, "", "Year");
//			//년도_E

//			//사번_S
//			oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");
//			//사번_E


//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			//사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;


//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY415_CreateItems_Error:

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY415_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			//    Call oForm.EnableMenu("1283", False)         '// 삭제
//			//    Call oForm.EnableMenu("1287", True)          '// 복제
//			//'    Call oForm.EnableMenu("1286", True)         '// 닫기
//			//    Call oForm.EnableMenu("1284", True)         '// 취소
//			//    Call oForm.EnableMenu("1293", True)         '// 행삭제

//			return;
//			PH_PY415_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY415_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY415_FormItemEnabled();
//				//        Call PH_PY415_AddMatrixRow
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY415_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY415_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY415_FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;
//			string CLTCOD = null;

//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				//폼 DocEntry 세팅
//				//        Call PH_PY415_FormClear

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				//년도 세팅
//				//        Call oDS_PH_PY415A.setValue("U_StdYear", 0, Format(Date, "YYYY"))
//				//UPGRADE_WARNING: oForm.Items(Year).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Year").Specific.VALUE = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYY")) - 1;

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

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
//			PH_PY415_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			short loopCount = 0;
//			//For Loop 용 (VALIDATE Event에서 사용)
//			string yyyy = null;
//			//년도
//			string Result = null;
//			string FullName = null;
//			//성명

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						//                If pval.ItemUID = "1" Then
//						//                    If oForm.Mode = fm_ADD_MODE Then
//						//                        If PH_PY415_DataValidCheck = False Then
//						//                            BubbleEvent = False
//						//                        End If
//						//
//						//                        '//해야할일 작업
//						//                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                        If PH_PY415_DataValidCheck = False Then
//						//                            BubbleEvent = False
//						//                        End If
//						//                        '//해야할일 작업
//						//
//						//                    ElseIf oForm.Mode = fm_OK_MODE Then
//						//                    End If
//						//                End If
//						if (pval.ItemUID == "BtnCalC") {

//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY415_DataValidCheck() == false) {
//									BubbleEvent = false;
//									return;
//								}

//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								yyyy = oForm.Items.Item("Year").Specific.VALUE;
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								FullName = oForm.Items.Item("FullName").Specific.VALUE;
//								//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE)) {
//									if (MDC_Globals.Sbo_Application.MessageBox(yyyy + "년 전사원 정산계산을 하시겠습니까?", 2, "Yes", "No") == 2) {
//										BubbleEvent = false;
//										return;
//									}
//								} else {
//									if (MDC_Globals.Sbo_Application.MessageBox(yyyy + "년 (" + FullName + ") 계산을 하시겠습니까?", 2, "Yes", "No") == 2) {
//										BubbleEvent = false;
//										return;
//									}
//								}
//								//
//								//                        '//해야할일 작업
//								//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								yyyy = oForm.Items.Item("Year").Specific.VALUE;
//								sQry = "select b.U_UseYN from [@PS_HR200L] b where b.code ='87' and b.u_code ='" + yyyy + "'";
//								oRecordSet.DoQuery(sQry);

//								//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								Result = oRecordSet.Fields.Item(0).Value;
//								if (Result != "Y") {
//									MDC_Globals.Sbo_Application.MessageBox("계산불가 년도입니다. 담당자에게 문의바랍니다.");
//								}
//								if (Result == "Y") {
//									PH_PY415_MTX01();
//									PH_PY415_Calc();
//								}
//							}

//						}

//					} else if (pval.BeforeAction == false) {
//						//                If pval.ItemUID = "1" Then
//						//                    If oForm.Mode = fm_ADD_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY415_FormItemEnabled
//						//                            Call PH_PY415_AddMatrixRow
//						//                        End If
//						//                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY415_FormItemEnabled
//						//                            Call PH_PY415_AddMatrixRow
//						//                        End If
//						//                    ElseIf oForm.Mode = fm_OK_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY415_FormItemEnabled
//						//                        End If
//						//                    End If
//						//                End If
//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2

//					if (pval.BeforeAction == true) {

//						if (pval.CharPressed == Convert.ToDouble("9")) {
//							if (pval.ItemUID == "MSTCOD") {
//								//UPGRADE_WARNING: oForm.Items(MSTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.VALUE)) {
//									MDC_Globals.Sbo_Application.ActivateMenuItem(("7425"));
//									BubbleEvent = false;
//								}
//							}


//						}

//					} else if (pval.Before_Action == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat01":
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
//							//                    Call PH_PY415_AddMatrixRow
//							oMat1.AutoResizeColumns();
//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Mat01":
//								if (pval.Row > 0) {
//									oMat1.SelectRow(pval.Row, true, false);
//								}
//								break;
//						}

//						switch (pval.ItemUID) {
//							case "Mat01":
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

//					}
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
//					oForm.Freeze(true);
//					if (pval.BeforeAction == true) {

//						if (pval.ItemChanged == true) {

//						}

//					} else if (pval.BeforeAction == false) {

//						if (pval.ItemChanged == true) {

//							switch (pval.ItemUID) {
//								case "MSTCOD":
//									////사원명 찿아서 화면 표시 하기
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE) + "'";
//									oRecordSet.DoQuery(sQry);
//									//UPGRADE_WARNING: oForm.Items(FullName).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("FullName").Specific.String = Strings.Trim(oRecordSet.Fields.Item("U_FullName").Value);
//									break;

//							}

//						}

//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pval.BeforeAction == true) {
//					} else if (pval.BeforeAction == false) {
//						oMat1.LoadFromDataSource();

//						PH_PY415_FormItemEnabled();
//						PH_PY415_AddMatrixRow();
//						oMat1.AutoResizeColumns();

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
//						//UPGRADE_NOTE: oDS_PH_PY415A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY415A = null;
//						//UPGRADE_NOTE: oDS_PH_PY415B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY415B = null;

//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;

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

//						oMat1.AutoResizeColumns();

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
//						//                If pval.ItemUID = "Code" Then
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY415A", "Code")
//						//                End If
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


//			short loopCount = 0;
//			double FeeTot = 0;
//			double TuiTot = 0;
//			double Total = 0;

//			oForm.Freeze(true);

//			if ((pval.BeforeAction == true)) {
//				switch (pval.MenuUID) {
//					case "1283":
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

//					case "7169":
//						//엑셀 내보내기

//						//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
//						PH_PY415_AddMatrixRow();
//						break;

//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY415_FormItemEnabled();
//						PH_PY415_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY415_FormItemEnabled();
//						PH_PY415_AddMatrixRow();
//						oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY415_FormItemEnabled();
//						PH_PY415_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY415_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						break;

//					case "7169":
//						//엑셀 내보내기

//						//엑셀 내보내기 이후 처리
//						oForm.Freeze(true);
//						oDS_PH_PY415B.RemoveRecord(oDS_PH_PY415B.Size - 1);
//						oMat1.LoadFromDataSource();
//						oForm.Freeze(false);
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
//				case "Mat01":
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

//		public void PH_PY415_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY415B.GetValue("U_LineNum", oRow - 1)))) {
//					if (oDS_PH_PY415B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY415B.InsertRecord((oRow));
//					}
//					oDS_PH_PY415B.Offset = oRow;
//					oDS_PH_PY415B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					oDS_PH_PY415B.SetValue("U_ColReg01", oRow, "");
//					oDS_PH_PY415B.SetValue("U_ColReg02", oRow, "");
//					oDS_PH_PY415B.SetValue("U_ColReg03", oRow, "");
//					oDS_PH_PY415B.SetValue("U_ColSum01", oRow, "");
//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY415B.Offset = oRow - 1;
//					oDS_PH_PY415B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					oDS_PH_PY415B.SetValue("U_ColReg01", oRow - 1, "");
//					oDS_PH_PY415B.SetValue("U_ColReg02", oRow - 1, "");
//					oDS_PH_PY415B.SetValue("U_ColReg03", oRow - 1, "");
//					oDS_PH_PY415B.SetValue("U_ColSum01", oRow - 1, "");
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY415B.Offset = oRow;
//				oDS_PH_PY415B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				oDS_PH_PY415B.SetValue("U_ColReg01", oRow, "");
//				oDS_PH_PY415B.SetValue("U_ColReg02", oRow, "");
//				oDS_PH_PY415B.SetValue("U_ColReg03", oRow, "");
//				oDS_PH_PY415B.SetValue("U_ColSum01", oRow, "");
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY415_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY415_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY415'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY415_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY415_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//사업장
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//년도
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("Year").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("Year").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//
//			//    oMat1.FlushToDataSource
//			//    '// Matrix 마지막 행 삭제(DB 저장시)
//			//    If oDS_PH_PY415B.Size > 1 Then oDS_PH_PY415B.RemoveRecord (oDS_PH_PY415B.Size - 1)
//			//
//			//    oMat1.LoadFromDataSource

//			functionReturnValue = true;
//			return functionReturnValue;


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY415_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY415_MTX01()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;
//			string Param05 = null;


//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("Year").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("MSTCOD").Specific.VALUE;

//			if (string.IsNullOrEmpty(Strings.Trim(Param03))) {
//				Param03 = "%";
//			}
//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//			//    If Param02 <= "2013" Then
//			//        sQry = "EXEC PH_PY415_01 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
//			//    Else
//			//        '//2014년 변경
//			//        sQry = "EXEC PH_PY415_2015 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
//			//    End If
//			// 2017귀속
//			if (Param02 >= "2018") {
//				sQry = "EXEC PH_PY415_2018 '" + Param01 + "','" + Param02 + "','" + Param03 + "'";
//			}

//			oRecordSet.DoQuery(sQry);

//			oMat1.Clear();
//			oMat1.FlushToDataSource();
//			oMat1.LoadFromDataSource();

//			if ((oRecordSet.RecordCount == 0)) {
//				oMat1.Clear();
//				goto PH_PY415_MTX01_Exit;
//			}

//			for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//				if (i != 0) {
//					oDS_PH_PY415B.InsertRecord((i));
//				}
//				oDS_PH_PY415B.Offset = i;
//				oDS_PH_PY415B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_PH_PY415B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("MSTCOD").Value);
//				oDS_PH_PY415B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("FullName").Value);
//				oDS_PH_PY415B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("TeamCode").Value);
//				oDS_PH_PY415B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("TeamName").Value);
//				oDS_PH_PY415B.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("RspCode").Value);
//				oDS_PH_PY415B.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("RspName").Value);
//				oDS_PH_PY415B.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("ClsCode").Value);
//				oDS_PH_PY415B.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("ClsName").Value);


//				oRecordSet.MoveNext();
//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";

//			}


//			oMat1.LoadFromDataSource();
//			oMat1.AutoResizeColumns();
//			oForm.Update();

//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY415_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			if ((ProgressBar01 != null)) {
//				ProgressBar01.Stop();
//			}
//			MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
//			return;
//			PH_PY415_MTX01_Error:
//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY415_Validate(string ValidateType)
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
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY415A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY415A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY415_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY415_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY415_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY415_Calc()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;
//			string Param05 = null;


//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("Year").Specific.VALUE;

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("계산시작!", oRecordSet.RecordCount, false);



//			for (i = 1; i <= oMat1.VisualRowCount; i++) {
//				//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Param03 = oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.VALUE;

//				//        If Param02 = "2014" And (oMat1.Columns("FullName").Cells(i).Specific.VALUE = "박용철" Or oMat1.Columns("FullName").Cells(i).Specific.VALUE = "국요한") Then
//				//            '//2014년 재정산시 박용철, 국요한 제외
//				//        Else
//				//            If Param02 <= "2013" Then
//				//                sQry = "EXEC PH_PY415_02 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
//				//            Else
//				//                '//2015 연말정산
//				//                sQry = "EXEC PH_PY415_2015_02 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
//				//            End If
//				//            Call oRecordSet.DoQuery(sQry)
//				//        End If

//				if (Param02 >= "2018") {
//					sQry = "EXEC PH_PY415_2018_02 '" + Param01 + "','" + Param02 + "','" + Param03 + "'";
//					oRecordSet.DoQuery(sQry);
//				}

//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oMat1.VisualRowCount + "건 조회중...!";
//			}

//			ProgressBar01.Value = 0;

//			for (i = 1; i <= oMat1.VisualRowCount; i++) {
//				//UPGRADE_WARNING: oMat1.Columns().Cells().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				Param03 = oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.VALUE;

//				//        If Param02 = "2014" And (oMat1.Columns("FullName").Cells(i).Specific.VALUE = "박용철" Or oMat1.Columns("FullName").Cells(i).Specific.VALUE = "국요한") Then
//				//            '//2014년 재정산시 박용철, 국요한 제외
//				//        Else
//				//            If Param02 <= "2013" Then
//				//
//				//            Else
//				//                '//2015 연말정산_표준세액공제 적용받을시 계산
//				//                sQry = "EXEC PH_PY415_2015_03 '" & Param01 & "','" & Param02 & "','" & Param03 & "'"
//				//            End If
//				//            Call oRecordSet.DoQuery(sQry)
//				//        End If

//				if (Param02 >= "2018") {
//					sQry = "EXEC PH_PY415_2018_03 '" + Param01 + "','" + Param02 + "','" + Param03 + "'";
//					oRecordSet.DoQuery(sQry);
//				}


//				ProgressBar01.Value = ProgressBar01.Value + 1;
//				ProgressBar01.Text = ProgressBar01.Value + "/" + oMat1.VisualRowCount + "건 조회중...!";
//			}

//			//2018년(2017귀속)부터 기부조정계산 추가 (정산계산은 사원별로 하고 기부조정은 전체로함)

//			if (Param02 >= "2018") {
//				sQry = "EXEC PH_PY415_2018_04 '" + Param01 + "','" + Param02 + "'";
//				oRecordSet.DoQuery(sQry);
//			}

//			MDC_Globals.Sbo_Application.MessageBox("정산계산을 완료하였습니다.");

//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return;
//			PH_PY415_Calc_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			if ((ProgressBar01 != null)) {
//				ProgressBar01.Stop();
//			}
//			MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
//			return;
//			PH_PY415_Calc_Error:
//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY415_Calc_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}
//	}
//}
