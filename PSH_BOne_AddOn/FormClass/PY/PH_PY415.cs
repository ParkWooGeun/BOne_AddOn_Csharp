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
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
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

                

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
          //      oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                oForm.Visible = true;
                PH_PY415_CreateItems();
                PH_PY415_EnableMenus();
                PH_PY415_SetDocument(oFormDocEntry01);
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
            string sQry;
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
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY415_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY415_FormItemEnabled();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY415_FormItemEnabled();

                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
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
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.Items.Item("Year").Specific.Value = Convert.ToString(DateTime.Now.Year - 1);

                    oForm.EnableMenu("1281", true); ////문서찾기
                    oForm.EnableMenu("1282", false); ////문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.EnableMenu("1281", false); ////문서찾기
                    oForm.EnableMenu("1282", true); ////문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
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
                        case "7169": //엑셀 내보내기
                            PH_PY415_AddMatrixRow(); //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
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
                        case "1281": // 문서찾기
                            PH_PY415_FormItemEnabled();
                            PH_PY415_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": // 문서추가
                            PH_PY415_FormItemEnabled();
                            PH_PY415_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY415_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
                            break;
                        case "7169": //엑셀 내보내기
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.Value.Trim() + "'";
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

                            yyyy = oForm.Items.Item("Year").Specific.Value.Trim();
                            FullName = oForm.Items.Item("FullName").Specific.Value.Trim();
                            if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.Trim()))
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
                        if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.Value.Trim()))
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
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value.Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //년도
                if (string.IsNullOrEmpty(oForm.Items.Item("Year").Specific.Value.Trim()))
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

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                Year = oForm.Items.Item("Year").Specific.Value.Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.Trim();

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

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                Year = oForm.Items.Item("Year").Specific.Value.Trim();

                for (i = 1; i <= oMat1.VisualRowCount; i++)
                {
                    MSTCOD = oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value.Trim();

                    //        If Param02 = "2014" And (oMat1.Columns("FullName").Cells(i).Specific.Value = "박용철" Or oMat1.Columns("FullName").Cells(i).Specific.Value = "국요한") Then
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
                    MSTCOD = oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value.Trim();

                    //        If Param02 = "2014" And (oMat1.Columns("FullName").Cells(i).Specific.Value = "박용철" Or oMat1.Columns("FullName").Cells(i).Specific.Value = "국요한") Then
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
