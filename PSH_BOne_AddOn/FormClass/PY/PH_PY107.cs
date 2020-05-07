using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여기준일설정
    /// </summary>
    internal class PH_PY107 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;

        private SAPbouiCOM.DBDataSource oDS_PH_PY107A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY107B;

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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY107.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY107_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY107");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                //oForm.Visible = true;
                PH_PY107_CreateItems();
                PH_PY107_EnableMenus();
                PH_PY107_SetDocument(oFromDocEntry01);
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
        private void PH_PY107_CreateItems()
        {
            int i = 0;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY107A = oForm.DataSources.DBDataSources.Item("@PH_PY107A");
                oDS_PH_PY107B = oForm.DataSources.DBDataSources.Item("@PH_PY107B");

                oMat1 = oForm.Items.Item("Mat1").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat1.AutoResizeColumns();

                // 사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 적용년월
                oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);

                // 월급직일수계산
                oForm.Items.Item("INPDAY").Specific.ValidValues.Add("1", "30일기준");
                oForm.Items.Item("INPDAY").Specific.ValidValues.Add("2", "말일기준");
                oForm.Items.Item("INPDAY").DisplayDesc = true;

                // 메트릭스 SET
                // 급여지급
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P213' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    while (!(oRecordSet.EoF))
                    {
                        oMat1.Columns.Item("PAYSEL").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("PAYSEL").DisplayDesc = true;

                // 시작월
                oMat1.Columns.Item("STRMON").ValidValues.Add("0", "당월");
                oMat1.Columns.Item("STRMON").ValidValues.Add("-1", "전월");
                oMat1.Columns.Item("STRMON").ValidValues.Add("1", "익월");
                oMat1.Columns.Item("STRMON").DisplayDesc = true;

                // 지급월
                oMat1.Columns.Item("JIGMON").ValidValues.Add("0", "당월");
                oMat1.Columns.Item("JIGMON").ValidValues.Add("-1", "전월");
                oMat1.Columns.Item("JIGMON").ValidValues.Add("1", "익월");
                oMat1.Columns.Item("JIGMON").DisplayDesc = true;

                // 상여한도월
                oMat1.Columns.Item("BNSEMM").ValidValues.Add("0", "당월");
                oMat1.Columns.Item("BNSEMM").ValidValues.Add("-1", "전월");
                oMat1.Columns.Item("BNSEMM").ValidValues.Add("1", "익월");
                oMat1.Columns.Item("BNSEMM").DisplayDesc = true;

                // 급여한도월
                oMat1.Columns.Item("PAYEMM").ValidValues.Add("0", "당월");
                oMat1.Columns.Item("PAYEMM").ValidValues.Add("-1", "전월");
                oMat1.Columns.Item("PAYEMM").ValidValues.Add("1", "익월");
                oMat1.Columns.Item("PAYEMM").DisplayDesc = true;

                // 시작일
                for (i = 1; i <= 31; i++)
                {
                    oMat1.Columns.Item("STRDAY").ValidValues.Add(Convert.ToString(i), i + "일");
                }

                // 지급일
                oMat1.Columns.Item("JIGDAY").ValidValues.Add(Convert.ToString(0), "말일");
                for (i = 1; i <= 31; i++)
                {
                    oMat1.Columns.Item("JIGDAY").ValidValues.Add(Convert.ToString(i), i + "일");
                }

                // 상여한도기준일
                oMat1.Columns.Item("BNSEDD").ValidValues.Add(Convert.ToString(0), "말일");
                for (i = 1; i <= 31; i++)
                {
                    oMat1.Columns.Item("BNSEDD").ValidValues.Add(Convert.ToString(i), i + "일");
                }

                // 급여한도기준일
                oMat1.Columns.Item("PAYEDD").ValidValues.Add(Convert.ToString(0), "말일");
                for (i = 1; i <= 31; i++)
                {
                    oMat1.Columns.Item("PAYEDD").ValidValues.Add(Convert.ToString(i), i + "일");
                }

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY107_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private void PH_PY107_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY107_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY107_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY107_FormItemEnabled();
                    PH_PY107_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY107_FormItemEnabled();

                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY107_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            PH_PY107_FormItemEnabled();
                            PH_PY107_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281":
                            // 문서찾기
                            PH_PY107_FormItemEnabled();
                            PH_PY107_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            // 문서추가
                            PH_PY107_FormItemEnabled();
                            PH_PY107_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY107_FormItemEnabled();
                            break;
                        case "1293":
                            // 행삭제
                            // [MAT1 용]
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while ((i <= oDS_PH_PY107B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY107B.GetValue("U_PAYSEL", i)))
                                    {
                                        oDS_PH_PY107B.RemoveRecord((i));
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY107B.Size; i++)
                                {
                                    oDS_PH_PY107B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            PH_PY107_AddMatrixRow();
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
        private void PH_PY107_AddMatrixRow()
        {
            int oRow = 0;

            try
            {
                oForm.Freeze(true);
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY107B.GetValue("U_PAYSEL", oRow - 1).Trim()))
                    {
                        if (oDS_PH_PY107B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY107B.InsertRecord((oRow));
                        }
                        oDS_PH_PY107B.Offset = oRow;
                        oDS_PH_PY107B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY107B.SetValue("U_PAYSEL", oRow, "");
                        oDS_PH_PY107B.SetValue("U_STRMON", oRow, "");
                        oDS_PH_PY107B.SetValue("U_STRDAY", oRow, "");
                        oDS_PH_PY107B.SetValue("U_JIGMON", oRow, "");
                        oDS_PH_PY107B.SetValue("U_JIGDAY", oRow, "");
                        oDS_PH_PY107B.SetValue("U_BNSEMM", oRow, "");
                        oDS_PH_PY107B.SetValue("U_BNSEDD", oRow, "");
                        oDS_PH_PY107B.SetValue("U_PAYEMM", oRow, "");
                        oDS_PH_PY107B.SetValue("U_PAYEDD", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY107B.Offset = oRow - 1;
                        oDS_PH_PY107B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY107B.SetValue("U_PAYSEL", oRow - 1, "");
                        oDS_PH_PY107B.SetValue("U_STRMON", oRow - 1, "");
                        oDS_PH_PY107B.SetValue("U_STRDAY", oRow - 1, "");
                        oDS_PH_PY107B.SetValue("U_JIGMON", oRow - 1, "");
                        oDS_PH_PY107B.SetValue("U_JIGDAY", oRow - 1, "");
                        oDS_PH_PY107B.SetValue("U_BNSEMM", oRow - 1, "");
                        oDS_PH_PY107B.SetValue("U_BNSEDD", oRow - 1, "");
                        oDS_PH_PY107B.SetValue("U_PAYEMM", oRow - 1, "");
                        oDS_PH_PY107B.SetValue("U_PAYEDD", oRow - 1, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY107B.Offset = oRow;
                    oDS_PH_PY107B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY107B.SetValue("U_PAYSEL", oRow, "");
                    oDS_PH_PY107B.SetValue("U_STRMON", oRow, "");
                    oDS_PH_PY107B.SetValue("U_STRDAY", oRow, "");
                    oDS_PH_PY107B.SetValue("U_JIGMON", oRow, "");
                    oDS_PH_PY107B.SetValue("U_JIGDAY", oRow, "");
                    oDS_PH_PY107B.SetValue("U_BNSEMM", oRow, "");
                    oDS_PH_PY107B.SetValue("U_BNSEDD", oRow, "");
                    oDS_PH_PY107B.SetValue("U_PAYEMM", oRow, "");
                    oDS_PH_PY107B.SetValue("U_PAYEDD", oRow, "");
                    oMat1.LoadFromDataSource();
                }

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY107_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                    PH_PY107_FormItemEnabled();
                    PH_PY107_AddMatrixRow();
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
                        if (pVal.ItemUID == "Mat1" & pVal.ColUID == "PAYSEL")
                        {
                            PH_PY107_AddMatrixRow();

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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY107A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY107B);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY107_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }

                            ////해야할일 작업
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY107_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                            ////해야할일 작업
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
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
                                PH_PY107_FormItemEnabled();
                                PH_PY107_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY107_FormItemEnabled();
                                PH_PY107_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY107_FormItemEnabled();
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
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY107_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("INPDAY").Enabled = true;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", true); ////문서찾기
                    oForm.EnableMenu("1282", false); ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("INPDAY").Enabled = true;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.EnableMenu("1281", false); ////문서찾기
                    oForm.EnableMenu("1282", true); ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("INPDAY").Enabled = false;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);
                    oForm.EnableMenu("1281", true);  ////문서찾기
                    oForm.EnableMenu("1282", true); ////문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY301_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private bool PH_PY107_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            int k = 0;
            string Chk_Data = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.ChkYearMonth(oDS_PH_PY107A.GetValue("U_YM", 0)) == false)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("적용년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (string.IsNullOrEmpty(oDS_PH_PY107A.GetValue("U_INPDAY", 0).Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("월급직일수계산은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("INPDAY").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }
                // Code & Name 생성
                oDS_PH_PY107A.SetValue("Code", 0, oDS_PH_PY107A.GetValue("U_CLTCOD", 0).Trim() + oDS_PH_PY107A.GetValue("U_YM", 0).Trim());
                oDS_PH_PY107A.SetValue("Name", 0, oDS_PH_PY107A.GetValue("U_CLTCOD", 0).Trim() + oDS_PH_PY107A.GetValue("U_YM", 0).Trim());

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (dataHelpClass.Value_ChkYn("[@PH_PY107A]", "Code", "'" + oDS_PH_PY107A.GetValue("Code", 0).Trim() + "'".Trim(),"") == false)
                    {
                        PSH_Globals.SBO_Application.StatusBar.SetText("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return functionReturnValue;
                    }
                }

                // 라인
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        // 구분
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("PAYSEL").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("PAYSEL").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        // 시작월
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("STRMON").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("STRMON").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        // 시작일
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("STRDAY").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("시작일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("STRDAY").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        // 지급월
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("JIGMON").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("지급월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("JIGMON").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        // 지급일
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("JIGDAY").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("지급일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("JIGDAY").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        // 상여한도월
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("BNSEMM").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("상여한도월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("BNSEMM").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        // 상여한도일
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("BNSEDD").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("상여한도일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("BNSEDD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        // 급여한도월
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("PAYEMM").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("급여한도월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("PAYEMM").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        // 급여한도일
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("PAYEDD").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("급여한도일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("PAYEDD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }

                        ////중복체크
                        Chk_Data = oDS_PH_PY107B.GetValue("U_PAYSEL", i).Trim();
                        for (k = i + 1; k <= oMat1.VisualRowCount - 2; k++)
                        {
                            oDS_PH_PY107B.Offset = k;
                            if (Chk_Data.Trim() == oDS_PH_PY107B.GetValue("U_PAYSEL", k).Trim())
                            {
                                PSH_Globals.SBO_Application.SetStatusBarMessage("내용이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                oMat1.Columns.Item("PAYSEL").Cells.Item(k + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                return functionReturnValue;
                            }
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();
                //// Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY107B.Size > 1)
                    oDS_PH_PY107B.RemoveRecord((oDS_PH_PY107B.Size - 1));

                oMat1.LoadFromDataSource();
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY107_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                functionReturnValue = false;
                return functionReturnValue;
            }
            finally
            {
            }
            return functionReturnValue;
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
//	internal class PH_PY107
//	{
//////********************************************************************************
//////  File           : PH_PY107.cls
//////  Module         : 인사관리 > 급여관리
//////  Desc           : 급상여기준일설정
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

////'// 그리드 사용시
////Public oGrid1           As SAPbouiCOM.Grid
////Public oDS_PH_PY107     As SAPbouiCOM.DataTable
////
////'// 매트릭스 사용시
//		public SAPbouiCOM.Matrix oMat1;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY107A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY107B;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY107.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY107_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY107");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			oForm.DataBrowser.BrowseBy = "Code";

//			oForm.Freeze(true);
//			PH_PY107_CreateItems();
//			PH_PY107_EnableMenus();
//			PH_PY107_SetDocument(oFromDocEntry01);
//			//    Call PH_PY107_FormResize

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

//		private bool PH_PY107_CreateItems()
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


//			//    '//매트릭스--------------------------------------------------------------------------------------
//			oDS_PH_PY107A = oForm.DataSources.DBDataSources("@PH_PY107A");
//			////헤더
//			oDS_PH_PY107B = oForm.DataSources.DBDataSources("@PH_PY107B");
//			////라인

//			oMat1 = oForm.Items.Item("Mat1").Specific;
//			//

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//			oMat1.AutoResizeColumns();
//			//    '//----------------------------------------------------------------------------------------------

//			////----------------------------------------------------------------------------------------------
//			//// 아이템 설정
//			////----------------------------------------------------------------------------------------------
//			////사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			////기준일계산
//			oCombo = oForm.Items.Item("INPDAY").Specific;
//			oCombo.ValidValues.Add("1", "30일기준");
//			oCombo.ValidValues.Add("2", "말일기준");
//			oForm.Items.Item("INPDAY").DisplayDesc = true;

//			////급여지급
//			oColumn = oMat1.Columns.Item("PAYSEL");
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P213' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				while (!(oRecordSet.EoF)) {
//					oColumn.ValidValues.Add(Strings.Trim(oRecordSet.Fields.Item(0).Value), Strings.Trim(oRecordSet.Fields.Item(1).Value));
//					oRecordSet.MoveNext();
//				}
//			}
//			oColumn.DisplayDesc = true;

//			////시작월
//			oColumn = oMat1.Columns.Item("STRMON");
//			oColumn.ValidValues.Add("0", "당월");
//			oColumn.ValidValues.Add("-1", "전월");
//			oColumn.ValidValues.Add("1", "익월");
//			oColumn.DisplayDesc = true;

//			////지급월
//			oColumn = oMat1.Columns.Item("JIGMON");
//			oColumn.ValidValues.Add("0", "당월");
//			oColumn.ValidValues.Add("-1", "전월");
//			oColumn.ValidValues.Add("1", "익월");
//			oColumn.DisplayDesc = true;

//			////상여한도월
//			oColumn = oMat1.Columns.Item("BNSEMM");
//			oColumn.ValidValues.Add("0", "당월");
//			oColumn.ValidValues.Add("-1", "전월");
//			oColumn.ValidValues.Add("1", "익월");
//			oColumn.DisplayDesc = true;

//			////급여한도월
//			oColumn = oMat1.Columns.Item("PAYEMM");
//			oColumn.ValidValues.Add("0", "당월");
//			oColumn.ValidValues.Add("-1", "전월");
//			oColumn.ValidValues.Add("1", "익월");
//			oColumn.DisplayDesc = true;

//			//// 시작일
//			oColumn = oMat1.Columns.Item("STRDAY");
//			for (i = 1; i <= 31; i++) {
//				oColumn.ValidValues.Add(Convert.ToString(i), i + "일");
//			}

//			//// 지급일
//			oColumn = oMat1.Columns.Item("JIGDAY");
//			oColumn.ValidValues.Add(Convert.ToString(0), "말일");
//			for (i = 1; i <= 31; i++) {
//				oColumn.ValidValues.Add(Convert.ToString(i), i + "일");
//			}

//			//// 상여한도기준일
//			oColumn = oMat1.Columns.Item("BNSEDD");
//			oColumn.ValidValues.Add(Convert.ToString(0), "말일");
//			for (i = 1; i <= 31; i++) {
//				oColumn.ValidValues.Add(Convert.ToString(i), i + "일");
//			}

//			//// 급여한도기준일
//			oColumn = oMat1.Columns.Item("PAYEDD");
//			oColumn.ValidValues.Add(Convert.ToString(0), "말일");
//			for (i = 1; i <= 31; i++) {
//				oColumn.ValidValues.Add(Convert.ToString(i), i + "일");
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
//			PH_PY107_CreateItems_Error:

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
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY107_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY107_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", true);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", true);
//			////행삭제

//			return;
//			PH_PY107_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY107_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY107_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY107_FormItemEnabled();
//				PH_PY107_AddMatrixRow();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY107_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.Value = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY107_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY107_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY107_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;

//			 // ERROR: Not supported in C#: OnErrorStatement



//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				oForm.Items.Item("YM").Enabled = true;
//				oForm.Items.Item("INPDAY").Enabled = true;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");


//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				oForm.Items.Item("YM").Enabled = true;
//				oForm.Items.Item("INPDAY").Enabled = true;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가
//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = false;
//				oForm.Items.Item("YM").Enabled = false;
//				oForm.Items.Item("INPDAY").Enabled = false;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY107_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY107_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY107_DataValidCheck() == false) {
//									BubbleEvent = false;
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (PH_PY107_DataValidCheck() == false) {
//									BubbleEvent = false;
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//							}
//						}
//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemUID == "1") {
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (pval.ActionSuccess == true) {
//									//                            Call PH_PY107_FormItemEnabled
//									//                            Call PH_PY107_AddMatrixRow
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (pval.ActionSuccess == true) {
//									//                            Call PH_PY107_FormItemEnabled
//									//                            Call PH_PY107_AddMatrixRow
//								}
//							} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//								if (pval.ActionSuccess == true) {
//									//                            Call PH_PY107_FormItemEnabled
//								}
//							}
//						}

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat1":
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
//							if (pval.ItemUID == "Mat1" & pval.ColUID == "PAYSEL") {
//								PH_PY107_AddMatrixRow();
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
//							case "Mat1":
//								break;
//							//                    If pval.Row > 0 Then
//							//                        Call oMat1.SelectRow(pval.Row, True, False)
//							//                    End If
//						}

//						switch (pval.ItemUID) {
//							case "Mat1":
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
//						oMat1.LoadFromDataSource();

//						PH_PY107_FormItemEnabled();
//						PH_PY107_AddMatrixRow();

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
//						//                Set oDS_PH_PY107A = Nothing
//						//                Set oDS_PH_PY107B = Nothing

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
//						MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@PH_PY107A", ref "Code");
//						////접속자 권한에 따른 사업장 보기
//						break;
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY107_FormItemEnabled();
//						PH_PY107_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY107_FormItemEnabled();
//						PH_PY107_AddMatrixRow();
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY107_FormItemEnabled();
//						PH_PY107_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY107_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						//// [MAT1 용]
//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							oMat1.FlushToDataSource();

//							while ((i <= oDS_PH_PY107B.Size - 1)) {
//								if (string.IsNullOrEmpty(oDS_PH_PY107B.GetValue("U_PAYSEL", i))) {
//									oDS_PH_PY107B.RemoveRecord((i));
//									i = 0;
//								} else {
//									i = i + 1;
//								}
//							}

//							for (i = 0; i <= oDS_PH_PY107B.Size; i++) {
//								oDS_PH_PY107B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//							}

//							oMat1.LoadFromDataSource();
//						}
//						PH_PY107_AddMatrixRow();
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

//		public void PH_PY107_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1 용]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY107B.GetValue("U_PAYSEL", oRow - 1)))) {
//					if (oDS_PH_PY107B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY107B.InsertRecord((oRow));
//					}
//					oDS_PH_PY107B.Offset = oRow;
//					oDS_PH_PY107B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					oDS_PH_PY107B.SetValue("U_PAYSEL", oRow, "");
//					oDS_PH_PY107B.SetValue("U_STRMON", oRow, "");
//					oDS_PH_PY107B.SetValue("U_STRDAY", oRow, "");
//					oDS_PH_PY107B.SetValue("U_JIGMON", oRow, "");
//					oDS_PH_PY107B.SetValue("U_JIGDAY", oRow, "");
//					oDS_PH_PY107B.SetValue("U_BNSEMM", oRow, "");
//					oDS_PH_PY107B.SetValue("U_BNSEDD", oRow, "");
//					oDS_PH_PY107B.SetValue("U_PAYEMM", oRow, "");
//					oDS_PH_PY107B.SetValue("U_PAYEDD", oRow, "");
//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY107B.Offset = oRow - 1;
//					oDS_PH_PY107B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					oDS_PH_PY107B.SetValue("U_PAYSEL", oRow - 1, "");
//					oDS_PH_PY107B.SetValue("U_STRMON", oRow - 1, "");
//					oDS_PH_PY107B.SetValue("U_STRDAY", oRow - 1, "");
//					oDS_PH_PY107B.SetValue("U_JIGMON", oRow - 1, "");
//					oDS_PH_PY107B.SetValue("U_JIGDAY", oRow - 1, "");
//					oDS_PH_PY107B.SetValue("U_BNSEMM", oRow - 1, "");
//					oDS_PH_PY107B.SetValue("U_BNSEDD", oRow - 1, "");
//					oDS_PH_PY107B.SetValue("U_PAYEMM", oRow - 1, "");
//					oDS_PH_PY107B.SetValue("U_PAYEDD", oRow - 1, "");
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY107B.Offset = oRow;
//				oDS_PH_PY107B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				oDS_PH_PY107B.SetValue("U_PAYSEL", oRow, "");
//				oDS_PH_PY107B.SetValue("U_STRMON", oRow, "");
//				oDS_PH_PY107B.SetValue("U_STRDAY", oRow, "");
//				oDS_PH_PY107B.SetValue("U_JIGMON", oRow, "");
//				oDS_PH_PY107B.SetValue("U_JIGDAY", oRow, "");
//				oDS_PH_PY107B.SetValue("U_BNSEMM", oRow, "");
//				oDS_PH_PY107B.SetValue("U_BNSEDD", oRow, "");
//				oDS_PH_PY107B.SetValue("U_PAYEMM", oRow, "");
//				oDS_PH_PY107B.SetValue("U_PAYEDD", oRow, "");
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY107_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY107_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY107_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY107'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.Value = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
//			}
//			return;
//			PH_PY107_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY107_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY107_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			int i = 0;
//			int k = 0;
//			string sQry = null;
//			string Chk_Data = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			functionReturnValue = false;
//			//    '//----------------------------------------------------------------------------------
//			//    '//필수 체크
//			//    '//----------------------------------------------------------------------------------
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY107A.GetValue("U_CLTCOD", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				return functionReturnValue;
//			}

//			if (MDC_SetMod.ChkYearMonth(ref oDS_PH_PY107A.GetValue("U_YM", 0)) == false) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("적용년도는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				return functionReturnValue;
//			}

//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY107A.GetValue("U_INPDAY", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("월급직일수계산은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("INPDAY").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				return functionReturnValue;
//			}
//			//
//			oDS_PH_PY107A.SetValue("Code", 0, Strings.Trim(oDS_PH_PY107A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY107A.GetValue("U_YM", 0)));
//			oDS_PH_PY107A.SetValue("Name", 0, Strings.Trim(oDS_PH_PY107A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY107A.GetValue("U_YM", 0)));

//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				if (MDC_SetMod.Value_ChkYn("[@PH_PY107A]", "Code", "'" + Strings.Trim(oDS_PH_PY107A.GetValue("Code", 0)) + "'") == false) {
//					MDC_Globals.Sbo_Application.StatusBar.SetText("이미 저장되어져 있는 헤더의 내용과 일치합니다", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					return functionReturnValue;
//				}
//			}

//			//// 라인 ---------------------------
//			if (oMat1.VisualRowCount > 1) {
//				for (i = 1; i <= oMat1.VisualRowCount - 1; i++) {
//					////구분
//					//UPGRADE_WARNING: oMat1.Columns(PAYSEL).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("PAYSEL").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("지급구분은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("PAYSEL").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}
//					////시작월
//					//UPGRADE_WARNING: oMat1.Columns(STRMON).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("STRMON").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("STRMON").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}
//					////시작일
//					//UPGRADE_WARNING: oMat1.Columns(STRDAY).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("STRDAY").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("시작일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("STRDAY").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}
//					////지급월
//					//UPGRADE_WARNING: oMat1.Columns(JIGMON).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("JIGMON").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("지급월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("JIGMON").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}
//					////지급일
//					//UPGRADE_WARNING: oMat1.Columns(JIGDAY).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("JIGDAY").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("지급일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("JIGDAY").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}
//					////상여한도월
//					//UPGRADE_WARNING: oMat1.Columns(BNSEMM).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("BNSEMM").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("상여한도월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("BNSEMM").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}
//					////상여한도일
//					//UPGRADE_WARNING: oMat1.Columns(BNSEDD).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("BNSEDD").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("상여한도일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("BNSEDD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}
//					////급여한도월
//					//UPGRADE_WARNING: oMat1.Columns(PAYEMM).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("PAYEMM").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("급여한도월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("PAYEMM").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}
//					////급여한도일
//					//UPGRADE_WARNING: oMat1.Columns(PAYEDD).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("PAYEDD").Cells.Item(i).Specific.Value)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("급여한도일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("PAYEDD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						return functionReturnValue;
//					}

//					////중복체크
//					Chk_Data = Strings.Trim(oDS_PH_PY107B.GetValue("U_PAYSEL", i));
//					for (k = i + 1; k <= oMat1.VisualRowCount - 2; k++) {
//						oDS_PH_PY107B.Offset = k;
//						if (Strings.Trim(Chk_Data) == Strings.Trim(oDS_PH_PY107B.GetValue("U_PAYSEL", k))) {
//							MDC_Globals.Sbo_Application.SetStatusBarMessage("내용이 중복입력되었습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//							oMat1.Columns.Item("PAYSEL").Cells.Item(k + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//							return functionReturnValue;
//						}
//					}
//				}
//			} else {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			oMat1.FlushToDataSource();
//			//// Matrix 마지막 행 삭제(DB 저장시)
//			if (oDS_PH_PY107B.Size > 1)
//				oDS_PH_PY107B.RemoveRecord((oDS_PH_PY107B.Size - 1));
//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY107_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY107_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		public bool PH_PY107_Validate(string ValidateType)
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = true;
//			object i = null;
//			int j = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY107A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY107A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY107_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY107_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY107_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY107_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}
//	}
//}
