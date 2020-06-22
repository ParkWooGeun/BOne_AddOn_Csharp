using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using Microsoft.VisualBasic;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 호봉표등록
    /// </summary>
    internal class PH_PY105 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY105A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY105B;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        private bool CheckDataApply;  // 적용버턴 실행여부
        private string CLTCOD;        // 사업장
        private string YM;            // 적용연월

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY105.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }
                oFormUniqueID = "PH_PY105_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY105");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY105_CreateItems();
                PH_PY105_EnableMenus();
                PH_PY105_SetDocument(oFromDocEntry01);
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
        private void PH_PY105_CreateItems()
        {
            int i = 0;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY105A = oForm.DataSources.DBDataSources.Item("@PH_PY105A"); // 헤더
                oDS_PH_PY105B = oForm.DataSources.DBDataSources.Item("@PH_PY105B"); // 라인

                oMat1 = oForm.Items.Item("Mat1").Specific;
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                CheckDataApply = false;

                // 사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 직급
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P129' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat1.Columns.Item("JIGCOD").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("JIGCOD").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY105_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY105_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);  // 제거
                oForm.EnableMenu("1284", false); // 취소
                oForm.EnableMenu("1293", true);  // 행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY105_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// PH_PY105_SetDocument
        /// </summary>
        private void PH_PY105_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY105_FormItemEnabled();
                    PH_PY105_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY105_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY105_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY105_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = false;

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    // 귀속년월
                    oForm.Items.Item("YM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");

                    oForm.EnableMenu("1281", true);                    // 문서찾기
                    oForm.EnableMenu("1282", false);                   // 문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = false;

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false);                   // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = false;

                    // 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1281", true);                    // 문서찾기
                    oForm.EnableMenu("1282", true);                    // 문서추가

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY105_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY105_AddMatrixRow()
        {
            int oRow = 0;

            try
            {
                oForm.Freeze(true);
                // [Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY105B.GetValue("U_HOBCOD", oRow - 1))))
                    {
                        if (oDS_PH_PY105B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY105B.InsertRecord((oRow));
                        }
                        oDS_PH_PY105B.Offset = oRow;
                        oDS_PH_PY105B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY105B.SetValue("U_JIGCOD", oRow, "");
                        oDS_PH_PY105B.SetValue("U_HOBCOD", oRow, "");
                        oDS_PH_PY105B.SetValue("U_HOBNAM", oRow, "");
                        oDS_PH_PY105B.SetValue("U_STDAMT", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_BNSAMT", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow, Convert.ToString(0));
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY105B.Offset = oRow - 1;
                        oDS_PH_PY105B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY105B.SetValue("U_JIGCOD", oRow - 1, "");
                        oDS_PH_PY105B.SetValue("U_HOBCOD", oRow - 1, "");
                        oDS_PH_PY105B.SetValue("U_HOBNAM", oRow - 1, "");
                        oDS_PH_PY105B.SetValue("U_STDAMT", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_BNSAMT", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow - 1, Convert.ToString(0));
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY105B.Offset = oRow;
                    oDS_PH_PY105B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY105B.SetValue("U_JIGCOD", oRow, "");
                    oDS_PH_PY105B.SetValue("U_HOBCOD", oRow, "");
                    oDS_PH_PY105B.SetValue("U_HOBNAM", oRow, "");
                    oDS_PH_PY105B.SetValue("U_STDAMT", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_BNSAMT", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow, Convert.ToString(0));
                    oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow, Convert.ToString(0));
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY105_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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

                //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    //Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY105A", "Code");
                            // 접속자 권한에 따른 사업장 보기
                            break;

                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY105_FormItemEnabled();
                            PH_PY105_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281":                            // 문서찾기
                            PH_PY105_FormItemEnabled();
                            PH_PY105_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":                            // 문서추가
                            PH_PY105_FormItemEnabled();
                            PH_PY105_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY105_FormItemEnabled();
                            break;
                        case "1293":
                            // 행삭제
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while ((i <= oDS_PH_PY105B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY105B.GetValue("U_LineNum", i)))
                                    {
                                        oDS_PH_PY105B.RemoveRecord((i));
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }
                                for (i = 0; i <= oDS_PH_PY105B.Size; i++)
                                {
                                    oDS_PH_PY105B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }
                                oMat1.LoadFromDataSource();
                            }
                            PH_PY105_AddMatrixRow();
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;
            string FullName = string.Empty;
            string FindYN = string.Empty;

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY105_DataValidCheck() == false)
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
                        if (pVal.ActionSuccess == true)
                        {
                            if (CheckDataApply == true)
                            {
                                PH_PY105_DataApply(CLTCOD, YM);
                                CheckDataApply = false;
                            }
                            PH_PY105_FormItemEnabled();
                            PH_PY105_AddMatrixRow();
                        }
                    }
                    if (pVal.ItemUID == "Btn_UPLOAD")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY105_Excel_Upload);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    if (pVal.ItemUID == "Btn_Apply")
                    {
                        CLTCOD = Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0));
                        YM = Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0));
                        if (oMat1.RowCount > 1)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                CheckDataApply = true;
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                PH_PY105_DataApply(CLTCOD, YM);
                            }
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("호봉표 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                    PH_PY105_FormItemEnabled();
                    PH_PY105_AddMatrixRow();
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
        /// Raise_EVENT_KEY_DOWN
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
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
                        if (pVal.ItemUID == "Mat1" & pVal.ColUID == "HOBCOD")
                        {
                            PH_PY105_AddMatrixRow();
                            oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY105A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY105B);
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
        /// Raise_EVENT_CHOOSE_FROM_LIST 이벤트
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY105_DataValidCheck(string ItemUID)
        {
            bool functionReturnValue = false;
            int i = 0;
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                // 헤더 ---------------------------
                // 사업장
                if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0))))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                // 적용시작월
                if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0))))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("적용시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                // Code & Name 생성
                oDS_PH_PY105A.SetValue("Code", 0, Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0)));
                oDS_PH_PY105A.SetValue("NAME", 0, Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0)));

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY105A]", "'" + oDS_PH_PY105A.GetValue("Code", 0) + "'","")))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 존재하는 코드입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return functionReturnValue;
                    }
                }
                // 라인 ---------------------------
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        // 호봉코드
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("HOBCOD").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("호봉코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("HOBCOD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }
                        // 호봉명
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("HOBNAM").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("내역 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("HOBNAM").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
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

                // Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY105B.Size > 1)
                    oDS_PH_PY105B.RemoveRecord((oDS_PH_PY105B.Size - 1));

                oMat1.LoadFromDataSource();

                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY105_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                functionReturnValue = false;
                return functionReturnValue;
            }
            finally
            {
            }
        }

        ///// <summary>
        ///// PH_PY105_Excel_Upload 
        ///// </summary>
        //[STAThread]
        //private void PH_PY105_Excel_Upload()
        //{
        //    int i = 0;
        //    int ErrNum = 0;
        //    int TOTCNT = 0;
        //    int V_StatusCnt = 0;
        //    int oProValue = 0;
        //    int tRow = 0;
        //    string sPrice = string.Empty;
        //    string sFile = string.Empty;
        //    string OneRec = string.Empty;
        //    string sQry = string.Empty;

        //    short columnCount  = 15;  // 엑셀 컬럼수
        //    short columnCount2 = 15;  // 엑셀 컬럼수
        //    int loopCount = 0;

        //    PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

        //    CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

        //    commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xls;*.xlsx"));
        //    commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("모든 파일", "*.*"));

        //    if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
        //    {
        //        sFile = commonOpenFileDialog.FileName;
        //    }
        //    else //Cancel 버튼 클릭
        //    {
        //        return;
        //    }

        //    //엑셀 Object 연결
        //    //암시적 객체참조 시 Excel.exe 메모리 반환이 안됨, 아래와 같이 명시적 참조로 선언
        //    Microsoft.Office.Interop.Excel.ApplicationClass xlapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
        //    Microsoft.Office.Interop.Excel.Workbooks xlwbs = xlapp.Workbooks;
        //    Microsoft.Office.Interop.Excel.Workbook xlwb = xlwbs.Open(sFile);  // sFile
        //    Microsoft.Office.Interop.Excel.Sheets xlshs = xlwb.Worksheets;
        //    Microsoft.Office.Interop.Excel.Worksheet xlsh = (Microsoft.Office.Interop.Excel.Worksheet)xlshs[1];
        //    Microsoft.Office.Interop.Excel.Range xlCell = xlsh.Cells;
        //    Microsoft.Office.Interop.Excel.Range xlRange = xlsh.UsedRange;
        //    Microsoft.Office.Interop.Excel.Range xlRow = xlRange.Rows;

        //    SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    SAPbouiCOM.ProgressBar ProgressBar01 = null;

        //    try
        //    {
        //        oForm = PSH_Globals.SBO_Application.Forms.ActiveForm;
        //        if (string.IsNullOrEmpty(sFile))
        //        {
        //            ErrNum = 1;
        //            throw new Exception();
        //        }
        //        else
        //        {
        //            if ( codeHelpClass.Right(sFile, 3) == "xls" | codeHelpClass.Right(sFile, 4) == "xlsx" )
        //            {
        //                oDS_PH_PY105A.SetValue("U_Comments", 0, sFile);
        //            }
        //            else
        //            {
        //                ErrNum = 2;
        //                throw new Exception();
        //            }
        //        }

        //        Microsoft.Office.Interop.Excel.Range[] t = new Microsoft.Office.Interop.Excel.Range[columnCount2 + 1];
        //        for (loopCount = 1; loopCount <= columnCount2; loopCount++)
        //        {
        //            t[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[1, loopCount];
        //        }

        //        if (t[1].Value.ToString().Trim() != "직급")
        //        {
        //            ErrNum = 3;
        //            throw new Exception();
        //        }
        //        if (t[2].Value.ToString().Trim() != "호봉코드")
        //        {
        //            ErrNum = 4;
        //            throw new Exception();
        //        }
        //        if (t[3].Value.ToString().Trim() != "호봉명")
        //        {
        //            ErrNum = 5;
        //            throw new Exception();
        //        }
        //        if (t[4].Value.ToString().Trim() != "급여기본")
        //        {
        //            ErrNum = 6;
        //            throw new Exception();
        //        }
        //        if (t[5].Value.ToString().Trim() != "상여기본")
        //        {
        //            ErrNum = 7;
        //            throw new Exception();
        //        }
        //        if (t[6].Value.ToString().Trim() != "제수당01")
        //        {
        //            ErrNum = 8;
        //            throw new Exception();
        //        }
        //        if (t[7].Value.ToString().Trim() != "제수당02")
        //        {
        //            ErrNum = 9;
        //            throw new Exception();
        //        }
        //        if (t[8].Value.ToString().Trim() != "제수당03")
        //        {
        //            ErrNum = 10;
        //            throw new Exception();
        //        }
        //        if (t[9].Value.ToString().Trim() != "제수당04")
        //        {
        //            ErrNum = 11;
        //            throw new Exception();
        //        }
        //        if (t[10].Value.ToString().Trim() != "제수당05")
        //        {
        //            ErrNum = 12;
        //            throw new Exception();
        //        }
        //        if (t[11].Value.ToString().Trim() != "제수당06")
        //        {
        //            ErrNum = 13;
        //            throw new Exception();
        //        }
        //        if (t[12].Value.ToString().Trim() != "제수당07")
        //        {
        //            ErrNum = 14;
        //            throw new Exception();
        //        }
        //        if (t[13].Value.ToString().Trim() != "제수당08")
        //        {
        //            ErrNum = 15;
        //            throw new Exception();
        //        }
        //        if (t[14].Value.ToString().Trim() != "제수당09")
        //        {
        //            ErrNum = 16;
        //            throw new Exception();
        //        }
        //        if (t[15].Value.ToString().Trim() != "제수당10")
        //        {
        //            ErrNum = 17;
        //            throw new Exception();
        //        }

        //        for (loopCount = 1; loopCount <= columnCount2; loopCount++)
        //        {
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(t[loopCount]); //메모리 해제
        //        }

        //        ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("데이터 읽는중...!", 50, false);
        //        // 최대값 구하기
        //        TOTCNT = xlsh.UsedRange.Rows.Count - 1;
        //        V_StatusCnt = TOTCNT / 50;
        //        oProValue = 1;
        //        tRow = 1;

        //        // 테이블 생성
        //        sQry = "EXEC PH_PY105_TEMP_CHK";
        //        oRecordSet.DoQuery(sQry);

        //        for (i = 2; i <= xlsh.UsedRange.Rows.Count; i++)
        //        {
        //            Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

        //            for (loopCount = 1; loopCount <= columnCount; loopCount++)
        //            {
        //                r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[i, loopCount];
        //            }

        //            sQry = "EXEC PH_PY105 '" + r[1].Value.ToString().Trim() + "','" + r[2].Value.ToString().Trim() + "','" + r[3].Value.ToString().Trim() + "','" + r[4].Value.ToString().Trim() + "','" + r[5].Value.ToString().Trim() + "','";
        //            sQry = sQry + r[6].Value.ToString().Trim() + "','" + r[7].Value.ToString().Trim() + "','" + r[8].Value.ToString().Trim() + "','" + r[9].Value.ToString().Trim() + "','" + r[10].Value.ToString().Trim() + "','";
        //            sQry = sQry + r[11].Value.ToString().Trim() + "','" + r[12].Value.ToString().Trim() + "','" + r[13].Value.ToString().Trim() + "','" + r[14].Value.ToString().Trim() + "','" + r[15].Value.ToString().Trim() + "'";

        //            oRecordSet.DoQuery(sQry);

        //            if ((TOTCNT > 50 & tRow == oProValue * V_StatusCnt) | TOTCNT <= 50)
        //            {
        //                ProgressBar01.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
        //                oProValue = oProValue + 1;
        //                ProgressBar01.Value = oProValue;
        //            }
        //            tRow = tRow + 1;

        //            for (loopCount = 1; loopCount <= columnCount; loopCount++)
        //            {
        //                System.Runtime.InteropServices.Marshal.ReleaseComObject(r[loopCount]); //메모리 해제
        //            }
        //        }

        //        oMat1.Clear();
        //        oMat1.FlushToDataSource();

        //        // 임시데이터 데이타 검색
        //        sQry = "SELECT JIGCOD, HOBCOD, HOBNAM, STDAMT, BNSAMT, EXTAMT01, EXTAMT02, EXTAMT03, EXTAMT04, EXTAMT05, ";
        //        sQry = sQry + " EXTAMT06, EXTAMT07, EXTAMT08, EXTAMT09, EXTAMT10 FROM PH_PY105_TEMP ";
        //        oRecordSet.DoQuery(sQry);

        //        if (oRecordSet.RecordCount > 0)
        //        {
        //            for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
        //            {
        //                oDS_PH_PY105B.InsertRecord((i));
        //                oDS_PH_PY105B.Offset = i;
        //                oDS_PH_PY105B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //                oDS_PH_PY105B.SetValue("U_JIGCOD", i, oRecordSet.Fields.Item(0).Value);
        //                oDS_PH_PY105B.SetValue("U_HOBCOD", i, oRecordSet.Fields.Item(1).Value);
        //                oDS_PH_PY105B.SetValue("U_HOBNAM", i, oRecordSet.Fields.Item(2).Value);
        //                oDS_PH_PY105B.SetValue("U_STDAMT", i, oRecordSet.Fields.Item(3).Value);
        //                oDS_PH_PY105B.SetValue("U_BNSAMT", i, oRecordSet.Fields.Item(4).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT01", i, oRecordSet.Fields.Item(5).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT02", i, oRecordSet.Fields.Item(6).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT03", i, oRecordSet.Fields.Item(7).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT04", i, oRecordSet.Fields.Item(8).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT05", i, oRecordSet.Fields.Item(9).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT06", i, oRecordSet.Fields.Item(10).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT07", i, oRecordSet.Fields.Item(11).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT08", i, oRecordSet.Fields.Item(12).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT09", i, oRecordSet.Fields.Item(13).Value);
        //                oDS_PH_PY105B.SetValue("U_EXTAMT10", i, oRecordSet.Fields.Item(14).Value);
        //                oRecordSet.MoveNext();
        //            }

        //        }

        //        oMat1.LoadFromDataSource();
        //        PH_PY105_AddMatrixRow();

        //        PSH_Globals.SBO_Application.StatusBar.SetText("엑셀을 불러왔습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

        //        ProgressBar01.Stop();
        //    }
        //    catch (Exception ex)
        //    {
        //        if ((ProgressBar01 != null))
        //        {
        //            ProgressBar01.Stop();
        //            ProgressBar01 = null;
        //        }

        //        if (ErrNum == 1)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("파일을 선택해 주세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 2)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("엑셀파일이 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 3)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("A열 첫번째 행 타이틀은 직급", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 4)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("B열 두번째 행 타이틀은 호봉코드", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 5)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("C열 세번째 행 타이틀은 호봉명", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 6)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("D열 세번째 행 타이틀은 급여기본", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 7)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("E열 세번째 행 타이틀은 상여기본", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 8)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("F열 세번째 행 타이틀은 제수당01", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 9)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("G열 세번째 행 타이틀은 제수당02", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 10)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("H열 세번째 행 타이틀은 제수당03", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 11)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("I열 세번째 행 타이틀은 제수당04", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 12)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("J열 세번째 행 타이틀은 제수당05", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 13)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("K열 세번째 행 타이틀은 제수당06", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 14)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("L열 세번째 행 타이틀은 제수당07", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 15)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("M열 세번째 행 타이틀은 제수당08", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 16)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("N열 세번째 행 타이틀은 제수당09", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else if (ErrNum == 17)
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("O열 세번째 행 타이틀은 제수당10", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY105_Excel_Upload_ERROR" + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }

        //        //xlapp.Quit();
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
        //        //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);

        //    }
        //    finally
        //    {
        //        xlapp.Quit();
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);

        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
        //    }
        //}

        /// <summary>
        /// 엑셀자료 업로드
        /// </summary>
        [STAThread]
        private void PH_PY105_Excel_Upload()
        {
            int rowCount = 0;
            int loopCount = 0;
            string sFile = string.Empty;
            
            bool sucessFlag = false;
            short columnCount = 15; //엑셀 컬럼수

            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xls;*.xlsx"));
            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("모든 파일", "*.*"));

            if (commonOpenFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                sFile = commonOpenFileDialog.FileName;
            }
            else //Cancel 버튼 클릭
            {
                return;
            }

            if (string.IsNullOrEmpty(sFile))
            {
                //PSH_Globals.SBO_Application.StatusBar.SetText("파일을 선택해 주세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return;
            }
            else
            {
                oForm.Items.Item("Comments").Specific.VALUE = sFile;
            }

            //엑셀 Object 연결
            //암시적 객체참조 시 Excel.exe 메모리 반환이 안됨, 아래와 같이 명시적 참조로 선언
            Microsoft.Office.Interop.Excel.ApplicationClass xlapp = new Microsoft.Office.Interop.Excel.ApplicationClass();
            Microsoft.Office.Interop.Excel.Workbooks xlwbs = xlapp.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook xlwb = xlwbs.Open(sFile);
            Microsoft.Office.Interop.Excel.Sheets xlshs = xlwb.Worksheets;
            Microsoft.Office.Interop.Excel.Worksheet xlsh = (Microsoft.Office.Interop.Excel.Worksheet)xlshs[1];
            Microsoft.Office.Interop.Excel.Range xlCell = xlsh.Cells;
            Microsoft.Office.Interop.Excel.Range xlRange = xlsh.UsedRange;
            Microsoft.Office.Interop.Excel.Range xlRow = xlRange.Rows;

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("시작!", xlRow.Count, false);

            oForm.Freeze(true);

            oMat1.Clear();
            oMat1.FlushToDataSource();
            oMat1.LoadFromDataSource();

            try
            {
                for (rowCount = 2; rowCount <= xlRow.Count; rowCount++)
                {
                    if (rowCount - 2 != 0)
                    {
                        oDS_PH_PY105B.InsertRecord(rowCount - 2);
                    }

                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[rowCount, loopCount];
                    }

                    oDS_PH_PY105B.Offset = rowCount - 2;
                    oDS_PH_PY105B.SetValue("U_LineNum", rowCount - 2, Convert.ToString(rowCount - 1));
                    oDS_PH_PY105B.SetValue("U_JIGCOD", rowCount - 2, Convert.ToString(r[1].Value)); //직급
                    oDS_PH_PY105B.SetValue("U_HOBCOD", rowCount - 2, Convert.ToString(r[2].Value)); //호봉코드
                    oDS_PH_PY105B.SetValue("U_HOBNAM", rowCount - 2, Convert.ToString(r[3].Value)); //호봉명
                    oDS_PH_PY105B.SetValue("U_STDAMT", rowCount - 2, Convert.ToString(r[4].Value)); //급여기본
                    oDS_PH_PY105B.SetValue("U_BNSAMT", rowCount - 2, Convert.ToString(r[5].Value)); //상여기본
                    oDS_PH_PY105B.SetValue("U_EXTAMT01", rowCount - 2, Convert.ToString(r[6].Value)); //제수당01
                    oDS_PH_PY105B.SetValue("U_EXTAMT02", rowCount - 2, Convert.ToString(r[7].Value)); //제수당02
                    oDS_PH_PY105B.SetValue("U_EXTAMT03", rowCount - 2, Convert.ToString(r[8].Value)); //제수당03
                    oDS_PH_PY105B.SetValue("U_EXTAMT04", rowCount - 2, Convert.ToString(r[9].Value)); //제수당04
                    oDS_PH_PY105B.SetValue("U_EXTAMT05", rowCount - 2, Convert.ToString(r[10].Value)); //제수당05
                    oDS_PH_PY105B.SetValue("U_EXTAMT06", rowCount - 2, Convert.ToString(r[11].Value)); //제수당06
                    oDS_PH_PY105B.SetValue("U_EXTAMT07", rowCount - 2, Convert.ToString(r[12].Value)); //제수당07
                    oDS_PH_PY105B.SetValue("U_EXTAMT08", rowCount - 2, Convert.ToString(r[13].Value)); //제수당08
                    oDS_PH_PY105B.SetValue("U_EXTAMT09", rowCount - 2, Convert.ToString(r[14].Value)); //제수당09
                    oDS_PH_PY105B.SetValue("U_EXTAMT10", rowCount - 2, Convert.ToString(r[15].Value)); //제수당10

                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "건 Loding...!";

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(r[loopCount]); //메모리 해제
                    }
                }

                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();

                PH_PY105_AddMatrixRow();
                sucessFlag = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox("[PH_PY105_Excel_Upload_Error]" + (char)13 + ex.Message);
                sucessFlag = false;
            }
            finally
            {
                //액셀개체 닫음
                xlapp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRow);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCell);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsh);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlshs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlwbs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlapp);

                ProgressBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);

                if (sucessFlag == true)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("엑셀 Loding 완료", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }

                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY105_DataApply
        /// </summary>
        /// <returns></returns>
        private bool PH_PY105_DataApply(string CLTCOD, string YM)
        {
            bool functionReturnValue = false;
            string sQry = string.Empty;
            string Tablename = string.Empty;
            string sTablename = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oMat1.FlushToDataSource();
                sTablename = "@PH_PY001A_" + YM;
                // 조회용
                sQry = " SELECT Count(*) FROM SYSOBJECTS WHERE Name = '" + sTablename + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item(0).Value == 0)
                {
                    Tablename = "[@PH_PY001A_" + YM + "]";
                    // 테이블 생성용
                    sQry = "Exec PH_PY105_99 '" + Tablename + "'";
                    oRecordSet.DoQuery(sQry);
                }

                sQry = "";
                sQry = sQry + " Update T2 ";
                sQry = sQry + " SET T2.U_STDAMT = T1.U_STDAMT, T2.U_BNSAMT = T1.U_BNSAMT, T2.U_HOBYMM = T0.U_YM";
                sQry = sQry + " FROM [@PH_PY105A] T0";
                sQry = sQry + " INNER JOIN [@PH_PY105B] T1 ON T0.Code = T1.Code";
                sQry = sQry + " INNER JOIN [@PH_PY001A] T2 ON T2.U_JIGCOD = T1.U_JIGCOD AND T2.U_HOBONG = T1.U_HOBCOD";
                sQry = sQry + " WHERE T0.U_YM = '" + YM + "'";
                sQry = sQry + " And T2.U_status <> '5' ";
                sQry = sQry + " And Not Exists (Select * From [@PH_PY001A] T3 ";
                sQry = sQry + " Where T2.Code = T3.Code";
                sQry = sQry + " And T3.U_status <> '5'";
                sQry = sQry + " And dbo.PH_PY_PAYPEAK_YEAR(T3.U_CLTCOD,'" + YM + "',T3.Code) > 0 )";
                // 호봉등록 년월에 임금피크제대상은 임금조정을 안함.
                // sQry = sQry & " AND T0.U_YM = '" & YM & "'"
                oRecordSet.DoQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("해당 직급에 대헤 인사마스터에 금액이 적용 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY105_DataApply_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                functionReturnValue = false;
                return functionReturnValue;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// PH_PY105_DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY105_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            string sQry = string.Empty;
            string Tablename = string.Empty;
            string sTablename = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                // 헤더
                // 사업장
                if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0))))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                // 적용시작월
                if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0))))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("적용시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                // Code & Name 생성
                oDS_PH_PY105A.SetValue("Code", 0, Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0)));
                oDS_PH_PY105A.SetValue("NAME", 0, Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0)));

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY105A]", "'" + oDS_PH_PY105A.GetValue("Code", 0) + "'","")))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 존재하는 코드입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return functionReturnValue;
                    }
                }

                // 라인
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        // 호봉코드
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("HOBCOD").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("호봉코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("HOBCOD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
                        }
                        // 호봉명
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("HOBNAM").Cells.Item(i).Specific.VALUE))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("내역 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("HOBNAM").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            functionReturnValue = false;
                            return functionReturnValue;
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

                // Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY105B.Size > 1)
                    oDS_PH_PY105B.RemoveRecord((oDS_PH_PY105B.Size - 1));

                oMat1.LoadFromDataSource();

                functionReturnValue = true;
                return functionReturnValue;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY105_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                functionReturnValue = false;
                return functionReturnValue;
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
//	internal class PH_PY105
//	{
//////********************************************************************************
//////  File           : PH_PY105.cls
//////  Module         : 급여관리 > 급여관리
//////  Desc           : 호봉표등록
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY105A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY105B;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//			////적용버턴 실행여부
//		private bool CheckDataApply;
//			////사업장
//		private string CLTCOD;
//			////적용연월
//		private string YM;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY105.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY105_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY105");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			oForm.DataBrowser.BrowseBy = "Code";

//			oForm.Freeze(true);
//			PH_PY105_CreateItems();
//			PH_PY105_EnableMenus();
//			PH_PY105_SetDocument(oFromDocEntry01);
//			//    Call PH_PY105_FormResize

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

//		private bool PH_PY105_CreateItems()
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

//			oDS_PH_PY105A = oForm.DataSources.DBDataSources("@PH_PY105A");
//			oDS_PH_PY105B = oForm.DataSources.DBDataSources("@PH_PY105B");


//			oMat1 = oForm.Items.Item("Mat1").Specific;
//			////@PH_PY105B

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//			oMat1.AutoResizeColumns();

//			CheckDataApply = false;

//			//// 사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			//    sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'"
//			//    Call SetReDataCombo(oForm, sQry, oCombo)
//			//    oCombo.Select 0, psk_Index
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;




//			//// 직급
//			oColumn = oMat1.Columns.Item("JIGCOD");
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P129' AND U_UseYN= 'Y'";
//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount > 0) {
//				for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//					oColumn.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
//					oRecordSet.MoveNext();
//				}
//			}
//			oColumn.DisplayDesc = true;

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
//			PH_PY105_CreateItems_Error:

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
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY105_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", true);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", true);
//			////행삭제

//			return;
//			PH_PY105_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY105_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY105_FormItemEnabled();
//				PH_PY105_AddMatrixRow();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY105_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY105_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY105_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;

//			 // ERROR: Not supported in C#: OnErrorStatement



//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				oForm.Items.Item("YM").Enabled = true;
//				oForm.Items.Item("Comments").Enabled = false;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");
//				/// 귀속년월
//				//UPGRADE_WARNING: oForm.Items(YM).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("YM").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMM");

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = true;
//				oForm.Items.Item("YM").Enabled = true;
//				oForm.Items.Item("Comments").Enabled = false;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가
//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				oForm.Items.Item("CLTCOD").Enabled = false;
//				oForm.Items.Item("YM").Enabled = false;
//				oForm.Items.Item("Comments").Enabled = false;

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY105_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//								if (PH_PY105_DataValidCheck() == false) {
//									BubbleEvent = false;
//								}
//							}
//						}

//					} else if (pval.BeforeAction == false) {
//						if (pval.ItemUID == "1") {
//							if (pval.ActionSuccess == true) {
//								if (CheckDataApply == true) {
//									PH_PY105_DataApply(ref CLTCOD, ref YM);
//									CheckDataApply = false;
//								}
//								PH_PY105_FormItemEnabled();
//								PH_PY105_AddMatrixRow();
//							}
//						}
//						if (pval.ItemUID == "Btn_UPLOAD") {
//							PH_PY105_Excel_Upload();
//						}
//						if (pval.ItemUID == "Btn_Apply") {
//							CLTCOD = Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0));
//							YM = Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0));
//							if (oMat1.RowCount > 1) {
//								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//									CheckDataApply = true;
//									oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//								} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//									PH_PY105_DataApply(ref CLTCOD, ref YM);
//								}
//							} else {
//								MDC_Globals.Sbo_Application.SetStatusBarMessage("호봉표 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
//								if (pval.Row > 0) {
//									oMat1.SelectRow(pval.Row, true, false);
//								}
//								break;
//						}

//						switch (pval.ItemUID) {
//							case "Mat1":
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
//							if (pval.ItemUID == "Mat1" & pval.ColUID == "HOBCOD") {
//								PH_PY105_AddMatrixRow();
//								oMat1.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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

//						PH_PY105_FormItemEnabled();
//						PH_PY105_AddMatrixRow();
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
//						//UPGRADE_NOTE: oDS_PH_PY105A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY105A = null;
//						//UPGRADE_NOTE: oDS_PH_PY105B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY105B = null;

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
//						//                If pval.ItemUID = "Code" Then
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY105A", "Code")
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
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string sQry = null;
//			int i = 0;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
//						MDC_SetMod.AuthorityCheck(ref oForm, ref "CLTCOD", ref "@PH_PY105A", ref "Code");
//						////접속자 권한에 따른 사업장 보기
//						break;

//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY105_FormItemEnabled();
//						PH_PY105_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY105_FormItemEnabled();
//						PH_PY105_AddMatrixRow();
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY105_FormItemEnabled();
//						PH_PY105_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY105_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent, ref oMat1, ref oDS_PH_PY105B, ref "U_JIGCOD");
//						PH_PY105_AddMatrixRow();
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

//		public void PH_PY105_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY105B.GetValue("U_HOBCOD", oRow - 1)))) {
//					if (oDS_PH_PY105B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY105B.InsertRecord((oRow));
//					}
//					oDS_PH_PY105B.Offset = oRow;
//					oDS_PH_PY105B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					oDS_PH_PY105B.SetValue("U_JIGCOD", oRow, "");
//					oDS_PH_PY105B.SetValue("U_HOBCOD", oRow, "");
//					oDS_PH_PY105B.SetValue("U_HOBNAM", oRow, "");
//					oDS_PH_PY105B.SetValue("U_STDAMT", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_BNSAMT", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow, Convert.ToString(0));
//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY105B.Offset = oRow - 1;
//					oDS_PH_PY105B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					oDS_PH_PY105B.SetValue("U_JIGCOD", oRow - 1, "");
//					oDS_PH_PY105B.SetValue("U_HOBCOD", oRow - 1, "");
//					oDS_PH_PY105B.SetValue("U_HOBNAM", oRow - 1, "");
//					oDS_PH_PY105B.SetValue("U_STDAMT", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_BNSAMT", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow - 1, Convert.ToString(0));
//					oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow - 1, Convert.ToString(0));
//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY105B.Offset = oRow;
//				oDS_PH_PY105B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				oDS_PH_PY105B.SetValue("U_JIGCOD", oRow, "");
//				oDS_PH_PY105B.SetValue("U_HOBCOD", oRow, "");
//				oDS_PH_PY105B.SetValue("U_HOBNAM", oRow, "");
//				oDS_PH_PY105B.SetValue("U_STDAMT", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_BNSAMT", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT01", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT02", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT03", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT04", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT05", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT06", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT07", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT08", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT09", oRow, Convert.ToString(0));
//				oDS_PH_PY105B.SetValue("U_EXTAMT10", oRow, Convert.ToString(0));
//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY105_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY105_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY105'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY105_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY105_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//// 헤더 ---------------------------
//			////사업장
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			////적용시작월
//			if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0)))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("적용시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//// Code & Name 생성
//			oDS_PH_PY105A.SetValue("Code", 0, Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0)));
//			oDS_PH_PY105A.SetValue("NAME", 0, Strings.Trim(oDS_PH_PY105A.GetValue("U_CLTCOD", 0)) + Strings.Trim(oDS_PH_PY105A.GetValue("U_YM", 0)));

//			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//				//UPGRADE_WARNING: MDC_SetMod.Get_ReData(Code, Code, [PH_PY105A], ' & oDS_PH_PY105A.GetValue(Code, 0) & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				if (!string.IsNullOrEmpty(MDC_SetMod.Get_ReData("Code", "Code", "[@PH_PY105A]", "'" + oDS_PH_PY105A.GetValue("Code", 0) + "'"))) {
//					MDC_Globals.Sbo_Application.SetStatusBarMessage("이미 존재하는 코드입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//					return functionReturnValue;
//				}
//			}

//			//// 라인 ---------------------------
//			if (oMat1.VisualRowCount > 1) {
//				for (i = 1; i <= oMat1.VisualRowCount - 1; i++) {
//					////호봉코드
//					//UPGRADE_WARNING: oMat1.Columns(HOBCOD).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("HOBCOD").Cells.Item(i).Specific.VALUE)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("호봉코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("HOBCOD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						functionReturnValue = false;
//						return functionReturnValue;
//					}
//					////호봉명
//					//UPGRADE_WARNING: oMat1.Columns(HOBNAM).Cells(i).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//					if (string.IsNullOrEmpty(oMat1.Columns.Item("HOBNAM").Cells.Item(i).Specific.VALUE)) {
//						MDC_Globals.Sbo_Application.SetStatusBarMessage("내역 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//						oMat1.Columns.Item("HOBNAM").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						functionReturnValue = false;
//						return functionReturnValue;
//					}
//				}
//			} else {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			oMat1.FlushToDataSource();

//			//// Matrix 마지막 행 삭제(DB 저장시)
//			if (oDS_PH_PY105B.Size > 1)
//				oDS_PH_PY105B.RemoveRecord((oDS_PH_PY105B.Size - 1));

//			oMat1.LoadFromDataSource();


//			functionReturnValue = true;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY105_DataValidCheck_Error:




//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY105_MTX01()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;

//			string Param01 = null;
//			string Param02 = null;
//			string Param03 = null;
//			string Param04 = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param01 = oForm.Items.Item("Param01").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param02 = oForm.Items.Item("Param01").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param03 = oForm.Items.Item("Param01").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Param04 = oForm.Items.Item("Param01").Specific.VALUE;

//			sQry = "SELECT 10";
//			oRecordSet.DoQuery(sQry);

//			oMat1.Clear();
//			oMat1.FlushToDataSource();
//			oMat1.LoadFromDataSource();

//			if ((oRecordSet.RecordCount == 0)) {
//				MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
//				goto PH_PY105_MTX01_Exit;
//			}

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//			for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//				if (i != 0) {
//					oDS_PH_PY105B.InsertRecord((i));
//				}
//				oDS_PH_PY105B.Offset = i;
//				oDS_PH_PY105B.SetValue("U_COL01", i, oRecordSet.Fields.Item(0).Value);
//				oDS_PH_PY105B.SetValue("U_COL02", i, oRecordSet.Fields.Item(1).Value);
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
//			PH_PY105_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			if ((ProgressBar01 != null)) {
//				ProgressBar01.Stop();
//			}
//			return;
//			PH_PY105_MTX01_Error:
//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY105_Validate(string ValidateType)
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
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY105A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY105A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY105_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY105_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY105_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//////행삭제 (FormUID, pval, BubbleEvent, 매트릭스 이름, 디비데이터소스, 데이터 체크 필드명)
//		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent, ref SAPbouiCOM.Matrix oMat, ref SAPbouiCOM.DBDataSource DBData, ref string CheckField)
//		{

//			int i = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((oLastColRow > 0)) {
//				if (pval.BeforeAction == true) {

//				} else if (pval.BeforeAction == false) {
//					if (oMat.RowCount != oMat.VisualRowCount) {
//						oMat.FlushToDataSource();

//						while ((i <= DBData.Size - 1)) {
//							if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i))) {
//								DBData.RemoveRecord((i));
//								i = 0;
//							} else {
//								i = i + 1;
//							}
//						}

//						for (i = 0; i <= DBData.Size; i++) {
//							DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//						}

//						oMat.LoadFromDataSource();
//					}
//				}
//			}
//			return;
//			Raise_EVENT_ROW_DELETE_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		private void PH_PY105_Excel_Upload()
//		{

//			int i = 0;
//			int j = 0;
//			string sPrice = null;
//			string sFile = null;
//			string OneRec = null;
//			string sQry = null;

//			Microsoft.Office.Interop.Excel.Application xl = default(Microsoft.Office.Interop.Excel.Application);
//			Microsoft.Office.Interop.Excel.Workbook xlwb = default(Microsoft.Office.Interop.Excel.Workbook);
//			Microsoft.Office.Interop.Excel.Worksheet xlsh = default(Microsoft.Office.Interop.Excel.Worksheet);

//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.Form oForm = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			oForm = MDC_Globals.Sbo_Application.Forms.ActiveForm;


//			//// 임시 테이블 생성
//			//    sQry = "EXEC MDC_MM_TEMP_CPD102 "
//			//
//			//    oRecordset.DoQuery sQry

//			//UPGRADE_WARNING: FileListBoxForm.OpenDialog() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sFile = My.MyProject.Forms.FileListBoxForm.OpenDialog(ref FileListBoxForm, ref "*.xls", ref "파일선택", ref "C:\\");

//			if (string.IsNullOrEmpty(sFile)) {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("파일을 선택해 주세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				return;
//			} else {
//				if (Strings.Mid(Strings.Right(sFile, 4), 1, 3) == "xls" | Strings.Mid(Strings.Right(sFile, 5), 1, 4) == "xlsx") {
//					oDS_PH_PY105A.SetValue("U_Comments", 0, sFile);
//				} else {
//					MDC_Globals.Sbo_Application.StatusBar.SetText("엑셀파일이 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//					return;
//				}
//			}

//			//엑셀 Object 연결
//			xl = Interaction.CreateObject("excel.application");
//			xlwb = xl.Workbooks.Open(sFile, , true);
//			xlsh = xlwb.Worksheets("호봉등록");

//			//UPGRADE_WARNING: xlsh.Cells(1, 1).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 1).VALUE != "직급") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("A열 첫번째 행 타이틀은 직급", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//UPGRADE_WARNING: xlsh.Cells(1, 2).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 2).VALUE != "호봉코드") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("B열 두번째 행 타이틀은 호봉코드", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//UPGRADE_WARNING: xlsh.Cells(1, 3).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 3).VALUE != "호봉명") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("C열 세번째 행 타이틀은 호봉명", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//UPGRADE_WARNING: xlsh.Cells(1, 4).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 4).VALUE != "급여기본") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("D열 세번째 행 타이틀은 급여기본", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//UPGRADE_WARNING: xlsh.Cells(1, 5).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 5).VALUE != "상여기본") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("E열 세번째 행 타이틀은 상여기본", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//UPGRADE_WARNING: xlsh.Cells(1, 6).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 6).VALUE != "제수당01") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("F열 세번째 행 타이틀은 제수당01", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			//UPGRADE_WARNING: xlsh.Cells(1, 7).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 7).VALUE != "제수당02") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("G열 세번째 행 타이틀은 제수당02", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}
//			//UPGRADE_WARNING: xlsh.Cells(1, 8).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 8).VALUE != "제수당03") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("H열 세번째 행 타이틀은 제수당03", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}
//			//UPGRADE_WARNING: xlsh.Cells(1, 9).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 9).VALUE != "제수당04") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("I열 세번째 행 타이틀은 제수당04", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}
//			//UPGRADE_WARNING: xlsh.Cells(1, 10).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 10).VALUE != "제수당05") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("J열 세번째 행 타이틀은 제수당05", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}
//			//UPGRADE_WARNING: xlsh.Cells(1, 11).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 11).VALUE != "제수당06") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("K열 세번째 행 타이틀은 제수당06", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}
//			//UPGRADE_WARNING: xlsh.Cells(1, 12).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 12).VALUE != "제수당07") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("L열 세번째 행 타이틀은 제수당07", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}
//			//UPGRADE_WARNING: xlsh.Cells(1, 13).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 13).VALUE != "제수당08") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("M열 세번째 행 타이틀은 제수당08", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}
//			//UPGRADE_WARNING: xlsh.Cells(1, 14).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 14).VALUE != "제수당09") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("N열 세번째 행 타이틀은 제수당09", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}
//			//UPGRADE_WARNING: xlsh.Cells(1, 15).VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (xlsh.Cells._Default(1, 15).VALUE != "제수당10") {
//				MDC_Globals.Sbo_Application.StatusBar.SetText("O열 세번째 행 타이틀은 제수당10", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
//				xlwb.Close();
//				//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlwb = null;
//				//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xl = null;
//				//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				xlsh = null;
//				return;
//			}

//			////테이블 생성
//			sQry = "EXEC PH_PY105_TEMP_CHK";
//			oRecordSet.DoQuery(sQry);

//			for (i = 2; i <= xlsh.UsedRange.Rows.Count; i++) {
//				//UPGRADE_WARNING: xlsh.Cells(i, 5) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 4) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 3) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 2) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = "EXEC PH_PY105 '" + xlsh.Cells._Default(i, 1) + "','" + xlsh.Cells._Default(i, 2) + "','" + xlsh.Cells._Default(i, 3) + "','" + xlsh.Cells._Default(i, 4) + "','" + xlsh.Cells._Default(i, 5) + "','";
//				//UPGRADE_WARNING: xlsh.Cells(i, 10) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 9) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 8) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 7) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + xlsh.Cells._Default(i, 6) + "','" + xlsh.Cells._Default(i, 7) + "','" + xlsh.Cells._Default(i, 8) + "','" + xlsh.Cells._Default(i, 9) + "','" + xlsh.Cells._Default(i, 10) + "','";
//				//UPGRADE_WARNING: xlsh.Cells(i, 15) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 14) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 13) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells(i, 12) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				//UPGRADE_WARNING: xlsh.Cells() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				sQry = sQry + xlsh.Cells._Default(i, 11) + "','" + xlsh.Cells._Default(i, 12) + "','" + xlsh.Cells._Default(i, 13) + "','" + xlsh.Cells._Default(i, 14) + "','" + xlsh.Cells._Default(i, 15) + "'";

//				oRecordSet.DoQuery(sQry);
//			}

//			oMat1.Clear();
//			oMat1.FlushToDataSource();

//			//// 임시데이터 데이타 검색
//			sQry = "SELECT JIGCOD, HOBCOD, HOBNAM, STDAMT, BNSAMT, EXTAMT01, EXTAMT02, EXTAMT03, EXTAMT04, EXTAMT05, ";
//			sQry = sQry + " EXTAMT06, EXTAMT07, EXTAMT08, EXTAMT09, EXTAMT10 FROM PH_PY105_TEMP ";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.RecordCount > 0) {
//				for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//					oDS_PH_PY105B.InsertRecord((i));
//					oDS_PH_PY105B.Offset = i;
//					oDS_PH_PY105B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//					oDS_PH_PY105B.SetValue("U_JIGCOD", i, oRecordSet.Fields.Item(0).Value);
//					oDS_PH_PY105B.SetValue("U_HOBCOD", i, oRecordSet.Fields.Item(1).Value);
//					oDS_PH_PY105B.SetValue("U_HOBNAM", i, oRecordSet.Fields.Item(2).Value);
//					oDS_PH_PY105B.SetValue("U_STDAMT", i, oRecordSet.Fields.Item(3).Value);
//					oDS_PH_PY105B.SetValue("U_BNSAMT", i, oRecordSet.Fields.Item(4).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT01", i, oRecordSet.Fields.Item(5).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT02", i, oRecordSet.Fields.Item(6).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT03", i, oRecordSet.Fields.Item(7).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT04", i, oRecordSet.Fields.Item(8).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT05", i, oRecordSet.Fields.Item(9).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT06", i, oRecordSet.Fields.Item(10).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT07", i, oRecordSet.Fields.Item(11).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT08", i, oRecordSet.Fields.Item(12).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT09", i, oRecordSet.Fields.Item(13).Value);
//					oDS_PH_PY105B.SetValue("U_EXTAMT10", i, oRecordSet.Fields.Item(14).Value);
//					oRecordSet.MoveNext();
//				}

//			}

//			oMat1.LoadFromDataSource();
//			PH_PY105_AddMatrixRow();

//			MDC_Globals.Sbo_Application.StatusBar.SetText("엑셀을 불러왔습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);


//			//액셀개체 닫음
//			xlwb.Close();
//			//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlwb = null;
//			//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xl = null;
//			//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlsh = null;

//			//진행바 초기화
//			return;
//			Err_Renamed:

//			xlwb.Close();
//			//UPGRADE_NOTE: xlwb 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlwb = null;
//			//UPGRADE_NOTE: xl 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xl = null;
//			//UPGRADE_NOTE: xlsh 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			xlsh = null;
//		}


//		private bool PH_PY105_DataApply(ref string CLTCOD, ref string YM)
//		{
//			bool functionReturnValue = false;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;
//			string Tablename = null;
//			string sTablename = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			 // ERROR: Not supported in C#: OnErrorStatement


//			functionReturnValue = false;

//			oMat1.FlushToDataSource();

//			sTablename = "@PH_PY001A_" + YM;
//			//조회용


//			sQry = " SELECT Count(*) FROM SYSOBJECTS WHERE Name = '" + sTablename + "'";
//			oRecordSet.DoQuery(sQry);


//			if (oRecordSet.Fields.Item(0).Value == 0) {
//				Tablename = "[@PH_PY001A_" + YM + "]";
//				////테이블 생성용
//				sQry = "Exec PH_PY105_99 '" + Tablename + "'";
//				oRecordSet.DoQuery(sQry);
//			}


//			sQry = "";
//			sQry = sQry + " Update T2 ";
//			sQry = sQry + " SET T2.U_STDAMT = T1.U_STDAMT, T2.U_BNSAMT = T1.U_BNSAMT, T2.U_HOBYMM = T0.U_YM";
//			sQry = sQry + " FROM [@PH_PY105A] T0";
//			sQry = sQry + " INNER JOIN [@PH_PY105B] T1 ON T0.Code = T1.Code";
//			sQry = sQry + " INNER JOIN [@PH_PY001A] T2 ON T2.U_JIGCOD = T1.U_JIGCOD AND T2.U_HOBONG = T1.U_HOBCOD";
//			// AND T2.U_CLTCOD = T0.U_CLTCOD"
//			sQry = sQry + " WHERE T0.U_YM = '" + YM + "'";
//			sQry = sQry + " And T2.U_status <> '5' ";
//			sQry = sQry + " And Not Exists (Select * From [@PH_PY001A] T3 ";
//			sQry = sQry + " Where T2.Code = T3.Code";
//			sQry = sQry + " And T3.U_status <> '5'";
//			sQry = sQry + " And dbo.PH_PY_PAYPEAK_YEAR(T3.U_CLTCOD,'" + YM + "',T3.Code) > 0 )";
//			//호봉등록 년월에 임금피크제대상은 임금조정을 안함.
//			//sQry = sQry & " AND T0.U_YM = '" & YM & "'"
//			oRecordSet.DoQuery(sQry);

//			//    Sbo_Application.SetStatusBarMessage "해당 직급에 대헤 인사마스터에 금액이 적용 되었습니다." & Err.Description, bmt_Short, False
//			MDC_Globals.Sbo_Application.StatusBar.SetText("해당 직급에 대헤 인사마스터에 금액이 적용 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY105_DataApply_Error:

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY105_DataApply_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}
//	}
//}
