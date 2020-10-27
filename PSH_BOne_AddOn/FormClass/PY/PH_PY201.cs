using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 정년임박자 휴가경비 등록
    /// </summary>
    internal class PH_PY201 : PSH_BaseClass
    {
        public string oFormUniqueID01;
        //public SAPbouiCOM.Form oForm;

        public SAPbouiCOM.Matrix oMat1;

        private SAPbouiCOM.DBDataSource oDS_PH_PY201A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY201B;

        private string oLastItemUID01;
        private string oLastColUID01;
        private int oLastColRow01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY201.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY201_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY201");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PH_PY201_CreateItems();
                PH_PY201_EnableMenus();
                PH_PY201_SetDocument(oFormDocEntry01);
                ////PH_PY201_FormResize();
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
        private void PH_PY201_CreateItems()
        {
            try
            {
                oDS_PH_PY201A = oForm.DataSources.DBDataSources.Item("@PH_PY201A");
                oDS_PH_PY201B = oForm.DataSources.DBDataSources.Item("@PH_PY201B");

                oMat1 = oForm.Items.Item("Mat01").Specific;
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMat1.AutoResizeColumns();

                oForm.Items.Item("CLTCOD").DisplayDesc = true;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY201_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY201_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1287", true); // 복제
                //oForm.EnableMenu("1286", True); // 닫기
                oForm.EnableMenu("1284", true); // 취소
                oForm.EnableMenu("1293", true); // 행삭제
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY201_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY201_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY201_FormItemEnabled();
                    PH_PY201_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY201_FormItemEnabled();

                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY201_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {

            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY201_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Comment").Enabled = true;

                    //폼 DocEntry 세팅
                    PH_PY201_FormClear();

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가

                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("Comment").Enabled = false;

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Comment").Enabled = true;
                    // oMat1.Columns("MSTCOD").Editable = False

                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); //접속자에 따른 권한별 사업장 콤보박스세팅

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY201_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메트릭스 Row 추가
        /// </summary>
        private void PH_PY201_AddMatrixRow()
        {
            int oRow = 0;

            try
            {
                oForm.Freeze(true);

                //[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY201B.GetValue("U_MSTCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY201B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY201B.InsertRecord(oRow);
                        }
                        oDS_PH_PY201B.Offset = oRow;
                        oDS_PH_PY201B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY201B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY201B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY201B.SetValue("U_TeamCode", oRow, "");
                        oDS_PH_PY201B.SetValue("U_TeamName", oRow, "");
                        oDS_PH_PY201B.SetValue("U_RspCode", oRow, "");
                        oDS_PH_PY201B.SetValue("U_RspName", oRow, "");
                        oDS_PH_PY201B.SetValue("U_ClsCode", oRow, "");
                        oDS_PH_PY201B.SetValue("U_ClsName", oRow, "");
                        oDS_PH_PY201B.SetValue("U_Amount", oRow, Convert.ToString(0));
                        oDS_PH_PY201B.SetValue("U_Comment", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY201B.Offset = oRow - 1;
                        oDS_PH_PY201B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY201B.SetValue("U_MSTCOD", oRow - 1, "");
                        oDS_PH_PY201B.SetValue("U_MSTNAM", oRow - 1, "");
                        oDS_PH_PY201B.SetValue("U_TeamCode", oRow - 1, "");
                        oDS_PH_PY201B.SetValue("U_TeamName", oRow - 1, "");
                        oDS_PH_PY201B.SetValue("U_RspCode", oRow - 1, "");
                        oDS_PH_PY201B.SetValue("U_RspName", oRow - 1, "");
                        oDS_PH_PY201B.SetValue("U_ClsCode", oRow - 1, "");
                        oDS_PH_PY201B.SetValue("U_ClsName", oRow - 1, "");
                        oDS_PH_PY201B.SetValue("U_Amount", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY201B.SetValue("U_Comment", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY201B.Offset = oRow;
                    oDS_PH_PY201B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY201B.SetValue("U_MSTCOD", oRow, "");
                    oDS_PH_PY201B.SetValue("U_MSTNAM", oRow, "");
                    oDS_PH_PY201B.SetValue("U_TeamCode", oRow, "");
                    oDS_PH_PY201B.SetValue("U_TeamName", oRow, "");
                    oDS_PH_PY201B.SetValue("U_RspCode", oRow, "");
                    oDS_PH_PY201B.SetValue("U_RspName", oRow, "");
                    oDS_PH_PY201B.SetValue("U_ClsCode", oRow, "");
                    oDS_PH_PY201B.SetValue("U_ClsName", oRow, "");
                    oDS_PH_PY201B.SetValue("U_Amount", oRow, Convert.ToString(0));
                    oDS_PH_PY201B.SetValue("U_Comment", oRow, "");
                    oMat1.LoadFromDataSource();
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY201_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// FormClear : DocEntry Setting
        /// </summary>
        private void PH_PY201_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY201'", "");

                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY201_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck : 입력데이터의 Valid Check
        /// </summary>
        /// <returns></returns>
        private bool PH_PY201_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            short ErrNum = 0;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY201A.GetValue("U_CLTCOD", 0).ToString().Trim())) //사업장
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oDS_PH_PY201A.GetValue("U_DocDate", 0).ToString().Trim())) //지급일자
                {
                    ErrNum = 2;
                    throw new Exception();
                }

                //라인
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("MSTCOD").Cells.Item(i).Specific.Value)) //사번
                        {
                            ErrNum = 3;
                            throw new Exception();
                        }

                        if (oMat1.Columns.Item("Amount").Cells.Item(i).Specific.Value == "0") //금액
                        {
                            ErrNum = 4;
                            throw new Exception();
                        }
                    }
                }
                else
                {
                    ErrNum = 5;
                    throw new Exception();
                }

                oMat1.FlushToDataSource();
                if (oDS_PH_PY201B.Size > 1)
                {
                    oDS_PH_PY201B.RemoveRecord(oDS_PH_PY201B.Size - 1); // Matrix 마지막 행 삭제(DB 저장시)
                }
                oMat1.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                functionReturnValue = false;

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급일자는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage(i + 1 + "번째 라인의 사번은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oMat1.Columns.Item("MSTCOD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 4)
                {

                    PSH_Globals.SBO_Application.SetStatusBarMessage(i + 1 + "번째 금액은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oMat1.Columns.Item("Amount").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (ErrNum == 5)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY201_DataValidCheck_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }

            return functionReturnValue;
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

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                ////case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                ////    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                ////case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                ////    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

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
                ////    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                ////    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                ////    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //    //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    //    break;

                //    //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    //    break;
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
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY201_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY201_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
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
                                PH_PY201_FormItemEnabled();
                                PH_PY201_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY201_FormItemEnabled();
                                PH_PY201_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY201_FormItemEnabled();
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
                oForm.Freeze(false);
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "MSTCOD" && pVal.CharPressed == 9)
                        {
                            if (string.IsNullOrEmpty(oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat1.AutoResizeColumns();
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
                    //switch (pVal.ItemUID)
                    //{
                    //    case "Mat01":
                    //        if (pVal.Row > 0)
                    //        {
                    //            oMat1.SelectRow(pVal.Row, true, false);
                    //        }
                    //        break;
                    //}

                    switch (pVal.ItemUID)
                    {
                        case "Mat01":
                            if (pVal.Row > 0)
                            {
                                oMat1.SelectRow(pVal.Row, true, false);

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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_DOUBLE_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LINK_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "CntcCode":
                                oDS_PH_PY201A.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));
                                break;

                            case "Mat01":

                                if (pVal.ColUID == "MSTCOD")
                                {
                                    oMat1.FlushToDataSource();

                                    oDS_PH_PY201B.SetValue("U_MSTNAM", pVal.Row - 1, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value + "'", ""));
                                    oDS_PH_PY201B.SetValue("U_TeamCode", pVal.Row - 1, dataHelpClass.Get_ReData("U_TeamCode", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value + "'", ""));
                                    oDS_PH_PY201B.SetValue("U_TeamName", pVal.Row - 1, dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oDS_PH_PY201B.GetValue("U_TeamCode", pVal.Row - 1) + "'", " AND Code = '1'"));
                                    oDS_PH_PY201B.SetValue("U_RspCode", pVal.Row - 1, dataHelpClass.Get_ReData("U_RspCode", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value + "'", ""));
                                    oDS_PH_PY201B.SetValue("U_RspName", pVal.Row - 1, dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oDS_PH_PY201B.GetValue("U_RspCode", pVal.Row - 1) + "'", " AND Code = '2'"));
                                    oDS_PH_PY201B.SetValue("U_ClsCode", pVal.Row - 1, dataHelpClass.Get_ReData("U_ClsCode", "Code", "[@PH_PY001A]", "'" + oMat1.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value + "'", ""));
                                    oDS_PH_PY201B.SetValue("U_ClsName", pVal.Row - 1, dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oDS_PH_PY201B.GetValue("U_ClsCode", pVal.Row - 1) + "'", " AND Code = '9'"));
                                    oDS_PH_PY201B.SetValue("U_Amount", pVal.Row - 1, dataHelpClass.Get_ReData("TOP 1 U_Num1", "Code", "[@PS_HR200l]", "'P238' ORDER BY U_Num1 DESC", ""));

                                    oMat1.LoadFromDataSource();

                                    PH_PY201_AddMatrixRow();
                                }

                                oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oMat1.AutoResizeColumns();
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat1.LoadFromDataSource();
                    PH_PY201_FormItemEnabled();
                    PH_PY201_AddMatrixRow();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY201A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY201B);
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat1.AutoResizeColumns();
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
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    //원본 소스(VB6.0 주석처리되어 있음)
                    //if(pVal.ItemUID == "Code")
                    //{
                    //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY201A", "Code", "", 0, "", "", "");
                    //}
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            int i = 0;

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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY201A", "DocEntry"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY201_FormItemEnabled();
                            PH_PY201_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        //Case "1293":
                        //  Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent);
                        case "1281": //문서찾기
                            PH_PY201_FormItemEnabled();
                            PH_PY201_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY201_FormItemEnabled();
                            PH_PY201_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY201_FormItemEnabled();
                            break;
                        case "1293": // 행삭제
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while (i <= oDS_PH_PY201B.Size - 1)
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY201B.GetValue("U_LineNum", i)))
                                    {
                                        oDS_PH_PY201B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY201B.Size; i++)
                                {
                                    oDS_PH_PY201B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            PH_PY201_AddMatrixRow();
                            break;
                        case "1287": //복제

                            oForm.Freeze(true);
                            oDS_PH_PY201A.SetValue("DocEntry", 0, "");

                            for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
                            {
                                oMat1.FlushToDataSource();
                                oDS_PH_PY201B.SetValue("DocEntry", i, "");
                                oDS_PH_PY201B.SetValue("U_PayYN", i, "N");
                                oMat1.LoadFromDataSource();
                            }

                            oForm.Items.Item("Quarter").Enabled = true;
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
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            //string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                            //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                            //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                            //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
                            //36
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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

                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = pVal.ColUID;
                        oLastColRow01 = pVal.Row;
                    }
                }
                else
                {
                    oLastItemUID01 = pVal.ItemUID;
                    oLastColUID01 = "";
                    oLastColRow01 = 0;
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
    }
}
