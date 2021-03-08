using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 근태시간구분등록
    /// </summary>
    internal class PH_PY002 : PSH_BaseClass
    {
        private string oFormUniqueID01;
        private SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY002A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY002B;
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY002.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY002_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY002");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "Code";

                oForm.Freeze(true);
                PH_PY002_CreateItems();
                PH_PY002_EnableMenus();
                PH_PY002_SetDocument(oFormDocEntry01);
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
                oForm.ActiveItem = "CLTCOD";
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PH_PY002_CreateItems()
        {
            string sQry;
            string CLTCOD;
            int i;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                oDS_PH_PY002A = oForm.DataSources.DBDataSources.Item("@PH_PY002A");
                oDS_PH_PY002B = oForm.DataSources.DBDataSources.Item("@PH_PY002B");

                oMat1 = oForm.Items.Item("Mat1").Specific;
                
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                //사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
                CLTCOD = dataHelpClass.Get_ReData("Branch", "USER_CODE", "OUSR", "'" + PSH_Globals.oCompany.UserName + "'", "");
                oForm.Items.Item("CLTCOD").Specific.Select(CLTCOD, SAPbouiCOM.BoSearchKey.psk_ByValue);
                
                //근무형태
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P154' AND U_UseYN = 'Y' ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("SType").Specific, "");
                oForm.Items.Item("SType").DisplayDesc = true;

                //근무조(=근무구분)
                oForm.Items.Item("Shift").DisplayDesc = true;

                //요일구분
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P202' AND U_UseYN = 'Y'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat1.Columns.Item("DayType").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("DayType").DisplayDesc = true;

                //시간구분
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P157' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                    {
                        oMat1.Columns.Item("TimeType").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                        oRecordSet.MoveNext();
                    }
                }
                oMat1.Columns.Item("TimeType").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY002_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY002_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY002_FormItemEnabled();
                    PH_PY002_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY002_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY002_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메트릭스 Row 추가
        /// </summary>
        private void PH_PY002_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);

                ////[Mat1]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY002B.GetValue("U_FromTime", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY002B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY002B.InsertRecord(oRow);
                        }
                        oDS_PH_PY002B.Offset = oRow;
                        oDS_PH_PY002B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY002B.SetValue("U_DayType", oRow, "");
                        oDS_PH_PY002B.SetValue("U_TimeType", oRow, "");
                        oDS_PH_PY002B.SetValue("U_NextDay", oRow, "N");
                        oDS_PH_PY002B.SetValue("U_FromTime", oRow, "");
                        oDS_PH_PY002B.SetValue("U_ToTime", oRow, "");
                        oDS_PH_PY002B.SetValue("U_Minus", oRow, "N");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY002B.Offset = oRow - 1;
                        oDS_PH_PY002B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY002B.SetValue("U_DayType", oRow - 1, "");
                        oDS_PH_PY002B.SetValue("U_TimeType", oRow - 1, "");
                        oDS_PH_PY002B.SetValue("U_NextDay", oRow - 1, "N");
                        oDS_PH_PY002B.SetValue("U_FromTime", oRow - 1, "");
                        oDS_PH_PY002B.SetValue("U_ToTime", oRow - 1, "");
                        oDS_PH_PY002B.SetValue("U_Minus", oRow - 1, "N");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY002B.Offset = oRow;
                    oDS_PH_PY002B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY002B.SetValue("U_DayType", oRow, "");
                    oDS_PH_PY002B.SetValue("U_TimeType", oRow, "");
                    oDS_PH_PY002B.SetValue("U_NextDay", oRow, "N");
                    oDS_PH_PY002B.SetValue("U_FromTime", oRow, "");
                    oDS_PH_PY002B.SetValue("U_ToTime", oRow, "");
                    oDS_PH_PY002B.SetValue("U_Minus", oRow, "N");
                    oMat1.LoadFromDataSource();
                }

            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DataValidCheck : 입력데이터의 Valid Check
        /// </summary>
        /// <returns></returns>
        private bool PH_PY002_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            short errNum = 0;
            string sQry;
            string tCode;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY002A.GetValue("U_CLTCOD", 0).ToString().Trim())) //사업장
                {
                    errNum = 1;
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oDS_PH_PY002A.GetValue("U_SType", 0).ToString().Trim())) //근무형태
                {
                    errNum = 2;
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oDS_PH_PY002A.GetValue("U_Shift", 0).ToString().Trim())) //근무구분
                {
                    errNum = 3;
                    throw new Exception();
                }

                //코드 생성
                tCode = oDS_PH_PY002A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY002A.GetValue("U_SType", 0).ToString().Trim() + oDS_PH_PY002A.GetValue("U_Shift", 0).ToString().Trim();

                //코드 중복 체크
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    sQry = "SELECT CODE FROM [@PH_PY002A] WHERE CODE = '" + tCode + "'";
                    oRecordSet.DoQuery(sQry);
                    if (oRecordSet.RecordCount > 0)
                    {
                        errNum = 4;
                        throw new Exception();
                    }
                    else
                    {
                        oDS_PH_PY002A.SetValue("Code", 0, tCode);
                        oDS_PH_PY002A.SetValue("Name", 0, tCode);
                    }
                }

                //매트릭스 체크
                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("DayType").Cells.Item(i).Specific.Value)) //일구분
                        {
                            errNum = 5;
                            throw new Exception();
                        }

                        if (string.IsNullOrEmpty(oMat1.Columns.Item("TimeType").Cells.Item(i).Specific.Value)) //시간구분
                        {
                            errNum = 6;
                            throw new Exception();
                        }

                        if (string.IsNullOrEmpty(oMat1.Columns.Item("FromTime").Cells.Item(i).Specific.Value)) //시작시간
                        {
                            errNum = 7;
                            throw new Exception();
                        }

                        if (string.IsNullOrEmpty(oMat1.Columns.Item("ToTime").Cells.Item(i).Specific.Value)) //종료시간
                        {
                            errNum = 8;
                            throw new Exception();
                        }
                    }
                }
                else
                {
                    errNum = 9;
                    throw new Exception();
                }

                oMat1.FlushToDataSource();
                if (oDS_PH_PY002B.Size > 1)
                {
                    oDS_PH_PY002B.RemoveRecord(oDS_PH_PY002B.Size - 1); //Matrix 마지막 행 삭제(DB 저장시)
                }
                oMat1.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch(Exception ex)
            {
                if (errNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat1.Columns.Item("CLTCOD").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("근무형태는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat1.Columns.Item("SType").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("근무구분은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat1.Columns.Item("Shift").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 4)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("코드가 존재합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errNum == 5)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("일구분은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat1.Columns.Item("DayType").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 6)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("시간구분은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat1.Columns.Item("TimeType").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 7)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("시작시간은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat1.Columns.Item("FromTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 8)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("종료시간은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat1.Columns.Item("ToTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errNum == 9)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("라인 데이터가 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY002_DataValidCheck_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                ////    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                ////    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                ////    break;

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
                            if (PH_PY002_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
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
                        if (pVal.ActionSuccess == true)
                        {
                            PH_PY002_FormItemEnabled();
                            PH_PY002_AddMatrixRow();
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
                        case "Mat1":
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
            string sQry;
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
                        if (pVal.ItemUID == "SType")
                        {
                            if (oForm.Items.Item("Shift").Specific.ValidValues.Count > 0)
                            {
                                for (int i = oForm.Items.Item("Shift").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("Shift").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = 'P155' AND U_Char1 = '" + oForm.Items.Item("SType").Specific.Value + "'";
                            sQry += " ORDER BY U_Code";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Shift").Specific, "");

                            oForm.Items.Item("Shift").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            oForm.Items.Item("Shift").DisplayDesc = true;
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ColUID == "FromTime")
                        {
                            PH_PY002_AddMatrixRow();
                            oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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

                    PH_PY002_FormItemEnabled();
                    PH_PY002_AddMatrixRow();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY002A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY002B);
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY002A", "Code"); //접속자 권한에 따른 사업장 보기
                            break;
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY002_FormItemEnabled();
                            PH_PY002_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY002_FormItemEnabled();
                            PH_PY002_AddMatrixRow();
                            break;
                        case "1282": //문서추가
                            PH_PY002_FormItemEnabled();
                            PH_PY002_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY002_FormItemEnabled();
                            break;
                        case "1293": //행삭제
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while (i <= oDS_PH_PY002B.Size - 1)
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY002B.GetValue("U_DayType", i)))
                                    {
                                        oDS_PH_PY002B.RemoveRecord(i);
                                        i = 0;
                                    }
                                    else
                                    {
                                        i += 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY002B.Size; i++)
                                {
                                    oDS_PH_PY002B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            PH_PY002_AddMatrixRow();
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
            int i;
            string sQry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            //부서
                            if (oForm.Items.Item("Shift").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("Shift").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("Shift").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "  SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry += " WHERE Code = 'P155' AND U_Char1 = '" + oDS_PH_PY002A.GetValue("U_SType", 0).ToString().Trim() + "'";
                            sQry += " ORDER BY U_Code";

                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("Shift").Specific, "");
                            oForm.Items.Item("Shift").DisplayDesc = true;
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
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
