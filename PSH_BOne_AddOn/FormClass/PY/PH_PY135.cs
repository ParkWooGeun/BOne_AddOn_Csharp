using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급상여분개처리
    /// </summary>
    internal class PH_PY135 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY135A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY135B;

        private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLast_Col_UID; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLast_Col_Row; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY135.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY135_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY135");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.DataBrowser.BrowseBy = "DocEntry";
                
                oForm.Freeze(true);
                PH_PY135_CreateItems();
                PH_PY135_ComboBox_Setting();
                PH_PY135_EnableMenus();
                PH_PY135_SetDocument(oFormDocEntry01);
                PH_PY135_FormResize();
                PH_PY135_FormItemEnabled();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY135_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                oDS_PH_PY135A = oForm.DataSources.DBDataSources.Item("@PH_PY135A");
                oDS_PH_PY135B = oForm.DataSources.DBDataSources.Item("@PH_PY135B");

                oMat01 = oForm.Items.Item("Mat01").Specific; //메트릭스 개체 할당
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
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
        /// 콤보박스 Setting
        /// </summary>
        private void PH_PY135_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            string sQry = string.Empty;

            try
            {
                oForm.Freeze(true);

                ////////////Header//////////_S
                //지급종류
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("2", "상여");
                oForm.Items.Item("JOBTYP").DisplayDesc = true;
                //oForm.Items.Item("JOBTYP").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

                //지급구분
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN = 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "");
                oForm.Items.Item("JOBGBN").DisplayDesc = true;
                //oForm.Items.Item("JOBGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                ////////////Header//////////_E

                ////////////매트릭스//////////_S
                //사업장
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
                ////////////매트릭스//////////_E
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
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY135_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); //삭제
                oForm.EnableMenu("1286", false); //닫기(미지원)
                oForm.EnableMenu("1287", false); //복제
                oForm.EnableMenu("1285", false); //복원
                oForm.EnableMenu("1284", true); //취소
                oForm.EnableMenu("1293", false); //행삭제(미지원)
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY135_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY135_FormItemEnabled();
                    PH_PY135_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY135_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
        private void PH_PY135_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PH_PY135A.SetValue("U_YM", 0, DateTime.Now.AddMonths(-1).ToString("yyyyMM")); //급여계산은 항상 익월에 실시하므로, 전월로 기본 세팅
                oForm.Items.Item("JOBTYP").Specific.Select("1", BoSearchKey.psk_ByValue); //지급종류
                oForm.Items.Item("JOBGBN").Specific.Select("1", BoSearchKey.psk_ByValue); //지급구분

                oDS_PH_PY135A.SetValue("U_DebitT", 0, "0");
                oDS_PH_PY135A.SetValue("U_CreditT", 0, "0");

                //oDS_PH_PY135A.SetValue("U_JOBTYP", 0, "급여"); 
                //oDS_PH_PY135A.SetValue("U_JOBGBN", 0, "정기"); 

                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY135'", "");

                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY135_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("JOBTYP").Enabled = true; //지급종류
                    oForm.Items.Item("YM").Enabled = true; //지급년월
                    oForm.Items.Item("JOBGBN").Enabled = true; //지급구분
                    oForm.Items.Item("JdtDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Btn02").Enabled = true;
                    oForm.Items.Item("Btn03").Enabled = true;

                    //폼 세팅
                    PH_PY135_FormClear();

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("JOBTYP").Enabled = true; //지급종류
                    oForm.Items.Item("YM").Enabled = true; //지급년월
                    oForm.Items.Item("JOBGBN").Enabled = true; //지급구분
                    oForm.Items.Item("JdtDate").Enabled = true;

                    oForm.Items.Item("Mat01").Enabled = true;

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    //dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("JOBTYP").Enabled = false; //지급종류
                    oForm.Items.Item("YM").Enabled = false; //지급년월
                    oForm.Items.Item("JOBGBN").Enabled = false; //지급구분
                    oForm.Items.Item("Btn03").Enabled = true;

                    //        If oForm.Items("JdtDate").Specific.Value <> "" Then '분개일자가 등록이 되어 있는 경우
                    oForm.Items.Item("JdtDate").Enabled = true;
                    //        Else
                    //            oForm.Items("JdtDate").Enabled = False
                    //        End If

                    oForm.Items.Item("Mat01").Enabled = false;

                    //접속자에 따른 권한별 사업장 콤보박스세팅
                    //dataHelpClass.CLTCOD_Select(oForm, "CLTCOD",true);

                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int ErrNum = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY135_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY135_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Btn01") //조회버튼
                    {
                        PH_PY135_MTX01();
                    }
                    else if (pVal.ItemUID == "Btn02") //분개확정 버튼
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("JdtDate").Specific.Value))
                            {
                                ErrNum = 1;
                                PH_PY135_Item_Error_Message(ErrNum);
                                BubbleEvent = false;
                                return;
                            }
                            else if (oForm.Items.Item("Status").Specific.Value == "C")
                            {
                                ErrNum = 2;
                                PH_PY135_Item_Error_Message(ErrNum);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (PH_PY135_Create_oJournalEntries(1) == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        else
                        {
                            dataHelpClass.MDC_GF_Message("먼저 갱신한 후 분개 처리 바랍니다.", "W");
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.ItemUID == "Btn03") //분개취소 버튼
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("JdtDate").Specific.Value))
                            {
                                ErrNum = 1;
                                PH_PY135_Item_Error_Message(ErrNum);
                                BubbleEvent = false;
                                return;
                            }
                            else if (oForm.Items.Item("JdtCC").Specific.Value != "Y")
                            {
                                ErrNum = 3;
                                PH_PY135_Item_Error_Message(ErrNum);
                                BubbleEvent = false;
                                return;
                            }
                            else if (oForm.Items.Item("Status").Specific.Value == "C")
                            {
                                ErrNum = 2;
                                PH_PY135_Item_Error_Message(ErrNum);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (PH_PY135_Cancel_oJournalEntries(1) == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        else
                        {
                            dataHelpClass.MDC_GF_Message("먼저 갱신한 후 분개 처리 바랍니다.", "W");
                            BubbleEvent = false;
                            return;
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
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                PSH_Globals.SBO_Application.ActivateMenuItem("1291"); //이동(최종데이타)
                            }
                            else if (pVal.Action_Success == false)
                            {
                                PH_PY135_FormItemEnabled();
                                PH_PY135_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY135_FormItemEnabled();
                                PH_PY135_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY135_FormItemEnabled();
                            }
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
        /// PH_PY135_Item_Error_Message
        /// </summary>
        /// <param name="ErrNum"></param>
        private void PH_PY135_Item_Error_Message(int ErrNum)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                throw new Exception();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    dataHelpClass.MDC_GF_Message("분개처리일을 먼저 입력하세요.", "E");
                }
                else if (ErrNum == 2)
                {
                    dataHelpClass.MDC_GF_Message("문서가 Close 또는 Cancel 되었습니다.", "E");
                }
                else if (ErrNum == 3)
                {
                    dataHelpClass.MDC_GF_Message("분개생성:Y일 때 취소 할 수 있습니다.", "E");
                }
                else if (ErrNum == 4)
                {
                    dataHelpClass.MDC_GF_Message("거래처코드와, 사업장을 먼저 입력하세요.", "E");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ProfCode"); //배부규칙
                }
                else if (pVal.BeforeAction == false)
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
                    if (pVal.ItemChanged == true)
                    {
                        PH_PY135_FlushToItemValue(pVal.ItemUID);
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
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
                    PH_PY135_FormResize();
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
                    PH_PY135_FormItemEnabled();
                    PH_PY135_AddMatrixRow();
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
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLast_Item_UID = pVal.ItemUID;
                        oLast_Col_UID = pVal.ColUID;
                        oLast_Col_Row = pVal.Row;
                    }
                }
                else
                {
                    oLast_Item_UID = pVal.ItemUID;
                    oLast_Col_UID = "";
                    oLast_Col_Row = 0;
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLast_Item_UID = pVal.ItemUID;
                            oLast_Col_UID = pVal.ColUID;
                            oLast_Col_Row = pVal.Row;

                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    else
                    {
                        oLast_Item_UID = pVal.ItemUID;
                        oLast_Col_UID = "";
                        oLast_Col_Row = 0;
                    }
                }
                else if (pVal.BeforeAction == false)
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
        /// 행삭제(사용자 메소드로 구현)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent)
        {
            int i = 0;

            try
            {
                if (oLast_Col_Row > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }

                        oMat01.FlushToDataSource();
                        oDS_PH_PY135B.RemoveRecord(oDS_PH_PY135B.Size - 1);
                        oMat01.LoadFromDataSource();

                        if (oMat01.RowCount == 0)
                        {
                            PH_PY135_AddMatrixRow();
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY135B.GetValue("U_DocNo", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PH_PY135_AddMatrixRow();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY135A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY135B);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
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
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            PH_PY135_FormItemEnabled();
                            PH_PY135_AddMatrixRow();
                            break;
                        case "1282": //추가
                            PH_PY135_FormItemEnabled();
                            PH_PY135_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
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
            string sQry = string.Empty;

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

                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oLast_Item_UID = pVal.ItemUID;
                        oLast_Col_UID = pVal.ColUID;
                        oLast_Col_Row = pVal.Row;
                    }
                }
                else
                {
                    oLast_Item_UID = pVal.ItemUID;
                    oLast_Col_UID = "";
                    oLast_Col_Row = 0;
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
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY135_FlushToItemValue(string oUID, int oRow = 0, string oCol = "")
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {
                    case "Mat01":
                        oMat01.AutoResizeColumns();
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        private bool PH_PY135_DataValidCheck()
        {
            bool functionReturnValue = true;
            int i = 0;
            short errCode = 0;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //사업장
                if (string.IsNullOrEmpty(oDS_PH_PY135A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    errCode = 1;
                    throw new Exception();
                }

                //지급년월
                if (string.IsNullOrEmpty(oDS_PH_PY135A.GetValue("U_YM", 0).ToString().Trim()))
                {
                    errCode = 2;
                    throw new Exception();
                }

                //라인
                if (oMat01.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        ////배부규칙
                        //if (string.IsNullOrEmpty(oMat01.Columns.Item("ProfCode").Cells.Item(i).Specific.Value) & Convert.ToInt32(oMat01.Columns.Item("DocNo").Cells.Item(i).Specific.Value) != 0)
                        //{
                        //    PSH_Globals.SBO_Application.SetStatusBarMessage("배부규칙은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        //    oMat01.Columns.Item("ProfCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //    functionReturnValue = false;
                        //    return functionReturnValue;
                        //}
                    }
                }
                else
                {
                    errCode = 10;
                    throw new Exception();
                }

                oMat01.FlushToDataSource();
                //Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY135B.Size > 1)
                {
                    oDS_PH_PY135B.RemoveRecord(oDS_PH_PY135B.Size - 1);
                }
                oMat01.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                functionReturnValue = false;

                if (errCode == 1)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == 2)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("지급년월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == 10)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY001_MTX01, 메트릭스에 데이터 로드
        /// </summary>
        private void PH_PY135_MTX01()
        {
            short i = 0;
            string sQry = string.Empty;
            short ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            string CLTCOD = string.Empty; //사업장
            string jobType = string.Empty; //지급종류
            string stdYM = string.Empty; //지급년월
            string jobGBN = string.Empty; //지급구분

            double totalDebit = 0; //차변 합계
            double totalCredit = 0; //대변 합계

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim().ToString(); //사업장
            jobType = oForm.Items.Item("JOBTYP").Specific.Value.Trim().ToString(); //지급종류
            stdYM = oForm.Items.Item("YM").Specific.Value.Trim().ToString(); //지급년월
            jobGBN = oForm.Items.Item("JOBGBN").Specific.Value.Trim().ToString(); //지급구분

            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet01.RecordCount, false);

            try
            {
                oForm.Freeze(true);

                sQry = "            EXEC [PH_PY135_01] ";
                sQry = sQry + "'" + CLTCOD + "',"; //사업장
                sQry = sQry + "'" + jobType + "',"; //지급종류
                sQry = sQry + "'" + stdYM + "',"; //지급년월
                sQry = sQry + "'" + jobGBN + "'"; //지급구분

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY135B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    ErrNum = 1;
                    oDS_PH_PY135A.SetValue("U_DebitT", 0, "0");
                    oDS_PH_PY135A.SetValue("U_CreditT", 0, "0");
                    PH_PY135_AddMatrixRow();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY135B.Size)
                    {
                        oDS_PH_PY135B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY135B.Offset = i;

                    oDS_PH_PY135B.SetValue("U_LineNum", i, (i + 1).ToString());
                    oDS_PH_PY135B.SetValue("U_CLTCOD", i, oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim()); //사업장
                    oDS_PH_PY135B.SetValue("U_ShortCD", i, oRecordSet01.Fields.Item("ShortCD").Value.ToString().Trim()); //GL계정
                    oDS_PH_PY135B.SetValue("U_ShortNM", i, oRecordSet01.Fields.Item("ShortNM").Value.ToString().Trim()); //GL계정명
                    oDS_PH_PY135B.SetValue("U_AcctCode", i, oRecordSet01.Fields.Item("AcctCode").Value.ToString().Trim()); //관리계정
                    oDS_PH_PY135B.SetValue("U_AcctName", i, oRecordSet01.Fields.Item("AcctName").Value.ToString().Trim()); //관리계정명
                    oDS_PH_PY135B.SetValue("U_Debit", i, oRecordSet01.Fields.Item("Debit").Value.ToString().Trim()); //차변
                    oDS_PH_PY135B.SetValue("U_Credit", i, oRecordSet01.Fields.Item("Credit").Value.ToString().Trim()); //대변
                    oDS_PH_PY135B.SetValue("U_ProfCode", i, oRecordSet01.Fields.Item("ProfCode").Value.ToString().Trim()); //배부규칙
                    oDS_PH_PY135B.SetValue("U_ProfName", i, oRecordSet01.Fields.Item("ProfName").Value.ToString().Trim()); //배부규칙명
                    oDS_PH_PY135B.SetValue("U_LineMemo", i, oRecordSet01.Fields.Item("LineMemo").Value.ToString().Trim()); //적요

                    totalDebit = totalDebit + Convert.ToDouble(oRecordSet01.Fields.Item("Debit").Value);
                    totalCredit = totalCredit + Convert.ToDouble(oRecordSet01.Fields.Item("Credit").Value);

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oDS_PH_PY135A.SetValue("U_DebitT", 0, totalDebit.ToString());
                oDS_PH_PY135A.SetValue("U_CreditT", 0, totalCredit.ToString());

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                PH_PY135_AddMatrixRow();
                ProgressBar01.Stop();
            }
            catch (Exception ex)
            {
                ProgressBar01.Stop();
                if (ErrNum == 1)
                {
                    dataHelpClass.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY135_FormResize
        /// </summary>
        private void PH_PY135_FormResize()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                oMat01.AutoResizeColumns();
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
        /// PH_PY135_Create_oJournalEntries(분개)
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool PH_PY135_Create_oJournalEntries(int ChkType)
        {
            bool functionReturnValue = true;

            int i = 0;
            int ErrCode = 0;

            string ErrMsg = string.Empty;
            string RetVal = string.Empty;
            string sTransId = string.Empty;
            string sShortName = string.Empty; //GL계정
            string sAcctCode = string.Empty; //관리계정
            string sDocDate = string.Empty;
            string sPrcCode = string.Empty;
            string sPrcName = string.Empty;
            string BPLid = string.Empty;
            
            double sDebit = 0;
            double sCredit = 0;

            string sLineMemo = string.Empty;
            string sQry = string.Empty;
            string sCC = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.JournalEntries oJournal = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                //sDocDate =  DateTime.ParseExact(oDS_PH_PY135A.GetValue("U_JdtDate", 0), "yyyy-MM-DD", null).ToString("yyyy-MM-dd");
                sDocDate = oDS_PH_PY135A.GetValue("U_JdtDate", 0);

                // var _with1 = f_oJournalEntries;
                oJournal.ReferenceDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null); //전기일
                oJournal.DueDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null); 
                oJournal.TaxDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    sShortName = oMat01.Columns.Item("ShortCD").Cells.Item(i).Specific.Value; //GL계정
                    sAcctCode = oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.Value; //관리계정
                    sDebit = Convert.ToDouble(oMat01.Columns.Item("Debit").Cells.Item(i).Specific.Value); //차변
                    sCredit = Convert.ToDouble(oMat01.Columns.Item("Credit").Cells.Item(i).Specific.Value); //대변
                    sPrcCode = oMat01.Columns.Item("ProfCode").Cells.Item(i).Specific.Value; //배부규칙
                    sPrcName = oMat01.Columns.Item("ProfName").Cells.Item(i).Specific.Value; //배부규칙명
                    sLineMemo = oMat01.Columns.Item("LineMemo").Cells.Item(i).Specific.Value; //적요
                    BPLid = oForm.Items.Item("CLTCOD").Specific.Value.Trim().ToString();
                    oJournal.Lines.Add();

                    if (!string.IsNullOrEmpty(sAcctCode))
                    {
                        oJournal.Lines.SetCurrentLine(i - 1);
                        oJournal.Lines.ShortName = sShortName; //G/L계정                            
                        oJournal.Lines.AccountCode = sAcctCode; //관리계정
                        oJournal.Lines.Debit = sDebit; //차변
                        oJournal.Lines.Credit = sCredit; //대변
                        oJournal.Lines.CostingCode = sPrcCode; //배부규칙
                        oJournal.Lines.UserFields.Fields.Item("U_OcrName").Value = sPrcName; //배부규칙명
                        oJournal.Lines.LineMemo = sLineMemo; //비고

                        if (sShortName.Substring(1, 1) == "5") //비용계정의 법정지출증빙코드
                        {
                            oJournal.Lines.UserFields.Fields.Item("U_BillCode").Value = "P90015"; //법정지출증빙코드
                            oJournal.Lines.UserFields.Fields.Item("U_BillName").Value = "품의서"; //법정지출증빙명
                        }

                        if (sShortName == "66062") //GL계정이 명성상사(창원사업장 세탁업체)일 경우
                        {
                            oJournal.Lines.UserFields.Fields.Item("U_VatBP").Value = sShortName; //거래처
                            oJournal.Lines.UserFields.Fields.Item("U_VatRegN").Value = "606-35-33636"; //사업자번호
                            oJournal.Lines.UserFields.Fields.Item("U_VatBPName").Value = "명성상사"; //거래처명
                        }

                        oJournal.UserFields.Fields.Item("U_BPLId").Value = BPLid;  //사업장
                    }
                }

                RetVal = oJournal.Add().ToString(); //완료
                if (0 != Convert.ToInt32(RetVal))
                {
                    PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception();
                }

                sCC = "Y";

                if (ChkType == 1)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransId);
                    sQry = "        UPDATE  [@PH_PY135A] ";
                    sQry = sQry + " SET     U_JdtNo = '" + sTransId + "',";
                    sQry = sQry + "         U_JdtDate = '" + sDocDate + "',";
                    sQry = sQry + "         U_JdtCC = '" + sCC + "'";
                    sQry = sQry + " WHERE   DocNum = '" + oDS_PH_PY135A.GetValue("DocNum", 0).ToString().Trim() + "'";

                    oRecordSet01.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PH_PY135A.SetValue("U_JdtNo", 0, sTransId);
                oDS_PH_PY135A.SetValue("U_JdtCC", 0, sCC);
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                functionReturnValue = false;

                if (ErrCode == -5002)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ErrMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournal);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY135_Create_oJournalEntries(분개취소-역분개)
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool PH_PY135_Cancel_oJournalEntries(int ChkType)
        {
            bool functionReturnValue = true;
            int ErrCode = 0;
            int RetVal = 0;
            string ErrMsg = string.Empty;
            string sTransId = string.Empty;
            string sCardCode = string.Empty;
            string sDocDate = string.Empty;
            string sCC = string.Empty;
            string sQry = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.JournalEntries oJournal = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();

                string jdtNo = oDS_PH_PY135A.GetValue("U_JdtNo", 0).Trim();

                if (oJournal.GetByKey(Convert.ToInt32(jdtNo)) == false)
                {
                    PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception();
                }
                
                RetVal = oJournal.Cancel(); //완료
                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
                    throw new Exception();
                }

                sCC = "N";

                if (ChkType == 1)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransId);
                    sQry = "        UPDATE  [@PH_PY135A]";
                    sQry = sQry + " SET     U_JdtCanNo = '" + sTransId + "',";
                    sQry = sQry + "         U_JdtCC = '" + sCC + "'";
                    sQry = sQry + " WHERE   DocNum = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    oRecordSet01.DoQuery(sQry);

                    if ((PSH_Globals.oCompany.InTransaction == true))
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PH_PY135A.SetValue("U_JdtCanNo", 0, sTransId);
                oDS_PH_PY135A.SetValue("U_JdtCC", 0, sCC);

                oForm.Items.Item("Btn02").Enabled = false;
                oForm.Items.Item("Btn03").Enabled = false;

                dataHelpClass.MDC_GF_Message("성공적으로 분개취소되었습니다.", "S");
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                functionReturnValue = false;
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oJournal);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        private void PH_PY135_AddMatrixRow()
        {
            int oRow = 0;
            oForm.Freeze(true);

            try
            {
                //[Mat1]
                oMat01.FlushToDataSource();
                oRow = oMat01.VisualRowCount;

                if (oMat01.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY135B.GetValue("U_CLTCOD", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY135B.Size <= oMat01.VisualRowCount)
                        {
                            oDS_PH_PY135B.InsertRecord(oRow);
                        }
                        oDS_PH_PY135B.Offset = oRow;
                        oDS_PH_PY135B.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                        oDS_PH_PY135B.SetValue("U_CLTCOD", oRow, "");
                        oDS_PH_PY135B.SetValue("U_AcctCode", oRow, "");
                        oDS_PH_PY135B.SetValue("U_AcctName", oRow, "");
                        oDS_PH_PY135B.SetValue("U_Debit", oRow, "0");
                        oDS_PH_PY135B.SetValue("U_Credit", oRow, "0");
                        oDS_PH_PY135B.SetValue("U_ProfCode", oRow, "");
                        oDS_PH_PY135B.SetValue("U_ProfName", oRow, "");
                        oDS_PH_PY135B.SetValue("U_LineMemo", oRow, "");
                        oMat01.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY135B.Offset = oRow - 1;
                        oDS_PH_PY135B.SetValue("U_LineNum", oRow - 1, oRow.ToString());
                        oDS_PH_PY135B.SetValue("U_AcctCode", oRow - 1, "");
                        oDS_PH_PY135B.SetValue("U_AcctName", oRow - 1, "");
                        oDS_PH_PY135B.SetValue("U_Debit", oRow - 1, "0");
                        oDS_PH_PY135B.SetValue("U_Credit", oRow - 1, "0");
                        oDS_PH_PY135B.SetValue("U_ProfCode", oRow - 1, "");
                        oDS_PH_PY135B.SetValue("U_ProfName", oRow - 1, "");
                        oDS_PH_PY135B.SetValue("U_LineMemo", oRow - 1, "");
                        oMat01.LoadFromDataSource();
                    }
                }
                else if (oMat01.VisualRowCount == 0)
                {
                    oDS_PH_PY135B.Offset = oRow;
                    oDS_PH_PY135B.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                    oDS_PH_PY135B.SetValue("U_AcctCode", oRow, "");
                    oDS_PH_PY135B.SetValue("U_AcctName", oRow, "");
                    oDS_PH_PY135B.SetValue("U_Debit", oRow, "0");
                    oDS_PH_PY135B.SetValue("U_Credit", oRow, "0");
                    oDS_PH_PY135B.SetValue("U_ProfCode", oRow, "");
                    oDS_PH_PY135B.SetValue("U_ProfName", oRow, "");
                    oDS_PH_PY135B.SetValue("U_LineMemo", oRow, "");
                    oMat01.LoadFromDataSource();
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
    }
}
