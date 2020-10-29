using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 공용분개처리
    /// </summary>
    internal class PH_PY034 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.DBDataSource oDS_PH_PY034A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY034B;

        private string oLast_Item_UID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLast_Col_UID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLast_Col_Row;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

    /// <summary>
    /// Form 호출
    /// </summary>
    /// <param name="oFormDocEntry01"></param>
    public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY034.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY034_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY034");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                //***************************************************************
                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                oForm.DataBrowser.BrowseBy = "DocEntry";
                //***************************************************************

                oForm.Freeze(true);
                PH_PY034_CreateItems();
                PH_PY034_ComboBox_Setting();
                PH_PY034_EnableMenus();
                PH_PY034_SetDocument(oFormDocEntry01);
                PH_PY034_FormResize();
                PH_PY034_FormItemEnabled();
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
        private void PH_PY034_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                oDS_PH_PY034A = oForm.DataSources.DBDataSources.Item("@PH_PY034A");
                oDS_PH_PY034B = oForm.DataSources.DBDataSources.Item("@PH_PY034B");
                //메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.Items.Item("CLTCOD").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY034_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 콤보박스 Setting
        /// </summary>
        private void PH_PY034_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);

                ////////////매트릭스//////////
                //사업장
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("CLTCOD"), "SELECT BPLId, BPLName FROM OBPL order by BPLId","","");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY034_ComboBox_Setting_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY034_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false);                //삭제
                oForm.EnableMenu("1286", false);                //닫기(미지원)
                oForm.EnableMenu("1287", false);                //복제
                oForm.EnableMenu("1285", false);                //복원
                oForm.EnableMenu("1284", true);                //취소
                oForm.EnableMenu("1293", false);                //행삭제(미지원)
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY034_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY034_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY034_FormItemEnabled();
                    PH_PY034_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY034_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY034_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY034_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY034'", "");

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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY034_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY034_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {

                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("TeamCode").Enabled = true;
                    oForm.Items.Item("FrDt").Enabled = true;
                    oForm.Items.Item("ToDt").Enabled = true;
                    oForm.Items.Item("JdtDate").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Btn02").Enabled = true;
                    oForm.Items.Item("Btn03").Enabled = true;

                    //폼 DocEntry 세팅
                    PH_PY034_FormClear();

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);
                    oForm.EnableMenu("1281", true);                    ////문서찾기
                    oForm.EnableMenu("1282", false);                   ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {

                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("TeamCode").Enabled = true;
                    oForm.Items.Item("FrDt").Enabled = true;
                    oForm.Items.Item("ToDt").Enabled = true;
                    oForm.Items.Item("JdtDate").Enabled = true;

                    oForm.Items.Item("Mat01").Enabled = true;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                   // dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false);
                    ////문서찾기
                    oForm.EnableMenu("1282", true);
                    ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {

                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("TeamCode").Enabled = false;
                    oForm.Items.Item("FrDt").Enabled = false;
                    oForm.Items.Item("ToDt").Enabled = false;
                    oForm.Items.Item("Btn03").Enabled = true;

                    //        If oForm.Items("JdtDate").Specific.Value <> "" Then '분개일자가 등록이 되어 있는 경우
                    oForm.Items.Item("JdtDate").Enabled = true;
                    //        Else
                    //            oForm.Items("JdtDate").Enabled = False
                    //        End If

                    oForm.Items.Item("Mat01").Enabled = false;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    //dataHelpClass.CLTCOD_Select(oForm, "CLTCOD",true);

                    oForm.EnableMenu("1281", true);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                    ////문서추가

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY034_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PH_PY034_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY034_DataValidCheck() == false)
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
                        PH_PY034_MTX01();
                    }
                    else if (pVal.ItemUID == "Btn02") //분개확정 버튼
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("JdtDate").Specific.Value))
                            {
                                PH_PY034_Item_Error_Message(1);
                                BubbleEvent = false;
                                return;
                            }
                            else if (oForm.Items.Item("Status").Specific.Value == "C")
                            {
                                PH_PY034_Item_Error_Message(2);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (PH_PY034_Create_oJournalEntries(1) == false)
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
                                PH_PY034_Item_Error_Message(1);
                                BubbleEvent = false;
                                return;
                            }
                            else if (oForm.Items.Item("JdtCC").Specific.Value != "Y")
                            {
                                PH_PY034_Item_Error_Message(3);
                                BubbleEvent = false;
                                return;
                            }
                            else if (oForm.Items.Item("Status").Specific.Value == "C")
                            {
                                PH_PY034_Item_Error_Message(2);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (PH_PY034_Cancel_oJournalEntries(1) == false)
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
                                PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                            }
                            else if (pVal.Action_Success == false)
                            {
                                PH_PY034_FormItemEnabled();
                                PH_PY034_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY034_FormItemEnabled();
                                PH_PY034_AddMatrixRow();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PH_PY034_FormItemEnabled();
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
        /// PH_PY034_Item_Error_Message
        /// </summary>
        /// <param name="ErrNum"></param>
        private void PH_PY034_Item_Error_Message(int ErrNum)
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
                    dataHelpClass.MDC_GF_Message( "문서가 Close 또는 Cancel 되었습니다.", "E");
                }
                else if (ErrNum == 3)
                {
                    dataHelpClass.MDC_GF_Message( "분개생성:Y일 때 취소 할 수 있습니다.", "E");
                }
                else if (ErrNum == 4)
                {
                    dataHelpClass.MDC_GF_Message( "거래처코드와, 사업장을 먼저 입력하세요.", "E");
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
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "TeamCode", ""); //팀정보
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ProfCode"); //배부규칙
                }
                else if (pVal.BeforeAction == false)
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        PH_PY034_FlushToItemValue(pVal.ItemUID);
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    PH_PY034_FormResize();
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
                    PH_PY034_FormItemEnabled();
                    PH_PY034_AddMatrixRow();
                    ////UDO방식
                    oMat01.AutoResizeColumns();
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            int i;

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
                        oDS_PH_PY034B.RemoveRecord(oDS_PH_PY034B.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PH_PY034_AddMatrixRow();
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PH_PY034B.GetValue("U_DocNo", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PH_PY034_AddMatrixRow();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ROW_DELETE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY034A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY034B);
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
                            PH_PY034_FormItemEnabled();
                            PH_PY034_AddMatrixRow();
                            break;
                        case "1282": //추가
                            PH_PY034_FormItemEnabled();
                            PH_PY034_AddMatrixRow();
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
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                            ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                            ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }


        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        private void PH_PY034_FlushToItemValue(string oUID)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {
                    case "Mat01":
                        oMat01.AutoResizeColumns();
                        break;

                    case "CLTCOD":
                        break;

                    case "TeamCode":
                        oForm.Items.Item("TeamName").Specific.Value = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item("TeamCode").Specific.Value + "'", " AND Code = '1'"); //팀명
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY034_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY034_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY034A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //팀
                if (string.IsNullOrEmpty(oDS_PH_PY034A.GetValue("U_TeamCode", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("팀정보는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("TeamCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //기간(시작)
                if (string.IsNullOrEmpty(oDS_PH_PY034A.GetValue("U_FrDt", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("기간(시작) 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("FrDt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //기간(종료)
                if (string.IsNullOrEmpty(oDS_PH_PY034A.GetValue("U_ToDt", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("기간(종료)는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("ToDt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                //라인
                if (oMat01.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        //배부규칙
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("ProfCode").Cells.Item(i).Specific.Value) && Convert.ToInt32(oMat01.Columns.Item("DocNo").Cells.Item(i).Specific.Value) != 0)
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("배부규칙은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat01.Columns.Item("ProfCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }

                        //적요
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineMemo").Cells.Item(i).Specific.Value) && Convert.ToInt32(oMat01.Columns.Item("DocNo").Cells.Item(i).Specific.Value) != 0)
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("적요는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat01.Columns.Item("LineMemo").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return functionReturnValue;
                }

                oMat01.FlushToDataSource();
                //Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY034B.Size > 1)
                {
                    oDS_PH_PY034B.RemoveRecord(oDS_PH_PY034B.Size - 1);
                } 
                oMat01.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY034_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        public void PH_PY034_MTX01()
        {
            short i;
            string sQry;
            short ErrNum = 0;
            string CLTCOD;          //사업장
            string TeamCode;        //팀
            string FrDt;            //기간(시작)
            string ToDt;            //기간(종료)
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim().ToString();                                //사업장
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.Trim().ToString();                            //부서
                FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim().Replace(".", "");          //시작일자
                ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim().Replace(".", "");          //종료일자

                sQry = "EXEC [PH_PY034_01] '";
                sQry += CLTCOD + "','";                //사업장
                sQry += TeamCode + "','";                //부서
                sQry += FrDt + "','";                //시작일자
                sQry += ToDt + "'";                //종료일자

                oRecordSet01.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY034B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet01.RecordCount == 0)
                {
                    ErrNum = 1;
                    PH_PY034_FormItemEnabled();
                    PH_PY034_AddMatrixRow();
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY034B.Size)
                    {
                        oDS_PH_PY034B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY034B.Offset = i;

                    oDS_PH_PY034B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY034B.SetValue("U_DocNo", i, oRecordSet01.Fields.Item("DocNo").Value.ToString().Trim());                          //관리번호
                    oDS_PH_PY034B.SetValue("U_CLTCOD", i, oRecordSet01.Fields.Item("CLTCOD").Value.ToString().Trim());                        //사업장
                    oDS_PH_PY034B.SetValue("U_DestNo1", i, oRecordSet01.Fields.Item("DestNo1").Value.ToString().Trim());                      //출장번호1
                    oDS_PH_PY034B.SetValue("U_DestNo2", i, oRecordSet01.Fields.Item("DestNo2").Value.ToString().Trim());                      //출장번호2
                    oDS_PH_PY034B.SetValue("U_AcctCode", i, oRecordSet01.Fields.Item("AcctCode").Value.ToString().Trim());                    //G/L계정
                    oDS_PH_PY034B.SetValue("U_AcctName", i, oRecordSet01.Fields.Item("AcctName").Value.ToString().Trim());                    //G/L계정명
                    oDS_PH_PY034B.SetValue("U_Debit", i, oRecordSet01.Fields.Item("Debit").Value.ToString().Trim());                          //차변
                    oDS_PH_PY034B.SetValue("U_Credit", i, oRecordSet01.Fields.Item("Credit").Value.ToString().Trim());                        //대변
                    oDS_PH_PY034B.SetValue("U_ProfCode", i, oRecordSet01.Fields.Item("ProfCode").Value.ToString().Trim());                    //배부규칙
                    oDS_PH_PY034B.SetValue("U_Spender", i, oRecordSet01.Fields.Item("Spender").Value.ToString().Trim());                      //사용인
                    oDS_PH_PY034B.SetValue("U_LineMemo", i, oRecordSet01.Fields.Item("LineMemo").Value.ToString().Trim());                    //적요

                    oRecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                PH_PY034_AddMatrixRow();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    dataHelpClass.MDC_GF_Message("조회 결과가 없습니다. 확인하세요.", "W");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY034_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

            }
            finally
            {
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY034_FormResize
        /// </summary>
        private void PH_PY034_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY034_FormResize_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }


        /// <summary>
        /// PH_PY034_Create_oJournalEntries
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool PH_PY034_Create_oJournalEntries(int ChkType)
        {
            bool functionReturnValue = false;
            
            int i;
            int ErrCode = 0;
            string ErrMsg = string.Empty;
            string RetVal;
            string sTransId = string.Empty;
            string sAcctCode;
            string sDocDate;
            string sPrcCode;
            string sSpender;
            string BPLid;
            double sDebit;
            double sCredit;

            //사용인
            string sLineMemo;
            string sQry;
            string sCC;

            string errCode = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.JournalEntries oJournal = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                oMat01.FlushToDataSource();
                sDocDate = oDS_PH_PY034A.GetValue("U_JdtDate", 0);
                oJournal.ReferenceDate = DateTime.ParseExact(sDocDate,"yyyyMMdd",null);
                oJournal.DueDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);
                oJournal.TaxDate = DateTime.ParseExact(sDocDate, "yyyyMMdd", null);

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    sAcctCode = oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.Value;    //G/L 계정
                    sDebit = Convert.ToDouble(oMat01.Columns.Item("Debit").Cells.Item(i).Specific.Value);          //차변
                    sCredit = Convert.ToDouble(oMat01.Columns.Item("Credit").Cells.Item(i).Specific.Value);        //대변
                    sPrcCode = oMat01.Columns.Item("ProfCode").Cells.Item(i).Specific.Value;     //배부규칙
                    sSpender = oMat01.Columns.Item("Spender").Cells.Item(i).Specific.Value;      //사용인
                    sLineMemo = oMat01.Columns.Item("LineMemo").Cells.Item(i).Specific.Value;    //적요
                    BPLid = oForm.Items.Item("CLTCOD").Specific.Value.Trim().ToString();
                    oJournal.Lines.Add();

                    if (!string.IsNullOrEmpty(sAcctCode))
                    {
                        oJournal.Lines.SetCurrentLine(i - 1);
                        oJournal.Lines.AccountCode = sAcctCode;               //관리계정
                        oJournal.Lines.ShortName = sAcctCode;                 //G/L계정/BP 코드
                        oJournal.Lines.Debit = sDebit;                        //차변
                        oJournal.Lines.Credit = sCredit;                      //대변
                        oJournal.Lines.CostingCode = sPrcCode;                //배부규칙
                        oJournal.Lines.LineMemo = sLineMemo;                  //비고

                        oJournal.Lines.UserFields.Fields.Item("U_spender").Value = sSpender;                                         //사용인
                        oJournal.Lines.UserFields.Fields.Item("U_BillCode").Value = "P90010";                                        //법정지출증빙코드
                        oJournal.Lines.UserFields.Fields.Item("U_BillName").Value = "규정";                                          //법정지출증빙명
                        oJournal.UserFields.Fields.Item("U_BPLId").Value = BPLid;  //사업장
                    }
                }
                //완료
                RetVal = oJournal.Add().ToString();
                if (0 != Convert.ToInt32(RetVal))
                {
                    PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
                    errCode = "1";
                    throw new Exception();
                }

                sCC = "Y";

                if (ChkType == 1)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransId);
                    sQry = "  UPDATE  [@PH_PY034A] ";
                    sQry += " SET     U_JdtNo = '" + sTransId + "',";
                    sQry += "         U_JdtDate = '" + sDocDate + "',";
                    sQry += "         U_JdtCC = '" + sCC + "'";
                    sQry += " WHERE   DocNum = '" + oDS_PH_PY034A.GetValue("DocNum", 0).ToString().Trim() + "'";

                    oRecordSet01.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PH_PY034A.SetValue("U_JdtNo", 0, sTransId);
                oDS_PH_PY034A.SetValue("U_JdtCC", 0, sCC);
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + ErrCode + "]" + ErrMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY034_Create_oJournalEntries_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        /// PH_PY034_Create_oJournalEntries
        /// </summary>
        /// <param name="ChkType"></param>
        /// <returns></returns>
        private bool PH_PY034_Cancel_oJournalEntries(int ChkType)
        {
            bool functionReturnValue = false;
            int ErrCode = 0;
            int RetVal;
            string ErrMsg = string.Empty;
            string sTransId = string.Empty;
            //string sCardCode;
            //string sDocDate;
            string sCC;
            string sQry;
            string errCode = string.Empty;

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

                string jdtNo = oDS_PH_PY034A.GetValue("U_JdtNo", 0).Trim();
                
                if (oJournal.GetByKey(Convert.ToInt32(jdtNo)) == false)
                {
                    PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
                    errCode = "1";
                    throw new Exception();
                }
                //완료
                RetVal = oJournal.Cancel();
                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
                    errCode = "1";
                    throw new Exception();
                }

                sCC = "N";

                if (ChkType == 1)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out sTransId);
                    sQry = "  UPDATE  [@PH_PY034A]";
                    sQry += " SET     U_JdtCanNo = '" + sTransId + "',";
                    sQry += "         U_JdtCC = '" + sCC + "'";
                    sQry += " WHERE   DocNum = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    oRecordSet01.DoQuery(sQry);

                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }

                oDS_PH_PY034A.SetValue("U_JdtCanNo", 0, sTransId);
                oDS_PH_PY034A.SetValue("U_JdtCC", 0, sCC);

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

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + ErrCode + "]" + ErrMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY034_Cancel_oJournalEntries_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
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
        public void PH_PY034_AddMatrixRow()
        {
            int oRow;
            
            try
            {
                oForm.Freeze(true);
                ////[Mat1]
                oMat01.FlushToDataSource();
                oRow = oMat01.VisualRowCount;

                if (oMat01.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY034B.GetValue("U_DocNo", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY034B.Size <= oMat01.VisualRowCount)
                        {
                            oDS_PH_PY034B.InsertRecord(oRow);
                        }
                        oDS_PH_PY034B.Offset = oRow;
                        oDS_PH_PY034B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY034B.SetValue("U_DocNo", oRow, "");
                        oDS_PH_PY034B.SetValue("U_CLTCOD", oRow, "");
                        oDS_PH_PY034B.SetValue("U_DestNo1", oRow, "");
                        oDS_PH_PY034B.SetValue("U_DestNo2", oRow, "");
                        oDS_PH_PY034B.SetValue("U_AcctCode", oRow, "");
                        oDS_PH_PY034B.SetValue("U_AcctName", oRow, "");
                        oDS_PH_PY034B.SetValue("U_Debit", oRow, Convert.ToString(0));
                        oDS_PH_PY034B.SetValue("U_Credit", oRow, Convert.ToString(0));
                        oDS_PH_PY034B.SetValue("U_ProfCode", oRow, "");
                        oDS_PH_PY034B.SetValue("U_Spender", oRow, "");
                        oDS_PH_PY034B.SetValue("U_LineMemo", oRow, "");
                        oMat01.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY034B.Offset = oRow - 1;
                        oDS_PH_PY034B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY034B.SetValue("U_DocNo", oRow - 1, "");
                        oDS_PH_PY034B.SetValue("U_CLTCOD", oRow - 1, "");
                        oDS_PH_PY034B.SetValue("U_DestNo1", oRow - 1, "");
                        oDS_PH_PY034B.SetValue("U_DestNo2", oRow - 1, "");
                        oDS_PH_PY034B.SetValue("U_AcctCode", oRow - 1, "");
                        oDS_PH_PY034B.SetValue("U_AcctName", oRow - 1, "");
                        oDS_PH_PY034B.SetValue("U_Debit", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY034B.SetValue("U_Credit", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY034B.SetValue("U_ProfCode", oRow - 1, "");
                        oDS_PH_PY034B.SetValue("U_Spender", oRow, "");
                        oDS_PH_PY034B.SetValue("U_LineMemo", oRow - 1, "");
                        oMat01.LoadFromDataSource();
                    }
                }
                else if (oMat01.VisualRowCount == 0)
                {
                    oDS_PH_PY034B.Offset = oRow;
                    oDS_PH_PY034B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY034B.SetValue("U_DocNo", oRow, "");
                    oDS_PH_PY034B.SetValue("U_CLTCOD", oRow, "");
                    oDS_PH_PY034B.SetValue("U_DestNo1", oRow, "");
                    oDS_PH_PY034B.SetValue("U_DestNo2", oRow, "");
                    oDS_PH_PY034B.SetValue("U_AcctCode", oRow, "");
                    oDS_PH_PY034B.SetValue("U_AcctName", oRow, "");
                    oDS_PH_PY034B.SetValue("U_Debit", oRow, Convert.ToString(0));
                    oDS_PH_PY034B.SetValue("U_Credit", oRow, Convert.ToString(0));
                    oDS_PH_PY034B.SetValue("U_ProfCode", oRow, "");
                    oDS_PH_PY034B.SetValue("U_Spender", oRow, "");
                    oDS_PH_PY034B.SetValue("U_LineMemo", oRow, "");
                    oMat01.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY034_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
    }
}

