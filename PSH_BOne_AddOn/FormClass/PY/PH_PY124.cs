using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 베네피아 금액등록
    /// </summary>
    internal class PH_PY124 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY124A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY124B;

        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        ////적용버턴 실행여부
        private bool CheckDataApply;
        ////사업장
        private string CLTCOD;
        ////적용연월
        private string YM;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY124.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY124_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY124");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                //***************************************************************
                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                oForm.DataBrowser.BrowseBy = "Code";
                //***************************************************************

                oForm.Freeze(true);
                PH_PY124_CreateItems();
                PH_PY124_EnableMenus();
                PH_PY124_SetDocument(oFromDocEntry01);
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
        private void PH_PY124_CreateItems()
        {
            string sQry = string.Empty;
            int i = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                oDS_PH_PY124A = oForm.DataSources.DBDataSources.Item("@PH_PY124A");
                oDS_PH_PY124B = oForm.DataSources.DBDataSources.Item("@PH_PY124B");

                oMat1 = oForm.Items.Item("Mat1").Specific;
                ////@PH_PY124B

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                CheckDataApply = false;

                ////// 사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY124_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true);                ////제거
                oForm.EnableMenu("1284", false);                ////취소
                oForm.EnableMenu("1293", true);                ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY124_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY124_FormItemEnabled();
                    PH_PY124_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY124_FormItemEnabled();
                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY124_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY124_FormItemEnabled()
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
                    oForm.Items.Item("FieldCo").Enabled = true;
                    oForm.Items.Item("Mat1").Enabled = true;
                    oForm.Items.Item("Btn_Apply").Enabled = true;
                    oForm.Items.Item("Btn_Cancel").Enabled = false;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", true);                      ////문서찾기
                    oForm.EnableMenu("1282", false);                     ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = true;
                    oForm.Items.Item("YM").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = false;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.EnableMenu("1281", false);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                     ////문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    oForm.Items.Item("CLTCOD").Enabled = false;
                    oForm.Items.Item("YM").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = false;

                    if (oForm.Items.Item("StatYN").Specific.VALUE == "Y")
                    {
                        oForm.Items.Item("FieldCo").Enabled = false;
                        oForm.Items.Item("Mat1").Enabled = false;
                        oForm.Items.Item("Btn_Apply").Enabled = false;
                        oForm.Items.Item("Btn_Cancel").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("FieldCo").Enabled = true;
                        oForm.Items.Item("Mat1").Enabled = true;
                        oForm.Items.Item("Btn_Apply").Enabled = true;
                        oForm.Items.Item("Btn_Cancel").Enabled = false;
                    }

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", false);

                    oForm.EnableMenu("1281", true);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                    ////문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
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

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PH_PY124_DataValidCheck() == false)
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
                                PH_PY124_DataApply(CLTCOD, YM);
                                CheckDataApply = false;
                            }
                            PH_PY124_FormItemEnabled();
                            PH_PY124_AddMatrixRow();
                        }
                    }
                    if (pVal.ItemUID == "Btn_UPLOAD")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY124_Excel_Upload);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start(); 
                        PH_PY124_AddMatrixRow();
                    }
                    if (pVal.ItemUID == "Btn_Cancel")
                    {
                        PH_PY124_DataCancel(CLTCOD, YM);
                    }
                    if (pVal.ItemUID == "Btn_Apply")
                    {
                        CLTCOD = oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim();
                        YM = oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim();
                        if (oMat1.RowCount > 1)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                CheckDataApply = true;
                                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                PH_PY124_DataApply(CLTCOD, YM);
                            }
                        }
                        else
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("베네피아 자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            string FieldCo = string.Empty;
            int ErrCode = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
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
                        //if (pVal.ItemUID == "Mat1" & pVal.ColUID == "HOBCOD")
                        //{
                        //    PH_PY124_AddMatrixRow();
                        //    oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        //}
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ItemUID == "YM")
                            {

                                CLTCOD = oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim();
                                YM = codeHelpClass.Right(oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim(), 4);

                                if (!string.IsNullOrEmpty(oDS_PH_PY124A.GetValue("U_FieldCo", 0).ToString().Trim()))
                                {
                                    FieldCo = " = '" + oDS_PH_PY124A.GetValue("U_FieldCo", 0).ToString().Trim();
                                }
                                else
                                {
                                    FieldCo = " like '%";
                                }
                                sQry = "select U_Sequence from [@PH_PY109Z] where code ='" + CLTCOD + YM + "111'";
                                oRecordSet.DoQuery(sQry);
                                if ((oRecordSet.RecordCount == 0))
                                {
                                    ErrCode = 1;
                                    throw new Exception();
                                   
                                }
                                else
                                {
                                    sQry = "select distinct U_Sequence,U_PDName from [@PH_PY109Z] where code ='" + CLTCOD + YM + "111' and u_sequence" + FieldCo + "' order by 1";
                                    dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("FieldCo").Specific, "");
                                    oForm.Items.Item("FieldCo").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                    oForm.Items.Item("FieldCo").DisplayDesc = true;
                                }
                            }
                        }
                       
                    }
                }
            }
            catch (Exception ex)
            {
                if(ErrCode == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("급상여변동자료 입력은 필수입니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
            }
        }

        /// <summary>
        /// Raise_EVENT_MATRIX_LOAD
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
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
                    PH_PY124_FormItemEnabled();
                    PH_PY124_AddMatrixRow();
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

                    SubMain.Remove_Forms(oFormUniqueID);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY124A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY124B);
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
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

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
                            CLTCOD = oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim();
                            YM = oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim();
                            PH_PY124_DataCancel(CLTCOD, YM);
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
                            dataHelpClass.AuthorityCheck(oForm, "CLTCOD", "@PH_PY124A", "Code");                            ////접속자 권한에 따른 사업장 보기
                            PH_PY124_FormItemEnabled();
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY124_FormItemEnabled();
                            PH_PY124_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281":
                            ////문서찾기
                            PH_PY124_FormItemEnabled();
                            PH_PY124_AddMatrixRow();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY124_FormItemEnabled();
                            PH_PY124_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY124_FormItemEnabled();

                            CLTCOD = oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim();
                            YM = codeHelpClass.Right(oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim(), 4);
                            break;

                        case "1293":
                            //// 행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent, oMat1, oDS_PH_PY124B, "U_JIGCOD");
                            PH_PY124_AddMatrixRow();
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                               ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                                ////34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                            ////35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                            ////36
                            break;
                    }
                }
                else if ((BusinessObjectInfo.BeforeAction == false))
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                              ////33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                               ////34
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
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        public void PH_PY124_AddMatrixRow()
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
                    if (!string.IsNullOrEmpty(oDS_PH_PY124B.GetValue("U_Seq", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY124B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY124B.InsertRecord((oRow));
                        }
                        oDS_PH_PY124B.Offset = oRow;
                        oDS_PH_PY124B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY124B.SetValue("U_Seq", oRow, "");
                        oDS_PH_PY124B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY124B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY124B.SetValue("U_BeneAmt", oRow, Convert.ToString(0));
                        oDS_PH_PY124B.SetValue("U_BillAmt", oRow, Convert.ToString(0));
                        oDS_PH_PY124B.SetValue("U_CardAmt", oRow, Convert.ToString(0));
                        oDS_PH_PY124B.SetValue("U_TotAmt", oRow, Convert.ToString(0));
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY124B.Offset = oRow - 1;
                        oDS_PH_PY124B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY124B.SetValue("U_Seq", oRow, "");
                        oDS_PH_PY124B.SetValue("U_MSTCOD", oRow, "");
                        oDS_PH_PY124B.SetValue("U_MSTNAM", oRow, "");
                        oDS_PH_PY124B.SetValue("U_BeneAmt", oRow, Convert.ToString(0));
                        oDS_PH_PY124B.SetValue("U_BillAmt", oRow, Convert.ToString(0));
                        oDS_PH_PY124B.SetValue("U_CardAmt", oRow, Convert.ToString(0));
                        oDS_PH_PY124B.SetValue("U_TotAmt", oRow, Convert.ToString(0));
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY124B.Offset = oRow;
                    oDS_PH_PY124B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY124B.SetValue("U_Seq", oRow, "");
                    oDS_PH_PY124B.SetValue("U_MSTCOD", oRow, "");
                    oDS_PH_PY124B.SetValue("U_MSTNAM", oRow, "");
                    oDS_PH_PY124B.SetValue("U_BeneAmt", oRow, Convert.ToString(0));
                    oDS_PH_PY124B.SetValue("U_BillAmt", oRow, Convert.ToString(0));
                    oDS_PH_PY124B.SetValue("U_CardAmt", oRow, Convert.ToString(0));
                    oDS_PH_PY124B.SetValue("U_TotAmt", oRow, Convert.ToString(0));
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        public bool PH_PY124_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (string.IsNullOrEmpty(oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                ////적용시작월
                if (string.IsNullOrEmpty(oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("적용시작월은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("YM").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                //// Code & Name 생성
                oDS_PH_PY124A.SetValue("Code", 0, oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim());
                oDS_PH_PY124A.SetValue("NAME", 0, oDS_PH_PY124A.GetValue("U_CLTCOD", 0).ToString().Trim() + oDS_PH_PY124A.GetValue("U_YM", 0).ToString().Trim());

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (!string.IsNullOrEmpty(dataHelpClass.Get_ReData("Code", "Code", "[@PH_PY124A]", "'" + oDS_PH_PY124A.GetValue("Code", 0).ToString().Trim() + "'", "")))
                    {
                        PSH_Globals.SBO_Application.SetStatusBarMessage("이미 존재하는 코드입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        return functionReturnValue;
                    }
                }

                //// 라인 ---------------------------
                if (oMat1.VisualRowCount >= 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {

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
                if (oDS_PH_PY124B.Size > 1)
                    oDS_PH_PY124B.RemoveRecord((oDS_PH_PY124B.Size - 1));

                oMat1.LoadFromDataSource();
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                functionReturnValue = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// PH_PY124_MTX01
        /// </summary>
        private void PH_PY124_MTX01()
        {
            int i = 0;
            string sQry = string.Empty;
            int ErrNum = 0;

            string Param01 = string.Empty;
            string Param02 = string.Empty;
            string Param03 = string.Empty;
            string Param04 = string.Empty;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

            try
            {
                oForm.Freeze(true);
                Param01 = oForm.Items.Item("Param01").Specific.VALUE;
                Param02 = oForm.Items.Item("Param01").Specific.VALUE;
                Param03 = oForm.Items.Item("Param01").Specific.VALUE;
                Param04 = oForm.Items.Item("Param01").Specific.VALUE;

                sQry = "SELECT 10";
                oRecordSet.DoQuery(sQry);

                oMat1.Clear();
                oMat1.FlushToDataSource();
                oMat1.LoadFromDataSource();

                if ((oRecordSet.RecordCount == 0))
                {
                    throw new Exception();
                }
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PH_PY124B.InsertRecord((i));
                    }
                    oDS_PH_PY124B.Offset = i;
                    oDS_PH_PY124B.SetValue("U_COL01", i, oRecordSet.Fields.Item(0).Value);
                    oDS_PH_PY124B.SetValue("U_COL02", i, oRecordSet.Fields.Item(1).Value);
                    oRecordSet.MoveNext();
                    ProgressBar01.Value = ProgressBar01.Value + 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
                oForm.Update();

                ProgressBar01.Stop();
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY001_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                ProgressBar01.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PH_PY124_FormClear
        /// </summary>
        public void PH_PY124_FormClear()
        {
            string DocEntry = string.Empty;

            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY124'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY124_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY124_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            functionReturnValue = true;

            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY124A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
                {
                    functionReturnValue = false;
                    ErrNumm = 1;
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
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }
            return functionReturnValue;
        }

        /// <summary>
        /// ROW_DELETE(Raise_FormMenuEvent에서 호출)
        /// 해당 클래스에서는 사용되지 않음
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pval, bool BubbleEvent, SAPbouiCOM.Matrix oMat, SAPbouiCOM.DBDataSource DBData, string CheckField)
        {
            int i = 0;

            try
            {
                if (oLastColRow > 0)
                {
                    if (pval.BeforeAction == true)
                    {

                    }
                    else if (pval.BeforeAction == false)
                    {
                        if (oMat.RowCount != oMat.VisualRowCount)
                        {
                            oMat.FlushToDataSource();

                            while (i <= DBData.Size - 1)
                            {
                                if (string.IsNullOrEmpty(DBData.GetValue(CheckField, i)))
                                {
                                    DBData.RemoveRecord(i);
                                    i = 0;
                                }
                                else
                                {
                                    i = i + 1;
                                }
                            }

                            for (i = 0; i <= DBData.Size; i++)
                            {
                                DBData.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                            }

                            oMat.LoadFromDataSource();
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
        /// 매트릭스 행 추가
        /// </summary>
        [STAThread]
        public void PH_PY124_Excel_Upload()
        {
            int loopCount = 0;
            int j = 0;
            int CheckLine = 0;
            int i = 0;
            bool sucessFlag = false;
            short columnCount = 7; //엑셀 컬럼수
            short columnCount2 = 7; //엑셀 컬럼수
            string sFile = null;
            string sQry = null;
            double TOTCNT = 0;
            int V_StatusCnt = 0;
            int oProValue = 0;
            int tRow = 0;
            bool CheckYN = true;
            string temp1 = string.Empty;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            CommonOpenFileDialog commonOpenFileDialog = new CommonOpenFileDialog();

            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xls;*.xlsx"));
            commonOpenFileDialog.Filters.Add(new CommonFileDialogFilter("모든 파일", "*.*"));
            commonOpenFileDialog.IsFolderPicker = false;

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
                // PSH_Globals.SBO_Application.StatusBar.SetText("파일을 선택해 주세요.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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

            SAPbouiCOM.ProgressBar ProgressBar01 = null;
            oForm.Freeze(true);

            oMat1.Clear();
            oMat1.FlushToDataSource();
            oMat1.LoadFromDataSource();
            try
            {
                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("시작!", xlRow.Count, false);
                Microsoft.Office.Interop.Excel.Range[] t = new Microsoft.Office.Interop.Excel.Range[columnCount2 + 1];
                for (loopCount = 1; loopCount <= columnCount2; loopCount++)
                {
                    t[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[1, loopCount];
                }

                // 첫 타이틀 비교
                if (Convert.ToString(t[1].Value) != "일련번호")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("A열 첫번째 행 타이틀은 일련번호", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[2].Value) != "사번")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("B열 두번째 행 타이틀은 사번", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[3].Value) != "이름")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("C열 세번째 행 타이틀은 이름", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[4].Value) != "베네피아")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("D열 세번째 행 타이틀은 베네피아", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[5].Value) != "영수증")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("E열 세번째 행 타이틀은 영수증", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[6].Value) != "복지카드(국내)")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("F열 세번째 행 타이틀은 복지카드(국내)", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }
                if (Convert.ToString(t[7].Value) != "총합계(원)")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("G열 세번째 행 타이틀은 총합계(원)", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception();
                }

                //프로그레스 바    ///////////////////////////////////////
                ProgressBar01.Text = "데이터 읽는중...!";

                //최대값 구하기 ///////////////////////////////////////
                TOTCNT = xlsh.UsedRange.Rows.Count;

                V_StatusCnt = Convert.ToInt32(Math.Round(TOTCNT / 50, 0));
                oProValue = 1;
                tRow = 1;
                /////////////////////////////////////////////////////

                for (i = 2; i <= xlsh.UsedRange.Rows.Count; i++)
                {
                    Microsoft.Office.Interop.Excel.Range[] r = new Microsoft.Office.Interop.Excel.Range[columnCount + 1];

                    for (loopCount = 1; loopCount <= columnCount; loopCount++)
                    {
                        r[loopCount] = (Microsoft.Office.Interop.Excel.Range)xlCell[i, loopCount];
                    }
                    CheckYN = false;
                    for (j = 0; j <= oDS_PH_PY124B.Size - 1; j++)
                    {

                        if (Convert.ToString(r[1].Value) == oDS_PH_PY124B.GetValue("U_MSTCOD", j).ToString().Trim())
                        {
                            CheckYN = true;
                            CheckLine = j;
                        }
                    }

                    ////마지막행 제거
                    if (string.IsNullOrEmpty(oDS_PH_PY124B.GetValue("U_MSTCOD", oDS_PH_PY124B.Size - 1).ToString().Trim()))
                    {
                        oDS_PH_PY124B.RemoveRecord(oDS_PH_PY124B.Size - 1);
                    }

                    oDS_PH_PY124B.InsertRecord(oDS_PH_PY124B.Size);
                    oDS_PH_PY124B.Offset = oDS_PH_PY124B.Size - 1;
                    oDS_PH_PY124B.SetValue("U_LineNum", oDS_PH_PY124B.Size - 1, Convert.ToString(r[1].Value));
                    oDS_PH_PY124B.SetValue("U_Seq", oDS_PH_PY124B.Size - 1, Convert.ToString(r[1].Value));
                    oDS_PH_PY124B.SetValue("U_MSTCOD", oDS_PH_PY124B.Size - 1, Convert.ToString(r[2].Value));
                    oDS_PH_PY124B.SetValue("U_MSTNAM", oDS_PH_PY124B.Size - 1, Convert.ToString(r[3].Value));
                    oDS_PH_PY124B.SetValue("U_BeneAmt", oDS_PH_PY124B.Size - 1, Convert.ToString(r[4].Value));
                    oDS_PH_PY124B.SetValue("U_BillAmt", oDS_PH_PY124B.Size - 1, Convert.ToString(r[5].Value));
                    oDS_PH_PY124B.SetValue("U_CardAmt", oDS_PH_PY124B.Size - 1, Convert.ToString(r[6].Value));
                    oDS_PH_PY124B.SetValue("U_TotAmt", oDS_PH_PY124B.Size - 1, Convert.ToString(r[7].Value));


                    //oDS_PH_PY124B.InsertRecord(oDS_PH_PY124B.Size - 1);
                    //oDS_PH_PY124B.Offset = oDS_PH_PY124B.Size - 1;
                    //oDS_PH_PY124B.SetValue("U_LineNum"  , oDS_PH_PY124B.Size - 1 , Convert.ToString(r[1].Value));
                    //oDS_PH_PY124B.SetValue("U_Seq"      , oDS_PH_PY124B.Size - 1 , Convert.ToString(r[1].Value));
                    //oDS_PH_PY124B.SetValue("U_MSTCOD"   , oDS_PH_PY124B.Size - 1, Convert.ToString(r[2].Value));
                    //oDS_PH_PY124B.SetValue("U_MSTNAM"   , oDS_PH_PY124B.Size - 1, Convert.ToString(r[3].Value));
                    //oDS_PH_PY124B.SetValue("U_BeneAmt"  , oDS_PH_PY124B.Size - 1, Convert.ToString(r[4].Value));
                    //oDS_PH_PY124B.SetValue("U_BillAmt"  , oDS_PH_PY124B.Size - 1, Convert.ToString(r[5].Value));
                    //oDS_PH_PY124B.SetValue("U_CardAmt"  , oDS_PH_PY124B.Size - 1, Convert.ToString(r[6].Value));
                    //oDS_PH_PY124B.SetValue("U_TotAmt"   , oDS_PH_PY124B.Size - 1, Convert.ToString(r[7].Value));

                    if ((TOTCNT > 50 & tRow == oProValue * V_StatusCnt) | TOTCNT <= 50)
                    {
                        ProgressBar01.Text = tRow + "/ " + TOTCNT + " 건 처리중...!";
                        ProgressBar01.Value = ProgressBar01.Value + 1;
                    }
                    tRow = tRow + 1;
                }
                ProgressBar01.Value = ProgressBar01.Value + 1;
                ProgressBar01.Text = ProgressBar01.Value + "/" + (xlRow.Count - 1) + "건 Loding...!";

                ////라인번호 재정의
                for (i = 0; i <= oDS_PH_PY124B.Size - 1; i++)
                {
                    oDS_PH_PY124B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                }
                oMat1.LoadFromDataSource();
                oMat1.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY124_Excel_Upload:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
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
        /// PH_PY124_DataApply
        /// </summary>
        /// <param name="CLTCOD"></param>
        /// <param name="YM"></param>
        /// <returns></returns>
        private bool PH_PY124_DataApply(string CLTCOD, string YM)
        {
            bool functionReturnValue = false;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string AMTLen = string.Empty;

            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            try
            {
                oMat1.FlushToDataSource();

                if (oForm.Items.Item("FieldCo").Specific.VALUE.ToString().Trim().Length == 1)
                {
                    AMTLen = Convert.ToString(Convert.ToDouble("0") + oForm.Items.Item("FieldCo").Specific.VALUE).ToString().Trim();
                }
                else
                {
                    AMTLen = oForm.Items.Item("FieldCo").Specific.VALUE.ToString().Trim();
                }

                sQry = "";
                sQry = sQry + " update [@PH_PY109B]";
                sQry = sQry + " set U_AMT" + AMTLen + "=isnull(U_AMT" + AMTLen + ",0)  + isnull(b.U_TotAmt,0)";
                sQry = sQry + " from [@PH_PY109B] a left join [@PH_PY124B] b on left(a.code,1) = left(b.code,1) and SUBSTRING(a.code,2,4) = right(b.code,4) and a.U_MSTCOD  = b.U_MSTCOD";
                sQry = sQry + " where a.code ='" + CLTCOD + codeHelpClass.Right(YM, 4) + "111'";

                oRecordSet.DoQuery(sQry);

                sQry = "";
                sQry = sQry + " update [@PH_PY124A] set U_statYN = 'Y' where U_NaviDoc ='" +oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + oForm.Items.Item("YM").Specific.VALUE.ToString().Trim() + "'";

                oRecordSet.DoQuery(sQry);

                oForm.Items.Item("StatYN").Specific.VALUE = "Y";
                oForm.Items.Item("Test").Click((SAPbouiCOM.BoCellClickType.ct_Regular));

                oForm.Items.Item("FieldCo").Enabled = false;
                oForm.Items.Item("Mat1").Enabled = false;
                oForm.Items.Item("Btn_Apply").Enabled = false;
                oForm.Items.Item("Btn_Cancel").Enabled = true;

                PSH_Globals.SBO_Application.StatusBar.SetText("급상여변동 자료에 금액이 적용 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }
            return functionReturnValue;
        }


        /// <summary>
        /// PH_PY124_DataCancel
        /// </summary>
        /// <param name="CLTCOD"></param>
        /// <param name="YM"></param>
        /// <returns></returns>
        private bool PH_PY124_DataCancel(string CLTCOD, string YM)
        {
            bool functionReturnValue = false;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string AMTLen = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            try
            {
                oMat1.FlushToDataSource();

                if (oForm.Items.Item("FieldCo").Specific.VALUE.ToString().Trim().Length == 1)
                {
                    AMTLen = Convert.ToString(Convert.ToDouble("0") + oForm.Items.Item("FieldCo").Specific.VALUE).ToString().Trim();
                }
                else
                {
                    AMTLen = oForm.Items.Item("FieldCo").Specific.VALUE.ToString().Trim();
                }

                sQry = "";
                sQry = sQry + " update [@PH_PY109B]";
                sQry = sQry + " set U_AMT" + AMTLen + "=isnull(U_AMT" + AMTLen + ",0)  - isnull(b.U_TotAmt,0)";
                sQry = sQry + " from [@PH_PY109B] a left join [@PH_PY124B] b on left(a.code,1) = left(b.code,1) and SUBSTRING(a.code,2,4) = right(b.code,4) and a.U_MSTCOD  = b.U_MSTCOD";
                sQry = sQry + " where a.code ='" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + codeHelpClass.Right(oForm.Items.Item("YM").Specific.VALUE.ToString().Trim(), 4) + "111'";

                oRecordSet.DoQuery(sQry);

                sQry = "";
                sQry = sQry + " update [@PH_PY124A] set U_statYN = 'N' where U_NaviDoc ='" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + oForm.Items.Item("YM").Specific.VALUE.ToString().Trim() + "'";

                oRecordSet.DoQuery(sQry);

                oForm.Items.Item("StatYN").Specific.VALUE = "N";
                oForm.Items.Item("FieldCo").Enabled = true;
                oForm.Items.Item("Mat1").Enabled = true;
                oForm.Items.Item("Btn_Apply").Enabled = true;
                oForm.Items.Item("Btn_Cancel").Enabled = false;

                PSH_Globals.SBO_Application.StatusBar.SetText("급상여변동 자료에 금액이 적용 되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
            return functionReturnValue;
        }

    }
}

