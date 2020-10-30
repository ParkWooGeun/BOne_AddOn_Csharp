using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 퇴직금 분개장 생성
    /// </summary>
    internal class PH_PY116 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY116A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY116B;

        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oSTRDAT;
        private string oENDDAT;
        private string oJOBTYP = string.Empty;
        private string oCLTCOD;
        private string oMSTCOD;
        private double oTOTDEB;
        private double oTOTCRE;

        private string oDocDate;
        private string oREMARK;
        private string oDocNum;
        private string oDIM3 = string.Empty;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY116.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY116_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY116");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                //***************************************************************
                //화면키값(화면에서 유일키값을 담고 있는 아이템의 Uid값)
                oForm.DataBrowser.BrowseBy = "DocEntry";
                //***************************************************************

                oForm.Freeze(true);
                PH_PY116_CreateItems();
                PH_PY116_EnableMenus();
                PH_PY116_SetDocument(oFormDocEntry01);
                oForm.Update();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY116_CreateItems()
        {
            try
            {
                oForm.Freeze(true);
                
                oDS_PH_PY116A = oForm.DataSources.DBDataSources.Item("@PH_PY116A"); //헤더
                oDS_PH_PY116B = oForm.DataSources.DBDataSources.Item("@PH_PY116B"); //라인
                oMat1 = oForm.Items.Item("Mat1").Specific;
                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                //사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //Matrix
                oMat1 = oForm.Items.Item("Mat1").Specific;

                oMat1.Columns.Item("AcctCode").ExtendedObject.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY116_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY116_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false);                ////제거
                oForm.EnableMenu("1284", false);                ////취소
                oForm.EnableMenu("1293", true);                ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY116_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY116_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY116_FormItemEnabled();
                    PH_PY116_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY116_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY116_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY116_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);

                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                {
                    PH_PY116_FormClear();
                    oForm.ActiveItem = "CLTCOD";
                    oForm.Items.Item("DocEntry").Enabled = false;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    //// 지급일자
                    oDS_PH_PY116A.SetValue("U_STRDAT", 0, DateTime.Now.ToString("yyyyMM")+"01");
                    oDS_PH_PY116A.SetValue("U_STRDAT", 0, DateTime.Now.ToString("yyyyMMdd"));

                    //// 전기일자
                    oDS_PH_PY116A.SetValue("U_DOCDATE", 0, DateTime.Now.ToString("yyyyMMdd"));

                    oForm.EnableMenu("1281", true);                     ////문서찾기
                    oForm.EnableMenu("1282", false);                    ////문서추가

                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.Items.Item("DocEntry").Enabled = true;

                    oForm.EnableMenu("1281", false);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                     ////문서추가
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
                {
                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                    oForm.Items.Item("DocEntry").Enabled = false;

                    oForm.EnableMenu("1281", true);                    ////문서찾기
                    oForm.EnableMenu("1282", true);                    ////문서추가

                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY116_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                        if (PH_PY116_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn2")
                    {
                        if (Execution_Process() == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                        else
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.ItemUID == "CBtn2")
                    {
                        if (oForm.Items.Item("MSTCOD").Enabled == true)
                        {
                            oForm.Items.Item("MSTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                            BubbleEvent = false;
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            PH_PY116_FormItemEnabled();
                        }
                    }
                    //분개장 문서생성
                    if (pVal.ItemUID == "Btn1")
                    {
                        if (!string.IsNullOrEmpty(oDS_PH_PY116A.GetValue("DocEntry", 0)))
                        {
                            if (dataHelpClass.Value_ChkYn("[@PH_PY116A]", "DocEntry", "'" + oDS_PH_PY116A.GetValue("DocEntry", 0).ToString().Trim() + "'","") == false)
                            {
                                DI_oJournalEntries();
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("저장된 문서만 분개 생성이 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                if (pVal.BeforeAction == true & pVal.ColUID == "AcctCode" & pVal.CharPressed == 9)
                {
                      if (pVal.BeforeAction == true & pVal.ColUID == "AcctCode" & pVal.CharPressed == 9)
                    {
                        if (dataHelpClass.Value_ChkYn("[OACT]", "FormatCode", "'" + oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String + "'", "") == true | string.IsNullOrEmpty(oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
                        {
                            PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                            BubbleEvent = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(dataHelpClass);
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
                if (pVal.BeforeAction == false && pVal.ItemChanged == true)
                {
                    if (pVal.ItemUID == "JOBTYP")
                    {
                        if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "R01")
                        {
                            oForm.DataSources.UserDataSources.Item("STRDAT").ValueEx = DateTime.Now.ToString("yyyyMM") + "01";
                            oForm.DataSources.UserDataSources.Item("ENDDAT").ValueEx = DateTime.Now.ToString("yyyyMMdd");
                        }
                        else
                        {
                            oForm.DataSources.UserDataSources.Item("STRDAT").ValueEx = DateTime.Now.ToString("yyyyMMdd");
                            oForm.DataSources.UserDataSources.Item("ENDDAT").ValueEx = DateTime.Now.ToString("yyyyMMdd");
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string STRDAT;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.BeforeAction == false && pVal.ItemChanged == true)
                {
                    if (pVal.ItemUID == "MSTCOD")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item("MSTCOD").Specific.String))
                        {
                            oForm.Items.Item("MSTNAM").Specific.Value = "";
                        }
                        else
                        {
                            oForm.Items.Item("MSTNAM").Specific.Value = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.String.ToString().Trim() + "'","");
                        }
                        oForm.Items.Item("MSTCOD").Update();
                        oForm.Items.Item("MSTNAM").Update();
                    }
                    if (pVal.ItemUID == "STRDAT" || pVal.ItemUID == "ENDDAT")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.String))
                        {
                            oDS_PH_PY116A.SetValue("U_STRDAT", 0, DateTime.Now.ToString("yyyyMMdd"));
                            oDS_PH_PY116A.SetValue("U_ENDDAT", 0, DateTime.Now.ToString("yyyyMMdd"));
                        }
                        else if (pVal.ItemUID == "STRDAT")
                        {
                            STRDAT = oForm.Items.Item("STRDAT").Specific.Value;
                            oDS_PH_PY116A.SetValue("U_ENDDAT", 0, STRDAT);
                        }
                        oForm.Items.Item(pVal.ItemUID).Update();
                    }
                    if (pVal.ItemUID == "Mat1")
                    {
                        if (pVal.ColUID == "LineNum" || pVal.ColUID == "AcctCode" || pVal.ColUID == "AcctName" || pVal.ColUID == "ShortNam" || pVal.ColUID == "Debit" || pVal.ColUID == "Credit")
                        {
                            oMat1.FlushToDataSource();
                            oDS_PH_PY116B.Offset = pVal.Row - 1;
                            switch (pVal.ColUID)
                            {
                                case "AcctCode":
                                    oDS_PH_PY116B.SetValue("U_ShortNam", pVal.Row - 1, oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PH_PY116B.SetValue("U_AcctNam", pVal.Row - 1, dataHelpClass.Get_ReData("AcctName", "AcctCode", "OACT", "'" + oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'",""));
                                    break;
                                case "Debit":
                                case "Credit":
                                    oDS_PH_PY116B.SetValue("U_Debit", pVal.Row - 1, String.Format("{0:#,###}", oMat1.Columns.Item("Debit").Cells.Item(pVal.Row).Specific.Value));
                                    oDS_PH_PY116B.SetValue("U_Credit", pVal.Row - 1, String.Format("{0:#,###}", oMat1.Columns.Item("Credit").Cells.Item(pVal.Row).Specific.Value));
                                    break;
                            }
                            oMat1.SetLineData(pVal.Row);
                            TOTAL_AMT();
                            oDS_PH_PY116B.Offset = pVal.Row - 1;
                            if (pVal.Row == oMat1.RowCount && !string.IsNullOrEmpty(oDS_PH_PY116B.GetValue("U_AcctCode", pVal.Row - 1).ToString().Trim()))
                            {
                                PH_PY116_AddMatrixRow();
                                oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
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
                    case "Grid1":
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY116A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY116B);
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
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY116_FormItemEnabled();
                            PH_PY116_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY116_FormItemEnabled();
                            PH_PY116_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY116_FormItemEnabled();
                            PH_PY116_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY116_FormItemEnabled();
                            break;
                        case "1293": //행삭제
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
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY116_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY116'", "");
                if (Convert.ToInt32(DocEntry) == 0)
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY116_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY116_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i;
            
            try
            {
                if (!string.IsNullOrEmpty(oForm.Items.Item("DOCNUM").Specific.Value))
                {
                    if (PSH_Globals.SBO_Application.MessageBox("이미 분개된 자료입니다. 다시 분개하시겠습니까?", 1, "예", "아니오") == 2)
                    {
                        return functionReturnValue;
                    }
                }

                oDocDate = oForm.Items.Item("DOCDATE").Specific.Value;
                oREMARK = oForm.Items.Item("REMARK").Specific.Value;
                
                if (string.IsNullOrEmpty(oDocDate))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("전기일은 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("DOCDATE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (oForm.Items.Item("TOTDEB").Specific.Value == 0 & oForm.Items.Item("TOTCRE").Specific.Value)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("금액이 0 입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("TOTDEB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (oForm.Items.Item("TOTDEB").Specific.Value != oForm.Items.Item("TOTCRE").Specific.Value)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("차변과 대변금액이 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (oForm.Items.Item("TOTDEB").Specific.Value != oForm.Items.Item("TOTPAY").Specific.Value)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("차변과 총지급액이 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();

                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 0; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        oDS_PH_PY116B.Offset = i;
                        if (string.IsNullOrEmpty(oDS_PH_PY116B.GetValue("U_AcctCode", i).ToString().Trim()) && (Convert.ToDouble(oDS_PH_PY116B.GetValue("U_Debit", i)) != 0 || Convert.ToDouble(oDS_PH_PY116B.GetValue("U_Credit", i)) != 0))
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("계정 코드는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            oMat1.Columns.Item("AcctCode").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                        else if (!string.IsNullOrEmpty(oDS_PH_PY116B.GetValue("U_AcctCode", i).ToString().Trim()) && (Convert.ToDouble(oDS_PH_PY116B.GetValue("U_Debit", i)) == 0 && Convert.ToDouble(oDS_PH_PY116B.GetValue("U_Credit", i)) == 0))
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("차변과 대변금액이 0 입니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            oMat1.Columns.Item("AcctCode").Cells.Item(i + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("분개 자료가 없습니다. 확인하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return functionReturnValue;
                }

                oMat1.FlushToDataSource();
                //Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY116B.Size > 1)
                {
                    oDS_PH_PY116B.RemoveRecord(oDS_PH_PY116B.Size - 1);
                }
                oMat1.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY116_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY116_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            
            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY116A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
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
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNumm == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else 
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 매트릭스 행 추가
        /// </summary>
        public void PH_PY116_AddMatrixRow()
        {
            int oRow;

            try
            {
                oForm.Freeze(true);

                //[Mat1 용]
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY116B.GetValue("U_AcctCode", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY116B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY116B.InsertRecord(oRow);
                        }
                        oDS_PH_PY116B.Offset = oRow;
                        oDS_PH_PY116B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY116B.SetValue("U_AcctCode", oRow, "");
                        oDS_PH_PY116B.SetValue("U_AcctName", oRow, "");
                        oDS_PH_PY116B.SetValue("U_ShortNam", oRow, "");
                        oDS_PH_PY116B.SetValue("U_Debit", oRow, "0");
                        oDS_PH_PY116B.SetValue("U_Credit", oRow, "0");
                        oDS_PH_PY116B.SetValue("U_Project", oRow, "");
                        oDS_PH_PY116B.SetValue("U_CostCent", oRow, "");
                        oDS_PH_PY116B.SetValue("U_Comments", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY116B.Offset = oRow - 1;
                        oDS_PH_PY116B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY116B.SetValue("U_AcctCode", oRow - 1, "");
                        oDS_PH_PY116B.SetValue("U_AcctName", oRow - 1, "");
                        oDS_PH_PY116B.SetValue("U_ShortNam", oRow - 1, "");
                        oDS_PH_PY116B.SetValue("U_Debit", oRow - 1, "0");
                        oDS_PH_PY116B.SetValue("U_Credit", oRow - 1, "0");
                        oDS_PH_PY116B.SetValue("U_Project", oRow - 1, "");
                        oDS_PH_PY116B.SetValue("U_CostCent", oRow - 1, "");
                        oDS_PH_PY116B.SetValue("U_Comments", oRow - 1, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY116B.Offset = oRow;
                    oDS_PH_PY116B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY116B.SetValue("U_AcctCode", oRow, "");
                    oDS_PH_PY116B.SetValue("U_AcctName", oRow, "");
                    oDS_PH_PY116B.SetValue("U_ShortNam", oRow, "");
                    oDS_PH_PY116B.SetValue("U_Debit", oRow, "0");
                    oDS_PH_PY116B.SetValue("U_Credit", oRow, "0");
                    oDS_PH_PY116B.SetValue("U_Project", oRow, "");
                    oDS_PH_PY116B.SetValue("U_CostCent", oRow, "");
                    oDS_PH_PY116B.SetValue("U_Comments", oRow, "");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY116_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DI_oJournalEntries
        /// </summary>
        private void DI_oJournalEntries()
        {
            SAPbobsCOM.JournalVouchers f_oJournalEntries = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sQry;
            int RetVal;
            int nErr = 0;
            string ErrMsg = string.Empty;
            int i;
            int k;
            string AcctCode;
            string shortCode;
            double Credit;
            double Debit;
            string LineMemo;
            string Project;
            string Dimenz1;
            string Dimenz2;
            string Dimenz3 = string.Empty;
            string errCode = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oDocDate.ToString().Trim()))
                {
                    oDocDate = DateTime.Now.ToString("yyyyMMdd");
                }
                else
                {
                    oDocDate = string.Format("yyyy-MM-dd", oDocDate);
                }
                ///// 분개장문서

                PSH_Globals.oCompany.StartTransaction();
                var _with1 = f_oJournalEntries;
                //// 전표전체적용
                //.Series = "14"             '/ 시리즈:주요 분개개체번호
                //.Series = MDC_SetMod.Get_Series_No("30")    '시리즈:주요 분개개체번호
                _with1.JournalEntries.DueDate = Convert.ToDateTime(oDocDate);                //"04/26/2007"   '/ 만기일
                _with1.JournalEntries.TaxDate = Convert.ToDateTime(oDocDate);                //"04/26/2007"   '/ 과세일
                _with1.JournalEntries.TaxDate = Convert.ToDateTime(oDocDate);                //"04/26/2007" '전기일"
                _with1.JournalEntries.ReferenceDate = Convert.ToDateTime(oDocDate);          //"04/26/2007" '전기일"
                _with1.JournalEntries.Memo = oREMARK.ToString().Trim();
                // .ProjectCode = "103"     '/ 프로젝트코드
                //.AutoVAT = tYES          '/ 부가세코드사용
                i = 1;
                oMat1.FlushToDataSource();
                for (k = 0; k <= oMat1.VisualRowCount - 1; k++)
                {
                    oDS_PH_PY116B.Offset = k;
                    if (!string.IsNullOrEmpty(oDS_PH_PY116B.GetValue("U_Col02", k).ToString().Trim()))
                    {
                        if (i != 1)
                        {
                            _with1.JournalEntries.Lines.Add();
                            _with1.JournalEntries.Lines.SetCurrentLine((k));
                        }
                        AcctCode = oDS_PH_PY116B.GetValue("U_Col02", k).ToString().Trim();
                        shortCode = oDS_PH_PY116B.GetValue("U_Col04", k).ToString().Trim();
                        Debit = Convert.ToDouble(oDS_PH_PY116B.GetValue("U_Col05", k).Replace(",", ""));
                        Credit = Convert.ToDouble(oDS_PH_PY116B.GetValue("U_Col06", k).Replace(",", ""));
                        LineMemo = oDS_PH_PY116B.GetValue("U_Col09", k).ToString().Trim();
                        Project = oDS_PH_PY116B.GetValue("U_Col11", k).ToString().Trim();
                        Dimenz1 = oDS_PH_PY116B.GetValue("U_Col07", k).ToString().Trim();
                        Dimenz2 = oDS_PH_PY116B.GetValue("U_Col08", k).ToString().Trim();
                        if (oDIM3 == "Y")
                        {
                            Dimenz3 = oDS_PH_PY116B.GetValue("U_Col10", k).ToString().Trim();
                        }

                        _with1.JournalEntries.Lines.AccountCode = dataHelpClass.Get_ReData("AcctCode", "FormatCode", "[OACT]", "'" + AcctCode + "'","");
                        if (dataHelpClass.Value_ChkYn("[OACT]", "FormatCode", "'" + shortCode.ToString().Trim() + "'","") == false)
                        {
                            _with1.JournalEntries.Lines.ShortName = dataHelpClass.Get_ReData("AcctCode", "FormatCode", "[OACT]", "'" + shortCode.ToString().Trim() + "'","");
                        }
                        else
                        {
                            _with1.JournalEntries.Lines.ShortName = shortCode.ToString().Trim();
                        }
                        _with1.JournalEntries.Lines.Credit = Credit;
                        _with1.JournalEntries.Lines.Debit = Debit;
                        _with1.JournalEntries.Lines.ProjectCode = Project.ToString().Trim();                        //프로젝트
                        _with1.JournalEntries.Lines.CostingCode = Dimenz1.ToString().Trim();                        //차원1
                        _with1.JournalEntries.Lines.CostingCode2 = Dimenz2.ToString().Trim();                       //차원2
                        _with1.JournalEntries.Lines.LineMemo = LineMemo.ToString().Trim();
                        if (oDIM3 == "Y")
                        {
                            _with1.JournalEntries.Lines.CostingCode3 = Dimenz3.ToString().Trim();
                            //차원3
                        }
                        i += 1;
                    }
                }
                RetVal = _with1.JournalEntries.Add();
                RetVal = _with1.Add();

                //Check Error
                if (0 != RetVal)
                {
                    PSH_Globals.oCompany.GetLastError(out nErr, out ErrMsg);
                    errCode = "1";
                    throw new Exception();
                }
                PSH_Globals.oCompany.GetNewObjectCode(out oDocNum);

                sQry = "EXEC ZPY309_INSERT '" + oDocNum + "', 'ZPY405', '" + oDocDate.ToString().Trim() + "', " + oTOTDEB + ", " + oTOTCRE + ", '" + oSTRDAT.ToString().Trim() + "', '" + oENDDAT.ToString().Trim() + "', '" + oJOBTYP.ToString().Trim() + "', '', '" + oCLTCOD.ToString().Trim() + "', '" + oMSTCOD.ToString().Trim() + "', '', ''";
                oRecordSet.DoQuery(sQry);

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                oForm.Items.Item("DOCNUM").Specific.Value = codeHelpClass.Left(oDocNum, Convert.ToInt32(oDocNum.Length.ToString()) - 1) + "-" + codeHelpClass.Right(oDocNum, 1);
                oForm.Items.Item("DocNum").Update();
                PSH_Globals.SBO_Application.StatusBar.SetText("분개장 문서가 생성되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + nErr + "]" + ErrMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("DI_oJournalEntries_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(codeHelpClass);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(dataHelpClass);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(f_oJournalEntries);
            }
        }

        /// <summary>
        /// Execution_Process
        /// </summary>
        private bool Execution_Process()
        {
            string sQry;
            bool functionReturnValue = false;
            short ErrNum = 0;
            int i;
            
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oSTRDAT = oDS_PH_PY116A.GetValue("U_STRDAT", 0).ToString().Trim();
                oENDDAT = oDS_PH_PY116A.GetValue("U_ENDDAT", 0).ToString().Trim();
                oCLTCOD = oDS_PH_PY116A.GetValue("U_CLTCOD", 0).ToString().Trim();
                oMSTCOD = oDS_PH_PY116A.GetValue("U_MSTCOD", 0).ToString().Trim();

                if (string.IsNullOrEmpty(oMSTCOD))
                {
                    oMSTCOD = "%";
                }

                if (string.IsNullOrEmpty(oSTRDAT) || string.IsNullOrEmpty(oENDDAT))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                sQry = "EXEC ZPY309_QUERY 'ZPY405', '" + oSTRDAT + "', '" + oENDDAT + "', '" + oJOBTYP + "', '', '" + oCLTCOD + "', '" + oMSTCOD + "', '', '%'";

                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    oDocNum = oRecordSet.Fields.Item("DOCNUM").Value;
                }
                else
                {
                    oDocNum = "";
                }
                oForm.Items.Item("DocNum").Specific.Value = oDocNum;

                oDS_PH_PY116B.Clear();
                oMat1.LoadFromDataSource();
                i = 0;

                if (oJOBTYP == "R01")
                {
                    sQry = " EXEC ZPY405_1 '" + oSTRDAT + "', '" + oENDDAT + "'," + "'" + oCLTCOD + "', '" + oMSTCOD + "'";
                }
                else
                {
                    sQry = " EXEC ZPY405_2 '" + oSTRDAT + "', '" + oENDDAT + "'," + "'" + oCLTCOD + "', '" + oMSTCOD + "'";
                }
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount == 0)
                {
                    ErrNum = 2;
                    throw new Exception();
                }
                oTOTDEB = 0;
                oTOTCRE = 0;
                while (!oRecordSet.EoF)
                {
                    oDS_PH_PY116B.InsertRecord(i);
                    oDS_PH_PY116B.Offset = i;
                    oDS_PH_PY116B.SetValue("U_Col01", i, Convert.ToString(i + 1));
                    oDS_PH_PY116B.SetValue("U_Col02", i, oRecordSet.Fields.Item("AcctCode").Value);
                    oDS_PH_PY116B.SetValue("U_Col03", i, oRecordSet.Fields.Item("AcctName").Value);
                    oDS_PH_PY116B.SetValue("U_Col04", i, oRecordSet.Fields.Item("ShortName").Value);
                    oDS_PH_PY116B.SetValue("U_Col05", i, oRecordSet.Fields.Item("Debit").Value); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("Debit").Value, "#,###,###,##0"));
                    //총급여액
                    oDS_PH_PY116B.SetValue("U_Col06", i, oRecordSet.Fields.Item("Credit").Value); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("Credit").Value, "#,###,###,##0"));
                    oDS_PH_PY116B.SetValue("U_Col08", i, oRecordSet.Fields.Item("U_PNLCOD").Value);
                    oDS_PH_PY116B.SetValue("U_Col09", i, oRecordSet.Fields.Item("U_Remark").Value);
                    //차원3이 활성화된 경우에만...
                    if (oDIM3 == "Y")
                    {
                        oDS_PH_PY116B.SetValue("U_Col10", i, oRecordSet.Fields.Item("U_DIM3CD").Value);
                    }
                    oDS_PH_PY116B.SetValue("U_Col11", i, oRecordSet.Fields.Item("U_PRJCOD").Value);
                    oTOTDEB += oRecordSet.Fields.Item("Debit").Value;
                    oTOTCRE += oRecordSet.Fields.Item("Credit").Value;
                    i += 1;
                    oRecordSet.MoveNext();
                }
                oMat1.LoadFromDataSource();

                PH_PY116_AddMatrixRow();

                oForm.Items.Item("TOTPAY").Specific.Value = oTOTDEB;
                oForm.Items.Item("TOTGON").Specific.Value = oTOTCRE;
                oForm.Items.Item("TOTDEB").Specific.Value = oTOTDEB;
                oForm.Items.Item("TOTCRE").Specific.Value = oTOTCRE;

                PSH_Globals.SBO_Application.StatusBar.SetText("작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지급일자는 필수입니다. 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조건과 일치하는 자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("Execution_Process_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return functionReturnValue;
        }

        /// <summary>
        /// TOTAL_AMT
        /// </summary>
        private void TOTAL_AMT()
        {
            short oRow;
            oTOTDEB = 0;
            oTOTCRE = 0;
            try
            {
                for (oRow = 1; oRow <= oMat1.VisualRowCount; oRow++)
                {
                    oDS_PH_PY116B.Offset = oRow - 1;
                    oTOTDEB = oTOTDEB + Convert.ToDouble(oDS_PH_PY116B.GetValue("U_Debit", oRow - 1).Replace(",", ""));
                    oTOTCRE = oTOTCRE + Convert.ToDouble(oDS_PH_PY116B.GetValue("U_Credit", oRow - 1).Replace(",", ""));
                }
                oForm.Items.Item("TOTDEB").Specific.Value = oTOTDEB;
                oForm.Items.Item("TOTCRE").Specific.Value = oTOTCRE;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("TOTAL_AMT_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}

