
using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 급(상)여 분개장 생성
    /// </summary>
    internal class PH_PY113 : PSH_BaseClass
    {
        public string oFormUniqueID;

        public SAPbouiCOM.Matrix oMat1;
        private SAPbouiCOM.DBDataSource oDS_PH_PY113A;
        private SAPbouiCOM.DBDataSource oDS_PH_PY113B;

        private string oLastItemUID;     //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow;         //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        private string oCLTCOD;
        private string oYM;
        private string oJOBTYP;
        private string oJOBGBN;
        private string oPAYSEL;
        private double oTOTDEB;
        private double oTOTCRE;
        private double oTOTPAY;
        private double oTOTGON;

        private string oDocDate;
        private string oREMARK;
        private string oDocNum;
        private string oDIM3;

        public string ItemUID { get; private set; }

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY113.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY113_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY113");

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
                PH_PY113_CreateItems();
                PH_PY113_EnableMenus();
                PH_PY113_SetDocument(oFromDocEntry01);
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
        private void PH_PY113_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                
                ////----------------------------------------------------------------------------------------------
                //// 데이터셋정의
                ////----------------------------------------------------------------------------------------------
                //    '//테이블이 있을경우 데이터셋(Matrix)
                oDS_PH_PY113A = oForm.DataSources.DBDataSources.Item("@PH_PY113A");
                ////헤더
                oDS_PH_PY113B = oForm.DataSources.DBDataSources.Item("@PH_PY113B");
                ////라인

                oMat1 = oForm.Items.Item("Mat1").Specific;

                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat1.AutoResizeColumns();

                ////----------------------------------------------------------------------------------------------
                //// 아이템 설정
                ////----------------------------------------------------------------------------------------------
                ////사업장
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                ////지급종류
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("1", "급여");
                oForm.Items.Item("JOBTYP").Specific.ValidValues.Add("2", "상여");
                oForm.Items.Item("JOBTYP").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("JOBTYP").DisplayDesc = true;

                ////지급구분
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("JOBGBN").Specific, "");
                oForm.Items.Item("JOBGBN").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("JOBGBN").DisplayDesc = true;


                ////지급대상자구분
                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' ORDER BY CAST(U_Code AS NUMERIC) ";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("PAYSEL").Specific, "");
                oForm.Items.Item("PAYSEL").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("PAYSEL").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oForm.Items.Item("PAYSEL").DisplayDesc = true;

                //// 전기일자
                oDS_PH_PY113A.SetValue("U_DOCDATE", 0, DateTime.Now.ToString("yyyyMMdd"));

                /// Matrix
                oMat1 = oForm.Items.Item("Mat1").Specific;

                oMat1.Columns.Item("AcctCode").ExtendedObject.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private void PH_PY113_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false);                ////제거
                oForm.EnableMenu("1284", false);                ////취소
                oForm.EnableMenu("1293", true);                ////행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY113_SetDocument(string oFromDocEntry01)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if ((string.IsNullOrEmpty(oFromDocEntry01)))
                {
                    PH_PY113_FormItemEnabled();
                    PH_PY113_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY113_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFromDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY113_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY113_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
				    PH_PY113_FormClear();
				    oForm.ActiveItem = "CLTCOD";
				    oForm.Items.Item("DocEntry").Enabled = false;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

				    //// 귀속년월
				    oDS_PH_PY113A.SetValue("U_YM", 0, DateTime.Now.ToString("yyyyMM"));
				    ////지급종류
                    oForm.Items.Item("JOBTYP").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
				    ////지급구분
                    oForm.Items.Item("JOBGBN").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
				    ////지급대상자구분
                    oForm.Items.Item("PAYSEL").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				    oForm.EnableMenu("1281", true);				////문서찾기
				    oForm.EnableMenu("1282", false);    		////문서추가

			    }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
				    oForm.Items.Item("DocEntry").Enabled = true;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

				    oForm.EnableMenu("1281", false);				////문서찾기
				    oForm.EnableMenu("1282", true);				////문서추가
			    }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
				    oForm.Items.Item("DocEntry").Enabled = false;

                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

				    oForm.EnableMenu("1281", true);				////문서찾기
				    oForm.EnableMenu("1282", true);				////문서추가
			    }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                        if (PH_PY113_DataValidCheck() == false)
                        {
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn2")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
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
                        else
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("추가 모드에서만 조회가 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (pVal.ActionSuccess == true)
                        {
                            PH_PY113_FormItemEnabled();
                        }
                    }

                    /// 분개장 문서생성
                    if (pVal.ItemUID == "Btn1")
                    {
                        if (!string.IsNullOrEmpty(oDS_PH_PY113A.GetValue("DocEntry", 0)))
                        {
                            if (dataHelpClass.Value_ChkYn("[@PH_PY113A]", "DocEntry", "'" + oDS_PH_PY113A.GetValue("DocEntry", 0).ToString().Trim() + "'","") == false)
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
                    if (dataHelpClass.Value_ChkYn( "[OACT]", "FormatCode", "'" + oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String + "'", "") == true | string.IsNullOrEmpty(oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
                    {
                        PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
                        BubbleEvent = false;
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
        /// Raise_EVENT_VALIDATE
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (pVal.BeforeAction == false & pVal.ItemChanged == true)
                {
                    if (pVal.ItemUID == "ENDDAT" | pVal.ItemUID == "MSTCOD")
                    {
                        if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.String))
                        {
                            oForm.Items.Item(pVal.ItemUID).Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                        }
                        oForm.Items.Item(pVal.ItemUID).Update();
                    }
                }
                else if (pVal.BeforeAction == false & pVal.ItemChanged == true)
                {
                    if (pVal.ItemUID == "Mat1")
                    {
                        if ((pVal.ColUID == "AcctCode" | pVal.ColUID == "AcctName" | pVal.ColUID == "ShortNam" | pVal.ColUID == "Debit" | pVal.ColUID == "Credit"))
                        {

                            oMat1.FlushToDataSource();
                            oDS_PH_PY113B.Offset = pVal.Row - 1;
                            switch (pVal.ColUID)
                            {
                                case "AcctCode":
                                    //계정코드
                                    oDS_PH_PY113B.SetValue("U_Col04", pVal.Row - 1, oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PH_PY113B.SetValue("U_Col12", pVal.Row - 1, dataHelpClass.Get_ReData("AcctName", "AcctCode", "OACT", "'" + oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'",""));
                                    break;
                                case "Debit":
                                case "Credit":
                                    oDS_PH_PY113B.SetValue("U_Col05", pVal.Row - 1, String.Format("{0:#,###}", oMat1.Columns.Item("Debit").Cells.Item(pVal.Row).Specific.Value));
                                    oDS_PH_PY113B.SetValue("U_Col06", pVal.Row - 1, String.Format("{0:#,###}", oMat1.Columns.Item("Credit").Cells.Item(pVal.Row).Specific.Value));
                                    break;
                            }
                            oMat1.SetLineData(pVal.Row);
                            TOTAL_AMT();
                            oDS_PH_PY113B.Offset = pVal.Row - 1;
                            if (pVal.Row == oMat1.RowCount & !string.IsNullOrEmpty(oDS_PH_PY113B.GetValue("U_AcctCode", pVal.Row - 1).ToString().Trim()))
                            {
                                PH_PY113_AddMatrixRow();
                                oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
        /// Raise_EVENT_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {

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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY113A);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY113B);
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
                            break;
                    }
                }
                else if ((pVal.BeforeAction == false))
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY113_FormItemEnabled();
                            PH_PY113_AddMatrixRow();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;

                        case "1281":
                            ////문서찾기
                            PH_PY113_FormItemEnabled();
                            PH_PY113_AddMatrixRow();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282":
                            ////문서추가
                            PH_PY113_FormItemEnabled();
                            PH_PY113_AddMatrixRow();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY113_FormItemEnabled();
                            break;
                        case "1293":
                            //// 행삭제
                            //// [MAT1 용]
                            if (oMat1.RowCount != oMat1.VisualRowCount)
                            {
                                oMat1.FlushToDataSource();

                                while ((i <= oDS_PH_PY113B.Size - 1))
                                {
                                    if (string.IsNullOrEmpty(oDS_PH_PY113B.GetValue("U_AcctCode", i)))
                                    {
                                        oDS_PH_PY113B.RemoveRecord((i));
                                        i = 0;
                                    }
                                    else
                                    {
                                        i = i + 1;
                                    }
                                }

                                for (i = 0; i <= oDS_PH_PY113B.Size; i++)
                                {
                                    oDS_PH_PY113B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                                }

                                oMat1.LoadFromDataSource();
                            }
                            PH_PY113_AddMatrixRow();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm.Freeze(false);
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
            int i = 0;
            string sQry = string.Empty;

            try
            {
                if ((BusinessObjectInfo.BeforeAction == true))
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
                else if ((BusinessObjectInfo.BeforeAction == false))
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
        private void PH_PY113_FormClear()
        {
            string DocEntry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY113'", "");
                if (Convert.ToInt32(DocEntry) == 0)
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY113_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY113_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
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

                ////----------------------------------------------------------------------------------
                ////필수 체크
                ////----------------------------------------------------------------------------------
                if (string.IsNullOrEmpty(oDS_PH_PY113A.GetValue("U_DOCDATE", 0).ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("전기일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("DOCDATE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (Convert.ToDouble(oDS_PH_PY113A.GetValue("U_TOTDEB", 0).ToString().Trim()) == 0)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("차변합계가 0입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("TOTDEB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (Convert.ToDouble(oDS_PH_PY113A.GetValue("U_TOTCRE", 0).ToString().Trim()) == 0)
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("대변합계가 0입니다. ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("TOTCRE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (oDS_PH_PY113A.GetValue("U_TOTDEB", 0).ToString().Trim() != oDS_PH_PY113A.GetValue("U_TOTCRE", 0).ToString().Trim())
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("차변과 대변금액이 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (oDS_PH_PY113A.GetValue("U_TOTDEB", 0).ToString().Trim() != oDS_PH_PY113A.GetValue("U_TOTPAY", 0).ToString().Trim())
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("차변과 총지급액이 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                if (oMat1.VisualRowCount > 1)
                {
                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
                    {
                        if (string.IsNullOrEmpty(oMat1.Columns.Item("U_AcctCode").Cells.Item(i).Specific.Value))
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("계정과목코드 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("U_AcctCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }

                        if (Convert.ToInt32(oMat1.Columns.Item("U_Debit").Cells.Item(i).Specific.Value) == 0)
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("차변금액이 0 입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("U_Debit").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }

                        if (Convert.ToInt32(oMat1.Columns.Item("U_Credit").Cells.Item(i).Specific.Value) == 0)
                        {
                            PSH_Globals.SBO_Application.SetStatusBarMessage("대변금액이 0 입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            oMat1.Columns.Item("U_Credit").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            return functionReturnValue;
                        }
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("분개 자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    return functionReturnValue;
                }
                oMat1.FlushToDataSource();
                //// Matrix 마지막 행 삭제(DB 저장시)
                if (oDS_PH_PY113B.Size > 1)
                    oDS_PH_PY113B.RemoveRecord((oDS_PH_PY113B.Size - 1));
                oMat1.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                functionReturnValue = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY113_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            functionReturnValue = true;

            short ErrNumm = 0;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY113A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    functionReturnValue = false;
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
        /// 매트릭스 행 추가
        /// </summary>
        public void PH_PY113_AddMatrixRow()
        {
            int oRow = 0;
            try
            {
                oForm.Freeze(true);
                oMat1.FlushToDataSource();
                oRow = oMat1.VisualRowCount;

                if (oMat1.VisualRowCount > 0)
                {
                    if (!string.IsNullOrEmpty(oDS_PH_PY113B.GetValue("U_AcctCode", oRow - 1).ToString().Trim()))
                    {
                        if (oDS_PH_PY113B.Size <= oMat1.VisualRowCount)
                        {
                            oDS_PH_PY113B.InsertRecord((oRow));
                        }
                        oDS_PH_PY113B.Offset = oRow;
                        oDS_PH_PY113B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                        oDS_PH_PY113B.SetValue("U_AcctCode", oRow, "");
                        oDS_PH_PY113B.SetValue("U_AcctName", oRow, "");
                        oDS_PH_PY113B.SetValue("U_ShortNam", oRow, "");
                        oDS_PH_PY113B.SetValue("U_Debit", oRow, Convert.ToString(0));
                        oDS_PH_PY113B.SetValue("U_Credit", oRow, Convert.ToString(0));
                        oDS_PH_PY113B.SetValue("U_Project", oRow, "");
                        oDS_PH_PY113B.SetValue("U_CostCent", oRow, "");
                        oDS_PH_PY113B.SetValue("U_Comments", oRow, "");
                        oMat1.LoadFromDataSource();
                    }
                    else
                    {
                        oDS_PH_PY113B.Offset = oRow - 1;
                        oDS_PH_PY113B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
                        oDS_PH_PY113B.SetValue("U_AcctCode", oRow - 1, "");
                        oDS_PH_PY113B.SetValue("U_AcctName", oRow - 1, "");
                        oDS_PH_PY113B.SetValue("U_ShortNam", oRow - 1, "");
                        oDS_PH_PY113B.SetValue("U_Debit", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY113B.SetValue("U_Credit", oRow - 1, Convert.ToString(0));
                        oDS_PH_PY113B.SetValue("U_Project", oRow - 1, "");
                        oDS_PH_PY113B.SetValue("U_CostCent", oRow - 1, "");
                        oDS_PH_PY113B.SetValue("U_Comments", oRow - 1, "");
                        oMat1.LoadFromDataSource();
                    }
                }
                else if (oMat1.VisualRowCount == 0)
                {
                    oDS_PH_PY113B.Offset = oRow;
                    oDS_PH_PY113B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                    oDS_PH_PY113B.SetValue("U_AcctCode", oRow, "");
                    oDS_PH_PY113B.SetValue("U_AcctName", oRow, "");
                    oDS_PH_PY113B.SetValue("U_ShortNam", oRow, "");
                    oDS_PH_PY113B.SetValue("U_Debit", oRow, Convert.ToString(0));
                    oDS_PH_PY113B.SetValue("U_Credit", oRow, Convert.ToString(0));
                    oDS_PH_PY113B.SetValue("U_Project", oRow, "");
                    oDS_PH_PY113B.SetValue("U_CostCent", oRow, "");
                    oDS_PH_PY113B.SetValue("U_Comments", oRow, "");
                    oMat1.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_AddMatrixRow_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
            string sQry = string.Empty;
            int RetVal = 0;
            int nErr = 0;
            string ErrMsg = string.Empty;
            int i = 0;
            int k = 0;
            string AcctCode = string.Empty;
            string shortCode = string.Empty;
            double Credit = 0;
            double Debit = 0;
            string LineMemo = string.Empty;
            string Project = string.Empty;
            string Dimenz1 = string.Empty;
            string Dimenz2 = string.Empty;
            string Dimenz3 = string.Empty;

            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (string.IsNullOrEmpty(oDocDate.ToString().Trim()))
                {
                    oDocDate = DateTime.Now.ToString("yyyy-MM-dd");
                }
                else
                {
                    oDocDate = Convert.ToDateTime(oDocDate).ToString("yyyy-MM-dd");
                }

                // 재무관리>분개 =계정정보
                PSH_Globals.oCompany.StartTransaction();
                var _with1 = f_oJournalEntries;
                // 전표전체적용
                _with1.JournalEntries.DueDate = Convert.ToDateTime(oDocDate);                       //"04/26/2007"   '/ 만기일
                _with1.JournalEntries.TaxDate = Convert.ToDateTime(oDocDate);                       //"04/26/2007"   '/ 과세일
                _with1.JournalEntries.ReferenceDate = Convert.ToDateTime(oDocDate);                 //"04/26/2007"   '/ 전기일
                _with1.JournalEntries.Memo = oREMARK.ToString().Trim();

                i = 1;
                oMat1.FlushToDataSource();
                for (k = 0; k <= oMat1.VisualRowCount - 1; k++)
                {
                    oDS_PH_PY113B.Offset = k;
                    if (!string.IsNullOrEmpty(oDS_PH_PY113B.GetValue("U_AcctCode", k).ToString().Trim()))
                    {
                        if (i != 1)
                        {
                            _with1.JournalEntries.Lines.Add();
                            _with1.JournalEntries.Lines.SetCurrentLine((k));
                        }
                        AcctCode = oDS_PH_PY113B.GetValue("U_AcctCode", k);
                        shortCode = oDS_PH_PY113B.GetValue("U_ShortNam", k).ToString().Trim();
                        Debit = Convert.ToDouble(oDS_PH_PY113B.GetValue("U_Debit", k).Replace(",", ""));
                        Credit = Convert.ToDouble(oDS_PH_PY113B.GetValue("U_Credit", k).Replace(",", ""));
                        LineMemo = oDS_PH_PY113B.GetValue("U_Comments", k).ToString().Trim();
                        Project = oDS_PH_PY113B.GetValue("U_Prject", k).ToString().Trim();
                        _with1.JournalEntries.Lines.AccountCode = dataHelpClass.Get_ReData("AcctCode", "FormatCode", "[OACT]", "'" + AcctCode.ToString().Trim() + "'","");
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
                        _with1.JournalEntries.Lines.ProjectCode = Project.ToString().Trim();
                        _with1.JournalEntries.Lines.LineMemo = LineMemo.ToString().Trim();
                        i = i + 1;
                    }
                }
                RetVal = _with1.Add();

                //Check Error
                if ((0 != RetVal))
                {
                    PSH_Globals.oCompany.GetLastError(out nErr, out ErrMsg);
                    throw new Exception();
                    //저장시 에러 발생
                }
                PSH_Globals.oCompany.GetNewObjectCode(out oDocNum);

                sQry = "EXEC PH_PY113_INSERT '" + oDocNum.ToString().Trim() + "', 'PH_PY113', '" + oDocDate.ToString().Trim() + "', " + oTOTDEB + ", " + oTOTCRE + ", '" + oYM.ToString().Trim() + "', '" + oJOBGBN.ToString().Trim() + "', '" + oJOBTYP.ToString().Trim() + "', '" + oPAYSEL.ToString().Trim() + "', '" + oCLTCOD.ToString().Trim() + "'";

                oRecordSet.DoQuery(sQry);

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                oForm.Items.Item("DOCNUM").Specific.Value = codeHelpClass.Left(oDocNum, Convert.ToInt32(oDocNum.Length.ToString()) - 1) + "-" + codeHelpClass.Right(oDocNum, 1);
                oForm.Items.Item("DOCNUM").Update();

                // MsgBox ("완료!")
                PSH_Globals.SBO_Application.StatusBar.SetText("분개장 문서가 생성되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction((SAPbobsCOM.BoWfTransOpt.wf_RollBack));
                }
                PSH_Globals.SBO_Application.StatusBar.SetText("DI_oJournalEntries_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            string sQry = string.Empty;
            bool functionReturnValue = false;
            short ErrNum = 0;
            int i = 0;
            PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oYM = oForm.Items.Item("YM").Specific.Value;
                if (string.IsNullOrEmpty(oYM.ToString().Trim()))
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                oCLTCOD = oDS_PH_PY113A.GetValue("U_CLTCOD", 0);
                oJOBTYP = oDS_PH_PY113A.GetValue("U_JOBTYP", 0);
                oJOBGBN = oDS_PH_PY113A.GetValue("U_JOBGBN", 0);
                oPAYSEL = oDS_PH_PY113A.GetValue("U_PAYSEL", 0);

                // 초기화

                oDS_PH_PY113B.Clear();
                oMat1.LoadFromDataSource();
                i = 0;
                sQry = " EXEC PH_PY113 '" + oCLTCOD.ToString().Trim() + "','" + oYM.ToString().Trim() + "', '" + oJOBTYP.ToString().Trim() + "', '";
                sQry = sQry + oJOBGBN.ToString().Trim() + "', '" + oPAYSEL.ToString().Trim() + "'";

                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount == 0)
                {
                    oForm.Items.Item("DOCNUM").Specific.Value = "";
                    ErrNum = 2;
                    throw new Exception();

                }
                oTOTDEB = 0;
                oTOTCRE = 0;
                while (!(oRecordSet.EoF))
                {
                    oDS_PH_PY113B.InsertRecord((i));
                    oDS_PH_PY113B.Offset = i;
                    oDS_PH_PY113B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY113B.SetValue("U_AcctCode", i, oRecordSet.Fields.Item("AcctCode").Value);
                    oDS_PH_PY113B.SetValue("U_AcctName", i, oRecordSet.Fields.Item("AcctName").Value);
                    oDS_PH_PY113B.SetValue("U_ShortNam", i, oRecordSet.Fields.Item("ShortNam").Value);
                    oDS_PH_PY113B.SetValue("U_Debit", i, String.Format("{0:#,###}", oRecordSet.Fields.Item("Debit").Value));
                    // 총급여액
                    oDS_PH_PY113B.SetValue("U_Credit", i, String.Format("{0:#,###}", oRecordSet.Fields.Item("Credit").Value));
                    oDS_PH_PY113B.SetValue("U_Project", i, oRecordSet.Fields.Item("U_Project").Value);
                    oDS_PH_PY113B.SetValue("U_CostCent", i, oRecordSet.Fields.Item("U_PNLCOD").Value);
                    oDS_PH_PY113B.SetValue("U_Comments", i, oRecordSet.Fields.Item("U_Remark").Value);

                    oTOTDEB = oTOTDEB + oRecordSet.Fields.Item("Debit").Value;
                    oTOTCRE = oTOTCRE + oRecordSet.Fields.Item("Credit").Value;
                    i = i + 1;
                    oRecordSet.MoveNext();
                }
                oMat1.LoadFromDataSource();

                PH_PY113_AddMatrixRow();

                // 분개No 조회
                sQry = "EXEC PH_PY113_QUERY 'PH_PY113', '" + oCLTCOD.ToString().Trim() + "','" + oYM.ToString().Trim() + "', '";
                sQry = sQry + oJOBGBN.ToString().Trim() + "', '" + oJOBTYP.ToString().Trim() + "', '" + oPAYSEL.ToString().Trim() + "'";

                oRecordSet.DoQuery(sQry);
                if (oRecordSet.RecordCount > 0)
                {
                    oDocNum = oRecordSet.Fields.Item("DOCNUM").Value;
                }
                else
                {
                    oDocNum = "";
                }
                oForm.Items.Item("DOCNUM").Specific.Value = oDocNum;
                oForm.Items.Item("TOTPAY").Specific.Value = oTOTDEB;
                oForm.Items.Item("TOTGON").Specific.Value = oTOTCRE;
                oForm.Items.Item("TOTDEB").Specific.Value = oTOTDEB;
                oForm.Items.Item("TOTCRE").Specific.Value = oTOTCRE;

                PSH_Globals.SBO_Application.StatusBar.SetText("작업을 완료하였습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회월을 입력하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 2)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조건과 일치하는 자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                else if (ErrNum == 3)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("자사코드가 없습니다. 선택하여 주십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
            short oRow = 0;
            oTOTDEB = 0;
            oTOTCRE = 0;
            try
            {
                for (oRow = 1; oRow <= oMat1.VisualRowCount; oRow++)
                {
                    oDS_PH_PY113B.Offset = oRow - 1;
                    oTOTDEB = oTOTDEB + Convert.ToDouble(oDS_PH_PY113B.GetValue("U_Debit", oRow - 1).Replace(",", ""));
                    oTOTCRE = oTOTCRE + Convert.ToDouble(oDS_PH_PY113B.GetValue("U_Credit", oRow - 1).Replace(",", ""));
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


//using Microsoft.VisualBasic;
 //using Microsoft.VisualBasic.Compatibility;
 //using System;
 //using System.Collections;
 //using System.Data;
 //using System.Diagnostics;
 //using System.Drawing;
 //using System.Windows.Forms;
 //using SAPbouiCOM;
 //using PSH_BOne_AddOn.Data;

//namespace PSH_BOne_AddOn
//{
//    internal class PH_PY113 : PSH_BaseClass
//    {
//        ////********************************************************************************
//        ////  File           : PH_PY113.cls
//        ////  Module         : 급여관리 > 급(상)여 분개장 생성
//        ////  Desc           : 급(상)여 분개장 생성
//        ////********************************************************************************

//        #region 변수선언
//        public string oFormUniqueID;
//        //public SAPbouiCOM.Form oForm;

//        //'// 매트릭스 사용시
//        public SAPbouiCOM.Matrix oMat1;

//        private SAPbouiCOM.DBDataSource oDS_PH_PY113A;
//        private SAPbouiCOM.DBDataSource oDS_PH_PY113B;

//        private string oLastItemUID;
//        private string oLastColUID;
//        private int oLastColRow;

//        private string oCLTCOD;
//        private string oYM;
//        private string oJOBTYP;
//        private string oJOBGBN;
//        private string oPAYSEL;
//        private double oTOTDEB;
//        private double oTOTCRE;
//        private double oTOTPAY;
//        private double oTOTGON;

//        private string oDocDate;
//        private string oREMARK;
//        private string oDocNum;
//        private string oDIM3;
//        #endregion


//        /// <param name="oFromDocEntry01"></param>
//        public override void LoadForm(string oFromDocEntry01)
//        {
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            try
//            {
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY113.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                for (int loopCount = 1; loopCount <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); loopCount++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[loopCount - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[loopCount - 1].nodeValue = 16;
//                }

//                oFormUniqueID = "PH_PY113_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY113");

//                string strXml = null;
//                strXml = oXmlDoc.xml.ToString();

//                PSH_Globals.SBO_Application.LoadBatchActions(ref strXml);
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                oForm.DataBrowser.BrowseBy = "DocEntry";

//                oForm.Freeze(true);
//                PH_PY113_CreateItems();
//                PH_PY113_EnableMenus();
//                PH_PY113_SetDocument(oFromDocEntry01);
//                //    Call PH_PY113_FormResize
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("LoadForm_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Update();
//                oForm.Freeze(false);
//                oForm.Visible = true;
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
//            }
//        }

//        private void PH_PY113_CreateItems()
//        {
//            string sQry = string.Empty;

//            SAPbouiCOM.CheckBox oCheck = null;
//            //SAPbouiCOM.EditText oEdit = null;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbouiCOM.Column oColumn = null;
//            //SAPbouiCOM.Columns oColumns = null;
//            SAPbouiCOM.OptionBtn optBtn = null;
//            SAPbouiCOM.LinkedButton oLink = null;

//            SAPbobsCOM.Recordset oRecordSet = null;

//            oForm.Freeze(true);

//            try
//            {
//                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//                oDS_PH_PY113A = oForm.DataSources.DBDataSources.Item("@PH_PY113A");
//                oDS_PH_PY113B = oForm.DataSources.DBDataSources.Item("@PH_PY113B");

//                oMat1 = oForm.Items.Item("Mat1").Specific;

//                oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//                oMat1.AutoResizeColumns();

//                //----------------------------------------------------------------------------------------------
//                // 기본사항
//                //----------------------------------------------------------------------------------------------
//                oForm.AutoManaged = true;
//                PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//                dataHelpClass.AutoManaged(oForm, "Code,FullName,CLTCOD");

//                //사업장
//                oCombo = oForm.Items.Item("CLTCOD").Specific;
//                oForm.Items.Item("CLTCOD").DisplayDesc = true;

//                ////지급종류
//                oCombo = oForm.Items.Item("JOBTYP").Specific;
//                oCombo.ValidValues.Add("1", "급여");
//                oCombo.ValidValues.Add("2", "상여");
//                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
//                oForm.Items.Item("JOBTYP").DisplayDesc = true;

//                ////지급구분
//                oCombo = oForm.Items.Item("JOBGBN").Specific;
//                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P212' AND U_UseYN= 'Y'";
//                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
//                oCombo.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
//                oForm.Items.Item("JOBGBN").DisplayDesc = true;


//                ////지급대상자구분
//                oCombo = oForm.Items.Item("PAYSEL").Specific;
//                sQry = " SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code='P213' ORDER BY CAST(U_Code AS NUMERIC) ";
//                dataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "N");
//                oCombo.ValidValues.Add("%", "전체");
//                oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//                oForm.Items.Item("PAYSEL").DisplayDesc = true;

//                //// 전기일자
//                //DateTime.ParseExact(today, "yyyyMMdd", null); //생년월일
//                oDS_PH_PY113A.SetValue("U_DOCDATE", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD"));

//                /// Matrix
//                oMat1 = oForm.Items.Item("Mat1").Specific;

//                oColumn = oMat1.Columns.Item("AcctCode");
//                oLink = oColumn.ExtendedObject;
//                oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GLAccounts;
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//            finally
//            {
//                oForm.Freeze(false);

//                if (oForm.Visible == false)
//                {
//                    oForm.Visible = true;
//                }

//                oForm.Update();
//                //메모리 해제
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCheck);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(optBtn);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//            }
//        }

//        private void PH_PY113_EnableMenus()
//        {
//            try
//            {
//                oForm.EnableMenu("1283", false);                ////제거
//                oForm.EnableMenu("1284", false);                ////취소
//                oForm.EnableMenu("1293", true);                ////행삭제
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_EnableMenus_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//        }

//        private void PH_PY113_SetDocument(string oFromDocEntry01)
//        {
//            try
//            {
//                if (string.IsNullOrEmpty(oFromDocEntry01))
//                {
//                    PH_PY113_FormItemEnabled();
//                    PH_PY113_AddMatrixRow();
//                }
//                else
//                {
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                    PH_PY113_FormItemEnabled();

//                    oForm.Items.Item("DocEntry").Specific.Value = oFromDocEntry01;
//                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY113_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public void PH_PY113_FormItemEnabled()
//        {
//            //SAPbouiCOM.ComboBox oCombo = null;

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                oForm.Freeze(true);
//                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//                {
//                    PH_PY113_FormClear();
//                    oForm.ActiveItem = "CLTCOD";
//                    oForm.Items.Item("DocEntry").Enabled = false;

//                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                    DataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

//                    //// 귀속년월
//                    oDS_PH_PY113A.SetValue("U_YM", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMM"));

//                    ////지급종류
//                    if (oForm.Items.Item("JOBTYP").Specific.ValidValues.Count > 0)
//                    {
//                        oForm.Items.Item("JOBTYP").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_ByValue);
//                    }
//                    //oCombo = oForm.Items.Item("JOBTYP").Specific;
//                    //oCombo.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

//                    ////지급구분
//                    if (oForm.Items.Item("JOBGBN").Specific.ValidValues.Count > 0)
//                    {
//                        oForm.Items.Item("JOBGBN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_ByValue);
//                    }
//                    //oCombo = oForm.Items.Item("JOBGBN").Specific;
//                    //oCombo.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);

//                    ////지급대상자구분
//                    if (oForm.Items.Item("PAYSEL").Specific.ValidValues.Count > 0)
//                    {
//                        oForm.Items.Item("PAYSEL").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);
//                    }
//                    //oCombo = oForm.Items.Item("PAYSEL").Specific;
//                    //oCombo.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

//                    oForm.EnableMenu("1281", true);                    ////문서찾기
//                    oForm.EnableMenu("1282", false);                    ////문서추가
//                }
//                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//                {
//                    oForm.Items.Item("DocEntry").Enabled = true;

//                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                    DataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

//                    oForm.EnableMenu("1281", false);                    ////문서찾기
//                    oForm.EnableMenu("1282", true);                    ////문서추가
//                }
//                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//                {
//                    oForm.Items.Item("DocEntry").Enabled = false;

//                    //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                    DataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

//                    oForm.EnableMenu("1281", true);                    ////문서찾기
//                    oForm.EnableMenu("1282", true);                    ////문서추가
//                }
//                oForm.Freeze(false);
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_FormItemEnabled_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//                //메모리 해제
//                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
//            }
//        }

//        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            string sQry = string.Empty;
//            int i = 0;

//            //SAPbouiCOM.ComboBox oCombo = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            switch (pVal.EventType)
//            {
//                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                    ////1
//                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                    ////2
//                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                    ////3
//                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
//                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
//                    break;

//                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
//                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
//                    break;
//            }
//        }

//        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                    if (pVal.ItemUID == "1")
//                    {
//                        if (PH_PY113_DataValidCheck() == false)
//                        {
//                            BubbleEvent = false;
//                        }
//                    }
//                    if (pVal.ItemUID == "Btn2")
//                    {
//                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                        {
//                            if (Execution_Process() == false)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                            else
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                        }
//                        else
//                        {
//                            PSH_Globals.SBO_Application.SetStatusBarMessage("추가 모드에서만 조회가 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                        }
//                    }
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                    if (pVal.ItemUID == "1")
//                    {
//                        if (pVal.ActionSuccess == true)
//                        {
//                            PH_PY113_FormItemEnabled();
//                        }
//                    }

//                    /// 분개장 문서생성
//                    if (pVal.ItemUID == "Btn1")
//                    {
//                        if (!string.IsNullOrEmpty(oDS_PH_PY113A.GetValue("DocEntry", 0)))
//                        {
//                            if (DataHelpClass.Value_ChkYn("[@PH_PY113A]", "DocEntry", "'" + Strings.Trim(oDS_PH_PY113A.GetValue("DocEntry", 0)) + "'", "") == false)
//                            {
//                                if (DI_oJournalEntries() == false)
//                                {
//                                    //functionReturnValue = false;
//                                    //return functionReturnValue;
//                                }
//                            }
//                            else
//                            {
//                                PSH_Globals.SBO_Application.SetStatusBarMessage("저장된 문서만 분개 생성이 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                            }
//                        }
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_ITEM_PRESSED_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();
//            try
//            {
//                if (pVal.BeforeAction == true & pVal.ColUID == "AcctCode" & pVal.CharPressed == 9)
//                {
//                    if (DataHelpClass.Value_ChkYn("[OACT]", "FormatCode", "'" + oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String + "'", "") == true | string.IsNullOrEmpty(oMat1.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
//                    {
//                        PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
//                        BubbleEvent = false;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_KEY_DOWN_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                switch (pVal.ItemUID)
//                {
//                    case "Mat1":
//                    case "Grid1":
//                        if (pVal.Row > 0)
//                        {
//                            oLastItemUID = pVal.ItemUID;
//                            oLastColUID = pVal.ColUID;
//                            oLastColRow = pVal.Row;
//                        }
//                        break;
//                    default:
//                        oLastItemUID = pVal.ItemUID;
//                        oLastColUID = "";
//                        oLastColRow = 0;
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_GOT_FOCUS_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            string tSex = string.Empty;
//            string errCode = string.Empty;

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (pVal.BeforeAction == false & pVal.ItemChanged == true)
//                {
//                    if (pVal.ItemUID == "ENDDAT" | pVal.ItemUID == "MSTCOD")
//                    {
//                        //UPGRADE_WARNING: oForm.Items(pVal.ItemUid).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.String))
//                        {
//                            //UPGRADE_WARNING: oForm.Items(pVal.ItemUid).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            oForm.Items.Item(pVal.ItemUID).Specific.Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");
//                        }
//                        oForm.Items.Item(pVal.ItemUID).Update();
//                    }
//                }
//                else if (pVal.BeforeAction == false & pVal.ItemChanged == true)
//                {
//                    if (pVal.ItemUID == "Mat1")
//                    {
//                        if ((pVal.ColUID == "AcctCode" | pVal.ColUID == "AcctName" | pVal.ColUID == "ShortNam" | pVal.ColUID == "Debit" | pVal.ColUID == "Credit"))
//                        {

//                            oMat1.FlushToDataSource();
//                            oDS_PH_PY113B.Offset = pVal.Row - 1;
//                            switch (pVal.ColUID)
//                            {
//                                case "AcctCode":
//                                    //계정코드                                    
//                                    oDS_PH_PY113B.SetValue("U_Col04", pVal.Row - 1, oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Specific.Value);
//                                    oDS_PH_PY113B.SetValue("U_Col12", pVal.Row - 1, DataHelpClass.Get_ReData("AcctName", "AcctCode", "OACT", "'" + Strings.Trim(oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Specific.Value) + "'", ""));
//                                    break;
//                                case "Debit":
//                                case "Credit":
//                                    oDS_PH_PY113B.SetValue("U_Col05", pVal.Row - 1, oMat1.Columns.Item("Debit").Cells.Item(pVal.Row).Specific.Value);
//                                    oDS_PH_PY113B.SetValue("U_Col06", pVal.Row - 1, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oMat1.Columns.Item("Credit").Cells.Item(pVal.Row).Specific.Value, "#,###,###,##0"));
//                                    break;
//                            }
//                            oMat1.SetLineData(pVal.Row);
//                            TOTAL_AMT();
//                            oDS_PH_PY113B.Offset = pVal.Row - 1;
//                            if (pVal.Row == oMat1.RowCount & !string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY113B.GetValue("U_AcctCode", pVal.Row - 1))))
//                            {
//                                PH_PY113_AddMatrixRow();
//                                oMat1.Columns.Item("AcctCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            }
//                        }
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                if (errCode == "1")
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("가족사항 - 잘못된 주민번호입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//                else if (errCode == "2")
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("잘못된 주민번호입니다. : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }

//                BubbleEvent = false;
//            }
//            finally
//            {
//            }
//        }

//        private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                    SubMain.Remove_Forms(oFormUniqueID);

//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY113A);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY113B);
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat1);
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_UNLOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
//        {
//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                oForm.Freeze(true);
//                int i = 0;

//                if ((pVal.BeforeAction == true))
//                {
//                    switch (pVal.MenuUID)
//                    {
//                        case "1283":
//                            if (PSH_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2)
//                            {
//                                BubbleEvent = false;
//                                return;
//                            }
//                            break;
//                        case "1284":
//                            break;
//                        case "1286":
//                            break;
//                        case "1293":
//                            break;
//                        case "1281":
//                            break;
//                        case "1282":
//                            break;
//                        case "1288":
//                        case "1289":
//                        case "1290":
//                        case "1291":
//                            break;
//                    }
//                }
//                else if ((pVal.BeforeAction == false))
//                {
//                    switch (pVal.MenuUID)
//                    {
//                        case "1283":
//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                            PH_PY113_FormItemEnabled();
//                            PH_PY113_AddMatrixRow();
//                            break;
//                        case "1284":
//                            break;
//                        case "1286":
//                            break;
//                        //            Case "1293":
//                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
//                        case "1281":
//                            ////문서찾기
//                            PH_PY113_FormItemEnabled();
//                            PH_PY113_AddMatrixRow();
//                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            break;
//                        case "1282":
//                            ////문서추가
//                            PH_PY113_FormItemEnabled();
//                            PH_PY113_AddMatrixRow();
//                            break;
//                        case "1288":
//                        case "1289":
//                        case "1290":
//                        case "1291":
//                            PH_PY113_FormItemEnabled();
//                            break;
//                        case "1293":
//                            //// 행삭제
//                            //// [MAT1 용]
//                            if (oMat1.RowCount != oMat1.VisualRowCount)
//                            {
//                                oMat1.FlushToDataSource();

//                                while ((i <= oDS_PH_PY113B.Size - 1))
//                                {
//                                    if (string.IsNullOrEmpty(oDS_PH_PY113B.GetValue("U_AcctCode", i)))
//                                    {
//                                        oDS_PH_PY113B.RemoveRecord((i));
//                                        i = 0;
//                                    }
//                                    else
//                                    {
//                                        i = i + 1;
//                                    }
//                                }

//                                for (i = 0; i <= oDS_PH_PY113B.Size; i++)
//                                {
//                                    oDS_PH_PY113B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                                }

//                                oMat1.LoadFromDataSource();
//                            }
//                            PH_PY113_AddMatrixRow();
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormMenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        {
//            try
//            {
//                if ((BusinessObjectInfo.BeforeAction == true))
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                            ////33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                            ////34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                            ////35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                            ////36
//                            break;
//                    }
//                }
//                else if ((BusinessObjectInfo.BeforeAction == false))
//                {
//                    switch (BusinessObjectInfo.EventType)
//                    {
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//                            ////33
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//                            ////34
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//                            ////35
//                            break;
//                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//                            ////36
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pVal.BeforeAction == true)
//                {
//                }
//                else if (pVal.BeforeAction == false)
//                {
//                }

//                switch (pVal.ItemUID)
//                {
//                    case "Mat1":
//                        if (pVal.Row > 0)
//                        {
//                            oLastItemUID = pVal.ItemUID;
//                            oLastColUID = pVal.ColUID;
//                            oLastColRow = pVal.Row;
//                        }
//                        break;
//                    default:
//                        oLastItemUID = pVal.ItemUID;
//                        oLastColUID = "";
//                        oLastColRow = 0;
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//            }
//        }

//        public void PH_PY113_AddMatrixRow()
//        {
//            int oRow = 0;
//            oForm.Freeze(true);

//            ////[Mat1 용]
//            oMat1.FlushToDataSource();
//            oRow = oMat1.VisualRowCount;

//            try
//            {
//                if (oMat1.VisualRowCount > 0)
//                {
//                    if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY113B.GetValue("U_AcctCode", oRow - 1))))
//                    {
//                        if (oDS_PH_PY113B.Size <= oMat1.VisualRowCount)
//                        {
//                            oDS_PH_PY113B.InsertRecord((oRow));
//                        }
//                        oDS_PH_PY113B.Offset = oRow;
//                        oDS_PH_PY113B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                        oDS_PH_PY113B.SetValue("U_AcctCode", oRow, "");
//                        oDS_PH_PY113B.SetValue("U_AcctName", oRow, "");
//                        oDS_PH_PY113B.SetValue("U_ShortNam", oRow, "");
//                        oDS_PH_PY113B.SetValue("U_Debit", oRow, Convert.ToString(0));
//                        oDS_PH_PY113B.SetValue("U_Credit", oRow, Convert.ToString(0));
//                        oDS_PH_PY113B.SetValue("U_Project", oRow, "");
//                        oDS_PH_PY113B.SetValue("U_CostCent", oRow, "");
//                        oDS_PH_PY113B.SetValue("U_Comments", oRow, "");
//                        oMat1.LoadFromDataSource();
//                    }
//                    else
//                    {
//                        oDS_PH_PY113B.Offset = oRow - 1;
//                        oDS_PH_PY113B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//                        oDS_PH_PY113B.SetValue("U_AcctCode", oRow - 1, "");
//                        oDS_PH_PY113B.SetValue("U_AcctName", oRow - 1, "");
//                        oDS_PH_PY113B.SetValue("U_ShortNam", oRow - 1, "");
//                        oDS_PH_PY113B.SetValue("U_Debit", oRow - 1, Convert.ToString(0));
//                        oDS_PH_PY113B.SetValue("U_Credit", oRow - 1, Convert.ToString(0));
//                        oDS_PH_PY113B.SetValue("U_Project", oRow - 1, "");
//                        oDS_PH_PY113B.SetValue("U_CostCent", oRow - 1, "");
//                        oDS_PH_PY113B.SetValue("U_Comments", oRow - 1, "");
//                        oMat1.LoadFromDataSource();
//                    }
//                }
//                else if (oMat1.VisualRowCount == 0)
//                {
//                    oDS_PH_PY113B.Offset = oRow;
//                    oDS_PH_PY113B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//                    oDS_PH_PY113B.SetValue("U_AcctCode", oRow, "");
//                    oDS_PH_PY113B.SetValue("U_AcctName", oRow, "");
//                    oDS_PH_PY113B.SetValue("U_ShortNam", oRow, "");
//                    oDS_PH_PY113B.SetValue("U_Debit", oRow, Convert.ToString(0));
//                    oDS_PH_PY113B.SetValue("U_Credit", oRow, Convert.ToString(0));
//                    oDS_PH_PY113B.SetValue("U_Project", oRow, "");
//                    oDS_PH_PY113B.SetValue("U_CostCent", oRow, "");
//                    oDS_PH_PY113B.SetValue("U_Comments", oRow, "");
//                    oMat1.LoadFromDataSource();
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY113_AddMatrixRow_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        public void PH_PY113_FormClear()
//        {
//            string DocEntry = null;
//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY113'", "");
//                if (Convert.ToDouble(DocEntry) == 0)
//                {
//                    oForm.Items.Item("DocEntry").Specific.Value = 1;
//                }
//                else
//                {
//                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY113_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public bool PH_PY113_DataValidCheck()
//        {
//            bool functionReturnValue = false;
//            functionReturnValue = false;

//            int i = 0;
//            string sQry = string.Empty;
//            //SAPbobsCOM.Recordset oRecordSet = null;

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                if (!string.IsNullOrEmpty(oForm.Items.Item("DOCNUM").Specific.Value))
//                {
//                    if (PSH_Globals.SBO_Application.MessageBox("이미 분개된 자료입니다. 다시 분개하시겠습니까?", 1, "예", "아니오") == 2)
//                    {
//                        return functionReturnValue;
//                    }
//                }

//                oDocDate = oForm.Items.Item("DOCDATE").Specific.Value;
//                oREMARK = oForm.Items.Item("REMARK").Specific.Value;

//                ////----------------------------------------------------------------------------------
//                ////필수 체크
//                ////----------------------------------------------------------------------------------
//                if (string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY113A.GetValue("U_DOCDATE", 0))))
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("전기일은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    oForm.Items.Item("DOCDATE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    return functionReturnValue;
//                }

//                if (Convert.ToDouble(Strings.Trim(oDS_PH_PY113A.GetValue("U_TOTDEB", 0))) == 0)
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("차변합계가 0입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    oForm.Items.Item("TOTDEB").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    return functionReturnValue;
//                }

//                if (Convert.ToDouble(Strings.Trim(oDS_PH_PY113A.GetValue("U_TOTCRE", 0))) == 0)
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("대변합계가 0입니다. ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    oForm.Items.Item("TOTCRE").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    return functionReturnValue;
//                }

//                if (Strings.Trim(oDS_PH_PY113A.GetValue("U_TOTDEB", 0)) != Strings.Trim(oDS_PH_PY113A.GetValue("U_TOTCRE", 0)))
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("차변과 대변금액이 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    return functionReturnValue;
//                }

//                if (Strings.Trim(oDS_PH_PY113A.GetValue("U_TOTDEB", 0)) != Strings.Trim(oDS_PH_PY113A.GetValue("U_TOTPAY", 0)))
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("차변과 총지급액이 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                    return functionReturnValue;
//                }

//                if (oMat1.VisualRowCount > 1)
//                {
//                    for (i = 1; i <= oMat1.VisualRowCount - 1; i++)
//                    {
//                        if (string.IsNullOrEmpty(oMat1.Columns.Item("U_AcctCode").Cells.Item(i).Specific.Value))
//                        {
//                            PSH_Globals.SBO_Application.SetStatusBarMessage("계정과목코드 필수입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                            oMat1.Columns.Item("U_AcctCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            return functionReturnValue;
//                        }

//                        if (Conversion.Val(oMat1.Columns.Item("U_Debit").Cells.Item(i).Specific.Value) == 0)
//                        {
//                            PSH_Globals.SBO_Application.SetStatusBarMessage("차변금액이 0 입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                            oMat1.Columns.Item("U_Debit").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            return functionReturnValue;
//                        }

//                        if (Conversion.Val(oMat1.Columns.Item("U_Credit").Cells.Item(i).Specific.Value) == 0)
//                        {
//                            PSH_Globals.SBO_Application.SetStatusBarMessage("대변금액이 0 입니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                            oMat1.Columns.Item("U_Credit").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            return functionReturnValue;
//                        }
//                    }
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("분개 자료가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    return functionReturnValue;
//                }

//                oMat1.FlushToDataSource();
//                //// Matrix 마지막 행 삭제(DB 저장시)
//                if (oDS_PH_PY113B.Size > 1)
//                    oDS_PH_PY113B.RemoveRecord((oDS_PH_PY113B.Size - 1));
//                oMat1.LoadFromDataSource();

//                functionReturnValue = true;
//            }
//            catch (Exception ex)
//            {
//                functionReturnValue = false;
//                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY113_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//            }

//            return functionReturnValue;
//        }


//        public bool PH_PY113_Validate(string ValidateType)
//        {
//            bool functionReturnValue = false;
//            functionReturnValue = true;

//            string errCode = string.Empty;

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            object i = null;
//            int j = 0;

//            try
//            {
//                if (DataHelpClass.Get_ReData("Canceled", "DocEntry", "[@PH_PY113A]", "'" & oForm.Items.Item("DocEntry").Specific.VALUE & "'", "") == "Y")
//                {
//                    errCode = "1";
//                    throw new Exception();
//                }

//                if (ValidateType == "수정")
//                {

//                }
//                else if (ValidateType == "행삭제")
//                {

//                }
//                else if (ValidateType == "취소")
//                {

//                }
//            }
//            catch (Exception ex)
//            {
//                if (errCode == "1")
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY113_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//                functionReturnValue = false;
//            }
//            finally
//            {
//            }
//            return functionReturnValue;
//        }

//        private void PH_PY113_Print_Report01()
//        {

//            string DocNum = string.Empty;
//            int ErrNum = 0;
//            string WinTitle = string.Empty;
//            string ReportName = string.Empty;
//            string sQry = string.Empty;

//            string BPLID = string.Empty;
//            string ItmBsort = string.Empty;
//            string DocDate = string.Empty;

//            Form.PSH_FormHelpClass FormHelpClass = new Form.PSH_FormHelpClass();

//            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            try
//            {
//                /// ODBC 연결 체크
//                if (ConnectODBC() == false)
//                {
//                    throw new Exception();
//                }

//                /// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//                WinTitle = "[S142] 발주서";
//                ReportName = "S142_1.rpt";

//                sQry = "EXEC PH_PY113_1 '" + oForm.Items.Item("8").Specific.Value + "'";
//                FormHelpClass.gRpt_Formula = new string[2];
//                FormHelpClass.gRpt_Formula_Value = new string[2];
//                PSH_Globals.gRpt_SRptSqry = new string[2];
//                PSH_Globals.gRpt_SRptName = new string[2];
//                PSH_Globals.gRpt_SFormula = new string[2, 2];
//                PSH_Globals.gRpt_SFormula_Value = new string[2, 2];
//                FormHelpClass.CrystalReportOpen();
//                /// Formula 수식필드
//                /// SubReport

//                /// Procedure 실행"
//                sQry = "EXEC [PS_PP820_01] '" + BPLID + "', '" + ItmBsort + "', '" + DocDate + "'";

//                oRecordSet.DoQuery(sQry);
//                if (oRecordSet.RecordCount == 0)
//                {
//                    if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false)
//                    {
//                        PSH_Globals.SBO_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    }
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("조회된 데이터가 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY113_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//            }
//        }


//        private bool DI_oJournalEntries()
//        {
//            bool functionReturnValue = false;
//            functionReturnValue = true;

//            string sQry = string.Empty;
//            int RetVal = 0;
//            int nErr = 0;
//            string ErrMsg = string.Empty;
//            int i = 0;
//            int k = 0;
//            string AcctCode = string.Empty;
//            string shortCode = string.Empty;
//            double Credit = 0;
//            double Debit = 0;
//            string LineMemo = string.Empty;
//            string Project = string.Empty;
//            string Dimenz1 = string.Empty;
//            string Dimenz2 = string.Empty;
//            string Dimenz3 = string.Empty;

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//            ///// 분개장문서
//            //    Set f_oJournalEntries = oCompany.GetBusinessObject(oJournalEntries)
//            SAPbobsCOM.JournalVouchers f_oJournalEntries = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);

//            try
//            {
//                if (string.IsNullOrEmpty(Strings.Trim(oDocDate)))
//                {
//                    oDocDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyy-mm-dd");
//                }
//                else
//                {
//                    oDocDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDocDate, "0000-00-00");
//                }

//                /// 재무관리>분개 =계정정보
//                PSH_Globals.oCompany.StartTransaction();
//                var _with1 = f_oJournalEntries;
//                //// 전표전체적용
//                //        .Series = "14"                             '/ 시리즈:주요 분개개체번호
//                //        .Series = MDC_Setmod.Get_Series_No("30")    '시리즈:주요 분개개체번호
//                _with1.JournalEntries.DueDate = Convert.ToDateTime(oDocDate);
//                //"04/26/2007"   '/ 만기일
//                _with1.JournalEntries.TaxDate = Convert.ToDateTime(oDocDate);
//                //"04/26/2007"   '/ 과세일
//                _with1.JournalEntries.ReferenceDate = Convert.ToDateTime(oDocDate);
//                //"04/26/2007"   '/ 전기일
//                _with1.JournalEntries.Memo = Strings.Trim(oREMARK);
//                //        .ProjectCode = "103"                       '/ 프로젝트코드
//                //        .AutoVAT = tYES                            '/ 부가세코드사용
//                i = 1;
//                oMat1.FlushToDataSource();
//                for (k = 0; k <= oMat1.VisualRowCount - 1; k++)
//                {
//                    oDS_PH_PY113B.Offset = k;
//                    if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY113B.GetValue("U_AcctCode", k))))
//                    {
//                        if (i != 1)
//                        {
//                            _with1.JournalEntries.Lines.Add();
//                            _with1.JournalEntries.Lines.SetCurrentLine((k));
//                        }
//                        AcctCode = oDS_PH_PY113B.GetValue("U_AcctCode", k);
//                        shortCode = Strings.Trim(oDS_PH_PY113B.GetValue("U_ShortNam", k));
//                        Debit = Conversion.Val(Strings.Replace(oDS_PH_PY113B.GetValue("U_Debit", k), ",", ""));
//                        Credit = Conversion.Val(Strings.Replace(oDS_PH_PY113B.GetValue("U_Credit", k), ",", ""));
//                        LineMemo = Strings.Trim(oDS_PH_PY113B.GetValue("U_Comments", k));
//                        Project = Strings.Trim(oDS_PH_PY113B.GetValue("U_Prject", k));
//                        //                Dimenz1 = Trim$(oDS_PH_PY113B.GetValue("U_Col07", k))
//                        //                Dimenz2 = Trim$(oDS_PH_PY113B.GetValue("U_Col08", k))
//                        //                If oDIM3 = "Y" Then Dimenz3 = Trim$(oDS_PH_PY113B.GetValue("U_Col10", k))

//                        //  .Lines.SetCurrentLine (0)
//                        //UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                        _with1.JournalEntries.Lines.AccountCode = DataHelpClass.Get_ReData("AcctCode", "FormatCode", "[OACT]", "'" + Strings.Trim(AcctCode) + "'", "");
//                        if (DataHelpClass.Value_ChkYn("[OACT]", "FormatCode", "'" + Strings.Trim(shortCode) + "'", "") == false)
//                        {
//                            //UPGRADE_WARNING: MDC_SetMod.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                            _with1.JournalEntries.Lines.ShortName = DataHelpClass.Get_ReData("AcctCode", "FormatCode", "[OACT]", "'" + Strings.Trim(shortCode) + "'", "");
//                        }
//                        else
//                        {
//                            _with1.JournalEntries.Lines.ShortName = Strings.Trim(shortCode);
//                        }
//                        _with1.JournalEntries.Lines.Credit = Credit;
//                        _with1.JournalEntries.Lines.Debit = Debit;
//                        _with1.JournalEntries.Lines.ProjectCode = Strings.Trim(Project);
//                        //                .JournalEntries.Lines.CostingCode = Trim$(Dimenz1)  '차원1
//                        //                .JournalEntries.Lines.CostingCode2 = Trim$(Dimenz2) '차원2
//                        _with1.JournalEntries.Lines.LineMemo = Strings.Trim(LineMemo);
//                        //                If oDIM3 = "Y" Then
//                        //                    .JournalEntries.Lines.CostingCode3 = Trim$(Dimenz3)  '차원3
//                        //                End If
//                        i = i + 1;
//                    }
//                }

//                RetVal = _with1.Add();
//                //Check Error
//                if ((0 != RetVal))
//                {
//                    PSH_Globals.oCompany.GetLastError(out nErr, out ErrMsg);
//                    PSH_Globals.SBO_Application.SetStatusBarMessage(ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                }
//                else
//                {
//                    PSH_Globals.oCompany.GetNewObjectCode(out oDocNum);

//                    sQry = "EXEC PH_PY113_INSERT '" + Strings.Trim(oDocNum) + "', 'PH_PY113', '" + Strings.Trim(oDocDate) + "', " + oTOTDEB + ", " + oTOTCRE + ", '" + Strings.Trim(oYM) + "', '" + Strings.Trim(oJOBGBN) + "', '" + Strings.Trim(oJOBTYP) + "', '" + Strings.Trim(oPAYSEL) + "', '" + Strings.Trim(oCLTCOD) + "'";

//                    oRecordSet.DoQuery(sQry);

//                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
//                    //UPGRADE_WARNING: oForm.Items(DOCNUM).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oForm.Items.Item("DOCNUM").Specific.Value = Strings.Left(oDocNum, Strings.Len(Strings.Trim(oDocNum)) - 1) + "-" + Strings.Right(oDocNum, 1);
//                    oForm.Items.Item("DOCNUM").Update();

//                    /// MsgBox ("완료!")
//                    PSH_Globals.SBO_Application.StatusBar.SetText("분개장 문서가 생성되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

//                    functionReturnValue = true;
//                }
//            }
//            catch (Exception ex)
//            {
//                if (PSH_Globals.oCompany.InTransaction)
//                {
//                    PSH_Globals.oCompany.EndTransaction((SAPbobsCOM.BoWfTransOpt.wf_RollBack));
//                }
//                PSH_Globals.SBO_Application.StatusBar.SetText("oJournalEntries : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

//                functionReturnValue = false;
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(f_oJournalEntries);
//            }

//            return functionReturnValue;
//        }

//        private bool Execution_Process()
//        {
//            bool functionReturnValue = false;

//            SAPbobsCOM.Recordset oRecordSet = null;
//            string sQry = null;
//            short ErrNum = 0;
//            int i = 0;
//            /// Check
//            ErrNum = 0;

//            try
//            {
//                oYM = oForm.Items.Item("YM").Specific.Value;
//                if (string.IsNullOrEmpty(Strings.Trim(oYM)))
//                {
//                    ErrNum = 1;
//                    throw new Exception();
//                }
//                oCLTCOD = oDS_PH_PY113A.GetValue("U_CLTCOD", 0);
//                oJOBTYP = oDS_PH_PY113A.GetValue("U_JOBTYP", 0);
//                oJOBGBN = oDS_PH_PY113A.GetValue("U_JOBGBN", 0);
//                oPAYSEL = oDS_PH_PY113A.GetValue("U_PAYSEL", 0);

//                /// 초기화
//                oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//                oDS_PH_PY113B.Clear();
//                oMat1.LoadFromDataSource();
//                i = 0;
//                sQry = " EXEC PH_PY113 '" + Strings.Trim(oCLTCOD) + "','" + Strings.Trim(oYM) + "', '" + Strings.Trim(oJOBTYP) + "', '";
//                sQry = sQry + Strings.Trim(oJOBGBN) + "', '" + Strings.Trim(oPAYSEL) + "'";

//                oRecordSet.DoQuery(sQry);
//                if (oRecordSet.RecordCount == 0)
//                {
//                    //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oForm.Items.Item("DOCNUM").Specific.Value = "";
//                    ErrNum = 2;
//                    throw new Exception();

//                }
//                oTOTDEB = 0;
//                oTOTCRE = 0;
//                while (!(oRecordSet.EoF))
//                {
//                    oDS_PH_PY113B.InsertRecord((i));
//                    oDS_PH_PY113B.Offset = i;
//                    oDS_PH_PY113B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//                    oDS_PH_PY113B.SetValue("U_AcctCode", i, oRecordSet.Fields.Item("AcctCode").Value);
//                    oDS_PH_PY113B.SetValue("U_AcctName", i, oRecordSet.Fields.Item("AcctName").Value);
//                    oDS_PH_PY113B.SetValue("U_ShortNam", i, oRecordSet.Fields.Item("ShortNam").Value);
//                    oDS_PH_PY113B.SetValue("U_Debit", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("Debit").Value, "#,###,###,##0"));
//                    /// 총급여액
//                    oDS_PH_PY113B.SetValue("U_Credit", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("Credit").Value, "#,###,###,##0"));
//                    oDS_PH_PY113B.SetValue("U_Project", i, oRecordSet.Fields.Item("U_Project").Value);
//                    oDS_PH_PY113B.SetValue("U_CostCent", i, oRecordSet.Fields.Item("U_PNLCOD").Value);
//                    oDS_PH_PY113B.SetValue("U_Comments", i, oRecordSet.Fields.Item("U_Remark").Value);

//                    oTOTDEB = oTOTDEB + oRecordSet.Fields.Item("Debit").Value;
//                    oTOTCRE = oTOTCRE + oRecordSet.Fields.Item("Credit").Value;
//                    i = i + 1;
//                    oRecordSet.MoveNext();
//                }
//                oMat1.LoadFromDataSource();

//                PH_PY113_AddMatrixRow();

//                /// 분개No 조회
//                sQry = "EXEC PH_PY113_QUERY 'PH_PY113', '" + Strings.Trim(oCLTCOD) + "','" + Strings.Trim(oYM) + "', '";
//                sQry = sQry + Strings.Trim(oJOBGBN) + "', '" + Strings.Trim(oJOBTYP) + "', '" + Strings.Trim(oPAYSEL) + "'";

//                oRecordSet.DoQuery(sQry);
//                if (oRecordSet.RecordCount > 0)
//                {
//                    //UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oDocNum = oRecordSet.Fields.Item("DOCNUM").Value;
//                }
//                else
//                {
//                    oDocNum = "";
//                }

//                oForm.Items.Item("DOCNUM").Specific.Value = oDocNum;
//                oForm.Items.Item("TOTPAY").Specific.Value = oTOTDEB;
//                oForm.Items.Item("TOTGON").Specific.Value = oTOTCRE;
//                oForm.Items.Item("TOTDEB").Specific.Value = oTOTDEB;
//                oForm.Items.Item("TOTCRE").Specific.Value = oTOTCRE;

//                /// End
//                PSH_Globals.SBO_Application.StatusBar.SetText("작업을 완료하였습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
//                functionReturnValue = true;
//            }
//            catch (Exception ex)
//            {
//                functionReturnValue = false;
//                if (ErrNum == 1)
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("조회월을 입력하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//                else if (ErrNum == 2)
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("조건과 일치하는 자료가 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//                else if (ErrNum == 3)
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("자사코드가 없습니다. 선택하여 주십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//                else
//                {
//                    PSH_Globals.SBO_Application.StatusBar.SetText("Execution_Process 실행 중 오류가 발생했습니다." + Strings.Space(10) + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//                }
//            }
//            finally
//            {
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//            }

//            return functionReturnValue;
//        }

//        private void TOTAL_AMT()
//        {
//            short oRow = 0;
//            oTOTDEB = 0;
//            oTOTCRE = 0;


//            for (oRow = 1; oRow <= oMat1.VisualRowCount; oRow++)
//            {
//                oDS_PH_PY113B.Offset = oRow - 1;
//                oTOTDEB = oTOTDEB + Conversion.Val(Strings.Replace(oDS_PH_PY113B.GetValue("U_Debit", oRow - 1), ",", ""));
//                oTOTCRE = oTOTCRE + Conversion.Val(Strings.Replace(oDS_PH_PY113B.GetValue("U_Credit", oRow - 1), ",", ""));
//            }
//            ///
//            //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("TOTDEB").Specific.Value = oTOTDEB;
//            //UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            oForm.Items.Item("TOTCRE").Specific.Value = oTOTCRE;
//        }
//    }
//}