using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using Microsoft.VisualBasic;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 개인별 대부금 잔액현황
    /// </summary>
    internal class PH_PY315 : PSH_BaseClass
    {
        public string oFormUniqueID;
        public SAPbouiCOM.Matrix oMat01;

        //public SAPbouiCOM.Form oForm;

        //private SAPbouiCOM.DBDataSource oDS_PH_PY315A; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PH_PY315B; //등록라인

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        public override void LoadForm(string oFromDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY315.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY315_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY315");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy="DocEntry" '//UDO방식일때

                oForm.Freeze(true);
                PH_PY315_CreateItems();
                PH_PY315_EnableMenus();
                PH_PY315_SetDocument(oFromDocEntry01);

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
        private void PH_PY315_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                oDS_PH_PY315B = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

                // 메트릭스 개체 할당
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 상환일자
                oForm.DataSources.UserDataSources.Add("RpmtDate", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("RpmtDate").Specific.DataBind.SetBound(true, "", "RpmtDate");
                oForm.Items.Item("RpmtDate").Specific.VALUE = DateTime.Now.ToString("yyyyMMdd");

                // 사번
                oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

                // 성명
                oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY315_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
        private void PH_PY315_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", false); // 삭제
                oForm.EnableMenu("1286", false); // 닫기
                oForm.EnableMenu("1287", false); // 복제
                oForm.EnableMenu("1285", false); // 복원
                oForm.EnableMenu("1284", false); // 취소
                oForm.EnableMenu("1293", false); // 행삭제
                oForm.EnableMenu("1281", false);
                oForm.EnableMenu("1282", true);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY315_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFromDocEntry01"></param>
        private void PH_PY315_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PH_PY315_FormItemEnabled();
                    //Call PH_PY315_AddMatrixRow(0, True) '//UDO방식일때
                }
                else
                {
                    //oForm.Mode = fm_FIND_MODE
                    //PH_PY315_FormItemEnabled
                    //oForm.Items("DocEntry").Specific.Value = oFromDocEntry01
                    //oForm.Items("1").Click ct_Regular
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY315_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY315_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY315_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        private void PH_PY315_MTX01()
        {
            int i = 0;
            string sQry = null;
            short ErrNum = 0;

            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string CLTCOD = string.Empty; //사업장
            string RpmtDate = string.Empty;
            string CntcCode = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

            try
            {
                oForm.Freeze(true);

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();   // 사업장
                RpmtDate = oForm.Items.Item("RpmtDate").Specific.VALUE.Trim();
                CntcCode = oForm.Items.Item("CntcCode").Specific.VALUE.Trim();

                sQry = "            EXEC [PH_PY315_01] ";
                sQry = sQry + "'" + CLTCOD + "',"; //사업장
                sQry = sQry + "'" + RpmtDate + "',";
                sQry = sQry + "'" + CntcCode + "'";

                oRecordSet.DoQuery(sQry);

                oMat01.Clear();
                oDS_PH_PY315B.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (oRecordSet.RecordCount == 0)
                {
                    ErrNum = 1;
                    oMat01.Clear();
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    throw new Exception();
                }

                for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
                {
                    if (i + 1 > oDS_PH_PY315B.Size)
                    {
                        oDS_PH_PY315B.InsertRecord(i);
                    }

                    oMat01.AddRow();
                    oDS_PH_PY315B.Offset = i;

                    oDS_PH_PY315B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PH_PY315B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("TeamName").Value);   // 부서
                    oDS_PH_PY315B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("JIGName").Value);    // 직급
                    oDS_PH_PY315B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CntcCode").Value);   // 사번
                    oDS_PH_PY315B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("CntcName").Value);   // 성명
                    oDS_PH_PY315B.SetValue("U_ColDt01",  i, Convert.ToDateTime(oRecordSet.Fields.Item("LoanDate").Value.ToString().Trim()).ToString("yyyyMMdd"));   // 대출일자
                    oDS_PH_PY315B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("NewLoanAmt").Value); // 신규대출액
                    oDS_PH_PY315B.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("IntRate").Value);    // 금리
                    oDS_PH_PY315B.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("TotLoanAmt").Value); // 대출총액
                    oDS_PH_PY315B.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("PrRmainAmt").Value); // 전월대출잔액
                    oDS_PH_PY315B.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("RpmtAmt").Value);    // 당월상환원금
                    oDS_PH_PY315B.SetValue("U_ColSum05", i, oRecordSet.Fields.Item("RpmtInt").Value);    // 당월상환이자
                    oDS_PH_PY315B.SetValue("U_ColSum06", i, oRecordSet.Fields.Item("RmainAmt").Value);   // 미상환잔액
                    oDS_PH_PY315B.SetValue("U_ColSum07", i, oRecordSet.Fields.Item("TotRpmtAmt").Value); // 당월상환총액


                    oRecordSet.MoveNext();
                    ProgBar01.Value = ProgBar01.Value + 1;
                    ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";

                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                ProgBar01.Stop();

            }
            catch (Exception ex)
            {
                ProgBar01.Stop();

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY315_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oForm.Freeze(false);
                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PH_PY315_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {
                    case "CntcCode":
                        oForm.Items.Item("CntcName").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("CntcCode").Specific.VALUE + "'", "");
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY315_FlushToItemValue_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        private void PH_PY315_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty; //사업장
            System.DateTime RpmtDate = default(System.DateTime);
            string CntcCode = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();   // 사업장
            RpmtDate = DateTime.ParseExact(oForm.Items.Item("RpmtDate").Specific.Value, "yyyyMMdd", null);
            CntcCode = oForm.Items.Item("CntcCode").Specific.VALUE.Trim();

            try
            {
                WinTitle = "[PH_PY315] 개인별 대부금 잔액현황";
                ReportName = "PH_PY315_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();//Parameter List
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                // Formula
               // dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'")));

                // Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@RpmtDate", RpmtDate));
                dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                                                             //    Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "PH_PY315")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "BtnSearch") //조회 버튼
                    {
                        if (PH_PY315_DataValidCheck() == true)
                        {
                            PH_PY315_MTX01();
                        }
                    }
                    else if (pVal.ItemUID == "BtnPrint1")
                    {
                        if (PH_PY315_DataValidCheck() == true)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PH_PY315_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "PH_PY315")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
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
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "RpmtDate", ""); //상환일자
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PH_PY315_DataValidCheck()
        {
            bool functionReturnValue = false;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }

                // 상환일자
                if (string.IsNullOrEmpty(oForm.Items.Item("RpmtDate").Specific.VALUE.Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("상환일자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("RpmtDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    functionReturnValue = false;
                    return functionReturnValue;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY315_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                functionReturnValue = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
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
                    PH_PY315_FlushToItemValue(pVal.ItemUID, 0, "");
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
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
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                        }
                        else
                        {
                            PH_PY315_FlushToItemValue(pVal.ItemUID, 0, "");

                            //if (pVal.ItemUID == "MSTCOD")
                            //{
                            //    oForm.Items.Item("MSTNAM").Specific.VALUE = dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item("MSTCOD").Specific.VALUE + "'", ""); //성명
                            //}
                            //else if (pVal.ItemUID == "ShiftDatCd")
                            //{
                            //    oForm.Items.Item("ShiftDatNm").Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" + oForm.Items.Item("ShiftDatCd").Specific.VALUE + "'", " AND T0.Code = 'P154' AND T0.U_UseYN = 'Y'"); //근무형태

                            //}
                            //else if (pVal.ItemUID == "GNMUJOCd")
                            //{
                            //    oForm.Items.Item("GNMUJONm").Specific.VALUE = dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L] AS T0", "'" + oForm.Items.Item("GNMUJOCd").Specific.VALUE + "'", " AND T0.Code = 'P155' AND T0.U_UseYN = 'Y'"); //근무조
                            //}
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
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
                    PH_PY315_FormItemEnabled();
                    //PH_PY315_AddMatrixRow(oMat01.VisualRowCount);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
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
                    //If (pval.ItemUID = "ItemCode") Then
                    //  Dim oDataTable01 As SAPbouiCOM.DataTable
                    //  Set oDataTable01 = pval.SelectedObjects
                    //  oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
                    //  Set oDataTable01 = Nothing
                    //End If
                    //If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
                    //  Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY315A", "U_CardCode,U_CardName")
                    //End If
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
                        case "7169": //엑셀 내보내기
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
                            //Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            break;
                        case "7169": //엑셀 내보내기
                            //엑셀 내보내기 이후 처리_S
                            oForm.Freeze(true);
                            oDS_PH_PY315B.RemoveRecord(oDS_PH_PY315B.Size - 1);
                            oMat01.LoadFromDataSource();
                            oForm.Freeze(false);
                            //엑셀 내보내기 이후 처리_E
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

        #region Raise_FormMenuEvent (최종테스트 후 주석 삭제 필요, 2019.05.17 송명규)
        //		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			string sQry = null;
        //			SAPbobsCOM.Recordset RecordSet01 = null;
        //			RecordSet01 = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //			////BeforeAction = True
        //			if ((pval.BeforeAction == true)) {

        //			////BeforeAction = False
        //			} else if ((pval.BeforeAction == false)) {

        //			}
        //			return;
        //			Raise_FormMenuEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormMenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormDataEvent (최종테스트 후 주석 삭제 필요, 2019.05.17 송명규)
        //		public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			////BeforeAction = True
        //			if ((BusinessObjectInfo.BeforeAction == true)) {

        //			////BeforeAction = False
        //			} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //			}
        //			return;
        //			Raise_FormDataEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_RightClickEvent (최종테스트 후 주석 삭제 필요, 2019.05.17 송명규)
        //		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {
        //			}

        //			return;
        //			Raise_RightClickEvent_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion

        #region Raise_FormItemEvent (최종테스트 후 주석 삭제 필요, 2019.05.17 송명규)
        //		private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_ITEM_PRESSED_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_KEY_DOWN_Error:

        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {


        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_CLICK_Error:

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			return;
        //			Raise_EVENT_COMBO_SELECT_Error:
        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_DOUBLE_CLICK_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			oForm.Freeze(true);

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}

        //			oForm.Freeze(false);

        //			return;
        //			Raise_EVENT_VALIDATE_Error:

        //			oForm.Freeze(false);
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_MATRIX_LOAD_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_RESIZE_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {

        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}


        //		private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement


        //			return;
        //			Raise_EVENT_GOT_FOCUS_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}

        //		private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //		{
        //			 // ERROR: Not supported in C#: OnErrorStatement

        //			if (pval.BeforeAction == true) {
        //			} else if (pval.BeforeAction == false) {

        //			}
        //			return;
        //			Raise_EVENT_FORM_UNLOAD_Error:
        //			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        #endregion
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
//	internal class PH_PY315
//	{
//////********************************************************************************
//////  File           : PH_PY315.cls
//////  Module         : 인사관리 > 기타관리 > 대부금관리
//////  Desc           : 개인별 대부금 잔액현황
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

//		public SAPbouiCOM.Matrix oMat1;
//		public SAPbouiCOM.Matrix oMat2;

//		private SAPbouiCOM.DBDataSource oDS_PH_PY315A;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY315B;
//		private SAPbouiCOM.DBDataSource oDS_PH_PY315C;

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY315.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY315_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY315");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "DocEntry"

//			oForm.Freeze(true);
//			PH_PY315_CreateItems();
//			PH_PY315_EnableMenus();
//			PH_PY315_SetDocument(oFromDocEntry01);
//			PH_PY315_FormResize();

//			oForm.Update();
//			oForm.Freeze(false);

//			//    oForm.Items("Folder01").Specific.Select '폼이 로드 될 때 Folder01이 선택됨

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

//		private bool PH_PY315_CreateItems()
//		{
//			bool functionReturnValue = false;

//			string sQry = null;
//			int i = 0;

//			SAPbouiCOM.EditText oEdit = null;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbouiCOM.Column oColumn = null;
//			SAPbouiCOM.Columns oColumns = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//    Set oDS_PH_PY315A = oForm.DataSources.DBDataSources("@PH_PY315A")
//			oDS_PH_PY315B = oForm.DataSources.DBDataSources("@PS_USERDS01");
//			//    Set oDS_PH_PY315C = oForm.DataSources.DBDataSources("@PS_USERDS02")

//			oMat1 = oForm.Items.Item("Mat01").Specific;
//			//    Set oMat2 = oForm.Items("Mat02").Specific

//			oMat1.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
//			oMat1.AutoResizeColumns();

//			//    oMat2.SelectionMode = ms_Auto
//			//    oMat2.AutoResizeColumns

//			//----------조회 조건----------
//			//사업장_S
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
//			//사업장_E

//			//상환일자_S
//			oForm.DataSources.UserDataSources.Add("RpmtDate", SAPbouiCOM.BoDataType.dt_DATE);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("RpmtDate").Specific.DataBind.SetBound(true, "", "RpmtDate");
//			//상환일자_E

//			//사번_S
//			oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");
//			//사번_E

//			//성명_S
//			oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");
//			//성명_E
//			//----------조회 조건----------

//			////----------------------------------------------------------------------------------------------
//			//// 기본사항
//			////----------------------------------------------------------------------------------------------

//			//사업장
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			return functionReturnValue;
//			PH_PY315_CreateItems_Error:

//			//UPGRADE_NOTE: oEdit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oEdit = null;
//			//UPGRADE_NOTE: oCombo 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oCombo = null;
//			//UPGRADE_NOTE: oColumn 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumn = null;
//			//UPGRADE_NOTE: oColumns 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oColumns = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY315_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			//    Call oForm.EnableMenu("1283", False)         '// 삭제
//			//    Call oForm.EnableMenu("1287", True)          '// 복제
//			//'    Call oForm.EnableMenu("1286", True)         '// 닫기
//			//    Call oForm.EnableMenu("1284", True)         '// 취소
//			//    Call oForm.EnableMenu("1293", True)         '// 행삭제

//			return;
//			PH_PY315_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY315_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY315_FormItemEnabled();
//				//        Call PH_PY315_AddMatrixRow1
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY315_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY315_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY315_FormItemEnabled()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			SAPbouiCOM.ComboBox oCombo = null;
//			string CLTCOD = null;

//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				//폼 DocEntry 세팅
//				//        Call PH_PY315_FormClear

//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				//일자 세팅
//				//UPGRADE_WARNING: oForm.Items(RpmtDate).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("RpmtDate").Specific.VALUE = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {
//				//// 접속자에 따른 권한별 사업장 콤보박스세팅
//				MDC_SetMod.CLTCOD_Select(ref oForm, ref "CLTCOD", ref false);

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY315_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			//    Dim loopCount As Integer 'For Loop 용 (VALIDATE Event에서 사용)
//			//    Dim GovID1 As String '주민등록번호 앞자리(VALIDATE Event에서 사용)
//			//    Dim GovID2 As String '주민등록번호 뒷자리(VALIDATE Event에서 사용)
//			//    Dim GovID As String '주민등록번호 전체(VALIDATE Event에서 사용)
//			//    Dim Sex As String '성별(VALIDATE Event에서 사용)
//			//    Dim PayCnt As Integer '지급횟수(VALIDATE Event에서 사용)
//			//    Dim FeeTot As Double '입학금계(VALIDATE Event에서 사용)
//			//    Dim TuiTot As Double '등록금계(VALIDATE Event에서 사용)
//			//    Dim Total As Double '총계(VALIDATE Event에서 사용)

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pval.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pval.BeforeAction == true) {
//						//                If pval.ItemUID = "1" Then
//						//                    If oForm.Mode = fm_ADD_MODE Then
//						//                        If PH_PY315_DataValidCheck = False Then
//						//                            BubbleEvent = False
//						//                        End If
//						//
//						//                        '//해야할일 작업
//						//                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                        If PH_PY315_DataValidCheck = False Then
//						//                            BubbleEvent = False
//						//                        End If
//						//                        '//해야할일 작업
//						//
//						//                    ElseIf oForm.Mode = fm_OK_MODE Then
//						//                    End If
//						//                End If
//						if (pval.ItemUID == "BtnSearch") {

//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY315_DataValidCheck() == false) {
//									BubbleEvent = false;
//									return;
//								}
//								//
//								//                        '//해야할일 작업
//								PH_PY315_MTX01();
//								//                        Call PH_PY315_MTX02

//							}

//						}

//						if (pval.ItemUID == "BtnPrint1") {

//							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//								if (PH_PY315_DataValidCheck() == false) {
//									BubbleEvent = false;
//									return;
//								}
//								//
//								//                        '//해야할일 작업
//								PH_PY315_Print_Report01();

//							}

//						}

//					} else if (pval.BeforeAction == false) {
//						//                If pval.ItemUID = "1" Then
//						//                    If oForm.Mode = fm_ADD_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY315_FormItemEnabled
//						//                            Call PH_PY315_AddMatrixRow1
//						//                        End If
//						//                    ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY315_FormItemEnabled
//						//                            Call PH_PY315_AddMatrixRow1
//						//                        End If
//						//                    ElseIf oForm.Mode = fm_OK_MODE Then
//						//                        If pval.ActionSuccess = True Then
//						//                            Call PH_PY315_FormItemEnabled
//						//                        End If
//						//                    End If
//						//                End If

//						//                '폴더를 사용할 때는 필수 소스_S
//						//                'Folder01이 선택되었을 때
//						//                If pval.ItemUID = "Folder01" Then
//						//
//						//                    oForm.PaneLevel = 1
//						//
//						//                End If
//						//
//						//                'Folder02가 선택되었을 때
//						//                If pval.ItemUID = "Folder02" Then
//						//
//						//                    oForm.PaneLevel = 2
//						//
//						//                End If
//						//                '폴더를 사용할 때는 필수 소스_E

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2

//					if (pval.BeforeAction == true) {

//						if (pval.ItemUID == "Mat01") {

//							if (pval.ColUID == "Name" & pval.CharPressed == Convert.ToDouble("9")) {

//								//UPGRADE_WARNING: oMat1.Columns.Item(Name).Cells(pval.Row).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								if (string.IsNullOrEmpty(oMat1.Columns.Item("Name").Cells.Item(pval.Row).Specific.VALUE)) {
//									MDC_Globals.Sbo_Application.ActivateMenuItem("7425");
//									BubbleEvent = false;
//								}

//							}

//						} else if (pval.ItemUID == "CntcCode" & pval.CharPressed == Convert.ToDouble("9")) {

//							//UPGRADE_WARNING: oForm.Items(CntcCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//							if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.VALUE)) {
//								MDC_Globals.Sbo_Application.ActivateMenuItem("7425");
//								BubbleEvent = false;
//							}

//						}

//					} else if (pval.Before_Action == false) {

//					}
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pval.ItemUID) {
//						case "Mat01":
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
//							//                    Call PH_PY315_AddMatrixRow1
//							oMat1.AutoResizeColumns();
//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true) {
//						//                Select Case pval.ItemUID
//						//                Case "Mat01"
//						//                    If pval.Row > 0 Then
//						//                        Call oMat1.SelectRow(pval.Row, True, False)
//						//                    End If
//						//                End Select

//						switch (pval.ItemUID) {
//							case "Mat01":
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

//						if (pval.ItemChanged == true) {

//						}

//					} else if (pval.BeforeAction == false) {

//						if (pval.ItemChanged == true) {

//							switch (pval.ItemUID) {

//								case "CntcCode":
//									//Call oDS_PH_PY315A.setValue("U_CntcName", 0, MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" & oForm.Items(pval.ItemUID).Specific.Value & "'"))

//									//UPGRADE_WARNING: oForm.Items(CntcName).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									oForm.Items.Item("CntcName").Specific.VALUE = MDC_GetData.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pval.ItemUID).Specific.VALUE + "'");
//									break;

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

//						PH_PY315_FormItemEnabled();
//						//                Call PH_PY315_AddMatrixRow1
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
//						//UPGRADE_NOTE: oDS_PH_PY315A 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY315A = null;
//						//UPGRADE_NOTE: oDS_PH_PY315B 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY315B = null;
//						//UPGRADE_NOTE: oDS_PH_PY315C 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY315C = null;

//						//UPGRADE_NOTE: oMat1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat1 = null;
//						//UPGRADE_NOTE: oMat2 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oMat2 = null;

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

//						PH_PY315_FormResize();
//						oMat1.AutoResizeColumns();

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
//						//                    Call MDC_CF_DBDatasourceReturn(pval, pval.FormUID, "@PH_PY315A", "Code")
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
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement


//			short loopCount = 0;
//			double FeeTot = 0;
//			double TuiTot = 0;
//			double Total = 0;

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
//						break;

//					case "7169":
//						//엑셀 내보내기

//						//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
//						PH_PY315_AddMatrixRow1();
//						break;

//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY315_FormItemEnabled();
//						PH_PY315_AddMatrixRow1();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY315_FormItemEnabled();
//						PH_PY315_AddMatrixRow1();
//						oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY315_FormItemEnabled();
//						PH_PY315_AddMatrixRow1();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY315_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제

//						if (oMat1.RowCount != oMat1.VisualRowCount) {
//							oMat1.FlushToDataSource();

//							while ((i <= oDS_PH_PY315B.Size - 1)) {
//								if (string.IsNullOrEmpty(oDS_PH_PY315B.GetValue("U_LineNum", i))) {
//									oDS_PH_PY315B.RemoveRecord((i));
//									i = 0;
//								} else {
//									i = i + 1;
//								}
//							}

//							for (i = 0; i <= oDS_PH_PY315B.Size; i++) {
//								oDS_PH_PY315B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//							}

//							oMat1.LoadFromDataSource();
//						}
//						PH_PY315_AddMatrixRow1();
//						break;

//					case "7169":
//						//엑셀 내보내기

//						//엑셀 내보내기 이후 처리
//						oForm.Freeze(true);
//						oDS_PH_PY315B.RemoveRecord(oDS_PH_PY315B.Size - 1);
//						oMat1.LoadFromDataSource();
//						oForm.Freeze(false);
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
//				case "Mat01":
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

//		public void PH_PY315_AddMatrixRow1()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1]
//			oMat1.FlushToDataSource();
//			oRow = oMat1.VisualRowCount;

//			if (oMat1.VisualRowCount > 0) {
//				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PH_PY315B.GetValue("U_LineNum", oRow - 1)))) {
//					if (oDS_PH_PY315B.Size <= oMat1.VisualRowCount) {
//						oDS_PH_PY315B.InsertRecord((oRow));
//					}
//					oDS_PH_PY315B.Offset = oRow;
//					oDS_PH_PY315B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//					oDS_PH_PY315B.SetValue("U_ColReg01", oRow, "");
//					//부서
//					oDS_PH_PY315B.SetValue("U_ColReg02", oRow, "");
//					//직급
//					oDS_PH_PY315B.SetValue("U_ColReg03", oRow, "");
//					//사번
//					oDS_PH_PY315B.SetValue("U_ColReg04", oRow, "");
//					//성명
//					oDS_PH_PY315B.SetValue("U_ColDt01", oRow, "");
//					//대출일자
//					oDS_PH_PY315B.SetValue("U_ColSum01", oRow, Convert.ToString(0));
//					//신규대출액
//					oDS_PH_PY315B.SetValue("U_ColQty01", oRow, Convert.ToString(0));
//					//금리
//					oDS_PH_PY315B.SetValue("U_ColSum02", oRow, Convert.ToString(0));
//					//대출총액
//					oDS_PH_PY315B.SetValue("U_ColSum03", oRow, Convert.ToString(0));
//					//전월대출잔액
//					oDS_PH_PY315B.SetValue("U_ColSum04", oRow, Convert.ToString(0));
//					//당월상환원금
//					oDS_PH_PY315B.SetValue("U_ColSum05", oRow, Convert.ToString(0));
//					//당월상환이자
//					oDS_PH_PY315B.SetValue("U_ColSum06", oRow, Convert.ToString(0));
//					//미상환잔액
//					oDS_PH_PY315B.SetValue("U_ColSum07", oRow, Convert.ToString(0));
//					//당월상환총액

//					oMat1.LoadFromDataSource();
//				} else {
//					oDS_PH_PY315B.Offset = oRow - 1;
//					oDS_PH_PY315B.SetValue("U_LineNum", oRow - 1, Convert.ToString(oRow));
//					oDS_PH_PY315B.SetValue("U_ColReg01", oRow - 1, "");
//					//부서
//					oDS_PH_PY315B.SetValue("U_ColReg02", oRow - 1, "");
//					//직급
//					oDS_PH_PY315B.SetValue("U_ColReg03", oRow - 1, "");
//					//사번
//					oDS_PH_PY315B.SetValue("U_ColReg04", oRow - 1, "");
//					//성명
//					oDS_PH_PY315B.SetValue("U_ColDt01", oRow - 1, "");
//					//대출일자
//					oDS_PH_PY315B.SetValue("U_ColSum01", oRow - 1, Convert.ToString(0));
//					//신규대출액
//					oDS_PH_PY315B.SetValue("U_ColQty01", oRow - 1, Convert.ToString(0));
//					//금리
//					oDS_PH_PY315B.SetValue("U_ColSum02", oRow - 1, Convert.ToString(0));
//					//대출총액
//					oDS_PH_PY315B.SetValue("U_ColSum03", oRow - 1, Convert.ToString(0));
//					//전월대출잔액
//					oDS_PH_PY315B.SetValue("U_ColSum04", oRow - 1, Convert.ToString(0));
//					//당월상환원금
//					oDS_PH_PY315B.SetValue("U_ColSum05", oRow - 1, Convert.ToString(0));
//					//당월상환이자
//					oDS_PH_PY315B.SetValue("U_ColSum06", oRow - 1, Convert.ToString(0));
//					//미상환잔액
//					oDS_PH_PY315B.SetValue("U_ColSum07", oRow - 1, Convert.ToString(0));
//					//당월상환총액

//					oMat1.LoadFromDataSource();
//				}
//			} else if (oMat1.VisualRowCount == 0) {
//				oDS_PH_PY315B.Offset = oRow;
//				oDS_PH_PY315B.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//				oDS_PH_PY315B.SetValue("U_ColReg01", oRow, "");
//				//부서
//				oDS_PH_PY315B.SetValue("U_ColReg02", oRow, "");
//				//직급
//				oDS_PH_PY315B.SetValue("U_ColReg03", oRow, "");
//				//사번
//				oDS_PH_PY315B.SetValue("U_ColReg04", oRow, "");
//				//성명
//				oDS_PH_PY315B.SetValue("U_ColDt01", oRow, "");
//				//대출일자
//				oDS_PH_PY315B.SetValue("U_ColSum01", oRow, Convert.ToString(0));
//				//신규대출액
//				oDS_PH_PY315B.SetValue("U_ColQty01", oRow, Convert.ToString(0));
//				//금리
//				oDS_PH_PY315B.SetValue("U_ColSum02", oRow, Convert.ToString(0));
//				//대출총액
//				oDS_PH_PY315B.SetValue("U_ColSum03", oRow, Convert.ToString(0));
//				//전월대출잔액
//				oDS_PH_PY315B.SetValue("U_ColSum04", oRow, Convert.ToString(0));
//				//당월상환원금
//				oDS_PH_PY315B.SetValue("U_ColSum05", oRow, Convert.ToString(0));
//				//당월상환이자
//				oDS_PH_PY315B.SetValue("U_ColSum06", oRow, Convert.ToString(0));
//				//미상환잔액
//				oDS_PH_PY315B.SetValue("U_ColSum07", oRow, Convert.ToString(0));
//				//당월상환총액

//				oMat1.LoadFromDataSource();
//			}

//			oForm.Freeze(false);
//			return;
//			PH_PY315_AddMatrixRow1_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_AddMatrixRow1_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY315_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY315'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY315_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY315_FormResize()
//		{
//			//******************************************************************************
//			//Function ID : PH_PY315_FormResize()
//			//해당모듈 : PH_PY315
//			//기능 : Form의 크기 변경 시 아이템들의 위치 및 크기 동적 변경
//			//인수 : 없음
//			//반환값 : 없음
//			//특이사항 : 없음
//			//******************************************************************************
//			 // ERROR: Not supported in C#: OnErrorStatement


//			//    oForm.Items("Rect01").Height = oForm.Height - 180
//			//    oForm.Items("Rect01").Width = oForm.Width - 35

//			return;
//			PH_PY315_FormResize_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY315_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//사업장
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//상환일자
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (string.IsNullOrEmpty(Strings.Trim(oForm.Items.Item("RpmtDate").Specific.VALUE))) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("상환일자는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				oForm.Items.Item("RpmtDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//				functionReturnValue = false;
//				return functionReturnValue;
//			}

//			//회차
//			//    If Trim(oForm.Items("Count").Specific.VALUE) = "" Then
//			//        Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
//			//        oForm.Items("Count").CLICK ct_Regular
//			//        PH_PY315_DataValidCheck = False
//			//        Exit Function
//			//    End If


//			//    '라인
//			//    If oMat1.VisualRowCount > 1 Then
//			//        For i = 1 To oMat1.VisualRowCount - 1
//			//
//			//            '학교
//			//            If oMat1.Columns("SchCls").Cells(i).Specific.Value = "" Then
//			//                Sbo_Application.SetStatusBarMessage "학교는 필수입니다.", bmt_Short, True
//			//                oMat1.Columns("SchCls").Cells(i).CLICK ct_Regular
//			//                PH_PY315_DataValidCheck = False
//			//                Exit Function
//			//            End If
//			//
//			//            '학교명
//			//            If oMat1.Columns("SchName").Cells(i).Specific.Value = "" Then
//			//                Sbo_Application.SetStatusBarMessage "학교명은 필수입니다.", bmt_Short, True
//			//                oMat1.Columns("SchName").Cells(i).CLICK ct_Regular
//			//                PH_PY315_DataValidCheck = False
//			//                Exit Function
//			//            End If
//			//
//			//            '학년
//			//            If oMat1.Columns("Grade").Cells(i).Specific.Value = "" Then
//			//                Sbo_Application.SetStatusBarMessage "학년은 필수입니다.", bmt_Short, True
//			//                oMat1.Columns("Grade").Cells(i).CLICK ct_Regular
//			//                PH_PY315_DataValidCheck = False
//			//                Exit Function
//			//            End If
//			//
//			//            '회차
//			//            If oMat1.Columns("Count").Cells(i).Specific.Value = "" Then
//			//                Sbo_Application.SetStatusBarMessage "회차는 필수입니다.", bmt_Short, True
//			//                oMat1.Columns("Count").Cells(i).CLICK ct_Regular
//			//                PH_PY315_DataValidCheck = False
//			//                Exit Function
//			//            End If
//			//
//			//        Next
//			//    Else
//			//        Sbo_Application.SetStatusBarMessage "라인 데이터가 없습니다.", bmt_Short, True
//			//        PH_PY315_DataValidCheck = False
//			//        Exit Function
//			//    End If

//			oMat1.FlushToDataSource();
//			//// Matrix 마지막 행 삭제(DB 저장시)
//			if (oDS_PH_PY315B.Size > 1)
//				oDS_PH_PY315B.RemoveRecord((oDS_PH_PY315B.Size - 1));

//			oMat1.LoadFromDataSource();

//			functionReturnValue = true;
//			return functionReturnValue;


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY315_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY315_MTX01()
//		{

//			////메트릭스에 데이터 로드

//			int i = 0;
//			string sQry = null;

//			string CLTCOD = null;
//			string RpmtDate = null;
//			string CntcCode = null;

//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);
//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			RpmtDate = oForm.Items.Item("RpmtDate").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CntcCode = oForm.Items.Item("CntcCode").Specific.VALUE;

//			SAPbouiCOM.ProgressBar ProgressBar01 = null;
//			ProgressBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", oRecordSet.RecordCount, false);

//			sQry = "EXEC PH_PY315_01 '" + CLTCOD + "','" + RpmtDate + "','" + CntcCode + "'";
//			oRecordSet.DoQuery(sQry);

//			oMat1.Clear();
//			oMat1.FlushToDataSource();
//			oMat1.LoadFromDataSource();

//			if ((oRecordSet.RecordCount == 0)) {
//				oMat1.Clear();
//				goto PH_PY315_MTX01_Exit;
//			}

//			for (i = 0; i <= oRecordSet.RecordCount - 1; i++) {
//				if (i != 0) {
//					oDS_PH_PY315B.InsertRecord((i));
//				}
//				oDS_PH_PY315B.Offset = i;
//				oDS_PH_PY315B.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//				oDS_PH_PY315B.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("TeamName").Value);
//				//부서
//				oDS_PH_PY315B.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("JIGName").Value);
//				//직급
//				oDS_PH_PY315B.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CntcCode").Value);
//				//사번
//				oDS_PH_PY315B.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("CntcName").Value);
//				//성명
//				oDS_PH_PY315B.SetValue("U_ColDt01", i, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oRecordSet.Fields.Item("LoanDate").Value, "YYYYMMDD"));
//				//대출일자
//				oDS_PH_PY315B.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("NewLoanAmt").Value);
//				//신규대출액
//				oDS_PH_PY315B.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("IntRate").Value);
//				//금리
//				oDS_PH_PY315B.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("TotLoanAmt").Value);
//				//대출총액
//				oDS_PH_PY315B.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("PrRmainAmt").Value);
//				//전월대출잔액
//				oDS_PH_PY315B.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("RpmtAmt").Value);
//				//당월상환원금
//				oDS_PH_PY315B.SetValue("U_ColSum05", i, oRecordSet.Fields.Item("RpmtInt").Value);
//				//당월상환이자
//				oDS_PH_PY315B.SetValue("U_ColSum06", i, oRecordSet.Fields.Item("RmainAmt").Value);
//				//미상환잔액
//				oDS_PH_PY315B.SetValue("U_ColSum07", i, oRecordSet.Fields.Item("TotRpmtAmt").Value);
//				//당월상환총액

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
//			PH_PY315_MTX01_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			if ((ProgressBar01 != null)) {
//				ProgressBar01.Stop();
//			}
//			MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "W");
//			return;
//			PH_PY315_MTX01_Error:
//			ProgressBar01.Stop();
//			//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgressBar01 = null;
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY315_Validate(string ValidateType)
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
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY315A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY315A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY315_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY315_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY315_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY315_Print_Report01()
//		{

//			string DocNum = null;
//			short ErrNum = 0;
//			string WinTitle = null;
//			string ReportName = null;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			string CLTCOD = null;
//			string RpmtDate = null;
//			string CntcCode = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//			/// ODBC 연결 체크
//			if (ConnectODBC() == false) {
//				goto PH_PY315_Print_Report01_Error;
//			}


//			////인자 MOVE , Trim 시키기..
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			RpmtDate = oForm.Items.Item("RpmtDate").Specific.VALUE;
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CntcCode = oForm.Items.Item("CntcCode").Specific.VALUE;

//			/// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//			WinTitle = "[PH_PY315] 개인별 대부금 잔액현황";
//			ReportName = "PH_PY315_01.rpt";
//			MDC_Globals.gRpt_Formula = new string[1];
//			MDC_Globals.gRpt_Formula_Value = new string[1];
//			MDC_Globals.gRpt_SRptSqry = new string[1];
//			MDC_Globals.gRpt_SRptName = new string[1];

//			/// Formula 수식필드


//			/// SubReport

//			//    ReDim gRpt_SFormula(0, 0)
//			//    ReDim gRpt_SFormula_Value(0, 0)


//			/// Procedure 실행"
//			sQry = "EXEC PH_PY315_01 '" + CLTCOD + "','" + RpmtDate + "','" + CntcCode + "'";

//			oRecordSet.DoQuery(sQry);
//			if (oRecordSet.RecordCount == 0) {
//				ErrNum = 1;
//				goto PH_PY315_Print_Report01_Error;
//			}

//			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "N", sQry, "1", "Y", "V", , 1) == false) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			PH_PY315_Print_Report01_Error:

//			if (ErrNum == 1) {
//				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet = null;
//				MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
//			} else {
//				//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//				oRecordSet = null;
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY315_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//		}
//	}
//}
