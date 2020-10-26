using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 일출근현황
    /// </summary>
    internal class PH_PY560 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY560.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY560_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY560");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY560_CreateItems();
                PH_PY560_FormItemEnabled();
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
                oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY560_CreateItems()
        {
            string sQry = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // 기준일자
                oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
                oForm.DataSources.UserDataSources.Item("DocDate").Value = DateTime.Now.ToString("yyyyMMdd");

                // 부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                // G5연장제외 CHK_BOX 
                oForm.DataSources.UserDataSources.Add("G5_YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("G5_YN").Specific.DataBind.SetBound(true, "", "G5_YN");
                oForm.Items.Item("G5_YN").Specific.Checked = true;

                // 무재해일수 조회용 CHK_BOX
                oForm.DataSources.UserDataSources.Add("NoAccdnt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("NoAccdnt").Specific.DataBind.SetBound(true, "", "NoAccdnt");
                oForm.Items.Item("NoAccdnt").Specific.Checked = false;

                // 평일_석식
                oForm.DataSources.UserDataSources.Add("SukS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SukS").Specific.DataBind.SetBound(true, "", "SukS");

                // 평일_야식
                oForm.DataSources.UserDataSources.Add("YaS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("YaS").Specific.DataBind.SetBound(true, "", "YaS");

                // 평일_당직자
                oForm.DataSources.UserDataSources.Add("DangJik", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("DangJik").Specific.DataBind.SetBound(true, "", "DangJik");

                // 휴일_중식
                oForm.DataSources.UserDataSources.Add("JoongS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("JoongS").Specific.DataBind.SetBound(true, "", "JoongS");

                // 통근버스_주간
                oForm.DataSources.UserDataSources.Add("Bus1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Bus1").Specific.ValidValues.Add("", "");
                oForm.Items.Item("Bus1").Specific.ValidValues.Add("운행", "운행");
                oForm.Items.Item("Bus1").Specific.ValidValues.Add("미운행", "미운행");
                oForm.Items.Item("Bus1").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 통근버스_주잔
                oForm.DataSources.UserDataSources.Add("Bus2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Bus2").Specific.ValidValues.Add("", "");
                oForm.Items.Item("Bus2").Specific.ValidValues.Add("운행", "운행");
                oForm.Items.Item("Bus2").Specific.ValidValues.Add("운행(20:00)", "운행(20:00)");
                oForm.Items.Item("Bus2").Specific.ValidValues.Add("운행(20:30)", "운행(20:30)");
                oForm.Items.Item("Bus2").Specific.ValidValues.Add("미운행", "미운행");
                oForm.Items.Item("Bus2").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 통근버스_야간
                oForm.DataSources.UserDataSources.Add("Bus3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Bus3").Specific.ValidValues.Add("", "");
                oForm.Items.Item("Bus3").Specific.ValidValues.Add("운행(05:30)", "운행(05:30)");
                oForm.Items.Item("Bus3").Specific.ValidValues.Add("운행(08:00)", "운행(08:00)");
                oForm.Items.Item("Bus3").Specific.ValidValues.Add("운행(08:30)", "운행(08:30)");
                oForm.Items.Item("Bus3").Specific.ValidValues.Add("미운행", "미운행");
                oForm.Items.Item("Bus3").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY560_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        ///// <summary>
        ///// 화면의 아이템 Enable 설정
        ///// </summary>
        public void PH_PY560_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY560_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        // </summary>
        // <param name="FormUID">Form UID</param>
        // <param name="pVal">이벤트 </param>
        // <param name="BubbleEvent">Bubble Event</param>
        public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                    ////2
                    break;


                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                    ////3

                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                    ////4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                    ////7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                    ////8
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
                    ////9
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                                                             //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
                    ////12
                    break;


                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                    ////16
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                    ////18
                    break;


                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                    ////19
                    break;


                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                    ////20
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                                                             // Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
                    ////22
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
                    ////23
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                                                                  // Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
                    ////37
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
                    ////38
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag:
                    ////39
                    break;

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
                    if (pVal.ItemUID == "Btn01")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY560_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            int i = 0;
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
                            //사업장이 바뀌면 부서와 담당 재설정
                            case "CLTCOD":
                                ////부서
                                ////삭제
                                if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");

                                break;
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            //case "MSTCOD":
                            //    sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim() + "'";
                            //    oRecordSet.DoQuery(sQry);
                            //    oForm.Items.Item("MSTNAME").Specific.VALUE = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();
                            //    break;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY560_Print_Report01()
        {
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty;
            string DocDate = string.Empty;
            string TeamCode = string.Empty;
            string SukS = string.Empty;
            string YaS = string.Empty;
            string DangJ = string.Empty;
            string JoongS = string.Empty;
            string Bus1 = string.Empty;
            string Bus2 = string.Empty;
            string Bus3 = string.Empty;
            string G5_YN = string.Empty;
            string NoAccdnt = string.Empty;

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.VALUE.Trim();
            DocDate = oForm.Items.Item("DocDate").Specific.VALUE.Trim();
            TeamCode = oForm.Items.Item("TeamCode").Specific.VALUE.Trim();
            SukS = oForm.Items.Item("SukS").Specific.VALUE.Trim();
            YaS = oForm.Items.Item("YaS").Specific.VALUE.Trim();
            DangJ = oForm.Items.Item("DangJik").Specific.VALUE.Trim();
            JoongS = oForm.Items.Item("JoongS").Specific.VALUE.Trim();
            Bus1 = oForm.Items.Item("Bus1").Specific.VALUE.Trim();
            Bus2 = oForm.Items.Item("Bus2").Specific.VALUE.Trim();
            Bus3 = oForm.Items.Item("Bus3").Specific.VALUE.Trim();

            // Chk1,Chk2는 안씀
            

            if (oForm.DataSources.UserDataSources.Item("G5_YN").Value == "Y")
            {
                G5_YN = "Y";
            }
            else
            {
                G5_YN = "N";
            }

            if (oForm.DataSources.UserDataSources.Item("NoAccdnt").Value == "Y")
            {
                NoAccdnt = "Y";
            }
            else
            {
                NoAccdnt = "N";
            }

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
                    

            try
            {
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackSubReportFormula = new List<PSH_DataPackClass>();

                //	평일,휴일Check
                sQry = "Select b.U_DayType From [@PH_PY003A] a INNER JOIN [@PH_PY003B] b ON a.Code = b.Code WHERE a.U_CLTCOD = '" + CLTCOD + "' AND B.U_DATE = '" + DocDate + "'";
                oRecordSet.DoQuery(sQry);

                if (oRecordSet.Fields.Item("U_DayType").Value.ToString().Trim() != "2")
                {
                    // 평일
                    WinTitle = "[PH_PY560] 일출근현황";
                    ReportName = "PH_PY560_01.rpt";

                    //Formula
                    dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //
                    dataPackFormula.Add(new PSH_DataPackClass("@DocDate", DocDate.Substring(0, 4) + "-" + DocDate.Substring(4, 2) + "-" + DocDate.Substring(6, 2)));

                    //SubReport Formula
                    dataPackSubReportFormula.Add(new PSH_DataPackClass("@SukS", SukS, "PH_PY560_SUB3"));
                    dataPackSubReportFormula.Add(new PSH_DataPackClass("@YaS", YaS, "PH_PY560_SUB3"));
                    dataPackSubReportFormula.Add(new PSH_DataPackClass("@DangJ", DangJ, "PH_PY560_SUB3"));
                    dataPackSubReportFormula.Add(new PSH_DataPackClass("@Bus1", Bus1, "PH_PY560_SUB3"));
                    dataPackSubReportFormula.Add(new PSH_DataPackClass("@Bus2", Bus2, "PH_PY560_SUB3"));
                    dataPackSubReportFormula.Add(new PSH_DataPackClass("@Bus3", Bus3, "PH_PY560_SUB3"));

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@DocDate", DocDate)); 
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode)); 
                    
                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD, "PH_PY560_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@DocDate", DocDate, "PH_PY560_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY560_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD, "PH_PY560_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@DocDate", DocDate, "PH_PY560_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY560_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@G5_YN", G5_YN, "PH_PY560_SUB3"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, dataPackSubReportFormula, WinTitle, ReportName);
                }
                else
                {
                    // 휴일
                    WinTitle = "[PH_PY560] 일출근현황(휴일)";
                    if (NoAccdnt == "N")
                    {
                        ReportName = "PH_PY560_05.rpt";
                    }
                    else
                    {
                        ReportName = "PH_PY560_06.rpt";
                    }
                
                    //Formula
                    dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //
                    dataPackFormula.Add(new PSH_DataPackClass("@DocDate", DocDate.Substring(0, 4) + "-" + DocDate.Substring(4, 2) + "-" + DocDate.Substring(6, 2)));
                    dataPackFormula.Add(new PSH_DataPackClass("@JoongS", JoongS));
                    dataPackFormula.Add(new PSH_DataPackClass("@SukS", SukS));
                    dataPackFormula.Add(new PSH_DataPackClass("@YaS", YaS));
                    dataPackFormula.Add(new PSH_DataPackClass("@DangJ", DangJ));
                    dataPackFormula.Add(new PSH_DataPackClass("@Bus1", Bus1));
                    dataPackFormula.Add(new PSH_DataPackClass("@Bus2", Bus2));
                    dataPackFormula.Add(new PSH_DataPackClass("@Bus3", Bus3));

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@DocDate", DocDate));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY560_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
