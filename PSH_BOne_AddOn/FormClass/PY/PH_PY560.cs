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
        public override void LoadForm(string oFromDocEntry01)
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
                                                             // Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
//	internal class PH_PY560
//	{
//////********************************************************************************
//////  File           : PH_PY560.cls
//////  Module         : PH
//////  Desc           : 일출근현황
//////  작성자         : NGY
//////  DATE           : 2012.12.03
//////********************************************************************************

//		public string oFormUniqueID;
//		public SAPbouiCOM.Form oForm;

////'// 그리드 사용시
////Public oGrid1           As SAPbouiCOM.Grid
////Public oDS_PH_PY560     As SAPbouiCOM.DataTable
////
////'// 매트릭스 사용시
////Public oMat1 As SAPbouiCOM.Matrix
////Private oDS_PH_PY560A As SAPbouiCOM.DBDataSource
////Private oDS_PH_PY560B As SAPbouiCOM.DBDataSource

//		private string oLastItemUID;
//		private string oLastColUID;
//		private int oLastColRow;

//		public void LoadForm(string oFromDocEntry01 = "")
//		{

//			int i = 0;
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oXmlDoc.load(MDC_Globals.SP_Path + "\\" + MDC_Globals.SP_Screen + "\\PH_PY560.srf");
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (GetTotalFormsCount());
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue + (GetCurrentFormsCount() * 10);
//			oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue + (GetCurrentFormsCount() * 10);
//			for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++) {
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//				oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//			}
//			oFormUniqueID = "PH_PY560_" + GetTotalFormsCount();
//			SubMain.AddForms(this, oFormUniqueID, "PH_PY560");
//			MDC_Globals.Sbo_Application.LoadBatchActions(out (oXmlDoc.xml));
//			oForm = MDC_Globals.Sbo_Application.Forms.Item(oFormUniqueID);

//			oForm.SupportedModes = -1;
//			oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//			//    oForm.DataBrowser.BrowseBy = "Code"

//			oForm.Freeze(true);
//			PH_PY560_CreateItems();
//			PH_PY560_EnableMenus();
//			PH_PY560_SetDocument(oFromDocEntry01);
//			//    Call PH_PY560_FormResize

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

//		private bool PH_PY560_CreateItems()
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


//			////----------------------------------------------------------------------------------------------
//			//// 아이템 설정
//			////----------------------------------------------------------------------------------------------

//			////사업장
//			oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo = oForm.Items.Item("CLTCOD").Specific;
//			oCombo.DataBind.SetBound(true, "", "CLTCOD");

//			oForm.Items.Item("CLTCOD").DisplayDesc = true;

//			//// 접속자에 따른 권한별 사업장 콤보박스세팅
//			MDC_SetMod.CLTCOD_Select(oForm, "CLTCOD");

//			////기준일자
//			oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
//			//UPGRADE_WARNING: oForm.Items().Specific.DataBind 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
//			oForm.DataSources.UserDataSources.Item("DocDate").Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYYMMDD");

//			////부서
//			oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//			oCombo = oForm.Items.Item("TeamCode").Specific;
//			oCombo.DataBind.SetBound(true, "", "TeamCode");
//			//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//			MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//			oForm.Items.Item("TeamCode").DisplayDesc = true;


//			////주간
//			oCombo = oForm.Items.Item("Bus1").Specific;
//			oCombo.ValidValues.Add("", "");
//			oCombo.ValidValues.Add("운행", "운행");
//			oCombo.ValidValues.Add("미운행", "미운행");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			////주잔
//			oCombo = oForm.Items.Item("Bus2").Specific;
//			oCombo.ValidValues.Add("", "");
//			oCombo.ValidValues.Add("운행", "운행");
//			oCombo.ValidValues.Add("운행(20:00)", "운행(20:00)");
//			oCombo.ValidValues.Add("운행(20:30)", "운행(20:30)");
//			oCombo.ValidValues.Add("미운행", "미운행");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			////야간
//			oCombo = oForm.Items.Item("Bus3").Specific;
//			oCombo.ValidValues.Add("", "");
//			oCombo.ValidValues.Add("운행(05:30)", "운행(05:30)");
//			oCombo.ValidValues.Add("운행(08:00)", "운행(08:00)");
//			oCombo.ValidValues.Add("운행(08:30)", "운행(08:30)");
//			oCombo.ValidValues.Add("미운행", "미운행");
//			oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//			//   ' 통근버스운행
//			//    '//체크박스1
//			//    Call oForm.DataSources.UserDataSources.Add("Chk1", dt_SHORT_TEXT, 1)
//			//    Set oCheck = oForm.Items("Chk1").Specific
//			//    oCheck.DataBind.SetBound True, "", "Chk1"
//			//
//			//    '/체크박스2
//			//    Call oForm.DataSources.UserDataSources.Add("Chk2", dt_SHORT_TEXT, 1)
//			//    Set oCheck = oForm.Items("Chk2").Specific
//			//    oCheck.DataBind.SetBound True, "", "Chk2"

//			//G5 포함
//			oForm.DataSources.UserDataSources.Add("G5_YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			oCheck = oForm.Items.Item("G5_YN").Specific;
//			oCheck.DataBind.SetBound(true, "", "G5_YN");
//			//UPGRADE_WARNING: oForm.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("G5_YN").Specific.Checked = true;

//			//무재해일수 조회용(2014.09.11 송명규 추가)
//			oForm.DataSources.UserDataSources.Add("NoAccdnt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
//			oCheck = oForm.Items.Item("NoAccdnt").Specific;
//			oCheck.DataBind.SetBound(true, "", "NoAccdnt");
//			//UPGRADE_WARNING: oForm.Items().Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			oForm.Items.Item("NoAccdnt").Specific.Checked = false;

//			//    '//매트릭스컬럼
//			//    Set oColumn = oMat1.Columns("FILD01")
//			//    oColumn.Editable = True


//			////커서를 첫번째 ITEM으로 지정
//			//    oForm.ActiveItem = "CLTCOD"
//			oForm.Items.Item("DocDate").Click();


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
//			PH_PY560_CreateItems_Error:

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
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_CreateItems_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}

//		private void PH_PY560_EnableMenus()
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.EnableMenu("1283", true);
//			////제거
//			oForm.EnableMenu("1284", false);
//			////취소
//			oForm.EnableMenu("1293", true);
//			////행삭제

//			return;
//			PH_PY560_EnableMenus_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_EnableMenus_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		private void PH_PY560_SetDocument(string oFromDocEntry01)
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement


//			if ((string.IsNullOrEmpty(oFromDocEntry01))) {
//				PH_PY560_FormItemEnabled();
//				PH_PY560_AddMatrixRow();
//			} else {
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//				PH_PY560_FormItemEnabled();
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//				oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//			}
//			return;
//			PH_PY560_SetDocument_Error:

//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_SetDocument_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY560_FormItemEnabled()
//		{
//			SAPbouiCOM.ComboBox oCombo = null;

//			 // ERROR: Not supported in C#: OnErrorStatement



//			oForm.Freeze(true);
//			if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {

//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", false);
//				////문서추가

//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)) {

//				oForm.EnableMenu("1281", false);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가
//			} else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)) {


//				oForm.EnableMenu("1281", true);
//				////문서찾기
//				oForm.EnableMenu("1282", true);
//				////문서추가

//			}
//			oForm.Freeze(false);
//			return;
//			PH_PY560_FormItemEnabled_Error:

//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_FormItemEnabled_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
//						if (pval.ItemUID == "Btn01") {
//							PH_PY560_Print_Report01();
//						}
//					} else if (pval.BeforeAction == false) {


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
//						case "Mat01":
//						case "Grid01":
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
//							switch (pval.ItemUID) {
//								////근무형태가 바뀌면 근무조 재설정
//								case "ShiftDat":
//									oCombo = oForm.Items.Item("GNMUJO").Specific;
//									////삭제
//									if (oCombo.ValidValues.Count > 0) {
//										for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//											oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//										}
//									}
//									//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("ShiftDat").Specific.VALUE + "'";
//									MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);
//									break;
//								////사업장이 바뀌면 부서와 담당 재설정
//								case "CLTCOD":
//									////부서
//									oCombo = oForm.Items.Item("TeamCode").Specific;
//									////삭제
//									if (oCombo.ValidValues.Count > 0) {
//										for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//											oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//										}
//									}
//									////현재 사업장으로 다시 Qry
//									//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//									sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//									MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//									break;
//							}

//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pval.BeforeAction == true) {
//						switch (pval.ItemUID) {
//							case "Mat01":
//								break;
//							//                    If pval.Row > 0 Then
//							//                        Call oMat1.SelectRow(pval.Row, True, False)
//							//                    End If
//						}

//						switch (pval.ItemUID) {
//							case "Mat01":
//							case "Grid01":
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
//							switch (pval.ItemUID) {
//								//                        Case "Code"
//								//                            '//사원명 찿아서 화면 표시 하기
//								//                            sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" & Trim(oForm.Items("Code").Specific.Value) & "'"
//								//                            oRecordSet.DoQuery sQry
//								//                            oForm.Items("CodeName").Specific.String = Trim(oRecordSet.Fields("U_FullName").Value)
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
//						//oMat1.LoadFromDataSource

//						PH_PY560_FormItemEnabled();
//						PH_PY560_AddMatrixRow();

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
//						//                Set oDS_PH_PY560A = Nothing
//						//                Set oDS_PH_PY560B = Nothing

//						//Set oMat1 = Nothing
//						//Set oGrid1 = Nothing

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
//				}
//			} else if ((pval.BeforeAction == false)) {
//				switch (pval.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY560_FormItemEnabled();
//						PH_PY560_AddMatrixRow();
//						break;
//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY560_FormItemEnabled();
//						PH_PY560_AddMatrixRow();
//						oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//						break;
//					case "1282":
//						////문서추가
//						PH_PY560_FormItemEnabled();
//						PH_PY560_AddMatrixRow();
//						break;
//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY560_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						//                '// [MAT1 용]
//						//                 If oMat1.RowCount <> oMat1.VisualRowCount Then
//						//                    oMat1.FlushToDataSource
//						//
//						//                    While (i <= oDS_PH_PY560B.Size - 1)
//						//                        If oDS_PH_PY560B.GetValue("U_FILD01", i) = "" Then
//						//                            oDS_PH_PY560B.RemoveRecord (i)
//						//                            i = 0
//						//                        Else
//						//                            i = i + 1
//						//                        End If
//						//                    Wend
//						//
//						//                    For i = 0 To oDS_PH_PY560B.Size
//						//                        Call oDS_PH_PY560B.setValue("U_LineNum", i, i + 1)
//						//                    Next i
//						//
//						//                    oMat1.LoadFromDataSource
//						//End If
//						PH_PY560_AddMatrixRow();
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

//		public void PH_PY560_AddMatrixRow()
//		{
//			int oRow = 0;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oForm.Freeze(true);

//			////[Mat1 용]
//			//oMat1.FlushToDataSource
//			//oRow = oMat1.VisualRowCount
//			//
//			//    If oMat1.VisualRowCount > 0 Then
//			//        If Trim(oDS_PH_PY560B.GetValue("U_FILD01", oRow - 1)) <> "" Then
//			//            If oDS_PH_PY560B.Size <= oMat1.VisualRowCount Then
//			//                oDS_PH_PY560B.InsertRecord (oRow)
//			//            End If
//			//            oDS_PH_PY560B.Offset = oRow
//			//            oDS_PH_PY560B.setValue "U_LineNum", oRow, oRow + 1
//			//            oDS_PH_PY560B.setValue "U_FILD01", oRow, ""
//			//            oDS_PH_PY560B.setValue "U_FILD02", oRow, ""
//			//            oDS_PH_PY560B.setValue "U_FILD03", oRow, 0
//			//            oMat1.LoadFromDataSource
//			//        Else
//			//            oDS_PH_PY560B.Offset = oRow - 1
//			//            oDS_PH_PY560B.setValue "U_LineNum", oRow - 1, oRow
//			//            oDS_PH_PY560B.setValue "U_FILD01", oRow - 1, ""
//			//            oDS_PH_PY560B.setValue "U_FILD02", oRow - 1, ""
//			//            oDS_PH_PY560B.setValue "U_FILD03", oRow - 1, 0
//			//            oMat1.LoadFromDataSource
//			//        End If
//			//    ElseIf oMat1.VisualRowCount = 0 Then
//			//        oDS_PH_PY560B.Offset = oRow
//			//        oDS_PH_PY560B.setValue "U_LineNum", oRow, oRow + 1
//			//        oDS_PH_PY560B.setValue "U_FILD01", oRow, ""
//			//        oDS_PH_PY560B.setValue "U_FILD02", oRow, ""
//			//        oDS_PH_PY560B.setValue "U_FILD03", oRow, 0
//			//        oMat1.LoadFromDataSource
//			//    End If

//			oForm.Freeze(false);
//			return;
//			PH_PY560_AddMatrixRow_Error:
//			oForm.Freeze(false);
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_AddMatrixRow_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public void PH_PY560_FormClear()
//		{
//			 // ERROR: Not supported in C#: OnErrorStatement

//			string DocEntry = null;
//			//UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocEntry = MDC_GetData.Get_ReData(ref "AutoKey", ref "ObjectCode", ref "ONNM", ref "'PH_PY560'", ref "");
//			if (Convert.ToDouble(DocEntry) == 0) {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//			} else {
//				//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//			}
//			return;
//			PH_PY560_FormClear_Error:
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_FormClear_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}

//		public bool PH_PY560_DataValidCheck()
//		{
//			bool functionReturnValue = false;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			functionReturnValue = false;
//			int i = 0;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			//    '//----------------------------------------------------------------------------------
//			//    '//필수 체크
//			//    '//----------------------------------------------------------------------------------
//			//    If Trim(oDS_PH_PY560A.GetValue("Code", 0)) = "" Then
//			//        Sbo_Application.SetStatusBarMessage "사원번호는 필수입니다.", bmt_Short, True
//			//        oForm.Items("Code").CLICK ct_Regular
//			//        PH_PY560_DataValidCheck = False
//			//        Exit Function
//			//    End If
//			//
//			//    oMat1.FlushToDataSource
//			//    '// Matrix 마지막 행 삭제(DB 저장시)
//			//    If oDS_PH_PY560B.Size > 1 Then oDS_PH_PY560B.RemoveRecord (oDS_PH_PY560B.Size - 1)
//			//    oMat1.LoadFromDataSource

//			functionReturnValue = true;
//			return functionReturnValue;


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY560_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		public bool PH_PY560_Validate(string ValidateType)
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
//			//UPGRADE_WARNING: MDC_Company_Common.GetValue(SELECT Canceled FROM [PH_PY560A] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.VALUE & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			if (MDC_Company_Common.GetValue("SELECT Canceled FROM [@PH_PY560A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y") {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//				functionReturnValue = false;
//				goto PH_PY560_Validate_Exit;
//			}
//			//
//			if (ValidateType == "수정") {

//			} else if (ValidateType == "행삭제") {

//			} else if (ValidateType == "취소") {

//			}
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY560_Validate_Exit:
//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return functionReturnValue;
//			PH_PY560_Validate_Error:
//			functionReturnValue = false;
//			MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}


//		private void PH_PY560_Print_Report01()
//		{

//			string DocNum = null;
//			short ErrNum = 0;
//			string WinTitle = null;
//			string ReportName = null;
//			string sQry = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			string CLTCOD = null;
//			string DocDate = null;
//			string TeamCode = null;

//			string SukS = null;
//			string YaS = null;
//			string DangJ = null;
//			string Chk1 = null;
//			string Chk2 = null;
//			string JoongS = null;
//			string Bus1 = null;
//			string Bus2 = null;
//			string Bus3 = null;
//			string BusAdd = null;
//			string G5_YN = null;
//			string NoAccdnt = null;

//			SAPbouiCOM.CheckBox oCheck = null;

//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			SAPbouiCOM.ProgressBar ProgBar01 = null;
//			ProgBar01 = MDC_Globals.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

//			/// ODBC 연결 체크
//			if (ConnectODBC() == false) {
//				goto PH_PY560_Print_Report01_Error;
//			}


//			////인자 MOVE , Trim 시키기..
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DocDate = Strings.Trim(oForm.Items.Item("DocDate").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			TeamCode = Strings.Trim(oForm.Items.Item("TeamCode").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			SukS = Strings.Trim(oForm.Items.Item("SukS").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			YaS = Strings.Trim(oForm.Items.Item("YaS").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			DangJ = Strings.Trim(oForm.Items.Item("DangJik").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			JoongS = Strings.Trim(oForm.Items.Item("JoongS").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Bus1 = Strings.Trim(oForm.Items.Item("Bus1").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Bus2 = Strings.Trim(oForm.Items.Item("Bus2").Specific.VALUE);
//			//UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//			Bus3 = Strings.Trim(oForm.Items.Item("Bus3").Specific.VALUE);
//			// BusAdd = Trim(oForm.Items("BusAdd").Specific.VALUE) '추가운행횟수

//			//    If oForm.DataSources.UserDataSources.Item("Chk1").VALUE = "Y" Then
//			//       Chk1 = "05:30운행 "
//			//    Else
//			//        Chk1 = ""
//			//    End If
//			//
//			//    If oForm.DataSources.UserDataSources.Item("Chk2").VALUE = "Y" Then
//			//       Chk2 = "08:30운행"
//			//    Else: Chk2 = ""
//			//    End If

//			if (oForm.DataSources.UserDataSources.Item("G5_YN").Value == "Y") {
//				G5_YN = "Y";
//			} else {
//				G5_YN = "N";
//			}

//			if (oForm.DataSources.UserDataSources.Item("NoAccdnt").Value == "Y") {
//				NoAccdnt = "Y";
//			} else {
//				NoAccdnt = "N";
//			}


//			////평일,휴일Check
//			sQry = "Select b.U_DayType From [@PH_PY003A] a INNER JOIN [@PH_PY003B] b ON a.Code = b.Code WHERE a.U_CLTCOD = '" + CLTCOD + "' AND B.U_DATE = '" + DocDate + "'";
//			oRecordSet.DoQuery(sQry);

//			if (oRecordSet.Fields.Item("U_DayType").Value != "2") {
//				//------------------------------평일-------------------------------------
//				WinTitle = "[PH_PY560] 일출근현황(평일)";
//				ReportName = "PH_PY560_01.rpt";
//				MDC_Globals.gRpt_Formula = new string[3];
//				MDC_Globals.gRpt_Formula_Value = new string[3];
//				/// Formula 수식필드

//				MDC_Globals.gRpt_Formula[1] = "CLTCOD";
//				sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" + CLTCOD + "'";
//				oRecordSet.DoQuery(sQry);
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.gRpt_Formula_Value[1] = oRecordSet.Fields.Item(0).Value;

//				MDC_Globals.gRpt_Formula[2] = "DocDate";
//				MDC_Globals.gRpt_Formula_Value[2] = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DocDate, "####-##-##");
//				MDC_Globals.gRpt_SRptSqry = new string[4];
//				MDC_Globals.gRpt_SRptName = new string[4];
//				MDC_Globals.gRpt_SFormula = new string[4, 11];
//				MDC_Globals.gRpt_SFormula_Value = new string[4, 11];


//				/// SubReport


//				MDC_Globals.gRpt_SFormula[1, 1] = "";
//				MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

//				MDC_Globals.gRpt_SFormula[2, 1] = "CLTCOD";
//				MDC_Globals.gRpt_SFormula_Value[2, 1] = CLTCOD;

//				MDC_Globals.gRpt_SFormula[3, 1] = "SukS";
//				MDC_Globals.gRpt_SFormula_Value[3, 1] = SukS;

//				MDC_Globals.gRpt_SFormula[3, 2] = "YaS";
//				MDC_Globals.gRpt_SFormula_Value[3, 2] = YaS;

//				MDC_Globals.gRpt_SFormula[3, 3] = "DangJ";
//				MDC_Globals.gRpt_SFormula_Value[3, 3] = DangJ;

//				//            gRpt_SFormula(3, 4) = "Chk1"
//				//            gRpt_SFormula_Value(3, 4) = Chk1
//				//
//				//            gRpt_SFormula(3, 5) = "Chk2"
//				//            gRpt_SFormula_Value(3, 5) = Chk2

//				MDC_Globals.gRpt_SFormula[3, 4] = "Bus1";
//				MDC_Globals.gRpt_SFormula_Value[3, 4] = Bus1;

//				MDC_Globals.gRpt_SFormula[3, 5] = "Bus2";
//				MDC_Globals.gRpt_SFormula_Value[3, 5] = Bus2;

//				MDC_Globals.gRpt_SFormula[3, 6] = "Bus3";
//				MDC_Globals.gRpt_SFormula_Value[3, 6] = Bus3;


//				sQry = "EXEC [PH_PY560_02] '" + CLTCOD + "', '" + DocDate + "', '" + TeamCode + "'";
//				MDC_Globals.gRpt_SRptSqry[1] = sQry;
//				MDC_Globals.gRpt_SRptName[1] = "PH_PY560_SUB1";

//				sQry = "EXEC [PH_PY560_04] '" + CLTCOD + "', '" + DocDate + "', '" + TeamCode + "','" + G5_YN + "'";
//				MDC_Globals.gRpt_SRptSqry[3] = sQry;
//				MDC_Globals.gRpt_SRptName[3] = "PH_PY560_SUB3";


//				/// Procedure 실행"
//				sQry = "EXEC [PH_PY560_01] '" + CLTCOD + "', '" + DocDate + "', '" + TeamCode + "'";

//				//--------평일끝------------------------------------------------------------------------
//			} else {
//				//------------------------------휴일-------------------------------------
//				WinTitle = "[PH_PY560] 일출근현황(휴일)";
//				ReportName = "PH_PY560_05.rpt";
//				MDC_Globals.gRpt_Formula = new string[11];
//				MDC_Globals.gRpt_Formula_Value = new string[11];

//				/// Formula 수식필드

//				MDC_Globals.gRpt_Formula[1] = "CLTCOD";
//				sQry = "SELECT U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y' AND U_Code = '" + CLTCOD + "'";
//				oRecordSet.DoQuery(sQry);
//				//UPGRADE_WARNING: oRecordSet.Fields().VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//				MDC_Globals.gRpt_Formula_Value[1] = oRecordSet.Fields.Item(0).Value;

//				MDC_Globals.gRpt_Formula[2] = "DocDate";
//				MDC_Globals.gRpt_Formula_Value[2] = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DocDate, "####-##-##");

//				MDC_Globals.gRpt_Formula[3] = "JoongS";
//				MDC_Globals.gRpt_Formula_Value[3] = JoongS;

//				MDC_Globals.gRpt_Formula[4] = "SukS";
//				MDC_Globals.gRpt_Formula_Value[4] = SukS;

//				MDC_Globals.gRpt_Formula[5] = "YaS";
//				MDC_Globals.gRpt_Formula_Value[5] = YaS;

//				MDC_Globals.gRpt_Formula[6] = "DangJ";
//				MDC_Globals.gRpt_Formula_Value[6] = DangJ;

//				MDC_Globals.gRpt_Formula[7] = "Bus1";
//				MDC_Globals.gRpt_Formula_Value[7] = Bus1;

//				MDC_Globals.gRpt_Formula[8] = "Bus2";
//				MDC_Globals.gRpt_Formula_Value[8] = Bus2;

//				MDC_Globals.gRpt_Formula[9] = "Bus3";
//				MDC_Globals.gRpt_Formula_Value[9] = Bus3;
//				MDC_Globals.gRpt_SRptSqry = new string[2];
//				MDC_Globals.gRpt_SRptName = new string[2];
//				MDC_Globals.gRpt_SFormula = new string[2, 2];
//				MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

//				//             gRpt_Formula(10) = "BusAdd"
//				//             gRpt_Formula_Value(10) = BusAdd


//				/// SubReport


//				/// Procedure 실행"
//				if (NoAccdnt == "N") {
//					sQry = "EXEC [PH_PY560_05] '" + CLTCOD + "', '" + DocDate + "', '" + TeamCode + "'";
//				} else {
//					sQry = "EXEC [PH_PY560_06] '" + CLTCOD + "', '" + DocDate + "', '" + TeamCode + "'";
//				}

//				//--------휴일끝------------------------------------------------------------------------
//			}

//			//    oRecordSet.DoQuery sQry
//			//    If oRecordSet.RecordCount = 0 Then
//			//        ErrNum = 1
//			//        GoTo PH_PY560_Print_Report01_Error
//			//    End If

//			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V", , 1) == false) {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			return;
//			PH_PY560_Print_Report01_Error:


//			ProgBar01.Value = 100;
//			ProgBar01.Stop();
//			//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			ProgBar01 = null;

//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;

//			if (ErrNum == 1) {
//				MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
//			} else {
//				MDC_Globals.Sbo_Application.SetStatusBarMessage("PH_PY560_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			}

//		}
//	}
//}
