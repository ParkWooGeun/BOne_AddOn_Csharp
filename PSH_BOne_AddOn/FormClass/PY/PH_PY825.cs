using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 전문직O/T현황
    /// </summary>
    internal class PH_PY825 : PSH_BaseClass
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY825.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY825_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY825");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY825_CreateItems();
                PH_PY825_FormItemEnabled();
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
        private void PH_PY825_CreateItems()
        {
            try
            {
                oForm.Freeze(true);

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //지급년월
                oForm.DataSources.UserDataSources.Add("FYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("FYM").Specific.String = DateTime.Now.ToString("yyyyMM");

                oForm.DataSources.UserDataSources.Add("TYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
                oForm.Items.Item("TYM").Specific.String = DateTime.Now.ToString("yyyyMM");
                                
                ////출력선택 RAD
                oForm.DataSources.UserDataSources.Add("OptionDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Rad01").Specific.ValOn = "1";
                oForm.Items.Item("Rad01").Specific.ValOff = "0";
                oForm.Items.Item("Rad01").Specific.DataBind.SetBound(true, "", "OptionDS");

                oForm.Items.Item("Rad01").Specific.Selected = true;

                oForm.Items.Item("Rad02").Specific.ValOn = "2";
                oForm.Items.Item("Rad02").Specific.ValOff = "0";
                oForm.Items.Item("Rad02").Specific.DataBind.SetBound(true, "", "OptionDS");
                oForm.Items.Item("Rad02").Specific.GroupWith(("Rad01"));

                oForm.Items.Item("Rad03").Specific.ValOn = "3";
                oForm.Items.Item("Rad03").Specific.ValOff = "0";
                oForm.Items.Item("Rad03").Specific.DataBind.SetBound(true, "", "OptionDS");
                oForm.Items.Item("Rad03").Specific.GroupWith(("Rad02"));

                //기준일자
                oForm.DataSources.UserDataSources.Add("StdDt", SAPbouiCOM.BoDataType.dt_DATE);
                oForm.Items.Item("StdDt").Specific.DataBind.SetBound(true, "", "StdDt");
                this.oForm.Items.Item("StdDt").Enabled = false;

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY825_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PH_PY825_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY825_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Raise_FormItemEvent
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">이벤트 </param>
        /// <param name="BubbleEvent">Bubble Event</param>
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
                                                              // Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY825_Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "Rad01")
                    {
                        oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱  안하면 에러남
                        oForm.Items.Item("StdDt").Enabled = false;
                    }
                    else if (pVal.ItemUID == "Rad02")
                    {
                        oForm.ActiveItem = "CLTCOD"; //사업장 콤보박스로 포커싱  안하면 에러남
                        oForm.Items.Item("StdDt").Enabled = false;
                    }
                    else if (pVal.ItemUID == "Rad03")
                    {
                        oForm.Items.Item("StdDt").Enabled = true;
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
        private void PH_PY825_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty;
            string FYM = string.Empty;
            string TYM = string.Empty;
            System.DateTime StdDt = default(System.DateTime);
            string JIGTYP = string.Empty;
            string OptBtnValue = string.Empty;

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
            FYM = oForm.Items.Item("FYM").Specific.Value.Trim();
            TYM = oForm.Items.Item("TYM").Specific.Value.Trim();

            OptBtnValue = oForm.DataSources.UserDataSources.Item("OptionDS").Value;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PH_PY825] 전문직O/T현황";
                if (OptBtnValue == "1")
                {
                    ReportName = "PH_PY825_01.rpt";
                }
                else if (OptBtnValue == "2")
                {
                    ReportName = "PH_PY825_05.rpt";
                }
                else
                {
                    ReportName = "PH_PY825_07.rpt";
                }
                
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장
                dataPackFormula.Add(new PSH_DataPackClass("@FYM", FYM.Substring(0, 4) + "-" + FYM.Substring(4, 2))); //시작년월
                dataPackFormula.Add(new PSH_DataPackClass("@TYM", TYM.Substring(0, 4) + "-" + TYM.Substring(4, 2))); //종료년월

                if (OptBtnValue == "1") // PH_PY825_01.rpt
                {
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD, "PH_PY825_SUB01"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@FYM", FYM, "PH_PY825_SUB01"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TYM", TYM, "PH_PY825_SUB01"));

                    dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@FYM", FYM)); //시작년월
                    dataPackParameter.Add(new PSH_DataPackClass("@TYM", TYM)); //종료년월

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }
                else if (OptBtnValue == "2")  // PH_PY825_05.rpt
                {
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD, "PH_PY825_SUB01"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@FYM", FYM, "PH_PY825_SUB01"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TYM", TYM, "PH_PY825_SUB01"));

                    dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@FYM", FYM)); //시작년월
                    dataPackParameter.Add(new PSH_DataPackClass("@TYM", TYM)); //종료년월

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }
                else // PH_PY825_07
                {
                     StdDt = DateTime.ParseExact(oForm.Items.Item("StdDt").Specific.Value, "yyyyMMdd", null);

                    dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@FYM", FYM)); //시작년월
                    dataPackParameter.Add(new PSH_DataPackClass("@TYM", TYM)); //종료년월
                    dataPackParameter.Add(new PSH_DataPackClass("@StdDt", StdDt)); // 기준일자

                    formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY825_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
