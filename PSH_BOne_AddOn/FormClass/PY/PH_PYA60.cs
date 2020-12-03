using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 학자금신청내역(집계)
    /// </summary>
    internal class PH_PYA60 : PSH_BaseClass
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PYA60.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PYA60_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PYA60");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PYA60_CreateItems();
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
        private void PH_PYA60_CreateItems()
        {
            string sQry = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = null;
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P144' AND U_UseYN= 'Y'";
                oRecordSet.DoQuery(sQry);
                oForm.Items.Item("CLTCOD").Specific.ValidValues.Add("%", "전 사업장");
                while (!(oRecordSet.EoF))
                {
                    oForm.Items.Item("CLTCOD").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.Trim(), oRecordSet.Fields.Item(1).Value.Trim());
                    oRecordSet.MoveNext();
                }
                oForm.Items.Item("CLTCOD").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //시작일자
                oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("StdYear").Specific.String = DateTime.Now.ToString("yyyy");

                //분기
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("01", "1/4 혹은 1학기");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("02", "2/4");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("03", "3/4 혹은 2학기");
                oForm.Items.Item("Quarter").Specific.ValidValues.Add("04", "4/4");
                oForm.Items.Item("Quarter").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Quarter").DisplayDesc = true;

                //회차
                oForm.Items.Item("Count").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("Count").Specific.ValidValues.Add("01", "1차");
                oForm.Items.Item("Count").Specific.ValidValues.Add("02", "2차");
                oForm.Items.Item("Count").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("Count").DisplayDesc = true;

                //입학금포함여부
                oForm.Items.Item("EntFeeYN").Specific.ValidValues.Add("01", "포함");
                oForm.Items.Item("EntFeeYN").Specific.ValidValues.Add("02", "제외");
                oForm.Items.Item("EntFeeYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oForm.Items.Item("EntFeeYN").DisplayDesc = true;

                //출력선택 RAD
                oForm.DataSources.UserDataSources.Add("OptionDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Rad01").Specific.ValOn = "1";
                oForm.Items.Item("Rad01").Specific.ValOff = "0";
                oForm.Items.Item("Rad01").Specific.DataBind.SetBound(true, "", "OptionDS");
                oForm.Items.Item("Rad01").Specific.Selected = true;

                oForm.Items.Item("Rad02").Specific.ValOn = "2";
                oForm.Items.Item("Rad02").Specific.ValOff = "0";
                oForm.Items.Item("Rad02").Specific.DataBind.SetBound(true, "", "OptionDS");
                oForm.Items.Item("Rad02").Specific.GroupWith("Rad01");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PYA60_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                    ////2
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                    ////3

                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                    ////4
                    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                    ////7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:                    ////8
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:                    ////9
                    break;

                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                //    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:                    ////12
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                    ////16
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                    ////18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                    ////19
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:                    ////20
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:                    ////22
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:                    ////23
                    break;

                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:                    ////37
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:                    ////38
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag:                    ////39
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
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PYA60_Print_Report01);
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
        private void PH_PYA60_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty;
            string yyyy = string.Empty;
            string StdYear = string.Empty;
            string Quarter = string.Empty;
            string Count = string.Empty;
            string EntFeeYN = string.Empty;
            string OptBtnValue = string.Empty;

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
            yyyy = oForm.Items.Item("StdYear").Specific.Value.Trim();
            Quarter = oForm.Items.Item("Quarter").Specific.Value.Trim();
            Count = oForm.Items.Item("Count").Specific.Value.Trim();
            EntFeeYN = oForm.Items.Item("EntFeeYN").Specific.Value.Trim();
            OptBtnValue = oForm.DataSources.UserDataSources.Item("OptionDS").Value;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

                WinTitle = "학자금신청내역(집계)";

                if (OptBtnValue == "1")
                {
                    ReportName = "PH_PYA60_01.rpt";
                }
                else
                {
                    ReportName = "PH_PYA60_02.rpt";
                }

                //dataPackFormula

                //사업장명
                if (CLTCOD == "%")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", "전 사업장")); //년도
                }
                else
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //년도
                }

                //년
                dataPackFormula.Add(new PSH_DataPackClass("@YYYY", yyyy));

                //분기
                if (Quarter == "01")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@BUNGI", "1/4 혹은 1학기"));
                }
                else if (Quarter == "02")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@BUNGI", "2/4"));
                }
                else if (Quarter == "03")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@BUNGI", "3/4 혹은 2학기"));
                }
                else if (Quarter == "04")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@BUNGI", "4/4"));
                }
                else
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@BUNGI", "전체"));
                }

                //회차
                if (Count == "01")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@CHASU", "1차"));
                }
                else if (Count == "02")
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@CHASU", "2차"));
                }
                else
                {
                    dataPackFormula.Add(new PSH_DataPackClass("@CHASU", "전체"));
                }

                //dataPackParameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@StdYear", yyyy)); //등록기간(시작)
                dataPackParameter.Add(new PSH_DataPackClass("@Quarter", Quarter)); //등록기간(종료)
                dataPackParameter.Add(new PSH_DataPackClass("@Count", Count)); //등록기간(종료)
                dataPackParameter.Add(new PSH_DataPackClass("@EntFeeYN", EntFeeYN)); //등록기간(종료)

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PYA60_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
