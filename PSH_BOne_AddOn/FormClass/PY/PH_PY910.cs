using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 소득공제신고서출력
    /// </summary>
    internal class PH_PY910 : PSH_BaseClass
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY910.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY910_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY910");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY910_CreateItems();
                PH_PY910_FormItemEnabled();
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
        private void PH_PY910_CreateItems()
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

                // 년도
                oForm.DataSources.UserDataSources.Add("YYYY", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
                oForm.Items.Item("YYYY").Specific.DataBind.SetBound(true, "", "YYYY");
                oForm.DataSources.UserDataSources.Item("YYYY").Value = Convert.ToString(DateTime.Now.Year - 1);

                // 부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                // 담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");
                oForm.Items.Item("RspCode").DisplayDesc = true;

                // 반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 성명
                oForm.DataSources.UserDataSources.Add("MSTNAME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("MSTNAME").Specific.DataBind.SetBound(true, "", "MSTNAME");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY910_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        ///  <summary>
        ///  화면의 아이템 Enable 설정
        ///  </summary>
        public void PH_PY910_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        /// <param name="pVal">pVal</param>
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
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
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
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
                    ////22
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
                    ////23
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
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
        /// COMBO_SELECT 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;
            int i = 0;
            SAPbobsCOM.Recordset oRecordSet = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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

                                ////담당
                                ////삭제
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                ////반
                                ////삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;

                            ////부서가 바뀌면 담당 재설정
                            case "TeamCode":
                                ////담당은 그 부서의 담당만 표시
                                ////삭제
                                if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.VALUE + "' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");

                                ////반
                                ////삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;

                            ////담당이 바뀌면 반 재설정
                            case "RspCode":
                                ////반
                                ////삭제
                                if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                ////현재 사업장으로 다시 Qry
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                            case "MSTCOD":
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.VALUE.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("MSTNAME").Specific.VALUE = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();
                                break;
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
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY910_Print_Report01);
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
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PH_PY910_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty;
            string YYYY = string.Empty;
            string TeamCode = string.Empty;
            string RspCode = string.Empty;
            string ClsCode = string.Empty;
            string MSTCOD = string.Empty;
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
            YYYY = oForm.Items.Item("YYYY").Specific.Value.Trim();
            TeamCode = oForm.Items.Item("TeamCode").Specific.Value.Trim();
            RspCode = oForm.Items.Item("RspCode").Specific.Value.Trim();
            ClsCode = oForm.Items.Item("ClsCode").Specific.Value.Trim();
            MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.Trim();
            
            try
            {
                if (YYYY == "2013")
                {

                    //2013년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2013년";
                    ReportName = "PH_PY910_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB6"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }

                if (YYYY == "2014")
                {

                    //2014년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2014년";
                    ReportName = "PH_PY910_14_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB6"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }

                if (YYYY == "2015")
                {

                    //2015년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2015년";
                    ReportName = "PH_PY910_15_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB6"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }

                if (YYYY == "2016")
                {

                    //2016년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2016년";
                    ReportName = "PH_PY910_16_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB6"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }

                if (YYYY == "2017")
                {

                    //2017년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2017년";
                    ReportName = "PH_PY910_17_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB6"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }

                if (YYYY == "2018")
                {

                    //2018년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2018년";
                    ReportName = "PH_PY910_18_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB11"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB21"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB52"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB6"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }

                if (YYYY == "2019")
                {

                    //2019년귀속
                    WinTitle = "[PH_PY910] 소득공제신고서출력 2019년";
                    ReportName = "PH_PY910_19_01.rpt";

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB11"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB21"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB52"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY910_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY910_SUB6"));

                    formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
////using Microsoft.Office.Interop;
//using SAPbobsCOM;
//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;

//namespace PSH_BOne_AddOn
//{
//    internal class PH_PY910 : PSH_BaseClass
//    {
//        ////********************************************************************************
//        ////  File           : PH_PY910.cls
//        ////  Module         : PH
//        ////  Desc           : 소득공제신고서출력
//        ////  작성자         : NGY
//        ////  DATE           : 2013.12.19, 201612
//        ////********************************************************************************

//        public string oFormUniqueID;
//        //public SAPbouiCOM.Form oForm;

//        //'// 그리드 사용시
//        //Public oGrid1           As SAPbouiCOM.Grid
//        //Public oDS_PH_PY910     As SAPbouiCOM.DataTable
//        //
//        //'// 매트릭스 사용시
//        //Public oMat1 As SAPbouiCOM.Matrix
//        //Private oDS_PH_PY910A As SAPbouiCOM.DBDataSource
//        //Private oDS_PH_PY910B As SAPbouiCOM.DBDataSource

//        private string oLastItemUID;
//        private string oLastColUID;
//        private int oLastColRow;

//        public override void LoadForm(string oFromDocEntry01 = "")
//        {
//            int i = 0;
//            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//            try
//            {
//                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY910.srf");
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue.ToString() + "_" + (SubMain.Get_TotalFormsCount());
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//                {
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//                }
//                oFormUniqueID = "PH_PY910_" + SubMain.Get_TotalFormsCount();
//                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY910");

//                string strXml = string.Empty;
//                strXml = oXmlDoc.xml.ToString();

//                PSH_Globals.SBO_Application.LoadBatchActions(ref strXml);
//                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//                oForm.SupportedModes = -1;
//                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                //    oForm.DataBrowser.BrowseBy = "Code"

//                oForm.Freeze(true);
//                PH_PY910_CreateItems();
//                PH_PY910_EnableMenus();
//                PH_PY910_SetDocument(oFromDocEntry01);
//                //    Call PH_PY910_FormResize
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("Form_Load Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }
//            finally
//            {
//                oForm.Update();
//                oForm.Freeze(false);
//                oForm.Visible = true;
//                //메모리 해제
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
//            }           
//        }

//        private void PH_PY910_CreateItems()
//        {
//            string sQry = string.Empty;

//            SAPbouiCOM.ComboBox oCombo = null;            

//            SAPbobsCOM.Recordset oRecordSet = null;

//            oForm.Freeze(true);

//            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            try
//            {
//                PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//                ////사업장
//                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oCombo = oForm.Items.Item("CLTCOD").Specific;
//                oCombo.DataBind.SetBound(true, "", "CLTCOD");

//                oForm.Items.Item("CLTCOD").DisplayDesc = true;

//                //// 접속자에 따른 권한별 사업장 콤보박스세팅
//                DataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

//                oForm.Items.Item("YYYY").Specific.String = Convert.ToDouble(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "YYYY")) - 1;

//                ////부서
//                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oCombo = oForm.Items.Item("TeamCode").Specific;
//                oCombo.DataBind.SetBound(true, "", "TeamCode");

//                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");
//                oForm.Items.Item("TeamCode").DisplayDesc = true;

//                ////담당
//                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oCombo = oForm.Items.Item("RspCode").Specific;
//                oCombo.DataBind.SetBound(true, "", "RspCode");

//                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");
//                oForm.Items.Item("RspCode").DisplayDesc = true;

//                ////반
//                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oCombo = oForm.Items.Item("ClsCode").Specific;
//                oCombo.DataBind.SetBound(true, "", "ClsCode");
//                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";

//                sQry = sQry + " AND U_Char3 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");
//                oForm.Items.Item("ClsCode").DisplayDesc = true;

//                ////커서를 첫번째 ITEM으로 지정
//                oForm.ActiveItem = "CLTCOD";
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo); //메모리 해제
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
//            }           
//        }

//        private void PH_PY910_EnableMenus()
//        {
//            try
//            {
//                oForm.EnableMenu("1283", true);
//                ////제거
//                oForm.EnableMenu("1284", false);
//                ////취소
//                oForm.EnableMenu("1293", true);
//                ////행삭제
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        private void PH_PY910_SetDocument(string oFromDocEntry01)
//        {
//            try
//            {
//                if ((string.IsNullOrEmpty(oFromDocEntry01)))
//                {
//                    PH_PY910_FormItemEnabled();
//                    //PH_PY910_AddMatrixRow();
//                }
//                else
//                {
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                    PH_PY910_FormItemEnabled();
//                    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oForm.Items.Item("Code").Specific.VALUE = oFromDocEntry01;
//                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public void PH_PY910_FormItemEnabled()
//        {
//            try
//            {
//                oForm.Freeze(true);
//                if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
//                {

//                    oForm.EnableMenu("1281", true);
//                    ////문서찾기
//                    oForm.EnableMenu("1282", false);
//                    ////문서추가

//                }
//                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
//                {

//                    oForm.EnableMenu("1281", false);
//                    ////문서찾기
//                    oForm.EnableMenu("1282", true);
//                    ////문서추가
//                }
//                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
//                {


//                    oForm.EnableMenu("1281", true);
//                    ////문서찾기
//                    oForm.EnableMenu("1282", true);
//                    ////문서추가

//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
//        {
//            string sQry = null;
//            int i = 0;
//            SAPbouiCOM.ComboBox oCombo = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                switch (pval.EventType)
//                {
//                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//                        ////1

//                        if (pval.BeforeAction == true)
//                        {
//                            if (pval.ItemUID == "Btn01")
//                            {
//                                //PH_PY910_DataValidCheck();
//                                PH_PY910_Print_Report01();
//                            }
//                        }
//                        else if (pval.BeforeAction == false)
//                        {


//                        }
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//                        ////2
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//                        ////3
//                        switch (pval.ItemUID)
//                        {
//                            case "Mat01":
//                            case "Grid01":
//                                if (pval.Row > 0)
//                                {
//                                    oLastItemUID = pval.ItemUID;
//                                    oLastColUID = pval.ColUID;
//                                    oLastColRow = pval.Row;
//                                }
//                                break;
//                            default:
//                                oLastItemUID = pval.ItemUID;
//                                oLastColUID = "";
//                                oLastColRow = 0;
//                                break;
//                        }
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//                        ////4
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//                        ////5
//                        oForm.Freeze(true);
//                        if (pval.BeforeAction == true)
//                        {

//                        }
//                        else if (pval.BeforeAction == false)
//                        {
//                            if (pval.ItemChanged == true)
//                            {
//                                switch (pval.ItemUID)
//                                {
//                                    ////사업장이 바뀌면 부서와 담당 재설정
//                                    case "CLTCOD":
//                                        ////부서
//                                        oCombo = oForm.Items.Item("TeamCode").Specific;
//                                        ////삭제
//                                        if (oCombo.ValidValues.Count > 0)
//                                        {
//                                            for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1)
//                                            {
//                                                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//                                            }
//                                        }
//                                        ////현재 사업장으로 다시 Qry
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                                        DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");

//                                        ////담당
//                                        oCombo = oForm.Items.Item("RspCode").Specific;
//                                        ////삭제
//                                        if (oCombo.ValidValues.Count > 0)
//                                        {
//                                            for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1)
//                                            {
//                                                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//                                            }
//                                        }
//                                        ////현재 사업장으로 다시 Qry
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                                        DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");

//                                        ////반
//                                        oCombo = oForm.Items.Item("ClsCode").Specific;
//                                        ////삭제
//                                        if (oCombo.ValidValues.Count > 0)
//                                        {
//                                            for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1)
//                                            {
//                                                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//                                            }
//                                        }
//                                        ////현재 사업장으로 다시 Qry
//                                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = sQry + " AND U_Char3 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
//                                        DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");
//                                        break;

//                                    ////부서가 바뀌면 담당 재설정
//                                    case "TeamCode":
//                                        ////담당은 그 부서의 담당만 표시
//                                        oCombo = oForm.Items.Item("RspCode").Specific;
//                                        ////삭제
//                                        if (oCombo.ValidValues.Count > 0)
//                                        {
//                                            for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1)
//                                            {
//                                                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//                                            }
//                                        }
//                                        ////현재 사업장으로 다시 Qry
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.VALUE + "' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                                        DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");

//                                        ////반
//                                        oCombo = oForm.Items.Item("ClsCode").Specific;
//                                        ////삭제
//                                        if (oCombo.ValidValues.Count > 0)
//                                        {
//                                            for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1)
//                                            {
//                                                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//                                            }
//                                        }
//                                        ////현재 사업장으로 다시 Qry
//                                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = sQry + " AND U_Char3 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
//                                        DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");
//                                        break;

//                                    ////담당이 바뀌면 반 재설정
//                                    case "RspCode":
//                                        ////반
//                                        oCombo = oForm.Items.Item("ClsCode").Specific;
//                                        ////삭제
//                                        if (oCombo.ValidValues.Count > 0)
//                                        {
//                                            for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1)
//                                            {
//                                                oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//                                            }
//                                        }
//                                        ////현재 사업장으로 다시 Qry
//                                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = sQry + " AND U_Char3 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "'";
//                                        //UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
//                                        DataHelpClass.SetReDataCombo(oForm, sQry, oCombo, "Y");
//                                        break;

//                                }
//                            }
//                        }

//                        oForm.Freeze(false);
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_CLICK:
//                        ////6
//                        if (pval.BeforeAction == true)
//                        {
//                            switch (pval.ItemUID)
//                            {
//                                case "Mat01":
//                                    break;
//                                    //                    If pval.Row > 0 Then
//                                    //                        Call oMat1.SelectRow(pval.Row, True, False)
//                                    //                    End If
//                            }

//                            switch (pval.ItemUID)
//                            {
//                                case "Mat01":
//                                case "Grid01":
//                                    if (pval.Row > 0)
//                                    {
//                                        oLastItemUID = pval.ItemUID;
//                                        oLastColUID = pval.ColUID;
//                                        oLastColRow = pval.Row;
//                                    }
//                                    break;
//                                default:
//                                    oLastItemUID = pval.ItemUID;
//                                    oLastColUID = "";
//                                    oLastColRow = 0;
//                                    break;
//                            }
//                        }
//                        else if (pval.BeforeAction == false)
//                        {

//                        }
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//                        ////7
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//                        ////8
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:
//                        ////9
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//                        ////10
//                        oForm.Freeze(true);
//                        if (pval.BeforeAction == true)
//                        {

//                        }
//                        else if (pval.BeforeAction == false)
//                        {
//                            if (pval.ItemChanged == true)
//                            {
//                                switch (pval.ItemUID)
//                                {
//                                    case "MSTCOD":
//                                        ////사원명 찿아서 화면 표시 하기
//                                        //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE) + "'";
//                                        oRecordSet.DoQuery(sQry);
//                                        //UPGRADE_WARNING: oForm.Items(MSTNAME).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                                        oForm.Items.Item("MSTNAME").Specific.String = Strings.Trim(oRecordSet.Fields.Item("U_FullName").Value);
//                                        break;

//                                }
//                            }
//                        }
//                        oForm.Freeze(false);
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//                        ////11
//                        if (pval.BeforeAction == true)
//                        {
//                        }
//                        else if (pval.BeforeAction == false)
//                        {
//                            //oMat1.LoadFromDataSource

//                            PH_PY910_FormItemEnabled();
//                            //PH_PY910_AddMatrixRow();

//                        }
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:
//                        ////12
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
//                        ////16
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//                        ////17
//                        if (pval.BeforeAction == true)
//                        {
//                        }
//                        else if (pval.BeforeAction == false)
//                        {
//                            SubMain.Remove_Forms(oFormUniqueID);
//                            //UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                            oForm = null;
//                            //                Set oDS_PH_PY910A = Nothing
//                            //                Set oDS_PH_PY910B = Nothing

//                            //Set oMat1 = Nothing
//                            //Set oGrid1 = Nothing

//                        }
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//                        ////18
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//                        ////19
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
//                        ////20
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//                        ////21
//                        if (pval.BeforeAction == true)
//                        {

//                        }
//                        else if (pval.BeforeAction == false)
//                        {

//                        }
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:
//                        ////22
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
//                        ////23
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//                        ////27
//                        if (pval.BeforeAction == true)
//                        {

//                        }
//                        else if (pval.Before_Action == false)
//                        {

//                        }
//                        break;
//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:
//                        ////37
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_GRID_SORT:
//                        ////38
//                        break;

//                    //----------------------------------------------------------
//                    case SAPbouiCOM.BoEventTypes.et_Drag:
//                        ////39
//                        break;

//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_Raise_MenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo);
//                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
//            }            
//        }

//        public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
//        {
//            try
//            {
//                oForm.Freeze(true);

//                if ((pval.BeforeAction == true))
//                {
//                    switch (pval.MenuUID)
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
//                else if ((pval.BeforeAction == false))
//                {
//                    switch (pval.MenuUID)
//                    {
//                        case "1283":
//                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//                            PH_PY910_FormItemEnabled();
//                            //PH_PY910_AddMatrixRow();
//                            break;
//                        case "1284":
//                            break;
//                        case "1286":
//                            break;
//                        //            Case "1293":
//                        //                Call Raise_EVENT_ROW_DELETE(FormUID, pval, BubbleEvent)
//                        case "1281":
//                            ////문서찾기
//                            PH_PY910_FormItemEnabled();
//                            //PH_PY910_AddMatrixRow();
//                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                            break;
//                        case "1282":
//                            ////문서추가
//                            PH_PY910_FormItemEnabled();
//                            //PH_PY910_AddMatrixRow();
//                            break;
//                        case "1288":
//                        case "1289":
//                        case "1290":
//                        case "1291":
//                            PH_PY910_FormItemEnabled();
//                            break;
//                        case "1293":
//                            //// 행삭제
//                            //                '// [MAT1 용]
//                            //                 If oMat1.RowCount <> oMat1.VisualRowCount Then
//                            //                    oMat1.FlushToDataSource
//                            //
//                            //                    While (i <= oDS_PH_PY910B.Size - 1)
//                            //                        If oDS_PH_PY910B.GetValue("U_FILD01", i) = "" Then
//                            //                            oDS_PH_PY910B.RemoveRecord (i)
//                            //                            i = 0
//                            //                        Else
//                            //                            i = i + 1
//                            //                        End If
//                            //                    Wend
//                            //
//                            //                    For i = 0 To oDS_PH_PY910B.Size
//                            //                        Call oDS_PH_PY910B.setValue("U_LineNum", i, i + 1)
//                            //                    Next i
//                            //
//                            //                    oMat1.LoadFromDataSource
//                            //End If
//                            //PH_PY910_AddMatrixRow();
//                            break;
//                    }
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_Raise_MenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                oForm.Freeze(false);
//            }
//        }

//        public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
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
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
//        {
//            try
//            {
//                if (pval.BeforeAction == true)
//                {
//                }
//                else if (pval.BeforeAction == false)
//                {
//                }
//                switch (pval.ItemUID)
//                {
//                    case "Mat01":
//                        if (pval.Row > 0)
//                        {
//                            oLastItemUID = pval.ItemUID;
//                            oLastColUID = pval.ColUID;
//                            oLastColRow = pval.Row;
//                        }
//                        break;
//                    default:
//                        oLastItemUID = pval.ItemUID;
//                        oLastColUID = "";
//                        oLastColRow = 0;
//                        break;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public void PH_PY910_FormClear()
//        {
//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                string DocEntry = null;
//                //UPGRADE_WARNING: MDC_GetData.Get_ReData() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY910'", "");
//                if (Convert.ToDouble(DocEntry) == 0)
//                {
//                    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
//                }
//                else
//                {
//                    //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        public bool PH_PY910_Validate(string ValidateType)
//        {
//            bool functionReturnValue = false;
//            functionReturnValue = true;

//            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                if (DataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY910A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.VALUE + "'", 0, 1) == "Y")
//                {
//                    PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//                    functionReturnValue = false;
//                    goto PH_PY910_Validate_Exit;
//                }
//                //
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
//                functionReturnValue = false;

//                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY910_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            PH_PY910_Validate_Exit:
//            return functionReturnValue;
//        }

//        private void PH_PY910_Print_Report01()
//        {

//            string DocNum = null;
//            short ErrNum = 0;
//            string WinTitle = null;
//            string ReportName = null;
//            string sQry = null;
//            SAPbobsCOM.Recordset oRecordSet = null;

//            // ERROR: Not supported in C#: OnErrorStatement


//            string CLTCOD = null;
//            string yyyy = null;
//            string TeamCode = null;
//            string RspCode = null;
//            string ClsCode = null;
//            string MSTCOD = null;


//            oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//            /// ODBC 연결 체크
//            if (ConnectODBC() == false)
//            {
//                goto PH_PY910_Print_Report01_Error;
//            }


//            ////인자 MOVE , Trim 시키기..
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            CLTCOD = Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE);
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            yyyy = Strings.Trim(oForm.Items.Item("YYYY").Specific.VALUE);
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            TeamCode = Strings.Trim(oForm.Items.Item("TeamCode").Specific.VALUE);
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            RspCode = Strings.Trim(oForm.Items.Item("RspCode").Specific.VALUE);
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            ClsCode = Strings.Trim(oForm.Items.Item("ClsCode").Specific.VALUE);
//            //UPGRADE_WARNING: oForm.Items().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//            MSTCOD = Strings.Trim(oForm.Items.Item("MSTCOD").Specific.VALUE);


//            /// Crystal /~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~/

//            //2013년
//            if (yyyy == "2013")
//            {

//                WinTitle = "[PH_PY910] 소득공제신고서출력 2013년";
//                ReportName = "PH_PY910_01.rpt";
//                PSH_Globals.gRpt_Formula = new string[4];
//                PSH_Globals.gRpt_Formula_Value = new string[4];
//                PSH_Globals.gRpt_SRptSqry = new string[7];
//                PSH_Globals.gRpt_SRptName = new string[7];
//                PSH_Globals.gRpt_SFormula = new string[7, 2];
//                PSH_Globals.gRpt_SFormula_Value = new string[7, 2];

//                /// Formula 수식필드

//                /// SubReport


//                sQry = "EXEC [PH_PY910_02] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[1] = sQry;
//                PSH_Globals.gRpt_SRptName[1] = "PH_PY910_SUB1";

//                sQry = "EXEC [PH_PY910_03] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[2] = sQry;
//                PSH_Globals.gRpt_SRptName[2] = "PH_PY910_SUB2";

//                sQry = "EXEC [PH_PY910_04] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[3] = sQry;
//                PSH_Globals.gRpt_SRptName[3] = "PH_PY910_SUB3";

//                sQry = "EXEC [PH_PY910_05] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[4] = sQry;
//                PSH_Globals.gRpt_SRptName[4] = "PH_PY910_SUB4";

//                sQry = "EXEC [PH_PY910_06] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[5] = sQry;
//                PSH_Globals.gRpt_SRptName[5] = "PH_PY910_SUB5";

//                sQry = "EXEC [PH_PY910_07] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[6] = sQry;
//                PSH_Globals.gRpt_SRptName[6] = "PH_PY910_SUB6";

//                /// Procedure 실행"
//                sQry = "EXEC [PH_PY910_01] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";

//            }
//            //2014년
//            if (yyyy == "2014")
//            {

//                WinTitle = "[PH_PY910] 소득공제신고서출력 2014년";
//                ReportName = "PH_PY910_14_01.rpt";
//                PSH_Globals.gRpt_Formula = new string[4];
//                PSH_Globals.gRpt_Formula_Value = new string[4];
//                PSH_Globals.gRpt_SRptSqry = new string[8];
//                PSH_Globals.gRpt_SRptName = new string[8];
//                PSH_Globals.gRpt_SFormula = new string[8, 2];
//                PSH_Globals.gRpt_SFormula_Value = new string[8, 2];

//                /// Formula 수식필드

//                /// SubReport


//                sQry = "EXEC [PH_PY910_14_02] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[1] = sQry;
//                PSH_Globals.gRpt_SRptName[1] = "PH_PY910_SUB1";

//                sQry = "EXEC [PH_PY910_14_03] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[2] = sQry;
//                PSH_Globals.gRpt_SRptName[2] = "PH_PY910_SUB2";

//                sQry = "EXEC [PH_PY910_14_04] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[3] = sQry;
//                PSH_Globals.gRpt_SRptName[3] = "PH_PY910_SUB3";

//                sQry = "EXEC [PH_PY910_14_05] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[4] = sQry;
//                PSH_Globals.gRpt_SRptName[4] = "PH_PY910_SUB4";

//                sQry = "EXEC [PH_PY910_14_06] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[5] = sQry;
//                PSH_Globals.gRpt_SRptName[5] = "PH_PY910_SUB5";

//                sQry = "EXEC [PH_PY910_14_061] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[6] = sQry;
//                PSH_Globals.gRpt_SRptName[6] = "PH_PY910_SUB51";

//                sQry = "EXEC [PH_PY910_14_07] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[7] = sQry;
//                PSH_Globals.gRpt_SRptName[7] = "PH_PY910_SUB6";

//                /// Procedure 실행"
//                sQry = "EXEC [PH_PY910_14_01] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";

//            }
//            //2015년귀속
//            if (yyyy == "2015")
//            {

//                WinTitle = "[PH_PY910] 소득공제신고서출력 2015년 귀속";
//                ReportName = "PH_PY910_15_01.rpt";
//                PSH_Globals.gRpt_Formula = new string[4];
//                PSH_Globals.gRpt_Formula_Value = new string[4];
//                PSH_Globals.gRpt_SRptSqry = new string[8];
//                PSH_Globals.gRpt_SRptName = new string[8];
//                PSH_Globals.gRpt_SFormula = new string[8, 2];
//                PSH_Globals.gRpt_SFormula_Value = new string[8, 2];

//                /// Formula 수식필드

//                /// SubReport


//                sQry = "EXEC [PH_PY910_15_02] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[1] = sQry;
//                PSH_Globals.gRpt_SRptName[1] = "PH_PY910_SUB1";

//                sQry = "EXEC [PH_PY910_15_03] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[2] = sQry;
//                PSH_Globals.gRpt_SRptName[2] = "PH_PY910_SUB2";

//                sQry = "EXEC [PH_PY910_15_04] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[3] = sQry;
//                PSH_Globals.gRpt_SRptName[3] = "PH_PY910_SUB3";

//                sQry = "EXEC [PH_PY910_15_05] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[4] = sQry;
//                PSH_Globals.gRpt_SRptName[4] = "PH_PY910_SUB4";

//                sQry = "EXEC [PH_PY910_15_06] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[5] = sQry;
//                PSH_Globals.gRpt_SRptName[5] = "PH_PY910_SUB5";

//                sQry = "EXEC [PH_PY910_15_061] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[6] = sQry;
//                PSH_Globals.gRpt_SRptName[6] = "PH_PY910_SUB51";

//                sQry = "EXEC [PH_PY910_15_07] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[7] = sQry;
//                PSH_Globals.gRpt_SRptName[7] = "PH_PY910_SUB6";

//                /// Procedure 실행"
//                sQry = "EXEC [PH_PY910_15_01] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";

//            }

//            //2016년귀속
//            if (yyyy == "2016")
//            {

//                WinTitle = "[PH_PY910] 소득공제신고서출력 2016년 귀속";
//                ReportName = "PH_PY910_16_01.rpt";
//                PSH_Globals.gRpt_Formula = new string[4];
//                PSH_Globals.gRpt_Formula_Value = new string[4];
//                PSH_Globals.gRpt_SRptSqry = new string[8];
//                PSH_Globals.gRpt_SRptName = new string[8];
//                PSH_Globals.gRpt_SFormula = new string[8, 2];
//                PSH_Globals.gRpt_SFormula_Value = new string[8, 2];

//                /// Formula 수식필드

//                /// SubReport


//                sQry = "EXEC [PH_PY910_16_02] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[1] = sQry;
//                PSH_Globals.gRpt_SRptName[1] = "PH_PY910_SUB1";

//                sQry = "EXEC [PH_PY910_16_03] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[2] = sQry;
//                PSH_Globals.gRpt_SRptName[2] = "PH_PY910_SUB2";

//                sQry = "EXEC [PH_PY910_16_04] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[3] = sQry;
//                PSH_Globals.gRpt_SRptName[3] = "PH_PY910_SUB3";

//                sQry = "EXEC [PH_PY910_16_05] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[4] = sQry;
//                PSH_Globals.gRpt_SRptName[4] = "PH_PY910_SUB4";

//                sQry = "EXEC [PH_PY910_16_06] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[5] = sQry;
//                PSH_Globals.gRpt_SRptName[5] = "PH_PY910_SUB5";

//                sQry = "EXEC [PH_PY910_16_061] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[6] = sQry;
//                PSH_Globals.gRpt_SRptName[6] = "PH_PY910_SUB51";

//                sQry = "EXEC [PH_PY910_16_07] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[7] = sQry;
//                PSH_Globals.gRpt_SRptName[7] = "PH_PY910_SUB6";

//                /// Procedure 실행"
//                sQry = "EXEC [PH_PY910_16_01] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";

//            }

//            //2017년귀속
//            if (yyyy == "2017")
//            {

//                WinTitle = "[PH_PY910] 소득공제신고서출력 2017년 귀속";
//                ReportName = "PH_PY910_17_01.rpt";
//                PSH_Globals.gRpt_Formula = new string[4];
//                PSH_Globals.gRpt_Formula_Value = new string[4];
//                PSH_Globals.gRpt_SRptSqry = new string[8];
//                PSH_Globals.gRpt_SRptName = new string[8];
//                PSH_Globals.gRpt_SFormula = new string[8, 2];
//                PSH_Globals.gRpt_SFormula_Value = new string[8, 2];

//                /// Formula 수식필드

//                /// SubReport


//                sQry = "EXEC [PH_PY910_17_02] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[1] = sQry;
//                PSH_Globals.gRpt_SRptName[1] = "PH_PY910_SUB1";

//                sQry = "EXEC [PH_PY910_17_03] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[2] = sQry;
//                PSH_Globals.gRpt_SRptName[2] = "PH_PY910_SUB2";

//                sQry = "EXEC [PH_PY910_17_04] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[3] = sQry;
//                PSH_Globals.gRpt_SRptName[3] = "PH_PY910_SUB3";

//                sQry = "EXEC [PH_PY910_17_05] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[4] = sQry;
//                PSH_Globals.gRpt_SRptName[4] = "PH_PY910_SUB4";

//                sQry = "EXEC [PH_PY910_17_06] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[5] = sQry;
//                PSH_Globals.gRpt_SRptName[5] = "PH_PY910_SUB5";

//                sQry = "EXEC [PH_PY910_17_061] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[6] = sQry;
//                PSH_Globals.gRpt_SRptName[6] = "PH_PY910_SUB51";

//                sQry = "EXEC [PH_PY910_17_07] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[7] = sQry;
//                PSH_Globals.gRpt_SRptName[7] = "PH_PY910_SUB6";

//                /// Procedure 실행"
//                sQry = "EXEC [PH_PY910_17_01] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";

//            }

//            //2018년귀속
//            if (yyyy == "2018")
//            {

//                WinTitle = "[PH_PY910] 소득공제신고서출력 2018년 귀속";
//                ReportName = "PH_PY910_18_01.rpt";
//                PSH_Globals.gRpt_Formula = new string[4];
//                PSH_Globals.gRpt_Formula_Value = new string[4];
//                PSH_Globals.gRpt_SRptSqry = new string[11];
//                PSH_Globals.gRpt_SRptName = new string[11];
//                PSH_Globals.gRpt_SFormula = new string[11, 2];
//                PSH_Globals.gRpt_SFormula_Value = new string[11, 2];

//                /// Formula 수식필드

//                /// SubReport


//                sQry = "EXEC [PH_PY910_18_02] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[1] = sQry;
//                PSH_Globals.gRpt_SRptName[1] = "PH_PY910_SUB1";

//                sQry = "EXEC [PH_PY910_18_021] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[2] = sQry;
//                PSH_Globals.gRpt_SRptName[2] = "PH_PY910_SUB11";

//                sQry = "EXEC [PH_PY910_18_03] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[3] = sQry;
//                PSH_Globals.gRpt_SRptName[3] = "PH_PY910_SUB2";

//                sQry = "EXEC [PH_PY910_18_031] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[4] = sQry;
//                PSH_Globals.gRpt_SRptName[4] = "PH_PY910_SUB21";

//                sQry = "EXEC [PH_PY910_18_04] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[5] = sQry;
//                PSH_Globals.gRpt_SRptName[5] = "PH_PY910_SUB3";

//                sQry = "EXEC [PH_PY910_18_05] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[6] = sQry;
//                PSH_Globals.gRpt_SRptName[6] = "PH_PY910_SUB4";

//                sQry = "EXEC [PH_PY910_18_06] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[7] = sQry;
//                PSH_Globals.gRpt_SRptName[7] = "PH_PY910_SUB5";

//                sQry = "EXEC [PH_PY910_18_061] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[8] = sQry;
//                PSH_Globals.gRpt_SRptName[8] = "PH_PY910_SUB51";

//                sQry = "EXEC [PH_PY910_18_062] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[9] = sQry;
//                PSH_Globals.gRpt_SRptName[9] = "PH_PY910_SUB52";

//                sQry = "EXEC [PH_PY910_18_07] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//                PSH_Globals.gRpt_SRptSqry[10] = sQry;
//                PSH_Globals.gRpt_SRptName[10] = "PH_PY910_SUB6";

//                /// Procedure 실행"
//                sQry = "EXEC [PH_PY910_18_01] '" + CLTCOD + "', '" + yyyy + "',  '" + TeamCode + "', '" + RspCode + "', '" + ClsCode + "', '" + MSTCOD + "'";
//            }

//            oRecordSet.DoQuery(sQry);
//            if (oRecordSet.RecordCount == 0)
//            {
//                ErrNum = 1;
//                goto PH_PY910_Print_Report01_Error;
//            }

//            if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V", , 1) == false)
//            {
//                PSH_Globals.SBO_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }

//            //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//            oRecordSet = null;
//            return;
//        PH_PY910_Print_Report01_Error:

//            if (ErrNum == 1)
//            {
//                //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oRecordSet = null;
//                MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다. 확인해 주세요.", ref "E");
//            }
//            else
//            {
//                //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//                oRecordSet = null;
//                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY910_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//            }

//        }
//    }
//}