using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 원천징수영수증출력
    /// </summary>
    internal class PH_PY920 : PSH_BaseClass
    {
        private string oFormUniqueID01;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY920.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY920_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY920");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY920_CreateItems();
                PH_PY920_FormItemEnabled();
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
        private void PH_PY920_CreateItems()
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

                // 출력구분
                oForm.DataSources.UserDataSources.Add("Div", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Div").Specific.ValidValues.Add("1", "소득자보관용");
                oForm.Items.Item("Div").Specific.ValidValues.Add("2", "발행자보관용");
                oForm.Items.Item("Div").Specific.ValidValues.Add("3", "발행자보고용");
                oForm.Items.Item("Div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 재직구분
                oForm.DataSources.UserDataSources.Add("Div1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Div1").Specific.ValidValues.Add("1", "전체");
                oForm.Items.Item("Div1").Specific.ValidValues.Add("2", "재직자");
                oForm.Items.Item("Div1").Specific.ValidValues.Add("3", "퇴직자");
                oForm.Items.Item("Div1").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);


                // 출력선택 RAD
                oForm.DataSources.UserDataSources.Add("OptionDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Rad01").Specific.ValOn = "1";
                oForm.Items.Item("Rad01").Specific.ValOff = "0";
                oForm.Items.Item("Rad01").Specific.DataBind.SetBound(true, "", "OptionDS");

                oForm.Items.Item("Rad01").Specific.Selected = true;

                oForm.Items.Item("Rad02").Specific.ValOn = "2";
                oForm.Items.Item("Rad02").Specific.ValOff = "0";
                oForm.Items.Item("Rad02").Specific.DataBind.SetBound(true, "", "OptionDS");
                oForm.Items.Item("Rad02").Specific.GroupWith(("Rad01"));

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY920_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        ///  <summary>
        ///  화면의 아이템 Enable 설정
        ///  </summary>
        private void PH_PY920_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY920_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:                     //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:                         //2
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                        //3
                                        break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                       //4
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:                     //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK:                            //6
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:                     //7
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:              //8
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED:          //9
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE:                         //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:                      //11
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD:                  //12
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                        //16
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                      //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:                    //18
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:                  //19
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:                       //20
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:                      //21
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN:                    //22
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:                //23
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:                 //27
                    break;

                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED:                   //37
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT:                        //38
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag:                             //39
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
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
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
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
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
                                sQry  = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '9' AND U_UseYN= 'Y'";
                                sQry += " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.Trim() + "'";
                                sQry += " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.VALUE + "'";
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
                    SubMain.Remove_Forms(oFormUniqueID01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                }
                else if (pVal.Before_Action == false)
                {   
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
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY920_Print_Report01);
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
        private void PH_PY920_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty;
            string YYYY = string.Empty;
            string TeamCode = string.Empty;
            string RspCode = string.Empty;
            string ClsCode = string.Empty;
            string MSTCOD = string.Empty;
            string Div = string.Empty;
            string Gubun = string.Empty;
            string Div1 = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
                YYYY = oForm.Items.Item("YYYY").Specific.Value.Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.Trim();
                ClsCode = oForm.Items.Item("ClsCode").Specific.Value.Trim();
                MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.Trim();
                Div = oForm.Items.Item("Div").Specific.Value.Trim();
                Gubun = oForm.DataSources.UserDataSources.Item("OptionDS").Value; //출력구분
                Div1 = oForm.Items.Item("Div1").Specific.Value.Trim();   //재직구분

                if (Convert.ToInt16(YYYY) >= 2020)
                {
                    //2020년귀속
                    WinTitle = "[PH_PY920] 원천징수영수증출력 2020년";

                    if (Gubun == "1")
                    {
                        ReportName = "PH_PY920_20_01.rpt";
                    }
                    else
                    {
                        ReportName = "PH_PY920_20_02.rpt";
                    }

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장
                    dataPackFormula.Add(new PSH_DataPackClass("@YYYY", YYYY)); //년
                    dataPackFormula.Add(new PSH_DataPackClass("@Div", Div));

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));
                    dataPackParameter.Add(new PSH_DataPackClass("@Div", Div1));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB11"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB21"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB52"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB6"));

                    if (Gubun == "1")
                    {
                        formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                    }
                    else
                    {
                        formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
                    }
                }
                else
                {
                    //2019년귀속
                    WinTitle = "[PH_PY920] 원천징수영수증출력 2019년";

                    if (Gubun == "1")
                    {
                        ReportName = "PH_PY920_19_01.rpt";
                    }
                    else
                    {
                        ReportName = "PH_PY920_19_02.rpt";
                    }

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

                    //Formula
                    //dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장
                    dataPackFormula.Add(new PSH_DataPackClass("@YYYY", YYYY)); //년
                    dataPackFormula.Add(new PSH_DataPackClass("@Div", Div));

                    //Parameter
                    dataPackParameter.Add(new PSH_DataPackClass("@saup", CLTCOD)); //사업장
                    dataPackParameter.Add(new PSH_DataPackClass("@yyyy", YYYY));
                    dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                    dataPackParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD));
                    dataPackParameter.Add(new PSH_DataPackClass("@Div", Div1));

                    //SubReport Parameter
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB1"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB1"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB11"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB11"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB2"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB2"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB21"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB21"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB3"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB3"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB4"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB4"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB5"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB5"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB51"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB51"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB52"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB52"));

                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@saup", CLTCOD, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@yyyy", YYYY, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PH_PY920_SUB6"));
                    dataPackSubReportParameter.Add(new PSH_DataPackClass("@sabun", MSTCOD, "PH_PY920_SUB6"));

                    if (Gubun == "1")
                    {
                        formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                    }
                    else
                    {
                        formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY920_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
