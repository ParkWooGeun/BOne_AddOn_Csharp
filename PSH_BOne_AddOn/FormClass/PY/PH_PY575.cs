using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 근태기찰현황
    /// </summary>
    internal class PH_PY575 : PSH_BaseClass
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
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY575.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY575_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY575");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PH_PY575_CreateItems();
                PH_PY575_FormItemEnabled();
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
        private void PH_PY575_CreateItems()
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
                oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMM") + "01";

                oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
                oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");

                // 부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "Y");
                oForm.Items.Item("TeamCode").DisplayDesc = true;

                // 담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "Y");
                oForm.Items.Item("RspCode").DisplayDesc = true;

                // 반
                oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

                // 사번
                oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

                // 성명
                oForm.DataSources.UserDataSources.Add("FullName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("FullName").Specific.DataBind.SetBound(true, "", "FullName");

                // 근무형태
                oForm.DataSources.UserDataSources.Add("ShiftDat", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ShiftDat").Specific.DataBind.SetBound(true, "", "ShiftDat");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P154' AND U_UseYN= 'Y'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ShiftDat").Specific, "Y");
                oForm.Items.Item("ShiftDat").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 근무조
                oForm.DataSources.UserDataSources.Add("GNMUJO", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("GNMUJO").Specific.DataBind.SetBound(true, "", "GNMUJO");
                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("ShiftDat").Specific.Value.Trim() + "'";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("GNMUJO").Specific, "Y");
                oForm.Items.Item("GNMUJO").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 구분
                oForm.DataSources.UserDataSources.Add("Gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("Y", "근태이상");
                oForm.Items.Item("Gubun").Specific.ValidValues.Add("N", "정상");
                oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // 근태확인완료구분
                oForm.DataSources.UserDataSources.Add("Conf", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("Conf").Specific.ValidValues.Add("%", "전체");
                oForm.Items.Item("Conf").Specific.ValidValues.Add("Y", "확인");
                oForm.Items.Item("Conf").Specific.ValidValues.Add("N", "미확인");
                oForm.Items.Item("Conf").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                // G5제외
                oForm.DataSources.UserDataSources.Add("G5_YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                oForm.Items.Item("G5_YN").Specific.DataBind.SetBound(true, "", "G5_YN");
                oForm.Items.Item("G5_YN").Specific.Checked = false;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PH_PY575_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        ///  <summary>
        ///  화면의 아이템 Enable 설정
        ///  </summary>
        public void PH_PY575_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY575_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '1' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
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
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
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
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value + "'";
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
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = '2' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.Value + "' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
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
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value + "'";
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
                                sQry = sQry + " AND U_Char3 = '" + oForm.Items.Item("CLTCOD").Specific.Value.Trim() + "'";
                                sQry = sQry + " AND U_Char1 = '" + oForm.Items.Item("RspCode").Specific.Value + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ClsCode").Specific, "Y");
                                break;
                            //근무형태가 바뀌면 근무조 재설정
                            case "ShiftDat":
                                // 삭제
                                if (oForm.Items.Item("GNMUJO").Specific.ValidValues.Count > 0)
                                {
                                    for (i = oForm.Items.Item("GNMUJO").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                    {
                                        oForm.Items.Item("GNMUJO").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                                sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] WHERE Code = 'P155' AND U_UseYN= 'Y' AND U_Char1 = '" + oForm.Items.Item("ShiftDat").Specific.Value + "'";
                                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("GNMUJO").Specific, "Y");
                                
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
                                sQry = "SELECT U_FullName FROM [@PH_PY001A] WHERE Code =  '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
                                oRecordSet.DoQuery(sQry);
                                oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item("U_FullName").Value.ToString().Trim();
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
                        System.Threading.Thread thread = new System.Threading.Thread(PH_PY575_Print_Report01);
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
        private void PH_PY575_Print_Report01()
        {
            string WinTitle = string.Empty;
            string ReportName = string.Empty;

            string CLTCOD = string.Empty;
            string DocDateFr = string.Empty;
            string DocDateTo = string.Empty;
            string TeamCode = string.Empty;
            string RspCode = string.Empty;
            string ClsCode = string.Empty;
            string MSTCOD = string.Empty;
            string ShiftDat = string.Empty;
            string GNMUJO = string.Empty;
            string Gubun = string.Empty;
            string Conf = string.Empty;
            string G5_YN = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim();
            DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.Trim();
            DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.Trim();
            TeamCode = oForm.Items.Item("TeamCode").Specific.Value.Trim();
            RspCode = oForm.Items.Item("RspCode").Specific.Value.Trim();
            ClsCode = oForm.Items.Item("ClsCode").Specific.Value.Trim();
            MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.Trim();
            ShiftDat = oForm.Items.Item("ShiftDat").Specific.Value.Trim();
            GNMUJO = oForm.Items.Item("GNMUJO").Specific.Value.Trim();
            Gubun = oForm.Items.Item("Gubun").Specific.Value.Trim();
            Conf = oForm.Items.Item("Conf").Specific.Value.Trim();

            if (oForm.DataSources.UserDataSources.Item("G5_YN").Value == "Y")
            {
                G5_YN = "Y";
            }
            else
            {
                G5_YN = "N";
            }

            try
            {
                WinTitle = "[PH_PY575] 근태기찰현황";
                ReportName = "PH_PY575_01.rpt";

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();//Parameter List
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); 
                dataPackFormula.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr.Substring(0, 4) + "-" + DocDateFr.Substring(4, 2) + "-" + DocDateFr.Substring(6, 2))); 
                dataPackFormula.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo.Substring(0, 4) + "-" + DocDateTo.Substring(4, 2) + "-" + DocDateTo.Substring(6, 2))); 
                dataPackFormula.Add(new PSH_DataPackClass("@Gubun", Gubun));

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr)); 
                dataPackParameter.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo)); 
                dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode)); 
                dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
                dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
                dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD", MSTCOD));
                dataPackParameter.Add(new PSH_DataPackClass("@ShiftDat", ShiftDat));
                dataPackParameter.Add(new PSH_DataPackClass("@GNMUJO", GNMUJO));
                dataPackParameter.Add(new PSH_DataPackClass("@Gubun", Gubun));
                dataPackParameter.Add(new PSH_DataPackClass("@Conf", Conf));
                dataPackParameter.Add(new PSH_DataPackClass("@G5_YN", G5_YN));

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY575_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
