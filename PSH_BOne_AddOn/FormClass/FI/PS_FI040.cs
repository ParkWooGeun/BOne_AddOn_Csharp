using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 분말부자재비용분석
    /// </summary>
    internal class PS_FI040 : PSH_BaseClass
    {
        private string oFormUniqueID;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI040.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_FI040_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_FI040");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_FI040_CreateItems();
                PS_FI040_ComboBox_Setting();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        private void PS_FI040_CreateItems()
        {
            try
            {
                oForm.DataSources.UserDataSources.Add("StrDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("StrDate").Specific.DataBind.SetBound(true, "", "StrDate");
                oForm.DataSources.UserDataSources.Item("StrDate").Value = DateTime.Now.ToString("yyyyMMdd");

                oForm.DataSources.UserDataSources.Add("EndDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oForm.Items.Item("EndDate").Specific.DataBind.SetBound(true, "", "EndDate");
                oForm.DataSources.UserDataSources.Item("EndDate").Value = DateTime.Now.ToString("yyyyMMdd");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_FI040_ComboBox_Setting()
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sQry = String.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
                oRecordSet01.DoQuery(sQry);
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("0", "전체 사업장");
                while (!oRecordSet01.EoF)
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //상태
                oForm.Items.Item("Status").Specific.ValidValues.Add("A", "전체");
                oForm.Items.Item("Status").Specific.ValidValues.Add("G", "보관");
                oForm.Items.Item("Status").Specific.ValidValues.Add("F", "부도");
                oForm.Items.Item("Status").Specific.ValidValues.Add("P", "입금");
                oForm.Items.Item("Status").Specific.ValidValues.Add("C", "추심");
                oForm.Items.Item("Status").Specific.ValidValues.Add("D", "할인");
                oForm.Items.Item("Status").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("StrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01); //메모리 해제
            }
        }

        /// <summary>
        /// 리포트 조회
        /// </summary>
        [STAThread]
        private void PS_FI040_Print_Report01()
        {
            string WinTitle = String.Empty;
            string ReportName = String.Empty;
            string StrDate = String.Empty;
            string EndDate = String.Empty;
            string FrDate = String.Empty;
            string ToDate = String.Empty;
            string SCardCode = String.Empty;
            string ECardCode = String.Empty;
            string BPLId = String.Empty;
            string Status = String.Empty;
            string IssuePsn = String.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List

                StrDate = oForm.Items.Item("StrDate").Specific.Value.ToString().Trim();
                EndDate = oForm.Items.Item("EndDate").Specific.Value.ToString().Trim();

                FrDate = oForm.Items.Item("FrDate").Specific.Value.ToString().Trim();
                ToDate = oForm.Items.Item("ToDate").Specific.Value.ToString().Trim();
                SCardCode = oForm.Items.Item("SCardCode").Specific.Value.ToString().Trim();
                ECardCode = oForm.Items.Item("ECardCode").Specific.Value.ToString().Trim();
                BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
                Status = oForm.Items.Item("Status").Specific.Selected.Value.ToString().Trim();
                IssuePsn = oForm.Items.Item("IssuePsn").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(StrDate))
                {
                    StrDate = "19000101";
                }

                if (string.IsNullOrEmpty(EndDate))
                {
                    EndDate = "21001231";
                }

                if (string.IsNullOrEmpty(FrDate))
                {
                    FrDate = "19000101";
                }

                if (string.IsNullOrEmpty(ToDate))
                {
                    ToDate = "21001231";
                }

                if (string.IsNullOrEmpty(SCardCode))
                {
                    SCardCode = "1";
                }

                if (string.IsNullOrEmpty(ECardCode))
                {
                    ECardCode = "ZZZZZZZZ";
                }

                if (string.IsNullOrEmpty(IssuePsn))
                {
                    IssuePsn = "%";
                }

                WinTitle = "[PS_FI040] 만기일자별 받을어음 현황";
                ReportName = "PS_FI040_01.RPT";

                dataPackParameter.Add(new PSH_DataPackClass("@StrDate", DateTime.ParseExact(StrDate, "yyyyMMdd", null)));
                dataPackParameter.Add(new PSH_DataPackClass("@EndDate", DateTime.ParseExact(EndDate, "yyyyMMdd", null)));
                dataPackParameter.Add(new PSH_DataPackClass("@FrDate", DateTime.ParseExact(FrDate, "yyyyMMdd", null)));
                dataPackParameter.Add(new PSH_DataPackClass("@ToDate", DateTime.ParseExact(ToDate, "yyyyMMdd", null)));
                dataPackParameter.Add(new PSH_DataPackClass("@SCardCode", SCardCode));
                dataPackParameter.Add(new PSH_DataPackClass("@ECardCode", ECardCode));
                dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
                dataPackParameter.Add(new PSH_DataPackClass("@Status", Status));
                dataPackParameter.Add(new PSH_DataPackClass("@IssuePsn", IssuePsn));

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@StrDate", StrDate == "19000101" ? "All" : codeHelpClass.Left(StrDate, 4) + "-" + codeHelpClass.Mid(StrDate, 4, 2) + "-" + codeHelpClass.Right(StrDate, 2)));
                dataPackFormula.Add(new PSH_DataPackClass("@EndDate", EndDate == "21001231" ? "All" : codeHelpClass.Left(EndDate, 4) + "-" + codeHelpClass.Mid(EndDate, 4, 2) + "-" + codeHelpClass.Right(EndDate, 2)));
                dataPackFormula.Add(new PSH_DataPackClass("@BPLId", dataHelpClass.Get_ReData("BPLName", "BPLId", "OBPL", BPLId, ""))); //사업장
                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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

                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                //case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                //    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

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

                    //case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    //    break;

                    //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
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
            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(PS_FI040_Print_Report01);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "SCardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("SCardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "ECardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("ECardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "IssuePsn")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("IssuePsn").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
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
                }
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
        /// FormMenuEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            int i = 0;
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
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1293": //행삭제
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
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    switch (BusinessObjectInfo.EventType)
                    {
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                            break;
                        case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }
    }
}
