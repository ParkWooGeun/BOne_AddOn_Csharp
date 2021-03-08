using System;
using System.Collections.Generic;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 입사자대장
    /// </summary>
    internal class PH_PY505 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        public override void LoadForm(string oFormDocEntry01)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY505.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue.ToString() + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PH_PY505_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PH_PY505");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //oForm.DataBrowser.BrowseBy = "Code"

                oForm.Freeze(true);
                PH_PY505_CreateItems();
                PH_PY505_EnableMenus();
                PH_PY505_SetDocument(oFormDocEntry01);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("Form_Load Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Update();
                oForm.Freeze(false);
                oForm.Visible = true;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PH_PY505_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = DataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PH_PY505'", "");
                if (Convert.ToDouble(DocEntry) == 0)
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_FormClear_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Validate
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PH_PY505_Validate(string ValidateType)
        {
            bool functionReturnValue = false;
            string errCode = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PH_PY505A] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errCode = "1";
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

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_Validate_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 리포트 출력
        /// </summary>
        [STAThread]
        private void PH_PY505_Print_Report01()
        {
            string WinTitle;
            string ReportName;

            string CLTCOD; //사업장
            string DocDateFr; //등록기간(시작)
            string DocDateTo; //등록기간(종료)
            
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                WinTitle = "[PH_PY505] 입사자대장";
                ReportName = "PH_PY505_01.rpt";

                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Selected.Value.ToString().Trim(); //사업장
                DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.Trim();
                DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.Trim();

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

                //Formula
                dataPackFormula.Add(new PSH_DataPackClass("@CLTCOD", dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", CLTCOD, "and Code = 'P144' AND U_UseYN= 'Y'"))); //사업장
                dataPackFormula.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr.Substring(0, 4) + "-" + DocDateFr.Substring(4, 2) + "-" + DocDateFr.Substring(6, 2))); //시작일
                dataPackFormula.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo.Substring(0, 4) + "-" + DocDateTo.Substring(4, 2) + "-" + DocDateTo.Substring(6, 2))); //종료일

                //Parameter
                dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD)); //사업장
                dataPackParameter.Add(new PSH_DataPackClass("@DocDateFr", DocDateFr)); //등록기간(시작)
                dataPackParameter.Add(new PSH_DataPackClass("@DocDateTo", DocDateTo)); //등록기간(종료)

                formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        /// <returns></returns>
        private void PH_PY505_CreateItems()
        {
            SAPbouiCOM.ComboBox oCombo = null;
            PSH_DataHelpClass DataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);

                // 사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oCombo = oForm.Items.Item("CLTCOD").Specific;
                oCombo.DataBind.SetBound(true, "", "CLTCOD");

                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                //  접속자에 따른 권한별 사업장 콤보박스세팅
                DataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true);

                // 시작일자
                oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 10);                
                oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMMdd");

                // 종료일자
                oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 10);                
                oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
                oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY645_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCombo); //메모리 해제
            }
        }

        /// <summary>
        /// 메뉴 아이콘 Enable
        /// </summary>
        private void PH_PY505_EnableMenus()
        {
            try
            {
                oForm.EnableMenu("1283", true); //제거
                oForm.EnableMenu("1284", false); //취소
                oForm.EnableMenu("1293", true); //행삭제
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_EnableMenus_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면세팅
        /// </summary>
        /// <param name="oFormDocEntry01"></param>
        private void PH_PY505_SetDocument(string oFormDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry01))
                {
                    PH_PY505_FormItemEnabled();
                    //PH_PY505_AddMatrixRow();
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PH_PY505_FormItemEnabled();                    
                    oForm.Items.Item("Code").Specific.Value = oFormDocEntry01;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_SetDocument_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY505_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", false); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //문서찾기
                    oForm.EnableMenu("1282", true); //문서추가
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                        if (pVal.BeforeAction == true)
                        {
                            if (pVal.ItemUID == "Btn01")
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(PH_PY505_Print_Report01);
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start();
                            }
                        }
                        else if (pVal.BeforeAction == false)
                        {
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                        break;
                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                        switch (pVal.ItemUID)
                        {
                            case "Mat01":
                            case "Grid01":
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
                        break;
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                        if (pVal.BeforeAction == true)
                        {
                            switch (pVal.ItemUID)
                            {
                                case "Mat01":
                                case "Grid01":
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
                        else if (pVal.BeforeAction == false)
                        {
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                        if (pVal.BeforeAction == true)
                        {
                        }
                        else if (pVal.BeforeAction == false)
                        {
                            PH_PY505_FormItemEnabled();
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                        if (pVal.BeforeAction == true)
                        {
                        }
                        else if (pVal.BeforeAction == false)
                        {
                            SubMain.Remove_Forms(oFormUniqueID);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST://27
                        break;
                    case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                        break;
                    case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                        break;
                    case SAPbouiCOM.BoEventTypes.et_Drag: //39
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_Raise_ItemEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
            try
            {
                oForm.Freeze(true);

                if (pVal.BeforeAction == true)
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
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1283":
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                            PH_PY505_FormItemEnabled();
                            break;
                        case "1284":
                            break;
                        case "1286":
                            break;
                        case "1281": //문서찾기
                            PH_PY505_FormItemEnabled();
                            oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //문서추가
                            PH_PY505_FormItemEnabled();
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                            PH_PY505_FormItemEnabled();
                            break;
                        case "1293":
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_Raise_MenuEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_Raise_FormDataEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    case "Mat01":
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
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY505_Raise_RightClickEvent_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}