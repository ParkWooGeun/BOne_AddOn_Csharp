using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 마스터승인권한관리
    /// </summary>
    internal class PS_SY010 : PSH_BaseClass
    {
        private string oFormUniqueID;
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.Grid oGrid2;
        public SAPbouiCOM.DataTable oDS_PS_SY010H;
        public SAPbouiCOM.DataTable oDS_PS_SY010L;

        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private SAPbouiCOM.BoFormMode oForm_Mode;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SY010.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_SY010_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_SY010");

                string strXml = null;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                //oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_SY010_CreateItems();
                PS_SY010_ComboBox_Setting();
                PS_SY010_Initialization();

                oForm.EnableMenu(("1281"), false); //찾기
                oForm.EnableMenu(("1282"), false); //추가
                oForm.EnableMenu(("1293"), false); //행삭제
                oForm.EnableMenu(("1283"), false); //삭제
                oForm.EnableMenu(("1286"), false); //닫기
                oForm.EnableMenu(("1287"), false); //복제
                oForm.EnableMenu(("1285"), false); //복원
                oForm.EnableMenu(("1284"), false); //취소
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
        private void PS_SY010_CreateItems()
        {
            try
            {
                oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");

                oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMM01");

                oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");

                oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PS_SY010H");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PS_SY010H");
                oDS_PS_SY010H = oForm.DataSources.DataTables.Item("PS_SY010H");

                oGrid2 = oForm.Items.Item("Grid02").Specific;
                oForm.DataSources.DataTables.Add("PS_SY010L");

                oGrid2.DataTable = oForm.DataSources.DataTables.Item("PS_SY010L");
                oDS_PS_SY010L = oForm.DataSources.DataTables.Item("PS_SY010L");

            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SY010_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                // 사업장
                sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }

                // 모듈
                sQry = "select distinct b.Code, a.name";
                sQry = sQry + " from [@PS_SY005H] a inner join [@PS_SY005L] b on a.Code = b.Code and b.U_UseYN ='Y'";
                sQry = sQry + " Where b.U_AppUser = '" + PSH_Globals.oCompany.UserName + "'";
                sQry = sQry + "   and len(a.Code) <> '2'";

                oRecordSet01.DoQuery(sQry);
                while (!(oRecordSet01.EoF))
                {
                    oForm.Items.Item("Module").Specific.ValidValues.Add(oRecordSet01.Fields.Item(0).Value.ToString().Trim(), oRecordSet01.Fields.Item(1).Value.ToString().Trim());
                    oRecordSet01.MoveNext();
                }
                oForm.Items.Item("Module").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// PS_SY010_Initialization
        /// </summary>
        private void PS_SY010_Initialization()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //아이디별 사업장 세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Update_PurchaseDemand
        /// </summary>
        /// <returns></returns>
        private bool PS_SY010_Update_PurchaseDemand(SAPbouiCOM.ItemEvent pVal)
        {
            bool returnValue = false;
            string errMessage = string.Empty;
            short i;
            string sQry;
            string codeValue;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
                {
                    if (oForm.Items.Item("Module").Specific.Selected.Value == "S150")
                    {

                        for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                        {
                            if (oDS_PS_SY010H.Columns.Item("OKYN").Cells.Item(i).Value == "N")
                            {
                                codeValue = oDS_PS_SY010H.Columns.Item("품목코드").Cells.Item(i).Value;

                                sQry = "EXEC [PS_SY010_03] '" + codeValue + "','" + oForm.Items.Item("Module").Specific.Value + "','" + PSH_Globals.oCompany.UserSignature + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                        }
                        PSH_Globals.SBO_Application.MessageBox("수정완료");
                        oForm.Items.Item("Btn02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else if (oForm.Items.Item("Module").Specific.Selected.Value == "S134")
                    {

                        for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                        {
                            if (oDS_PS_SY010H.Columns.Item("OKYN").Cells.Item(i).Value == "N")
                            {
                                codeValue = oDS_PS_SY010H.Columns.Item("거래처코드").Cells.Item(i).Value;

                                sQry = "EXEC [PS_SY010_03] '" + codeValue + "','" + oForm.Items.Item("Module").Specific.Value + "','" + PSH_Globals.oCompany.UserSignature + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                        }
                        PSH_Globals.SBO_Application.MessageBox("수정완료");
                        oForm.Items.Item("Btn02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else if (oForm.Items.Item("Module").Specific.Selected.Value == "OCRD")
                    {
                        for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                        {
                            if (oDS_PS_SY010H.Columns.Item("OKYN").Cells.Item(i).Value == "N")
                            {
                                codeValue = oDS_PS_SY010H.Columns.Item("거래처코드").Cells.Item(i).Value;
                                sQry = "EXEC [PS_SY010_03] '" + codeValue + "','" + oForm.Items.Item("Module").Specific.Value + "','" + PSH_Globals.oCompany.UserSignature + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                        }
                        PSH_Globals.SBO_Application.MessageBox("수정완료");
                        oForm.Items.Item("Btn02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else if (oForm.Items.Item("Module").Specific.Selected.Value == "CO800")
                    {

                        for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                        {
                            if (oDS_PS_SY010H.Columns.Item("OKYN").Cells.Item(i).Value == "N")
                            {
                                codeValue = oDS_PS_SY010H.Columns.Item("문서번호").Cells.Item(i).Value.ToString().Trim();
                                sQry = "EXEC [PS_SY010_03] '" + codeValue + "','" + oForm.Items.Item("Module").Specific.Value + "','" + PSH_Globals.oCompany.UserSignature + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                        }
                        PSH_Globals.SBO_Application.MessageBox("수정완료");
                        oForm.Items.Item("Btn02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                    else if (oForm.Items.Item("Module").Specific.Selected.Value == "SD030")
                    {

                        for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                        {
                            if (oDS_PS_SY010H.Columns.Item("OKYN").Cells.Item(i).Value == "N")
                            {
                                codeValue = oDS_PS_SY010H.Columns.Item("출하요청문서번호").Cells.Item(i).Value.ToString().Trim();
                                sQry = "EXEC [PS_SY010_03] '" + codeValue + "','" + oForm.Items.Item("Module").Specific.Value + "','" + PSH_Globals.oCompany.UserSignature + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                        }
                        PSH_Globals.SBO_Application.MessageBox("수정완료");
                        oForm.Items.Item("Btn02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }
                else
                {
                    dataHelpClass.MDC_GF_Message("데이터가 존재하지 않습니다.!", "S");
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            return returnValue;
        }

        /// <summary>
        /// LoadData
        /// </summary>
        private void PS_SY010_LoadData()
        {
            int iRow;
            string sQry;
            string Module_Renamed;
            string DocDateFr;
            string DocDateTo;
            string errMessage = string.Empty;

            try
            {
                oForm.Freeze(true);
                DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
                DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
                Module_Renamed = oForm.Items.Item("Module").Specific.Value.ToString().Trim();


                if (string.IsNullOrEmpty(DocDateFr))
                    DocDateFr = DateTime.Now.AddDays(-90).ToString("yyyyMM01").Trim();
                if (string.IsNullOrEmpty(DocDateTo))
                    DocDateTo = DateTime.Now.ToString("yyyyMMdd");

                sQry = "EXEC [PS_SY010_01] '" + DocDateFr + "','" + DocDateTo + "','" + PSH_Globals.oCompany.UserName + "','" + Module_Renamed + "'";

                oDS_PS_SY010H.ExecuteQuery(sQry);

                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

                PS_SY010_TitleSetting(iRow);
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_SY010_TitleSetting(int iRow)
        {
            int i;
            SAPbouiCOM.ComboBoxColumn oComboCol = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                for (i = 0; i < oGrid1.DataTable.Columns.Count - 1; i++)
                {

                    switch (oGrid1.Columns.Item(i).TitleObject.Caption)
                    {
                        case "OKYN":
                            oGrid1.Columns.Item("OKYN").Editable = true;
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            oComboCol = (SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("OKYN");

                            oComboCol.ValidValues.Add("Y", "대상");
                            oComboCol.ValidValues.Add("N", "확인완료");

                            oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                            break;

                        case "거래처코드":
                            oGrid1.Columns.Item("거래처코드").Type = BoGridColumnType.gct_EditText;
                            EditTextColumn col1 = (EditTextColumn)oGrid1.Columns.Item("거래처코드");
                            col1.Editable = false;
                            col1.LinkedObjectType = "2"; // Link to BusinessPartner
                            break;

                        case "품목코드":
                            oGrid1.Columns.Item("품목코드").Type = BoGridColumnType.gct_EditText;
                            EditTextColumn col2 = (EditTextColumn)oGrid1.Columns.Item("품목코드");
                            col2.Editable = false;
                            col2.LinkedObjectType = "4"; // Link to ItemMaster
                            break;

                        default:
                            oGrid1.Columns.Item(oGrid1.Columns.Item(i).TitleObject.Caption).Editable = false;
                            break;
                    }

                }
                if (oGrid1.Columns.Count > 0)
                {
                    oGrid1.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// LoadCaption
        /// </summary>
        private void PS_SY010_LoadCaption()
        {
            try
            {
                if (oForm_Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("Btn01").Specific.Caption = "확인";
                }
                else if (oForm_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                {
                    oForm.Items.Item("Btn01").Specific.Caption = "갱신";
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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

                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                        if (oForm_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SY010_Update_PurchaseDemand(pVal) == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oForm_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_SY010_LoadCaption();
                        }
                        else if (oForm_Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Close();
                        }
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        oDS_PS_SY010L.Clear();
                        PS_SY010_LoadData();

                        oForm_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        PS_SY010_LoadCaption();
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "BPLId")
                    {
                        oDS_PS_SY010H.Clear();
                        oDS_PS_SY010L.Clear();
                    }
                    else if (pVal.ItemUID == "Grid01")
                    {
                        oForm_Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        PS_SY010_LoadCaption();
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Grid01")
                        {
                            oForm_Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            PS_SY010_LoadCaption();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                BubbleEvent = false;
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid2);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SY010H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SY010L);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }

        /// <summary>
        /// Raise_EVENT_DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry = string.Empty;

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Grid01")
                    {
                        if (oForm.Items.Item("Module").Specific.Selected.Value == "S134")
                        {
                            sQry = "EXEC [PS_SY010_02] '" + oGrid1.DataTable.Columns.Item("거래처코드").Cells.Item(pVal.Row).Value + "','" + oForm.Items.Item("Module").Specific.Value + "'";
                            oDS_PS_SY010L.ExecuteQuery(sQry);
                            oGrid2.AutoResizeColumns();
                        }
                        else if (oForm.Items.Item("Module").Specific.Selected.Value == "S150")
                        {
                            sQry = "EXEC [PS_SY010_02] '" + oGrid1.DataTable.Columns.Item("품목코드").Cells.Item(pVal.Row).Value + "','" + oForm.Items.Item("Module").Specific.Value + "'";
                            oDS_PS_SY010L.ExecuteQuery(sQry);
                            oGrid2.AutoResizeColumns();
                        }
                        else if (oForm.Items.Item("Module").Specific.Selected.Value == "OCRD")
                        {
                            sQry = "EXEC [PS_SY010_02] '" + oGrid1.DataTable.Columns.Item("거래처코드").Cells.Item(pVal.Row).Value + "','" + oForm.Items.Item("Module").Specific.Value + "'";
                            oDS_PS_SY010L.ExecuteQuery(sQry);
                            oGrid2.AutoResizeColumns();
                        }
                        else if (oForm.Items.Item("Module").Specific.Selected.Value == "CO800")
                        {
                            sQry = "EXEC [PS_SY010_02] '" + oGrid1.DataTable.Columns.Item("문서번호").Cells.Item(pVal.Row).Value + "','" + oForm.Items.Item("Module").Specific.Value + "'";
                            oDS_PS_SY010L.ExecuteQuery(sQry);
                            oGrid2.AutoResizeColumns();
                        }
                        else if (oForm.Items.Item("Module").Specific.Selected.Value == "SD030")
                        {
                            sQry = "EXEC [PS_SY010_02] '" + oGrid1.DataTable.Columns.Item("출하요청문서번호").Cells.Item(pVal.Row).Value + "','" + oForm.Items.Item("Module").Specific.Value + "'";
                            oDS_PS_SY010L.ExecuteQuery(sQry);
                            oGrid2.AutoResizeColumns();
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
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
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(다음)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(최초)
                        case "1291": //레코드이동(최종)
                        case "1287": //복제
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
            }
        }
    }
}
