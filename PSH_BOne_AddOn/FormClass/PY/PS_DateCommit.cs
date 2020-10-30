using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 날짜 변경 승인
    /// </summary>
    internal class PS_DateCommit : PSH_BaseClass
    {
        public string oFormUniqueID01;

        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PS_DateCommit;

        private string oLastItemUID;
        private string oLastColUID;
        private int oLastColRow;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFormDocEntry01)
        {
            int i;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_DateCommit.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PS_DateCommit_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PS_DateCommit");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

                oForm.Freeze(true);
                PS_DateCommit_CreateItems();
                PS_DateCommit_FormItemEnabled();
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
        private void PS_DateCommit_CreateItems()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PS_DateCommit");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PS_DateCommit");
                oDS_PS_DateCommit = oForm.DataSources.DataTables.Item("PS_DateCommit");

                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("일자", SAPbouiCOM.BoFieldsType.ft_Date);
                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("요일", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("근태구분", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("부서", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("담당", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("위해일수", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                //oForm.DataSources.DataTables.Item("PS_DateCommit").Columns.Add("위해코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                ////사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                // ObjectCode
                sQry = "select Code, Name from [@PS_SY005H] where len(Code) in (2,3)";
                dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("ObjectCode").Specific, "Y");
                oForm.Items.Item("ObjectCode").DisplayDesc = true;
                oForm.Items.Item("ObjectCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");
                oForm.Items.Item("FrDt").Specific.Value = DateTime.Now.ToString("yyyyMM01");

                oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE, 8);
                oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");
                oForm.Items.Item("ToDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

                oForm.DataSources.UserDataSources.Add("Comments", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
                oForm.Items.Item("Comments").Specific.DataBind.SetBound(true, "", "Comments");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_DateCommit_CreateItems_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        public void PS_DateCommit_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PS_DateCommit_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "Btn_Serch")
                    {
                        if (PS_DateCommit_DataValidCheck() == true)
                        {
                            PS_DateCommit_DataFind();
                        }
                        else
                        {
                            BubbleEvent = false;
                        }
                    }
                    if (pVal.ItemUID == "Btn_Save")
                    {
                        if (PS_DateCommit_DataSave() == false)
                        {
                            BubbleEvent = false;
                        }
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
        /// DataValidCheck
        /// </summary>
        /// <returns></returns>
        public bool PS_DateCommit_DataValidCheck()
        {
            bool functionReturnValue = false;
            
            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim()))
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    return functionReturnValue;
                }

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_DateCommit_DataValidCheck_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
            }

            return functionReturnValue;
        }

        /// <summary>
        /// 데이터 조회
        /// </summary>
        /// <returns></returns>
        private void PS_DateCommit_DataFind()
        {
            int iRow;
            string sQry;
            string CLTCODE;
            string Grantor;
            string ObjectCode;
            string FrDate;
            string ToDate;

            CLTCODE = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
            Grantor = PSH_Globals.oCompany.UserName;
            ObjectCode = oForm.Items.Item("ObjectCode").Specific.Value.Trim();
            FrDate = oForm.Items.Item("FrDt").Specific.Value.Trim();
            ToDate = oForm.Items.Item("ToDt").Specific.Value.Trim();

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                sQry = "Exec PS_DateCommit_01 '" + CLTCODE + "','" + Grantor + "',";
                sQry = sQry + "'" + ObjectCode + "','" + FrDate + "','" + ToDate + "'";
                oDS_PS_DateCommit.ExecuteQuery(sQry);

                iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

                PS_DateCommit_TitleSetting();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_DateCommit_DataFind_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }

        /// <summary>
        /// DataSave
        /// </summary>
        /// <returns></returns>
        private bool PS_DateCommit_DataSave()
        {
            bool functionReturnValue = false;
            int i;
            int ErrNum = 0;
            string sQry;
            string CLTCOD;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                functionReturnValue = false;
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.Trim();
                if (PSH_Globals.SBO_Application.MessageBox("저장하시겠습니까?", 2, "Yes", "No") == 2)
                {
                    ErrNum = 1;
                    throw new Exception();
                }
                if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
                {
                    for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                    {
                        if (oDS_PS_DateCommit.Columns.Item("OKYN").Cells.Item(i).Value != "N")
                        {
                            sQry = "UPDATE PSH_DateChange SET OKYN = '" + oDS_PS_DateCommit.Columns.Item("OKYN").Cells.Item(i).Value + "', ApprDate = convert(char(8), GETDATE(), 112)";
                            sQry += " where ObjectCode ='" + oForm.Items.Item("ObjectCode").Specific.Value.Trim() + "'";
                            sQry += "  and OKYN = 'N'";
                            sQry += "  and DocEntry = '" + oDS_PS_DateCommit.Columns.Item("DocEntry").Cells.Item(i).Value + "'";
                            sQry += "  and LineId ="  + oDS_PS_DateCommit.Columns.Item("LineId").Cells.Item(i).Value;

                            oRecordSet.DoQuery(sQry);

                            if (oDS_PS_DateCommit.Columns.Item("OKYN").Cells.Item(i).Value == "Y")
                            {
                                sQry = "EXEC [PS_DateCommit_02] '" + oForm.Items.Item("ObjectCode").Specific.Value.Trim() + "'";
                                sQry += ",'" + oDS_PS_DateCommit.Columns.Item("DocEntry").Cells.Item(i).Value + "'";
                                sQry += ",'" + oDS_PS_DateCommit.Columns.Item("LineId").Cells.Item(i).Value + "'";
                                sQry += ",'" + oDS_PS_DateCommit.Columns.Item("DocDate").Cells.Item(i).Value + "'";
                                sQry += ",'" + oDS_PS_DateCommit.Columns.Item("DueDate").Cells.Item(i).Value + "'";
                                sQry += ",'" + oDS_PS_DateCommit.Columns.Item("TaxDate").Cells.Item(i).Value + "'";
                                oRecordSet.DoQuery(sQry);
                            }
                        }
                        
                    }
                    PSH_Globals.SBO_Application.MessageBox("저장되었습니다.");
                    functionReturnValue = true;
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox("데이터가 존재하지 않습니다.");
                }
            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.MessageBox("등록 취소되었습니다.");
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("PS_DateCommit_DataSave_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
            return functionReturnValue;
        }

        /// <summary>
        /// 그리드 타이블 변경
        /// </summary>
        /// <returns></returns>
        private void PS_DateCommit_TitleSetting()
        {
            int i;

            string[] COLNAM = new string[9];

            SAPbouiCOM.ComboBoxColumn oComboCol = null;
            try
            {
                oForm.Freeze(true);

                COLNAM[0] = "구분";
                COLNAM[1] = "등록일자";
                COLNAM[2] = "문서번호";
                COLNAM[3] = "라인번호";
                COLNAM[4] = "등록자";
                COLNAM[5] = "전기일";
                COLNAM[6] = "만기일";
                COLNAM[7] = "증빙일";
                COLNAM[8] = "OKYN";

                for (i = 0; i <= (COLNAM.Length - 1); i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    switch (COLNAM[i])
                    {
                        case "OKYN":
                            oGrid1.Columns.Item(i).Editable = true;
                            oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                            oComboCol = (SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("OKYN");

                            oComboCol.ValidValues.Add("N", "대상"); 
                            oComboCol.ValidValues.Add("Y", "승인");
                            oComboCol.ValidValues.Add("C", "반려");
                            oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;

                            break;
                        default:

                            oGrid1.Columns.Item(i).Editable = false;
                            break;
                    }
                }
                oGrid1.AutoResizeColumns();
            }
            
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.SetStatusBarMessage("PS_DateCommit_TitleSetting_Error:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oComboCol);
            }
        }

        /// <summary>
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Grid01":
                            if (pVal.Row >= 0)
                            {
                                switch (pVal.ItemUID)
                                {
                                    case "Grid01":
                                        PS_DateCommit_Comments(pVal.Row);
                                        break;
                                }
                            }
                            break;
                    }

                    switch (pVal.ItemUID)
                    {
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
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_CLICK_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// PS_DateChange_MTX02
        /// </summary>
        private void PS_DateCommit_Comments(int oRow)
        {
            int sRow;
            int ErrNum = 0;
            string sQry;
            string BPLId;
            string CreateUser;
            string ObjectCode;
            string Grantor;
            int LineId;
            int DocEntry;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                sRow = oRow;
                BPLId = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                CreateUser = oDS_PS_DateCommit.Columns.Item("CreateUser").Cells.Item(oRow).Value;
                DocEntry = oDS_PS_DateCommit.Columns.Item("DocEntry").Cells.Item(oRow).Value;
                Grantor = PSH_Globals.oCompany.UserName;
                ObjectCode = oForm.Items.Item("ObjectCode").Specific.Value.ToString().Trim(); 
                LineId = oDS_PS_DateCommit.Columns.Item("LineId").Cells.Item(oRow).Value;

                sQry = "SELECT  Comments ";
                sQry += "FROM PSH_DateChange ";
                sQry += "WHERE ObjectCode = '"+ ObjectCode + "' ";
                sQry += "AND BPLId = '" + BPLId + "' ";
                sQry += "AND Grantor = '" + Grantor + "' ";
                sQry += "AND DocEntry = '" + DocEntry + "'  ";
                sQry += "AND LineId = '" + LineId + "'  ";
                sQry += "AND CreateUser = '" + CreateUser + "' ";

                oRecordSet.DoQuery(sQry);

                oForm.Items.Item("Comments").Specific.Value = oRecordSet.Fields.Item("Comments").Value;

            }
            catch (Exception ex)
            {
                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PS_DateChange_MTX02_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            string sQry;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {

                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        switch (pVal.ItemUID)
                        {
                            case "MSTCOD":
                                sQry = "SELECT U_FullName from [@PH_PY001A] Where Code = '" + oForm.Items.Item("MSTCOD").Specific.Value + "'";
                                oRecordSet.DoQuery(sQry);
                                if (oRecordSet.RecordCount > 0)
                                {
                                    oForm.Items.Item("FullName").Specific.Value = oRecordSet.Fields.Item(0).Value;

                                }
                                break;

                            case "Grid01":
                                switch (oForm.Items.Item("CLTCOD").Specific.Value.Trim())
                                {
                                    case "1":
                                        if (Convert.ToDouble(oDS_PS_DateCommit.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 0 & Convert.ToDouble(oDS_PS_DateCommit.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 1)
                                        {
                                            oDS_PS_DateCommit.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value = 0;
                                            PSH_Globals.SBO_Application.SetStatusBarMessage("0 또는 1만 입력 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                        }
                                        oDS_PS_DateCommit.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
                                        break;
                                    case "2":
                                        if (Convert.ToDouble(oDS_PS_DateCommit.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 0.5 & Convert.ToDouble(oDS_PS_DateCommit.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) != 1)
                                        {
                                            oDS_PS_DateCommit.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value = 0;
                                            oDS_PS_DateCommit.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
                                            PSH_Globals.SBO_Application.SetStatusBarMessage("0.5 또는 1 만 입력 가능합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                        }
                                        else
                                        {
                                            if (Convert.ToDouble(oDS_PS_DateCommit.Columns.Item("DangerNu").Cells.Item(pVal.Row).Value) >= 0.5)
                                            {
                                                oDS_PS_DateCommit.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "56";
                                                //// '// 위해등급 6급
                                            }
                                            else
                                            {
                                                oDS_PS_DateCommit.Columns.Item("DangerCD").Cells.Item(pVal.Row).Value = "";
                                                ////"56" '// 위해등급 6급
                                            }
                                        }
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_VALIDATE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid1);
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
    }
}