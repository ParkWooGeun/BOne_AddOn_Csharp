using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using Microsoft.VisualBasic;


namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 기본업무등록
    /// </summary>
    internal class PH_PY016 : PSH_BaseClass
    {
        public string oFormUniqueID01;

        //'// 그리드 사용시
        public SAPbouiCOM.Grid oGrid1;
        public SAPbouiCOM.DataTable oDS_PH_PY016;

        /// <summary>
        /// 화면 호출
        /// </summary>
        public override void LoadForm(string oFromDocEntry01)
        {
            int i = 0;
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();
            try
            {
                oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PH_PY016.srf");
                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID01 = "PH_PY016_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID01, "PH_PY016");

                string strXml = string.Empty;
                strXml = oXmlDoc.xml.ToString();

                PSH_Globals.SBO_Application.LoadBatchActions(strXml);
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                oForm.Freeze(true);
                PH_PY016_CreateItems();
                PH_PY016_FormItemEnabled();
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
        private void PH_PY016_CreateItems()
        {
            string sQry = string.Empty;
            try
            {
                oForm.Freeze(true);

                oGrid1 = oForm.Items.Item("Grid01").Specific;
                oForm.DataSources.DataTables.Add("PH_PY016");

                oGrid1.DataTable = oForm.DataSources.DataTables.Item("PH_PY016");
                oDS_PH_PY016 = oForm.DataSources.DataTables.Item("PH_PY016");

                //그리드 타이틀 
                oForm.DataSources.DataTables.Item("PH_PY016").Columns.Add("부서", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY016").Columns.Add("담당", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY016").Columns.Add("사번", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY016").Columns.Add("성명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY016").Columns.Add("위해코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY016").Columns.Add("위해명", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY016").Columns.Add("업무코드", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                oForm.DataSources.DataTables.Item("PH_PY016").Columns.Add("기본업무", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);

                //사업장
                oForm.DataSources.UserDataSources.Add("CLTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CLTCOD").Specific.DataBind.SetBound(true, "", "CLTCOD");
                oForm.Items.Item("CLTCOD").DisplayDesc = true;

                ////부서
                oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");
                oForm.Items.Item("TeamCode").DisplayDesc = true;
                //
                //// 담당
                oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");
                oForm.Items.Item("RspCode").DisplayDesc = true;

                
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY016_CreateItems_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PH_PY016_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                dataHelpClass.CLTCOD_Select(oForm, "CLTCOD", true); // 접속자에 따른 권한별 사업장 콤보박스세팅
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText("PPH_PY016_FormItemEnabled_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PH_PY016);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                //    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                //    break;

                ////case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                //    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

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
                    if (pVal.ItemUID == "Btn_Serch")
                    {
                        PH_PY016_DataFind();
                    }
                    if (pVal.ItemUID == "Btn_Save")
                    {
                        PH_PY016_DataSave();
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
                if (pVal.BeforeAction == true)
                {

                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "CLTCOD")
                        {
                            ////기본사항 - 부서 (사업장에 따른 부서변경)

                            if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "' And U_UseYN = 'Y'";
                            sQry = sQry + " ORDER BY U_Seq";

                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("TeamCode").Specific, "N");

                            oForm.Items.Item("TeamCode").DisplayDesc = true;
                        }
                        if (pVal.ItemUID == "TeamCode")
                        {
                            ////담당 (사업장에 따른 담당변경)
                            if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
                            {
                                for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
                                {
                                    oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                }
                            }

                            sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                            sQry = sQry + " WHERE Code = '2' AND U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE.ToString().Trim() + "' And U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.VALUE.ToString().Trim() + "' And U_UseYN = 'Y'";
                            sQry = sQry + " Order By U_Seq";
                            dataHelpClass.SetReDataCombo(oForm, sQry, oForm.Items.Item("RspCode").Specific, "N");
                            oForm.Items.Item("RspCode").DisplayDesc = true;
                        }
                        if (pVal.ItemUID == "Grid01")
                        {

                        }
                    }
                }
                if (oGrid1.Columns.Count > 0)
                {
                    oGrid1.AutoResizeColumns();
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

        private void PH_PY016_DataFind()
            {
            short ErrNum = 0;
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string TeamCode = string.Empty;
            string RspCode = string.Empty;
            int j, i = 0;
            string[] COLNAM = new string[8];

            SAPbouiCOM.ComboBoxColumn oComboCol = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                oForm.Freeze(true);
                CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
                TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
                RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(CLTCOD))
                {
                    ErrNum = 1;
                    throw new Exception();
                }

                sQry = "Exec PH_PY016 '" + CLTCOD + "','" + TeamCode + "',";
                sQry = sQry + "'" + RspCode + "'";
                oDS_PH_PY016.ExecuteQuery(sQry);

                //그리드 콤보박스
                COLNAM[0] = "부서";
                COLNAM[1] = "담당";
                COLNAM[2] = "사번";
                COLNAM[3] = "성명";
                COLNAM[4] = "위해코드";
                COLNAM[5] = "위해명";
                COLNAM[6] = "업무코드";
                COLNAM[7] = "기본업무";

                for (i = 0; i <= Information.UBound(COLNAM); i++)
                {
                    oGrid1.Columns.Item(i).TitleObject.Caption = COLNAM[i];
                    if (i >= 0 & i < Information.UBound(COLNAM))
                    {
                        oGrid1.Columns.Item(i).Editable = false;
                    }
                    else if (i == Information.UBound(COLNAM))
                    {
                        oGrid1.Columns.Item(i).Editable = true;
                        oGrid1.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
                        oComboCol = (SAPbouiCOM.ComboBoxColumn)oGrid1.Columns.Item("Activity");

                        sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
                        sQry = sQry + " WHERE U_Char2 = '" + oForm.Items.Item("CLTCOD").Specific.VALUE + "' And Code = 'P127' AND U_UseYN= 'Y' Order by U_Seq";
                        oRecordSet.DoQuery(sQry);
                        if (oRecordSet.RecordCount > 0)
                        {
                            oComboCol.ValidValues.Add("", "");
                            for (j = 0; j <= oRecordSet.RecordCount - 1; j++)
                            {
                                oComboCol.ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
                                oRecordSet.MoveNext();
                            }
                        }
                        oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
                    }
                }
            }
            catch (Exception ex)
            {

                if (ErrNum == 1)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장이 없습니다. 확인바랍니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY016_DataFind_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                oGrid1.AutoResizeColumns();
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oComboCol);
            }
         }

        private void PH_PY016_DataSave()
        {
            string sQry = string.Empty;
            string CLTCOD = string.Empty;
            string TeamCode = string.Empty;
            string RspCode = string.Empty;
            int i = 0;

            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            try
            {
                if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
                {
                    for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
                    {
                        sQry = " UPDATE [@PH_PY001A] SET U_Activity = '" + oDS_PH_PY016.Columns.Item("Activity").Cells.Item(i).Value + "'";
                        sQry = sQry + ",                 U_UpdtProg = 'PH_PY016_Y'";
                        sQry = sQry + ",                 U_UserSign2 = '" + PSH_Globals.oCompany.UserSignature.ToString() + "'";
                        sQry = sQry + " WHERE Code = '" + oDS_PH_PY016.Columns.Item("MSTCOD").Cells.Item(i).Value + "'";
                        oRecordSet.DoQuery(sQry);
                    }
                    PSH_Globals.SBO_Application.SetStatusBarMessage("기본업무가 변경되었습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
                else
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("데이터가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
            }
            catch (Exception ex)
            {

                PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY016_DataSave_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);

            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }
        }
    }
}



//		public void Raise_FormItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//		{
//			string sQry = null;
//			int i = 0;
//			SAPbouiCOM.ComboBox oCombo = null;
//			SAPbobsCOM.Recordset oRecordSet = null;

//			 // ERROR: Not supported in C#: OnErrorStatement


//			oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//			switch (pVal.EventType) {
//				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//					////1

//					if (pVal.BeforeAction == true) {
//						if (pVal.ItemUID == "Btn_Serch") {
//							if (PH_PY016_DataValidCheck() == true) {
//								PH_PY016_DataFind();
//							} else {
//								BubbleEvent = false;
//							}
//						}
//						if (pVal.ItemUID == "Btn_Save") {
//							if (PH_PY016_DataSave() == false) {
//								BubbleEvent = false;
//							}
//						}
//					} else if (pVal.BeforeAction == false) {
//						//                If oForm.Mode = fm_ADD_MODE Then
//						//                    If pVal.ActionSuccess = True Then
//						//                        Call PH_PY016_FormItemEnabled
//						//                    End If
//						//                ElseIf oForm.Mode = fm_UPDATE_MODE Then
//						//                    If pVal.ActionSuccess = True Then
//						//                        Call PH_PY016_FormItemEnabled
//						//                    End If
//						//                ElseIf oForm.Mode = fm_OK_MODE Then
//						//                    If pVal.ActionSuccess = True Then
//						//                        Call PH_PY016_FormItemEnabled
//						//                    End If
//						//                End If

//					}
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//					////2
//					break;

//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//					////3
//					switch (pVal.ItemUID) {
//						case "Grid01":
//							if (pVal.Row > 0) {
//								oLastItemUID = pVal.ItemUID;
//								oLastColUID = pVal.ColUID;
//								oLastColRow = pVal.Row;
//							}
//							break;
//						default:
//							oLastItemUID = pVal.ItemUID;
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
//					if (pVal.BeforeAction == true) {

//					} else if (pVal.BeforeAction == false) {
//						if (pVal.ItemChanged == true) {
//							if (pVal.ItemUID == "CLTCOD") {
//								////기본사항 - 부서 (사업장에 따른 부서변경)
//								oCombo = oForm.Items.Item("TeamCode").Specific;

//								if (oCombo.ValidValues.Count > 0) {
//									for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//										oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//									}
//								}

//								sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//								//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sQry = sQry + " WHERE Code = '1' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "' And U_UseYN = 'Y'";
//								sQry = sQry + " ORDER BY U_Seq";
//								MDC_SetMod.SetReDataCombo(oForm, sQry, oCombo);

//								oForm.Items.Item("TeamCode").DisplayDesc = true;



//							}
//							if (pVal.ItemUID == "TeamCode") {
//								////담당 (사업장에 따른 담당변경)

//								oCombo = oForm.Items.Item("RspCode").Specific;

//								if (oCombo.ValidValues.Count > 0) {
//									for (i = oCombo.ValidValues.Count - 1; i >= 0; i += -1) {
//										oCombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
//									}
//								}

//								sQry = "SELECT U_Code, U_CodeNm FROM [@PS_HR200L] ";
//								//UPGRADE_WARNING: oForm.Items.Item(TeamCode).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								//UPGRADE_WARNING: oForm.Items.Item().Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//								sQry = sQry + " WHERE Code = '2' AND U_Char2 = '" + Strings.Trim(oForm.Items.Item("CLTCOD").Specific.VALUE) + "' And U_Char1 = '" + oForm.Items.Item("TeamCode").Specific.VALUE + "' And U_UseYN = 'Y'";
//								sQry = sQry + " Order By U_Seq";
//								MDC_SetMod.SetReDataCombo(oForm, sQry, ref oCombo, ref "Y");
//								oForm.Items.Item("RspCode").DisplayDesc = true;
//							}
//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_CLICK:
//					////6
//					if (pVal.BeforeAction == true) {
//						switch (pVal.ItemUID) {
//							case "Grid01":
//								if (pVal.Row > 0) {
//									//                        Call oGrid1.SelectRow(pVal.Row, True, False)

//								}
//								break;
//						}

//						switch (pVal.ItemUID) {
//							case "Grid01":
//								if (pVal.Row > 0) {
//									oLastItemUID = pVal.ItemUID;
//									oLastColUID = pVal.ColUID;
//									oLastColRow = pVal.Row;
//								}
//								break;
//							default:
//								oLastItemUID = pVal.ItemUID;
//								oLastColUID = "";
//								oLastColRow = 0;
//								break;
//						}
//					} else if (pVal.BeforeAction == false) {

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
//					if (pVal.BeforeAction == true) {

//					} else if (pVal.BeforeAction == false) {
//						if (pVal.ItemChanged == true) {

//						}
//					}
//					oForm.Freeze(false);
//					break;
//				//----------------------------------------------------------
//				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//					////11
//					if (pVal.BeforeAction == true) {
//					} else if (pVal.BeforeAction == false) {
//						PH_PY016_FormItemEnabled();
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
//					if (pVal.BeforeAction == true) {
//					} else if (pVal.BeforeAction == false) {
//						SubMain.RemoveForms(oFormUniqueID);
//						//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oForm = null;
//						//UPGRADE_NOTE: oDS_PH_PY016 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oDS_PH_PY016 = null;
//						//UPGRADE_NOTE: oGrid1 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//						oGrid1 = null;

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
//					if (pVal.BeforeAction == true) {

//					} else if (pVal.BeforeAction == false) {

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
//					if (pVal.BeforeAction == true) {
//					} else if (pVal.Before_Action == false) {
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
//			MDC_Globals.SBO_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}


//		public void Raise_FormMenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//		{
//			int i = 0;
//			 // ERROR: Not supported in C#: OnErrorStatement

//			oForm.Freeze(true);

//			if ((pVal.BeforeAction == true)) {
//				switch (pVal.MenuUID) {
//					case "1283":
//						if (MDC_Globals.SBO_Application.MessageBox("현재 화면내용전체를 제거 하시겠습니까? 복구할 수 없습니다.", 2, "Yes", "No") == 2) {
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
//			} else if ((pVal.BeforeAction == false)) {
//				switch (pVal.MenuUID) {
//					case "1283":
//						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//						PH_PY016_FormItemEnabled();
//						break;

//					case "1284":
//						break;
//					case "1286":
//						break;
//					//            Case "1293":
//					//                Call Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent)
//					case "1281":
//						////문서찾기
//						PH_PY016_FormItemEnabled();
//						break;
//					case "1282":
//						////문서추가
//						PH_PY016_FormItemEnabled();
//						break;

//					case "1288":
//					case "1289":
//					case "1290":
//					case "1291":
//						PH_PY016_FormItemEnabled();
//						break;
//					case "1293":
//						//// 행삭제
//						break;

//				}
//			}
//			oForm.Freeze(false);
//			return;
//			Raise_FormMenuEvent_Error:
//			oForm.Freeze(false);
//			MDC_Globals.SBO_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
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


//			MDC_Globals.SBO_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);

//		}

//		public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//		{

//			 // ERROR: Not supported in C#: OnErrorStatement


//			if (pVal.BeforeAction == true) {
//			} else if (pVal.BeforeAction == false) {
//			}
//			switch (pVal.ItemUID) {
//				case "Grid01":
//					if (pVal.Row > 0) {
//						oLastItemUID = pVal.ItemUID;
//						oLastColUID = pVal.ColUID;
//						oLastColRow = pVal.Row;
//					}
//					break;
//				default:
//					oLastItemUID = pVal.ItemUID;
//					oLastColUID = "";
//					oLastColRow = 0;
//					break;
//			}
//			return;
//			Raise_RightClickEvent_Error:

//			MDC_Globals.SBO_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//		}



//public bool PH_PY016_DataValidCheck()
//{
//    bool functionReturnValue = false;
//    // ERROR: Not supported in C#: OnErrorStatement

//    functionReturnValue = false;
//    int i = 0;
//    string sQry = null;
//    SAPbobsCOM.Recordset oRecordSet = null;

//    oRecordSet = MDC_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//    //UPGRADE_WARNING: oForm.Items(CLTCOD).Specific.VALUE 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//    if (string.IsNullOrEmpty(oForm.Items.Item("CLTCOD").Specific.VALUE))
//    {
//        MDC_Globals.SBO_Application.SetStatusBarMessage("사업장은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        oForm.Items.Item("CLTCOD").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        functionReturnValue = false;
//        return functionReturnValue;
//    }

//    functionReturnValue = true;
//    return functionReturnValue;



//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			PH_PY016_DataValidCheck_Error:


//			//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//			oRecordSet = null;
//			functionReturnValue = false;
//			MDC_Globals.SBO_Application.SetStatusBarMessage("PH_PY016_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//			return functionReturnValue;
//		}
