﻿using System;

using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 통합재무제표용 계정 관리
	/// </summary>
	internal class PS_CO658 : PSH_BaseClass
	{
		public string oFormUniqueID;
		//public SAPbouiCOM.Form oForm;
		public SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_CO658H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_CO658L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private string oDocType01;
		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_CO658.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_CO658_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_CO658");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";
				
				oForm.Freeze(true);
                PS_CO658_CreateItems();
                PS_CO658_ComboBox_Setting();
                PS_CO658_CF_ChooseFromList();
                PS_CO658_EnableMenus();
                PS_CO658_SetDocument(oFromDocEntry01);
                PS_CO658_FormResize();

                oForm.EnableMenu("1283", true); //삭제
				oForm.EnableMenu("1287", true); //복제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", true); //행삭제
			}
			catch(Exception ex)
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
        private void PS_CO658_CreateItems()
        {
            try
            {
                oDS_PS_CO658H = oForm.DataSources.DBDataSources.Item("@PS_CO658H");
                oDS_PS_CO658L = oForm.DataSources.DBDataSources.Item("@PS_CO658L");
                oMat01 = oForm.Items.Item("Mat01").Specific;

                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_CO658_ComboBox_Setting()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PS_CO658", "Mat01", "UseYN", "Y", "Y");
                dataHelpClass.Combo_ValidValues_Insert("PS_CO658", "Mat01", "UseYN", "N", "N");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("UseYN"), "PS_CO658", "Mat01", "UseYN", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "0", "계산");
                dataHelpClass.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "1", "차변잔액");
                dataHelpClass.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "2", "차변합계");
                dataHelpClass.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "3", "대변잔액");
                dataHelpClass.Combo_ValidValues_Insert("PS_CO658", "Mat01", "CRStd", "4", "대변합계");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("CRStd"), "PS_CO658", "Mat01", "CRStd", false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList
        /// </summary>
        private void PS_CO658_CF_ChooseFromList()
        {
            try
            {

            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_CO658_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFromDocEntry01">DocEntry</param>
        private void PS_CO658_SetDocument(string oFromDocEntry01)
        {
            try
            {
                if (string.IsNullOrEmpty(oFromDocEntry01))
                {
                    PS_CO658_FormItemEnabled();
                    PS_CO658_AddMatrixRow(0, true);
                    oForm.Items.Item("Code").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    //oForm.Mode = fm_FIND_MODE;
                    //PS_CO658_FormItemEnabled();
                    //oForm.Items("DocEntry").Specific.VALUE = oFromDocEntry01;
                    //oForm.Items("1").Click(ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_CO658_FormResize()
        {
            try
            {
                oMat01.AutoResizeColumns();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 모드에 따른 아이템 설정
        /// </summary>
        private void PS_CO658_FormItemEnabled()
        {
            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("Code").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    PS_CO658_FormClear();
                    
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("Code").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;

                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Items.Item("Code").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;

                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                }

                oMat01.AutoResizeColumns();
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
        /// 
        /// </summary>
        /// <param name="oRow">행 번호</param>
        /// <param name="RowIserted">행 추가 여부</param>
        private void PS_CO658_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                
                if (RowIserted == false) //행추가여부
                {
                    oDS_PS_CO658L.InsertRecord(oRow);
                }

                oMat01.AddRow();
                oDS_PS_CO658L.Offset = oRow;
                oDS_PS_CO658L.SetValue("U_LineNum", oRow, (oRow + 1).ToString());
                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        
        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_CO658_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_CO658'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.VALUE = DocEntry;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수 사항 check
        /// </summary>
        /// <returns></returns>
        private bool PS_CO658_DataValidCheck()
        {
            bool functionReturnValue = false;
            int i = 0;
            string errCode = string.Empty;

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_CO658_FormClear();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("Code").Specific.VALUE)) //구분코드 미입력
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("Name").Specific.VALUE)) //구분명 미입력
                {
                    errCode = "2";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1) //라인정보 미입력
                {
                    errCode = "3";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("AcctCode").Cells.Item(i).Specific.VALUE))
                    {
                        errCode = "4";
                        throw new Exception();
                    }

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("Contents").Cells.Item(i).Specific.VALUE))
                    {
                        errCode = "5";
                        throw new Exception();
                    }
                }

                oMat01.FlushToDataSource();
                oDS_PS_CO658L.RemoveRecord(oDS_PS_CO658L.Size - 1);
                oMat01.LoadFromDataSource();

                functionReturnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("구분코드가 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("구분명이 입력되지 않았습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("계정코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oMat01.Columns.Item("AcctCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.SetStatusBarMessage("목차제목은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                    oMat01.Columns.Item("Contents").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {

            }

            return functionReturnValue;
        }

        /// <summary>
        /// FlushToItemValue(사용자의 Event에 따른 화면 Item의 유동적인 세팅)
        /// </summary>
        /// <param name="oUID"></param>
        /// <param name="oRow"></param>
        /// <param name="oCol"></param>
        private void PS_CO658_FlushToItemValue(string oUID, int oRow, string oCol)
        {
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                switch (oUID)
                {
                    case "Mat01":
                        if (oCol == "AcctCode")
                        {
                            oMat01.FlushToDataSource();
                            oDS_PS_CO658L.SetValue("U_" + oCol, oRow - 1, oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE);
                            oDS_PS_CO658L.SetValue("U_AcctName", oRow - 1, dataHelpClass.Get_ReData("AcctName", "AcctCode", "[OACT]", "'" + oMat01.Columns.Item(oCol).Cells.Item(oRow).Specific.VALUE + "'", ""));

                            if (oMat01.RowCount == oRow && !string.IsNullOrEmpty(oDS_PS_CO658L.GetValue("U_" + oCol, oRow - 1).ToString().Trim()))
                            {
                                PS_CO658_AddMatrixRow(oRow, false);
                            }
                            oMat01.LoadFromDataSource();
                        }

                        oMat01.Columns.Item("UseYN").Cells.Item(oRow).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                        oMat01.Columns.Item(oCol).Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oMat01.AutoResizeColumns();

                        break;
                    case "Code":
                        break;
                }
            }
            catch(Exception ex)
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

                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                    //Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                    //Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
                    Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    //Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
                    //Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
                    //Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                    //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    //Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                    //Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                    //Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                    break;

                case SAPbouiCOM.BoEventTypes.et_Drag: //39
                    //Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
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
                    if (pVal.ItemUID == "PS_CO658")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }

                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_CO658_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry01 = oForm.Items.Item("Code").Specific.VALUE;
                            oFormMode01 = oForm.Mode;

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_CO658_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry01 = oForm.Items.Item("Code").Specific.VALUE;
                            oFormMode01 = oForm.Mode;

                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
				{
                    if (pVal.ItemUID == "PS_CO658")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }

                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_CO658_FormItemEnabled();
                                PS_CO658_AddMatrixRow(0, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_CO658_FormItemEnabled();
                            }
                        }
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

        /// <summary>
        /// KEY_DOWN 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();            

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "AcctCode")
                        {
                            dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "AcctCode"); //계정
                        }
                    }
                    else if (pVal.ItemUID == "Code")
                    {
                        //dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "AcctCode", ""); //계정 포맷서치 설정
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;

                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    else
                    {
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
                        PS_CO658_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }
            finally
            {
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_CO658_FormItemEnabled();
                    PS_CO658_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
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
        /// RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_CO658_FormResize();
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent)
        {
            int i = 0;

            try
            {
                if (pVal.BeforeAction == true)
                {
                }
                else if (pVal.BeforeAction == false)
                {
                    for (i = 1; i <= oMat01.VisualRowCount; i++)
                    {
                        oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.VALUE = i;
                    }

                    oMat01.FlushToDataSource();
                    oDS_PS_CO658L.RemoveRecord(oDS_PS_CO658L.Size - 1);
                    oMat01.LoadFromDataSource();

                    if (oMat01.RowCount == 0)
                    {
                        PS_CO658_AddMatrixRow(0, false);
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(oDS_PS_CO658L.GetValue("U_AcctCode", oMat01.RowCount - 1).ToString().Trim()))
                        {
                            PS_CO658_AddMatrixRow(oMat01.RowCount, false);
                        }
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
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
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
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_CO658_FormItemEnabled();
                            oForm.Items.Item("Name").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_CO658_FormItemEnabled();
                            PS_CO658_AddMatrixRow(0, true);
                            break;
                        case "1288": //레코드이동(최초)
                        case "1289": //레코드이동(이전)
                        case "1290": //레코드이동(다음)
                        case "1291": //레코드이동(최종)
                            PS_CO658_FormItemEnabled();
                            break;
                        case "1287": //복제
                            oForm.Freeze(true);
                            PS_CO658_FormClear();
                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_CO658L.SetValue("Code", i, "");
                                oMat01.LoadFromDataSource();
                            }
                            oForm.Freeze(false);
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
            }
        }


        #region (최종테스트 후 주석 삭제 필요, 2020.09.08 송명규)

        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	switch (pval.EventType)
        //	{
        //		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //			////1
        //			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //			////2
        //			Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //			////5
        //			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CLICK:
        //			////6
        //			Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //			////7
        //			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //			////8
        //			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //			////10
        //			Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //			////11
        //			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //			////18
        //			break;
        //		////et_FORM_ACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //			////19
        //			break;
        //		////et_FORM_DEACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //			////20
        //			Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //			////27
        //			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //			////3
        //			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //			////4
        //			break;
        //		////et_LOST_FOCUS
        //		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //			////17
        //			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
        //			break;
        //	}
        //	return;
        //Raise_ItemEvent_Error:

        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	short i = 0;

        //	////BeforeAction = True
        //	if ((pval.BeforeAction == true))
        //	{

        //	}
        //	else if ((pval.BeforeAction == false))
        //	{

        //	}
        //	return;
        //Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true))
        //	{

        //	}
        //	else if ((BusinessObjectInfo.BeforeAction == false))
        //	{
        //		switch (BusinessObjectInfo.EventType)
        //		{
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	}
        //	return;
        //Raise_FormDataEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //    	if (pval.BeforeAction == true)
        //    	{
        //    		        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then

        //                        Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //                        Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //                        MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //                        MenuCreationParams01.uniqueID = "MenuUID"
        //                        MenuCreationParams01.String = "메뉴명"
        //                        MenuCreationParams01.Enabled = True
        //                        Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //                    End If
        //      }
        //    	else if (pval.BeforeAction == false)
        //    	{
        //    		        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //                      Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //                  End If
        //      }

        //	return;
        //Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{
        //		
        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}

        //	return;
        //Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement
        //	return;
        //Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	oForm01.Freeze(true);
        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //		if (pval.ItemChanged == true)
        //		{

        //		}
        //	}
        //	oForm01.Freeze(false);
        //	return;
        //Raise_EVENT_COMBO_SELECT_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_DOUBLE_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	object oTempClass = null;
        //	if (pval.BeforeAction == true)
        //	{
        //		if (pval.ItemUID == "Mat01")
        //		{

        //		}
        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement


        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;

        //	oForm01.Freeze(true);
        //	if (pval.BeforeAction == true)
        //	{
        //		if (pval.ItemChanged == true)
        //		{

        //		}
        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	oForm01.Freeze(false);
        //	return;
        //Raise_EVENT_VALIDATE_Error:
        //	oForm01.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{
        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pval = null, ref bool BubbleEvent = false)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{

        //	}
        //	return;
        //Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	if (pval.BeforeAction == true)
        //	{

        //	}
        //	else if (pval.BeforeAction == false)
        //	{
        //	}
        //	return;
        //Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //	// ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if ((oLastColRow01 > 0))
        //	{
        //		if (pval.BeforeAction == true)
        //		{
        //			////행삭제전 행삭제가능여부검사
        //		}
        //		else if (pval.BeforeAction == false)
        //		{

        //		}
        //	}
        //	return;
        //Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_CO658_MTX01 (사용 안됨)
        ///// <summary>
        ///// 데이터 로드(사용안됨)
        ///// </summary>
        //private void PS_CO658_MTX01()
        //{
        //    int i = 0;
        //    string Query01 = string.Empty;
        //    string Param01 = string.Empty;
        //    string Param02 = string.Empty;
        //    string Param03 = string.Empty;
        //    string Param04 = string.Empty;

        //    string errCode = string.Empty;

        //    SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("조회시작!", 100, false);

        //    try
        //    {
        //        oForm.Freeze(true);

        //        Param01 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
        //        Param02 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
        //        Param03 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();
        //        Param04 = oForm.Items.Item("Param01").Specific.Value.ToString().Trim();

        //        Query01 = "SELECT 10";
        //        oRecordSet01.DoQuery(Query01);

        //        oMat01.Clear();
        //        oMat01.FlushToDataSource();
        //        oMat01.LoadFromDataSource();

        //        if (oRecordSet01.RecordCount == 0)
        //        {
        //            errCode = "1";
        //            throw new Exception();
        //        }

        //        for (i = 0; i <= oRecordSet01.RecordCount - 1; i++)
        //        {
        //            if (i != 0)
        //            {
        //                oDS_PS_CO658L.InsertRecord((i));
        //            }
        //            oDS_PS_CO658L.Offset = i;
        //            oDS_PS_CO658L.SetValue("U_COL01", i, oRecordSet01.Fields.Item(0).Value);
        //            oDS_PS_CO658L.SetValue("U_COL02", i, oRecordSet01.Fields.Item(1).Value);
        //            oRecordSet01.MoveNext();
        //            ProgBar01.Value = ProgBar01.Value + 1;
        //            ProgBar01.Text = ProgBar01.Value + "/" + oRecordSet01.RecordCount + "건 조회중...!";
        //        }

        //        oMat01.LoadFromDataSource();
        //        oMat01.AutoResizeColumns();
        //        oForm.Update();
        //    }
        //    catch (Exception ex)
        //    {
        //        ProgBar01.Stop();

        //        if (errCode == "1")
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("조회 결과가 없습니다. 확인하세요.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
        //        }
        //        else
        //        {
        //            PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY204_MTX01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //        }
        //    }
        //    finally
        //    {
        //        oForm.Freeze(false);
        //        if (ProgBar01 != null)
        //        {
        //            ProgBar01.Stop();
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
        //        }
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
        //    }
        //}
        #endregion

        #region PS_CO658_Validate (사용 안됨)
        ///// <summary>
        ///// Validate(사용안됨)
        ///// </summary>
        ///// <param name="ValidateType"></param>
        ///// <returns></returns>
        //private bool PS_CO658_Validate(string ValidateType)
        //{
        //    bool functionReturnValue = false;

        //    SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    try
        //    {
        //        if (ValidateType == "수정")
        //        {
        //            //삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
        //        }
        //        else if (ValidateType == "행삭제")
        //        {
        //            //행삭제전 행삭제가능여부검사
        //        }
        //        else if (ValidateType == "취소")
        //        {
        //        }

        //        functionReturnValue = true;
        //    }
        //    catch(Exception ex)
        //    {
        //        PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
        //    }

        //    return functionReturnValue;
        //}
        #endregion
        
        #endregion

    }
}
