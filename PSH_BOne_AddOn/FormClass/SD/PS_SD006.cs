using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 판매계획등록(분말)
	/// </summary>
	internal class PS_SD006 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD006H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD006L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		
		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD006.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD006_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD006");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "Code";

				oForm.Freeze(true);

				PS_SD006_CreateItems();
				PS_SD006_SetComboBox();
				PS_SD006_Initialize();
                PS_SD006_AddMatrixRow(0, true);

                oForm.EnableMenu("1283", true); //삭제
				oForm.EnableMenu("1287", true); //복제
				oForm.EnableMenu("1286", false); //닫기
				oForm.EnableMenu("1284", false); //취소
				oForm.EnableMenu("1293", true); //행삭제
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
            }
        }

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_SD006_CreateItems()
        {
            try
            {
                oDS_PS_SD006H = oForm.DataSources.DBDataSources.Item("@PS_SD006H");
                oDS_PS_SD006L = oForm.DataSources.DBDataSources.Item("@PS_SD006L");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("T_Wgt1", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt1").Specific.DataBind.SetBound(true, "", "T_Wgt1");

                oForm.DataSources.UserDataSources.Add("T_Wgt2", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt2").Specific.DataBind.SetBound(true, "", "T_Wgt2");

                oForm.DataSources.UserDataSources.Add("T_Wgt3", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt3").Specific.DataBind.SetBound(true, "", "T_Wgt3");

                oForm.DataSources.UserDataSources.Add("T_Wgt4", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt4").Specific.DataBind.SetBound(true, "", "T_Wgt4");

                oForm.DataSources.UserDataSources.Add("T_Wgt5", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt5").Specific.DataBind.SetBound(true, "", "T_Wgt5");

                oForm.DataSources.UserDataSources.Add("T_Wgt6", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt6").Specific.DataBind.SetBound(true, "", "T_Wgt6");

                oForm.DataSources.UserDataSources.Add("T_Wgt7", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt7").Specific.DataBind.SetBound(true, "", "T_Wgt7");

                oForm.DataSources.UserDataSources.Add("T_Wgt8", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt8").Specific.DataBind.SetBound(true, "", "T_Wgt8");

                oForm.DataSources.UserDataSources.Add("T_Wgt9", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt9").Specific.DataBind.SetBound(true, "", "T_Wgt9");

                oForm.DataSources.UserDataSources.Add("T_Wgt10", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt10").Specific.DataBind.SetBound(true, "", "T_Wgt10");

                oForm.DataSources.UserDataSources.Add("T_Wgt11", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt11").Specific.DataBind.SetBound(true, "", "T_Wgt11");

                oForm.DataSources.UserDataSources.Add("T_Wgt12", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("T_Wgt12").Specific.DataBind.SetBound(true, "", "T_Wgt12");

                oForm.DataSources.UserDataSources.Add("T_Amt1", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt1").Specific.DataBind.SetBound(true, "", "T_Amt1");

                oForm.DataSources.UserDataSources.Add("T_Amt2", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt2").Specific.DataBind.SetBound(true, "", "T_Amt2");

                oForm.DataSources.UserDataSources.Add("T_Amt3", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt3").Specific.DataBind.SetBound(true, "", "T_Amt3");

                oForm.DataSources.UserDataSources.Add("T_Amt4", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt4").Specific.DataBind.SetBound(true, "", "T_Amt4");

                oForm.DataSources.UserDataSources.Add("T_Amt5", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt5").Specific.DataBind.SetBound(true, "", "T_Amt5");

                oForm.DataSources.UserDataSources.Add("T_Amt6", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt6").Specific.DataBind.SetBound(true, "", "T_Amt6");

                oForm.DataSources.UserDataSources.Add("T_Amt7", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt7").Specific.DataBind.SetBound(true, "", "T_Amt7");

                oForm.DataSources.UserDataSources.Add("T_Amt8", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt8").Specific.DataBind.SetBound(true, "", "T_Amt8");

                oForm.DataSources.UserDataSources.Add("T_Amt9", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt9").Specific.DataBind.SetBound(true, "", "T_Amt9");

                oForm.DataSources.UserDataSources.Add("T_Amt10", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt10").Specific.DataBind.SetBound(true, "", "T_Amt10");

                oForm.DataSources.UserDataSources.Add("T_Amt11", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt11").Specific.DataBind.SetBound(true, "", "T_Amt11");

                oForm.DataSources.UserDataSources.Add("T_Amt12", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("T_Amt12").Specific.DataBind.SetBound(true, "", "T_Amt12");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_SD006_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                //사업장
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면 초기화
        /// </summary>
        private void PS_SD006_Initialize()
        {
            try
            {
                oForm.Items.Item("Year").Specific.Value = DateTime.Now.ToString("yyyy");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD006_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                if (RowIserted == false)
                {
                    oDS_PS_SD006L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD006L.Offset = oRow;
                oDS_PS_SD006L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_SD006_CheckDataValid()
        {
            bool returnValue = false;
            string errMessage = string.Empty;

            try
            {
                if (string.IsNullOrEmpty(oDS_PS_SD006H.GetValue("U_CardCode", 0)))
                {
                    errMessage = "거래처는 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oDS_PS_SD006H.GetValue("U_BPLId", 0)))
                {
                    errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oDS_PS_SD006H.GetValue("U_Year", 0)))
                {
                    errMessage = "년도는 필수입력사항입니다.확인하세요.";
                    throw new Exception();
                }

                if (oDS_PS_SD006L.Size > 1)
                {
                    oDS_PS_SD006L.RemoveRecord(oDS_PS_SD006L.Size - 1);
                }
                oMat01.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
            }

            return returnValue;
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
            string cardCode;
            string year;
            string code;
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD006_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                cardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                                year = oForm.Items.Item("Year").Specific.Value.ToString().Trim();

                                code = codeHelpClass.Right(year, 2) + cardCode;

                                oDS_PS_SD006H.SetValue("Code", 0, code);
                            }
                        }
                    }
                }
                else if (pVal.BeforeAction == false)
                {
                    if (pVal.ItemUID == "1")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PSH_Globals.SBO_Application.ActivateMenuItem("1282");
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
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "CardCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
                            }
                        }
                        else if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "ItemCode")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    PS_SM010 tempForm = new PS_SM010();
                                    tempForm.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
                                    BubbleEvent = false;
                                }
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
            string sQry;
            string errMessage = string.Empty;

            int i = 0;

            double T_Wgt1 = 0;
            double T_Wgt2 = 0;
            double T_Wgt3 = 0;
            double T_Wgt4 = 0;
            double T_Wgt5 = 0;
            double T_Wgt6 = 0;
            double T_Wgt7 = 0;
            double T_Wgt8 = 0;
            double T_Wgt9 = 0;
            double T_Wgt10 = 0;
            double T_Wgt11 = 0;
            double T_Wgt12 = 0;

            double T_Amt1 = 0;
            double T_Amt2 = 0;
            double T_Amt3 = 0;
            double T_Amt4 = 0;
            double T_Amt5 = 0;
            double T_Amt6 = 0;
            double T_Amt7 = 0;
            double T_Amt8 = 0;
            double T_Amt9 = 0;
            double T_Amt10 = 0;
            double T_Amt11 = 0;
            double T_Amt12 = 0;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
                        if (pVal.ItemUID == "CardCode") //거래처명조회
                        {
                            
                            sQry = "SELECT CardName FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
                            oRecordSet01.DoQuery(sQry);
                            oForm.Items.Item("CardName").Specific.Value = oRecordSet01.Fields.Item(0).Value.ToString().Trim();
                        }

                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.FlushToDataSource();
                            if (pVal.ColUID == "ItemCode")
                            {
                                if ((pVal.Row == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
                                {
                                    oMat01.FlushToDataSource();
                                    PS_SD006_AddMatrixRow(oMat01.RowCount, false);
                                    oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }

                                sQry = "  Select    ItemName ";
                                sQry += " From      OITM ";
                                sQry += " Where     ItemCode = '" + oDS_PS_SD006L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                oMat01.FlushToDataSource();

                                if (oRecordSet01.RecordCount == 0)
                                {
                                    oDS_PS_SD006L.SetValue("U_ItemName", pVal.Row - 1, "");

                                    errMessage = "조회 결과가 없습니다. 확인하세요.";
                                    oMat01.LoadFromDataSource();
                                    throw new Exception();
                                }

                                oDS_PS_SD006L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemName").Value.ToString().Trim());
                                oMat01.LoadFromDataSource();
                                oMat01.AutoResizeColumns();
                            }
                            else
                            {
                                if (pVal.ColUID == "Wgt1" || pVal.ColUID == "Prc1")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt1", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt1").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc1").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt2" || pVal.ColUID == "Prc2")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt2", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt2").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc2").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt3" || pVal.ColUID == "Prc3")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt3", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt3").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc3").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt4" || pVal.ColUID == "Prc4")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt4", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt4").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc4").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt5" || pVal.ColUID == "Prc5")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt5", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt5").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc5").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt6" || pVal.ColUID == "Prc6")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt6", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt6").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc6").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt7" || pVal.ColUID == "Prc7")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt7", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt7").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc7").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt8" || pVal.ColUID == "Prc8")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt8", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt8").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc8").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt9" || pVal.ColUID == "Prc9")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt9", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt9").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc9").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt10" || pVal.ColUID == "Prc10")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt10", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt10").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc10").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt11" || pVal.ColUID == "Prc11")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt11", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt11").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc11").Cells.Item(i + 1).Specific.Value));
                                }
                                else if (pVal.ColUID == "Wgt12" || pVal.ColUID == "Prc12")
                                {
                                    oDS_PS_SD006L.SetValue("U_Amt12", pVal.Row - 1, Convert.ToDouble(oMat01.Columns.Item("Wgt12").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Prc12").Cells.Item(i + 1).Specific.Value));
                                }

                                oMat01.LoadFromDataSourceEx();
                                oMat01.AutoResizeColumns();

                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    T_Wgt1 += Convert.ToDouble(oMat01.Columns.Item("Wgt1").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt2 += Convert.ToDouble(oMat01.Columns.Item("Wgt2").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt3 += Convert.ToDouble(oMat01.Columns.Item("Wgt3").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt4 += Convert.ToDouble(oMat01.Columns.Item("Wgt4").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt5 += Convert.ToDouble(oMat01.Columns.Item("Wgt5").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt6 += Convert.ToDouble(oMat01.Columns.Item("Wgt6").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt7 += Convert.ToDouble(oMat01.Columns.Item("Wgt7").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt8 += Convert.ToDouble(oMat01.Columns.Item("Wgt8").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt9 += Convert.ToDouble(oMat01.Columns.Item("Wgt9").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt10 += Convert.ToDouble(oMat01.Columns.Item("Wgt10").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt11 += Convert.ToDouble(oMat01.Columns.Item("Wgt11").Cells.Item(i + 1).Specific.Value);
                                    T_Wgt12 += Convert.ToDouble(oMat01.Columns.Item("Wgt12").Cells.Item(i + 1).Specific.Value);

                                    T_Amt1 += Convert.ToDouble(oMat01.Columns.Item("Amt1").Cells.Item(i + 1).Specific.Value);
                                    T_Amt2 += Convert.ToDouble(oMat01.Columns.Item("Amt2").Cells.Item(i + 1).Specific.Value);
                                    T_Amt3 += Convert.ToDouble(oMat01.Columns.Item("Amt3").Cells.Item(i + 1).Specific.Value);
                                    T_Amt4 += Convert.ToDouble(oMat01.Columns.Item("Amt4").Cells.Item(i + 1).Specific.Value);
                                    T_Amt5 += Convert.ToDouble(oMat01.Columns.Item("Amt5").Cells.Item(i + 1).Specific.Value);
                                    T_Amt6 += Convert.ToDouble(oMat01.Columns.Item("Amt6").Cells.Item(i + 1).Specific.Value);
                                    T_Amt7 += Convert.ToDouble(oMat01.Columns.Item("Amt7").Cells.Item(i + 1).Specific.Value);
                                    T_Amt8 += Convert.ToDouble(oMat01.Columns.Item("Amt8").Cells.Item(i + 1).Specific.Value);
                                    T_Amt9 += Convert.ToDouble(oMat01.Columns.Item("Amt9").Cells.Item(i + 1).Specific.Value);
                                    T_Amt10 += Convert.ToDouble(oMat01.Columns.Item("Amt10").Cells.Item(i + 1).Specific.Value);
                                    T_Amt11 += Convert.ToDouble(oMat01.Columns.Item("Amt11").Cells.Item(i + 1).Specific.Value);
                                    T_Amt12 += Convert.ToDouble(oMat01.Columns.Item("Amt12").Cells.Item(i + 1).Specific.Value);
                                }

                                oForm.Items.Item("T_Wgt1").Specific.Value = T_Wgt1;
                                oForm.Items.Item("T_Wgt2").Specific.Value = T_Wgt2;
                                oForm.Items.Item("T_Wgt3").Specific.Value = T_Wgt3;
                                oForm.Items.Item("T_Wgt4").Specific.Value = T_Wgt4;
                                oForm.Items.Item("T_Wgt5").Specific.Value = T_Wgt5;
                                oForm.Items.Item("T_Wgt6").Specific.Value = T_Wgt6;
                                oForm.Items.Item("T_Wgt7").Specific.Value = T_Wgt7;
                                oForm.Items.Item("T_Wgt8").Specific.Value = T_Wgt8;
                                oForm.Items.Item("T_Wgt9").Specific.Value = T_Wgt9;
                                oForm.Items.Item("T_Wgt10").Specific.Value = T_Wgt10;
                                oForm.Items.Item("T_Wgt11").Specific.Value = T_Wgt11;
                                oForm.Items.Item("T_Wgt12").Specific.Value = T_Wgt12;

                                oForm.Items.Item("T_Amt1").Specific.Value = T_Amt1;
                                oForm.Items.Item("T_Amt2").Specific.Value = T_Amt2;
                                oForm.Items.Item("T_Amt3").Specific.Value = T_Amt3;
                                oForm.Items.Item("T_Amt4").Specific.Value = T_Amt4;
                                oForm.Items.Item("T_Amt5").Specific.Value = T_Amt5;
                                oForm.Items.Item("T_Amt6").Specific.Value = T_Amt6;
                                oForm.Items.Item("T_Amt7").Specific.Value = T_Amt7;
                                oForm.Items.Item("T_Amt8").Specific.Value = T_Amt8;
                                oForm.Items.Item("T_Amt9").Specific.Value = T_Amt9;
                                oForm.Items.Item("T_Amt10").Specific.Value = T_Amt10;
                                oForm.Items.Item("T_Amt11").Specific.Value = T_Amt11;
                                oForm.Items.Item("T_Amt12").Specific.Value = T_Amt12;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
            int i;

            double T_Wgt1 = 0;
            double T_Wgt2 = 0;
            double T_Wgt3 = 0;
            double T_Wgt4 = 0;
            double T_Wgt5 = 0;
            double T_Wgt6 = 0;
            double T_Wgt7 = 0;
            double T_Wgt8 = 0;
            double T_Wgt9 = 0;
            double T_Wgt10 = 0;
            double T_Wgt11 = 0;
            double T_Wgt12 = 0;

            double T_Amt1 = 0;
            double T_Amt2 = 0;
            double T_Amt3 = 0;
            double T_Amt4 = 0;
            double T_Amt5 = 0;
            double T_Amt6 = 0;
            double T_Amt7 = 0;
            double T_Amt8 = 0;
            double T_Amt9 = 0;
            double T_Amt10 = 0;
            double T_Amt11 = 0;
            double T_Amt12 = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        T_Wgt1 += Convert.ToDouble(oMat01.Columns.Item("Wgt1").Cells.Item(i + 1).Specific.Value);
                        T_Wgt2 += Convert.ToDouble(oMat01.Columns.Item("Wgt2").Cells.Item(i + 1).Specific.Value);
                        T_Wgt3 += Convert.ToDouble(oMat01.Columns.Item("Wgt3").Cells.Item(i + 1).Specific.Value);
                        T_Wgt4 += Convert.ToDouble(oMat01.Columns.Item("Wgt4").Cells.Item(i + 1).Specific.Value);
                        T_Wgt5 += Convert.ToDouble(oMat01.Columns.Item("Wgt5").Cells.Item(i + 1).Specific.Value);
                        T_Wgt6 += Convert.ToDouble(oMat01.Columns.Item("Wgt6").Cells.Item(i + 1).Specific.Value);
                        T_Wgt7 += Convert.ToDouble(oMat01.Columns.Item("Wgt7").Cells.Item(i + 1).Specific.Value);
                        T_Wgt8 += Convert.ToDouble(oMat01.Columns.Item("Wgt8").Cells.Item(i + 1).Specific.Value);
                        T_Wgt9 += Convert.ToDouble(oMat01.Columns.Item("Wgt9").Cells.Item(i + 1).Specific.Value);
                        T_Wgt10 += Convert.ToDouble(oMat01.Columns.Item("Wgt10").Cells.Item(i + 1).Specific.Value);
                        T_Wgt11 += Convert.ToDouble(oMat01.Columns.Item("Wgt11").Cells.Item(i + 1).Specific.Value);
                        T_Wgt12 += Convert.ToDouble(oMat01.Columns.Item("Wgt12").Cells.Item(i + 1).Specific.Value);

                        T_Amt1 += Convert.ToDouble(oMat01.Columns.Item("Amt1").Cells.Item(i + 1).Specific.Value);
                        T_Amt2 += Convert.ToDouble(oMat01.Columns.Item("Amt2").Cells.Item(i + 1).Specific.Value);
                        T_Amt3 += Convert.ToDouble(oMat01.Columns.Item("Amt3").Cells.Item(i + 1).Specific.Value);
                        T_Amt4 += Convert.ToDouble(oMat01.Columns.Item("Amt4").Cells.Item(i + 1).Specific.Value);
                        T_Amt5 += Convert.ToDouble(oMat01.Columns.Item("Amt5").Cells.Item(i + 1).Specific.Value);
                        T_Amt6 += Convert.ToDouble(oMat01.Columns.Item("Amt6").Cells.Item(i + 1).Specific.Value);
                        T_Amt7 += Convert.ToDouble(oMat01.Columns.Item("Amt7").Cells.Item(i + 1).Specific.Value);
                        T_Amt8 += Convert.ToDouble(oMat01.Columns.Item("Amt8").Cells.Item(i + 1).Specific.Value);
                        T_Amt9 += Convert.ToDouble(oMat01.Columns.Item("Amt9").Cells.Item(i + 1).Specific.Value);
                        T_Amt10 += Convert.ToDouble(oMat01.Columns.Item("Amt10").Cells.Item(i + 1).Specific.Value);
                        T_Amt11 += Convert.ToDouble(oMat01.Columns.Item("Amt11").Cells.Item(i + 1).Specific.Value);
                        T_Amt12 += Convert.ToDouble(oMat01.Columns.Item("Amt12").Cells.Item(i + 1).Specific.Value);
                    }

                    oForm.Items.Item("T_Wgt1").Specific.Value = T_Wgt1;
                    oForm.Items.Item("T_Wgt2").Specific.Value = T_Wgt2;
                    oForm.Items.Item("T_Wgt3").Specific.Value = T_Wgt3;
                    oForm.Items.Item("T_Wgt4").Specific.Value = T_Wgt4;
                    oForm.Items.Item("T_Wgt5").Specific.Value = T_Wgt5;
                    oForm.Items.Item("T_Wgt6").Specific.Value = T_Wgt6;
                    oForm.Items.Item("T_Wgt7").Specific.Value = T_Wgt7;
                    oForm.Items.Item("T_Wgt8").Specific.Value = T_Wgt8;
                    oForm.Items.Item("T_Wgt9").Specific.Value = T_Wgt9;
                    oForm.Items.Item("T_Wgt10").Specific.Value = T_Wgt10;
                    oForm.Items.Item("T_Wgt11").Specific.Value = T_Wgt11;
                    oForm.Items.Item("T_Wgt12").Specific.Value = T_Wgt12;

                    oForm.Items.Item("T_Amt1").Specific.Value = T_Amt1;
                    oForm.Items.Item("T_Amt2").Specific.Value = T_Amt2;
                    oForm.Items.Item("T_Amt3").Specific.Value = T_Amt3;
                    oForm.Items.Item("T_Amt4").Specific.Value = T_Amt4;
                    oForm.Items.Item("T_Amt5").Specific.Value = T_Amt5;
                    oForm.Items.Item("T_Amt6").Specific.Value = T_Amt6;
                    oForm.Items.Item("T_Amt7").Specific.Value = T_Amt7;
                    oForm.Items.Item("T_Amt8").Specific.Value = T_Amt8;
                    oForm.Items.Item("T_Amt9").Specific.Value = T_Amt9;
                    oForm.Items.Item("T_Amt10").Specific.Value = T_Amt10;
                    oForm.Items.Item("T_Amt11").Specific.Value = T_Amt11;
                    oForm.Items.Item("T_Amt12").Specific.Value = T_Amt12;

                    oMat01.AutoResizeColumns();
                    PS_SD006_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD006H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD006L);
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
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    oMat01.AutoResizeColumns();
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
        /// 행삭제 체크 메소드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (int i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_SD006L.RemoveRecord(oDS_PS_SD006L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_SD006_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_SD006L.GetValue("U_ItemCode", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_SD006_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
                else if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            break;
                        case "1286": //닫기
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            oForm.Freeze(true);
                            PS_SD006_Initialize();
                            oForm.Freeze(false);
                            break;
                        case "1282": //추가
                            oForm.Freeze(true);
                            PS_SD006_Initialize();
                            PS_SD006_AddMatrixRow(0, true);
                            oForm.Freeze(false);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            break;
                        case "1287": //복제
                            oForm.Freeze(true);
                            oDS_PS_SD006H.SetValue("Code", 0, "");

                            for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                oMat01.FlushToDataSource();
                                oDS_PS_SD006L.SetValue("Code", i, "");
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
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
    }
}
