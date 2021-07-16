using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작번별청구현황조회
	/// </summary>
	internal class PS_SD052 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD052.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD052_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD052");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_SD052_CreateItems();
				PS_SD052_ComboBox_Setting();
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
		/// PS_SD052_CreateItems
		/// </summary>
		private void PS_SD052_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//품목코드(작번)
				oForm.DataSources.UserDataSources.Add("OrdNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum").Specific.DataBind.SetBound(true, "", "OrdNum");

				//품목명
				oForm.DataSources.UserDataSources.Add("FrgnName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("FrgnName").Specific.DataBind.SetBound(true, "", "FrgnName");

				//품목코드(자재)
				oForm.DataSources.UserDataSources.Add("SItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("SItemCode").Specific.DataBind.SetBound(true, "", "SItemCode");

				//품목명(자재)
				oForm.DataSources.UserDataSources.Add("SItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("SItemName").Specific.DataBind.SetBound(true, "", "SItemName");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				//품의구분
				oForm.DataSources.UserDataSources.Add("POType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("POType").Specific.DataBind.SetBound(true, "", "POType");

				//품의완료여부
				oForm.DataSources.UserDataSources.Add("POYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("POYN").Specific.DataBind.SetBound(true, "", "POYN");

				//부자재품의 제외 여부
				oForm.DataSources.UserDataSources.Add("POType20YN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("POType20YN").Specific.DataBind.SetBound(true, "", "POType20YN");

				//실적금액계
				oForm.DataSources.UserDataSources.Add("TResAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("TResAmt").Specific.DataBind.SetBound(true, "", "TResAmt");

				//품의금액계
				oForm.DataSources.UserDataSources.Add("T030Amt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("T030Amt").Specific.DataBind.SetBound(true, "", "T030Amt");

				//검수금액계
				oForm.DataSources.UserDataSources.Add("T070Amt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("T070Amt").Specific.DataBind.SetBound(true, "", "T070Amt");

				oForm.Items.Item("POType20YN").Specific.Checked = true;
				oForm.Items.Item("FrDt").Specific.Value = DateTime.Now.ToString("yyyyMM01");
				oForm.Items.Item("ToDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
				oForm.Items.Item("OrdNum").Click();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD052_ComboBox_Setting
		/// </summary>
		private void PS_SD052_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//거래처구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'C100'";
				oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, sQry, "", false, false);
				oForm.Items.Item("CardType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품목구분
				sQry = " SELECT      U_Minor AS [Code], ";
				sQry += "             U_CdName AS [Name]";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'S002'";
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, sQry, "", false, false);
				oForm.Items.Item("ItemType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품의구분
				sQry = " SELECT      Code AS [Code],";
				sQry += "             Name AS [Name]";
				sQry += " FROM        [@PSH_ORDTYP]";
				sQry += " WHERE       Code IN ('10','20','30','40')"; //4개 품의에 대해서만 조회
				sQry += sQry + " ORDER BY    Code";
				oForm.Items.Item("POType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("POType").Specific, sQry, "", false, false);
				oForm.Items.Item("POType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//품의완료여부
				oForm.Items.Item("POYN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("POYN").Specific.ValidValues.Add("Y", "품의완료");
				oForm.Items.Item("POYN").Specific.ValidValues.Add("N", "품의미완료");
				oForm.Items.Item("POYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD052_MTX01
		/// </summary>
		private void PS_SD052_MTX01()
		{
			int loopCount01;
			string sQry;
			string errMessage = string.Empty;

			string BPLID;	   //사업장
			string FrDt;	   //기간(Fr)
			string ToDt;	   //기간(To)
			string OrdNum;	   //품목코드(작번)
			string SItemCode;  //품목코드(자재)
			string CardType;   //거래처구분
			string ItemType;   //품목구분
			string POType;	   //품의구분
			string POYN;	   //품의완료여부
			string POType20YN; //부자재품의 제외 여부

			double TResultAmt = 0; //실적금액계
			double TMM030Amt = 0;  //품의금액계
			double TMM070Amt = 0;  //검수금액계

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();
				SItemCode = oForm.Items.Item("SItemCode").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				POType = oForm.Items.Item("POType").Specific.Value.ToString().Trim();
				POYN = oForm.Items.Item("POYN").Specific.Value.ToString().Trim();

				if (oForm.Items.Item("POType20YN").Specific.Checked == true)
                {
					POType20YN = "Y";
				}
				else
                {
					POType20YN = "N";
				}

				ProgressBar01.Text = "조회시작!";

				sQry = " EXEC PS_SD052_01 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += OrdNum + "','";
				sQry += SItemCode + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += POType + "','";
				sQry += POYN + "','";
				sQry += POType20YN + "'";

				oGrid.DataTable.Clear();
				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

				oGrid.Columns.Item(29).RightJustified = true; //요청수량
				oGrid.Columns.Item(30).RightJustified = true; //실적단가
				oGrid.Columns.Item(31).RightJustified = true; //실적금액
				oGrid.Columns.Item(34).RightJustified = true; //품의수량/중량
				oGrid.Columns.Item(35).RightJustified = true; //품의금액
				oGrid.Columns.Item(38).RightJustified = true; //대비견적금액1
				oGrid.Columns.Item(40).RightJustified = true; //대비견적금액2
				oGrid.Columns.Item(42).RightJustified = true; //대비견적금액3
				oGrid.Columns.Item(44).RightJustified = true; //가입고수량/중량
				oGrid.Columns.Item(45).RightJustified = true; //가입고금액
				oGrid.Columns.Item(51).RightJustified = true; //검수수량/중량
				oGrid.Columns.Item(52).RightJustified = true; //검수금액
				oGrid.Columns.Item(59).RightJustified = true; //수주금액

				for (loopCount01 = 0; loopCount01 <= oGrid.Rows.Count - 1; loopCount01++)
				{
					if (oGrid.DataTable.GetValue(31, loopCount01).ToString().Trim() == "")
                    {
						TResultAmt += 0;
					}
					else
                    {
						TResultAmt += Convert.ToDouble(oGrid.DataTable.GetValue(31, loopCount01).ToString().Trim());
					}

					if (oGrid.DataTable.GetValue(35, loopCount01).ToString().Trim() == "")
					{
						TMM030Amt += 0;
					}
					else
					{
						TMM030Amt += Convert.ToDouble(oGrid.DataTable.GetValue(35, loopCount01));
					}

					if (oGrid.DataTable.GetValue(52, loopCount01).ToString().Trim() == "")
					{
						TMM070Amt += 0;
					}
					else
					{
						TMM070Amt += Convert.ToDouble(oGrid.DataTable.GetValue(52, loopCount01));
					}
				}

				oForm.Items.Item("TResAmt").Specific.Value = TResultAmt;
				oForm.Items.Item("T030Amt").Specific.Value = TMM030Amt;
				oForm.Items.Item("T070Amt").Specific.Value = TMM070Amt;

				if (oGrid.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid.AutoResizeColumns();
				oForm.Update();
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
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
			switch (pVal.EventType)
			{
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED: //1
                    Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    //Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "BtnSearch")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_SD052_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
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
		}

		/// <summary>
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "SItemCode", "");
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "OrdNum")
					{
						oForm.Items.Item("FrgnName").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", "");
					}
					else if (pVal.ItemUID == "SItemCode")
					{
						oForm.Items.Item("SItemName").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", "");
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_UNLOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
