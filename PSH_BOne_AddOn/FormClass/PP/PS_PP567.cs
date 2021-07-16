using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 장비이익률관리 등록 현황 조회
	/// </summary>
	internal class PS_PP567 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DataTable oDS_PS_PP567A;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP567.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP567_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP567");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP567_CreateItems();
				PS_PP567_ComboBox_Setting();

				oForm.Items.Item("BtnPrt01").Visible = false;
				oForm.Items.Item("CardCode").Click();
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
		/// PS_PP567_CreateItems
		/// </summary>
		private void PS_PP567_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("PS_PP567A");
				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_PP567A");
				oDS_PS_PP567A = oForm.DataSources.DataTables.Item("PS_PP567A");

				//수주년월(시작)
				oForm.DataSources.UserDataSources.Add("FrYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("FrYM").Specific.DataBind.SetBound(true, "", "FrYM");

				//수주년월(종료)
				oForm.DataSources.UserDataSources.Add("ToYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("ToYM").Specific.DataBind.SetBound(true, "", "ToYM");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

				//수주처
				oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

				//수주처명
				oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

				//작번
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemSpec").Specific.DataBind.SetBound(true, "", "ItemSpec");

				//자체/외주
				oForm.DataSources.UserDataSources.Add("InOut", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("InOut").Specific.DataBind.SetBound(true, "", "InOut");

				//수주금액
				oForm.DataSources.UserDataSources.Add("OrdAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("OrdAmt").Specific.DataBind.SetBound(true, "", "OrdAmt");

				//생산완료포함CheckBox
				oForm.DataSources.UserDataSources.Add("CmltYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CmltYN").Specific.DataBind.SetBound(true, "", "CmltYN");

				oForm.Items.Item("FrYM").Specific.VALUE = DateTime.Now.AddMonths(-6).ToString("yyyyMM"); //6개월전
				oForm.Items.Item("ToYM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");
				oForm.Items.Item("OrdAmt").Specific.VALUE = "100000000";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP567_ComboBox_Setting
		/// </summary>
		private void PS_PP567_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//거래처구분
				sQry = " SELECT      U_Minor,";
				sQry += "             U_CdName";
				sQry += " FROM        [@PS_SY001L]";
				sQry += " WHERE       Code = 'C100'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, sQry, "", false, false);
				oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//자체/외주
				oForm.Items.Item("InOut").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("InOut").Specific.ValidValues.Add("IN", "자체");
				oForm.Items.Item("InOut").Specific.ValidValues.Add("OUT", "외주");
				oForm.Items.Item("InOut").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP567_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP567_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "CardCode":  //거래처
						oForm.Items.Item("CardName").Specific.VALUE = dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'", "");
						break;

					case "ItemCode": //작번,규격
						oForm.Items.Item("ItemName").Specific.VALUE = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'", "");
						oForm.Items.Item("ItemSpec").Specific.VALUE = dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'", "");
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP567_MTX01
		/// </summary>
		private void PS_PP567_MTX01()
		{
			string sQry;
			string errMessage = string.Empty;

			string FrYM;	 //기준년월(시작)
			string ToYM;	 //기준년월(종료)
			string CardType; //거래처구분
			string CardCode; //수주처
			string ItemCode; //작번
			string InOut;	 //자체/외주
			double OrdAmt;	 //수주금액
			string CmltYN;	 //생산완료

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			try
			{
				oForm.Freeze(true);

				FrYM     = oForm.Items.Item("FrYM").Specific.Value.ToString().Trim();
				ToYM     = oForm.Items.Item("ToYM").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				InOut    = oForm.Items.Item("InOut").Specific.Selected.Value.ToString().Trim();
				OrdAmt = Convert.ToDouble(oForm.Items.Item("OrdAmt").Specific.Value.ToString().Trim());

				if (oForm.Items.Item("CmltYN").Specific.Checked == true)
				{
					CmltYN = "Y";
				}
				else
				{
					CmltYN = "N";
				}

				ProgressBar01.Text = "조회 중...";

				sQry = " EXEC PS_PP567_01 '";
				sQry += FrYM + "','";
				sQry += ToYM + "','";
				sQry += CardType + "','";
				sQry += CardCode + "','";
				sQry += ItemCode + "','";
				sQry += InOut + "','";
				sQry += OrdAmt + "','";
				sQry += CmltYN + "'";

				oGrid.DataTable.Clear();
				oDS_PS_PP567A.ExecuteQuery(sQry);

				oGrid.Columns.Item(6).RightJustified = true;
				oGrid.Columns.Item(8).RightJustified = true;
				oGrid.Columns.Item(9).RightJustified = true;
				oGrid.Columns.Item(10).RightJustified = true;
				oGrid.Columns.Item(11).RightJustified = true;
				oGrid.Columns.Item(12).RightJustified = true;
				oGrid.Columns.Item(13).RightJustified = true;
				oGrid.Columns.Item(14).RightJustified = true;
				oGrid.Columns.Item(15).RightJustified = true;
				oGrid.Columns.Item(16).RightJustified = true;
				oGrid.Columns.Item(17).RightJustified = true;

				if (oGrid.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oGrid.AutoResizeColumns();
				oForm.Update();
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP567_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_PP567_Print_Report01()
		{
			string WinTitle;
			string ReportName;

			string FrYM;     //기준년월(시작)
			string ToYM;     //기준년월(종료)
			string CardType; //거래처구분
			string CardCode; //수주처
			string ItemCode; //작번
			string InOut;    //자체/외주
			double OrdAmt;   //수주금액
			string CmltYN;   //생산완료

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				FrYM = oForm.Items.Item("FrYM").Specific.Value.ToString().Trim();
				ToYM = oForm.Items.Item("ToYM").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				InOut = oForm.Items.Item("InOut").Specific.Selected.Value.ToString().Trim();
				OrdAmt = Convert.ToDouble(oForm.Items.Item("OrdAmt").Specific.Value.ToString().Trim());

				if (oForm.Items.Item("CmltYN").Specific.Checked == true)
				{
					CmltYN = "Y";
				}
				else
				{
					CmltYN = "N";
				}

				WinTitle = "[PS_PP567] 레포트";
				ReportName = "PS_PP567_01.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@FrYM", FrYM));
				dataPackParameter.Add(new PSH_DataPackClass("@ToYM", ToYM));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@InOut", InOut));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdAmt", OrdAmt));
				dataPackParameter.Add(new PSH_DataPackClass("@CmltYN", CmltYN));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP567_Print_Report02
		/// </summary>
		[STAThread]
		private void PS_PP567_Print_Report02()
		{
			string WinTitle;
			string ReportName;

			string FrYM;     //기준년월(시작)
			string ToYM;     //기준년월(종료)
			string CardType; //거래처구분
			string CardCode; //수주처
			string ItemCode; //작번
			string InOut;    //자체/외주
			double OrdAmt;   //수주금액
			string CmltYN;   //생산완료

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				FrYM = oForm.Items.Item("FrYM").Specific.Value.ToString().Trim();
				ToYM = oForm.Items.Item("ToYM").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Selected.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				InOut = oForm.Items.Item("InOut").Specific.Selected.Value.ToString().Trim();
				OrdAmt = Convert.ToDouble(oForm.Items.Item("OrdAmt").Specific.Value.ToString().Trim());

				if (oForm.Items.Item("CmltYN").Specific.Checked == true)
				{
					CmltYN = "Y";
				}
				else
				{
					CmltYN = "N";
				}

				WinTitle = "[PS_PP567] 레포트";
				ReportName = "PS_PP567_02.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@FrYM", FrYM));
				dataPackParameter.Add(new PSH_DataPackClass("@ToYM", ToYM));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@InOut", InOut));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdAmt", OrdAmt));
				dataPackParameter.Add(new PSH_DataPackClass("@CmltYN", CmltYN));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
							PS_PP567_MTX01();
						}
					}
				}
				else if (pVal.ItemUID == "BtnPrt01")
				{
					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP567_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
				else if (pVal.ItemUID == "BtnPrt02")
				{
					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP567_Print_Report02);
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP567_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
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
						PS_PP567_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP567A);
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
						case "1285": //복원
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
						case "1285": //복원
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
		}

		/// <summary>
		/// Raise_FormDataEvent
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
