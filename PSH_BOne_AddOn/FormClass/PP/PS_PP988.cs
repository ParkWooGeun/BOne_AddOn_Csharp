using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작번별 비용명세[결산/계산/완료]
	/// </summary>
	internal class PS_PP988 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;

		/// <summary>
		/// 화면 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{

			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP988.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP988_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP988");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP988_CreateItems();
				PS_PP988_SetComboBox();
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
		/// PS_PP988_CreateItems
		/// </summary>
		private void PS_PP988_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");
				oForm.Items.Item("FrDt").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");
				oForm.Items.Item("ToDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

				//품목코드
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				//팀
				oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

				//담당
				oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

				//반
				oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

				//완료금액구분
				oForm.DataSources.UserDataSources.Add("WkCltCls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("WkCltCls").Specific.DataBind.SetBound(true, "", "WkCltCls");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP988_SetComboBox
		/// </summary>
		private void PS_PP988_SetComboBox()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//품목구분
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
				sQry = "  SELECT     U_Minor AS [Code], ";
				sQry += "               U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'S002'";
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, sQry, "", false, false);
				oForm.Items.Item("ItemType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//거래처구분
				oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
				sQry = " SELECT     U_Minor AS [Code], ";
				sQry += "               U_CdName AS [Name]";
				sQry += " FROM      [@PS_SY001L]";
				sQry += " WHERE     Code = 'C100'";
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, sQry, "", false, false);
				oForm.Items.Item("CardType").Specific.Select("%", SAPbouiCOM.BoSearchKey.psk_ByValue);

				//팀
				oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
				sQry = "   SELECT      U_Code AS [Code],";
				sQry += "                 U_CodeNm As [Name]";
				sQry += "  FROM       [@PS_HR200L]";
				sQry += "  WHERE      Code = '1'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "                 AND U_Char2 = '2'";	//부산사업장(2)으로 고정
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
				oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//완료금액 구분
				oForm.Items.Item("WkCltCls").Specific.ValidValues.Add("01", "완료시점금액");
				oForm.Items.Item("WkCltCls").Specific.ValidValues.Add("02", "누계금액");
				oForm.Items.Item("WkCltCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP988_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP988_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i;
			string sQry;

			string TeamCode;
			string RspCode;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "TeamCode":
						TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//담당콤보세팅
						oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
						sQry =  " SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '2'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char1 = '" + TeamCode + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
						oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "RspCode":
						TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
						RspCode  = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						//반콤보세팅
						oForm.Items.Item("ClsCode").Specific.ValidValues.Add("%", "전체");
						sQry =  " SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '9'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char1 = '" + RspCode + "'";
						sQry += "                 AND U_Char2 = '" + TeamCode + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode").Specific, sQry, "", false, false);
						oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "ItemCode":
						sQry = "  SELECT       FrgnName";
						sQry += " FROM        OITM";
						sQry = sQry + " WHERE       ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item("FrgnName").Value.ToString().Trim();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP988_MTX01
		/// </summary>
		private void PS_PP988_MTX01()
		{
			string sQry;
			string errMessage = string.Empty;
			string FrDt;
			string ToDt;
			string CardType;
			string ItemType;
			string ItemCode;
			string TeamCode;
			string RspCode;
			string ClsCode;
			string WkCltCls;
			string CntcCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				CntcCode = dataHelpClass.User_MSTCOD();

				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();
				WkCltCls = oForm.Items.Item("WkCltCls").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC PS_PP988_91 '";
				sQry += CntcCode + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CardType + "','";
				sQry += ItemType + "','";
				sQry += ItemCode + "','";
				sQry += TeamCode + "','";
				sQry += RspCode + "','";
				sQry += ClsCode + "','";
				sQry += WkCltCls + "'";

				oGrid.DataTable.Clear();
				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

				oGrid.Columns.Item(7).RightJustified = true;
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
				oGrid.Columns.Item(18).RightJustified = true;
				oGrid.Columns.Item(19).RightJustified = true;
				oGrid.Columns.Item(20).RightJustified = true;
				oGrid.Columns.Item(21).RightJustified = true;
				oGrid.Columns.Item(22).RightJustified = true;
				oGrid.Columns.Item(23).RightJustified = true;
				oGrid.Columns.Item(24).RightJustified = true;
				oGrid.Columns.Item(25).RightJustified = true;
				oGrid.Columns.Item(26).RightJustified = true;
				oGrid.Columns.Item(27).RightJustified = true;
				oGrid.Columns.Item(28).RightJustified = true;
				oGrid.Columns.Item(29).RightJustified = true;
				oGrid.Columns.Item(30).RightJustified = true;
				oGrid.Columns.Item(32).RightJustified = true;
				oGrid.Columns.Item(33).RightJustified = true;
				oGrid.Columns.Item(34).RightJustified = true;
				oGrid.Columns.Item(35).RightJustified = true;

				oGrid.Columns.Item(13).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //[결산]계, 노랑
				oGrid.Columns.Item(21).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //[계산]계, 노랑
				oGrid.Columns.Item(29).BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 125)); //[완료]계, 노랑

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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP988_PrintReport
		/// </summary>
		[STAThread]
		private void PS_PP988_PrintReport()
		{
			string WinTitle;
			string ReportName;
			string FrDt;
			string ToDt;
			string CardType;
			string ItemType;
			string ItemCode;
			string TeamCode;
			string RspCode;
			string ClsCode;
			string CntcCode;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				CntcCode = dataHelpClass.User_MSTCOD();

				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP988] 레포트";
				ReportName = "PS_PP988_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@CardType", CardType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
				dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
		/// <param name="PVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent PVal, ref bool BubbleEvent)
		{
			try
			{
				if (PVal.BeforeAction == true)
				{
					if (PVal.ItemUID == "BtnSearch")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP988_MTX01();
						}
					}
					else if (PVal.ItemUID == "BtnPrint")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP988_PrintReport);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (PVal.BeforeAction == false)
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
		/// <param name="PVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent PVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (PVal.BeforeAction == true)
				{
				}
				else if (PVal.BeforeAction == false)
				{
					PS_PP988_FlushToItemValue(PVal.ItemUID, PVal.Row, PVal.ColUID);
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
		/// <param name="PVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent PVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (PVal.BeforeAction == true)
				{
				}
				else if (PVal.BeforeAction == false)
				{
					PS_PP988_FlushToItemValue(PVal.ItemUID, PVal.Row, PVal.ColUID);
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
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
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
