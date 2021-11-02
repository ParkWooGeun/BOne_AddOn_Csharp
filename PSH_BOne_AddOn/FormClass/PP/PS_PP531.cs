using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 재고자산 월령 조회
	/// </summary>
	internal class PS_PP531 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DataTable oDS_PS_PP531A;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP531.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP531_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP531");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP531_CreateItems();
				PS_PP531_ComboBox_Setting();
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
		/// PS_PP531_CreateItems
		/// </summary>
		private void PS_PP531_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("PS_PP531A");
				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_PP531A");
				oDS_PS_PP531A = oForm.DataSources.DataTables.Item("PS_PP531A");

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM").Specific.DataBind.SetBound(true, "", "StdYM");
				oForm.Items.Item("StdYM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");

				//제품구분
				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemClass", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemClass").Specific.DataBind.SetBound(true, "", "ItemClass");

				//조회구분
				oForm.DataSources.UserDataSources.Add("SrchType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("SrchType").Specific.DataBind.SetBound(true, "", "SrchType");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP531_ComboBox_Setting
		/// </summary>
		private void PS_PP531_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", dataHelpClass.User_BPLID(), false, false);

				//제품구분
				sQry = " SELECT  U_Minor,";
				sQry += "         U_CdName";
				sQry += " FROM    [@PS_SY001L]";
				sQry += " WHERE   Code = 'P221'";
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, sQry, "%", false, false);

				//품목구분
				sQry = " SELECT  U_Minor,";
				sQry += "         U_CdName";
				sQry += " FROM    [@PS_SY001L]";
				sQry += " WHERE   Code = 'P222'";
				oForm.Items.Item("ItemClass").Specific.ValidValues.Add("%", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemClass").Specific, sQry, "%", false, false);

				//조회구분
				sQry = " SELECT  U_Minor,";
				sQry += "         U_CdName";
				sQry += " FROM    [@PS_SY001L]";
				sQry += " WHERE   Code = 'P223'";
				dataHelpClass.Set_ComboList(oForm.Items.Item("SrchType").Specific, sQry, "01", false, false);

				//조회결과
				sQry = " SELECT  U_Minor,";
				sQry += "         U_CdName";
				sQry += " FROM    [@PS_SY001L]";
				sQry += " WHERE   Code = 'P224'";
				dataHelpClass.Set_ComboList(oForm.Items.Item("SrchRslt").Specific, sQry, "01", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP531_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_PP531_DataValidCheck()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				//제품구분
				if (oForm.Items.Item("ItemType").Specific.Selected.VALUE == "%")
				{
					errMessage = "제품구분을 선택하세요.";
					throw new Exception();
				}

				//품목구분
				if (oForm.Items.Item("ItemClass").Specific.Selected.VALUE == "%")
				{
					errMessage = "재고구분을 선택하세요.";
					throw new Exception();
				}
				functionReturnValue = true;
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
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP531_FormResize
		/// </summary>
		private void PS_PP531_FormResize()
		{
			try
			{
				if (oGrid.Columns.Count > 0)
				{
					oGrid.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP531_MTX01
		/// </summary>
		private void PS_PP531_MTX01()
		{
			string BPLId;       //사업장
			string StdYM;       //기준년월
			string ItemType;    //제품구분
			string ItemClass;   //재고구분
			string SrchType;    //조회구분
			string SrchRslt;    //조회결과
			string sQry = string.Empty;
			string errMessage = string.Empty;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Selected.Value.ToString().Trim();
				ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
				SrchType = oForm.Items.Item("SrchType").Specific.Selected.Value.ToString().Trim();
				SrchRslt = oForm.Items.Item("SrchRslt").Specific.Selected.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중...";

				//월령조회
				if (SrchType == "01")
				{
					sQry = " EXEC PS_PP531_01 ";
					sQry += "'" + BPLId + "',";     //사업장
					sQry += "'" + StdYM + "',";     //기준년월
					sQry += "'" + ItemType + "',";  //제품구분
					sQry += "'" + ItemClass + "',"; //재고구분
					sQry += "'" + SrchRslt + "'";   //조회결과
				}

				//기초조회
				if (SrchType == "02")
				{
					sQry = " EXEC PS_PP531_02 ";
					sQry += "'" + BPLId + "',";     //사업장
					sQry += "'" + StdYM + "',";     //기준년월
					sQry += "'" + ItemType + "',";  //제품구분
					sQry += "'" + ItemClass + "',"; //재고구분
					sQry += "'" + SrchRslt + "'";   //조회결과
				}

				oGrid.DataTable.Clear();
				oDS_PS_PP531A.ExecuteQuery(sQry);

				//필드정렬(조회구분(월령))
				if (SrchType == "01")
				{
					oGrid.Columns.Item(6).RightJustified = true;
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
				}

				//필드정렬(조회구분(기초))
				if (SrchType == "02")
				{
					oGrid.Columns.Item(5).RightJustified = true;
					oGrid.Columns.Item(6).RightJustified = true;
					oGrid.Columns.Item(9).RightJustified = true;
				}

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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP531_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_PP531_Print_Report01()
		{
			string WinTitle = null;
			string ReportName = null;

			string BPLId;       //사업장
			string StdYM;       //기준년월
			string ItemType;    //제품구분
			string ItemClass;   //재고구분
			string SrchType;    //조회구분
			string SrchRslt;    //조회결과

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Selected.Value.ToString().Trim();
				ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
				SrchType = oForm.Items.Item("SrchType").Specific.Selected.Value.ToString().Trim();
				SrchRslt = oForm.Items.Item("SrchRslt").Specific.Selected.Value.ToString().Trim();

				//월령조회
				if (SrchType == "01")
				{
					WinTitle = "[PS_PP531] 월령조회";
					ReportName = "PS_PP531_01.rpt";
				}

				//기초조회
				if (SrchType == "02")
				{
					WinTitle = "[PS_PP531] 기초조회";
					ReportName = "PS_PP531_02.rpt";
				}

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemType", ItemType));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemClass", ItemClass));
				dataPackParameter.Add(new PSH_DataPackClass("@SrchRslt", SrchRslt));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                   // Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    //Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                    //Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                    //Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
							if (PS_PP531_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							else
							{
								PS_PP531_MTX01();
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}

					if (pVal.ItemUID == "BtnPrint")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP531_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							else
							{
								System.Threading.Thread thread = new System.Threading.Thread(PS_PP531_Print_Report01);
								thread.SetApartmentState(System.Threading.ApartmentState.STA);
								thread.Start();
							}
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
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP531_FormResize();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_CHOOSE_FROM_LIST
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					oForm.Update();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_GOT_FOCUS
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP531A);
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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
				if ((BusinessObjectInfo.BeforeAction == true))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:   //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:    //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
							break;
					}
				}
				else if ((BusinessObjectInfo.BeforeAction == false))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:   //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:    //34
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
		}
	}
}
