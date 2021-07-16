using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 월생산계획및원자재공급현황
	/// </summary>
	internal class PS_PP220 : PSH_BaseClass
	{
		private string oFormUniqueID;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP220.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP220_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP220");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);
				CreateItems();
				ComboBox_Setting();

				oForm.EnableMenu(("1283"), false); // 삭제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1287"), false); // 복제
				oForm.EnableMenu(("1284"), true);  // 취소
				oForm.EnableMenu(("1293"), false); // 행삭제
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
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
			try
			{
				oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");
				oForm.DataSources.UserDataSources.Item("YM").Value = DateTime.Now.ToString("yyyyMM");

				oForm.DataSources.UserDataSources.Add("Rad01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad04", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad06", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad07", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad08", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad09", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad10", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

				oForm.Items.Item("Rad01").Specific.DataBind.SetBound(true, "", "Rad01");
				oForm.Items.Item("Rad02").Specific.DataBind.SetBound(true, "", "Rad02");
				oForm.Items.Item("Rad03").Specific.DataBind.SetBound(true, "", "Rad03");
				oForm.Items.Item("Rad04").Specific.DataBind.SetBound(true, "", "Rad04");
				oForm.Items.Item("Rad05").Specific.DataBind.SetBound(true, "", "Rad05");
				oForm.Items.Item("Rad06").Specific.DataBind.SetBound(true, "", "Rad06");
				oForm.Items.Item("Rad07").Specific.DataBind.SetBound(true, "", "Rad07");
				oForm.Items.Item("Rad08").Specific.DataBind.SetBound(true, "", "Rad08");
				oForm.Items.Item("Rad09").Specific.DataBind.SetBound(true, "", "Rad09");
				oForm.Items.Item("Rad10").Specific.DataBind.SetBound(true, "", "Rad10");

				oForm.Items.Item("Rad01").Specific.ValOn = "10";
				oForm.Items.Item("Rad01").Specific.ValOff = "0";
				oForm.Items.Item("Rad01").Specific.Selected = true;

				oForm.Items.Item("Rad02").Specific.ValOn = "20";
				oForm.Items.Item("Rad02").Specific.ValOff = "0";
				oForm.Items.Item("Rad02").Specific.GroupWith(("Rad01"));

				oForm.Items.Item("Rad03").Specific.ValOn = "30";
				oForm.Items.Item("Rad03").Specific.ValOff = "0";
				oForm.Items.Item("Rad03").Specific.GroupWith(("Rad01"));

				oForm.Items.Item("Rad04").Specific.ValOn = "40";
				oForm.Items.Item("Rad04").Specific.ValOff = "0";
				oForm.Items.Item("Rad04").Specific.GroupWith(("Rad01"));

				oForm.Items.Item("Rad05").Specific.ValOn = "50";
				oForm.Items.Item("Rad05").Specific.ValOff = "0";
				oForm.Items.Item("Rad05").Specific.GroupWith(("Rad01"));

				oForm.Items.Item("Rad06").Specific.ValOn = "60";
				oForm.Items.Item("Rad06").Specific.ValOff = "0";
				oForm.Items.Item("Rad06").Specific.GroupWith(("Rad01"));

				oForm.Items.Item("Rad07").Specific.ValOn = "70";
				oForm.Items.Item("Rad07").Specific.ValOff = "0";
				oForm.Items.Item("Rad07").Specific.GroupWith(("Rad01"));

				oForm.Items.Item("Rad08").Specific.ValOn = "80";
				oForm.Items.Item("Rad08").Specific.ValOff = "0";
				oForm.Items.Item("Rad08").Specific.GroupWith(("Rad01"));

				oForm.Items.Item("Rad09").Specific.ValOn = "90";
				oForm.Items.Item("Rad09").Specific.ValOff = "0";
				oForm.Items.Item("Rad09").Specific.GroupWith(("Rad01"));

				oForm.Items.Item("Rad10").Specific.ValOn = "95";
				oForm.Items.Item("Rad10").Specific.ValOff = "0";
				oForm.Items.Item("Rad10").Specific.GroupWith(("Rad01"));
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		private void ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("%", "전체");
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
		/// Print_Query
		/// </summary>
		[STAThread]
		private void Print_Query()
		{
			string WinTitle = string.Empty;
			string ReportName = string.Empty;

			string BPLId;
			string YM;

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				YM = oForm.Items.Item("YM").Specific.Value.ToString().Trim();

				if (oForm.Items.Item("Rad01").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_01] 월생산계획(거래처순)";
					ReportName = "PS_PP220_01.RPT";
				}
				else if (oForm.Items.Item("Rad02").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_03] 월생산계획(원자재순)";
					ReportName = "PS_PP220_03.RPT";
				}
				else if (oForm.Items.Item("Rad03").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_02] 원자재공급현황";
					ReportName = "PS_PP220_02.RPT";
				}
				else if (oForm.Items.Item("Rad04").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_04] 생산계획대비생산현황";
					ReportName = "PS_PP220_04.RPT";
				}
				else if (oForm.Items.Item("Rad05").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_05] 생산계획대비출고현황";
					ReportName = "PS_PP220_05.RPT";
				}
				else if (oForm.Items.Item("Rad06").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_06] 생산계획대비생산현황(일자별)";
					ReportName = "PS_PP220_06.RPT";
				}
				else if (oForm.Items.Item("Rad07").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_07] 월생산계획(거래처순)-창원";
					ReportName = "PS_PP220_07.RPT";
				}
				else if (oForm.Items.Item("Rad08").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_07] 월생산계획(원자재순)-창원";
					ReportName = "PS_PP220_07.RPT";
				}
				else if (oForm.Items.Item("Rad09").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_08] 생산계획대비생산현황(타입별)";
					ReportName = "PS_PP220_08.RPT";
				}
				else if (oForm.Items.Item("Rad10").Specific.Selected == true)
				{
					WinTitle = "[PS_PP220_09] 생산계획대비출고현황(타입별)";
					ReportName = "PS_PP220_09.RPT";
				}

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@YM", YM.Substring(0, 4) + "-" + YM.Substring(4, 2)));

				// Parameter
				if (oForm.Items.Item("Rad07").Specific.Selected == true)
				{
					dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
					dataPackParameter.Add(new PSH_DataPackClass("@YM", YM));
					dataPackParameter.Add(new PSH_DataPackClass("@Gubun", "1"));
				}
				else if (oForm.Items.Item("Rad08").Specific.Selected == true)
				{
					dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
					dataPackParameter.Add(new PSH_DataPackClass("@YM", YM));
					dataPackParameter.Add(new PSH_DataPackClass("@Gubun", "2"));
				}
				else
                {
					dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
					dataPackParameter.Add(new PSH_DataPackClass("@YM", YM));
				}

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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
					}
					else if (pVal.ItemUID == "Btn01")
					{
						System.Threading.Thread thread = new System.Threading.Thread(Print_Query);
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
