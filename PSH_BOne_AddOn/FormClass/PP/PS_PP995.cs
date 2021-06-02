using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 분말LABEL출력
	/// </summary>
	internal class PS_PP995 : PSH_BaseClass
	{
		private string oFormUniqueID;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP995.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP995_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP995");

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
				oForm.DataSources.UserDataSources.Add("MDate", SAPbouiCOM.BoDataType.dt_DATE, 10);
				oForm.Items.Item("MDate").Specific.DataBind.SetBound(true, "", "MDate");
				oForm.DataSources.UserDataSources.Item("MDate").Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("Title", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("Title").Specific.DataBind.SetBound(true, "", "Title");
				oForm.DataSources.UserDataSources.Item("Title").Value = "Copper Powder";
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
			try
			{
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("1", "Cu분말");
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("2", "청동분말");
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("3", "주석분말");
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("9", "기타(입력)");
				oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Print_Query
		/// </summary>
		[STAThread]
		private void Print_Query()
		{
			string WinTitle;
			string ReportName;

			string Gubun;
			string ItemName;
			string Gwgt;
			string Nwgt;
			string MDate;
			string Lot;
			string Title;

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				Gubun    = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();
				ItemName = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				Gwgt     = oForm.Items.Item("Gwgt").Specific.Value.ToString().Trim();
				Nwgt     = oForm.Items.Item("Nwgt").Specific.Value.ToString().Trim();
				MDate    = oForm.Items.Item("MDate").Specific.Value.ToString().Trim();
				Lot      = oForm.Items.Item("Lot").Specific.Value.ToString().Trim();
				Title    = oForm.Items.Item("Title").Specific.Value.ToString().Trim();

				WinTitle = "분말LABEL출력[PS_PP995_01]";
				ReportName = "PS_PP995_01.RPT";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@Gubun", Gubun));
				dataPackFormula.Add(new PSH_DataPackClass("@ItemName", ItemName));
				dataPackFormula.Add(new PSH_DataPackClass("@Gwgt", Gwgt));
				dataPackFormula.Add(new PSH_DataPackClass("@Nwgt", Nwgt));
				dataPackFormula.Add(new PSH_DataPackClass("@MDate", MDate.Substring(0, 4) + "-" + MDate.Substring(4, 2) + "-" + MDate.Substring(6, 2)));
				dataPackFormula.Add(new PSH_DataPackClass("@Lot", Lot));
				dataPackFormula.Add(new PSH_DataPackClass("@Title", Title));
				dataPackFormula.Add(new PSH_DataPackClass("@BarCode", Lot));

				// Parameter
				// Qry는 필요없는데 에러로안해 그냥하나 만듬
				dataPackParameter.Add(new PSH_DataPackClass("@Gubun", Gubun));

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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string Gubun;

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					Gubun = oForm.Items.Item("Gubun").Specific.Value.ToString().Trim();
					if (Gubun == "1")
					{
						oForm.DataSources.UserDataSources.Item("Title").Value = "Copper Powder";
					}
					else if (Gubun == "2")
					{
						oForm.DataSources.UserDataSources.Item("Title").Value = "혼 합  청 동 분";
					}
					else if (Gubun == "3")
					{
						oForm.DataSources.UserDataSources.Item("Title").Value = "Sn Powder";
					}
					else
					{
						oForm.DataSources.UserDataSources.Item("Title").Value = "";
					}
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
