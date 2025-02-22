using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 고정자산현황
	/// </summary>
	internal class PS_FX240 : PSH_BaseClass
	{
		private string oFormUniqueID;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FX240.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_FX240_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_FX240");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);
				PS_FX240_CreateItems();
				PS_FX240_ComboBox_Setting();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
		/// PS_FX240_CreateItems
		/// </summary>
		private void PS_FX240_CreateItems()
		{
			try
			{
				oForm.DataSources.UserDataSources.Add("YM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
				oForm.Items.Item("YM").Specific.DataBind.SetBound(true, "", "YM");
				oForm.DataSources.UserDataSources.Item("YM").Value = DateTime.Now.ToString("yyyy-MM");

				oForm.DataSources.UserDataSources.Add("YMF", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 7);
				oForm.Items.Item("YMF").Specific.DataBind.SetBound(true, "", "YMF");
				oForm.DataSources.UserDataSources.Item("YMF").Value = DateTime.Now.ToString("yyyy") + "-01";

				oForm.DataSources.UserDataSources.Add("Rad01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Rad04", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

				oForm.Items.Item("Rad01").Specific.DataBind.SetBound(true, "", "Rad01");
				oForm.Items.Item("Rad02").Specific.DataBind.SetBound(true, "", "Rad02");
				oForm.Items.Item("Rad03").Specific.DataBind.SetBound(true, "", "Rad03");
				oForm.Items.Item("Rad03").Specific.DataBind.SetBound(true, "", "Rad04");

				oForm.Items.Item("Rad01").Specific.ValOn = "10";
				oForm.Items.Item("Rad01").Specific.ValOff = "0";
				oForm.Items.Item("Rad01").Specific.Selected = true;

				oForm.Items.Item("Rad02").Specific.ValOn = "20";
				oForm.Items.Item("Rad02").Specific.ValOff = "0";
				oForm.Items.Item("Rad02").Specific.GroupWith("Rad01");

				oForm.Items.Item("Rad03").Specific.ValOn = "30";
				oForm.Items.Item("Rad03").Specific.ValOff = "0";
				oForm.Items.Item("Rad03").Specific.GroupWith("Rad02");

				oForm.Items.Item("Rad04").Specific.ValOn = "30";
				oForm.Items.Item("Rad04").Specific.ValOff = "0";
				oForm.Items.Item("Rad04").Specific.GroupWith("Rad03");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX240_ComboBox_Setting
		/// </summary>
		private void PS_FX240_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 이력구분  1: 전체, 2:폐기매각 제외, 3:폐기매각, 매각대기 제외
				oForm.Items.Item("Div").Specific.ValidValues.Add("1", "전체");
				oForm.Items.Item("Div").Specific.ValidValues.Add("2", "폐기매각제외");
				oForm.Items.Item("Div").Specific.ValidValues.Add("3", "폐기매각,매각대기제외");
				oForm.Items.Item("Div").Specific.ValidValues.Add("4", "폐기매각,매각대기만");
				oForm.Items.Item("Div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				// 출력구분  0: 전체, 1:메인장비로 집계
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("0", "전체");
				oForm.Items.Item("Gubun").Specific.ValidValues.Add("1", "메인장비로집계");
				oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_FX240_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_FX240_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("YM").Specific.Value.ToString().Trim()))
				{
					errMessage = "상각년월는 필수사항입니다. 입력하여 주십시오.";
					throw new Exception();
				}
				if (oForm.Items.Item("YM").Specific.Value.ToString().Trim().Length != 7)
				{
					errMessage = "상각년월는 7자리 입니다(YYYY-MM). 확인하여 주십시오.";
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_FX240_Print_Query
		/// </summary>
		[STAThread]
		private void PS_FX240_Print_Query()
		{
			string WinTitle = string.Empty;
			string ReportName = string.Empty;
			string BPLId;
			string YM;
			string YMf;
			string Div;
			string Gubun;
			string BPLName;
			string sQry;
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				YM = codeHelpClass.Left(oForm.Items.Item("YM").Specific.Value.ToString().Trim(), 4) + codeHelpClass.Right(oForm.Items.Item("YM").Specific.Value.ToString().Trim(), 2);
				YMf = codeHelpClass.Left(oForm.Items.Item("YMF").Specific.Value.ToString().Trim(), 4) + codeHelpClass.Right(oForm.Items.Item("YMF").Specific.Value.ToString().Trim(), 2);
				Div = oForm.Items.Item("Div").Specific.Selected.Value.ToString().Trim();
				Gubun = oForm.Items.Item("Gubun").Specific.Selected.Value.ToString().Trim();

				sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + BPLId + "'";
				oRecordSet.DoQuery(sQry);
				BPLName = oRecordSet.Fields.Item(0).Value.ToString().Trim();

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@BPLId", BPLName));
				dataPackFormula.Add(new PSH_DataPackClass("@YM", YMf.Substring(0, 4) + "-" + YMf.Substring(4, 2) + " ~ " + YM.Substring(0, 4) + "-" + YM.Substring(4, 2)));

				if (oForm.Items.Item("Rad01").Specific.Selected == true)
				{
					WinTitle = "고정자산현황 [PS_FX240_01]";
					ReportName = "PS_FX240_01.RPT";
					// Parameter
					dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
					dataPackParameter.Add(new PSH_DataPackClass("@YM", YM));
					dataPackParameter.Add(new PSH_DataPackClass("@Div", Div));
					dataPackParameter.Add(new PSH_DataPackClass("@Gubun", Gubun));
				}
				else if (oForm.Items.Item("Rad02").Specific.Selected == true)
				{
					WinTitle = "고정자산현황 [PS_FX240_02]";
					ReportName = "PS_FX240_02.RPT";
					// Parameter
					dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
					dataPackParameter.Add(new PSH_DataPackClass("@YM", YM));
					dataPackParameter.Add(new PSH_DataPackClass("@Div", Div));
					dataPackParameter.Add(new PSH_DataPackClass("@Gubun", Gubun));
				}
				else if (oForm.Items.Item("Rad03").Specific.Selected == true)
				{
					WinTitle = "고정자산현황 [PS_FX240_03]";
					ReportName = "PS_FX240_03.RPT";
					// Parameter
					dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
					dataPackParameter.Add(new PSH_DataPackClass("@YM", YM));
					dataPackParameter.Add(new PSH_DataPackClass("@Div", Div));
					dataPackParameter.Add(new PSH_DataPackClass("@Gubun", Gubun));
				}
				else if (oForm.Items.Item("Rad04").Specific.Selected == true)
				{
					WinTitle = "고정자산현황 [PS_FX240_04]";
					ReportName = "PS_FX240_04.RPT";
					// Parameter
					dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
					dataPackParameter.Add(new PSH_DataPackClass("@YMf", YMf));
					dataPackParameter.Add(new PSH_DataPackClass("@YM", YM));
				}

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
				//case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
				//    Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
				//    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
					//    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
					//    break;
					//case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
					//    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
					//    break;
					//case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
					//    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
					//    break;
					//case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
					//    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
					//    break;
					//case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
					//    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
					//    break;
					//case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
					//    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
					//    break;
					//case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
					//    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
					//    break;
					//case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
					//    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
					//    break;
					//case SAPbouiCOM.BoEventTypes.et_Drag: //39
					//    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
					//    break;
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
						if (PS_FX240_HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_FX240_Print_Query);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (oForm.Items.Item("Rad01").Specific.Selected == true)
					{
						oForm.Items.Item("YMF").Visible = false;
					}
					else if (oForm.Items.Item("Rad02").Specific.Selected == true)
					{
						oForm.Items.Item("YMF").Visible = false;
					}
					else if (oForm.Items.Item("Rad03").Specific.Selected == true)
					{
						oForm.Items.Item("YMF").Visible = false;
					}
					else if (oForm.Items.Item("Rad04").Specific.Selected == true)
					{
						oForm.Items.Item("YMF").Visible = true;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_MenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_MenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1283": //삭제
							break;
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1283": //삭제
							break;
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1287": // 복제
							break;
						case "7169": //엑셀 내보내기
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Freeze(false);
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
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
