using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn.Core
{
	/// <summary>
	/// 분개장문서 사용안함.
	/// </summary>
	internal class S393 : PSH_BaseClass
	{
		private SAPbouiCOM.Matrix oMat;
		private int oMatRow;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="formUID"></param>
		public override void LoadForm(string formUID)
		{
			try
			{
				oForm = PSH_Globals.SBO_Application.Forms.Item(formUID);
				oForm.Freeze(true);
				SubMain.Add_Forms(this, formUID, "S393");
				oMat = oForm.Items.Item("76").Specific;
				S393_CreateItems();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
			}
		}

		private void S393_CreateItems()
		{
			SAPbouiCOM.Item newItem = null;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//회계전표 버튼
				newItem = oForm.Items.Add("Btn01", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 90;
				newItem.Top = oForm.Items.Item("2").Top;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 100;
				newItem.Specific.Caption = "회계 전표";

				//전결 라디오 버튼
				newItem = oForm.Items.Add("RBtn01", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 200;
				newItem.Top = oForm.Items.Item("2").Top - 8;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 50;
				newItem.Specific.Caption = "담당";

				//전결 라디오 버튼
				newItem = oForm.Items.Add("RBtn02", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 200;
				newItem.Top = oForm.Items.Item("2").Top + 10;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 50;
				newItem.Specific.Caption = "차장";

				//전결 라디오 버튼
				newItem = oForm.Items.Add("RBtn03", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 250;
				newItem.Top = oForm.Items.Item("2").Top - 8;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 70;
				newItem.Specific.Caption = "팀장";

				//전결 라디오 버튼
				newItem = oForm.Items.Add("RBtn04", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 250;
				newItem.Top = oForm.Items.Item("2").Top + 10;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 70;
				newItem.Specific.Caption = "사업부장";

				//전결 라디오 버튼
				newItem = oForm.Items.Add("RBtn05", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 325;
				newItem.Top = oForm.Items.Item("2").Top - 8;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 70;
				newItem.Specific.Caption = "전무";

				//전결 라디오 버튼
				newItem = oForm.Items.Add("RBtn06", SAPbouiCOM.BoFormItemTypes.it_OPTION_BUTTON);
				newItem.Left = oForm.Items.Item("2").Left + 325;
				newItem.Top = oForm.Items.Item("2").Top + 10;
				newItem.Height = oForm.Items.Item("2").Height;
				newItem.Width = 70;
				newItem.Specific.Caption = "사장";

				//라디오버튼
				oForm.DataSources.UserDataSources.Add("RadioBtn01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("RBtn01").Specific.ValOn = "1";
				oForm.Items.Item("RBtn01").Specific.ValOff = "0";
				oForm.Items.Item("RBtn01").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RBtn01").Specific.Selected = true;

				oForm.Items.Item("RBtn02").Specific.ValOn = "2";
				oForm.Items.Item("RBtn02").Specific.ValOff = "0";
				oForm.Items.Item("RBtn02").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RBtn02").Specific.GroupWith(("RBtn01"));

				oForm.Items.Item("RBtn03").Specific.ValOn = "3";
				oForm.Items.Item("RBtn03").Specific.ValOff = "0";
				oForm.Items.Item("RBtn03").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RBtn03").Specific.GroupWith(("RBtn01"));

				oForm.Items.Item("RBtn04").Specific.ValOn = "4";
				oForm.Items.Item("RBtn04").Specific.ValOff = "0";
				oForm.Items.Item("RBtn04").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RBtn04").Specific.GroupWith(("RBtn01"));

				oForm.Items.Item("RBtn05").Specific.ValOn = "5";
				oForm.Items.Item("RBtn05").Specific.ValOff = "0";
				oForm.Items.Item("RBtn05").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RBtn05").Specific.GroupWith(("RBtn01"));

				oForm.Items.Item("RBtn06").Specific.ValOn = "6";
				oForm.Items.Item("RBtn06").Specific.ValOff = "0";
				oForm.Items.Item("RBtn06").Specific.DataBind.SetBound(true, "", "RadioBtn01");
				oForm.Items.Item("RBtn06").Specific.GroupWith(("RBtn01"));

				oForm.DataSources.UserDataSources.Item("RadioBtn01").Value = "0";

				newItem = oForm.Items.Add("Static01", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				newItem.Left = oForm.Items.Item("2006").Left + 93;
				newItem.Top = oForm.Items.Item("2006").Top;
				newItem.Height = oForm.Items.Item("2006").Height;
				newItem.Width = oForm.Items.Item("2006").Width;
				newItem.Specific.Caption = "사업장";

				newItem = oForm.Items.Add("BPLId01", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
				newItem.Left = oForm.Items.Item("2007").Left + 93;
				newItem.Top = oForm.Items.Item("2007").Top;
				newItem.Height = oForm.Items.Item("2007").Height;
				newItem.Width = oForm.Items.Item("2007").Width + 40;
				newItem.DisplayDesc = true;
				newItem.Specific.DataBind.SetBound(true, "OBTF", "U_BPLId");

				sQry = "SELECT BPLId, BPLName From [OBPL] Order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					newItem.Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				newItem.Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//사업장-ComboBox
				newItem = oForm.Items.Add("Static02", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				newItem.Left = oForm.Items.Item("2001").Left + 161;
				newItem.Top = oForm.Items.Item("2001").Top;
				newItem.Height = oForm.Items.Item("2001").Height;
				newItem.Width = oForm.Items.Item("2001").Width;
				newItem.FromPane = 2;
				newItem.ToPane = 2;
				newItem.Specific.Caption = "사업장";

				newItem = oForm.Items.Add("BPLId02", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
				newItem.Left = oForm.Items.Item("2000").Left + 161;
				newItem.Top = oForm.Items.Item("2000").Top;
				newItem.Height = oForm.Items.Item("2000").Height;
				newItem.Width = oForm.Items.Item("2000").Width;
				newItem.FromPane = 2;
				newItem.ToPane = 2;
				newItem.DisplayDesc = true;
				newItem.Specific.DataBind.SetBound(true, "BTF1", "U_BPLId");

				newItem = oForm.Items.Add("AddonText", SAPbouiCOM.BoFormItemTypes.it_STATIC);
				newItem.Top = oForm.Items.Item("1").Top - 12;
				newItem.Left = oForm.Items.Item("1").Left;
				newItem.Height = 12;
				newItem.Width = 70;
				newItem.FontSize = 10;
				newItem.Specific.Caption = "Addon running";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(newItem);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// S393_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void S393_FlushToItemValue(string oUID, int oRow = 0, string oCol = "")
		{
			int i;

			try
			{
				oForm.Freeze(true);
				switch (oUID)
				{
					case "BPLId02":
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							if (oMatRow == i)
							{
								if (!string.IsNullOrEmpty(oMat.Columns.Item("1").Cells.Item(i).Specific.Value.ToString().Trim()))
								{
									oMat.Columns.Item("U_BPLId").Cells.Item(i).Specific.Select(oForm.Items.Item("BPLId02").Specific.Selected.Value.ToString().Trim());
								}
							}
						}
						break;
				}

				if (oUID == "76")
				{
					switch (oCol)
					{
						case "U_BPLId":
							oForm.Items.Item("BPLId02").Specific.Select(oMat.Columns.Item("U_BPLId").Cells.Item(oRow).Specific.Value.ToString().Trim());
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
		/// S393_Form_Resize
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void S393_Form_Resize(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Items.Item("Static01").Left = oForm.Items.Item("2006").Left + 93;
				oForm.Items.Item("BPLId01").Left =  oForm.Items.Item("2007").Left + 93;
				oForm.Items.Item("Static02").Top =  oForm.Items.Item("2001").Top;
				oForm.Items.Item("Static02").Left = oForm.Items.Item("2001").Left + 161;
				oForm.Items.Item("BPLId02").Top =   oForm.Items.Item("2000").Top;
				oForm.Items.Item("BPLId02").Left =  oForm.Items.Item("2000").Left + 161;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// S393_Print_Report01
		/// </summary>
		[STAThread]
		private void S393_Print_Report01()
		{
			string TransId;
			string WinTitle;
			string ReportName;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
			
			try
			{
				TransId = oForm.Items.Item("5").Specific.Value.ToString().Trim();

				WinTitle = "회계전표 [PS_FI010]";
				ReportName = "PS_FI010_03.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				//Formula List
				dataPackFormula.Add(new PSH_DataPackClass("@RadioBtn01", oForm.DataSources.UserDataSources.Item("RadioBtn01").Value.ToString().Trim()));

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@TransId", TransId));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				//	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
					Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
				//	Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//	Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
				//    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
				//    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
				//    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
				//    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
				//    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
					Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
				//    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
				//    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
				//    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
				//    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
				//    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_Drag: //39
				//    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
				if (pVal.ItemChanged == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemUID == "Btn01")
					{
						System.Threading.Thread thread = new System.Threading.Thread(S393_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
					else if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
						{
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1291");
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
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
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "BPLId02")
						{
							S393_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						if (pVal.ItemUID == "76" && pVal.ColUID == "U_BPLId")
						{
							S393_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
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
		/// CLICK 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.Before_Action == true)
				{
					if (pVal.ItemUID == "76")
					{
						oMatRow = pVal.Row;
					}
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemUID == "76" && pVal.ColUID == "U_BPLId")
					{
						if (oMat.VisualRowCount > 1 && !string.IsNullOrEmpty(oMat.Columns.Item("U_BPLId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
						{
							oForm.Items.Item("BPLId02").Specific.Select(oMat.Columns.Item("U_BPLId").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
						}
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
					S393_Form_Resize(FormUID, ref pVal, ref BubbleEvent);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_RightClickEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="eventInfo"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
		{
			try
			{
				if (eventInfo.BeforeAction == true)
				{
					if (eventInfo.ItemUID == "76")
					{
						if (eventInfo.Row > 0)
						{
							oMatRow = eventInfo.Row;
						}
					}
				}
				else if (eventInfo.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1283": //삭제
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							break;
						case "7169": //엑셀 내보내기
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
							oForm.Items.Item("BPLId01").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}
	}
}
