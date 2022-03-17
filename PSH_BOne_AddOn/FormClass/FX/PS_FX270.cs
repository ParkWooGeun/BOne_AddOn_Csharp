using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 고정자산관리대장(개별)
	/// </summary>
	internal class PS_FX270 : PSH_BaseClass
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FX270.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_FX270_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_FX270");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_FX270_CreateItems();
				PS_FX270_ComboBox_Setting();
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
		/// PS_FX270_CreateItems
		/// </summary>
		private void PS_FX270_CreateItems()
		{
			try
			{
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("ClasCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("FixCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");
				oForm.Items.Item("ClasCode").Specific.DataBind.SetBound(true, "", "ClasCode");
				oForm.Items.Item("FixCode").Specific.DataBind.SetBound(true, "", "FixCode");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX270_ComboBox_Setting
		/// </summary>
		private void PS_FX270_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("", "");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("ClasCode").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ClasCode").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'FX001'", "", false, false);
				oForm.Items.Item("ClasCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX270_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_FX270_DataValidCheck()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "사업장은 필수입니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("FixCode").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("FixCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "자산코드는 필수입니다.";
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
		/// PS_FX270_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_FX270_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string BPLId;
			string FixCode;
			//string BPLName;
			//string sQry;
			//SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId   = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				FixCode = oForm.Items.Item("FixCode").Specific.Value.ToString().Trim();

				//sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + BPLId + "'";
				//oRecordSet.DoQuery(sQry);
				//BPLName = oRecordSet.Fields.Item(0).Value.ToString().Trim();

				WinTitle = "고정자산관리대장(개별) [PS_FX270_01]";
				ReportName = "PS_FX270_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
				List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

				//Formula

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@FixCode", FixCode));

				//SubReport Parameter
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@BPLId", BPLId, "PS_FX270_01_SUB1"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@FixCode", FixCode, "PS_FX270_01_SUB1"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@BPLId", BPLId, "PS_FX270_01_SUB2"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@FixCode", FixCode, "PS_FX270_01_SUB2"));

				formHelpClass.OpenCrystalReport(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText("PH_PY540_Print_Report01_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
				case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
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
					if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_FX270_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_FX270_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pval.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pval, ref BubbleEvent, "FixCode", "");
				}
				else if (pval.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			string FixCode;
			string SubCode;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				oForm.Freeze(true);
				if (pval.BeforeAction == true)
				{
					if (pval.ItemChanged == true)
					{
						if (pval.ItemUID == "FixCode")
						{
							FixCode = codeHelpClass.Left(oForm.Items.Item("FixCode").Specific.Value.ToString().Trim(), 6);
							SubCode = codeHelpClass.Right(oForm.Items.Item("FixCode").Specific.Value.ToString().Trim(), 3);

							sQry = "Select U_FixName From [@PS_FX005H] Where U_FixCode = '" + FixCode + "'";
							sQry = sQry + " and U_SubCode = '" + SubCode + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("FixName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}

						oForm.Items.Item(pval.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					}
				}
				else if (pval.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
