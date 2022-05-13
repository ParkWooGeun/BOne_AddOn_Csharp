using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 구매실적가격조회
	/// </summary>
	internal class PS_MM255 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM255.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM255_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM255");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_MM255_CreateItems();
				PS_MM255_ComboBox_Setting();
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
		/// PS_MM255_CreateItems
		/// </summary>
		private void PS_MM255_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;

				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("OrdType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("Size", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");
				oForm.Items.Item("OrdType").Specific.DataBind.SetBound(true, "", "OrdType");
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");
				oForm.Items.Item("Size").Specific.DataBind.SetBound(true, "", "Size");
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
				oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.AddYears(-1).ToString("yyyyMMdd");	//일년전

				oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
				oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM255_ComboBox_Setting
		/// </summary>
		private void PS_MM255_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From OBPL OBPL order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				// 품목구분
				oForm.Items.Item("OrdType").Specific.ValidValues.Add("0", "전체품목");
				sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("OrdType").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
		/// PS_MM255_MTX01
		/// </summary>
		private void PS_MM255_MTX01()
		{
			string sQry;
			string Param01;
			string Param02;
			string Param03;
			string Param04;
			string Param05;
			string Param06;
			string Param07;
			string Param08;
			string errMessage = string.Empty;

			try
			{
				oForm.Freeze(true);
				Param01 = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Param02 = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
				Param03 = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
				Param04 = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
				Param05 = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				Param06 = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
				Param07 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				Param08 = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();

				sQry = "EXEC PS_MM255_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "','" + Param07 + "','" + Param08 + "'";

				oGrid.DataTable.Clear();
				oForm.DataSources.DataTables.Item("DataTable").ExecuteQuery(sQry);
				oGrid.DataTable = oForm.DataSources.DataTables.Item("DataTable");

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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
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
                //case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
                //	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                //	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                //	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
		/// ITEM_PRESSED 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Button01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_MM255_MTX01();
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
		/// VALIDATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
						if (pVal.ItemUID == "CardCode")
						{
							sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
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
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
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
