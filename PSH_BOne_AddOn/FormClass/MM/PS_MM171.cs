using System;
using SAPbouiCOM;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 지체상금 대상 입고 리스트  PS_MM170 - SUB
	/// </summary>
	internal class PS_MM171 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_MM171L; //등록라인
		
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private SAPbouiCOM.Form oBaseForm01;
		private string oBaseItemUID01;
		private string oBaseColUID01;
		private int oBaseColRow01;

		/// <summary>
		///  Form 호출
		/// </summary>
		/// <param name="oForm02"></param>
		/// <param name="oItemUID02"></param>
		/// <param name="oColUID02"></param>
		/// <param name="oColRow02"></param>
		public void LoadForm(ref SAPbouiCOM.Form oForm02, string oItemUID02, string oColUID02, int oColRow02)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM171.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM171_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM171");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				oBaseForm01 = oForm02;
				oBaseItemUID01 = oItemUID02;
				oBaseColUID01 = oColUID02;
				oBaseColRow01 = oColRow02;

				PS_MM171_CreateItems();
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
		/// PS_MM171_CreateItems
		/// </summary>
		private void PS_MM171_CreateItems()
		{
			try
			{
				oDS_PS_MM171L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
				oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");
				oForm.DataSources.UserDataSources.Item("CardCode").Value = oBaseForm01.Items.Item("CardCode").Specific.Value;

				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");
				oForm.DataSources.UserDataSources.Item("BPLId").Value = oBaseForm01.Items.Item("BPLId").Specific.Value;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM171_MTX01
		/// </summary>
		private void PS_MM171_MTX01()
		{
			int i;
			int j;
			string Param01;
			string Param02;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				Param01 = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				Param02 = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

				sQry = "EXEC PS_MM171_01 '" + Param01 + "','" + Param02 + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}
				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				ProgressBar01.Text = "조회시작!";

				j = 0;
				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (Convert.ToDouble(oRecordSet.Fields.Item("RepayP").Value.ToString().Trim()) > 0)
					{
						if (j != 0)
						{
							oDS_PS_MM171L.InsertRecord(j);
						}

						oDS_PS_MM171L.Offset = j;
						oDS_PS_MM171L.SetValue("U_LineNum", j, Convert.ToString(j + 1));
						oDS_PS_MM171L.SetValue("U_ColReg01", j, Convert.ToString(false));
						oDS_PS_MM171L.SetValue("U_ColReg02", j, oRecordSet.Fields.Item("GRDocNum").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg03", j, oRecordSet.Fields.Item("GRLinNum").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg04", j, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg05", j, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg06", j, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg07", j, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColSum01", j, oRecordSet.Fields.Item("LinTotal").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColDt01", j, Convert.ToDateTime(oRecordSet.Fields.Item("ImDate").Value.ToString().Trim()).ToString("yyyyMMdd"));
						oDS_PS_MM171L.SetValue("U_ColDt02", j, Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd"));
						oDS_PS_MM171L.SetValue("U_ColReg08", j, oRecordSet.Fields.Item("LateDay").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColSum02", j, oRecordSet.Fields.Item("RepayP").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg09", j, oRecordSet.Fields.Item("DocType").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg10", j, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg11", j, oRecordSet.Fields.Item("PODocNum").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg12", j, oRecordSet.Fields.Item("Unit").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg13", j, oRecordSet.Fields.Item("Size").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg14", j, oRecordSet.Fields.Item("Qty").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg15", j, oRecordSet.Fields.Item("Unweight").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg16", j, oRecordSet.Fields.Item("ItmBsort").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg17", j, oRecordSet.Fields.Item("ItmMsort").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg18", j, oRecordSet.Fields.Item("ItemType").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg19", j, oRecordSet.Fields.Item("Quality").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg20", j, oRecordSet.Fields.Item("Mark").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg21", j, oRecordSet.Fields.Item("CallSize").Value.ToString().Trim());
						oDS_PS_MM171L.SetValue("U_ColReg22", j, oRecordSet.Fields.Item("ObasUnit").Value.ToString().Trim());
						j += 1;
					}
					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				oForm.Update();
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_MM171_SetBaseForm
		/// </summary>
		private void PS_MM171_SetBaseForm()
		{
			int i;
			int j;
			SAPbouiCOM.Matrix oMat02 = null;
			string errMessage = string.Empty;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oMat02 = oBaseForm01.Items.Item("Mat01").Specific;
					for (i = 1; i <= oMat.VisualRowCount; i++)
					{
						for (j = 1; j <= oMat02.VisualRowCount; j++)
						{
							if (oMat.Columns.Item("GRDocNum").Cells.Item(i).Specific.Value.ToString().Trim() == oMat02.Columns.Item("GRDocNum").Cells.Item(j).Specific.Value.ToString().Trim())
							{
								errMessage = "입고 번호가 이미 있습니다. 확인 후 재선택 하세요!.";
								throw new Exception();
							}
						}
						
						if (oMat.Columns.Item("V_0").Cells.Item(i).Specific.Checked == true)
						{
							oMat02.Columns.Item("GRDocNum").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("GRDocNum").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("GRLinNum").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("GRLinNum").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("PODocNum").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("PODocNum").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("ItemName").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("ItemName").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("LinTotal").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("LinTotal").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("ImDate").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("ImDate").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("DueDate").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("DueDate").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("LateDay").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("LateDay").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("Qty").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("Qty").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("Unit").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("Unit").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("Size").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("Size").Cells.Item(i).Specific.Value.ToString().Trim();

							oMat02.Columns.Item("Weight").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("Unweight").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("RepayP").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("RepayP").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("ItmBsort").Cells.Item(oBaseColRow01).Specific.Select(oMat.Columns.Item("ItmBsort").Cells.Item(i).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
							oMat02.Columns.Item("ItmMsort").Cells.Item(oBaseColRow01).Specific.Select(oMat.Columns.Item("ItmMsort").Cells.Item(i).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
							oMat02.Columns.Item("ItemType").Cells.Item(oBaseColRow01).Specific.Select(oMat.Columns.Item("ItemType").Cells.Item(i).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
							oMat02.Columns.Item("Quality").Cells.Item(oBaseColRow01).Specific.Select(oMat.Columns.Item("Quality").Cells.Item(i).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
							oMat02.Columns.Item("Mark").Cells.Item(oBaseColRow01).Specific.Select(oMat.Columns.Item("Mark").Cells.Item(i).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
							oMat02.Columns.Item("CallSize").Cells.Item(oBaseColRow01).Specific.Value = oMat.Columns.Item("CallSize").Cells.Item(i).Specific.Value.ToString().Trim();
							oMat02.Columns.Item("ObasUnit").Cells.Item(oBaseColRow01).Specific.Select(oMat.Columns.Item("ObasUnit").Cells.Item(i).Specific.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
							oBaseColRow01 += 1;
						}
					}
					oForm.Close();
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
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
				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
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
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
				//	Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE: //19
				//    Raise_EVENT_FORM_DEACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE: //20
				//    Raise_EVENT_FORM_CLOSE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//           case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//break;
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
			string errMessage = string.Empty;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_MM171_MTX01();
						}
					}
					if (pVal.ItemUID == "Btn02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_MM171_SetBaseForm(); //부모폼에 입력하는 작업
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{

					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
					BubbleEvent = false;
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
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
				if (pVal.ItemUID == "Mat01")
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
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM171L);
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
