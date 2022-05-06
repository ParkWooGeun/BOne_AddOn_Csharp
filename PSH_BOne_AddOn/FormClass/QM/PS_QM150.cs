using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 제안시상기준표등록
	/// </summary>
	internal class PS_QM150 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM150.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM150_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM150");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_QM150_CreateItems();
				PS_QM150_ComboBox_Setting();

				oForm.EnableMenu("1282", true);
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
		/// PS_QM150_CreateItems
		/// </summary>
		private void PS_QM150_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;

				//등급
				oForm.DataSources.UserDataSources.Add("DivNm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("DivNm").Specific.DataBind.SetBound(true, "", "DivNm");
				//등급
				oForm.DataSources.UserDataSources.Add("Grade", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Grade").Specific.DataBind.SetBound(true, "", "Grade");
				//순번
				oForm.DataSources.UserDataSources.Add("Num", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Num").Specific.DataBind.SetBound(true, "", "Num");
				//득점
				oForm.DataSources.UserDataSources.Add("MarkMin", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("MarkMin").Specific.DataBind.SetBound(true, "", "MarkMin");
				oForm.DataSources.UserDataSources.Add("MarkMax", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("MarkMax").Specific.DataBind.SetBound(true, "", "MarkMax");
				//시상액
				oForm.DataSources.UserDataSources.Add("PrizeAmt", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("PrizeAmt").Specific.DataBind.SetBound(true, "", "PrizeAmt");
				//시상액
				oForm.DataSources.UserDataSources.Add("Par", SAPbouiCOM.BoDataType.dt_PRICE);
				oForm.Items.Item("Par").Specific.DataBind.SetBound(true, "", "Par");
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM150_ComboBox_Setting
		/// </summary>
		private void PS_QM150_ComboBox_Setting()
		{
			string BPLId;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				BPLId = dataHelpClass.User_BPLID();
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLID, BPLName FROM OBPL order by BPLID", BPLId, false, false);

				//콤보에 기본값설정
				dataHelpClass.Combo_ValidValues_Insert("PS_QM150", "Div", "", "0", "정식");
				dataHelpClass.Combo_ValidValues_Insert("PS_QM150", "Div", "", "1", "약식");
				dataHelpClass.Combo_ValidValues_Insert("PS_QM150", "Div", "", "2", "등외");
				dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("Div").Specific, "PS_QM150", "Div", false);
				oForm.Items.Item("Div").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM150_MTX01
		/// </summary>
		private void PS_QM150_MTX01()
		{
			string BPLId;
			string sQry;
			string errMessage = string.Empty;

			try
			{
				oForm.Freeze(true);
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

				sQry = " EXEC [PS_QM150_01] ";
				sQry += "'" + BPLId + "'";

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
		/// 그리드 자료를 head에 로드
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_QM150_MTX02(string oUID, int oRow, string oCol)
		{
			int sRow;
			string Param01;
			string Param02;
			string Param03;
			string Param04;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				sRow = oRow;
				Param01 = oGrid.DataTable.Columns.Item("코드").Cells.Item(oRow).Value.ToString().Trim();
				Param02 = oGrid.DataTable.Columns.Item("구분명").Cells.Item(oRow).Value.ToString().Trim();
				Param03 = oGrid.DataTable.Columns.Item("등급").Cells.Item(oRow).Value.ToString().Trim();
				Param04 = oGrid.DataTable.Columns.Item("순번").Cells.Item(oRow).Value.ToString().Trim();

				sQry = "EXEC PS_QM150_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Items.Item("Div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
					oForm.Items.Item("DivNm").Specific.Value = "";
					oForm.Items.Item("Grade").Specific.Value = "";
					oForm.Items.Item("Num").Specific.Value = "";

					oForm.Items.Item("MarkMin").Specific.Value = "0";
					oForm.Items.Item("MarkMax").Specific.Value = "0";
					oForm.Items.Item("PrizeAmt").Specific.Value = "0";
					oForm.Items.Item("Par").Specific.Value = "0";

					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oForm.Items.Item("Div").Specific.Select(oRecordSet.Fields.Item("Div").Value.ToString().Trim(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.DataSources.UserDataSources.Item("DivNm").Value = oRecordSet.Fields.Item("DivNm").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("Grade").Value = oRecordSet.Fields.Item("Grade").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("Num").Value = oRecordSet.Fields.Item("Num").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("MarkMin").Value = oRecordSet.Fields.Item("MarkMin").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("MarkMax").Value = oRecordSet.Fields.Item("MarkMax").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("PrizeAmt").Value = oRecordSet.Fields.Item("PrizeAmt").Value.ToString().Trim();
				oForm.DataSources.UserDataSources.Item("Par").Value = oRecordSet.Fields.Item("Par").Value.ToString().Trim();

				oForm.ActiveItem = "MarkMin"; 
				oForm.Items.Item("Div").Enabled = false;
				oForm.Items.Item("Grade").Enabled = false;
				oForm.Items.Item("Num").Enabled = false;

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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_QM150_SAVE
		/// </summary>
		private void PS_QM150_SAVE()
		{
			string Grade;
			string Div;
			string DivNm;
			string num;
			decimal MarkMax;
			decimal MarkMin;
			decimal PrizeAmt;
			decimal Par;
			string BPLId;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				Div = oForm.Items.Item("Div").Specific.Value.ToString().Trim();
				DivNm = oForm.Items.Item("DivNm").Specific.Value.ToString().Trim();
				Grade = oForm.Items.Item("Grade").Specific.Value.ToString().Trim();
				num = oForm.Items.Item("Num").Specific.Value.ToString().Trim();
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

				MarkMin = Convert.ToDecimal(oForm.Items.Item("MarkMin").Specific.Value.ToString().Trim());
				MarkMax = Convert.ToDecimal(oForm.Items.Item("MarkMax").Specific.Value.ToString().Trim());
				PrizeAmt = Convert.ToDecimal(oForm.Items.Item("PrizeAmt").Specific.Value.ToString().Trim());
				Par = Convert.ToDecimal(oForm.Items.Item("Par").Specific.Value.ToString().Trim());

				if (string.IsNullOrEmpty(Div))
				{
					errMessage = "구분코드에러. 확인바랍니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(Grade))
				{
					errMessage = "등급에러.확인바랍니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(num))
				{
					errMessage = "순번에러. 확인바랍니다.";
					throw new Exception();
				}

				sQry = "    Select     Count(*) ";
				sQry += " From      [ZPS_QM150] ";
				sQry += " Where    Div = '" + Div + "'";
				sQry += "             And DivNm = '" + DivNm + "'";
				sQry += "             And Grade = '" + Grade + "'";
				sQry += "             And Num = '" + num + "'";
				sQry += "             And BPLID = '" + BPLId + "'";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 0)
				{
					sQry = "Update [ZPS_QM150] set ";
					sQry += "MarkMin = '" + MarkMin + "',";
					sQry += "MarkMax = '" + MarkMax + "',";
					sQry += "PrizeAmt = '" + PrizeAmt + "',";
					sQry += "Par = '" + Par + "'";
					sQry += " Where Div = '" + Div + "'";
					sQry += "          And DivNm = '" + DivNm + "'";
					sQry += "          And Grade = '" + Grade + "'";
					sQry += "          And Num = '" + num + "'";
					sQry += "          And BPLID = '" + BPLId + "'";
					oRecordSet.DoQuery(sQry);
				}
				else
				{
					sQry = "INSERT INTO [ZPS_QM150]";
					sQry += " (";
					sQry += "Div,";
					sQry += "DivNm,";
					sQry += "Grade,";
					sQry += "Num,";
					sQry += "MArkMin,";
					sQry += "MarkMax,";
					sQry += "PrizeAmt,";
					sQry += "Par,";
					sQry += "BPLID";
					sQry += " ) ";
					sQry += "VALUES(";
					sQry += "'" + Div + "',";
					sQry += "'" + DivNm + "',";
					sQry += "'" + Grade + "',";
					sQry += "'" + num + "',";
					sQry += MarkMin + ",";
					sQry += MarkMax + ",";
					sQry += PrizeAmt + ",";
					sQry += Par + ",";
					sQry += "'" + BPLId + "'";
					sQry += " ) ";
					oRecordSet.DoQuery(sQry);
				}

				PS_QM150_MTX01();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 선택된 자료 삭제
		/// </summary>
		private void PS_QM150_Delete()
		{
			string Grade;
			string Div;
			string DivNm;
			string num;
			int Cnt;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				Div = oForm.Items.Item("Div").Specific.Value.ToString().Trim();
				DivNm = oForm.Items.Item("DivNm").Specific.Value.ToString().Trim();
				Grade = oForm.Items.Item("Grade").Specific.Value.ToString().Trim();
				num = oForm.Items.Item("Num").Specific.Value.ToString().Trim();

				sQry = " Select Count(*) From [ZPS_QM150] Where Div = '" + Div + "' And DivNm = '" + DivNm + "' And Grade = '" + Grade + "' And Num = '" + num + "'";
				oRecordSet.DoQuery(sQry);
				Cnt = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim());

				if (Cnt > 0)
				{
					if (PSH_Globals.SBO_Application.MessageBox(" 선택한라인을 삭제하시겠습니까? ?", 2, "예", "아니오") == 1)
					{
						sQry = "Delete From [ZPS_QM150] Where Div = '" + Div + "' And DivNm = '" + DivNm + "' And Grade = '" + Grade + "' And Num = '" + num + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("Div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						oForm.Items.Item("DivNm").Specific.Value = "정식";
						oForm.Items.Item("Grade").Specific.Value = "";
						oForm.Items.Item("Num").Specific.Value = "";

						oForm.Items.Item("MarkMin").Specific.Value = "0";
						oForm.Items.Item("MarkMax").Specific.Value = "0";
						oForm.Items.Item("PrizeAmt").Specific.Value = "0";
						oForm.Items.Item("Par").Specific.Value = "0";

						oForm.Items.Item("Div").Enabled = true;
						oForm.Items.Item("Grade").Enabled = true;
						oForm.Items.Item("Num").Enabled = true;

						oForm.ActiveItem = "Div";
						oForm.Update();

						PS_QM150_MTX01();
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
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8	
				//	Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
				//case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//	Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
					if (pVal.ItemUID == "Btn_ret")
					{
						PS_QM150_MTX01();
					}
					else if (pVal.ItemUID == "Btn_save")
					{
						PS_QM150_SAVE();
					}
					else if (pVal.ItemUID == "Btn_delete")
					{
						PS_QM150_Delete();
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string Div;
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Div")
					{
						Div = oForm.Items.Item("Div").Specific.Value.ToString().Trim();

						if (Div == "0")
						{
							oForm.Items.Item("DivNm").Specific.Value = "정식";
						}
						else if (Div == "1")
						{
							oForm.Items.Item("DivNm").Specific.Value = "약식";
						}
						else if (Div == "2")
						{
							oForm.Items.Item("DivNm").Specific.Value = "등외";
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}
		
		/// <summary>
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Grid01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.Row >= 0)
							{
								PS_QM150_MTX02(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
						}
					}
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
							oForm.Items.Item("Div").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
							oForm.Items.Item("DivNm").Specific.Value = "정식";
							oForm.Items.Item("Grade").Specific.Value = "";
							oForm.Items.Item("Num").Specific.Value = "";
							oForm.Items.Item("MarkMin").Specific.Value = "0";
							oForm.Items.Item("MarkMax").Specific.Value = "0";
							oForm.Items.Item("PrizeAmt").Specific.Value = "0";
							oForm.Items.Item("Par").Specific.Value = "0";

							oForm.Items.Item("Div").Enabled = true;
							oForm.Items.Item("Grade").Enabled = true;
							oForm.Items.Item("Num").Enabled = true;
							oForm.ActiveItem = "Div";
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
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
						case "1293": //행삭제
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1287": //복제 
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
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
