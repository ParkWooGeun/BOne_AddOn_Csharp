using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 포장생산 작업공정 확인(수정)
	/// </summary>
	internal class PS_PP037 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP037L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP037.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP037_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP037");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP037_CreateItems();
				PS_PP037_SetComboBox();
				PS_PP037_Initialize();
				PS_PP037_SetDocument(oFormDocEntry);
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
		/// PS_PP037_CreateItems
		/// </summary>
		private void PS_PP037_CreateItems()
		{
			try
			{
				oDS_PS_PP037L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				oForm.DataSources.UserDataSources.Add("CpCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CpCode").Specific.DataBind.SetBound(true, "", "CpCode");

				oForm.Items.Item("Mat01").Enabled = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP037_SetComboBox
		/// </summary>
		private void PS_PP037_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				dataHelpClass.Set_ComboList(oForm.Items.Item("CpCode").Specific, "SELECT b.U_CpCode, b.U_CpName FROM [@PS_PP001H] a Inner Join [@PS_PP001L] b On a.Code = b.Code Where a.Code In ('CP701', 'CP702') ", "CP70101", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP037_Initialize
		/// </summary>
		private void PS_PP037_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP037_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_PP037_SetDocument(string oFormDocEntry)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_PP037_AddMatrixRow(0, true); 
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP037_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP037_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);

				if (RowIserted == false)
				{
					oDS_PS_PP037L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP037L.Offset = oRow;
				oDS_PS_PP037L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
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
		/// PS_PP037_MTX01
		/// </summary>
		private void PS_PP037_MTX01()
		{
			int i;
			string sQry;
			string errMessage = string.Empty;
			string Param01;
			string Param02;
			string Param03;
			double S_BQty = 0;
			double S_PQty = 0;
			double S_YQty = 0;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				Param01 = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				Param02 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				Param03 = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = "EXEC PS_PP037_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "'";

				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Items.Item("Mat01").Enabled = false;
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}
				else
				{
					oForm.Items.Item("Mat01").Enabled = true;
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i != 0)
					{
						oDS_PS_PP037L.InsertRecord(i);
					}
					oDS_PS_PP037L.Offset = i;
					oDS_PS_PP037L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP037L.SetValue("U_ColReg01", i, Convert.ToString(false));
					oDS_PS_PP037L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("LineId").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("DocDate").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());
															
					oDS_PS_PP037L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("BQty").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("PQty").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColQty03", i, oRecordSet.Fields.Item("YQty").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColQty04", i, oRecordSet.Fields.Item("NQty").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColQty05", i, oRecordSet.Fields.Item("ScrapWt").Value.ToString().Trim());
					oDS_PS_PP037L.SetValue("U_ColNum01", i, oRecordSet.Fields.Item("WorkTime").Value.ToString().Trim());

					S_BQty += Convert.ToDouble(oRecordSet.Fields.Item("BQty").Value.ToString().Trim());
					S_PQty += Convert.ToDouble(oRecordSet.Fields.Item("PQty").Value.ToString().Trim());
					S_YQty += Convert.ToDouble(oRecordSet.Fields.Item("YQty").Value.ToString().Trim());

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oForm.Items.Item("S_BQty").Specific.Value = S_BQty;
				oForm.Items.Item("S_PQty").Specific.Value = S_PQty;
				oForm.Items.Item("S_YQty").Specific.Value = S_YQty;

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP037_SetBaseForm
		/// </summary>
		private void PS_PP037_SetBaseForm()
		{
			int i;
			string sQry;
			string Param01;
			int Param02;
			double Param03;
			double Param04;
			double Param05;
			double Param06;
			double Param07;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				for (i = 1; i <= oMat.RowCount; i++)
				{
					if (oMat.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
					{
						Param01 = oMat.Columns.Item("DocEntry").Cells.Item(i).Specific.Value;
						Param02 = Convert.ToInt32(oMat.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString().Trim());
						Param03 = Convert.ToDouble(oMat.Columns.Item("BQty").Cells.Item(i).Specific.Value.ToString().Trim());
						Param04 = Convert.ToDouble(oMat.Columns.Item("PQty").Cells.Item(i).Specific.Value.ToString().Trim());
						Param05 = Convert.ToDouble(oMat.Columns.Item("YQty").Cells.Item(i).Specific.Value.ToString().Trim());
						Param06 = Convert.ToDouble(oMat.Columns.Item("NQty").Cells.Item(i).Specific.Value.ToString().Trim());
						Param07 = Convert.ToDouble(oMat.Columns.Item("ScrapWt").Cells.Item(i).Specific.Value.ToString().Trim());

						sQry = "EXEC PS_PP037_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "'";
						oRecordSet.DoQuery(sQry);
						PSH_Globals.SBO_Application.MessageBox("데이터를 수정하였습니다.");
					}
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
							PS_PP037_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
					else if (pVal.ItemUID == "Btn02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_PP037_SetBaseForm(); //자료수정
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
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;

			string YY;
			string MM;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oForm.Freeze(true);

							if (pVal.ColUID == "CHK")
							{
								oMat.FlushToDataSource();

								YY = codeHelpClass.Left(oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 4);
								MM = codeHelpClass.Right(codeHelpClass.Left(oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6), 2);

								sQry = "select PeriodStat from OFPR Where Name = '" + YY + "-" + MM + "'";
								oRecordSet.DoQuery(sQry);

								if (oRecordSet.Fields.Item(0).Value.ToString().Trim() != "N")
								{
									oDS_PS_PP037L.SetValue("U_ColReg01", pVal.Row - 1, Convert.ToString(false));
								}
								oMat.LoadFromDataSource();
							}
							oForm.Freeze(false);
						}
					}
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
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string Chk = string.Empty;
			int i;

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01" && pVal.Row == 0 && pVal.ColUID == "CHK")
					{
						oForm.Freeze(true);
						oMat.FlushToDataSource();
						if (string.IsNullOrEmpty(oDS_PS_PP037L.GetValue("U_ColReg01", 0).ToString().Trim()) || (oDS_PS_PP037L.GetValue("U_ColReg01", 0).ToString().Trim()) == "N")
						{
							Chk = "Y";
						}
						else if (oDS_PS_PP037L.GetValue("U_ColReg01", 0).ToString().Trim() == "Y")
						{
							Chk = "N";
						}
						for (i = 0; i <= oMat.VisualRowCount - 1; i++)
						{
							oDS_PS_PP037L.SetValue("U_ColReg01", i, Chk);
						}
						oMat.LoadFromDataSource();
						oForm.Freeze(false);
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LINK_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			string YY;
			string MM;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (!string.IsNullOrEmpty(oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.String))
					{
						YY = codeHelpClass.Left(oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 4);
						MM = codeHelpClass.Right(codeHelpClass.Left(oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6), 2);
						sQry = "select PeriodStat from OFPR Where Name = '" + YY + "-" + MM + "'";
						oRecordSet.DoQuery(sQry);

						if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "N")
						{
							PS_PP043 oTempClass = new PS_PP043();
							oTempClass.LoadForm(oMat.Columns.Item("DocEntry").Cells.Item(pVal.Row).Specific.String);
							BubbleEvent = false;
						}
					}
					else
					{
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			short i;
			double S_PQty = 0;
			double S_YQty = 0;
			string YY;
			string MM;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "PQty")
							{
								//생산량
								oMat.Columns.Item("YQty").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();

								for (i = 1; i <= oMat.VisualRowCount; i++)
								{
									S_PQty += Convert.ToDouble(oMat.Columns.Item("PQty").Cells.Item(i).Specific.Value.ToString().Trim());
									S_YQty += Convert.ToDouble(oMat.Columns.Item("YQty").Cells.Item(i).Specific.Value.ToString().Trim());
								}
								oForm.Items.Item("S_PQty").Specific.Value = S_PQty;
								oForm.Items.Item("S_YQty").Specific.Value = S_YQty;
							}
							else if (pVal.ColUID == "CHK")
							{
								YY = codeHelpClass.Left(oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 4);
								MM = codeHelpClass.Right(codeHelpClass.Left(oMat.Columns.Item("DocDate").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(), 6), 2);
								sQry = "select PeriodStat from OFPR Where Name = '" + YY + "-" + MM + "'";
								oRecordSet.DoQuery(sQry);

								if (oRecordSet.Fields.Item(0).Value != "N")
								{
									oMat.Columns.Item("CHK").Cells.Item(pVal.Row).Specific.Value = "N";
								}
							}
							else
							{
							}
							oForm.Update();
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							sQry = "SELECT ItemName FROM [OITM] WHERE ItemCode =  '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("ItemName").Specific.String = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP037L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
