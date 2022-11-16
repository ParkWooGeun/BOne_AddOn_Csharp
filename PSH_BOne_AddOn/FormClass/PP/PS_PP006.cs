using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 외주업체별 가공비단가등록
	/// </summary>
	internal class PS_PP006 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_PP006H; //등록헤더

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP006.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP006_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP006");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_PP006_CreateItems();
				PS_PP006_ComboBox_Setting();
				PS_PP006_CF_ChooseFromList();
				PS_PP006_EnableMenus();
				PS_PP006_SetDocument(oFormDocEntry);
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
		/// PS_PP006_CreateItems
		/// </summary>
		private void PS_PP006_CreateItems()
		{
			try
			{
				oDS_PS_PP006H = oForm.DataSources.DBDataSources.Item("@PS_PP006H");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("Rad01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.DataSources.UserDataSources.Add("Rad02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.DataSources.UserDataSources.Add("CtrDate", SAPbouiCOM.BoDataType.dt_DATE, 8);

				oForm.Items.Item("Rad01").Specific.DataBind.SetBound(true, "", "Rad01");
				oForm.Items.Item("Rad02").Specific.DataBind.SetBound(true, "", "Rad02");
				oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");
				oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");
				oForm.Items.Item("CtrDate").Specific.DataBind.SetBound(true, "", "CtrDate");

				oForm.Items.Item("Rad01").Specific.ValOn = "A";
				oForm.Items.Item("Rad01").Specific.ValOff = "0";
				oForm.Items.Item("Rad01").Specific.Selected = true;

				oForm.Items.Item("Rad02").Specific.ValOn = "B";
				oForm.Items.Item("Rad02").Specific.ValOff = "0";
				oForm.Items.Item("Rad02").Specific.GroupWith("Rad01");

				oForm.Settings.MatrixUID = "Mat01";

				// 서식세팅
				oForm.Settings.Enabled = true;
				oForm.Settings.EnableRowFormat = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_PP006_ComboBox_Setting
		/// </summary>
		private void PS_PP006_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Combo_ValidValues_Insert("PS_PP006", "Mat01", "Unit", "EA", "수량");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP006", "Mat01", "Unit", "KG", "중량");
				dataHelpClass.Combo_ValidValues_SetValueColumn(oMat.Columns.Item("Unit"), "PS_PP006", "Mat01", "Unit", false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_PP006_CF_ChooseFromList
		/// </summary>
		private void PS_PP006_CF_ChooseFromList()
		{
			SAPbouiCOM.ChooseFromListCollection oCFLs02 = null;
			SAPbouiCOM.ChooseFromList oCFL02 = null;
			SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams02 = null;
			SAPbouiCOM.EditText oEdit02 = null;
			// 품목 - 매트릭스
			SAPbouiCOM.ChooseFromListCollection oCFLs04 = null;
			SAPbouiCOM.ChooseFromList oCFL04 = null;
			SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams04 = null;
			SAPbouiCOM.Column oColumn04 = null;

			try
			{
				oEdit02 = oForm.Items.Item("ItemCode").Specific;
				oCFLs02 = oForm.ChooseFromLists;
				oCFLCreationParams02 = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

				oCFLCreationParams02.ObjectType = "4";
				oCFLCreationParams02.UniqueID = "CFLITEMCD02";
				oCFLCreationParams02.MultiSelection = false;
				oCFL02 = oCFLs02.Add(oCFLCreationParams02);

				oEdit02.ChooseFromListUID = "CFLITEMCD02";
				oEdit02.ChooseFromListAlias = "ItemCode";

				oColumn04 = oMat.Columns.Item("ItemCode");
				oCFLs04 = oForm.ChooseFromLists;
				oCFLCreationParams04 = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

				oCFLCreationParams04.ObjectType = "4";
				oCFLCreationParams04.UniqueID = "CFLITEMCD04";
				oCFLCreationParams04.MultiSelection = false;
				oCFL04 = oCFLs04.Add(oCFLCreationParams04);

				oColumn04.ChooseFromListUID = "CFLITEMCD04";
				oColumn04.ChooseFromListAlias = "ItemCode";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs02);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL02);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams02);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit02);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs04);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL04);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams04);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn04);
			}
		}

		/// <summary>
		/// PS_PP006_EnableMenus
		/// </summary>
		private void PS_PP006_EnableMenus()
		{
			try
			{
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
				oForm.EnableMenu("1293", true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP006_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_PP006_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_PP006_FormItemEnabled();
					PS_PP006_AddMatrixRow(0, true);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP006_FormItemEnabled
		/// </summary>
		private void PS_PP006_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.EnableMenu("1282", true); //추가
					oForm.Items.Item("Rad01").Enabled = true;
					oForm.Items.Item("Rad02").Enabled = true;
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("Rad01").Enabled = true;
					oForm.Items.Item("Rad02").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1282", true); //추가
					oForm.Items.Item("Rad01").Enabled = true;
					oForm.Items.Item("Rad02").Enabled = true;
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
		/// PS_PP006_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP006_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_PP006H.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP006H.Offset = oRow;
				oDS_PS_PP006H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
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
		/// PS_PP006_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_PP006_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			int j = 0;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					j = oMat.VisualRowCount - 1;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					j = oMat.VisualRowCount - 1;
				}

				for (i = 1; i <= j; i++)
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("eCardCod").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("eCardNam").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "외주 거래처 코드는 필수입니다.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("ItemName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "품목 코드는 필수입니다.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oMat.Columns.Item("Cprice").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("Cprice").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "계약 단가는 필수입니다.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oMat.Columns.Item("CtrDate").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("CtrDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "계약 일자는 필수입니다.";
						throw new Exception();
					}
				}

				oDS_PS_PP006H.RemoveRecord(oDS_PS_PP006H.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_PP006_SetBaseForm01();
				}

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					PS_PP006_SetBaseForm02();
				}

				ReturnValue = true;
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
			return ReturnValue;
		}

		/// <summary>
		/// 메트릭스에 데이터 로드
		/// </summary>
		private void PS_PP006_MTX01()
		{
			int i;
			string Param01;
			string Param02;
			string Param03;
			string Param04 = string.Empty;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				Param01 = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				Param02 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				Param03 = oForm.Items.Item("CtrDate").Specific.Value.ToString().Trim();

				if (oForm.DataSources.UserDataSources.Item("Rad01").Value == "A")
				{
					Param04 = "A";
				}
				else if (oForm.DataSources.UserDataSources.Item("Rad01").Value == "B")
				{
					Param04 = "B";
				}

				sQry = "EXEC PS_PP006_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i != 0)
					{
						oDS_PS_PP006H.InsertRecord(i);
					}
					oDS_PS_PP006H.Offset = i;
					oDS_PS_PP006H.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP006H.SetValue("Code", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_eCardCod", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_eCardNam", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_ItemCode", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_ItemName", i, oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_Unit", i, oRecordSet.Fields.Item(5).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_Cprice", i, oRecordSet.Fields.Item(6).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_CtrDate", i, oRecordSet.Fields.Item(7).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_CpCode", i, oRecordSet.Fields.Item(8).Value.ToString().Trim());
					oDS_PS_PP006H.SetValue("U_CpName", i, oRecordSet.Fields.Item(9).Value.ToString().Trim());
					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				oForm.Update();
				PS_PP006_AddMatrixRow(oMat.VisualRowCount, false);
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
		/// PS_PP006_SetBaseForm01
		/// </summary>
		private void PS_PP006_SetBaseForm01()
		{
			int i;
			string Param01;
			string Param02;
			string Param03;
			string Param04;
			string Param05;
			string Param06;
			string Param07;
			double Param08;
			string Param09;
			string Param10;
			string Param11;
			string sQry;
			string sQry1;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				for (i = 1; i <= oMat.RowCount; i++)
				{
					sQry = "SELECT ISNULL(MAX(Convert(Numeric, Code)), 0) + 1 FROM [@PS_PP006H]";
					oRecordSet.DoQuery(sQry);

					Param01 = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					Param02 = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					Param03 = oMat.Columns.Item("eCardCod").Cells.Item(i).Specific.Value.ToString().Trim();
					Param04 = oMat.Columns.Item("eCardNam").Cells.Item(i).Specific.Value.ToString().Trim();
					Param05 = oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
					Param06 = oMat.Columns.Item("ItemName").Cells.Item(i).Specific.Value.ToString().Trim();
					Param07 = oMat.Columns.Item("Unit").Cells.Item(i).Specific.Value.ToString().Trim();
					Param08 = Convert.ToDouble(oMat.Columns.Item("Cprice").Cells.Item(i).Specific.Value.ToString().Trim());
					Param09 = oMat.Columns.Item("CtrDate").Cells.Item(i).Specific.Value.ToString().Trim();
					Param10 = oMat.Columns.Item("CpCode").Cells.Item(i).Specific.Value.ToString().Trim();
					Param11 = oMat.Columns.Item("CpName").Cells.Item(i).Specific.Value.ToString().Trim();

					sQry1 = "EXEC PS_PP006_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "', '" + Param08 + "', '" + Param09 + "', '" + Param10 + "', '" + Param11 + "'";
					oRecordSet01.DoQuery(sQry1);
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("데이터를 추가하였습니다", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				oDS_PS_PP006H.RemoveRecord(oDS_PS_PP006H.Size - 1);
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();
				oMat.Clear();
				PS_PP006_AddMatrixRow(0, true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
			}
		}

		/// <summary>
		/// PS_PP006_SetBaseForm02
		/// </summary>
		private void PS_PP006_SetBaseForm02()
		{
			int i;
			string Param01;
			string Param02;
			string Param03;
			string Param04;
			string Param05;
			string Param06;
			double Param07;
			string Param08;
			string Param09;
			string Param10;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				for (i = 1; i <= oMat.RowCount; i++)
				{
					Param01 = oMat.Columns.Item("Code").Cells.Item(i).Specific.Value.ToString().Trim();
					Param02 = oMat.Columns.Item("eCardCod").Cells.Item(i).Specific.Value.ToString().Trim();
					Param03 = oMat.Columns.Item("eCardNam").Cells.Item(i).Specific.Value.ToString().Trim();
					Param04 = oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
					Param05 = oMat.Columns.Item("ItemName").Cells.Item(i).Specific.Value.ToString().Trim();
					Param06 = oMat.Columns.Item("Unit").Cells.Item(i).Specific.Value.ToString().Trim();
					Param07 = Convert.ToDouble(oMat.Columns.Item("Cprice").Cells.Item(i).Specific.Value.ToString().Trim());
					Param08 = oMat.Columns.Item("CtrDate").Cells.Item(i).Specific.Value.ToString().Trim();
					Param09 = oMat.Columns.Item("CpCode").Cells.Item(i).Specific.Value.ToString().Trim();
					Param10 = oMat.Columns.Item("CpName").Cells.Item(i).Specific.Value.ToString().Trim();
					sQry = " EXEC PS_PP006_03 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', ";
					sQry += "'" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "', '" + Param08 + "', '" + Param09 + "', '" + Param10 + "'";
					oRecordSet.DoQuery(sQry);
					PSH_Globals.SBO_Application.StatusBar.SetText("데이터를 수정 하였습니다", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				}
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
					Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//	Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
					Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
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
                //           case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
                //Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_KEY_DOWN: //22
                //    Raise_EVENT_FORM_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
					if (pVal.ItemUID == "Btn01")
					{
						PS_PP006_MTX01();
					}
					else if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP006_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP006_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
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
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "eCardCod");
					if (pVal.ColUID == "CpCode")
					{
						if (string.IsNullOrEmpty(oMat.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
						{
							dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CpCode");
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "Unit")
							{
								oDS_PS_PP006H.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP006H.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_PP006_AddMatrixRow(pVal.Row, false);
								}
							}
							else
							{
								oDS_PS_PP006H.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							}
						}
						else
						{
							if (pVal.ItemUID == "CardCode")
							{
								oDS_PS_PP006H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
							else if (pVal.ItemUID == "CardCode")
							{
								oDS_PS_PP006H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
								oDS_PS_PP006H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
							}
							else
							{
								oDS_PS_PP006H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
						}
						oMat.LoadFromDataSource();
						oMat.AutoResizeColumns();
						oForm.Update();
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								if (pVal.ColUID == "eCardCod")
								{
									oDS_PS_PP006H.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									if (oMat.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_PP006H.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
									{
										PS_PP006_AddMatrixRow(pVal.Row, false);
									}

									oMat.LoadFromDataSource();

									sQry = "SELECT CardName FROM [OCRD] WHERE CardCode = '" + oMat.Columns.Item("eCardCod").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
									oRecordSet.DoQuery(sQry);
									oMat.Columns.Item("eCardNam").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
								}
								else if (pVal.ColUID == "CpCode")
								{
									oDS_PS_PP006H.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									oDS_PS_PP006H.SetValue("U_CpName", pVal.Row - 1, dataHelpClass.Get_ReData("U_CpName", "U_CpCode", "[@PS_PP001L]", "'" + oDS_PS_PP006H.GetValue("U_CpCode", pVal.Row - 1).ToString().Trim() + "'", ""));
									oMat.LoadFromDataSource();
								}
							}

							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (pVal.ItemUID == "CardCode")
						{
							sQry = "SELECT CardName FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP006_FormItemEnabled();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "ItemCode")
					{
						oForm.DataSources.UserDataSources.Item("ItemCode").Value = oDataTable01.Columns.Item(0).Cells.Item(0).Value;
						oForm.DataSources.UserDataSources.Item("ItemName").Value = oDataTable01.Columns.Item(1).Cells.Item(0).Value;
					}
					else if (pVal.ItemUID == "CardCode")
					{
						oForm.DataSources.UserDataSources.Item("CardCode").Value = oDataTable01.Columns.Item(0).Cells.Item(0).Value;
						oForm.DataSources.UserDataSources.Item("CardName").Value = oDataTable01.Columns.Item(1).Cells.Item(0).Value;
					}

					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.ColUID == "eCardCod")
						{
							oMat.FlushToDataSource();

							oDS_PS_PP006H.SetValue("U_eCardCod", pVal.Row - 1, oDataTable01.Columns.Item("CardCode").Cells.Item(0).Value.ToString().Trim());
							oDS_PS_PP006H.SetValue("U_eCardNam", pVal.Row - 1, oDataTable01.Columns.Item("CardName").Cells.Item(0).Value.ToString().Trim());

							if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP006H.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
							{
								PS_PP006_AddMatrixRow(pVal.Row, false);
							}

							oMat.LoadFromDataSource();
						}
						else if (pVal.ColUID == "ItemCode")
						{
							oMat.FlushToDataSource();

							oDS_PS_PP006H.SetValue("U_ItemCode", pVal.Row - 1, oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value.ToString().Trim());
							oDS_PS_PP006H.SetValue("U_ItemName", pVal.Row - 1, oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value.ToString().Trim());
							//                Next i

							if (oMat.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_PP006H.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
							{
								PS_PP006_AddMatrixRow(pVal.Row, false);
							}

							oMat.LoadFromDataSource();
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP006H);
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
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_RightClickEvent(string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
				}

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
							oMat.Clear();
							oMat.FlushToDataSource();
							oMat.LoadFromDataSource();
							PS_PP006_FormItemEnabled();
							PS_PP006_AddMatrixRow(0, true);
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
