using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 외주반출등록 PS_MM130_서브폼
	/// </summary>
	internal class PS_MM131 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_MM131L; //등록라인

		//부모폼
		private SAPbouiCOM.Form oBaseForm01;
		private string oBaseItemUID01;
		private string oBaseColUID01;
		private int oBaseColRow01;
		private string oBaseTradeType01;
		private string oRadioGrp;

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oMatRow01;

		/// <summary>
		///  LoadForm
		/// </summary>
		/// <param name="oForm02"></param>
		/// <param name="oItemUID02"></param>
		/// <param name="oColUID02"></param>
		/// <param name="oColRow02"></param>
		/// <param name="RadioGrp"></param>
		public void LoadForm(ref SAPbouiCOM.Form oForm02, string oItemUID02, string oColUID02, int oColRow02, string RadioGrp)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM131.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM131_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM131");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				oBaseForm01 = oForm02;
				oBaseItemUID01 = oItemUID02;
				oBaseColUID01 = oColUID02;
				oBaseColRow01 = oColRow02;
				oRadioGrp = RadioGrp;

				PS_MM131_CreateItems();
				PS_MM131_ComboBox_Setting();
				PS_MM131_CF_ChooseFromList();
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
		/// PS_MM131_CreateItems
		/// </summary>
		private void PS_MM131_CreateItems()
		{
			try
			{
				oDS_PS_MM131L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.Items.Item("Mat01").Enabled = true;

				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

				oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");

				oForm.DataSources.UserDataSources.Add("Mark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("Mark").Specific.DataBind.SetBound(true, "", "Mark");

				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				oForm.DataSources.UserDataSources.Add("Size", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("Size").Specific.DataBind.SetBound(true, "", "Size");

				oForm.DataSources.UserDataSources.Add("CpCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CpCode").Specific.DataBind.SetBound(true, "", "CpCode");

				oForm.DataSources.UserDataSources.Add("CpName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CpName").Specific.DataBind.SetBound(true, "", "CpName");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM131_ComboBox_Setting
		/// </summary>
		private void PS_MM131_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("ItmBsort").Specific.ValidValues.Add("", "전체");
				sQry = "SELECT Code, Name FROM [@PSH_ITMBSORT] order by Code";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("ItmBsort").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT]", "", false, true);
				dataHelpClass.Set_ComboList(oForm.Items.Item("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] order by Code", "", false, true);
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, "SELECT Code, Name FROM [@PSH_SHAPE] order by Code", "", false, true);
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
		/// PS_MM131_CF_ChooseFromList
		/// </summary>
		private void PS_MM131_CF_ChooseFromList()
		{
			SAPbouiCOM.ChooseFromListCollection oCFLs = null;
			SAPbouiCOM.ChooseFromList oCFL = null;
			SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
			SAPbouiCOM.EditText oEdit = null;

			try
			{
				oEdit = oForm.Items.Item("ItemCode").Specific;
				oCFLs = oForm.ChooseFromLists;
				oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

				oCFLCreationParams.ObjectType = "4";
				oCFLCreationParams.UniqueID = "CFLITEMCD";
				oCFLCreationParams.MultiSelection = false;
				oCFL = oCFLs.Add(oCFLCreationParams);

				oEdit.ChooseFromListUID = "CFLITEMCD";
				oEdit.ChooseFromListAlias = "ItemCode";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				if (oCFLs != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
				}
				if (oCFL != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
				}
				if (oCFLCreationParams != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
				}
				if (oEdit != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
				}
			}
		}

		/// <summary>
		/// PS_MM131_Select_Data
		/// </summary>
		private void PS_MM131_Select_Data()
		{
			int i;
			int j;
			string Param01;
			string Param02;
			string Param03;
			string Param04;
			string Param05;
			string Param06;
			string Param07;
			string Param08;
			string errMessage = string.Empty;
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				Param01 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				Param02 = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();
				Param03 = oForm.Items.Item("ItmMsort").Specific.Value.ToString().Trim();
				Param04 = oForm.Items.Item("Mark").Specific.Value.ToString().Trim();
				Param05 = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				Param06 = oForm.Items.Item("FrgnName").Specific.Value.ToString().Trim();
				Param07 = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
				Param08 = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(Param02))
				{
					errMessage = "대분류는 필수입니다.";
					throw new Exception();
				}

                if (oRadioGrp == "B" && (Param02 == "102") && string.IsNullOrEmpty(Param08))
                {
                    errMessage = "부품 재공반출은 공정코드를 선택해야 합니다.";
                    throw new Exception();
                }

                if (oRadioGrp == "A")
				{
					sQry = "EXEC PS_MM131_01 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "'";
				}
				else if (oRadioGrp == "B")
				{
					sQry = "EXEC PS_MM131_02 '" + Param01 + "', '" + Param02 + "', '" + Param03 + "', '" + Param04 + "', '" + Param05 + "', '" + Param06 + "', '" + Param07 + "', '" + Param08 + "'";
				}

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

				ProgressBar01.Text = "조회시작!";

				j = 0;
				for ( i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (oRadioGrp == "A")
					{
						if (j != 0)
						{
							oDS_PS_MM131L.InsertRecord(j);
						}

						oDS_PS_MM131L.Offset = j;
						oDS_PS_MM131L.SetValue("U_LineNum", j, Convert.ToString(j + 1));
						oDS_PS_MM131L.SetValue("U_ColReg01", j, Convert.ToString(false));
						oDS_PS_MM131L.SetValue("U_ColReg02", j, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg15", j, oRecordSet.Fields.Item("JisNo").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg03", j, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg04", j, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg13", j, oRecordSet.Fields.Item("ItemCod2").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg14", j, oRecordSet.Fields.Item("ItemNam2").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg05", j, oRecordSet.Fields.Item("Size").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColQty04", j, oRecordSet.Fields.Item("UnWeight").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg10", j, oRecordSet.Fields.Item("Mark").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColNum01", j, oRecordSet.Fields.Item("SelQty").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColQty01", j, oRecordSet.Fields.Item("SelWt").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColNum02", j, oRecordSet.Fields.Item("OutQty").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColQty02", j, oRecordSet.Fields.Item("OutWt").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg06", j, oRecordSet.Fields.Item("OutWhCd").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg07", j, oRecordSet.Fields.Item("OutWhNm").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg08", j, oRecordSet.Fields.Item("InWhCd").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg09", j, oRecordSet.Fields.Item("InWhNm").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColNum03", j, oRecordSet.Fields.Item("PosQty").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColQty03", j, oRecordSet.Fields.Item("PosWt").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg11", j, oRecordSet.Fields.Item("PP030HNo").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg12", j, oRecordSet.Fields.Item("PP030MNo").Value.ToString().Trim());

						sQry = "Select U_CpCode, U_CpName From [@PS_PP030M] Where DocEntry = '" + oRecordSet.Fields.Item("PP030HNo").Value.ToString().Trim() + "' And U_Sequence = '" + oRecordSet.Fields.Item("PP030MNo").Value.ToString().Trim() + "'";
						oRecordSet02.DoQuery(sQry);

						if (oRecordSet02.RecordCount > 0)
						{
							oDS_PS_MM131L.SetValue("U_ColReg16", j, oRecordSet02.Fields.Item("U_CpCode").Value.ToString().Trim());
							oDS_PS_MM131L.SetValue("U_ColReg17", j, oRecordSet02.Fields.Item("U_CpName").Value.ToString().Trim());
						}

						j += 1;
					}
					else if (oRadioGrp == "B" && Convert.ToDouble(oRecordSet.Fields.Item("PosQty").Value.ToString().Trim()) > 0)
					{
						if (j != 0)
						{
							oDS_PS_MM131L.InsertRecord(j);
						}

						oDS_PS_MM131L.Offset = j;
						oDS_PS_MM131L.SetValue("U_LineNum", j, Convert.ToString(j + 1));
						oDS_PS_MM131L.SetValue("U_ColReg01", j, Convert.ToString(false));
						oDS_PS_MM131L.SetValue("U_ColReg02", j, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg15", j, oRecordSet.Fields.Item("JisNo").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg03", j, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg04", j, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg13", j, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg14", j, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg05", j, oRecordSet.Fields.Item("Size").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColQty04", j, oRecordSet.Fields.Item("UnWeight").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg10", j, oRecordSet.Fields.Item("Mark").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColNum01", j, oRecordSet.Fields.Item("SelQty").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColQty01", j, oRecordSet.Fields.Item("SelWt").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColNum02", j, oRecordSet.Fields.Item("OutQty").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColQty02", j, oRecordSet.Fields.Item("OutWt").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg06", j, oRecordSet.Fields.Item("OutWhCd").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg07", j, oRecordSet.Fields.Item("OutWhNm").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg08", j, oRecordSet.Fields.Item("InWhCd").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg09", j, oRecordSet.Fields.Item("InWhNm").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColNum03", j, oRecordSet.Fields.Item("PosQty").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColQty03", j, oRecordSet.Fields.Item("PosWt").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg11", j, oRecordSet.Fields.Item("PP030HNo").Value.ToString().Trim());
						oDS_PS_MM131L.SetValue("U_ColReg12", j, oRecordSet.Fields.Item("PP030MNo").Value.ToString().Trim());

						sQry = "Select U_CpCode, U_CpName From [@PS_PP030M] Where DocEntry = '" + oRecordSet.Fields.Item("PP030HNo").Value.ToString().Trim() + "' And U_Sequence = '" + oRecordSet.Fields.Item("PP030MNo").Value.ToString().Trim() + "'";
						oRecordSet02.DoQuery(sQry);

						if (oRecordSet02.RecordCount > 0)
						{
							oDS_PS_MM131L.SetValue("U_ColReg16", j, oRecordSet02.Fields.Item("U_CpCode").Value.ToString().Trim());
							oDS_PS_MM131L.SetValue("U_ColReg17", j, oRecordSet02.Fields.Item("U_CpName").Value.ToString().Trim());
						}

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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_MM131_SetBaseForm
		/// </summary>
		private void PS_MM131_SetBaseForm()
		{
			int i;
			int j = 0;
			SAPbouiCOM.Matrix oBaseMat01 = null;
			SAPbouiCOM.DBDataSource oDS_PS_MM130L = null;

			try
			{
				oDS_PS_MM130L = oBaseForm01.DataSources.DBDataSources.Item("@PS_MM130L");
				oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;

				if (oBaseForm01 == null)
				{
				}
				else if (oBaseForm01.TypeEx == "PS_MM130")
				{
					//품목선택품목
					for (i = 1; i <= oMat.VisualRowCount; i++)
					{
						if (oMat.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							oBaseMat01.Columns.Item("OrdNum").Cells.Item(oBaseMat01.VisualRowCount).Specific.Value = oMat.Columns.Item("OrdNum").Cells.Item(i).Specific.Value.ToString().Trim();
							oDS_PS_MM130L.SetValue("U_ItemCode", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_ItemName", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("ItemName").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_Size", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("Size").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_UnWeight", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("UnWeight").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_Qty", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_Weight", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("SelWt").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_Mark", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("Mark").Cells.Item(i).Specific.Value.ToString().Trim());
							if (oRadioGrp == "A")
							{
								oDS_PS_MM130L.SetValue("U_OutGbn", oBaseMat01.VisualRowCount - 2, "10");
								oDS_PS_MM130L.SetValue("U_OutItmCd", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("OutItmCd").Cells.Item(i).Specific.Value.ToString().Trim());
								oDS_PS_MM130L.SetValue("U_OutItmNm", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("OutItmNm").Cells.Item(i).Specific.Value.ToString().Trim());
							}
							else if (oRadioGrp == "B")
							{
								oDS_PS_MM130L.SetValue("U_OutGbn", oBaseMat01.VisualRowCount - 2, "20");
								oDS_PS_MM130L.SetValue("U_OutItmCd", oBaseMat01.VisualRowCount - 2, "");
								oDS_PS_MM130L.SetValue("U_OutItmNm", oBaseMat01.VisualRowCount - 2, "");
							}
							oDS_PS_MM130L.SetValue("U_OutQty", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("PosQty").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_OutWt", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("PosWt").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_OutWhCd", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("OutWhCd").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_OutWhNm", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("OutWhNm").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_InWhCd", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("InWhCd").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_InWhNm", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("InWhNm").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_PP030HNo", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_PP030MNo", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("PP030MNo").Cells.Item(i).Specific.Value.ToString().Trim());

							oDS_PS_MM130L.SetValue("U_CpCode", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("CpCode").Cells.Item(i).Specific.Value.ToString().Trim());
							oDS_PS_MM130L.SetValue("U_CpName", oBaseMat01.VisualRowCount - 2, oMat.Columns.Item("CpName").Cells.Item(i).Specific.Value.ToString().Trim());
							j += 1;
						}
						else
						{
						}
					}

					oBaseMat01.LoadFromDataSource();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM130L);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oBaseMat01);
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
					if (pVal.ItemUID == "Btn02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_MM131_Select_Data();
						}
					}
					else if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_MM131_SetBaseForm(); //부모폼에입력
							oForm.Close();
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
			string RadioGrp = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CpCode", "");
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string ItemCode01;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.BeforeAction == true)
					{
						if (pVal.ItemChanged == true)
						{
							if (pVal.ItemUID == "Mat01")
							{
								if (pVal.ColUID == "SelQty")
								{
									if (Convert.ToDouble(oMat.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) <= 0)
									{
										oMat.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = "0";
										oMat.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = "0";
									}
									else
									{
										ItemCode01 = oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
										//EA자체품
										if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "101" || dataHelpClass.GetItem_SbasUnit(ItemCode01) == "601")
										{
											oMat.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim(); //EAUOM
										}
										else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "102")
										{
											oMat.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToString(Convert.ToDouble(oMat.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(ItemCode01))); //KGSPEC
										}
										else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "201")
										{
											oMat.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value =
												Convert.ToString((Convert.ToDouble(dataHelpClass.GetItem_Spec1(ItemCode01))
																  - Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)))
																  * Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)) * 0.02808
																  * Convert.ToDouble(dataHelpClass.GetItem_Spec3(ItemCode01)) / 1000)
																  * Convert.ToDouble(oMat.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()); //KG단중
										}
										else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "202")
										{
											oMat.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value =
												Convert.ToString(System.Math.Round(Convert.ToDouble(oMat.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim())
																				   * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)) / 1000, 0)); //KG입력
										}
										else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "203")
										{
											oMat.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 1;
											oMat.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 1;
										}
										else
										{
											oMat.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 1;
										}
									}
									oForm.Update();
								}
								else if (pVal.ColUID == "PosQty")
								{
									if (Convert.ToDouble(oMat.Columns.Item("PosQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) > 0)
									{
										sQry = "Select U_ObasUnit FROM OITM WHERE ItemCode = '" + oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
										oRecordSet.DoQuery(sQry);

										sQry = "Select OnHand, U_Qty FROM OITW WHERE ItemCode = '" + oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "' AND WhsCode = '101'";
										oRecordSet02.DoQuery(sQry);

										if (oRecordSet.Fields.Item(0).Value.ToString().Trim().Substring(0, 1) == "1")
										{
											if (Convert.ToDouble(oRecordSet02.Fields.Item(0).Value.ToString().Trim()) > 0 && !string.IsNullOrEmpty(oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
											{
												oMat.Columns.Item("PosWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("PosQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
											}
											else
											{
												oMat.Columns.Item("PosWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("PosQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
											}
										}
										else if (oRecordSet.Fields.Item(0).Value.ToString().Trim().Substring(0, 1) == "2")
										{
											if (Convert.ToDouble(oRecordSet02.Fields.Item(1).Value.ToString().Trim()) > 0 && !string.IsNullOrEmpty(oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
											{
												oMat.Columns.Item("PosWt").Cells.Item(pVal.Row).Specific.Value = Convert.ToString((Convert.ToDouble(oRecordSet02.Fields.Item(0).Value.ToString().Trim()) / Convert.ToDouble(oRecordSet02.Fields.Item(1).Value.ToString().Trim())) * Convert.ToDouble(oMat.Columns.Item("PosQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()));
											}
											else
											{
												oMat.Columns.Item("PosWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("PosQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
											}
										}
										else
										{
										}

										oMat.Columns.Item("PosWt").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
									}
								}
							}
							else
							{
								sQry = "Select U_CpName From [@PS_PP001L] Where U_CpCode = '" + oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								oForm.Items.Item("CpName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							}
						}
					}
					else if (pVal.BeforeAction == false)
					{
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
				oForm.Freeze(false);
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
			int i;
			int sCount;
			int sSeq;
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
						if (pVal.ItemUID == "ItmBsort")
						{
							sCount = oForm.Items.Item("ItmMsort").Specific.ValidValues.Count;
							sSeq = sCount;

							for (i = 1; i <= sCount; i++)
							{
								oForm.Items.Item("ItmMsort").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
								sSeq -= 1;
							}

							oForm.Items.Item("ItmMsort").Specific.ValidValues.Add("", "전체");
							sQry = "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] Where U_rCode = '" + oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							while (!oRecordSet.EoF)
							{
								oForm.Items.Item("ItmMsort").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oRecordSet.MoveNext();
							}
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
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
					if (pVal.ItemUID == "Mat01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.Row > 0)
							{
								oMat.SelectRow(pVal.Row, true, false);
								oMatRow01 = pVal.Row;
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
		/// Raise_EVENT_CHOOSE_FROM_LIST
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "ItemCode")
					{
						if (oDataTable01 == null)
						{
						}
						else
						{
							oForm.DataSources.UserDataSources.Item("ItemCode").Value = oDataTable01.Columns.Item(0).Cells.Item(0).Value.ToString().Trim();
							oForm.DataSources.UserDataSources.Item("ItemName").Value = oDataTable01.Columns.Item(1).Cells.Item(0).Value.ToString().Trim();
						}
					}
					oForm.Update();
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
			//		System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm); SUB 에서 닫으면 에러남
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM131L);
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
