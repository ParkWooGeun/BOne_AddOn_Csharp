using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	///  제품BOM 조회 재고장
	/// </summary>
	internal class PS_SM030 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
			
		private SAPbouiCOM.DBDataSource oDS_PS_SM030H;  //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SM030L;  //등록라인

		//부모폼
		private SAPbouiCOM.Form oBaseForm01;
		private string oBaseItemUID01;
		private string oBaseColUID01;
		private int oBaseColRow01;
		private string oBaseItemCode01;

		private string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private int oMat01Row01;
		private int oMat02Row02;

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oForm02"></param>
		/// <param name="oItemUID02"></param>
		/// <param name="oColUID02"></param>
		/// <param name="oColRow02"></param>
		/// <param name="oItemCode02"></param>
		public void LoadForm(SAPbouiCOM.Form oForm02, string oItemUID02, string oColUID02, int oColRow02, string oItemCode02)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SM030.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SM030_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SM030");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				oBaseForm01 = oForm02;
				oBaseItemUID01 = oItemUID02;
				oBaseColUID01 = oColUID02;
				oBaseColRow01 = oColRow02;
				oBaseItemCode01 = oItemCode02;

				PS_SM030_CreateItems();
				PS_SM030_ComboBox_Setting();
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
		/// PS_SM030_CreateItems
		/// </summary>
		private void PS_SM030_CreateItems()
		{
			try
			{
				oDS_PS_SM030H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_SM030L = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();
				oMat02.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 30);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				if (!string.IsNullOrEmpty(oBaseItemCode01))
				{
					oForm.Items.Item("ItemCode").Specific.Value = oBaseItemCode01;
				}

				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				oForm.DataSources.UserDataSources.Add("StockType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("StockType").Specific.DataBind.SetBound(true, "", "StockType");

				oForm.DataSources.UserDataSources.Add("TradeType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("TradeType").Specific.DataBind.SetBound(true, "", "TradeType");

				oForm.DataSources.UserDataSources.Add("ItemGpCd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ItemGpCd").Specific.DataBind.SetBound(true, "", "ItemGpCd");

				oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

				oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");

				oForm.DataSources.UserDataSources.Add("Size", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Size").Specific.DataBind.SetBound(true, "", "Size");

				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				oForm.DataSources.UserDataSources.Add("Mark", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Mark").Specific.DataBind.SetBound(true, "", "Mark");

				oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

				oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");
				oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");

				oForm.DataSources.UserDataSources.Add("Check01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				oForm.Items.Item("ChkBox01").Specific.ValOn = "Y";
				oForm.Items.Item("ChkBox01").Specific.ValOff = "N";
				oForm.Items.Item("ChkBox01").Specific.DataBind.SetBound(true, "", "Check01");
				oForm.DataSources.UserDataSources.Item("Check01").Value = "N";

				oForm.Items.Item("Mat01").Enabled = false;
				oForm.Items.Item("Mat02").Enabled = false;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SM030_ComboBox_Setting
		/// </summary>
		private void PS_SM030_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Combo_ValidValues_Insert("PS_SM030", "StockType", "", "1", "재고있는품목");
				dataHelpClass.Combo_ValidValues_Insert("PS_SM030", "StockType", "", "2", "전체");
				dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("StockType").Specific, "PS_SM030", "StockType", false);
				oForm.Items.Item("StockType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);

				dataHelpClass.Combo_ValidValues_Insert("PS_SM030", "TradeType", "", "", "전체");
				dataHelpClass.Combo_ValidValues_Insert("PS_SM030", "TradeType", "", "1", "일반");
				dataHelpClass.Combo_ValidValues_Insert("PS_SM030", "TradeType", "", "2", "임가공");
				dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("TradeType").Specific, "PS_SM030", "TradeType", false);

				oForm.Items.Item("ItmBsort").Specific.ValidValues.Add("선택", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", false, false);

				oForm.Items.Item("ItmMsort").Specific.ValidValues.Add("선택", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", false, false);

				oForm.Items.Item("ItemType").Specific.ValidValues.Add("선택", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", false, false);

				oForm.Items.Item("Mark").Specific.ValidValues.Add("선택", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", false, false);

				oForm.Items.Item("ItemGpCd").Specific.ValidValues.Add("선택", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemGpCd").Specific, "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]", "", false, false);

				oForm.Items.Item("ItmBsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("ItmMsort").Specific.Select("선택", SAPbouiCOM.BoSearchKey.psk_ByValue);
				//oForm.Items.Item("ItmMsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("Mark").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("ItemGpCd").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				oForm.Items.Item("TradeType").Enabled = false;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SM030_MTX01
		/// </summary>
		private void PS_SM030_MTX01()
		{
			int i;
			string sQry;
			string errMessage = string.Empty;

			string Param01;
			string Param02;
			string Param03;
			string Param04;
			string Param05;
			string Param06;
			string Param07;
			string Param08;
			string Param09;
            string Param10;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				Param01 = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				Param02 = oForm.Items.Item("StockType").Specific.Selected.Value.ToString().Trim();
				Param03 = oForm.Items.Item("TradeType").Specific.Selected.Value.ToString().Trim();
				Param04 = oForm.Items.Item("ItmBsort").Specific.Selected.Value.ToString().Trim();
				Param05 = oForm.Items.Item("ItmMsort").Specific.Selected.Value.ToString().Trim();
				Param06 = oForm.Items.Item("Size").Specific.Value.ToString().Trim();
				Param07 = oForm.Items.Item("ItemType").Specific.Selected.Value.ToString().Trim();
				Param08 = oForm.Items.Item("Mark").Specific.Selected.Value.ToString().Trim();
				Param09 = oForm.Items.Item("ItemName").Specific.Value.ToString().Trim();
				Param10 = oForm.Items.Item("ItemGpCd").Specific.Selected.Value.ToString().Trim();

				if (oBaseForm01 == null)
				{
					sQry = "EXEC PS_SM030_01 '" + Param01 + "','','','" + Param02 + "','','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
				}
				else if (oBaseForm01.Type == Convert.ToDouble("149") || oBaseForm01.Type == Convert.ToDouble("139") || oBaseForm01.Type == Convert.ToDouble("140") 
					     || oBaseForm01.Type == Convert.ToDouble("180") || oBaseForm01.Type == Convert.ToDouble("133") || oBaseForm01.Type == Convert.ToDouble("179") 
						 || oBaseForm01.Type == Convert.ToDouble("60091"))
				{   //판매Y,구매,재고타입(1:재고있는것만,2:전체),거래타입(1:일반,2:임가공)
					sQry = "EXEC PS_SM030_01 '" + Param01 + "','Y','','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
				} 
				else if (oBaseForm01.Type == Convert.ToDouble("142") || oBaseForm01.Type == Convert.ToDouble("143") || oBaseForm01.Type == Convert.ToDouble("182") || oBaseForm01.Type == Convert.ToDouble("141") || oBaseForm01.Type == Convert.ToDouble("181") || oBaseForm01.Type == Convert.ToDouble("60092"))
				{   //판매,구매Y
					sQry = "EXEC PS_SM030_01 '" + Param01 + "','','Y','" + Param02 + "','" + Param03 + "','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
				}
				else 
				{
					sQry = "EXEC PS_SM030_01 '" + Param01 + "','','','" + Param02 + "','','" + Param04 + "','" + Param05 + "','" + Param06 + "','" + Param07 + "','" + Param08 + "','" + Param09 + "','" + Param10 + "'";
				}
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				oMat02.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

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

				ProgressBar01.Text = "조회중...";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i != 0)
					{
						oDS_PS_SM030H.InsertRecord(i);
					}
					oDS_PS_SM030H.Offset = i;
					oDS_PS_SM030H.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_SM030H.SetValue("U_ColReg01", i, Convert.ToString(false));
					oDS_PS_SM030H.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
					oDS_PS_SM030H.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
					oDS_PS_SM030H.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("OnHand").Value.ToString().Trim());
					oDS_PS_SM030H.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("IsCommited").Value.ToString().Trim());
					oDS_PS_SM030H.SetValue("U_ColQty03", i, oRecordSet.Fields.Item("OnOrder").Value.ToString().Trim());
					oDS_PS_SM030H.SetValue("U_ColQty04", i, oRecordSet.Fields.Item("OnEnabled").Value.ToString().Trim());
					oDS_PS_SM030H.SetValue("U_ColNum01", i, Convert.ToString(0));

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
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
		/// PS_SM030_MTX02
		/// </summary>
		private void PS_SM030_MTX02()
		{
			int i;
			string sQry;
			string Param01;
			string errMessage = string.Empty;

			double rate_Renamed ;
			double StdQty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				Param01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value.ToString().Trim();
				StdQty = Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(oMat01Row01).Specific.Value.ToString().Trim());

				sQry = "EXEC PS_SM030_02 '" + Param01 + "'";
				oRecordSet.DoQuery(sQry);

				oMat02.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Items.Item("Mat02").Enabled = false;
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}
				else
				{
					oForm.Items.Item("Mat02").Enabled = true;
				}

				//품목이 멀티이면 수량,중량필드 비활성화
				if (dataHelpClass.GetItem_ItmBsort(Param01) == "104" || dataHelpClass.GetItem_ItmBsort(Param01) == "302")  //멀티
				{
					oMat02.Columns.Item("SelQty").Editable = false;
					oMat02.Columns.Item("SelWeight").Editable = false;
					
				}
				else  //그외품목은 수량선택가능
				{
					oMat02.Columns.Item("SelQty").Editable = true;
					oMat02.Columns.Item("SelWeight").Editable = true;
				}

				ProgressBar01.Text = "조회중...";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i != 0)
					{
						oDS_PS_SM030L.InsertRecord(i);
					}
					oDS_PS_SM030L.Offset = i;
					oDS_PS_SM030L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_SM030L.SetValue("U_ColReg01", i, "Y");
					oDS_PS_SM030L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
					oDS_PS_SM030L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
					oDS_PS_SM030L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("Size").Value.ToString().Trim());
					oDS_PS_SM030L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("Ut").Value.ToString().Trim());
					oDS_PS_SM030L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("OutSize").Value.ToString().Trim());

					if (StdQty != 0)
					{
						oDS_PS_SM030L.SetValue("U_ColNum01", i, Convert.ToString(StdQty));
						rate_Renamed = StdQty / Convert.ToDouble(oRecordSet.Fields.Item("StdQty").Value.ToString().Trim());

						oDS_PS_SM030L.SetValue("U_ColNum02", i, Convert.ToString(Convert.ToInt32(rate_Renamed * oRecordSet.Fields.Item("SelQty").Value) * -1) * -1);
						oDS_PS_SM030L.SetValue("U_ColQty01", i, Convert.ToString((System.Math.Round(Convert.ToInt32((rate_Renamed * oRecordSet.Fields.Item("SelQty").Value) * -1) * -1) / Convert.ToInt32(oRecordSet.Fields.Item("SelQty").Value)) * Convert.ToInt32(oRecordSet.Fields.Item("SelWeight").Value), 2));
					}
					else
					{
						oDS_PS_SM030L.SetValue("U_ColNum01", i, oRecordSet.Fields.Item("StdQty").Value.ToString().Trim());
						oDS_PS_SM030L.SetValue("U_ColNum02", i, oRecordSet.Fields.Item("SelQty").Value.ToString().Trim());
						oDS_PS_SM030L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("SelWeight").Value.ToString().Trim());
					}
					oDS_PS_SM030L.SetValue("U_ColNum03", i, oRecordSet.Fields.Item("StdQty").Value.ToString().Trim());
					oDS_PS_SM030L.SetValue("U_ColNum04", i, oRecordSet.Fields.Item("SelQty").Value.ToString().Trim());
					oDS_PS_SM030L.SetValue("U_ColNum05", i, oRecordSet.Fields.Item("StdQty").Value.ToString().Trim());
					oDS_PS_SM030L.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("SelWeight").Value.ToString().Trim());
					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();
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
		/// PS_SM030_SetBaseForm
		/// </summary>
		private void PS_SM030_SetBaseForm()
		{
			int i;
			string ItemCode01;
			SAPbouiCOM.Matrix oBaseMat01;
			double Qty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (oBaseForm01.Type == Convert.ToDouble("133") || oBaseForm01.Type == Convert.ToDouble("139") || oBaseForm01.Type == Convert.ToDouble("140") 
					|| oBaseForm01.Type == Convert.ToDouble("141") || oBaseForm01.Type == Convert.ToDouble("142") || oBaseForm01.Type == Convert.ToDouble("143") 
					|| oBaseForm01.Type == Convert.ToDouble("149") || oBaseForm01.Type == Convert.ToDouble("179") || oBaseForm01.Type == Convert.ToDouble("180") 
					|| oBaseForm01.Type == Convert.ToDouble("181") || oBaseForm01.Type == Convert.ToDouble("182") || oBaseForm01.Type == Convert.ToDouble("60091") 
					|| oBaseForm01.Type == Convert.ToDouble("60092"))
				{
					oBaseMat01 = oBaseForm01.Items.Item("38").Specific;
					for (i = 1; i <= oMat01.RowCount; i++)
					{
						if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							if (Convert.ToDouble(oMat01.Columns.Item("SelWeight").Cells.Item(i).Specific.Value.ToString().Trim()) <= 0)
							{
								//중량이 선택되지 않은품목
								ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
								oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.Value = ItemCode01;	//품목
								oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								oBaseColRow01 += 1;
							}
							else
							{
								ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
								oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.Value = ItemCode01; //품목
								oBaseMat01.Columns.Item("U_Qty").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim()); //수량 수량을 변경하면 중량이 자동변경된다.
								oBaseMat01.Columns.Item("U_Unweight").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)); //단중
								oBaseMat01.Columns.Item("11").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat01.Columns.Item("SelWeight").Cells.Item(i).Specific.Value.ToString().Trim()); //중량
								oBaseMat01.Columns.Item("14").Cells.Item(oBaseColRow01).Specific.Value = dataHelpClass.GetValue("EXEC PS_SBO_GETPRICE '" + oBaseForm01.Items.Item("4").Specific.Value.ToString().Trim() + "','" + ItemCode01 + "'", 0, 1);
								oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								oBaseColRow01 += 1;
							}
						}
					}
					//배치선택품목
					for (i = 1; i <= oMat02.RowCount; i++)
					{
						if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							if (Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value.ToString().Trim()) <= 0)
							{
								//중량이 선택되지 않은품목
								ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value.ToString().Trim();
								oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.Value = ItemCode01; //품목
								oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								oBaseColRow01 += 1;
							}
							else
							{
								ItemCode01 = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value.ToString().Trim();
								oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.Value = ItemCode01; //품목
								oBaseMat01.Columns.Item("U_Qty").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim());//수량 수량을 변경하면 중량이 자동변경된다.
								oBaseMat01.Columns.Item("U_Unweight").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)); //단중
								oBaseMat01.Columns.Item("11").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value.ToString().Trim()); //중량
								oBaseMat01.Columns.Item("14").Cells.Item(oBaseColRow01).Specific.Value = dataHelpClass.GetValue("EXEC PS_SBO_GETPRICE '" + oBaseForm01.Items.Item("4").Specific.Value.ToString().Trim() + "','" + ItemCode01 + "'", 0, 1);
								oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								oBaseColRow01 += 1;
							}
						}
					}
				}
				else if (oBaseForm01.TypeEx == "720")
				{
					oBaseMat01 = oBaseForm01.Items.Item("13").Specific;
					//매트릭스
					//품목선택품목
					for (i = 1; i <= oMat01.RowCount; i++)
					{
						if (oMat01.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim(); //품목
							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oBaseColRow01 += 1;
						}
					}
					//배치선택품목
					for (i = 1; i <= oMat02.RowCount; i++)
					{
						if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01).Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value.ToString().Trim(); //품목
							oBaseMat01.Columns.Item("1").Cells.Item(oBaseColRow01 + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oBaseColRow01 += 1;
						}
					}
				}
				else if (oBaseForm01.TypeEx == "PS_MM090")  //자재기타출고
				{
					oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
					//매트릭스
					//배치선택품목
					for (i = 1; i <= oMat02.RowCount; i++)
					{
						if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							oBaseMat01.Columns.Item("ItemName").Cells.Item(oBaseColRow01).Specific.Value = oMat02.Columns.Item("ItemName").Cells.Item(i).Specific.Value.ToString().Trim(); //품목
							oBaseMat01.Columns.Item("Size").Cells.Item(oBaseColRow01).Specific.Value = oMat02.Columns.Item("Size").Cells.Item(i).Specific.Value.ToString().Trim();		   //규격
							oBaseMat01.Columns.Item("Unit").Cells.Item(oBaseColRow01).Specific.Value = oMat02.Columns.Item("Ut").Cells.Item(i).Specific.Value.ToString().Trim();		   //단위
							oBaseMat01.Columns.Item("Qty").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim());		//수량
							oBaseMat01.Columns.Item("Weight").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value.ToString().Trim());	//재적

							if (i == 1)
							{
								oBaseForm01.Items.Item("Comment1").Specific.Value = oMat02.Columns.Item("OutSize").Cells.Item(i).Specific.Value.ToString().Trim() + " (" + oMat02.Columns.Item("StdQty").Cells.Item(i).Specific.Value.ToString().Trim() + "EA" + ")"; //수량
							}
							oBaseColRow01 += 1;
							oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked = false;
						}
					}

					
				}
				else if (oBaseForm01.TypeEx == "PS_MM005") //자재청구등록
				{
					oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
					//매트릭스
					//배치선택품목
					for (i = 1; i <= oMat02.RowCount; i++)
					{
						if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();	//품목
							oBaseMat01.Columns.Item("Qty").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim()); //수량
							oBaseMat01.Columns.Item("Weight").Cells.Item(oBaseColRow01).Specific.Value = (Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value.ToString().Trim())); //재적

							if (i == 1)
							{
								oBaseMat01.Columns.Item("Comments").Cells.Item(oBaseColRow01).Specific.Value = oMat02.Columns.Item("OutSize").Cells.Item(i).Specific.Value.ToString().Trim() + " (" + oMat02.Columns.Item("StdQty").Cells.Item(i).Specific.Value.ToString().Trim() + "EA";	//수량
							}
							oBaseColRow01 += 1;
							oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked = false;
						}
					}
					
				}
				else if (oBaseForm01.TypeEx == "PS_MM135")  //포장사업부 외주반출등록
				{
					oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific;
					//매트릭스
					Qty = Convert.ToDouble(oBaseForm01.Items.Item("Qty").Specific.Value.ToString().Trim());
					//배치선택품목
					for (i = 1; i <= oMat02.RowCount; i++)
					{
						if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							oBaseMat01.Columns.Item("Qty").Cells.Item(oBaseColRow01).Specific.Value = Qty; //제품수량
							oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oBaseItemCode01; //제품품목
							oBaseMat01.Columns.Item("OutItmCd").Cells.Item(oBaseColRow01).Specific.Value = oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim(); //품목

							if (Qty <= 0)
							{
								oBaseMat01.Columns.Item("OutQty").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim());	//수량
								oBaseMat01.Columns.Item("OutWt").Cells.Item(oBaseColRow01).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value.ToString().Trim());	//재적
							}
							else
							{
								oBaseMat01.Columns.Item("OutQty").Cells.Item(oBaseColRow01).Specific.Value = System.Math.Round(Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim()) / Convert.ToDouble(oMat02.Columns.Item("StdQty").Cells.Item(i).Specific.Value.ToString().Trim()) * Qty, 0);   //수량
								oBaseMat01.Columns.Item("OutWt").Cells.Item(oBaseColRow01).Specific.Value = System.Math.Round(Convert.ToDouble(oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value.ToString().Trim()) / Convert.ToDouble(oMat02.Columns.Item("StdQty").Cells.Item(i).Specific.Value.ToString().Trim()) * Qty, 2); //재적
							}

							oBaseColRow01 += 1;
							oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked = false;
						}
					}
				}
				else if (oBaseForm01.TypeEx == "PS_MM097")  //포장 원자재 재고조사등록
				{
					oBaseMat01 = oBaseForm01.Items.Item("Mat01").Specific; //매트릭스
					//배치선택품목
					for (i = 1; i <= oMat02.RowCount; i++)
					{
						if (oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked == true)
						{
							oBaseMat01.Columns.Item("ItemCode").Cells.Item(oBaseColRow01).Specific.Value = oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim(); //품목

							oBaseColRow01 += 1;
							oMat02.Columns.Item("CHK").Cells.Item(i).Specific.Checked = false;
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
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
				if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					if (pVal.ItemUID == "Button01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_SM030_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
					if (pVal.ItemUID == "Button02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_SM030_SetBaseForm();	//부모폼에입력
							if (oForm.DataSources.UserDataSources.Item("Check01").Value.ToString().Trim() == "N")
							{
								oForm.Close();
							}
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
					if (pVal.ItemUID == "PS_SM030")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int i;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "ItmBsort")
					{
						for (i = 0; i <= oForm.Items.Item("ItmMsort").Specific.ValidValues.Count - 1; i++)
						{
							oForm.Items.Item("ItmMsort").Specific.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
						}
						oForm.Items.Item("ItmMsort").Specific.ValidValues.Add("선택", "선택");
						dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] WHERE U_rCode = '" + oForm.Items.Item("ItmBsort").Specific.Selected.Value.ToString().Trim() + "' ORDER BY U_Code", "", false, false);
						if (oForm.Items.Item("ItmMsort").Specific.ValidValues.Count > 0)
						{
							//oForm.Items.Item("ItmMsort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
							oForm.Items.Item("ItmMsort").Specific.Select("선택", SAPbouiCOM.BoSearchKey.psk_ByValue);
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
				oForm.Freeze(false);
			}
		}

		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
								oMat01.SelectRow(pVal.Row, true, false);
								oMat01Row01 = pVal.Row;
								// ???
								if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) == "Y") //배치를 사용하는품목
								{
									PS_SM030_MTX02(); 
								}
								else if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) == "N") //배치를 사용하지 않는품목
								{
									PS_SM030_MTX02();
								}
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
					if (pVal.ItemUID == "Mat02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.Row > 0)
							{
								oMat02.SelectRow(pVal.Row, true, false);
								oMat02Row02 = pVal.Row;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
					if (pVal.ItemUID == "Opt01")
					{
						oForm.Settings.MatrixUID = "Mat01";
						oForm.Settings.Enabled = true;
						oForm.Settings.EnableRowFormat = true;
					}
					if (pVal.ItemUID == "Opt02")
					{
						oForm.Settings.MatrixUID = "Mat02";
						oForm.Settings.Enabled = true;
						oForm.Settings.EnableRowFormat = true;
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
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row == 0)
						{
							oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
							oMat01.FlushToDataSource();
						}
					}
					if (pVal.ItemUID == "Mat02")
					{
						if (pVal.Row == 0)
						{
							oMat02.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
							oMat02.FlushToDataSource();
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
		/// Raise_EVENT_MATRIX_LINK_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (!string.IsNullOrEmpty(oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.String))
						{
							PS_MM002 oTempClass = new PS_MM002();
							oTempClass.LoadForm(dataHelpClass.User_BPLID(), oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.String);
							BubbleEvent = false;
						}
						else
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			double SelQty;
			double rate_Renamed;
			int i;
			string ItemCode01;
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
							if (pVal.ColUID == "SelQty")  //기준수량 수정에따른 BOM수량 변동
							{
								if ( Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) <= 0)
								{
									oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = 0;
									oMat01.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 0;

									for (i = 1; i <= oMat02.RowCount; i++)
									{
										oMat02.Columns.Item("StdQty").Cells.Item(i).Specific.Value = oMat02.Columns.Item("OStdQty").Cells.Item(i).Specific.Value.ToString().Trim();
										oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value = oMat02.Columns.Item("OSelQty").Cells.Item(i).Specific.Value.ToString().Trim();
										oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value = oMat02.Columns.Item("OSelWeight").Cells.Item(i).Specific.Value.ToString().Trim();
									}
								}
								else
								{
									SelQty = Convert.ToDouble(oMat01.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());

									for (i = 1; i <= oMat02.RowCount; i++)
									{
										if (SelQty == 0)
										{
											rate_Renamed = 1;
										}
										else
										{
											rate_Renamed = SelQty / Convert.ToDouble(oMat02.Columns.Item("OStdQty").Cells.Item(i).Specific.Value.ToString().Trim());
										}

										oMat02.Columns.Item("StdQty").Cells.Item(i).Specific.Value = SelQty;
										oMat02.Columns.Item("SelQty").Cells.Item(i).Specific.Value = Convert.ToDouble(rate_Renamed * Convert.ToDouble(oMat02.Columns.Item("OSelQty").Cells.Item(i).Specific.Value.ToString().Trim()) * -1) * -1;
										oMat02.Columns.Item("SelWeight").Cells.Item(i).Specific.Value = System.Math.Round(((Convert.ToDouble((rate_Renamed * Convert.ToDouble(oMat02.Columns.Item("OSelQty").Cells.Item(i).Specific.Value.ToString().Trim())) * -1) * -1) / Convert.ToDouble(oMat02.Columns.Item("OSelQty").Cells.Item(i).Specific.Value.ToString().Trim())) * Convert.ToDouble(oMat02.Columns.Item("OSelWeight").Cells.Item(i).Specific.Value.ToString().Trim()), 2);
									}
								}
								oForm.Update();
							}
						}
						else if (pVal.ItemUID == "Mat02")
						{
							if (pVal.ColUID == "SelQty")
							{
								if (Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) <= 0)
								{
									oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = 0;
									oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 0;
								}
								else
								{
									ItemCode01 = oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
									
									if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "101") //EA자체품
									{
										oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									}
									else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "102")  //EAUOM
									{
										oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(ItemCode01));
									}
									else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "201") //KGSPEC
									{
										oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(ItemCode01)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(ItemCode01)) / 1000) * Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									}
									else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "202") //KG단중
									{
										oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)) / 1000, 2);
									}
									else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "203") //KG입력
									{
									}
								}
								oForm.Update();
							}
							if (pVal.ColUID == "StdQty")
							{
								if (Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) <= 0)
								{
									oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = 0;
									oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = 0;
								}
								else
								{
									ItemCode01 = oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
									
									if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "101") //EA자체품
									{
										oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									}
									else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "102")  //EAUOM
									{
										oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(ItemCode01));
									}
									else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "201")  //KGSPEC
									{
										oMat02.Columns.Item("SelWeight").Cells.Item(pVal.Row).Specific.Value = (Convert.ToDouble(dataHelpClass.GetItem_Spec1(ItemCode01)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(ItemCode01)) / 1000) * Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									}
									else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "202")  //KG단중 
									{
										oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value = Convert.ToDouble(oMat02.Columns.Item("StdQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) / Convert.ToDouble(oMat02.Columns.Item("StdQty2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) * Convert.ToDouble(oMat02.Columns.Item("SelQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									}
									else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "203") //KG입력
									{
									}
								}
								oForm.Update();
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
            {
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					oForm.Items.Item("Mat01").Top = 70;
					oForm.Items.Item("Mat01").Height = (oForm.Height / 2) - 70;
					oForm.Items.Item("Mat01").Left = 7;
					oForm.Items.Item("Mat01").Width = oForm.Width - 21;
					oForm.Items.Item("Mat02").Top = (oForm.Height / 2) + 10;
					oForm.Items.Item("Mat02").Height = (oForm.Height / 2) - 75;
					oForm.Items.Item("Mat02").Left = 7;
					oForm.Items.Item("Mat02").Width = oForm.Width - 21;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
				if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02")
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
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);  //해제하면 에러남 (확인(종료)시)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SM030H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SM030L);
                }
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_FormMenuEvent
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
