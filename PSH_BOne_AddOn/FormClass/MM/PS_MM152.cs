using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 가공품 출고
	/// </summary>
	internal class PS_MM152 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_MM152H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM152L; //등록라인

		private struct PS_PP040s
		{
			public string PP030HNo;
			public string PP030MNo;
			public string OrdNum;
			public string OrdGbn;
			public string ItemCode;
			public string ItemName;
			public string Sequence;
			public string CpCode;
			public string CpName;
			public string Chk;
		}

		private PS_PP040s[] PS_PP040_Renamed;

		private struct PS_PP040DocEntrys
		{
			public string PP040DocEntry;
		}

		private PS_PP040DocEntrys[] PS_PP040DocEntry;

		private int oDocEntryNext;
		private string oCheck;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM152.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM152_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM152");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_MM152_CreateItems();
				PS_MM152_ComboBox_Setting();
				PS_MM152_Initialization();
				PS_MM152_FormClear();
				PS_MM152_Add_MatrixRow(0, true);
				PS_MM152_FormItemEnabled();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1287", true);  // 복제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
		/// PS_MM152_CreateItems
		/// </summary>
		private void PS_MM152_CreateItems()
		{
			try
			{
				oDS_PS_MM152H = oForm.DataSources.DBDataSources.Item("@PS_MM152H");
				oDS_PS_MM152L = oForm.DataSources.DBDataSources.Item("@PS_MM152L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM152_ComboBox_Setting
		/// </summary>
		private void PS_MM152_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				oMat.Columns.Item("OutGbn").ValidValues.Add("10", "원재료");
				oMat.Columns.Item("OutGbn").ValidValues.Add("20", "재공");

				oMat.Columns.Item("ReStatus").ValidValues.Add("Y", "완료");
				oMat.Columns.Item("ReStatus").ValidValues.Add("N", "미완료");
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
		/// PS_MM152_Initialization
		/// </summary>
		private void PS_MM152_Initialization()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oDS_PS_MM152H.SetValue("U_BPLId", 0, "1");
				oDS_PS_MM152H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
				oDS_PS_MM152H.SetValue("U_OKYNC", 0, "N");

				if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
				{
					oDS_PS_MM152H.SetValue("U_CardCode", 0, PSH_Globals.oCompany.UserName);

					sQry = "Select CardName From OCRD Where CardCode = '" + oDS_PS_MM152H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);

					oDS_PS_MM152H.SetValue("U_CardName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
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
		/// PS_MM152_FormClear
		/// </summary>
		private void PS_MM152_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM152'", "");

				if (Convert.ToDouble(DocEntry) == 0)
				{
					oForm.Items.Item("DocEntry").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM152_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM152_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_MM152L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_MM152L.Offset = oRow;
				oDS_PS_MM152L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM152_FormItemEnabled
		/// </summary>
		private void PS_MM152_FormItemEnabled()
		{
			try
			{
				oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					if (PSH_Globals.oCompany.UserName == "66302" || PSH_Globals.oCompany.UserName == "71090")
					{
						oForm.Items.Item("BPLId").Enabled = false;
						oForm.Items.Item("CardCode").Enabled = false;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("OKYNC").Enabled = false;
						oForm.Items.Item("DocEntry").Enabled = false;
						oMat.Columns.Item("OtDocLin").Editable = true;
						oMat.Columns.Item("OutQty").Editable = true;
						oMat.Columns.Item("OutWt").Editable = true;
						oMat.Columns.Item("NQty").Editable = true;
						oMat.Columns.Item("NWeight").Editable = true;
						oMat.Columns.Item("CPWt").Editable = true;
						oMat.Columns.Item("CPWtName").Editable = true;
						oMat.Columns.Item("MUseQty").Editable = true;
						oMat.Columns.Item("MUseWt").Editable = true;
						oMat.Columns.Item("ReStatus").Editable = true;
						oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					}
					else
					{
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("OKYNC").Enabled = false;
						oForm.Items.Item("DocEntry").Enabled = false;
						oMat.Columns.Item("OtDocLin").Editable = true;
						oMat.Columns.Item("OutQty").Editable = true;
						oMat.Columns.Item("OutWt").Editable = true;
						oMat.Columns.Item("NQty").Editable = true;
						oMat.Columns.Item("NWeight").Editable = true;
						oMat.Columns.Item("CPWt").Editable = true;
						oMat.Columns.Item("CPWtName").Editable = true;
						oMat.Columns.Item("MUseQty").Editable = true;
						oMat.Columns.Item("MUseWt").Editable = true;
						oMat.Columns.Item("ReStatus").Editable = true;
					}
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("OKYNC").Enabled = true;
					oForm.Items.Item("DocEntry").Enabled = true;
					oMat.Columns.Item("OtDocLin").Editable = false;
					oMat.Columns.Item("OutQty").Editable = false;
					oMat.Columns.Item("OutWt").Editable = false;
					oMat.Columns.Item("NQty").Editable = false;
					oMat.Columns.Item("NWeight").Editable = false;
					oMat.Columns.Item("CPWt").Editable = false;
					oMat.Columns.Item("CPWtName").Editable = false;
					oMat.Columns.Item("MUseQty").Editable = false;
					oMat.Columns.Item("MUseWt").Editable = false;
					oMat.Columns.Item("ReStatus").Editable = false;
					oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					if (PSH_Globals.oCompany.UserName == oForm.Items.Item("CardCode").Specific.Value.ToString().Trim())
					{
						if (oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() != "Y")
						{
							oForm.Items.Item("BPLId").Enabled = false;
							oForm.Items.Item("CardCode").Enabled = false;
							oForm.Items.Item("DocDate").Enabled = true;
							oForm.Items.Item("OKYNC").Enabled = false;
							oForm.Items.Item("DocEntry").Enabled = false;
							oMat.Columns.Item("OtDocLin").Editable = false;
							oMat.Columns.Item("OutQty").Editable = false;
							oMat.Columns.Item("OutWt").Editable = false;
							oMat.Columns.Item("NQty").Editable = false;
							oMat.Columns.Item("NWeight").Editable = false;
							oMat.Columns.Item("CPWt").Editable = true;
							oMat.Columns.Item("CPWtName").Editable = true;
							oMat.Columns.Item("MUseQty").Editable = true;
							oMat.Columns.Item("MUseWt").Editable = true;
							oMat.Columns.Item("ReStatus").Editable = true;
							oForm.EnableMenu("1284", true); // 취소
							oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else
						{
							oForm.Items.Item("BPLId").Enabled = false;
							oForm.Items.Item("CardCode").Enabled = false;
							oForm.Items.Item("DocDate").Enabled = false;
							oForm.Items.Item("OKYNC").Enabled = false;
							oForm.Items.Item("DocEntry").Enabled = false;
							oMat.Columns.Item("OtDocLin").Editable = false;
							oMat.Columns.Item("OutQty").Editable = false;
							oMat.Columns.Item("OutWt").Editable = false;
							oMat.Columns.Item("NQty").Editable = false;
							oMat.Columns.Item("NWeight").Editable = false;
							oMat.Columns.Item("CPWt").Editable = false;
							oMat.Columns.Item("CPWtName").Editable = false;
							oMat.Columns.Item("MUseQty").Editable = false;
							oMat.Columns.Item("MUseWt").Editable = false;
							oMat.Columns.Item("ReStatus").Editable = false;
							oForm.EnableMenu("1284", false); // 취소
						}
					}
					else
					{
						if (oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() != "Y")
						{
							oForm.Items.Item("BPLId").Enabled = false;
							oForm.Items.Item("CardCode").Enabled = false;
							oForm.Items.Item("DocDate").Enabled = true;

							if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
							{
								oForm.Items.Item("OKYNC").Enabled = false;
							}
							else
							{
								if (oDS_PS_MM152H.GetValue("CanCeled", 0).ToString().Trim() == "Y")
								{
									oForm.Items.Item("OKYNC").Enabled = false;
								}
								else
								{
									oForm.Items.Item("OKYNC").Enabled = true;
								}
							}

							oForm.Items.Item("DocEntry").Enabled = false;
							oMat.Columns.Item("OtDocLin").Editable = true;
							oMat.Columns.Item("OutQty").Editable = true;
							oMat.Columns.Item("OutWt").Editable = true;
							oMat.Columns.Item("NQty").Editable = true;
							oMat.Columns.Item("NWeight").Editable = true;
							oMat.Columns.Item("CPWt").Editable = true;
							oMat.Columns.Item("CPWtName").Editable = true;
							oMat.Columns.Item("MUseQty").Editable = true;
							oMat.Columns.Item("MUseWt").Editable = true;
							oMat.Columns.Item("ReStatus").Editable = true;
							oForm.EnableMenu("1284", true); // 취소
						}
						else
						{
							oForm.Items.Item("BPLId").Enabled = false;
							oForm.Items.Item("CardCode").Enabled = false;
							oForm.Items.Item("DocDate").Enabled = false;

							if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
							{
								oForm.Items.Item("OKYNC").Enabled = false;
							}
							else
							{
								if (oDS_PS_MM152H.GetValue("CanCeled", 0).ToString().Trim() == "Y")
								{
									oForm.Items.Item("OKYNC").Enabled = false;
								}
								else
								{
									oForm.Items.Item("OKYNC").Enabled = true;
								}
							}

							oForm.Items.Item("DocEntry").Enabled = false;
							oMat.Columns.Item("OtDocLin").Editable = false;
							oMat.Columns.Item("OutQty").Editable = false;
							oMat.Columns.Item("OutWt").Editable = false;
							oMat.Columns.Item("NQty").Editable = false;
							oMat.Columns.Item("NWeight").Editable = false;
							oMat.Columns.Item("CPWt").Editable = false;
							oMat.Columns.Item("CPWtName").Editable = false;
							oMat.Columns.Item("MUseQty").Editable = false;
							oMat.Columns.Item("MUseWt").Editable = false;
							oMat.Columns.Item("ReStatus").Editable = false;
							oForm.EnableMenu("1284", false); // 취소
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
		/// PS_MM152_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_MM152_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int sRow;
			double MDUseQty;
			double MDUseUnWt;
			double MDUseWt;
			double MUseWt;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				sRow = oRow;

				switch (oUID)
				{
					case "CardCode":
						sQry = "Select CardName From OCRD Where CardCode = '" + oDS_PS_MM152H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oDS_PS_MM152H.SetValue("U_CardName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
						break;
					case "Mat01":
						if (oCol == "OtDocLin")
						{
							oMat.FlushToDataSource();
							if ((oRow == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("OtDocLin").Cells.Item(oRow).Specific.Value.ToString().Trim()))
							{
								oMat.FlushToDataSource();
								PS_MM152_Add_MatrixRow(oMat.RowCount, false);
								oMat.Columns.Item("OtDocLin").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}

							sQry = "EXEC [PS_MM152_01] '" + oDS_PS_MM152H.GetValue("U_BPLId", 0).ToString().Trim() + "', '" + oDS_PS_MM152H.GetValue("U_CardCode", 0).ToString().Trim() + "', ";
							sQry += "'" + oDS_PS_MM152L.GetValue("U_OtDocLin", oRow - 1).ToString().Trim() + "' , '2'";
							oRecordSet.DoQuery(sQry);

							oMat.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim();
							oMat.Columns.Item("ItemName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim();
							oMat.Columns.Item("Size").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_Size").Value.ToString().Trim();
							oMat.Columns.Item("Mark").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_Mark").Value.ToString().Trim();
							oMat.Columns.Item("OutItmCd").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_OutItmCd").Value.ToString().Trim();
							oMat.Columns.Item("OutItmNm").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_OutItmNm").Value.ToString().Trim();
							oMat.Columns.Item("OutGbn").Cells.Item(oRow).Specific.Select(oRecordSet.Fields.Item("U_OutGbn").Value.ToString().Trim());
							oMat.Columns.Item("InQty").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_InQty").Value.ToString().Trim();
							oMat.Columns.Item("InWt").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_InWt").Value.ToString().Trim();
							oMat.Columns.Item("MTUseQty").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_MTUseQty").Value.ToString().Trim();
							oMat.Columns.Item("MTUseWt").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_MTUseWt").Value.ToString().Trim();
							oMat.Columns.Item("MDUseQty").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_MDUseQty").Value.ToString().Trim();
							oMat.Columns.Item("MDUseWt").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_MDUseWt").Value.ToString().Trim();
							oMat.Columns.Item("OutDoc").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_OutDoc").Value.ToString().Trim();
							oMat.Columns.Item("OutLine").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_OutLine").Value.ToString().Trim();
							oMat.Columns.Item("JakNum").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_OrdNum").Value.ToString().Trim();
							oMat.Columns.Item("ReStatus").Cells.Item(oRow).Specific.Select("N");
							oMat.Columns.Item("CpCode").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_CpCode").Value.ToString().Trim();
							oMat.Columns.Item("CpName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_CpName").Value.ToString().Trim();
							oMat.Columns.Item("TCpCode").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_TCpCode").Value.ToString().Trim();
							oMat.Columns.Item("TCpName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_TCpName").Value.ToString().Trim();
						}
						else if (oCol == "MUseQty")
						{
							MDUseQty = Convert.ToDouble(oMat.Columns.Item("MDUseQty").Cells.Item(oRow).Specific.Value.ToString().Trim());
							MDUseWt = Convert.ToDouble(oMat.Columns.Item("MDUseWt").Cells.Item(oRow).Specific.Value.ToString().Trim());

							if (MDUseQty != 0)
							{
								MDUseUnWt = MDUseWt / MDUseQty;
							}
							else
							{
								MDUseUnWt = 0;
							}

							MUseWt = MDUseUnWt * Convert.ToDouble(oMat.Columns.Item("MUseQty").Cells.Item(oRow).Specific.Value.ToString().Trim());

							if (MUseWt > MDUseWt)
							{
								oMat.Columns.Item("MUseWt").Cells.Item(oRow).Specific.Value = Convert.ToString(MDUseWt);
							}
							else
							{
								oMat.Columns.Item("MUseWt").Cells.Item(oRow).Specific.Value = Convert.ToString(MUseWt);
							}

							oMat.Columns.Item("MUseQty").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						break;
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
		/// PS_MM152_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM152_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_MM152H.GetValue("U_BPLId", 0).ToString().Trim()))
				{
					errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_MM152H.GetValue("U_CardCode", 0).ToString().Trim()))
				{
					errMessage = "외주거래처는 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_MM152H.GetValue("U_DocDate", 0).ToString().Trim()))
				{
					errMessage = "전기일자는 필수입력사항입니다. 확인하세요.";
					throw new Exception();
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
		/// PS_MM152_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM152_MatrixSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;
			int i;
			double MUseQty;
			double MUseWt;
			double MDUseWt;
			double MDUseQty;
			string OutDoc;
			string OutLine;
			string ItemCode;
			double MM130Wt;
			double MM152Wt;
			double MM132Wt;
			double OutWt;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}
				else if (oMat.VisualRowCount == 1)
				{
					if (string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_OtDocLin", 0).ToString().Trim()))
					{
						errMessage = "첫라인에 반출문서-행 번호가 없습니다. 확인하세요.";
						throw new Exception();
					}
				}

				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_OtDocLin", i).ToString().Trim()))
					{
						errMessage = "반출문서 - 행 필수사항입니다.확인하세요.";
						throw new Exception();
					}
					if (Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_ScrapWt", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MUseWt", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_Loss", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_Sample", i).ToString().Trim()) == 0)
					{
						errMessage = "" + i + 1 + "번 라인의 출고중량은 0보다 커야만 합니다. 확인하세요.";
						throw new Exception();
					}
				}

				MUseQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MUseQty", i).ToString().Trim());
				MUseWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MUseWt", i).ToString().Trim());
				MDUseQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MDUseQty", i).ToString().Trim());
				MDUseWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MDUseWt", i).ToString().Trim());

				if (MUseWt > MDUseWt)
				{
					errMessage = "" + i + 1 + "번 라인의 원재료사용중량은 원재료잔량중량보다 클수 없습니다. 확인하세요.";
					throw new Exception();
				}
				//출고중량 체크
				OutDoc = oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim();
				OutLine = oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim();
				ItemCode = oDS_PS_MM152L.GetValue("U_ItemCode", i).ToString().Trim();

				sQry = "Select U_ItmBsort from OITM Where ItemCode = '" + ItemCode + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "101")
				{
					//휘팅일경우 출고중량과 반납중량을 check
					//반출중량
					sQry = "Select b.U_OutWt from [@PS_MM130H] a, [@PS_MM130L] b Where a.DocEntry = b.DocEntry and a.Canceled = 'N' and a.U_OKYNC <> 'C' and a.U_OutDoc = '" + OutDoc + "' and b.U_LineNum = '" + OutLine + "'";
					oRecordSet.DoQuery(sQry);

					MM130Wt = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim());
					//외주입고중량
					sQry = "Select Isnull(Sum(b.U_OutWt + b.U_ScrapWt),0) from [@PS_MM152H] a, [@PS_MM152L] b Where a.DocEntry = b.DocEntry and a.Canceled = 'N' and a.U_OKYNC <> 'C' And b.U_OutDoc = '" + OutDoc + "' and b.U_OutLine = '" + OutLine + "'";
					oRecordSet.DoQuery(sQry);
					// 기입고량 + 현재입력량(스크랩포함)
					MM152Wt = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim());

					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
					{
						//현재입력되어있는 출고량
						sQry = "Select Isnull(b.U_OutWt,0) from [@PS_MM152H] a, [@PS_MM152L] b ";
						sQry = sQry + "Where a.DocEntry = b.DocEntry and a.Canceled = 'N' and a.U_OKYNC <> 'C' And b.U_OutDoc = '" + OutDoc + "' and b.U_OutLine = '" + OutLine + "' ";
						sQry = sQry + "And a.DocEntry = '" + oDS_PS_MM152L.GetValue("DocEntry", i).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						OutWt = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim());
					}
					else
					{
						OutWt = 0;
					}

					MM152Wt = MM152Wt - OutWt + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim());
					//반품량
					sQry = "Select Isnull(Sum(b.U_ReWt),0) from [@PS_MM132H] a, [@PS_MM132L] b Where a.DocEntry = b.DocEntry and a.Canceled = 'N' and a.U_OKYNC <> 'C' And b.U_OutDoc = '" + OutDoc + "' and b.U_OutLine = '" + OutLine + "'";
					oRecordSet.DoQuery(sQry);

					MM132Wt = Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim());

					if (MM130Wt < (MM152Wt + MM132Wt))
					{
						errMessage = "반출중량보다 입고중량이 큽니다. 확인하세요.";
						throw new Exception();
					}
				}

				oMat.LoadFromDataSource();
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_MM152_Delete_EmptyRow
		/// </summary>
		private void PS_MM152_Delete_EmptyRow()
		{
			int i;

			try
			{
				oMat.FlushToDataSource();

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_OtDocLin", i).ToString().Trim()))
					{
						oDS_PS_MM152L.RemoveRecord(i);
					}
				}

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM152_Add_PS_PP040
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool PS_MM152_Add_PS_PP040(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool ReturnValue = false;
			int i;
			int j;
			int lTypeCount;
			string BPLID;
			double OutQty;
			double OutWt;
			string JakNum;
			double NQty;
			double NWeight; //불량수량, 중량
			string CpCode;
			string DocDate;
			string PP040H_DocEntry;
			string DocType;
			string OutGbn;
			string TCpCode;
			int AutoKey;
			int iTemp01;
			string sTemp01;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				PSH_Globals.oCompany.StartTransaction();

				oMat.FlushToDataSource();

				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();

				if (oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() == "C")
				{
					for (i = 0; i <= oMat.RowCount - 2; i++)
					{
						if (!string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_OtDocLin", i).ToString().Trim()) && !string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_PP040Doc", i).ToString().Trim()))
						{
							OutQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutQty", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NQty", i).ToString().Trim()); //출고수량 + 불량수량
							OutWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim()); //출고중량 + 불량중량

							sQry = "Update [@PS_PP040H] Set Status = 'C', Canceled = 'Y' Where DocEntry = '" + oDS_PS_MM152L.GetValue("U_PP040Doc", i).ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							sQry = "Update [@PS_MM130L] ";
							sQry += "Set U_InQty = IsNull(U_InQty, 0) - " + OutQty + ", U_InWt = IsNull(U_InWt, 0) - " + OutWt + " ";
							sQry += "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry ";
							sQry += "Where b.U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
							sQry += "And a.U_LineNum = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
							oRecordSet.DoQuery(sQry);

							sQry = "Update [@PS_MM152L] Set U_ReStatus = 'N' From [@PS_MM152H] a Where [@PS_MM152L].DocEntry = a.DocEntry ";
							sQry += "And a.Canceled = 'N' And a.U_OKYNC <> 'C' And [@PS_MM152L].DocEntry <> '" + oDS_PS_MM152L.GetValue("DocEntry", i).ToString().Trim() + "' ";
							sQry += "And [@PS_MM152L].U_OtDocLin = '" + oDS_PS_MM152L.GetValue("U_OtDocLin", i).ToString().Trim() + "' ";
							sQry += "And [@PS_MM152L].U_ReStatus = 'Y' ";
							oRecordSet.DoQuery(sQry);

							oDS_PS_MM152L.SetValue("U_PP040Doc", i, "");
						}
					}
				}
				else if (oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
				{
					for (i = 0; i <= oMat.RowCount - 2; i++)
					{
						if (!string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_OtDocLin", i).ToString().Trim()) && string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_PP040Doc", i).ToString().Trim()))
						{
							DocDate = oDS_PS_MM152H.GetValue("U_DocDate", 0).ToString().Trim();
							JakNum = oDS_PS_MM152L.GetValue("U_JakNum", i).ToString().Trim();

							OutQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutQty", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_Sample", i).ToString().Trim()); //외주수량 + 시료수량
							OutWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_Sample", i).ToString().Trim()); //외주중량 + 시료수량

							NQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NQty", i).ToString().Trim());
							NWeight = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim());

							OutGbn = oDS_PS_MM152L.GetValue("U_OutGbn", i).ToString().Trim();
							CpCode = oDS_PS_MM152L.GetValue("U_CpCode", i).ToString().Trim();
							TCpCode = oDS_PS_MM152L.GetValue("U_TCpCode", i).ToString().Trim();

							sQry = "EXEC [PS_MM152_02] '" + JakNum + "','" + OutGbn + "', '10', '" + CpCode + "','" + TCpCode + "'";
							oRecordSet.DoQuery(sQry);

							if (oRecordSet.RecordCount == 0)
							{
								errMessage = "작업지시가 없습니다. 확인하세요.";
								throw new Exception();
							}

							lTypeCount = 0;
							while (!oRecordSet.EoF)
							{
								Array.Resize(ref PS_PP040_Renamed, lTypeCount + 1);
								PS_PP040_Renamed[lTypeCount].PP030HNo = oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim();
								PS_PP040_Renamed[lTypeCount].PP030MNo = oRecordSet.Fields.Item("LineId").Value.ToString().Trim();
								PS_PP040_Renamed[lTypeCount].OrdNum = JakNum;
								PS_PP040_Renamed[lTypeCount].Sequence = oRecordSet.Fields.Item("U_Sequence").Value.ToString().Trim();
								PS_PP040_Renamed[lTypeCount].OrdGbn = oRecordSet.Fields.Item("U_OrdGbn").Value.ToString().Trim();
								PS_PP040_Renamed[lTypeCount].ItemCode = oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim();
								PS_PP040_Renamed[lTypeCount].ItemName = oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim();
								PS_PP040_Renamed[lTypeCount].CpCode = oRecordSet.Fields.Item("U_CpCode").Value.ToString().Trim();
								PS_PP040_Renamed[lTypeCount].CpName = oRecordSet.Fields.Item("U_CpName").Value.ToString().Trim();
								PS_PP040_Renamed[lTypeCount].Chk = "N";
								lTypeCount += 1;
								oRecordSet.MoveNext();
							}
							// DocEntry
							sQry = "Select AutoKey From [ONNM] Where ObjectCode = 'PS_PP040'";
							oRecordSet.DoQuery(sQry);

							PP040H_DocEntry = oRecordSet.Fields.Item("AutoKey").Value.ToString().Trim();
							AutoKey = Convert.ToInt32(PP040H_DocEntry) + 1;

							Array.Resize(ref PS_PP040DocEntry, i + 1);
							PS_PP040DocEntry[i].PP040DocEntry = PP040H_DocEntry;

							//문서타입 하드코딩
							if (PS_PP040_Renamed[0].OrdGbn.ToString().Trim() == "104" || PS_PP040_Renamed[0].OrdGbn.ToString().Trim() == "107")
							{
								DocType = "20";
							}
							else
							{
								DocType = "10";
							}
							// Insert PS_PP040H
							sQry = " INSERT INTO [@PS_PP040H]";
							sQry += " (";
							sQry += " DocEntry,";
							sQry += " DocNum,";
							sQry += " Period,";
							sQry += " Series,";
							sQry += " Object,";
							sQry += " UserSign,";
							sQry += " CreateDate,";
							sQry += " CreateTime,";
							sQry += " DataSource,";
							sQry += " U_OrdType,";
							sQry += " U_OrdGbn,";
							sQry += " U_BPLId,";
							sQry += " U_ItemCode,";
							sQry += " U_ItemName,";
							sQry += " U_OrdNum,";
							sQry += " U_OrdSub1,";
							sQry += " U_OrdSub2,";
							sQry += " U_PP030HNo,";
							sQry += " U_DocType,";
							sQry += " Canceled,";
							sQry += " U_DocDate";
							sQry += " ) ";
							sQry += "VALUES(";
							sQry += "'" + PP040H_DocEntry + "',";
							sQry += "'" + PP040H_DocEntry + "',";
							sQry += "'22',";
							sQry += "'-1',";
							sQry += "'PS_PP040',";
							sQry += "'1',";
							sQry += "'" + DateTime.Now.ToString("yyyyMMdd") + "',";
							sQry += "'1000',";
							sQry += "'I',";
							sQry += "'30',";
							sQry += "'" + PS_PP040_Renamed[0].OrdGbn + "',";
							sQry += "'" + BPLID + "',";
							sQry += "'" + PS_PP040_Renamed[0].ItemCode + "',";
							sQry += "'" + PS_PP040_Renamed[0].ItemName + "',";
							sQry += "'" + JakNum + "',";
							sQry += "'00',";
							sQry += "'000',";
							sQry += "'" + PS_PP040_Renamed[0].PP030HNo + "',";
							sQry += "'" + DocType + "',";
							sQry += "'N',";
							sQry += "'" + DocDate + "'";
							sQry += ")";
							oRecordSet.DoQuery(sQry);

							// Insert PS_PP040M
							sQry = "INSERT INTO [@PS_PP040M]";
							sQry += " (";
							sQry += " DocEntry,";
							sQry += " LineId,";
							sQry += " VisOrder,";
							sQry += " Object,";
							sQry += " U_LineNum,";
							sQry += " U_LineId,";
							sQry += " U_WorkCode,";
							sQry += " U_WorkName";
							sQry += " ) ";
							sQry += "VALUES(";
							sQry += "'" + PP040H_DocEntry + "',";
							sQry += "'1',";
							sQry += "'0',";
							sQry += "'PS_PP040',";
							sQry += "'1',";
							sQry += "'1',";
							sQry += "'9999999',";
							sQry += "'조정'";
							sQry += ")";
							oRecordSet.DoQuery(sQry);

							//for (j = 0; j <= (Information.UBound(PS_PP040_Renamed)); j++)   
							// Check 해 주세요..... 황
							for (j = 0; j <= PS_PP040_Renamed.Length - 2; j++)
							{
								if (PS_PP040_Renamed[j].Chk == "N")
								{
									iTemp01 = j + 1;
									sTemp01 = PS_PP040_Renamed[j].PP030HNo + "-" + PS_PP040_Renamed[j].PP030MNo;
									if (iTemp01 > 1)
									{
										NQty = 0;
										NWeight = 0;
									}
									// Insert PS_PP040L
									sQry = " INSERT INTO [@PS_PP040L]";
									sQry += " (";
									sQry += " DocEntry,";
									sQry += " LineId,";
									sQry += " VisOrder,";
									sQry += " Object,";
									sQry += " U_LineNum,";
									sQry += " U_LineId,";
									sQry += " U_OrdMgNum,";
									sQry += " U_Sequence,";
									sQry += " U_CpCode,";
									sQry += " U_CpName,";
									sQry += " U_OrdGbn,";
									sQry += " U_BPLId,";
									sQry += " U_ItemCode,";
									sQry += " U_ItemName,";
									sQry += " U_OrdNum,";
									sQry += " U_OrdSub1,";
									sQry += " U_OrdSub2,";
									sQry += " U_PP030HNo,";
									sQry += " U_PP030MNo,";
									sQry += " U_PQty,";
									sQry += " U_PWeight,";
									sQry += " U_YQty,";
									sQry += " U_YWeight,";
									sQry += " U_NQty,";
									sQry += " U_NWeight,";
									sQry += " U_PSum";
									sQry += " ) ";
									sQry += "VALUES(";
									sQry += "'" + PP040H_DocEntry + "',";
									sQry += "'" + iTemp01 + "',";
									sQry += "'" + j + "',";
									sQry += "'PS_PP040',";
									sQry += "'" + iTemp01 + "',";
									sQry += "'" + iTemp01 + "',";
									sQry += "'" + sTemp01 + "',";
									sQry += "'" + PS_PP040_Renamed[j].Sequence + "',";
									sQry += "'" + PS_PP040_Renamed[j].CpCode + "',";
									sQry += "'" + PS_PP040_Renamed[j].CpName + "',";
									sQry += "'" + PS_PP040_Renamed[j].OrdGbn + "',";
									sQry += "'" + BPLID + "',";
									sQry += "'" + PS_PP040_Renamed[j].ItemCode + "',";
									sQry += "'" + PS_PP040_Renamed[j].ItemName + "',";
									sQry += "'" + JakNum + "',";
									sQry += "'000',";
									sQry += "'00',";
									sQry += "'" + PS_PP040_Renamed[j].PP030HNo + "',";
									sQry += "'" + PS_PP040_Renamed[j].PP030MNo + "',";
									sQry += "'" + OutQty + NQty + "',";
									sQry += "'" + OutWt + NWeight + "',";
									sQry += "'" + OutQty + "',";
									sQry += "'" + OutWt + "',";
									sQry += "'" + NQty + "',";
									sQry += "'" + NWeight + "',";
									sQry += "'" + 0 + "'";
									sQry += ")";
									oRecordSet.DoQuery(sQry);

									// Insert PS_PP040N
									sQry = " INSERT INTO [@PS_PP040N]";
									sQry += " (";
									sQry += " DocEntry,";
									sQry += " LineId,";
									sQry += " VisOrder,";
									sQry += " Object,";
									sQry += " U_LineNum,";
									sQry += " U_LineId,";
									sQry += " U_OrdMgNum,";
									sQry += " U_CpCode,";
									sQry += " U_CpName";
									sQry += " ) ";
									sQry += "VALUES(";
									sQry += "'" + PP040H_DocEntry + "',";
									sQry += "'" + iTemp01 + "',";
									sQry += "'" + j + "',";
									sQry += "'PS_PP040',";
									sQry += "'" + iTemp01 + "',";
									sQry += "'" + iTemp01 + "',";
									sQry += "'" + sTemp01 + "',";
									sQry += "'" + PS_PP040_Renamed[j].CpCode + "',";
									sQry += "'" + PS_PP040_Renamed[j].CpName + "'";
									sQry += ")";
									oRecordSet.DoQuery(sQry);

									PS_PP040_Renamed[j].Chk = "Y";
								}
							}

							sQry = "Update [ONNM] Set AutoKey = '" + AutoKey + "' Where ObjectCode = 'PS_PP040'";
							oRecordSet.DoQuery(sQry);

							NQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NQty", i).ToString().Trim());
							NWeight = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim());

							sQry = " Update [@PS_MM130L] ";
							sQry += "Set U_InQty = IsNull(U_InQty, 0) + " + OutQty + NQty + ", U_InWt = IsNull(U_InWt, 0) + " + OutWt + NWeight + " ";
							sQry += "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry ";
							sQry += "Where b.U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
							sQry += "And a.U_LineNum = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
							oRecordSet.DoQuery(sQry);

							oDS_PS_MM152L.SetValue("U_PP040Doc", i, PS_PP040DocEntry[i].PP040DocEntry);
						}
					}
				}

				oMat.LoadFromDataSource();

				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("외주업체 출고등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_MM152_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_MM152_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string DocNum;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocNum = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

				WinTitle = "거래명세표 [PS_MM152_01]";
				ReportName = "PS_MM152_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocNum", DocNum));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
					Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//	Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
			int i;
			int DocEntryNext = 0;
			double ScrapWt;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_MM152_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_MM152_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							//스크랩 중량 계산
							oMat.FlushToDataSource();

							for (i = 0; i <= oMat.RowCount - 2; i++)
							{
								if (oDS_PS_MM152L.GetValue("U_ReStatus", i).ToString().Trim() == "Y")
								{
									sQry = "Select U_ItmBSort From [OITM] Where ItemCode = '" + oDS_PS_MM152L.GetValue("U_ItemCode", i).ToString().Trim() + "'";
									oRecordSet.DoQuery(sQry);

									//휘팅일때만
									if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "101")
									{
										sQry = " Select a.U_OutDoc, b.U_LineNum, b.U_OutWt - Isnull(c.U_OutWt,0) - Isnull(b.U_ReWt,0) ";
										sQry += "from [@PS_MM130H] a Inner Join [@PS_MM130L] b On a.DocEntry = b.DocEntry ";
										sQry += "Inner Join (Select t1.U_OutDoc, ";
										sQry += "t1.U_OutLine, ";
										sQry += "Isnull(Sum(t1.U_OutWt),0) As U_OutWt ";
										sQry += "From [@PS_MM152H] t Inner Join [@PS_MM152L] t1 On t.DocEntry = t1.DocEntry ";
										sQry += "Where t.Canceled = 'N' ";
										sQry += "and t.U_OKYNC <> 'C' ";
										sQry += "Group by t1.U_OutDoc, ";
										sQry += "t1.U_OutLine ) c On a.U_OutDoc = c.U_OutDoc And b.U_LineNum = c.U_OutLine ";
										sQry += "where a.U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
										sQry += "and b.U_LineNum = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
										oRecordSet.DoQuery(sQry);

										//스크랩 중량 = 미입고 잔량 - 현재 입고 중량
										ScrapWt = Convert.ToDouble(oRecordSet.Fields.Item(2).Value.ToString().Trim()) - Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim());

										if (oRecordSet.RecordCount == 0)
										{
											sQry = " Select U_OutDoc, U_LineNum, IsNull(U_OutWt, 0) - IsNull(U_ReWt, 0) ";
											sQry += "From [@PS_MM130L] ";
											sQry += "Where U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
											sQry += "And U_LineNum = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
											oRecordSet.DoQuery(sQry);

											ScrapWt = Convert.ToDouble(oRecordSet.Fields.Item(2).Value.ToString().Trim()) - Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim());
										}

										oDS_PS_MM152L.SetValue("U_ScrapWt", i, Convert.ToString(ScrapWt));
									}
								}
							}

							oMat.LoadFromDataSource();
							PS_MM152_Delete_EmptyRow();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PSH_Globals.oCompany.InTransaction == true)
							{
								PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
							}
							PSH_Globals.oCompany.StartTransaction();

							if (PS_MM152_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() != "C")
							{
								if (PS_MM152_MatrixSpaceLineDel() == false)
								{
									BubbleEvent = false;
									return;
								}

								//스크랩 중량 계산
								oMat.FlushToDataSource();

								for (i = 0; i <= oMat.RowCount - 2; i++)
								{
									if (oDS_PS_MM152L.GetValue("U_ReStatus", i).ToString().Trim() == "Y")
									{
										sQry = "Select U_ItmBSort From [OITM] Where ItemCode = '" + oDS_PS_MM152L.GetValue("U_ItemCode", i).ToString().Trim() + "'";
										oRecordSet.DoQuery(sQry);
										//휘팅일때만
										if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "101")
										{
											sQry = " Select  b.U_OutDoc, b.U_OutLine, IsNull(d.U_OutWt, 0) - Sum(IsNull(b.U_OutWt, 0)) - IsNull(d.U_ReWt, 0) ";
											sQry += "From [@PS_MM152H] a, ";
											sQry += "[@PS_MM152L] b, ";
											sQry += "[@PS_MM130H] c, ";
											sQry += "[@PS_MM130L] d ";
											sQry += "where a.DocEntry = b.DocEntry ";
											sQry += "and c.DocEntry = d.DocEntry ";
											sQry += "and b.U_OutDoc = c.U_OutDoc ";
											sQry += "and b.U_OutLine = d.U_LineNum ";
											sQry += "and b.U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
											sQry += "and b.U_OutLine = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
											sQry += "and a.Canceled = 'N' ";
											sQry += "and a.U_OKYNC <> 'C' ";
											sQry += "Group by b.U_OutDoc, b.U_OutLine, IsNull(d.U_OutWt, 0), IsNull(d.U_ReWt, 0) ";
											oRecordSet.DoQuery(sQry);

											//스크랩 중량 = 미입고 잔량 - 현재 입고 중량
											ScrapWt = Convert.ToDouble(oRecordSet.Fields.Item(2).Value.ToString().Trim());

											//수정전 입고중량
											sQry = " Select U_OutWt From [@PS_MM152L] Where DocEntry = '" + oDS_PS_MM152L.GetValue("DocEntry", i).ToString().Trim() + "' ";
											sQry += "And U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
											sQry += "And U_OutLine = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
											oRecordSet.DoQuery(sQry);

											ScrapWt = ScrapWt - Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim());

											if (oRecordSet.RecordCount == 0)
											{
												sQry = " Select U_OutDoc, U_LineNum, IsNull(U_OutWt, 0) - IsNull(U_ReWt, 0) ";
												sQry += "From [@PS_MM130L] ";
												sQry += "Where U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
												sQry += "And U_LineNum = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
												oRecordSet.DoQuery(sQry);

												ScrapWt = Convert.ToDouble(oRecordSet.Fields.Item(2).Value.ToString().Trim()) - Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i));
											}

											oDS_PS_MM152L.SetValue("U_ScrapWt", i, Convert.ToString(ScrapWt));
										}
									}
								}

								oMat.LoadFromDataSource();
							}
							if (PS_MM152_Add_PS_PP040(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}
							else
							{
								if (PSH_Globals.oCompany.InTransaction == true)
								{
									PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
								}
							}

							PS_MM152_Delete_EmptyRow();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim()))
							{
							}
							else
							{
								DocEntryNext = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());

								sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" + DocEntryNext + "'";
								oRecordSet.DoQuery(sQry);

								if (PSH_Globals.oCompany.UserName != oRecordSet.Fields.Item(0).Value.ToString().Trim())
								{
									if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
									{
										errMessage = "해당 문서 번호는 없는 다른 거래처의 문서입니다. 확실한 문서번호를 입력하세요.";
										throw new Exception();
									}
								}
							}
						}
					}
					else if (pVal.ItemUID == "Btn01")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_MM152_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE & pVal.Action_Success == true)
						{
							PS_MM152_Initialization();
							PS_MM152_FormClear();
							PS_MM152_FormItemEnabled();
							PS_MM152_Add_MatrixRow(0, true);
						}
					}
				}
			}
			catch (Exception ex)
			{
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "CardCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "OtDocLin")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
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
							PS_MM152_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "OtDocLin" || pVal.ColUID == "MUseQty")
							{
								PS_MM152_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							else if (pVal.ColUID == "OutQty")
							{
								sQry = "Select U_ItmBSort From OITM Where ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "102")
								{
									oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								}
							}
							else if (pVal.ColUID == "NQty")
							{
								sQry = "Select U_ItmBSort From OITM Where ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "102")
								{
									oMat.Columns.Item("NWeight").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("NQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								}
							}

							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ReStatus")
					{
						if (oMat.Columns.Item("OutGbn").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "10")
						{
							if (oMat.Columns.Item("ReStatus").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "Y")
							{
								oMat.Columns.Item("MUseQty").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("MDUseQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								oMat.Columns.Item("MUseWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("MDUseWt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
							}
							else if (oMat.Columns.Item("ReStatus").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "N")
							{
								oMat.Columns.Item("MUseQty").Cells.Item(pVal.Row).Specific.Value = "0";
								oMat.Columns.Item("MUseWt").Cells.Item(pVal.Row).Specific.Value = "0";
							}
						}
						else if (oMat.Columns.Item("OutGbn").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "20")
						{
							oMat.Columns.Item("MUseQty").Cells.Item(pVal.Row).Specific.Value = "0";
							oMat.Columns.Item("MUseWt").Cells.Item(pVal.Row).Specific.Value = "0";
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
					PS_MM152_Add_MatrixRow(oMat.RowCount, false);
					PS_MM152_FormItemEnabled();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM152H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM152L);
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
			int DocEntry;
			int DocEntryMax;
			int DocEntryNext;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								BubbleEvent = false;
								return;
							}
							break;
						case "1286": //닫기
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
                            if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim()))
                            {
                                DocEntry = 0;
                            }
                            else
                            {
                                DocEntry = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
                            }

                            sQry = "Select Max(DocEntry) From [@PS_MM152H]";
                            oRecordSet.DoQuery(sQry);

                            DocEntryMax = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim());

                            if (pVal.MenuUID == "1288") //다음
							{
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                {
                                    PSH_Globals.SBO_Application.ActivateMenuItem("1290");
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                                {
                                    if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim()))
                                    {
                                        PSH_Globals.SBO_Application.ActivateMenuItem("1290");
                                        BubbleEvent = false;
                                        return;
                                    }
                                }

                            One_More_Check_1288:
                                DocEntryNext = DocEntry + 1;
                                if (DocEntryNext > DocEntryMax)
                                {
                                    DocEntry = 0;
                                    goto One_More_Check_1288;
                                }

                                sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" + DocEntryNext + "'";
                                oRecordSet.DoQuery(sQry);

								if (PSH_Globals.oCompany.UserName != oRecordSet.Fields.Item(0).Value.ToString().Trim())
								{
									if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
									{
										DocEntry = DocEntry + 1;
										goto One_More_Check_1288;
									}
									else
									{
										oCheck = "False";
										return;
									}
								}
								else
								{
									oCheck = "True";
									oDocEntryNext = DocEntryNext;
								}
                            }

							else if (pVal.MenuUID == "1289")  //이전
							{
								if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("1291");
									BubbleEvent = false;
									return;
								}
								else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
								{
									if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim()))
									{
										PSH_Globals.SBO_Application.ActivateMenuItem("1291");
										BubbleEvent = false;
										return;
									}
								}

							One_More_Check_1289:
								DocEntryNext = DocEntry - 1;
								if (DocEntryNext < 1)
								{
									DocEntry = DocEntryMax + 1;
									goto One_More_Check_1289;
								}

								sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" + DocEntryNext + "'";
								oRecordSet.DoQuery(sQry);

								if (PSH_Globals.oCompany.UserName != oRecordSet.Fields.Item(0).Value.ToString().Trim())
								{
									if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
									{
										DocEntry = DocEntry - 1;
										goto One_More_Check_1289;
									}
									else
									{
										oCheck = "False";
										return;
									}
								}
								else
								{
									oCheck = "True";
									oDocEntryNext = DocEntryNext;
								}
							}

							else if (pVal.MenuUID == "1290") //맨첨
							{
								DocEntryNext = 0;

							One_More_Check_1290:
								DocEntryNext += 1;

								sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" + DocEntryNext + "'";
								oRecordSet.DoQuery(sQry);

								if (PSH_Globals.oCompany.UserName != oRecordSet.Fields.Item(0).Value.ToString().Trim())
								{
									if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
									{
										goto One_More_Check_1290;
									}
									else
									{
										oCheck = "False";
										return;
									}
								}
								else
								{
									oCheck = "True";
									oDocEntryNext = DocEntryNext;
								}
							}

							else if (pVal.MenuUID == "1291") //맨뒤
							{
								DocEntryNext = DocEntryMax + 1;

							One_More_Check_1291:
								DocEntryNext -= 1;

								sQry = "Select U_CardCode From [@PS_MM152H] Where DocEntry = '" + DocEntryNext + "'";
								oRecordSet.DoQuery(sQry);

								if (PSH_Globals.oCompany.UserName != oRecordSet.Fields.Item(0).Value.ToString().Trim())
								{
									if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
									{
										goto One_More_Check_1291;
									}
									else
									{
										oCheck = "False";
										return;
									}
								}
								else
								{
									oCheck = "True";
									oDocEntryNext = DocEntryNext;
								}
							}
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
							oDS_PS_MM152H.SetValue("U_BPLId", 0, "1");
							PS_MM152_FormItemEnabled();
							if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
							{
								oForm.Items.Item("CardCode").Specific.Value = PSH_Globals.oCompany.UserName;

								sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							}
							oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oForm.Items.Item("CardCode").Enabled = false;
							break;
						case "1282": //추가
							PS_MM152_Initialization();
							PS_MM152_FormClear();
							PS_MM152_FormItemEnabled();
							PS_MM152_Add_MatrixRow(0, true);
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							PS_MM152_FormClear();
							oDS_PS_MM152H.SetValue("Status", 0, "O");
							oDS_PS_MM152H.SetValue("Canceled", 0, "N");
							PS_MM152_FormItemEnabled();
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							if (oCheck == "True")
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("1281");
								oForm.Items.Item("DocEntry").Specific.Value = oDocEntryNext;
								oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								oCheck = "False";
								oDocEntryNext = 0;
							}
							PS_MM152_FormItemEnabled();
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
