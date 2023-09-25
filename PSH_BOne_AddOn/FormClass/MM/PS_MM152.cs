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
				oForm.EnableMenu("1284", false);  // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM152_ComboBox_Setting
		/// </summary>
		private void PS_MM152_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oMat.Columns.Item("OutGbn").ValidValues.Add("10", "원재료");
				oMat.Columns.Item("OutGbn").ValidValues.Add("20", "재공");

				oMat.Columns.Item("ReStatus").ValidValues.Add("Y", "완료");
				oMat.Columns.Item("ReStatus").ValidValues.Add("N", "미완료");

				//Action(Matrix)
				sQry = "  SELECT      U_Minor, ";
				sQry += "             U_CdName ";
				sQry += " FROM        [@PS_SY001L] ";
				sQry += " WHERE       Code = 'A009'";
				sQry += "             AND ISNULL(U_UseYN, 'Y') = 'Y'";
				sQry += "             AND U_Minor <> 'D'";
				sQry += " ORDER BY    U_Seq";

				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("Action"), sQry, "", "");
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
		/// PS_MM152_Initialization
		/// </summary>
		private void PS_MM152_Initialization()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet1 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				//oDS_PS_MM152H.SetValue("U_BPLId", 0, "1");
				oDS_PS_MM152H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
				oDS_PS_MM152H.SetValue("U_OKYNC", 0, "N");

				sQry = "SELECT COUNT(*) FROM [@PS_SY005H] A INNER JOIN [@PS_SY005L] B ON A.Code = B.Code where A.Code ='M152' AND B.U_UseYN ='Y' AND B.U_AppUser = '" + PSH_Globals.oCompany.UserName + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "1")
				{
					oDS_PS_MM152H.SetValue("U_CardCode", 0, PSH_Globals.oCompany.UserName);

					sQry = "Select CardName From OCRD Where CardCode = '" + oDS_PS_MM152H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
					oRecordSet1.DoQuery(sQry);

					oDS_PS_MM152H.SetValue("U_CardName", 0, oRecordSet1.Fields.Item(0).Value.ToString().Trim());
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet1);
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM152_FormItemEnabled
		/// </summary>
		private void PS_MM152_FormItemEnabled()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			try
			{
				if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "2")
				{
					oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oMat.Columns.Item("HeatNo").Visible = true;
					oMat.Columns.Item("DNQty").Visible = true;
					oMat.Columns.Item("DNCode").Visible = true;
					oMat.Columns.Item("AttPath").Visible = true;
					oMat.Columns.Item("Action").Visible = true;
					oMat.Columns.Item("QCOKDate").Visible = true;
					oMat.Columns.Item("MSTCOD").Visible = true;
					oMat.Columns.Item("MSTNAM").Visible = true;



					oMat.Columns.Item("ScrapWt").Visible = false;
					oMat.Columns.Item("CPWt").Visible = false;
					oMat.Columns.Item("CPWtName").Visible = false;
					oMat.Columns.Item("Sample").Visible = false;
					oMat.Columns.Item("Loss").Visible = false;

					PS_MM152_Add_MatrixRow(0, true);
				}
				else
				{
					oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oMat.Columns.Item("HeatNo").Visible = false;
					oMat.Columns.Item("DNQty").Visible = false;
					oMat.Columns.Item("DNCode").Visible = false;
					oMat.Columns.Item("AttPath").Visible = false;
					oMat.Columns.Item("Action").Visible = false;
					oMat.Columns.Item("QCOKDate").Visible = false;
					oMat.Columns.Item("MSTCOD").Visible = false;
					oMat.Columns.Item("MSTNAM").Visible = false;

					oMat.Columns.Item("ScrapWt").Visible = true;
					oMat.Columns.Item("CPWt").Visible = true;
					oMat.Columns.Item("CPWtName").Visible = true;
					oMat.Columns.Item("Sample").Visible = true;
					oMat.Columns.Item("Loss").Visible = true;

					PS_MM152_Add_MatrixRow(0, true);
				}
				oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				sQry = "SELECT COUNT(*) FROM [@PS_SY005H] A INNER JOIN [@PS_SY005L] B ON A.Code = B.Code where A.Code ='M152' AND B.U_UseYN ='Y' AND B.U_AppUser = '" + PSH_Globals.oCompany.UserName + "'";
				oRecordSet.DoQuery(sQry);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					if(oRecordSet.Fields.Item(0).Value.ToString().Trim() == "1")
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
						
						oMat.Columns.Item("QCOKDate").Editable = false;
						oMat.Columns.Item("MSTCOD").Editable = false;
						oMat.Columns.Item("DNQty").Editable = false;
						oMat.Columns.Item("DNCode").Editable = false;
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
						oMat.Columns.Item("QCOKDate").Editable = true;
						oMat.Columns.Item("MSTCOD").Editable = true;
						oMat.Columns.Item("DNQty").Editable = true;
						oMat.Columns.Item("DNCode").Editable = true;
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

							if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "1")
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

							if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "1")
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			string sQry;
			double MDUseQty;
			double MDUseUnWt;
			double MDUseWt;
			double MUseWt;
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
							oMat.Columns.Item("HeatNo").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("U_HeatNo").Value.ToString().Trim();
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_MM152H.GetValue("U_BPLId", 0).ToString().Trim()))
				{
					errMessage = "사업장은 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_MM152H.GetValue("U_DocDate", 0).ToString().Trim()))
				{
					errMessage = "전기일자는 필수입력사항입니다. 확인하세요.";
					throw new Exception();
				}
				// 마감일자 Check
				else if (dataHelpClass.Check_Finish_Status(oDS_PS_MM152H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM152H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
				{
					errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_MM152H.GetValue("U_CardCode", 0).ToString().Trim()))
				{
					errMessage = "외주거래처는 필수입력사항입니다. 확인하세요.";
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
			int i;
			string sQry;
			string OutDoc;
			string OutLine;
			string ItemCode;
			string errMessage = string.Empty;
			double MOutWt;
			double MUseWt;
			double MDUseWt;
			double MM130Wt;
			double MM152Wt;
			double MM132Wt;
			double MM132Qty;
			double MM152Qty;
			double OutWt;
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
						errMessage = "'반출문서 - 행'dev 필수사항입니다.확인하세요.";
						throw new Exception();
					}
					if (Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_ScrapWt", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MUseWt", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_Loss", i).ToString().Trim()) == 0 && Convert.ToDouble(oDS_PS_MM152L.GetValue("U_Sample", i).ToString().Trim()) == 0)
					{
						errMessage = "" + (i + 1) + "번 라인의 출고중량은 0보다 커야만 합니다. 확인하세요.";
						throw new Exception();
					}
				}
				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					MOutWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim());
					MUseWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MUseWt", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim());
					MDUseWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MDUseWt", i).ToString().Trim());

					if (MUseWt > MDUseWt || MOutWt > MDUseWt)
					{
						errMessage = "" + (i + 1) + "번 라인의 원재료사용중량은 원재료잔량중량보다 클수 없습니다. 확인하세요.";
						throw new Exception();
					}
				}

				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					MM132Qty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutQty", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NQty", i).ToString().Trim());
					MM152Qty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MDUseQty", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim());
					if (MM132Qty > MM152Qty)
					{
						errMessage = "입고수량보다 반출수량이 큽니다. 확인하세요.";
						throw new Exception();
					}
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
				else if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "105")
				{
					for (i = 0; i <= oMat.VisualRowCount - 2; i++)
					{
						MOutWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutQty", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NQty", i).ToString().Trim());
						MUseWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim()) + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim());
						MDUseWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MDUseWt", i).ToString().Trim());

						if(MOutWt == Convert.ToDouble(oDS_PS_MM152L.GetValue("U_MDUseQty", i).ToString().Trim()))
                        {
							if (MUseWt != MDUseWt)
							{
								errMessage = "" + (i + 1) + "번 라인의 최종납품시 (납품중량 + 불량중량 = 원재료잔량중량)이 되어야합니다 . 확인하세요.";
								throw new Exception();
							}
						}
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
		/// PS_MM152_etBaseForm
		/// </summary>
		private void PS_MM152_SaveAttach(int pRow)
		{
			string sFileFullPath;
			string sFilePath;
			string sFileName;
			string SaveFolders;
			string sourceFile;
			string targetFile;
			string errMessage = string.Empty;

			try
			{
				sFileFullPath = PS_MM152_OpenFileSelectDialog();//OpenFileDialog를 쓰레드로 실행

				SaveFolders = "\\\\191.1.1.220\\Attach\\PS_MM152";
				sFileName = System.IO.Path.GetFileName(sFileFullPath); //파일명
				sFilePath = System.IO.Path.GetDirectoryName(sFileFullPath); //파일명을 제외한 전체 경로

				sourceFile = System.IO.Path.Combine(sFilePath, sFileName);
				targetFile = System.IO.Path.Combine(SaveFolders, sFileName);
				oMat.FlushToDataSource();

				if (System.IO.File.Exists(targetFile) || !string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_AttPath", pRow - 1).ToString().Trim())) //서버에 기존파일이 존재하는지 체크
				{
					if (PSH_Globals.SBO_Application.MessageBox("파일이 존재합니다. 교체하시겠습니까?", 2, "Yes", "No") == 1)
					{
						System.IO.File.Delete(targetFile); //삭제
					}
					else
					{
						return;
					}
				}
				oDS_PS_MM152L.SetValue("U_AttPath", pRow - 1, SaveFolders + "\\" + sFileName); //첨부파일 경로 등록

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				System.IO.File.Copy(sourceFile, targetFile, true); //파일 복사 (여기서 오류발생)
				PSH_Globals.SBO_Application.MessageBox("업로드 되었습니다.");
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
		}

		/// <summary>
		/// OpenFileSelectDialog 호출(쓰레드를 이용하여 비동기화)
		/// OLE 호출을 수행하려면 현재 스레드를 STA(단일 스레드 아파트) 모드로 설정해야 합니다.
		/// </summary>
		[STAThread]
		private string PS_MM152_OpenFileSelectDialog()
		{
			string returnFileName = string.Empty;

			var thread = new System.Threading.Thread(() =>
			{
				System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
				openFileDialog.InitialDirectory = "C:\\";
				openFileDialog.Filter = "All files (*.*)|*.*";
				openFileDialog.FilterIndex = 1; //FilterIndex는 1부터 시작
				openFileDialog.RestoreDirectory = true;

				if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
				{
					returnFileName = openFileDialog.FileName;
				}
			});

			thread.SetApartmentState(System.Threading.ApartmentState.STA);
			thread.Start();
			thread.Join();

			return returnFileName;
		}

		/// <summary>
		/// PS_MM152_etBaseForm
		/// </summary>
		private void PS_MM152_OpenAttach(int pRow)
		{
			string AttachPath;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				AttachPath = oDS_PS_MM152L.GetValue("U_AttPath", pRow - 1).ToString().Trim();

				if (string.IsNullOrEmpty(AttachPath))
				{
					PSH_Globals.SBO_Application.MessageBox("첨부파일이 없습니다.");
				}
				else
				{
					System.Diagnostics.ProcessStartInfo process = new System.Diagnostics.ProcessStartInfo(AttachPath);
					process.UseShellExecute = true;
					process.Verb = "open";

					System.Diagnostics.Process.Start(process);
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
			int AutoKey;
			int iTemp01;
			string JakNum;
			string BPLID;
			string CpCode;
			string DocDate;
			string PP040H_DocEntry;
			string DocType;
			string OutGbn;
			string TCpCode;
			string sTemp01;
			string errMessage = string.Empty;
			string sQry;
			double OutQty;
			double DNQty;
			double OutWt;
			double NQty;
			double NWeight; //불량수량, 중량
			string sQry1;
			double SelWt;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet1 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
							OutQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutQty", i).ToString().Trim()); //출고수량 + 불량수량 
							OutWt = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim()); //출고중량 + 불량중량 

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

							if (BPLID == "1")
							{
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

								for (j = 0; j <= PS_PP040_Renamed.Length - 1; j++)
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
										sQry += OutQty + ",";
										sQry += OutWt + ",";
										sQry += OutQty + ",";
										sQry += OutWt + ",";
										sQry += NQty + ",";
										sQry += NWeight + ",";
										sQry += 0;
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

										sQry = "Update [ONNM] Set AutoKey = '" + AutoKey + "' Where ObjectCode = 'PS_PP040'";
										oRecordSet.DoQuery(sQry);

										NQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NQty", i).ToString().Trim());

										sQry = " Update [@PS_MM130L] ";
										sQry += "Set U_InQty = IsNull(U_InQty, 0) + " + OutQty + "+" + NQty + ", U_InWt = IsNull(U_InWt, 0) + " + OutWt;
										sQry += " From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry ";
										sQry += "Where b.U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
										sQry += "And a.U_LineNum = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
										oRecordSet.DoQuery(sQry);

										oDS_PS_MM152L.SetValue("U_PP040Doc", i, PS_PP040DocEntry[i].PP040DocEntry);
									}
								}
							}
							else //부산 작업지시등록
							{
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
								sQry += "'80',";
								sQry += "'" + PS_PP040_Renamed[0].OrdGbn + "',";
								sQry += "'" + BPLID + "',";
								sQry += "'10',";
								sQry += "'N',";
								sQry += "'" + DocDate + "'";
								sQry += ")";
								oRecordSet.DoQuery(sQry);

								for (j = 0; j <= PS_PP040_Renamed.Length - 1; j++)
								{
									NQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NQty", i).ToString().Trim());  //외주불량수량
									DNQty = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_DNQty", i).ToString().Trim()); //당사불량수량
									NWeight = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_NWeight", i).ToString().Trim()); //외주불량중량

									if (PS_PP040_Renamed[j].Chk == "N")
									{
										iTemp01 = j + 1;
										sTemp01 = PS_PP040_Renamed[j].PP030HNo + "-" + PS_PP040_Renamed[j].PP030MNo;
										if (iTemp01 > 1)
										{
											NQty = 0;
											NWeight = 0;
										}
										sQry1 = "SELECT U_SelWt AS U_SelWt FROM [@PS_PP030H] WHERE U_OrdNum ='" + PS_PP040_Renamed[j].ItemCode + "'";
										oRecordSet1.DoQuery(sQry1);

										SelWt = Convert.ToDouble(oRecordSet1.Fields.Item(0).Value.ToString().Trim());

										// Insert PS_PP040L
										sQry = " INSERT INTO [@PS_PP040L]";
										sQry += " (";
										sQry += " DocEntry,";
										sQry += " LineId,";
										sQry += " VisOrder,";
										sQry += " Object,";
										sQry += " U_SelWt,";
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
										sQry += " U_WorkCls,";
										sQry += " U_ScrapWt,";
										sQry += " U_WorkTime,";
										sQry += " U_PSum,";
										sQry += " U_Comment";
										sQry += " ) ";
										sQry += "VALUES(";
										sQry += "'" + PP040H_DocEntry + "',";
										sQry += "'" + iTemp01 + "',";
										sQry += "'" + j + "',";
										sQry += "'PS_PP040',";
										sQry += "'" + SelWt + "',";
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
										sQry += OutQty + NQty + ",";
										sQry += OutWt + NWeight + ",";
										sQry += OutQty - DNQty+ ",";
										sQry += OutWt - NWeight + ",";
										sQry += NQty + DNQty + ",";
										sQry += NWeight + ",";
										sQry += "'A',";
										sQry += 0 + ",";
										sQry += 0 + ",";
										sQry += 0 + ",";
										sQry += "'A'";
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
										sQry += " U_FailQty,";
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
										sQry += "'" + NQty + "',";
										sQry += "'" + PS_PP040_Renamed[j].CpCode + "',";
										sQry += "'" + PS_PP040_Renamed[j].CpName + "'";
										sQry += ")";
										oRecordSet.DoQuery(sQry);

										PS_PP040_Renamed[j].Chk = "Y";

										sQry = "Update [ONNM] Set AutoKey = '" + AutoKey + "' Where ObjectCode = 'PS_PP040'";
										oRecordSet.DoQuery(sQry);

									
										sQry = " Update [@PS_MM130L] ";
										sQry += "Set U_InQty = IsNull(U_InQty, 0) + " + OutQty + "+" + NQty + "+" + DNQty + ", U_InWt = IsNull(U_InWt, 0) + " + OutWt + "+" + NWeight;
										sQry += " From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry ";
										sQry += "Where b.U_OutDoc = '" + oDS_PS_MM152L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
										sQry += "And a.U_LineNum = '" + oDS_PS_MM152L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
										oRecordSet.DoQuery(sQry);

										oDS_PS_MM152L.SetValue("U_PP040Doc", i, PS_PP040DocEntry[i].PP040DocEntry);
									}
								}
							}
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet1);
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
		/// 입고DI
		/// </summary>
		/// <returns></returns>
		private bool PS_MM152_DI_API()
		{
			bool returnValue = false;
			string errCode = string.Empty;
			string errDIMsg = string.Empty;
			int errDICode = 0;
			int i;
			int RetVal;
			int LineNumCount;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Documents oDIObject = null;
			SAPbouiCOM.ProgressBar ProgBar01 = null;
			try
			{
				ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

				PSH_Globals.oCompany.StartTransaction();

				//현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
				if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
				{
					errCode = "2";
					throw new Exception();
				}

				LineNumCount = 0;
				oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
				if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
				{
					oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
				}
				oDIObject.UserFields.Fields.Item("Comments").Value = "가공품출고(외주업체용)_PS_MM152 문서번호:" + oForm.Items.Item("DocEntry").Specific.Value + " 자동불출취소";

				for (i = 0; i < oMat.VisualRowCount - 1; i++)
				{
					oDIObject.Lines.Add();
					oDIObject.Lines.SetCurrentLine(LineNumCount);
					oDIObject.Lines.ItemCode = oDS_PS_MM152L.GetValue("U_OutItmCd", i).ToString().Trim();
					oDIObject.Lines.WarehouseCode = "802";
					oDIObject.Lines.Quantity = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim());
					oDIObject.Lines.Price = Convert.ToDouble(0);
					oDIObject.Lines.LineTotal = Convert.ToDouble(0);
					oDIObject.Lines.UserFields.Fields.Item("PriceBefDi").Value = Convert.ToDouble(0);
					oDIObject.Lines.UserFields.Fields.Item("U_OrdNum").Value = oDS_PS_MM152L.GetValue("U_ItemCode", i).ToString().Trim();
					oDIObject.Lines.UserFields.Fields.Item("U_sSize").Value = oDS_PS_MM152L.GetValue("U_HeatNo", i).ToString().Trim();
					LineNumCount += 1;
				}

					RetVal = oDIObject.Add();

				if (RetVal == 0)
				{
					//PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

					//for (i = 1; i <= oMat.VisualRowCount; i++)
					//{
					//	oMat.Columns.Item("InDoc").Cells.Item(i).Specific.Value = Convert.ToString(afterDIDocNum);
					//	oMat.Columns.Item("InNum").Cells.Item(i).Specific.Value = i;
					//}
				}
				else
				{
					PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
					errCode = "1";
					throw new Exception();
				}

				PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();

				returnValue = true;
			}
			catch (Exception ex)
			{
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}

				if (errCode == "1")
				{
					PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
				}
				else if (errCode == "2")
				{
					PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
				}
				else if (errCode == "3")
				{
					//PS_MM180_InterfaceB1toR3에서 오류 발생하면 해당 메소드에서 오류 메시지 출력, 이 분기문에서는 별도 메시지 출력 안함
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
				}
			}
			finally
			{
				if (oDIObject != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
				}

				if (ProgBar01 != null)
				{
					ProgBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
				}
			}

			return returnValue;
		}

		/// <summary>
		/// 출고DI
		/// </summary>
		/// <returns></returns>
		private bool PS_MM152_DI_API2()
		{
			bool returnValue = false;
			string errCode = string.Empty;
			string errDIMsg = string.Empty;
			int errDICode = 0;
			int i;
			int RetVal;
			int LineNumCount;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Documents oDIObject = null;
			SAPbouiCOM.ProgressBar ProgBar01 = null;
			try
			{
				ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

				PSH_Globals.oCompany.StartTransaction();

				//현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
				if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
				{
					errCode = "2";
					throw new Exception();
				}

				LineNumCount = 0;
				oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
				if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
				{
					oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
				}
				oDIObject.UserFields.Fields.Item("Comments").Value = "가공품출고(외주업체용)_PS_MM152 문서번호:" + oForm.Items.Item("DocEntry").Specific.Value + " 자동불출";

				for (i = 0; i < oMat.VisualRowCount - 1; i++)
				{
					oDIObject.Lines.Add();
					oDIObject.Lines.SetCurrentLine(LineNumCount);
					oDIObject.Lines.ItemCode = oDS_PS_MM152L.GetValue("U_OutItmCd", i).ToString().Trim();
					oDIObject.Lines.WarehouseCode = "802";
					oDIObject.Lines.Quantity = Convert.ToDouble(oDS_PS_MM152L.GetValue("U_OutWt", i).ToString().Trim());
					oDIObject.Lines.Price = Convert.ToDouble(0);
					oDIObject.Lines.LineTotal = Convert.ToDouble(0);
					oDIObject.Lines.UserFields.Fields.Item("PriceBefDi").Value = Convert.ToDouble(0);
					oDIObject.Lines.UserFields.Fields.Item("U_OrdNum").Value = oDS_PS_MM152L.GetValue("U_ItemCode", i).ToString().Trim();
					oDIObject.Lines.UserFields.Fields.Item("U_sSize").Value = oDS_PS_MM152L.GetValue("U_HeatNo", i).ToString().Trim();
					LineNumCount += 1;
				}

				RetVal = oDIObject.Add();

				if (RetVal == 0)
				{
					//PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

					//for (i = 1; i <= oMat.VisualRowCount; i++)
					//{
					//	oMat.Columns.Item("InDoc").Cells.Item(i).Specific.Value = "";
					//	oMat.Columns.Item("InNum").Cells.Item(i).Specific.Value = "";
					//}
				}
				else
				{
					PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
					errCode = "1";
					throw new Exception();
				}

				PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();

				returnValue = true;
			}
			catch (Exception ex)
			{
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}

				if (errCode == "1")
				{
					PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
				}
				else if (errCode == "2")
				{
					PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
				}
				else if (errCode == "3")
				{
					//PS_MM180_InterfaceB1toR3에서 오류 발생하면 해당 메소드에서 오류 메시지 출력, 이 분기문에서는 별도 메시지 출력 안함
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
				}
			}
			finally
			{
				if (oDIObject != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
				}

				if (ProgBar01 != null)
				{
					ProgBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
				}
			}

			return returnValue;
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
			string sQry1;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet1 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
							oMat.FlushToDataSource();
							sQry = "SELECT U_OKYNC FROM [@PS_MM152H] WHERE DocEntry ='" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() +"'";
							oRecordSet.DoQuery(sQry);

							sQry = "SELECT COUNT(*) FROM[@PS_SY005H] A INNER JOIN[@PS_SY005L] B ON A.Code = B.Code where A.Code = 'MM152' AND B.U_UseYN = 'Y' AND B.U_AppUser = '" + PSH_Globals.oCompany.UserName + "'";
							oRecordSet1.DoQuery(sQry);
							if(oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
                            {
								if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "Y")
								{
									if (oRecordSet1.Fields.Item(0).Value.ToString().Trim() == "1")
									{
										for (i = 0; i <= oMat.RowCount - 2; i++)
										{
											sQry = "UPDATE [@PS_MM152L] SET U_DNQty ='" + Convert.ToDouble(oDS_PS_MM152L.GetValue("U_DNQty", i).ToString().Trim()) + "',";
											sQry += " U_QCOKDate = '" + oDS_PS_MM152L.GetValue("U_QCOKDate", i).ToString().Trim() + "' WHERE U_LineNum ='" + (i+1) + "'";
											oRecordSet.DoQuery(sQry);
										}
									}
									else
									{
										errMessage = "이미 승인되어있습니다. 승인된 문서는 재차 승인할수없습니다.";
										PSH_Globals.SBO_Application.MessageBox(errMessage);
										BubbleEvent = false;
										return;
									}
								}
							}

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

							if (oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y" || oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() == "C")
                            {
								sQry = "select count(*) from [@PS_SY005H] A INNER JOIN [@PS_SY005L] B ON A.Code = B.Code WHERE A.Code ='MM158' AND B.U_UseYN ='Y' AND B.U_AppUser ='" + PSH_Globals.oCompany.UserName + "'";
								oRecordSet.DoQuery(sQry);

								for (i = 0; i <= oMat.RowCount - 2; i++)
								{
									if (string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_PP040Doc", i).ToString().Trim()) || oDS_PS_MM152H.GetValue("U_OKYNC", 0).ToString().Trim() == "C")
									{
										//시스템 코드등록에 자재담당자로 등록되어있고, 입고문서가 없으면 승인처리 가능
										if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "1")
										{
											sQry1 = "Select U_ItmBSort From [OITM] Where ItemCode = '" + oDS_PS_MM152L.GetValue("U_ItemCode", i).ToString().Trim() + "'";
											oRecordSet1.DoQuery(sQry1);
											if (oRecordSet1.Fields.Item(0).Value.ToString().Trim() == "105" && oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() != "C")
											{
												if (!string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_MSTNAM", i).ToString().Trim()))
												{
													if (PS_MM152_Add_PS_PP040(ref pVal) == false)
													{
														BubbleEvent = false;
														return;
													}
													if (!string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_OutItmCd", i).ToString().Trim()))
													{
														if (PS_MM152_DI_API2() == false) //자동불출
														{
															BubbleEvent = false;
															return;
														}
													}
												}
											}
											else
											{
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
												if (oRecordSet1.Fields.Item(0).Value.ToString().Trim() == "105" && !string.IsNullOrEmpty(oDS_PS_MM152L.GetValue("U_OutItmCd", i).ToString().Trim()))
                                                {
													if (PS_MM152_DI_API() == false) //자동불출취소
													{
														BubbleEvent = false;
														return;
													}
												}
											}
										}
										else
										{
											errMessage = "자재담당자만 승인또는 승인취소가 가능합니다.";
											BubbleEvent = false;
											throw new System.Exception();
										}
									}
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

								if (PSH_Globals.oCompany.UserName.ToString().Trim() != oRecordSet.Fields.Item(0).Value.ToString().Trim())
								{
									sQry1 = "SELECT COUNT(*) FROM [@PS_SY005H] A INNER JOIN [@PS_SY005L] B ON A.Code = B.Code where A.Code ='M152' AND B.U_UseYN ='Y' AND B.U_AppUser = '" + PSH_Globals.oCompany.UserName + "'";
									oRecordSet1.DoQuery(sQry1);
									if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "1")
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
			oForm.Freeze(true);
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
							if (pVal.ColUID == "QCOKDate")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									oDS_PS_MM152L.SetValue("U_QCOKDate", pVal.Row - 1, DateTime.Now.ToString("yyyyMMdd"));
									BubbleEvent = false;
								}
                                oMat.LoadFromDataSource();
                                oMat.AutoResizeColumns();
								oMat.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}

							if (pVal.ColUID == "MSTCOD")
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
								oMat.AutoResizeColumns();
								//oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}
							else if (pVal.ColUID == "OutQty")
							{
								sQry = "Select U_ItmBSort From OITM Where ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "102")
								{
									oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								}
                                else
                                {
									oMat.Columns.Item("MUseQty").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								}
							}
							else if (pVal.ColUID == "OutWt")
							{
								sQry = "Select U_ItmBSort From OITM Where ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "105")
								{
									oMat.Columns.Item("MUseWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								}
							}
							else if (pVal.ColUID == "NQty")
							{
								sQry = "Select U_ItmBSort From OITM Where ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "102" || oRecordSet.Fields.Item(0).Value.ToString().Trim() == "105")
								{
									oMat.Columns.Item("NWeight").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("NQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								}
							}
							else if (pVal.ColUID == "MSTCOD")
							{
								sQry = "Select U_FULLNAME From [@PH_PY001A] Where Code = '" + oMat.Columns.Item("MSTCOD").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								oMat.Columns.Item("MSTNAM").Cells.Item(pVal.Row).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
			string errmsg = string.Empty;
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					
				}
				else if (pVal.BeforeAction == false)
				{
					if(pVal.ItemUID == "BPLId")
                    {
						if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "2")
						{
							oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oMat.Columns.Item("HeatNo").Visible = true;
							oMat.Columns.Item("DNQty").Visible = true;
							oMat.Columns.Item("AttPath").Visible = true;
							oMat.Columns.Item("Action").Visible = true;
							oMat.Columns.Item("QCOKDate").Visible = true;
							oMat.Columns.Item("MSTCOD").Visible = true;
							oMat.Columns.Item("MSTNAM").Visible = true;



							oMat.Columns.Item("ScrapWt").Visible = false;
							oMat.Columns.Item("CPWt").Visible = false;
							oMat.Columns.Item("CPWtName").Visible = false;
							oMat.Columns.Item("Sample").Visible = false;
							oMat.Columns.Item("Loss").Visible = false;

							PS_MM152_Add_MatrixRow(0, true);
						}
						else
						{
							oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oMat.Columns.Item("HeatNo").Visible = false;
							oMat.Columns.Item("DNQty").Visible = false;
							oMat.Columns.Item("AttPath").Visible = false;
							oMat.Columns.Item("Action").Visible = false;
							oMat.Columns.Item("QCOKDate").Visible = false;
							oMat.Columns.Item("MSTCOD").Visible = false;
							oMat.Columns.Item("MSTNAM").Visible = false;

							oMat.Columns.Item("ScrapWt").Visible = true;
							oMat.Columns.Item("CPWt").Visible = true;
							oMat.Columns.Item("CPWtName").Visible = true;
							oMat.Columns.Item("Sample").Visible = true;
							oMat.Columns.Item("Loss").Visible = true;

							PS_MM152_Add_MatrixRow(0, true);
						}
					}

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
					else if (pVal.ItemUID == "Mat01" && pVal.ColUID == "Action")
					{
						if (oMat.Columns.Item("Action").Cells.Item(pVal.Row).Specific.Value == "S")
						{
							if(oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() == "Y")
                            {
								errmsg = "승인된 문서는 PDF파일 저장이 불가능합니다.";
								throw new Exception();
							}
                            else
                            {
								PS_MM152_SaveAttach(pVal.Row);
							}
						}
						else if (oMat.Columns.Item("Action").Cells.Item(pVal.Row).Specific.Value == "O")
						{
							PS_MM152_OpenAttach(pVal.Row);
						}
					}
				}
			}
			catch (Exception ex)
			{
				if (errmsg != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errmsg);
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
					oMat.AutoResizeColumns();
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
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
							// 마감일자 Check
							if (dataHelpClass.Check_Finish_Status(oDS_PS_MM152H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM152H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
							{
								PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 취소할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
								BubbleEvent = false;
								return;
							}
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								BubbleEvent = false;
								return;
							}
							break;
						case "1286": //닫기
							// 마감일자 Check
							if (dataHelpClass.Check_Finish_Status(oDS_PS_MM152H.GetValue("U_BPLId", 0).ToString().Trim(), oDS_PS_MM152H.GetValue("U_DocDate", 0).ToString().Trim().Substring(0, 6)) == false)
							{
								PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 닫기할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
								BubbleEvent = false;
								return;
							}
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							sQry = "SELECT COUNT(*) FROM [@PS_SY005H] A INNER JOIN [@PS_SY005L] B ON A.Code = B.Code where A.Code ='M152' AND B.U_UseYN ='Y' AND B.U_AppUser = '" + PSH_Globals.oCompany.UserName + "'";
							oRecordSet.DoQuery(sQry);
							if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "1")
							{
								Raise_EVENT_RECORD_MOVE(FormUID, ref pVal, ref BubbleEvent);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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

		/// <summary>
		/// 네비게이션 메소드(Raise_FormMenuEvent 에서 사용)
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_RECORD_MOVE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			string docEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				docEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim(); //현재문서번호

				if (pVal.MenuUID == "1288") //다음
				{
					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
					{
						PSH_Globals.SBO_Application.ActivateMenuItem("1290");
						return;
					}
					else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
					{
						if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))
						{
							PSH_Globals.SBO_Application.ActivateMenuItem("1290");
							return;
						}
					}
					else
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
						oForm.Items.Item("DocEntry").Enabled = true;
						sQry = "  Select min(DocEntry)";
						sQry += "  From [@PS_MM152H]";
						sQry += " Where U_CardCode = '" + PSH_Globals.oCompany.UserName + "'";
						sQry += "   AND DocEntry > " + docEntry;

						oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(sQry, 0, 1);
						oForm.Items.Item("1").Enabled = true;
						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("DocEntry").Enabled = false;
					}
				}
				else if (pVal.MenuUID == "1289") //이전
				{
					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
					{
						PSH_Globals.SBO_Application.ActivateMenuItem("1291");
						return;
					}
					else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
					{
						if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))
						{
							PSH_Globals.SBO_Application.ActivateMenuItem("1291");
							return;
						}
					}
					else
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
						oForm.Items.Item("DocEntry").Enabled = true;
						sQry = "  Select max(DocEntry)";
						sQry += "  From [@PS_MM152H]";
						sQry += " Where U_CardCode = '" + PSH_Globals.oCompany.UserName + "'";
						sQry += "   AND DocEntry < " + docEntry;

						oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(sQry, 0, 1);
						oForm.Items.Item("1").Enabled = true;
						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("DocEntry").Enabled = false;
					}
				}
				else if (pVal.MenuUID == "1290") //최초
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					oForm.Items.Item("DocEntry").Enabled = true;
					sQry = "  Select Min(DocEntry)";
					sQry += "  From [@PS_MM152H]";
					sQry += " Where U_CardCode = '" + PSH_Globals.oCompany.UserName + "'";

					oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(sQry, 0, 1);
					oForm.Items.Item("1").Enabled = true;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("DocEntry").Enabled = false;
				}
				else if (pVal.MenuUID == "1291") //최종
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					oForm.Items.Item("DocEntry").Enabled = true;
					sQry = "  Select Max(DocEntry)";
					sQry += "  From [@PS_MM152H]";
					sQry += " Where U_CardCode = '" + PSH_Globals.oCompany.UserName + "'";

					oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(sQry, 0, 1);
					oForm.Items.Item("1").Enabled = true;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("DocEntry").Enabled = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				BubbleEvent = false;
			}
		}
	}
}
