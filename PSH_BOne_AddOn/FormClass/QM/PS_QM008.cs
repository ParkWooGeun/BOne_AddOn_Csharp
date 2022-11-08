using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 품질관리>(분말)제품검사성적서 등록
	/// </summary>
	internal class PS_QM008 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_QM008H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM008L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM008.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM008_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM008");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_QM008_CreateItems();
				PS_QM008_ComboBox_Setting();
				PS_QM008_Initial_Setting();
				PS_QM008_EnableMenus();
				PS_QM008_SetDocument(oFormDocEntry);

				oForm.EnableMenu("1283", true); // 삭제
				oForm.EnableMenu("1287", true); // 복제
				oForm.EnableMenu("1286", true); // 닫기
				oForm.EnableMenu("1284", true); // 취소
				oForm.EnableMenu("1293", true); // 행삭제
				oForm.Items.Item("InspNo").Click(); //검사의뢰번호 포커서
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
		/// PS_QM008_CreateItems
		/// </summary>
		private void PS_QM008_CreateItems()
		{
			try
			{
				oDS_PS_QM008H = oForm.DataSources.DBDataSources.Item("@PS_QM008H");
				oDS_PS_QM008L = oForm.DataSources.DBDataSources.Item("@PS_QM008L");
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
		/// PS_QM008_ComboBox_Setting
		/// </summary>
		private void PS_QM008_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

				//합부판정
				oForm.Items.Item("PASSYN").Specific.ValidValues.Add("Y", "합격");
				oForm.Items.Item("PASSYN").Specific.ValidValues.Add("N", "불합격");
				oForm.Items.Item("PASSYN").Specific.ValidValues.Add("S", "특채");
				oForm.Items.Item("PASSYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("VIEWYN").Specific.ValidValues.Add("Y", "출력");
				oForm.Items.Item("VIEWYN").Specific.ValidValues.Add("N", "미출력");
				oForm.Items.Item("VIEWYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("SOVIEWYN").Specific.ValidValues.Add("Y", "출력");
				oForm.Items.Item("SOVIEWYN").Specific.ValidValues.Add("N", "미출력");
				oForm.Items.Item("SOVIEWYN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("ReleaYN").Specific.ValidValues.Add("Y", "금지");
				oForm.Items.Item("ReleaYN").Specific.ValidValues.Add("N", "허용");
				oForm.Items.Item("ReleaYN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM008_Initial_Setting
		/// </summary>
		private void PS_QM008_Initial_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				//일자
				oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("PPYN").Specific.Value = "N";
				oForm.Items.Item("QMYN").Specific.Value = "N";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM008_EnableMenus
		/// </summary>
		private void PS_QM008_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM008_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_QM008_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_QM008_FormItemEnabled();
					PS_QM008_AddMatrixRow(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_QM008_FormItemEnabled();
					oForm.Items.Item("DocEntry").Specific.Value = oFromDocEntry01;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM008_FormItemEnabled
		/// </summary>
		private void PS_QM008_FormItemEnabled()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{ 
					oForm.Items.Item("Mat01").Enabled = true;
					PS_QM008_FormClear();
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
					oForm.Items.Item("CardCode").Specific.Value = "12549";

					sQry = "select count(*) from [@PS_SY001L] where code ='QM008_01' and U_Minor = '" + PSH_Globals.oCompany.UserName + "'";
					oRecordSet.DoQuery(sQry);

					if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) < 1)
					{
						oForm.Items.Item("1").Enabled = false;
					}
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocEntry").Specific.Value = "";
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가
					oForm.Items.Item("1").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					oForm.Items.Item("Remark").Click();

					sQry = "select count(*) from [@PS_SY001L] where code ='QM008_01' and U_Minor = '" + PSH_Globals.oCompany.UserName + "'";
					oRecordSet.DoQuery(sQry);

					if (Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) < 1)
					{
						oForm.Items.Item("1").Enabled = false;
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
		/// PS_QM008_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM008_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_QM008L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM008L.Offset = oRow;
				oDS_PS_QM008L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_QM008_CopyMatrixRow
		/// </summary>
		private void PS_QM008_CopyMatrixRow()
		{
			int i;

			try
			{
				oForm.Freeze(true);
				oDS_PS_QM008H.SetValue("DocEntry", 0, "");
				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_QM008H.SetValue("DocEntry", i, "");
					oMat.LoadFromDataSource();
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
		/// PS_QM008_FormClear
		/// </summary>
		private void PS_QM008_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM008'", "");
				if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
				{
					oForm.Items.Item("DocEntry").Specific.Value = "1";
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
		/// PS_QM008_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_QM008_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			decimal SPEC_MIN;
			decimal SPEC_MAX;
			decimal PVALUE;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM008_FormClear();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "검사일자가 입력되지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("CardSeq").Specific.Value.ToString().Trim()))
				{
					errMessage = "고객번호가 입력되지 않았습니다.";
					throw new Exception();
				}
				if (oMat.VisualRowCount == 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				sQry = " select Count(b.lineid) as CntLineid ";
				sQry += "  from [@PS_QM007H] a inner join [@PS_QM007L] b on a.DocEntry = b.DocEntry and b.U_UseYN ='Y' and Canceled ='N'";
				sQry += " Where a.U_CardCode ='" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
				sQry += "   and a.U_ItemCode ='" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
				sQry += "   and a.U_CardSeq ='" + oForm.Items.Item("CardSeq").Specific.Value.ToString().Trim() + "'";
				oRecordSet.DoQuery(sQry);

				if (oMat.VisualRowCount-1 != Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()))
				{
					errMessage = "성적서등록 화면내 검사항목과 사양서화면내 검사항목이 불일치합니다.";
					throw new Exception();
				}

				if (oForm.Items.Item("PASSYN").Specific.Value.ToString().Trim() == "Y") //합격일때만 검사치수CHECK
				{
					for (i = 1; i <= oMat.VisualRowCount - 1; i++)
					{
						if (oMat.Columns.Item("InspBal").Cells.Item(i).Specific.Value.ToString().Trim() != "Y")
						{
							PVALUE = Convert.ToDecimal(oMat.Columns.Item("Value").Cells.Item(i).Specific.Value.ToString().Trim());
							SPEC_MIN = Convert.ToDecimal(oMat.Columns.Item("InspMin").Cells.Item(i).Specific.Value.ToString().Trim());
							SPEC_MAX = Convert.ToDecimal(oMat.Columns.Item("InspMax").Cells.Item(i).Specific.Value.ToString().Trim());

							if (SPEC_MIN == 0 && SPEC_MAX == 0)
							{
								if (PVALUE != 0)
								{
									oMat.Columns.Item("Value").Cells.Item(i).Click();
									errMessage = "검사치수를 확인 하십시요.";
									throw new Exception();
								}
							}
							else if (PVALUE < SPEC_MIN || PVALUE > SPEC_MAX)
							{
								oMat.Columns.Item("Value").Cells.Item(i).Click();
								errMessage = "검사치수를 확인 하십시요.";
								throw new Exception();
							}
						}
					}

					if (oForm.Items.Item("IpdoChk").Specific.Value.ToString().Trim() == "Y")
					{
						//입도분포CHECK = 합이 100
						PVALUE = 0;
						for (i = 1; i <= oMat.VisualRowCount - 1; i++)
						{
							if (oMat.Columns.Item("InspItem").Cells.Item(i).Specific.Value.ToString().Trim() == "입도분포")
							{
								PVALUE += Convert.ToDecimal(oMat.Columns.Item("Value").Cells.Item(i).Specific.Value.ToString().Trim());
							}
						}

						if (PVALUE != 100)
						{
							errMessage = "입도분포의 합이(100)이 아닙니다. 확인 하십시요.";
							throw new Exception();
						}
					}
				}

				oMat.FlushToDataSource();
				oDS_PS_QM008L.RemoveRecord(oDS_PS_QM008L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM008_FormClear();
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_QM008_LoadData
		/// </summary>
		private void PS_QM008_LoadData()
		{
			int i;
			string Sintern = string.Empty;
			string IpdoChk = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				sQry = "Select b.U_InspItem, b.U_InspItNm, b.U_InspUnit, U_InspSpec, b.U_InspMeth, b.U_InspMin, b.U_InspMax, b.U_InspBal, a.U_Sintern, a.U_IpdoChk ";
				sQry += " From [@PS_QM007H] a INNER JOIN [@PS_QM007L] b ON a.DocEntry = b.DocEntry AND a.Canceled = 'N' ";
				sQry += "Where a.U_ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "' ";
				sQry += "  AND a.U_CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "' ";
				sQry += "  And a.U_CardSeq  = '" + oForm.Items.Item("CardSeq").Specific.Value.ToString().Trim() + "' ";
				sQry += "  AND b.U_UseYN = 'Y' Order By b.U_Seqno ";
				oRecordSet.DoQuery(sQry);

				oDS_PS_QM008L.Clear();
				oMat.Clear();
				oMat.FlushToDataSource();

				i = 0;
				while (!oRecordSet.EoF)
				{
					oDS_PS_QM008L.InsertRecord(i);
					oDS_PS_QM008L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM008L.SetValue("U_InspItem", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_QM008L.SetValue("U_InspItNm", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_QM008L.SetValue("U_InspUnit", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_QM008L.SetValue("U_InspSpec", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_QM008L.SetValue("U_InspMeth", i, oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_QM008L.SetValue("U_InspMin", i, oRecordSet.Fields.Item(5).Value.ToString().Trim());
					oDS_PS_QM008L.SetValue("U_InspMax", i, oRecordSet.Fields.Item(6).Value.ToString().Trim());
					oDS_PS_QM008L.SetValue("U_InspBal", i, oRecordSet.Fields.Item(7).Value.ToString().Trim());
					Sintern = oRecordSet.Fields.Item(8).Value.ToString().Trim();
					IpdoChk = oRecordSet.Fields.Item(9).Value.ToString().Trim();
					i += 1;
					oRecordSet.MoveNext();
				}

				oDS_PS_QM008H.SetValue("U_Sintern", 0, Sintern);
				oDS_PS_QM008H.SetValue("U_IpdoChk", 0, IpdoChk);

				oMat.LoadFromDataSource();
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
		/// PS_QM008_Version_Combo
		/// </summary>
		/// <param name="Version"></param>
		private void PS_QM008_Version_Combo(string Version)
		{
			bool comboCount = false;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				sQry = "select U_Minor, U_CdName from [@PS_SY001L] where code = 'Q017' and U_RelCd ='";
				sQry += oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
				oRecordSet.DoQuery(sQry);
				
				if (oForm.Items.Item("Version").Specific.ValidValues.Count > 0)
				{
					comboCount = true;
				}
				else
				{
					comboCount = false;
				}

				dataHelpClass.Set_ComboList(oForm.Items.Item("Version").Specific, sQry, Version, comboCount, false);
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
		/// PS_QM008_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_QM008_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string DocEntry;
			string VIEWYN;
			string SoviewYN;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				VIEWYN = oForm.Items.Item("VIEWYN").Specific.Value.ToString().Trim();
				SoviewYN = oForm.Items.Item("SOVIEWYN").Specific.Value.ToString().Trim();

				WinTitle = "[PS_QM008_10] 검사성적서 출력(한글)";
				ReportName = "PS_QM008_10.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));
				dataPackParameter.Add(new PSH_DataPackClass("@VIEWYN", VIEWYN));
				dataPackParameter.Add(new PSH_DataPackClass("@SOVIEWYN", SoviewYN));
				dataPackParameter.Add(new PSH_DataPackClass("@Gubun", "Q"));
				dataPackParameter.Add(new PSH_DataPackClass("@Lang", "K"));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM008_Print_Report02
		/// </summary>
		[STAThread]
		private void PS_QM008_Print_Report02()
		{
			string WinTitle;
			string ReportName;
			string DocEntry;
			string VIEWYN;
			string SoviewYN;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				VIEWYN = oForm.Items.Item("VIEWYN").Specific.Value.ToString().Trim();
				SoviewYN = oForm.Items.Item("SOVIEWYN").Specific.Value.ToString().Trim();

				WinTitle = "[PS_QM008_20] 검사성적서 출력(한문)";
				ReportName = "PS_QM008_20.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));
				dataPackParameter.Add(new PSH_DataPackClass("@VIEWYN", VIEWYN));
				dataPackParameter.Add(new PSH_DataPackClass("@SOVIEWYN", SoviewYN));
				dataPackParameter.Add(new PSH_DataPackClass("@Gubun", "Q"));
				dataPackParameter.Add(new PSH_DataPackClass("@Lang", "C"));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM008_Print_Report03
		/// </summary>
		[STAThread]
		private void PS_QM008_Print_Report03()
		{
			string WinTitle;
			string ReportName;
			string DocEntry;
			string VIEWYN;
			string SoviewYN;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				VIEWYN = oForm.Items.Item("VIEWYN").Specific.Value.ToString().Trim();
				SoviewYN = oForm.Items.Item("SOVIEWYN").Specific.Value.ToString().Trim();

				WinTitle = "[PS_QM008_30] 검사성적서 출력(영문)";
				ReportName = "PS_QM008_30.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));
				dataPackParameter.Add(new PSH_DataPackClass("@VIEWYN", VIEWYN));
				dataPackParameter.Add(new PSH_DataPackClass("@SOVIEWYN", SoviewYN));
				dataPackParameter.Add(new PSH_DataPackClass("@Gubun", "Q"));
				dataPackParameter.Add(new PSH_DataPackClass("@Lang", "E"));

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
				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
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
			string InspNo;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_QM008_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							//Start) SY001 - Q017에 등록된 거래처는 무조건 검사버전관리에 값을 입력해야한다.
							sQry = "select distinct U_RelCd from [@PS_SY001L] where code ='Q017' and U_RelCd ='" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "' ";
							oRecordSet.DoQuery(sQry);

							if (oRecordSet.RecordCount > 0)
							{
								if (string.IsNullOrEmpty(oForm.Items.Item("Version").Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.MessageBox("검사버전관리에 값을 입력하세요.");
									oForm.Items.Item("Version").Click();
									BubbleEvent = false;
									return;
								}
							}
							
							InspNo = oForm.Items.Item("InspNo").Specific.Value.ToString().Trim();
							//검사입력되면 검사완료 Sign Update
							sQry = "Update [Z_PACKING_PD] Set QM006YN = 'Y' Where InspNo = '" + InspNo + "'";
							oRecordSet.DoQuery(sQry);

							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							PS_QM008_Version_Combo(oDS_PS_QM008H.GetValue("U_Version", 0).ToString().Trim());
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM008_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}

							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
					}
					if (pVal.ItemUID == "Button01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM008_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					if (pVal.ItemUID == "Button02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
							{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM008_Print_Report02);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					if (pVal.ItemUID == "Button03")
					{
			            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM008_Print_Report03);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_QM008_FormItemEnabled();
								PS_QM008_AddMatrixRow(0, true);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							PS_QM008_Version_Combo(oDS_PS_QM008H.GetValue("U_Version", 0).ToString().Trim());
							if (pVal.ActionSuccess == true)
							{
								PS_QM008_FormItemEnabled();
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
			}
		}

		/// <summary>
		/// KEY_DOWN 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "CntcCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "CardCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "InspNo")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("InspNo").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "CardSeq")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CardSeq").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
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
		/// Raise_EVENT_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oLastItemUID01 = pVal.ItemUID;
							oLastColUID01 = pVal.ColUID;
							oLastColRow01 = pVal.Row;
							oMat.SelectRow(pVal.Row, true, false);
						}
					}
					else
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = "";
						oLastColRow01 = 0;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					// 출고금지여부 해당된사람만 입력가능 (SY001 Q016 코드 등록해야함)
					if (pVal.ItemUID == "ReleaYN")
					{
						sQry = "select U_Minor from [@PS_SY001L] where 1=1   and code ='Q016'   and U_Minor ='" + PSH_Globals.oCompany.UserName + "'";
						oRecordSet.DoQuery(sQry);

						if (oRecordSet.RecordCount != 0)
						{
						}
						else
						{
							PSH_Globals.SBO_Application.MessageBox("해당 항목 수정불가능합니다.");
							oForm.Items.Item("ReleaYN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
							oForm.Items.Item("InspNo").Click();
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
			decimal SPEC_MIN;
			decimal SPEC_MAX;
			decimal PVALUE;
			string CardCode;
			string ItemCode;
			string errMessage = string.Empty;
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
							if (pVal.ColUID == "Value") //검사치수 CHECK
							{
								if (oMat.Columns.Item("InspBal").Cells.Item(pVal.Row).Specific.Value != "Y")
								{
									PVALUE   = Convert.ToDecimal(oMat.Columns.Item("Value").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									SPEC_MIN = Convert.ToDecimal(oMat.Columns.Item("InspMin").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									SPEC_MAX = Convert.ToDecimal(oMat.Columns.Item("InspMax").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());

									if (SPEC_MIN == 0 && SPEC_MAX == 0)
									{
										if (PVALUE != 0)
										{
											errMessage = "검사치수와 검사규격을 확인하여 주십시오.";
											throw new Exception();
										}
									}
									else if (PVALUE < SPEC_MIN || PVALUE > SPEC_MAX)
									{
										errMessage = "검사치수와 검사규격을 확인하여 주십시오.";
										throw new Exception();
									}
								}
							}
						}
						else if (pVal.ItemUID == "InspNo") //검사의뢰번호
						{
							sQry = "Select Itemcode, Quantity = Sum(Quantity), CardName, CardSeq, CardCode From [Z_PACKING_PD] Where QM006YN = 'N' AND InspNo = '" + oForm.Items.Item("InspNo").Specific.Value.ToString().Trim() + "' Group by Itemcode, CardName, CardSeq, CardCode ";
                            oRecordSet.DoQuery(sQry);

                            oDS_PS_QM008H.SetValue("U_Weight", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oForm.Items.Item("ItemCode").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oForm.Items.Item("CardCode").Specific.Value = oRecordSet.Fields.Item(4).Value.ToString().Trim();

							if (!string.IsNullOrEmpty(oRecordSet.Fields.Item(2).Value.ToString().Trim()))
							{
								if (oRecordSet.Fields.Item(3).Value.ToString().Trim() == "00")
								{
									oForm.Items.Item("Remark").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
								}
								else
								{
									oForm.Items.Item("Remark").Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim() + " #" + oRecordSet.Fields.Item(3).Value.ToString().Trim();
								}
							}
							else
							{
								oForm.Items.Item("Remark").Specific.Value = "";
							}

							oForm.Items.Item("ItemCode").Enabled = false;
						}
						else if (pVal.ItemUID == "CntcCode") //사번
						{
							oDS_PS_QM008H.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
						}
						else if (pVal.ItemUID == "ItemCode") //품목코드
						{
							sQry = "Select ItemName, FrgnName, U_Size From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oDS_PS_QM008H.SetValue("U_ItemName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_QM008H.SetValue("U_FrgnName", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_QM008H.SetValue("U_Size", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());

							CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
							ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

							sQry = "Select Count(*), Max(U_CardSeq) From [@PS_QM007H] Where U_CardCode = '" + CardCode + "' And U_ItemCode = '" + ItemCode + "'";
							oRecordSet.DoQuery(sQry);

							if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 1)
							{
								oDS_PS_QM008H.SetValue("U_CardSeq", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oForm.Items.Item("CardSeq").Enabled = false;
								PS_QM008_LoadData();
							}
							else
							{
								oForm.Items.Item("CardSeq").Enabled = true;
								oDS_PS_QM008H.SetValue("U_CardSeq", 0, "");
							}
						}
						else if (pVal.ItemUID == "CardCode") //거래처코드
						{
							sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" + oDS_PS_QM008H.GetValue("U_CardCode", 0).ToString().Trim() +"'";
							oRecordSet.DoQuery(sQry);

							oDS_PS_QM008H.SetValue("U_CardName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
							ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

							sQry = "Select Count(*), Max(U_CardSeq) From [@PS_QM007H] Where U_CardCode = '" + CardCode + "' And U_ItemCode = '" + ItemCode + "'";
							oRecordSet.DoQuery(sQry);

							if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 1)
							{
								oDS_PS_QM008H.SetValue("U_CardSeq", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oForm.Items.Item("CardSeq").Enabled = false;
								PS_QM008_LoadData();
							}
							else
							{
								oForm.Items.Item("CardSeq").Enabled = true;
								oDS_PS_QM008H.SetValue("U_CardSeq", 0, "");
							}
						}
						else if (pVal.ItemUID == "CardSeq") //거래처 순번
						{
							if (!string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()) && !string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()) && !string.IsNullOrEmpty(oForm.Items.Item("CardSeq").Specific.Value.ToString().Trim()))
							{
								PS_QM008_LoadData();
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "InspNo" || pVal.ItemUID == "CardCode") //검사의뢰번호
					{
						PS_QM008_Version_Combo(oDS_PS_QM008H.GetValue("U_Version", 0));
					}
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_CLICK 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_QM008_FormItemEnabled();
					PS_QM008_AddMatrixRow(oMat.VisualRowCount, false);
					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_MATRIX_LOAD_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM008H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM008L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_ROW_DELETE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_QM008L.RemoveRecord(oDS_PS_QM008L.Size - 1);
						oMat.LoadFromDataSource();

						if (oMat.RowCount == 0)
						{
							PS_QM008_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_QM008L.GetValue("U_InspItem", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_QM008_AddMatrixRow(oMat.RowCount, false);
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
			string InspNo;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1283": //삭제
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
							{   //검사삭제시 검사완료 Update
								InspNo = oForm.Items.Item("InspNo").Specific.Value.ToString().Trim();
								sQry = "Update [Z_PACKING_PD] Set QM006YN = 'N' Where InspNo = '" + InspNo + "'";
								oRecordSet.DoQuery(sQry);
							}
							break;
						case "1284": //취소
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
							{
								//검사삭제시 검사완료 Update
								InspNo = oForm.Items.Item("InspNo").Specific.Value.ToString().Trim();
								sQry = "Update [Z_PACKING_PD] Set QM006YN = 'N' Where InspNo = '" + InspNo + "'";
								oRecordSet.DoQuery(sQry);
							}
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							PS_QM008_FormItemEnabled();
							break;
						case "1282": //추가
							PS_QM008_FormItemEnabled();
							PS_QM008_AddMatrixRow(0, true);
							break;
						case "1287": //복제 
							PS_QM008_CopyMatrixRow();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM008_FormItemEnabled();
							PS_QM008_Version_Combo(oDS_PS_QM008H.GetValue("U_Version", 0));
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
