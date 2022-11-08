using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 멀티제품검사성적서등록(신)
	/// </summary>
	internal class PS_QM610 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_QM610H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM610L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private string Last_InspPrsn;
		private string Last_CardCode;
		private string Last_CardName;
		private string Last_PrsnName;

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM610.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 25;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 21;
				}

				oFormUniqueID = "PS_QM610_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM610");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_QM610_CreateItems();
				PS_QM610_ComboBox_Setting();
				PS_QM610_EnableMenus();
				PS_QM610_SetDocument(oFormDocEntry);
				PS_QM610_Initial_Setting();

				oForm.EnableMenu("1283", true); // 삭제
				oForm.EnableMenu("1287", true); // 복제
				oForm.EnableMenu("1286", true); // 닫기
				oForm.EnableMenu("1284", true); // 취소
				oForm.EnableMenu("1282", true); // 추가
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
		/// PS_QM610_CreateItems
		/// </summary>
		private void PS_QM610_CreateItems()
		{
			try
			{
				oDS_PS_QM610H = oForm.DataSources.DBDataSources.Item("@PS_QM610H");
				oDS_PS_QM610L = oForm.DataSources.DBDataSources.Item("@PS_QM610L");
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
		/// PS_QM610_ComboBox_Setting
		/// </summary>
		private void PS_QM610_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

				// 인장강도단위
				oForm.Items.Item("TS_Gbn").Specific.ValidValues.Add("10", "Kgf/m2");
				oForm.Items.Item("TS_Gbn").Specific.ValidValues.Add("20", "N/mm2");

				// 외관검사
				oForm.Items.Item("M_Visual").Specific.ValidValues.Add("Y", "합격");
				oForm.Items.Item("M_Visual").Specific.ValidValues.Add("N", "불합격");
				oForm.Items.Item("M_Visual").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				// 포장상태
				oForm.Items.Item("PackYN").Specific.ValidValues.Add("Y", "양호");
				oForm.Items.Item("PackYN").Specific.ValidValues.Add("N", "부적합");
				oForm.Items.Item("PackYN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);

				// 합부판정
				oForm.Items.Item("PassYN").Specific.ValidValues.Add("Y", "합격");
				oForm.Items.Item("PassYN").Specific.ValidValues.Add("N", "불합격");
				oForm.Items.Item("PassYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM610_EnableMenus
		/// </summary>
		private void PS_QM610_EnableMenus()
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
		/// PS_QM610_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_QM610_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_QM610_FormItemEnabled();
					PS_QM610_AddMatrixRow(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_QM610_FormItemEnabled();
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
		/// PS_QM610_Initial_Setting
		/// </summary>
		private void PS_QM610_Initial_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //사업장
				oForm.Items.Item("InspDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //일자
				oForm.Items.Item("OrdNum").Click(); //포커서
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM610_FormItemEnabled
		/// </summary>
		private void PS_QM610_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM610_FormClear();
					oForm.Items.Item("Mat01").Enabled = true;
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("OrdNum").Enabled = true;
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가

					oForm.Items.Item("M_Visual").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //외관검사
					oForm.Items.Item("PackYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //포장상태
					oForm.Items.Item("PassYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //합부판정
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("OrdNum").Enabled = true;
					oForm.Items.Item("InspDate").Specific.Value = "";
					oForm.Items.Item("Mat01").Enabled = false;
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("OrdNum").Enabled = false;
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
		/// PS_QM610_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM610_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_QM610L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM610L.Offset = oRow;
				oDS_PS_QM610L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_QM610_CopyMatrixRow
		/// </summary>
		private void PS_QM610_CopyMatrixRow()
		{
			int i;

			try
			{
				oDS_PS_QM610H.SetValue("DocEntry", 0, "");
				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_QM610H.SetValue("DocEntry", i, "");
					oMat.LoadFromDataSource();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM610_FormClear
		/// </summary>
		private void PS_QM610_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM610'", "");
				if (string.IsNullOrEmpty(DocEntry) | DocEntry == "0")
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
		/// PS_QM610_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_QM610_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			string SPEC;
			decimal SPEC_MIN;
			decimal SPEC_MAX;
			decimal VAL_MIN;
			decimal VAL_MAX;
			string errMessage = string.Empty;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM610_FormClear();
				}
				//일자 미입력시
				if (string.IsNullOrEmpty(oForm.Items.Item("InspDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "검사일자가 입력되지 않았습니다.";
					throw new Exception();
				}
				//거래처순번 미입력시
				if (string.IsNullOrEmpty(oForm.Items.Item("CardSeq").Specific.Value.ToString().Trim()))
				{
					errMessage = "거래처순번이 입력되지 않았습니다.";
					throw new Exception();
				}
				//검사자 미입력시
				if (string.IsNullOrEmpty(oForm.Items.Item("InspPrsn").Specific.Value.ToString().Trim()))
				{
					errMessage = "검사자가 입력되지 않았습니다.";
					throw new Exception();
				}
				//라인정보 미입력 시
				if (oMat.VisualRowCount == 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				if (oForm.Items.Item("PassYN").Specific.Selected.Value.ToString().Trim() == "Y")
				{
					for (i = 1; i <= oMat.VisualRowCount - 1; i++)
					{
						SPEC = oMat.Columns.Item("InspSpec").Cells.Item(i).Specific.Value.ToString().Trim();
						VAL_MIN = Convert.ToDecimal(oMat.Columns.Item("ValMin").Cells.Item(i).Specific.Value.ToString().Trim());
						VAL_MAX = Convert.ToDecimal(oMat.Columns.Item("ValMax").Cells.Item(i).Specific.Value.ToString().Trim());
						SPEC_MIN = Convert.ToDecimal(oMat.Columns.Item("InspMin").Cells.Item(i).Specific.Value.ToString().Trim());
						SPEC_MAX = Convert.ToDecimal(oMat.Columns.Item("InspMax").Cells.Item(i).Specific.Value.ToString().Trim());

						if (SPEC == "MAX")
						{
							// MIN 0
							if (VAL_MIN != 0)
							{
								oMat.Columns.Item("ValMin").Cells.Item(i).Click();
								errMessage = "이항목은 MIN값이 있울수 없습니다 MAX만 입력 하세요.";
								throw new Exception();
							}
							// MAX만 Check
							if (VAL_MAX < SPEC_MIN || VAL_MAX > SPEC_MAX)
							{
								oMat.Columns.Item("ValMax").Cells.Item(i).Click();
								errMessage = "검사치수를 확인 하십시요.";
								throw new Exception();
							}
						}
						else
						{
							if (VAL_MIN < SPEC_MIN || VAL_MIN > SPEC_MAX || VAL_MAX > SPEC_MAX || VAL_MAX < SPEC_MIN)
							{
								oMat.Columns.Item("ValMin").Cells.Item(i).Click();
								errMessage = "검사치수를 확인 하십시요.";
								throw new Exception();
							}
						}
					}
				}

				oMat.FlushToDataSource();
				oDS_PS_QM610L.RemoveRecord(oDS_PS_QM610L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM610_FormClear();
				}

				Last_CardCode = oDS_PS_QM610H.GetValue("U_CardCode", 0).ToString().Trim();
				Last_CardName = oDS_PS_QM610H.GetValue("U_CardName", 0).ToString().Trim();
				Last_InspPrsn = oDS_PS_QM610H.GetValue("U_InspPrsn", 0).ToString().Trim();
				Last_PrsnName = oDS_PS_QM610H.GetValue("U_PrsnName", 0).ToString().Trim();

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
		/// 사양서DATA를 기본셋팅
		/// </summary>
		private void PS_QM610_LoadData()
		{
			int i;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "Select b.U_InspItem, c.U_CdName, b.U_InspItNm, d.U_CdName, U_InspSpec, b.U_InspMin, b.U_InspMax ";
				sQry += " From [@PS_QM600H] a INNER JOIN [@PS_QM600L] b ON a.DocEntry = b.DocEntry AND a.Canceled = 'N' ";
				sQry += " LEFT  JOIN [@PS_SY001L] c ON c.Code = 'Q600' AND c.U_Minor = b.U_InspItem ";
				sQry += " LEFT  JOIN [@PS_SY001L] d ON d.Code = 'Q600' AND d.U_Minor = b.U_InspItNm ";
				sQry += "Where a.U_ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "' ";
				sQry += "  AND a.U_CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "' ";
				sQry += "  And a.U_CardSeq  = '" + oForm.Items.Item("CardSeq").Specific.Value.ToString().Trim() + "' ";
				sQry += "  AND b.U_UseYN = 'Y' Order By b.U_Seqno ";
				oRecordSet.DoQuery(sQry);

				//원소재성분이 입력[PS_QM030]된 Data Select
				sQry = "select Top 1 U_ChemC_Cu, U_ChemC_Fe, U_ChemC_P from [OBTN] ";
				sQry += "where InDate >  dateadd(YEAR, -9, getdate())";
				sQry += "  and DistNumber like '" + oDS_PS_QM610H.GetValue("U_LotNo", 0).ToString().Trim() + "' + '%'";
				oRecordSet01.DoQuery(sQry);

				oDS_PS_QM610L.Clear();
				oMat.Clear();
				oMat.FlushToDataSource();

				i = 0;
				while (!oRecordSet.EoF)
				{
					oDS_PS_QM610L.InsertRecord(i);
					oDS_PS_QM610L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM610L.SetValue("U_InspItem", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_QM610L.SetValue("U_ItemDscr", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_QM610L.SetValue("U_InspItNm", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_QM610L.SetValue("U_ItNmDscr", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_QM610L.SetValue("U_InspSpec", i, oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_QM610L.SetValue("U_InspMin", i, oRecordSet.Fields.Item(5).Value.ToString().Trim());
					oDS_PS_QM610L.SetValue("U_InspMax", i, oRecordSet.Fields.Item(6).Value.ToString().Trim());

					//원소재성분이 입력[PS_QM030]된 Data Select
					if (oRecordSet.Fields.Item(3).Value.ToString().Trim() == "Cu" && oRecordSet.Fields.Item(4).Value.ToString().Trim() != "R")
					{
						oDS_PS_QM610L.SetValue("U_ValMax", i, oRecordSet01.Fields.Item(0).Value.ToString().Trim());
					}
					if (oRecordSet.Fields.Item(3).Value.ToString().Trim() == "Fe")
					{
						oDS_PS_QM610L.SetValue("U_ValMax", i, oRecordSet01.Fields.Item(1).Value.ToString().Trim());
					}
					if (oRecordSet.Fields.Item(3).Value.ToString().Trim() == "P")
					{
						oDS_PS_QM610L.SetValue("U_ValMax", i, oRecordSet01.Fields.Item(2).Value.ToString().Trim());
					}

					//기존저장된 Data select
					if (oRecordSet.Fields.Item(3).Value.ToString().Trim() == "D" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "F˚" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "Edge burr" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "Cross bow" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "R" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "T.S (kgf/㎟)" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "T.S (N/㎟)" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "E.L (%)" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "V.H (Hv)" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "S.R (㎛) Rmax" ||
						oRecordSet.Fields.Item(3).Value.ToString().Trim() == "S.R (㎛) Rz")
					{
						sQry = "select TOP 1 b.U_ValMin, b.U_ValMax ";
						sQry += "From [@PS_QM610H] a INNER JOIN[@PS_QM610L] b ON a.DocEntry = b.DocEntry ";
						sQry += "where a.U_ItemCode = '" + oDS_PS_QM610H.GetValue("U_ItemCode", 0).ToString().Trim() + "' ";
						sQry += "and a.U_LotNo = '" + oDS_PS_QM610H.GetValue("U_LotNo", 0).ToString().Trim() + "' ";
						sQry += "and b.U_ItNmDscr = '" + oRecordSet.Fields.Item(3).Value.ToString().Trim() + "' ";
						oRecordSet02.DoQuery(sQry);

						oDS_PS_QM610L.SetValue("U_ValMin", i, oRecordSet02.Fields.Item(0).Value.ToString().Trim());
						oDS_PS_QM610L.SetValue("U_ValMax", i, oRecordSet02.Fields.Item(1).Value.ToString().Trim());
					}

					//기본값 SET
					if (oRecordSet.Fields.Item(3).Value.ToString().Trim() == "E.C (%)")
					{
						oDS_PS_QM610L.SetValue("U_ValMin", i, "85");
						oDS_PS_QM610L.SetValue("U_ValMax", i, "85");
					}

					i += 1;
					oRecordSet.MoveNext();
				}

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
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
				//    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_QM610_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}

							PS_QM610_FormItemEnabled();
							PS_QM610_Initial_Setting();

							oDS_PS_QM610H.SetValue("U_CardCode", 0, Last_CardCode);
							oDS_PS_QM610H.SetValue("U_CardName", 0, Last_CardName);
							oDS_PS_QM610H.SetValue("U_InspPrsn", 0, Last_InspPrsn);
							oDS_PS_QM610H.SetValue("U_PrsnName", 0, Last_PrsnName);
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM610_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
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
								PS_QM610_FormItemEnabled();
								PS_QM610_AddMatrixRow(0, true);
								PS_QM610_Initial_Setting();

								oDS_PS_QM610H.SetValue("U_CardCode", 0, Last_CardCode);
								oDS_PS_QM610H.SetValue("U_CardName", 0, Last_CardName);
								oDS_PS_QM610H.SetValue("U_InspPrsn", 0, Last_InspPrsn);
								oDS_PS_QM610H.SetValue("U_PrsnName", 0, Last_PrsnName);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_QM610_FormItemEnabled();
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
						if (pVal.ItemUID == "InspPrsn")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("InspPrsn").Specific.Value.ToString().Trim()))
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
						else if (pVal.ItemUID == "OrdNum")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim()))
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
			string SPEC;
			decimal SPEC_MIN;
			decimal SPEC_MAX;
			decimal VAL_MIN;
			decimal VAL_MAX;
			string CardCode;
			string ItemCode;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						//검사치수 CHECK
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "ValMin")
							{
								SPEC = oMat.Columns.Item("InspSpec").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								VAL_MIN = Convert.ToDecimal(oMat.Columns.Item("ValMin").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								SPEC_MIN = Convert.ToDecimal(oMat.Columns.Item("InspMin").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								SPEC_MAX = Convert.ToDecimal(oMat.Columns.Item("InspMax").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());

								if (SPEC == "MAX")
								{
									oMat.FlushToDataSource();
									oDS_PS_QM610L.SetValue("U_ValMin", pVal.Row - 1, "0");
									oMat.LoadFromDataSource();
									errMessage = "이 항목은 MAX에 입력 하십시요.";
									oForm.Items.Item("ValMax").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
									throw new Exception();
								}
								else
								{
									if (VAL_MIN < SPEC_MIN || VAL_MIN > SPEC_MAX)
									{
										oMat.FlushToDataSource();
										oDS_PS_QM610L.SetValue("U_ValMin", pVal.Row - 1, "0");
										oMat.LoadFromDataSource();

										errMessage = "검사치수와 검사규격을 확인하여 주십시오.";
										throw new Exception();
									}
								}
							}

							if (pVal.ColUID == "ValMax")
							{
								SPEC = oMat.Columns.Item("InspSpec").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								VAL_MAX = Convert.ToDecimal(oMat.Columns.Item("ValMax").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								SPEC_MIN = Convert.ToDecimal(oMat.Columns.Item("InspMin").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								SPEC_MAX = Convert.ToDecimal(oMat.Columns.Item("InspMax").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());

								if (VAL_MAX < SPEC_MIN || VAL_MAX > SPEC_MAX)
								{
									oMat.FlushToDataSource();
									oDS_PS_QM610L.SetValue("U_ValMin", pVal.Row - 1, "0");
									oMat.LoadFromDataSource();
									errMessage = "검사치수와 검사규격을 확인하여 주십시오.";
									throw new Exception();
								}
							}
						}
						else if (pVal.ItemUID == "OrdNum") //작지번호
						{
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								sQry = "select * from [@PS_QM610H] where canceled ='N' and U_OrdNum = '" + oDS_PS_QM610H.GetValue("U_OrdNum", 0).ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								if (oRecordSet.RecordCount != 0)
								{
									oDS_PS_QM610H.SetValue("U_OrdNum", 0, "");
									errMessage = "이미 동일한 작지번호가 존재합니다. 확인하여 주십시오.";
									throw new Exception();
								}
							}

							sQry = "select a.U_ItemCode, a.U_ItemName, b.U_BatchNum, c.FrgnName, c.U_Spec1 ";
							sQry += " from [@PS_PP030H] a INNER JOIN  [@PS_PP030L] b on a.docentry = b.docentry ";
							sQry += "                     LEFT  JOIN  [OITM] c on c.ItemCode = a.U_ItemCode ";
							sQry += " where U_OrdNum = '" + oDS_PS_QM610H.GetValue("U_OrdNum", 0).ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oDS_PS_QM610H.SetValue("U_ItemCode", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_QM610H.SetValue("U_ItemName", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_QM610H.SetValue("U_LotNo", 0, codeHelpClass.Left(oRecordSet.Fields.Item(2).Value.ToString().Trim(), 8));
							oDS_PS_QM610H.SetValue("U_FrgnName", 0, oRecordSet.Fields.Item(3).Value.ToString().Trim());
							oDS_PS_QM610H.SetValue("U_Size", 0, oRecordSet.Fields.Item(4).Value.ToString().Trim());

							CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
							ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

							sQry = "Select Count(*), Max(U_CardSeq), MAX(U_TS_Gbn)  From [@PS_QM600H] Where U_CardCode = '" + CardCode + "' And U_ItemCode = '" + ItemCode + "'";
							oRecordSet.DoQuery(sQry);

							if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 1)
							{
								oDS_PS_QM610H.SetValue("U_TS_Gbn", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());
								oDS_PS_QM610H.SetValue("U_CardSeq", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oForm.Items.Item("CardSeq").Enabled = false;
								PS_QM610_LoadData();
							}
						}
						else if (pVal.ItemUID == "InspPrsn") //사번
						{
							oDS_PS_QM610H.SetValue("U_PrsnName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
						}
						else if (pVal.ItemUID == "ItemCode") //품목코드
						{
							sQry = "Select ItemName, FrgnName, U_Size From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oDS_PS_QM610H.SetValue("U_ItemName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_QM610H.SetValue("U_FrgnName", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_QM610H.SetValue("U_Size", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());

							CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
							ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

							sQry = "Select Count(*), Max(U_CardSeq) From [@PS_QM007H] Where U_CardCode = '" + CardCode + "' And U_ItemCode = '" + ItemCode + "'";
							oRecordSet.DoQuery(sQry);

							if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 1)
							{
								oDS_PS_QM610H.SetValue("U_CardSeq", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oForm.Items.Item("CardSeq").Enabled = false;
								PS_QM610_LoadData();
							}
							else
							{
								oForm.Items.Item("CardSeq").Enabled = true;
								oDS_PS_QM610H.SetValue("U_CardSeq", 0, "");
							}
						}
						else if (pVal.ItemUID == "CardCode") //거래처코드
						{
							sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" + oDS_PS_QM610H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oDS_PS_QM610H.SetValue("U_CardName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());

							CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
							ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

							sQry = "Select Count(*), Max(U_CardSeq), MAX(U_TS_Gbn)  From [@PS_QM600H] Where U_CardCode = '" + CardCode + "' And U_ItemCode = '" + ItemCode + "'";
							oRecordSet.DoQuery(sQry);

							if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 1)
							{
								oDS_PS_QM610H.SetValue("U_TS_Gbn", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());
								oDS_PS_QM610H.SetValue("U_CardSeq", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oForm.Items.Item("CardSeq").Enabled = false;
								PS_QM610_LoadData();
							}
							else
							{
								oDS_PS_QM610H.SetValue("U_TS_Gbn", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());
								oForm.Items.Item("CardSeq").Enabled = true;
								oDS_PS_QM610H.SetValue("U_CardSeq", 0, "");
							}
						}
						else if (pVal.ItemUID == "CardSeq") //거래처 순번
						{
							if (!string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()) &&
								!string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()) &&
								!string.IsNullOrEmpty(oForm.Items.Item("CardSeq").Specific.Value.ToString().Trim()))
							{
								PS_QM610_LoadData();
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
					PS_QM610_FormItemEnabled();
					PS_QM610_AddMatrixRow(oMat.VisualRowCount, false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM610H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM610L);
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
						case "1283":
							//제거
							if (PSH_Globals.SBO_Application.MessageBox("문서를 제거(삭제) 하시겠습니까?", 1, "예", "아니오") == 1)
							{
							}
							else
							{
								BubbleEvent = false;
								return;
							}
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
							PSH_Globals.SBO_Application.MessageBox("행삭제를 할수 없습니다. ");
							BubbleEvent = false;
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
							break;
						case "1281": //찾기
							PS_QM610_FormItemEnabled();
							oForm.Items.Item("OrdNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_QM610_FormItemEnabled();
							PS_QM610_AddMatrixRow(0, true);
							PS_QM610_Initial_Setting();
							oDS_PS_QM610H.SetValue("U_CardCode", 0, Last_CardCode);
							oDS_PS_QM610H.SetValue("U_CardName", 0, Last_CardName);
							oDS_PS_QM610H.SetValue("U_InspPrsn", 0, Last_InspPrsn);
							oDS_PS_QM610H.SetValue("U_PrsnName", 0, Last_PrsnName);
							break;
						case "1287": //복제
							PS_QM610_CopyMatrixRow();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM610_FormItemEnabled();
							oForm.EnableMenu("1282", true);
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
