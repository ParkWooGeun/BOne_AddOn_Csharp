using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 품질검사등록
	/// </summary>
	internal class PS_QM410 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_QM410H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM410L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM410.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM410_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM410");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_QM410_CreateItems();
				PS_QM410_ComboBox_Setting();
				PS_QM410_EnableMenus();
				PS_QM410_SetDocument(oFormDocEntry);
				oForm.Items.Item("DrawNo").Click();

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
		/// PS_QM410_CreateItems
		/// </summary>
		private void PS_QM410_CreateItems()
		{
			try
			{
				oDS_PS_QM410H = oForm.DataSources.DBDataSources.Item("@PS_QM410H");
				oDS_PS_QM410L = oForm.DataSources.DBDataSources.Item("@PS_QM410L");
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
		/// PS_QM410_ComboBox_Setting
		/// </summary>
		private void PS_QM410_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//자체/수입
				oForm.Items.Item("Class").Specific.ValidValues.Add("%", "선택");
				oForm.Items.Item("Class").Specific.ValidValues.Add("1", "수입품검사");
				oForm.Items.Item("Class").Specific.ValidValues.Add("2", "자체품검사");
				oForm.Items.Item("Class").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//검사종류
				sQry = "     SELECT  U_Minor AS [Code],";
				sQry += "             U_CdName AS [Name]";
				sQry += "  FROM   [@PS_SY001L]";
				sQry += "  WHERE  Code = 'Q014'";
				oForm.Items.Item("ChkCls").Specific.ValidValues.Add("%", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ChkCls").Specific, sQry, "", false, false);
				oForm.Items.Item("ChkCls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//매트릭스_검사항목
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName + '[' + ISNULL(U_RelCd,'') + ']' AS [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q011'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("ChkPnt"), sQry, "", "");

				//매트릭스_부호(최소)
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName AS [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q013'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("MinMark"), sQry, "", "");

				//매트릭스_부호(최대)
				sQry = "     SELECT      U_Minor AS [Code],";
				sQry += "                 U_CdName AS [Name]";
				sQry += "  FROM       [@PS_SY001L]";
				sQry += "  WHERE      Code = 'Q013'";
				sQry += "                 AND U_UseYN = 'Y'";
				sQry += "  ORDER BY  U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("MaxMark"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM410_EnableMenus
		/// </summary>
		private void PS_QM410_EnableMenus()
		{
			try
			{
				////메뉴활성화
				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", true);  // 닫기
				oForm.EnableMenu("1287", true);  // 복제
				oForm.EnableMenu("1285", true);  // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM410_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_QM410_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_QM410_FormItemEnabled();
					PS_QM410_AddMatrixRow(0, true);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM410_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_QM410_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string StdPnt; //기준치수
			string MinPnt; //공차(최소)
			string MaxPnt; //공차(최대)
			string InputData; //입력한 측정값
			string MinMark;
			string MaxMark;
			string ChkPnt; //검사항목
			string ChkSeq;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				switch (oUID)
				{
					case "Mat01":
						oMat.FlushToDataSource();
						if (oMat.RowCount == oRow && !string.IsNullOrEmpty(oDS_PS_QM410L.GetValue("U_ChkPnt", oRow - 1).ToString().Trim()))
						{
							PS_QM410_AddMatrixRow(oRow, false);
						}

						StdPnt = oDS_PS_QM410L.GetValue("U_StdPnt", oRow - 1).ToString().Trim();
						MinMark = oDS_PS_QM410L.GetValue("U_MinMark", oRow - 1).ToString().Trim();
						MinPnt = oDS_PS_QM410L.GetValue("U_MinPnt", oRow - 1).ToString().Trim();
						MaxMark = oDS_PS_QM410L.GetValue("U_MaxMark", oRow - 1).ToString().Trim();
						MaxPnt = oDS_PS_QM410L.GetValue("U_MaxPnt", oRow - 1).ToString().Trim();
						InputData = oDS_PS_QM410L.GetValue("U_" + oCol, oRow - 1).ToString().Trim();
						ChkPnt = oDS_PS_QM410L.GetValue("U_ChkPnt", oRow - 1).ToString().Trim();

						if (oCol == "Data01" || oCol == "Data02" || oCol == "Data03" || oCol == "Data04" || oCol == "Data05" || oCol == "Data06" || oCol == "Data07" || oCol == "Data08" || oCol == "Data09" || oCol == "Data10" || oCol == "Data11" || oCol == "Data12" || oCol == "Data13" || oCol == "Data14" || oCol == "Data15")
						{
							//입력 값이 존재하고 제목(또는 빈행)이 아닐 때만 체크
							if (!string.IsNullOrEmpty(InputData) && Convert.ToDouble(ChkPnt) < 90)
							{
								if (PS_QM410_CheckInputData(StdPnt, MinPnt, MaxPnt, InputData, MinMark, MaxMark) == false)
								{
									if (PSH_Globals.SBO_Application.MessageBox("입력한 측정치는 불량입니다. 등록하시겠습니까?", 1, "예", "아니오") == 1)
									{
									}
									else
									{
										oMat.Columns.Item(oCol).Cells.Item(oRow).Specific.Value = ""; //빈값 세팅
									}
								}
							}
						}
						oMat.AutoResizeColumns();
						break;
					case "CardCode":
						oDS_PS_QM410H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oDS_PS_QM410H.GetValue("U_CardCode", 0).ToString().Trim() + "'", "")); //납품처
						break;
					case "CntcCode":
						oDS_PS_QM410H.SetValue("U_CntcName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oDS_PS_QM410H.GetValue("U_CntcCode", 0).ToString().Trim() + "'", "")); //검사자성명
						break;
					case "ItemCode":
						oDS_PS_QM410H.SetValue("U_ItemName", 0, dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oDS_PS_QM410H.GetValue("U_ItemCode", 0).ToString().Trim() + "'", ""));	//작번
						//검사순번 조회
						sQry = "    SELECT      ISNULL(MAX(U_ChkSeq), 0) AS [ChkSeq]";
						sQry += " FROM       [@PS_QM410H]";
						sQry += " WHERE     U_ItemCode = '" + oDS_PS_QM410H.GetValue("U_ItemCode", 0).ToString().Trim() + "'";
						sQry += "                AND Status = 'O'";
						oRecordSet.DoQuery(sQry);

						if (Convert.ToInt32(oRecordSet.Fields.Item("ChkSeq").Value.ToString().Trim()) == 0)
						{
							ChkSeq = "1";
						}
						else
						{
							ChkSeq = Convert.ToString(Convert.ToInt32(oRecordSet.Fields.Item("ChkSeq").Value.ToString().Trim()) + 1);
						}
						oDS_PS_QM410H.SetValue("U_ChkSeq", 0, ChkSeq); 

						//납품처 자동등록
						sQry = "    SELECT       T0.CardCode,";
						sQry += "                 T0.CardName";
						sQry += " FROM        ORDR AS T0";
						sQry += "                 INNER JOIN";
						sQry += "                 RDR1 AS T1";
						sQry += "                     ON T0.DocEntry = T1.DocEntry";
						sQry += " WHERE       T0.Canceled = 'N'";
						sQry += "                 AND T1.ItemCode = '" + oDS_PS_QM410H.GetValue("U_ItemCode", 0).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oDS_PS_QM410H.SetValue("U_CardCode", 0, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim()); //거래처코드
						oDS_PS_QM410H.SetValue("U_CardName", 0, oRecordSet.Fields.Item("CardName").Value.ToString().Trim()); //거래처명
						break;
					case "TotalQty":
						PS_QM410_CalculateQty(Convert.ToDouble(oDS_PS_QM410H.GetValue("U_TotalQty", 0).ToString().Trim()), Convert.ToDouble(oDS_PS_QM410H.GetValue("U_PassQty", 0).ToString().Trim()));
						break;
					case "PassQty":
						PS_QM410_CalculateQty(Convert.ToDouble(oDS_PS_QM410H.GetValue("U_TotalQty", 0).ToString().Trim()), Convert.ToDouble(oDS_PS_QM410H.GetValue("U_PassQty", 0).ToString().Trim()));
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
		/// PS_QM410_FormItemEnabled
		/// </summary>
		private void PS_QM410_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					PS_QM410_FormClear();
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("DrawNo").Click();
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocEntry").Specific.Value = "";
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가
					oForm.Items.Item("DrawNo").Click();
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
				}
				oMat.AutoResizeColumns();
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
		/// PS_QM410_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM410_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_QM410L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM410L.Offset = oRow;
				oDS_PS_QM410L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oDS_PS_QM410L.SetValue("U_Check", oRow, "Y");
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
		/// /PS_QM410_CopyMatrixRow
		/// </summary>
		private void PS_QM410_CopyMatrixRow()
		{
			int i;
			string DocEntry;
			string ChkSeq;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM410'", "");
				ChkSeq = dataHelpClass.Get_ReData("MAX(U_ChkSeq)", "U_ItemCode", "[@PS_QM410H]", "'" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'", "AND Status = 'O'");
				oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
				oForm.Items.Item("ChkSeq").Specific.Value = Convert.ToString(Convert.ToDouble(ChkSeq) + 1);

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_QM410H.SetValue("DocEntry", i, DocEntry);
					oMat.LoadFromDataSource();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM410_FormClear
		/// </summary>
		private void PS_QM410_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM410'", "");

				if (string.IsNullOrEmpty(DocEntry) | DocEntry == "0")
				{
					oForm.Items.Item("DocEntry").Specific.Value = "1";
				}
				else
				{
					oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
				}

				oDS_PS_QM410H.SetValue("U_ChkDate", 0, DateTime.Now.ToString("yyyyMMdd"));
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM410_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_QM410_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try {
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM410_FormClear();
				}

				if (Convert.ToDouble(oDS_PS_QM410H.GetValue("U_TotalQty", 0).ToString().Trim()) < Convert.ToDouble(oDS_PS_QM410H.GetValue("U_PassQty", 0).ToString().Trim()))
				{
					errMessage = "합격수량이 전체수량보다 많습니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DrawNo").Specific.Value.ToString().Trim()))
				{
					errMessage = "도면번호가 입력되지 않았습니다.";
					throw new Exception();
				}
				if (oForm.Items.Item("Class").Specific.Value.ToString().Trim() == "%")
				{
					errMessage = "자체/수입이 선택되지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "납품처가 입력되지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ChkDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "검사일자가 입력되지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("TotalQty").Specific.Value.ToString().Trim()))
				{
					errMessage = "전체수량이 입력되지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("PassQty").Specific.Value.ToString().Trim()))
				{
					errMessage = "합격수량이 입력되지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("BadQty").Specific.Value.ToString().Trim()))
				{
					errMessage = "불합격수량이 입력되지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "검사자정보가 입력되지 않았습니다.";
					throw new Exception();
				}
				//수입품검사일 경우
				if (oForm.Items.Item("Class").Specific.Value.ToString().Trim() == "1")
				{
					//가입고번호가 없으면
					if (string.IsNullOrEmpty(oForm.Items.Item("MM050HNo").Specific.Value.ToString().Trim()))
					{
						//선검사사유 필수
						if (string.IsNullOrEmpty(oForm.Items.Item("PreChkNt").Specific.Value.ToString().Trim()))
						{
							oForm.Items.Item("PreChkNt").Click();
							errMessage = "[가입고번호]를 입력하지 않을 경우 [선검사사유]는 필수입니다.";
							throw new Exception();
						}
					}
				}
				if (oForm.Items.Item("ChkCls").Specific.Value.ToString().Trim() == "%")
				{
					errMessage = "검사종류가 선택되지 않았습니다.";
					throw new Exception();
				}
				if (oMat.VisualRowCount == 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if (Convert.ToDouble(oMat.Columns.Item("ChkPnt").Cells.Item(i).Specific.Selected.Value.ToString().Trim()) < 90)
					{
						if (string.IsNullOrEmpty(oMat.Columns.Item("ChkPnt").Cells.Item(i).Specific.Selected.Value.ToString().Trim()))
						{
							oMat.Columns.Item("ChkPnt").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							errMessage = "검사항목은 필수입니다.";
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oMat.Columns.Item("StdPnt").Cells.Item(i).Specific.Value.ToString().Trim()))
						{
							oMat.Columns.Item("StdPnt").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							errMessage = "기준치수는 필수입니다.";
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oMat.Columns.Item("MinPnt").Cells.Item(i).Specific.Value.ToString().Trim()))
						{
							oMat.Columns.Item("MinPnt").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							errMessage = "최소공차는 필수입니다.";
							throw new Exception();
						}
						if (string.IsNullOrEmpty(oMat.Columns.Item("MaxPnt").Cells.Item(i).Specific.Value.ToString().Trim()))
						{
							oMat.Columns.Item("MaxPnt").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							errMessage = "최대공차는 필수입니다.";
							throw new Exception();
						}
					}
				}

				oDS_PS_QM410L.RemoveRecord(oDS_PS_QM410L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM410_FormClear();
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
		private void PS_QM410_MTX01()
		{
			int i;
			string DrawNo;
			string ReviNo;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				DrawNo = oForm.Items.Item("DrawNo").Specific.Value.ToString().Trim();
				ReviNo = oForm.Items.Item("ReviNo").Specific.Value.ToString().Trim();

				sQry = "       EXEC PS_QM410_01 '";
				sQry += DrawNo + "', '";
				sQry += ReviNo + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "검사기준이 존재하지 않습니다.";
					throw new Exception();
				}

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i != 0)
					{
						oDS_PS_QM410L.InsertRecord(i);
					}
					oDS_PS_QM410L.Offset = i;
					oDS_PS_QM410L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM410L.SetValue("U_Check", i, "Y");                  //선택
					oDS_PS_QM410L.SetValue("U_ChkPnt", i, oRecordSet.Fields.Item("ChkPnt").Value.ToString().Trim());                    //검사항목
					oDS_PS_QM410L.SetValue("U_StdPnt", i, oRecordSet.Fields.Item("StdPnt").Value.ToString().Trim());                    //기준치수
					oDS_PS_QM410L.SetValue("U_MinMark", i, oRecordSet.Fields.Item("MinMark").Value.ToString().Trim());                  //부호(최소)
					oDS_PS_QM410L.SetValue("U_MinPnt", i, oRecordSet.Fields.Item("MinPnt").Value.ToString().Trim());                    //공차(최소)
					oDS_PS_QM410L.SetValue("U_MaxMark", i, oRecordSet.Fields.Item("MaxMark").Value.ToString().Trim());                  //부호(최대)
					oDS_PS_QM410L.SetValue("U_MaxPnt", i, oRecordSet.Fields.Item("MaxPnt").Value.ToString().Trim());                    //공차(최대)
					oDS_PS_QM410L.SetValue("U_Data01", i, oRecordSet.Fields.Item("Data01").Value.ToString().Trim());                    //Data01
					oDS_PS_QM410L.SetValue("U_Data02", i, oRecordSet.Fields.Item("Data02").Value.ToString().Trim());                    //Data02
					oDS_PS_QM410L.SetValue("U_Data03", i, oRecordSet.Fields.Item("Data03").Value.ToString().Trim());                    //Data03
					oDS_PS_QM410L.SetValue("U_Data04", i, oRecordSet.Fields.Item("Data04").Value.ToString().Trim());                    //Data04
					oDS_PS_QM410L.SetValue("U_Data05", i, oRecordSet.Fields.Item("Data05").Value.ToString().Trim());                    //Data05
					oDS_PS_QM410L.SetValue("U_Data06", i, oRecordSet.Fields.Item("Data06").Value.ToString().Trim());                    //Data06
					oDS_PS_QM410L.SetValue("U_Data07", i, oRecordSet.Fields.Item("Data07").Value.ToString().Trim());                    //Data07
					oDS_PS_QM410L.SetValue("U_Data08", i, oRecordSet.Fields.Item("Data08").Value.ToString().Trim());                    //Data08
					oDS_PS_QM410L.SetValue("U_Data09", i, oRecordSet.Fields.Item("Data09").Value.ToString().Trim());                    //Data09
					oDS_PS_QM410L.SetValue("U_Data10", i, oRecordSet.Fields.Item("Data10").Value.ToString().Trim());                    //Data10
					oDS_PS_QM410L.SetValue("U_Data11", i, oRecordSet.Fields.Item("Data11").Value.ToString().Trim());                    //Data11
					oDS_PS_QM410L.SetValue("U_Data12", i, oRecordSet.Fields.Item("Data12").Value.ToString().Trim());                    //Data12
					oDS_PS_QM410L.SetValue("U_Data13", i, oRecordSet.Fields.Item("Data13").Value.ToString().Trim());                    //Data13
					oDS_PS_QM410L.SetValue("U_Data14", i, oRecordSet.Fields.Item("Data14").Value.ToString().Trim());                    //Data14
					oDS_PS_QM410L.SetValue("U_Data15", i, oRecordSet.Fields.Item("Data15").Value.ToString().Trim());                    //Data15
					oRecordSet.MoveNext();

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				oForm.Update();
				PS_QM410_AddMatrixRow(oMat.VisualRowCount, false);
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
		/// 불량수량 계산
		/// </summary>
		/// <param name="pTotalQty"></param>
		/// <param name="pPassQty"></param>
		private void PS_QM410_CalculateQty(double pTotalQty, double pPassQty)
		{
			string errMessage = string.Empty;

			try
			{
				if (pTotalQty < pPassQty)
				{
					errMessage = "합격수량이 전체수량보다 큽니다. 확인하세요.";
					throw new Exception();
				}
				oDS_PS_QM410H.SetValue("U_BadQty", 0, Convert.ToString(pTotalQty - pPassQty));
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
		/// 입력한 측정치수의 합격불합격 여부 체크
		/// </summary>
		/// <param name="pStdPnt"></param>
		/// <param name="pMinPnt"></param>
		/// <param name="pMaxPnt"></param>
		/// <param name="pInputData"></param>
		/// <param name="pMinMark"></param>
		/// <param name="pMaxMark"></param>
		/// <returns></returns>
		private bool PS_QM410_CheckInputData(string pStdPnt, string pMinPnt, string pMaxPnt, string pInputData, string pMinMark, string pMaxMark)
		{
			bool ReturnValue = false;
			string MinData; //최소기준치
			string MaxData; //최대기준치

			try
			{
				//최소기준치 계산
				if (pMinMark == "03")
				{
					MinData = Convert.ToString(Convert.ToDouble(pStdPnt) + Convert.ToDouble(pMinPnt));
				}
				else if (pMinMark == "02")
				{
					MinData = Convert.ToString(Convert.ToDouble(pStdPnt) - Convert.ToDouble(pMinPnt));
				}
				else
				{
					MinData = pStdPnt;
				}

				//최대기준치 계산
				if (pMaxMark == "03")
				{
					MaxData = Convert.ToString(Convert.ToDouble(pStdPnt) + Convert.ToDouble(pMaxPnt));
				}
				else if (pMaxMark == "02")
				{
					MaxData = Convert.ToString(Convert.ToDouble(pStdPnt) - Convert.ToDouble(pMaxPnt));
				}
				else
				{
					MaxData = pStdPnt;
				}

				if (Convert.ToDouble(pInputData) < Convert.ToDouble(MinData) || Convert.ToDouble(pInputData) > Convert.ToDouble(MaxData))
				{
					ReturnValue = false;
				}
				else
				{
					ReturnValue = true;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			return ReturnValue;
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
				//    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_QM410_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM410_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
					}
					else if (pVal.ItemUID == "BtnOpen")
					{
						if (string.IsNullOrEmpty(oForm.Items.Item("DrawNo").Specific.Value.ToString().Trim()))
						{
							errMessage = "도면번호를 입력하십시오.";
							throw new Exception();
						}

						PS_QM410_MTX01();
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
								PS_QM410_FormItemEnabled();
								PS_QM410_AddMatrixRow(0, true);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_QM410_FormItemEnabled();
							}
						}
					}
					else if (pVal.ItemUID == "Mat01")
					{
						if (pVal.ColUID == "Check")
						{
							oMat.FlushToDataSource();
						}
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
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.ColUID == "BatchNum")
						{
							dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "BatchNum");
						}
					}
					else
					{
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "DrawNo", "");
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "MM050HNo", "");
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
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
						PS_QM410_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
		/// VALIDATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemChanged == true)
					{
						PS_QM410_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						if (pVal.ItemUID == "Mat01")
						{
							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
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
					PS_QM410_FormItemEnabled();
					PS_QM410_AddMatrixRow(oMat.VisualRowCount, false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM410H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM410L);
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
						oDS_PS_QM410L.RemoveRecord(oDS_PS_QM410L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_QM410_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_QM410L.GetValue("U_ChkPnt", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_QM410_AddMatrixRow(oMat.RowCount, false);
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
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
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
							PS_QM410_FormItemEnabled();
							oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_QM410_FormItemEnabled();
							PS_QM410_AddMatrixRow(0, true);
							break;
						case "1287": //복제
							PS_QM410_CopyMatrixRow();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM410_FormItemEnabled();
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
