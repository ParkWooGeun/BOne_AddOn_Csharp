using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 일일 생산계획 등록
	/// </summary>
	internal class PS_PP930 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP930H;

		private string oLastItemUID01;
		private string oLastColUID01;
		private int oLastColRow01;
		private int oLast_Mode;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP930.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}
				oFormUniqueID = "PS_PP930_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP930");

				string strXml = null;
				strXml = oXmlDoc.xml.ToString();

				PSH_Globals.SBO_Application.LoadBatchActions(strXml);
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP930_CreateItems();
				PS_PP930_ComboBox_Setting();
				Add_MatrixRow(0, true);

				oForm.EnableMenu(("1283"), false); // 삭제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1287"), false); // 복제
				oForm.EnableMenu(("1285"), false); // 복원
				oForm.EnableMenu(("1284"), true);  // 취소
				oForm.EnableMenu(("1293"), true);  // 행삭제
				oForm.EnableMenu(("1281"), false);
				oForm.EnableMenu(("1282"), true);
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
		/// PS_PP930_CreateItems
		/// </summary>
		private void PS_PP930_CreateItems()
		{
			try
			{
				oDS_PS_PP930H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();
				oForm.Items.Item("YYYYMM").Specific.VALUE = DateTime.Now.ToString("yyyyMM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP930_ComboBox_Setting
		/// </summary>
		private void PS_PP930_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "Select Code, Name From [@PSH_ITMBSORT] Where U_PudYN = 'Y'", "102", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP930H.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP930H.Offset = oRow;
				oDS_PS_PP930H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// LoadData
		/// </summary>
		private void LoadData()
		{
			int i;
			string sQry;

			string YYYYMM;
			string ItmBsort;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim();
				ItmBsort = oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim();

				oForm.Freeze(true);

				for (i = 1; i <= 31; i++)
				{
					oForm.Items.Item("D" + i).Specific.Value = "0";
					oForm.Items.Item("N" + i).Specific.Value = "0";
				}

				sQry = "select dd = right(Convert(char(8),U_DocDate,112),2), D = max(U_DTime), N = max(U_NTime) from [@PS_PP930H] Where Convert(Char(6),U_DocDate,112) = '" + YYYYMM + "' And U_ItmBsort = '" + ItmBsort + "' group by right(Convert(char(8),U_DocDate,112),2)";
				oRecordSet.DoQuery(sQry);

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					oForm.Items.Item("D" + Convert.ToString(i + 1)).Specific.Value = oRecordSet.Fields.Item(1).Value.ToString().Trim();
					oForm.Items.Item("N" + Convert.ToString(i + 1)).Specific.Value = oRecordSet.Fields.Item(2).Value.ToString().Trim();
					oRecordSet.MoveNext();
				}

				sQry = "EXEC [PS_PP930_01] '" + YYYYMM + "','" + ItmBsort + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP930H.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP930H.Size)
					{
						oDS_PS_PP930H.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP930H.Offset = i;

					oDS_PS_PP930H.SetValue("U_ColRgl01", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColRgl02", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColRgl03", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColRgl04", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColSum01", i, oRecordSet.Fields.Item(4).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColSum02", i, oRecordSet.Fields.Item(5).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColSum03", i, oRecordSet.Fields.Item(6).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg01", i, oRecordSet.Fields.Item(7).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg02", i, oRecordSet.Fields.Item(8).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg03", i, oRecordSet.Fields.Item(9).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg04", i, oRecordSet.Fields.Item(10).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg05", i, oRecordSet.Fields.Item(11).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg06", i, oRecordSet.Fields.Item(12).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg07", i, oRecordSet.Fields.Item(13).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg08", i, oRecordSet.Fields.Item(14).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg09", i, oRecordSet.Fields.Item(15).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg10", i, oRecordSet.Fields.Item(16).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg11", i, oRecordSet.Fields.Item(17).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg12", i, oRecordSet.Fields.Item(18).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg13", i, oRecordSet.Fields.Item(19).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg14", i, oRecordSet.Fields.Item(20).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg15", i, oRecordSet.Fields.Item(21).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg16", i, oRecordSet.Fields.Item(22).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg17", i, oRecordSet.Fields.Item(23).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg18", i, oRecordSet.Fields.Item(24).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg19", i, oRecordSet.Fields.Item(25).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg20", i, oRecordSet.Fields.Item(26).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg21", i, oRecordSet.Fields.Item(27).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg22", i, oRecordSet.Fields.Item(28).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg23", i, oRecordSet.Fields.Item(29).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg24", i, oRecordSet.Fields.Item(30).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg25", i, oRecordSet.Fields.Item(31).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg26", i, oRecordSet.Fields.Item(32).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg27", i, oRecordSet.Fields.Item(33).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg28", i, oRecordSet.Fields.Item(34).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg29", i, oRecordSet.Fields.Item(35).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg30", i, oRecordSet.Fields.Item(36).Value.ToString().Trim());
					oDS_PS_PP930H.SetValue("U_ColReg31", i, oRecordSet.Fields.Item(37).Value.ToString().Trim());

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// CalcData  일일생산량 계산
		/// </summary>
		private void CalcData()
		{
			int i;
			int Lastdd;
			string errMessage = string.Empty;

			Double[] D = new Double[31];
			Double[] N = new Double[31];

			string NInwon;
			string DInwon;
			string MkCapa;

			string YYYYMM;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim() + "01";

				sQry = "Select right(Convert(Char(8),DateAdd(dd, -1, DateAdd(mm, 1,'" + YYYYMM + "')),112),2)";
				oRecordSet.DoQuery(sQry);
				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					errMessage = "날짜 계산이 틀립니다. 확인하세요.:";
					throw new Exception();
				}
				else
				{
					Lastdd = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim());
				}

				oForm.Freeze(true);

				for (i = 1; i <= Lastdd; i++)
				{
					D[i - 1] = Convert.ToDouble(oForm.Items.Item("D" + i).Specific.Value.ToString().Trim());
					N[i - 1] = Convert.ToDouble(oForm.Items.Item("N" + i).Specific.Value.ToString().Trim());
				}

				for (i = 1; i <= oMat.RowCount; i++)
				{
					DInwon = oMat.Columns.Item("DInwon").Cells.Item(i).Specific.Value.ToString().Trim();
					NInwon = oMat.Columns.Item("NInwon").Cells.Item(i).Specific.Value.ToString().Trim();
					MkCapa = oMat.Columns.Item("MkCapa").Cells.Item(i).Specific.Value.ToString().Trim();

					oDS_PS_PP930H.SetValue("U_ColSum01", i - 1, DInwon);
					oDS_PS_PP930H.SetValue("U_ColSum02", i - 1, NInwon);
					oDS_PS_PP930H.SetValue("U_ColSum03", i - 1, MkCapa);
					oDS_PS_PP930H.SetValue("U_ColReg01", i - 1, Convert.ToString((Convert.ToDouble(D[0]) + Convert.ToDouble(N[0])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg02", i - 1, Convert.ToString((Convert.ToDouble(D[1]) + Convert.ToDouble(N[1])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg03", i - 1, Convert.ToString((Convert.ToDouble(D[2]) + Convert.ToDouble(N[2])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg04", i - 1, Convert.ToString((Convert.ToDouble(D[3]) + Convert.ToDouble(N[3])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg05", i - 1, Convert.ToString((Convert.ToDouble(D[4]) + Convert.ToDouble(N[4])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg06", i - 1, Convert.ToString((Convert.ToDouble(D[5]) + Convert.ToDouble(N[5])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg07", i - 1, Convert.ToString((Convert.ToDouble(D[6]) + Convert.ToDouble(N[6])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg08", i - 1, Convert.ToString((Convert.ToDouble(D[7]) + Convert.ToDouble(N[7])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg09", i - 1, Convert.ToString((Convert.ToDouble(D[8]) + Convert.ToDouble(N[8])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg10", i - 1, Convert.ToString((Convert.ToDouble(D[9]) + Convert.ToDouble(N[9])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg11", i - 1, Convert.ToString((Convert.ToDouble(D[10]) + Convert.ToDouble(N[10])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg12", i - 1, Convert.ToString((Convert.ToDouble(D[11]) + Convert.ToDouble(N[11])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg13", i - 1, Convert.ToString((Convert.ToDouble(D[12]) + Convert.ToDouble(N[12])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg14", i - 1, Convert.ToString((Convert.ToDouble(D[13]) + Convert.ToDouble(N[13])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg15", i - 1, Convert.ToString((Convert.ToDouble(D[14]) + Convert.ToDouble(N[14])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg16", i - 1, Convert.ToString((Convert.ToDouble(D[15]) + Convert.ToDouble(N[15])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg17", i - 1, Convert.ToString((Convert.ToDouble(D[16]) + Convert.ToDouble(N[16])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg18", i - 1, Convert.ToString((Convert.ToDouble(D[17]) + Convert.ToDouble(N[17])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg19", i - 1, Convert.ToString((Convert.ToDouble(D[18]) + Convert.ToDouble(N[18])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg20", i - 1, Convert.ToString((Convert.ToDouble(D[19]) + Convert.ToDouble(N[19])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg21", i - 1, Convert.ToString((Convert.ToDouble(D[20]) + Convert.ToDouble(N[20])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg22", i - 1, Convert.ToString((Convert.ToDouble(D[21]) + Convert.ToDouble(N[21])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg23", i - 1, Convert.ToString((Convert.ToDouble(D[22]) + Convert.ToDouble(N[22])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg24", i - 1, Convert.ToString((Convert.ToDouble(D[23]) + Convert.ToDouble(N[23])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg25", i - 1, Convert.ToString((Convert.ToDouble(D[24]) + Convert.ToDouble(N[24])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg26", i - 1, Convert.ToString((Convert.ToDouble(D[25]) + Convert.ToDouble(N[25])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg27", i - 1, Convert.ToString((Convert.ToDouble(D[26]) + Convert.ToDouble(N[26])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg28", i - 1, Convert.ToString((Convert.ToDouble(D[27]) + Convert.ToDouble(N[27])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg29", i - 1, Convert.ToString((Convert.ToDouble(D[28]) + Convert.ToDouble(N[28])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg30", i - 1, Convert.ToString((Convert.ToDouble(D[29]) + Convert.ToDouble(N[29])) * Convert.ToDouble(MkCapa)));
					oDS_PS_PP930H.SetValue("U_ColReg31", i - 1, Convert.ToString((Convert.ToDouble(D[30]) + Convert.ToDouble(N[30])) * Convert.ToDouble(MkCapa)));
				}

				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oMat.Clear();
				oMat.LoadFromDataSource();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Add_PurchaseDemand
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool Add_PurchaseDemand(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool functionReturnValue = false;

			int i;
			int j;
			string sQry;
			string errMessage = string.Empty;

			string DocDate;
			string ItemCode;
			string ItmMsort;
			string BPLId;
			string YYYYMM;
			string DocEntry;
			string ItmMname;
			string ItemName;
			string ymd;
			int ItmBsort;

			string NInwon;
			string DInwon;
			string MkCapa;
			int Lastdd;

			object[] D = new object[31];
			object[] N = new object[31];
			object[] MP = new object[31];
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLId = "1";
				YYYYMM = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim();
				ymd = oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim() + "01";
				ItmBsort = Convert.ToInt32(oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim());

				//저장시 삭제후 재등록
				sQry = "Delete From [@PS_PP930H] Where Convert(Char(6),U_DocDate,112) ='" + YYYYMM + "'";
				oRecordSet.DoQuery(sQry);

				//해당월 최종일자 계산
				sQry = "Select right(Convert(Char(8),DateAdd(dd, -1, DateAdd(mm, 1,'" + ymd + "')),112),2)";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
				{
					errMessage = "날짜 계산이 틀립니다. 확인하세요.";
					throw new Exception();
				}
				else
				{
					Lastdd = Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim());
				}

				//주야간 근무시간
				for (i = 1; i <= Lastdd; i++)
				{
					D[i - 1] = oForm.Items.Item("D" + i).Specific.Value.ToString().Trim();
					N[i - 1] = oForm.Items.Item("N" + i).Specific.Value.ToString().Trim();
				}

				ProgressBar01.Text = "자장중.....";

				for (i = 1; i <= oMat.RowCount; i++)
				{
					ItmMsort = oMat.Columns.Item("ItmMsort").Cells.Item(i).Specific.Value.ToString().Trim();
					ItmMname = oMat.Columns.Item("ItmMname").Cells.Item(i).Specific.Value.ToString().Trim();
					ItemCode = oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
					ItemName = oMat.Columns.Item("ItemName").Cells.Item(i).Specific.Value.ToString().Trim();
					DInwon = oMat.Columns.Item("DInwon").Cells.Item(i).Specific.Value.ToString().Trim();
					NInwon = oMat.Columns.Item("NInwon").Cells.Item(i).Specific.Value.ToString().Trim();
					MkCapa = oMat.Columns.Item("MkCapa").Cells.Item(i).Specific.Value.ToString().Trim();

					for (j = 1; j <= Lastdd; j++)
					{
						sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_PP930H]";
						oRecordSet.DoQuery(sQry);
						if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
						{
							DocEntry = Convert.ToString(1);
						}
						else
						{
							DocEntry = Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
						}
						MP[j - 1] = oMat.Columns.Item("D" + j).Cells.Item(i).Specific.Value.ToString().Trim();

						if (j < 10)
						{
							DocDate = YYYYMM + "0" + j;
						}
						else
						{
							DocDate = YYYYMM + j;
						}

						sQry = "INSERT INTO [@PS_PP930H]";
						sQry += " (";
						sQry += " DocEntry,";
						sQry += " DocNum,";
						sQry += " U_BPLId,";
						sQry += " U_DocDate,";
						sQry += " U_ItmBsort,";
						sQry += " U_ItmBname,";
						sQry += " U_ItmMsort,";
						sQry += " U_ItmMname,";
						sQry += " U_ItemCode,";
						sQry += " U_ItemName,";
						sQry += " U_DInwon,";
						sQry += " U_NInwon,";
						sQry += " U_MkCapa,";
						sQry += " U_PQTy,";
						sQry += " U_DTime,";
						sQry += " U_NTime";
						sQry += " ) ";
						sQry += "VALUES(";
						sQry += DocEntry + ",";
						sQry += DocEntry + ",";
						sQry += "'" + BPLId + "',";
						sQry += "'" + DocDate + "',";
						sQry += "'" + ItmBsort + "',";
						sQry += "'" + ItmBsort + "',";
						sQry += "'" + ItmMsort + "',";
						sQry += "'" + ItmMname + "',";
						sQry += "'" + ItemCode + "',";
						sQry += "'" + ItemName + "',";
						sQry += "'" + DInwon + "',";
						sQry += "'" + NInwon + "',";
						sQry += "'" + Convert.ToDouble(MkCapa) + "',";
						sQry += "'" + Convert.ToDouble(MP[j - 1]) + "',";
						sQry += "'" + Convert.ToDouble(D[j - 1]) + "',";
						sQry += "'" + Convert.ToDouble(N[j - 1]) + "'";
						sQry += ")";
						oRecordSet.DoQuery(sQry);
					}

					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oMat.RowCount + "건 저장중...!";
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				functionReturnValue = true;
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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}

			return functionReturnValue;
		}

		/// <summary>
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("YYYYMM").Specific.Value.ToString().Trim()))
				{
					errMessage = "년월은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ItmBsort").Specific.Value.ToString().Trim()))
				{
					errMessage = "품목대분류는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}

				functionReturnValue = true;
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
			return functionReturnValue;
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
					//Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					//Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
					//Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
					//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
					//Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					//Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
					//Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
					//Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
					////Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "Btn_save")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (Add_PurchaseDemand(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							LoadData();
							oLast_Mode = Convert.ToInt16(oForm.Mode);
						}
					}
					else if (pVal.ItemUID == "Btn_Calc")
					{
						if (HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}
						CalcData();
					}
					else if (pVal.ItemUID == "Btn_ret")
					{
						if (HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}
						LoadData();
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
			try
			{
				if (pVal.CharPressed == 9)
				{
					if (pVal.ItemUID == "ItemCode")
					{
						if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
						{
							PSH_Globals.SBO_Application.ActivateMenuItem("7425");
							BubbleEvent = false;
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP930H);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1287": //복제
							break;
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
				if (BusinessObjectInfo.BeforeAction == true)
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
				else if (BusinessObjectInfo.BeforeAction == false)
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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
