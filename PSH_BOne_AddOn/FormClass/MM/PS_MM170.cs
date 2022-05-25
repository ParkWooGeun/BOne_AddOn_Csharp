using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 지체상금등록
	/// </summary>
	internal class PS_MM170 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_MM170H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM170L; //등록라인
		
		private string oLast_Item_UID; //클래스에서 선택한 마지막 아이템 Uid값
		private int oSeq;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM170.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM170_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM170");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);
				PS_MM170_CreateItems();
				PS_MM170_ComboBox_Setting();
				PS_MM170_Initial_Setting();
				PS_MM170_FormItemEnabled();
				PS_MM170_FormClear();
				PS_MM170_AddMatrixRow(0, oMat.RowCount, true);

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", true);  // 복제
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
				oForm.EnableMenu("1293", true);
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
		/// PS_MM170_CreateItems
		/// </summary>
		private void PS_MM170_CreateItems()
		{
			try
			{
				oDS_PS_MM170H = oForm.DataSources.DBDataSources.Item("@PS_MM170H");
				oDS_PS_MM170L = oForm.DataSources.DBDataSources.Item("@PS_MM170L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oDS_PS_MM170H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
				oDS_PS_MM170H.SetValue("U_JdtDate", 0, DateTime.Now.ToString("yyyyMMdd"));
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM170_ComboBox_Setting
		/// </summary>
		private void PS_MM170_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장 리스트
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);

				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				// 지체상금유무
				oMat.Columns.Item("RepayYN").ValidValues.Add("Y", "부여");
				oMat.Columns.Item("RepayYN").ValidValues.Add("N", "면제");

				// 품목대분류
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] ORDER BY Code", "", "");
				// 품목중분류
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("ItmMsort"), "SELECT U_Code, U_CodeName FROM [@PSH_ITMMSORT] ORDER BY U_Code", "", "");
				// 형태타입
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("ItemType"), "SELECT Code, Name FROM [@PSH_SHAPE] ORDER BY Code", "", "");
				// 질별
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("Quality"), "SELECT Code, Name FROM [@PSH_QUALITY] ORDER BY Code", "", "");
				// 인증기호
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("Mark"), "SELECT Code, Name FROM [@PSH_MARK] ORDER BY Code", "", "");
				// 매입기준단위
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("ObasUnit"), "SELECT Code, Name FROM [@PSH_UOMORG] ORDER BY Code", "", "");
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
		/// PS_MM170_Initial_Setting
		/// </summary>
		private void PS_MM170_Initial_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				// 인수자
				oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM170_FormItemEnabled
		/// </summary>
		private void PS_MM170_FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("JdtDate").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = true;
					oForm.Items.Item("JdtDate").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("JdtDate").Enabled = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM170_FormClear
		/// </summary>
		private void PS_MM170_FormClear()
		{
			string DocNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM170'", "");
				if (Convert.ToDouble(DocNum) == 0)
				{
					oDS_PS_MM170H.SetValue("DocNum", 0, "1");
				}
				else
				{
					oDS_PS_MM170H.SetValue("DocNum", 0, DocNum); // 화면에 적용이 안되기 때문
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM170_AddMatrixRow
		/// </summary>
		/// <param name="oSeq"></param>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM170_AddMatrixRow(short oSeq, int oRow, bool RowIserted)
		{
			try
			{
				switch (oSeq)
				{
					case 0:
						oMat.AddRow(); // 매트릭스에 새로운 로를 추가한다.
						oDS_PS_MM170L.SetValue("LineId", oRow, Convert.ToString(oRow + 1));
						oMat.LoadFromDataSource();
						break;
					case 1:
						oDS_PS_MM170L.InsertRecord(oRow);
						oDS_PS_MM170L.SetValue("LIneId", oRow, Convert.ToString(oRow + 1));
						oMat.LoadFromDataSource();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM170_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_MM170_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			try
			{
				switch (oUID)
				{
					case "Mat01":
						if ((oRow == oMat.RowCount || oMat.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat.Columns.Item("GRDocNum").Cells.Item(oRow).Specific.Value.ToString().Trim()))
						{
							oMat.FlushToDataSource();
							PS_MM170_AddMatrixRow(1, oMat.RowCount, true);
							oMat.Columns.Item("GRDocNum").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM170_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM170_HeaderSpaceLineDel()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_MM170H.GetValue("U_CardCode", 0).ToString().Trim()) || string.IsNullOrEmpty(oDS_PS_MM170H.GetValue("U_BPLId", 0).ToString().Trim()) || string.IsNullOrEmpty(oDS_PS_MM170H.GetValue("U_DocDate", 0).ToString().Trim()))
				{
					errMessage = "거래처코드, 사업장, 요청일자는 필수입력 사항입니다. 확인하세요.";
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
		/// PS_MM170_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM170_MatrixSpaceLineDel()
		{
			bool ReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount <= 1)
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}

				if (oMat.VisualRowCount > 0)
				{
					for (i = 0; i <= oMat.VisualRowCount - 2; i++)
					{
						oDS_PS_MM170L.Offset = i;
						if (string.IsNullOrEmpty(oDS_PS_MM170L.GetValue("U_GRDocNum", i).ToString().Trim()))
						{
							errMessage = "입고문서 데이터가 없습니다. 확인하세요.";
							throw new Exception();
						}

						if (string.IsNullOrEmpty(oDS_PS_MM170L.GetValue("U_RepayYN", i).ToString().Trim()))
						{
							errMessage = "지체상금유무 데이터가 없습니다. 확인하세요.";
							throw new Exception();
						}

						if (string.IsNullOrEmpty(oDS_PS_MM170L.GetValue("U_RepayP", i).ToString().Trim()))
						{
							errMessage = "지체금액 데이터가 없습니다. 확인하세요.";
							throw new Exception();
						}
					}
				}

				if (string.IsNullOrEmpty(oDS_PS_MM170L.GetValue("U_GRDocNum", oMat.VisualRowCount - 1).ToString().Trim()))
				{
					oDS_PS_MM170L.RemoveRecord(oMat.VisualRowCount - 1);
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
			return ReturnValue;
		}

		//private bool PS_MM170_Create_oJournalEntries(ref short ChkType)
		//{
		//	bool ReturnValue = false;
		//	// ERROR: Not supported in C#: OnErrorStatement

		//	SAPbobsCOM.JournalEntries f_oJournalEntries = null;

		//	int i = 0;
		//	short ErrNum = 0;
		//	int ErrCode = 0;
		//	string ErrMsg = null;
		//	string RetVal = null;
		//	string sTransId = null;

		//	string SCardCode = null;
		//	string sDocDate = null;
		//	decimal sSum = default(decimal);
		//	string sCC = null;
		//	SAPbobsCOM.Recordset oRecordSet = null;
		//	string sQry = null;

		//	if ((SubMain.Sbo_Company.InTransaction == true))
		//	{
		//		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
		//	}
		//	SubMain.Sbo_Company.StartTransaction();

		//	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	oMat.FlushToDataSource();

		//	SCardCode = oDS_PS_MM170H.GetValue("U_CardCode", 0));
		//	sDocDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDS_PS_MM170H.GetValue("U_JdtDate", 0), "0000-00-00");
		//	sSum = Convert.ToDecimal(oDS_PS_MM170H.GetValue("U_Sum", 0)));

		//	f_oJournalEntries = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

		//	var _with1 = f_oJournalEntries;
		//	_with1.ReferenceDate = Convert.ToDateTime(sDocDate);
		//	//전기일
		//	_with1.DueDate = Convert.ToDateTime(sDocDate);
		//	_with1.TaxDate = Convert.ToDateTime(sDocDate);

		//	_with1.Lines.Add();
		//	_with1.Lines.SetCurrentLine(0);
		//	_with1.Lines.AccountCode = "21101010";
		//	//관리계정
		//	_with1.Lines.ShortName = SCardCode;
		//	//G/L계정/BP 코드
		//	_with1.Lines.Debit = sSum;
		//	//차변
		//	_with1.Lines.LineMemo = "지체상금";

		//	_with1.Lines.SetCurrentLine(1);
		//	_with1.Lines.AccountCode = "43125020";
		//	//잡이익-기타
		//	_with1.Lines.ShortName = "43125020";
		//	_with1.Lines.Credit = sSum;
		//	//대변
		//	_with1.Lines.LineMemo = "지체상금";
		//	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//	_with1.UserFields.Fields.Item("U_BPLId").Value = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
		//	//사업장

		//	//// 완료
		//	RetVal = Convert.ToString(f_oJournalEntries.Add());
		//	if ((0 != Convert.ToDouble(RetVal)))
		//	{
		//		SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
		//		goto PS_MM170_Create_oJournalEntries;
		//	}

		//	sCC = "Y";

		//	if (ChkType == 1)
		//	{
		//		SubMain.Sbo_Company.GetNewObjectCode(out sTransId);
		//		sQry = "Update [@PS_MM170H] Set U_JdtNo = '" + sTransId + "', U_JdtDate = '" + sDocDate + "', U_Sum = '" + sSum + "', U_JdtCC = '" + sCC + "' ";
		//		sQry = sQry + "Where DocNum = '" + oDS_PS_MM170H.GetValue("DocNum", 0)) + "'";
		//		oRecordSet.DoQuery(sQry);
		//		if ((SubMain.Sbo_Company.InTransaction == true))
		//		{
		//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
		//		}
		//	}

		//	oDS_PS_MM170H.SetValue("U_JdtNo", 0, sTransId);
		//	oDS_PS_MM170H.SetValue("U_JdtDate", 0, Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Now, "yyyymmdd"));
		//	oDS_PS_MM170H.SetValue("U_Sum", 0, Convert.ToString(sSum));
		//	oDS_PS_MM170H.SetValue("U_JdtCC", 0, sCC);

		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	//UPGRADE_NOTE: f_oJournalEntries 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	f_oJournalEntries = null;
		//	ReturnValue = true;

		//	oForm.Items.Item("Btn02").Enabled = false;
		//	oForm.Items.Item("Btn03").Enabled = true;
		//	return ReturnValue;
		//PS_MM170_Create_oJournalEntries:


		//	/////////////////////////////////////////////////////////////////////////////////////////
		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	//UPGRADE_NOTE: f_oJournalEntries 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	f_oJournalEntries = null;
		//	if (SubMain.Sbo_Company.InTransaction)
		//	{
		//		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
		//	}
		//	ReturnValue = false;
		//	MDC_Com.MDC_GF_Message(ref "PS_MM170_Create_oJournalEntries:" + Err().Description + ErrMsg, ref "E");
		//	return ReturnValue;
		//}

		//private bool PS_MM170_Cancel_oJournalEntries(ref short ChkType)
		//{
		//	bool ReturnValue = false;
		//	// ERROR: Not supported in C#: OnErrorStatement

		//	SAPbobsCOM.JournalEntries f_oJournalEntries = null;

		//	int i = 0;
		//	short ErrNum = 0;
		//	int ErrCode = 0;
		//	string ErrMsg = null;
		//	int RetVal = 0;
		//	string sTransId = null;

		//	string SCardCode = null;
		//	string sDocDate = null;
		//	decimal sSum = default(decimal);
		//	string sCC = null;

		//	SAPbobsCOM.Recordset oRecordSet = null;
		//	string sQry = null;

		//	if ((SubMain.Sbo_Company.InTransaction == true))
		//	{
		//		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
		//	}

		//	SubMain.Sbo_Company.StartTransaction();

		//	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	oMat.FlushToDataSource();

		//	//UPGRADE_NOTE: f_oJournalEntries 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	f_oJournalEntries = null;
		//	f_oJournalEntries = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

		//	if ((f_oJournalEntries.GetByKey(Convert.ToInt32(oDS_PS_MM170H.GetValue("U_JdtNo", 0)))) == false))
		//	{
		//		SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
		//		goto PS_MM170_Cancel_oJournalEntries;
		//	}

		//	//// 완료
		//	RetVal = f_oJournalEntries.Cancel();
		//	if ((0 != RetVal))
		//	{
		//		SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
		//		goto PS_MM170_Cancel_oJournalEntries;
		//	}

		//	sCC = "N";

		//	if (ChkType == 1)
		//	{
		//		SubMain.Sbo_Company.GetNewObjectCode(out sTransId);
		//		sQry = "Update [@PS_MM170H] Set U_JdtCan = '" + sTransId + "', U_Sum = '" + sSum + "', U_JdtCC = '" + sCC + "' ";
		//		sQry = sQry + "Where DocNum = '" + oDS_PS_MM170H.GetValue("DocNum", 0)) + "'";
		//		oRecordSet.DoQuery(sQry);

		//		if ((SubMain.Sbo_Company.InTransaction == true))
		//		{
		//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
		//		}
		//	}

		//	//    oForm.Update
		//	oDS_PS_MM170H.SetValue("U_JdtCan", 0, sTransId);
		//	oDS_PS_MM170H.SetValue("U_Sum", 0, Convert.ToString(sSum));
		//	oDS_PS_MM170H.SetValue("U_JdtCC", 0, sCC);

		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	//UPGRADE_NOTE: f_oJournalEntries 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	f_oJournalEntries = null;
		//	ReturnValue = true;

		//	oForm.Items.Item("Btn02").Enabled = false;
		//	oForm.Items.Item("Btn03").Enabled = false;
		//	return ReturnValue;
		//PS_MM170_Cancel_oJournalEntries:


		//	/////////////////////////////////////////////////////////////////////////////////////////
		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	//UPGRADE_NOTE: f_oJournalEntries 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	f_oJournalEntries = null;
		//	if (SubMain.Sbo_Company.InTransaction)
		//	{
		//		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
		//	}
		//	ReturnValue = false;
		//	MDC_Com.MDC_GF_Message(ref "PS_MM170_Create_oJournalEntries:" + Err().Description + ErrMsg, ref "E");
		//	return ReturnValue;
		//}

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
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
                case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE: //18
                    Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
			string errMessage = string.Empty;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM170_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_MM170_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
					}
					else if (pVal.ItemUID == "Btn02") // DI API - 분개 생성
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("JdtDate").Specific.Value.ToString().Trim()))
							{
								errMessage = "분개처리일을 먼저 입력하세요.";
								throw new Exception();
							}
							else if (oForm.Items.Item("Status").Specific.Value.ToString().Trim() == "C")
							{
								errMessage = "문서가 Close 또는 Cancel 되었습니다.";
								throw new Exception();
							}
							else
							{
								//if (PS_MM170_Create_oJournalEntries(ref 1) == false)
								//{
								//	BubbleEvent = false;
								//	return;
								//}
							}
						}
						else
						{
							errMessage = "먼저 저장한 후 분개 처리 바랍니다.";
							throw new Exception();
						}
					}
					else if (pVal.ItemUID == "Btn03") // DI API - 분개 취소
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("JdtDate").Specific.Value.ToString().Trim()))
							{
								errMessage = "분개처리일을 먼저 입력하세요.";
								throw new Exception();
							}
							else if (oForm.Items.Item("JdtCC").Specific.Value.ToString().Trim() != "Y")
							{
								errMessage = "분개생성:Y일 때 취소 할 수 있습니다.";
								throw new Exception();
							}
							else if (oForm.Items.Item("Status").Specific.Value.ToString().Trim() == "C")
							{
								errMessage = "문서가 Close 또는 Cancel 되었습니다.";
								throw new Exception();
							}
							else
							{
								//if (PS_MM170_Cancel_oJournalEntries(ref 1) == false)
								//{
								//	BubbleEvent = false;
								//	return;
								//}
							}
						}
						else
						{
							errMessage = "먼저 저장한 후 분개 처리 바랍니다.";
							throw new Exception();
						}
					}
					else
					{
						if (pVal.ItemChanged == true)
						{
							if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItemCode")
							{
								PS_MM170_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == true)
						{
							PS_MM170_FormItemEnabled();
							PS_MM170_FormClear();
							PS_MM170_AddMatrixRow(0, oMat.RowCount, true);
							oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Action_Success == false)
						{
							PS_MM170_FormItemEnabled();
							PS_MM170_AddMatrixRow(1, oMat.RowCount, true);
						}
					}
				}
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
					BubbleEvent = false;
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
			string errMessage = string.Empty;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
					{
						if (pVal.ItemUID == "CardCode" && pVal.CharPressed == 9)
						{
							oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							PSH_Globals.SBO_Application.ActivateMenuItem("7425");
							BubbleEvent = false;
						}
					}

					// 입고번호 - 서브폼 호출
					if (pVal.ItemUID == "Mat01" && pVal.ColUID == "GRDocNum" && pVal.CharPressed == 9)
					{
						if (string.IsNullOrEmpty(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()) || string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
							{
								errMessage = "거래처코드와, 사업장을 먼저 입력하세요.";
								throw new Exception();
							}
							else
							{
								PS_MM171 TempForm01 = new PS_MM171();
								TempForm01.LoadForm(ref oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
							}
							BubbleEvent = false;
						}
					}

					// 담당자
					if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
					{
						if (pVal.ItemUID == "CntcCode" && pVal.CharPressed == 9)
						{
							oForm.Items.Item("CntcCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							PSH_Globals.SBO_Application.ActivateMenuItem("7425");
							BubbleEvent = false;
						}
					}

					// 품목코드
					if (pVal.ItemUID == "Mat01" && pVal.ColUID == "ItemCode" && pVal.CharPressed == 9)
					{
						if (string.IsNullOrEmpty(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.String))
						{
							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							PSH_Globals.SBO_Application.ActivateMenuItem("7425");
							BubbleEvent = false;
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.Action_Success == true)
					{
						oSeq = 1;
					}
				}
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
					BubbleEvent = false;
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
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
					oLast_Item_UID = pVal.ItemUID;
				}
				else
				{
					oLast_Item_UID = pVal.ItemUID;
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
			int i;

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01" && pVal.ColUID == "RepayYN")
					{
						oForm.Items.Item("Sum").Specific.Value = "";
						for (i = 1; i <= oMat.VisualRowCount - 1; i++)
						{
							if (oMat.Columns.Item("RepayYN").Cells.Item(i).Specific.Value.ToString().Trim() == "Y")
							{
								oForm.Items.Item("Sum").Specific.Value = Convert.ToString(Convert.ToDouble(oForm.Items.Item("Sum").Specific.Value.ToString().Trim()) + Convert.ToDouble(oMat.Columns.Item("RepayP").Cells.Item(i).Specific.Value.ToString().Trim()));
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
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01" && pVal.ColUID == "GRDocNum")
						{
							PS_MM170_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					if (pVal.ItemUID == "CntcCode")
					{
						sQry = "Select U_FULLNAME, U_MSTCOD From [OHEM] Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					}
					
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01" && pVal.ColUID == "GRDocNum")
						{
							PS_MM170_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "CardCode" && pVal.ItemChanged == true)
					{
						sQry = "Select CardName From [OCRD] Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						oDS_PS_MM170H.SetValue("U_CardName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					}

					if (pVal.ItemUID == "DocDate")
					{
						//전기일자를 분개일자와 동일하게...
						oDS_PS_MM170H.SetValue("U_JdtDate", 0, oForm.Items.Item("DocDate").Specific.Value.ToString().Trim());
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
					PS_MM170_AddMatrixRow(1, oMat.VisualRowCount, true);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_ACTIVATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_ACTIVATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (oSeq == 1)
					{
						oSeq = 0;
					}
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM170H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM170L);
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
		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (PSH_Globals.SBO_Application.MessageBox("정말 삭제 하시겠습니까?", 1, "OK", "NO") != 1)
					{
						BubbleEvent = false;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					for (i = 1; i <= oMat.VisualRowCount; i++)
					{
						oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
					}
					oMat.FlushToDataSource();
					oDS_PS_MM170L.RemoveRecord(oDS_PS_MM170L.Size - 1);
					oMat.LoadFromDataSource();
					if (oMat.RowCount == 0)
					{
						PS_MM170_AddMatrixRow(0, oMat.RowCount, true);
					}
					else
					{
						if (!string.IsNullOrEmpty(oDS_PS_MM170L.GetValue("U_ItemCode", oMat.RowCount - 1).ToString().Trim()))
						{
							PS_MM170_AddMatrixRow(1, oMat.RowCount, true);
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
							Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
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
							PS_MM170_FormItemEnabled();
							break;
						case "1282": //추가
							PS_MM170_FormItemEnabled();
							PS_MM170_FormClear();
							PS_MM170_AddMatrixRow(0, oMat.RowCount, true);
							oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
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
							PS_MM170_FormItemEnabled();
							if (oMat.VisualRowCount > 0)
							{
								if (!string.IsNullOrEmpty(oMat.Columns.Item("GRDocNum").Cells.Item(oMat.VisualRowCount).Specific.Value.ToString().Trim()))
								{
									PS_MM170_AddMatrixRow(1, oMat.RowCount, true);
								}
							}
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
