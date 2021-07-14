using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작업외 공수 등록
	/// </summary>
	internal class PS_PP060 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP060H; //등록헤더

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Mode;

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		public override void LoadForm(string oFromDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP060.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP060_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP060");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP060_CreateItems();
				PS_PP060_ComboBox_Setting();
				Add_MatrixRow(0, true);
				LoadCaption();
				FormItemEnabled();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", true);	 // 취소
				oForm.EnableMenu("1293", true);	 // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
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
		/// PS_PP060_CreateItems
		/// </summary>
		private void PS_PP060_CreateItems()
		{
			try
			{
				oDS_PS_PP060H = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();
				oMat.Columns.Item("DocNum").Visible = false; //DocNum Hidden 처리(2015.01.27 송명규)

				//유저데이타 속성 선언 날짜형식
				oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
				oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");

				//일자 Set
				oForm.Items.Item("DocDateFr").Specific.Value = DateTime.Now.AddDays(-1).ToString("yyyyMMdd"); 
				oForm.Items.Item("DocDateTo").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.Items.Item("CntcCode").Click();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP060_ComboBox_Setting
		/// </summary>
		private void PS_PP060_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				oForm.Items.Item("WorkGbn").Specific.ValidValues.Add("", "");
				dataHelpClass.Set_ComboList(oForm.Items.Item("WorkGbn").Specific, "select Code, Name from [@PSH_ITMBSORT] Where U_PudYN = 'Y' order by Code", "101", false, false);

				oForm.Items.Item("OrdType").Specific.ValidValues.Add("10", "실동");
				oForm.Items.Item("OrdType").Specific.ValidValues.Add("20", "비가동");
				oForm.Items.Item("OrdType").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				//공정구분 입력(시스템코드로 수정, 2014.11.18 송명규)
				sQry = "    SELECT      U_Minor,";
				sQry += "                U_CdName";
				sQry += " FROM       [@PS_SY001L]";
				sQry += " WHERE     Code = 'P206'";
				sQry += "                AND U_UseYN = 'Y'";
				sQry += " ORDER BY  U_Seq";
				oForm.Items.Item("CpGbn").Specific.ValidValues.Add("%", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CpGbn").Specific, sQry, "", false, false);
				oForm.Items.Item("CpGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				sQry = "    SELECT      U_Minor,";
				sQry += "                U_CdName";
				sQry += " FROM       [@PS_SY001L]";
				sQry += " WHERE     Code = 'P206'";
				sQry += "                AND U_UseYN = 'Y'";
				sQry += " ORDER BY  U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("CpGbn"), sQry, "", "");

				//작업구분(매트릭스) 2013.05.01 송명규 추가
				sQry = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' ORDER BY Code";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("WorkGbn"), sQry, "", "");

				//문서구분(매트릭스) 2013.05.01 송명규 추가
				oMat.Columns.Item("OrdType").ValidValues.Add("10", "실동");
				oMat.Columns.Item("OrdType").ValidValues.Add("20", "비가동");

				//사업장(매트릭스) 2013.05.01 송명규 추가
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");

				sQry = "SELECT B.U_Minor, B.U_CdName From [@PS_SY001H] A, [@PS_SY001L] B WHERE A.CODE = B.CODE AND A.CODE = 'P005' Order By U_Minor ";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oMat.Columns.Item("NCode").ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
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
					oDS_PS_PP060H.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_PP060H.Offset = oRow;
				oDS_PS_PP060H.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// LoadCaption
		/// </summary>
		private void LoadCaption()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("Btn_save").Specific.Caption = "추가";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("Btn_save").Specific.Caption = "수정";
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// FormItemEnabled
		/// </summary>
		private void FormItemEnabled()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("OrdType").Enabled = true;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("CpGbn").Enabled = true;
					oForm.Items.Item("CpCode").Enabled = true;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("WorkGbn").Enabled = true;
					oForm.Items.Item("DocDateFr").Enabled = true;
					oForm.Items.Item("DocDateTo").Enabled = true;
					oMat.Columns.Item("BPLId").Editable = false;
					oMat.Columns.Item("CpGbn").Editable = false;
					oMat.Columns.Item("CpCode").Editable = false;
					oMat.Columns.Item("CpName").Editable = false;
					oMat.Columns.Item("CntcCode").Editable = true;
					oMat.Columns.Item("CntcName").Editable = false;
					oMat.Columns.Item("ItmBsort").Editable = true;
					oMat.Columns.Item("DocDate").Editable = true;
					oMat.Columns.Item("WorkNote").Editable = true;
					oMat.Columns.Item("WorkTime").Editable = true;
					oMat.Columns.Item("WorkGbn").Editable = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("OrdType").Enabled = true;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("CpGbn").Enabled = true;
					oForm.Items.Item("CpCode").Enabled = true;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("WorkGbn").Enabled = true;
					oForm.Items.Item("DocDateFr").Enabled = true;
					oForm.Items.Item("DocDateTo").Enabled = true;
					oMat.Columns.Item("BPLId").Editable = false;
					oMat.Columns.Item("CpGbn").Editable = false;
					oMat.Columns.Item("CpCode").Editable = false;
					oMat.Columns.Item("CpName").Editable = false;
					oMat.Columns.Item("CntcCode").Editable = false;
					oMat.Columns.Item("CntcName").Editable = false;
					oMat.Columns.Item("ItmBsort").Editable = false;
					oMat.Columns.Item("DocDate").Editable = false;
					oMat.Columns.Item("WorkNote").Editable = false;
					oMat.Columns.Item("WorkTime").Editable = false;
					oMat.Columns.Item("WorkGbn").Editable = false;
					oForm.Items.Item("Btn_ret").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
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
				if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "10" && oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim() == "%")
                {
					errMessage = "공정구분은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
                {
					errMessage = "사업장은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "10" && oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim() != "60" 
					&& oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim() != "70" && oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim() != "80" 
					   && string.IsNullOrEmpty(oForm.Items.Item("CpCode").Specific.Value.ToString().Trim()))
                {
					errMessage = "공정코드는 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim()))
                {
					errMessage = "작업구분은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("OrdType").Specific.Value.ToString().Trim()))
                {
					errMessage = "문서구분은 필수사항입니다. 확인하세요.";
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
		/// MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool MatrixSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;
			int i;
			// 헤드사항을 입력
			string CpName;
			string CpGbn;
			string OrdType;
			string BPLID;
			string CpCode;
			string WorkGbn;

			try
			{
				OrdType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CpGbn = oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				CpName = oForm.Items.Item("CpName").Specific.Value.ToString().Trim();
				WorkGbn = oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim();

				//메트릭스 읽을때 선언
				oMat.FlushToDataSource();

				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP060H.GetValue("U_ColReg07", i).ToString().Trim()))
					{
						oDS_PS_PP060H.RemoveRecord(i); // Mat01에 마지막라인(빈라인) 삭제
					}
				}
				oMat.LoadFromDataSource();
				// 라인
				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인 데이터가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					if (string.IsNullOrEmpty(oDS_PS_PP060H.GetValue("U_ColReg07", i).ToString().Trim()))
					{
						errMessage = Convert.ToString(i + 1) + "번 라인의 사원코드가 없습니다. 확인하세요.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oDS_PS_PP060H.GetValue("U_ColQty01", i).ToString().Trim()))
					{
						errMessage = Convert.ToString(i + 1) + "번 라인의 시간이 없습니다. 확인하세요.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oDS_PS_PP060H.GetValue("U_ColDt01", i).ToString().Trim()))
					{
						errMessage = Convert.ToString(i + 1) + "번 라인의 등록일자가 없습니다. 확인하세요.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oDS_PS_PP060H.GetValue("U_ColReg15", i).ToString().Trim()))
					{
						if (OrdType == "20")
						{
							errMessage = Convert.ToString(i + 1) + "번 라인의 비가동코드가 없습니다. 확인하세요.";
							throw new Exception();
						}
					}
				}
				//메트릭스의 값변경후 선언
				oMat.LoadFromDataSource();
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
		/// FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				switch (oUID)
				{
					case "CntcCode": //사번
						sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;

					case "Mat01": //메트릭스
						if (oCol == "CntcCode") //사번코드
						{
							sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oMat.Columns.Item("CntcCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oMat.Columns.Item("CntcName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							oMat.Columns.Item("DocDate").Cells.Item(oRow).Specific.Value = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");
							oMat.FlushToDataSource();

							if (oMat.RowCount == oRow && !string.IsNullOrEmpty(oDS_PS_PP060H.GetValue("U_ColReg07", oRow - 1).ToString().Trim())) 
							{
								Add_MatrixRow(oRow, false);
							}
						}
						else if (oCol == "FixCode")
						{
							sQry = "    SELECT       U_FixName AS [FixName],";
							sQry += "                 U_TempChr1 As FixCode2";
							sQry += " FROM        [@PS_FX005H]";
							sQry += " WHERE       U_FixCode + '-' + U_SubCode = '" + oMat.Columns.Item("FixCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oMat.Columns.Item("FixName").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("FixName").Value.ToString().Trim();
							oMat.Columns.Item("FixCode2").Cells.Item(oRow).Specific.Value = oRecordSet.Fields.Item("FixCode2").Value.ToString().Trim();
							oMat.FlushToDataSource();
						}
						oMat.AutoResizeColumns();
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
		/// PS_PP060_OpenItemRegist
		/// </summary>
		/// <param name="pRow"></param>
		private void PS_PP060_OpenItemRegist(int pRow)
		{
			string StdNo;	 //기준문서번호
			string BPLID;	 //사업장
			string CpCode;	 //공정코드
			string CpName;	 //공정명
			string WkCode;	 //작업자사번
			string WkName;	 //작업자명
			string FixCode;	 //자산코드
			string FixName;	 //자산명
			string DocDate;	 //등록일자
			string WorkTime; //작업시간
			string WorkNote; //작업내용

			try
			{
				StdNo = oMat.Columns.Item("DocEntry").Cells.Item(pRow).Specific.Value.ToString().Trim();
				BPLID = oMat.Columns.Item("BPLId").Cells.Item(pRow).Specific.Value.ToString().Trim();
				CpCode = oMat.Columns.Item("CpCode").Cells.Item(pRow).Specific.Value.ToString().Trim();
				CpName = oMat.Columns.Item("CpName").Cells.Item(pRow).Specific.Value.ToString().Trim();
				WkCode = oMat.Columns.Item("CntcCode").Cells.Item(pRow).Specific.Value.ToString().Trim();
				WkName = oMat.Columns.Item("CntcName").Cells.Item(pRow).Specific.Value.ToString().Trim();
				FixCode = oMat.Columns.Item("FixCode").Cells.Item(pRow).Specific.Value.ToString().Trim();
				FixName = oMat.Columns.Item("FixName").Cells.Item(pRow).Specific.Value.ToString().Trim();
				DocDate = oMat.Columns.Item("DocDate").Cells.Item(pRow).Specific.Value.ToString().Trim();
				WorkTime = oMat.Columns.Item("WorkTime").Cells.Item(pRow).Specific.Value.ToString().Trim();
				WorkNote = oMat.Columns.Item("WorkNote").Cells.Item(pRow).Specific.Value.ToString().Trim();

				PS_PP061 oTempClass = new PS_PP061();
				oTempClass.LoadForm(StdNo, BPLID, CpCode, CpName, WkCode, WkName, FixCode, FixName, DocDate, WorkTime, WorkNote);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// LoadData
		/// 조회데이타 가져오기
		/// </summary>
		private void LoadData()
		{
			int i;
			string sQry;
			string errMessage = string.Empty;

			string DocDateFr;
			string CntcCode;
			string CpGbn;
			string OrdType;
			string BPLID;
			string CpCode;
			string WorkGbn;
			string DocDateTo;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				OrdType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CpGbn = oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				WorkGbn = oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim();
				DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
				DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(BPLID))
                {
					BPLID = "%";
				}
				if (string.IsNullOrEmpty(CpGbn))
                {
					CpGbn = "%";
				}
				if (string.IsNullOrEmpty(CpCode))
                {
					CpCode = "%";
				}
				if (string.IsNullOrEmpty(CntcCode))
                {
					CntcCode = "%";
				}
				if (string.IsNullOrEmpty(WorkGbn))
                {
					WorkGbn = "%";
				}
				if (string.IsNullOrEmpty(DocDateFr))
                {
					DocDateFr = "19000101";
				}
				if (string.IsNullOrEmpty(DocDateTo))
                {
					DocDateTo = "20991231";
				}

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = "EXEC [PS_PP060_01] '" + OrdType + "', '" + BPLID + "','" + CpGbn + "','" + CpCode + "','" + CntcCode + "','" + WorkGbn + "','" + DocDateFr + "','" + DocDateTo + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP060H.Clear();
				oMat.FlushToDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					Add_MatrixRow(0, true);
					LoadCaption();
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP060H.Size)
					{
						oDS_PS_PP060H.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP060H.Offset = i;

					oDS_PS_PP060H.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP060H.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());  	//문서번호
					oDS_PS_PP060H.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("DocNum").Value.ToString().Trim());		//문서번호2
					oDS_PS_PP060H.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("U_CpGbn").Value.ToString().Trim());		//공정구분
					oDS_PS_PP060H.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("U_CpCode").Value.ToString().Trim());	//공정코드
					oDS_PS_PP060H.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("U_CpName").Value.ToString().Trim());	//공정명
					oDS_PS_PP060H.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("U_CntcCode").Value.ToString().Trim());	//사원코드
					oDS_PS_PP060H.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("U_CntcName").Value.ToString().Trim());	//사원명
					oDS_PS_PP060H.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("U_FixCode").Value.ToString().Trim());	//자산코드
					oDS_PS_PP060H.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("U_FixName").Value.ToString().Trim());	//장비명
					oDS_PS_PP060H.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("U_FixCode2").Value.ToString().Trim());	//자산번호(구)
					oDS_PS_PP060H.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("U_ItmBsort").Value.ToString().Trim());	//금형번호
					oDS_PS_PP060H.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("U_BPLId").Value.ToString().Trim());		//사업장
					oDS_PS_PP060H.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("U_DocDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //등록일자
					oDS_PS_PP060H.SetValue("U_ColReg15", i, oRecordSet.Fields.Item("U_NCode").Value.ToString().Trim());		//비가동코드
					oDS_PS_PP060H.SetValue("U_ColReg16", i, oRecordSet.Fields.Item("U_WorkNote").Value.ToString().Trim());  //작업내용
					oDS_PS_PP060H.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("U_WorkTime").Value.ToString().Trim());  //작업시간
					oDS_PS_PP060H.SetValue("U_ColReg18", i, oRecordSet.Fields.Item("U_WorkGbn").Value.ToString().Trim());   //작업구분
					oDS_PS_PP060H.SetValue("U_ColReg19", i, oRecordSet.Fields.Item("U_OrdType").Value.ToString().Trim());   //문서구분
					oDS_PS_PP060H.SetValue("U_ColReg20", i, oRecordSet.Fields.Item("ItmRegYN").Value.ToString().Trim());	//자재등록여부
					oRecordSet.MoveNext();

                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// DeleteData
		/// </summary>
		private void DeleteData()
		{
			int i;
			string sQry;
			string ItmRegYN;
			string DocEntry;
			string DocNum;
			string Check;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat.FlushToDataSource();

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					Check = oDS_PS_PP060H.GetValue("U_ColReg03", i).ToString().Trim();

					if (Check == "Y")
					{
						DocEntry = oDS_PS_PP060H.GetValue("U_ColReg01", i).ToString().Trim();
						DocNum = oDS_PS_PP060H.GetValue("U_ColReg02", i).ToString().Trim();
						ItmRegYN = oDS_PS_PP060H.GetValue("U_ColReg20", i).ToString().Trim();

						sQry = "Delete From [@PS_PP060H] where DocEntry = '" + DocEntry + "' and DocNum = '" + DocNum + "'";

						if (ItmRegYN == "N")
						{
							oRecordSet.DoQuery(sQry); //소요자재가 등록되지 않은 행만 삭제
						}
					}
				}
				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();
				Add_MatrixRow(0, true);
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
		/// UpdateData
		///  데이타 UPDATE
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool UpdateData(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool functionReturnValue = false;

			int i;
			int j = 0;
			string sQry;
			string errMessage = string.Empty;

			string BPLID;
			string DocEntry;
			decimal WorkTime;
			string OrdType;
			string WorkGbn;
			string ItmBsort;
			string CntcCode;
			string CpCode;
			string CpGbn;
			string CpName;
			string CntcName;
			string DocDate;
			string WorkNote;
			string NCode;
			string FixCode;
			string FixName;
			string FixCode2;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat.FlushToDataSource();

				OrdType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CpGbn = oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				CpName = oForm.Items.Item("CpName").Specific.Value.ToString().Trim();
				WorkGbn = oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim();

				for (i = 1; i <= oMat.RowCount; i++)
				{
					if (oMat.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)
					{
						j += 1;
					}
				}

				if (j <= 0)
				{
					errMessage = "수정할 항목이 없습니다. 수정하실려면 항목을 선택을 하세요!";
					throw new Exception();
				}

				for (i = 1; i <= oMat.RowCount; i++)
				{
					if (oMat.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)
					{
						j += 1;
						DocEntry = oMat.Columns.Item("DocEntry").Cells.Item(i).Specific.Value.ToString().Trim();

						CntcCode = oMat.Columns.Item("CntcCode").Cells.Item(i).Specific.Value.ToString().Trim();
						CntcName = oMat.Columns.Item("CntcName").Cells.Item(i).Specific.Value.ToString().Trim();

						FixCode = oMat.Columns.Item("FixCode").Cells.Item(i).Specific.Value.ToString().Trim();
						FixName = oMat.Columns.Item("FixName").Cells.Item(i).Specific.Value.ToString().Trim();
						FixCode2 = oMat.Columns.Item("FixCode2").Cells.Item(i).Specific.Value.ToString().Trim();

						ItmBsort = oMat.Columns.Item("ItmBsort").Cells.Item(i).Specific.Value.ToString().Trim();
						NCode = oMat.Columns.Item("NCode").Cells.Item(i).Specific.Value.ToString().Trim();
						WorkNote = oMat.Columns.Item("WorkNote").Cells.Item(i).Specific.Value.ToString().Trim();
						WorkTime = Convert.ToDecimal(oMat.Columns.Item("WorkTime").Cells.Item(i).Specific.Value.ToString().Trim());
						DocDate = oMat.Columns.Item("DocDate").Cells.Item(i).Specific.Value.ToString().Trim();

						sQry = " Update [@PS_PP060H]";
						sQry += " set ";
						sQry += " U_CntcCode = '" + CntcCode + "',";
						sQry += " U_CntcName = '" + CntcName + "',";
						sQry += " U_FixCode = '" + FixCode + "',";
						sQry += " U_FixName = '" + FixName + "',";
						sQry += " U_FixCode2 = '" + FixCode2 + "',";
						sQry += " U_ItmBsort = '" + ItmBsort + "',";
						sQry += " U_DocDate  = '" + DocDate + "',";
						sQry += " U_CpCode  = '" + CpCode + "',";
						sQry += " U_CpName  = '" + CpName + "',";
						sQry += " U_NCode = '" + NCode + "',";
						sQry += " U_WorkNote = '" + WorkNote + "',";
						sQry += " U_WorkTime = '" + WorkTime + "'";
						sQry += " Where DocEntry = '" + DocEntry + "'";
						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("작업외 공수수정 완료!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
				oMat.FlushToDataSource();
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
			}
			return functionReturnValue;
		}

		/// <summary>
		/// Add_PurchaseDemand
		/// 데이타 INSERT
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool Add_PurchaseDemand(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool functionReturnValue = false;

			int i;
			string sQry;

			string BPLID;
			string DocNum;
			string DocEntry;
			string OrdType;
			string LineNum;
			decimal WorkTime;
			string WorkNote;
			string DocDate;
			string CntcName;
			string CpName;
			string CpGbn;
			string CpCode;
			string CntcCode;
			string ItmBsort;
			string WorkGbn;
			string NCode;
			string FixCode;
			string FixName;
			string FixCode2;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				oMat.FlushToDataSource();

				OrdType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CpGbn = oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim();
				CpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();
				CpName = oForm.Items.Item("CpName").Specific.Value.ToString().Trim();
				WorkGbn = oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim();

				if (OrdType == "20")
				{
					//비가동입력
					CpGbn = "";
					CpCode = "";
					CpName = "";
				}

				for (i = 0; i <= oMat.RowCount - 2; i++)
				{
					DocDate = oDS_PS_PP060H.GetValue("U_ColDt01", i).ToString().Trim();

					sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_PP060H] where Left(DocEntry, 6) = Left('" + DocDate + "', 6)";
					oRecordSet.DoQuery(sQry);
					if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
					{
						DocEntry = codeHelpClass.Left(DocDate, 6) + "0001";
					}
					else
					{
						DocEntry = Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
					}

					DocNum = Convert.ToString(i + 1);
					LineNum = Convert.ToString(i + 1);

					CntcCode = oDS_PS_PP060H.GetValue("U_ColReg07", i).ToString().Trim();
					CntcName = oDS_PS_PP060H.GetValue("U_ColReg08", i).ToString().Trim();
					FixCode = oDS_PS_PP060H.GetValue("U_ColReg09", i).ToString().Trim();
					FixName = oDS_PS_PP060H.GetValue("U_ColReg10", i).ToString().Trim();
					FixCode2 = oDS_PS_PP060H.GetValue("U_ColReg11", i).ToString().Trim();
					ItmBsort = oDS_PS_PP060H.GetValue("U_ColReg12", i).ToString().Trim();
					NCode = oDS_PS_PP060H.GetValue("U_ColReg15", i).ToString().Trim();
					WorkNote = oDS_PS_PP060H.GetValue("U_ColReg16", i).ToString().Trim();
					WorkTime = Convert.ToDecimal(oDS_PS_PP060H.GetValue("U_ColQty01", i).ToString().Trim());

					sQry = "INSERT INTO [@PS_PP060H]";
					sQry += " (";
					sQry += " DocEntry,";
					sQry += " DocNum,";
					sQry += " U_BPLId,";
					sQry += " U_LineNum,";
					sQry += " U_CpGbn,";
					sQry += " U_CpCode,";
					sQry += " U_CpName,";
					sQry += " U_CntcCode,";
					sQry += " U_CntcName,";
					sQry += " U_FixCode,";
					sQry += " U_FixName,";
					sQry += " U_FixCode2,";
					sQry += " U_ItmBsort,";
					sQry += " U_DocDate,";
					sQry += " U_WorkNote,";
					sQry += " U_WorkTime,";
					sQry += " U_WorkGbn,";
					sQry += " U_OrdType,";
					sQry += " U_NCode,";
					sQry += " U_ItmRegYN";
					sQry += " ) ";
					sQry += "VALUES(";
					sQry += DocEntry + ",";
					sQry += DocNum + ",";
					sQry += "'" + BPLID + "',";
					sQry += "'" + LineNum + "',";
					sQry += "'" + CpGbn + "',";
					sQry += "'" + CpCode + "',";
					sQry += "'" + CpName + "',";
					sQry += "'" + CntcCode + "',";
					sQry += "'" + CntcName + "',";
					sQry += "'" + FixCode + "',";
					sQry += "'" + FixName + "',";
					sQry += "'" + FixCode2 + "',";
					sQry += "'" + ItmBsort + "',";
					sQry += "'" + DocDate + "',";
					sQry += "'" + WorkNote + "',";
					sQry += "'" + WorkTime + "',";
					sQry += "'" + WorkGbn + "',";
					sQry += "'" + OrdType + "',";
					sQry += "'" + NCode + "',";
					sQry += "'N'";
					sQry += ")";
					oRecordSet.DoQuery(sQry);
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("작업외 공수등록 완료!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
							if (MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (Add_PurchaseDemand(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}

							oMat.Clear();
							oMat.FlushToDataSource();
							oMat.LoadFromDataSource();
							Add_MatrixRow(0, true);
							oLast_Mode = Convert.ToInt32(oForm.Mode);
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (UpdateData(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}

							LoadData();
						}
					}
					else if (pVal.ItemUID == "Btn_ret")
					{
						if (HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}

						oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						LoadCaption();
						LoadData();
					}
					else if (pVal.ItemUID == "Btn_del")
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?                              ※소요자재가 등록된 행은 삭제가 불가능합니다. 소요자재 정보를 먼저 삭제하세요.", 1, "예", "아니오") == 1)
						{
							DeleteData();
							LoadData();
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
						if (pVal.ItemUID == "CntcCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "CpCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
							else if (pVal.ColUID == "CntcCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("CntcCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
							else if (pVal.ColUID == "FixCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("FixCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
							oMat.SelectRow(pVal.Row, true, false);
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int i;
			string sQry;

			int sCount;
			int sSeq;
			string sCode = string.Empty;
			string SCpCode = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "CpGbn")
					{
						sCount = oForm.Items.Item("CpCode").Specific.ValidValues.Count;
						sSeq = sCount;
						for (i = 1; i <= sCount; i++)
						{
							oForm.Items.Item("CpCode").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
							sSeq -= 1;
						}

						//공정구분에 따른 공정코드변경
						switch (oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim())
						{
							case "10":
								//멀티게이지-금형래핑
								oForm.Items.Item("WorkGbn").Specific.Select("104", SAPbouiCOM.BoSearchKey.psk_ByValue);
								oForm.Items.Item("ItmBsort").Enabled = true;
								oMat.Columns.Item("ItmBsort").Editable = true;
								break;
							case "20":
								//멀티게이지-외경연삭
								oForm.Items.Item("WorkGbn").Specific.Select("104", SAPbouiCOM.BoSearchKey.psk_ByValue);
								oForm.Items.Item("ItmBsort").Enabled = false;
								oMat.Columns.Item("ItmBsort").Editable = false;
								break;
							case "30":
								//멀티게이지-포장
								oForm.Items.Item("WorkGbn").Specific.Select("104", SAPbouiCOM.BoSearchKey.psk_ByValue);
								oForm.Items.Item("ItmBsort").Enabled = false;
								oMat.Columns.Item("ItmBsort").Editable = false;
								break;
							case "40":
								//휘팅-바렐
								oForm.Items.Item("WorkGbn").Specific.Select("101", SAPbouiCOM.BoSearchKey.psk_ByValue);
								oForm.Items.Item("ItmBsort").Enabled = false;
								oMat.Columns.Item("ItmBsort").Editable = false;
								break;
							case "50":
								//휘팅-포장
								oForm.Items.Item("WorkGbn").Specific.Select("101", SAPbouiCOM.BoSearchKey.psk_ByValue);
								oForm.Items.Item("ItmBsort").Enabled = false;
								oMat.Columns.Item("ItmBsort").Editable = false;
								break;
							case "60":
							case "70":
							case "80":
								//검사공수 입력
								oForm.Items.Item("WorkGbn").Specific.Select("105", SAPbouiCOM.BoSearchKey.psk_ByValue);
								break;
							default:
								break;
						}
					}
					else if (pVal.ItemUID == "WorkGbn")
					{
						if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "20")
						{
							//비가동
						}
						else
						{
							//실동
							sCount = oForm.Items.Item("CpCode").Specific.ValidValues.Count;
							sSeq = sCount;
							for (i = 1; i <= sCount; i++)
							{
								oForm.Items.Item("CpCode").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
								sSeq -= 1;
							}

							switch (oForm.Items.Item("WorkGbn").Specific.Value.ToString().Trim())
							{
								case "101":
									sCode = "CP301";
									SCpCode = "%";
									break;
								case "102":
									sCode = "CP401";
									SCpCode = "%";
									break;
								case "104":
									sCode = "CP501";
									SCpCode = "%";
									break;
								case "105":
								case "106":
									sCode = "%";
									SCpCode = "%";
									break;

								case "107":
									sCode = "CP101";
									SCpCode = "%";
									break;
							}

							sQry = "SELECT U_CpCode, U_CpName From [@PS_PP001L] Where Code like '" + sCode + "' and U_CpCode like '" + SCpCode + "' Order by Code";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("CpCode").Specific.ValidValues.Add("", "");

							while (!(oRecordSet.EoF))
							{
								oForm.Items.Item("CpCode").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oRecordSet.MoveNext();
							}

							switch (oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim())
							{
								case "10":
									//v-mill
									SCpCode = "CP50101";
									break;
								case "20":
									// 콤보박스에 값을 지정 멀티게이지-외경연삭일때는 FRM공정을 SD380
									SCpCode = "CP50104";
									break;
								case "30":
									// 콤보박스에 값을 지정 멀티게이지-포장일때는 FRM공정을 SD380
									SCpCode = "CP50107";
									break;
								case "40":
									//휘팅바렐
									SCpCode = "CP30112";
									break;
								case "50":
									//휘팅포장
									SCpCode = "CP30114";
									break;
								case "60":
								case "70":
								case "80":
									oForm.Items.Item("CpCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
									break;
								default:
									// 콤보박스에 첫데이타를 SD380로
									oForm.Items.Item("CpCode").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
									SCpCode = oForm.Items.Item("CpCode").Specific.Value.ToString().Trim();

									sQry = "SELECT U_CpName From [@PS_PP001L] Where U_CpCode = '" + SCpCode + "' Order by Code";
									oRecordSet.DoQuery(sQry);

									oForm.Items.Item("CpName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
									break;
							}

							if (oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim() != "60" && oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim() != "70" 
								&& oForm.Items.Item("CpGbn").Specific.Value.ToString().Trim() != "80")
							{
								oForm.Items.Item("CpCode").Specific.Select(SCpCode, SAPbouiCOM.BoSearchKey.psk_ByValue);
								sQry = "SELECT U_CpName From [@PS_PP001L] Where U_CpCode = '" + SCpCode + "' Order by Code";
								oRecordSet.DoQuery(sQry);

								oForm.Items.Item("CpName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
							}
						}
					}
					else if (pVal.ItemUID == "OrdType")
					{
						if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() == "10")
						{
							// 실동입력
							oForm.Items.Item("CpGbn").Enabled = true;
							oForm.Items.Item("CpCode").Enabled = true;
							oForm.Items.Item("ItmBsort").Enabled = true;
							oMat.Columns.Item("NCode").Editable = false;
						}
						else
						{
							// 비가동입력
							oForm.Items.Item("CpGbn").Enabled = false;
							oForm.Items.Item("CpGbn").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("CpCode").Enabled = false;
							oForm.Items.Item("CpCode").Specific.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
							oForm.Items.Item("ItmBsort").Enabled = false;
							oMat.Columns.Item("NCode").Editable = true;
						}
					}
					else if (pVal.ItemUID == "CpCode") // 공정코드
					{
						sQry = "SELECT U_CpName From [@PS_PP001L] Where U_CpCode = '" + oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						oForm.Items.Item("CpName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					}
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
						if (pVal.Row == -1)
						{
						}
						else
						{
							PS_PP060_OpenItemRegist(pVal.Row);
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
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CntcCode")
						{
							FlushToItemValue(pVal.ItemUID, 0, "");
						}
						else if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "CntcCode")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							else if (pVal.ColUID == "WorkTime")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							else if (pVal.ColUID == "FixCode")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							else if (pVal.ColUID == "WorkNote")
							{
								FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
							}
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
							}
							else
							{
								oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							}
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP060H);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
						oDS_PS_PP060H.RemoveRecord(oDS_PS_PP060H.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							Add_MatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_PP060H.GetValue("U_ColReg07", oMat.RowCount - 1).ToString().Trim()))
							{
								Add_MatrixRow(oMat.RowCount, false);
							}
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							//추가버튼 클릭시 메트릭스 insertrow
							oMat.Clear();
							oMat.FlushToDataSource();
							oMat.LoadFromDataSource();

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							Add_MatrixRow(0, true);
							BubbleEvent = false;
							LoadCaption();
							break;
						case "1285": //복원
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1287": //복제
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "7169": //엑셀 내보내기
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
