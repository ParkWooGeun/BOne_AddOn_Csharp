using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 기성매출현황조회
	/// </summary>
	internal class PS_SD056 : PSH_BaseClass
	{
		private string oFormUniqueID;
		public SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_SD056L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oLast_Mode;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD056.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD056_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD056");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_SD056_CreateItems();
				PS_SD056_ComboBox_Setting();
				PS_SD056_FormResize();
				PS_SD056_LoadCaption();
				PS_SD056_Initial_Setting();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
		/// PS_SD056_CreateItems
		/// </summary>
		private void PS_SD056_CreateItems()
		{
			try
			{
				oDS_PS_SD056L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//입력정보
				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM").Specific.DataBind.SetBound(true, "", "StdYM");

				//원가계산실행전
				oForm.DataSources.UserDataSources.Add("CostExBf", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("CostExBf").Specific.DataBind.SetBound(true, "", "CostExBf");

				//조회정보
				//작번
				oForm.DataSources.UserDataSources.Add("OrdNumS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNumS").Specific.DataBind.SetBound(true, "", "OrdNumS");

				//품명
				oForm.DataSources.UserDataSources.Add("FrgnNameS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("FrgnNameS").Specific.DataBind.SetBound(true, "", "FrgnNameS");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYMS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYMS").Specific.DataBind.SetBound(true, "", "StdYMS");

				//출력구분
				oForm.DataSources.UserDataSources.Add("PrintOpt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("PrintOpt").Specific.DataBind.SetBound(true, "", "PrintOpt");

				//당월기성매출계
				oForm.DataSources.UserDataSources.Add("CurSlTotal", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("CurSlTotal").Specific.DataBind.SetBound(true, "", "CurSlTotal");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD056_ComboBox_Setting
		/// </summary>
		private void PS_SD056_ComboBox_Setting()
		{
			try
			{
				//출력구분
				oForm.Items.Item("PrintOpt").Specific.ValidValues.Add("01", "전체 리스트");
				oForm.Items.Item("PrintOpt").Specific.ValidValues.Add("02", "작번별 리스트");
				oForm.Items.Item("PrintOpt").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD056_FormResize
		/// </summary>
		private void PS_SD056_FormResize()
		{
			try
			{
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD056_LoadCaption
		/// Form의 Mode에 따라 추가, 확인, 갱신 버튼 이름 변경
		/// </summary>
		private void PS_SD056_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "확정등록";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("BtnAdd").Specific.Caption = "수정";
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
		/// PS_SD056_Initial_Setting
		/// </summary>
		private void PS_SD056_Initial_Setting()
		{
			try
			{
				oMat.Columns.Item("StdYM").Visible = false;

				oForm.Items.Item("StdYM").Specific.Value = DateTime.Now.ToString("yyyyMM");
				oForm.Items.Item("StdYMS").Specific.Value = DateTime.Now.ToString("yyyyMM");
				oForm.Items.Item("StdYM").Click();
				oForm.Items.Item("CostExBf").Specific.Checked = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD056_CheckAll
		/// </summary>
		private void PS_SD056_CheckAll()
		{
			string CheckType;
			int loopCount;
			double CurSlTotal = 0;

			try
			{
				oForm.Freeze(true);
				CheckType = "Y";

				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_SD056L.GetValue("U_ColReg08", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break; // TODO: might not be correct. Was : Exit For
					}
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_SD056L.Offset = loopCount;
					if (CheckType == "N")
					{

						oDS_PS_SD056L.SetValue("U_ColReg08", loopCount, "Y");
						CurSlTotal += System.Math.Round(Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum10", loopCount).ToString().Trim()), 0);
					}
					else
					{
						oDS_PS_SD056L.SetValue("U_ColReg08", loopCount, "N");
						CurSlTotal = 0;
					}
				}
				oMat.LoadFromDataSource();
				oForm.Items.Item("CurSlTotal").Specific.Value = CurSlTotal;
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
		/// PS_SD056_CalculateCurSalseTotal
		/// 당월기성매출합계 계산
		/// </summary>
		private void PS_SD056_CalculateCurSalseTotal()
		{
			int loopCount;
			double CurSlTotal = 0;

			try
			{
				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_SD056L.GetValue("U_ColReg08", loopCount).ToString().Trim() == "Y")
					{
						CurSlTotal += System.Math.Round(Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum10", loopCount).ToString().Trim()), 0);
					}
				}
				oForm.Items.Item("CurSlTotal").Specific.Value = CurSlTotal;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD056_CheckBeforeSearch
		/// 필수입력사항 체크
		/// </summary>
		/// <param name="pItemUID"></param>
		/// <returns></returns>
		private bool PS_SD056_CheckBeforeSearch(string pItemUID)
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (pItemUID == "BtnSearch1")
				{
					if (string.IsNullOrEmpty(oForm.Items.Item("StdYM").Specific.Value.ToString().Trim()))
					{
						errMessage = "입력정보의 기준년월은 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
				}
				else if (pItemUID == "BtnSearch2")
				{
					if (string.IsNullOrEmpty(oForm.Items.Item("StdYMS").Specific.Value.ToString().Trim()))
					{
						errMessage = "조회정보의 기준년월은 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
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
		/// PS_SD056_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_SD056_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			double PP030Amt; //품의금액
			double ResRate;  //진도율
			double Amount;   //금액
			short loopCount;
			double TotalAmt = 0; //금액 합계
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "Mat01":
						if (oCol == "ResRate")
						{
							oMat.FlushToDataSource();

							PP030Amt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum01", oRow - 1).ToString().Trim());
							ResRate = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColQty01", oRow - 1).ToString().Trim());
							Amount = PP030Amt * ResRate / 100;
							oDS_PS_SD056L.SetValue("U_ColSum02", oRow - 1, Convert.ToString(Amount));

							for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
							{
								TotalAmt += Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum02", loopCount).ToString().Trim());
							}
							oForm.Items.Item("Total").Specific.Value = TotalAmt;
							oMat.LoadFromDataSource();
						}
						oMat.AutoResizeColumns();
						break;

					case "OrdNum":
						oForm.Items.Item("FrgnName").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
						break;

					case "OrdNumS":
						oForm.Items.Item("FrgnNameS").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(oUID).Specific.Value.ToString().Trim() + "'", "");
						break;

				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD056_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_SD056_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_SD056L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_SD056L.Offset = oRow;
				oDS_PS_SD056L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD056_MTX01
		/// 데이터 조회
		/// </summary>
		/// <param name="pItemUID"></param>
		private void PS_SD056_MTX01(string pItemUID)
		{
			int i;
			string sQry;
			string errMessage = string.Empty;

			string StdYM;       //기준년월
			string CostExBf;    //원가계산실행전
			string OrdNumS;     //작번(조회)
			string StdYMS;      //기준년월(조회)
			string CntcCode;    //사용자 사번
			double CurSlTotal = 0;  //당월기성매출계

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pItemUID == "BtnSearch1")
				{
					StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();
					CntcCode = dataHelpClass.User_MSTCOD();
					if (oForm.Items.Item("CostExBf").Specific.Checked == true)
					{
						CostExBf = "Y";
					}
					else
					{
						CostExBf = "N";
					}

					sQry = " EXEC [PS_SD056_01] '";
					sQry += StdYM + "','";
					sQry += CostExBf + "'";
					oRecordSet.DoQuery(sQry);

					oMat.Clear();
					oDS_PS_SD056L.Clear();
					oMat.FlushToDataSource();
					oMat.LoadFromDataSource();

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_SD056_LoadCaption();
						errMessage = "조회 결과가 없습니다. 확인하세요.";
						throw new Exception();
					}

					for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
					{
						if (i + 1 > oDS_PS_SD056L.Size)
						{
							oDS_PS_SD056L.InsertRecord(i);
						}

						oMat.AddRow();
						oDS_PS_SD056L.Offset = i;

						oDS_PS_SD056L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
						oDS_PS_SD056L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());     //선택
						oDS_PS_SD056L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("StdYM").Value.ToString().Trim());     //기준년월
						oDS_PS_SD056L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());    //작번
						oDS_PS_SD056L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("FrgnName").Value.ToString().Trim());  //품명
						oDS_PS_SD056L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("Spec").Value.ToString().Trim());      //규격
						oDS_PS_SD056L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("InOut").Value.ToString().Trim());     //자체/외주
						oDS_PS_SD056L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());  //수주처코드
						oDS_PS_SD056L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());  //수주처명
						oDS_PS_SD056L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //수주일
						oDS_PS_SD056L.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //납기일
						oDS_PS_SD056L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("LineTotal").Value.ToString().Trim()); //수주금액
						oDS_PS_SD056L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("InAmt").Value.ToString().Trim());     //당월누계원가
						oDS_PS_SD056L.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("SD055Amt").Value.ToString().Trim());  //기성매입
						oDS_PS_SD056L.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("SD054Amt").Value.ToString().Trim());  //추가예상공수
						oDS_PS_SD056L.SetValue("U_ColSum05", i, oRecordSet.Fields.Item("AddAmt").Value.ToString().Trim());    //추가자재비
						oDS_PS_SD056L.SetValue("U_ColSum06", i, oRecordSet.Fields.Item("AddCostAmt").Value.ToString().Trim());//추가경비계
						oDS_PS_SD056L.SetValue("U_ColSum07", i, oRecordSet.Fields.Item("TCostAmt").Value.ToString().Trim());  //총공사예정비
						oDS_PS_SD056L.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("IngRate").Value.ToString().Trim());   //진척도
						oDS_PS_SD056L.SetValue("U_ColSum08", i, oRecordSet.Fields.Item("CurTSlAmt").Value.ToString().Trim()); //당월누계기성매출
						oDS_PS_SD056L.SetValue("U_ColSum09", i, oRecordSet.Fields.Item("PreTSlAmt").Value.ToString().Trim()); //전월누계기성매출
						oDS_PS_SD056L.SetValue("U_ColSum10", i, oRecordSet.Fields.Item("CurSlAmt").Value.ToString().Trim());  //당월기성매출
						oDS_PS_SD056L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());    //작번
						oDS_PS_SD056L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("OKYN").Value.ToString().Trim());      //수주상세금액승인(여부)

						CurSlTotal += System.Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("CurSlAmt").Value.ToString().Trim()), 0);

						oRecordSet.MoveNext();
						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
					}

					oForm.Items.Item("CurSlTotal").Specific.Value = CurSlTotal;
					oMat.LoadFromDataSource();
					oMat.AutoResizeColumns();
				}
				else if (pItemUID == "BtnSearch2")
				{
					OrdNumS = oForm.Items.Item("OrdNumS").Specific.Value.ToString().Trim();
					StdYMS = oForm.Items.Item("StdYMS").Specific.Value.ToString().Trim();
					CntcCode = dataHelpClass.User_MSTCOD();

					sQry = " EXEC [PS_SD056_02] '";
					sQry += OrdNumS + "','";
					sQry += StdYMS + "'";
					oRecordSet.DoQuery(sQry);

					oMat.Clear();
					oDS_PS_SD056L.Clear();
					oMat.FlushToDataSource();
					oMat.LoadFromDataSource();

					if (oRecordSet.RecordCount == 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_SD056_LoadCaption();
						errMessage = "조회 결과가 없습니다. 확인하세요.";
						throw new Exception();
					}

					for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
					{
						if (i + 1 > oDS_PS_SD056L.Size)
						{
							oDS_PS_SD056L.InsertRecord(i);
						}

						oMat.AddRow();
						oDS_PS_SD056L.Offset = i;

						oDS_PS_SD056L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
						oDS_PS_SD056L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());                        //선택
						oDS_PS_SD056L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("StdYM").Value.ToString().Trim());                        //기준년월
						oDS_PS_SD056L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());                       //작번
						oDS_PS_SD056L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("FrgnName").Value.ToString().Trim());                     //품명
						oDS_PS_SD056L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("Spec").Value.ToString().Trim());                     //규격
						oDS_PS_SD056L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("InOut").Value.ToString().Trim());                        //자체/외주
						oDS_PS_SD056L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());                     //수주처코드
						oDS_PS_SD056L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());                     //수주처명
						oDS_PS_SD056L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //수주일
						oDS_PS_SD056L.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //납기일
						oDS_PS_SD056L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("LineTotal").Value.ToString().Trim());                        //수주금액
						oDS_PS_SD056L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("InAmt").Value.ToString().Trim());                        //당월누계원가
						oDS_PS_SD056L.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("SD055Amt").Value.ToString().Trim());                     //기성매입
						oDS_PS_SD056L.SetValue("U_ColSum04", i, oRecordSet.Fields.Item("SD054Amt").Value.ToString().Trim());                     //추가예상공수
						oDS_PS_SD056L.SetValue("U_ColSum05", i, oRecordSet.Fields.Item("AddAmt").Value.ToString().Trim());                       //추가자재비
						oDS_PS_SD056L.SetValue("U_ColSum06", i, oRecordSet.Fields.Item("AddCostAmt").Value.ToString().Trim());                       //추가경비계
						oDS_PS_SD056L.SetValue("U_ColSum07", i, oRecordSet.Fields.Item("TotalCostAmt").Value.ToString().Trim());                     //총공사예정비
						oDS_PS_SD056L.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("IngRate").Value.ToString().Trim());                      //진척도
						oDS_PS_SD056L.SetValue("U_ColSum08", i, oRecordSet.Fields.Item("CurTotalSalesAmt").Value.ToString().Trim());                     //당월누계기성매출
						oDS_PS_SD056L.SetValue("U_ColSum09", i, oRecordSet.Fields.Item("PreTotalSalesAmt").Value.ToString().Trim());                     //전월누계기성매출
						oDS_PS_SD056L.SetValue("U_ColSum10", i, oRecordSet.Fields.Item("CurSalesAmt").Value.ToString().Trim());                      //당월기성매출
						oDS_PS_SD056L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());                       //작번
						oDS_PS_SD056L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("OKYN").Value.ToString().Trim());                     //수주상세금액승인(여부)

						CurSlTotal += System.Math.Round(Convert.ToDouble(oRecordSet.Fields.Item("CurSalesAmt").Value.ToString().Trim()), 0);

						oRecordSet.MoveNext();
						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";

					}
					oForm.Items.Item("CurSlTotal").Specific.Value = CurSlTotal;

					oMat.LoadFromDataSource();
					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				ProgressBar01.Stop();  //stop 안하면 오래결림.

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
		/// PS_SD056_AddData
		/// 데이터 INSERT
		/// </summary>
		/// <returns></returns>
		private bool PS_SD056_AddData()
		{
			bool functionReturnValue = false;

			int loopCount;
			string sQry;

			string StdYM;		//기준년월
			string OrdNum;		//작번
			string FrgnName;	//품명
			string SPEC;		//규격
			string InOut;		//자체/외주
			string CardCode;	//수주처코드
			string CardName;	//수주처명
			string DocDate;		//수주일
			string DueDate;		//납기일
			double LineTotal;	//수주금액
			double InAmt;		//당월누계원가
			double SD055Amt;	//기성매입
			double SD054Amt;	//추가예상공수
			double AddAmt;		//추가자재비
			double AddCostAmt;	//추가경비계
			double TCostAmt;	//총공사예정비
			double IngRate;		//진척도
			double CurTSlAmt;	//당월누계기성매출
			double PreTSlAmt;	//전월누계기성매출
			double CurSlAmt;	//당월기성매출
			string CntcCode;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				CntcCode = dataHelpClass.User_MSTCOD(); //사용자사번

				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{

					if (oDS_PS_SD056L.GetValue("U_ColReg08", loopCount).ToString().Trim() == "Y")
					{

						StdYM    = oDS_PS_SD056L.GetValue("U_ColReg01", loopCount).ToString().Trim();
						OrdNum   = oDS_PS_SD056L.GetValue("U_ColReg02", loopCount).ToString().Trim();
						FrgnName = oDS_PS_SD056L.GetValue("U_ColReg03", loopCount).ToString().Trim();
						SPEC     = oDS_PS_SD056L.GetValue("U_ColReg04", loopCount).ToString().Trim();
						InOut    = oDS_PS_SD056L.GetValue("U_ColReg07", loopCount).ToString().Trim();
						CardCode = oDS_PS_SD056L.GetValue("U_ColReg05", loopCount).ToString().Trim();
						CardName = oDS_PS_SD056L.GetValue("U_ColReg06", loopCount).ToString().Trim();
						DocDate  = oDS_PS_SD056L.GetValue("U_ColDt01", loopCount).ToString().Trim();
						DueDate  = oDS_PS_SD056L.GetValue("U_ColDt02", loopCount).ToString().Trim();
						LineTotal = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum01", loopCount).ToString().Trim());
						InAmt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum02", loopCount).ToString().Trim());
						SD055Amt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum03", loopCount).ToString().Trim());
						SD054Amt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum04", loopCount).ToString().Trim());
						AddAmt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum05", loopCount).ToString().Trim());
						AddCostAmt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum06", loopCount).ToString().Trim());
						TCostAmt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum07", loopCount).ToString().Trim());
						IngRate = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColQty02", loopCount).ToString().Trim());
						CurTSlAmt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum08", loopCount).ToString().Trim());
						PreTSlAmt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum09", loopCount).ToString().Trim());
						CurSlAmt = Convert.ToDouble(oDS_PS_SD056L.GetValue("U_ColSum10", loopCount).ToString().Trim());

						sQry = " EXEC [PS_SD056_03] '";
						sQry += StdYM + "','";
						sQry += OrdNum + "','";
						sQry += FrgnName + "','";
						sQry += SPEC + "','";
						sQry += InOut + "','";
						sQry += CardCode + "','";
						sQry += CardName + "','";
						sQry += DocDate + "','";
						sQry += DueDate + "','";
						sQry += LineTotal + "','";
						sQry += InAmt + "','";
						sQry += SD055Amt + "','";
						sQry += SD054Amt + "','";
						sQry += AddAmt + "','";
						sQry += AddCostAmt + "','";
						sQry += TCostAmt + "','";
						sQry += IngRate + "','";
						sQry += CurTSlAmt + "','";
						sQry += PreTSlAmt + "','";
						sQry += CurSlAmt + "'";
						oRecordSet.DoQuery(sQry);
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// PS_SD056_DeleteDataAll
		/// 전체 삭제
		/// </summary>
		private void PS_SD056_DeleteDataAll()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;
			string StdYM; //기준년월

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oMat.VisualRowCount == 0)
				{
					errMessage = "삭제대상이 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					StdYM = oDS_PS_SD056L.GetValue("U_ColReg01", loopCount).ToString().Trim();
					sQry = " EXEC [PS_SD056_06] '";
					sQry += StdYM + "'";
					oRecordSet.DoQuery(sQry);
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("전체 삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		}

		/// <summary>
		/// PS_SD056_DeleteDataCheck
		/// 선택 삭제
		/// </summary>
		private void PS_SD056_DeleteDataCheck()
		{
			int	loopCount;
			string sQry;
			string errMessage = string.Empty;

			string StdYM;	//기준년월
			string OrdNum;  //작번

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oMat.VisualRowCount == 0)
				{
					errMessage = "삭제대상이 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_SD056L.GetValue("U_ColReg08", loopCount).ToString().Trim() == "Y")
					{
						StdYM  = oDS_PS_SD056L.GetValue("U_ColReg01", loopCount).ToString().Trim();
						OrdNum = oDS_PS_SD056L.GetValue("U_ColReg02", loopCount).ToString().Trim();

						sQry = " EXEC [PS_SD056_07] '";
						sQry +=  StdYM + "','";
						sQry +=  OrdNum + "'";
						oRecordSet.DoQuery(sQry);
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("선택 삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		}

		/// <summary>
		/// PS_SD056_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_SD056_Print_Report01()
		{
			int i;
			string WinTitle;
			string ReportName = string.Empty;
			string sQry;

			string OrdNumS;	 //작번(조회)
			string StdYMS;	 //기준년월(조회)
			string PrintOpt; //출력구분
			string CntcCode;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				OrdNumS  = oForm.Items.Item("OrdNumS").Specific.Value.ToString().Trim();
				StdYMS   = oForm.Items.Item("StdYMS").Specific.Value.ToString().Trim();
				PrintOpt = oForm.Items.Item("PrintOpt").Specific.Value.ToString().Trim();
				CntcCode = dataHelpClass.User_MSTCOD(); //조회자사번

				//임시테이블에 check된항목저장하기 위한 기 정보 삭제
				sQry = "DELETE [Z_PS_SD056_02] WHERE CntcCode = '" + CntcCode + "' AND StdYM = '" + StdYMS + "'";
				oRecordSet.DoQuery(sQry);

				//임시테이블에 check된항목저장
				oMat.FlushToDataSource();
				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					if (oDS_PS_SD056L.GetValue("U_ColReg08", i).ToString().Trim() == "Y")
					{
						sQry = "INSERT INTO [Z_PS_SD056_02] VALUES ('" + CntcCode + "', '" + oDS_PS_SD056L.GetValue("U_ColReg01", i).ToString().Trim() + "','" + oDS_PS_SD056L.GetValue("U_ColReg02", i).ToString().Trim() + "')";
						oRecordSet.DoQuery(sQry);
					}
				}

				WinTitle = "[PS_SD056] 기성매출";

				if (PrintOpt == "01")
				{
					ReportName = "PS_SD056_01.rpt";
				}
				else if (PrintOpt == "02")
				{
					ReportName = "PS_SD056_02.rpt";
				}

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@CntcCode", CntcCode));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYMS));
				dataPackParameter.Add(new PSH_DataPackClass("@OrdNum", OrdNumS));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
					if (pVal.ItemUID == "BtnAdd")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_SD056_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD056_LoadCaption();
							oLast_Mode = Convert.ToInt32(oForm.Mode);
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_SD056_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							PS_SD056_MTX01("BtnSearch2");
							PS_SD056_LoadCaption();
						}
					}
					else if (pVal.ItemUID == "BtnSearch1")
					{
						if (PS_SD056_CheckBeforeSearch(pVal.ItemUID) == false)
						{
							BubbleEvent = false;
							return;
						}
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_SD056_LoadCaption();
						PS_SD056_MTX01(pVal.ItemUID);
					}
					else if (pVal.ItemUID == "BtnSearch2")
					{
						if (PS_SD056_CheckBeforeSearch(pVal.ItemUID) == false)
						{
							BubbleEvent = false;
							return;
						}
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						PS_SD056_LoadCaption();
						PS_SD056_MTX01(pVal.ItemUID);
					}
					else if (pVal.ItemUID == "BtnAllDel")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 전체 기성매출 정보를 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
						{
							PS_SD056_DeleteDataAll();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD056_LoadCaption();
						}
					}
					else if (pVal.ItemUID == "BtnChkDel")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 선택한 작번의 기성매출 정보를 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
						{
							PS_SD056_DeleteDataCheck();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD056_LoadCaption();
						}
					}
					else if (pVal.ItemUID == "BtnChk")
					{
						PS_SD056_CheckAll();
					}
					else if (pVal.ItemUID == "BtnPrint")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_SD056_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
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
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat.SelectRow(pVal.Row, true, false);
						}
						if (pVal.ColUID == "Check")
						{
							//선택한 작번의 당월기성매출 합계 계산
							PS_SD056_CalculateCurSalseTotal();
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_SD056_FlushToItemValue(pVal.ItemUID, 0, "");
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
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							PS_SD056_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else
						{
							PS_SD056_FlushToItemValue(pVal.ItemUID, 0, "");
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
		private void Raise_EVENT_FORM_RESIZE(string FormUID, SAPbouiCOM.ItemEvent pVal, bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_SD056_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD056L);
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
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							BubbleEvent = false;
							PS_SD056_LoadCaption();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							break;
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							PS_SD056_Add_MatrixRow(oMat.VisualRowCount, false);
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
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							oDS_PS_SD056L.RemoveRecord(oDS_PS_SD056L.Size - 1);
							oMat.LoadFromDataSource();
							oForm.Freeze(false);
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
