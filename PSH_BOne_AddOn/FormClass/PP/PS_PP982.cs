using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 반별가동률 조회(개별, 집계)
	/// </summary>
	internal class PS_PP982 : PSH_BaseClass
	{
		private string oFormUniqueID;

		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Grid oGrid03;
		private SAPbouiCOM.Matrix oMat01;

		private SAPbouiCOM.DataTable oDS_PS_PP982L;
		private SAPbouiCOM.DataTable oDS_PS_PP982M;
		private SAPbouiCOM.DataTable oDS_PS_PP982N;
		private SAPbouiCOM.DBDataSource oDS_PS_PP982O;

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private short CheckBoxCount; //반 체크박스 최대수량

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry01"></param>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP982.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP982_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP982");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP982_CreateItems();
				PS_PP982_ComboBox_Setting();
				PS_PP982_FormResize();
				PS_PP982_FormReset();

				oForm.EnableMenu(("1283"), false); // 삭제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1287"), false); // 복제
				oForm.EnableMenu(("1285"), false); // 복원
				oForm.EnableMenu(("1284"), true);  // 취소
				oForm.EnableMenu(("1293"), false); // 행삭제
				oForm.EnableMenu(("1281"), false);
				oForm.EnableMenu(("1282"), true);
				CheckBoxCount = 12;	//반 체크박스 최대 수량 초기화
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Items.Item("Folder01").Specific.Select();
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP982_CreateItems
		/// </summary>
		private void PS_PP982_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;
				oGrid03 = oForm.Items.Item("Grid03").Specific;

				oForm.DataSources.DataTables.Add("PS_PP982L");
				oForm.DataSources.DataTables.Add("PS_PP982M");
				oForm.DataSources.DataTables.Add("PS_PP982N");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_PP982L");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_PP982M");
				oGrid03.DataTable = oForm.DataSources.DataTables.Item("PS_PP982N");

				oDS_PS_PP982L = oForm.DataSources.DataTables.Item("PS_PP982L");
				oDS_PS_PP982M = oForm.DataSources.DataTables.Item("PS_PP982M");
				oDS_PS_PP982N = oForm.DataSources.DataTables.Item("PS_PP982N");

				oDS_PS_PP982O = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				//// 메트릭스 개체 할당
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat01.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//일자Fr
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//일자To
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//사번
				oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

				//성명
				oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

				//반 체크박스
				oForm.DataSources.UserDataSources.Add("Check01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check01").Specific.DataBind.SetBound(true, "", "Check01");

				oForm.DataSources.UserDataSources.Add("Check02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check02").Specific.DataBind.SetBound(true, "", "Check02");

				oForm.DataSources.UserDataSources.Add("Check03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check03").Specific.DataBind.SetBound(true, "", "Check03");

				oForm.DataSources.UserDataSources.Add("Check04", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check04").Specific.DataBind.SetBound(true, "", "Check04");

				oForm.DataSources.UserDataSources.Add("Check05", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check05").Specific.DataBind.SetBound(true, "", "Check05");

				oForm.DataSources.UserDataSources.Add("Check06", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check06").Specific.DataBind.SetBound(true, "", "Check06");

				oForm.DataSources.UserDataSources.Add("Check07", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check07").Specific.DataBind.SetBound(true, "", "Check07");

				oForm.DataSources.UserDataSources.Add("Check08", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check08").Specific.DataBind.SetBound(true, "", "Check08");

				oForm.DataSources.UserDataSources.Add("Check09", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check09").Specific.DataBind.SetBound(true, "", "Check09");

				oForm.DataSources.UserDataSources.Add("Check10", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check10").Specific.DataBind.SetBound(true, "", "Check10");

				oForm.DataSources.UserDataSources.Add("Check11", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check11").Specific.DataBind.SetBound(true, "", "Check11");

				oForm.DataSources.UserDataSources.Add("Check12", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("Check12").Specific.DataBind.SetBound(true, "", "Check12");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP982_ComboBox_Setting
		/// </summary>
		private void PS_PP982_ComboBox_Setting()
		{
			string BPLID;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				BPLID = dataHelpClass.User_BPLID();
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP982_FormResize
		/// </summary>
		private void PS_PP982_FormResize()
		{
			try
			{
				oForm.Freeze(true);

				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 25;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

				if (oGrid01.Columns.Count > 0)
				{
					oGrid01.AutoResizeColumns();
				}

				if (oGrid02.Columns.Count > 0)
				{
					oGrid02.AutoResizeColumns();
				}

				if (oGrid03.Columns.Count > 0)
				{
					oGrid03.AutoResizeColumns();
				}

				oMat01.AutoResizeColumns();
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
		/// PS_PP982_FormReset
		/// </summary>
		private void PS_PP982_FormReset()
		{
			string User_BPLId;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				User_BPLId = dataHelpClass.User_BPLID();

				//헤더 초기화
				oForm.DataSources.UserDataSources.Item("FrDt").Value = DateTime.Now.ToString("yyyyMM") + "01"; 
				oForm.DataSources.UserDataSources.Item("ToDt").Value = DateTime.Now.ToString("yyyyMMdd");

				//라인 초기화
				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();
				PS_PP982_Add_MatrixRow01(0, true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP982_Add_MatrixRow01
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP982_Add_MatrixRow01(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP982O.InsertRecord(oRow);
				}

				oMat01.AddRow();
				oDS_PS_PP982O.Offset = oRow;
				oDS_PS_PP982O.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat01.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP982_SelectGrid01  반별자료 조회
		/// </summary>
		private void PS_PP982_SelectGrid01()
		{
			short loopCount;
			string sQry;
			string errMessage = String.Empty;

			string BPLID;
			string FrDt;
			string ToDt;
			string CntcCode;
			string ClsCodeChk;
			string CheckBoxID;
		
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrDt     = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt     = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				ClsCodeChk = "";

				for (loopCount = 1; loopCount <= CheckBoxCount; loopCount++)
				{
					if (loopCount >= 1 & loopCount < 10)
					{
						CheckBoxID = "Check0" + loopCount;
					}
					else
					{
						CheckBoxID = "Check" + loopCount;
					}

					if (oForm.DataSources.UserDataSources.Item(CheckBoxID).Value == "Y")
					{
						ClsCodeChk += codeHelpClass.Left(codeHelpClass.Right(oForm.Items.Item(CheckBoxID).Specific.Caption.Split('-')[0], 5), 4) + ",";
					}
				}

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP982_01 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CntcCode + "','";
				sQry += ClsCodeChk + "'";

				oGrid01.DataTable.Clear();
				oDS_PS_PP982L.ExecuteQuery(sQry);

				oGrid01.Columns.Item(7).RightJustified = true;
				oGrid01.Columns.Item(8).RightJustified = true;
				oGrid01.Columns.Item(9).RightJustified = true;
				oGrid01.Columns.Item(10).RightJustified = true;
				oGrid01.Columns.Item(11).RightJustified = true;
				oGrid01.Columns.Item(12).RightJustified = true;
				oGrid01.Columns.Item(13).RightJustified = true;
				oGrid01.Columns.Item(14).RightJustified = true;

				if (oGrid01.Rows.Count == 1)
				{
					errMessage = "반별자료 결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid01.AutoResizeColumns();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP982_SelectGrid02  집계자료(주별)
		/// </summary>
		private void PS_PP982_SelectGrid02()
		{
			short loopCount;
			string sQry;
			string errMessage = String.Empty;

			string BPLID;
			string FrDt;
			string ToDt;
			string CntcCode;
			string ClsCodeChk;
			string CheckBoxID;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				ClsCodeChk = "";

				for (loopCount = 1; loopCount <= CheckBoxCount; loopCount++)
				{
					if (loopCount >= 1 & loopCount < 10)
					{
						CheckBoxID = "Check0" + loopCount;
					}
					else
					{
						CheckBoxID = "Check" + loopCount;
					}

					if (oForm.DataSources.UserDataSources.Item(CheckBoxID).Value == "Y")
					{
						ClsCodeChk += codeHelpClass.Left(codeHelpClass.Right(oForm.Items.Item(CheckBoxID).Specific.Caption.Split('-')[0], 5), 4) + ",";
					}
				}

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP982_02 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CntcCode + "','";
				sQry += ClsCodeChk + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_PP982M.ExecuteQuery(sQry);

				oGrid02.Columns.Item(2).RightJustified = true;
				oGrid02.Columns.Item(3).RightJustified = true;
				oGrid02.Columns.Item(4).RightJustified = true;
				oGrid02.Columns.Item(5).RightJustified = true;
				oGrid02.Columns.Item(6).RightJustified = true;
				oGrid02.Columns.Item(7).RightJustified = true;
				oGrid02.Columns.Item(8).RightJustified = true;

				if (oGrid02.Rows.Count == 1)
				{
					errMessage = "집계자료(주별) 결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid02.AutoResizeColumns();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP982_SelectGrid03 집계자료(일별)
		/// </summary>
		private void PS_PP982_SelectGrid03()
		{
			short loopCount;
			string sQry;
			string errMessage = String.Empty;

			string BPLID;
			string FrDt;
			string ToDt;
			string CntcCode;
			string ClsCodeChk;
			string CheckBoxID;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				ClsCodeChk = "";

				for (loopCount = 1; loopCount <= CheckBoxCount; loopCount++)
				{
					if (loopCount >= 1 & loopCount < 10)
					{
						CheckBoxID = "Check0" + loopCount;
					}
					else
					{
						CheckBoxID = "Check" + loopCount;
					}

					if (oForm.DataSources.UserDataSources.Item(CheckBoxID).Value == "Y")
					{
						ClsCodeChk += codeHelpClass.Left(codeHelpClass.Right(oForm.Items.Item(CheckBoxID).Specific.Caption.Split('-')[0], 5), 4) + ",";
					}
				}

				sQry = " EXEC PS_PP982_03 '";
				sQry += BPLID + "','";
				sQry += FrDt + "','";
				sQry += ToDt + "','";
				sQry += CntcCode + "','";
				sQry += ClsCodeChk + "'";

				oGrid03.DataTable.Clear();
				oDS_PS_PP982N.ExecuteQuery(sQry);
				//    oGrid03.DataTable = oForm.DataSources.DataTables.Item("DataTable")

				oGrid03.Columns.Item(3).RightJustified = true;
				oGrid03.Columns.Item(4).RightJustified = true;
				oGrid03.Columns.Item(5).RightJustified = true;
				oGrid03.Columns.Item(6).RightJustified = true;
				oGrid03.Columns.Item(7).RightJustified = true;
				oGrid03.Columns.Item(8).RightJustified = true;

				if (oGrid03.Rows.Count == 1)
				{
					errMessage = "집계자료(일별) 결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid03.AutoResizeColumns();
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
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP982_SelectMatrix01  예외작번 관리 조회
		/// </summary>
		private void PS_PP982_SelectMatrix01()
		{
			short i;
			string sQry;
			string errMessage = String.Empty;

			string BPLID;  //사업장

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC [PS_PP982_04] '";
				sQry += BPLID + "'";

				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_PP982O.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP982_Add_MatrixRow01(0, true);
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP982O.Size)
					{
						oDS_PS_PP982O.InsertRecord(i);
					}

					oMat01.AddRow();
					oDS_PS_PP982O.Offset = i;

					oDS_PS_PP982O.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP982O.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Select").Value.ToString().Trim());   //선택
					oDS_PS_PP982O.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("CntcCode").Value.ToString().Trim()); //사번
					oDS_PS_PP982O.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim()); //성명
					oDS_PS_PP982O.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("TeamName").Value.ToString().Trim()); //팀
					oDS_PS_PP982O.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("RspName").Value.ToString().Trim());  //담당
					oDS_PS_PP982O.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("ClsName").Value.ToString().Trim());  //반
					oDS_PS_PP982O.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("Comments").Value.ToString().Trim()); //비고

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";

				}

				PS_PP982_Add_MatrixRow01(oMat01.VisualRowCount, false);

				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();

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
		/// PS_PP982_InsertMatrix01
		/// </summary>
		private void PS_PP982_InsertMatrix01()
		{
			short loopCount;
			string sQry;

			string BPLID;	 //사업장
			string CntcCode; //사번
			string CntcName; //성명
			string Comments; //비고

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();

				oMat01.FlushToDataSource();

				ProgressBar01.Text = "저장중...";

				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP982O.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						CntcCode = oDS_PS_PP982O.GetValue("U_ColReg02", loopCount).ToString().Trim(); //사번
						CntcName = oDS_PS_PP982O.GetValue("U_ColReg03", loopCount).ToString().Trim(); //성명
						Comments = oDS_PS_PP982O.GetValue("U_ColReg07", loopCount).ToString().Trim(); //비고

						sQry = " EXEC [PS_PP982_05] '";
						sQry += BPLID + "','";
						sQry += CntcCode + "','";
						sQry += CntcName + "','";
						sQry += Comments + "'";

						oRecordSet.DoQuery(sQry);

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + (oMat01.VisualRowCount - 1) + "건 저장중...";
					}
				}

				PSH_Globals.SBO_Application.MessageBox("저장 완료!");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP982_DeleteMatrix01
		/// </summary>
		private void PS_PP982_DeleteMatrix01()
		{
			short loopCount;
			string sQry;

			string BPLID;    //사업장
			string CntcCode; //사번

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();

				oMat01.FlushToDataSource();

				ProgressBar01.Text = "삭제중...";

				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP982O.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						CntcCode = oDS_PS_PP982O.GetValue("U_ColReg02", loopCount).ToString().Trim();

						sQry = "  EXEC [PS_PP982_06] '";
						sQry += BPLID + "','";
						sQry += CntcCode + "'";

						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.MessageBox("삭제 완료!");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP982_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP982_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			short loopCount01;
			short loopCount02;
			string sQry;
			string BPLID;
			string CheckBoxID;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "Mat01":
						oMat01.FlushToDataSource();

						if (oCol == "CntcCode")
						{
							sQry = " SELECT      T0.Code AS [CntcCode],";		 //사번
							sQry += "             T0.U_FullName AS [CntcName],"; //성명
							sQry += "             T1.U_CodeNm AS [TeamName],";	 //팀
							sQry += "             T2.U_CodeNm AS [RspName],";	 //담당
							sQry += "             T3.U_CodeNm AS [ClsName]";	 //반
							sQry += " FROM        [@PH_PY001A] AS T0";
							sQry += "             LEFT JOIN";
							sQry += "             [@PS_HR200L] AS T1";
							sQry += "                 ON T0.U_TeamCode = T1.U_Code";
							sQry += "                 AND T1.Code = '1'";
							sQry += "             LEFT JOIN";
							sQry += "             [@PS_HR200L] AS T2";
							sQry += "                 ON T0.U_RspCode = T2.U_Code";
							sQry += "                 AND T2.Code = '2'";
							sQry += "             LEFT JOIN";
							sQry += "             [@PS_HR200L] AS T3";
							sQry += "                 ON T0.U_ClsCode = T3.U_Code";
							sQry += "                 AND T3.Code = '9'";
							sQry += " WHERE       T0.Code = '" + oDS_PS_PP982O.GetValue("U_ColReg02", oRow - 1).ToString().Trim() + "'";

							oRecordSet.DoQuery(sQry);

							oDS_PS_PP982O.SetValue("U_ColReg01", oRow - 1, "Y");							                            //선택
							oDS_PS_PP982O.SetValue("U_ColReg02", oRow - 1, oRecordSet.Fields.Item("CntcCode").Value.ToString().Trim());	//사번
							oDS_PS_PP982O.SetValue("U_ColReg03", oRow - 1, oRecordSet.Fields.Item("CntcName").Value.ToString().Trim());	//성명
							oDS_PS_PP982O.SetValue("U_ColReg04", oRow - 1, oRecordSet.Fields.Item("TeamName").Value.ToString().Trim());	//팀
							oDS_PS_PP982O.SetValue("U_ColReg05", oRow - 1, oRecordSet.Fields.Item("RspName").Value.ToString().Trim());	//담당
							oDS_PS_PP982O.SetValue("U_ColReg06", oRow - 1, oRecordSet.Fields.Item("ClsName").Value.ToString().Trim());	//반

							oMat01.LoadFromDataSource();

							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("CntcCode").Cells.Item(oRow).Specific.Value.ToString().Trim()))
								{
									PS_PP982_Add_MatrixRow01(oMat01.RowCount, false);
								}
							}
						}
						oMat01.AutoResizeColumns();
						break;
					case "BPLID":

						BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();

						//반 체크박스 초기화
						for (loopCount01 = 1; loopCount01 <= CheckBoxCount; loopCount01++)
						{
							if (loopCount01 >= 1 & loopCount01 < 10)
							{
								CheckBoxID = "Check0" + loopCount01;
							}
							else
							{
								CheckBoxID = "Check" + loopCount01;
							}
							oForm.Items.Item(CheckBoxID).Specific.Caption = "정보 없음";
						}

						//반 체크박스 제목 바인딩
						sQry = "  SELECT      U_CodeNm + '[' + U_Code +']-' + U_Comment1 + '[' + U_Char1 + ']-' + U_Comment2 + '[' + U_Char2 + ']' AS [FullClsName]";
						sQry += " FROM        [@PS_HR200L]";
						sQry += " WHERE       Code = '9'";
						sQry += "             AND U_Char3 = '" + BPLID + "'";
						sQry += "             AND U_UseYN = 'Y'";
						sQry += " ORDER BY    U_Seq";

						oRecordSet.DoQuery(sQry);

						for (loopCount02 = 1; loopCount02 <= oRecordSet.RecordCount; loopCount02++)
						{
							if (loopCount02 >= 1 & loopCount02 < 10)
							{
								CheckBoxID = "Check0" + loopCount02;
							}
							else
							{
								CheckBoxID = "Check" + loopCount02;
							}
							oForm.Items.Item(CheckBoxID).Specific.Caption = oRecordSet.Fields.Item("FullClsName").Value.ToString().Trim();
							oRecordSet.MoveNext();
						}
						break;
					case "CntcCode":
						oForm.DataSources.UserDataSources.Item("CntcName").Value = dataHelpClass.Get_ReData("U_FULLNAME", "U_MSTCOD", "[OHEM]", "'" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'", "");
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
			}
		}

		/// <summary>
		/// PS_PP982_CheckAll  체크박스 전체 선택(해제)
		/// </summary>
		private void PS_PP982_CheckAll()
		{
			string CheckType;
			short loopCount;

			try
			{
				oForm.Freeze(true);
				CheckType = "Y";
				oMat01.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 2; loopCount++)
				{
					if (oDS_PS_PP982O.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break; // TODO: might not be correct. Was : Exit For
					}
				}

				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 2; loopCount++)
				{
					oDS_PS_PP982O.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_PP982O.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_PP982O.SetValue("U_ColReg01", loopCount, "N");
					}
				}

				oMat01.LoadFromDataSource();
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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
                   // Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnSrch01")
					{
						PS_PP982_SelectGrid01(); //반별자료
						PS_PP982_SelectGrid02(); //집계자료(주별)
						PS_PP982_SelectGrid03(); //집계자료(일별)
					}
					else if (pVal.ItemUID == "BtnPrt01")  //버턴이 없음 ..
					{
						//System.Threading.Thread thread = new System.Threading.Thread(PS_PP982_Print_Report01);
						//thread.SetApartmentState(System.Threading.ApartmentState.STA);
						//thread.Start();
					}

					if (pVal.ItemUID == "BtnSrch04")
					{
						PS_PP982_SelectMatrix01();
					}
					else if (pVal.ItemUID == "BtnSave04")
					{
						PS_PP982_InsertMatrix01();
						PS_PP982_SelectMatrix01();
					}
					else if (pVal.ItemUID == "BtnDel04")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제후 복구는 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
						{
							PS_PP982_DeleteMatrix01();
							PS_PP982_SelectMatrix01();
						}
					}

					if (pVal.ItemUID == "BtnAll")
					{
						PS_PP982_CheckAll();  // 버턴 없음
					}
				}
				else if (pVal.BeforeAction == false)
				{
					//폴더를 사용할 때는 필수 소스
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
					}
					if (pVal.ItemUID == "Folder03")
					{
						oForm.PaneLevel = 3;
					}
					if (pVal.ItemUID == "Folder04")
					{
						oForm.PaneLevel = 4;
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CntcCode");
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
							oMat01.SelectRow(pVal.Row, true, false);
							oLastItemUID01 = pVal.ItemUID;
							oLastColUID01 = pVal.ColUID;
							oLastColRow01 = pVal.Row;
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
			try
			{
				if (pVal.Before_Action == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.Row > 0)
						{
							oMat01.SelectRow(pVal.Row, true, false);
						}
					}
				}
				else if (pVal.Before_Action == false)
				{

					if (pVal.ItemChanged == true)
					{
						PS_PP982_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
					}
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
							PS_PP982_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else
						{
							PS_PP982_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					PS_PP982_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid03);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP982L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP982M);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP982N);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP982O);
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
						case "7169":                            //엑셀 내보내기
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
						case "7169": //엑셀 내보내기
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
