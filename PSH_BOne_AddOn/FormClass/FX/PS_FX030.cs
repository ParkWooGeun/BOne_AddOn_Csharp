using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 고정자산현황및상태등록
	/// </summary>
	internal class PS_FX030 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_FX030L;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FX030.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_FX030_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_FX030");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);
				PS_FX030_CreateItems();
				PS_FX030_ComboBox_Setting();
				PS_FX030_Initialization();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", true);  // 취소
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
		/// PS_FX030_CreateItems
		/// </summary>
		private void PS_FX030_CreateItems()
		{
			try
			{
				oDS_PS_FX030L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("BPLID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID").Specific.DataBind.SetBound(true, "", "BPLID");

				//기준년도
				oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
				oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");

				//기준년월
				oForm.DataSources.UserDataSources.Add("StdYM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 6);
				oForm.Items.Item("StdYM").Specific.DataBind.SetBound(true, "", "StdYM");

				//자산분류
				oForm.DataSources.UserDataSources.Add("FixType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("FixType").Specific.DataBind.SetBound(true, "", "FixType");

				//폐기대상 조회
				oForm.DataSources.UserDataSources.Add("ChkDisu", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("ChkDisu").Specific.DataBind.SetBound(true, "", "ChkDisu");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX030_ComboBox_Setting
		/// </summary>
		private void PS_FX030_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLID").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//자산분류
				sQry = "     SELECT     U_Minor,";
				sQry += "                U_CdName";
				sQry += "  FROM      [@PS_SY001L]";
				sQry += "  WHERE    Code = 'FX001'";
				sQry += "                AND U_UseYN = 'Y'";
				oForm.Items.Item("FixType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("FixType").Specific, sQry, "", false, false);
				oForm.Items.Item("FixType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//상태(Line)
				sQry = "    SELECT      U_Minor, ";
				sQry += "                U_CdName";
				sQry += " FROM       [@PS_SY001L]";
				sQry += " WHERE      Code = 'FX004'";
				sQry += "                AND U_UseYN = 'Y'";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("Status"), sQry, "", "");
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
		/// PS_FX030_Initialization
		/// </summary>
		private void PS_FX030_Initialization()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLID").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("StdYear").Specific.Value = DateTime.Now.ToString("yyyy");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX030_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_FX030_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				switch (oUID)
				{
					case "TeamCD":
						sQry = "SELECT  T1.U_CodeNm FROM [@PS_HR200H] AS T0 Inner Join [@PS_HR200L] AS T1 ON T0.Code = T1.Code ";
						sQry += " WHERE   T0.Name = '부서' AND T1.U_UseYN = 'Y' AND T1.U_Char2 = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim() + "' ";
						sQry += " And T1.U_Code = '" + oForm.Items.Item("TeamCD").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("TeamNM").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
					case "RspCD":
						sQry = "SELECT  T1.U_CodeNm FROM [@PS_HR200H] AS T0 Inner Join [@PS_HR200L] AS T1 ON T0.Code = T1.Code ";
						sQry += " WHERE   T0.Name = '담당' AND T1.U_UseYN = 'Y' AND T1.U_Char2 = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim() + "' ";
						sQry += " And T1.U_Char1 = '" + oForm.Items.Item("TeamCD").Specific.Value.ToString().Trim() + "'";
						sQry += " And T1.U_Code = '" + oForm.Items.Item("RspCD").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("RspNM").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
					case "StdYear":
						sQry = "SELECT MAX(U_YM) AS [MaxYM] FROM [@PS_FX020H] WHERE U_BPLId = '" + oForm.Items.Item("BPLID").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);

						if (string.IsNullOrEmpty(oForm.Items.Item("StdYM").Specific.Value.ToString().Trim()))
						{
							oForm.DataSources.UserDataSources.Item("StdYM").Value = oRecordSet.Fields.Item("MaxYM").Value.ToString().Trim();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_FX030_MTX01
		/// </summary>
		private void PS_FX030_MTX01()
		{
			int i;
			string errMessage = string.Empty;
			string BPLID;   //사업장
			string StdYear; //기준년도
			string StdYM;   //기준년월
			string FixType; //자산분류
			string ChkDisu; //폐기대상 조회
			string TeamCd;  //팀코드
			string RspCd;   //담당코드
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();
				FixType = oForm.Items.Item("FixType").Specific.Value.ToString().Trim();
				TeamCd = oForm.Items.Item("TeamCD").Specific.Value.ToString().Trim();
				RspCd = oForm.Items.Item("RspCD").Specific.Value.ToString().Trim();

				if (oForm.Items.Item("ChkDisu").Specific.Checked == true)
				{
					ChkDisu = "Y";
				}
				else
				{
					ChkDisu = "N";
				}

				ProgressBar01.Text = "조회시작!";

				if (ChkDisu == "N")  //폐기대상 조회 체크 안됨
				{
					sQry = "  EXEC [PS_FX030_01] '";
					sQry += BPLID + "','";
					sQry += StdYear + "','";
					sQry += StdYM + "','";
					sQry += FixType + "','";
					sQry += TeamCd + "','";
					sQry += RspCd + "'";
				}
				else //폐기대상 조회 체크
				{
					sQry = "  EXEC [PS_FX030_02] '";
					sQry += BPLID + "','";
					sQry += StdYear + "','";
					sQry += StdYM + "','";
					sQry += FixType + "','";
					sQry += TeamCd + "','";
					sQry += RspCd + "'";
				}
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_FX030L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_FX030L.Size)
					{
						oDS_PS_FX030L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_FX030L.Offset = i;

					oDS_PS_FX030L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_FX030L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());   //선택
					oDS_PS_FX030L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("FixCode").Value.ToString().Trim()); //자산코드
					oDS_PS_FX030L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("FixName").Value.ToString().Trim()); //품명
					oDS_PS_FX030L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("FixSpec").Value.ToString().Trim()); //규격
					oDS_PS_FX030L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("PostQty").Value.ToString().Trim()); //수량
					oDS_PS_FX030L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("PostDate").Value.ToString().Trim()).ToString("yyyyMMdd"));  //구입일자
					oDS_PS_FX030L.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet.Fields.Item("SPostDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //기초취득일
					oDS_PS_FX030L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("LongYear").Value.ToString().Trim()); //내용년수
					oDS_PS_FX030L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("PostAmt").Value.ToString().Trim());  //취득금액
					oDS_PS_FX030L.SetValue("U_ColSum03", i, oRecordSet.Fields.Item("CurAmt").Value.ToString().Trim());   //현재잔액
					oDS_PS_FX030L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("TeamName").Value.ToString().Trim()); //부서
					oDS_PS_FX030L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("RspName").Value.ToString().Trim());  //담당
					oDS_PS_FX030L.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("ClsName").Value.ToString().Trim());  //반
					oDS_PS_FX030L.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("Qty").Value.ToString().Trim());      //현황
					oDS_PS_FX030L.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("Status").Value.ToString().Trim());   //상태
					oDS_PS_FX030L.SetValue("U_ColReg15", i, oRecordSet.Fields.Item("Comment").Value.ToString().Trim());  //비고
					oDS_PS_FX030L.SetValue("U_ColReg16", i, oRecordSet.Fields.Item("FixType").Value.ToString().Trim());  //자산분류
					oDS_PS_FX030L.SetValue("U_ColReg17", i, oRecordSet.Fields.Item("CDate").Value.ToString().Trim());    //등록일
					oDS_PS_FX030L.SetValue("U_ColReg18", i, oRecordSet.Fields.Item("CUName").Value.ToString().Trim());   //등록자이름
					oDS_PS_FX030L.SetValue("U_ColReg19", i, oRecordSet.Fields.Item("UDate").Value.ToString().Trim());    //수정일
					oDS_PS_FX030L.SetValue("U_ColReg20", i, oRecordSet.Fields.Item("UUName").Value.ToString().Trim());   //수정자이름
					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				ProgressBar01.Stop();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01); //메모리 해제
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_FX030_AddData
		/// </summary>
		private void PS_FX030_AddData()
		{
			int loopCount;
			string BPLID;   //사업장
			string StdYear; //기준년도
			string FixCode; //자산코드
			string Qty;     //현황(수량)
			string Status;  //상태
			string Comment; //비고
			string UserSign; //UserSign
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();
				UserSign = PSH_Globals.oCompany.UserSignature.ToString();

				oMat.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oMat.Columns.Item("Check").Cells.Item(loopCount + 1).Specific.Checked == true)
					{
						FixCode = oDS_PS_FX030L.GetValue("U_ColReg02", loopCount).ToString().Trim(); //자산코드
						Qty = oDS_PS_FX030L.GetValue("U_ColQty02", loopCount).ToString().Trim();     //현황
						Status = oDS_PS_FX030L.GetValue("U_ColReg14", loopCount).ToString().Trim();  //상태
						Comment = oDS_PS_FX030L.GetValue("U_ColReg15", loopCount).ToString().Trim(); //비고

						sQry = "                EXEC [PS_FX030_03] ";
						sQry += "'" + BPLID + "',";   //사업장
						sQry += "'" + StdYear + "',"; //기준년도
						sQry += "'" + FixCode + "',"; //자산코드
						sQry += "'" + Qty + "',";     //현황
						sQry += "'" + Status + "',";  //상태
						sQry += "'" + Comment + "',"; //비고
						sQry += "'" + UserSign + "'"; //UserSign
						oRecordSet.DoQuery(sQry);
					}
				}

				PSH_Globals.SBO_Application.StatusBar.SetText("등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
			}

			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_FX030_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_FX030_Print_Report01()
		{
			string WinTitle;
			string ReportName = string.Empty;
			string BPLID;   //사업장
			string StdYear; //기준년도
			string StdYM;   //기준년월
			string FixType; //자산분류
			string ChkDisu; //폐기대상 조회
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();
				FixType = oForm.Items.Item("FixType").Specific.Value.ToString().Trim();
				ChkDisu = (oForm.Items.Item("ChkDisu").Specific.Checked == true ? "Y" : "N");

				WinTitle = "[PS_FX030_02] 고정자산 재물조사";

				if (BPLID == "1")
				{
				}
				else if (BPLID == "2")
				{
					ReportName = "PS_FX030_02.rpt";
				}
				else if (BPLID == "3")
				{
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYear", StdYear));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));
				dataPackParameter.Add(new PSH_DataPackClass("@FixType", FixType));
				dataPackParameter.Add(new PSH_DataPackClass("@ChkDisu", ChkDisu));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX030_Print_Report02
		/// </summary>
		[STAThread]
		private void PS_FX030_Print_Report02()
		{
			string WinTitle;
			string ReportName = string.Empty;
			string BPLID;   //사업장
			string StdYear; //기준년도
			string StdYM;   //기준년월
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();

				WinTitle = "[PS_FX030_05] 고정자산 보유현황";

				if (BPLID == "1")
				{
				}
				else if (BPLID == "2")
				{
					ReportName = "PS_FX030_05.rpt";
				}
				else if (BPLID == "3")
				{
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYear", StdYear));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX030_Print_Report03
		/// </summary>
		[STAThread]
		private void PS_FX030_Print_Report03()
		{
			string WinTitle;
			string ReportName = string.Empty;
			string BPLID;   //사업장
			string StdYear; //기준년도
			string StdYM;   //기준년월
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();

				WinTitle = "[PS_FX030_08] 고정자산 보유현황";

				if (BPLID == "1")
				{
				}
				else if (BPLID == "2")
				{
					ReportName = "PS_FX030_08.rpt";
				}
				else if (BPLID == "3")
				{
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYear", StdYear));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX030_Print_Report04
		/// </summary>
		[STAThread]
		private void PS_FX030_Print_Report04()
		{
			string WinTitle;
			string ReportName = string.Empty;
			string BPLID;   //사업장
			string StdYear; //기준년도
			string StdYM;   //기준년월
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLID").Specific.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();
				StdYM = oForm.Items.Item("StdYM").Specific.Value.ToString().Trim();

				WinTitle = "[PS_FX030_11] 결과현황 일괄조회";

				if (BPLID == "1")
				{
				}
				else if (BPLID == "2")
				{
					ReportName = "PS_FX030_11.rpt";
				}
				else if (BPLID == "3")
				{
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>(); //Parameter
				List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>(); //SubReport

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYear", StdYear));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYM", StdYM));

				//SubReport Parameter
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@BPLID", BPLID, "PS_FX030_SUB1"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@StdYear", StdYear, "PS_FX030_SUB1"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@StdYM", StdYM, "PS_FX030_SUB1"));

				formHelpClass.OpenCrystalReport(dataPackParameter, dataPackSubReportParameter, WinTitle, ReportName);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_FX030_CheckAll
		/// </summary>
		private void PS_FX030_CheckAll()
		{
			int loopCount;
			string CheckType;
			try
			{
				oForm.Freeze(true);
				CheckType = "Y";

				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_FX030L.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break; // TODO: might not be correct. Was : Exit For
					}
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_FX030L.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_FX030L.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_FX030L.SetValue("U_ColReg01", loopCount, "N");
					}
				}

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
		/// PS_FX030_Add_MatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_FX030_Add_MatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_FX030L.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_FX030L.Offset = oRow;
				oDS_PS_FX030L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat.LoadFromDataSource();
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
				//case SAPbouiCOM.BoEventTypes.et_KEY_DOWN: //2
				//	Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
				//    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
				//    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
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
				//case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT:
				//    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
				//    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
				//    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_Drag: //39
				//    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
					if (pVal.ItemUID == "BtnSrch")
					{
						PS_FX030_MTX01();
					}
					else if (pVal.ItemUID == "BtnSave")
					{
						PS_FX030_AddData();
					}
					else if (pVal.ItemUID == "BtnPrint")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_FX030_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
					else if (pVal.ItemUID == "BtnPrt02")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_FX030_Print_Report02);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
					else if (pVal.ItemUID == "BtnPrt03")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_FX030_Print_Report03);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
					else if (pVal.ItemUID == "BtnPrt04")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_FX030_Print_Report04);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
					}
					else if (pVal.ItemUID == "BtnAll")
					{
						PS_FX030_CheckAll();
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
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					oForm.Update();
				}
				else if (pVal.BeforeAction == false)
				{
					PS_FX030_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FX030L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_MenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public void Raise_MenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1283": //삭제
							break;
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							PS_FX030_Add_MatrixRow(oMat.VisualRowCount, false);
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1283": //삭제
							break;
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1287": // 복제
							break;
						case "7169": //엑셀 내보내기
							oDS_PS_FX030L.RemoveRecord(oDS_PS_FX030L.Size - 1);
							oMat.LoadFromDataSource();
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
		/// Raise_FormDataEvent
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
