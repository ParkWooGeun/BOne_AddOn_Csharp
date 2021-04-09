using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 가동율조회
	/// </summary>
	internal class PS_PP981 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.DBDataSource oDS_PS_PP981L; //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP981M;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP981.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP981_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP981");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP981_CreateItems();
				PS_PP981_ComboBox_Setting();

				oForm.EnableMenu(("1283"), false); // 삭제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1287"), false); // 복제
				oForm.EnableMenu(("1285"), false); // 복원
				oForm.EnableMenu(("1284"), true);  // 취소
				oForm.EnableMenu(("1293"), false); // 행삭제
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
		/// PS_PP981_CreateItems
		/// </summary>
		private void PS_PP981_CreateItems()
		{
			try
			{
				oDS_PS_PP981L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oDS_PS_PP981M = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

				//// 메트릭스 개체 할당
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				//ms_NotSupported
				oMat02.AutoResizeColumns();

				//부서
				oForm.DataSources.UserDataSources.Add("TeamCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("TeamCode").Specific.DataBind.SetBound(true, "", "TeamCode");

				//담당
				oForm.DataSources.UserDataSources.Add("RspCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("RspCode").Specific.DataBind.SetBound(true, "", "RspCode");

				//반
				oForm.DataSources.UserDataSources.Add("ClsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ClsCode").Specific.DataBind.SetBound(true, "", "ClsCode");

				//사번
				oForm.DataSources.UserDataSources.Add("MSTCOD", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("MSTCOD").Specific.DataBind.SetBound(true, "", "MSTCOD");

				//성명
				oForm.DataSources.UserDataSources.Add("MSTNAM", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("MSTNAM").Specific.DataBind.SetBound(true, "", "MSTNAM");

				//기준년도
				oForm.DataSources.UserDataSources.Add("StdYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
				oForm.Items.Item("StdYear").Specific.DataBind.SetBound(true, "", "StdYear");
				oForm.Items.Item("StdYear").Specific.VALUE = DateTime.Now.ToString("yyyy");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP981_ComboBox_Setting
		/// </summary>
		private void PS_PP981_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				oForm.Items.Item("CLTCOD").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CLTCOD").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				oForm.Items.Item("CLTCOD").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP981_Add_MatrixRow01
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP981_Add_MatrixRow01(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP981L.InsertRecord(oRow);
				}

				oMat01.AddRow();
				oDS_PS_PP981L.Offset = oRow;
				oDS_PS_PP981L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat01.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP981_Add_MatrixRow02
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP981_Add_MatrixRow02(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP981M.InsertRecord(oRow);
				}

				oMat02.AddRow();
				oDS_PS_PP981M.Offset = oRow;
				oDS_PS_PP981M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat02.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP981_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_PP981_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = String.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("StdYear").Specific.Value.ToString().Trim()))
				{
					errMessage = "기준년도는 필수사항입니다. 확인하세요.";
					oForm.Items.Item("StdYear").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
		/// PS_PP981_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP981_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			int i;
			string sQry;

			string CLTCOD;
			string TeamCode;
			string RspCode;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "CLTCOD":
						CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("TeamCode").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("TeamCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("TeamCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						// 부서콤보세팅
						oForm.Items.Item("TeamCode").Specific.ValidValues.Add("%", "전체");
						sQry =  "  SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '1'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char2 = '" + CLTCOD + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("TeamCode").Specific, sQry, "", false, false);
						oForm.Items.Item("TeamCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "TeamCode":
						TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
						
						if (oForm.Items.Item("RspCode").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("RspCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("RspCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						oForm.Items.Item("RspCode").Specific.ValidValues.Add("%", "전체");
						sQry =  "  SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '2'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char1 = '" + TeamCode + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("RspCode").Specific, sQry, "", false, false);
						oForm.Items.Item("RspCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;

					case "RspCode":
						TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
						RspCode  = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();

						if (oForm.Items.Item("ClsCode").Specific.ValidValues.Count > 0)
						{
							for (i = oForm.Items.Item("ClsCode").Specific.ValidValues.Count - 1; i >= 0; i += -1)
							{
								oForm.Items.Item("ClsCode").Specific.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index);
							}
						}

						oForm.Items.Item("ClsCode").Specific.ValidValues.Add("%", "전체");
						sQry = "   SELECT      U_Code AS [Code],";
						sQry += "                 U_CodeNm As [Name]";
						sQry += "  FROM       [@PS_HR200L]";
						sQry += "  WHERE      Code = '9'";
						sQry += "                 AND U_UseYN = 'Y'";
						sQry += "                 AND U_Char1 = '" + RspCode + "'";
						sQry += "                 AND U_Char2 = '" + TeamCode + "'";
						sQry += "  ORDER BY  U_Seq";
						dataHelpClass.Set_ComboList(oForm.Items.Item("ClsCode").Specific, sQry, "", false, false);
						oForm.Items.Item("ClsCode").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
						break;


					case "MSTCOD":
						sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("MSTNAM").Specific.VALUE = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
		/// PS_PP981_MTX01
		/// </summary>
		private void PS_PP981_MTX01()
		{
			short i;
			string sQry;
			string errMessage = String.Empty;

			string CLTCOD;   //사업장
			string TeamCode; //부서
			string RspCode;  //담당
			string ClsCode;  //반
			string MSTCOD;   //사번
			string StdYear;  //기준년도

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
				RspCode = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim();
				ClsCode = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();
				MSTCOD = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
				StdYear = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC [PS_PP981_01] '" + CLTCOD + "','" + TeamCode + "','" + RspCode + "','" + ClsCode + "','" + MSTCOD + "','" + StdYear + "','" + "1', 'N'";
				oRecordSet.DoQuery(sQry);

				// 가동율 내역
				oMat01.Clear();
				oDS_PS_PP981L.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP981_Add_MatrixRow01(0, true);
					errMessage = "가동율 내역 조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP981L.Size)
					{
						oDS_PS_PP981L.InsertRecord(i);
					}

					oMat01.AddRow();
					oDS_PS_PP981L.Offset = i;

					oDS_PS_PP981L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP981L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Class").Value.ToString().Trim()); //구분
					oDS_PS_PP981L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("Mon01").Value.ToString().Trim()); //1월
					oDS_PS_PP981L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("Mon02").Value.ToString().Trim()); //2월
					oDS_PS_PP981L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("Mon03").Value.ToString().Trim()); //3월
					oDS_PS_PP981L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("Mon04").Value.ToString().Trim()); //4월
					oDS_PS_PP981L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("Mon05").Value.ToString().Trim()); //5월
					oDS_PS_PP981L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("Mon06").Value.ToString().Trim()); //6월
					oDS_PS_PP981L.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("Mon07").Value.ToString().Trim()); //7월
					oDS_PS_PP981L.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("Mon08").Value.ToString().Trim()); //8월
					oDS_PS_PP981L.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("Mon09").Value.ToString().Trim()); //9월
					oDS_PS_PP981L.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("Mon10").Value.ToString().Trim()); //10월
					oDS_PS_PP981L.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("Mon11").Value.ToString().Trim()); //11월
					oDS_PS_PP981L.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("Mon12").Value.ToString().Trim()); //12월
					oDS_PS_PP981L.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("Total").Value.ToString().Trim()); //계

					oRecordSet.MoveNext();
				}
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();

				// 비가동 내역
				sQry = "EXEC [PS_PP981_01] '" + CLTCOD + "','" + TeamCode + "','" + RspCode + "','" + ClsCode + "','" + MSTCOD + "','" + StdYear + "','" + "2', 'N'";
				oRecordSet.DoQuery(sQry);
				oMat02.Clear();
				oDS_PS_PP981M.Clear();
				oMat02.FlushToDataSource();
				oMat02.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP981_Add_MatrixRow02(0, true);
					errMessage = "비가동 내역 조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP981M.Size)
					{
						oDS_PS_PP981M.InsertRecord(i);
					}

					oMat02.AddRow();
					oDS_PS_PP981M.Offset = i;

					oDS_PS_PP981M.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP981M.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Class").Value.ToString().Trim()); //구분
					oDS_PS_PP981M.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("Mon01").Value.ToString().Trim()); //1월
					oDS_PS_PP981M.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("Mon02").Value.ToString().Trim()); //2월
					oDS_PS_PP981M.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("Mon03").Value.ToString().Trim()); //3월
					oDS_PS_PP981M.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("Mon04").Value.ToString().Trim()); //4월
					oDS_PS_PP981M.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("Mon05").Value.ToString().Trim()); //5월
					oDS_PS_PP981M.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("Mon06").Value.ToString().Trim()); //6월
					oDS_PS_PP981M.SetValue("U_ColReg08", i, oRecordSet.Fields.Item("Mon07").Value.ToString().Trim()); //7월
					oDS_PS_PP981M.SetValue("U_ColReg09", i, oRecordSet.Fields.Item("Mon08").Value.ToString().Trim()); //8월
					oDS_PS_PP981M.SetValue("U_ColReg10", i, oRecordSet.Fields.Item("Mon09").Value.ToString().Trim()); //9월
					oDS_PS_PP981M.SetValue("U_ColReg11", i, oRecordSet.Fields.Item("Mon10").Value.ToString().Trim()); //10월
					oDS_PS_PP981M.SetValue("U_ColReg12", i, oRecordSet.Fields.Item("Mon11").Value.ToString().Trim()); //11월
					oDS_PS_PP981M.SetValue("U_ColReg13", i, oRecordSet.Fields.Item("Mon12").Value.ToString().Trim()); //12월
					oDS_PS_PP981M.SetValue("U_ColReg14", i, oRecordSet.Fields.Item("Total").Value.ToString().Trim()); //계

					oRecordSet.MoveNext();
				}

				oMat02.LoadFromDataSource();
				oMat02.AutoResizeColumns();

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
				oForm.Update();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP981_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_PP981_Print_Report01()
		{
			string WinTitle;
			string ReportName;

			string CLTCOD;	 //사업장
			string TeamCode; //부서
			string RspCode;	 //담당
			string ClsCode;	 //반
			string MSTCOD;	 //사번
			string StdYear;  //기준년도

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				CLTCOD = oForm.Items.Item("CLTCOD").Specific.Value.ToString().Trim();
				TeamCode = oForm.Items.Item("TeamCode").Specific.Value.ToString().Trim();
				RspCode  = oForm.Items.Item("RspCode").Specific.Value.ToString().Trim(); ;
				ClsCode  = oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim();
				MSTCOD   = oForm.Items.Item("MSTCOD").Specific.Value.ToString().Trim();
				StdYear  = oForm.Items.Item("StdYear").Specific.Value.ToString().Trim();

				WinTitle = "[PS_PP981] 레포트";
				ReportName = "PS_PP981_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>(); //Formula List
				List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>();

				//Formula

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD));
				dataPackParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode));
				dataPackParameter.Add(new PSH_DataPackClass("@RspCode", RspCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode));
				dataPackParameter.Add(new PSH_DataPackClass("@MSTCOD", MSTCOD));
				dataPackParameter.Add(new PSH_DataPackClass("@StdYear", StdYear));
				dataPackParameter.Add(new PSH_DataPackClass("@QryCls", "1"));
				dataPackParameter.Add(new PSH_DataPackClass("@ReportYN", "N"));

				//SubReport Parameter
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@CLTCOD", CLTCOD, "PS_PP981_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@TeamCode", TeamCode, "PS_PP981_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@RspCode", RspCode, "PS_PP981_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@ClsCode", ClsCode, "PS_PP981_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@MSTCOD", MSTCOD, "PS_PP981_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@StdYear", StdYear, "PS_PP981_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@QryCls", "2", "PS_PP981_SUB_01"));
				dataPackSubReportParameter.Add(new PSH_DataPackClass("@ReportYN", "Y", "PS_PP981_SUB_01"));

				formHelpClass.CrystalReportOpen(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnSearch")
					{
						if (PS_PP981_HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							PS_PP981_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnPrint")
					{
						if (PS_PP981_HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_PP981_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
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
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemChanged == true)
					{
						PS_PP981_FlushToItemValue(pVal.ItemUID, 0, "");
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
						if (pVal.Row == 0)
						{
							oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
							oMat01.FlushToDataSource();
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (pval.BeforeAction == true)
				{
					if (pval.ItemChanged == true)
					{

						if ((pval.ItemUID == "Mat01"))
						{
						}
						else
						{
							PS_PP981_FlushToItemValue(pval.ItemUID, 0, "");
						}
					}
				}
				else if (pval.BeforeAction == false)
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP981L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP981M);
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
																//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							oForm.Freeze(true);
							PS_PP981_Add_MatrixRow01(oMat01.VisualRowCount, false);
							PS_PP981_Add_MatrixRow02(oMat02.VisualRowCount, false);
							oForm.Freeze(false);
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
									 //엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							oDS_PS_PP981L.RemoveRecord(oDS_PS_PP981L.Size - 1);
							oMat01.LoadFromDataSource();
							oDS_PS_PP981M.RemoveRecord(oDS_PS_PP981M.Size - 1);
							oMat02.LoadFromDataSource();
							oForm.Freeze(false);
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
