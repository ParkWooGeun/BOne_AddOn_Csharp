using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 공수 잔량 조회
	/// </summary>
	internal class PS_PP991 : PSH_BaseClass
	{
		private string oFormUniqueID;

		private SAPbouiCOM.Grid oGrid01;
		private SAPbouiCOM.Grid oGrid02;
		private SAPbouiCOM.Matrix oMat01;

		private SAPbouiCOM.DataTable oDS_PS_PP991L;
		private SAPbouiCOM.DataTable oDS_PS_PP991M;
		private SAPbouiCOM.DBDataSource oDS_PS_PP991O;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP991.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP991_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP991");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP991_CreateItems();
				PS_PP991_ComboBox_Setting();
				PS_PP991_FormItemEnabled();
				PS_PP991_FormResize();

				oForm.EnableMenu(("1283"), false); // 삭제
				oForm.EnableMenu(("1286"), false); // 닫기
				oForm.EnableMenu(("1287"), false); // 복제
				oForm.EnableMenu(("1285"), false); // 복원
				oForm.EnableMenu(("1284"), true);  // 취소
				oForm.EnableMenu(("1293"), false); // 행삭제
				oForm.EnableMenu(("1281"), false);
				oForm.EnableMenu(("1282"), true);

				PS_PP991_FormReset();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Items.Item("Folder01").Specific.Select(); //폼이 로드 될 때 Folder01이 선택됨
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_PP991_CreateItems
		/// </summary>
		private void PS_PP991_CreateItems()
		{
			try
			{
				oGrid01 = oForm.Items.Item("Grid01").Specific;
				oGrid02 = oForm.Items.Item("Grid02").Specific;

				oForm.DataSources.DataTables.Add("PS_PP991L");
				oForm.DataSources.DataTables.Add("PS_PP991M");

				oGrid01.DataTable = oForm.DataSources.DataTables.Item("PS_PP991L");
				oGrid02.DataTable = oForm.DataSources.DataTables.Item("PS_PP991M");

				oDS_PS_PP991L = oForm.DataSources.DataTables.Item("PS_PP991L");
				oDS_PS_PP991M = oForm.DataSources.DataTables.Item("PS_PP991M");

				oDS_PS_PP991O = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				// 메트릭스 개체 할당
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
				oMat01.AutoResizeColumns();

				//표준공수VS실적공수
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID01").Specific.DataBind.SetBound(true, "", "BPLID01");

				//기준일자
				oForm.DataSources.UserDataSources.Add("StdDt01", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("StdDt01").Specific.DataBind.SetBound(true, "", "StdDt01");

				//작번별 주별 공수 현황
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID02").Specific.DataBind.SetBound(true, "", "BPLID02");

				//기준일자
				oForm.DataSources.UserDataSources.Add("StdDt02", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("StdDt02").Specific.DataBind.SetBound(true, "", "StdDt02");

				//예외 작번 관리
				//사업장
				oForm.DataSources.UserDataSources.Add("BPLID03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLID03").Specific.DataBind.SetBound(true, "", "BPLID03");

				//기준일자
				oForm.DataSources.UserDataSources.Add("StdDt03", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("StdDt03").Specific.DataBind.SetBound(true, "", "StdDt03");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP991_ComboBox_Setting
		/// </summary>
		private void PS_PP991_ComboBox_Setting()
		{
			string BPLID;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			try
			{
				BPLID = dataHelpClass.User_BPLID();

				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID01").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID02").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLID03").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", BPLID, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP991_FormItemEnabled
		/// </summary>
		private void PS_PP991_FormItemEnabled()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BPLID02").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
					PS_PP991_FlushToItemValue("BPLID02", 0, "");//팀, 담당, 반 콤보박스 강제 설정
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
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
		/// PS_PP991_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP991_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "Mat01":
						oMat01.FlushToDataSource();

						if (oCol == "PP030Entry")
						{
							sQry = " SELECT      T0.U_OrdNum AS [OrdNum],";    //작번
							sQry += "             T0.U_OrdSub1 AS [OrdSub1],";  //서브작번1
							sQry += "             T0.U_OrdSub2 AS [OrdSub2],";  //서브작번2
							sQry += "             T0.U_JakMyung AS [ItemName]"; //품명
							sQry += " FROM        [@PS_PP030H] AS T0";
							sQry += " WHERE       T0.DocEntry = '" + oDS_PS_PP991O.GetValue("U_ColReg02", oRow - 1).ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oDS_PS_PP991O.SetValue("U_ColReg01", oRow - 1, "Y");                                                        //선택
							oDS_PS_PP991O.SetValue("U_ColReg03", oRow - 1, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());   //작번
							oDS_PS_PP991O.SetValue("U_ColReg04", oRow - 1, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());  //서브작번1
							oDS_PS_PP991O.SetValue("U_ColReg05", oRow - 1, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());  //서브작번2
							oDS_PS_PP991O.SetValue("U_ColReg06", oRow - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품명
							oMat01.LoadFromDataSource();

							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								if ((oRow == oMat01.RowCount || oMat01.VisualRowCount == 0) && !string.IsNullOrEmpty(oMat01.Columns.Item("PP030Entry").Cells.Item(oRow).Specific.Value.ToString().Trim()))
								{
									PS_PP991_Add_MatrixRow01(oMat01.RowCount, false);
								}
							}
						}
						oMat01.AutoResizeColumns();
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
		/// PS_PP991_Add_MatrixRow01
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP991_Add_MatrixRow01(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP991O.InsertRecord(oRow);
				}

				oMat01.AddRow();
				oDS_PS_PP991O.Offset = oRow;
				oDS_PS_PP991O.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat01.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP991_FormResize
		/// </summary>
		private void PS_PP991_FormResize()
		{
			try
			{
				oForm.Freeze(true);

				//그룹박스 크기 동적 할당
				oForm.Items.Item("GrpBox01").Height = oForm.Items.Item("Grid01").Height + 68;
				oForm.Items.Item("GrpBox01").Width = oForm.Items.Item("Grid01").Width + 30;

				if (oGrid01.Columns.Count > 0)
				{
					oGrid01.AutoResizeColumns();
				}

				if (oGrid02.Columns.Count > 0)
				{
					oGrid02.AutoResizeColumns();
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
		/// PS_PP991_FormReset
		/// </summary>
		private void PS_PP991_FormReset()
		{
			try
			{
				oForm.Freeze(true);

				//헤더 초기화
				oForm.DataSources.UserDataSources.Item("StdDt01").Value = DateTime.Now.ToString("yyyyMMdd"); //기준일자
				oForm.DataSources.UserDataSources.Item("StdDt02").Value = DateTime.Now.ToString("yyyyMMdd"); //기준일자
				oForm.DataSources.UserDataSources.Item("StdDt03").Value = DateTime.Now.ToString("yyyyMMdd"); //기준일자

				//라인 초기화
				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();
				PS_PP991_Add_MatrixRow01(0, true);
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
		/// PS_PP991_CheckAll  체크박스 전체 선택(해제)
		/// </summary>
		private void PS_PP991_CheckAll()
		{
			string CheckType;
			int loopCount;

			try
			{
				oForm.Freeze(true);

				CheckType = "Y";

				oMat01.FlushToDataSource();

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 2; loopCount++)
                {
                    if (oDS_PS_PP991O.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
                    {
                        CheckType = "N";
                        break;
                    }
                }

                for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 2; loopCount++)
				{
					oDS_PS_PP991O.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_PP991O.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_PP991O.SetValue("U_ColReg01", loopCount, "N");
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
		/// PS_PP991_SelectGrid01  표준공수VS실적공수
		/// </summary>
		private void PS_PP991_SelectGrid01()
		{
			string sQry;
			string errMessage = String.Empty;

			string BPLID;
			string StdDt;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);


			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID01").Specific.Value.ToString().Trim();
				StdDt = oForm.Items.Item("StdDt01").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP991_01 '";
				sQry += BPLID + "','";
				sQry += StdDt + "'";

				oGrid01.DataTable.Clear();
				oDS_PS_PP991L.ExecuteQuery(sQry);

				oGrid01.Columns.Item(7).RightJustified = true;
				oGrid01.Columns.Item(9).RightJustified = true;
				oGrid01.Columns.Item(10).RightJustified = true;
				oGrid01.Columns.Item(12).RightJustified = true;
				oGrid01.Columns.Item(14).RightJustified = true;

				if (oGrid01.Rows.Count == 1)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oGrid01.AutoResizeColumns();
				oForm.Update();
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP991_SelectGrid02  작번별 주별 공수 현황 조회
		/// </summary>
		private void PS_PP991_SelectGrid02()
		{
			string sQry;
			string errMessage = String.Empty;

			string BPLID;
			string StdDt;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);


			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID02").Specific.Value.ToString().Trim();
				StdDt = oForm.Items.Item("StdDt02").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC PS_PP991_02 '";
				sQry += BPLID + "','";
				sQry += StdDt + "'";

				oGrid02.DataTable.Clear();
				oDS_PS_PP991M.ExecuteQuery(sQry);

				oGrid02.Columns.Item(6).RightJustified = true;
				oGrid02.Columns.Item(7).RightJustified = true;
				oGrid02.Columns.Item(8).RightJustified = true;
				oGrid02.Columns.Item(9).RightJustified = true;
				oGrid02.Columns.Item(10).RightJustified = true;
				oGrid02.Columns.Item(11).RightJustified = true;
				oGrid02.Columns.Item(12).RightJustified = true;
				oGrid02.Columns.Item(13).RightJustified = true;
				oGrid02.Columns.Item(14).RightJustified = true;
				oGrid02.Columns.Item(15).RightJustified = true;
				oGrid02.Columns.Item(16).RightJustified = true;
				oGrid02.Columns.Item(17).RightJustified = true;

				if (oGrid02.Rows.Count == 1)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				oGrid02.AutoResizeColumns();
				oForm.Update();
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP991_SelectMatrix01  예외작번 관리 조회
		/// </summary>
		private void PS_PP991_SelectMatrix01()
		{
			int i;
			string sQry;
			string errMessage = String.Empty;

			string BPLID;
			string StdDt;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLID03").Specific.Value.ToString().Trim();
				StdDt = oForm.Items.Item("StdDt03").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				sQry = "EXEC [PS_PP991_03]";

				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_PP991O.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP991_Add_MatrixRow01(0, true);
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP991O.Size)
					{
						oDS_PS_PP991O.InsertRecord(i);
					}

					oMat01.AddRow();
					oDS_PS_PP991O.Offset = i;

					oDS_PS_PP991O.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP991O.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Select").Value.ToString().Trim());		//선택
					oDS_PS_PP991O.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("PP030Entry").Value.ToString().Trim());	//작업지시번호
					oDS_PS_PP991O.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());		//작번
					oDS_PS_PP991O.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());		//서브작번1
					oDS_PS_PP991O.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());		//서브작번2
					oDS_PS_PP991O.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());	//품명

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				PS_PP991_Add_MatrixRow01(oMat01.VisualRowCount, false);
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
				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
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
		/// PS_PP991_InsertMatrix01  저장
		/// </summary>
		private void PS_PP991_InsertMatrix01()
		{
			short loopCount;
			string sQry;

			int PP030Entry;	    //작업지시번호
			string OrdNum;		//작번
			string OrdSub1;		//서브작번1
			string OrdSub2;		//서브작번2
			string ItemName;    //품명

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat01.FlushToDataSource();
				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP991O.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						PP030Entry = Convert.ToInt32(oDS_PS_PP991O.GetValue("U_ColReg02", loopCount).ToString().Trim());	//작업지시번호
						OrdNum     = oDS_PS_PP991O.GetValue("U_ColReg03", loopCount).ToString().Trim();	//작번
						OrdSub1    = oDS_PS_PP991O.GetValue("U_ColReg04", loopCount).ToString().Trim();	//서브작번1
						OrdSub2    = oDS_PS_PP991O.GetValue("U_ColReg05", loopCount).ToString().Trim();	//서브작번2
						ItemName   = oDS_PS_PP991O.GetValue("U_ColReg06", loopCount).ToString().Trim();	//품명

						sQry = "      EXEC [PS_PP991_04] '";
						sQry += PP030Entry + "','";
						sQry += OrdNum + "','";
						sQry += OrdSub1 + "','";
						sQry += OrdSub2 + "','";
						sQry += ItemName + "'";
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
		/// PS_PP991_DeleteMatrix01  삭제
		/// </summary>
		private void PS_PP991_DeleteMatrix01()
		{
			short loopCount;
			string sQry;
			string PP030Entry;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oMat01.FlushToDataSource();

				ProgressBar01.Text = "삭제중...";

				for (loopCount = 0; loopCount <= oMat01.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_PP991O.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						PP030Entry = oDS_PS_PP991O.GetValue("U_ColReg02", loopCount).ToString().Trim();

						sQry = " EXEC [PS_PP991_05] '";
						sQry += PP030Entry + "'";
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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
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
					if (pVal.ItemUID == "BtnSrch01")
					{
						PS_PP991_SelectGrid01();
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						PS_PP991_SelectGrid02();
					}
					else if (pVal.ItemUID == "BtnSrch03")
					{
						PS_PP991_SelectMatrix01();
					}
					else if (pVal.ItemUID == "BtnSave03")
					{
						PS_PP991_InsertMatrix01();
						PS_PP991_SelectMatrix01();
					}
					else if (pVal.ItemUID == "BtnDel03")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제후 복구는 불가능합니다. 삭제하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
						{
							PS_PP991_DeleteMatrix01();
							PS_PP991_SelectMatrix01();
						}
					}
					else if (pVal.ItemUID == "BtnAll")
					{
						PS_PP991_CheckAll();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
					}
					if (pVal.ItemUID == "Folder03")
					{
						oForm.PaneLevel = 3;
						oForm.DefButton = "BtnSrch03";
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PP030Entry");
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
						PS_PP991_FlushToItemValue(pVal.ItemUID, 0, "");
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
							PS_PP991_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else
						{
							PS_PP991_FlushToItemValue(pVal.ItemUID, 0, "");
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
					PS_PP991_FormResize();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP991L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP991M);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP991O);
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
