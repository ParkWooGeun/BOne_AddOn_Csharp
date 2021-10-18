using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 기성매입등록및조회
	/// </summary>
	internal class PS_SD061 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DBDataSource oDS_PS_SD061A; //등록용 Matrix
		private SAPbouiCOM.DataTable oDS_PS_SD061B; //조회용 Grid
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD061.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD061_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD061");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_SD061_CreateItems();
				PS_SD061_SetComboBox();
				PS_SD061_ResizeForm();
				PS_SD061_AddMatrixRow(0, true);
				PS_SD061_LoadCaption();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);

				PS_SD061_ResetForm(); //폼초기화
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Freeze(false);
				oForm.Items.Item("Folder01").Specific.Select();	//폼이 로드 될 때 Folder01이 선택됨
				oForm.Items.Item("StdYear01").Click();
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_SD061_CreateItems
		/// </summary>
		private void PS_SD061_CreateItems()
		{
			try
			{
				oDS_PS_SD061A = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("PS_SD061B");
				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_SD061B");
				oDS_PS_SD061B = oForm.DataSources.DataTables.Item("PS_SD061B");

				//기준년도
				oForm.DataSources.UserDataSources.Add("StdYear01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
				oForm.Items.Item("StdYear01").Specific.DataBind.SetBound(true, "", "StdYear01");

				//기준월
				oForm.DataSources.UserDataSources.Add("StdMth01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
				oForm.Items.Item("StdMth01").Specific.DataBind.SetBound(true, "", "StdMth01");

				//작번
				oForm.DataSources.UserDataSources.Add("OrdNum01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("OrdNum01").Specific.DataBind.SetBound(true, "", "OrdNum01");

				//품명
				oForm.DataSources.UserDataSources.Add("ItemName01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemName01").Specific.DataBind.SetBound(true, "", "ItemName01");

				//규격
				oForm.DataSources.UserDataSources.Add("ItemSpec01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemSpec01").Specific.DataBind.SetBound(true, "", "ItemSpec01");

				//조회
				//기준년도
				oForm.DataSources.UserDataSources.Add("StdYear02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 4);
				oForm.Items.Item("StdYear02").Specific.DataBind.SetBound(true, "", "StdYear02");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD061_SetComboBox
		/// </summary>
		private void PS_SD061_SetComboBox()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//기준월
				sQry = " SELECT      U_Code AS [Code],";
				sQry += "             U_CodeNm AS [Name]";
				sQry += " FROM        [@PS_HR200L]";
				sQry += " WHERE       Code = '4'";
				sQry += "             AND U_UseYN = 'Y'";
				sQry += " ORDER BY    U_Seq";
				oForm.Items.Item("StdMth01").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("StdMth01").Specific, sQry, "", false, false);
				oForm.Items.Item("StdMth01").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD061_ResizeForm
		/// </summary>
		private void PS_SD061_ResizeForm()
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
		/// PS_SD061_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_SD061_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_SD061A.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_SD061A.Offset = oRow;
				oDS_PS_SD061A.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oDS_PS_SD061A.SetValue("U_ColReg01", oRow, "Y");
				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_SD061_LoadCaption
		/// </summary>
		private void PS_SD061_LoadCaption()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BtnAdd01").Specific.Caption = "저장";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("BtnAdd01").Specific.Caption = "수정";
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
		/// PS_SD061_ResetForm
		/// 화면 초기화
		/// </summary>
		private void PS_SD061_ResetForm()
		{
			try
			{
				oForm.Freeze(true);

				//헤더 초기화
				oForm.DataSources.UserDataSources.Item("StdYear01").Value = DateTime.Now.ToString("yyyy"); //기준년도(등록)
				oForm.DataSources.UserDataSources.Item("StdMth01").Value = "%"; //기준월
				oForm.DataSources.UserDataSources.Item("StdYear02").Value = DateTime.Now.ToString("yyyy"); //기준년도(조회)

				//라인 초기화
				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();
				PS_SD061_AddMatrixRow(0, true);
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
		/// PS_SD061_CheckAll
		/// </summary>
		private void PS_SD061_CheckAll()
		{
			string CheckType;
			short loopCount;

			try
			{
				oForm.Freeze(true);
				CheckType = "Y";

				oMat.FlushToDataSource();

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (oDS_PS_SD061A.GetValue("U_ColReg01", loopCount).ToString().Trim() == "N")
					{
						CheckType = "N";
						break; // TODO: might not be correct. Was : Exit For
					}
				}

				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					oDS_PS_SD061A.Offset = loopCount;
					if (CheckType == "N")
					{
						oDS_PS_SD061A.SetValue("U_ColReg01", loopCount, "Y");
					}
					else
					{
						oDS_PS_SD061A.SetValue("U_ColReg01", loopCount, "N");
					}
				}

				oMat.LoadFromDataSource();
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
		/// PS_SD061_MatrixSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_SD061_MatrixSpaceLineDel()
		{
			bool returnValue = false;
			string errMessage = string.Empty;
			int i;

			try
			{
				oMat.FlushToDataSource();
				
				for (i = 0; i <= oMat.VisualRowCount - 2; i++) //마지막 빈행 제외를 위해 2를 뺌
				{
					if (oDS_PS_SD061A.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
					{
						if (string.IsNullOrEmpty(oDS_PS_SD061A.GetValue("U_ColReg02", i).ToString().Trim())) //작번
						{
							errMessage = "" + i + 1 + "번 라인의 작번이 없습니다. 확인하세요.";
							throw new Exception();
						}
						else if (string.IsNullOrEmpty(oDS_PS_SD061A.GetValue("U_ColSum01", i).ToString().Trim()))
						{
							errMessage = "" + i + 1 + "번 라인의 금액이 없습니다. 확인하세요.";
							throw new Exception();
						}
						else if (string.IsNullOrEmpty(oDS_PS_SD061A.GetValue("U_ColReg05", i).ToString().Trim()))
						{
							errMessage = "" + i + 1 + "번 라인의 기준년월이 없습니다. 확인하세요.";
							throw new Exception();
						}
					}
				}
				returnValue = true;
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
			return returnValue;
		}

		/// <summary>
		/// PS_SD061_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_SD061_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			string ItemName;
			string ItemSpec;
			string Amount;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				switch (oUID)
				{
					case "Mat01":
						oMat.FlushToDataSource();

						if (oCol == "OrdNum")
						{
							oDS_PS_SD061A.SetValue("U_ColReg02", oRow - 1, oMat.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim());
							oDS_PS_SD061A.SetValue("U_ColReg05", oRow - 1, oForm.Items.Item("StdYear01").Specific.Value.ToString().Trim() + oForm.Items.Item("StdMth01").Specific.Value.ToString().Trim());

							//품목마스터 조회
							sQry = " SELECT  T0.FrgnName AS [ItemName], ";
							sQry += "         T0.U_Size AS [ItemSpec] ";
							sQry += " FROM    [OITM] AS T0 ";
							sQry += " WHERE   T0.ItemCode = '" + oMat.Columns.Item(oCol).Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							ItemName = oRecordSet.Fields.Item("ItemName").Value.ToString().Trim();
							ItemSpec = oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim();
							Amount = "0";

							oDS_PS_SD061A.SetValue("U_ColReg03", oRow - 1, ItemName); //품명
							oDS_PS_SD061A.SetValue("U_ColReg04", oRow - 1, ItemSpec); //규격
							oDS_PS_SD061A.SetValue("U_ColSum01", oRow - 1, Amount);	  //금액

							//행 추가
							if (oMat.RowCount == oRow & !string.IsNullOrEmpty(oDS_PS_SD061A.GetValue("U_ColReg02", oRow - 1).ToString().Trim()))
							{
								PS_SD061_AddMatrixRow(oRow, false);
							}
							oMat.Columns.Item("Amount").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}

						oMat.LoadFromDataSource();			
						oMat.Columns.Item(oCol).Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oMat.AutoResizeColumns();
						break;

					case "OrdNum01":
						oForm.Items.Item("ItemName01").Specific.Value = dataHelpClass.Get_ReData("FrgnName", "ItemCode", "OITM", "'" + oForm.Items.Item("OrdNum01").Specific.Value + "'", "");
						oForm.Items.Item("ItemSpec01").Specific.Value = dataHelpClass.Get_ReData("U_Size", "ItemCode", "OITM", "'" + oForm.Items.Item("OrdNum01").Specific.Value + "'", "");
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
		/// PS_SD061_MTX01
		/// 등록된 데이터 조회
		/// </summary>
		private void PS_SD061_MTX01()
		{
			int i;
			string sQry;
			string errMessage = string.Empty;
			string StdYear; //기준년도
			string StdMth;	//기준월
			string OrdNum;  //작번
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				StdYear = oForm.Items.Item("StdYear01").Specific.Value.ToString().Trim();
				StdMth = oForm.Items.Item("StdMth01").Specific.Value.ToString().Trim();
				OrdNum = oForm.Items.Item("OrdNum01").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회시작!";

				oForm.Freeze(true);

				sQry = " EXEC [PS_SD061_01] '";
				sQry +=  StdYear + "','";
				sQry +=  StdMth + "','";
				sQry +=  OrdNum + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_SD061A.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_SD061_AddMatrixRow(0, true);
					PS_SD061_LoadCaption();
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_SD061A.Size)
					{
						oDS_PS_SD061A.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_SD061A.Offset = i;

					oDS_PS_SD061A.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_SD061A.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("Check").Value.ToString().Trim());					//선택
					oDS_PS_SD061A.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());					//작번
					oDS_PS_SD061A.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());					//품명
					oDS_PS_SD061A.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim());					//규격
					oDS_PS_SD061A.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("Amount").Value.ToString().Trim());					//금액
					oDS_PS_SD061A.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("StdYM").Value.ToString().Trim());					//기준년월
					oDS_PS_SD061A.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("ChkAmtInc").Value.ToString().Trim());                   //검수금액포함

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				PS_SD061_AddMatrixRow(oMat.VisualRowCount, false);
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
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_SD061_MTX02
		/// 월별 조회
		/// </summary>
		private void PS_SD061_MTX02()
		{
			string sQry;
			string errMessage = string.Empty;
			string StdYear; //기준년도

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				StdYear = oForm.Items.Item("StdYear02").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회 중......!";

				sQry = " EXEC PS_SD061_04 '";
				sQry += StdYear + "'";

				oGrid.DataTable.Clear();
				oDS_PS_SD061B.ExecuteQuery(sQry);

				oGrid.Columns.Item(4).RightJustified = true;
				oGrid.Columns.Item(5).RightJustified = true;
				oGrid.Columns.Item(6).RightJustified = true;
				oGrid.Columns.Item(7).RightJustified = true;
				oGrid.Columns.Item(8).RightJustified = true;
				oGrid.Columns.Item(9).RightJustified = true;
				oGrid.Columns.Item(10).RightJustified = true;
				oGrid.Columns.Item(11).RightJustified = true;
				oGrid.Columns.Item(12).RightJustified = true;
				oGrid.Columns.Item(13).RightJustified = true;
				oGrid.Columns.Item(14).RightJustified = true;
				oGrid.Columns.Item(15).RightJustified = true;
				oGrid.Columns.Item(16).RightJustified = true;
				oGrid.Columns.Item(17).RightJustified = true;

				if (oGrid.Rows.Count == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				oGrid.AutoResizeColumns();
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
		/// PS_SD061_DeleteData
		/// 기본정보 삭제
		/// </summary>
		private void PS_SD061_DeleteData()
		{
			int loopCount;
			string sQry;

			string OrdNum;
			string StdYM;

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				ProgressBar01.Text = "삭제 중...";
				oMat.FlushToDataSource();
				//마지막 빈행 제외를 위해 2를 뺌
				for (loopCount = 0; loopCount <= oMat.VisualRowCount - 2; loopCount++)
				{
					if (oDS_PS_SD061A.GetValue("U_ColReg01", loopCount).ToString().Trim() == "Y")
					{
						StdYM  = oDS_PS_SD061A.GetValue("U_ColReg05", loopCount).ToString().Trim();
						OrdNum = oDS_PS_SD061A.GetValue("U_ColReg02", loopCount).ToString().Trim();

						sQry = "EXEC [PS_SD061_03] '";
						sQry += StdYM + "','";
						sQry += OrdNum + "'";
						oRecordSet.DoQuery(sQry);
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("삭제 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
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
		/// PS_SD061_AddData
		/// 데이터 INSERT, UPDATE(기존 데이터가 존재하면 UPDATE, 아니면 INSERT)
		/// </summary>
		/// <returns></returns>
		private bool PS_SD061_AddData()
		{
			bool returnValue = false;

			int i;
			string sQry;

			string StdYM;	 //기준년월
			string OrdNum;	 //품목코드(작번)
			string ItemName; //품명
			string ItemSpec; //규격
			decimal Amount;  //금액
			string ChkAmtInc; //검수금액포함

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				ProgressBar01.Text = "저장 중...";
				oMat.FlushToDataSource();
				//마지막 빈행 제외를 위해 2를 뺌
				for (i = 0; i <= oMat.VisualRowCount - 2; i++)
				{
					if (oDS_PS_SD061A.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
					{
						StdYM    = oDS_PS_SD061A.GetValue("U_ColReg05", i).ToString().Trim(); //기준년월
						OrdNum   = oDS_PS_SD061A.GetValue("U_ColReg02", i).ToString().Trim(); //품목코드(작번)
						ItemName = oDS_PS_SD061A.GetValue("U_ColReg03", i).ToString().Trim().Replace("'", "`"); //품명
						ItemSpec = oDS_PS_SD061A.GetValue("U_ColReg04", i).ToString().Trim().Replace("'", "`"); //규격
						Amount = Convert.ToDecimal(oDS_PS_SD061A.GetValue("U_ColSum01", i).ToString().Trim()); //금액
						ChkAmtInc = oDS_PS_SD061A.GetValue("U_ColReg06", i).ToString().Trim(); //검수금액포함

						sQry = "EXEC [PS_SD061_02] '";
						sQry += StdYM + "','";
						sQry += OrdNum + "','";
						sQry += ItemName + "','";
						sQry += ItemSpec + "','";
						sQry += Amount + "','";
						sQry += ChkAmtInc + "'";
						oRecordSet.DoQuery(sQry);

						ProgressBar01.Value += 1;
						ProgressBar01.Text = ProgressBar01.Value + "/" + Convert.ToInt32(oMat.VisualRowCount - 1) + "건 저장중...";
					}
				}
				PSH_Globals.SBO_Application.StatusBar.SetText("저장 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				returnValue = true;
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
			return returnValue;
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
					if (pVal.ItemUID == "BtnAdd01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_SD061_MatrixSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_SD061_AddData() == false)
							{
								BubbleEvent = false;
								return;
							}
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD061_LoadCaption();
							PS_SD061_MTX01();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							PS_SD061_ResetForm();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD061_LoadCaption();
							PS_SD061_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSrch01")
					{
						PS_SD061_MTX01();
					}
					else if (pVal.ItemUID == "BtnSrch02")
					{
						PS_SD061_MTX02();
					}
					else if (pVal.ItemUID == "BtnDel01")
					{
						if (PSH_Globals.SBO_Application.MessageBox("삭제 후에는 복구가 불가능합니다. 삭제하시겠습니까?", 1, "예", "아니오") == 1)
						{
							PS_SD061_DeleteData();
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
							PS_SD061_LoadCaption();
							PS_SD061_MTX01();
						}
					}
					else if (pVal.ItemUID == "BtnSel01")
					{
						PS_SD061_CheckAll();
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "Folder01")
					{
						oForm.PaneLevel = 1;
						oForm.DefButton = "BtnSrch01";
						oForm.Items.Item("BtnAdd01").Enabled = true;
						oForm.Items.Item("BtnDel01").Enabled = true;
					}
					if (pVal.ItemUID == "Folder02")
					{
						oForm.PaneLevel = 2;
						oForm.DefButton = "BtnSrch02";
						oForm.Items.Item("BtnAdd01").Enabled = false;
						oForm.Items.Item("BtnDel01").Enabled = false;
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum01", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdNum");
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
					PS_SD061_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
							PS_SD061_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
						}
						else
						{
							PS_SD061_FlushToItemValue(pVal.ItemUID, 0, "");
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
					PS_SD061_ResizeForm();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD061A);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD061B);
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
							PS_SD061_ResetForm();
							break;
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
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
							break;
						case "1282": //추가
							break;
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							break;
						case "1287": //복제
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
	}
}
