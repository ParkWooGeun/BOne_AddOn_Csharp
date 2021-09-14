using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 구매요청승인
	/// </summary>
	internal class PS_MM006 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DataTable oDS_PS_MM006H;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM006.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM006_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM006");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);

				PS_MM006_CreateItems();
				PS_MM006_SetComboBox();
				PS_MM006_Initialize();
				PS_MM006_LoadCaption();

				oForm.EnableMenu("1281", false); // 찾기
				oForm.EnableMenu("1282", false); // 추가
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", false); // 취소
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_MM006_CreateItems
		/// </summary>
		private void PS_MM006_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;
				oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

				oForm.DataSources.DataTables.Add("PS_MM006");

				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_MM006");
				oDS_PS_MM006H = oForm.DataSources.DataTables.Item("PS_MM006");

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//청구인(사번)
				oForm.DataSources.UserDataSources.Add("CntcCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CntcCode").Specific.DataBind.SetBound(true, "", "CntcCode");

				//청구인(성명)
				oForm.DataSources.UserDataSources.Add("CntcName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CntcName").Specific.DataBind.SetBound(true, "", "CntcName");

				//품목구분
				oForm.DataSources.UserDataSources.Add("OrdType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OrdType").Specific.DataBind.SetBound(true, "", "OrdType");

				//사용처부서
				oForm.DataSources.UserDataSources.Add("DeptCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
				oForm.Items.Item("DeptCode").Specific.DataBind.SetBound(true, "", "DeptCode");

				//요청일자(FR)
				oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
				
				//요청일자(TO)
				oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
				
				//청구번호(FR)
				oForm.DataSources.UserDataSources.Add("CgNumFr", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CgNumFr").Specific.DataBind.SetBound(true, "", "CgNumFr");

				//청구번호(TO)
				oForm.DataSources.UserDataSources.Add("CgNumTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CgNumTo").Specific.DataBind.SetBound(true, "", "CgNumTo");

				//품목코드
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				//품목이름
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//대분류
				oForm.DataSources.UserDataSources.Add("ItmBSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItmBSort").Specific.DataBind.SetBound(true, "", "ItmBSort");

				//중분류
				oForm.DataSources.UserDataSources.Add("ItmMSort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItmMSort").Specific.DataBind.SetBound(true, "", "ItmMSort");

				//품목타입
				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				//결재여부
				oForm.DataSources.UserDataSources.Add("OKYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 5);
				oForm.Items.Item("OKYN").Specific.DataBind.SetBound(true, "", "OKYN");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM006_SetComboBox
		/// </summary>
		private void PS_MM006_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 품목구분
				sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("OrdType").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim().ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("OrdType").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim().ToString().Trim());
					oRecordSet.MoveNext();
				}

				//대분류
				sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("ItmBSort").Specific.ValidValues.Add("ALL", "ALL");
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim().ToString().Trim());
					oRecordSet.MoveNext();
				}

				//품목타입
				sQry = "SELECT Code, Name From [@PSH_SHAPE] Order by Code";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("ALL", "ALL");
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("ItemType").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim().ToString().Trim());
					oRecordSet.MoveNext();
				}

				//결재여부
				oForm.Items.Item("OKYN").Specific.ValidValues.Add("Y", "결재");
				oForm.Items.Item("OKYN").Specific.ValidValues.Add("N", "미결재");
				oForm.Items.Item("OKYN").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
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
		/// PS_MM006_Initialize
		/// </summary>
		private void PS_MM006_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.DataSources.UserDataSources.Item("BPLId").Value = dataHelpClass.User_BPLID(); //사업장
				oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.AddMonths(-3).ToString("yyyyMM") + "01"; //요청일자(FR)
				oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd"); //요청일자(TO)
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM006_LoadCaption
		/// </summary>
		private void PS_MM006_LoadCaption()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("Btn01").Specific.Caption = "확인";
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
				{
					oForm.Items.Item("Btn01").Specific.Caption = "갱신";
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM006_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_MM006_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "CntcCode":
						sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.DataSources.UserDataSources.Item("CntcName").Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
					case "ItemCode":
						sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.DataSources.UserDataSources.Item("ItemName").Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
		/// PS_MM006_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_MM006_DelHeaderSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("OrdType").Specific.Value.ToString().Trim().ToString().Trim()))
				{
					errMessage = "품목구분은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim().ToString().Trim()))
				{
					errMessage = "사업장은 필수사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim().ToString().Trim()))
					{
						errMessage = "청구인은 필수사항입니다. 확인하세요.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim().ToString().Trim()))
					{
						errMessage = "청구부서는 필수사항입니다. 확인하세요.";
						throw new Exception();
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
				}
			}
			return returnValue;
		}

		/// <summary>
		/// PS_MM006_UpdatePurchaseDemand
		/// </summary>
		/// <returns></returns>
		private bool PS_MM006_UpdatePurchaseDemand()
		{
			bool returnValue = false;
			short i;
			string sQry;
			string OkDate;
			string OkYN;
			string CgNum;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
				{
					for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
					{
						if (oDS_PS_MM006H.Columns.Item("견적여부").Cells.Item(i).Value == "N")
						{
							OkYN = oDS_PS_MM006H.Columns.Item("결재여부").Cells.Item(i).Value.ToString().Trim();
							OkDate = (OkYN == "N") ? "" : oDS_PS_MM006H.Columns.Item("결재일").Cells.Item(i).Value.ToString("yyyyMMdd");
							CgNum = oDS_PS_MM006H.Columns.Item("청구번호").Cells.Item(i).Value.ToString().Trim();

							sQry = "UPDATE [@PS_MM005H] ";
							sQry += "SET ";
							sQry += "U_OKYN = '" + OkYN + "', ";
                            if (OkDate == "")
                            {
                                sQry += "U_OKDate = NULL ";
                            }
                            else
                            {
								sQry += "U_OKDate = '" + OkDate + "' ";
							}
                            sQry += " Where DocEntry = '" + CgNum + "' ";

							oRecordSet.DoQuery(sQry);

							// 이력저장
							if (OkYN == "Y")
							{
								sQry = "Exec [PS_Table_history] '" + CgNum + "','MM005','" + PSH_Globals.oCompany.UserSignature + "'";
								oRecordSet.DoQuery(sQry);
							}
						}
					}

					PS_MM006_LoadData("구매요청승인 변경 완료");
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("데이터가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}

				returnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return returnValue;
		}

		/// <summary>
		/// 메트릭스 데이터 조회
		/// </summary>
		/// <param name="message">SetText 메시지</param>
		private void PS_MM006_LoadData(string message)
		{
			string sQry;
			string ItemType;
			string ItmBsort;
			string CgNumTo;
			string DeptCode;
			string BPLId;
			string OrdType;
			string CntcCode;
			string CgNumFr;
			string ItemCode;
			string ItmMsort;
			string OkYN;
			string DocDateFr;
			string DocDateTo;
			SAPbouiCOM.ProgressBar ProgressBar = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				OrdType = string.IsNullOrEmpty(oForm.Items.Item("OrdType").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("OrdType").Specific.Value.ToString().Trim();
				BPLId = string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CntcCode = string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				DeptCode = string.IsNullOrEmpty(oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim();
				DocDateFr = string.IsNullOrEmpty(oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim()) ? DateTime.Now.AddMonths(-3).ToString("yyyyMM") + "01" : oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
				DocDateTo = string.IsNullOrEmpty(oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim()) ? DocDateTo = DateTime.Now.ToString("yyyyMMdd") : oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
				CgNumFr = string.IsNullOrEmpty(oForm.Items.Item("CgNumFr").Specific.Value.ToString().Trim()) ? "0000000000" : oForm.Items.Item("CgNumFr").Specific.Value.ToString().Trim();
				CgNumTo = string.IsNullOrEmpty(oForm.Items.Item("CgNumTo").Specific.Value.ToString().Trim()) ? "9999999999" : oForm.Items.Item("CgNumTo").Specific.Value.ToString().Trim();
				ItemCode = string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				ItmBsort = string.IsNullOrEmpty(oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();
				ItmMsort = string.IsNullOrEmpty(oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim();
				ItemType = string.IsNullOrEmpty(oForm.Items.Item("ItemType").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				OkYN = string.IsNullOrEmpty(oForm.Items.Item("OKYN").Specific.Value.ToString().Trim()) ? "%" : oForm.Items.Item("OKYN").Specific.Value.ToString().Trim();

				sQry = "EXEC [PS_MM006_01] '";
				sQry += OrdType + "','";
				sQry += BPLId + "','";
				sQry += CntcCode + "','";
				sQry += DeptCode + "','";
				sQry += DocDateFr + "','";
				sQry += DocDateTo + "','";
				sQry += CgNumFr + "','";
				sQry += CgNumTo + "','";
				sQry += ItemCode + "','";
				sQry += ItmBsort + "','";
				sQry += ItmMsort + "','";
				sQry += ItemType + "','";
				sQry += OkYN + "','";
				sQry += PSH_Globals.oCompany.UserName + "'";

				oDS_PS_MM006H.ExecuteQuery(sQry);

				PS_MM006_SetTitle();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar);
				if (message != "")
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				}
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_MM006_SetTitle
		/// </summary>
		private void PS_MM006_SetTitle()
		{
			int i;
			int ColumnCnt;
			string BPLId;

			SAPbouiCOM.ComboBoxColumn oComboCol = null;

			try
			{
				oForm.Freeze(true);

				ColumnCnt = Convert.ToInt32(oDS_PS_MM006H.Columns.Item("ColumnCnt").Cells.Item(0).Value.ToString().Trim());
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim().ToString().Trim();

				for (i = 0; i <= ColumnCnt; i++)
				{
					switch (oGrid.Columns.Item(i).TitleObject.Caption)
					{
						case "본수":
							oGrid.Columns.Item(i).RightJustified = true;
							oGrid.Columns.Item(i).Editable = false;
							break;
						case "수량/중량":
							oGrid.Columns.Item(i).RightJustified = true;
							oGrid.Columns.Item(i).Editable = false;
							break;
						case "결재일":
							oGrid.Columns.Item(i).Editable = false;
							break;
						case "결재여부":
							oGrid.Columns.Item(i).Editable = true;
							oGrid.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
							oComboCol = (ComboBoxColumn)oGrid.Columns.Item("결재여부");
							oComboCol.ValidValues.Add("Y", "결재");
							oComboCol.ValidValues.Add("N", "미결재");
							oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
							break;
						case "최종입고수량":
							oGrid.Columns.Item(i).RightJustified = true;
							oGrid.Columns.Item(i).Editable = false;
							break;
						case "현재고수량":
							oGrid.Columns.Item(i).RightJustified = true;
							oGrid.Columns.Item(i).Editable = false;
							break;
						default:
							oGrid.Columns.Item(i).Editable = false;
							break;
					}
				}

				oGrid.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Freeze(false);
				if (oComboCol != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oComboCol);
				}
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
        /// ITEM_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM006_UpdatePurchaseDemand() == false)
							{
								BubbleEvent = false;
								return;
							}

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            PS_MM006_LoadCaption();
                        }
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							oLastItemUID01 = pVal.ItemUID;
							oForm.Close();
						}
					}
					else if (pVal.ItemUID == "Btn02")
					{
						if (PS_MM006_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}

						PS_MM006_LoadData("");

						oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
						PS_MM006_LoadCaption();
					}
					else if (pVal.ItemUID == "Btn03")
					{
						if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
						{
							oForm.Freeze(true);

							for (int i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
							{
								if (oDS_PS_MM006H.Columns.Item("견적여부").Cells.Item(i).Value.ToString().Trim() == "N")
								{
									if (oGrid.DataTable.GetValue("결재여부", i).ToString().Trim() == "Y")
									{
										oGrid.DataTable.Columns.Item("결재여부").Cells.Item(i).Value = "N";
										oDS_PS_MM006H.Columns.Item("결재일").Cells.Item(i).Value = "";
									}
									else
									{
										oGrid.DataTable.Columns.Item("결재여부").Cells.Item(i).Value = "Y";
										oDS_PS_MM006H.Columns.Item("결재일").Cells.Item(i).Value = DateTime.Now.ToString("yyyyMMdd");
									}
								}
							}
							oForm.Freeze(false);
						}
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						PS_MM006_LoadCaption();
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
		/// KEY_DOWN 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.Before_Action == true)
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
						else if (pVal.ItemUID == "ItemCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
							{
								PS_SM010 ChildForm01 = new PS_SM010();
								ChildForm01.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row);
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "DeptCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
					}
				}
				else if (pVal.Before_Action == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// COMBO_SELECT 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
			try
			{
				oForm.Freeze(true);

				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemUID == "OrdType" || pVal.ItemUID == "BPLId")
					{
						oDS_PS_MM006H.Clear();
					}
					else if (pVal.ItemUID == "Grid01")
					{
						if (oDS_PS_MM006H.Columns.Item("견적여부").Cells.Item(pVal.Row).Value.ToString().Trim() == "Y")
						{
							oDS_PS_MM006H.Columns.Item("결재여부").Cells.Item(pVal.Row).Value = oDS_PS_MM006H.Columns.Item("OKYN").Cells.Item(pVal.Row).Value.ToString().Trim();
							oDS_PS_MM006H.Columns.Item("결재일").Cells.Item(pVal.Row).Value = oDS_PS_MM006H.Columns.Item("OKDate").Cells.Item(pVal.Row).Value;
						}
						else
						{
							if (oDS_PS_MM006H.Columns.Item("결재여부").Cells.Item(pVal.Row).Value.ToString().Trim() == "Y")
							{
								oDS_PS_MM006H.Columns.Item("결재일").Cells.Item(pVal.Row).Value = DateTime.Now.ToString("yyyyMMdd");
							}
							else
							{
								oDS_PS_MM006H.Columns.Item("결재일").Cells.Item(pVal.Row).Value = "";
							}
						}
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						PS_MM006_LoadCaption();
						oGrid.AutoResizeColumns();
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
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
						if (pVal.ItemUID == "CntcCode")
						{
							PS_MM006_FlushToItemValue(pVal.ItemUID, 0, "");
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							PS_MM006_FlushToItemValue(pVal.ItemUID, 0, "");
						}
						else if (pVal.ItemUID == "Grid01")
						{
							if (oDS_PS_MM006H.Columns.Item("견적여부").Cells.Item(pVal.Row).Value.ToString().Trim() == "Y")
							{
								oDS_PS_MM006H.Columns.Item("결재여부").Cells.Item(pVal.Row).Value = oDS_PS_MM006H.Columns.Item("OKYN").Cells.Item(pVal.Row).Value.ToString().Trim();
								oDS_PS_MM006H.Columns.Item("결재일").Cells.Item(pVal.Row).Value = oDS_PS_MM006H.Columns.Item("OKDate").Cells.Item(pVal.Row).Value.ToString().Trim();
							}
							else
							{
							}

							oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_MM006_LoadCaption();
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
		/// FORM_RESIZE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (oGrid.Rows.Count > 0)
					{
						oGrid.AutoResizeColumns();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText("Raise_EVENT_FORM_RESIZE_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					if (oLastItemUID01 != "Btn01") //확인 버튼을 클릭해서 Form을 닫을 경우 제외
					{
						System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					}
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oGrid);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM006H);
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
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
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
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
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
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD: //33
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }
    }
}