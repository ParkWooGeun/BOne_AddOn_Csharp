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
		private SAPbouiCOM.BoFormMode oForm1_Mode;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM006.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
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

				CreateItems();
				ComboBox_Setting();
				Initialization();
				LoadCaption();

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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;
				oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

				oForm.DataSources.DataTables.Add("PS_MM006");

				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_MM006");
				oDS_PS_MM006H = oForm.DataSources.DataTables.Item("PS_MM006");

				oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
				oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.AddMonths(-3).ToString("yyyyMM") + "01";

				oForm.DataSources.UserDataSources.Add("DocDateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateTo").Specific.DataBind.SetBound(true, "", "DocDateTo");
				oForm.DataSources.UserDataSources.Item("DocDateTo").Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		private void ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 품목구분
				sQry = "SELECT Code, Name From [@PSH_ORDTYP] Order by Code";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("OrdType").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim().ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("OrdType").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);

				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim().ToString().Trim());
					oRecordSet.MoveNext();
				}

				//대분류
				sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("ItmBSort").Specific.ValidValues.Add("ALL", "ALL");
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim().ToString().Trim());
					oRecordSet.MoveNext();
				}

				//품목타입
				sQry = "SELECT Code, Name From [@PSH_SHAPE] Order by Code";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("ALL", "ALL");
				while (!(oRecordSet.EoF))
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
		/// Initialization
		/// </summary>
		private void Initialization()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//아이디별 사업장 세팅
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//아이디별 사번 세팅
				oForm.Items.Item("CntcCode").Specific.Value = "";
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
				if (oForm1_Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("Btn01").Specific.Caption = "확인";
				}
				else if (oForm1_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
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
				switch (oUID)
				{
					case "CntcCode":
						sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim().ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim();
						break;
					case "ItemCode":
						sQry = "Select ItemName From OITM Where ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim().ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim().ToString().Trim();
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
		/// HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
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
				if (oForm1_Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// Update_PurchaseDemand
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool Update_PurchaseDemand(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool functionReturnValue = false;

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
							OkDate = oDS_PS_MM006H.Columns.Item("결재일").Cells.Item(i).Value.ToString().Trim();
							CgNum = oDS_PS_MM006H.Columns.Item("청구번호").Cells.Item(i).Value.ToString().Trim();

							sQry = "UPDATE [@PS_MM005H] ";
							sQry += "SET ";
							sQry += "U_OKYN = '" + OkYN + "', ";
                            if (string.IsNullOrEmpty(OkDate))
                            {
                                sQry += "U_OKDate = NULL ";
                            }
                            else
                            {
								sQry += "U_OKDate = '" + OkDate.Substring(0, 10) + "' ";
							}
                            sQry += " Where DocEntry = '" + CgNum + "' ";

							oRecordSet.DoQuery(sQry);

							// 이력저장
							if (OkYN == "Y")
							{
								sQry = "Exec [PS_Table_history] '" + CgNum + "','MM005','" + PSH_Globals.oCompany.UserSignature.ToString() + "'";
								oRecordSet.DoQuery(sQry);
							}
						}
					}

					PSH_Globals.SBO_Application.StatusBar.SetText("구매요청승인 변경 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
					oForm.Items.Item("Btn02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("데이터가 존재하지 않습니다.!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}

				functionReturnValue = true;
				oForm1_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// LoadData
		/// </summary>
		private void LoadData()
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
			int iRow;
			SAPbouiCOM.ProgressBar ProgressBar = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				OrdType = oForm.Items.Item("OrdType").Specific.Value.ToString().Trim().ToString().Trim();
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim().ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim().ToString().Trim();
				DeptCode = oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim().ToString().Trim();
				DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim().ToString().Trim();
				DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim().ToString().Trim();
				CgNumFr = oForm.Items.Item("CgNumFr").Specific.Value.ToString().Trim().ToString().Trim();
				CgNumTo = oForm.Items.Item("CgNumTo").Specific.Value.ToString().Trim().ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim().ToString().Trim();
				ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim().ToString().Trim();
				ItmMsort = oForm.Items.Item("ItmMSort").Specific.Value.ToString().Trim().ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim().ToString().Trim();
				OkYN = oForm.Items.Item("OKYN").Specific.Value.ToString().Trim().ToString().Trim();

				if (string.IsNullOrEmpty(OrdType))
				{
					OrdType = "%";
				}
				if (string.IsNullOrEmpty(BPLId))
				{
					BPLId = "%";
				}
				if (string.IsNullOrEmpty(CntcCode))
				{
					CntcCode = "%";
				}
				if (string.IsNullOrEmpty(DeptCode))
				{
					DeptCode = "%";
				}
				if (string.IsNullOrEmpty(DocDateFr))
				{
					DocDateFr = DateTime.Now.AddMonths(-3).ToString("yyyy-MM-") + "01";
				}
				if (string.IsNullOrEmpty(DocDateTo))
				{
					DocDateTo = DateTime.Now.ToString("yyyy-MM-dd");
				}
				if (string.IsNullOrEmpty(CgNumFr))
				{
					CgNumFr = "0000000000";
				}
				if (string.IsNullOrEmpty(CgNumTo))
				{
					CgNumTo = "9999999999";
				}
				if (string.IsNullOrEmpty(ItemCode))
				{
					ItemCode = "%";
				}
				if (string.IsNullOrEmpty(ItmBsort) || ItmBsort == "ALL")
				{
					ItmBsort = "%";
				}
				if (string.IsNullOrEmpty(ItmMsort) || ItmMsort == "ALL")
				{
					ItmMsort = "%";
				}
				if (string.IsNullOrEmpty(ItemType) || ItemType == "ALL")
				{
					ItemType = "%";
				}
				if (string.IsNullOrEmpty(OkYN) | OkYN == "ALL")
				{
					OkYN = "%";
				}

				ProgressBar.Text = "조회시작!";

				sQry = "EXEC [PS_MM006_01] '" + OrdType + "','" + BPLId + "','" + CntcCode + "','" + DeptCode + "','" + DocDateFr + "',";
				sQry += "'" + DocDateTo + "','" + CgNumFr + "','" + CgNumTo + "','" + ItemCode + "','" + ItmBsort + "','" + ItmMsort + "','" + ItemType + "','" + OkYN + "','" + PSH_Globals.oCompany.UserName.ToString() + "'";

				oDS_PS_MM006H.ExecuteQuery(sQry);

				iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;

				TitleSetting(iRow);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				ProgressBar.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// TitleSetting
		/// </summary>
		/// <param name="iRow"></param>
		private void TitleSetting(int iRow)
		{
			int i;
			int ColumnCnt;
			string BPLId;

			SAPbouiCOM.ComboBoxColumn oComboCol;

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
							//결재일 수정 불가(2015.05.13 송명규 수정, 류석균 요청)
							oGrid.Columns.Item(i).RightJustified = true;
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
			int i;

            try
            {
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Btn01")
					{
						if (oForm1_Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (Update_PurchaseDemand(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}

							oForm1_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							LoadCaption();
						}
						else if (oForm1_Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							oForm.Close();
						}
					}
					else if (pVal.ItemUID == "Btn02")
					{
						if (HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}

						LoadData();

						oForm1_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
						LoadCaption();
					}
					else if (pVal.ItemUID == "Btn03")
					{
						if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
						{
							oForm.Freeze(true);

							for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
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
						oForm1_Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						LoadCaption();
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
						oForm.Freeze(true);
						oDS_PS_MM006H.Clear();
						oForm.Freeze(false);
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
						oForm1_Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						LoadCaption();
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
							FlushToItemValue(pVal.ItemUID, 0, "");
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							FlushToItemValue(pVal.ItemUID, 0, "");
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

							oForm1_Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							LoadCaption();
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
					oGrid.AutoResizeColumns();
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
					if (oForm != null)
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
