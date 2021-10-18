using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 외주입고 품질검사
	/// </summary>
	internal class PS_MM153 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Grid oGrid;
		private SAPbouiCOM.DataTable oDS_PS_MM153H;
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM153.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM153_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM153");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);

				PS_MM153_CreateItems();
				PS_MM153_ComboBox_Setting();
				PS_MM153_Initialization();
				PS_MM153_LoadCaption();

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
		/// PS_MM153_CreateItems
		/// </summary>
		private void PS_MM153_CreateItems()
		{
			try
			{
				oGrid = oForm.Items.Item("Grid01").Specific;
				oForm.DataSources.DataTables.Add("PS_MM153");
				oGrid.DataTable = oForm.DataSources.DataTables.Item("PS_MM153");
				oDS_PS_MM153H = oForm.DataSources.DataTables.Item("PS_MM153");

				oForm.DataSources.UserDataSources.Add("DocDateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDateFr").Specific.DataBind.SetBound(true, "", "DocDateFr");
				oForm.DataSources.UserDataSources.Item("DocDateFr").Value = DateTime.Now.ToString("yyyyMMdd");

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
		/// PS_MM153_ComboBox_Setting
		/// </summary>
		private void PS_MM153_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
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
		/// PS_MM153_Initialization
		/// </summary>
		private void PS_MM153_Initialization()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//아이디별 사업장 세팅
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//아이디별 사번 세팅
				oForm.Items.Item("CardCode").Specific.Value = "";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM153_LoadCaption
		/// </summary>
		private void PS_MM153_LoadCaption()
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
		/// PS_MM153_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_MM153_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "CardCode":
						sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() +"'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
		/// PS_MM153_Update_PurchaseDemand
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool PS_MM153_Update_PurchaseDemand(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool returnValue = false;

			int i;
			string DocNo;
			string OkYN;
			string OkDate;
			string LineNum;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
				{
					for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
					{
						if (oDS_PS_MM153H.Columns.Item("청구여부").Cells.Item(i).Value.ToString().Trim() == "N")
						{
							OkYN = oDS_PS_MM153H.Columns.Item("검사여부").Cells.Item(i).Value.ToString().Trim();
							OkDate = Convert.ToString(oDS_PS_MM153H.Columns.Item("검사일").Cells.Item(i).Value);
							DocNo = oDS_PS_MM153H.Columns.Item("입고번호").Cells.Item(i).Value.ToString().Trim();
							LineNum = oDS_PS_MM153H.Columns.Item("입고순번").Cells.Item(i).Value.ToString().Trim();

							sQry = "UPDATE [@PS_MM152L] ";
							sQry += "SET ";
							sQry += "U_QCOKYN = '" + OkYN + "', ";
							if (string.IsNullOrEmpty(OkDate))
							{
								sQry += "U_QCOKDate = NULL ";
							}
							else
							{
								sQry += "U_QCOKDate = '" + OkDate.Substring(0, 10) + "' ";
							}
							sQry += " Where DocEntry = '" + DocNo + "' and U_LineNum = '" + LineNum + "' ";
							oRecordSet.DoQuery(sQry);
						}
					}

					PSH_Globals.SBO_Application.StatusBar.SetText("외주입고 품질검사 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
			//		oForm.Items.Item("Btn02").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("데이터가 존재하지 않습니다.!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}

				returnValue = true;
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
			return returnValue;
		}

		/// <summary>
		/// PS_MM153_LoadData
		/// </summary>
		private void PS_MM153_LoadData()
		{
			string CardCode;
			string BPLID;
			string OkYN;
			string DocDateFr;
			string DocDateTo;
			int iRow;
			string sQry;
			SAPbouiCOM.ProgressBar ProgressBar = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				DocDateFr = oForm.Items.Item("DocDateFr").Specific.Value.ToString().Trim();
				DocDateTo = oForm.Items.Item("DocDateTo").Specific.Value.ToString().Trim();
				OkYN = oForm.Items.Item("OKYN").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(BPLID))
                {
					BPLID = "%";
				}
				if (string.IsNullOrEmpty(CardCode))
                {
					CardCode = "%";
				}
				if (string.IsNullOrEmpty(DocDateFr))
                {
					DocDateFr = DateTime.Now.AddMonths(-3).ToString("yyyy-MM-") + "01";
				}
				if (string.IsNullOrEmpty(DocDateTo))
                {
					DocDateTo = DateTime.Now.ToString("yyyy-MM-dd");
				}
				if (string.IsNullOrEmpty(OkYN) || OkYN == "ALL")
                {
					OkYN = "%";
				}

				ProgressBar.Text = "조회시작!";

				sQry = "EXEC [PS_MM153_01] '" + BPLID + "','" + CardCode + "','" + DocDateFr + "','" + DocDateTo + "','" + OkYN + "'";

				oDS_PS_MM153H.ExecuteQuery(sQry);
				iRow = oForm.DataSources.DataTables.Item(0).Rows.Count;
				PS_MM153_TitleSetting(iRow);

				oGrid.Columns.Item(6).RightJustified = true;
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
		/// PS_MM153_TitleSetting
		/// </summary>
		/// <param name="iRow"></param>
		private void PS_MM153_TitleSetting(int iRow)
		{
			int i;
			int ColumnCnt;
			SAPbouiCOM.ComboBoxColumn oComboCol;

			try
			{
				oForm.Freeze(true);

				ColumnCnt = Convert.ToInt32(oDS_PS_MM153H.Columns.Item("ColumnCnt").Cells.Item(0).Value.ToString().Trim());

				for (i = 0; i <= ColumnCnt; i++)
				{
					switch (oGrid.Columns.Item(i).TitleObject.Caption)
					{
						case "검사일":
							oGrid.Columns.Item(i).Editable = true;
							oGrid.Columns.Item(i).RightJustified = true;
							break;
						case "검사여부":
							oGrid.Columns.Item(i).Editable = true;
							oGrid.Columns.Item(i).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox;
							oComboCol = (ComboBoxColumn)oGrid.Columns.Item("검사여부");

							oComboCol.ValidValues.Add("Y", "검사");
							oComboCol.ValidValues.Add("N", "미검사");

							oComboCol.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description;
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
                   // Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
							if (PS_MM153_Update_PurchaseDemand(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}

							oForm1_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PS_MM153_LoadCaption();
						}
						else if (oForm1_Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							oForm.Close();
						}
					}
					else if (pVal.ItemUID == "Btn02")
					{
						PS_MM153_LoadData();

						oForm1_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
						PS_MM153_LoadCaption();
					}
					else if (pVal.ItemUID == "Btn03")
					{
						if (oForm.DataSources.DataTables.Item(0).Rows.Count > 0)
						{
							oForm.Freeze(true);
							for (i = 0; i <= oForm.DataSources.DataTables.Item(0).Rows.Count - 1; i++)
							{
								if (oDS_PS_MM153H.Columns.Item("청구여부").Cells.Item(i).Value.ToString().Trim() == "N")
								{
									if (oGrid.DataTable.GetValue("검사여부", i).ToString().Trim() == "Y")
									{
										oGrid.DataTable.Columns.Item("검사여부").Cells.Item(i).Value = "N";
										oDS_PS_MM153H.Columns.Item("검사일").Cells.Item(i).Value = "";
									}
									else
									{
										oGrid.DataTable.Columns.Item("검사여부").Cells.Item(i).Value = "Y";
										oDS_PS_MM153H.Columns.Item("검사일").Cells.Item(i).Value = DateTime.Now.ToString("yyyyMMdd");
									}
								}
							}
							oForm.Freeze(false);
						}
						oForm1_Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						PS_MM153_LoadCaption();
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
			string errMessage = string.Empty;

			try
			{
				oForm.Freeze(true);

				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemUID == "BPLId")
					{
						oDS_PS_MM153H.Clear();
					}
					else if (pVal.ItemUID == "Grid01")
					{
						if (oDS_PS_MM153H.Columns.Item("청구여부").Cells.Item(pVal.Row).Value.ToString().Trim() == "Y")
						{
							if (string.IsNullOrEmpty(oDS_PS_MM153H.Columns.Item("OKDate").Cells.Item(pVal.Row).Value))
							{
								errMessage = "OKDate가 있어야 합니다..";
								throw new Exception();
							}
							oDS_PS_MM153H.Columns.Item("검사여부").Cells.Item(pVal.Row).Value = oDS_PS_MM153H.Columns.Item("OKYN").Cells.Item(pVal.Row).Value;
							oDS_PS_MM153H.Columns.Item("검사일").Cells.Item(pVal.Row).Value = oDS_PS_MM153H.Columns.Item("OKDate").Cells.Item(pVal.Row).Value;
						}
						else
						{
							if (oDS_PS_MM153H.Columns.Item("검사여부").Cells.Item(pVal.Row).Value.ToString().Trim() == "Y")
							{
								oDS_PS_MM153H.Columns.Item("검사일").Cells.Item(pVal.Row).Value = DateTime.Now.ToString("yyyyMMdd");
							}
							else
							{
								oDS_PS_MM153H.Columns.Item("검사일").Cells.Item(pVal.Row).Value = "";
							}
						}
						oForm1_Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
						PS_MM153_LoadCaption();
					}
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
							PS_MM153_FlushToItemValue(pVal.ItemUID, 0, "");
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							PS_MM153_FlushToItemValue(pVal.ItemUID, 0, "");
						}
						else if (pVal.ItemUID == "Grid01")
						{
							if (oDS_PS_MM153H.Columns.Item("청구여부").Cells.Item(pVal.Row).Value.ToString().Trim() == "Y")
							{
								oDS_PS_MM153H.Columns.Item("검사여부").Cells.Item(pVal.Row).Value = oDS_PS_MM153H.Columns.Item("OKYN").Cells.Item(pVal.Row).Value.ToString().Trim();
								oDS_PS_MM153H.Columns.Item("검사일").Cells.Item(pVal.Row).Value = oDS_PS_MM153H.Columns.Item("OKDate").Cells.Item(pVal.Row).Value;
							}
							else
							{
							}

							oForm1_Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							PS_MM153_LoadCaption();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM153H);
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
