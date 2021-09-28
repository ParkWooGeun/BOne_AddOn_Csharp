using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 휘팅벌크포장
	/// </summary>
	internal class PS_PP070 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;			
		private SAPbouiCOM.DBDataSource oDS_PS_PP070H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP070L; //등록라인
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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP070.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP070_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP070");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry"; 

				oForm.Freeze(true);

				PS_PP070_CreateItems();
				PS_PP070_SetComboBox();
				PS_PP070_EnableMenus();
				PS_PP070_SetDocument(oFormDocEntry);
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
		/// PS_PP070_CreateItems
		/// </summary>
		private void PS_PP070_CreateItems()
		{
			try
			{
				oDS_PS_PP070H = oForm.DataSources.DBDataSources.Item("@PS_PP070H");
				oDS_PS_PP070L = oForm.DataSources.DBDataSources.Item("@PS_PP070L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP070_SetComboBox
		/// </summary>
		private void PS_PP070_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", false, false);

				oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP070_EnableMenus
		/// </summary>
		private void PS_PP070_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP070_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_PP070_SetDocument(string oFormDocEntry)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_PP070_EnableFormItem();
					PS_PP070_AddMatrixRow(0, true);
				}
				else
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP070_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP070_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);

				if (RowIserted == false)
				{
					oDS_PS_PP070L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP070L.Offset = oRow;
				oDS_PS_PP070L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_PP070_EnableFormItem
		/// </summary>
		private void PS_PP070_EnableFormItem()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_PP070_ClearForm();
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("BPLId").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
					oForm.Items.Item("OrdGbn").Specific.Select("101", SAPbouiCOM.BoSearchKey.psk_ByValue);
					oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
					oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("OrdGbn").Enabled = false;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가
					oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("OrdGbn").Enabled = true;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.EnableMenu("1282", true); //추가

					if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP075L] WHERE SUBSTRING(U_PP070No,1,CHARINDEX('-',U_PP070No) -1) = '" + oDS_PS_PP070H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1)) > 0)
					{
						oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("DocEntry").Enabled = false;
						oForm.Items.Item("BPLId").Enabled = false;
						oForm.Items.Item("OrdGbn").Enabled = false;
						oForm.Items.Item("CntcCode").Enabled = false;
						oForm.Items.Item("DocDate").Enabled = false;
						oForm.Items.Item("Mat01").Enabled = false;
					}
					else
					{
						oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("DocEntry").Enabled = false;
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("OrdGbn").Enabled = true;
						oForm.Items.Item("CntcCode").Enabled = true;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("Mat01").Enabled = true;
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
		/// PS_PP070_ClearForm
		/// </summary>
		private void PS_PP070_ClearForm()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP070'", "");
				if (Convert.ToDouble(DocEntry) == 0)
				{
					oForm.Items.Item("DocEntry").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP070_CheckDataValid
		/// </summary>
		/// <returns></returns>
		private bool PS_PP070_CheckDataValid()
		{
			bool returnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "담당자는 필수입니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "작성일자는 필수입니다.";
					throw new Exception();
				}
				if (oMat.VisualRowCount <= 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}
				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("PP030No").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						errMessage = "작지문서는 필수입니다.";
						throw new Exception();
					}
					if (Convert.ToDouble(oMat.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim()) <= 0)
					{
						errMessage = "선택수량은 필수입니다.";
						throw new Exception();
					}
					if (Convert.ToDouble(oMat.Columns.Item("SelQty").Cells.Item(i).Specific.Value.ToString().Trim()) > Convert.ToDouble(oMat.Columns.Item("CpQty").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						errMessage = "선택수량이 재공수량을 초과합니다.";
						throw new Exception();
					}
				}

				if (PS_PP070_Validate("검사01") == false)
				{
					return returnValue;
				}

				oDS_PS_PP070L.RemoveRecord(oDS_PS_PP070L.Size - 1);
				oMat.LoadFromDataSource();
				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
				{
					PS_PP070_ClearForm();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			return returnValue;
		}

		/// <summary>
		/// PS_PP070_Validate
		/// </summary>
		/// <param name="ValidateType"></param>
		/// <returns></returns>
		private bool PS_PP070_Validate(string ValidateType)
		{
			bool returnValue = false;
			bool Exist;

			int i;
			int j;
			string sQry;
			string errMessage = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (ValidateType == "검사01")
				{
					//입력된 행에 대해
					for (i = 1; i <= oMat.VisualRowCount - 1; i++)
					{
						if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat.Columns.Item("PP030No").Cells.Item(i).Specific.Value.ToString().Trim() + "'", 0, 1)) <= 0)
						{
							errMessage = "작업지시문서가 존재하지 않습니다.";
							throw new Exception();
						}
					}

					//삭제된 행을 찾아서 삭제가능성 검사 , 만약 입력된행이 수정이 불가능하도록 변경이 필요하다면 삭제된행 찾는구문 제거
					sQry = "SELECT DocEntry,LineId,U_ItemCode FROM [@PS_PP070L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);
					for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
					{
						Exist = false;

						for (j = 1; j <= oMat.RowCount - 1; j++)
						{
							if (Convert.ToInt32(oRecordSet.Fields.Item(1).Value.ToString().Trim()) == Convert.ToInt32(oMat.Columns.Item("LineId").Cells.Item(j).Specific.Value.ToString().Trim()) && !string.IsNullOrEmpty(oMat.Columns.Item("LineId").Cells.Item(j).Specific.Value.ToString().Trim()))
							{
								Exist = true;
							}
						}
						//삭제된 행중
						if (Exist == false)
						{
							if (Convert.ToInt32((dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP075L] WHERE U_PP070No = '" + Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + "-" + Convert.ToInt32(oRecordSet.Fields.Item(1).Value.ToString().Trim()) + "'", 0, 1))) > 0)
							{
								errMessage = "삭제된행이 다른사용자에 의해 이동번호등록 되었습니다. 적용할 수 없습니다.";
								throw new Exception();
							}
						}
						oRecordSet.MoveNext();
					}
				}
				else if (ValidateType == "행삭제")
				{
					//행삭제전 행삭제가능여부검사 (추가,수정모드일때행삭제가능검사)
					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
					{
						//새로추가된 행인경우, 삭제하여도 무방하다
						if (!string.IsNullOrEmpty(oMat.Columns.Item("LineId").Cells.Item(oLastColRow01).Specific.Value.ToString().Trim()))
						{
							if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP075L] WHERE U_PP070No = '" + Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim()) + "-" + Convert.ToInt32(oMat.Columns.Item("LineId").Cells.Item(oLastColRow01).Specific.Value.ToString().Trim()) + "'", 0, 1)) > 0)
							{
								errMessage = "이미 이동번호등록된 행입니다. 삭제할 수 없습니다.";
								throw new Exception();
							}
						}
					}
				}
				else if (ValidateType == "취소")
				{
					sQry = "SELECT DocEntry,LineId FROM [@PS_PP070L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
					oRecordSet.DoQuery(sQry);

					for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
					{
						if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP075L] a inner join [@PS_PP075H] b on a.docentry=b.docentry WHERE b.canceled<>'Y' and a.U_PP070No = '" + Convert.ToInt32(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + "-" + Convert.ToInt32(oRecordSet.Fields.Item(1).Value.ToString().Trim()) + "'", 0, 1)) > 0)
						{
							errMessage = "이미 이동번호등록된 행입니다. 삭제할 수 없습니다.";
							throw new Exception();
						}
						oRecordSet.MoveNext();
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
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
            {
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
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP070_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP070_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_PP070_EnableFormItem();
								PS_PP070_AddMatrixRow(0, true); 
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_PP070_EnableFormItem();
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
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string errMessage = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ColUID == "PP030No")
					{
						PS_PP071 oTempClass = new PS_PP071();   //작지조회
						oTempClass.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim(), oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim());
					}
					else
					{
						dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
					}

					if (pVal.ColUID == "PP030No")
					{
						if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택")
						{
							errMessage = "사업장은 필수입니다.";
							throw new Exception();
						}
						else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
						{
							errMessage = "작업구분은 필수입니다.";
							throw new Exception();
						}
						else
						{
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int i;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "PP030No")
							{
								sQry = "EXEC PS_PP070_04 '" + oMat.Columns.Item("PP030No").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
								{
									oDS_PS_PP070L.SetValue("U_PP030No", pVal.Row - 1, oRecordSet.Fields.Item("PP030No").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_OrdNum", pVal.Row - 1, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_OrdSub1", pVal.Row - 1, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_OrdSub2", pVal.Row - 1, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_PP030HNo", pVal.Row - 1, oRecordSet.Fields.Item("PP030HNo").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_PP030MNo", pVal.Row - 1, oRecordSet.Fields.Item("PP030MNo").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_BPLId", pVal.Row - 1, oRecordSet.Fields.Item("BPLId").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_CpCode", pVal.Row - 1, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_CpName", pVal.Row - 1, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());
									oDS_PS_PP070L.SetValue("U_CpQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item("CpQty").Value.ToString().Trim())));
									oDS_PS_PP070L.SetValue("U_CpWt", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item("CpWt").Value.ToString().Trim())));
									oDS_PS_PP070L.SetValue("U_SelQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item("SelQty").Value.ToString().Trim())));
									oDS_PS_PP070L.SetValue("U_SelWt", pVal.Row - 1, "0");
									oDS_PS_PP070L.SetValue("U_LineId", pVal.Row - 1, oRecordSet.Fields.Item("LineId").Value.ToString().Trim());
									oRecordSet.MoveNext();
								}
								if (oMat.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_PP070L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_PP070_AddMatrixRow(pVal.Row, false);
								}
								oMat.LoadFromDataSource();
								oMat.AutoResizeColumns();
								oForm.Update();
							}
							else if (pVal.ColUID == "SelQty")
							{
								if (Convert.ToDouble(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) <= 0)
								{
									PSH_Globals.SBO_Application.MessageBox("수량이 0보다 작습니다.");
									oDS_PS_PP070L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP070L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
								}
								else if (Convert.ToDouble(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) > Convert.ToDouble(oMat.Columns.Item("CpQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.MessageBox("재공수량을 초과합니다.");
									oDS_PS_PP070L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP070L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
								}
								else
								{
									oDS_PS_PP070L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim())));
								}
							}
							else
							{
								oDS_PS_PP070L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							}
						}
						else
						{
							if (pVal.ItemUID == "DocEntry")
							{
								oDS_PS_PP070H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
							else if (pVal.ItemUID == "CntcCode")
							{
								oDS_PS_PP070H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
								oDS_PS_PP070H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", 0, 1));
							}
							else
							{
								oDS_PS_PP070H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
						}
						oMat.LoadFromDataSource();
						oMat.AutoResizeColumns();
						oForm.Update();
						if (pVal.ItemUID == "Mat01")
						{
							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else
						{
							oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Raise_EVENT_MATRIX_LOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					PS_PP070_EnableFormItem();
					PS_PP070_AddMatrixRow(oMat.VisualRowCount, false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP070H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP070L);
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
			try
			{
				int i = 0;
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
						if (PS_PP070_Validate("행삭제") == false)
						{
							BubbleEvent = false;
							return;
						}
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_PP070L.RemoveRecord(oDS_PS_PP070L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_PP070_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_PP070L.GetValue("U_PP030No", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_PP070_AddMatrixRow(oMat.RowCount,false);
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
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
							{
								if (PS_PP070_Validate("취소") == false)
								{
									BubbleEvent = false;
									return;
								}
								if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
								{
									BubbleEvent = false;
									return;
								}
							}
							else
							{
								PSH_Globals.SBO_Application.MessageBox("현재 모드에서는 취소할수 없습니다.");
								BubbleEvent = false;
								return;
							}
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							PS_PP070_EnableFormItem();
							oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_PP070_EnableFormItem();
							PS_PP070_AddMatrixRow(0, true);
							break;
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							PS_PP070_EnableFormItem();
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
