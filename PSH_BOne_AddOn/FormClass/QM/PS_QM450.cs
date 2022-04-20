using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 검사여부등록
	/// </summary>
	internal class PS_QM450 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_QM450H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM450L; //등록라인
		
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

		/// <summary>
		/// Form 호출
		/// </summary>
		public override void LoadForm(string oFormDocEntry)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM450.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM450_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM450");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_QM450_CreateItems();
				PS_QM450_EnableMenus();
				PS_QM450_SetDocument(oFormDocEntry);
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
		/// PS_QM450_CreateItems
		/// </summary>
		private void PS_QM450_CreateItems()
		{
			try
			{
				oDS_PS_QM450H = oForm.DataSources.DBDataSources.Item("@PS_QM450H");
				oDS_PS_QM450L = oForm.DataSources.DBDataSources.Item("@PS_QM450L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM450_EnableMenus
		/// </summary>
		private void PS_QM450_EnableMenus()
		{
			try
			{
				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", true);  // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", true);  // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
				oForm.EnableMenu("1281", false); 
				oForm.EnableMenu("1282", true);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM450_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_QM450_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_QM450_FormItemEnabled();
					PS_QM450_AddMatrixRow(0, true);
				}
				else
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM450_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_QM450_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);
				switch (oUID)
				{
					case "Mat01":
						oMat.FlushToDataSource();

						if (oCol == "ItemCode")
						{
							if (oMat.RowCount == oRow && !string.IsNullOrEmpty(oDS_PS_QM450L.GetValue("U_ItemCode", oRow - 1).ToString().Trim()))
							{
								PS_QM450_AddMatrixRow(oRow, false);
							}

							sQry = " EXEC PS_QM450_01 '";
							sQry += oMat.Columns.Item("ItemCode").Cells.Item(oRow).Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oDS_PS_QM450L.SetValue("U_ItemName", oRow - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
							oDS_PS_QM450L.SetValue("U_Spec", oRow - 1, oRecordSet.Fields.Item("ItemSpec").Value.ToString().Trim());
							oDS_PS_QM450L.SetValue("U_OrdQty", oRow - 1, oRecordSet.Fields.Item("OrdQty").Value.ToString().Trim());
							oDS_PS_QM450L.SetValue("U_ChkYQty", oRow - 1, oRecordSet.Fields.Item("ChkYQty").Value.ToString().Trim());
						}
						
						oMat.LoadFromDataSource();
						oMat.AutoResizeColumns();
						break;

					case "ChkCode":
						oDS_PS_QM450H.SetValue("U_ChkName", 0, dataHelpClass.Get_ReData("U_FullName", "Code", "[@PH_PY001A]", "'" + oDS_PS_QM450H.GetValue("U_ChkCode", 0) + "'", ""));
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_QM450_FormItemEnabled
		/// </summary>
		private void PS_QM450_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					PS_QM450_FormClear();
					oForm.EnableMenu("1281", true);	 //찾기
					oForm.EnableMenu("1282", false); //추가
					oDS_PS_QM450H.SetValue("U_ChkDate", 0, DateTime.Now.ToString("yyyyMMdd"));
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocEntry").Specific.Value = "";
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
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

		// <summary>
		/// /PS_QM450_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM450_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_QM450L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM450L.Offset = oRow;
				oDS_PS_QM450L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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

		// <summary>
		/// /PS_QM450_CopyMatrixRow
		/// </summary>
		private void PS_QM450_CopyMatrixRow()
		{
			int i;
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM450'", "");
				oForm.Items.Item("DocEntry").Specific.Value = DocEntry;

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_QM450H.SetValue("DocEntry", i, Convert.ToString(DocEntry));
					oMat.LoadFromDataSource();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM450_FormClear
		/// </summary>
		private void PS_QM450_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM450'", "");

				if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM450_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_QM450_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM450_FormClear();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("ChkCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "검사자정보가 입력되지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ChkDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "검사일자가 입력되지 않았습니다.";
					throw new Exception();
				}
				if (oMat.VisualRowCount == 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "작번은 필수입니다.";
						throw new Exception();
					}
					if (Convert.ToDouble(oMat.Columns.Item("ChkQty").Cells.Item(i).Specific.Value) == 0)
					{
						oMat.Columns.Item("ChkQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "검사수량을 입력하지 않았습니다.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oMat.Columns.Item("ChkNo").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("ChkNo").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "성적서번호는 필수입니다. 확인하십시오.";
						throw new Exception();
					}
					if (PS_QM450_CheckDate(i) == false)
					{
						errMessage = Convert.ToString(i) + "행 작번의 가입고가 등록되지 않았거나, 가입고일자가 더 빠릅니다." + "자재창고에 문의하십시오." + "전체문서가 등록되지 않습니다.";
						throw new Exception();
					}
					if (Convert.ToDouble(oMat.Columns.Item("ChkYQty").Cells.Item(i).Specific.Value.ToString().Trim()) + Convert.ToDouble(oMat.Columns.Item("ChkQty").Cells.Item(i).Specific.Value.ToString().Trim()) > Convert.ToDouble(oMat.Columns.Item("OrdQty").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						errMessage = i + "행 작번은 검사여부등록이 완료되었습니다." + "수주수량과 검사완료수량을 확인하십시오." + "전체문서가 등록되지 않습니다.";
						throw new Exception();
					}
				}

				oDS_PS_QM450L.RemoveRecord(oDS_PS_QM450L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM450_FormClear();
				}

				ReturnValue = true;
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
			return ReturnValue;
		}

		/// <summary>
		/// 선행프로세스와 일자 비교
		/// </summary>
		/// <param name="pRow"></param>
		/// <returns></returns>
		private bool PS_QM450_CheckDate(int pRow)
		{
			bool ReturnValue = false;
			string BaseEntry;
			string BaseLine;
			string DocType;
			string CurDocDate;
			string OrdNum;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				OrdNum = oMat.Columns.Item("ItemCode").Cells.Item(pRow).Specific.Value.ToString().Trim();
				BaseEntry = "";
				BaseLine = "";
				DocType = "PS_QM450" + "-" + OrdNum;
				CurDocDate = oForm.Items.Item("ChkDate").Specific.Value.ToString().Trim();

				sQry = " EXEC PS_Z_CHECK_DATE '";
				sQry += BaseEntry + "','";
				sQry += BaseLine + "','";
				sQry += DocType + "','";
				sQry += CurDocDate + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.Fields.Item("ReturnValue").Value == "False")
				{
					ReturnValue = false;
				}
				else
				{
					ReturnValue = true;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
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
                //case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                //    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                //    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
					Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
                //    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
                //    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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
                //case SAPbouiCOM.BoEventTypes.et_FORM_MENU_HILIGHT: //23
                //    Raise_EVENT_FORM_MENU_HILIGHT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                //    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED: //37
                //    Raise_EVENT_PICKER_CLICKED(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_GRID_SORT: //38
                //    Raise_EVENT_GRID_SORT(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                //case SAPbouiCOM.BoEventTypes.et_Drag: //39
                //    Raise_EVENT_Drag(FormUID, ref pVal, ref BubbleEvent);
                //    break;
                case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_QM450_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM450_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
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
								PS_QM450_FormItemEnabled();
								PS_QM450_AddMatrixRow(0, true);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_QM450_FormItemEnabled();
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.ColUID == "ItemCode")
						{
							dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "ItemCode");
						}
					}
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ChkCode", ""); //검사자
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
							oLastItemUID01 = pVal.ItemUID;
							oLastColUID01 = pVal.ColUID;
							oLastColRow01 = pVal.Row;
							oMat.SelectRow(pVal.Row, true, false);
						}
					}
					else
					{
						oLastItemUID01 = pVal.ItemUID;
						oLastColUID01 = "";
						oLastColRow01 = 0;
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
							oMat.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
							oMat.FlushToDataSource();
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
		/// VALIDATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (pVal.ItemChanged == true)
					{
						PS_QM450_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);

						if (pVal.ItemUID == "Mat01")
						{
							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
						}
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
					PS_QM450_FormItemEnabled();
					PS_QM450_AddMatrixRow(oMat.VisualRowCount, false);
					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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

					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM450H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM450L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_QM450L.RemoveRecord(oDS_PS_QM450L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_QM450_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_QM450L.GetValue("U_ItemCode", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_QM450_AddMatrixRow(oMat.RowCount, false);
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
							PS_QM450_FormItemEnabled();
							oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_QM450_FormItemEnabled();
							PS_QM450_AddMatrixRow(0, true);
							break;
						case "1287": //복제
							PS_QM450_CopyMatrixRow();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM450_FormItemEnabled();
							break;
						case "7169": //엑셀 내보내기
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
