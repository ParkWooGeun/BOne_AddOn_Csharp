using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 설비정비일지등록
	/// </summary>
	internal class PS_PP285 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.DBDataSource oDS_PS_PP285H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP285L; //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP285M; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oMat01Row01;
		private int oMat02Row02;
		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP285.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP285_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP285");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);

				PS_PP285_CreateItems();
				PS_PP285_SetComboBox();
				PS_PP285_EnableMenus();
				PS_PP285_SetDocument(oFormDocEntry);
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
		/// PS_PP285_CreateItems
		/// </summary>
		private void PS_PP285_CreateItems()
		{
			try
			{
				oDS_PS_PP285H = oForm.DataSources.DBDataSources.Item("@PS_PP285H");
				oDS_PS_PP285L = oForm.DataSources.DBDataSources.Item("@PS_PP285L");
				oDS_PS_PP285M = oForm.DataSources.DBDataSources.Item("@PS_PP285M");

				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oMat02 = oForm.Items.Item("Mat02").Specific;
				oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat02.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP285_SetComboBox
		/// </summary>
		private void PS_PP285_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
				// 반
				dataHelpClass.Combo_ValidValues_Insert("PS_PP285", "ClsCode", "", "10", "동력");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP285", "ClsCode", "", "20", "공무");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP285", "ClsCode", "", "90", "공통");
				dataHelpClass.Combo_ValidValues_SetValueItem((oForm.Items.Item("ClsCode").Specific), "PS_PP285", "ClsCode", false);

				dataHelpClass.Combo_ValidValues_Insert("PS_PP285", "Mat01", "WorkDiv", "선택", "선택");
				dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkDiv"), "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE CODE = 'P220' order by U_Seq", "", "");

				dataHelpClass.Combo_ValidValues_Insert("PS_PP285", "Mat01", "EndYN", "0", "계속");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP285", "Mat01", "EndYN", "1", "완료");
				dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("EndYN"), "PS_PP285", "Mat01", "EndYN", false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// EnableMenus
		/// </summary>
		private void PS_PP285_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, true, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false); //메뉴설정
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP285_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_PP285_SetDocument(string oFormDocEntry)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_PP285_EnableFormItem();
					PS_PP285_AddMatrixRow01(0, true);
					PS_PP285_AddMatrixRow02(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_PP285_EnableFormItem();
					oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP285_AddMatrixRow01
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP285_AddMatrixRow01(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				
				if (RowIserted == false) //행추가여부
				{
					oDS_PS_PP285L.InsertRecord(oRow);
				}
				oMat01.AddRow();
				oDS_PS_PP285L.Offset = oRow;
				oDS_PS_PP285L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_PP285_AddMatrixRow02
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP285_AddMatrixRow02(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				
				if (RowIserted == false) //행추가여부
				{
					oDS_PS_PP285M.InsertRecord(oRow);
				}
				oMat02.AddRow();
				oDS_PS_PP285M.Offset = oRow;
				oDS_PS_PP285M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
				oMat02.LoadFromDataSource();
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
		/// PS_PP285_CheckDataValid
		/// </summary>
		/// <returns></returns>
		private bool PS_PP285_CheckDataValid()
		{
			bool returnValue = false;
			string errMessage = string.Empty;
			int i;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_PP285_ClearForm();
				}
				// 작업반
				if (string.IsNullOrEmpty(oForm.Items.Item("ClsCode").Specific.Value.ToString().Trim()))
				{
					PSH_Globals.SBO_Application.MessageBox("작업반을 입력 하세요.");
					oForm.Items.Item("ClsCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "1";
					throw new Exception();
				}

				for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oMat01.Columns.Item("Code1").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						PSH_Globals.SBO_Application.MessageBox("설비코드(대)는 필수입니다.");
						oMat01.Columns.Item("Code1").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "2";
						throw new Exception();
					}

					if (string.IsNullOrEmpty(oMat01.Columns.Item("WorkDiv").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						PSH_Globals.SBO_Application.MessageBox("작업구분은 필수입니다.");
						oMat01.Columns.Item("WorkDiv").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "3";
						throw new Exception();
					}

					if (string.IsNullOrEmpty(oMat01.Columns.Item("EndYN").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						PSH_Globals.SBO_Application.MessageBox("완료/계속은 필수입니다.");
						oMat01.Columns.Item("EndYN").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "4";
						throw new Exception();
					}

					if (oMat01.Columns.Item("WorkDiv").Cells.Item(i).Specific.Value.ToString().Trim() == "50" 
						&& string.IsNullOrEmpty(oMat01.Columns.Item("PP284LN").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						PSH_Globals.SBO_Application.MessageBox("계획예방작업선택시 계획연결을 꼭 등록하여야 합니다.");
						oMat01.Columns.Item("PP284LN").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "5";
						throw new Exception();
					}
				}

				oDS_PS_PP285L.RemoveRecord(oDS_PS_PP285L.Size - 1);
				oMat01.LoadFromDataSource();
				oDS_PS_PP285M.RemoveRecord(oDS_PS_PP285M.Size - 1);
				oMat02.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_PP285_ClearForm();
				}

				returnValue = true;
			}
			catch (Exception ex)
			{
				if (errMessage != string.Empty)
				{
				//	PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			return returnValue;
		}

		/// <summary>
		/// PS_PP285_ClearForm
		/// </summary>
		private void PS_PP285_ClearForm()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP285'", "");
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
		/// PS_PP285_EnableFormItem
		/// </summary>
		private void PS_PP285_EnableFormItem()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					oForm.Items.Item("Mat02").Enabled = true;
					oForm.Items.Item("1").Enabled = true;
					PS_PP285_ClearForm();
					
					oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); // 사업장
				    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); // 일자
				} 
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);	 //추가
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
					oForm.Items.Item("Mat02").Enabled = true;
					oForm.Items.Item("1").Enabled = true;
				}
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.EnableMenu("1282", true); //추가
					if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP285H] WHERE DocEntry = '" + oDS_PS_PP285H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "Y")
					{
						oForm.Items.Item("DocEntry").Enabled = false;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("Mat01").Enabled = true;
						oForm.Items.Item("Mat02").Enabled = true;
						oForm.Items.Item("1").Enabled = true;
					}
					else
					{
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
		/// PS_PP285_Validate
		/// </summary>
		/// <param name="ValidateType"></param>
		/// <returns></returns>
		private bool PS_PP285_Validate(string ValidateType)
		{
			bool returnValue = false;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			string errMessage = string.Empty;

			try
			{
				if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP285H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1) == "Y")
				{
					errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.";
					throw new Exception();
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
				if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02")
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
				if (pVal.ItemUID == "Mat01")
				{
					if (pVal.Row > 0)
					{
						oMat01Row01 = pVal.Row;
					}
				}
				else if (pVal.ItemUID == "Mat02")
				{
					if (pVal.Row > 0)
					{
						oMat02Row02 = pVal.Row;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
							if (PS_PP285_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP285_CheckDataValid() == false)
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
					//취소버튼 누를시 저장할 자료가 있으면 메시지 표시
					if (pVal.ItemUID == "2")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (oMat01.VisualRowCount > 1)
							{
								if (PSH_Globals.SBO_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 1, "예", "아니오") != 1)
								{
									BubbleEvent = false;
									return;
								}
							}
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
								PS_PP285_EnableFormItem();
								PS_PP285_AddMatrixRow01(0, true);
								PS_PP285_AddMatrixRow02(0, true);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
								{
									oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
									oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
									PS_PP285_EnableFormItem();
									oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
									oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}
								PS_PP285_EnableFormItem();
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
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.ColUID == "RspCode")
						{
							if (string.IsNullOrEmpty(oMat01.Columns.Item("RspCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						if (pVal.ColUID == "Code1")
						{
							if (string.IsNullOrEmpty(oMat01.Columns.Item("Code1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						if (pVal.ColUID == "PP284LN")
						{
							if (string.IsNullOrEmpty(oMat01.Columns.Item("PP284LN").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) 
								&& oMat01.Columns.Item("WorkDiv").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() == "50")
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
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
							oMat01Row01 = pVal.Row;
						}
					}
					if (pVal.ItemUID == "Mat02")
					{
						if (pVal.Row > 0)
						{
							oMat02.SelectRow(pVal.Row, true, false);
							oMat02Row02 = pVal.Row;
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
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
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
							oMat01.FlushToDataSource();

							if (pVal.ColUID == "RspCode") //담당명
							{
								sQry = "  SELECT      U_CdName ";
								sQry += " FROM        [@PS_SY001L]";
								sQry += " WHERE       Code = 'P003'";
								sQry += "   AND       U_Minor = '" + oMat01.Columns.Item("RspCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								oDS_PS_PP285L.SetValue("U_RspName", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
								
								oDS_PS_PP285L.SetValue("U_EndYN", pVal.Row - 1, "1"); //완료/계속 필드에 기본으로 "1"완료로 SET
								oDS_PS_PP285L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP285L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_PP285_AddMatrixRow01(pVal.Row, false);
								}
							}
							else if (pVal.ColUID == "Code1")  // 설비코드(대분류)조회
							{
								sQry = "  Select t1.U_Name1 ";
								sQry += " From [@PS_PP280H] t0 INNER JOIN [@PS_PP280L] t1 ON t0.DocEntry = t1.DocEntry ";
								sQry += " Where t0.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
								sQry += "   and t1.U_Code1 = '" + oMat01.Columns.Item("Code1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								oDS_PS_PP285L.SetValue("U_Name1", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							}
							else if (pVal.ColUID == "Code2") // 설비코드(중분류)조회
							{
								sQry = "  Select t1.U_Name2 ";
								sQry += " From [@PS_PP281H] t0 INNER JOIN [@PS_PP281L] t1 ON t0.DocEntry = t1.DocEntry ";
								sQry += " Where t0.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
								sQry += "   and t0.U_Code1 = '" + oMat01.Columns.Item("Code1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								sQry += "   and t1.U_Code2 = '" + oMat01.Columns.Item("Code2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								oDS_PS_PP285L.SetValue("U_Name2", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							}
							else if (pVal.ColUID == "Code3") // 설비코드(세분류)조회
							{
								sQry = "  Select t1.U_Name3 ";
								sQry += " From [@PS_PP282H] t0 INNER JOIN [@PS_PP282L] t1 ON t0.DocEntry = t1.DocEntry ";
								sQry += " Where t0.U_BPLId = '" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "'";
								sQry += "   and t0.U_Code1 = '" + oMat01.Columns.Item("Code1").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								sQry += "   and t0.U_Code2 = '" + oMat01.Columns.Item("Code2").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								sQry += "   and t1.U_Code3 = '" + oMat01.Columns.Item("Code3").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);
								oDS_PS_PP285L.SetValue("U_Name3", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							}
							oDS_PS_PP285L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
						}
						else if (pVal.ItemUID == "Mat02")
						{
							oMat02.FlushToDataSource();

							if (pVal.ColUID == "CntcCode")
							{
								oDS_PS_PP285M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_PP285M.SetValue("U_CntcName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_FullName FROM [@PH_PY001A] WHERE Code = '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", 0, 1));
								if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP285M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_PP285_AddMatrixRow02(pVal.Row, false);
								}
							}
							else
							{
								oDS_PS_PP285M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							}
						}
						else
						{
							if ((pVal.ItemUID == "DocEntry"))
							{
								oDS_PS_PP285H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
							else
							{
								oDS_PS_PP285H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
						}

                        oMat01.LoadFromDataSource();
						oMat02.LoadFromDataSource();
						oMat01.AutoResizeColumns();
						oMat02.AutoResizeColumns();
                        oForm.Update();

                        if (pVal.ItemUID == "Mat01")
						{
							oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (pVal.ItemUID == "Mat02")
						{
							oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
					PS_PP285_EnableFormItem();
					if (pVal.ItemUID == "Mat01")
					{
						PS_PP285_AddMatrixRow01(oMat01.VisualRowCount, false); 
					}
					else if (pVal.ItemUID == "Mat02")
					{
						PS_PP285_AddMatrixRow02(oMat02.VisualRowCount, false); 
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_FORM_RESIZE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					oForm.Items.Item("Mat02").Height = oForm.Height - (oForm.Items.Item("Mat01").Height + 180);
					oMat02.AutoResizeColumns();
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
				if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02")
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
				if (pVal.ItemUID == "Mat01")
				{
					if (pVal.Row > 0)
					{
						oMat01Row01 = pVal.Row;
					}
				}
				else if (pVal.ItemUID == "Mat02")
				{
					if (pVal.Row > 0)
					{
						oMat02Row02 = pVal.Row;
					}
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP285H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP285L);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP285M);
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
			int i;

			try
			{
				if ((oLastColRow01 > 0))
				{
					if (pVal.BeforeAction == true)
					{
					}
					else if (pVal.BeforeAction == false)
					{
						if (oLastItemUID01 == "Mat01")
						{
							for (i = 1; i <= oMat01.VisualRowCount; i++)
							{
								oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
							}
							oMat01.FlushToDataSource();
							oDS_PS_PP285L.RemoveRecord(oDS_PS_PP285L.Size - 1);
							oMat01.LoadFromDataSource();
							if (oMat01.RowCount == 0)
							{
								PS_PP285_AddMatrixRow01(0, false);
							}
							else
							{
								if (!string.IsNullOrEmpty(oDS_PS_PP285L.GetValue("U_Code1", oMat01.RowCount - 1).ToString().Trim()))
								{
									PS_PP285_AddMatrixRow01(oMat01.RowCount, false);
								}
							}
						}
						else if (oLastItemUID01 == "Mat02")
						{
							for (i = 1; i <= oMat02.VisualRowCount; i++)
							{
								oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
							}
							oMat02.FlushToDataSource();
							oDS_PS_PP285M.RemoveRecord(oDS_PS_PP285M.Size - 1);
							oMat02.LoadFromDataSource();
							if (oMat02.RowCount == 0)
							{
								PS_PP285_AddMatrixRow02(0, false);
							}
							else
							{
								if (!string.IsNullOrEmpty(oDS_PS_PP285M.GetValue("U_CntcCode", oMat02.RowCount - 1).ToString().Trim()))
								{
									PS_PP285_AddMatrixRow02(oMat02.RowCount, false);
								}
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
		/// FormMenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if ((pVal.BeforeAction == true))
				{
					switch (pVal.MenuUID)
					{
						case "1284": //취소
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
							{
								if ((PS_PP285_Validate("취소") == false))
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_PP285_EnableFormItem();
							break;

					}
				}
				else if ((pVal.BeforeAction == false))
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
							PS_PP285_EnableFormItem(); 
							oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_PP285_EnableFormItem(); 
							PS_PP285_AddMatrixRow01(0, true); 
							PS_PP285_AddMatrixRow02(0, true); 
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_PP285_EnableFormItem();
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
