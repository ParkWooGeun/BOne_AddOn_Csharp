using System;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 표준등록-SUB
	/// </summary>
	internal class PS_QM205 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_QM205H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM205L; //사업부실적

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

		/// <summary>
		/// Form 호출 PS_QM200만 호출
		/// </summary>
		/// <param name="prmPS_QM200"></param>
		/// <param name="prmRow"></param>
		/// <param name="prmLineID"></param>
		public void LoadForm(PS_QM200 prmPS_QM200, int prmRow, int prmLineID)
		{
			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM205.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM205_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM205");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_QM205_CreateItems();
				PS_QM205_ComboBox_Setting();
				PS_QM205_EnableMenus();
				PS_QM205_SetDocument(prmPS_QM200, prmRow, prmLineID);

				oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1286", true);	 //닫기
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1284", true);	 //취소
				oForm.EnableMenu("1293", true);	 //행삭제
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
		/// PS_QM205_CreateItems
		/// </summary>
		private void PS_QM205_CreateItems()
		{
			try
			{
				oDS_PS_QM205H = oForm.DataSources.DBDataSources.Item("@PS_QM205H");
				oDS_PS_QM205L = oForm.DataSources.DBDataSources.Item("@PS_QM205L");
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
		/// PS_QM205_ComboBox_Setting
		/// </summary>
		private void PS_QM205_ComboBox_Setting()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//문서타입
				sQry = "     SELECT      T1.U_Minor,";
				sQry += "                 T1.U_CdName";
				sQry += "  FROM       [@PS_SY001H] AS T0";
				sQry += "                 INNER JOIN";
				sQry += "                 [@PS_SY001L] AS T1";
				sQry += "                     ON T0.Code = T1.Code";
				sQry += "  WHERE      T0.Code = 'Q200'";
				oForm.Items.Item("DocType").Specific.ValidValues.Add("", "선택");
				dataHelpClass.Set_ComboList(oForm.Items.Item("DocType").Specific, sQry, "", false, false);
				oForm.Items.Item("DocType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM205_EnableMenus
		/// </summary>
		private void PS_QM205_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM205_SetDocument
		/// </summary>
		/// <param name="prmPS_QM200"></param>
		/// <param name="prmRow"></param>
		/// <param name="prmLineID"></param>
		private void PS_QM205_SetDocument(PS_QM200 prmPS_QM200, int prmRow, int prmLineID)
		{
			string DocType;
			string DocCode;
			int BasEntry;
			int BasLine;
			string DeptCode;
			string DeptName;
			string StdName;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (prmPS_QM200 == null)
				{
					PS_QM205_FormItemEnabled();
					PS_QM205_AddMatrixRow01(0, true);
				}
				else
				{
					DocType = prmPS_QM200.oForm.Items.Item("DocType").Specific.Value.ToString().Trim();
					DocCode = prmPS_QM200.oMat.Columns.Item("DocCode").Cells.Item(prmRow).Specific.Value.ToString().Trim();
					BasEntry = Convert.ToInt32(prmPS_QM200.oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim());
					BasLine = Convert.ToInt32(prmLineID);
					DeptCode = prmPS_QM200.oMat.Columns.Item("DeptCode").Cells.Item(prmRow).Specific.Value.ToString().Trim();
					DeptName = prmPS_QM200.oMat.Columns.Item("DeptName").Cells.Item(prmRow).Specific.Value.ToString().Trim();
					StdName = prmPS_QM200.oMat.Columns.Item("StdName").Cells.Item(prmRow).Specific.Value.ToString().Trim();

					//기존문서가 존재하는지 검사
					sQry = " SELECT      COUNT(*) AS [Cnt]";
					sQry += " FROM        [@PS_QM205H] AS T0";
					sQry += "                 INNER JOIN";
					sQry += "                 [@PS_QM205L] AS T1";
					sQry += "                     ON T0.DocEntry = T1.DocEntry";
					sQry += " WHERE       T0.U_DocType = '" + DocType + "'";
					sQry += "                 AND T0.U_BasEntry = " + BasEntry;
					sQry += "                 AND T0.U_BasLine = " + BasLine;
					oRecordSet.DoQuery(sQry);

					if (Convert.ToDouble(oRecordSet.Fields.Item("Cnt").Value.ToString().Trim()) > 0)
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
						oForm.Items.Item("DocType").Specific.Select(DocType, SAPbouiCOM.BoSearchKey.psk_ByValue);
						oForm.Items.Item("BasEntry").Specific.Value = BasEntry;
						oForm.Items.Item("BasLine").Specific.Value = BasLine;

						PS_QM205_FormItemEnabled();
						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					}
					else
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
						PS_QM205_FormItemEnabled();
						PS_QM205_AddMatrixRow01(0, true);

						oForm.Items.Item("DocType").Specific.Select(DocType, SAPbouiCOM.BoSearchKey.psk_ByValue);
						oForm.Items.Item("DocCode").Specific.Value = DocCode;
						oForm.Items.Item("BasEntry").Specific.Value = BasEntry;
						oForm.Items.Item("BasLine").Specific.Value = BasLine;
						oForm.Items.Item("DeptCode").Specific.Value = DeptCode;
						oForm.Items.Item("DeptName").Specific.Value = DeptName;
						oForm.Items.Item("StdName").Specific.Value = StdName;
					}
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
		}

		/// <summary>
		/// PS_QM205_FormItemEnabled
		/// </summary>
		private void PS_QM205_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocType").Enabled = false;
					oForm.Items.Item("DocCode").Enabled = true;
					oForm.Items.Item("DeptCode").Enabled = true;
					oForm.Items.Item("DeptName").Enabled = true;
					oForm.Items.Item("BasEntry").Enabled = false;
					oForm.Items.Item("BasLine").Enabled = false;
					oForm.Items.Item("StdName").Enabled = true;
					oForm.Items.Item("DocEntry").Enabled = false;
					PS_QM205_FormClear();
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocType").Enabled = true;
					oForm.Items.Item("DocCode").Enabled = true;
					oForm.Items.Item("DeptCode").Enabled = true;
					oForm.Items.Item("DeptName").Enabled = true;
					oForm.Items.Item("BasEntry").Enabled = true;
					oForm.Items.Item("BasLine").Enabled = true;
					oForm.Items.Item("StdName").Enabled = false;
					oForm.Items.Item("DocEntry").Specific.Value = "";
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocType").Enabled = false;
					oForm.Items.Item("DocCode").Enabled = true;
					oForm.Items.Item("DeptCode").Enabled = true;
					oForm.Items.Item("DeptName").Enabled = true;
					oForm.Items.Item("BasEntry").Enabled = false;
					oForm.Items.Item("BasLine").Enabled = false;
					oForm.Items.Item("StdName").Enabled = true;
					oForm.Items.Item("DocEntry").Enabled = false;
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
		/// PS_QM205_AddMatrixRow01
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM205_AddMatrixRow01(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_QM205L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM205L.Offset = oRow;
				oDS_PS_QM205L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_QM205_FormClear
		/// </summary>
		private void PS_QM205_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM205'", "");
				if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
				{
					oForm.Items.Item("DocEntry").Specific.Value = "1";
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
		/// PS_QM205_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_QM205_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM205_FormClear();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DocType").Specific.Selected.Value.ToString().Trim()))
				{
					errMessage = "문서타입을 입력하지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DocCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "문서코드를 입력하지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DeptCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "부서코드를 입력하지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DeptName").Specific.Value.ToString().Trim()))
				{
					errMessage = "부서명을 입력하지 않았습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("StdName").Specific.Value.ToString().Trim()))
				{
					errMessage = "표준명을 입력하지 않았습니다.";
					throw new Exception();
				}
				if (oMat.VisualRowCount == 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("CrtDate").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("CrtDate").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "개정일자를 입력하지 않았습니다. 확인하십시오.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oMat.Columns.Item("EmpCode").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("EmpCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "기안자사번을 입력하지 않았습니다. 확인하십시오.";
						throw new Exception();
					}
					if (string.IsNullOrEmpty(oMat.Columns.Item("CrtCmt").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("CrtCmt").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "개정내용을 입력하지 않았습니다. 확인하십시오.";
						throw new Exception();
					}
				}

				oDS_PS_QM205L.RemoveRecord(oDS_PS_QM205L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM205_FormClear();
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
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				case SAPbouiCOM.BoEventTypes.et_CLICK: //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8	
				//	Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
					Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
					break;
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
							if (PS_QM205_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM205_DataValidCheck() == false)
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
								PS_QM205_FormItemEnabled();
								PS_QM205_AddMatrixRow01(0, true);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_QM205_FormItemEnabled();
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
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "Mat01")
						{
							dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "EmpCode");
						}
						else
						{
							dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "DeptCode", "");
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
		/// VALIDATE 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string Code;
			string Name;
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
							oMat.FlushToDataSource();

							if (pVal.ColUID == "CrtDate")
							{
								if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_QM205L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_QM205_AddMatrixRow01(pVal.Row, false);
								}
							}
							else if (pVal.ColUID == "EmpCode")
							{
								Code = oDS_PS_QM205L.GetValue("U_EmpCode", pVal.Row - 1).ToString().Trim();

								sQry = "       SELECT      U_FullName AS [Name]";
								sQry += "  FROM        [OHEM]";
								sQry += "  WHERE      U_MSTCOD = '" + Code + "'";
								oRecordSet.DoQuery(sQry);

								Name = oRecordSet.Fields.Item("Name").Value.ToString().Trim();
								oDS_PS_QM205L.SetValue("U_EmpName", pVal.Row - 1, Name);
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
						else
						{
							if (pVal.ItemUID == "DeptCode")
							{
								oDS_PS_QM205H.SetValue("U_DeptName", 0, dataHelpClass.Get_ReData("U_CodeNm", "U_Code", "[@PS_HR200L]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", "AND Code = '1'"));
							}
						}

						oForm.Update();
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
					if (pVal.ItemUID == "Mat01")
					{
						PS_QM205_FormItemEnabled();
						PS_QM205_AddMatrixRow01(oMat.VisualRowCount, false);
						oMat.AutoResizeColumns();
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM205H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM205L);
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
						oDS_PS_QM205L.RemoveRecord(oDS_PS_QM205L.Size - 1);
						oMat.LoadFromDataSource();

						if (oMat.RowCount == 0)
						{
							PS_QM205_AddMatrixRow01(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_QM205L.GetValue("U_CrtDate", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_QM205_AddMatrixRow01(oMat.RowCount, false);
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
						case "1283": //삭제
							break;
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
							PS_QM205_FormItemEnabled();
							oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_QM205_FormItemEnabled();
							PS_QM205_AddMatrixRow01(0, true);
							break;
						case "1287": //복제 
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM205_FormItemEnabled();
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
