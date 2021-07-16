using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 포장사업팀 제품검사서 등록
	/// </summary>
	internal class PS_QM066 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_QM066H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM066L; //등록라인

		private string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM066.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM066_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM066");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry"; //UDO방식일때

				oForm.Freeze(true);

				PS_QM066_CreateItems();
				PS_QM066_ComboBox_Setting();
				PS_QM066_Initial_Setting();
				PS_QM066_EnableMenus();
				PS_QM066_SetDocument(oFormDocEntry);

				oForm.EnableMenu(("1283"), true); // 삭제
				oForm.EnableMenu(("1287"), true); // 복제
				oForm.EnableMenu(("1286"), true); // 닫기
				oForm.EnableMenu(("1284"), true); // 취소
				oForm.EnableMenu(("1293"), true); // 행삭제
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
		/// PS_QM066_CreateItems
		/// </summary>
		private void PS_QM066_CreateItems()
		{
			try
			{
				oDS_PS_QM066H = oForm.DataSources.DBDataSources.Item("@PS_QM066H");
				oDS_PS_QM066L = oForm.DataSources.DBDataSources.Item("@PS_QM066L");
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
		/// PS_QM066_ComboBox_Setting
		/// </summary>
		private void PS_QM066_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_QM066_Initial_Setting
		/// </summary>
		private void PS_QM066_Initial_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				//일자
				oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_QM066_EnableMenus
		/// </summary>
		private void PS_QM066_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, false, false, false, false, false, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_QM066_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_QM066_SetDocument(string oFormDocEntry)
		{
			try
			{
				if ((string.IsNullOrEmpty(oFormDocEntry)))
				{
					PS_QM066_FormItemEnabled();
					PS_QM066_AddMatrixRow(0, true); //UDO방식일때
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_QM066_FormItemEnabled();
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
		/// PS_QM066_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM066_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);

				if (RowIserted == false)
				{
					oDS_PS_QM066L.InsertRecord((oRow));
				}
				oMat.AddRow();
				oDS_PS_QM066L.Offset = oRow;
				oDS_PS_QM066L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_QM066_FormItemEnabled
		/// </summary>
		private void PS_QM066_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);

				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					PS_QM066_FormClear(); //UDO방식
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
				}
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
				{
					oForm.Items.Item("DocEntry").Specific.Value = "";
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가
				}
				else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE))
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
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
		/// PS_QM066_FormClear
		/// </summary>
		private void PS_QM066_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM066'", "");

				if (string.IsNullOrEmpty(DocEntry) | DocEntry == "0")
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
		/// PS_QM066_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_QM066_DataValidCheck()
		{
			bool functionReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
				{
					PS_QM066_FormClear();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "일자가 입력되지 않았습니다.";
					throw new Exception();
				}

				if (oMat.VisualRowCount <= 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				if (oForm.Items.Item("YQty").Specific.Value.ToString().Trim() != oMat.Columns.Item("YQty").Cells.Item(oMat.VisualRowCount - 1).Specific.Value.ToString().Trim())
				{
					errMessage = "합격수량이 일치하지 않습니다.";
					throw new Exception();
				}

				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if ((string.IsNullOrEmpty(oMat.Columns.Item("InspItNm").Cells.Item(i).Specific.Value.ToString().Trim())))
					{
						errMessage = "검사항목명은 필수입니다.";
						throw new Exception();
					}
				}

				oMat.FlushToDataSource();
				oDS_PS_QM066L.RemoveRecord(oDS_PS_QM066L.Size - 1);
				oMat.LoadFromDataSource();

				if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
				{
					PS_QM066_FormClear();
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
		/// LoadData
		/// </summary>
		/// <param name="oCode"></param>
		private void LoadData(string oCode)
		{
			int i;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "Select b.U_Page, b.U_InspItem, b.U_InspItNm, b.U_InspSpec";
				sQry += " From [@PS_QM065H] a INNER JOIN [@PS_QM065L] b ON a.Code = b.Code ";
				sQry += "Where a.Code = '" + oCode + "' Order By b.U_Seqno ";
				oRecordSet.DoQuery(sQry);

				oDS_PS_QM066L.Clear();
				oMat.Clear();
				oMat.FlushToDataSource();

				i = 0;
				while (!(oRecordSet.EoF))
				{
					oDS_PS_QM066L.InsertRecord(i);
					oDS_PS_QM066L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM066L.SetValue("U_Page", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_QM066L.SetValue("U_InspItem", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_QM066L.SetValue("U_InspItNm", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_QM066L.SetValue("U_InspSpec", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					i += 1;
					oRecordSet.MoveNext();
				}
				oMat.LoadFromDataSource();
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
		/// PS_QM066_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_QM066_Print_Report01()
		{
			string sQry;
			string WinTitle = string.Empty;
			string ReportName = string.Empty;
			string TAQL1;
			string TAQL2;
			string CardCode;
			string ItmMsort;
			string DocEntry;

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItmMsort = codeHelpClass.Left(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim(), 5);

				sQry = "Select Isnull(U_TAQL1,''), Isnull(U_TAQL2,'') From [@PS_QM066H] Where DocEntry = '" + DocEntry + "'";
				oRecordSet.DoQuery(sQry);

				TAQL1 = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				TAQL2 = oRecordSet.Fields.Item(1).Value.ToString().Trim();

				if (CardCode == "12575")
				{
					if (ItmMsort == "10807")
					{
						WinTitle = "[PS_QM066_10] 검사기록서출력_동래철선묶음상자";
						ReportName = "PS_QM066_10.rpt";
					}
					else if (ItmMsort == "10802")
					{
						WinTitle = "[PS_QM066_11] 검사기록서출력_동래파래트";
						ReportName = "PS_QM066_11.rpt";
					}
					else if (ItmMsort == "10801")
					{
						WinTitle = "[PS_QM066_12] 검사기록서출력_동래목상자";
						ReportName = "PS_QM066_12.rpt";
					}
				}
				else
				{
					if (string.IsNullOrEmpty(TAQL1) && string.IsNullOrEmpty(TAQL2))
					{
						WinTitle = "[PS_QM066_01] 검사기록서 출력";
						ReportName = "PS_QM066_01.rpt";
					}
					else
					{
						WinTitle = "[PS_QM066_02] 검사기록서 출력";
						ReportName = "PS_QM066_02.rpt";
					}
				}

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
							if (PS_QM066_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM066_DataValidCheck() == false)
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
					if (pVal.ItemUID == "Button01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM066_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
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
								PS_QM066_FormItemEnabled();
								PS_QM066_AddMatrixRow(0, true); //UDO방식일때
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_QM066_FormItemEnabled();
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
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "CardCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "ODLNNum")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ODLNNum").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem(("7425"));
								BubbleEvent = false;
							}
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
			string ODLNNum;
			int i;
			int YQty;
			string ItmMsort;
			string OutSize;

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "FailQty")
							{
								oMat.FlushToDataSource();
								oDS_PS_QM066L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								YQty = Convert.ToInt32(oMat.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) - Convert.ToInt32(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_QM066L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(YQty)); //합격수량

								for (i = pVal.Row; i <= oMat.VisualRowCount - 1; i++)
								{
									if (!string.IsNullOrEmpty(oDS_PS_QM066L.GetValue("U_InspItem", i).ToString().Trim()))
									{
										oDS_PS_QM066L.SetValue("U_PQty", i, Convert.ToString(YQty)); //검사수량
										oDS_PS_QM066L.SetValue("U_YQty", i, Convert.ToString(YQty)); //합격수량
									}
								}
								oMat.LoadFromDataSource();
							}
							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oMat.AutoResizeColumns();
						}
						else if ((pVal.ItemUID == "CardCode"))
						{
							oDS_PS_QM066H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("cardname", "cardcode", "[ocrd]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
						}
						else if ((pVal.ItemUID == "ODLNNum"))
						{
							ODLNNum = oForm.Items.Item("ODLNNum").Specific.Value.ToString().Trim();

							sQry = " select a.ItemCode, a.Dscription, b.U_Spec4, b.U_ItmMsort from [DLN1] a Inner join OITM b On a.ItemCode = b.ItemCode ";
							sQry += " where a.DocEntry = left('" + ODLNNum + "', charindex('-', '" + ODLNNum + "') -1) ";
							sQry += " and a.LineNum  = Substring('" + ODLNNum + "', charindex('-', '" + ODLNNum + "') + 1, 10) ";
							oRecordSet.DoQuery(sQry);

							oDS_PS_QM066H.SetValue("U_ItemCode", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_QM066H.SetValue("U_ItemName", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_QM066H.SetValue("U_OutSize", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());

							OutSize = oRecordSet.Fields.Item(2).Value.ToString().Trim();
							ItmMsort = oRecordSet.Fields.Item(3).Value.ToString().Trim();

							sQry = " select DocEntry, U_SizeNo, U_Sample1, U_Sample2, U_TAQL1, U_TAQL2 from [@PS_QM065H] ";
							sQry += " where U_ItmMsort = '" + ItmMsort + "' and U_OutSize = '" + OutSize + "'";
							oRecordSet.DoQuery(sQry);

							if (oRecordSet.RecordCount >= 1)
							{
								oDS_PS_QM066H.SetValue("U_SizeNo", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oDS_PS_QM066H.SetValue("U_Sample1", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());
								oDS_PS_QM066H.SetValue("U_Sample2", 0, oRecordSet.Fields.Item(3).Value.ToString().Trim());
								oDS_PS_QM066H.SetValue("U_TAQL1", 0, oRecordSet.Fields.Item(4).Value.ToString().Trim());
								oDS_PS_QM066H.SetValue("U_TAQL2", 0, oRecordSet.Fields.Item(5).Value.ToString().Trim());
								LoadData(oRecordSet.Fields.Item(0).Value.ToString().Trim());
							}
							else
							{
								sQry = " select DocEntry, U_SizeNo, U_Sample1, U_Sample2, U_TAQL1, U_TAQL2 from [@PS_QM065H] ";
								sQry += " where U_ItmMsort = '" + ItmMsort + "'";
								oRecordSet.DoQuery(sQry);

								if (oRecordSet.RecordCount >= 1)
								{
									oDS_PS_QM066H.SetValue("U_SizeNo", 0, oRecordSet.Fields.Item(1).Value.ToString().Trim());
									oDS_PS_QM066H.SetValue("U_Sample1", 0, oRecordSet.Fields.Item(2).Value.ToString().Trim());
									oDS_PS_QM066H.SetValue("U_Sample2", 0, oRecordSet.Fields.Item(3).Value.ToString().Trim());
									oDS_PS_QM066H.SetValue("U_TAQL1", 0, oRecordSet.Fields.Item(4).Value.ToString().Trim());
									oDS_PS_QM066H.SetValue("U_TAQL2", 0, oRecordSet.Fields.Item(5).Value.ToString().Trim());
									LoadData(oRecordSet.Fields.Item(0).Value.ToString().Trim());
								}
							}
						}
						else if ((pVal.ItemUID == "PQty"))
						{
							oMat.FlushToDataSource();

							for (i = 0; i <= oMat.VisualRowCount - 1; i++)
							{
								if (!string.IsNullOrEmpty(oDS_PS_QM066L.GetValue("U_InspItNm", i)))
								{
									oDS_PS_QM066L.SetValue("U_Sample", i, oForm.Items.Item("PQty").Specific.Value.ToString().Trim()); //시료수
									oDS_PS_QM066L.SetValue("U_PQty", i, oForm.Items.Item("PQty").Specific.Value.ToString().Trim());   //검사수량
									oDS_PS_QM066L.SetValue("U_YQty", i, oForm.Items.Item("PQty").Specific.Value.ToString().Trim());   //합격수량
								}
							}
							oMat.LoadFromDataSource();
						}
						else if ((pVal.ItemUID == "ReqNo"))
						{
							oForm.Items.Item("LotNo").Specific.Value = oForm.Items.Item("ReqNo").Specific.Value.ToString().Trim();
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
					PS_QM066_FormItemEnabled();
					PS_QM066_AddMatrixRow(oMat.VisualRowCount, false); //UDO방식
					oMat.AutoResizeColumns();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM066H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM066L);
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if ((oLastColRow01 > 0))
				{
					if (pval.BeforeAction == true)
					{
						//행삭제전 행삭제가능여부검사
					}
					else if (pval.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_QM066L.RemoveRecord(oDS_PS_QM066L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_QM066_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_QM066L.GetValue("U_CntcCode", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_QM066_AddMatrixRow(oMat.RowCount, false);
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
						case "1285": //복원
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
						case "1281": //찾기
							PS_QM066_FormItemEnabled();
							break;
						case "1282": //추가
							PS_QM066_Initial_Setting();
							PS_QM066_FormItemEnabled(); //UDO방식
							PS_QM066_AddMatrixRow(0, true); //UDO방식
							break;
						case "1287": //복제
							oDS_PS_QM066H.SetValue("DocEntry", 0, "");
							for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
							{
								oMat.FlushToDataSource();
								oDS_PS_QM066H.SetValue("DocEntry", i, "");
								oMat.LoadFromDataSource();
							}
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM066_FormItemEnabled();
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
				if (BusinessObjectInfo.BeforeAction == true)
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
				else if (BusinessObjectInfo.BeforeAction == false)
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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}
	}
}
