using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 품질관리>(부품) 제품검사서 등록
	/// </summary>
	internal class PS_QM056 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_QM056H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_QM056L; //등록라인
		
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		private string oDocEntry01;
		private SAPbouiCOM.BoFormMode oFormMode01;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_QM056.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_QM056_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_QM056");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_QM056_CreateItems();
				PS_QM056_ComboBox_Setting();
				PS_QM056_Initial_Setting();
				PS_QM056_EnableMenus();
				PS_QM056_SetDocument(oFormDocEntry);

				oForm.EnableMenu("1283", true); // 삭제
				oForm.EnableMenu("1287", true); // 복제
				oForm.EnableMenu("1286", true); // 닫기
				oForm.EnableMenu("1284", true); // 취소
				oForm.EnableMenu("1293", true); // 행삭제
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
		/// PS_QM056_CreateItems
		/// </summary>
		private void PS_QM056_CreateItems()
		{
			try
			{
				oDS_PS_QM056H = oForm.DataSources.DBDataSources.Item("@PS_QM056H");
				oDS_PS_QM056L = oForm.DataSources.DBDataSources.Item("@PS_QM056L");
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
		/// PS_QM056_ComboBox_Setting
		/// </summary>
		private void PS_QM056_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false); //사업장
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM056_Initial_Setting
		/// </summary>
		private void PS_QM056_Initial_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				oForm.Items.Item("CardCode").Specific.Value = "12521";
				oForm.Items.Item("AQL1").Specific.Value = "100";
				oForm.Items.Item("AQL2").Specific.Value = "100";
				oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM056_EnableMenus
		/// </summary>
		private void PS_QM056_EnableMenus()
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
		/// PS_QM056_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_QM056_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_QM056_FormItemEnabled();
					PS_QM056_AddMatrixRow(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_QM056_FormItemEnabled();
					oForm.Items.Item("DocEntry").Specific.Value = oFromDocEntry01;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_QM056_FormItemEnabled
		/// </summary>
		private void PS_QM056_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("Mat01").Enabled = true;
					PS_QM056_FormClear();
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
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

		/// <summary>
		/// PS_QM056_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_QM056_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_QM056L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_QM056L.Offset = oRow;
				oDS_PS_QM056L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_QM056_CopyMatrixRow
		/// </summary>
		private void PS_QM056_CopyMatrixRow()
		{
			int i;

			try
			{
				oForm.Freeze(true);
				oDS_PS_QM056H.SetValue("DocEntry", 0, "");

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
				{
					oMat.FlushToDataSource();
					oDS_PS_QM056H.SetValue("DocEntry", i, "");
					oMat.LoadFromDataSource();
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
		/// PS_QM056_FormClear
		/// </summary>
		private void PS_QM056_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_QM056'", "");

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
		/// PS_QM056_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_QM056_DataValidCheck()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;
			int i;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM056_FormClear();
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
					if (string.IsNullOrEmpty(oMat.Columns.Item("InspItem").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("InspItem").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "검사항목은 필수입니다.";
						throw new Exception();
					}
				}

				oMat.FlushToDataSource();
				oDS_PS_QM056L.RemoveRecord(oDS_PS_QM056L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_QM056_FormClear();
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
		/// PS_QM056_LoadData
		/// </summary>
		private void PS_QM056_LoadData()
		{
			int i;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				sQry = "Select b.U_InspItem, b.U_InspItNm, b.U_InspSpec, b.U_GageNo ";
				sQry += " From [@PS_QM055H] a INNER JOIN [@PS_QM055L] b ON a.Code = b.Code ";
				sQry += "Where a.U_ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "' Order By b.U_Seqno ";
				oRecordSet.DoQuery(sQry);

				oDS_PS_QM056L.Clear();
				oMat.Clear();
				oMat.FlushToDataSource();

				i = 0;
				while (!oRecordSet.EoF)
				{
					oDS_PS_QM056L.InsertRecord(i);
					oDS_PS_QM056L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_QM056L.SetValue("U_InspItem", i, oRecordSet.Fields.Item(0).Value.ToString().Trim());
					oDS_PS_QM056L.SetValue("U_InspItNm", i, oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oDS_PS_QM056L.SetValue("U_InspSpec", i, oRecordSet.Fields.Item(2).Value.ToString().Trim());
					oDS_PS_QM056L.SetValue("U_GageNo", i, oRecordSet.Fields.Item(3).Value.ToString().Trim());
					i += 1;
					oRecordSet.MoveNext();
				}

				oMat.LoadFromDataSource();
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
		/// PS_QM056_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_QM056_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string DocEntry;
			string ItemCode;
			string sQry;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

				sQry = "Select U_AQLYN From [@PS_QM055H] Where U_ItemCode = '" + ItemCode + "'";
				oRecordSet.DoQuery(sQry);

				if (oRecordSet.Fields.Item(0).Value == "N")
				{
					WinTitle = "[PS_QM056_01] 검사기록서 출력";
					ReportName = "PS_QM056_01.rpt";
				}
				else
				{
					WinTitle = "[PS_QM056_02] 검사기록서 출력";
					ReportName = "PS_QM056_02.rpt";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
		/// PS_QM056_Print_Report02
		/// </summary>
		[STAThread]
		private void PS_QM056_Print_Report02()
		{
			string WinTitle;
			string ReportName;
			string DocEntry;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

				WinTitle = "[PS_QM056_05] 보증서및의뢰서출력";
				ReportName = "PS_QM056_05.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
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
                //	Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
							if (PS_QM056_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_QM056_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
							oFormMode01 = oForm.Mode;
						}
					}
					if (pVal.ItemUID == "Button01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM056_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					if (pVal.ItemUID == "Button02")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_QM056_Print_Report02);
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
								PS_QM056_FormItemEnabled();
								PS_QM056_AddMatrixRow(0, true);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_QM056_FormItemEnabled();
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
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
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
			int i;
			Double YQty;
			string AQLYN;
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
							if (pVal.ColUID == "FailQty") //불량수량
							{
								oMat.FlushToDataSource();

								oDS_PS_QM056L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								YQty = Convert.ToDouble(oMat.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) - Convert.ToDouble(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_QM056L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(YQty));

								for (i = pVal.Row; i <= oMat.VisualRowCount - 1; i++)
								{
									if (!string.IsNullOrEmpty(oDS_PS_QM056L.GetValue("U_InspItem", i).ToString().Trim()))
									{
										oDS_PS_QM056L.SetValue("U_PQty", i, Convert.ToString(YQty)); //검사수량
										oDS_PS_QM056L.SetValue("U_YQty", i, Convert.ToString(YQty)); //합격수량
									}
								}
								oMat.LoadFromDataSource();
							}
							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oMat.AutoResizeColumns();
						}
						else if (pVal.ItemUID == "CardCode") //납품처
						{
							oDS_PS_QM056H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("cardname", "cardcode", "[ocrd]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
						}
						else if (pVal.ItemUID == "ItemCode") //품목코드
						{
							oDS_PS_QM056H.SetValue("U_ItemName", 0, dataHelpClass.Get_ReData("U_ItemName", "U_ItemCode", "[@PS_QM055H]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
							oDS_PS_QM056H.SetValue("U_Size", 0, dataHelpClass.Get_ReData("U_Size", "U_ItemCode", "[@PS_QM055H]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
							oDS_PS_QM056H.SetValue("U_OutSize", 0, dataHelpClass.Get_ReData("U_OutSize", "U_ItemCode", "[@PS_QM055H]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
							AQLYN = dataHelpClass.Get_ReData("U_AQLYN", "U_ItemCode", "[@PS_QM055H]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", "");

							if (string.IsNullOrEmpty(AQLYN))
							{
								AQLYN = "N";
							}

							if (AQLYN == "N")
							{
								oForm.Items.Item("LotSize").Enabled = false;
								oForm.Items.Item("Sample").Enabled = false;
								oForm.Items.Item("AQL1").Enabled = false;
								oForm.Items.Item("AQL2").Enabled = false;
								oForm.Items.Item("LotSize").Specific.Value = "0";
								oForm.Items.Item("Sample").Specific.Value = "0";
								oForm.Items.Item("AQL1").Specific.Value = "";
								oForm.Items.Item("AQL2").Specific.Value = "";
							}
							else
							{
								oForm.Items.Item("LotSize").Enabled = true;
								oForm.Items.Item("Sample").Enabled = true;
								oForm.Items.Item("AQL1").Enabled = true;
								oForm.Items.Item("AQL2").Enabled = true;
							}

							PS_QM056_LoadData();
						}
						else if (pVal.ItemUID == "PQty") //검사수량
						{
							oMat.FlushToDataSource();

							for (i = 0; i <= oMat.VisualRowCount - 1; i++)
							{
								if (!string.IsNullOrEmpty(oDS_PS_QM056L.GetValue("U_InspItem", i).ToString().Trim()))
								{
									oDS_PS_QM056L.SetValue("U_PQty", i, oForm.Items.Item("PQty").Specific.Value.ToString().Trim()); //검사수량
									oDS_PS_QM056L.SetValue("U_YQty", i, oForm.Items.Item("PQty").Specific.Value.ToString().Trim()); //합격수량
								}
							}

							oMat.LoadFromDataSource();
						}
						else if (pVal.ItemUID == "ReqNo") //의뢰번호
						{
							oForm.Items.Item("LotNo").Specific.Value = oForm.Items.Item("ReqNo").Specific.Value.ToString().Trim();
						}
					}
					else if (pVal.BeforeAction == false)
					{
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
					PS_QM056_FormItemEnabled();
					PS_QM056_AddMatrixRow(oMat.VisualRowCount, false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM056H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_QM056L);
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pval.BeforeAction == true)
					{
					}
					else if (pval.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_QM056L.RemoveRecord(oDS_PS_QM056L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_QM056_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_QM056L.GetValue("U_CntcCode", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_QM056_AddMatrixRow(oMat.RowCount, false);
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
							PS_QM056_FormItemEnabled();
							break;
						case "1282": //추가
							PS_QM056_FormItemEnabled();
							PS_QM056_AddMatrixRow(0, true);
							break;
						case "1287": //복제
							PS_QM056_CopyMatrixRow();
							break;
						case "1288": //레코드이동(최초)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(다음)
						case "1291": //레코드이동(최종)
							PS_QM056_FormItemEnabled();
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
