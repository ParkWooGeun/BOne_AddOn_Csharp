using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 분개전표 연결발행
	/// </summary>
	internal class PS_FI420 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_FI420L;  //등록헤더

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm(string oFormDocEntry01)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI420.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI420_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI420");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				CreateItems();
				ComboBox_Setting();
				Initialization();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc01); //메모리 해제
			}
		}

		/// <summary>
		/// CreateItems
		/// </summary>
		private void CreateItems()
		{
			SAPbouiCOM.OptionBtn optBtn = null;

			try
			{
				oDS_PS_FI420L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				oForm.DataSources.UserDataSources.Add("PntGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("PntGbn").Specific.DataBind.SetBound(true, "", "PntGbn");

				oForm.DataSources.UserDataSources.Add("DocType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("DocType").Specific.DataBind.SetBound(true, "", "DocType");

				oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");

				oForm.DataSources.UserDataSources.Add("OptionDS01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				optBtn = oForm.Items.Item("Rad01").Specific;
				optBtn.ValOn = "1";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");

				//optBtn.Selected = True

				optBtn = oForm.Items.Item("Rad02").Specific;
				optBtn.ValOn = "2";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");
				optBtn.GroupWith("Rad01");

				optBtn = oForm.Items.Item("Rad03").Specific;
				optBtn.ValOn = "3";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");
				optBtn.GroupWith("Rad01");

				optBtn = oForm.Items.Item("Rad04").Specific;
				optBtn.ValOn = "4";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");
				optBtn.GroupWith("Rad01");

				optBtn = oForm.Items.Item("Rad05").Specific;
				optBtn.ValOn = "5";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS01");
				optBtn.GroupWith("Rad01");

				oForm.DataSources.UserDataSources.Add("OptionDS11", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
				optBtn = oForm.Items.Item("Rad11").Specific;
				optBtn.ValOn = "1";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");

				//optBtn.Selected = True

				optBtn = oForm.Items.Item("Rad12").Specific;
				optBtn.ValOn = "2";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");
				optBtn.GroupWith("Rad11");

				optBtn = oForm.Items.Item("Rad13").Specific;
				optBtn.ValOn = "3";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");
				optBtn.GroupWith("Rad11");

				optBtn = oForm.Items.Item("Rad14").Specific;
				optBtn.ValOn = "4";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");
				optBtn.GroupWith("Rad11");

				optBtn = oForm.Items.Item("Rad15").Specific;
				optBtn.ValOn = "5";
				optBtn.ValOff = "0";
				optBtn.DataBind.SetBound(true, "", "OptionDS11");
				optBtn.GroupWith("Rad11");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(optBtn);
			}
		}

		/// <summary>
		/// ComboBox_Setting
		/// </summary>
		private void ComboBox_Setting()
		{
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				//콤보에 기본값설정
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				// 전표유형
				oForm.Items.Item("DocType").Specific.ValidValues.Add("24", "입금");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("46", "지급");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("13", "판매");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("99", "기타(입금,지급,판매,제외)");
				oForm.Items.Item("DocType").Specific.ValidValues.Add("00", "전체");
				oForm.Items.Item("DocType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				oForm.Items.Item("PntGbn").Specific.ValidValues.Add("10", "연결발행");
				oForm.Items.Item("PntGbn").Specific.ValidValues.Add("20", "개별발행");
				oForm.Items.Item("PntGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
			string sQry = String.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "CntcCode":
						sQry = "Select lastName + firstName From OHEM Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
		/// LoadData
		/// </summary>
		private void LoadData()
		{
			int i;
			string sQry;
			string BPLID;
			string DocDate;
			string DocType;
			string errCode = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocType = oForm.Items.Item("DocType").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

				if (String.IsNullOrEmpty(DocDate))
				{
					PSH_Globals.SBO_Application.MessageBox("전기일자는 필수입력사항 입니다. 확인하세요.");
					return;
				}

				sQry = "EXEC [PS_FI420_01] '" + BPLID + "','" + DocType + "','" + DocDate + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oDS_PS_FI420L.Clear();

				if (oRecordSet.RecordCount == 0)
				{
					errCode = "1";
					throw new Exception();
				}

				oForm.Freeze(true);

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_FI420L.Size)
					{
						oDS_PS_FI420L.InsertRecord(i);
					}

					oMat01.AddRow();
					oDS_PS_FI420L.Offset = i;
					oDS_PS_FI420L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_FI420L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColDt01", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd"));
					oDS_PS_FI420L.SetValue("U_ColDt02", i, Convert.ToDateTime(oRecordSet.Fields.Item("DocDueDate").Value.ToString().Trim()).ToString("yyyyMMdd"));
					oDS_PS_FI420L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("DocTotal").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("JrnlMemo").Value.ToString().Trim());
					oDS_PS_FI420L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("TransId").Value.ToString().Trim());

					oRecordSet.MoveNext();
				}

				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();

			}
			catch (Exception ex)
			{
				if (errCode == "1")
				{
					PSH_Globals.SBO_Application.MessageBox("조회 결과가 없습니다. 확인하세요.");
				}
                else
                {
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				if (ProgBar01 != null)
				{
					ProgBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
				}
				
				oForm.Freeze(false);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// Print_Report01
		/// </summary>
		[STAThread]
		private void Print_Report01()
		{
			int i;
			string WinTitle;
			string ReportName;
			string DocType;
			string sQry;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocType = oForm.Items.Item("DocType").Specific.Value.ToString().Trim();

				WinTitle = "회계전표 [PS_FI420]";

				if (oForm.Items.Item("PntGbn").Specific.Value.ToString().Trim() == "20")
				{
					ReportName = "PS_FI420_02.rpt";
				}
				else
				{
					ReportName = "PS_FI420_01.rpt";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				// Formula 수식필드
				dataPackFormula.Add(new PSH_DataPackClass("@RadBtn01", oForm.DataSources.UserDataSources.Item("OptionDS01").Value));
				dataPackFormula.Add(new PSH_DataPackClass("@RadBtn11", oForm.DataSources.UserDataSources.Item("OptionDS11").Value));

				sQry = "Delete [Z_PS_FI420]";
				oRecordSet.DoQuery(sQry);

				oMat01.FlushToDataSource();
				for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
				{
					if (oDS_PS_FI420L.GetValue("U_ColReg01", i).ToString().Trim() == "Y")
					{
						sQry = "Insert [Z_PS_FI420] values ('" + oDS_PS_FI420L.GetValue("U_ColReg06", i).ToString().Trim() + "')";
						oRecordSet.DoQuery(sQry);
					}
				}

				// 조회조건문

				//// 조회조건문  (원본)
				//sQry = "EXEC [PS_FI420_02] '" + oForm.Items.Item("DocType").Specific.Value.ToString().Trim() + "'";
				//oRecordSet.DoQuery(sQry);
				////    If oRecordSet01.RecordCount = 0 Then
				////        ErrNum = 1
				////        GoTo Print_Report01_Error
				////    End If
				//if (oForm.Items.Item("DocType").Specific.Value.ToString().Trim() == "13")
				//{
				//	sQry = " Select * From  ZPS_FI420_TEMP Order by U_RptItm01,TransId, Convert(Numeric(12,0),Line_Id)";
				//}
				//else
				//{
				//	sQry = "Select  * From  ZPS_FI420_TEMP Order by TransId, Convert(Numeric(12,0),Line_Id) ";
				//}

				// 마이그레션시 FI420_02로 통합해서 새로작성  2020.09.21

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocType", DocType));

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
					Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
					//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
					//Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					//Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
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
			try
			{
				if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Btn01")
                    {
                        System.Threading.Thread thread = new System.Threading.Thread(Print_Report01);
                        thread.SetApartmentState(System.Threading.ApartmentState.STA);
                        thread.Start();
                    }
                    else if (pVal.ItemUID == "Btn02")
                    {
                        LoadData();
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
                            if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
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
			finally
			{
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
                    if (pVal.ItemChanged == true)
                    {
                        oForm.Freeze(true);
                        if (pVal.ItemUID == "BPLId" || pVal.ItemUID == "DocType")
                        {
                            oMat01.Clear();
                            oDS_PS_FI420L.Clear();
                        }
                        oForm.Freeze(false);
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
		/// DOUBLE_CLICK 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string Check = string.Empty;

			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
                    if (pVal.ItemUID == "Mat01" && pVal.Row == Convert.ToDouble("0") && pVal.ColUID == "Check")
                    {
                        oForm.Freeze(true);
                        oMat01.FlushToDataSource();
                        if (string.IsNullOrEmpty(oDS_PS_FI420L.GetValue("U_ColReg01", 0).ToString().Trim()) || oDS_PS_FI420L.GetValue("U_ColReg01", 0).ToString().Trim() == "N")
                        {
                            Check = "Y";
                        }
                        else if (oDS_PS_FI420L.GetValue("U_ColReg01", 0).ToString().Trim() == "Y")
                        {
                            Check = "N";
                        }

                        for (int i = 0; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            oDS_PS_FI420L.SetValue("U_ColReg01", i, Check);
                        }
                        oMat01.LoadFromDataSource();
                        oForm.Freeze(false);
                    }
                }
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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
					SubMain.Remove_Forms(oFormUniqueID01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI420L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}

		/// <summary>
		/// Raise_FormMenuEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormMenuEvent(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1293":							//행삭제
							break;
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284":							//취소
							break;
						case "1286":							//닫기
							break;
						case "1293":							//행삭제
							break;
						case "1281":							//찾기
							break;
						case "1282":							//추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":							//레코드이동버튼
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
			}
		}

		/// <summary>
		/// Raise_FormDataEvent
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:							//33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:							//34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:						//35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:						//36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:							//33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:							//34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:						//35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:						//36
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
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
			}
		}
	}
}
