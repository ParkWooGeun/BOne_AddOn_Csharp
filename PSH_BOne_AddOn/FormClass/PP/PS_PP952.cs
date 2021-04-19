using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 지체건 상세내역
	/// </summary>
	internal class PS_PP952 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP952L;  //등록라인

		/// <summary>
		/// LoadForm
		/// </summary>
		public override void LoadForm()
		{

			int i;
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP952.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP952_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP952");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP952_CreateItems();
				PS_PP952_ComboBox_Setting();

				oForm.EnableMenu(("1283"), false);				//// 삭제
				oForm.EnableMenu(("1286"), false);				//// 닫기
				oForm.EnableMenu(("1287"), false);				//// 복제
				oForm.EnableMenu(("1285"), false);				//// 복원
				oForm.EnableMenu(("1284"), false);				//// 취소
				oForm.EnableMenu(("1293"), false);				//// 행삭제
				oForm.EnableMenu(("1281"), false);
				oForm.EnableMenu(("1282"), true);
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
		/// PS_PP952_CreateItems
		/// </summary>
		private void PS_PP952_CreateItems()
		{
			try
			{
				oDS_PS_PP952L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				// 메트릭스 개체 할당
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//기준일자
				oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
				oForm.DataSources.UserDataSources.Item("DocDate").Value = DateTime.Now.ToString("yyyyMMdd");

				//구분
				oForm.DataSources.UserDataSources.Add("Cls", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Cls").Specific.DataBind.SetBound(true, "", "Cls");

				//지체일수
				oForm.DataSources.UserDataSources.Add("DayDiff", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("DayDiff").Specific.DataBind.SetBound(true, "", "DayDiff");

				//외주구분
				oForm.DataSources.UserDataSources.Add("OutYN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("OutYN").Specific.DataBind.SetBound(true, "", "OutYN");

			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP952_ComboBox_Setting
		/// </summary>
		private void PS_PP952_ComboBox_Setting()
	    {
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//구분
				oForm.Items.Item("Cls").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("Cls").Specific.ValidValues.Add("01", "안강사업장");
				oForm.Items.Item("Cls").Specific.ValidValues.Add("02", "부산사업장");
				oForm.Items.Item("Cls").Specific.ValidValues.Add("03", "울산사업장");
				oForm.Items.Item("Cls").Specific.ValidValues.Add("04", "보그워너");
				oForm.Items.Item("Cls").Specific.ValidValues.Add("05", "에네스지");
				oForm.Items.Item("Cls").Specific.ValidValues.Add("06", "기타");
				oForm.Items.Item("Cls").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//지체일수
				oForm.Items.Item("DayDiff").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("DayDiff").Specific.ValidValues.Add("01", "1~10일");
				oForm.Items.Item("DayDiff").Specific.ValidValues.Add("02", "11~20일");
				oForm.Items.Item("DayDiff").Specific.ValidValues.Add("03", "21~30일");
				oForm.Items.Item("DayDiff").Specific.ValidValues.Add("04", "30일 이상");
				oForm.Items.Item("DayDiff").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//외주구분
				oForm.Items.Item("OutYN").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("OutYN").Specific.ValidValues.Add("N", "자체");
				oForm.Items.Item("OutYN").Specific.ValidValues.Add("Y", "외주");
				oForm.Items.Item("OutYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//매트릭스-외주구분
				dataHelpClass.Combo_ValidValues_Insert("PS_PP952", "Mat01", "OutYN", "Y", "외주");
				dataHelpClass.Combo_ValidValues_Insert("PS_PP952", "Mat01", "OutYN", "N", "자체");
				dataHelpClass.Combo_ValidValues_SetValueColumn(oMat.Columns.Item("OutYN"), "PS_PP952", "Mat01", "OutYN", false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP952_MTX01  데이터 조회
		/// </summary>
		private void PS_PP952_MTX01()
		{
			short i;
			string sQry;
			string errMessage = String.Empty;

			string DocDate;	//기준일자
			string Cls;		//구분
			string DayDiff;	//지체일수
			string OutYN;   //외주구분

			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				Cls = oForm.Items.Item("Cls").Specific.Value.ToString().Trim();
				DayDiff = oForm.Items.Item("DayDiff").Specific.Value.ToString().Trim();
				OutYN = oForm.Items.Item("OutYN").Specific.Value.ToString().Trim();

				if (Cls == "%")
				{
					Cls = "";
				}

				if (DayDiff == "%")
				{
					DayDiff = "";
				}

				if (OutYN == "%")
				{
					OutYN = "";
				}

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC [PS_PP952_01] '" + DocDate + "','" + Cls + "','" + DayDiff + "','" + OutYN + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP952L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP952L.Size)
					{
						oDS_PS_PP952L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP952L.Offset = i;

					oDS_PS_PP952L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP952L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());      //작번
					oDS_PS_PP952L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());    //거래처코드
					oDS_PS_PP952L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());    //거래처명
					oDS_PS_PP952L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("OutYN").Value.ToString().Trim());       //외주구분
					oDS_PS_PP952L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("ORDRQty").Value.ToString().Trim());     //수주수량
					oDS_PS_PP952L.SetValue("U_ColQty02", i, oRecordSet.Fields.Item("PP080Qty").Value.ToString().Trim());    //생산완료수량
					oDS_PS_PP952L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("DueDate").Value.ToString().Trim());     //납기일
					oDS_PS_PP952L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("DayDiff").Value.ToString().Trim());     //지체일
					oDS_PS_PP952L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("Comment").Value.ToString().Trim());     //비고

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
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
				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// private
		/// </summary>
		[STAThread]
		private void PS_PP952_Print_Report01()
		{
			string WinTitle;
			string ReportName;

			string DocDate; //기준일자
			string Cls;     //구분
			string DayDiff; //지체일수
			string OutYN;   //외주구분

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				Cls = oForm.Items.Item("Cls").Specific.Value.ToString().Trim();
				DayDiff = oForm.Items.Item("DayDiff").Specific.Value.ToString().Trim();
				OutYN = oForm.Items.Item("OutYN").Specific.Value.ToString().Trim();

				if (Cls == "%")
				{
					Cls = "";
				}

				if (DayDiff == "%")
				{
					DayDiff = "";
				}

				if (OutYN == "%")
				{
					OutYN = "";
				}

				WinTitle = "[PS_PP952] 레포트";
				ReportName = "PS_PP952_01.rpt";

				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocDate", DateTime.ParseExact(DocDate, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@Cls", Cls));
				dataPackParameter.Add(new PSH_DataPackClass("@DayDiff", DayDiff));
				dataPackParameter.Add(new PSH_DataPackClass("@OutYN", OutYN));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter, dataPackFormula);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    //Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "BtnSearch")
					{
						PS_PP952_MTX01();
					}
					else if (pVal.ItemUID == "BtnPrint")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP952_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
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
							oMat.SelectRow(pVal.Row, true, false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP952L);
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
						case "1285": //복원
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
						case "1285": //복원
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
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
							break;
					}
				}
				else if (BusinessObjectInfo.BeforeAction == false)
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:    //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:     //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:  //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:  //36
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
