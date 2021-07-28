using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 납기조회
	/// </summary>
	internal class PS_PP980 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP980L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP980.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP980_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP980");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_PP980_CreateItems();
				PS_PP980_SetComboBox();
				PS_PP980_AddMatrixRow(0, true);

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1285", false); // 복원
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", false); // 행삭제
				oForm.EnableMenu("1281", false);
				oForm.EnableMenu("1282", true);
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
		/// PS_PP980_CreateItems
		/// </summary>
		private void PS_PP980_CreateItems()
		{
			try
			{
				oDS_PS_PP980L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//기간(시작)
				oForm.DataSources.UserDataSources.Add("FrDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDate").Specific.DataBind.SetBound(true, "", "FrDate");
				oForm.Items.Item("FrDate").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";

				//기간(종료)
				oForm.DataSources.UserDataSources.Add("ToDate", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDate").Specific.DataBind.SetBound(true, "", "ToDate");
				oForm.Items.Item("ToDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

				//일자기준
				oForm.DataSources.UserDataSources.Add("DateStd", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("DateStd").Specific.DataBind.SetBound(true, "", "DateStd");

				//거래처구분
				oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

				//자체/외주
				oForm.DataSources.UserDataSources.Add("InOutGbn", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("InOutGbn").Specific.DataBind.SetBound(true, "", "InOutGbn");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
				oForm.Items.Item("ItemType").Specific.DataBind.SetBound(true, "", "ItemType");

				//납기준수건수
				oForm.DataSources.UserDataSources.Add("AssignYCnt", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 20);
				oForm.Items.Item("AssignYCnt").Specific.DataBind.SetBound(true, "", "AssignYCnt");

				//생산완료건수
				oForm.DataSources.UserDataSources.Add("CompltYCnt", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 20);
				oForm.Items.Item("CompltYCnt").Specific.DataBind.SetBound(true, "", "CompltYCnt");

				//납기준수율
				oForm.DataSources.UserDataSources.Add("AssignRate", SAPbouiCOM.BoDataType.dt_QUANTITY);
				oForm.Items.Item("AssignRate").Specific.DataBind.SetBound(true, "", "AssignRate");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP980_SetComboBox
		/// </summary>
		private void PS_PP980_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//거래처구분
				oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
				oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//품목구분
				oForm.Items.Item("ItemType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", false, false);
				oForm.Items.Item("ItemType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//일자기준
				oForm.Items.Item("DateStd").Specific.ValidValues.Add("%", "선택");
				oForm.Items.Item("DateStd").Specific.ValidValues.Add("1", "수주납기일");
				oForm.Items.Item("DateStd").Specific.ValidValues.Add("2", "생산완료일");
				oForm.Items.Item("DateStd").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//자체/외주
				oForm.Items.Item("InOutGbn").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("InOutGbn").Specific.ValidValues.Add("IN", "자체");
				oForm.Items.Item("InOutGbn").Specific.ValidValues.Add("OUT", "외주");
				oForm.Items.Item("InOutGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP980_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP980_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				if (RowIserted == false)
				{
					oDS_PS_PP980L.InsertRecord(oRow);
				}

				oMat.AddRow();
				oDS_PS_PP980L.Offset = oRow;
				oDS_PS_PP980L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));

				oMat.LoadFromDataSource();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP980_MTX01
		/// </summary>
		private void PS_PP980_MTX01()
		{
			short i;
			string sQry ;
			string errMessage = string.Empty;
			int AssignYCnt = 0;	//납기준수건수
			int CompltYCnt;	//생산완료건수
			double AssignRate;	//납기준수율
			string DateStd;		//일자기준
			string FrDate;		//기간(시작)
			string ToDate;		//기간(종료)
			string CardType;	//거래처구분
			string ItemType;	//품목구분
			string InOutGbn;    //자체/외주
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);

				DateStd = oForm.Items.Item("DateStd").Specific.Value.ToString().Trim();
				FrDate = oForm.Items.Item("FrDate").Specific.Value.ToString().Trim();
				ToDate = oForm.Items.Item("ToDate").Specific.Value.ToString().Trim();
				CardType = oForm.Items.Item("CardType").Specific.Value.ToString().Trim();
				ItemType = oForm.Items.Item("ItemType").Specific.Value.ToString().Trim();
				InOutGbn = oForm.Items.Item("InOutGbn").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = "EXEC [PS_PP980_01] '" + DateStd + "','" + FrDate + "','" + ToDate + "','" + CardType + "','" + ItemType + "','" + InOutGbn + "', '1'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_PP980L.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
					PS_PP980_AddMatrixRow(0, true);

					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_PP980L.Size)
					{
						oDS_PS_PP980L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_PP980L.Offset = i;

					oDS_PS_PP980L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_PP980L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim()); //품목코드
					oDS_PS_PP980L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim()); //품목명
					oDS_PS_PP980L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("Spec").Value.ToString().Trim());     //규격
					oDS_PS_PP980L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("DueDate").Value.ToString().Trim());  //수주납기일
					oDS_PS_PP980L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("DocDate").Value.ToString().Trim());  //생산완료일
					oDS_PS_PP980L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("InOutGbn").Value.ToString().Trim()); //자체/외주 구분
					oDS_PS_PP980L.SetValue("U_ColReg07", i, oRecordSet.Fields.Item("AssignYN").Value.ToString().Trim()); //준수여부

					//준수여부 건수 카운트
					if (oRecordSet.Fields.Item("AssignYN").Value.ToString().Trim() == "Y")
					{
						AssignYCnt += 1;
					}

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				CompltYCnt = oRecordSet.RecordCount;	    //생산완료 건수
				AssignRate = AssignYCnt / CompltYCnt * 100;	//납기준수율

				oForm.Items.Item("AssignYCnt").Specific.Value = AssignYCnt; //납기준수건수
				oForm.Items.Item("CompltYCnt").Specific.Value = CompltYCnt; //생산완료건수
				oForm.Items.Item("AssignRate").Specific.Value = AssignRate; //납기준수율
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
		/// PS_PP980_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP980_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("FrDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "기간(시작)은 필수사항입니다. 확인하세요.";
					oForm.Items.Item("FrDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("ToDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "기간(종료)은 필수사항입니다. 확인하세요.";
					oForm.Items.Item("ToDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DateStd").Specific.Value.ToString().Trim()))
				{
					errMessage = "일자기준은 필수사항입니다. 확인하세요.";
					oForm.Items.Item("DateStd").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					throw new Exception();
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
		/// Raise_EVENT_ITEM_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ITEM_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
					if (pval.ItemUID == "BtnSearch")
					{
						if (PS_PP980_DelHeaderSpaceLine() == false)
						{
							BubbleEvent = false;
							return;
						}
						else
						{
							PS_PP980_MTX01();
						}
					}
				}
				else if (pval.BeforeAction == false)
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
					if (pval.ItemUID == "Mat01")
					{
						if (pval.Row > 0)
						{
							oMat.SelectRow(pval.Row, true, false);
						}
					}
				}
				else if (pval.BeforeAction == false)
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// Raise_EVENT_DOUBLE_CLICK
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
					if (pval.ItemUID == "Mat01")
					{
						if (pval.Row == 0)
						{
							oMat.Columns.Item(pval.ColUID).TitleObject.Sortable = true;
							oMat.FlushToDataSource();
						}
					}
				}
				else if (pval.BeforeAction == false)
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
		/// <param name="pval"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_FORM_UNLOAD(string FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
		{
			try
			{
				if (pval.BeforeAction == true)
				{
				}
				else if (pval.BeforeAction == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP980L);
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
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							oForm.Freeze(true);
							PS_PP980_AddMatrixRow(oMat.VisualRowCount, false);
							oForm.Freeze(false);
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
						case "7169": //엑셀 내보내기
							//엑셀 내보내기 이후 처리
							oForm.Freeze(true);
							oDS_PS_PP980L.RemoveRecord(oDS_PS_PP980L.Size - 1);
							oMat.LoadFromDataSource();
							oForm.Freeze(false);
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
