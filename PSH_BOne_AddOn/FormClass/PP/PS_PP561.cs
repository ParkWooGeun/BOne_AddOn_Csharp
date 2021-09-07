using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 장비별 예상금액(월별) 등록
	/// </summary>
	internal class PS_PP561 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;			
		private SAPbouiCOM.DBDataSource oDS_PS_PP561H; //등록헤더 			
		private SAPbouiCOM.DBDataSource oDS_PS_PP561L; //등록라인
		private string oLastItemUID01;//클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;//마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;//마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP561.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP561_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP561");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry"; 

				oForm.Freeze(true);

				PS_PP561_CreateItems();
				PS_PP561_SetComboBox();
				PS_PP561_EnableMenus();
				PS_PP561_SetDocument(oFormDocEntry);

				oForm.Items.Item("OrdNum").Click();
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
		/// PS_PP561_CreateItems
		/// </summary>
		private void PS_PP561_CreateItems()
		{
			try
			{
				oDS_PS_PP561H = oForm.DataSources.DBDataSources.Item("@PS_PP561H");
				oDS_PS_PP561L = oForm.DataSources.DBDataSources.Item("@PS_PP561L");
				oMat = oForm.Items.Item("Mat01").Specific;

				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;

				oForm.Items.Item("StdYM").Specific.Value = DateTime.Now.ToString("yyyyMM");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP561_SetComboBox
		/// </summary>
		private void PS_PP561_SetComboBox()
		{
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//Matrix
				//비용구분
				sQry = " SELECT      T0.U_Minor, ";
				sQry += "                T0.U_CdName";
				sQry += " FROM       [@PS_SY001L] T0";
				sQry += " WHERE      T0.Code = 'P210'";
				sQry += "                AND T0.U_UseYN = 'Y'";
				sQry += " ORDER BY  T0.U_Seq";
				dataHelpClass.GP_MatrixSetMatComboList(oMat.Columns.Item("AmtCls"), sQry, "", "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP561_EnableMenus
		/// </summary>
		private void PS_PP561_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, false, false, false, true, false, true, true, true, true,	false, false, false, false, false, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP561_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_PP561_SetDocument(string oFormDocEntry)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_PP561_EnableFormItem();
					PS_PP561_AddMatrixRow(0, true); 
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
		/// PS_PP561_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_PP561_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				
				if (RowIserted == false) //행추가여부
				{
					oDS_PS_PP561L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_PP561L.Offset = oRow;
				oDS_PS_PP561L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_PP561_EnableFormItem
		/// </summary>
		private void PS_PP561_EnableFormItem()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					oForm.Items.Item("BtnSrch").Enabled = true;
					PS_PP561_ClearForm();
					oForm.EnableMenu("1281", true);     //찾기
					oForm.EnableMenu("1282", false);    //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocEntry").Specific.Value = "";
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = false;
					oForm.Items.Item("BtnSrch").Enabled = false;
					oForm.EnableMenu("1281", false);     //찾기
					oForm.EnableMenu("1282", true);      //추가
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Mat01").Enabled = true;
					oForm.Items.Item("BtnSrch").Enabled = false;
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
		/// PS_PP561_ClearForm
		/// </summary>
		private void PS_PP561_ClearForm()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP561'", "");

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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP561_CheckDataValid
		/// </summary>
		/// <returns></returns>
		private bool PS_PP561_CheckDataValid()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_PP561_ClearForm();
				}
				
				if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim()))
				{
					errMessage = "작번을 입력하지 않았습니다.";
					throw new Exception();
				}

				if (string.IsNullOrEmpty(oForm.Items.Item("StdYM").Specific.Value.ToString().Trim()))
				{
					errMessage = "기준년월을 입력하지 않았습니다.";
					throw new Exception();
				}

				if (oMat.VisualRowCount == 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}

				oDS_PS_PP561L.RemoveRecord(oDS_PS_PP561L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_PP561_ClearForm();
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
		/// PS_PP561_MTX01
		/// </summary>
		private void PS_PP561_MTX01()
		{
			int loopCount;
			string sQry;
			string errMessage = string.Empty;
			string OrdNum;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				OrdNum = oForm.Items.Item("OrdNum").Specific.Value.ToString().Trim();

				ProgressBar01.Text = "조회중...";

				sQry = " EXEC [PS_PP561_01] '";
				sQry += OrdNum + "'";

				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					errMessage = "결과가 존재하지 않습니다.";
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_PP561L.InsertRecord(loopCount);
					}
					oDS_PS_PP561L.Offset = loopCount;

					oDS_PS_PP561L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));
					oDS_PS_PP561L.SetValue("U_AmtCls", loopCount, oRecordSet.Fields.Item("AmtCls").Value.ToString().Trim());   //비용구분
					oDS_PS_PP561L.SetValue("U_OrdNum", loopCount, oRecordSet.Fields.Item("OrdNum").Value.ToString().Trim());   //작번
					oDS_PS_PP561L.SetValue("U_OrdSub1", loopCount, oRecordSet.Fields.Item("OrdSub1").Value.ToString().Trim()); //서브작번1
					oDS_PS_PP561L.SetValue("U_OrdSub2", loopCount, oRecordSet.Fields.Item("OrdSub2").Value.ToString().Trim()); //서즈작번2
					oDS_PS_PP561L.SetValue("U_OrdName", loopCount, oRecordSet.Fields.Item("OrdName").Value.ToString().Trim()); //작명
					oDS_PS_PP561L.SetValue("U_AddAmt", loopCount, oRecordSet.Fields.Item("AddAmt").Value.ToString().Trim());   //목표금액
					oDS_PS_PP561L.SetValue("U_Comment", loopCount, oRecordSet.Fields.Item("Comment").Value.ToString().Trim()); //비고

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oDS_PS_PP561H.SetValue("U_TAddAmt", 0, "0"); //합계금액 초기화

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				oForm.Update();
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
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_PP561_Validate
		/// </summary>
		/// <param name="ValidateType"></param>
		/// <returns></returns>
		private bool PS_PP561_Validate(string ValidateType)
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (ValidateType == "수정")
				{
				}
				else if (ValidateType == "행삭제")
				{
					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
					{
						if (string.IsNullOrEmpty(oMat.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value.ToString().Trim()))
						{
						}
						else
						{
							errMessage = "해당 내용은 행삭제할수 없습니다.";
							throw new Exception();
						}
					}
				}
				else if (ValidateType == "취소")
				{
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
		/// PS_PP561_SUM 금액 합계
		/// </summary>
		/// <param name="pAmtType"></param>
		private void PS_PP561_SUM(string pAmtType = "")
		{
			short loopCount;
			double TotalAddAmt = 0;

			try
			{
				for (loopCount = 1; loopCount <= oMat.VisualRowCount - 1; loopCount++)
				{
					if (pAmtType == "AddAmt")
					{
						TotalAddAmt += Convert.ToDouble(oDS_PS_PP561L.GetValue("U_AddAmt", loopCount - 1));  //예상금액
					}
					else if (string.IsNullOrEmpty(pAmtType))
					{
						TotalAddAmt += Convert.ToDouble(oDS_PS_PP561L.GetValue("U_AddAmt", loopCount - 1));  //예상금액
					}
				}
				if (pAmtType == "AddAmt")
				{
					oDS_PS_PP561H.SetValue("U_TAddAmt", 0, Convert.ToString(TotalAddAmt));		//예상금액
				}
				else if (string.IsNullOrEmpty(pAmtType))
				{
					oDS_PS_PP561H.SetValue("U_TAddAmt", 0, Convert.ToString(TotalAddAmt));		//예상금액
				}
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
			switch (pVal.EventType) {
				case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:				//1
					Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:					//2
					Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:				//5
					Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_CLICK:					    //6
					Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:				//7
					//Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:		//8
					//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE:					//10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:				//11
					Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:				//18
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:			//19
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:				//20
					break;
				case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:			//27
					break;
				case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:					//3
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:					//4
					break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:				//17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
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
			string errMessage = string.Empty;

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_PP561_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP561_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
					else if (pVal.ItemUID == "BtnSrch")
					{
						if (!string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
						{
							PS_PP561_MTX01();
							PS_PP561_AddMatrixRow(oMat.RowCount, false);
						}
						else
						{
							BubbleEvent = false;
							errMessage = "조회하려는 작번을 입력하세요.";
							throw new Exception();
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
								PS_PP561_EnableFormItem();
								PS_PP561_AddMatrixRow(0, true); 
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_PP561_EnableFormItem();
							}
						}
					}

					if (pVal.ItemUID == "Mat01")
					{
						if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (oMat.RowCount >= pVal.Row) //빈 Select 필드에 클릭했을 때 생기는 오류 수정을 위한 구문
							{
							}
						}
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.ColUID == "BatchNum")
						{
							dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "BatchNum");
						}
					}
					else
					{
						if (pVal.ItemUID == "OrdNum")
						{
							dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "OrdNum", "");
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							oMat.FlushToDataSource();

							if (pVal.ColUID == "AmtCls")
							{
								//기타작업
								oDS_PS_PP561L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value.ToString().Trim());

								if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP561L.GetValue("U_AmtCls", pVal.Row - 1).ToString().Trim()))
								{
									PS_PP561_AddMatrixRow(pVal.Row, false);
								}
							}
							else
							{
								oDS_PS_PP561L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value.ToString().Trim());
							}

							oMat.LoadFromDataSource();
							oMat.AutoResizeColumns();
						}
						else
						{
						}

						oForm.Update();
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
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

							if (pVal.ColUID == "AmtCls")
							{
								//비용구분을 선택하면
								if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP561L.GetValue("U_AmtType", pVal.Row - 1).ToString().Trim()))
								{
									PS_PP561_AddMatrixRow(pVal.Row, false);
								}
							}
							else if (pVal.ColUID == "AddAmt")
							{
								PS_PP561_SUM("AddAmt");
							}

							oMat.LoadFromDataSource();
							oMat.AutoResizeColumns();
						}
						else
						{
							if (pVal.ItemUID == "DocEntry")
							{
								oDS_PS_PP561H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
							else if (pVal.ItemUID == "OrdNum")
							{
								oDS_PS_PP561H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
								oDS_PS_PP561H.SetValue("U_ItemName", 0, dataHelpClass.Get_ReData("FrgnName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
								oDS_PS_PP561H.SetValue("U_ItemSpec", 0, dataHelpClass.Get_ReData("U_Size", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
							}
							else
							{
								oDS_PS_PP561H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}

							oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
					PS_PP561_EnableFormItem();
					PS_PP561_AddMatrixRow(oMat.VisualRowCount, false); 
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP561H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP561L);
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
		private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.IMenuEvent pVal, bool BubbleEvent)
		{
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
						if (PS_PP561_Validate("행삭제") == false)
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
						oDS_PS_PP561L.RemoveRecord(oDS_PS_PP561L.Size - 1);
						PS_PP561_SUM();
						oMat.LoadFromDataSource();
						oForm.Update();

						if (oMat.RowCount == 0)
						{
							PS_PP561_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_PP561L.GetValue("U_AmtCls", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_PP561_AddMatrixRow(oMat.RowCount, false);
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
				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
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
							Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
							break;
						case "1281": //찾기
							PS_PP561_EnableFormItem(); 
							oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_PP561_EnableFormItem(); 
							PS_PP561_AddMatrixRow(0, true); 
							oForm.Items.Item("StdYM").Specific.Value = DateTime.Now.ToString("yyyyMM");
							oForm.Items.Item("OrdNum").Click();
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_PP561_EnableFormItem();
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
