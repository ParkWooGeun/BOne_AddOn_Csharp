using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 벌크반품등록
	/// </summary>
	internal class PS_PP777 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP777H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP777L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP777.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP777_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP777");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);

				PS_PP777_CreateItems();
				PS_PP777_SetComboBox();
				PS_PP777_ClearForm();
				PS_PP777_AddMatrixRow(1, 0, true); //oMat

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", true);  // 취소
				oForm.EnableMenu("1293", true);  // 행삭제
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
		/// PS_PP777_CreateItems
		/// </summary>
		private void PS_PP777_CreateItems()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oDS_PS_PP777H = oForm.DataSources.DBDataSources.Item("@PS_PP777H");
				oDS_PS_PP777L = oForm.DataSources.DBDataSources.Item("@PS_PP777L");
				oMat = oForm.Items.Item("Mat01").Specific;

				oDS_PS_PP777H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));

				//담당자
				oDS_PS_PP777H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD());
				PS_PP777_FlushToItemValue("CntcCode", 0, "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP777_SetComboBox
		/// </summary>
		private void PS_PP777_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				// 작업구분
				sQry = "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("OrdGbn").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(3, SAPbouiCOM.BoSearchKey.psk_Index);
				oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP777_ClearForm
		/// </summary>
		private void PS_PP777_ClearForm()
		{
			string DocNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP777'", "");
				if (Convert.ToDouble(DocNum) == 0)
				{
					oForm.Items.Item("DocNum").Specific.Value = 1;
				}
				else
				{
					oForm.Items.Item("DocNum").Specific.Value = DocNum;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
		}

		/// <summary>
		/// PS_PP777_AddMatrixRow
		/// </summary>
		/// <param name="oMat1"></param>
		/// <param name="oRow"></param>
		/// <param name="Insert_YN"></param>
		private void PS_PP777_AddMatrixRow(short oMat1, int oRow, bool Insert_YN)
		{
			try
			{
				switch (oMat1)
				{
					case 1:
						if (Insert_YN == false)
						{
							oRow = oMat.RowCount;
							oDS_PS_PP777L.InsertRecord(oRow);
						}
						//수입내역
						oDS_PS_PP777L.Offset = oRow;
						oDS_PS_PP777L.SetValue("LineId", oRow, Convert.ToString(oRow + 1));
						oDS_PS_PP777L.SetValue("U_PP070HL", oRow, "");
						oDS_PS_PP777L.SetValue("U_MovDocNo", oRow, "");
						oDS_PS_PP777L.SetValue("U_PP070No", oRow, "");
						oDS_PS_PP777L.SetValue("U_PP070NoL", oRow, "");
						oDS_PS_PP777L.SetValue("U_PorNum", oRow, "");
						oDS_PS_PP777L.SetValue("U_ItemCode", oRow, "");
						oDS_PS_PP777L.SetValue("U_ItemName", oRow, "");
						oDS_PS_PP777L.SetValue("U_Size", oRow, "");
						oDS_PS_PP777L.SetValue("U_PkQty", oRow, "");
						oDS_PS_PP777L.SetValue("U_PkWt", oRow, "");
						oDS_PS_PP777L.SetValue("U_OPkQty", oRow, "");
						oDS_PS_PP777L.SetValue("U_OPkWt", oRow, "");
						oDS_PS_PP777L.SetValue("U_InDate", oRow, "");
						oMat.LoadFromDataSource();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
		}

		/// <summary>
		/// PS_PP777_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP777_FlushToItemValue(string oUID, int oRow, string oCol)
		{
			string DocNum;
			string LineId;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (oUID)
				{
					case "CntcCode":
						sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oDS_PS_PP777H.GetValue("U_CntcCode", 0).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oDS_PS_PP777H.SetValue("U_CntcName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
						break;
				}
				//Line
				if (oUID == "Mat01")
				{
					switch (oCol)
					{
						case "PP070HL":
							oMat.FlushToDataSource();
							oDS_PS_PP777L.Offset = oRow - 1;

							DocNum = oMat.Columns.Item("PP070HL").Cells.Item(oRow).Specific.String.Split('-')[0];
							LineId = oMat.Columns.Item("PP070HL").Cells.Item(oRow).Specific.String.Split('-')[1];

							sQry = "PS_PP777_02 '" + DocNum + "','" + LineId + "'";
							oRecordSet.DoQuery(sQry);
							oDS_PS_PP777L.SetValue("U_MovDocNo", oRow - 1, oRecordSet.Fields.Item("MovDocNo").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_PP070No", oRow - 1, oRecordSet.Fields.Item("PP070No").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_PP070NoL", oRow - 1, oRecordSet.Fields.Item("PP070NoL").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_PorNum", oRow - 1, oRecordSet.Fields.Item("PorNum").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_ItemCode", oRow - 1, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_ItemName", oRow - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_Size", oRow - 1, oRecordSet.Fields.Item("Size").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_PkQty", oRow - 1, oRecordSet.Fields.Item("PkQty").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_PkWt", oRow - 1, oRecordSet.Fields.Item("PkWt").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_OPkQty", oRow - 1, oRecordSet.Fields.Item("OPkQty").Value.ToString().Trim());
							oDS_PS_PP777L.SetValue("U_OPkWt", oRow - 1, oRecordSet.Fields.Item("OPkWt").Value.ToString().Trim());
							oMat.SetLineData(oRow);

							if (oRow == oMat.RowCount && !string.IsNullOrEmpty(oDS_PS_PP777L.GetValue("U_PP070HL", oRow - 1).ToString().Trim()))
							{
								// 다음 라인 추가
								PS_PP777_AddMatrixRow(1, 0, false);
								oMat.Columns.Item("PP070HL").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}
							break;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP777_EnableFormItem
		/// </summary>
		private void PS_PP777_EnableFormItem()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = true;
					oForm.Items.Item("CntcCode").Enabled = false;
					oForm.Items.Item("DocDate").Enabled = false;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
		}

		/// <summary>
		/// PS_PP777_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP777_DelHeaderSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_PP777H.GetValue("U_BPLId", 0).ToString().Trim()))
                {
					errMessage = "사업장은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP777H.GetValue("U_CntcCode", 0).ToString().Trim()))
                {
					errMessage = "담당자는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP777H.GetValue("U_CntcName", 0).ToString().Trim()))
                {
					errMessage = "담장자명이 없습니다. 담당자코드를 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP777H.GetValue("U_DocDate", 0).ToString().Trim()))
                {
					errMessage = "작성일자는 필수사항입니다. 확인하여 주십시오.";
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
		/// PS_PP777_DelMatrixSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP777_DelMatrixSpaceLine()
		{
			bool returnValue = false;
			string errMessage = string.Empty;
			int i;

			try
			{
				oMat.FlushToDataSource();

                // 라인
                // MAT01에 값이 있는지 확인 (ErrorNumber : 1)
                if (oMat.VisualRowCount == 1)
                {
                    errMessage = "라인 데이터가 없습니다. 확인하여 주십시오.";
                    throw new Exception();
                }
                //마지막 행 하나를 빼고 i=0부터 시작하므로 하나를 빼므로
                //oMat.RowCount - 2가 된다..반드시 들어 가야 하는 필수값을 확인한다
                if (oMat.VisualRowCount > 0)
				{
					// Mat1에 입력값이 올바르게 들어갔는지 확인 (ErrorNumber : 2)
					for (i = 0; i <= oMat.VisualRowCount - 2; i++)
					{
						oDS_PS_PP777L.Offset = i;
						if (string.IsNullOrEmpty(oDS_PS_PP777L.GetValue("U_PP070HL", i).ToString().Trim()))
						{
							errMessage = "벌크포장문서 번호는 필수입니다. 확인하여 주십시오.";
							throw new Exception();
						}
					}
				}
				//맨마지막에 데이터를 삭제하는 이유는 행을 추가 할경우에 디비데이터소스에
				//이미 데이터가 들어가 있기 때문에 저장시에는 마지막 행(DB데이터 소스에)을 삭제한다
				if (oMat.VisualRowCount > 0)
				{
					oDS_PS_PP777L.RemoveRecord(oDS_PS_PP777L.Size - 1); // Mat1에 마지막라인(빈라인) 삭제
				}
				//행을 삭제하였으니 DB데이터 소스를 다시 가져온다
				oMat.LoadFromDataSource();
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
                    //Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    //Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
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
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_PP777_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_PP777_DelMatrixSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
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
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1282");
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							PS_PP777_EnableFormItem();
							PS_PP777_AddMatrixRow(1, oMat.RowCount, false);
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
						//헤더
						if (pVal.ItemUID == "CntcCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						//라인
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "PP070HL")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("PP070HL").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemChanged == true)
					{
						//헤더
						if (pVal.ItemUID == "CntcCode")
						{
							PS_PP777_FlushToItemValue(pVal.ItemUID, 0, "");
						}
						//라인
						if (pVal.ItemUID == "Mat01" && (pVal.ColUID == "PP070HL"))
						{
							PS_PP777_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP777H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP777L);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

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
							break;
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1285": //복원
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
							if (oMat.RowCount != oMat.VisualRowCount)
							{
								for (int i = 0; i <= oMat.VisualRowCount - 1; i++)
								{
									oMat.Columns.Item("LineId").Cells.Item(i + 1).Specific.Value = i + 1;
								}
								oMat.FlushToDataSource();
								oDS_PS_PP777L.RemoveRecord(oDS_PS_PP777L.Size - 1); // Mat1에 마지막라인(빈라인) 삭제
								oMat.Clear();
								oMat.LoadFromDataSource();
							}
							break;
						case "1281": //찾기
							PS_PP777_EnableFormItem();
							oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_PP777_EnableFormItem();
							PS_PP777_ClearForm();
							oDS_PS_PP777H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
							PS_PP777_AddMatrixRow(1, 0, true);
							oForm.Items.Item("BPLId").Enabled = true;
							oForm.Items.Item("BPLId").Specific.Select(3, SAPbouiCOM.BoSearchKey.psk_Index);

							oForm.Items.Item("OrdGbn").Enabled = true;
							oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

							oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

							oForm.Items.Item("BPLId").Enabled = false;
							oForm.Items.Item("OrdGbn").Enabled = false;
							oDS_PS_PP777H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD());
							PS_PP777_FlushToItemValue("CntcCode", 0, "");
							break;
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							PS_PP777_EnableFormItem();
							if (oMat.VisualRowCount > 0)
							{
								if (!string.IsNullOrEmpty(oMat.Columns.Item("PP070No").Cells.Item(oMat.VisualRowCount).Specific.Value.ToString().Trim()))
								{
									if (oDS_PS_PP777H.GetValue("Status", 0).ToString().Trim() == "O")
									{
										PS_PP777_AddMatrixRow(1, oMat.RowCount, false);
									}
								}
							}
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
