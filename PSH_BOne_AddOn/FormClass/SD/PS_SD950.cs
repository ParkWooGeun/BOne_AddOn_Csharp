using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 채권만기일관리
	/// </summary>
	internal class PS_SD950 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_SD950H;//등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD950L;//등록라인
		
		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD950.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD950_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD950");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);

				PS_SD950_CreateItems();
				PS_SD950_SetComboBox();
				PS_SD950_Initialize();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				oForm.Update();
				oForm.Visible = true;
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_SD950_CreateItems
		/// </summary>
		private void PS_SD950_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oDS_PS_SD950H = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");
				oDS_PS_SD950L = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

				//매트릭스 초기화
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//만기일 시작
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//만기일 종료
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//거래처코드
				oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

				//거래처명
				oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

				//AR문서상태
				oForm.DataSources.UserDataSources.Add("DocStatus", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("DocStatus").Specific.DataBind.SetBound(true, "", "DocStatus");
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
		/// PS_SD950_SetComboBox
		/// </summary>
		private void PS_SD950_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				//사업장 콤보박스 세팅
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//수주문서상태
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("O", "미결");
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("C", "종료");
				oForm.Items.Item("DocStatus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//문서상태
				oMat.Columns.Item("DocStatus").ValidValues.Add("O", "미결");
				oMat.Columns.Item("DocStatus").ValidValues.Add("C", "종료");
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
		/// PS_SD950_Initialize
		/// </summary>
		private void PS_SD950_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장 사용자의 소속 사업장 선택
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);

				//날짜 설정
				oForm.Items.Item("FrDt").Specific.Value = DateTime.Now.ToString("yyyyMM") + "01";
				oForm.Items.Item("ToDt").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
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
		/// PS_SD950_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_SD950_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_SD950L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_SD950L.Offset = oRow;
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
		/// PS_SD950_MTX01
		/// </summary>
		private void PS_SD950_MTX01()
		{
			int loopCount;
			int ErrCode = 0;
			string sQry;

			string BPLID;           //사업장
			string FrDt;            //만기일시작
			string ToDt;            //만기일종료
			string CardCode;        //거래처
			string DocStatus;       //문서상태

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				DocStatus = oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim();

				if (DocStatus == "%")
				{
					DocStatus = "";
				}

				oForm.Freeze(true);

				sQry = "EXEC PS_SD950_01 '" + BPLID + "','" + FrDt + "','" + ToDt + "','" + CardCode + "','" + DocStatus + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					ErrCode = 1;
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_SD950L.InsertRecord(loopCount);
					}
					oDS_PS_SD950L.Offset = loopCount;

					oDS_PS_SD950L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                                //라인번호
					oDS_PS_SD950L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("DocEntry").Value.ToString().Trim());    //AR송장번호
					oDS_PS_SD950L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("DocDate").Value.ToString().Trim());     //전기일
					oDS_PS_SD950L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("DueDate").Value.ToString().Trim());     //만기일
					oDS_PS_SD950L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());    //거래처코드
					oDS_PS_SD950L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());    //거래처명
					oDS_PS_SD950L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("Currency").Value.ToString().Trim());    //통화
					oDS_PS_SD950L.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("LineTotal").Value.ToString().Trim());   //금액
					oDS_PS_SD950L.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("VatSum").Value.ToString().Trim());      //부가세
					oDS_PS_SD950L.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("Total").Value.ToString().Trim());       //총계
					oDS_PS_SD950L.SetValue("U_ColPrc01", loopCount, oRecordSet.Fields.Item("TotalFC").Value.ToString().Trim());     //총계(외화)
					oDS_PS_SD950L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ReceiptsDt").Value.ToString().Trim());  //입금일자
					oDS_PS_SD950L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("DelayDay").Value.ToString().Trim());    //지연일수
					oDS_PS_SD950L.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("Receipts").Value.ToString().Trim());    //입금액
					oDS_PS_SD950L.SetValue("U_ColPrc02", loopCount, oRecordSet.Fields.Item("ReceiptsFC").Value.ToString().Trim());  //입금액(외화)
					oDS_PS_SD950L.SetValue("U_ColSum05", loopCount, oRecordSet.Fields.Item("AdjAmt").Value.ToString().Trim());      //조정금액
					oDS_PS_SD950L.SetValue("U_ColPrc03", loopCount, oRecordSet.Fields.Item("AdjAmtFC").Value.ToString().Trim());    //조정금액(외화)
					oDS_PS_SD950L.SetValue("U_ColSum06", loopCount, oRecordSet.Fields.Item("RecTotal").Value.ToString().Trim());    //회수금액총계
					oDS_PS_SD950L.SetValue("U_ColPrc04", loopCount, oRecordSet.Fields.Item("RecTotalFC").Value.ToString().Trim());  //회수금액총계(외화)
					oDS_PS_SD950L.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("PayMth").Value.ToString().Trim());      //지급수단
					oDS_PS_SD950L.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("BoeDueDate").Value.ToString().Trim());  //어음만기일
					oDS_PS_SD950L.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("DocStatus").Value.ToString().Trim());   //AR송장문서상태

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrCode == 1)
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("결과가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					if (pVal.ItemUID == "Btn01")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							PS_SD950_MTX01();	//매트릭스에 데이터 로드
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
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

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "CardCode")
						{
							sQry = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value.String().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.String().Trim();
						}

						oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
		/// Raise_EVENT_FORM_UNLOAD
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD950H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD950L);
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
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "7169": //엑셀 내보내기
							PS_SD950_AddMatrixRow(oMat.VisualRowCount, false); //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
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
							oDS_PS_SD950L.RemoveRecord(oDS_PS_SD950L.Size - 1);
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
			finally
			{
			}
		}
	}
}
