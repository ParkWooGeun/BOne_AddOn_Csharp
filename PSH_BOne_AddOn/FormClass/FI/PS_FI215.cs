using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;
using PSH_BOne_AddOn.Form;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 지급현황
	/// </summary>
	internal class PS_FI215 : PSH_BaseClass
	{
		private string oFormUniqueID01;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_FI215L;  //등록라인
		private string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			int i = 0;
			MSXML2.DOMDocument oXmlDoc01 = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc01.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_FI215.srf");
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc01.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc01.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc01.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID01 = "PS_FI215_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID01, "PS_FI215");                   // 폼추가
				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc01.xml.ToString()); // 폼할당
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID01);
				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_FI215_CreateItems();
				PS_FI215_ComboBox_Setting();
				PS_FI215_Initial_Setting();
				PS_FI215_SetDocument(oFormDocEntry);
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
		/// PS_FI215_CreateItems
		/// </summary>
		/// <returns></returns>
		private void PS_FI215_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oDS_PS_FI215L = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

				// 매트릭스 초기화
				oMat01 = oForm.Items.Item("Mat01").Specific;
				oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat01.AutoResizeColumns();

				//사업장_S
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");
				//사업장_E

				//만기일 시작_S
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//만기일 종료_S
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");

				//거래처코드_S
				oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

				//거래처명_S
				oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

				//AR문서상태_S
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
		/// PS_FI215_ComboBox_Setting
		/// </summary>
		private void PS_FI215_ComboBox_Setting()
		{
			SAPbouiCOM.Column oColumn = null;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				//사업장 콤보박스 세팅_S
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//수주문서상태 세팅_S
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("O", "미결");
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("C", "종료");
				oForm.Items.Item("DocStatus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//문서상태
				oColumn = oMat01.Columns.Item("DocStatus");
				oColumn.ValidValues.Add("O", "미결");
				oColumn.ValidValues.Add("C", "종료");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_FI215_Initial_Setting
		/// </summary>
		private void PS_FI215_Initial_Setting()
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
		/// PS_FI215_FormItemEnabled
		/// </summary>
		private void PS_FI215_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
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
		/// PS_FI215_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_FI215_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				//행추가여부
				if (RowIserted == false)
				{
					oDS_PS_FI215L.InsertRecord(oRow);
				}
				oMat01.AddRow();
				oDS_PS_FI215L.Offset = oRow;
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
		/// PS_FI215_FormClear
		/// </summary>
		private void PS_FI215_FormClear()
		{
			string DocEntry = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_FI215'", "");
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
			finally
			{
			}
		}

		/// <summary>
		/// PS_FI215_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_FI215_SetDocument(string oFormDocEntry)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_FI215_FormItemEnabled();
				}
				else
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
		/// PS_FI215_MTX01
		/// </summary>
		private void PS_FI215_MTX01()
		{
			//메트릭스에 데이터 로드
			int loopCount = 0;
			int ErrNum = 0;
			string sQry = string.Empty;

			string BPLID = string.Empty;            //사업장
			string FrDt = string.Empty;             //만기일시작
			string ToDt = string.Empty;             //만기일종료
			string CardCode = string.Empty;         //거래처
			string DocStatus = string.Empty;        //문서상태

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);

				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				DocStatus = oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim();

				if (DocStatus == "%")
				{
					DocStatus = "";
				}
				sQry = "EXEC PS_FI215_01 '" + BPLID + "','" + FrDt + "','" + ToDt + "','" + CardCode + "','" + DocStatus + "'";
				oRecordSet.DoQuery(sQry);

				oMat01.Clear();
				oMat01.FlushToDataSource();
				oMat01.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat01.Clear();
					ErrNum = 1;
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_FI215L.InsertRecord(loopCount);
					}
					oDS_PS_FI215L.Offset = loopCount;

					oDS_PS_FI215L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                            //라인번호
					oDS_PS_FI215L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("DocEntry").Value);                  //AR송장번호
					oDS_PS_FI215L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("DocDate").Value);                   //전기일
					oDS_PS_FI215L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("DueDate").Value);                   //만기일
					oDS_PS_FI215L.SetValue("U_ColReg12", loopCount, oRecordSet.Fields.Item("TaxDate").Value);                   //증빙일
					oDS_PS_FI215L.SetValue("U_ColReg13", loopCount, oRecordSet.Fields.Item("PayDate").Value);                   //지급예정일
					oDS_PS_FI215L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("CardCode").Value);                  //거래처코드
					oDS_PS_FI215L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("CardName").Value);                  //거래처명
					oDS_PS_FI215L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("Currency").Value);                  //통화
					oDS_PS_FI215L.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("LineTotal").Value);                 //금액
					oDS_PS_FI215L.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("VatSum").Value);                    //부가세
					oDS_PS_FI215L.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("Total").Value);                     //총계
					oDS_PS_FI215L.SetValue("U_ColPrc01", loopCount, oRecordSet.Fields.Item("TotalFC").Value);                   //총계(외화)
					oDS_PS_FI215L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ReceiptsDt").Value);                //입금일자
					oDS_PS_FI215L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("DelayDay").Value);                  //지연일수
					oDS_PS_FI215L.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("Receipts").Value);                  //입금액
					oDS_PS_FI215L.SetValue("U_ColPrc02", loopCount, oRecordSet.Fields.Item("ReceiptsFC").Value);                //입금액(외화)
					oDS_PS_FI215L.SetValue("U_ColSum05", loopCount, oRecordSet.Fields.Item("AdjAmt").Value);                    //조정금액
					oDS_PS_FI215L.SetValue("U_ColPrc03", loopCount, oRecordSet.Fields.Item("AdjAmtFC").Value);                  //조정금액(외화)
					oDS_PS_FI215L.SetValue("U_ColSum06", loopCount, oRecordSet.Fields.Item("RecTotal").Value);                  //회수금액총계
					oDS_PS_FI215L.SetValue("U_ColPrc04", loopCount, oRecordSet.Fields.Item("RecTotalFC").Value);                //회수금액총계(외화)
					oDS_PS_FI215L.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("PayMth").Value);                    //지급수단
					oDS_PS_FI215L.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("BoeDueDate").Value);                //어음만기일
					oDS_PS_FI215L.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("DocStatus").Value);                 //AR송장문서상태

					oRecordSet.MoveNext();
				}

				oMat01.LoadFromDataSource();
				oMat01.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
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
				oForm.Freeze(false);
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
					System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_FI215_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_FI215_Print_Report01()
		{
			// 안씀
			string WinTitle = string.Empty;
			string ReportName = string.Empty;
			string sQry = string.Empty;

			string BPLID = string.Empty;            //사업장
			string ItemClass = string.Empty;        //품목구분
			string TradeType = string.Empty;        //거래형태
			string FrDt = string.Empty;             //납기일시작
			string ToDt = string.Empty;             //납기일종료
			string CardCode = string.Empty;         //거래처
			string ItemCode = string.Empty;         //품목코드(작번)
			string DocStatus = string.Empty;        //문서상태
			string Chk01 = string.Empty;            //미출고
			string Chk02 = string.Empty;            //미납품

			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToStyring().Trim();
				ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToStyring().Trim();
				TradeType = oForm.Items.Item("TradeType").Specific.Selected.Value.ToStyring().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToStyring().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToStyring().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToStyring().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToStyring().Trim();
				DocStatus = oForm.Items.Item("DocStatus").Specific.Selected.Value.ToStyring().Trim();

				if (oForm.Items.Item("Chk01").Specific.Checked == true)
				{
					Chk01 = "1";
				}
				else
				{
					Chk01 = "0";
				}

				if (oForm.Items.Item("Chk02").Specific.Checked == true)
				{
					Chk02 = "1";
				}
				else
				{
					Chk02 = "0";
				}

				if (oForm.Items.Item("ItemClass").Specific.Selected.Value == "%")
				{
					ItemClass = "";

				}
				else
				{
					ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToStyring().Trim();
				}

				if (oForm.Items.Item("TradeType").Specific.Selected.Value == "%")
				{
					TradeType = "";

				}
				else
				{
					TradeType = oForm.Items.Item("TradeType").Specific.Selected.Value.ToStyring().Trim();
				}

				if (oForm.Items.Item("DocStatus").Specific.Selected.Value == "%")
				{
					DocStatus = "";

				}
				else
				{
					DocStatus = oForm.Items.Item("DocStatus").Specific.Selected.Value.ToStyring().Trim();
				}

				WinTitle = "[PS_FI215] 레포트";
				ReportName = "PS_FI215.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();
				List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass>();

				// Formula

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLID", BPLID));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemClass", ItemClass));
				dataPackParameter.Add(new PSH_DataPackClass("@TradeType", TradeType));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", FrDt));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", ToDt));
				dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@DocStatus", DocStatus));
				dataPackParameter.Add(new PSH_DataPackClass("@Chk01", Chk01));
				dataPackParameter.Add(new PSH_DataPackClass("@Chk02", Chk02));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
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
		/// Raise_FormItemEvent
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		public override void Raise_FormItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				switch (pVal.EventType)
				{
					case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:					//1
						Raise_EVENT_ITEM_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:						//2
						Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                      //3
						Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:                     //4
						break;
					case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:					//5
						//Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CLICK:						    //6
						Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:					//7
						Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:			//8
						//Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_VALIDATE:						//10
						Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:					//11
						Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:                    //17
						Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:					//18
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:				//19
						break;
					case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:					//20
						//Raise_EVENT_RESIZE(FormUID, pVal, BubbleEvent);
						break;
					case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:				//27
						//Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
						break;
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
							PS_FI215_MTX01();
						}
					}
					else if (pVal.ItemUID == "Btn_Print")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_FI215_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); // 거래처코드 포맷서치 활성
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", ""); // 품목코드(작번) 포맷서치 활성
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
							oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
							oMat01.FlushToDataSource();
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
			string sQry = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if ((pVal.ItemUID == "CardCode"))
						{
							sQry = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if ((pVal.ItemUID == "ItemCode"))
						{
							sQry = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if ((pVal.ItemUID == "CntcCode"))
						{
							sQry = "SELECT U_FULLNAME, U_MSTCOD FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
					PS_FI215_FormItemEnabled();
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
			finally
			{
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
					SubMain.Remove_Forms(oFormUniqueID01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_FI215L);
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
						case "1284":                            //취소
							break;
						case "1286":                            //닫기
							break;
						case "1293":                            //행삭제
							break;
						case "1281":                            //찾기
							break;
						case "1282":                            //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                            //레코드이동버튼
							break;
						case "7169":                            //엑셀 내보내기
							PS_FI215_AddMatrixRow(oMat01.VisualRowCount, false); //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1284":                            //취소
							break;
						case "1286":                            //닫기
							break;
						case "1293":                            //행삭제
							break;
						case "1281":                            //찾기
							break;
						case "1282":                            //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291":                            //레코드이동버튼
							break;
						case "7169":                            //엑셀 내보내기
							oForm.Freeze(true); //엑셀 내보내기 이후 처리
							oDS_PS_FI215L.RemoveRecord(oDS_PS_FI215L.Size - 1);
							oMat01.LoadFromDataSource();
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
				if ((BusinessObjectInfo.BeforeAction == true))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                     //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                      //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                   //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                   //36
							break;
					}
				}
				else if ((BusinessObjectInfo.BeforeAction == false))
				{
					switch (BusinessObjectInfo.EventType)
					{
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                     //33
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:                      //34
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                   //35
							break;
						case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:                   //36
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
			finally
			{
			}
		}
	}
}
