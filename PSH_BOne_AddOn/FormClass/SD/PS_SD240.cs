using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 미납품현황(품목별)
	/// </summary>
	internal class PS_SD240 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_SD240L;//등록라인
		
		private string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		/// <summary>
		/// LoadForm
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD240.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD240_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD240");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;

				oForm.Freeze(true);
				PS_SD240_CreateItems();
				PS_SD240_SetComboBox();
				PS_SD240_Initialize();
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
		/// PS_SD240_CreateItems
		/// </summary>
		private void PS_SD240_CreateItems()
		{
			try
			{
				oForm.Freeze(true);

				oDS_PS_SD240L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");

				//매트릭스 초기화
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				//사업장
				oForm.DataSources.UserDataSources.Add("BPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("BPLId").Specific.DataBind.SetBound(true, "", "BPLId");

				//품목구분
				oForm.DataSources.UserDataSources.Add("ItemClass", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemClass").Specific.DataBind.SetBound(true, "", "ItemClass");

				//거래형태
				oForm.DataSources.UserDataSources.Add("TradeType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("TradeType").Specific.DataBind.SetBound(true, "", "TradeType");

				//납기일 시작
				oForm.DataSources.UserDataSources.Add("FrDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("FrDt").Specific.DataBind.SetBound(true, "", "FrDt");

				//납기일 종료
				oForm.DataSources.UserDataSources.Add("ToDt", SAPbouiCOM.BoDataType.dt_DATE);
				oForm.Items.Item("ToDt").Specific.DataBind.SetBound(true, "", "ToDt");
				//납기일 종료_E

				//거래처코드
				oForm.DataSources.UserDataSources.Add("CardCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("CardCode").Specific.DataBind.SetBound(true, "", "CardCode");

				//거래처명
				oForm.DataSources.UserDataSources.Add("CardName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("CardName").Specific.DataBind.SetBound(true, "", "CardName");

				//품목코드(작번)
				oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
				oForm.Items.Item("ItemCode").Specific.DataBind.SetBound(true, "", "ItemCode");

				//품목명
				oForm.DataSources.UserDataSources.Add("ItemName", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 200);
				oForm.Items.Item("ItemName").Specific.DataBind.SetBound(true, "", "ItemName");

				//수주문서상태
				oForm.DataSources.UserDataSources.Add("DocStatus", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("DocStatus").Specific.DataBind.SetBound(true, "", "DocStatus");

				//미출고
				oForm.DataSources.UserDataSources.Add("Chk01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Chk01").Specific.DataBind.SetBound(true, "", "Chk01");

				//미납품
				oForm.DataSources.UserDataSources.Add("Chk02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("Chk02").Specific.DataBind.SetBound(true, "", "Chk02");
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
		/// PS_SD240_SetComboBox
		/// </summary>
		private void PS_SD240_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				//사업장
				oForm.Items.Item("BPLId").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL] ORDER BY BPLId", "", false, false);
				oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//품목구분
				oForm.Items.Item("ItemClass").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("ItemClass").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'S002' ORDER BY Code", "", false, false);
				oForm.Items.Item("ItemClass").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//거래형태
				oForm.Items.Item("TradeType").Specific.ValidValues.Add("%", "전체");
				dataHelpClass.Set_ComboList(oForm.Items.Item("TradeType").Specific, "SELECT  U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'T001' ORDER BY U_Minor", "", false, false);
				oForm.Items.Item("TradeType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//수주문서상태 세팅
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("%", "전체");
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("O", "미결");
				oForm.Items.Item("DocStatus").Specific.ValidValues.Add("C", "종료");
				oForm.Items.Item("DocStatus").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				//거래형태 콤보박스
				sQry = "SELECT  U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'T001' ORDER BY U_Minor";
				oRecordSet.DoQuery(sQry);

				while (!oRecordSet.EoF)
				{
					oMat.Columns.Item("TradeType").ValidValues.Add(oRecordSet.Fields.Item(0).Value, oRecordSet.Fields.Item(1).Value);
					oRecordSet.MoveNext();
				}

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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// Initialize
		/// </summary>
		private void PS_SD240_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				//사업장 사용자의 소속 사업장 선택
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);  // 사업장

				//체크박스 설정
				oForm.Items.Item("Chk01").Specific.Checked = false;
				oForm.Items.Item("Chk02").Specific.Checked = true;

				//날짜 설정
				oForm.Items.Item("ToDt").Specific.Value = "";
				oForm.Items.Item("FrDt").Specific.Value = "";
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
		/// PS_SD240_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_SD240_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_SD240L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_SD240L.Offset = oRow;
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
		/// PS_SD240_MTX01
		/// </summary>
		private void PS_SD240_MTX01()
		{
			int loopCount;
			int ErrNum = 0;
			string sQry;
			string BPLID;           //사업장
			string ItemClass;       //품목구분
			string TradeType;       //거래형태
			string FrDt;            //납기일시작
			string ToDt;            //납기일종료
			string CardCode;        //거래처
			string ItemCode;        //품목코드(작번)
			string DocStatus;       //문서상태
			string Chk01;           //미출고
			string Chk02;           //미납품

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				BPLID = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
				TradeType = oForm.Items.Item("TradeType").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				DocStatus = oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim();

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
				if (ItemClass == "%")
				{
					ItemClass = "";
				}
				if (TradeType == "%")
				{
					TradeType = "";
				}
				if (DocStatus == "%")
				{
					DocStatus = "";
				}

				oForm.Freeze(true);

				sQry = "EXEC PS_SD240_01 '" + BPLID + "','" + ItemClass + "','" + TradeType + "','" + FrDt + "','" + ToDt + "','" + CardCode + "','" + ItemCode + "','" + DocStatus + "','" + Chk01 + "','" + Chk02 + "'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oMat.FlushToDataSource();
				oMat.LoadFromDataSource();

				if (oRecordSet.RecordCount == 0)
				{
					oMat.Clear();
					ErrNum = 1;
					throw new Exception();
				}

				for (loopCount = 0; loopCount <= oRecordSet.RecordCount - 1; loopCount++)
				{
					if (loopCount != 0)
					{
						oDS_PS_SD240L.InsertRecord(loopCount);
					}
					oDS_PS_SD240L.Offset = loopCount;

					oDS_PS_SD240L.SetValue("U_LineNum", loopCount, Convert.ToString(loopCount + 1));                                //라인번호
					oDS_PS_SD240L.SetValue("U_ColReg01", loopCount, oRecordSet.Fields.Item("SO_No").Value.ToString().Trim());       //오더번호
					oDS_PS_SD240L.SetValue("U_ColReg02", loopCount, oRecordSet.Fields.Item("LotNo").Value.ToString().Trim());       //주문번호
					oDS_PS_SD240L.SetValue("U_ColReg03", loopCount, oRecordSet.Fields.Item("TradeType").Value.ToString().Trim());   //거래형태
					oDS_PS_SD240L.SetValue("U_ColReg04", loopCount, oRecordSet.Fields.Item("CardCode").Value.ToString().Trim());    //거래처코드
					oDS_PS_SD240L.SetValue("U_ColReg05", loopCount, oRecordSet.Fields.Item("CardName").Value.ToString().Trim());    //거래처명
					oDS_PS_SD240L.SetValue("U_ColReg06", loopCount, oRecordSet.Fields.Item("ItemCode").Value.ToString().Trim());    //품목코드(작번)
					oDS_PS_SD240L.SetValue("U_ColReg07", loopCount, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());    //품명
					oDS_PS_SD240L.SetValue("U_ColReg08", loopCount, oRecordSet.Fields.Item("Spec").Value.ToString().Trim());        //규격
					oDS_PS_SD240L.SetValue("U_ColReg11", loopCount, oRecordSet.Fields.Item("Unit").Value.ToString().Trim());        //단위
					oDS_PS_SD240L.SetValue("U_ColDt01", loopCount, Convert.ToDateTime(oRecordSet.Fields.Item("DocDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //수주일
					oDS_PS_SD240L.SetValue("U_ColDt02", loopCount, Convert.ToDateTime(oRecordSet.Fields.Item("DueDate").Value.ToString().Trim()).ToString("yyyyMMdd")); //납기일
					oDS_PS_SD240L.SetValue("U_ColQty01", loopCount, oRecordSet.Fields.Item("UP_Qty").Value.ToString().Trim());      //미납수량
					oDS_PS_SD240L.SetValue("U_ColSum01", loopCount, oRecordSet.Fields.Item("UP_Amt").Value.ToString().Trim());      //미납금액
					oDS_PS_SD240L.SetValue("U_ColQty02", loopCount, oRecordSet.Fields.Item("SO_Qty").Value.ToString().Trim());      //수주수량
					oDS_PS_SD240L.SetValue("U_ColSum02", loopCount, oRecordSet.Fields.Item("SO_Amt").Value.ToString().Trim());      //수주금액
					oDS_PS_SD240L.SetValue("U_ColReg09", loopCount, oRecordSet.Fields.Item("Req_No").Value.ToString().Trim());      //생산의뢰번호
					oDS_PS_SD240L.SetValue("U_ColQty03", loopCount, oRecordSet.Fields.Item("Req_Qty").Value.ToString().Trim());     //생산의뢰수량
					oDS_PS_SD240L.SetValue("U_ColQty04", loopCount, oRecordSet.Fields.Item("Deli_Qty").Value.ToString().Trim());    //출고수량
					oDS_PS_SD240L.SetValue("U_ColSum03", loopCount, oRecordSet.Fields.Item("Deli_Amt").Value.ToString().Trim());    //출고금액
					oDS_PS_SD240L.SetValue("U_ColQty05", loopCount, oRecordSet.Fields.Item("AR_Qty").Value.ToString().Trim());      //납품수량
					oDS_PS_SD240L.SetValue("U_ColSum04", loopCount, oRecordSet.Fields.Item("AR_Amt").Value.ToString().Trim());      //납품금액
					oDS_PS_SD240L.SetValue("U_ColReg12", loopCount, oRecordSet.Fields.Item("EtcDate").Value.ToString().Trim());     //기타출고최종일자
					oDS_PS_SD240L.SetValue("U_ColReg13", loopCount, oRecordSet.Fields.Item("EtcQty").Value.ToString().Trim());      //기타출고수량
					oDS_PS_SD240L.SetValue("U_ColReg10", loopCount, oRecordSet.Fields.Item("DocStatus").Value.ToString().Trim());   //문서상태

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
			}
			catch (Exception ex)
			{
				if (ErrNum == 1)
				{
					PSH_Globals.SBO_Application.SetStatusBarMessage("결과가 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
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
		/// PS_SD240_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_SD240_Print_Report01()
		{
			string WinTitle;
			string ReportName;
			string BPLId;           //사업장
			string ItemClass;       //품목구분
			string TradeType;       //거래형태
			string FrDt;            //납기일시작
			string ToDt;            //납기일종료
			string CardCode;        //거래처
			string ItemCode;        //품목코드(작번)
			string DocStatus;       //문서상태
			string Chk01;           //미출고
			string Chk02;           //미납품

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
				TradeType = oForm.Items.Item("TradeType").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim() == "" ? "19000101" : oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim() == "" ? "99991231" : oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				DocStatus = oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim();

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
				if (ItemClass == "%")
				{
					ItemClass = "";
				}
				if (TradeType == "%")
				{
					TradeType = "";
				}
				if (DocStatus == "%")
				{
					DocStatus = "";
				}

				WinTitle = "[PS_SD240] 레포트";
				ReportName = "PS_SD240.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemClass", ItemClass));
				dataPackParameter.Add(new PSH_DataPackClass("@TradeType", TradeType));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@DocStatus", DocStatus));
				dataPackParameter.Add(new PSH_DataPackClass("@Chk01", Chk01));
				dataPackParameter.Add(new PSH_DataPackClass("@Chk02", Chk02));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
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
		/// PS_SD240_Print_Report02
		/// </summary>
		[STAThread]
		private void PS_SD240_Print_Report02()
		{
			string WinTitle;
			string ReportName;
			string BPLId;           //사업장
			string ItemClass;       //품목구분
			string TradeType;       //거래형태
			string FrDt;            //납기일시작
			string ToDt;            //납기일종료
			string CardCode;        //거래처
			string ItemCode;        //품목코드(작번)
			string DocStatus;       //문서상태
			string Chk01;           //미출고
			string Chk02;           //미납품

			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				BPLId = oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim();
				ItemClass = oForm.Items.Item("ItemClass").Specific.Selected.Value.ToString().Trim();
				TradeType = oForm.Items.Item("TradeType").Specific.Selected.Value.ToString().Trim();
				FrDt = oForm.Items.Item("FrDt").Specific.Value.ToString().Trim() == "" ? "19000101" : oForm.Items.Item("FrDt").Specific.Value.ToString().Trim();
				ToDt = oForm.Items.Item("ToDt").Specific.Value.ToString().Trim() == "" ? "99991231" : oForm.Items.Item("ToDt").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
				DocStatus = oForm.Items.Item("DocStatus").Specific.Selected.Value.ToString().Trim();

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
				if (ItemClass == "%")
				{
					ItemClass = "";
				}
				if (TradeType == "%")
				{
					TradeType = "";
				}
				if (DocStatus == "%")
				{
					DocStatus = "";
				}

				WinTitle = "[PS_SD240] 미납현황집계표";
				ReportName = "PS_SD240_02.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@BPLId", BPLId));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemClass", ItemClass));
				dataPackParameter.Add(new PSH_DataPackClass("@TradeType", TradeType));
				dataPackParameter.Add(new PSH_DataPackClass("@FrDt", DateTime.ParseExact(FrDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@ToDt", DateTime.ParseExact(ToDt, "yyyyMMdd", null)));
				dataPackParameter.Add(new PSH_DataPackClass("@CardCode", CardCode));
				dataPackParameter.Add(new PSH_DataPackClass("@ItemCode", ItemCode));
				dataPackParameter.Add(new PSH_DataPackClass("@DocStatus", DocStatus));
				dataPackParameter.Add(new PSH_DataPackClass("@Chk01", Chk01));
				dataPackParameter.Add(new PSH_DataPackClass("@Chk02", Chk02));

				formHelpClass.CrystalReportOpen(WinTitle, ReportName, dataPackParameter);
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
							PS_SD240_MTX01();
						}
					}
					else if (pVal.ItemUID == "Btn_Print")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_SD240_Print_Report01);
							thread.SetApartmentState(System.Threading.ApartmentState.STA);
							thread.Start();
						}
					}
					else if (pVal.ItemUID == "Btn_Print2")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							System.Threading.Thread thread = new System.Threading.Thread(PS_SD240_Print_Report02);
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
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", ""); //거래처코드 포맷서치 활성
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", ""); //품목코드(작번) 포맷서치 활성
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
							sQry = "SELECT CardName, CardCode FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							sQry = "SELECT FrgnName, ItemCode FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("ItemName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "CntcCode")
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
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					SubMain.Remove_Forms(oFormUniqueID);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD240L);
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
							PS_SD240_AddMatrixRow(oMat.VisualRowCount, false); //엑셀 내보내기 실행 시 매트릭스의 제일 마지막 행에 빈 행 추가
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
							oDS_PS_SD240L.RemoveRecord(oDS_PS_SD240L.Size - 1);
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
