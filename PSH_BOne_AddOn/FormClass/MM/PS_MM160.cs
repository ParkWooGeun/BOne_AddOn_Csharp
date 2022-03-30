using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 외주가공비집계 및 청구자동등록
	/// </summary>
	internal class PS_MM160 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_MM160L; //등록라인
		private SAPbouiCOM.BoFormMode oForm_Mode;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM160.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM160_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM160");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

				oForm.Freeze(true);
				PS_MM160_CreateItems();
				PS_MM160_ComboBox_Setting();
				PS_MM160_Initialization();
				PS_MM160_LoadCaption();

				oForm.EnableMenu("1283", false); // 삭제
				oForm.EnableMenu("1286", false); // 닫기
				oForm.EnableMenu("1287", false); // 복제
				oForm.EnableMenu("1284", false); // 취소
				oForm.EnableMenu("1293", false); // 행삭제
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
		/// PS_MM160_CreateItems
		/// </summary>
		private void PS_MM160_CreateItems()
		{
			try
			{
				oDS_PS_MM160L = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.AutoResizeColumns();

				oForm.DataSources.UserDataSources.Add("DocDate", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DocDate").Specific.DataBind.SetBound(true, "", "DocDate");
				oForm.DataSources.UserDataSources.Item("DocDate").Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("DateFr", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DateFr").Specific.DataBind.SetBound(true, "", "DateFr");
				oForm.DataSources.UserDataSources.Item("DateFr").Value = DateTime.Now.ToString("yyyyMM") + "01";

				oForm.DataSources.UserDataSources.Add("DateTo", SAPbouiCOM.BoDataType.dt_DATE, 8);
				oForm.Items.Item("DateTo").Specific.DataBind.SetBound(true, "", "DateTo");
				oForm.DataSources.UserDataSources.Item("DateTo").Value = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("DocTotal", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("DocTotal").Specific.DataBind.SetBound(true, "", "DocTotal");

				oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
				oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

				oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
				oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM160_ComboBox_Setting
		/// </summary>
		private void PS_MM160_ComboBox_Setting()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "SELECT BPLId, BPLName From [OBPL] order by BPLId";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}

				//대분류
				sQry = "SELECT Code, Name From [@PSH_ITMBSORT] Order by Code";
				oRecordSet.DoQuery(sQry);
				while (!(oRecordSet.EoF))
				{
					oForm.Items.Item("ItmBSort").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("ItmBSort").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_MM160_Initialization
		/// </summary>
		private void PS_MM160_Initialization()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM160_LoadCaption
		/// </summary>
		private void PS_MM160_LoadCaption()
		{
			try
			{
				if (oForm_Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("Btn01").Specific.Caption = "추가";
				}
				else if (oForm_Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.Items.Item("Btn01").Specific.Caption = "확인";
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM160_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		private void PS_MM160_FlushToItemValue(string oUID)
		{
			string sQry;
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
					case "CardCode":
						sQry = "Select CardName From OCRD Where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
		}

		/// <summary>
		/// PS_MM160_HeaderSpaceLineDel
		/// </summary>
		/// <returns></returns>
		private bool PS_MM160_HeaderSpaceLineDel()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{ 
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
                {
					errMessage = "사업장은 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
                {
					errMessage = "외주거래처는 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
				{
					errMessage = "요청일은 필수입력 사항입니다. 확인하세요.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
				{
					errMessage = "작성자는 필수입력 사항입니다. 확인하세요.";
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
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_MM160_LoadData
		/// </summary>
		private void PS_MM160_LoadData()
		{
			int i;
			string DateFr;
			string CntcCode;
			string BPLId;
			string CardCode;
			string DocDate;
			string DateTo;
			decimal DocTotal = 0;
			decimal SumWeight = 0;
			int SumQty = 0;
			string ItmBsort;
			string errMessage = string.Empty;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbouiCOM.ProgressBar ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

			try
			{
				oForm.Freeze(true);
				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				DateFr = oForm.Items.Item("DateFr").Specific.Value.ToString().Trim();
				DateTo = oForm.Items.Item("DateTo").Specific.Value.ToString().Trim();
				ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();

				if (string.IsNullOrEmpty(CardCode))
				{
					CardCode = "%";
				}
				if (string.IsNullOrEmpty(CntcCode))
				{
					CntcCode = "%";
				}
				if (string.IsNullOrEmpty(DateFr))
				{
					DateFr = "18990101";
				}
				if (string.IsNullOrEmpty(DateTo))
				{
					DateTo = "20991231";
				}

				sQry = "EXEC [PS_MM160_01] '" + BPLId + "', '" + CardCode + "', '" + CntcCode + "', '" + DocDate + "', '" + DateFr + "', '" + DateTo + "', '" + ItmBsort + "', '1'";
				oRecordSet.DoQuery(sQry);

				oMat.Clear();
				oDS_PS_MM160L.Clear();

				if (oRecordSet.RecordCount == 0)
				{
					oForm_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
					errMessage = "조회 결과가 없습니다. 확인하세요.";
					throw new Exception();
				}

				ProgressBar01.Text = "조회시작!";

				for (i = 0; i <= oRecordSet.RecordCount - 1; i++)
				{
					if (i + 1 > oDS_PS_MM160L.Size)
					{
						oDS_PS_MM160L.InsertRecord(i);
					}

					oMat.AddRow();
					oDS_PS_MM160L.Offset = i;
					oDS_PS_MM160L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
					oDS_PS_MM160L.SetValue("U_ColReg01", i, oRecordSet.Fields.Item("U_ItemCode").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColReg02", i, oRecordSet.Fields.Item("U_ItemName").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColNum01", i, oRecordSet.Fields.Item("U_WorkQty").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColQty01", i, oRecordSet.Fields.Item("U_WorkWt").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColSum02", i, oRecordSet.Fields.Item("Price").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColSum01", i, oRecordSet.Fields.Item("U_Cprice").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColReg03", i, oRecordSet.Fields.Item("U_CardCode").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColReg04", i, oRecordSet.Fields.Item("U_CardName").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColReg05", i, oRecordSet.Fields.Item("CpCode").Value.ToString().Trim());
					oDS_PS_MM160L.SetValue("U_ColReg06", i, oRecordSet.Fields.Item("CpName").Value.ToString().Trim());

					oRecordSet.MoveNext();
					ProgressBar01.Value += 1;
					ProgressBar01.Text = ProgressBar01.Value + "/" + oRecordSet.RecordCount + "건 조회중...!";
				}
				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();

				for (i = 0; i <= oMat.VisualRowCount - 1; i++)
                {
                    DocTotal += Convert.ToDecimal(oMat.Columns.Item("WorkP").Cells.Item(i + 1).Specific.Value.ToString().Trim());

                    if (string.IsNullOrEmpty(oMat.Columns.Item("WorkQty").Cells.Item(i + 1).Specific.Value.ToString().Trim()))
					{
					}
					else
					{
						SumQty += Convert.ToInt32(oMat.Columns.Item("WorkQty").Cells.Item(i + 1).Specific.Value.ToString().Trim());
					}
					SumWeight += Convert.ToDecimal(oMat.Columns.Item("WorkWt").Cells.Item(i + 1).Specific.Value.ToString().Trim());
				}
				oForm.Items.Item("DocTotal").Specific.Value = DocTotal;
				oForm.Items.Item("SumQty").Specific.Value = SumQty;
				oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
				oForm_Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
			}
			catch (Exception ex)
			{
				if (ProgressBar01 != null)
				{
					ProgressBar01.Stop();
				}
				if (errMessage != string.Empty)
				{
					PSH_Globals.SBO_Application.MessageBox(errMessage);
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				ProgressBar01.Stop();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_MM160_Add_MM005
		/// </summary>
		/// <param name="pVal"></param>
		/// <returns></returns>
		private bool PS_MM160_Add_MM005(ref SAPbouiCOM.ItemEvent pVal)
		{
			bool functionReturnValue = false;

			int i;
			string ItemCode;
			string CardName;
			string CntcName;
			string DocEntry;
			string BPLId;
			string CntcCode;
			string CardCode = string.Empty;
			string DocDate;
			string ItemName;
			int WorkQty;
			decimal WorkWt;
			decimal Price;
			decimal CPrice;
			string DateTo;
			string DateFr;
			string ItmBsort;
			string CpCode;
			string CpName;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oMat.FlushToDataSource();

				BPLId = oForm.Items.Item("BPLId").Specific.Value.ToString().Trim();
				DocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				CntcCode = oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim();
				CntcName = oForm.Items.Item("CntcName").Specific.Value.ToString().Trim();
				DateFr = oForm.Items.Item("DateFr").Specific.Value.ToString().Trim();
				DateTo = oForm.Items.Item("DateTo").Specific.Value.ToString().Trim();
				ItmBsort = oForm.Items.Item("ItmBSort").Specific.Value.ToString().Trim();

				PSH_Globals.oCompany.StartTransaction();

				for (i = 0; i <= oMat.RowCount - 1; i++)
				{
					sQry = "Select IsNull(Max(DocEntry), 0) From [@PS_MM005H] where Left(DocEntry, 6) = Left('" + DocDate + "', 6)";
					oRecordSet.DoQuery(sQry);

					if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) == 0)
					{
						DocEntry = DocDate.Substring(0, 6) + "0001";
					}
					else
					{
						DocEntry = Convert.ToString(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) + 1);
					}

					ItemCode = oDS_PS_MM160L.GetValue("U_ColReg01", i).ToString().Trim();
					ItemName = dataHelpClass.Make_ItemName(oDS_PS_MM160L.GetValue("U_ColReg02", i).ToString().Trim());
					if (string.IsNullOrEmpty(oDS_PS_MM160L.GetValue("U_ColNum01", i)))
					{
						WorkQty = 0;
					}
					else
					{
						WorkQty = Convert.ToInt32(oDS_PS_MM160L.GetValue("U_ColNum01", i));
					}
					WorkWt = Convert.ToDecimal(oDS_PS_MM160L.GetValue("U_ColQty01", i).ToString().Trim());
					CPrice = Convert.ToDecimal(oDS_PS_MM160L.GetValue("U_ColSum01", i).ToString().Trim());
					CardCode = oDS_PS_MM160L.GetValue("U_ColReg03", i).ToString().Trim();
					CardName = oDS_PS_MM160L.GetValue("U_ColReg04", i).ToString().Trim();
					CpCode = oDS_PS_MM160L.GetValue("U_ColReg05", i).ToString().Trim();
					CpName = oDS_PS_MM160L.GetValue("U_ColReg06", i).ToString().Trim();

					if (WorkWt == 0)
					{
						Price = 0;
					}
					else
					{
						Price = Math.Round(CPrice / WorkWt, 2);
					}

					sQry = "INSERT INTO [@PS_MM005H]";
					sQry += " (";
					sQry += " DocEntry,";
					sQry += " DocNum,";
					sQry += " UserSign,";
					sQry += " U_ItemCode,";
					sQry += " U_ItemName,";
					sQry += " U_Weight,";
					sQry += " U_BPLId,";
					sQry += " U_CgNum,";
					sQry += " U_DocDate,";
					sQry += " U_DueDate,";
					sQry += " U_CntcCode,";
					sQry += " U_CntcName,";
					sQry += " U_Auto,";
					sQry += " U_QCYN,";
					sQry += " U_OKYN,";
					sQry += " U_OrdType,";
					sQry += " U_ProcCode,";
					sQry += " U_ProcName,";
					sQry += " U_CardCode,";
					sQry += " U_CardName,";
					sQry += " U_Status";
					sQry += " ) ";
					sQry += "VALUES(";
					sQry += DocEntry + ",";
					sQry += DocEntry + ",";
					sQry += "'1',";
					sQry += "'" + ItemCode + "',";
					sQry += "'" + ItemName + "',";
					sQry += "'" + WorkQty + "',";
					sQry += "'" + BPLId + "',";
					sQry += "'" + DocEntry + "',";
					sQry += "'" + DocDate + "',";
					sQry += "'" + DocDate + "',";
					sQry += "'" + CntcCode + "',";
					sQry += "'" + CntcName + "',";
					sQry += "'N',";
					sQry += "'N',";
					sQry += "'Y',";
					sQry += "'30',";
					sQry += "'" + CpCode + "',";
					sQry += "'" + CpName + "',";
					sQry += "'" + CardCode + "',";
					sQry += "'" + CardName + "',";
					sQry += "'O'";
					sQry += ")";
					oRecordSet.DoQuery(sQry);
				}

				if (string.IsNullOrEmpty(CardCode))
                {
					CardCode = "%";
				}
				if (string.IsNullOrEmpty(CntcCode))
                {
					CntcCode = "%";
				}
				if (string.IsNullOrEmpty(DateFr))
                {
					DateFr = "18990101";
				}
				if (string.IsNullOrEmpty(DateTo))
                {
					DateTo = "20991231";
				}

				sQry = "EXEC [PS_MM160_01] '" + BPLId + "', '" + CardCode + "', '" + CntcCode + "', '" + DocDate + "', '" + DateFr + "', '" + DateTo + "', '" + ItmBsort + "', '2'";
				oRecordSet.DoQuery(sQry);

				PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				PSH_Globals.SBO_Application.StatusBar.SetText("청구 자동 등록 완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
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
					Raise_EVENT_KEY_DOWN(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS: //3
				//    Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
					Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//    Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
				//    Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
					Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
				//    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_DATASOURCE_LOAD: //12
				//    Raise_EVENT_DATASOURCE_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_FORM_LOAD: //16
				//    Raise_EVENT_FORM_LOAD(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD: //17
					Raise_EVENT_FORM_UNLOAD(FormUID, ref pVal, ref BubbleEvent);
					break;
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
				//    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
						if (oForm_Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_MM160_HeaderSpaceLineDel() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_MM160_Add_MM005(ref pVal) == false)
							{
								BubbleEvent = false;
								return;
							}

							oForm_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							oMat.Clear();
							oDS_PS_MM160L.Clear();
							PS_MM160_LoadCaption();
						}
						else if (oForm_Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							oForm.Close();
						}
					}
					else if (pVal.ItemUID == "Btn02")
					{
						if (PS_MM160_HeaderSpaceLineDel() == false)
						{
							BubbleEvent = false;
							return;
						}
						PS_MM160_LoadData();
						PS_MM160_LoadCaption();
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
						else if (pVal.ItemUID == "CntcCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
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
		/// COMBO_SELECT 이벤트
		/// </summary>
		/// <param name="FormUID">Form UID</param>
		/// <param name="pVal">ItemEvent 객체</param>
		/// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "BPLId")
					{
						oMat.Clear();
						oDS_PS_MM160L.Clear();
						oForm_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
						PS_MM160_LoadCaption();
					}
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
						if (pVal.ItemUID == "CardCode" || pVal.ItemUID == "DateFr" || pVal.ItemUID == "DateTo")
						{
							oMat.Clear();
							oDS_PS_MM160L.Clear();
							oForm_Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PS_MM160_LoadCaption();
							PS_MM160_FlushToItemValue(pVal.ItemUID);
						}
						else if (pVal.ItemUID == "CntcCode")
						{
							PS_MM160_FlushToItemValue(pVal.ItemUID);
						}
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM160L);
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
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
