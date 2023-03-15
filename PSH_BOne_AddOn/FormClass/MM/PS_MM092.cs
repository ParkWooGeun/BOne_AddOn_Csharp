using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 분말샘플출고등록
	/// </summary>
	internal class PS_MM092 : PSH_BaseClass
	{
		public string oFormUniqueID;
		public SAPbouiCOM.Matrix oMat;

		private SAPbouiCOM.DBDataSource oDS_PS_MM092H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM092L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private string oOutMan;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM092.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM092_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM092");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";
				////UDO방식일때

				oForm.Freeze(true);
				PS_MM092_CreateItems();
				PS_MM092_ComboBox_Setting();
				PS_MM092_EnableMenus();
				PS_MM092_SetDocument(oFormDocEntry);
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
				oForm.Items.Item("CardCode").Click();
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc); //메모리 해제
			}
		}

		/// <summary>
		/// PS_MM092_CreateItems
		/// </summary>
		private void PS_MM092_CreateItems()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oDS_PS_MM092H = oForm.DataSources.DBDataSources.Item("@PS_MM092H");
				oDS_PS_MM092L = oForm.DataSources.DBDataSources.Item("@PS_MM092L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + dataHelpClass.User_MSTCOD() + "'";
				oRecordSet.DoQuery(sQry);
				oForm.Items.Item("OutMan").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
				oOutMan = oRecordSet.Fields.Item(0).Value.ToString().Trim();

				oForm.Items.Item("InDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
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
		/// PS_MM092_ComboBox_Setting
		/// </summary>
		private void PS_MM092_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Combo_ValidValues_Insert("PS_MM092", "Title", "", "01", "송장");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM092", "Title", "", "02", "거래명세서");
				dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("Title").Specific, "PS_MM092", "Title", false);
				oForm.Items.Item("Title").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM [OBPL]  order by BPLId", "1", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM092_EnableMenus
		/// </summary>
		private void PS_MM092_EnableMenus()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM092_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_MM092_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_MM092_FormItemEnabled();
					PS_MM092_AddMatrixRow(0, true);
				}
				else
				{
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM092_FormItemEnabled
		/// </summary>
		private void PS_MM092_FormItemEnabled()
		{
			string User_BPLId;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				User_BPLId = dataHelpClass.User_BPLID();

				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("BPLId").Specific.Select(User_BPLId, SAPbouiCOM.BoSearchKey.psk_ByValue);
					oForm.Items.Item("Title").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
					oForm.Items.Item("OutMan").Specific.Value = oOutMan;
					oForm.Items.Item("InDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");

					PS_MM092_FormClear();

					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가

					oForm.Items.Item("Btn1").Enabled = false;
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("Mat01").Enabled = true;
					oForm.Items.Item("InDate").Enabled = true;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("OutMan").Enabled = true;
					oForm.Items.Item("RecMan").Enabled = true;
					oForm.Items.Item("PurPose").Enabled = true;
					oForm.Items.Item("Destin").Enabled = true;
					oForm.Items.Item("TranCard").Enabled = true;
					oForm.Items.Item("TranCode").Enabled = true;
					oForm.Items.Item("TranCost").Enabled = true;
					oForm.Items.Item("Title").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가

					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("InDate").Enabled = true;
					oForm.Items.Item("OutMan").Enabled = true;
					oForm.Items.Item("TranCard").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.EnableMenu("1282", true); //추가

					if (oDS_PS_MM092H.GetValue("Canceled", 0).ToString().Trim() == "Y")
					{
						oForm.Items.Item("Btn1").Enabled = false; //출력버튼비활성
						oForm.Items.Item("DocEntry").Enabled = false;
						oForm.Items.Item("CardCode").Enabled = false;
						oForm.Items.Item("Mat01").Enabled = false;
						oForm.Items.Item("InDate").Enabled = false;
						oForm.Items.Item("BPLId").Enabled = false;
						oForm.Items.Item("OutMan").Enabled = false;
						oForm.Items.Item("RecMan").Enabled = false;
						oForm.Items.Item("PurPose").Enabled = false;
						oForm.Items.Item("Destin").Enabled = false;
						oForm.Items.Item("TranCard").Enabled = false;
						oForm.Items.Item("TranCode").Enabled = false;
						oForm.Items.Item("TranCost").Enabled = false;
						oForm.Items.Item("Title").Enabled = false;
					}
					else
					{
						oForm.Items.Item("Btn1").Enabled = true;
						oForm.Items.Item("DocEntry").Enabled = false;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("Mat01").Enabled = true;
						oForm.Items.Item("InDate").Enabled = false;
						oForm.Items.Item("BPLId").Enabled = false;
						oForm.Items.Item("OutMan").Enabled = true;
						oForm.Items.Item("RecMan").Enabled = true;
						oForm.Items.Item("PurPose").Enabled = true;
						oForm.Items.Item("Destin").Enabled = true;
						oForm.Items.Item("TranCard").Enabled = true;
						oForm.Items.Item("TranCode").Enabled = true;
						oForm.Items.Item("TranCost").Enabled = true;
						oForm.Items.Item("Title").Enabled = true;
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
		/// PS_MM092_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM092_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_MM092L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_MM092L.Offset = oRow;
				oDS_PS_MM092L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_MM092_FormClear
		/// </summary>
		private void PS_MM092_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM092'", "");
				if (Convert.ToDouble(DocEntry) == 0)
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
		/// 데이터 필수값등등 체크
		/// </summary>
		/// <returns></returns>
		private bool PS_MM092_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			string errMessage = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("InDate").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("InDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "반출일은 필수입니다.";
					throw new Exception();
				}
				// 마감일자 Check
				else if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("InDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
				{
					errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.";
					oForm.Items.Item("InDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "고객코드는 필수입니다.";
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oForm.Items.Item("OutMan").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("OutMan").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "반출자는 필수입니다.";
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oForm.Items.Item("PurPose").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("PurPose").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "목적은 필수입니다.";
					throw new Exception();
				}

				if (oMat.VisualRowCount == 0)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}
				else
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(1).Specific.Value.ToString().Trim()))
					{
						errMessage = "Matrix값이 한줄이상은 있어야합니다.";
						throw new Exception();
					}
				}

				for (i = 1; i <= (oMat.VisualRowCount - 1); i++)
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("ItemName").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("ItemName").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "품명은 필수입니다.";
						throw new Exception();
					}
				}

				oMat.FlushToDataSource();
				oDS_PS_MM092L.RemoveRecord(oDS_PS_MM092L.Size - 1);
				oMat.LoadFromDataSource();

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_MM092_FormClear();
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
		/// PS_MM092_Validate
		/// </summary>
		/// <param name="ValidateType"></param>
		/// <returns></returns>
		private bool PS_MM092_Validate(string ValidateType)
		{
			bool ReturnValue = false;
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
							if (oForm.Items.Item("Canceled").Specific.Value.ToString().Trim() == "Y")
							{
								errMessage = "취소된문서는 수정할수 없습니다.";
								throw new Exception();
							}
						}
					}
				}
				else if (ValidateType == "취소")
				{
					if (oForm.Items.Item("Canceled").Specific.Value.ToString().Trim() == "Y")
					{
						errMessage = "이미취소된문서입니다.";
						throw new Exception();
					}
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
		/// PS_SMM092_Print_Report01
		/// </summary>
		[STAThread]
		private void PS_SMM092_Print_Report01()
		{
			string DocNum;
			string WinTitle;
			string ReportName;
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocNum = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				WinTitle = "[PS_MM092] 기타자재출고증출력";
				ReportName = "PS_MM092_01.rpt";

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocNum));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// 샘플출고
		/// </summary>
		/// <returns></returns>
		private bool PS_MM092_Add_oInventoryGenExit()
		{
			bool ReturnValue = false;
			int i;
			int RetVal;
			int ResultDocNum;
			double Quantity;
			string ItemCode;
			string BatchNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Documents DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit); //문서타입(입고)

			try
			{
				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				PSH_Globals.oCompany.StartTransaction();
				oMat.FlushToDataSource();

				DI_oInventoryGenExit.DocDate = DateTime.ParseExact(oForm.Items.Item("InDate").Specific.Value, "yyyyMMdd", null);
				DI_oInventoryGenExit.UserFields.Fields.Item("U_CardCode").Value = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				DI_oInventoryGenExit.UserFields.Fields.Item("U_CardName").Value = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();
				DI_oInventoryGenExit.UserFields.Fields.Item("U_IssueTyp").Value = "4"; //샘플

				for (i = 1; i <= oMat.RowCount; i++)
				{
					Quantity = Convert.ToDouble(oMat.Columns.Item("Weight").Cells.Item(i).Specific.Value.ToString().Trim());
                    ItemCode = oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim();
					BatchNum = oMat.Columns.Item("LotNo").Cells.Item(i).Specific.Value.ToString().Trim();

					DI_oInventoryGenExit.Lines.Add();
					DI_oInventoryGenExit.Lines.SetCurrentLine(i - 1);
					DI_oInventoryGenExit.Lines.ItemCode = ItemCode;
					DI_oInventoryGenExit.Lines.WarehouseCode = "101";
					DI_oInventoryGenExit.Lines.Quantity = Quantity;

					if (dataHelpClass.GetItem_ManBtchNum(oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim()) == "Y") //배치사용품목이면
					{
						DI_oInventoryGenExit.Lines.BatchNumbers.BatchNumber = BatchNum;
						DI_oInventoryGenExit.Lines.BatchNumbers.Quantity = Quantity;
						DI_oInventoryGenExit.Lines.BatchNumbers.Add();
					}
				}

				RetVal = DI_oInventoryGenExit.Add();

				if (RetVal == 0)
				{
					ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
					oDS_PS_MM092H.SetValue("U_OIGENo", 0, Convert.ToString(ResultDocNum));
				}
				else
				{
					throw new Exception();
				}

				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}

				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				if (DI_oInventoryGenExit != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenExit);
				}
			}
			return ReturnValue;
		}

		/// <summary>
		/// 샘플출고취소(입고)
		/// </summary>
		/// <returns></returns>
		private bool PS_MM092_Add_oInventoryGenEntry()
		{
			bool ReturnValue = false;
			int i;
			int RetVal;
			int ResultDocNum;
			double Quantity;
			string ItemCode;
			string BatchNum;
			string sQry;
			string WhsCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Documents DI_oInventoryGenEntry =  PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

			try
			{
				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				PSH_Globals.oCompany.StartTransaction();
				oMat.FlushToDataSource();

				sQry = "SELECT IGE1.ItemCode,IGE1.Quantity,IGE1.WhsCode,(SELECT BatchNum FROM [IBT1_LINK] WHERE BaseType = '60' AND BaseEntry = OIGE.DocEntry AND BaseLinNum = IGE1.LineNum) AS BatchNum ";
				sQry += " FROM [OIGE] OIGE LEFT JOIN [IGE1] IGE1 ON OIGE.DocEntry = IGE1.DocEntry WHERE OIGE.DocEntry = '" + oForm.Items.Item("OIGENo").Specific.Value.ToString().Trim() + "'";
				oRecordSet.DoQuery(sQry);

				DI_oInventoryGenEntry.DocDate = DateTime.ParseExact(oForm.Items.Item("InDate").Specific.Value, "yyyyMMdd", null);
				DI_oInventoryGenEntry.UserFields.Fields.Item("U_CardCode").Value = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
				DI_oInventoryGenEntry.UserFields.Fields.Item("U_CardName").Value = oForm.Items.Item("CardName").Specific.Value.ToString().Trim();
				DI_oInventoryGenEntry.UserFields.Fields.Item("U_IssueTyp").Value = "4"; //샘플
				DI_oInventoryGenEntry.UserFields.Fields.Item("U_CancDoc").Value = oForm.Items.Item("OIGENo").Specific.Value.ToString().Trim(); //입고(출고취소) 문서번호

				for (i = 1; i <= oRecordSet.RecordCount; i++)
				{
					ItemCode = oRecordSet.Fields.Item("ItemCode").Value;
					Quantity = oRecordSet.Fields.Item("Quantity").Value;
					BatchNum = oRecordSet.Fields.Item("BatchNum").Value;
					WhsCode = oRecordSet.Fields.Item("WhsCode").Value;

					DI_oInventoryGenEntry.Lines.Add();
					DI_oInventoryGenEntry.Lines.SetCurrentLine(i - 1);
					DI_oInventoryGenEntry.Lines.ItemCode = ItemCode;
					DI_oInventoryGenEntry.Lines.WarehouseCode = WhsCode;
					DI_oInventoryGenEntry.Lines.Quantity = Quantity;

					if (dataHelpClass.GetItem_ManBtchNum(oMat.Columns.Item("ItemCode").Cells.Item(i).Specific.Value.ToString().Trim()) == "Y") //배치사용품목이면
					{
						DI_oInventoryGenEntry.Lines.BatchNumbers.BatchNumber = BatchNum;
						DI_oInventoryGenEntry.Lines.BatchNumbers.Quantity = Quantity;
						DI_oInventoryGenEntry.Lines.BatchNumbers.Add();
					}
					oRecordSet.MoveNext();
				}

				RetVal = DI_oInventoryGenEntry.Add();

				if (RetVal == 0)
				{
					ResultDocNum = Convert.ToInt32(PSH_Globals.oCompany.GetNewObjectKey());
					dataHelpClass.DoQuery("UPDATE [@PS_MM092H] SET U_OIGNNo = '" + ResultDocNum + "' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'");
					oRecordSet.MoveFirst();
				}
				else
				{
					throw new Exception();
				}

				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}

				oMat.LoadFromDataSource();
				oMat.AutoResizeColumns();
				ReturnValue = true;
			}
			catch (Exception ex)
			{
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}

				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
			}
			finally
			{
				if (DI_oInventoryGenEntry != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenEntry);
				}
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return ReturnValue;
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
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
					Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
					break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                //case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                //	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //	break;
                //case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                //	Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                //	break;
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
				//           case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE: //21
				//Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
				//break;
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
							if (PS_MM092_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_MM092_Add_oInventoryGenExit() == false) //출고
							{
								PS_MM092_AddMatrixRow(oMat.VisualRowCount, false);
								BubbleEvent = false;
								return;
							}
							oOutMan = oForm.Items.Item("OutMan").Specific.Value.ToString().Trim();
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM092_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							oOutMan = oForm.Items.Item("OutMan").Specific.Value.ToString().Trim();
						}
					}
					else if (pVal.ItemUID == "Btn1") //출고증출력
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_SMM092_Print_Report01);
						thread.SetApartmentState(System.Threading.ApartmentState.STA);
						thread.Start();
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
								PS_MM092_FormItemEnabled();
								PS_MM092_AddMatrixRow(0, true);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_MM092_FormItemEnabled();
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
						else if (pVal.ItemUID == "OutMan")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("OutMan").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						else if (pVal.ItemUID == "TranCard")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("TranCard").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
					}

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PP092No");
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
		/// Raise_EVENT_COMBO_SELECT
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_COMBO_SELECT(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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
						}
						else
						{
							if (pVal.ItemUID == "Tonnage")
							{
								if (oForm.Items.Item("BPLId").Specific.Value == "1")
								{
									sQry = " Select U_RelCd From [@PS_SY001L] WHERE Code = 'M009' and U_Minor = '" + oForm.Items.Item("Tonnage").Specific.Selected.Value.ToString().Trim() + "'";
									oRecordSet.DoQuery(sQry);

									oDS_PS_MM092H.SetValue("U_TranCost", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
									oDS_PS_MM092H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value.ToString().Trim());
								}
							}
							else
							{
								oDS_PS_MM092H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value.ToString().Trim());
							}
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
							oMat.SelectRow(pVal.Row, true, false);
						}
					}
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "1")
					{
						oForm.EnableMenu("1281", true);
					}
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
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
							if (pVal.ColUID == "PP092No")
							{
								oDS_PS_MM092L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());

								if (oMat.RowCount == pVal.Row & !string.IsNullOrEmpty(oDS_PS_MM092L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_MM092_AddMatrixRow(pVal.Row, false);
								}

								sQry = " select  a.DocEntry,";
								sQry += " b.LineId,";
								sQry += " b.U_ItemCode,";
								sQry += " c.FrgnName,";
								sQry += " c.U_Size,";
								sQry += " c.SalUnitMsr,";
								sQry += " b.U_LotNo,";
								sQry += " b.U_Qty,";
								sQry += " B.U_Weight";
								sQry += " from [@PS_PP092H] a Inner Join [@PS_PP092L] b ON a.DocEntry = b.DocEntry and a.Canceled = 'N'";
								sQry += " Inner Join OITM c On b.U_ItemCode = c.ItemCode";
								sQry += " Where b.U_OutGbn = '20'";
								sQry += " AND Convert(Nvarchar(10),a.DocEntry) + '-' + Convert(Nvarchar(10),b.LineId) = '" + oMat.Columns.Item("PP092No").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								oDS_PS_MM092L.SetValue("U_PP092HNo", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
								oDS_PS_MM092L.SetValue("U_PP092LNo", pVal.Row - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oDS_PS_MM092L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet.Fields.Item(2).Value.ToString().Trim());
								oDS_PS_MM092L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item(3).Value.ToString().Trim());
								oDS_PS_MM092L.SetValue("U_Size", pVal.Row - 1, oRecordSet.Fields.Item(4).Value.ToString().Trim());
								oDS_PS_MM092L.SetValue("U_Unit", pVal.Row - 1, oRecordSet.Fields.Item(5).Value.ToString().Trim());
								oDS_PS_MM092L.SetValue("U_LotNo", pVal.Row - 1, oRecordSet.Fields.Item(6).Value.ToString().Trim());
								oDS_PS_MM092L.SetValue("U_Qty", pVal.Row - 1, oRecordSet.Fields.Item(7).Value.ToString().Trim());
								oDS_PS_MM092L.SetValue("U_Weight", pVal.Row - 1, oRecordSet.Fields.Item(8).Value.ToString().Trim());
							}
							else
							{
								oDS_PS_MM092L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
							}

							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else
						{
							if (pVal.ItemUID == "DocEntry")
							{
								oDS_PS_MM092H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
							else if (pVal.ItemUID == "CardCode")
							{
								oDS_PS_MM092H.SetValue("U_CardName", 0, dataHelpClass.Get_ReData("CardName", "CardCode", "[OCRD]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim() + "'", ""));
							}
						}

						oMat.LoadFromDataSource();
						oMat.AutoResizeColumns();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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
					PS_MM092_FormItemEnabled();
					PS_MM092_AddMatrixRow(oMat.VisualRowCount, false);
					oMat.AutoResizeColumns();
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
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
						if (PS_MM092_Validate("행삭제") == false)
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
						oDS_PS_MM092L.RemoveRecord(oDS_PS_MM092L.Size - 1);
						oMat.LoadFromDataSource();

						if (oMat.RowCount == 0)
						{
							PS_MM092_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_MM092L.GetValue("U_PP092No", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_MM092_AddMatrixRow(oMat.RowCount, false);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM092H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM092L);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1283": //삭제
							break;
						case "1284": //취소
							// 마감일자 Check
							if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("InDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
							{
								PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 취소할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
								BubbleEvent = false;
								return;
							}
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
							{
								if (PS_MM092_Validate("취소") == false)
								{
									BubbleEvent = false;
									return;
								}
								if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
								{
									BubbleEvent = false;
									return;
								}

								if (PS_MM092_Add_oInventoryGenEntry() == false)
								{
									BubbleEvent = false;
									return;
								}
							}
							else
							{
								PSH_Globals.SBO_Application.MessageBox("현재 모드에서는 취소할수 없습니다.");
								BubbleEvent = false;
								return;
							}
							break;
						case "1286": //닫기
							// 마감일자 Check
							if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("InDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
							{
								PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 닫기할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
								BubbleEvent = false;
								return;
							}
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "7169": //엑셀 내보내기
							break;
					}
				}
				else if (pVal.BeforeAction == false)
				{
					switch (pVal.MenuUID)
					{
						case "1281": //찾기
							PS_MM092_FormItemEnabled();
							break;
						case "1282": //추가
							PS_MM092_FormItemEnabled();
							PS_MM092_AddMatrixRow(0, true);
							break;
						case "1284": //취소
							break;
						case "1286": //닫기
							break;
						case "1287": // 복제
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_MM092_FormItemEnabled();
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
