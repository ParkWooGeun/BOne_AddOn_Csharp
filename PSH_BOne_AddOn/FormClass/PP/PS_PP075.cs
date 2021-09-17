using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 휘팅이동등록
	/// </summary>
	internal class PS_PP075 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_PP075H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP075L; //등록라인

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP075.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP075_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP075");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocNum";

				oForm.Freeze(true);

				PS_PP075_CreateItems();
				PS_PP075_SetComboBox();
				PS_PP075_ClearForm();
				PS_PP075_MakeMovDocNo();
				PS_PP075_AddMatrixRow(1, 0, true);

				oForm.EnableMenu("1283", false);	// 삭제
				oForm.EnableMenu("1286", false);	// 닫기
				oForm.EnableMenu("1287", false);	// 복제
				oForm.EnableMenu("1284", true);	// 취소
				oForm.EnableMenu("1293", true);	// 행삭제
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
		/// PS_PP075_CreateItems
		/// </summary>
		private void PS_PP075_CreateItems()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oDS_PS_PP075H = oForm.DataSources.DBDataSources.Item("@PS_PP075H");
				oDS_PS_PP075L = oForm.DataSources.DBDataSources.Item("@PS_PP075L");
				oMat = oForm.Items.Item("Mat01").Specific;

				oDS_PS_PP075H.SetValue("U_RegiDate", 0, DateTime.Now.ToString("yyyyMMdd"));

				//담당자
				oDS_PS_PP075H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD());
				PS_PP075_FlushToItemValue("CntcCode", 0, "");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP075_SetComboBox
		/// </summary>
		private void PS_PP075_SetComboBox()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				// 사업장
				sQry = "SELECT BPLId, BPLName From [OBPL] order by 1";
				oRecordSet.DoQuery(sQry);
				while (!oRecordSet.EoF)
				{
					oForm.Items.Item("BPLId").Specific.ValidValues.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim(), oRecordSet.Fields.Item(1).Value.ToString().Trim());
					oRecordSet.MoveNext();
				}
				oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
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
		/// PS_PP075_ClearForm
		/// </summary>
		private void PS_PP075_ClearForm()
		{
			string DocNum;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocNum = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP075'", "");
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP075_MakeMovDocNo
		/// 이동등록번호생성
		/// </summary>
		private void PS_PP075_MakeMovDocNo()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				sQry = "EXEC PS_PP075_01 '" + oDS_PS_PP075H.GetValue("U_RegiDate", 0) + "'"; //인수datetime
				oRecordSet.DoQuery(sQry);
				oDS_PS_PP075H.SetValue("U_MovDocNo", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
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
		/// PS_PP075_AddMatrixRow
		/// </summary>
		/// <param name="MChk"></param>
		/// <param name="oRow"></param>
		/// <param name="Insert_YN"></param>
		private void PS_PP075_AddMatrixRow(int MChk, int oRow, bool Insert_YN)
		{
			try
			{
				switch (MChk)
				{
					case 1:
						if (Insert_YN == false)
						{
							oRow = oMat.RowCount;
							oDS_PS_PP075L.InsertRecord(oRow);
						}
						oDS_PS_PP075L.Offset = oRow;
						oDS_PS_PP075L.SetValue("LineId", oRow, Convert.ToString(oRow + 1));
						oDS_PS_PP075L.SetValue("U_PP070No", oRow, "");
						oDS_PS_PP075L.SetValue("U_ItemCode", oRow, "");
						oDS_PS_PP075L.SetValue("U_ItemName", oRow, "");
						oDS_PS_PP075L.SetValue("U_Size", oRow, "");
						oDS_PS_PP075L.SetValue("U_Mark", oRow, "");
						oDS_PS_PP075L.SetValue("U_Qty", oRow, "");
						oDS_PS_PP075L.SetValue("U_Weight", oRow, "");
						oDS_PS_PP075L.SetValue("U_DocDate", oRow, "");
						oMat.LoadFromDataSource();
						break;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP075_EnableFormItem
		/// </summary>
		private void PS_PP075_EnableFormItem()
		{
			try
			{
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = true;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("DeliArea").Enabled = true;
					oForm.Items.Item("CarNo").Enabled = true;
					oForm.Items.Item("TransCom").Enabled = true;
					oForm.Items.Item("Fee").Enabled = true;
					oForm.Items.Item("RegiDate").Enabled = true;
					oMat.Columns.Item("PP070No").Editable = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					oForm.Items.Item("DocNum").Enabled = false;
					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("DeliArea").Enabled = true;
					oForm.Items.Item("CarNo").Enabled = true;
					oForm.Items.Item("TransCom").Enabled = true;
					oForm.Items.Item("Fee").Enabled = true;
					oForm.Items.Item("RegiDate").Enabled = true;
					oMat.Columns.Item("PP070No").Editable = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					if (oForm.Items.Item("Canceled").Specific.Value == "Y")
					{
						oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("DocNum").Enabled = false;
						oForm.Items.Item("BPLId").Enabled = false;
						oForm.Items.Item("CntcCode").Enabled = false;
						oForm.Items.Item("CardCode").Enabled = false;
						oForm.Items.Item("DeliArea").Enabled = false;
						oForm.Items.Item("CarNo").Enabled = false;
						oForm.Items.Item("TransCom").Enabled = false;
						oForm.Items.Item("Fee").Enabled = false;
						oForm.Items.Item("RegiDate").Enabled = false;
						oMat.Columns.Item("PP070No").Editable = false;
					}
					else
					{
						oForm.Items.Item("DocNum").Enabled = true;
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("CntcCode").Enabled = true;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("DeliArea").Enabled = true;
						oForm.Items.Item("CarNo").Enabled = true;
						oForm.Items.Item("TransCom").Enabled = true;
						oForm.Items.Item("Fee").Enabled = true;
						oForm.Items.Item("RegiDate").Enabled = true;
						oMat.Columns.Item("PP070No").Editable = false;
					}
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_PP075_FlushToItemValue
		/// </summary>
		/// <param name="oUID"></param>
		/// <param name="oRow"></param>
		/// <param name="oCol"></param>
		private void PS_PP075_FlushToItemValue(string oUID, int oRow, string oCol)
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
						sQry = "Select U_FULLNAME From OHEM Where U_MSTCOD = '" + oDS_PS_PP075H.GetValue("U_CntcCode", 0).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oDS_PS_PP075H.SetValue("U_CntcName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
						break;
					case "CardCode":
						sQry = "select cardname from ocrd where cardtype='C' and cardcode = '" + oDS_PS_PP075H.GetValue("U_CardCode", 0).ToString().Trim() + "'";
						oRecordSet.DoQuery(sQry);
						oDS_PS_PP075H.SetValue("U_CardName", 0, oRecordSet.Fields.Item(0).Value.ToString().Trim());
						break;
				}

				if (oUID == "Mat01")
				{
					switch (oCol)
					{
						case "PP070No":
							oMat.FlushToDataSource();
							oDS_PS_PP075L.Offset = oRow - 1;

							DocNum = oMat.Columns.Item("PP070No").Cells.Item(oRow).Specific.String.Split('-')[0]; 
							LineId = oMat.Columns.Item("PP070No").Cells.Item(oRow).Specific.String.Split('-')[1];

							sQry = " select b.U_ItemCode, b.U_ItemName, isnull(c.U_Size,''), isnull(c.U_Mark,''), isnull(d.name,''), b.U_SelQty, b.U_SelWt ";
							sQry += " from [@PS_PP070H] a inner join [@PS_PP070L] b on a.docentry = b.docentry ";
							sQry += " left  join OITM c on b.U_ItemCode = c.ItemCode ";
							sQry += " left  join [@PSH_MARK] d on c.U_Mark = d.Code ";
							sQry += " Where a.DocNum = '" + DocNum + "'";
							sQry += " and b.LineId = '" + LineId + "'";
							oRecordSet.DoQuery(sQry);

							oDS_PS_PP075L.SetValue("U_ItemCode", oRow - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
							oDS_PS_PP075L.SetValue("U_ItemName", oRow - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
							oDS_PS_PP075L.SetValue("U_Size", oRow - 1, oRecordSet.Fields.Item(2).Value.ToString().Trim());
							oDS_PS_PP075L.SetValue("U_Mark", oRow - 1, oRecordSet.Fields.Item(4).Value.ToString().Trim());
							oDS_PS_PP075L.SetValue("U_Qty", oRow - 1, oRecordSet.Fields.Item(5).Value.ToString().Trim());
							oDS_PS_PP075L.SetValue("U_Weight", oRow - 1, oRecordSet.Fields.Item(6).Value.ToString().Trim());
							oMat.SetLineData(oRow);

							if (oRow == oMat.RowCount & !string.IsNullOrEmpty(oDS_PS_PP075L.GetValue("U_PP070No", oRow - 1).ToString().Trim()))
							{
								PS_PP075_AddMatrixRow(1, 0, false);
								oMat.Columns.Item("PP070No").Cells.Item(oRow).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
		}

		/// <summary>
		/// PS_PP075_DelHeaderSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP075_DelHeaderSpaceLine()
		{
			bool functionReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oDS_PS_PP075H.GetValue("U_BPLId", 0).ToString().Trim()))
				{
					errMessage = "사업장은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();

				}
				if (string.IsNullOrEmpty(oDS_PS_PP075H.GetValue("U_CntcCode", 0).ToString().Trim()))
				{
					errMessage = "담당자코드은 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP075H.GetValue("U_CntcName", 0).ToString().Trim()))

				{
					errMessage = "담장자명이 없습니다. 담당자코드를 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP075H.GetValue("U_RegiDate", 0).ToString().Trim()))

				{
					errMessage = "등록일자는 필수사항입니다. 확인하여 주십시오.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oDS_PS_PP075H.GetValue("U_MovDocNo", 0).ToString().Trim()))

				{
					errMessage = "이동등록번호는 필수사항입니다. 확인하여 주십시오.";
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
		/// PS_PP075_DelMatrixSpaceLine
		/// </summary>
		/// <returns></returns>
		private bool PS_PP075_DelMatrixSpaceLine()
		{
			bool functionReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();
				if (oMat.VisualRowCount == 1)
				{
					errMessage = "라인 데이터가 없습니다. 확인하여 주십시오.";
					throw new Exception();
				}

				if (oMat.VisualRowCount > 0)
				{
					for (i = 0; i <= oMat.VisualRowCount - 2; i++)
					{
						oDS_PS_PP075L.Offset = i;
						if (string.IsNullOrEmpty(oDS_PS_PP075L.GetValue("U_PP070No", i).ToString().Trim()))
						{
							errMessage = "벌크포장문서 번호는 필수입니다. 확인하여 주십시오.";
							throw new Exception();
						}
					}
				}

				if (oMat.VisualRowCount > 0)
				{
					oDS_PS_PP075L.RemoveRecord(oDS_PS_PP075L.Size - 1);
				}
				oMat.LoadFromDataSource();

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
		/// PS_PP075_UpDatePP070
		/// </summary>
		/// <param name="A_B"></param>
		/// <returns></returns>
		private bool PS_PP075_UpDatePP070(string A_B)
		{
			bool functionReturnValue = false;
			int i;
			string DocEntry;
			string LineId;
			string sQry;
			string errMessage = string.Empty;

			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				switch (A_B)
				{
					case "A": //[@PS_PP070L]-MovDocNo 에 이동등록번호 update
						for (i = 0; i <= oMat.VisualRowCount - 1; i++)
						{
							oDS_PS_PP075L.Offset = i;

							DocEntry = oDS_PS_PP075L.GetValue("U_PP070No", i).ToString().Trim().Split('-')[0];
							LineId = oDS_PS_PP075L.GetValue("U_PP070No", i).ToString().Trim().Split('-')[1];

							sQry = "update [@PS_PP070L] set U_MovDocNo = '" + oDS_PS_PP075H.GetValue("U_MovDocNo", 0).ToString().Trim() + "' ";
							sQry += "Where docentry = '" + DocEntry + "' ";
							sQry += "and LineId = '" + LineId + "'";
							oRecordSet.DoQuery(sQry);
						}
						break;

					case "B"://[@PS_PP070L]-MovDocNo 에 이동등록번호 update 취소
						for (i = 0; i <= oMat.VisualRowCount - 2; i++)
						{
							oDS_PS_PP075L.Offset = i;

							DocEntry = oDS_PS_PP075L.GetValue("U_PP070No", i).ToString().Trim().Split('-')[0];
							LineId = oDS_PS_PP075L.GetValue("U_PP070No", i).ToString().Trim().Split('-')[1];

							//취소시 휘팅포장등록(PS_PP077)에서 사용되었으면 취소불가
							sQry = "select DocNum from [@PS_PP077H] ";
							sQry += "where isnull(Canceled,'') <> 'Y' ";
							sQry += "and U_MovDocNo = '" + oDS_PS_PP075H.GetValue("U_MovDocNo", 0).ToString().Trim() + "' ";
							sQry += "and U_PP070No = '" + DocEntry + "' ";
							sQry += "and U_PP070NoL = '" + LineId + "'";
							oRecordSet.DoQuery(sQry);

							//사용되었으면
							if (oRecordSet.RecordCount != 0)
							{
								errMessage = "포장처리등록[PS_PP077]에서 사용되어졌습니다. 취소할수 없습니다.";
								throw new Exception();
							}
							else
							{
								sQry = "update [@PS_PP070L] set U_MovDocNo = Null ";
								sQry = sQry + "Where docentry = '" + DocEntry + "' ";
								sQry = sQry + "and LineId = '" + LineId + "'";
								oRecordSet.DoQuery(sQry);
							}
						}
						break;
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
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_PP075_PrintReport
		/// </summary>
		[STAThread]
		private void PS_PP075_PrintReport()
		{
			string WinTitle;
			string ReportName;
			string sQry;
			int DocNum;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocNum = Convert.ToInt32(oForm.Items.Item("DocNum").Specific.Value.ToString().Trim());
				sQry = "SELECT COUNT(*) AS COUNT FROM [@PS_PP075H] H, [@PS_PP075L] L  WHERE H.DocEntry = L.DocEntry and H.DocNum = '" + DocNum + "'";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 10)
				{
					WinTitle = "[PS_PP075] 출고원부/반출증";
					ReportName = "PS_PP075_05.RPT";
				}
				else
				{
					WinTitle = "[PS_PP075] 출고원부/반출증";
					ReportName = "PS_PP075_01.RPT";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				// Formula 수식필드

				// Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocNum", DocNum));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
                    //Raise_EVENT_FORM_RESIZE(FormUID, pVal, BubbleEvent);
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
							if (PS_PP075_DelHeaderSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_PP075_DelMatrixSpaceLine() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_PP075_UpDatePP070("A") == false)
							{
								BubbleEvent = false;
								return;
							}
						}
					}
					else if (pVal.ItemUID == "Print")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_PP075_PrintReport);
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
							oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
							PSH_Globals.SBO_Application.ActivateMenuItem("1282");
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							PS_PP075_EnableFormItem();
							PS_PP075_AddMatrixRow(1, oMat.RowCount, false);
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
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "CntcCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						if (pVal.ItemUID == "CardCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "PP070No")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item("PP070No").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
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
						if (pVal.ItemUID == "CntcCode" || pVal.ItemUID == "CardCode")
						{
							PS_PP075_FlushToItemValue(pVal.ItemUID, 0, "");
						}
						else if (pVal.ItemUID == "RegiDate")
						{
							PS_PP075_MakeMovDocNo();
						}
						if (pVal.ItemUID == "Mat01" && (pVal.ColUID == "PP070No"))
						{
							PS_PP075_FlushToItemValue(pVal.ItemUID, pVal.Row, pVal.ColUID);
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP075H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP075L);
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
							if (PS_PP075_UpDatePP070("B") == false)
							{
								BubbleEvent = false;
								return;
							}
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
							if (PS_PP075_UpDatePP070("A") == false)
							{
								BubbleEvent = false;
								return;
							}
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
								oDS_PS_PP075L.RemoveRecord(oDS_PS_PP075L.Size - 1); // Mat1에 마지막라인(빈라인) 삭제
								oMat.Clear();
								oMat.LoadFromDataSource();
							}
							break;
						case "1281": //찾기
							PS_PP075_EnableFormItem();
							oForm.Items.Item("DocNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							break;
						case "1282": //추가
							PS_PP075_EnableFormItem();
							PS_PP075_ClearForm();
							oDS_PS_PP075H.SetValue("U_RegiDate", 0, DateTime.Now.ToString("yyyyMMdd"));
							PS_PP075_AddMatrixRow(1, 0, true);
							PS_PP075_MakeMovDocNo();
							oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_Index);
							oDS_PS_PP075H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD());
							PS_PP075_FlushToItemValue("CntcCode", 0, "");
							break;
						case "1288": //레코드이동(다음)
						case "1289": //레코드이동(이전)
						case "1290": //레코드이동(최초)
						case "1291": //레코드이동(최종)
							PS_PP075_EnableFormItem();
							if (oMat.VisualRowCount > 0)
							{
								if (!string.IsNullOrEmpty(oMat.Columns.Item("PP070No").Cells.Item(oMat.VisualRowCount).Specific.Value.ToString().Trim()))
								{
									if (oDS_PS_PP075H.GetValue("Status", 0).ToString().Trim() == "O")
									{
										PS_PP075_AddMatrixRow(1, oMat.RowCount, false);
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
