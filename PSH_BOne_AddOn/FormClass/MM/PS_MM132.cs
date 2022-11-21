using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 외주반품등록
	/// </summary>
	internal class PS_MM132 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_MM132H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM132L; //등록라인
		
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값 
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM132.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM132_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM132");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_MM132_CreateItems();
				PS_MM132_ComboBox_Setting();
				PS_MM132_Initial_Setting();
				PS_MM132_EnableMenus();
				PS_MM132_SetDocument(oFormDocEntry);
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
		/// PS_MM132_CreateItems
		/// </summary>
		private void PS_MM132_CreateItems()
		{
			try
			{
				oDS_PS_MM132H = oForm.DataSources.DBDataSources.Item("@PS_MM132H");
				oDS_PS_MM132L = oForm.DataSources.DBDataSources.Item("@PS_MM132L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				oForm.Settings.MatrixUID = "Mat01"; // 서식세팅
				oForm.Settings.Enabled = true;
				oForm.Settings.EnableRowFormat = true;

				oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM132_ComboBox_Setting
		/// </summary>
		private void PS_MM132_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Combo_ValidValues_Insert("PS_MM132", "Mat01", "OutGbn", "10", "원재료");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM132", "Mat01", "OutGbn", "20", "제공");
				dataHelpClass.Combo_ValidValues_SetValueColumn(oMat.Columns.Item("OutGbn"), "PS_MM132", "Mat01", "OutGbn", false);

				dataHelpClass.Combo_ValidValues_Insert("PS_MM132", "OKYNC", "", "N", "반품");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM132", "OKYNC", "", "Y", "승인");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM132", "OKYNC", "", "C", "반품취소");

				dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OKYNC").Specific, "PS_MM132", "OKYNC", false);
				oForm.Items.Item("OKYNC").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL WHERE BPLId = '1'  ORDER BY BPLId", "1", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM132_Initial_Setting
		/// </summary>
		private void PS_MM132_Initial_Setting()
		{
			try
			{
				if (PSH_Globals.oCompany.UserName == "66302" || PSH_Globals.oCompany.UserName == "71090")
				{
					oForm.Items.Item("CardCode").Specific.Value = PSH_Globals.oCompany.UserName;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM132_EnableMenus
		/// </summary>
		private void PS_MM132_EnableMenus()
		{
			try
			{
				oForm.EnableMenu("1283", false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM132_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_MM132_SetDocument(string oFromDocEntry01)
		{
			try
			{

				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_MM132_FormItemEnabled();
					PS_MM132_AddMatrixRow(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_MM132_FormItemEnabled();
					oForm.Items.Item("DocEntry").Specific.Value = oFromDocEntry01;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM132_FormItemEnabled
		/// </summary>
		private void PS_MM132_FormItemEnabled()
		{
			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_MM132_FormClear();
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.EnableMenu("1293", true);
					if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
					{
						oForm.EnableMenu("1284", false); //취소
						oForm.EnableMenu("1286", false); //닫기
						oForm.EnableMenu("1293", false); //행삭제
					}
					else
					{
						oForm.EnableMenu("1284", true); //취소
						oForm.EnableMenu("1286", true); //닫기
						oForm.EnableMenu("1293", true); //행삭제
					}

					oMat.Columns.Item("OutWhCd").Editable = true;
					oMat.Columns.Item("OutWhNm").Editable = true;
					oMat.Columns.Item("InWhCd").Editable = true;
					oMat.Columns.Item("InWhNm").Editable = true;
					oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					
					//외주업체
					if (PSH_Globals.oCompany.UserName == "66302" || PSH_Globals.oCompany.UserName == "71090" || PSH_Globals.oCompany.UserName == "66510")
					{
						if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "N")
						{
							oForm.Items.Item("CardCode").Enabled = true;
							oForm.Items.Item("CardCode").Specific.Value = PSH_Globals.oCompany.UserName;
							oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							oForm.Items.Item("CardCode").Enabled = false;
							oForm.Items.Item("DocDate").Enabled = true;
							oForm.Items.Item("OKYNC").Enabled = true;
							oMat.Columns.Item("OtDocLin").Editable = true;
							oMat.Columns.Item("ReQty").Editable = true;
							oMat.Columns.Item("ReWt").Editable = true;
						}
					}
					else
					{
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("OKYNC").Enabled = true;
						oMat.Columns.Item("OtDocLin").Editable = true;
						oMat.Columns.Item("ReQty").Editable = true;
						oMat.Columns.Item("ReWt").Editable = true;
					}

				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", false); //찾기
					oForm.EnableMenu("1282", true);  //추가
					oForm.EnableMenu("1293", true);

					if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
					{
						oForm.EnableMenu("1284", false); //취소
						oForm.EnableMenu("1286", false); //닫기
						oForm.EnableMenu("1293", false); //행삭제
					}
					else
					{
						oForm.EnableMenu("1284", true); //취소
						oForm.EnableMenu("1286", true); //닫기
						oForm.EnableMenu("1293", true); //행삭제
					}
					
					if (PSH_Globals.oCompany.UserName == "66302" || PSH_Globals.oCompany.UserName == "71090" || PSH_Globals.oCompany.UserName == "66510") //외주업체
					{
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("CardCode").Specific.Value = PSH_Globals.oCompany.UserName;
						oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("CardCode").Enabled = false;
					}
					else
					{
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("DocEntry").Enabled = true;
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("OKYNC").Enabled = true;
					}
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1281", true);   //찾기
					oForm.EnableMenu("1282", true);   //추가
					oForm.EnableMenu("1293", true);

					if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y" | oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "C")
					{
						oForm.EnableMenu("1284", false); //취소
						oForm.EnableMenu("1286", false); //닫기
						oForm.EnableMenu("1293", false); //행삭제
					}
					else
					{
						oForm.EnableMenu("1284", true); //취소
						oForm.EnableMenu("1286", true); //닫기
						oForm.EnableMenu("1293", true); //행삭제
					}

					oMat.Columns.Item("OtDocLin").Editable = false;
					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					
					if (PSH_Globals.oCompany.UserName == "66302" || PSH_Globals.oCompany.UserName == "71090" || PSH_Globals.oCompany.UserName == "66510") //외주업체
					{
						if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "C" || oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
						{
							oForm.Items.Item("CardCode").Enabled = false;
							oForm.Items.Item("DocDate").Enabled = false;
							oForm.Items.Item("OKYNC").Enabled = false;
							oMat.Columns.Item("OtDocLin").Editable = false;
							oMat.Columns.Item("ReQty").Editable = false;
							oMat.Columns.Item("ReWt").Editable = false;
							oMat.Columns.Item("OutWhCd").Editable = false;
							oMat.Columns.Item("OutWhNm").Editable = false;
							oMat.Columns.Item("InWhCd").Editable = false;
							oMat.Columns.Item("InWhNm").Editable = false;
						}
						else
						{
							oForm.Items.Item("CardCode").Enabled = true;
							oForm.Items.Item("DocDate").Enabled = true;
							oForm.Items.Item("OKYNC").Enabled = true;
							oMat.Columns.Item("OtDocLin").Editable = true;
							oMat.Columns.Item("ReQty").Editable = true;
							oMat.Columns.Item("ReWt").Editable = true;
						}
					}
					else
					{
						if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "N" || oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
						{
							oForm.Items.Item("OKYNC").Enabled = true;
							oMat.Columns.Item("OutWhCd").Editable = false;
							oMat.Columns.Item("OutWhNm").Editable = false;
							oMat.Columns.Item("InWhCd").Editable = false;
							oMat.Columns.Item("InWhNm").Editable = false;
						}
						else if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "C")
						{
							oForm.Items.Item("OKYNC").Enabled = false;
						}
						oForm.Items.Item("CardCode").Enabled = false;
					}
					if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "N")
					{
						oMat.Columns.Item("OutWhCd").Editable = true;
						oMat.Columns.Item("OutWhNm").Editable = true;
						oMat.Columns.Item("InWhCd").Editable = true;
						oMat.Columns.Item("InWhNm").Editable = true;
					}
					else if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
					{
						oMat.Columns.Item("OutWhCd").Editable = false;
						oMat.Columns.Item("OutWhNm").Editable = false;
						oMat.Columns.Item("InWhCd").Editable = false;
						oMat.Columns.Item("InWhNm").Editable = false;
					}
					else if (oDS_PS_MM132H.GetValue("U_OKYNC", 0).ToString().Trim() == "C")
					{
						oMat.Columns.Item("OutWhCd").Editable = false;
						oMat.Columns.Item("OutWhNm").Editable = false;
						oMat.Columns.Item("InWhCd").Editable = false;
						oMat.Columns.Item("InWhNm").Editable = false;
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
		/// PS_MM132_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM132_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_MM132L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_MM132L.Offset = oRow;
				oDS_PS_MM132L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_MM132_FormClear
		/// </summary>
		private void PS_MM132_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM132'", "");
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
		/// PS_MM132_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_MM132_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount <= 1)
				{
					errMessage = "라인이 존재하지 않습니다.";
					throw new Exception();
				}
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "사업장 코드는 필수입니다.";
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "외주거래처 코드는 필수입니다.";
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "전기일자는 필수입니다.";
					throw new Exception();
				}

				if (PSH_Globals.oCompany.UserName == "66302" || PSH_Globals.oCompany.UserName == "71090" || PSH_Globals.oCompany.UserName == "66510")
				{
					if (oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() != "N")
					{
						oForm.Items.Item("OKYNC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "이동승인상태 N- 반품만 선택 할 수 있습니다.";
						throw new Exception();
					}
				}

				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("OtDocLin").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("OtDocLin").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "반출문서는 필수입니다.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oMat.Columns.Item("ReQty").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("ReQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "반품수량은 필수입니다.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oMat.Columns.Item("ReWt").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("ReWt").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "반품중량은 필수입니다.";
						throw new Exception();
					}
				}

				oDS_PS_MM132L.RemoveRecord(oDS_PS_MM132L.Size - 1);
				oMat.LoadFromDataSource();
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_MM132_FormClear();
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
		/// PS_MM132_OutDocCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_MM132_OutDocCheck()
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;

			try
			{
				if (string.IsNullOrEmpty(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "사업장 코드를 먼저 입력하세요.";
					throw new Exception();
				}
				else if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value.ToString().Trim()))
				{
					oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "외주거래처 코드를 먼저 입력하세요.";
					throw new Exception();
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
		/// PS_MM132_Validate
		/// </summary>
		/// <param name="ValidateType"></param>
		/// <returns></returns>
		private bool PS_MM132_Validate(string ValidateType)
		{
			bool ReturnValue = false;
			string errMessage = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (ValidateType == "수정")
				{
				}
				else if (ValidateType == "행삭제")
				{
				}
				else if (ValidateType == "취소")
				{
					if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_MM132H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1) == "Y")
					{
						errMessage = "이미취소된 문서 입니다. 취소할수 없습니다.";
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
		/// PS_MM130_StockTrans
		/// </summary>
		/// <returns></returns>
		private bool PS_MM132_StockTrans()
		{
			bool returnValue = false;
			int i;
			int j = 0;
			int RetVal;
			int errDiCode = 0;
			string SDocEntry;
			string sQry;
			string errCode = string.Empty;
			string errDiMsg = string.Empty;
			double ReQty;
			double ReWt;
			SAPbobsCOM.StockTransfer oStockTrans = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				PSH_Globals.oCompany.StartTransaction();
				oMat.FlushToDataSource();

				oStockTrans.DocDate = DateTime.ParseExact(oForm.Items.Item("DocDate").Specific.Value, "yyyyMMdd", null);
				oStockTrans.FromWarehouse = oDS_PS_MM132L.GetValue("U_OutWhCd", 0).ToString().Trim();
				oStockTrans.Comments = "재고이전(" + oDS_PS_MM132H.GetValue("DocEntry", 0).ToString().Trim() + ") 재고이동 - PS_MM132 ";

				for (i = 0; i < oMat.VisualRowCount; i++)
				{
					oStockTrans.Lines.Add();
					oStockTrans.Lines.SetCurrentLine(j);
					oStockTrans.Lines.ItemCode = oDS_PS_MM132L.GetValue("U_OutItmCd", i).ToString().Trim();
					oStockTrans.Lines.WarehouseCode = oDS_PS_MM132L.GetValue("U_InWhCd", i).ToString().Trim();
					oStockTrans.Lines.Quantity = System.Math.Round(Convert.ToDouble(oDS_PS_MM132L.GetValue("U_ReWt", i).ToString().Trim()), 2);
					oStockTrans.Lines.UserFields.Fields.Item("U_UnWeight").Value = System.Math.Round(Convert.ToDouble(oDS_PS_MM132L.GetValue("U_UnWeight", i).ToString().Trim()), 2);
					oStockTrans.Lines.UserFields.Fields.Item("U_Qty").Value = oDS_PS_MM132L.GetValue("U_OutQty", i).ToString().Trim();
					oStockTrans.Lines.BatchNumbers.Notes = "재고이전(Addon)";
					j += 1;
				}
				RetVal = oStockTrans.Add();

				if (RetVal != 0)
				{
					PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
					errCode = "1";
					throw new Exception();
				}
				else
				{
					PSH_Globals.oCompany.GetNewObjectCode(out SDocEntry);
					oDS_PS_MM132H.SetValue("U_STDocNo", 0, SDocEntry);

					for (i = 0; i < oMat.VisualRowCount; i++)
					{
						ReQty = Convert.ToDouble(oDS_PS_MM132L.GetValue("U_ReQty", i));
						ReWt = Convert.ToDouble(oDS_PS_MM132L.GetValue("U_ReWt", i));

						sQry = "Update [@PS_MM130L] Set U_ReQty = IsNull(U_ReQty, 0) + " + ReQty + ", U_ReWt = IsNull(U_ReWt, 0) + " + ReWt + " ";
						sQry += "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry ";
						sQry += "Where b.U_OutDoc = '" + oDS_PS_MM132L.GetValue("U_OutDoc", i).ToString().Trim() + "' ";
						sQry += "And a.U_LineNum = '" + oDS_PS_MM132L.GetValue("U_OutLine", i).ToString().Trim() + "' ";
						oRecordSet.DoQuery(sQry);
					}
				}

				oMat.LoadFromDataSource();
				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}
				oMat.AutoResizeColumns();
				returnValue = true;
			}
			catch (Exception ex)
			{
				returnValue = false;
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				if (errCode == "1")
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oStockTrans);
			}
			return returnValue;
		}

		/// <summary>
		/// PS_MM130_Cancel_oStockTrans
		/// </summary>
		/// <returns></returns>
		private bool PS_MM132_Cancel_StockTrans()
		{
			bool returnValue = false;
			int i;
			int RetVal;
			int errDiCode = 0;
			string sQry;
			string errCode = string.Empty;
			string errDiMsg = string.Empty;
			double ReQty;
			double ReWt;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.StockTransfer oStockTrans = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

			try
			{
				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				PSH_Globals.oCompany.StartTransaction();

				if (oStockTrans.GetByKey(Convert.ToInt32(oDS_PS_MM132H.GetValue("U_STDocNo", 0).ToString().Trim())) == false)
				{
					PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
					throw new Exception();
				}
				
				RetVal = oStockTrans.Cancel();

				if (RetVal != 0)
				{
					PSH_Globals.oCompany.GetLastError(out errDiCode, out errDiMsg);
					errCode = "1";
					throw new Exception();
				}
				else
				{
                    oDS_PS_MM132H.SetValue("Status", 0, "C");
					oDS_PS_MM132H.SetValue("Canceled", 0, "Y");

                    for (i = 0; i <= oMat.VisualRowCount -1; i++)
					{
						ReQty = Convert.ToDouble(oDS_PS_MM132L.GetValue("U_ReQty", i));
						ReWt = Convert.ToDouble(oDS_PS_MM132L.GetValue("U_ReWt", i));

						sQry = "Update [@PS_MM130L] ";
						sQry += "Set U_ReQty = IsNull(U_ReQty, 0) - " + ReQty + ", U_ReWt = IsNull(U_ReWt, 0) - " + ReWt + " ";
						sQry += "From [@PS_MM130L] a Inner Join [@PS_MM130H] b On a.DocEntry = b.DocEntry ";
						sQry += "Where b.U_OutDoc = '" + oDS_PS_MM132L.GetValue("U_OutDoc", i).ToString() +"' ";
						sQry += "And a.U_LineNum = '" + oDS_PS_MM132L.GetValue("U_OutLine", i).ToString() +"' ";
						oRecordSet.DoQuery(sQry);
					}
					oForm.Items.Item("STDocCl").Specific.Value = dataHelpClass.Get_ReData("MAX(DocNum)", "ObjType", "OWTR", "'67'", "");
				}

				oMat.LoadFromDataSource();
				if (PSH_Globals.oCompany.InTransaction == true)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
				}
				oMat.AutoResizeColumns();
				returnValue = true;
			}
			catch (Exception ex)
			{
				returnValue = false;
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}
				if (errCode == "1")
				{
					PSH_Globals.SBO_Application.StatusBar.SetText("DI실행 중 오류 발생 : [" + errDiCode + "]" + errDiMsg, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
				else
				{
					PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oStockTrans);
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
					Raise_EVENT_GOT_FOCUS(FormUID, ref pVal, ref BubbleEvent);
					break;
				//case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
				//    Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
				//    break;
				//case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
				//	Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_CLICK: //6
				//	Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
				//	Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
				//	break;
				//case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
				//    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
				//    break;
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
				//	Raise_EVENT_FORM_ACTIVATE(FormUID, ref pVal, ref BubbleEvent);
				//	break;
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
                case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST: //27
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
                    break;
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
			int i;
			int j = 0;
			string errMessage = string.Empty;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						//외주 재고 이동 DI API
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_MM132_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							
						}

						//외주 재고 이동 취소 DI API
						else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) && oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() == "Y" && string.IsNullOrEmpty(oForm.Items.Item("STDocNo").Specific.Value.ToString().Trim()))
						{
							if (PS_MM132_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}

                            for (i = 0; i < oMat.RowCount; i++)
                            {
                                if (oMat.Columns.Item("OutGbn").Cells.Item(i + 1).Specific.Value.ToString().Trim() == "10")
                                {
                                    if (j == 0)
                                    {
                                        j += 1;

                                        if (PS_MM132_StockTrans() == false)
										{
											PS_MM132_AddMatrixRow(oMat.VisualRowCount, false);
											BubbleEvent = false;
											return;
										}
                                    }
                                }
                            }
                        }
						else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) && oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() == "C")
						{
							if (PS_MM132_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}

                            for (i = 0; i < oMat.RowCount; i++)
                            {
                                if (oMat.Columns.Item("OutGbn").Cells.Item(i + 1).Specific.Value.ToString().Trim() == "10")
                                {
                                    if (j == 0)
                                    {
                                        j += 1;

                                        if (PS_MM132_Cancel_StockTrans() == false)
										{
											PS_MM132_AddMatrixRow(0, true);
											BubbleEvent = false;
											return;
										}
                                    }
                                }
                            }
                        }
						//창원 문서 수정
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && (oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() == "N" && string.IsNullOrEmpty(oForm.Items.Item("STDocNo").Specific.Value.ToString().Trim())))
						{
							if (PS_MM132_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							
						}
						//창원
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && (oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() == "N" && !string.IsNullOrEmpty(oForm.Items.Item("STDocNo").Specific.Value.ToString().Trim())))
						{
							errMessage = "이미 승인되었습니다.";
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
								PS_MM132_FormItemEnabled();
								PS_MM132_AddMatrixRow(oMat.RowCount, true);
								oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
								oForm.Items.Item("BPLId").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
								oForm.Items.Item("OKYNC").Specific.Select("Y", SAPbouiCOM.BoSearchKey.psk_ByValue);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_MM132_FormItemEnabled();
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
					BubbleEvent = false;
				}
				else
				{
					PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
				}
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
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

					if (pVal.ItemUID == "Mat01" && pVal.ColUID == "OtDocLin")
					{
						if (PS_MM132_OutDocCheck() == true)
						{
							dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OtDocLin");
						}
					}
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OutWhCd");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "InWhCd");
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
		/// Raise_EVENT_VALIDATE
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			int c;
			string DocEntry;
			string LineId;
			string DocLin;
			string sQry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			SAPbobsCOM.Recordset oRecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemChanged == true)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "OtDocLin")
							{
								DocLin = oMat.Columns.Item("OtDocLin").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
								c = DocLin.IndexOf("-");
								DocEntry = DocLin.Substring(0, c);
								LineId = DocLin.Substring(c + 1);

								sQry = "EXEC PS_MM132_01 '" + DocEntry + "', '" + LineId + "'";
								oRecordSet.DoQuery(sQry);

								oMat.FlushToDataSource();

								oDS_PS_MM132L.SetValue("U_ItemCode", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item(1).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_Size", pVal.Row - 1, oRecordSet.Fields.Item(2).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_Mark", pVal.Row - 1, oRecordSet.Fields.Item(3).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_UnWeight", pVal.Row - 1, oRecordSet.Fields.Item(4).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutItmCd", pVal.Row - 1, oRecordSet.Fields.Item(5).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutItmNm", pVal.Row - 1, oRecordSet.Fields.Item(6).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutGbn", pVal.Row - 1, oRecordSet.Fields.Item(7).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutQty", pVal.Row - 1, oRecordSet.Fields.Item(8).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutWt", pVal.Row - 1, oRecordSet.Fields.Item(9).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_ReQty", pVal.Row - 1, oRecordSet.Fields.Item(10).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_ReWt", pVal.Row - 1, oRecordSet.Fields.Item(11).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutWhCd", pVal.Row - 1, oRecordSet.Fields.Item(12).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutWhNm", pVal.Row - 1, oRecordSet.Fields.Item(13).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_InWhCd", pVal.Row - 1, oRecordSet.Fields.Item(14).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_InWhNm", pVal.Row - 1, oRecordSet.Fields.Item(15).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_PP030HNo", pVal.Row - 1, oRecordSet.Fields.Item(16).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_PP030MNo", pVal.Row - 1, oRecordSet.Fields.Item(17).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutDoc", pVal.Row - 1, oRecordSet.Fields.Item(18).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OutLine", pVal.Row - 1, oRecordSet.Fields.Item(19).Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OrdNum", pVal.Row - 1, oRecordSet.Fields.Item(20).Value.ToString().Trim());

								oDS_PS_MM132L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());

								if ((oMat.RowCount == pVal.Row) && !string.IsNullOrEmpty(oDS_PS_MM132L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_MM132_AddMatrixRow(pVal.Row, false);
								}
							}
							else if (pVal.ColUID == "OutWhCd")
							{
								oDS_PS_MM132L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_OUtWhNm", pVal.Row - 1, dataHelpClass.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" + oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
								oMat.FlushToDataSource();
							}
							else if (pVal.ColUID == "InWhCd")
							{
								oDS_PS_MM132L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_MM132L.SetValue("U_InWhNm", pVal.Row - 1, dataHelpClass.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" + oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
								oMat.FlushToDataSource();
							}
							else if (pVal.ColUID == "ReQty")
							{
								oMat.FlushToDataSource();

								if (Convert.ToDouble(oMat.Columns.Item("ReQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) > 0)
								{
									sQry = "Select U_ObasUnit FROM OITM WHERE ItemCode = '" + oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
									oRecordSet.DoQuery(sQry);

									sQry = "Select OnHand, U_Qty FROM OITW WHERE ItemCode = '" + oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "' AND WhsCode = '801'";
									oRecordSet02.DoQuery(sQry);

									if (oRecordSet.Fields.Item(0).Value.ToString().Trim().Substring(0, 1) == "1")
									{
										oMat.Columns.Item("ReWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("ReQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
									}
									else if (oRecordSet.Fields.Item(0).Value.ToString().Trim().Substring(0, 1) == "2")
									{
										if (Convert.ToDouble(oRecordSet02.Fields.Item(1).Value.ToString().Trim()) > 0 && !string.IsNullOrEmpty(oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
										{
											oMat.Columns.Item("ReWt").Cells.Item(pVal.Row).Specific.Value = Convert.ToString((Convert.ToDouble(oRecordSet02.Fields.Item(0).Value.ToString().Trim()) / Convert.ToDouble(oRecordSet02.Fields.Item(1).Value.ToString().Trim())) * Convert.ToDouble(oMat.Columns.Item("ReQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()));
										}
										else
										{
											oMat.Columns.Item("ReWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("ReWt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
										}
									}

									oMat.LoadFromDataSource();
									oDS_PS_MM132L.SetValue("U_ReQty", pVal.Row - 1, oMat.Columns.Item("ReQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									oDS_PS_MM132L.SetValue("U_ReWt", pVal.Row - 1, oMat.Columns.Item("ReWt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									oMat.Columns.Item("ReWt").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}
							}
							else
							{
								oDS_PS_MM132L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oMat.LoadFromDataSource();
							}

							oMat.LoadFromDataSource();
							oMat.AutoResizeColumns();
							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (pVal.ItemUID == "CardCode")
						{
							sQry = "SELECT CardName FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "DocEntry")
						{
							oDS_PS_MM132H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
						}
					}
					oForm.Update();
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// 네비게이션 메소드(Raise_FormMenuEvent 에서 사용)
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_RECORD_MOVE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			string sQry;
			string docEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				docEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim(); //현재문서번호

				if (pVal.MenuUID == "1288") //다음
				{
					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
					{
						PSH_Globals.SBO_Application.ActivateMenuItem("1290");
						return;
					}
					else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
					{
						if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))
						{
							PSH_Globals.SBO_Application.ActivateMenuItem("1290");
							return;
						}
					}
					else
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
						oForm.Items.Item("DocEntry").Enabled = true;
						sQry = "  Select min(DocEntry)";
						sQry += "  From [@PS_MM132H]";
						sQry += " Where U_CardCode = '" + PSH_Globals.oCompany.UserName + "'";
						sQry += "   AND DocEntry > " + docEntry;

						oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(sQry, 0, 1);
						oForm.Items.Item("1").Enabled = true;
						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("DocEntry").Enabled = false;
					}
				}
				else if (pVal.MenuUID == "1289") //이전
				{
					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
					{
						PSH_Globals.SBO_Application.ActivateMenuItem("1291");
						return;
					}
					else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
					{
						if (string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))
						{
							PSH_Globals.SBO_Application.ActivateMenuItem("1291");
							return;
						}
					}
					else
					{
						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
						oForm.Items.Item("DocEntry").Enabled = true;
						sQry = "  Select max(DocEntry)";
						sQry += "  From [@PS_MM130H]";
						sQry += " Where U_CardCode = '" + PSH_Globals.oCompany.UserName + "'";
						sQry += "   AND DocEntry < " + docEntry;

						oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(sQry, 0, 1);
						oForm.Items.Item("1").Enabled = true;
						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("DocEntry").Enabled = false;
					}
				}
				else if (pVal.MenuUID == "1290") //최초
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					oForm.Items.Item("DocEntry").Enabled = true;
					sQry = "  Select Min(DocEntry)";
					sQry += "  From [@PS_MM130H]";
					sQry += " Where U_CardCode = '" + PSH_Globals.oCompany.UserName + "'";

					oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(sQry, 0, 1);
					oForm.Items.Item("1").Enabled = true;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("DocEntry").Enabled = false;
				}
				else if (pVal.MenuUID == "1291") //최종
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					oForm.Items.Item("DocEntry").Enabled = true;
					sQry = "  Select Max(DocEntry)";
					sQry += "  From [@PS_MM130H]";
					sQry += " Where U_CardCode = '" + PSH_Globals.oCompany.UserName + "'";

					oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(sQry, 0, 1);
					oForm.Items.Item("1").Enabled = true;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("DocEntry").Enabled = false;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
			{
				BubbleEvent = false;
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
					PS_MM132_FormItemEnabled();
					oMat.AutoResizeColumns();
					PS_MM132_AddMatrixRow(oMat.VisualRowCount, false);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// Raise_EVENT_CHOOSE_FROM_LIST
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; 
			
			try
			{
				if (pVal.BeforeAction == true)
				{
				}
				else if (pVal.BeforeAction == false)
				{
					if (pVal.ItemUID == "CardCode")
					{
						if (oDataTable01 == null)
						{
						}
						else
						{
							oDS_PS_MM132H.SetValue("U_CardCode", 0, oDataTable01.Columns.Item("CardCode").Cells.Item(0).Value.ToString().Trim());
							oDS_PS_MM132H.SetValue("U_CardName", 0, oDataTable01.Columns.Item("CardName").Cells.Item(0).Value.ToString().Trim());
							// 찾기나 문서이동 버튼 클릭 시에 갱신으로 바뀌지 않음
							if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							}
						}
					}

					oForm.Update();
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM132H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM132L);
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
		private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
		{
			int i;

			try
			{
				if (oLastColRow01 > 0)
				{
					if (pVal.BeforeAction == true)
					{
					}
					else if (pVal.BeforeAction == false)
					{
						for (i = 1; i <= oMat.VisualRowCount; i++)
						{
							oMat.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
						}
						oMat.FlushToDataSource();
						oDS_PS_MM132L.RemoveRecord(oDS_PS_MM132L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_MM132_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_MM132L.GetValue("U_OtDocLin", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_MM132_AddMatrixRow(oMat.RowCount, false);
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
						case "1281": //찾기
							break;
						case "1282": //추가
							break;
						case "1283": //삭제
							
							break;
						case "1284": //취소
							if (PS_MM132_Validate("취소") == false)
							{
								BubbleEvent = false;
								return;
							}
							break;
						case "1286": //닫기
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							if (PSH_Globals.oCompany.UserName.Substring(0, 1) == "6" || PSH_Globals.oCompany.UserName.Substring(0, 1) == "7")
							{
								Raise_EVENT_RECORD_MOVE(FormUID, ref pVal, ref BubbleEvent);
							}
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
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
							PS_MM132_FormItemEnabled();
							break;
						case "1282": //추가
							oDS_PS_MM132H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
							oDS_PS_MM132H.SetValue("U_BPLId", 0, "1");
							oDS_PS_MM132H.SetValue("U_OKYNC", 0, "N");
							PS_MM132_FormItemEnabled();
							PS_MM132_AddMatrixRow(0, true);
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
							PS_MM132_FormItemEnabled();
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
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
