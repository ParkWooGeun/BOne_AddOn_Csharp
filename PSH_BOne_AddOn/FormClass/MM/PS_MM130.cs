using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.DataPack;
using PSH_BOne_AddOn.Form;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 외주반출등록
	/// </summary>
	internal class PS_MM130 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
			
		private SAPbouiCOM.DBDataSource oDS_PS_MM130H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM130L; //등록라인

		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;  //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;     //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        // 입고 DI를 위한 정보를 가지는 구조체
		private struct StockInfos
		{
			public string CardCode;           //고객코드
			public string ItemCode;           //품목코드
			public string FromWarehouseCode;  //창고코드
			public string ToWarehouseCode;    //창고코드
			public double Weight;
			public double UnWeight;
			public string BatchNum;           //배치번호
			public double BatchWeight;        //배치중량
			public int Qty;                   //수량
			public string TransNo;            //재고이전문서번호
			public string Chk;
			public int MatrixRow;
			public string StockTransDocEntry; //재고이전문서번호
			public string StockTransLineNum;  //재고이전라인번호
			public string Indate;             //전기일
		}

		private StockInfos[] StockInfo;

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM130.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM130_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM130");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				PS_MM130_CreateItems();
				PS_MM130_ComboBox_Setting();
				PS_MM130_Initial_Setting();
				PS_MM130_CF_ChooseFromList();
				PS_MM130_EnableMenus();
				PS_MM130_SetDocument(oFormDocEntry);
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
		/// PS_MM130_CreateItems
		/// </summary>
		private void PS_MM130_CreateItems()
		{
			try
			{
				oDS_PS_MM130H = oForm.DataSources.DBDataSources.Item("@PS_MM130H");
				oDS_PS_MM130L = oForm.DataSources.DBDataSources.Item("@PS_MM130L");
				oMat = oForm.Items.Item("Mat01").Specific;
				oMat.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
				oMat.AutoResizeColumns();

				// 원재료반출
				oForm.Items.Item("Rad01").Specific.ValOn = "10";
				oForm.Items.Item("Rad01").Specific.ValOff = "0";
				oForm.Items.Item("Rad01").Specific.Selected = true;

				// 재공반출
				oForm.Items.Item("Rad02").Specific.ValOn = "20";
				oForm.Items.Item("Rad02").Specific.ValOff = "0";
				oForm.Items.Item("Rad02").Specific.GroupWith("Rad01");

				oForm.Settings.MatrixUID = "Mat01"; // 서식세팅
				oForm.Settings.Enabled = true;
				oForm.Settings.EnableRowFormat = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM130_ComboBox_Setting
		/// </summary>
		private void PS_MM130_ComboBox_Setting()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Combo_ValidValues_Insert("PS_MM130", "Mat01", "OutGbn", "10", "원재료");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM130", "Mat01", "OutGbn", "20", "제공");
				dataHelpClass.Combo_ValidValues_SetValueColumn(oMat.Columns.Item("OutGbn"), "PS_MM130", "Mat01", "OutGbn", false);

				dataHelpClass.Combo_ValidValues_Insert("PS_MM130", "OKYNC", "", "N", "재고이동");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM130", "OKYNC", "", "C", "이동취소");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM130", "OKYNC", "", "Y", "승인");
				dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OKYNC").Specific, "PS_MM130", "OKYNC", false);
				oForm.Items.Item("OKYNC").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL WHERE BPLId = '1' ORDER BY BPLId", "1", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM130_Initial_Setting
		/// </summary>
		private void PS_MM130_Initial_Setting()
		{
			string lcl_User_BPLId;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				lcl_User_BPLId = dataHelpClass.User_BPLID();

				if (lcl_User_BPLId == "1")
				{
					oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				}
				// 인수자
				oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
		}

		/// <summary>
		/// PS_MM130_CF_ChooseFromList
		/// </summary>
		private void PS_MM130_CF_ChooseFromList()
		{
			SAPbouiCOM.ChooseFromList oCFL02 = null;
			SAPbouiCOM.ChooseFromListCollection oCFLs02 = null;
			SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams02 = null;
			SAPbouiCOM.EditText oEdit02 = null;

			try
			{
				oEdit02 = oForm.Items.Item("ShipTo").Specific;
				oCFLs02 = oForm.ChooseFromLists;
				oCFLCreationParams02 = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

				//oCFLCreationParams02.ObjectType = Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_BusinessPartner);
				oCFLCreationParams02.ObjectType = "2";
				oCFLCreationParams02.UniqueID = "CFLSHIPCODE";
				oCFLCreationParams02.MultiSelection = false;
				oCFL02 = oCFLs02.Add(oCFLCreationParams02);

				oEdit02.ChooseFromListUID = "CFLSHIPCODE";
				oEdit02.ChooseFromListAlias = "CardCode";
			}
			catch(Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				if (oCFL02 != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL02);
				}
				if (oCFLs02 != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs02);
				}
				if (oCFLCreationParams02 != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams02);
				}
				if (oEdit02 != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit02);
				}
			}
		}

		/// <summary>
		/// PS_MM130_EnableMenus
		/// </summary>
		private void PS_MM130_EnableMenus()
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
		/// PS_MM130_SetDocument
		/// </summary>
		/// <param name="oFromDocEntry01"></param>
		private void PS_MM130_SetDocument(string oFromDocEntry01)
		{
			try
			{
				if (string.IsNullOrEmpty(oFromDocEntry01))
				{
					PS_MM130_FormItemEnabled();
					PS_MM130_AddMatrixRow(0, true);
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_MM130_FormItemEnabled();
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
		/// PS_MM130_FormItemEnabled
		/// </summary>
		private void PS_MM130_FormItemEnabled()
		{
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oForm.Freeze(true);
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_MM130_FormClear();

					oForm.EnableMenu("1281", true);	 //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.EnableMenu("1293", true); // 행삭제

					oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
					oForm.Items.Item("Purpose").Specific.Value = "외주가공";
					oForm.Items.Item("ShipCo").Specific.Value = "업체자가";

					PS_MM130_Initial_Setting();

					sQry = "Select U_FULLNAME, U_MSTCOD From [OHEM] Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
					oRecordSet.DoQuery(sQry);

					oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("OutDoc").Enabled = false;
					oForm.Items.Item("Print").Enabled = false;

					if (oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "C" || oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
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

					//외주업체
					if (PSH_Globals.oCompany.UserName == "66302" || PSH_Globals.oCompany.UserName == "71090" || PSH_Globals.oCompany.UserName == "66510")
					{
						oForm.Items.Item("BPLId").Enabled = false;
						oForm.Items.Item("Print").Enabled = false;
						oForm.Items.Item("Rad01").Enabled = false;
						oForm.Items.Item("Rad02").Enabled = false;
						oForm.Items.Item("CardCode").Enabled = false;
						oForm.Items.Item("CntcCode").Enabled = false;
						oForm.Items.Item("ShipTo").Enabled = false;
						oForm.Items.Item("CarNo").Enabled = false;
						oForm.Items.Item("ShipCo").Enabled = false;
						oForm.Items.Item("DocDate").Enabled = false;
						oForm.Items.Item("Fare").Enabled = false;
						oForm.Items.Item("ArrivePl").Enabled = false;
						oForm.Items.Item("OKYNC").Enabled = true;
						oMat.Columns.Item("OrdNum").Editable = false;
						oMat.Columns.Item("OutItmCd").Editable = false;
						oMat.Columns.Item("OutQty").Editable = false;
						oMat.Columns.Item("OutWt").Editable = false;
						oMat.Columns.Item("OutWhCd").Editable = false;
						oMat.Columns.Item("OutWhNm").Editable = false;
						oMat.Columns.Item("InWhCd").Editable = false;
						oMat.Columns.Item("InWhNm").Editable = false;
					}
					else
					{
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("CntcCode").Enabled = true;
						oForm.Items.Item("ShipTo").Enabled = true;
						oForm.Items.Item("CarNo").Enabled = true;
						oForm.Items.Item("ShipCo").Enabled = true;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("Fare").Enabled = true;
						oForm.Items.Item("ArrivePl").Enabled = true;
						oMat.Columns.Item("OrdNum").Editable = true;
						oMat.Columns.Item("OutQty").Editable = true;
						oMat.Columns.Item("OutWt").Editable = true;
						oMat.Columns.Item("OutWhCd").Editable = true;
						oMat.Columns.Item("OutWhNm").Editable = true;
						oMat.Columns.Item("InWhCd").Editable = true;
						oMat.Columns.Item("InWhNm").Editable = true;
						oForm.Items.Item("Rad01").Enabled = true;
						oForm.Items.Item("Rad02").Enabled = true;
						oForm.Items.Item("OKYNC").Enabled = true;
						oMat.Columns.Item("OrdNum").Editable = true;
					}
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.EnableMenu("1282", true); //추가
					oForm.EnableMenu("1293", true); //행삭제
					oForm.Items.Item("Print").Enabled = false;
					
					if (oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "C" || oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
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

					//외주업체
					if (PSH_Globals.oCompany.UserName == "66302" || PSH_Globals.oCompany.UserName == "71090" || PSH_Globals.oCompany.UserName == "66510")
					{
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("CardCode").Specific.Value = PSH_Globals.oCompany.UserName;
						oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						oForm.Items.Item("CardCode").Enabled = false;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("OKYNC").Enabled = true;
					}
					else
					{
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("DocEntry").Enabled = true;
						oForm.Items.Item("OutDoc").Enabled = true;
						oMat.Columns.Item("OrdNum").Editable = true;
						oForm.Items.Item("Rad01").Enabled = false;
						oForm.Items.Item("Rad02").Enabled = false;
						oForm.Items.Item("BPLId").Enabled = true;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("CntcCode").Enabled = true;
						oForm.Items.Item("ShipTo").Enabled = true;
						oForm.Items.Item("CarNo").Enabled = true;
						oForm.Items.Item("ShipCo").Enabled = true;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("Fare").Enabled = true;
						oForm.Items.Item("ArrivePl").Enabled = true;
						oForm.Items.Item("OKYNC").Enabled = true;
					}
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1281", true); //찾기
					oForm.EnableMenu("1282", true); //추가
					oForm.EnableMenu("1293", true); // 행삭제
					oForm.Items.Item("Print").Enabled = true;

					if (oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "N" || oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
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

					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("OutDoc").Enabled = false;
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Rad01").Enabled = false;
					oForm.Items.Item("Rad02").Enabled = false;
					oForm.Items.Item("DocEntry").Enabled = false;

					//외주업체
					if (PSH_Globals.oCompany.UserName== "66302" || PSH_Globals.oCompany.UserName == "71090" || PSH_Globals.oCompany.UserName == "66510")
					{
						oForm.Items.Item("BPLId").Enabled = false;
						oForm.Items.Item("CardCode").Enabled = false;
						oForm.Items.Item("CntcCode").Enabled = false;
						oForm.Items.Item("ShipTo").Enabled = false;
						oForm.Items.Item("CarNo").Enabled = false;
						oForm.Items.Item("ShipCo").Enabled = false;
						oForm.Items.Item("DocDate").Enabled = false;
						oForm.Items.Item("Fare").Enabled = false;
						oForm.Items.Item("ArrivePl").Enabled = false;
						if (oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "C" || oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
						{
							oForm.Items.Item("OKYNC").Enabled = false;
						}
						else
						{
							oForm.Items.Item("OKYNC").Enabled = true;
						}
						oMat.Columns.Item("OrdNum").Editable = false;
						oMat.Columns.Item("OutItmCd").Editable = false;
						oMat.Columns.Item("OutQty").Editable = false;
						oMat.Columns.Item("OutWt").Editable = false;
						oMat.Columns.Item("OutWhCd").Editable = false;
						oMat.Columns.Item("OutWhNm").Editable = false;
						oMat.Columns.Item("InWhCd").Editable = false;
						oMat.Columns.Item("InWhNm").Editable = false;
					}
					else
					{
						if (oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "C" || oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
						{
							if (oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "N")
							{
								oForm.Items.Item("OKYNC").Enabled = true;
							}
							else if (oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "C" || oDS_PS_MM130H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
							{
								oForm.Items.Item("OKYNC").Enabled = false;
							}
							oForm.Items.Item("CardCode").Enabled = false;
							oForm.Items.Item("CntcCode").Enabled = false;
							oForm.Items.Item("ShipTo").Enabled = false;
							oForm.Items.Item("CarNo").Enabled = false;
							oForm.Items.Item("ShipCo").Enabled = false;
							oForm.Items.Item("DocDate").Enabled = false;
							oForm.Items.Item("Fare").Enabled = false;
							oForm.Items.Item("ArrivePl").Enabled = false;
							oMat.Columns.Item("OrdNum").Editable = false;
							oMat.Columns.Item("OutQty").Editable = false;
							oMat.Columns.Item("OutWt").Editable = false;
							oMat.Columns.Item("OutWhCd").Editable = false;
							oMat.Columns.Item("OutWhNm").Editable = false;
							oMat.Columns.Item("InWhCd").Editable = false;
							oMat.Columns.Item("InWhNm").Editable = false;
						}
						else
						{
							oForm.Items.Item("Rad01").Enabled = true;
							oForm.Items.Item("Rad02").Enabled = true;
							oForm.Items.Item("CardCode").Enabled = true;
							oForm.Items.Item("CntcCode").Enabled = true;
							oForm.Items.Item("ShipTo").Enabled = true;
							oForm.Items.Item("CarNo").Enabled = true;
							oForm.Items.Item("ShipCo").Enabled = true;
							oForm.Items.Item("DocDate").Enabled = true;
							oForm.Items.Item("Fare").Enabled = true;
							oForm.Items.Item("ArrivePl").Enabled = true;
							oMat.Columns.Item("OrdNum").Editable = true;
							oMat.Columns.Item("OutQty").Editable = true;
							oMat.Columns.Item("OutWt").Editable = true;
							oMat.Columns.Item("OutWhCd").Editable = true;
							oMat.Columns.Item("OutWhNm").Editable = true;
							oMat.Columns.Item("InWhCd").Editable = true;
							oMat.Columns.Item("InWhNm").Editable = true;
							oForm.Items.Item("OKYNC").Enabled = true;
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
				oForm.Freeze(false);
			}
		}

		/// <summary>
		/// PS_MM130_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM130_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_MM130L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_MM130L.Offset = oRow;
				oDS_PS_MM130L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_MM130_FormClear
		/// </summary>
		private void PS_MM130_FormClear()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM130'", "");
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
		/// PS_MM130_DataValidCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_MM130_DataValidCheck()
		{
			bool ReturnValue = false;
			int i;
			string errMessage = string.Empty;

			try
			{
				oMat.FlushToDataSource();

				if (oMat.VisualRowCount == 1)
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
					if (oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() != "Y")
					{
						oForm.Items.Item("OKYNC").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "이동승인상태 Y - 승인만 선택 할 수 있습니다.";
						throw new Exception();
					}
				}

				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if (string.IsNullOrEmpty(oMat.Columns.Item("OrdNum").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("OrdNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "작지번호는 필수입니다.";
						throw new Exception();
					}
					else if (Convert.ToDouble(oMat.Columns.Item("OutQty").Cells.Item(i).Specific.Value.ToString().Trim()) < 0 || Convert.ToDouble(oMat.Columns.Item("OutWt").Cells.Item(i).Specific.Value.ToString().Trim()) < 0)
					{
						oMat.Columns.Item("OutQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "반출수(중)량은 필수입니다.";
						throw new Exception();
					}
					else if (string.IsNullOrEmpty(oMat.Columns.Item("TCpCode").Cells.Item(i).Specific.Value.ToString().Trim()))
					{
						oMat.Columns.Item("TCpCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "외주 마지막공정(외주공정구간)을 선택하셔야 합니다.";
						throw new Exception();
					}
				}

				oDS_PS_MM130L.RemoveRecord(oDS_PS_MM130L.Size - 1);
				oMat.LoadFromDataSource();
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_MM130_FormClear();
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
		/// PS_MM130_DataInsertCheck
		/// </summary>
		/// <returns></returns>
		private bool PS_MM130_DataInsertCheck()
		{
			bool ReturnValue = false;
			string oDate;
			string oOutDoc;
			int i;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				sQry = "SELECT ISNULL(MAX(U_OutDoc), 0) FROM [@PS_MM130H] WHERE SubString(U_OutDoc,1,8) = '" + oDate + "'";
				oRecordSet.DoQuery(sQry);

				oOutDoc = oRecordSet.Fields.Item(0).Value.ToString().Trim();

				if (Convert.ToDouble(oOutDoc) == 0)
				{
					oOutDoc = oDate + "001";
					oDS_PS_MM130H.SetValue("U_OutDoc", 0, oOutDoc);
				}
				else
				{
					oOutDoc = Convert.ToString(Convert.ToDouble(oOutDoc) + 1);
					oDS_PS_MM130H.SetValue("U_OutDoc", 0, oOutDoc);
				}

				oDS_PS_MM130H.SetValue("U_DocGbn", 0, "반출");

				for (i = 1; i <= oMat.VisualRowCount; i++)
				{
					if (Convert.ToDouble(oMat.Columns.Item("UnWeight").Cells.Item(i).Specific.Value.ToString().Trim()) == 0)
					{
						oMat.Columns.Item("UnWeight").Cells.Item(i).Specific.Value = "1";
					}

					oDS_PS_MM130L.SetValue("U_OutDoc", i - 1, oOutDoc);
				}

				if (oDS_PS_MM130H.GetValue("U_OutGbn", 0).ToString().Trim() == "10")
				{
					oDS_PS_MM130H.SetValue("U_OutGbn", 0, "10");
				}
				else if (oDS_PS_MM130H.GetValue("U_OutGbn", 0).ToString().Trim() == "20")
				{
					oDS_PS_MM130H.SetValue("U_OutGbn", 0, "20");
				}

				oMat.LoadFromDataSource();

				ReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet); //메모리 해제
			}
			return ReturnValue;
		}

		/// <summary>
		/// PS_MM130_Validate
		/// </summary>
		/// <param name="ValidateType"></param>
		/// <returns></returns>
		private bool PS_MM130_Validate(string ValidateType)
		{
			bool ReturnValue = false;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			string errMessage = string.Empty;

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
					if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_MM130H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1) == "Y")
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
		/// PS_MM130_Print_Query
		/// </summary>
		[STAThread]
		private void PS_MM130_Print_Query()
		{
			string WinTitle;
			string ReportName;
			string DocEntry;
			string sQry;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

			try
			{
				DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
				sQry = "SELECT COUNT(*) AS COUNT FROM [@PS_MM130L] WHERE DocEntry = '" + DocEntry + "'";
				oRecordSet.DoQuery(sQry);

				if (Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) > 7)
				{
					WinTitle = "[PS_MM235_20] 레포트";
					ReportName = "PS_MM235_20.rpt";
				}
				else
				{
					WinTitle = "[PS_MM235_10] 레포트";
					ReportName = "PS_MM235_10.rpt";
				}

				List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>();

				//Parameter
				dataPackParameter.Add(new PSH_DataPackClass("@DocEntry", DocEntry));

				formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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

		//private bool PS_MM130_StockTrans()
		//{
		//	bool ReturnValue = false;
		//	// ERROR: Not supported in C#: OnErrorStatement

		//	int RetVal = 0;
		//	int ErrCode = 0;
		//	string ErrMsg = null;
		//	string lQuery = null;
		//	SAPbobsCOM.Recordset lRecordSet = null;
		//	SAPbobsCOM.Recordset oRecordSet = null;
		//	int lMaxBatchNum = 0;
		//	//해당 품목의 최대 배치번호
		//	double lBatchWeight = 0;
		//	//배치별 중량
		//	short lTypeCount = 0;
		//	//전체 StockInfo 구조체배열의 RowCount
		//	object Q = null;
		//	object j = null;
		//	object i = null;
		//	object K = null;
		//	object z = null;
		//	int r = 0;
		//	int DocCnt = 0;
		//	string Chk1_Val = null;
		//	string sCur_ItemCode = null;
		//	string sNxt_ItemCode = null;
		//	string sCur_TrCardCode = null;
		//	string sCur_TrOutWhs = null;
		//	string sNxt_TrOutWhs = null;
		//	string sCur_TrInWhs = null;
		//	string sNxt_TrInWhs = null;
		//	string RtnDocNum = null;
		//	SAPbobsCOM.StockTransfer oStockTrans = null;
		//	SAPbouiCOM.ProgressBar oPrgBar = null;
		//	int StockTransLineCounter = 0;

		//	lRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	string BatchNum = null;

		//	ReturnValue = true;

		//	string sQry = null;
		//	SAPbobsCOM.Recordset oRecordSet = null;
		//	oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//	//// 객체 정의 및 데이터 할당

		//	string Query02 = null;
		//	SAPbobsCOM.Recordset oRecordset02 = null;
		//	oRecordset02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//	//// 객체 정의 및 데이터 할당

		//	for (i = 0; i <= oMat.RowCount - 1; i++)
		//	{
		//		Array.Resize(ref StockInfo, lTypeCount + 1);

		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		//UPGRADE_WARNING: oMat.Columns(OutGbn).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		if (oMat.Columns.Item("OutGbn").Cells.Item(i + 1).Specific.Value == "10")
		//		{

		//			oMat.FlushToDataSource();

		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].ItemCode = oDS_PS_MM130L.GetValue("U_OutItmCd", i));
		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].FromWarehouseCode = oDS_PS_MM130L.GetValue("U_OutWhCd", i));
		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].ToWarehouseCode = oDS_PS_MM130L.GetValue("U_InWhCd", i));
		//			//        StockInfo(lTypeCount).BatchNum = Trim(oDS_PS_MM130L.GetValue("U_BatchNum", i))

		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].Weight = System.Math.Round(Convert.ToDouble(oDS_PS_MM130L.GetValue("U_OutWt", i))), 2);
		//			//Quantity
		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].UnWeight = System.Math.Round(Convert.ToDouble(oDS_PS_MM130L.GetValue("U_UnWeight", i))), 2);
		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].BatchWeight = System.Math.Round(Convert.ToDouble(oDS_PS_MM130L.GetValue("U_OutQty", i))), 2);
		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].Qty = Conversion.Val(oDS_PS_MM130L.GetValue("U_OutQty", i)));
		//			//U_Qty

		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].TransNo = oForm.Items.Item("DocEntry").Specific.Value + (i + 1);
		//			StockInfo[lTypeCount].Chk = "N";
		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].MatrixRow = (i + 1);
		//			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[lTypeCount].Indate = oForm.Items.Item("DocDate").Specific.Value;
		//			lTypeCount = lTypeCount + 1;
		//		}
		//	}

		//	for (i = 0; i <= (Information.UBound(StockInfo)); i++)
		//	{
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		StockInfo[i].StockTransDocEntry = "";
		//	}

		//	SubMain.Sbo_Company.StartTransaction();
		//	for (i = 0; i <= (Information.UBound(StockInfo)); i++)
		//	{

		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		Chk1_Val = StockInfo[i].Chk;

		//		if (Chk1_Val != "N")
		//			goto Continue_First;

		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		sCur_TrOutWhs = StockInfo[i].FromWarehouseCode;

		//		oStockTrans = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
		//		//        oStockTrans.CardCode = StockInfo(i).CardCode
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		oStockTrans.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(StockInfo[i].Indate, "&&&&-&&-&&"));
		//		oStockTrans.FromWarehouse = sCur_TrOutWhs;
		//		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		oStockTrans.Comments = "재고이전" + oForm.Items.Item("DocEntry").Specific.Value + ".";

		//		StockTransLineCounter = -1;
		//		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//		for (K = i; K <= (Information.UBound(StockInfo)); K++)
		//		{
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			Chk1_Val = StockInfo[K].Chk;

		//			if (Chk1_Val != "N")
		//				goto Continue_Second;
		//			//            sCur_TrCardCode = StockInfo(K).CardCode
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			sNxt_TrOutWhs = StockInfo[K].FromWarehouseCode;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			sCur_ItemCode = StockInfo[K].ItemCode;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			sCur_TrInWhs = StockInfo[K].ToWarehouseCode;

		//			if ((sCur_TrOutWhs != sNxt_TrOutWhs))
		//			{
		//				goto Continue_Second;
		//			}

		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			if ((i != K))
		//			{
		//				oStockTrans.Lines.Add();
		//			}
		//			StockTransLineCounter = StockTransLineCounter + 1;
		//			//---------------------------------------------------------------------------< Line >----------
		//			var _with1 = oStockTrans.Lines;

		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.ItemCode = StockInfo[K].ItemCode;
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.UserFields.Fields.Item("U_Qty").Value = Convert.ToString(StockInfo[K].Qty));
		//			//// 수량
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.UserFields.Fields.Item("U_UnWeight").Value = Convert.ToString(StockInfo[K].UnWeight));
		//			//// 단중
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.Quantity = System.Math.Round(StockInfo[K].Weight, 2);
		//			//// 수량
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			_with1.WarehouseCode = StockInfo[K].ToWarehouseCode;
		//			////ManBatchNum = 'Y' 이면 배치번호를 입력하지 않는다.
		//			//                .UserFields("U_BatchNum").Value = StockInfo(K).BatchNum
		//			//                .BatchNumbers.BatchNumber = StockInfo(K).BatchNum
		//			//                .BatchNumbers.Quantity = Round(StockInfo(K).BatchWeight, 2)
		//			//                .BatchNumbers.Notes = "재고이전(Addon)"
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[K].Chk = "Y";
		//			/// 적용한 라인에 대한 표시
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[K].StockTransDocEntry = "Checked";
		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			StockInfo[K].StockTransLineNum = Convert.ToString(StockTransLineCounter);

		//			//UPGRADE_WARNING: K 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//			for (Q = K + 1; Q <= (Information.UBound(StockInfo)); Q++)
		//			{
		//				//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				Chk1_Val = StockInfo[Q].Chk;

		//				if (Chk1_Val != "N")
		//					goto Continue_Sixth;
		//				/// 체크2 에 않된건 Skip

		//				//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				sNxt_TrOutWhs = StockInfo[Q].FromWarehouseCode;
		//				//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				sNxt_ItemCode = StockInfo[Q].ItemCode;
		//				//UPGRADE_WARNING: Q 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		//				sNxt_TrInWhs = StockInfo[Q].ToWarehouseCode;
		//			Continue_Sixth:
		//			}
		//		Continue_Second:
		//		}
		//		//---------------------------------------------------------------------------------------------

		//		RetVal = oStockTrans.Add();
		//		if (RetVal == 0)
		//		{
		//			DocCnt = DocCnt + 1;
		//			SubMain.Sbo_Company.GetNewObjectCode(out RtnDocNum);
		//			////재고이전문서번호
		//			for (r = 0; r <= Information.UBound(StockInfo); r++)
		//			{
		//				if ((StockInfo[r].StockTransDocEntry == "Checked"))
		//				{
		//					StockInfo[r].StockTransDocEntry = RtnDocNum;
		//				}
		//			}
		//			//// 데이터 업데이트
		//		}
		//		else
		//		{
		//			goto PS_MM130_StockTrans_Error;
		//		}
		//	Continue_First:
		//	}
		//	//-----------------------------------------------------------------------------------------------< First For End

		//	if ((SubMain.Sbo_Company.InTransaction))
		//	{
		//		SubMain.Sbo_Company.EndTransaction((SAPbobsCOM.BoWfTransOpt.wf_Commit));
		//	}
		//	SubMain.Sbo_Application.SetStatusBarMessage(DocCnt + " 개의 재고이전 문서가 발행되었습니다 !", SAPbouiCOM.BoMessageTime.bmt_Short, false);
		//	//UPGRADE_NOTE: oStockTrans 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oStockTrans = null;
		//	//UPGRADE_NOTE: lRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	lRecordSet = null;
		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	return ReturnValue;
		//PS_MM130_StockTrans_Error:
		//	//************Error Process************
		//	if ((SubMain.Sbo_Company.InTransaction))
		//	{
		//		SubMain.Sbo_Company.EndTransaction((SAPbobsCOM.BoWfTransOpt.wf_RollBack));
		//	}
		//	SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
		//	SubMain.Sbo_Application.SetStatusBarMessage(ErrCode + " : " + ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, true);
		//	ReturnValue = false;
		//	//UPGRADE_NOTE: oStockTrans 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oStockTrans = null;
		//	//UPGRADE_NOTE: lRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	lRecordSet = null;
		//	//UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oRecordSet = null;
		//	return ReturnValue;
		//	//************Error Process************

		//}


		//private bool PS_MM130_UpdateUserField()
		//{
		//	bool ReturnValue = false;
		//	// ERROR: Not supported in C#: OnErrorStatement

		//	int i = 0;
		//	string lQuery = null;
		//	SAPbobsCOM.Recordset lRecordSet = null;
		//	lRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//	SAPbobsCOM.Recordset RecordSet01 = null;
		//	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	oDS_PS_MM130H.SetValue("U_STDocNum", 0, (StockInfo[i].StockTransDocEntry));
		//	//    oForm.Items("StoTrDoc").Specific.Value = StockInfo(i).StockTransDocEntry

		//	ReturnValue = true;
		//	return ReturnValue;
		//PS_MM130_UpdateUserField_Error:
		//	ReturnValue = false;
		//	return ReturnValue;
		//}

		//private bool PS_MM130_Update_Cancel()
		//{
		//	bool ReturnValue = false;
		//	// ERROR: Not supported in C#: OnErrorStatement

		//	int i = 0;
		//	string lQuery = null;
		//	SAPbobsCOM.Recordset lRecordSet = null;
		//	lRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
		//	SAPbobsCOM.Recordset RecordSet01 = null;
		//	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

		//	oDS_PS_MM130H.SetValue("U_STDocNum", 0, "");
		//	//    oForm.Items("StoTrDoc").Specific.Value = StockInfo(i).StockTransDocEntry

		//	ReturnValue = true;
		//	return ReturnValue;
		//PS_MM130_Update_Cancel_Error:
		//	ReturnValue = false;
		//	return ReturnValue;
		//}

		//private bool PS_MM130_Cancel_oStockTrans(ref short ChkType)
		//{
		//	bool ReturnValue = false;
		//	// ERROR: Not supported in C#: OnErrorStatement

		//	SAPbobsCOM.StockTransfer oStockTrans = null;

		//	int i = 0;
		//	int ErrCode = 0;
		//	int ErrNum = 0;
		//	int RetVal = 0;
		//	string ErrMsg = null;

		//	string DocEntry = null;

		//	DocEntry = oDS_PS_MM130H.GetValue("U_STDocNum", 0));

		//	if (!string.IsNullOrEmpty(oDS_PS_MM130H.GetValue("U_STDocNum", 0))))
		//	{
		//		SubMain.Sbo_Company.StartTransaction();
		//		//UPGRADE_NOTE: oStockTrans 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//		oStockTrans = null;
		//		oStockTrans = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

		//		////완료
		//		if ((oStockTrans.GetByKey(Convert.ToInt32(DocEntry)) == false))
		//		{
		//			SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
		//			goto PS_MM130_Cancel_oStockTrans_Error;
		//		}
		//		RetVal = oStockTrans.Cancel();
		//		if ((0 != RetVal))
		//		{
		//			SubMain.Sbo_Company.GetLastError(out ErrCode, out ErrMsg);
		//			ErrNum = 1;
		//			goto PS_MM130_Cancel_oStockTrans_Error;
		//		}

		//		if (ChkType == 1)
		//		{
		//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
		//		}
		//		else if (ChkType == 2)
		//		{
		//			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
		//		}
		//	}

		//	//UPGRADE_NOTE: oStockTrans 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oStockTrans = null;
		//	ReturnValue = true;
		//	return ReturnValue;
		//PS_MM130_Cancel_oStockTrans_Error:
		//	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		//	//UPGRADE_NOTE: oStockTrans 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		//	oStockTrans = null;
		//	if (SubMain.Sbo_Company.InTransaction)
		//		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
		//	ReturnValue = false;
		//	if (ErrNum == 1)
		//	{
		//		MDC_Com.MDC_GF_Message(ref "PS_MM130_Cancel_oStockTrans_Error:" + ErrCode + " - " + ErrMsg, ref "E");
		//	}
		//	else
		//	{
		//		MDC_Com.MDC_GF_Message(ref "PS_MM130_Cancel_oStockTrans_Error:" + Err().Number + " - " + Err().Description, ref "E");
		//	}
		//	return ReturnValue;
		//}

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
                //    Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                //    break;
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

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "1")
					{
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_MM130_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_MM130_DataInsertCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							//외주 재고 이동 DI API
							if (oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() == "N" && string.IsNullOrEmpty(oForm.Items.Item("STDocNum").Specific.Value.ToString().Trim()))
							{
								for (i = 0; i <= oMat.RowCount - 1; i++)
								{
									if (oMat.Columns.Item("OutGbn").Cells.Item(i + 1).Specific.Value.ToString().Trim() == "10")
									{
										if (j == 0)
										{
											j += 1;
											//if (PS_MM130_StockTrans() == true)
											//{
											//	PS_MM130_UpdateUserField();
											//}
											//else
											//{
											//	PS_MM130_AddMatrixRow(oMat.VisualRowCount, false);
											//	BubbleEvent = false;
											//	return;
											//}
										}
									}
								}
							}
						}
						//외주 재고 이동 취소 DI API
						else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) && oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() == "C" & !string.IsNullOrEmpty(oForm.Items.Item("STDocNum").Specific.Value.ToString().Trim()))
						{
							if (PS_MM130_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
							//if (PS_MM130_Cancel_oStockTrans(ref 2) == true)
							//{
							//	PS_MM130_Update_Cancel();
							//}
							//else
							//{
							//	PS_MM130_AddMatrixRow(0, false);
							//	BubbleEvent = false;
							//	return;
							//}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM130_DataValidCheck() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
						}
					}
					if (pVal.ItemUID == "Print")
					{
						System.Threading.Thread thread = new System.Threading.Thread(PS_MM130_Print_Query);
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
								PS_MM130_FormItemEnabled();
								PS_MM130_AddMatrixRow(oMat.RowCount, true);
								oForm.Items.Item("Rad01").Specific.Selected = true;
								oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
								oForm.Items.Item("BPLId").Specific.Select("0", SAPbouiCOM.BoSearchKey.psk_Index);
								oForm.Items.Item("OKYNC").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
							if (pVal.ActionSuccess == true)
							{
								PS_MM130_FormItemEnabled();
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
			string RadioGrp = string.Empty;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "OrdNum")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PS_MM131 TempForm01 = new PS_MM131();

									if (oDS_PS_MM130H.GetValue("U_OutGbn", 0).ToString().Trim() == "10")
									{
										RadioGrp = "A";
									} else if (oDS_PS_MM130H.GetValue("U_OutGbn", 0).ToString().Trim() == "20") 
									{
										RadioGrp = "B";
									}

									TempForm01.LoadForm(ref oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, RadioGrp);
									PS_MM130_AddMatrixRow(0, true);
									BubbleEvent = false;
								}
							}
							else if (pVal.ColUID == "TCpCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
						}
					}

					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
					dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CardCode", "");
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
			double Count;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			string sQry;
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
							if (pVal.ColUID == "OrdNum")
							{
								oDS_PS_MM130L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM130L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_MM130_AddMatrixRow(pVal.Row, false);
								}
								oMat.LoadFromDataSource();
							}
							else if (pVal.ColUID == "OutWhCd")
							{
								oDS_PS_MM130L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_MM130L.SetValue("U_OUtWhNm", pVal.Row - 1, dataHelpClass.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" + oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
								oMat.LoadFromDataSource();
							}
							else if (pVal.ColUID == "InWhCd")
							{
								oDS_PS_MM130L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_MM130L.SetValue("U_InWhNm", pVal.Row - 1, dataHelpClass.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" + oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
								oMat.LoadFromDataSource();
							}
							else if (pVal.ColUID == "OutQty")
							{
								Count = Convert.ToDouble(oMat.Columns.Item("UnWeight").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) * Convert.ToDouble(oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_MM130L.SetValue("U_OutQty", pVal.Row - 1, oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_MM130L.SetValue("U_OutWt", pVal.Row - 1, Convert.ToString(Count));
								oMat.LoadFromDataSource();

								oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								oMat.FlushToDataSource();

								if (Convert.ToDouble(oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) > 0)
								{
									sQry = "Select U_ObasUnit FROM OITM WHERE ItemCode = '" + oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
									oRecordSet.DoQuery(sQry);

									sQry = "Select OnHand, U_Qty FROM OITW WHERE ItemCode = '" + oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "' AND WhsCode = '101'";
									oRecordSet02.DoQuery(sQry);

									if (oRecordSet.Fields.Item(0).Value.ToString().Trim().Substring(0, 1) == "1")
									{
										if (Convert.ToDouble(oRecordSet02.Fields.Item(0).Value.ToString().Trim()) > 0 && !string.IsNullOrEmpty(oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
										{
											oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
										}
										else
										{
											oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
										}
									}
									else if (oRecordSet.Fields.Item(0).Value.ToString().Trim().Substring(0, 1) == "2")
									{
										if (Convert.ToDouble(oRecordSet02.Fields.Item(1).Value.ToString().Trim()) > 0 & !string.IsNullOrEmpty(oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
										{
											oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value = Convert.ToString((Convert.ToDouble(oRecordSet02.Fields.Item(0).Value.ToString().Trim()) / Convert.ToDouble(oRecordSet02.Fields.Item(1).Value.ToString().Trim())) * Convert.ToDouble(oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()));
										}
										else
										{
											oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value = oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim();
										}
									}

									oDS_PS_MM130L.SetValue("U_OutQty", pVal.Row - 1, oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									oDS_PS_MM130L.SetValue("U_OutWt", pVal.Row - 1, oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									oMat.LoadFromDataSource();

									oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}
							}
							else if (pVal.ColUID == "TCpCode")
							{
								sQry = "Select U_CpName from [@PS_PP001L] Where U_CpCode = '" + oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								oDS_PS_MM130L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_MM130L.SetValue("U_TCpName", pVal.Row - 1, oRecordSet.Fields.Item(0).Value.ToString().Trim());
								oMat.LoadFromDataSource();
							}
							else
							{
								oDS_PS_MM130L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oMat.LoadFromDataSource();
							}

							oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						}
						else if (pVal.ItemUID == "CardCode")
						{
							sQry = "SELECT CardName FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("CardName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "CntcCode")
						{
							sQry = "Select U_FULLNAME, U_MSTCOD From [OHEM] Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);

							oForm.Items.Item("CntcName").Specific.Value = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else
						{
							if (pVal.ItemUID == "DocEntry")
							{
								oDS_PS_MM130H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
							}
						}

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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet02);
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
					PS_MM130_FormItemEnabled();
					PS_MM130_AddMatrixRow(oMat.VisualRowCount, false);
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
			SAPbouiCOM.DataTable oDataTable02 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects;

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
							oDS_PS_MM130H.SetValue("U_CardCode", 0, oDataTable01.Columns.Item("CardCode").Cells.Item(0).Value.ToString().Trim());
							oDS_PS_MM130H.SetValue("U_CardName", 0, oDataTable01.Columns.Item("CardName").Cells.Item(0).Value.ToString().Trim());
							// 찾기나 문서이동 버튼 클릭 시에 갱신으로 바뀌지 않음
							if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							}
						}
					}
					else if (pVal.ItemUID == "ShipTo")
					{
						if (oDataTable02 == null)
						{
						}
						else
						{
							oDS_PS_MM130H.SetValue("U_ShipTo", 0, oDataTable02.Columns.Item("CardCode").Cells.Item(0).Value.ToString().Trim());
							oDS_PS_MM130H.SetValue("U_ShipNm", 0, oDataTable02.Columns.Item("CardName").Cells.Item(0).Value.ToString().Trim());
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM130H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM130L);
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
						oDS_PS_MM130L.RemoveRecord(oDS_PS_MM130L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_MM130_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_MM130L.GetValue("U_OrdNum", oMat.RowCount - 1).ToString().Trim())) 
							{
								PS_MM130_AddMatrixRow(oMat.RowCount, false);
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
							if (PS_MM130_Validate("취소") == false)
							{
								BubbleEvent = false;
								return;
							}
							if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
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
							PS_MM130_FormItemEnabled();
							break;
						case "1282": //추가
							oDS_PS_MM130H.SetValue("U_OutGbn", 0, "10");
							oDS_PS_MM130H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
							oDS_PS_MM130H.SetValue("U_BPLId", 0, "1");
							oDS_PS_MM130H.SetValue("U_OKYNC", 0, "N");
							PS_MM130_FormItemEnabled();
							PS_MM130_AddMatrixRow(0, true);
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
