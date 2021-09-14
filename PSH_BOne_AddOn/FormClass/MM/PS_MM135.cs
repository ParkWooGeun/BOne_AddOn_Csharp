using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 포장사업팀 외주반출등록
	/// </summary>
	internal class PS_MM135 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat;
		private SAPbouiCOM.DBDataSource oDS_PS_MM135H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_MM135L; //등록라인
		private string oLastItemUID01;  //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01;   //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01;      //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

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
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_MM135.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_MM135_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_MM135");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry"; 

				oForm.Freeze(true);

				PS_MM135_CreateItems();
				PS_MM135_SetComboBox();
				PS_MM135_Initialize();
				PS_MM135_CF_ChooseFromList();
				PS_MM135_EnableMenus();
				PS_MM135_SetDocument(oFormDocEntry);
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
		/// PS_MM135_CreateItems
		/// </summary>
		private void PS_MM135_CreateItems()
		{
			try
			{
				oDS_PS_MM135H = oForm.DataSources.DBDataSources.Item("@PS_MM135H");
				oDS_PS_MM135L = oForm.DataSources.DBDataSources.Item("@PS_MM135L");
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

				oForm.Settings.MatrixUID = "Mat01";	// 서식세팅
				oForm.Settings.Enabled = true;
				oForm.Settings.EnableRowFormat = true;

				oForm.Items.Item("DocDate").Specific.Value  = DateTime.Now.ToString("yyyyMMdd");

				oForm.DataSources.UserDataSources.Add("BOM_CHECK", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
				oForm.Items.Item("BOM_CHECK").Specific.DataBind.SetBound(true, "", "BOM_CHECK");
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM135_SetComboBox
		/// </summary>
		private void PS_MM135_SetComboBox()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				dataHelpClass.Combo_ValidValues_Insert("PS_MM135", "Mat01", "OutGbn", "10", "제품SET기준");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM135", "Mat01", "OutGbn", "20", "원재료기준");
				dataHelpClass.Combo_ValidValues_SetValueColumn(oMat.Columns.Item("OutGbn"), "PS_MM135", "Mat01", "OutGbn",false);

				dataHelpClass.Combo_ValidValues_Insert("PS_MM135", "OKYNC", "", "N", "재고이동");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM135", "OKYNC", "", "C", "이동취소");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM135", "OKYNC", "", "Y", "승인");
				dataHelpClass.Combo_ValidValues_Insert("PS_MM135", "OKYNC", "", "B", "반품");

				dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OKYNC").Specific, "PS_MM135", "OKYNC", false);
				oForm.Items.Item("OKYNC").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

				dataHelpClass.Set_ComboList(oForm.Items.Item("InWhCd").Specific, "SELECT [WhsCode], [WhsName] FROM OWHS", "803", false, false);

				dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL WHERE BPLId = '3' ORDER BY BPLId", "3", false, false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM135_Initialize
		/// </summary>
		private void PS_MM135_Initialize()
		{
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			string lcl_User_BPLId;

			try
			{
				// 사업장
				lcl_User_BPLId = dataHelpClass.User_BPLID();
				if (lcl_User_BPLId == "3")  // '3'포장 일때만
				{
					oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
				}
				oForm.Items.Item("CntcCode").Specific.Value  = dataHelpClass.User_MSTCOD(); // 인수자
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM135_CF_ChooseFromList
		/// </summary>
		private void PS_MM135_CF_ChooseFromList()
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

				oCFLCreationParams02.ObjectType = "2";
				oCFLCreationParams02.UniqueID = "CFLSHIPCODE";
				oCFLCreationParams02.MultiSelection = false;
				oCFL02 = oCFLs02.Add(oCFLCreationParams02);

				oEdit02.ChooseFromListUID = "CFLSHIPCODE";
				oEdit02.ChooseFromListAlias = "CardCode";
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
		/// PS_MM135_EnableMenus
		/// </summary>
		private void PS_MM135_EnableMenus()
		{
			try
			{	
				oForm.EnableMenu("1283", false);
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM135_SetDocument
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		private void PS_MM135_SetDocument(string oFormDocEntry)
		{
			try
			{
				if (string.IsNullOrEmpty(oFormDocEntry))
				{
					PS_MM135_EnableFormItem();
					PS_MM135_AddMatrixRow(0, true); 
				}
				else
				{
					oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
					PS_MM135_EnableFormItem();
					oForm.Items.Item("DocEntry").Specific.Value  = oFormDocEntry;
					oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM135_EnableFormItem
		/// </summary>
		private void PS_MM135_EnableFormItem()
		{
			try
			{
				oForm.Freeze(true);

				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_MM135_ClearForm(); 
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", false); //추가
					oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.EnableMenu("1293", true); // 행삭제
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("OutDoc").Enabled = false;
					oForm.Items.Item("BOM_CHECK").Specific.Checked = true;

					if (oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "C" || oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
					{
						oForm.EnableMenu("1284", false); //취소
						oForm.EnableMenu("1286", false); //닫기
						oForm.EnableMenu("1293", false); //행삭제
					}
					else
					{
						oForm.EnableMenu("1284", true);  //취소
						oForm.EnableMenu("1286", true);  //닫기
						oForm.EnableMenu("1293", true);  //행삭제
					}

					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("CntcCode").Enabled = true;
					oForm.Items.Item("ShipTo").Enabled = true;
					oForm.Items.Item("CarNo").Enabled = true;
					oForm.Items.Item("ShipCo").Enabled = true;
					oForm.Items.Item("DocDate").Enabled = true;
					oForm.Items.Item("Fare").Enabled = true;
					oForm.Items.Item("ArrivePl").Enabled = true;
					oForm.Items.Item("InWhCd").Enabled = true;
					oMat.Columns.Item("OutItmCd").Editable = true;
					oMat.Columns.Item("OutQty").Editable = true;
					oMat.Columns.Item("OutWt").Editable = true;
					oMat.Columns.Item("OutWhCd").Editable = true;
					oMat.Columns.Item("OutWhNm").Editable = true;
					oMat.Columns.Item("InWhCd").Editable = true;
					oMat.Columns.Item("InWhNm").Editable = false;
					oForm.Items.Item("Rad01").Enabled = false;
					oForm.Items.Item("Rad02").Enabled = false;
					oForm.Items.Item("OKYNC").Enabled = true;
				}
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
				{
					oForm.EnableMenu("1281", true);     //찾기
					oForm.EnableMenu("1282", true);     //추가
					oForm.EnableMenu("1293", true); // 행삭제

					if (oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "C")
					{
						oForm.EnableMenu("1284", false);  //취소
						oForm.EnableMenu("1286", false);  //닫기
						oForm.EnableMenu("1293", false);  //행삭제
					}
					else if (oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
					{
						if (!string.IsNullOrEmpty(oDS_PS_MM135H.GetValue("U_STDocNum", 0).ToString().Trim()))
						{
							oForm.EnableMenu("1284", false);  //취소
							oForm.EnableMenu("1286", false);  //닫기
							oForm.EnableMenu("1293", false);  //행삭제
						}
						else
						{
							oForm.EnableMenu("1284", true);  //취소
							oForm.EnableMenu("1286", true);  //닫기
							oForm.EnableMenu("1293", true);  //행삭제
						}
					}
					else
					{
						oForm.EnableMenu("1284", true);  //취소
						oForm.EnableMenu("1286", true);  //닫기
						oForm.EnableMenu("1293", true);  //행삭제
					}

					oForm.Items.Item("BPLId").Enabled = true;
					oForm.Items.Item("CardCode").Enabled = true;
					oForm.Items.Item("DocEntry").Enabled = true;
					oForm.Items.Item("OutDoc").Enabled = true;
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
				else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
				{
					oForm.EnableMenu("1281", true);  //찾기
					oForm.EnableMenu("1282", true);  //추가
					oForm.EnableMenu("1293", true);  // 행삭제

					if (oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "C")
					{
						oForm.EnableMenu("1284", false);  //취소
						oForm.EnableMenu("1286", false);  //닫기
						oForm.EnableMenu("1293", false);  //행삭제
					}
					else if (oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y")
					{
						if (!string.IsNullOrEmpty(oDS_PS_MM135H.GetValue("U_STDocNum", 0).ToString().Trim()))
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
					}
					else
					{
						oForm.EnableMenu("1284", true);  //취소
						oForm.EnableMenu("1286", true);  //닫기
						oForm.EnableMenu("1293", true);  //행삭제
					}

					oForm.Items.Item("BPLId").Enabled = false;
					oForm.Items.Item("Comments").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					oForm.Items.Item("OutDoc").Enabled = false;
					oForm.Items.Item("DocEntry").Enabled = false;
					oForm.Items.Item("Rad01").Enabled = false;
					oForm.Items.Item("Rad02").Enabled = false;
					oForm.Items.Item("DocEntry").Enabled = false;

					if (oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "C"
						|| oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y"
						|| oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "B")
					{
						if (oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "N")
						{
							oForm.Items.Item("OKYNC").Enabled = true;
						}
						else if (oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "C" 
							     || oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "Y" 
							     || oDS_PS_MM135H.GetValue("U_OKYNC", 0).ToString().Trim() == "B")
						{
							oForm.Items.Item("OKYNC").Enabled = false; //취소, 승인, 반품시 비활성화
						}
						oForm.Items.Item("CardCode").Enabled = false;
						oForm.Items.Item("CntcCode").Enabled = false;
						oForm.Items.Item("ShipTo").Enabled = false;
						oForm.Items.Item("CarNo").Enabled = false;
						oForm.Items.Item("ShipCo").Enabled = false;
						oForm.Items.Item("DocDate").Enabled = false;
						oForm.Items.Item("Fare").Enabled = false;
						oForm.Items.Item("ArrivePl").Enabled = false;
						oForm.Items.Item("InWhCd").Enabled = false;
						oMat.Columns.Item("OutQty").Editable = false;
						oMat.Columns.Item("OutWt").Editable = false;
						oMat.Columns.Item("OutWhCd").Editable = false;
						oMat.Columns.Item("OutWhNm").Editable = false;
						oMat.Columns.Item("InWhCd").Editable = true;
						oMat.Columns.Item("InWhNm").Editable = false;
					}
					else
					{
						oForm.Items.Item("Rad01").Enabled = false;
						oForm.Items.Item("Rad02").Enabled = false;
						oForm.Items.Item("CardCode").Enabled = true;
						oForm.Items.Item("CntcCode").Enabled = true;
						oForm.Items.Item("ShipTo").Enabled = true;
						oForm.Items.Item("CarNo").Enabled = true;
						oForm.Items.Item("ShipCo").Enabled = true;
						oForm.Items.Item("DocDate").Enabled = true;
						oForm.Items.Item("Fare").Enabled = true;
						oForm.Items.Item("ArrivePl").Enabled = true;
						oForm.Items.Item("InWhCd").Enabled = true;
						oMat.Columns.Item("OutItmCd").Editable = false;
						oMat.Columns.Item("OutQty").Editable = false;
						oMat.Columns.Item("OutWt").Editable = false;
						oMat.Columns.Item("OutWhCd").Editable = false;
						oMat.Columns.Item("OutWhNm").Editable = false;
						oMat.Columns.Item("InWhCd").Editable = false;
						oMat.Columns.Item("InWhNm").Editable = false;
						oForm.Items.Item("OKYNC").Enabled = true;
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

		/// <summary>
		/// PS_MM135_AddMatrixRow
		/// </summary>
		/// <param name="oRow"></param>
		/// <param name="RowIserted"></param>
		private void PS_MM135_AddMatrixRow(int oRow, bool RowIserted)
		{
			try
			{
				oForm.Freeze(true);
				if (RowIserted == false)
				{
					oDS_PS_MM135L.InsertRecord(oRow);
				}
				oMat.AddRow();
				oDS_PS_MM135L.Offset = oRow;
				oDS_PS_MM135L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
		/// PS_MM135_ClearForm
		/// </summary>
		private void PS_MM135_ClearForm()
		{
			string DocEntry;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
			
			try
			{
				DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_MM135'", "");
				if (Convert.ToDouble(DocEntry) == 0)
				{
					oForm.Items.Item("DocEntry").Specific.Value  = 1;
				}
				else
				{
					oForm.Items.Item("DocEntry").Specific.Value  = DocEntry;
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
		}

		/// <summary>
		/// PS_MM135_CheckDataValid
		/// </summary>
		private bool PS_MM135_CheckDataValid()
		{
			bool functionReturnValue = false;
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
				else if (string.IsNullOrEmpty(oForm.Items.Item("CpCode").Specific.Value.ToString().Trim()) && oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() != "B")
				{
					oForm.Items.Item("CpCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
					errMessage = "공정코드는 필수입니다.";
					throw new Exception();
				}
				for (i = 1; i <= oMat.VisualRowCount - 1; i++)
				{
					if (Convert.ToDouble(oMat.Columns.Item("OutQty").Cells.Item(i).Specific.Value.ToString().Trim()) < 0 
						|| Convert.ToDouble(oMat.Columns.Item("OutWt").Cells.Item(i).Specific.Value.ToString().Trim()) < 0 )
					{
						oMat.Columns.Item("OutQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
						errMessage = "반출수(중)량은 필수입니다.";
						throw new Exception();
					}
				}

				oDS_PS_MM135L.RemoveRecord(oDS_PS_MM135L.Size - 1);
				oMat.LoadFromDataSource();
				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
				{
					PS_MM135_ClearForm();
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
		/// PS_MM135_CanceloStockTrans
		/// </summary>
		/// <param name="ChkType"></param>
		/// <returns></returns>
		private bool PS_MM135_CanceloStockTrans(int ChkType)
		{
			bool functionReturnValue = false;
			SAPbobsCOM.StockTransfer oStockTrans = null;
			string ErrMsg;
			string DocEntry;
			string errMessage = string.Empty;
			int ErrCode;
			int RetVal;

			try
			{
				DocEntry = oDS_PS_MM135H.GetValue("U_STDocNum", 0).ToString().Trim();

				if (!string.IsNullOrEmpty(oDS_PS_MM135H.GetValue("U_STDocNum", 0).ToString().Trim()))
				{
					PSH_Globals.oCompany.StartTransaction();
					oStockTrans = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

					//완료
					if (oStockTrans.GetByKey(Convert.ToInt32(DocEntry)) == false)
					{
						PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
						throw new Exception();
					}
					RetVal = oStockTrans.Cancel();
					if (0 != RetVal)
					{
						PSH_Globals.oCompany.GetLastError(out ErrCode, out ErrMsg);
						errMessage = "DI실행 중 오류 발생";
						throw new Exception();
					}

					if (ChkType == 1)
					{
						PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
					}
					else if (ChkType == 2)
					{
						PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
					}
				}
				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				if (PSH_Globals.oCompany.InTransaction)
				{
					PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
				}

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
				if (oStockTrans != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oStockTrans);
				}
            }

			return functionReturnValue;
		}

		/// <summary>
		/// PS_MM135_CheckDataInsert
		/// </summary>
		/// <returns></returns>
		private bool PS_MM135_CheckDataInsert()
		{
			bool functionReturnValue = false;
			string sQry;
			string oDate;
			string oOutDoc;
			int i;
			SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			try
			{
				oDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
				sQry = "SELECT ISNULL(MAX(U_OutDoc), 0) FROM [@PS_MM135H] WHERE SubString(U_OutDoc,1,8) = '" + oDate + "'";
				oRecordSet.DoQuery(sQry);

				oOutDoc = oRecordSet.Fields.Item(0).Value.ToString().Trim();

				if (Convert.ToDouble(oOutDoc) == 0)
				{
					oOutDoc = oDate + "001";
					oDS_PS_MM135H.SetValue("U_OutDoc", 0, oOutDoc);
				}
				else
				{
					oOutDoc = Convert.ToString(Convert.ToDouble(oOutDoc) + 1);
					oDS_PS_MM135H.SetValue("U_OutDoc", 0, oOutDoc);
				}

				oDS_PS_MM135H.SetValue("U_DocGbn", 0, "반출");

				for (i = 1; i <= oMat.VisualRowCount; i++)
				{
					oDS_PS_MM135L.SetValue("U_OutDoc", i - 1, oOutDoc);
				}

				if (oDS_PS_MM135H.GetValue("U_OutGbn", 0).ToString().Trim() == "10")
				{
					oDS_PS_MM135H.SetValue("U_OutGbn", 0, "10");
				}
				else if (oDS_PS_MM135H.GetValue("U_OutGbn", 0).ToString().Trim() == "20")
				{
					oDS_PS_MM135H.SetValue("U_OutGbn", 0, "20");
				}
				oMat.LoadFromDataSource();

				functionReturnValue = true;
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
            {
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
			}
			return functionReturnValue;
		}

		/// <summary>
		/// PS_MM135_Validate
		/// </summary>
		/// <param name="ValidateType"></param>
		/// <returns></returns>
		private bool PS_MM135_Validate(string ValidateType)
		{
			bool functionReturnValue = false;

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
					if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_MM135H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1) == "Y")
					{
						errMessage = "이미취소된 문서 입니다. 취소할수 없습니다.";
						throw new Exception();
					}
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
		/// PS_MM135_UpdateCancel
		/// </summary>
		private void PS_MM135_UpdateCancel()
		{
			try
			{
				oDS_PS_MM135H.SetValue("U_STDocNum", 0, "");
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
		private void Raise_EVENT_ROW_DELETE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
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
						oDS_PS_MM135L.RemoveRecord(oDS_PS_MM135L.Size - 1);
						oMat.LoadFromDataSource();
						if (oMat.RowCount == 0)
						{
							PS_MM135_AddMatrixRow(0, false);
						}
						else
						{
							if (!string.IsNullOrEmpty(oDS_PS_MM135L.GetValue("U_ItemCode", oMat.RowCount - 1).ToString().Trim()))
							{
								PS_MM135_AddMatrixRow(oMat.RowCount, false);
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
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_CLICK: //6
                    //Raise_EVENT_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK: //7
                    //Raise_EVENT_DOUBLE_CLICK(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED: //8
                    Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_COLLAPSE_PRESSED: //9
                    //Raise_EVENT_MATRIX_COLLAPSE_PRESSED(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_VALIDATE: //10
                    Raise_EVENT_VALIDATE(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD: //11
                    Raise_EVENT_MATRIX_LOAD(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_CHOOSE_FROM_LIST(FormUID, ref pVal, ref BubbleEvent);
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
						if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
						{
							if (PS_MM135_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}

							if (PS_MM135_CheckDataInsert() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
							      && oForm.Items.Item("OKYNC").Specific.Value.ToString().Trim() == "C" 
								  && !string.IsNullOrEmpty(oForm.Items.Item("STDocNum").Specific.Value.ToString().Trim()))
						{
							if (PS_MM135_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
							if (PS_MM135_CanceloStockTrans(2) == true)
							{
								PS_MM135_UpdateCancel();
							}
							else
							{
								PS_MM135_AddMatrixRow(0, false);
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
						{
							if (PS_MM135_CheckDataValid() == false)
							{
								BubbleEvent = false;
								return;
							}
						}
						else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
						{
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
								PS_MM135_EnableFormItem();
								PS_MM135_AddMatrixRow(oMat.RowCount, true); 
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
								PS_MM135_EnableFormItem();
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
		/// Raise_EVENT_KEY_DOWN
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_KEY_DOWN(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			string ItemCode;
			PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

			try
			{
				oForm.Freeze(true);

				if (pVal.BeforeAction == true)
				{
					if (pVal.CharPressed == 9)
					{
						if (pVal.ItemUID == "Mat01")
						{
							if (pVal.ColUID == "ItemCode")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									ItemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();

									if (oForm.Items.Item("BOM_CHECK").Specific.Checked == true)
									{
										if (!string.IsNullOrEmpty(oForm.Items.Item("Qty").Specific.Value.ToString().Trim())
											&& !string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
										{
											PS_SM030 oTempClass = new PS_SM030(); //포장제품BOM
											oTempClass.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, ItemCode);
										}
										else
										{
											PSH_Globals.SBO_Application.SetStatusBarMessage("반출품목코드 또는 반출제품수량이 없습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
										}
									}
									else
									{
										PSH_Globals.SBO_Application.ActivateMenuItem("7425");
										BubbleEvent = false;
									}
								}
							}
							else if (pVal.ColUID == "OutItmCd")
							{
								if (string.IsNullOrEmpty(oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									PSH_Globals.SBO_Application.ActivateMenuItem("7425");
									BubbleEvent = false;
								}
							}
						}
						if (pVal.ItemUID == "ItemCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						if (pVal.ItemUID == "ItmGrpCd")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("ItmGrpCd").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
							}
						}
						if (pVal.ItemUID == "CpCode")
						{
							if (string.IsNullOrEmpty(oForm.Items.Item("CpCode").Specific.Value.ToString().Trim()))
							{
								PSH_Globals.SBO_Application.ActivateMenuItem("7425");
								BubbleEvent = false;
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
            {
				oForm.Freeze(false);
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
			int i;

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
						if (pVal.ItemUID == "InWhCd")
						{
							for (i = 1; i <= oMat.VisualRowCount - 1; i++)
							{
								oMat.Columns.Item("InWhCd").Cells.Item(i).Specific.Value  = oForm.Items.Item("InWhCd").Specific.Value.ToString().Trim();
							}
						}
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

		/// <summary>
		/// Raise_EVENT_MATRIX_LINK_PRESSED
		/// </summary>
		/// <param name="FormUID"></param>
		/// <param name="pVal"></param>
		/// <param name="BubbleEvent"></param>
		private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
		{
			try
			{
				if (pVal.BeforeAction == true)
				{
					if (pVal.ItemUID == "Mat01")
					{
						if (pVal.ColUID == "ItemCode")
						{
							if (!string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.String().Trim()))
							{
								PS_MM002 oTempClass = new PS_MM002();
								oTempClass.LoadForm(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								BubbleEvent = false;
							}
							else
							{
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
			int Qty;
			string OutGbn;
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
							if (pVal.ColUID == "ItemCode")
							{
								if (string.IsNullOrEmpty(oForm.Items.Item("Qty").Specific.Value.ToString().Trim()))
								{
									Qty = 0;
								}
								else
								{
									Qty = Convert.ToInt32(oForm.Items.Item("Qty").Specific.Value.ToString().Trim());
								}

								oDS_PS_MM135L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM135L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
								{
									PS_MM135_AddMatrixRow(pVal.Row, false);
								}

								sQry = "Select FrgnName, U_Size From OITM Where ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								oDS_PS_MM135L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet.Fields.Item("FrgnName").Value.ToString().Trim());
								oDS_PS_MM135L.SetValue("U_Size", pVal.Row - 1, oRecordSet.Fields.Item("U_Size").Value.ToString().Trim());
								if (string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									oDS_PS_MM135L.SetValue("U_Qty", pVal.Row - 1, "0");
								}
								else
								{
									oDS_PS_MM135L.SetValue("U_Qty", pVal.Row - 1, Convert.ToString(Qty));
								}
								oMat.LoadFromDataSource();
							}
							else if (pVal.ColUID == "OutItmCd")
							{
								OutGbn = oDS_PS_MM135H.GetValue("U_OutGbn", 0);
								oDS_PS_MM135L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								if (oForm.Items.Item("OKYNC").Specific.Value  == "B")
								{
									if (oMat.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_MM135L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
									{
										PS_MM135_AddMatrixRow(pVal.Row, false);
									}
								}
								sQry = "Select ItemName From OITM Where ItemCode = '" + oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
								oRecordSet.DoQuery(sQry);

								oDS_PS_MM135L.SetValue("U_OutItmNm", pVal.Row - 1, oRecordSet.Fields.Item("ItemName").Value.ToString().Trim());
								oDS_PS_MM135L.SetValue("U_OutGbn", pVal.Row - 1, OutGbn);

								oMat.Columns.Item("OutWhCd").Cells.Item(pVal.Row).Specific.Value  = "10" + oDS_PS_MM135H.GetValue("U_BPLId", 0).ToString().Trim();
								oMat.Columns.Item("InWhCd").Cells.Item(pVal.Row).Specific.Value  = oDS_PS_MM135H.GetValue("U_InWhCd", 0).ToString().Trim();

								if (string.IsNullOrEmpty(oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()))
								{
									oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
								}
								oMat.LoadFromDataSource();
							}
							else if (pVal.ColUID == "OutWhCd")
							{
								oDS_PS_MM135L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oDS_PS_MM135L.SetValue("U_OUtWhNm", pVal.Row - 1, dataHelpClass.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" + oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
								oMat.LoadFromDataSource();
							}
							else if (pVal.ColUID == "InWhCd")
							{
								oDS_PS_MM135L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value );
								oDS_PS_MM135L.SetValue("U_InWhNm", pVal.Row - 1, dataHelpClass.Get_ReData("WhsName", "WhsCode", "[OWHS]", "'" + oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'", ""));
								oMat.LoadFromDataSource();
							}
							else if (pVal.ColUID == "OutQty")
							{
								oDS_PS_MM135L.SetValue("U_OutQty", pVal.Row - 1, oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oMat.LoadFromDataSource();
								oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								oMat.FlushToDataSource();
								if (Convert.ToDouble(oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()) > 0)
								{
									sQry = "Select (b.U_Weight / b.U_Qty) FROM [@PS_MM002H] a Inner Join [@PS_MM002L] b On a.Code = b.Code WHERE a.U_ItemCode = '" + oMat.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "' and b.U_MItemCod = '" + oMat.Columns.Item("OutItmCd").Cells.Item(pVal.Row).Specific.Value.ToString().Trim() + "'";
									oRecordSet.DoQuery(sQry);

									oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value = System.Math.Round(Convert.ToDouble(oRecordSet.Fields.Item(0).Value.ToString().Trim()) * Convert.ToDouble(oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim()), 2);
									oDS_PS_MM135L.SetValue("U_OutQty", pVal.Row - 1, oMat.Columns.Item("OutQty").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									oDS_PS_MM135L.SetValue("U_OutWt", pVal.Row - 1, oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
									oMat.LoadFromDataSource();
									oMat.Columns.Item("OutWt").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
								}
							}
							else
							{
								oDS_PS_MM135L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Trim());
								oMat.LoadFromDataSource();
								oMat.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
							}
						}
						else if (pVal.ItemUID == "ItemCode")
						{
							sQry = "SELECT ItemName FROM OITM WHERE ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("ItemName").Specific.Value  = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "ItmGrpCd")
						{
							sQry = "SELECT U_CdName FROM [@PS_SY001L] WHERE Code ='M007' and U_Minor = '" + oForm.Items.Item("ItmGrpCd").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("ItmGrpNm").Specific.Value  = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "CardCode")
						{
							sQry = "SELECT CardName FROM [OCRD] WHERE CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CardName").Specific.Value  = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "CntcCode")
						{
							sQry = "Select U_FULLNAME, U_MSTCOD From [OHEM] Where U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CntcName").Specific.Value  = oRecordSet.Fields.Item(0).Value.ToString().Trim();
						}
						else if (pVal.ItemUID == "CpCode")
						{
							sQry = "SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode =  '" + oForm.Items.Item("CpCode").Specific.Value.ToString().Trim() + "'";
							oRecordSet.DoQuery(sQry);
							oForm.Items.Item("CpName").Specific.String = oRecordSet.Fields.Item("U_CpName").Value.ToString().Trim();
						}
						else
						{
							if (pVal.ItemUID == "DocEntry")
							{
								oDS_PS_MM135H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value.ToString().Trim());
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
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
					PS_MM135_EnableFormItem();
					PS_MM135_AddMatrixRow(oMat.VisualRowCount, false);
					oMat.AutoResizeColumns();
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
			SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당
			SAPbouiCOM.DataTable oDataTable02 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

			try
			{
				if (pVal.Before_Action == true)
				{
				}
				else if (pVal.Before_Action == false)
				{
					if (oDataTable01 != null) //SelectedObjects 가 null이 아닐때만 실행(ChooseFromList 팝업창을 취소했을 때 미실행)
					{
						if (pVal.ItemUID == "CardCode")
						{
							oDS_PS_MM135H.SetValue("U_CardCode", 0, oDataTable01.Columns.Item("CardCode").Cells.Item(0).Value.ToString().Trim());
							oDS_PS_MM135H.SetValue("U_CardName", 0, oDataTable01.Columns.Item("CardName").Cells.Item(0).Value.ToString().Trim());
							if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) // 찾기나 문서이동 버튼 클릭 시에 갱신으로 바뀌지 않음
							{
								oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							}
						}
					}
					if (oDataTable02 != null) //SelectedObjects 가 null이 아닐때만 실행(ChooseFromList 팝업창을 취소했을 때 미실행)
					{
						if (pVal.ItemUID == "ShipTo")
						{
							oDS_PS_MM135H.SetValue("U_ShipTo", 0, oDataTable02.Columns.Item("CardCode").Cells.Item(0).Value.ToString().Trim());
							oDS_PS_MM135H.SetValue("U_ShipNm", 0, oDataTable02.Columns.Item("CardName").Cells.Item(0).Value.ToString().Trim());
							if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
							{
								oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
							}
						}
					}
				}
				oForm.Update();
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
			}
			finally
            {
				if (oDataTable01 != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
				}
				if (oDataTable02 != null)
				{
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable02);
				}
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
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM135H);
					System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_MM135L);
				}
			}
			catch (Exception ex)
			{
				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
						case "1284": //취소
							if (PS_MM135_Validate("취소") == false)
							{
								BubbleEvent = false;
								return;
							}
							break;
						case "1286": //닫기
							break;
						case "1293": //행삭제
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
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
							Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
							break;
						case "1281": //찾기
							PS_MM135_EnableFormItem(); 
							break;
						case "1282": //추가
							oForm.Freeze(true);
							oDS_PS_MM135H.SetValue("U_OutGbn", 0, "10");
							oDS_PS_MM135H.SetValue("U_DocDate", 0, DateTime.Now.ToString("yyyyMMdd"));
							oDS_PS_MM135H.SetValue("U_BPLId", 0, "3");
							oDS_PS_MM135H.SetValue("U_OKYNC", 0, "Y");
							oDS_PS_MM135H.SetValue("U_InWhCd", 0, "803");
							PS_MM135_EnableFormItem(); 
							PS_MM135_AddMatrixRow(0, true); 
							oForm.Freeze(false);
							break;
						case "1288":
						case "1289":
						case "1290":
						case "1291": //레코드이동버튼
							PS_MM135_EnableFormItem();
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
