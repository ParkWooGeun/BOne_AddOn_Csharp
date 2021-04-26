using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 납품처리
	/// </summary>
	internal class PS_SD040 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD040H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD040L; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

		public class ItemInformation
		{
			public string ItemCode;
			public string BatchNum;
			public int Qty; //수량
			public double Weight; //중량
			public string Currency; //통화
			public double Price; //단가
			public double LineTotal; //총계
			public string WhsCode; //창고
			public int SD030HNum; //출하(선출)문서
			public int SD030LNum; //출하(선출)라인
			public int SD040HNum; //납품A문서
			public int SD040LNum; //납품A라인
			public int ORDRNum; //판매오더문서
			public int RDR1Num; //판매오더라인
			public bool Check; 
			public int ODLNNum; //납품문서 
			public int DLN1Num; //납품라인
			public int ORDNNum; //반품문서
			public int RDN1Num; //반품라인
		}

		public class BatchInformation
		{
			public string ItemCode; //품목코드
			public string WhsCode; //창고코드
			public string BatchNum; //배치번호
			public double Weight; //중량
		}

		/// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
		public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD040.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD040_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD040");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";
				
				oForm.Freeze(true);

				PS_SD040_CreateItems();
				PS_SD040_SetComboBox();
				PS_SD040_CF_ChooseFromList();
				PS_SD040_EnableMenus();
				PS_SD040_SetDocument(oFormDocEntry);
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
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
			}
		}

        /// <summary>
        /// 화면 Item 생성
        /// </summary>
        private void PS_SD040_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PS_SD040H = oForm.DataSources.DBDataSources.Item("@PS_SD040H");
                oDS_PS_SD040L = oForm.DataSources.DBDataSources.Item("@PS_SD040L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("SumSjQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumSjQty").Specific.DataBind.SetBound(true, "", "SumSjQty");

                oForm.DataSources.UserDataSources.Add("SumSjWt", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumSjWt").Specific.DataBind.SetBound(true, "", "SumSjWt");

                oForm.DataSources.UserDataSources.Add("SumQty", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("SumQty").Specific.DataBind.SetBound(true, "", "SumQty");

                oForm.DataSources.UserDataSources.Add("SumWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("SumWeight").Specific.DataBind.SetBound(true, "", "SumWeight");

                oForm.DataSources.UserDataSources.Add("HWeight", SAPbouiCOM.BoDataType.dt_QUANTITY);
                oForm.Items.Item("HWeight").Specific.DataBind.SetBound(true, "", "HWeight");

                oForm.Items.Item("Opt01").Specific.ValOn = "2";
                oForm.Items.Item("Opt01").Specific.ValOff = "0";
                oForm.Items.Item("Opt01").Specific.Selected = true;

                oForm.Items.Item("Opt02").Specific.ValOn = "1";
                oForm.Items.Item("Opt02").Specific.ValOff = "0";
                oForm.Items.Item("Opt02").Specific.GroupWith("Opt01");
                
                oForm.Items.Item("Opt03").Specific.ValOn = "3";
                oForm.Items.Item("Opt03").Specific.ValOff = "0";
                oForm.Items.Item("Opt03").Specific.GroupWith("Opt02");

                //담당자
                oDS_PS_SD040H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD());
                oDS_PS_SD040H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value + "'", 0, 1));
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_SD040_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PS_SD040", "DocType", false);
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("TrType").Specific, "PS_PS_SD040", "TrType", false);

                oForm.Items.Item("ExportYN").Specific.ValidValues.Add("N", "내수");
                oForm.Items.Item("ExportYN").Specific.ValidValues.Add("Y", "수출");
                oForm.Items.Item("ExportYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("Status"), "PS_SD040", "Mat01", "Status", false);
                dataHelpClass.Combo_ValidValues_SetValueItem((oForm.Items.Item("ProgStat").Specific), "PS_PS_SD040", "ProgStat", false);
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("TrType"), "PS_SD040", "Mat01", "TrType", false);
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "1", false, false);
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemGpCd"), "SELECT ItmsGrpCod,ItmsGrpNam FROM [OITB]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code,Name FROM [@PSH_ITMBSORT]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItemType"), "SELECT Code,Name FROM [@PSH_SHAPE]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Quality"), "SELECT Code,Name FROM [@PSH_QUALITY]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("Mark"), "SELECT Code,Name FROM [@PSH_MARK]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("SbasUnit"), "SELECT Code,Name FROM [@PSH_UOMORG]", "", "");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_SD040_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromListCollection oCFLs01 = null;
            SAPbouiCOM.Conditions oCons = null;
            SAPbouiCOM.Conditions oCons01 = null;
            SAPbouiCOM.Condition oCon = null;
            SAPbouiCOM.Condition oCon01 = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromList oCFL01 = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams01 = null;
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.EditText oEdit01 = null;
            SAPbouiCOM.Column oColumn = null;

            try
            {
                oEdit = oForm.Items.Item("DCardCod").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_BusinessPartner);
                oCFLCreationParams.UniqueID = "CFLCARDCODE";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oEdit.ChooseFromListUID = "CFLCARDCODE";
                oEdit.ChooseFromListAlias = "CardCode";

                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "CardType";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "C";
                oCFL.SetConditions(oCons);

                oEdit01 = oForm.Items.Item("CardCode").Specific;
                oCFLs01 = oForm.ChooseFromLists;
                oCFLCreationParams01 = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams01.ObjectType = Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_BusinessPartner);
                oCFLCreationParams01.UniqueID = "CFLCARD2CODE";
                oCFLCreationParams01.MultiSelection = false;
                oCFL01 = oCFLs01.Add(oCFLCreationParams01);

                oEdit01.ChooseFromListUID = "CFLCARD2CODE";
                oEdit01.ChooseFromListAlias = "CardCode";

                oCons01 = oCFL01.GetConditions();
                oCon01 = oCons01.Add();
                oCon01.Alias = "CardType";
                oCon01.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon01.CondVal = "C";
                oCFL01.SetConditions(oCons01);

                oColumn = oMat01.Columns.Item("WhsCode");
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = Convert.ToString(SAPbouiCOM.BoLinkedObject.lf_Warehouses);
                oCFLCreationParams.UniqueID = "CFLWAREHOUSES";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);
                oColumn.ChooseFromListUID = "CFLWAREHOUSES";
                oColumn.ChooseFromListAlias = "WhsCode";
            }
            catch(Exception ex)
            {

            }
            finally
            {
                if (oCFLs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                }

                if (oCFLs01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs01);
                }

                if (oCons != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons);
                }

                if (oCons01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons01);
                }

                if (oCon != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon);
                }

                if (oCon01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon01);
                }

                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }

                if (oCFL01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL01);
                }

                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }

                if (oCFLCreationParams01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams01);
                }

                if (oEdit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
                }

                if (oEdit01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit01);
                }

                if (oColumn != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oColumn);
                }
            }
        }

        /// <summary>
        /// 메뉴설정
        /// </summary>
        private void PS_SD040_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, true, true, true, true, true, true, false, false, false, false, false, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        private void PS_SD040_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_SD040_FormItemEnabled();
                    PS_SD040_AddMatrixRow(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_SD040_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 화면의 아이템 Enable 설정
        /// </summary>
        private void PS_SD040_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("DCardCod").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("DocType").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = true;
                    oForm.Items.Item("TrType").Enabled = true;
                    oForm.Items.Item("Opt01").Enabled = true;
                    oForm.Items.Item("Opt02").Enabled = true;
                    oForm.Items.Item("Opt03").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;

                    oMat01.AutoResizeColumns();
                    PS_SD040_SetDocEntry();
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("DocType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocType").Enabled = false;
                    oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("ExportYN").Specific.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oMat01.Columns.Item("SD030Num").Visible = true;
                    oMat01.Columns.Item("PackNo").Visible = true;
                    oMat01.Columns.Item("OrderNum").Visible = true;
                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("DueDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("ProgStat").Enabled = true;
                    oForm.Items.Item("ProgStat").Specific.Select("3", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("ProgStat").Enabled = false;
                    oForm.Items.Item("Opt01").Specific.Selected = true;
                    oForm.Items.Item("SumSjQty").Specific.Value = "0";
                    oForm.Items.Item("SumSjWt").Specific.Value = "0";
                    oForm.Items.Item("SumQty").Specific.Value = "0";
                    oForm.Items.Item("SumWeight").Specific.Value = "0";
                    oForm.Items.Item("HWeight").Specific.Value = "0";

                    if (oForm.Items.Item("Opt03").Specific.Selected == true)
                    {
                        oForm.Items.Item("PriChange").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("PriChange").Enabled = false;
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("DCardCod").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("DocType").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = true;
                    oForm.Items.Item("TrType").Enabled = true;
                    oForm.Items.Item("Opt01").Enabled = true;
                    oForm.Items.Item("Opt02").Enabled = true;
                    oForm.Items.Item("Opt03").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oMat01.AutoResizeColumns();
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oMat01.Columns.Item("SD030Num").Visible = true;
                    oMat01.Columns.Item("PackNo").Visible = true;
                    oMat01.Columns.Item("OrderNum").Visible = true;

                    if (oForm.Items.Item("Opt03").Specific.Selected == true)
                    {
                        oForm.Items.Item("PriChange").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("PriChange").Enabled = false;
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = false;
                    oForm.Items.Item("DCardCod").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = false;
                    oForm.Items.Item("DocType").Enabled = false;
                    oForm.Items.Item("CntcCode").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("DueDate").Enabled = false;
                    oForm.Items.Item("Comments").Enabled = true;
                    oForm.Items.Item("TrType").Enabled = false;
                    oForm.Items.Item("Opt01").Enabled = false;
                    oForm.Items.Item("Opt02").Enabled = false;
                    oForm.Items.Item("Opt03").Enabled = false;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oMat01.AutoResizeColumns();
                    oForm.Items.Item("DocDate").Enabled = false;
                    oForm.Items.Item("DueDate").Enabled = false;
                    oMat01.Columns.Item("SD030Num").Visible = false;
                    oMat01.Columns.Item("PackNo").Visible = false;
                    oMat01.Columns.Item("OrderNum").Visible = false;
                    oMat01.Columns.Item("Qty").Editable = false;
                    oMat01.Columns.Item("Weight").Editable = false;
                    oMat01.Columns.Item("WhsCode").Editable = false;
                    oMat01.Columns.Item("Comments").Editable = false;
                    
                    if (oForm.Items.Item("Opt03").Specific.Selected == true)
                    {
                        oForm.Items.Item("PriChange").Enabled = true;
                        if (PSH_Globals.oCompany.UserName == "PSH93" || PSH_Globals.oCompany.UserName == "manager")
                        {
                            oMat01.Columns.Item("Price").Editable = true;
                        }
                        else
                        {
                            oMat01.Columns.Item("Price").Editable = false;
                        }
                    }
                    else
                    {
                        oForm.Items.Item("PriChange").Enabled = false;
                        oMat01.Columns.Item("Price").Editable = false;
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DocEntry 초기화
        /// </summary>
        private void PS_SD040_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD040'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = "1";
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD040_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                if (RowIserted == false)
                {
                    oDS_PS_SD040L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD040L.Offset = oRow;
                oDS_PS_SD040L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat01.LoadFromDataSource();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 필수 입력 사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_SD040_CheckDataValid()
        {
            bool returnValue = false;
            int i;
            int j;
            string qeury;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                {
                    errMessage = "고객은 필수입니다.";
                    oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errMessage = "전기일은 필수입니다.";
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                //마감상태 체크_S
                if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value, oForm.TypeEx) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다." + (char)19 + "전기일을 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }
                //마감상태 체크_E

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {   
                    if (oForm.Items.Item("TrType").Specific.Value == "1") //멀티게이지가 아닐 때만 체크
                    {
                        //전기일 체크_S
                        qeury = "  SELECT   A.U_DocDate";
                        qeury += " FROM     [@PS_SD030H] AS A";
                        qeury += "          INNER JOIN";
                        qeury += "          [@PS_SD030L] AS B";
                        qeury += "              ON A.DocEntry = B.DocEntry";
                        qeury += " WHERE    A.DocEntry = " + oMat01.Columns.Item("SD030H").Cells.Item(i).Specific.Value;
                        qeury += "          AND B.U_LineNum = " + oMat01.Columns.Item("SD030L").Cells.Item(i).Specific.Value;

                        RecordSet01.DoQuery(qeury);

                        //생산완료의 전기일(출하요청의 전기일:생산완료의 전기일은 데이터의 구조상 JOIN을 할 수가 없음)보다 이전으로 입력불가(2011.12.07 송명규 추가)
                        if (RecordSet01.Fields.Item(0).Value.ToString("yyyyMMdd") > oForm.Items.Item("DocDate").Specific.Value) //TODO : 일자 비교 연산 오류 발생 확인
                        {
                            errMessage = "납품처리 전기일이 출하요청(" + oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value + ") 전기일보다 이전입니다." + (char)19 + "납품전기일을 확인하십시오.";
                            oMat01.Columns.Item("SD030Num").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        //전기일 체크_E
                    }

                    if (oForm.Items.Item("Opt01").Specific.Selected == true)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "출하(선출)요청은 필수입니다.";
                            oMat01.Columns.Item("SD030Num").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                    else if (oForm.Items.Item("Opt02").Specific.Selected == true)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("PackNo").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "포장번호는 필수입니다.";
                            oMat01.Columns.Item("PackNo").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                    else if (oForm.Items.Item("Opt03").Specific.Selected == true)
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("PackNo").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "포장번호는 필수입니다.";
                            oMat01.Columns.Item("PackNo").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    if (oForm.Items.Item("TrType").Specific.Value == "2") //거래형태 : 임가공
                    {
                        if (dataHelpClass.GetItem_ManBtchNum(oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value) == "Y") //해당품목이 배치를 사용하는 품목이면
                        {
                            if (string.IsNullOrEmpty(oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value))
                            {
                                errMessage = "임가공거래시 LotNo는 필수입니다.";
                                oMat01.Columns.Item("LotNo").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                        }
                    }

                    if (Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "중량(수량)은 필수입니다.";
                        oMat01.Columns.Item("Weight").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (string.IsNullOrEmpty(oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "창고는 필수입니다.";
                        oMat01.Columns.Item("WhsCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    for (j = i + 1; j <= oMat01.VisualRowCount - 1; j++)
                    {
                        if (oForm.Items.Item("Opt01").Specific.Selected == true)
                        {
                            if (oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value == oMat01.Columns.Item("SD030Num").Cells.Item(j).Specific.Value)
                            {
                                errMessage = "동일한 출하요청문서가 존재합니다.";
                                oMat01.Columns.Item("SD030Num").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }

                            if (codeHelpClass.Mid(oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value, 0, oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value.ToString().IndexOf("-")) != codeHelpClass.Mid(oMat01.Columns.Item("SD030Num").Cells.Item(j).Specific.Value, 0, oMat01.Columns.Item("SD030Num").Cells.Item(j).Specific.Value.ToString().IndexOf("-")))
                            {
                                if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "4") //구로사업소일경우에만 동일하지 않은 출하요청문서 막음
                                {
                                    errMessage = "동일하지않은 출하요청문서가 존재합니다.";
                                    oMat01.Columns.Item("SD030Num").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    throw new Exception();
                                }
                            }
                        }
                    }
                }
                //여신한도초과 Check(?)
                //풍산은 체크안함
                if (oForm.Items.Item("CardCode").Specific.Value != "12532")
                {
                    if (PS_SD040_ValidateCreditLine() == false)
                    {
                        errMessage = " ";
                        throw new Exception();
                    }
                }
                if (PS_SD040_Validate("검사") == false)
                {
                    errMessage = " ";
                    throw new Exception();
                }

                oDS_PS_SD040L.RemoveRecord(oDS_PS_SD040L.Size - 1);
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_SD040_SetDocEntry();
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage == " ")
                {
                    //프로세스 체크용 외부 메소드의 실행 결과는 각 메소드에서 메시지 출력, 해당 메소드에서는 메시지 처리 불필요
                }
                else if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 여신한도 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_SD040_ValidateCreditLine()
        {
            bool returnValue = false;
            string query;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Items.Item("Opt03").Specific.Selected == true) //분말제품
                {
                    query = "EXEC [S139_hando] '";
                    query += oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "', '";
                    query += oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'";
                    RecordSet01.DoQuery(query);

                    if (RecordSet01.RecordCount > 0)
                    {
                        if (Convert.ToDouble(RecordSet01.Fields.Item("OverAmt").Value) < 0) //TODO : 부호확인 필요, 여신한도금액
                        {
                            errMessage = "여신한도를 초과했습니다.";
                            throw new Exception();
                        }
                    }
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty) 
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_SD040_Validate(string ValidateType)
        {
            bool returnValue = false;
            object i = null;
            int j = 0;
            string Query01 = null;
            SAPbobsCOM.Recordset RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_SD040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
            {
                MDC_Com.MDC_GF_Message(ref "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", ref "W");
                functionReturnValue = false;
                goto PS_SD040_Validate_Exit;
            }

            if (ValidateType == "검사")
            {
                ////여신한도체크
                //        If MDC_PS_Common.GetValue("EXEC PS_SD040_04 '" & oForm.Items("CardCode").Specific.Value & "'", 0, 1) <= 0 Then
                //            Call MDC_Com.MDC_GF_Message("여신한도 부족합니다.", "W")
                //            PS_SD040_Validate = False
                //            GoTo PS_SD040_Validate_Exit
                //        End If

                ////입력된 행에 대해
                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    //UPGRADE_WARNING: oForm.Items(Opt01).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    if (oForm.Items.Item("Opt01").Specific.Selected == true)
                    {
                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        //UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [PS_SD030H] PS_SD030H LEFT JOIN [PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_SD030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_SD030L.LineId) = ' & oMat01.Columns(SD030Num).Cells(i).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_SD030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_SD030L.LineId) = '" + oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value + "'", 0, 1) <= 0)
                        {
                            MDC_Com.MDC_GF_Message(ref "출하(선출)요청문서가 존재하지 않습니다.", ref "W");
                            functionReturnValue = false;
                            goto PS_SD040_Validate_Exit;
                        }
                    }
                    //UPGRADE_WARNING: oForm.Items(Opt02).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    if (oForm.Items.Item("Opt02").Specific.Selected == true)
                    {

                    }
                }
            }
            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            RecordSet01 = null;
            return functionReturnValue;
        PS_SD040_Validate_Exit:
            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            RecordSet01 = null;
            return functionReturnValue;
        PS_SD040_Validate_Error:
            functionReturnValue = false;
            SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            return functionReturnValue;
        }
        #endregion


        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	switch (pVal.EventType) {
        //		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //			////1
        //			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //			////2
        //			Raise_EVENT_KEY_DOWN(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //			////5
        //			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CLICK:
        //			////6
        //			Raise_EVENT_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //			////7
        //			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //			////8
        //			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //			////10
        //			Raise_EVENT_VALIDATE(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //			////11
        //			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //			////18
        //			break;
        //		////et_FORM_ACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //			////19
        //			break;
        //		////et_FORM_DEACTIVATE
        //		case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //			////20
        //			Raise_EVENT_RESIZE(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //			////27
        //			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //			////3
        //			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //			////4
        //			break;
        //		////et_LOST_FOCUS
        //		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //			////17
        //			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pVal, ref BubbleEvent);
        //			break;
        //	}
        //	return;
        //	Raise_ItemEvent_Error:
        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	string sQry = null;

        //	////BeforeAction = True
        //	if ((pVal.BeforeAction == true)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				////납품취소, 반품
        //				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_SD040H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_SD040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y") {
        //						MDC_Com.MDC_GF_Message(ref "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //					}
        //					//UPGRADE_WARNING: oForm.Items(Opt02).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					////멀티이면서
        //					if (oForm.Items.Item("Opt02").Specific.Selected == true) {
        //						////멀티게이지 재고수량검사
        //						for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //							sQry = "Select Quantity = Sum(a.Quantity) From OIBT a Inner Join OITM b On a.ItemCode = b.ItemCode ";
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							sQry = sQry + " Where b.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + "'";
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							sQry = sQry + " And a.BatchNum = '" + oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value + "'";
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							sQry = sQry + " AND a.WhsCode = '" + oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value + "'";
        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue(sQry, 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if (MDC_PS_Common.GetValue(sQry, 0, 1) > 0) {
        //								//If MDC_PS_Common.GetValue("SELECT Quantity FROM [OIBT] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.Value & "' AND BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.Value & "' AND WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.Value & "'", 0, 1) > 0 Then

        //								//If MDC_PS_Common.GetValue("SELECT Sum(Quantity * (Case When Direction = '0' Then 1 Else -1 End))  FROM [IBT1] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.Value & "' AND BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.Value & "' AND WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.Value & "'", 0, 1) > 0 Then
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								MDC_Com.MDC_GF_Message(ref "멀티게이지 품목 : " + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + " 의 재고가 존재합니다.", ref "W");
        //								BubbleEvent = false;
        //								return;
        //							}
        //						}
        //					}
        //					//분말은 배치품목 수량을 분할로 납품반품되야되기때문에 재고가 있더라도 반품이 가능해야함(황영수 20190924)
        //					//                    If oForm.Items("Opt03").Specific.Selected = True Then '//분말이면서
        //					//                        For i = 1 To oMat01.VisualRowCount - 1 '//분말 재고수량검사
        //					//                            If MDC_PS_Common.GetValue("SELECT Quantity FROM [OIBT] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.Value & "' AND BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.Value & "' AND WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.Value & "'", 0, 1) > 0 Then
        //					//                                Call MDC_Com.MDC_GF_Message("분말 품목 : " & oMat01.Columns("ItemCode").Cells(i).Specific.Value & " 의 재고가 존재합니다.", "W")
        //					//                                BubbleEvent = False
        //					//                                Exit Sub
        //					//                            End If
        //					//                        Next
        //					//                    End If
        //					////AR송장처리된 문서존재유무검사
        //					for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(DLN1Num).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = ' & oMat01.Columns(ODLNNum).Cells(i).Specific.Value & ' AND BaseLine = ' & oMat01.Columns(DLN1Num).Cells(i).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = '" + oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value + "' AND BaseLine = '" + oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value + "'", 0, 1) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "AR송장처리된 문서가 존재합니다.", ref "W");
        //							BubbleEvent = false;
        //							return;
        //						}
        //					}
        //					////반품처리된 문서존재유무검사
        //					for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(DLN1Num).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = ' & oMat01.Columns(ODLNNum).Cells(i).Specific.Value & ' AND BaseLine = ' & oMat01.Columns(DLN1Num).Cells(i).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = '" + oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value + "' AND BaseLine = '" + oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value + "'", 0, 1) > 0) {
        //							MDC_Com.MDC_GF_Message(ref "반품처리된 문서가 존재합니다.", ref "W");
        //							BubbleEvent = false;
        //							return;
        //						}
        //					}
        //					if (PS_SD040_DI_API_03() == false) {
        //						BubbleEvent = false;
        //						return;
        //					}
        //				} else {
        //					MDC_Com.MDC_GF_Message(ref "현재 모드에서는 취소할수 없습니다.", ref "W");
        //					BubbleEvent = false;
        //					return;
        //				}
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			//                If oForm.Mode = fm_OK_MODE Then
        //			//                    If MDC_PS_Common.GetValue("SELECT Status FROM [@PS_SD040H] WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'", 0, 1) = "C" Then
        //			//                        Call MDC_Com.MDC_GF_Message("해당문서는 다른사용자에 의해 닫기되었습니다. 작업을 진행할수 없습니다.", "W")
        //			//                        BubbleEvent = False
        //			//                        Exit Sub
        //			//                    End If
        //			//                Else
        //			//                    MDC_Com.MDC_GF_Message "현재 모드에서는 취소할수 없습니다.", "W"
        //			//                    BubbleEvent = False
        //			//                    Exit Sub
        //			//                End If
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				break;
        //			case "1282":
        //				//추가
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((pVal.BeforeAction == false)) {
        //		switch (pVal.MenuUID) {
        //			case "1284":
        //				//취소
        //				break;
        //			case "1286":
        //				//닫기
        //				break;
        //			case "1293":
        //				//행삭제
        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
        //				break;
        //			case "1281":
        //				//찾기
        //				PS_SD040_FormItemEnabled();
        //				////UDO방식
        //				oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				break;
        //			case "1282":
        //				//추가
        //				PS_SD040_FormItemEnabled();
        //				////UDO방식
        //				PS_SD040_AddMatrixRow(0, ref true);
        //				////UDO방식
        //				////2010.12.06 추가
        //				//담당자
        //				oDS_PS_SD040H.SetValue("U_CntcCode", 0, MDC_PS_Common.User_MSTCOD());
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDS_PS_SD040H.SetValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value + "'", 0, 1));
        //				break;
        //			case "1288":
        //			case "1289":
        //			case "1290":
        //			case "1291":
        //				//레코드이동버튼
        //				PS_SD040_FormItemEnabled();
        //				////UDO방식
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_MenuEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////BeforeAction = True
        //	if ((BusinessObjectInfo.BeforeAction == true)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	////BeforeAction = False
        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
        //		switch (BusinessObjectInfo.EventType) {
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //				////33
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //				////34
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //				////35
        //				break;
        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //				////36
        //				break;
        //		}
        //	}
        //	return;
        //	Raise_FormDataEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {
        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;
        //			}
        //		} else {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = "";
        //			oLastColRow01 = 0;
        //		}
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //		//            MenuCreationParams01.uniqueID = "MenuUID"
        //		//            MenuCreationParams01.String = "메뉴명"
        //		//            MenuCreationParams01.Enabled = True
        //		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //		//        End If
        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {
        //				oLastItemUID01 = pVal.ItemUID;
        //				oLastColUID01 = pVal.ColUID;
        //				oLastColRow01 = pVal.Row;
        //			}
        //		} else {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = "";
        //			oLastColRow01 = 0;
        //		}
        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //		//        End If
        //	}
        //	return;
        //	Raise_RightClickEvent_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	object TempForm01 = null;
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (PS_SD040_CheckDataValid() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				//                납품처리 문서 단독으로 입력할 경우 주석처리 ("납품"문서 생성 안함)_S
        //				//UPGRADE_WARNING: oForm.Items(Opt01).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (oForm.Items.Item("Opt01").Specific.Selected == true) {
        //					////납품생성
        //					if (PS_SD040_DI_API_01() == false) {
        //						PS_SD040_AddMatrixRow(oMat01.RowCount);
        //						////UDO방식일때
        //						BubbleEvent = false;
        //						return;
        //					}
        //					//UPGRADE_WARNING: oForm.Items(Opt02).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				} else if (oForm.Items.Item("Opt02").Specific.Selected == true) {
        //					////멀티게이지납품생성
        //					if (PS_SD040_DI_API_02() == false) {
        //						PS_SD040_AddMatrixRow(oMat01.RowCount);
        //						////UDO방식일때
        //						BubbleEvent = false;
        //						return;
        //					}
        //					//UPGRADE_WARNING: oForm.Items(Opt03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				} else if (oForm.Items.Item("Opt03").Specific.Selected == true) {
        //					////분말 납품생성
        //					if (PS_SD040_DI_API_01() == false) {
        //						PS_SD040_AddMatrixRow(oMat01.RowCount);
        //						////UDO방식일때
        //						BubbleEvent = false;
        //						return;
        //					}
        //				}
        //				//                납품처리 문서 단독으로 입력할 경우 주석처리 ("납품"문서 생성 안함)_E
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				if (PS_SD040_CheckDataValid() == false) {
        //					BubbleEvent = false;
        //					return;
        //				}
        //				//----------------------------------------
        //				//출고가 생성안되었을때
        //				//                If oForm.Items("Opt01").Specific.Selected = True Then
        //				//                    If PS_SD040_DI_API_01 = False Then '//납품생성
        //				//                        Call PS_SD040_AddMatrixRow(oMat01.RowCount) '//UDO방식일때
        //				//                        BubbleEvent = False
        //				//                        Exit Sub
        //				//                    End If
        //				//                End If
        //				//
        //				//                If oForm.Items("Opt02").Specific.Selected = True Then
        //				//                    If PS_SD040_DI_API_02 = False Then '//멀티게이지납품생성
        //				//                        Call PS_SD040_AddMatrixRow(oMat01.RowCount) '//UDO방식일때
        //				//                        BubbleEvent = False
        //				//                        Exit Sub
        //				//                    End If
        //				//                End If
        //				//-----------------------------------------
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //			}
        //		}
        //		if (pVal.ItemUID == "Button01") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				PS_SD040_Print_Report01();
        //				//거래명세서
        //			}
        //		}
        //		if (pVal.ItemUID == "Button02") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(Opt03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//분말일경우
        //				if (oForm.Items.Item("Opt03").Specific.Selected == true) {
        //					PS_SD040_Print_Report03("N");
        //					//거래명세표(수량)
        //				} else {
        //					PS_SD040_Print_Report02("N");
        //					//출고증
        //				}
        //			}
        //		}

        //		if (pVal.ItemUID == "opt03") {
        //			PS_SD040_FormItemEnabled();
        //		}

        //		if (pVal.ItemUID == "PriChange") {
        //			PS_SD0040_ChangePrice();
        //			PS_SD040_AddMatrixRow(oMat01.RowCount);
        //			////UDO방식일때
        //		}

        //		if (pVal.ItemUID == "Button02_1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				//If oForm.Items("Opt03").Specific.Selected = True Then  '이영진과장 요청으로 모든 제품에 다 출력되도록 처리함.
        //				PS_SD040_Print_Report03_1("N");
        //				//거래명세표(금액)
        //				PS_SD0040_ChangePrice();
        //				// End If
        //			}
        //		}
        //		if (pVal.ItemUID == "Button03") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {

        //				TempForm01 = new PS_MM004();
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: TempForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				TempForm01.LoadForm("PS_SD040", oForm.Items.Item("DocEntry").Specific.Value);
        //				BubbleEvent = false;
        //				//

        //				//                Call PS_SD040_Print_Report02("Y") '출고증(운송장 포함)
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemUID == "1") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_SD040_FormItemEnabled();
        //					////Call PS_SD040_AddMatrixRow(oMat01.RowCount, True) '//UDO방식일때
        //					oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //					SubMain.Sbo_Application.ActivateMenuItem("1291");
        //				}
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
        //				if (pVal.ActionSuccess == true) {
        //					PS_SD040_FormItemEnabled();
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ITEM_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
        //		////사용자값활성
        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "ItemCode", "") '//사용자값활성
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.ColUID == "SD030Num") {
        //				//UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value)) {
        //					MDC_Com.MDC_GF_Message(ref "고객코드는 필수입니다.", ref "W");
        //					BubbleEvent = false;
        //					return;
        //				}
        //				MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "SD030Num");
        //				////사용자값활성
        //			}
        //		}
        //		////Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm, pVal, BubbleEvent, "Mat01", "OrderNum") '//사용자값활성
        //		////Call MDC_PS_Common.ActiveUserDefineValueAlways(oForm, pVal, BubbleEvent, "Mat01", "WhsCode") '//사용자값활성
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.ColUID == "PackNo") {
        //				//UPGRADE_WARNING: oForm.Items(Opt02).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (oForm.Items.Item("Opt02").Specific.Selected == true) {
        //					//UPGRADE_WARNING: oForm.Items(TrType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("TrType").Specific.Selected.Value != "2") {
        //						MDC_Com.MDC_GF_Message(ref "포장번호조회는 거래형태가 임가공이여야 합니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //					}

        //					//UPGRADE_WARNING: oForm.Items(DCardCod).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oForm.Items.Item("DCardCod").Specific.Value)) {
        //						MDC_Com.MDC_GF_Message(ref "납품처코드는 필수입니다.", ref "W");
        //						BubbleEvent = false;
        //						return;
        //					}
        //				}
        //				//UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value)) {
        //					MDC_Com.MDC_GF_Message(ref "고객코드는 필수입니다.", ref "W");
        //					BubbleEvent = false;
        //					return;
        //				}

        //				MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PackNo");
        //				////사용자값활성
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_KEY_DOWN_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	if (pVal.BeforeAction == true) {
        //		//        If pVal.ItemUID = "DocType" Then
        //		//            If (oForm.Items("DocType").Specific.Value = "1") Then
        //		//            ElseIf (oForm.Items("DocType").Specific.Value = "2") Then
        //		//                oForm.Items("BPLId").Enabled = True
        //		//                Call oForm.Items("BPLId").Specific.Select("2", psk_ByValue)
        //		//                oForm.Items("CardCode").Click ct_Regular
        //		//                oForm.Items("BPLId").Enabled = False
        //		//            End If
        //		//        End If
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_COMBO_SELECT_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.Row > 0) {
        //				oMat01.SelectRow(pVal.Row, true, false);
        //			}
        //		}
        //		if (pVal.ItemUID == "Opt01") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				oForm.Freeze(true);
        //				oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oForm.Items.Item("CardCode").Enabled = true;
        //				oForm.Items.Item("BPLId").Enabled = true;
        //				oForm.Items.Item("DCardCod").Enabled = true;
        //				oForm.Items.Item("TrType").Enabled = true;
        //				oMat01.Columns.Item("SD030Num").Visible = true;
        //				oMat01.Columns.Item("PackNo").Visible = false;
        //				oMat01.Columns.Item("OrderNum").Visible = true;
        //				oMat01.Columns.Item("SD030Num").Editable = true;
        //				oMat01.Columns.Item("Qty").Editable = true;
        //				oMat01.Columns.Item("Weight").Editable = true;
        //				oMat01.Columns.Item("UnWeight").Visible = true;
        //				oMat01.Columns.Item("Price").Editable = false;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //				oMat01.Clear();
        //				oMat01.FlushToDataSource();
        //				oMat01.LoadFromDataSource();
        //				PS_SD040_AddMatrixRow(0, ref true);

        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumSjQty").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumSjWt").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumQty").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumWeight").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("HWeight").Specific.Value = 0;
        //				oForm.Freeze(false);
        //			}
        //		}
        //		////멀티게이지의경우
        //		if (pVal.ItemUID == "Opt02") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				oForm.Freeze(true);
        //				oForm.Items.Item("TranCard").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oForm.Items.Item("CardCode").Enabled = true;
        //				oForm.Items.Item("BPLId").Enabled = true;
        //				oForm.Items.Item("DCardCod").Enabled = true;
        //				oForm.Items.Item("TrType").Enabled = false;
        //				oMat01.Columns.Item("SD030Num").Visible = false;
        //				oMat01.Columns.Item("PackNo").Visible = true;
        //				oMat01.Columns.Item("OrderNum").Visible = false;
        //				oMat01.Columns.Item("Qty").Editable = false;
        //				oMat01.Columns.Item("Weight").Editable = false;
        //				oMat01.Columns.Item("UnWeight").Visible = false;
        //				oMat01.Columns.Item("Price").Editable = true;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("TrType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //				oMat01.Clear();
        //				oMat01.FlushToDataSource();
        //				oMat01.LoadFromDataSource();
        //				PS_SD040_AddMatrixRow(0, ref true);

        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumSjQty").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumSjWt").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumQty").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumWeight").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("HWeight").Specific.Value = 0;
        //				oForm.Freeze(false);
        //			}
        //		}

        //		//분말
        //		if (pVal.ItemUID == "Opt03") {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				oForm.Freeze(true);
        //				oForm.Items.Item("TranCard").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oForm.Items.Item("CardCode").Enabled = true;
        //				oForm.Items.Item("BPLId").Enabled = true;
        //				oForm.Items.Item("DCardCod").Enabled = true;
        //				oForm.Items.Item("TrType").Enabled = false;
        //				oMat01.Columns.Item("SD030Num").Visible = true;
        //				oMat01.Columns.Item("PackNo").Visible = true;
        //				oMat01.Columns.Item("OrderNum").Visible = true;
        //				oMat01.Columns.Item("Qty").Editable = false;
        //				oMat01.Columns.Item("Weight").Editable = false;
        //				oMat01.Columns.Item("SD030Num").Editable = false;
        //				oMat01.Columns.Item("UnWeight").Visible = false;
        //				oMat01.Columns.Item("Price").Editable = false;
        //				oMat01.Columns.Item("WhsCode").Editable = true;
        //				oMat01.Columns.Item("Comments").Editable = true;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
        //				oMat01.Clear();
        //				oMat01.FlushToDataSource();
        //				oMat01.LoadFromDataSource();
        //				PS_SD040_AddMatrixRow(0, ref true);

        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumSjQty").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumSjWt").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumQty").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("SumWeight").Specific.Value = 0;
        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("HWeight").Specific.Value = 0;
        //				oForm.Freeze(false);
        //			}
        //		}


        //	} else if (pVal.BeforeAction == false) {
        //		if (pVal.ItemUID == "Opt01") {
        //			oForm.Items.Item("PriChange").Enabled = false;
        //		}

        //		////멀티게이지의경우
        //		if (pVal.ItemUID == "Opt02") {
        //			oForm.Items.Item("PriChange").Enabled = false;
        //		}

        //		//분말
        //		if (pVal.ItemUID == "Opt03") {
        //			oForm.Items.Item("PriChange").Enabled = true;
        //		}
        //		oForm.Freeze(false);
        //	}
        //	return;
        //	Raise_EVENT_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_DOUBLE_CLICK_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	object oTempClass = null;
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemUID == "Mat01") {
        //			if (pVal.ColUID == "SD030Num") {
        //				oTempClass = new PS_SD030();
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oTempClass.LoadForm(Strings.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, "-") - 1));
        //			}
        //			if (pVal.ColUID == "SD030H") {
        //				oTempClass = new PS_SD030();
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oTempClass.LoadForm(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	oForm.Freeze(true);
        //	int i = 0;
        //	string Query01 = null;
        //	string Work01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	SAPbobsCOM.Recordset RecordSet02 = null;
        //	SAPbobsCOM.Recordset RecordSet03 = null;
        //	string ItemCode01 = null;
        //	string sQry = null;
        //	int SumSjQty = 0;
        //	int SumQty = 0;
        //	decimal SumWeight = default(decimal);
        //	decimal SumSjWt = default(decimal);
        //	decimal HWeight = default(decimal);
        //	if (pVal.BeforeAction == true) {
        //		if (pVal.ItemChanged == true) {
        //			if ((pVal.ItemUID == "Mat01")) {
        //				if (pVal.ColUID == "SD030Num") {
        //					//UPGRADE_WARNING: oMat01.Columns(SD030Num).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value)) {
        //						goto Raise_EVENT_VALIDATE_Exit;
        //					}
        //					//UPGRADE_WARNING: oForm.Items(Opt03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("Opt03").Specific.Selected != true) {
        //						for (i = 1; i <= oMat01.RowCount; i++) {
        //							////현재 선택되어있는 행이 아니면
        //							if (pVal.Row != i) {
        //								//UPGRADE_WARNING: oMat01.Columns(SD030Num).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								//UPGRADE_WARNING: oMat01.Columns(SD030Num).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if ((oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value)) {
        //									MDC_Com.MDC_GF_Message(ref "동일한 출하(선출)요청이 존재합니다.", ref "W");
        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value = "";
        //									goto Raise_EVENT_VALIDATE_Exit;
        //								}
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								if ((Strings.Mid(oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value, "-") - 1) != Strings.Mid(oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value, "-") - 1))) {
        //									//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									//[2011.1.10추가]구로사업소일경우에만 동일하지 않은 출하요청문서 막음
        //									if (Strings.Trim(oForm.Items.Item("BPLId").Specific.Value) == "4") {
        //										MDC_Com.MDC_GF_Message(ref "동일하지않은 출하요청문서가 존재합니다.", ref "W");
        //										//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //										oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value = "";
        //										goto Raise_EVENT_VALIDATE_Exit;
        //									}
        //								}
        //							}
        //						}
        //					}
        //					RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					Query01 = "EXEC PS_SD040_01 '" + oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value + "'";
        //					RecordSet01.DoQuery(Query01);
        //					for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //						oDS_PS_SD040L.SetValue("U_SD030Num", pVal.Row - 1, RecordSet01.Fields.Item("SD030Num").Value);
        //						oDS_PS_SD040L.SetValue("U_OrderNum", pVal.Row - 1, RecordSet01.Fields.Item("OrderNum").Value);
        //						oDS_PS_SD040L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
        //						oDS_PS_SD040L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
        //						oDS_PS_SD040L.SetValue("U_ItemGpCd", pVal.Row - 1, RecordSet01.Fields.Item("ItemGpCd").Value);
        //						oDS_PS_SD040L.SetValue("U_ItmBsort", pVal.Row - 1, RecordSet01.Fields.Item("ItmBsort").Value);
        //						oDS_PS_SD040L.SetValue("U_ItmMsort", pVal.Row - 1, RecordSet01.Fields.Item("ItmMsort").Value);
        //						oDS_PS_SD040L.SetValue("U_Unit1", pVal.Row - 1, RecordSet01.Fields.Item("Unit1").Value);
        //						oDS_PS_SD040L.SetValue("U_Size", pVal.Row - 1, RecordSet01.Fields.Item("Size").Value);
        //						oDS_PS_SD040L.SetValue("U_ItemType", pVal.Row - 1, RecordSet01.Fields.Item("ItemType").Value);
        //						oDS_PS_SD040L.SetValue("U_Quality", pVal.Row - 1, RecordSet01.Fields.Item("Quality").Value);
        //						oDS_PS_SD040L.SetValue("U_Mark", pVal.Row - 1, RecordSet01.Fields.Item("Mark").Value);
        //						oDS_PS_SD040L.SetValue("U_SbasUnit", pVal.Row - 1, RecordSet01.Fields.Item("SbasUnit").Value);
        //						oDS_PS_SD040L.SetValue("U_LotNo", pVal.Row - 1, RecordSet01.Fields.Item("LotNo").Value);
        //						oDS_PS_SD040L.SetValue("U_SjQty", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("SjQty").Value)));
        //						oDS_PS_SD040L.SetValue("U_SjWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("SjWeight").Value)));
        //						oDS_PS_SD040L.SetValue("U_Qty", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Qty").Value)));
        //						oDS_PS_SD040L.SetValue("U_UnWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("UnWeight").Value)));
        //						oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Weight").Value)));
        //						oDS_PS_SD040L.SetValue("U_Currency", pVal.Row - 1, RecordSet01.Fields.Item("Currency").Value);
        //						oDS_PS_SD040L.SetValue("U_Price", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Price").Value)));
        //						oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("LinTotal").Value)));
        //						oDS_PS_SD040L.SetValue("U_WhsCode", pVal.Row - 1, RecordSet01.Fields.Item("WhsCode").Value);
        //						oDS_PS_SD040L.SetValue("U_WhsName", pVal.Row - 1, RecordSet01.Fields.Item("WhsName").Value);
        //						oDS_PS_SD040L.SetValue("U_Comments", pVal.Row - 1, RecordSet01.Fields.Item("Comments").Value);
        //						oDS_PS_SD040L.SetValue("U_SD030H", pVal.Row - 1, RecordSet01.Fields.Item("SD030H").Value);
        //						oDS_PS_SD040L.SetValue("U_SD030L", pVal.Row - 1, RecordSet01.Fields.Item("SD030L").Value);
        //						oDS_PS_SD040L.SetValue("U_TrType", pVal.Row - 1, RecordSet01.Fields.Item("TrType").Value);
        //						oDS_PS_SD040L.SetValue("U_ORDRNum", pVal.Row - 1, RecordSet01.Fields.Item("ORDRNum").Value);
        //						oDS_PS_SD040L.SetValue("U_RDR1Num", pVal.Row - 1, RecordSet01.Fields.Item("RDR1Num").Value);
        //						oDS_PS_SD040L.SetValue("U_Status", pVal.Row - 1, RecordSet01.Fields.Item("Status").Value);
        //						oDS_PS_SD040L.SetValue("U_LineId", pVal.Row - 1, RecordSet01.Fields.Item("LineId").Value);
        //						RecordSet01.MoveNext();
        //					}
        //					if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD040L.GetValue("U_SD030Num", pVal.Row - 1)))) {
        //						PS_SD040_AddMatrixRow((pVal.Row));
        //					}
        //					oMat01.LoadFromDataSource();
        //					oMat01.AutoResizeColumns();

        //					//2011.1.4 추가(yjh) : 출하요청(SD030)의 비고 내용을 납품(SD040)에 넣어준다
        //					if (oMat01.VisualRowCount > 0) {
        //						RecordSet03 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						sQry = "select U_Comments from [@PS_SD030H] where DocEntry = '" + Strings.Mid(oMat01.Columns.Item("SD030Num").Cells.Item(1).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item("SD030Num").Cells.Item(1).Specific.Value, "-") - 1) + "'";
        //						RecordSet03.DoQuery(sQry);
        //						//UPGRADE_WARNING: oForm.Items(Comments).Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: RecordSet03.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oForm.Items.Item("Comments").Specific.String = RecordSet03.Fields.Item("U_Comments").Value;
        //					}
        //					//추가 end------------------------------------------------------------------

        //					for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(SjQty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value)) {
        //							SumSjQty = SumSjQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumSjQty = SumSjQty + oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumSjWt = SumSjWt + oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value;
        //						//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //							SumQty = SumQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;
        //						//UPGRADE_WARNING: oMat01.Columns(UnWeight).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						HWeight = HWeight + (oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value * oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value / 1000);
        //					}
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumSjQty").Specific.Value = SumSjQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumSjWt").Specific.Value = SumSjWt;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("HWeight").Specific.Value = HWeight;

        //					if ((oMat01.RowCount > 1)) {
        //						oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						oForm.Items.Item("CardCode").Enabled = false;
        //						oForm.Items.Item("BPLId").Enabled = false;
        //						oForm.Items.Item("DCardCod").Enabled = true;
        //						oForm.Items.Item("TrType").Enabled = false;
        //					} else {
        //						oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						oForm.Items.Item("CardCode").Enabled = true;
        //						oForm.Items.Item("BPLId").Enabled = true;
        //						oForm.Items.Item("DCardCod").Enabled = true;
        //						oForm.Items.Item("TrType").Enabled = true;
        //					}
        //					oForm.Update();
        //					//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					RecordSet01 = null;
        //				} else if (pVal.ColUID == "PackNo") {
        //					//UPGRADE_WARNING: oMat01.Columns(PackNo).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value)) {
        //						goto Raise_EVENT_VALIDATE_Exit;
        //					}
        //					for (i = 1; i <= oMat01.RowCount; i++) {
        //						////현재 선택되어있는 행이 아니면
        //						if (pVal.Row != i) {
        //							//UPGRADE_WARNING: oMat01.Columns(PackNo).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							//UPGRADE_WARNING: oMat01.Columns(PackNo).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							if ((oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("PackNo").Cells.Item(i).Specific.Value)) {
        //								MDC_Com.MDC_GF_Message(ref "동일한 포장번호가 존재합니다.", ref "W");
        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value = "";
        //								goto Raise_EVENT_VALIDATE_Exit;
        //							}
        //						}
        //					}

        //					RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //					//UPGRADE_WARNING: oForm.Items(Opt02).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (oForm.Items.Item("Opt02").Specific.Selected == true) {
        //						Work01 = "1";
        //						//UPGRADE_WARNING: oForm.Items(Opt03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					} else if (oForm.Items.Item("Opt03").Specific.Selected == true) {
        //						Work01 = "3";
        //					}
        //					//UPGRADE_WARNING: oForm.Items().Specific.String 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					Query01 = "EXEC PS_SD040_06 '" + oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value + "', '" + Work01 + "', '" + Strings.Trim(oForm.Items.Item("CardCode").Specific.String) + "'";
        //					RecordSet01.DoQuery(Query01);
        //					if (RecordSet01.RecordCount <= 0) {
        //						MDC_Com.MDC_GF_Message(ref "포장번호정보가 존재하지 않습니다.", ref "W");
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value = "";
        //						goto Raise_EVENT_VALIDATE_Exit;
        //					} else {
        //						////해당 포장번호로 재고유무확인
        //						RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						Query01 = "EXEC PS_SD040_07 '" + oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value + "'";
        //						RecordSet02.DoQuery(Query01);
        //						if (RecordSet02.Fields.Item(0).Value == "Enabled") {
        //							////진행가능
        //						} else if (RecordSet02.Fields.Item(0).Value == "Disabled") {
        //							MDC_Com.MDC_GF_Message(ref "해당포장번호의 재고가 부족합니다.", ref "W");
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value = "";
        //							goto Raise_EVENT_VALIDATE_Exit;
        //						}
        //					}
        //					for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //						oDS_PS_SD040L.SetValue("U_PackNo", pVal.Row - 1 + i, RecordSet01.Fields.Item("PackNo").Value);
        //						if (Work01 == "3") {
        //							oDS_PS_SD040L.SetValue("U_SD030Num", pVal.Row - 1 + i, RecordSet01.Fields.Item("SD030Num").Value);
        //							oDS_PS_SD040L.SetValue("U_OrderNum", pVal.Row - 1 + i, RecordSet01.Fields.Item("OrderNum").Value);
        //							oDS_PS_SD040L.SetValue("U_SD030H", pVal.Row - 1 + i, RecordSet01.Fields.Item("SD030H").Value);
        //							oDS_PS_SD040L.SetValue("U_SD030L", pVal.Row - 1 + i, RecordSet01.Fields.Item("SD030L").Value);
        //							oDS_PS_SD040L.SetValue("U_ORDRNum", pVal.Row - 1 + i, RecordSet01.Fields.Item("ORDRNum").Value);
        //							oDS_PS_SD040L.SetValue("U_RDR1Num", pVal.Row - 1 + i, RecordSet01.Fields.Item("RDR1Num").Value);
        //						}
        //						oDS_PS_SD040L.SetValue("U_ItemCode", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItemCode").Value);
        //						oDS_PS_SD040L.SetValue("U_ItemName", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItemName").Value);
        //						oDS_PS_SD040L.SetValue("U_ItemGpCd", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItemGpCd").Value);
        //						oDS_PS_SD040L.SetValue("U_ItmBsort", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItmBsort").Value);
        //						oDS_PS_SD040L.SetValue("U_ItmMsort", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItmMsort").Value);
        //						oDS_PS_SD040L.SetValue("U_Unit1", pVal.Row - 1 + i, RecordSet01.Fields.Item("Unit1").Value);
        //						oDS_PS_SD040L.SetValue("U_Size", pVal.Row - 1 + i, RecordSet01.Fields.Item("Size").Value);
        //						oDS_PS_SD040L.SetValue("U_ItemType", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItemType").Value);
        //						oDS_PS_SD040L.SetValue("U_Quality", pVal.Row - 1 + i, RecordSet01.Fields.Item("Quality").Value);
        //						oDS_PS_SD040L.SetValue("U_Mark", pVal.Row - 1 + i, RecordSet01.Fields.Item("Mark").Value);
        //						oDS_PS_SD040L.SetValue("U_SbasUnit", pVal.Row - 1 + i, RecordSet01.Fields.Item("SbasUnit").Value);
        //						oDS_PS_SD040L.SetValue("U_LotNo", pVal.Row - 1 + i, RecordSet01.Fields.Item("LotNo").Value);
        //						oDS_PS_SD040L.SetValue("U_CoilNo", pVal.Row - 1 + i, RecordSet01.Fields.Item("CoilNo").Value);
        //						oDS_PS_SD040L.SetValue("U_PackWgt", pVal.Row - 1 + i, RecordSet01.Fields.Item("PackWgt").Value);
        //						oDS_PS_SD040L.SetValue("U_SjQty", pVal.Row - 1 + i, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("SjQty").Value)));
        //						oDS_PS_SD040L.SetValue("U_SjWeight", pVal.Row - 1 + i, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("SjWeight").Value)));
        //						oDS_PS_SD040L.SetValue("U_Qty", pVal.Row - 1 + i, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Qty").Value)));
        //						//                        Call oDS_PS_SD040L.setValue("U_UnWeight", pVal.Row - 1 + i, Val(RecordSet01.Fields("UnWeight").Value))
        //						oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1 + i, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Weight").Value)));
        //						oDS_PS_SD040L.SetValue("U_Currency", pVal.Row - 1 + i, RecordSet01.Fields.Item("Currency").Value);
        //						oDS_PS_SD040L.SetValue("U_Price", pVal.Row - 1 + i, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("Price").Value)));
        //						oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1 + i, Convert.ToString(Conversion.Val(RecordSet01.Fields.Item("LinTotal").Value)));
        //						oDS_PS_SD040L.SetValue("U_WhsCode", pVal.Row - 1 + i, RecordSet01.Fields.Item("WhsCode").Value);
        //						oDS_PS_SD040L.SetValue("U_WhsName", pVal.Row - 1 + i, RecordSet01.Fields.Item("WhsName").Value);
        //						//                        Call oDS_PS_SD040L.setValue("U_Comments", pVal.Row - 1 + i, RecordSet01.Fields("Comments").Value)
        //						oDS_PS_SD040L.SetValue("U_Status", pVal.Row - 1 + i, RecordSet01.Fields.Item("Status").Value);
        //						oDS_PS_SD040L.SetValue("U_LineId", pVal.Row - 1 + i, RecordSet01.Fields.Item("LineId").Value);
        //						PS_SD040_AddMatrixRow((pVal.Row + i));
        //						//                        If oMat01.RowCount = pVal.Row And Trim(oDS_PS_SD040L.GetValue("U_PackNo", pVal.Row - 1)) <> "" Then
        //						//                            PS_SD040_AddMatrixRow (pVal.Row)
        //						//                        End If
        //						//                        pVal.Row = pVal.Row + 1
        //						RecordSet01.MoveNext();
        //					}

        //					for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(SjQty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value)) {
        //							SumSjQty = SumSjQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumSjQty = SumSjQty + oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumSjWt = SumSjWt + oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value;
        //						//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //							SumQty = SumQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;
        //						//UPGRADE_WARNING: oMat01.Columns(UnWeight).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						HWeight = HWeight + (oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value * oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value / 1000);
        //					}
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumSjQty").Specific.Value = SumSjQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumSjWt").Specific.Value = SumSjWt;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("HWeight").Specific.Value = HWeight;


        //					if ((oMat01.RowCount > 1)) {
        //						oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						oForm.Items.Item("CardCode").Enabled = false;
        //						oForm.Items.Item("BPLId").Enabled = false;
        //						oForm.Items.Item("DCardCod").Enabled = false;
        //						oForm.Items.Item("TrType").Enabled = false;
        //					} else {
        //						oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //						oForm.Items.Item("CardCode").Enabled = true;
        //						oForm.Items.Item("BPLId").Enabled = true;
        //						oForm.Items.Item("DCardCod").Enabled = true;
        //						oForm.Items.Item("TrType").Enabled = true;
        //					}
        //					oForm.Update();
        //					//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					RecordSet01 = null;
        //					//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					RecordSet02 = null;
        //				} else if (pVal.ColUID == "Qty") {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)) {
        //						oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(0));
        //						oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(0));
        //						oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(0));
        //					} else {
        //						ItemCode01 = Strings.Trim(oDS_PS_SD040L.GetValue("U_ItemCode", pVal.Row - 1));
        //						////EA자체품
        //						if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
        //						////EAUOM
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(MDC_PS_Common.GetItem_Unit1(ItemCode01))));
        //						////KGSPEC
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
        //						////KG단중
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
        //						////KG선택
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //						}
        //						oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(Strings.Trim(oDS_PS_SD040L.GetValue("U_Weight", pVal.Row - 1))) * Convert.ToDouble(Strings.Trim(oDS_PS_SD040L.GetValue("U_Price", pVal.Row - 1)))));
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //					}

        //					for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(SjQty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value)) {
        //							SumSjQty = SumSjQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumSjQty = SumSjQty + oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumSjWt = SumSjWt + oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value;
        //						//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //							SumQty = SumQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;
        //						//UPGRADE_WARNING: oMat01.Columns(UnWeight).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						HWeight = HWeight + (oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value * oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value / 1000);
        //					}
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumSjQty").Specific.Value = SumSjQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumSjWt").Specific.Value = SumSjWt;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("HWeight").Specific.Value = HWeight;

        //				} else if (pVal.ColUID == "Weight") {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)) {
        //						oDS_PS_SD040L.SetValue("U_Qty", pVal.Row - 1, Convert.ToString(0));
        //						oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(0));
        //						oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(0));
        //					} else {
        //						ItemCode01 = Strings.Trim(oDS_PS_SD040L.GetValue("U_ItemCode", pVal.Row - 1));
        //						if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "101")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //						////EAUOM
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "102")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //						////KGSPEC
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "201")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Conversion.Val(MDC_PS_Common.GetItem_Spec1(ItemCode01)) - Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01))) * Conversion.Val(MDC_PS_Common.GetItem_Spec2(ItemCode01)) * 0.02808 * (Conversion.Val(MDC_PS_Common.GetItem_Spec3(ItemCode01)) / 1000) * Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
        //						////KG단중
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "202")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Conversion.Val(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(MDC_PS_Common.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
        //						////KG선택
        //						} else if ((MDC_PS_Common.GetItem_SbasUnit(ItemCode01) == "203")) {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //						}
        //						oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(Strings.Trim(oDS_PS_SD040L.GetValue("U_Weight", pVal.Row - 1))) * Convert.ToDouble(Strings.Trim(oDS_PS_SD040L.GetValue("U_Price", pVal.Row - 1)))));
        //					}

        //					for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //						//UPGRADE_WARNING: oMat01.Columns(SjQty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value)) {
        //							SumSjQty = SumSjQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumSjQty = SumSjQty + oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumSjWt = SumSjWt + oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value;
        //						//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //							SumQty = SumQty;
        //						} else {
        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //						}
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;
        //						//UPGRADE_WARNING: oMat01.Columns(UnWeight).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						HWeight = HWeight + (oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value * oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value / 1000);
        //					}
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumSjQty").Specific.Value = SumSjQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumSjWt").Specific.Value = SumSjWt;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oForm.Items.Item("HWeight").Specific.Value = HWeight;

        //				} else if (pVal.ColUID == "Price") {
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if ((Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)) {
        //						oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(0));
        //						oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(0));
        //					} else {
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //						//UPGRADE_WARNING: oMat01.Columns(Weight).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value * oMat01.Columns.Item("Weight").Cells.Item(pVal.Row).Specific.Value));
        //					}
        //					//                ElseIf pVal.ColUID = "WhsCode" Then
        //					//                    Call oDS_PS_SD040L.setValue("U_" & pVal.ColUID, pVal.Row - 1, oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value)
        //					//                    Call oDS_PS_SD040L.setValue("U_WhsName", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT WhsName FROM [OWHS] WHERE WhsCode = '" & oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value & "'", 0, 1))
        //				} else {
        //					oForm.Items.Item("TranCard").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
        //				}
        //				oMat01.LoadFromDataSource();
        //				oMat01.AutoResizeColumns();
        //				oForm.Update();
        //				oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //			} else {
        //				if ((pVal.ItemUID == "DocEntry")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD040H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //				} else if ((pVal.ItemUID == "CntcCode")) {
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDS_PS_SD040H.SetValue("U_CntcName", 0, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
        //				}
        //			}
        //		}
        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Exit:
        //	oForm.Freeze(false);
        //	return;
        //	Raise_EVENT_VALIDATE_Error:
        //	oForm.Freeze(false);
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int SumSjQty = 0;
        //	int i = 0;
        //	int SumQty = 0;
        //	decimal SumWeight = default(decimal);
        //	decimal SumSjWt = default(decimal);
        //	decimal HWeight = default(decimal);
        //	if (pVal.BeforeAction == true) {
        //	} else if (pVal.BeforeAction == false) {
        //		PS_SD040_FormItemEnabled();
        //		PS_SD040_AddMatrixRow(oMat01.VisualRowCount);
        //		////UDO방식

        //		for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
        //			//UPGRADE_WARNING: oMat01.Columns(SjQty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value)) {
        //				SumSjQty = SumSjQty;
        //			} else {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				SumSjQty = SumSjQty + oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value;
        //			}
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SumSjWt = SumSjWt + oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value;
        //			//UPGRADE_WARNING: oMat01.Columns(Qty).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value)) {
        //				SumQty = SumQty;
        //			} else {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				SumQty = SumQty + oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value;
        //			}
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			SumWeight = SumWeight + oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value;
        //			//UPGRADE_WARNING: oMat01.Columns(UnWeight).Cells(i + 1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			HWeight = HWeight + (oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value * oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value / 1000);
        //		}
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("SumSjQty").Specific.Value = SumSjQty;
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("SumSjWt").Specific.Value = SumSjWt;
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("SumQty").Specific.Value = SumQty;
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("SumWeight").Specific.Value = SumWeight;
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oForm.Items.Item("HWeight").Specific.Value = HWeight;
        //	}
        //	return;
        //	Raise_EVENT_MATRIX_LOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RESIZE
        //private void Raise_EVENT_RESIZE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {

        //	}
        //	return;
        //	Raise_EVENT_RESIZE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	SAPbouiCOM.DataTable oDataTable01 = null;
        //	if (pVal.BeforeAction == true) {

        //	} else if (pVal.BeforeAction == false) {
        //		if ((pVal.ItemUID == "CardCode" | pVal.ItemUID == "CardName")) {
        //			MDC_Com.MDC_GP_CF_DBDatasourceReturn(pVal, (pVal.FormUID), "@PS_SD040H", "U_CardCode,U_CardName");
        //		}
        //		if ((pVal.ItemUID == "DCardCod" | pVal.ItemUID == "DCardNam")) {
        //			MDC_Com.MDC_GP_CF_DBDatasourceReturn(pVal, (pVal.FormUID), "@PS_SD040H", "U_DCardCod,U_DCardNam");
        //		}
        //		if ((pVal.ItemUID == "Mat01")) {
        //			if ((pVal.ColUID == "WhsCode")) {
        //				//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (pVal.SelectedObjects == null) {
        //				} else {
        //					//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					oDataTable01 = pVal.SelectedObjects;
        //					oDS_PS_SD040L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
        //					oDS_PS_SD040L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
        //					//UPGRADE_NOTE: oDataTable01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //					oDataTable01 = null;
        //					//Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
        //					oMat01.LoadFromDataSource();
        //					oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				}
        //			}
        //		}
        //	}
        //	return;
        //	Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.ItemUID == "Mat01") {
        //		if (pVal.Row > 0) {
        //			oLastItemUID01 = pVal.ItemUID;
        //			oLastColUID01 = pVal.ColUID;
        //			oLastColRow01 = pVal.Row;
        //		}
        //	} else {
        //		oLastItemUID01 = pVal.ItemUID;
        //		oLastColUID01 = "";
        //		oLastColRow01 = 0;
        //	}
        //	return;
        //	Raise_EVENT_GOT_FOCUS_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	if (pVal.BeforeAction == true) {
        //	} else if (pVal.BeforeAction == false) {
        //		SubMain.RemoveForms(oFormUniqueID);
        //		//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oForm = null;
        //		//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oMat01 = null;
        //	}
        //	return;
        //	Raise_EVENT_FORM_UNLOAD_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	int i = 0;
        //	if ((oLastColRow01 > 0)) {
        //		if (pVal.BeforeAction == true) {
        //			////행삭제전 행삭제가능여부검사
        //			////추가,수정모드일때행삭제가능검사
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
        //				MDC_Com.MDC_GF_Message(ref "납품된행은 삭제할수 없습니다.", ref "W");
        //				BubbleEvent = false;
        //				return;
        //			}
        //			//UPGRADE_WARNING: oForm.Items(Opt03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oForm.Items(Opt02).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////멀티게이지, 분말
        //			if (oForm.Items.Item("Opt02").Specific.Selected == true | oForm.Items.Item("Opt03").Specific.Selected == true) {
        //				for (i = 0; i <= oDS_PS_SD040L.Size - 1; i++) {
        //					if (i == oDS_PS_SD040L.Size - 1) {
        //						break; // TODO: might not be correct. Was : Exit For
        //					}
        //					////선택된행과 같은값을 가지는 모든행
        //					if (oDS_PS_SD040L.GetValue("U_PackNo", i) == oDS_PS_SD040L.GetValue("U_PackNo", oLastColRow01 - 1)) {
        //						oDS_PS_SD040L.RemoveRecord(i);
        //						i = i - 1;
        //					}
        //				}
        //				for (i = 0; i <= oDS_PS_SD040L.Size - 1; i++) {
        //					oDS_PS_SD040L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
        //				}
        //				oMat01.LoadFromDataSource();
        //				////행삭제 정상처리

        //				if (oMat01.RowCount == 0) {
        //					PS_SD040_AddMatrixRow(0);
        //				} else {
        //					if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD040L.GetValue("U_SD030Num", oMat01.RowCount - 1)))) {
        //						PS_SD040_AddMatrixRow(oMat01.RowCount);
        //					}
        //				}
        //				if ((oMat01.RowCount > 1)) {
        //					oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					oForm.Items.Item("CardCode").Enabled = false;
        //					oForm.Items.Item("BPLId").Enabled = false;
        //					oForm.Items.Item("TrType").Enabled = false;
        //				} else {
        //					oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //					oForm.Items.Item("CardCode").Enabled = true;
        //					oForm.Items.Item("BPLId").Enabled = true;
        //					oForm.Items.Item("DCardCod").Enabled = true;
        //					oForm.Items.Item("TrType").Enabled = true;
        //				}
        //				oForm.Update();
        //				//                MDC_Com.MDC_GF_Message "멀티게이지는 포장번호 단위로 납품가능합니다. 행삭제 할수 없습니다.", "W"
        //				BubbleEvent = false;
        //				////멀티의경우 이미 행삭제가 되었으므로 행삭제를 하지 않는다.
        //				return;
        //			}
        //		} else if (pVal.BeforeAction == false) {
        //			for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //			}
        //			oMat01.FlushToDataSource();
        //			oDS_PS_SD040L.RemoveRecord(oDS_PS_SD040L.Size - 1);
        //			oMat01.LoadFromDataSource();
        //			if (oMat01.RowCount == 0) {
        //				PS_SD040_AddMatrixRow(0);
        //			} else {
        //				if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_SD040L.GetValue("U_SD030Num", oMat01.RowCount - 1)))) {
        //					PS_SD040_AddMatrixRow(oMat01.RowCount);
        //				}
        //			}
        //			if ((oMat01.RowCount > 1)) {
        //				oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oForm.Items.Item("CardCode").Enabled = false;
        //				oForm.Items.Item("BPLId").Enabled = false;
        //				oForm.Items.Item("TrType").Enabled = false;
        //			} else {
        //				oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //				oForm.Items.Item("CardCode").Enabled = true;
        //				oForm.Items.Item("BPLId").Enabled = true;
        //				oForm.Items.Item("DCardCod").Enabled = true;
        //				oForm.Items.Item("TrType").Enabled = true;
        //			}
        //			oForm.Update();
        //		}
        //	}
        //	return;
        //	Raise_EVENT_ROW_DELETE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion








        #region PS_SD040_MTX01
        //private void PS_SD040_MTX01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	////메트릭스에 데이터 로드
        //	int i = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string Param01 = null;
        //	string Param02 = null;
        //	string Param03 = null;
        //	string Param04 = null;
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param01 = Strings.Trim(oForm.Items.Item("Param01").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param02 = Strings.Trim(oForm.Items.Item("Param01").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param03 = Strings.Trim(oForm.Items.Item("Param01").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	Param04 = Strings.Trim(oForm.Items.Item("Param01").Specific.Value);

        //	Query01 = "SELECT 10";
        //	RecordSet01.DoQuery(Query01);

        //	if ((RecordSet01.RecordCount == 0)) {
        //		MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //		goto PS_SD040_MTX01_Exit;
        //	}
        //	oMat01.Clear();
        //	oMat01.FlushToDataSource();
        //	oMat01.LoadFromDataSource();

        //	SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //	ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

        //	for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
        //		if (i != 0) {
        //			oDS_PS_SD040L.InsertRecord((i));
        //		}
        //		oDS_PS_SD040L.Offset = i;
        //		oDS_PS_SD040L.SetValue("U_COL01", i, RecordSet01.Fields.Item(0).Value);
        //		oDS_PS_SD040L.SetValue("U_COL02", i, RecordSet01.Fields.Item(1).Value);
        //		RecordSet01.MoveNext();
        //		ProgressBar01.Value = ProgressBar01.Value + 1;
        //		ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
        //	}
        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();
        //	oForm.Update();

        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return;
        //	PS_SD040_MTX01_Exit:
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return;
        //	PS_SD040_MTX01_Error:
        //	ProgressBar01.Stop();
        //	//UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgressBar01 = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD040_DI_API_01
        //private bool PS_SD040_DI_API_01()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object j = null;
        //	object i = null;
        //	object K = null;
        //	int m = 0;
        //	SAPbobsCOM.Documents oDIObject = null;
        //	int RetVal = 0;
        //	double NeedWeight = 0;
        //	////필요중량
        //	double RemainWeight = 0;
        //	////잔여중량
        //	double SelectedWeight = 0;
        //	////선택중량
        //	int LineNumCount = 0;
        //	int ResultDocNum = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Company.StartTransaction();

        //	////배치정보
        //	BatchInformation = new BatchInformations[1];
        //	BatchInformationCount = 0;
        //	////품목정보
        //	ItemInformation = new ItemInformations[1];
        //	ItemInformationCount = 0;
        //	for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //		Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].BatchNum = oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value;
        //		////배치정보가져가기
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Qty = oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Weight = oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Currency = oMat01.Columns.Item("Currency").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Price = oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].LineTotal = oMat01.Columns.Item("LinTotal").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD040HNum = oForm.Items.Item("DocEntry").Specific.Value;
        //		//Val(oMat01.Columns("SD030H").Cells(i).Specific.Value) '문서번호(2016.10.20 송명규 수정, 출하요청 문서번호>납품처리 문서번호)
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD040LNum = i;
        //		//Val(oMat01.Columns("LineID").Cells(i).Specific.Value) Val(oMat01.Columns("SD030L").Cells(i).Specific.Value) '라인번호(2016.10.20 송명규 수정, 출하요청 라인번호>납품처리 라인번호)
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ORDRNum = Conversion.Val(oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].RDR1Num = Conversion.Val(oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value);
        //		ItemInformation[ItemInformationCount].Check = false;
        //		ItemInformationCount = ItemInformationCount + 1;
        //	}

        //	LineNumCount = 0;
        //	oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm.Items.Item("BPLId").Specific.Selected.Value));
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardCod").Value = Strings.Trim(oForm.Items.Item("DCardCod").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardNam").Value = Strings.Trim(oForm.Items.Item("DCardNam").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_TradeType").Value = Strings.Trim(oForm.Items.Item("TrType").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "&&&&-&&-&&"));
        //	}
        //	//UPGRADE_WARNING: oForm.Items(DueDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDueDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DueDate").Specific.Value, "&&&&-&&-&&"));
        //	}

        //	for (i = 0; i <= ItemInformationCount - 1; i++) {
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (ItemInformation[i].Check == true) {
        //			goto Continue_First;
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (i != 0) {
        //			oDIObject.Lines.Add();
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.WarehouseCode = ItemInformation[i].WhsCode;
        //		oDIObject.Lines.BaseType = 17;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseEntry = ItemInformation[i].ORDRNum;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseLine = ItemInformation[i].RDR1Num;
        //		oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD040";
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.UserFields.Fields.Item("U_BaseEntry").Value = ItemInformation[i].SD040HNum;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.UserFields.Fields.Item("U_BaseLine").Value = ItemInformation[i].SD040LNum;

        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		for (j = i; j <= ItemInformationCount - 1; j++) {
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (ItemInformation[j].Check == true) {
        //				goto Continue_Second;
        //			}
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if ((ItemInformation[i].ItemCode != ItemInformation[j].ItemCode | ItemInformation[i].ORDRNum != ItemInformation[j].ORDRNum | ItemInformation[i].RDR1Num != ItemInformation[j].RDR1Num)) {
        //				goto Continue_Second;
        //			}
        //			////같은것
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value + ItemInformation[j].Qty;
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation[j].Weight;
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.Currency = ItemInformation[j].Currency;
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.Price = oDIObject.Lines.Price + ItemInformation[j].Price;
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDIObject.Lines.LineTotal = oDIObject.Lines.LineTotal + ItemInformation[j].LineTotal;
        //			//UPGRADE_WARNING: oForm.Items(Opt03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oForm.Items(TrType).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////일반
        //			if ((oForm.Items.Item("TrType").Specific.Value == "1") & oForm.Items.Item("Opt03").Specific.Selected != true) {
        //				////DoNothing
        //				//UPGRADE_WARNING: oForm.Items(Opt03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				//UPGRADE_WARNING: oForm.Items(TrType).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			////임가공, 분말일때
        //			} else if ((oForm.Items.Item("TrType").Specific.Value == "2") | oForm.Items.Item("Opt03").Specific.Selected == true) {
        //				//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (MDC_PS_Common.GetItem_ManBtchNum(ItemInformation[j].ItemCode) == "Y") {
        //					//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					NeedWeight = ItemInformation[j].Weight;
        //					//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					Query01 = "EXEC PS_SD040_03 '" + ItemInformation[j].ItemCode + "','" + ItemInformation[j].BatchNum + "','" + ItemInformation[j].WhsCode + "'";
        //					RecordSet01.DoQuery(Query01);
        //					////재고배치조회
        //					for (K = 1; K <= RecordSet01.RecordCount; K++) {
        //						if ((NeedWeight <= 0)) {
        //							break; // TODO: might not be correct. Was : Exit For
        //						}
        //						//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //						RemainWeight = RecordSet01.Fields.Item("Weight").Value;
        //						for (m = 0; m <= BatchInformationCount - 1; m++) {
        //							if ((RecordSet01.Fields.Item("ItemCode").Value == BatchInformation[m].ItemCode & RecordSet01.Fields.Item("BatchNum").Value == BatchInformation[m].BatchNum & RecordSet01.Fields.Item("WhsCode").Value == BatchInformation[m].WhsCode)) {
        //								RemainWeight = RemainWeight - BatchInformation[m].Weight;
        //							}
        //						}
        //						////배치의중량이 남아있으면
        //						if ((RemainWeight > 0)) {

        //							//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							////부품
        //							if (MDC_PS_Common.GetItem_ItmBsort(ItemInformation[j].ItemCode) == "102" | MDC_PS_Common.GetItem_ItmBsort(ItemInformation[j].ItemCode) == "111") {
        //								////중량이 선택가능한 상황에서 여러배치를 선택할수 있다.
        //								//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet01.Fields.Item("BatchNum").Value;
        //								if ((NeedWeight > RemainWeight)) {
        //									SelectedWeight = RemainWeight;
        //									oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight;
        //								} else if ((NeedWeight <= RemainWeight)) {
        //									SelectedWeight = NeedWeight;
        //									oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight;
        //								}
        //								oDIObject.Lines.BatchNumbers.Add();
        //								NeedWeight = NeedWeight - RemainWeight;
        //								Array.Resize(ref BatchInformation, BatchInformationCount + 1);
        //								//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								BatchInformation[BatchInformationCount].ItemCode = RecordSet01.Fields.Item("ItemCode").Value;
        //								//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								BatchInformation[BatchInformationCount].BatchNum = RecordSet01.Fields.Item("BatchNum").Value;
        //								//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //								BatchInformation[BatchInformationCount].WhsCode = RecordSet01.Fields.Item("WhsCode").Value;
        //								BatchInformation[BatchInformationCount].Weight = SelectedWeight;
        //								BatchInformationCount = BatchInformationCount + 1;
        //								//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //							////MULTI
        //							} else if (MDC_PS_Common.GetItem_ItmBsort(ItemInformation[j].ItemCode) == "104" | MDC_PS_Common.GetItem_ItmBsort(ItemInformation[j].ItemCode) == "302") {
        //								////필요중량과 배치의 중량이 같아야 선택가능, 다르면 선택불가
        //								if ((NeedWeight == RemainWeight)) {
        //									SelectedWeight = NeedWeight;
        //									//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet01.Fields.Item("BatchNum").Value;
        //									oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight;
        //									oDIObject.Lines.BatchNumbers.Add();
        //									NeedWeight = NeedWeight - RemainWeight;
        //									Array.Resize(ref BatchInformation, BatchInformationCount + 1);
        //									//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									BatchInformation[BatchInformationCount].ItemCode = RecordSet01.Fields.Item("ItemCode").Value;
        //									//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									BatchInformation[BatchInformationCount].BatchNum = RecordSet01.Fields.Item("BatchNum").Value;
        //									//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //									BatchInformation[BatchInformationCount].WhsCode = RecordSet01.Fields.Item("WhsCode").Value;
        //									BatchInformation[BatchInformationCount].Weight = SelectedWeight;
        //									BatchInformationCount = BatchInformationCount + 1;
        //								}
        //							}
        //						////배치가 이미 선택된경우
        //						} else {
        //							////DoNothing 다음배치로이동
        //						}
        //						RecordSet01.MoveNext();
        //						////다음배치
        //					}
        //					if (NeedWeight > 0) {
        //						////MDC_Com.MDC_GF_Message "배치가 모두 선택되지 않았습니다. 재고를 확인하시기 바랍니다.", "W"
        //					}
        //				}
        //			}
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ItemInformation[j].DLN1Num = LineNumCount;
        //			//UPGRADE_WARNING: j 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			ItemInformation[j].Check = true;
        //			Continue_Second:
        //		}
        //		LineNumCount = LineNumCount + 1;
        //		Continue_First:
        //	}
        //	RetVal = oDIObject.Add();
        //	if (RetVal == 0) {
        //		ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //		for (i = 0; i <= ItemInformationCount - 1; i++) {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDS_PS_SD040L.SetValue("U_ODLNNum", i, Convert.ToString(ResultDocNum));
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDS_PS_SD040L.SetValue("U_DLN1Num", i, Convert.ToString(ItemInformation[i].DLN1Num));
        //			//            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD030H] SET U_ProgStat = '3' WHERE DocEntry = '" & ItemInformation(i).SD030HNum & "'") '//문서상태를 납품으로 변경
        //			//            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD030L] SET U_Status = 'C' WHERE DocEntry = '" & ItemInformation(i).SD030HNum & "' AND LineId = '" & ItemInformation(i).SD030LNum & "'") '//일부라도 납품되면 행이 클로즈된다.
        //		}
        //	} else {
        //		goto PS_SD040_DI_API_01_DI_Error;
        //	}

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //	}
        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();
        //	oForm.Update();
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //	PS_SD040_DI_API_01_DI_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_SD040_DI_API_01_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_DI_API_01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD040_DI_API_02
        //private bool PS_SD040_DI_API_02()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object j = null;
        //	object i = null;
        //	object K = null;
        //	int m = 0;
        //	SAPbobsCOM.Documents oDIObject = null;
        //	int RetVal = 0;
        //	double NeedWeight = 0;
        //	////필요중량
        //	double RemainWeight = 0;
        //	////잔여중량
        //	double SelectedWeight = 0;
        //	////선택중량
        //	int LineNumCount = 0;
        //	int ResultDocNum = 0;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);



        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Company.StartTransaction();

        //	////배치정보
        //	BatchInformation = new BatchInformations[1];
        //	BatchInformationCount = 0;
        //	////품목정보
        //	ItemInformation = new ItemInformations[1];
        //	ItemInformationCount = 0;
        //	for (i = 1; i <= oMat01.VisualRowCount; i++) {
        //		Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].BatchNum = oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value;
        //		////배치정보가져가기
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Qty = oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Weight = oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Currency = oMat01.Columns.Item("Currency").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].Price = oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].LineTotal = oMat01.Columns.Item("LinTotal").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD030HNum = Conversion.Val(oMat01.Columns.Item("SD030H").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD030LNum = Conversion.Val(oMat01.Columns.Item("SD030L").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ORDRNum = Conversion.Val(oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].RDR1Num = Conversion.Val(oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value);
        //		ItemInformation[ItemInformationCount].Check = false;
        //		ItemInformationCount = ItemInformationCount + 1;
        //	}

        //	LineNumCount = 0;
        //	oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm.Items.Item("BPLId").Specific.Selected.Value));
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardCod").Value = Strings.Trim(oForm.Items.Item("DCardCod").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardNam").Value = Strings.Trim(oForm.Items.Item("DCardNam").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_TradeType").Value = Strings.Trim(oForm.Items.Item("TrType").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "&&&&-&&-&&"));
        //	}
        //	//UPGRADE_WARNING: oForm.Items(DueDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDueDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DueDate").Specific.Value, "&&&&-&&-&&"));
        //	}

        //	for (i = 0; i <= ItemInformationCount - 1; i++) {
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (i != 0) {
        //			oDIObject.Lines.Add();
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.WarehouseCode = ItemInformation[i].WhsCode;

        //		oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD040";
        //		////별로 의미가 없을듯..
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = ItemInformation[i].Qty;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Quantity = ItemInformation[i].Weight;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Currency = ItemInformation[i].Currency;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Price = ItemInformation[i].Price;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.LineTotal = ItemInformation[i].LineTotal;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation[i].BatchNum;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BatchNumbers.Quantity = ItemInformation[i].Weight;
        //		oDIObject.Lines.BatchNumbers.Add();
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[i].DLN1Num = LineNumCount;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[i].Check = true;
        //		LineNumCount = LineNumCount + 1;
        //	}
        //	RetVal = oDIObject.Add();
        //	if (RetVal == 0) {
        //		ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //		for (i = 0; i <= ItemInformationCount - 1; i++) {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDS_PS_SD040L.SetValue("U_ODLNNum", i, Convert.ToString(ResultDocNum));
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			oDS_PS_SD040L.SetValue("U_DLN1Num", i, Convert.ToString(ItemInformation[i].DLN1Num));
        //		}
        //	} else {
        //		goto PS_SD040_DI_API_02_DI_Error;
        //	}

        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		if (PS_RFC_Sender() == true) {
        //			SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //		} else {
        //			goto PS_SD040_DI_API_02_DI_Error;
        //		}
        //	}

        //	oForm.Update();
        //	//    If Sbo_Company.InTransaction = True Then
        //	//
        //	//    End If

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //	PS_SD040_DI_API_02_DI_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}

        //	if (PS_RFC_Sender() == true) {
        //		SubMain.Sbo_Application.MessageBox("Di API 오류" + Err().Number + " - " + Err().Description + ")");
        //	}

        //	SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //	PS_SD040_DI_API_02_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_DI_API_02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_SD040_DI_API_03
        //private bool PS_SD040_DI_API_03()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;
        //	object j = null;
        //	object i = null;
        //	object K = null;
        //	int m = 0;
        //	SAPbobsCOM.Documents oDIObject = null;
        //	int RetVal = 0;
        //	double NeedWeight = 0;
        //	////필요중량
        //	double RemainWeight = 0;
        //	////잔여중량
        //	double SelectedWeight = 0;
        //	////선택중량
        //	int LineNumCount = 0;
        //	int ResultDocNum = 0;
        //	string BatchYN = null;

        //	string Query01 = null;
        //	string Query02 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	SAPbobsCOM.Recordset RecordSet02 = null;
        //	RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Company.StartTransaction();

        //	////배치정보
        //	BatchInformation = new BatchInformations[1];
        //	BatchInformationCount = 0;
        //	////품목정보
        //	ItemInformation = new ItemInformations[1];
        //	ItemInformationCount = 0;


        //	////배치번호 Y/N LotNo에 번호가 있으면 Batch = "Y"
        //	//UPGRADE_WARNING: oMat01.Columns(LotNo).Cells(1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (string.IsNullOrEmpty(oMat01.Columns.Item("LotNo").Cells.Item(1).Specific.Value)) {
        //		BatchYN = "N";
        //	} else {
        //		BatchYN = "Y";
        //	}

        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {

        //		Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ODLNNum = Conversion.Val(oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value);
        //		////납품
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].DLN1Num = Conversion.Val(oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD040HNum = Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value);
        //		////납품A
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD040LNum = Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value);
        //		ItemInformation[ItemInformationCount].Check = false;
        //		ItemInformationCount = ItemInformationCount + 1;

        //	}

        //	LineNumCount = 0;
        //	oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns);
        //	////반품DI '//납품의 취소시 반품, AR송장시, AR대변메모, AR송장이 지급시
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm.Items.Item("BPLId").Specific.Selected.Value));
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardCod").Value = Strings.Trim(oForm.Items.Item("DCardCod").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_DCardNam").Value = Strings.Trim(oForm.Items.Item("DCardNam").Specific.Value);
        //	//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oDIObject.UserFields.Fields.Item("U_TradeType").Value = Strings.Trim(oForm.Items.Item("TrType").Specific.Selected.Value);
        //	//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "&&&&-&&-&&"));
        //	}
        //	//UPGRADE_WARNING: oForm.Items(DueDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value)) {
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.DocDueDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DueDate").Specific.Value, "&&&&-&&-&&"));
        //	}

        //	for (i = 0; i <= ItemInformationCount - 1; i++) {

        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (i != 0) {
        //			//생산완료 등록 시 분할하여 등록한 데이터를 기준으로 납품처리 후 취소처리 할 때 판매오더의
        //			//이전 라인의 문서번호와 행번호를 체크하여 같으면 DI의 라인을 추가하지 않는 로직이 필요
        //			//해당 로직이 없으면 취소시 DI 에러 발생(2012.07.17 송명규 추가)
        //			//==>MG생산의 납품 취소시 에러가 남 아래 조건을 주석처리하면 반품처리가 됨. 배치번호문제인것 같음.
        //			//   배치제품이 안닌것만 조건을 걸어야 하지 않을까? 합니다. 2012.07.29 노근용
        //			if (BatchYN == "N") {
        //				////배치관리가 아닐때
        //				//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (ItemInformation[i - 1].ODLNNum != ItemInformation[i].ODLNNum & ItemInformation[i - 1].DLN1Num != ItemInformation[i].DLN1Num) {
        //					oDIObject.Lines.Add();
        //				}
        //			} else {
        //				//UPGRADE_WARNING: oForm.Items(Opt03).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if (oForm.Items.Item("Opt03").Specific.Selected == true) {
        //					//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					if (ItemInformation[i - 1].ODLNNum != ItemInformation[i].ODLNNum & ItemInformation[i - 1].DLN1Num != ItemInformation[i].DLN1Num) {
        //						////배치번호일때
        //						oDIObject.Lines.Add();
        //					}
        //				} else {
        //					oDIObject.Lines.Add();
        //				}
        //			}
        //		}
        //		////쿼리로 값구하기, 해당 납품라인에 대해 처리하기
        //		Query01 = "SELECT ";
        //		Query01 = Query01 + " ItemCode AS ItemCode,";
        //		Query01 = Query01 + " WhsCode AS WhsCode,";
        //		Query01 = Query01 + " U_Qty AS Qty,";
        //		Query01 = Query01 + " Quantity AS Weight,";
        //		Query01 = Query01 + " DocEntry AS ODLNNum,";
        //		Query01 = Query01 + " LineNum AS DLN1Num";
        //		Query01 = Query01 + " FROM";
        //		Query01 = Query01 + " [DLN1]";
        //		Query01 = Query01 + " WHERE";
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		Query01 = Query01 + " DocEntry = '" + ItemInformation[i].ODLNNum + "'";
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		Query01 = Query01 + " AND LineNum = '" + ItemInformation[i].DLN1Num + "'";
        //		////Query01 = Query01 & " AND LineStatus = 'O'"
        //		RecordSet01.DoQuery(Query01);

        //		//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.ItemCode = RecordSet01.Fields.Item("ItemCode").Value;
        //		//UPGRADE_WARNING: RecordSet01().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.WarehouseCode = RecordSet01.Fields.Item("WhsCode").Value;
        //		oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD040";
        //		////별로 의미가 없을듯..
        //		//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = RecordSet01.Fields.Item("Qty").Value;
        //		//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.Quantity = RecordSet01.Fields.Item("Weight").Value;
        //		oDIObject.Lines.BaseType = Convert.ToInt32("15");
        //		//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseEntry = RecordSet01.Fields.Item("ODLNNum").Value;
        //		//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseLine = RecordSet01.Fields.Item("DLN1Num").Value;
        //		if (MDC_PS_Common.GetItem_ManBtchNum(RecordSet01.Fields.Item("ItemCode").Value) == "Y") {
        //			////해당납품으로 출고된 배치내역 조회
        //			Query02 = "SELECT ";
        //			Query02 = Query02 + " BatchNum, Quantity";
        //			//Query02 = Query02 & " FROM [IBT1] "
        //			Query02 = Query02 + " FROM [IBT1_LINK] ";
        //			Query02 = Query02 + " WHERE BaseType = '15'";
        //			Query02 = Query02 + " AND BaseEntry = '" + RecordSet01.Fields.Item("ODLNNum").Value + "'";
        //			Query02 = Query02 + " AND BaseLinNum = '" + RecordSet01.Fields.Item("DLN1Num").Value + "'";
        //			RecordSet02.DoQuery(Query02);
        //			for (j = 0; j <= RecordSet02.RecordCount - 1; j++) {
        //				//UPGRADE_WARNING: RecordSet02.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet02.Fields.Item("BatchNum").Value;
        //				//UPGRADE_WARNING: RecordSet02.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oDIObject.Lines.BatchNumbers.Quantity = RecordSet02.Fields.Item("Quantity").Value;
        //				oDIObject.Lines.BatchNumbers.Add();
        //				RecordSet02.MoveNext();
        //			}
        //		}
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[i].RDN1Num = LineNumCount;
        //		LineNumCount = LineNumCount + 1;
        //		RecordSet01.MoveNext();
        //	}

        //	RetVal = oDIObject.Add();
        //	if (RetVal == 0) {
        //		ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());

        //		//        MDC_PS_Common.DoQuery ("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'")
        //		//        'Call RecordSet03.DoQuery("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'")
        //		//
        //		//        For i = 0 To ItemInformationCount - 1 '//납품문서번호 업데이트
        //		//            MDC_PS_Common.DoQuery ("UPDATE [@PS_SD040L] SET U_ORDNNum = '" & ResultDocNum & "', U_RDN1Num = '" & ItemInformation(i).RDN1Num & "' WHERE DocEntry = '" & ItemInformation(i).SD040HNum & "' AND LineId = '" & ItemInformation(i).SD040LNum & "'")
        //		//            'Call RecordSet03.DoQuery("UPDATE [@PS_SD040L] SET U_ORDNNum = '" & ResultDocNum & "', U_RDN1Num = '" & ItemInformation(i).RDN1Num & "' WHERE DocEntry = '" & ItemInformation(i).SD040HNum & "' AND LineId = '" & ItemInformation(i).SD040LNum & "'")
        //		//        Next

        //	} else {
        //		goto PS_SD040_DI_API_03_DI_Error;
        //	}

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

        //		//트랜젝션 속에 UPDATE 구문이 있어서 장시간 트랜젝션이 걸릴 경우 "시스템 호출 오류"가 발생하는 듯 함
        //		//트랜젝션 종료 후 정상적으로 Commit 되었을 때 상태 UPDATE 및 반품문서번호 UPDATE 되게 수정
        //		//2013.10.07 송명규 수정
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MDC_PS_Common.DoQuery(("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'"));
        //		//Call RecordSet03.DoQuery("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" & oForm.Items("DocEntry").Specific.Value & "'")

        //		////납품문서번호 업데이트
        //		for (i = 0; i <= ItemInformationCount - 1; i++) {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			MDC_PS_Common.DoQuery(("UPDATE [@PS_SD040L] SET U_ORDNNum = '" + ResultDocNum + "', U_RDN1Num = '" + ItemInformation[i].RDN1Num + "' WHERE DocEntry = '" + ItemInformation[i].SD040HNum + "' AND LineId = '" + ItemInformation[i].SD040LNum + "'"));
        //			//Call RecordSet03.DoQuery("UPDATE [@PS_SD040L] SET U_ORDNNum = '" & ResultDocNum & "', U_RDN1Num = '" & ItemInformation(i).RDN1Num & "' WHERE DocEntry = '" & ItemInformation(i).SD040HNum & "' AND LineId = '" & ItemInformation(i).SD040LNum & "'")
        //		}

        //	}
        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();
        //	oForm.Update();
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return functionReturnValue;
        //	PS_SD040_DI_API_03_DI_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return functionReturnValue;
        //	PS_SD040_DI_API_03_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_DI_API_03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	functionReturnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return functionReturnValue;
        //}
        #endregion



        #region PS_SD040_Print_Report01
        //private void PS_SD040_Print_Report01()
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocNum = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	int i = 0;
        //	string sQry01 = null;
        //	string Comments = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	//UPGRADE_WARNING: oForm.Items(ProgStat).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3") {
        //		MDC_Com.MDC_GF_Message(ref "문서상태가 납품이 아닙니다.", ref "W");
        //		return;
        //	}
        //	//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("BPLId").Specific.Selected.Value != "2") {
        //		MDC_Com.MDC_GF_Message(ref "사업장이 동래가 아닙니다.", ref "W");
        //		return;
        //	}
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		//UPGRADE_WARNING: oMat01.Columns(ItmBsort).Cells(i).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		if (oMat01.Columns.Item("ItmBsort").Cells.Item(i).Specific.Selected.Value != "105" & oMat01.Columns.Item("ItmBsort").Cells.Item(i).Specific.Selected.Value != "106") {
        //			MDC_Com.MDC_GF_Message(ref "품목이 기계공구,몰드가 아닙니다.", ref "W");
        //			return;
        //		}
        //	}

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry01 = "SELECT COUNT(*) AS COUNT FROM [@PS_SD040L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //	oRecordSet01.DoQuery(sQry01);
        //	if (Convert.ToDouble(Strings.Trim(oRecordSet01.Fields.Item(0).Value)) > 7) {

        //		WinTitle = "[PS_PP540_60] 레포트";
        //		ReportName = "PS_PP540_60.rpt";
        //		//sQry = "EXEC PS_PP540_60 '납품','" & oForm.Items("DocEntry").Specific.Value & "'"

        //		sQry = "EXEC PS_PP540_60 '납품','";
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sQry = sQry + oForm.Items.Item("DocEntry").Specific.Value + "','";
        //		sQry = sQry + SubMain.Sbo_Company.UserSignature + "'";
        //		MDC_Globals.gRpt_Formula = new string[2];
        //		MDC_Globals.gRpt_Formula_Value = new string[2];
        //		MDC_Globals.gRpt_SRptSqry = new string[2];
        //		MDC_Globals.gRpt_SRptName = new string[2];
        //		MDC_Globals.gRpt_SFormula = new string[2, 2];
        //		MDC_Globals.gRpt_SFormula_Value = new string[2, 2];


        //	} else {

        //		WinTitle = "[PS_PP540_10] 레포트";
        //		ReportName = "PS_PP540_10.rpt";
        //		sQry = "EXEC PS_PP540_10 '납품','";
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sQry = sQry + oForm.Items.Item("DocEntry").Specific.Value + "','";
        //		sQry = sQry + SubMain.Sbo_Company.UserSignature + "'";
        //		MDC_Globals.gRpt_Formula = new string[2];
        //		MDC_Globals.gRpt_Formula_Value = new string[2];
        //		MDC_Globals.gRpt_SRptSqry = new string[2];
        //		MDC_Globals.gRpt_SRptName = new string[2];
        //		MDC_Globals.gRpt_SFormula = new string[2, 2];
        //		MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //	}

        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	return;
        //	PS_SD040_Print_Report01_Error:

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_Print_Report01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD040_Print_Report02
        //private void PS_SD040_Print_Report02(string Chk)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocNum = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	string Sub_sQry = null;
        //	int i = 0;
        //	string sQry01 = null;
        //	string Comments = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	//UPGRADE_WARNING: oForm.Items(ProgStat).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3") {
        //		MDC_Com.MDC_GF_Message(ref "문서상태가 납품이 아닙니다.", ref "W");
        //		return;
        //	}
        //	//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1" & oForm.Items.Item("BPLId").Specific.Selected.Value != "4") {
        //		MDC_Com.MDC_GF_Message(ref "사업장이 창원,서울이 아닙니다.", ref "W");
        //		return;
        //	}
        //	//    For i = 1 To oMat01.VisualRowCount - 1
        //	//        If oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.Value <> "101" And oMat01.Columns("ItmBsort").Cells(i).Specific.Selected.Value <> "102" Then
        //	//            Call MDC_Com.MDC_GF_Message("품목이 휘팅,부품이 아닙니다.", "W")
        //	//            Exit Sub
        //	//        End If
        //	//    Next

        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry01 = "SELECT COUNT(*) AS COUNT FROM [@PS_SD040L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //	oRecordSet01.DoQuery(sQry01);
        //	if (Convert.ToDouble(Strings.Trim(oRecordSet01.Fields.Item(0).Value)) > 10) {
        //		//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oRecordSet01 = null;
        //		WinTitle = "[PS_SD220_20] 레포트";
        //		ReportName = "PS_SD220_20.rpt";
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sQry = "EXEC PS_SD220_20 '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //		MDC_Globals.gRpt_Formula = new string[5];
        //		MDC_Globals.gRpt_Formula_Value = new string[5];
        //		MDC_Globals.gRpt_SRptSqry = new string[2];
        //		MDC_Globals.gRpt_SRptName = new string[2];
        //		MDC_Globals.gRpt_SFormula = new string[2, 2];
        //		MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //		oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sQry01 = "Select BPLName FROM [OBPL] WHERE BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
        //		oRecordSet01.DoQuery(sQry01);
        //		MDC_Globals.gRpt_Formula[1] = "BPLName";
        //		MDC_Globals.gRpt_Formula[2] = "CardName";
        //		MDC_Globals.gRpt_Formula[3] = "DCardNam";
        //		MDC_Globals.gRpt_Formula[4] = "DocDate";
        //		//        Rpt_Formula(5) = "DocDate"

        //		MDC_Globals.gRpt_Formula_Value[1] = Strings.Trim(oRecordSet01.Fields.Item(0).Value);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MDC_Globals.gRpt_Formula_Value[2] = Strings.Trim(oForm.Items.Item("CardName").Specific.Value);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MDC_Globals.gRpt_Formula_Value[3] = Strings.Trim(oForm.Items.Item("DCardNam").Specific.Value);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MDC_Globals.gRpt_Formula_Value[4] = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "####-##-##");
        //	} else {
        //		WinTitle = "[PS_SD220_10] 레포트";
        //		ReportName = "PS_SD220_10.rpt";
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sQry = "EXEC PS_SD220_10 '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //		MDC_Globals.gRpt_Formula = new string[2];
        //		MDC_Globals.gRpt_Formula_Value = new string[2];
        //		MDC_Globals.gRpt_SRptSqry = new string[2];
        //		MDC_Globals.gRpt_SRptName = new string[2];
        //		MDC_Globals.gRpt_SFormula = new string[2, 2];
        //		MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //		MDC_Globals.gRpt_Formula[1] = "Chk";
        //		MDC_Globals.gRpt_Formula_Value[1] = Chk;

        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		Sub_sQry = "EXEC PS_SD220_11 '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //		MDC_Globals.gRpt_SRptSqry[1] = Sub_sQry;
        //		MDC_Globals.gRpt_SRptName[1] = "PS_SD220_SUB1";

        //	}

        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	return;
        //	PS_SD040_Print_Report02_Error:

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_Print_Report02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD040_Print_Report03
        //// 거래명세서(수량)
        //private void PS_SD040_Print_Report03(string Chk)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocNum = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	string Sub_sQry = null;
        //	int i = 0;
        //	string sQry01 = null;
        //	string Comments = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();

        //	//UPGRADE_WARNING: oForm.Items(ProgStat).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3") {
        //		MDC_Com.MDC_GF_Message(ref "문서상태가 납품이 아닙니다.", ref "W");
        //		return;
        //	}
        //	//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1" & oForm.Items.Item("BPLId").Specific.Selected.Value != "4") {
        //		MDC_Com.MDC_GF_Message(ref "사업장이 창원,서울이 아닙니다.", ref "W");
        //		return;
        //	}

        //	WinTitle = "분말거래명세표(수량)[PS_SD220_50]";
        //	ReportName = "PS_SD220_50.rpt";
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry = "EXEC PS_SD040_50 '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //	MDC_Globals.gRpt_Formula = new string[2];
        //	MDC_Globals.gRpt_Formula_Value = new string[2];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	return;
        //	PS_SD040_Print_Report03_Error:

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_Print_Report03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD040_Print_Report03_1
        //// 거래명세서(금액)
        //private void PS_SD040_Print_Report03_1(string Chk)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string DocNum = null;
        //	string WinTitle = null;
        //	string ReportName = null;
        //	string sQry = null;
        //	string Sub_sQry = null;
        //	int i = 0;
        //	string sQry01 = null;
        //	string Comments = null;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
        //	ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회 중...", 100, false);

        //	MDC_PS_Common.ConnectODBC();

        //	//UPGRADE_WARNING: oForm.Items(ProgStat).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3") {
        //		MDC_Com.MDC_GF_Message(ref "문서상태가 납품이 아닙니다.", ref "W");
        //		return;
        //	}
        //	//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1" & oForm.Items.Item("BPLId").Specific.Selected.Value != "4") {
        //		MDC_Com.MDC_GF_Message(ref "사업장이 창원,서울이 아닙니다.", ref "W");
        //		return;
        //	}

        //	WinTitle = "분말거래명세표(금액)[PS_SD220_51]";
        //	ReportName = "PS_SD220_51.rpt";
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sQry = "EXEC PS_SD040_50 '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
        //	MDC_Globals.gRpt_Formula = new string[2];
        //	MDC_Globals.gRpt_Formula_Value = new string[2];
        //	MDC_Globals.gRpt_SRptSqry = new string[2];
        //	MDC_Globals.gRpt_SRptName = new string[2];
        //	MDC_Globals.gRpt_SFormula = new string[2, 2];
        //	MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //	if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "Y", "V") == false) {
        //		SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	}

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	return;
        //	PS_SD040_Print_Report03_Error:

        //	ProgBar01.Value = 100;
        //	ProgBar01.Stop();
        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	ProgBar01 = null;

        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_Print_Report03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_SD0040_ChangePrice
        //public void PS_SD0040_ChangePrice()
        //{
        //	//******************************************************************************
        //	//Function ID : PS_SD0040_ChangePrice()
        //	//해당모듈 : PS_SD056
        //	//기능 : 단가수정
        //	//인수 : 없음
        //	//반환값 : 없음
        //	//특이사항 : 없음
        //	//******************************************************************************
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	string sQry = null;
        //	string DocEtnry = null;
        //	//기준년월

        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	//UPGRADE_WARNING: oForm.Items.Item().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEtnry = oForm.Items.Item("DocEntry").Specific.Value;

        //	sQry = "                EXEC [PS_SD040_08] ";
        //	sQry = sQry + "'" + DocEtnry + "'";
        //	//기준년월

        //	oRecordSet01.DoQuery(sQry);


        //	MDC_Com.MDC_GF_Message(ref "단가수정완료!", ref "S");

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;

        //	return;
        //	PS_SD0040_ChangePrice_Error:
        //	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	MDC_Com.MDC_GF_Message(ref "PS_SD056_DeleteDataAll_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //}
        #endregion


        #region PS_SD040_CheckDate
        //private bool PS_SD040_CheckDate(string pBaseEntry)
        //{
        //	bool functionReturnValue = false;
        //	//******************************************************************************
        //	//Function ID : PS_SD040_CheckDate()
        //	//해당모듈    : PS_SD040
        //	//기능        : 선행프로세스와 일자 비교
        //	//인수        : pBaseEntry : 기준문서번호
        //	//반환값      : True-선행프로세스보다 일자가 같거나 느릴 경우, False-선행프로세스보다 일자가 빠를 경우
        //	//특이사항    : 구현만 함, 사용은 하지 않음
        //	//******************************************************************************
        //	 // ERROR: Not supported in C#: OnErrorStatement


        //	string Query01 = null;
        //	short loopCount = 0;
        //	SAPbobsCOM.Recordset oRecordSet01 = null;
        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	string BaseEntry = null;
        //	string BaseLine = null;
        //	string DocType = null;
        //	string CurDocDate = null;

        //	string[] Entry = null;


        //	Entry = SAPbobsCOM.GTSResponseToExceedingEnum.Split(pBaseEntry, "-");
        //	BaseEntry = Entry[0];
        //	BaseLine = Entry[1];
        //	DocType = "PS_SD040";
        //	//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	CurDocDate = oMat01.Columns.Item("DocDate").Cells.Item(loopCount).Specific.Value;

        //	Query01 = "         EXEC PS_Z_CHECK_DATE '";
        //	Query01 = Query01 + BaseEntry + "','";
        //	Query01 = Query01 + BaseLine + "','";
        //	Query01 = Query01 + DocType + "','";
        //	Query01 = Query01 + CurDocDate + "'";

        //	oRecordSet01.DoQuery(Query01);

        //	if (oRecordSet01.Fields.Item("ReturnValue").Value == "False") {
        //		functionReturnValue = false;
        //	} else {
        //		functionReturnValue = true;
        //	}

        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oRecordSet01 = null;
        //	return functionReturnValue;
        //	PS_SD040_CheckDate_Error:

        //	functionReturnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD040_CheckDate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return functionReturnValue;
        //}
        #endregion

        #region PS_RFC_Sender
        //private bool PS_RFC_Sender()
        //{
        //	bool functionReturnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	functionReturnValue = true;

        //	string E_MESSAGE = null;
        //	string Client = null;
        //	//클라이언트(운영용:210, 테스트용:710)
        //	string ServerIP = null;
        //	//서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)

        //	object oFunction01 = null;
        //	SAPTableFactoryCtrl.Table sZSDN0001 = null;
        //	SAPTableFactoryCtrl.SAPTableFactory ddd = null;
        //	short i = 0;
        //	short j = 0;
        //	short l = 0;
        //	int Begew_i = 0;

        //	oMat01.FlushToDataSource();
        //	//Real
        //	Client = "210";
        //	ServerIP = "192.1.11.3";

        //	//Test
        //	//Client = "810"
        //	//ServerIP = "192.1.11.7"

        //	//1. 연결
        //	oSapConnection01 = Interaction.CreateObject("SAP.Functions");
        //	//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oSapConnection01.Connection.User = "ifuser";
        //	//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oSapConnection01.Connection.Password = "pdauser";
        //	//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oSapConnection01.Connection.Client = Client;
        //	//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oSapConnection01.Connection.ApplicationServer = ServerIP;
        //	//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oSapConnection01.Connection.Language = "KO";
        //	//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oSapConnection01.Connection.SystemNumber = "00";

        //	//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (!oSapConnection01.Connection.Logon(0, true)) {
        //		MDC_Com.MDC_GF_Message(ref "안강(R/3)서버에 접속할수 없습니다.", ref "E");
        //		goto RFC_Sender_Error;
        //	}

        //	////객체에 RFC Function을 할당
        //	//UPGRADE_WARNING: oSapConnection01.Add 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	oFunction01 = oSapConnection01.Add("ZPP_HOLDINGS_INTF_GI");

        //	////객체에 RFC Function 의 Table을 할당
        //	//UPGRADE_WARNING: oFunction01.Tables 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	sZSDN0001 = oFunction01.Tables.Item("ITAB");
        //	j = 1;
        //	for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {

        //		//UPGRADE_WARNING: sZSDN0001.Rows.Add 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sZSDN0001.Rows.Add();

        //		//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sZSDN0001[j, "ZLOTNO"] = Strings.Trim(oDS_PS_SD040L.GetValue("U_CoilNo", i));
        //		//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sZSDN0001[j, "ZORGDT"] = Strings.Trim(oDS_PS_SD040H.GetValue("U_DocDate", 0));
        //		//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sZSDN0001[j, "NTGEW"] = Strings.Trim(oDS_PS_SD040L.GetValue("U_Weight", i));
        //		//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sZSDN0001[j, "ZBOXWE"] = Strings.Trim(oDS_PS_SD040L.GetValue("U_PackWgt", i));
        //		//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		sZSDN0001[j, "BRGEW"] = Convert.ToDouble(oDS_PS_SD040L.GetValue("U_Weight", i)) + Convert.ToDouble(oDS_PS_SD040L.GetValue("U_PackWgt", i));

        //		// 마지막에 실행
        //		if (oMat01.VisualRowCount == i + 1) {
        //			//UPGRADE_WARNING: oFunction01.Call 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (!(oFunction01.Call)) {
        //				MDC_Com.MDC_GF_Message(ref "안강(R/3)서버 함수호출중 오류발생", ref "E");
        //				goto RFC_Sender_Error;
        //			} else {
        //				//UPGRADE_WARNING: oFunction01.Imports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				E_MESSAGE = oFunction01.Imports("E_MESSAGE").Value;
        //				if (Strings.Left(E_MESSAGE, 1) == "E") {
        //					goto RFC_Sender_Error;
        //				}
        //			}
        //		} else if (Strings.Trim(oDS_PS_SD040L.GetValue("U_PackNo", i)) != Strings.Trim(oDS_PS_SD040L.GetValue("U_PackNo", i + 1))) {
        //			//UPGRADE_WARNING: oFunction01.Call 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			if (!(oFunction01.Call)) {

        //				MDC_Com.MDC_GF_Message(ref "안강(R/3)서버 함수호출중 오류발생", ref "E");
        //				goto RFC_Sender_Error;

        //			} else {
        //				//UPGRADE_WARNING: oFunction01.Imports 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				E_MESSAGE = oFunction01.Imports("E_MESSAGE").Value;
        //				if (Strings.Left(E_MESSAGE, 1) == "E") {
        //					goto RFC_Sender_Error;
        //				}
        //				for (l = 1; l <= j; l++) {
        //					//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					sZSDN0001[l, "ZLOTNO"] = "";
        //					//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					sZSDN0001[l, "ZORGDT"] = "";
        //					//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					sZSDN0001[l, "NTGEW"] = "";
        //					//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					sZSDN0001[l, "ZBOXWE"] = "";
        //					//UPGRADE_WARNING: sZSDN0001() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //					sZSDN0001[l, "BRGEW"] = "";
        //				}
        //				j = 0;
        //			}
        //		}
        //		oDS_PS_SD040L.SetValue("U_TransYN", i, "Y");
        //		j = j + 1;
        //	}


        //	//UPGRADE_NOTE: sZSDN0001 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	sZSDN0001 = null;
        //	//UPGRADE_NOTE: oFunction01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oFunction01 = null;
        //	functionReturnValue = true;
        //	return functionReturnValue;
        //	RFC_Sender_Error:
        //	//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if ((oSapConnection01.Connection != null)) {
        //		//  If i = LastRow Then
        //		//UPGRADE_WARNING: oSapConnection01.Connection 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oSapConnection01.Connection.Logoff();
        //		//UPGRADE_NOTE: oSapConnection01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		oSapConnection01 = null;
        //		// End If
        //	}

        //	functionReturnValue = false;

        //	//UPGRADE_NOTE: oFunction01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oFunction01 = null;
        //	SubMain.Sbo_Application.MessageBox("R3 전송중 오류발생 (" + E_MESSAGE + ")");
        //	return functionReturnValue;
        //	//Sbo_Application.SetStatusBarMessage "RFC_Sender_Error: " & Err.Number & " - " & Err.Description, bmt_Short, True
        //}
        #endregion
    }
}
