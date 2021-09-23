using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using SAP.Middleware.Connector;
using System.Collections.Generic;

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
			//public int ORDNNum; //반품문서(DI API 실행 후 리턴받은 DocNum을 바로 UPDATE 하기 때문에 불필요)
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
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
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
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("ProgStat").Specific, "PS_PS_SD040", "ProgStat", false);
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

                oCFLCreationParams.ObjectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
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

                oCFLCreationParams01.ObjectType = "2"; //SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
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

                oCFLCreationParams.ObjectType = "64"; //SAPbouiCOM.BoLinkedObject.lf_Warehouses;
                oCFLCreationParams.UniqueID = "CFLWAREHOUSES";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);
                oColumn.ChooseFromListUID = "CFLWAREHOUSES";
                oColumn.ChooseFromListAlias = "WhsCode";
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
                    oForm.Items.Item("Button04").Enabled = false;

                    if (oForm.Items.Item("Opt03").Specific.Selected == true)
                    {
                        oForm.Items.Item("PriChange").Enabled = true;
                    }
                    else if (oForm.Items.Item("Opt02").Specific.Selected == true)
                    {
                        oForm.Items.Item("Button03").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("PriChange").Enabled = false;
                        oForm.Items.Item("Button03").Enabled = false;
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

                    if(oForm.Items.Item("Opt03").Specific.Selected == true)
                    {
                        oForm.Items.Item("PriChange").Enabled = true;
                    }
                    else if (oForm.Items.Item("Opt02").Specific.Selected == true)
                    {
                        oForm.Items.Item("Button03").Enabled = true;
                        oForm.Items.Item("Button04").Enabled = true;
                    }
                    else
                    {
                        oForm.Items.Item("PriChange").Enabled = false;
                        oForm.Items.Item("Button03").Enabled = false;
                        oForm.Items.Item("Button04").Enabled = false;
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
                        oForm.Items.Item("Button04").Enabled = false;
                        oForm.Items.Item("Button03").Enabled = true;
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
                    else if (oForm.Items.Item("Opt02").Specific.Selected == true)
                    {
                        oForm.Items.Item("Button04").Enabled = true;
                        oForm.Items.Item("Button03").Enabled = false;
                        oForm.Items.Item("PriChange").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("PriChange").Enabled = false;
                        oMat01.Columns.Item("Price").Editable = false;
                        oForm.Items.Item("Button04").Enabled = false;
                        oForm.Items.Item("Button03").Enabled = true;
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
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다." + (char)13 + "전기일을 확인하고, 회계부서로 문의하세요.";
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
                        DateTime sd030Date = Convert.ToDateTime(dataHelpClass.ConvertDateType(RecordSet01.Fields.Item(0).Value.ToString("yyyyMMdd"), "-")); //출하요청일
                        DateTime sd040Date = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-")); //납품처리일

                        if (sd030Date > sd040Date)
                        {
                            errMessage = "납품처리 전기일이 출하요청(" + oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value + ") 전기일보다 이전입니다." + (char)13 + "납품처리 전기일을 확인하십시오.";
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
                //여신한도초과 Check
                if (oForm.Items.Item("CardCode").Specific.Value != "12532") //풍산은 체크안함
                {
                    if (oForm.Mode == BoFormMode.fm_ADD_MODE) //최초 추가 모드일 때만 체크
                    {
                        if (PS_SD040_ValidateCreditLine() == false)
                        {
                            errMessage = " ";
                            throw new Exception();
                        }
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
                    //프로세스 체크용 외부 메소드의 실행 결과는 각 메소드에서 메시지 출력, 이 메소드에서는 메시지 처리 불필요
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
            double totalAmt = 0;
            double creditLine;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Items.Item("Opt01").Specific.Selected == true || oForm.Items.Item("Opt03").Specific.Selected == true) //MG제외
                {
                    query = "EXEC [PS_Z_CheckCreditLine] '";
                    query += oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "', '";
                    query += oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'";
                    RecordSet01.DoQuery(query);

                    if (RecordSet01.RecordCount > 0)
                    {
                        creditLine = Convert.ToDouble(RecordSet01.Fields.Item("OverAmt").Value);

                        //전체 금액 저장
                        for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            totalAmt += Convert.ToDouble(oMat01.Columns.Item("LinTotal").Cells.Item(i).Specific.Value);
                        }

                        if (creditLine < totalAmt) //여신한도(출고가능금액)보다 전체금액이 크면 출고 불가
                        {
                            errMessage = "해당 고객의 여신한도(현재 출고가능금액 : " + creditLine.ToString("#,##0.##") + ")를 초과했습니다." + (char)13 + "납품처리를 등록할 수 없습니다.";
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
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_SD040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할 수 없습니다.";
                    throw new Exception();
                }

                if (ValidateType == "검사")
                {
                    //입력된 행에 대해
                    for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (oForm.Items.Item("Opt01").Specific.Selected == true)
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD030H] PS_SD030H LEFT JOIN [@PS_SD030L] PS_SD030L ON PS_SD030H.DocEntry = PS_SD030L.DocEntry WHERE PS_SD030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_SD030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_SD030L.LineId) = '" + oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                            {
                                errMessage = "출하(선출)요청문서가 존재하지 않습니다.";
                                throw new Exception();
                            }
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
        /// 단가 수정
        /// </summary>
        private void PS_SD0040_ChangePrice()
        {
            string sQry;
            string DocEtnry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEtnry = oForm.Items.Item("DocEntry").Specific.Value;

                sQry = "EXEC [PS_SD040_08] ";
                sQry += "'" + DocEtnry + "'";
                oRecordSet01.DoQuery(sQry);

                PSH_Globals.SBO_Application.StatusBar.SetText("단가수정완료!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 선행프로세스와 일자 비교
        /// 구현만 하고 사용은 하지 않음
        /// </summary>
        /// <param name="pBaseEntry">기준문서번호</param>
        /// <returns>true-선행프로세스보다 일자가 같거나 느릴 경우, false-선행프로세스보다 일자가 빠를 경우</returns>
        private bool PS_SD040_CheckDate(string pBaseEntry)
        {
            bool returnValue = false;
            string query;
            string BaseEntry;
            string BaseLine;
            string DocType;
            string CurDocDate;
            string[] Entry;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                Entry = pBaseEntry.Split('-');
                BaseEntry = Entry[0];
                BaseLine = Entry[1];
                DocType = "PS_SD040";
                CurDocDate = oMat01.Columns.Item("DocDate").Cells.Item(0).Specific.Value;

                query = "EXEC PS_Z_CHECK_DATE '";
                query += BaseEntry + "','";
                query += BaseLine + "','";
                query += DocType + "','";
                query += CurDocDate + "'";

                oRecordSet01.DoQuery(query);

                if (oRecordSet01.Fields.Item("ReturnValue").Value == "False")
                {
                    returnValue = false;
                }
                else
                {
                    returnValue = true;
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 본사 데이터 전송
        /// </summary>
        /// <returns></returns>
        private void PS_SD040_InterfaceB1toR3()
        {
            string Client; //클라이언트(운영용:210, 테스트용:710)
            string ServerIP; //서버IP(운영용:192.1.11.3, 테스트용:192.1.11.7)
            int i;
            int j;
            string sQry;
            string errCode = string.Empty;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            RfcDestination rfcDest = null;
            RfcRepository rfcRep = null;

            try
            {
                oMat01.FlushToDataSource();
                ////Real
                Client = "210";
                ServerIP = "192.1.11.3";

                ////Test
                //Client = "810";
                //ServerIP = "192.1.11.7";

                //0. 연결
                if (dataHelpClass.SAPConnection(ref rfcDest, ref rfcRep, "PSC", ServerIP, Client, "ifuser", "pdauser") == false)
                {
                    errCode = "1";
                    throw new Exception();
                }

                //1. SAP R3 함수 호출(매개변수 전달)
                IRfcFunction oFunction = rfcRep.CreateFunction("ZPP_HOLDINGS_INTF_GI");
                IRfcTable oTable = oFunction.GetTable("ITAB"); //table 할당

                j = 1;
                oTable.Insert();
                for (i = 0; i <= oMat01.VisualRowCount -2; i++)
                {
                    if (oDS_PS_SD040L.GetValue("U_TransYN", i).ToString().Trim() != "Y")
                    {
                        //SetValue 매개변수용 변수(변수Type이 맞지 않으면 매개변수 전달시 SetValue 메소드 오류발생, 아래와 같이 매개변수에 값 저장후 SetValue에 전달)
                        string coilNo = oDS_PS_SD040L.GetValue("U_CoilNo", i).ToString().Trim();
                        string docDate = oDS_PS_SD040H.GetValue("U_DocDate", 0).ToString().Trim();
                        double weight = Convert.ToDouble(oDS_PS_SD040L.GetValue("U_Weight", i).ToString().Trim());
                        double packWgt = Convert.ToDouble(oDS_PS_SD040L.GetValue("U_PackWgt", i).ToString().Trim());
                        double totalWgt = Convert.ToDouble(oDS_PS_SD040L.GetValue("U_Weight", i)) + Convert.ToDouble(oDS_PS_SD040L.GetValue("U_PackWgt", i));

                        oTable.SetValue("ZLOTNO", coilNo);
                        oTable.SetValue("ZORGDT", docDate);
                        oTable.SetValue("NTGEW", weight);
                        oTable.SetValue("ZBOXWE", packWgt);
                        oTable.SetValue("BRGEW", totalWgt);
                        oTable.Append();

                        if (oMat01.VisualRowCount == i + 2) //마지막에 실행
                        {
                            errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
                            oFunction.Invoke(rfcDest); //Function 실행

                            errMessage = oFunction.GetValue("E_MESSAGE").ToString();

                            if (errMessage.Substring(0, 1) == "E")
                            {
                                sQry = "Update [@PS_SD040L] set U_TransYN ='E' where DocEntry ='" + oDS_PS_SD040L.GetValue("DocEntry", i).ToString().Trim() + "' and U_PackNo ='"+ oDS_PS_SD040L.GetValue("U_PackNo", i).ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                PSH_Globals.SBO_Application.MessageBox("패킹번호 [" + oDS_PS_SD040L.GetValue("U_PackNo", i).ToString().Trim() + "] 오류 :" + errMessage);
                            }
                            else
                            {
                                sQry = "Update [@PS_SD040L] set U_TransYN ='Y' where DocEntry ='" + oDS_PS_SD040L.GetValue("DocEntry", i).ToString().Trim() + "' and U_PackNo ='" + oDS_PS_SD040L.GetValue("U_PackNo", i).ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                        }
                        else if (oDS_PS_SD040L.GetValue("U_PackNo", i).ToString().Trim() != oDS_PS_SD040L.GetValue("U_PackNo", i + 1).ToString().Trim())
                        {
                            errCode = "2"; //SAP Function 실행 오류가 발생했을 때 에러코드로 처리하기 위해 이 위치에서 "2"를 대입
                            oFunction.Invoke(rfcDest); //Function 실행

                            errMessage = oFunction.GetValue("E_MESSAGE").ToString();

                            if (errMessage.Substring(0, 1) == "E")
                            {
                                sQry = "Update [@PS_SD040L] set U_TransYN ='E' where DocEntry ='" + oDS_PS_SD040L.GetValue("DocEntry", i).ToString().Trim() + "' and U_PackNo ='" + oDS_PS_SD040L.GetValue("U_PackNo", i).ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);

                                PSH_Globals.SBO_Application.MessageBox("패킹번호 [" + oDS_PS_SD040L.GetValue("U_PackNo", i).ToString().Trim() + "] 오류 :" +  errMessage);
                            }
                            else
                            {
                                sQry = "Update [@PS_SD040L] set U_TransYN ='Y' where DocEntry ='" + oDS_PS_SD040L.GetValue("DocEntry", i).ToString().Trim() + "' and U_PackNo ='" + oDS_PS_SD040L.GetValue("U_PackNo", i).ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                            }
                            oTable.Clear();
                            oTable.Insert();
                            j = 0;
                        }
                        j += 1;
                    }
                }
                PSH_Globals.SBO_Application.MessageBox("R3 인터페이스 완료!");
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("풍산 SAP R3에 로그온 할 수 없습니다. 관리자에게 문의 하세요.");
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("RFC Function 호출 오류");
                }
                else if (errCode == "3")
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 납품DI
        /// </summary>
        /// <returns></returns>
        private bool PS_SD040_DI_API_01()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int j;
            int i;
            int K;
            int m;
            int RetVal;
            double NeedWeight; //필요중량
            double RemainWeight; //잔여중량
            double SelectedWeight = 0; //선택중량
            int LineNumCount;
            string Query01;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<ItemInformation> itemInfoList = new List<ItemInformation>(); //품목정보
                List<BatchInformation> batchInfoList = new List<BatchInformation>(); //배치정보
                
                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    ItemInformation itemInfo = new ItemInformation
                    {
                        ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value,
                        BatchNum = oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value,
                        Qty = Convert.ToInt32(oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value),
                        Weight = Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value),
                        Currency = oMat01.Columns.Item("Currency").Cells.Item(i).Specific.Value,
                        Price = Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value),
                        LineTotal = Convert.ToDouble(oMat01.Columns.Item("LinTotal").Cells.Item(i).Specific.Value),
                        WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value,
                        SD040HNum = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value),
                        SD040LNum = i,
                        ORDRNum = Convert.ToInt32(oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value),
                        RDR1Num = Convert.ToInt32(oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value),
                        Check = false
                    };

                    itemInfoList.Add(itemInfo);
                }

                LineNumCount = 0;
                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim());
                oDIObject.CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_DCardCod").Value = oForm.Items.Item("DCardCod").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_DCardNam").Value = oForm.Items.Item("DCardNam").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_TradeType").Value = oForm.Items.Item("TrType").Specific.Selected.Value.ToString().Trim();

                if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                }
                if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value))
                {
                    oDIObject.DocDueDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DueDate").Specific.Value, "-"));
                }

                for (i = 0; i < itemInfoList.Count; i++)
                {
                    if (itemInfoList[i].Check == true)
                    {
                        continue;
                    }

                    if (i != 0)
                    {
                        oDIObject.Lines.Add();
                    }
                    oDIObject.Lines.ItemCode = itemInfoList[i].ItemCode;
                    oDIObject.Lines.WarehouseCode = itemInfoList[i].WhsCode;
                    oDIObject.Lines.BaseType = 17;
                    oDIObject.Lines.BaseEntry = itemInfoList[i].ORDRNum;
                    oDIObject.Lines.BaseLine = itemInfoList[i].RDR1Num;
                    oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD040";
                    oDIObject.Lines.UserFields.Fields.Item("U_BaseEntry").Value = itemInfoList[i].SD040HNum;
                    oDIObject.Lines.UserFields.Fields.Item("U_BaseLine").Value = itemInfoList[i].SD040LNum;

                    for (j = i; j < itemInfoList.Count; j++)
                    {
                        if (itemInfoList[j].Check == true || itemInfoList[i].ItemCode != itemInfoList[j].ItemCode || itemInfoList[i].ORDRNum != itemInfoList[j].ORDRNum || itemInfoList[i].RDR1Num != itemInfoList[j].RDR1Num)
                        {
                            continue;
                        }
                        
                        oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value + itemInfoList[j].Qty;
                        oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + itemInfoList[j].Weight;
                        oDIObject.Lines.Currency = itemInfoList[j].Currency;
                        oDIObject.Lines.Price = oDIObject.Lines.Price + itemInfoList[j].Price;
                        oDIObject.Lines.LineTotal = oDIObject.Lines.LineTotal + itemInfoList[j].LineTotal;
                        
                        if ((oForm.Items.Item("TrType").Specific.Value == "1") && oForm.Items.Item("Opt03").Specific.Selected != true) //일반
                        {
                            //DoNothing
                        }
                        else if ((oForm.Items.Item("TrType").Specific.Value == "2") || oForm.Items.Item("Opt03").Specific.Selected == true) //임가공, 분말일때
                        {
                            if (dataHelpClass.GetItem_ManBtchNum(itemInfoList[j].ItemCode) == "Y")
                            {
                                NeedWeight = itemInfoList[j].Weight;
                                Query01 = "EXEC PS_SD040_03 '" + itemInfoList[j].ItemCode + "','" + itemInfoList[j].BatchNum + "','" + itemInfoList[j].WhsCode + "'";
                                RecordSet01.DoQuery(Query01);
                                
                                for (K = 1; K <= RecordSet01.RecordCount; K++) //재고배치조회
                                {
                                    if (NeedWeight <= 0)
                                    {
                                        break;
                                    }

                                    RemainWeight = RecordSet01.Fields.Item("Weight").Value;
                                    for (m = 0; m < batchInfoList.Count; m++)
                                    {
                                        if (RecordSet01.Fields.Item("ItemCode").Value == batchInfoList[m].ItemCode && RecordSet01.Fields.Item("BatchNum").Value == batchInfoList[m].BatchNum && RecordSet01.Fields.Item("WhsCode").Value == batchInfoList[m].WhsCode)
                                        {
                                            RemainWeight -= batchInfoList[m].Weight;
                                        }
                                    }
                                    
                                    if (RemainWeight > 0) //배치의중량이 남아있으면
                                    {
                                        if (dataHelpClass.GetItem_ItmBsort(itemInfoList[j].ItemCode) == "102" || dataHelpClass.GetItem_ItmBsort(itemInfoList[j].ItemCode) == "111") //부품
                                        {
                                            //중량이 선택가능한 상황에서 여러배치 선택가능
                                            oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet01.Fields.Item("BatchNum").Value;

                                            if (NeedWeight > RemainWeight)
                                            {
                                                SelectedWeight = RemainWeight;
                                                oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight;
                                            }
                                            else if (NeedWeight <= RemainWeight)
                                            {
                                                SelectedWeight = NeedWeight;
                                                oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight;
                                            }
                                            oDIObject.Lines.BatchNumbers.Add();
                                            NeedWeight -= RemainWeight;

                                            BatchInformation batchInfo = new BatchInformation
                                            {
                                                ItemCode = RecordSet01.Fields.Item("ItemCode").Value,
                                                BatchNum = RecordSet01.Fields.Item("BatchNum").Value,
                                                WhsCode = RecordSet01.Fields.Item("WhsCode").Value,
                                                Weight = SelectedWeight
                                            };

                                            batchInfoList.Add(batchInfo);
                                        }
                                        else if (dataHelpClass.GetItem_ItmBsort(itemInfoList[j].ItemCode) == "104" || dataHelpClass.GetItem_ItmBsort(itemInfoList[j].ItemCode) == "302") //MULTI
                                        {
                                            //필요중량과 배치의 중량이 같아야 선택가능, 다르면 선택불가
                                            if (NeedWeight == RemainWeight)
                                            {
                                                SelectedWeight = NeedWeight;
                                                oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet01.Fields.Item("BatchNum").Value;
                                                oDIObject.Lines.BatchNumbers.Quantity = SelectedWeight;
                                                oDIObject.Lines.BatchNumbers.Add();
                                                NeedWeight -= RemainWeight;

                                                BatchInformation batchInfo = new BatchInformation
                                                {
                                                    ItemCode = RecordSet01.Fields.Item("ItemCode").Value,
                                                    BatchNum = RecordSet01.Fields.Item("BatchNum").Value,
                                                    WhsCode = RecordSet01.Fields.Item("WhsCode").Value,
                                                    Weight = SelectedWeight
                                                };

                                                batchInfoList.Add(batchInfo);
                                            }
                                        }
                                    }
                                    else //배치가 이미 선택된경우
                                    {
                                        //DoNothing 다음배치로이동
                                    }

                                    RecordSet01.MoveNext(); //다음배치
                                }
                            }
                        }
                        itemInfoList[j].DLN1Num = LineNumCount;
                        itemInfoList[j].Check = true;
                    }
                    LineNumCount += 1;
                }

                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                    for (i = 0; i < itemInfoList.Count; i++)
                    {
                        oDS_PS_SD040L.SetValue("U_ODLNNum", i, Convert.ToString(afterDIDocNum));
                        oDS_PS_SD040L.SetValue("U_DLN1Num", i, Convert.ToString(itemInfoList[i].DLN1Num));
                    }

                    //여신한도초과요청:납품처리여부 필드 업데이트(KEY-해당일자, 거래처코드)
                    Query01 = "  UPDATE	[@PS_SD080L]";
                    Query01 += " SET    U_SD040YN = 'Y'";
                    Query01 += "        FROM[@PS_SD080H] AS T0";
                    Query01 += "        INNER JOIN";
                    Query01 += "        [@PS_SD080L] AS T1";
                    Query01 += "            ON T0.DocEntry = T1.DocEntry";
                    Query01 += " WHERE  T0.U_DocDate = '" + oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'"; //해당일자
                    Query01 += "        AND T1.U_CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'"; //해당거래처코드
                    RecordSet01.DoQuery(Query01);
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                
                returnValue = true;
            }
            catch(Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 납품DI(멀티)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD040_DI_API_02()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int i;
            int RetVal;
            int LineNumCount;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<ItemInformation> itemInfoList = new List<ItemInformation>(); //품목정보
                List<BatchInformation> batchInfoList = new List<BatchInformation>(); //배치정보

                for (i = 1; i <= oMat01.VisualRowCount; i++)
                {
                    ItemInformation itemInfo = new ItemInformation
                    {
                        ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value,
                        BatchNum = oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value,
                        Qty = Convert.ToInt32(oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value),
                        Weight = Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value),
                        Currency = oMat01.Columns.Item("Currency").Cells.Item(i).Specific.Value,
                        Price = Convert.ToDouble(oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value),
                        LineTotal = Convert.ToDouble(oMat01.Columns.Item("LinTotal").Cells.Item(i).Specific.Value),
                        WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value,
                        SD030HNum = Convert.ToInt32(oMat01.Columns.Item("SD030H").Cells.Item(i).Specific.Value == "" ? "0" : oMat01.Columns.Item("SD030H").Cells.Item(i).Specific.Value),
                        SD030LNum = Convert.ToInt32(oMat01.Columns.Item("SD030L").Cells.Item(i).Specific.Value == "" ? "0" : oMat01.Columns.Item("SD030L").Cells.Item(i).Specific.Value),
                        ORDRNum = Convert.ToInt32(oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value == "" ? "0" : oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value),
                        RDR1Num = Convert.ToInt32(oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value == "" ? "0" : oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value),
                        Check = false
                    };
                    
                    itemInfoList.Add(itemInfo);
                }

                LineNumCount = 0;
                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim());
                oDIObject.CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_DCardCod").Value = oForm.Items.Item("DCardCod").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_DCardNam").Value = oForm.Items.Item("DCardNam").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_TradeType").Value = oForm.Items.Item("TrType").Specific.Selected.Value.ToString().Trim();

                if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                }

                if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value))
                {
                    oDIObject.DocDueDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DueDate").Specific.Value, "-"));
                }

                for (i = 0; i < itemInfoList.Count; i++)
                {
                    if (i != 0)
                    {
                        oDIObject.Lines.Add();
                    }
                    oDIObject.Lines.ItemCode = itemInfoList[i].ItemCode;
                    oDIObject.Lines.WarehouseCode = itemInfoList[i].WhsCode;

                    oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD040";
                    oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = itemInfoList[i].Qty;
                    oDIObject.Lines.Quantity = itemInfoList[i].Weight;
                    oDIObject.Lines.Currency = itemInfoList[i].Currency;
                    oDIObject.Lines.Price = itemInfoList[i].Price;
                    oDIObject.Lines.LineTotal = itemInfoList[i].LineTotal;
                    oDIObject.Lines.BatchNumbers.BatchNumber = itemInfoList[i].BatchNum;
                    oDIObject.Lines.BatchNumbers.Quantity = itemInfoList[i].Weight;
                    oDIObject.Lines.BatchNumbers.Add();
                    itemInfoList[i].DLN1Num = LineNumCount;
                    itemInfoList[i].Check = true;
                    LineNumCount += 1;
                }

                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                    for (i = 0; i < itemInfoList.Count; i++)
                    {
                        oDS_PS_SD040L.SetValue("U_ODLNNum", i, Convert.ToString(afterDIDocNum));
                        oDS_PS_SD040L.SetValue("U_DLN1Num", i, Convert.ToString(itemInfoList[i].DLN1Num));
                    }
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 반품(납품 취소)
        /// </summary>
        /// <returns></returns>
        private bool PS_SD040_DI_API_03()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int j;
            int i;
            int RetVal;
            int LineNumCount;
            string BatchYN;
            string Query01;
            string Query02;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Documents oDIObject = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                List<ItemInformation> itemInfoList = new List<ItemInformation>(); //품목정보
                List<BatchInformation> batchInfoList = new List<BatchInformation>(); //배치정보

                //배치번호 Y/N LotNo에 번호가 있으면 Batch = "Y"
                if (string.IsNullOrEmpty(oMat01.Columns.Item("LotNo").Cells.Item(1).Specific.Value))
                {
                    BatchYN = "N";
                }
                else
                {
                    BatchYN = "Y";
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    ItemInformation itemInfo = new ItemInformation
                    {
                        ODLNNum = Convert.ToInt32(oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value),
                        DLN1Num = Convert.ToInt32(oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value),
                        SD040HNum = Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value),
                        SD040LNum = Convert.ToInt32(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value),
                        Check = false
                    };

                    itemInfoList.Add(itemInfo);
                }

                LineNumCount = 0;
                oDIObject = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns);
                oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim());
                oDIObject.CardCode = oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_DCardCod").Value = oForm.Items.Item("DCardCod").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_DCardNam").Value = oForm.Items.Item("DCardNam").Specific.Value.ToString().Trim();
                oDIObject.UserFields.Fields.Item("U_TradeType").Value = oForm.Items.Item("TrType").Specific.Selected.Value.ToString().Trim();

                if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    oDIObject.DocDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-"));
                }
                if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value))
                {
                    oDIObject.DocDueDate = Convert.ToDateTime(dataHelpClass.ConvertDateType(oForm.Items.Item("DueDate").Specific.Value, "-"));
                }

                for (i = 0; i < itemInfoList.Count; i++)
                {
                    if (i != 0)
                    {
                        if (BatchYN == "N")
                        {
                            if (itemInfoList[i - 1].ODLNNum != itemInfoList[i].ODLNNum && itemInfoList[i - 1].DLN1Num != itemInfoList[i].DLN1Num) //배치관리가 아닐때
                            {
                                oDIObject.Lines.Add();
                            }
                        }
                        else
                        {
                            if (oForm.Items.Item("Opt03").Specific.Selected == true)
                            {
                                if (itemInfoList[i - 1].ODLNNum != itemInfoList[i].ODLNNum && itemInfoList[i - 1].DLN1Num != itemInfoList[i].DLN1Num) //배치번호일때
                                {
                                    oDIObject.Lines.Add();
                                }
                            }
                            else
                            {
                                oDIObject.Lines.Add();
                            }
                        }
                    }

                    //쿼리로 값구하기, 해당 납품라인에 대해 처리하기
                    Query01 = "  SELECT ItemCode AS ItemCode,";
                    Query01 += "        WhsCode AS WhsCode,";
                    Query01 += "        U_Qty AS Qty,";
                    Query01 += "        Quantity AS Weight,";
                    Query01 += "        DocEntry AS ODLNNum,";
                    Query01 += "        LineNum AS DLN1Num";
                    Query01 += " FROM   [DLN1]";
                    Query01 += " WHERE  DocEntry = '" + itemInfoList[i].ODLNNum + "'";
                    Query01 += "        AND LineNum = '" + itemInfoList[i].DLN1Num + "'";
                    RecordSet01.DoQuery(Query01);

                    oDIObject.Lines.ItemCode = RecordSet01.Fields.Item("ItemCode").Value;
                    oDIObject.Lines.WarehouseCode = RecordSet01.Fields.Item("WhsCode").Value;
                    oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD040";
                    oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = RecordSet01.Fields.Item("Qty").Value;
                    oDIObject.Lines.Quantity = RecordSet01.Fields.Item("Weight").Value;
                    oDIObject.Lines.BaseType = 15;
                    oDIObject.Lines.BaseEntry = RecordSet01.Fields.Item("ODLNNum").Value;
                    oDIObject.Lines.BaseLine = RecordSet01.Fields.Item("DLN1Num").Value;

                    if (dataHelpClass.GetItem_ManBtchNum(RecordSet01.Fields.Item("ItemCode").Value) == "Y")
                    {
                        //해당납품으로 출고된 배치내역 조회
                        Query02 = "  SELECT BatchNum, Quantity";
                        Query02 += " FROM   [IBT1_LINK] ";
                        Query02 += " WHERE  BaseType = '15'";
                        Query02 += "        AND BaseEntry = '" + RecordSet01.Fields.Item("ODLNNum").Value + "'";
                        Query02 += "        AND BaseLinNum = '" + RecordSet01.Fields.Item("DLN1Num").Value + "'";
                        RecordSet02.DoQuery(Query02);

                        for (j = 0; j <= RecordSet02.RecordCount - 1; j++)
                        {
                            oDIObject.Lines.BatchNumbers.BatchNumber = RecordSet02.Fields.Item("BatchNum").Value;
                            oDIObject.Lines.BatchNumbers.Quantity = RecordSet02.Fields.Item("Quantity").Value;
                            oDIObject.Lines.BatchNumbers.Add();
                            RecordSet02.MoveNext();
                        }
                    }

                    itemInfoList[i].RDN1Num = LineNumCount;
                    LineNumCount += 1;
                    RecordSet01.MoveNext();
                }

                RetVal = oDIObject.Add();

                if (RetVal == 0)
                {
                    PSH_Globals.oCompany.GetNewObjectCode(out string afterDIDocNum);

                    dataHelpClass.DoQuery("UPDATE [@PS_SD040H] SET U_ProgStat = '4' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'");

                    //납품문서번호 업데이트
                    for (i = 0; i < itemInfoList.Count; i++)
                    {
                        dataHelpClass.DoQuery("UPDATE [@PS_SD040L] SET U_ORDNNum = '" + afterDIDocNum + "', U_RDN1Num = '" + itemInfoList[i].RDN1Num + "' WHERE DocEntry = '" + itemInfoList[i].SD040HNum + "' AND LineId = '" + itemInfoList[i].SD040LNum + "'");
                    }

                    //여신한도초과요청:납품처리여부 필드 업데이트(KEY-해당일자, 거래처코드)
                    Query01 = "  UPDATE	[@PS_SD080L]";
                    Query01 += " SET    U_SD040YN = 'N'"; //납품처리여부 "N"으로 환원
                    Query01 += "        FROM[@PS_SD080H] AS T0";
                    Query01 += "        INNER JOIN";
                    Query01 += "        [@PS_SD080L] AS T1";
                    Query01 += "            ON T0.DocEntry = T1.DocEntry";
                    Query01 += " WHERE  T0.U_DocDate = '" + oForm.Items.Item("DocDate").Specific.Value.ToString().Trim() + "'"; //해당일자
                    Query01 += "        AND T1.U_CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'"; //해당거래처코드
                    RecordSet01.DoQuery(Query01);
                }
                else
                {
                    PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                    errCode = "1";
                    throw new Exception();
                }

                PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.MessageBox("DI실행 중 오류 발생 : [" + errDICode + "]" + (char)13 + errDIMsg);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.MessageBox("현재월의 전기기간이 잠겼습니다. 회계부서에 문의하세요.");
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
                if (oDIObject != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDIObject);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
            }

            return returnValue;
        }

        /// <summary>
        /// 거래명세표 출력
        /// </summary>
        [STAThread]
        private void PS_SD040_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            string sQry;
            int i;
            string errMessage = string.Empty;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3")
                {
                    errMessage = "문서상태가 납품이 아닙니다.";
                    throw new Exception();
                }
                else if (oForm.Items.Item("BPLId").Specific.Selected.Value != "2")
                {
                    errMessage = "사업장이 동래가 아닙니다.";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (oMat01.Columns.Item("ItmBsort").Cells.Item(i).Specific.Selected.Value != "105" && oMat01.Columns.Item("ItmBsort").Cells.Item(i).Specific.Selected.Value != "106")
                    {
                        errMessage = "품목이 기계공구,몰드가 아닙니다.";
                        throw new Exception();
                    }
                }
                
                sQry = "SELECT COUNT(*) AS COUNT FROM [@PS_SD040L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) > 7) //입력된 납품처리 테이블의 Line이 7행보다 클 경우(거래명세표 : 최대 7개의 품목만 표시)
                {
                    WinTitle = "[PS_PP540_60] 레포트";
                    ReportName = "PS_PP540_60.rpt";
                    //쿼리 : PS_PP540_60
                }
                else
                {
                    WinTitle = "[PS_PP540_10] 레포트";
                    ReportName = "PS_PP540_10.rpt";
                    //쿼리 : PS_PP540_10
                }

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocType", "납품"),
                    new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value),
                    new PSH_DataPackClass("@UserSign", PSH_Globals.oCompany.UserSignature)
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 출고증 출력
        /// </summary>
        /// <param name="Chk"></param>
        [STAThread]
        private void PS_SD040_Print_Report02(string Chk)
        {
            string WinTitle;
            string ReportName;
            string sQry;
            string errMessage = string.Empty;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3")
                {
                    errMessage = "문서상태가 납품이 아닙니다.";
                    throw new Exception();
                }
                if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1" & oForm.Items.Item("BPLId").Specific.Selected.Value != "4")
                {
                    errMessage = "사업장이 창원,서울이 아닙니다.";
                    throw new Exception();
                }

                sQry = "SELECT COUNT(*) AS COUNT FROM [@PS_SD040L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                oRecordSet01.DoQuery(sQry);

                if (Convert.ToDouble(oRecordSet01.Fields.Item(0).Value.ToString().Trim()) > 10)
                {
                    WinTitle = "[PS_SD220_20] 레포트";
                    ReportName = "PS_SD220_20.rpt";
                    //쿼리 : PS_SD220_20

                    sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
                    oRecordSet01.DoQuery(sQry);

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass> //Parameter List
                    {
                        new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value) //DocEntry
                    };

                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass> //Formula List
                    {
                        new PSH_DataPackClass("@BPLName", oRecordSet01.Fields.Item(0).Value), //BPLName(oForm.Items.Item("BPLId").Specific.String)
                        new PSH_DataPackClass("@CardName", oForm.Items.Item("CardName").Specific.Value), //CardName
                        new PSH_DataPackClass("@DCardNam", oForm.Items.Item("DCardNam").Specific.Value), //DCardNam
                        new PSH_DataPackClass("@DocDate", dataHelpClass.ConvertDateType(oForm.Items.Item("DocDate").Specific.Value, "-")) //DocDate
                    };

                    formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
                }
                else
                {
                    WinTitle = "[PS_SD220_10] 레포트";
                    ReportName = "PS_SD220_10.rpt";
                    //쿼리 : PS_SD220_10, 서브리포트 : PS_SD220_11

                    List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass> //Parameter List
                    { 
                        new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value) //DocEntry    
                    };

                    List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass> //Formula List
                    {
                        new PSH_DataPackClass("@Chk", Chk) //Chk
                    }; 

                    List<PSH_DataPackClass> dataPackSubReportParameter = new List<PSH_DataPackClass>  //SubReport Parameter List
                    {
                        new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value, "PS_SD220_SUB1") //DocEntry
                    };

                    formHelpClass.OpenCrystalReport(dataPackParameter, dataPackFormula, dataPackSubReportParameter, WinTitle, ReportName);
                }
            }
            catch (Exception ex)
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 거래명세서(수량)
        /// </summary>
        [STAThread]
        private void PS_SD040_Print_Report03()
        {
            string WinTitle;
            string ReportName;
            string errMessage = string.Empty;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3")
                {
                    errMessage = "문서상태가 납품이 아닙니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1" & oForm.Items.Item("BPLId").Specific.Selected.Value != "4")
                {
                    errMessage = "사업장이 창원,서울이 아닙니다.";
                    throw new Exception();
                }

                WinTitle = "분말거래명세표(수량)[PS_SD220_50]";
                ReportName = "PS_SD220_50.rpt";
                //쿼리 : PS_SD040_50

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass> //Parameter List
                {
                    new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value) //DocEntry
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
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
            }
        }

        /// <summary>
        /// 거래명세서(금액) 
        /// </summary>
        [STAThread]
        private void PS_SD040_Print_Report03_1()
        {
            string WinTitle;
            string ReportName;
            string errMessage = string.Empty;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3")
                {
                    errMessage = "문서상태가 납품이 아닙니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1" & oForm.Items.Item("BPLId").Specific.Selected.Value != "4")
                {
                    errMessage = "사업장이 창원,서울이 아닙니다.";
                    throw new Exception();
                }

                WinTitle = "분말거래명세표(금액)[PS_SD220_51]";
                ReportName = "PS_SD220_51.rpt";
                //쿼리 : PS_SD040_50

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass> //Parameter List
                {
                    new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value) //DocEntry
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter);
            }
            catch (Exception ex)
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
                            if (PS_SD040_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            //납품처리 문서 단독으로 입력할 경우 주석처리 ("납품"문서 생성 안함)_S
                            if (oForm.Items.Item("Opt01").Specific.Selected == true)
                            {
                                if (PS_SD040_DI_API_01() == false) //납품생성
                                {
                                    PS_SD040_AddMatrixRow(oMat01.RowCount, false);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else if (oForm.Items.Item("Opt02").Specific.Selected == true)
                            {
                                if (PS_SD040_DI_API_02() == false) //멀티게이지납품생성
                                {
                                    PS_SD040_AddMatrixRow(oMat01.RowCount, false);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else if (oForm.Items.Item("Opt03").Specific.Selected == true)
                            {
                                if (PS_SD040_DI_API_01() == false) //분말 납품생성
                                {
                                    PS_SD040_AddMatrixRow(oMat01.RowCount, false);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            //납품처리 문서 단독으로 입력할 경우 주석처리 ("납품"문서 생성 안함)_E
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD040_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "Button01") //거래명세서
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_SD040_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                    else if (pVal.ItemUID == "Button02") //출고증(수량)
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oForm.Items.Item("Opt03").Specific.Selected == true) //분말일경우
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(PS_SD040_Print_Report03);
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start();
                            }
                            else
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(() => PS_SD040_Print_Report02("N"));
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Button04") //B1 to R3 데이터 인터페이스
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oForm.Items.Item("Opt02").Specific.Selected == true) //분말일경우
                            {
                                PS_SD040_InterfaceB1toR3();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Button05") // 멀티 매트릭스 내 R3 전송 오류 메시지 초기화
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (oForm.Items.Item("Opt02").Specific.Selected == true) //분말일경우
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(PS_SD040_Print_Report03);
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start();
                            }
                            else
                            {
                                System.Threading.Thread thread = new System.Threading.Thread(() => PS_SD040_Print_Report02("N"));
                                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                                thread.Start();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "opt03")
                    {
                        PS_SD040_FormItemEnabled();
                    }
                    else if (pVal.ItemUID == "PriChange")
                    {
                        PS_SD0040_ChangePrice();
                        PS_SD040_AddMatrixRow(oMat01.RowCount, false);
                    }
                    else if (pVal.ItemUID == "Button02_1") //출고증(금액)
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_SD040_Print_Report03_1);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                            PS_SD0040_ChangePrice();
                        }
                    }
                    if (pVal.ItemUID == "Button03")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) //TODO : PS_MM004 구현 필요
                        {
                            PS_MM004 tempForm = new PS_MM004(); 
                            tempForm.LoadForm("PS_SD040", oForm.Items.Item("DocEntry").Specific.Value);
                            BubbleEvent = false;
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
                                PS_SD040_FormItemEnabled();
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                PSH_Globals.SBO_Application.ActivateMenuItem("1291");
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_SD040_FormItemEnabled();
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
                    
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "SD030Num")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("고객코드는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                BubbleEvent = false;
                                return;
                            }

                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "SD030Num");
                        }
                        else if (pVal.ColUID == "PackNo")
                        {
                            if (oForm.Items.Item("Opt02").Specific.Selected == true)
                            {
                                if (oForm.Items.Item("TrType").Specific.Selected.Value != "2")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("포장번호조회는 거래형태가 임가공이여야 합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }

                                if (string.IsNullOrEmpty(oForm.Items.Item("DCardCod").Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("납품처코드는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }
                            }

                            if (string.IsNullOrEmpty(oForm.Items.Item("CardCode").Specific.Value))
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("고객코드는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                BubbleEvent = false;
                                return;
                            }

                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "PackNo");
                        }
                    }
                }
                else if (pVal.Before_Action == false)
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
        /// GOT_FOCUS 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_GOT_FOCUS(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
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
                else if (pVal.Before_Action == false)
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                        }
                    }
                    else if (pVal.ItemUID == "Opt01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            oForm.Freeze(true);
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("CardCode").Enabled = true;
                            oForm.Items.Item("BPLId").Enabled = true;
                            oForm.Items.Item("DCardCod").Enabled = true;
                            oForm.Items.Item("TrType").Enabled = true;
                            oMat01.Columns.Item("SD030Num").Visible = true;
                            oMat01.Columns.Item("PackNo").Visible = false;
                            oMat01.Columns.Item("OrderNum").Visible = true;
                            oMat01.Columns.Item("SD030Num").Editable = true;
                            oMat01.Columns.Item("Qty").Editable = true;
                            oMat01.Columns.Item("Weight").Editable = true;
                            oMat01.Columns.Item("UnWeight").Visible = true;
                            oMat01.Columns.Item("Price").Editable = false;
                            oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_SD040_AddMatrixRow(0, true);

                            oForm.Items.Item("SumSjQty").Specific.Value = 0;
                            oForm.Items.Item("SumSjWt").Specific.Value = 0;
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Items.Item("SumWeight").Specific.Value = 0;
                            oForm.Items.Item("HWeight").Specific.Value = 0;
                            oForm.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Opt02") //멀티게이지
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            oForm.Freeze(true);
                            oForm.Items.Item("TranCard").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("CardCode").Enabled = true;
                            oForm.Items.Item("BPLId").Enabled = true;
                            oForm.Items.Item("DCardCod").Enabled = true;
                            oForm.Items.Item("TrType").Enabled = false;
                            oMat01.Columns.Item("SD030Num").Visible = false;
                            oMat01.Columns.Item("PackNo").Visible = true;
                            oMat01.Columns.Item("OrderNum").Visible = false;
                            oMat01.Columns.Item("Qty").Editable = false;
                            oMat01.Columns.Item("Weight").Editable = false;
                            oMat01.Columns.Item("UnWeight").Visible = false;
                            oMat01.Columns.Item("Price").Editable = true;
                            oForm.Items.Item("TrType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_SD040_AddMatrixRow(0, true);

                            oForm.Items.Item("SumSjQty").Specific.Value = 0;
                            oForm.Items.Item("SumSjWt").Specific.Value = 0;
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Items.Item("SumWeight").Specific.Value = 0;
                            oForm.Items.Item("HWeight").Specific.Value = 0;
                            oForm.Freeze(false);
                        }
                    }
                    else if (pVal.ItemUID == "Opt03") //분말
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            oForm.Freeze(true);
                            oForm.Items.Item("TranCard").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("CardCode").Enabled = true;
                            oForm.Items.Item("BPLId").Enabled = true;
                            oForm.Items.Item("DCardCod").Enabled = true;
                            oForm.Items.Item("TrType").Enabled = false;
                            oMat01.Columns.Item("SD030Num").Visible = true;
                            oMat01.Columns.Item("PackNo").Visible = true;
                            oMat01.Columns.Item("OrderNum").Visible = true;
                            oMat01.Columns.Item("Qty").Editable = false;
                            oMat01.Columns.Item("Weight").Editable = false;
                            oMat01.Columns.Item("SD030Num").Editable = false;
                            oMat01.Columns.Item("UnWeight").Visible = false;
                            oMat01.Columns.Item("Price").Editable = false;
                            oMat01.Columns.Item("WhsCode").Editable = true;
                            oMat01.Columns.Item("Comments").Editable = true;
                            oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oMat01.Clear();
                            oMat01.FlushToDataSource();
                            oMat01.LoadFromDataSource();
                            PS_SD040_AddMatrixRow(0, true);

                            oForm.Items.Item("SumSjQty").Specific.Value = 0;
                            oForm.Items.Item("SumSjWt").Specific.Value = 0;
                            oForm.Items.Item("SumQty").Specific.Value = 0;
                            oForm.Items.Item("SumWeight").Specific.Value = 0;
                            oForm.Items.Item("HWeight").Specific.Value = 0;
                            oForm.Freeze(false);
                        }
                    }

                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Opt01")
                    {
                        oForm.Items.Item("PriChange").Enabled = false;
                    }
                    else if (pVal.ItemUID == "Opt02") //멀티게이지
                    {
                        oForm.Items.Item("PriChange").Enabled = false;
                    }
                    else if (pVal.ItemUID == "Opt03") //분말
                    {
                        oForm.Items.Item("PriChange").Enabled = true;
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
        /// MATRIX_LINK_PRESSED 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LINK_PRESSED(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "SD030Num")
                        {
                            PS_SD030 tempForm = new PS_SD030();
                            tempForm.LoadForm(codeHelpClass.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 0, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().IndexOf("-")));
                        }
                        if (pVal.ColUID == "SD030H")
                        {
                            PS_SD030 tempForm = new PS_SD030();
                            tempForm.LoadForm(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                }
                else if (pVal.Before_Action == false)
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
        /// VALIDATE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string Query01;
            string Work01 = string.Empty;
            string ItemCode01;
            string sQry;
            double SumSjQty = 0;
            double SumQty = 0;
            double SumWeight = 0;
            double SumSjWt = 0;
            double HWeight = 0;
            string errMessage = string.Empty;

            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            SAPbobsCOM.Recordset RecordSet01 = null;
            SAPbobsCOM.Recordset RecordSet02 = null;
            SAPbobsCOM.Recordset RecordSet03 = null;

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "SD030Num")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    errMessage = " ";
                                    throw new Exception();
                                }

                                if (oForm.Items.Item("Opt03").Specific.Selected != true)
                                {
                                    for (i = 1; i <= oMat01.RowCount; i++)
                                    {
                                        if (pVal.Row != i) //현재 선택된 행 제외
                                        {
                                            if (oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value)
                                            {
                                                errMessage = "동일한 출하(선출)요청이 존재합니다.";
                                                oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value = "";
                                                throw new Exception();
                                            }

                                            if (codeHelpClass.Mid(oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value, 0, oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value.ToString().IndexOf("-")) != codeHelpClass.Mid(oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value, 0, oMat01.Columns.Item("SD030Num").Cells.Item(i).Specific.Value.ToString().IndexOf("-")))
                                            {
                                                if (oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() == "4") //구로영업소
                                                {
                                                    errMessage = "동일하지않은 출하요청문서가 존재합니다.";
                                                    oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value = "";
                                                    throw new Exception();
                                                }
                                            }
                                        }
                                    }
                                }
                                RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                Query01 = "EXEC PS_SD040_01 '" + oMat01.Columns.Item("SD030Num").Cells.Item(pVal.Row).Specific.Value + "'";
                                RecordSet01.DoQuery(Query01);
                                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                                {
                                    oDS_PS_SD040L.SetValue("U_SD030Num", pVal.Row - 1, RecordSet01.Fields.Item("SD030Num").Value);
                                    oDS_PS_SD040L.SetValue("U_OrderNum", pVal.Row - 1, RecordSet01.Fields.Item("OrderNum").Value);
                                    oDS_PS_SD040L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                    oDS_PS_SD040L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
                                    oDS_PS_SD040L.SetValue("U_ItemGpCd", pVal.Row - 1, RecordSet01.Fields.Item("ItemGpCd").Value);
                                    oDS_PS_SD040L.SetValue("U_ItmBsort", pVal.Row - 1, RecordSet01.Fields.Item("ItmBsort").Value);
                                    oDS_PS_SD040L.SetValue("U_ItmMsort", pVal.Row - 1, RecordSet01.Fields.Item("ItmMsort").Value);
                                    oDS_PS_SD040L.SetValue("U_Unit1", pVal.Row - 1, RecordSet01.Fields.Item("Unit1").Value);
                                    oDS_PS_SD040L.SetValue("U_Size", pVal.Row - 1, RecordSet01.Fields.Item("Size").Value);
                                    oDS_PS_SD040L.SetValue("U_ItemType", pVal.Row - 1, RecordSet01.Fields.Item("ItemType").Value);
                                    oDS_PS_SD040L.SetValue("U_Quality", pVal.Row - 1, RecordSet01.Fields.Item("Quality").Value);
                                    oDS_PS_SD040L.SetValue("U_Mark", pVal.Row - 1, RecordSet01.Fields.Item("Mark").Value);
                                    oDS_PS_SD040L.SetValue("U_SbasUnit", pVal.Row - 1, RecordSet01.Fields.Item("SbasUnit").Value);
                                    oDS_PS_SD040L.SetValue("U_LotNo", pVal.Row - 1, RecordSet01.Fields.Item("LotNo").Value);
                                    oDS_PS_SD040L.SetValue("U_SjQty", pVal.Row - 1, RecordSet01.Fields.Item("SjQty").Value);
                                    oDS_PS_SD040L.SetValue("U_SjWeight", pVal.Row - 1, RecordSet01.Fields.Item("SjWeight").Value);
                                    oDS_PS_SD040L.SetValue("U_Qty", pVal.Row - 1, RecordSet01.Fields.Item("Qty").Value);
                                    oDS_PS_SD040L.SetValue("U_UnWeight", pVal.Row - 1, RecordSet01.Fields.Item("UnWeight").Value);
                                    oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, RecordSet01.Fields.Item("Weight").Value);
                                    oDS_PS_SD040L.SetValue("U_Currency", pVal.Row - 1, RecordSet01.Fields.Item("Currency").Value);
                                    oDS_PS_SD040L.SetValue("U_Price", pVal.Row - 1, RecordSet01.Fields.Item("Price").Value);
                                    oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, RecordSet01.Fields.Item("LinTotal").Value);
                                    oDS_PS_SD040L.SetValue("U_WhsCode", pVal.Row - 1, RecordSet01.Fields.Item("WhsCode").Value);
                                    oDS_PS_SD040L.SetValue("U_WhsName", pVal.Row - 1, RecordSet01.Fields.Item("WhsName").Value);
                                    oDS_PS_SD040L.SetValue("U_Comments", pVal.Row - 1, RecordSet01.Fields.Item("Comments").Value);
                                    oDS_PS_SD040L.SetValue("U_SD030H", pVal.Row - 1, RecordSet01.Fields.Item("SD030H").Value);
                                    oDS_PS_SD040L.SetValue("U_SD030L", pVal.Row - 1, RecordSet01.Fields.Item("SD030L").Value);
                                    oDS_PS_SD040L.SetValue("U_TrType", pVal.Row - 1, RecordSet01.Fields.Item("TrType").Value);
                                    oDS_PS_SD040L.SetValue("U_ORDRNum", pVal.Row - 1, RecordSet01.Fields.Item("ORDRNum").Value);
                                    oDS_PS_SD040L.SetValue("U_RDR1Num", pVal.Row - 1, RecordSet01.Fields.Item("RDR1Num").Value);
                                    oDS_PS_SD040L.SetValue("U_Status", pVal.Row - 1, RecordSet01.Fields.Item("Status").Value);
                                    oDS_PS_SD040L.SetValue("U_LineId", pVal.Row - 1, RecordSet01.Fields.Item("LineId").Value);
                                    RecordSet01.MoveNext();
                                }
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_SD040L.GetValue("U_SD030Num", pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_SD040_AddMatrixRow(pVal.Row, false);
                                }

                                oMat01.LoadFromDataSource();
                                oMat01.AutoResizeColumns();

                                if (oMat01.VisualRowCount > 0)
                                {
                                    RecordSet03 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    sQry = "select U_Comments from [@PS_SD030H] where DocEntry = '" + codeHelpClass.Mid(oMat01.Columns.Item("SD030Num").Cells.Item(1).Specific.Value, 0, oMat01.Columns.Item("SD030Num").Cells.Item(1).Specific.Value.ToString().IndexOf("-")) + "'";
                                    RecordSet03.DoQuery(sQry);
                                    oForm.Items.Item("Comments").Specific.String = RecordSet03.Fields.Item("U_Comments").Value;
                                }

                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumSjQty += Convert.ToDouble(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value);
                                    }

                                    SumSjWt += Convert.ToDouble(oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value);

                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                    }

                                    SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                                    HWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value) / 1000;
                                }

                                oForm.Items.Item("SumSjQty").Specific.Value = Convert.ToString(SumSjQty);
                                oForm.Items.Item("SumSjWt").Specific.Value = Convert.ToString(SumSjWt);
                                oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(SumQty);
                                oForm.Items.Item("SumWeight").Specific.Value = Convert.ToString(SumWeight);
                                oForm.Items.Item("HWeight").Specific.Value = Convert.ToString(HWeight);

                                if (oMat01.RowCount > 1)
                                {
                                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oForm.Items.Item("CardCode").Enabled = false;
                                    oForm.Items.Item("BPLId").Enabled = false;
                                    oForm.Items.Item("DCardCod").Enabled = true;
                                    oForm.Items.Item("TrType").Enabled = false;
                                }
                                else
                                {
                                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oForm.Items.Item("CardCode").Enabled = true;
                                    oForm.Items.Item("BPLId").Enabled = true;
                                    oForm.Items.Item("DCardCod").Enabled = true;
                                    oForm.Items.Item("TrType").Enabled = true;
                                }
                                oForm.Update();
                            }
                            else if (pVal.ColUID == "PackNo")
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value))
                                {
                                    errMessage = " ";
                                    throw new Exception();
                                }

                                for (i = 1; i <= oMat01.RowCount; i++)
                                {
                                    if (pVal.Row != i) //현재 선택된 행 제외
                                    {
                                        if (oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("PackNo").Cells.Item(i).Specific.Value)
                                        {
                                            errMessage = "동일한 포장번호가 존재합니다.";
                                            oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value = "";
                                            throw new Exception();
                                        }
                                    }
                                }

                                RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                if (oForm.Items.Item("Opt02").Specific.Selected == true)
                                {
                                    Work01 = "1";
                                }
                                else if (oForm.Items.Item("Opt03").Specific.Selected == true)
                                {
                                    Work01 = "3";
                                }

                                Query01 = "EXEC PS_SD040_06 '" + oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value + "', '" + Work01 + "', '" + oForm.Items.Item("CardCode").Specific.String + "'";
                                RecordSet01.DoQuery(Query01);

                                if (RecordSet01.RecordCount <= 0)
                                {
                                    errMessage = "포장번호정보가 존재하지 않습니다.";
                                    oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value = "";
                                    throw new Exception();
                                }
                                else //해당 포장번호로 재고유무확인
                                {
                                    RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    Query01 = "EXEC PS_SD040_07 '" + oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value + "'";
                                    RecordSet02.DoQuery(Query01);

                                    if (RecordSet02.Fields.Item(0).Value == "Enabled")
                                    {
                                        //진행가능
                                    }
                                    else if (RecordSet02.Fields.Item(0).Value == "Disabled")
                                    {
                                        errMessage = "해당 포장번호의 재고가 부족합니다.";
                                        oMat01.Columns.Item("PackNo").Cells.Item(pVal.Row).Specific.Value = "";
                                        throw new Exception();
                                    }
                                }
                                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                                {
                                    oDS_PS_SD040L.SetValue("U_PackNo", pVal.Row - 1 + i, RecordSet01.Fields.Item("PackNo").Value);
                                    if (Work01 == "3")
                                    {
                                        oDS_PS_SD040L.SetValue("U_SD030Num", pVal.Row - 1 + i, RecordSet01.Fields.Item("SD030Num").Value);
                                        oDS_PS_SD040L.SetValue("U_OrderNum", pVal.Row - 1 + i, RecordSet01.Fields.Item("OrderNum").Value);
                                        oDS_PS_SD040L.SetValue("U_SD030H", pVal.Row - 1 + i, RecordSet01.Fields.Item("SD030H").Value);
                                        oDS_PS_SD040L.SetValue("U_SD030L", pVal.Row - 1 + i, RecordSet01.Fields.Item("SD030L").Value);
                                        oDS_PS_SD040L.SetValue("U_ORDRNum", pVal.Row - 1 + i, RecordSet01.Fields.Item("ORDRNum").Value);
                                        oDS_PS_SD040L.SetValue("U_RDR1Num", pVal.Row - 1 + i, RecordSet01.Fields.Item("RDR1Num").Value);
                                    }
                                    oDS_PS_SD040L.SetValue("U_ItemCode", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItemCode").Value);
                                    oDS_PS_SD040L.SetValue("U_ItemName", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItemName").Value);
                                    oDS_PS_SD040L.SetValue("U_ItemGpCd", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItemGpCd").Value);
                                    oDS_PS_SD040L.SetValue("U_ItmBsort", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItmBsort").Value);
                                    oDS_PS_SD040L.SetValue("U_ItmMsort", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItmMsort").Value);
                                    oDS_PS_SD040L.SetValue("U_Unit1", pVal.Row - 1 + i, RecordSet01.Fields.Item("Unit1").Value);
                                    oDS_PS_SD040L.SetValue("U_Size", pVal.Row - 1 + i, RecordSet01.Fields.Item("Size").Value);
                                    oDS_PS_SD040L.SetValue("U_ItemType", pVal.Row - 1 + i, RecordSet01.Fields.Item("ItemType").Value);
                                    oDS_PS_SD040L.SetValue("U_Quality", pVal.Row - 1 + i, RecordSet01.Fields.Item("Quality").Value);
                                    oDS_PS_SD040L.SetValue("U_Mark", pVal.Row - 1 + i, RecordSet01.Fields.Item("Mark").Value);
                                    oDS_PS_SD040L.SetValue("U_SbasUnit", pVal.Row - 1 + i, RecordSet01.Fields.Item("SbasUnit").Value);
                                    oDS_PS_SD040L.SetValue("U_LotNo", pVal.Row - 1 + i, RecordSet01.Fields.Item("LotNo").Value);
                                    oDS_PS_SD040L.SetValue("U_CoilNo", pVal.Row - 1 + i, RecordSet01.Fields.Item("CoilNo").Value);
                                    oDS_PS_SD040L.SetValue("U_PackWgt", pVal.Row - 1 + i, RecordSet01.Fields.Item("PackWgt").Value);
                                    oDS_PS_SD040L.SetValue("U_SjQty", pVal.Row - 1 + i, RecordSet01.Fields.Item("SjQty").Value);
                                    oDS_PS_SD040L.SetValue("U_SjWeight", pVal.Row - 1 + i, RecordSet01.Fields.Item("SjWeight").Value);
                                    oDS_PS_SD040L.SetValue("U_Qty", pVal.Row - 1 + i, RecordSet01.Fields.Item("Qty").Value);
                                    oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1 + i, RecordSet01.Fields.Item("Weight").Value);
                                    oDS_PS_SD040L.SetValue("U_Currency", pVal.Row - 1 + i, RecordSet01.Fields.Item("Currency").Value);
                                    oDS_PS_SD040L.SetValue("U_Price", pVal.Row - 1 + i, RecordSet01.Fields.Item("Price").Value);
                                    oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1 + i, RecordSet01.Fields.Item("LinTotal").Value);
                                    oDS_PS_SD040L.SetValue("U_WhsCode", pVal.Row - 1 + i, RecordSet01.Fields.Item("WhsCode").Value);
                                    oDS_PS_SD040L.SetValue("U_WhsName", pVal.Row - 1 + i, RecordSet01.Fields.Item("WhsName").Value);
                                    oDS_PS_SD040L.SetValue("U_Status", pVal.Row - 1 + i, RecordSet01.Fields.Item("Status").Value);
                                    oDS_PS_SD040L.SetValue("U_LineId", pVal.Row - 1 + i, RecordSet01.Fields.Item("LineId").Value);
                                    PS_SD040_AddMatrixRow(pVal.Row + i, false);
                                    RecordSet01.MoveNext();
                                }

                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumSjQty += Convert.ToDouble(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value);
                                    }

                                    SumSjWt += Convert.ToDouble(oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value);

                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                    }

                                    SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                                    HWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value) / 1000;
                                }

                                oForm.Items.Item("SumSjQty").Specific.Value = Convert.ToString(SumSjQty);
                                oForm.Items.Item("SumSjWt").Specific.Value = Convert.ToString(SumSjWt);
                                oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(SumQty);
                                oForm.Items.Item("SumWeight").Specific.Value = Convert.ToString(SumWeight);
                                oForm.Items.Item("HWeight").Specific.Value = Convert.ToString(HWeight);

                                if (oMat01.RowCount > 1)
                                {
                                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oForm.Items.Item("CardCode").Enabled = false;
                                    oForm.Items.Item("BPLId").Enabled = false;
                                    oForm.Items.Item("DCardCod").Enabled = false;
                                    oForm.Items.Item("TrType").Enabled = false;
                                }
                                else
                                {
                                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    oForm.Items.Item("CardCode").Enabled = true;
                                    oForm.Items.Item("BPLId").Enabled = true;
                                    oForm.Items.Item("DCardCod").Enabled = true;
                                    oForm.Items.Item("TrType").Enabled = true;
                                }

                                oForm.Update();
                            }
                            else if (pVal.ColUID == "Qty")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                    oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, "0");
                                    oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, "0");
                                }
                                else
                                {
                                    ItemCode01 = oDS_PS_SD040L.GetValue("U_ItemCode", pVal.Row - 1);

                                    if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "101") //EA자체품
                                    {
                                        oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "102") //EAUOM
                                    {
                                        oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(ItemCode01))));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "201") //KGSPEC
                                    {
                                        oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(dataHelpClass.GetItem_Spec1(ItemCode01)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(ItemCode01)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "202") //KG단중
                                    {
                                        oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "203") //KG선택
                                    {
                                    }
                                    oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oDS_PS_SD040L.GetValue("U_Weight", pVal.Row - 1).ToString().Trim()) * Convert.ToDouble(oDS_PS_SD040L.GetValue("U_Price", pVal.Row - 1).ToString().Trim())));
                                    oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }

                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumSjQty += Convert.ToDouble(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value);
                                    }

                                    SumSjWt += Convert.ToDouble(oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value);

                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                    }
                                    
                                    SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                                    HWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value) / 1000;
                                }
                                oForm.Items.Item("SumSjQty").Specific.Value = Convert.ToString(SumSjQty);
                                oForm.Items.Item("SumSjWt").Specific.Value = Convert.ToString(SumSjWt);
                                oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(SumQty);
                                oForm.Items.Item("SumWeight").Specific.Value = Convert.ToString(SumWeight);
                                oForm.Items.Item("HWeight").Specific.Value = Convert.ToString(HWeight);
                            }
                            else if (pVal.ColUID == "Weight")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oDS_PS_SD040L.SetValue("U_Qty", pVal.Row - 1, "0");
                                    oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                    oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, "0");
                                }
                                else
                                {
                                    ItemCode01 = oDS_PS_SD040L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
                                    if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "101")
                                    {
                                        oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "102") //EAUOM
                                    {
                                        oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "201") //KGSPEC
                                    {
                                        oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(dataHelpClass.GetItem_Spec1(ItemCode01)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(ItemCode01)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "202") //KG단중
                                    {
                                        oDS_PS_SD040L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
                                    }
                                    else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "203") //KG선택
                                    {
                                        oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oDS_PS_SD040L.GetValue("U_Weight", pVal.Row - 1).ToString().Trim()) * Convert.ToDouble(oDS_PS_SD040L.GetValue("U_Price", pVal.Row - 1).ToString().Trim())));
                                }

                                for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                                {
                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumSjQty += Convert.ToDouble(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value);
                                    }

                                    SumSjWt += Convert.ToDouble(oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value);

                                    if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                                    {
                                        SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                                    }
                                    
                                    SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                                    HWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value) / 1000;
                                }

                                oForm.Items.Item("SumSjQty").Specific.Value = Convert.ToString(SumSjQty);
                                oForm.Items.Item("SumSjWt").Specific.Value = Convert.ToString(SumSjWt);
                                oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(SumQty);
                                oForm.Items.Item("SumWeight").Specific.Value = Convert.ToString(SumWeight);
                                oForm.Items.Item("HWeight").Specific.Value = Convert.ToString(HWeight);
                            }
                            else if (pVal.ColUID == "Price")
                            {
                                if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                {
                                    oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                    oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, "0");
                                }
                                else
                                {
                                    oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_SD040L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(pVal.Row).Specific.Value)));
                                }
                            }
                            else
                            {
                                oForm.Items.Item("TranCard").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oDS_PS_SD040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                            oForm.Update();
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_SD040H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_SD040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_SD040H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                        }
                    }
                }
                else if (pVal.Before_Action == false)
                {
                }
            }
            catch (Exception ex)
            {
                if (errMessage == " ")
                {
                    //메시지 출력 없음
                }
                else if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }

                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);

                if (RecordSet01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                }

                if (RecordSet02 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
                }

                if (RecordSet03 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet03);
                }
            }
        }

        /// <summary>
        /// MATRIX_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_MATRIX_LOAD(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            double SumSjQty = 0;
            double SumQty = 0;
            double SumWeight = 0;
            double SumSjWt = 0;
            double HWeight = 0;

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_SD040_FormItemEnabled();
                    PS_SD040_AddMatrixRow(oMat01.VisualRowCount, false);
                    
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        if (!string.IsNullOrEmpty(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value))
                        {
                            SumSjQty += Convert.ToDouble(oMat01.Columns.Item("SjQty").Cells.Item(i + 1).Specific.Value);
                        }

                        SumSjWt += Convert.ToDouble(oMat01.Columns.Item("SjWeight").Cells.Item(i + 1).Specific.Value);

                        if (!string.IsNullOrEmpty(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value))
                        {
                            SumQty += Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(i + 1).Specific.Value);
                        }

                        SumWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value);
                        HWeight += Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i + 1).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item("UnWeight").Cells.Item(i + 1).Specific.Value) / 1000;
                    }
                    oForm.Items.Item("SumSjQty").Specific.Value = Convert.ToString(SumSjQty);
                    oForm.Items.Item("SumSjWt").Specific.Value = Convert.ToString(SumSjWt);
                    oForm.Items.Item("SumQty").Specific.Value = Convert.ToString(SumQty);
                    oForm.Items.Item("SumWeight").Specific.Value = Convert.ToString(SumWeight);
                    oForm.Items.Item("HWeight").Specific.Value = Convert.ToString(HWeight);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD040H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD040L);
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
        /// CHOOSE_FROM_LIST 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CHOOSE_FROM_LIST(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "CardCode" || pVal.ItemUID == "CardName")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_SD040H", "U_CardCode,U_CardName", "", 0, "", "", "");
                    }
                    else if (pVal.ItemUID == "DCardCod" || pVal.ItemUID == "DCardNam")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_SD040H", "U_DCardCod,U_DCardNam", "", 0, "", "", "");
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "WhsCode")
                        {
                            if (((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects != null)
                            {
                                SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당
                                oDS_PS_SD040L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
                                oDS_PS_SD040L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
                                oMat01.LoadFromDataSource();
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
            }
        }

        /// <summary>
        /// 행삭제 체크 메소드(Raise_FormMenuEvent 에서 사용)
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
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("납품된행은 삭제할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                            BubbleEvent = false;
                            return;
                        }

                        if (oForm.Items.Item("Opt02").Specific.Selected == true || oForm.Items.Item("Opt03").Specific.Selected == true) //멀티게이지, 분말
                        {
                            for (i = 0; i <= oDS_PS_SD040L.Size - 1; i++)
                            {
                                if (i == oDS_PS_SD040L.Size - 1)
                                {
                                    break;
                                }

                                if (oDS_PS_SD040L.GetValue("U_PackNo", i) == oDS_PS_SD040L.GetValue("U_PackNo", oLastColRow01 - 1)) //선택된행과 같은값을 가지는 모든행
                                {
                                    oDS_PS_SD040L.RemoveRecord(i);
                                    i -= 1;
                                }
                            }

                            for (i = 0; i <= oDS_PS_SD040L.Size - 1; i++)
                            {
                                oDS_PS_SD040L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                            }
                            oMat01.LoadFromDataSource();

                            if (oMat01.RowCount == 0)
                            {
                                PS_SD040_AddMatrixRow(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_SD040L.GetValue("U_SD030Num", oMat01.RowCount - 1).ToString().Trim()))
                                {
                                    PS_SD040_AddMatrixRow(oMat01.RowCount, false);
                                }
                            }

                            if (oMat01.RowCount > 1)
                            {
                                oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Items.Item("CardCode").Enabled = false;
                                oForm.Items.Item("BPLId").Enabled = false;
                                oForm.Items.Item("TrType").Enabled = false;
                            }
                            else
                            {
                                oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Items.Item("CardCode").Enabled = true;
                                oForm.Items.Item("BPLId").Enabled = true;
                                oForm.Items.Item("DCardCod").Enabled = true;
                                oForm.Items.Item("TrType").Enabled = true;
                            }
                            oForm.Update();
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_SD040L.RemoveRecord(oDS_PS_SD040L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_SD040_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_SD040L.GetValue("U_SD030Num", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_SD040_AddMatrixRow(oMat01.RowCount, false);
                            }
                        }

                        if (oMat01.RowCount > 1)
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("CardCode").Enabled = false;
                            oForm.Items.Item("BPLId").Enabled = false;
                            oForm.Items.Item("TrType").Enabled = false;
                        }
                        else
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("CardCode").Enabled = true;
                            oForm.Items.Item("BPLId").Enabled = true;
                            oForm.Items.Item("DCardCod").Enabled = true;
                            oForm.Items.Item("TrType").Enabled = true;
                        }
                        oForm.Update();
                    }
                }
            }
            catch(Exception ex)
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
            int i;
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284": //취소
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_SD040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }

                                if (oForm.Items.Item("Opt02").Specific.Selected == true) //멀티
                                {
                                    for (i = 1; i <= oMat01.VisualRowCount - 1; i++) //멀티게이지 재고수량검사
                                    {
                                        sQry = "Select Quantity = Sum(a.Quantity) From OIBT a Inner Join OITM b On a.ItemCode = b.ItemCode ";
                                        sQry = sQry + " Where b.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + "'";
                                        sQry = sQry + " And a.BatchNum = '" + oMat01.Columns.Item("LotNo").Cells.Item(i).Specific.Value + "'";
                                        sQry = sQry + " AND a.WhsCode = '" + oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value + "'";
                                        if (Convert.ToDouble(dataHelpClass.GetValue(sQry, 0, 1)) > 0)
                                        {
                                            PSH_Globals.SBO_Application.StatusBar.SetText("멀티게이지 품목 : " + oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + " 의 재고가 존재합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                }
                                //분말은 배치품목 수량을 분할로 납품반품되야되기때문에 재고가 있더라도 반품이 가능해야함(황영수 20190924)
                                //                    If oForm.Items("Opt03").Specific.Selected = True Then '//분말이면서
                                //                        For i = 1 To oMat01.VisualRowCount - 1 '//분말 재고수량검사
                                //                            If MDC_PS_Common.GetValue("SELECT Quantity FROM [OIBT] WHERE ItemCode = '" & oMat01.Columns("ItemCode").Cells(i).Specific.Value & "' AND BatchNum = '" & oMat01.Columns("LotNo").Cells(i).Specific.Value & "' AND WhsCode = '" & oMat01.Columns("WhsCode").Cells(i).Specific.Value & "'", 0, 1) > 0 Then
                                //                                Call MDC_Com.MDC_GF_Message("분말 품목 : " & oMat01.Columns("ItemCode").Cells(i).Specific.Value & " 의 재고가 존재합니다.", "W")
                                //                                BubbleEvent = False
                                //                                Exit Sub
                                //                            End If
                                //                        Next
                                //                    End If
                                
                                for (i = 1; i <= oMat01.VisualRowCount - 1; i++) //AR송장처리된 문서존재유무검사
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = '" + oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value + "' AND BaseLine = '" + oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        PSH_Globals.SBO_Application.StatusBar.SetText("AR송장처리된 문서가 존재합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                                
                                for (i = 1; i <= oMat01.VisualRowCount - 1; i++) //반품처리된 문서존재유무검사
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = '" + oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value + "' AND BaseLine = '" + oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        PSH_Globals.SBO_Application.StatusBar.SetText("반품처리된 문서가 존재합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                        BubbleEvent = false;
                                        return;
                                    }
                                }

                                if (PSH_Globals.SBO_Application.MessageBox("납품처리를 취소하시겠습니까?", 1, "Yes", "No") == 1)
                                {
                                    if (PS_SD040_DI_API_03() == false)
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                                else
                                {
                                    BubbleEvent = false;
                                }
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("현재 모드에서는 취소할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
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
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_SD040_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_SD040_FormItemEnabled();
                            PS_SD040_AddMatrixRow(0, true);
                            oDS_PS_SD040H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD()); //담당자
                            oDS_PS_SD040H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value + "'", 0, 1));
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_SD040_FormItemEnabled();
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
                        //Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE: //35
                        //Raise_EVENT_FORM_DATA_UPDATE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE: //36
                        //Raise_EVENT_FORM_DATA_DELETE(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
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
        /// RightClickEvent
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
                switch (pVal.ItemUID)
                {
                    case "Mat01":
                        if (pVal.Row > 0)
                        {
                            oLastItemUID01 = pVal.ItemUID;
                            oLastColUID01 = pVal.ColUID;
                            oLastColRow01 = pVal.Row;
                        }
                        break;
                    default:
                        oLastItemUID01 = pVal.ItemUID;
                        oLastColUID01 = "";
                        oLastColRow01 = 0;
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
    }
}
