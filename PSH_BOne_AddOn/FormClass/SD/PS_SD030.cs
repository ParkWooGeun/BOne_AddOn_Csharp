using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Form;
using PSH_BOne_AddOn.DataPack;
using System.Collections.Generic;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 출하(선출)요청등록
	/// </summary>
	internal class PS_SD030 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.DBDataSource oDS_PS_SD030H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_SD030L; //등록라인
		private string oDocType01;
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값

        #region 선출요청[PS_SD031]에서 사용
        //public class ItemInformation
        //{
        //	public string ItemCode; //수량
        //	public int Qty; //중량
        //	public double Weight; //통화
        //	public string Currency; //단가
        //	public double Price; //총계
        //	public double LineTotal; //창고
        //	public string WhsCode; //판매오더문서
        //	public int ORDRNum; //판매오더라인
        //	public int RDR1Num;
        //	public bool Check; //납품문서
        //	public int ODLNNum; //납품라인
        //	public int DLN1Num; //반품문서
        //	public int ORDNNum; //반품라인
        //	public int RDN1Num; //출하(선출)문서
        //	public int SD030HNum; //출하(선출)라인
        //	public int SD030LNum;
        //}
        #endregion

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_SD030.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_SD030_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_SD030");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
				oDocType01 = "출하요청";
                PS_SD030_CreateItems();
                PS_SD030_SetComboBox();
                PS_SD030_CF_ChooseFromList();
                PS_SD030_EnableMenus();
                PS_SD030_SetDocument(oFormDocEntry);
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
        private void PS_SD030_CreateItems()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oDS_PS_SD030H = oForm.DataSources.DBDataSources.Item("@PS_SD030H");
                oDS_PS_SD030L = oForm.DataSources.DBDataSources.Item("@PS_SD030L");
                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                if (oDocType01 == "출하요청")
                {
                    oForm.Title = "출하요청[PS_SD030]";
                    oForm.Items.Item("DocType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if ((oDocType01 == "선출요청"))
                {
                    oForm.Title = "선출요청[PS_SD031]";
                    oForm.Items.Item("DocType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                
                oDS_PS_SD030H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD()); //담당자
                oDS_PS_SD030H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value + "'", 0, 1));
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 콤보박스 세팅
        /// </summary>
        private void PS_SD030_SetComboBox()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "DocType", "", "1", "출하요청");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "DocType", "", "2", "선출요청");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PS_SD030", "DocType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "TrType", "", "1", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "TrType", "", "2", "임가공");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("TrType").Specific, "PS_PS_SD030", "TrType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "1", "출하요청");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "2", "선출요청");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "3", "납품");
                dataHelpClass.Combo_ValidValues_Insert("PS_PS_SD030", "ProgStat", "", "4", "반품");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("ProgStat").Specific, "PS_PS_SD030", "ProgStat", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_SD030", "Mat01", "Status", "O", "미결");
                dataHelpClass.Combo_ValidValues_Insert("PS_SD030", "Mat01", "Status", "C", "완료");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("Status"), "PS_SD030", "Mat01", "Status", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_SD030", "Mat01", "TrType", "1", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_SD030", "Mat01", "TrType", "2", "군납");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("TrType"), "PS_SD030", "Mat01", "TrType", false);

                oForm.Items.Item("Managed").Specific.ValidValues.Add("N", "미대상");
                oForm.Items.Item("Managed").Specific.ValidValues.Add("Y", "대상");
                oForm.Items.Item("Managed").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                oForm.Items.Item("CheckYN").Specific.ValidValues.Add("Y", "승인대기");
                oForm.Items.Item("CheckYN").Specific.ValidValues.Add("N", "승인완료");
                oForm.Items.Item("CheckYN").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_SD030_CF_ChooseFromList()
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
            catch (Exception ex)
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
        private void PS_SD030_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, true, true, true, true, true, true, false, false, false, false, true, false);
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
        private void PS_SD030_SetDocument(string oFormDocEntry)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_SD030_FormItemEnabled();
                    PS_SD030_AddMatrixRow(0, true);
                }
                else
                {
                    if (dataHelpClass.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" + oFormDocEntry + "'", 0, 1) == "1")
                    {
                        oForm.Title = "출하요청[PS_SD030]";
                        oForm.Items.Item("DocType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    else if ((dataHelpClass.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" + oFormDocEntry + "'", 0, 1) == "2"))
                    {
                        oForm.Title = "선출요청[PS_SD031]";
                        oForm.Items.Item("DocType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_SD030_FormItemEnabled();
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
        private void PS_SD030_FormItemEnabled()
        {
            int i;
            bool Enabled;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("TranCard").Enabled = true;
                    oForm.Items.Item("TranCode").Enabled = true;
                    oForm.Items.Item("Destin").Enabled = true;
                    oForm.Items.Item("TranCost").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = true;
                    oForm.Items.Item("TrType").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oMat01.AutoResizeColumns();
                    PS_SD030_SetDocEntry();
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    
                    if (oDocType01 == "출하요청")
                    {
                        oForm.Items.Item("DocType").Enabled = true;
                        oForm.Items.Item("DocType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("DocType").Enabled = false;
                        oForm.Items.Item("Button01").Visible = false;
                        oForm.Items.Item("Button04").Visible = false;
                        oForm.Items.Item("ProgStat").Enabled = true;
                        oForm.Items.Item("ProgStat").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("ProgStat").Enabled = false;
                        oMat01.Columns.Item("ODLNNum").Visible = false;
                        oMat01.Columns.Item("DLN1Num").Visible = false;
                        oForm.Update();
                    }
                    else if (oDocType01 == "선출요청")
                    {
                        oForm.Items.Item("DocType").Enabled = true;
                        oForm.Items.Item("DocType").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("DocType").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = true;
                        oForm.Items.Item("BPLId").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("TrType").Enabled = false;
                        oForm.Items.Item("Button01").Visible = true;
                        oForm.Items.Item("Button04").Visible = true;
                        oForm.Items.Item("ProgStat").Enabled = true;
                        oForm.Items.Item("ProgStat").Specific.Select("2", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        oForm.Items.Item("CardCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("ProgStat").Enabled = false;
                        oMat01.Columns.Item("ODLNNum").Visible = true;
                        oMat01.Columns.Item("DLN1Num").Visible = true;
                        oForm.Items.Item("Button03").Visible = false;
                    }
                    oForm.Items.Item("TrType").Specific.Select("1", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oDS_PS_SD030H.SetValue("U_CntcCode", 0, dataHelpClass.User_MSTCOD());
                    oDS_PS_SD030H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("CntcCode").Specific.Value + "'", 0, 1));
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("Button01").Enabled = false;
                    oForm.Items.Item("Button04").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("DocDate").Specific.String = DateTime.Now.ToString("yyyyMMdd");
                    oForm.Items.Item("DueDate").Specific.String = DateTime.Now.ToString("yyyyMMdd");

                    oForm.Items.Item("14").Visible = false;
                    oForm.Items.Item("16").Visible = false;
                    oForm.Items.Item("18").Visible = false;
                    oForm.Items.Item("20").Visible = false;
                    oForm.Items.Item("TranCard").Visible = false;
                    oForm.Items.Item("TranCode").Visible = false;
                    oForm.Items.Item("Destin").Visible = false;
                    oForm.Items.Item("TranCost").Visible = false;
                }
                else if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE))
                {
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("CardCode").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("TranCard").Enabled = true;
                    oForm.Items.Item("TranCode").Enabled = true;
                    oForm.Items.Item("Destin").Enabled = true;
                    oForm.Items.Item("TranCost").Enabled = true;
                    oForm.Items.Item("Comments").Enabled = true;
                    oForm.Items.Item("TrType").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oMat01.AutoResizeColumns();
                    oForm.EnableMenu("1281", false);
                    oForm.EnableMenu("1282", true);
                    oForm.Items.Item("Button01").Enabled = false;
                    oForm.Items.Item("Button04").Enabled = false;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;

                    oForm.Items.Item("14").Visible = false;
                    oForm.Items.Item("16").Visible = false;
                    oForm.Items.Item("18").Visible = false;
                    oForm.Items.Item("20").Visible = false;
                    oForm.Items.Item("TranCard").Visible = false;
                    oForm.Items.Item("TranCode").Visible = false;
                    oForm.Items.Item("Destin").Visible = false;
                    oForm.Items.Item("TranCost").Visible = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    
                    if (dataHelpClass.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" + oDS_PS_SD030H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "1") //출하요청일때
                    {
                        Enabled = false;
                        for (i = 0; i <= oDS_PS_SD030L.Size - 1; i++)
                        {
                            //매트릭스의 중량과 납품문서의 중량을 비교
                            if (Convert.ToDouble(oDS_PS_SD030L.GetValue("U_Weight", i).ToString().Trim()) > Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(U_Weight) FROM [@PS_SD040L] WHERE U_SD030H = '" + oDS_PS_SD030H.GetValue("DocEntry", 0).ToString().Trim() + "' AND U_SD030L = '" + oDS_PS_SD030L.GetValue("U_LineId", i).ToString().Trim() + "'", 0, 1)))
                            {
                                Enabled = true;
                            }
                        }
                        
                        if (Enabled == false) //문서 수정 불가능한 경우
                        {
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("CardCode").Enabled = false;
                            oForm.Items.Item("BPLId").Enabled = false;
                            oForm.Items.Item("CntcCode").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("DueDate").Enabled = false;
                            oForm.Items.Item("TranCard").Enabled = false;
                            oForm.Items.Item("TranCode").Enabled = false;
                            oForm.Items.Item("Destin").Enabled = false;
                            oForm.Items.Item("TranCost").Enabled = false;
                            oForm.Items.Item("Comments").Enabled = true;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("TrType").Enabled = false;
                            oMat01.AutoResizeColumns();
                            oForm.EnableMenu("1281", true);
                            oForm.EnableMenu("1282", false);
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("Button04").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("DueDate").Enabled = false;
                        }
                        else //문서 수정 가능한 경우
                        {
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("CardCode").Enabled = false;
                            oForm.Items.Item("BPLId").Enabled = false;
                            oForm.Items.Item("CntcCode").Enabled = true;
                            oForm.Items.Item("DocDate").Enabled = true;
                            oForm.Items.Item("DueDate").Enabled = true;
                            oForm.Items.Item("TranCard").Enabled = true;
                            oForm.Items.Item("TranCode").Enabled = true;
                            oForm.Items.Item("Destin").Enabled = true;
                            oForm.Items.Item("TranCost").Enabled = true;
                            oForm.Items.Item("Comments").Enabled = true;
                            oForm.Items.Item("TrType").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = true;
                            oMat01.AutoResizeColumns();
                            oForm.EnableMenu("1281", true);
                            oForm.EnableMenu("1282", false);
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("Button04").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("DueDate").Enabled = false;
                        }
                        oForm.Items.Item("14").Visible = false;
                        oForm.Items.Item("16").Visible = false;
                        oForm.Items.Item("18").Visible = false;
                        oForm.Items.Item("20").Visible = false;
                        oForm.Items.Item("TranCard").Visible = false;
                        oForm.Items.Item("TranCode").Visible = false;
                        oForm.Items.Item("Destin").Visible = false;
                        oForm.Items.Item("TranCost").Visible = false;
                    } 
                    else if (dataHelpClass.GetValue("SELECT U_DocType FROM [@PS_SD030H] WHERE DocEntry = '" + oDS_PS_SD030H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "2") //선출요청일때
                    {
                        if (dataHelpClass.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" + oDS_PS_SD030H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "3") //납품일때
                        {
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("CardCode").Enabled = false;
                            oForm.Items.Item("BPLId").Enabled = false;
                            oForm.Items.Item("CntcCode").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("DueDate").Enabled = false;
                            oForm.Items.Item("TranCard").Enabled = false;
                            oForm.Items.Item("TranCode").Enabled = false;
                            oForm.Items.Item("Destin").Enabled = false;
                            oForm.Items.Item("TranCost").Enabled = false;
                            oForm.Items.Item("Comments").Enabled = true;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("TrType").Enabled = false;
                            oMat01.AutoResizeColumns();
                            oForm.EnableMenu("1281", true);
                            oForm.EnableMenu("1282", false);
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("Button04").Enabled = true;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("DueDate").Enabled = false;

                            oForm.Items.Item("14").Visible = true;
                            oForm.Items.Item("16").Visible = true;
                            oForm.Items.Item("18").Visible = true;
                            oForm.Items.Item("20").Visible = true;
                            oForm.Items.Item("TranCard").Visible = true;
                            oForm.Items.Item("TranCode").Visible = true;
                            oForm.Items.Item("Destin").Visible = true;
                            oForm.Items.Item("TranCost").Visible = true;
                        }
                        else //납품이 아닐때
                        {
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("CardCode").Enabled = false;
                            oForm.Items.Item("BPLId").Enabled = false;
                            oForm.Items.Item("CntcCode").Enabled = true;
                            oForm.Items.Item("DocDate").Enabled = true;
                            oForm.Items.Item("DueDate").Enabled = true;
                            oForm.Items.Item("TranCard").Enabled = true;
                            oForm.Items.Item("TranCode").Enabled = true;
                            oForm.Items.Item("Destin").Enabled = true;
                            oForm.Items.Item("TranCost").Enabled = true;
                            oForm.Items.Item("Comments").Enabled = true;
                            oForm.Items.Item("TrType").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = true;
                            oMat01.AutoResizeColumns();
                            oForm.EnableMenu("1281", true);
                            oForm.EnableMenu("1282", false);
                            oForm.Items.Item("Button01").Enabled = true;
                            oForm.Items.Item("Button04").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("DueDate").Enabled = false;

                            oForm.Items.Item("14").Visible = false;
                            oForm.Items.Item("16").Visible = false;
                            oForm.Items.Item("18").Visible = false;
                            oForm.Items.Item("20").Visible = false;
                            oForm.Items.Item("TranCard").Visible = false;
                            oForm.Items.Item("TranCode").Visible = false;
                            oForm.Items.Item("Destin").Visible = false;
                            oForm.Items.Item("TranCost").Visible = false;
                        }
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
        private void PS_SD030_SetDocEntry()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_SD030'", "");
                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
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
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 행추가
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param>
        private void PS_SD030_AddMatrixRow(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_SD030L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_SD030L.Offset = oRow;
                oDS_PS_SD030L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
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
        private bool PS_SD030_CheckDataValid()
        {
            bool returnValue = false;
            int i;
            int j;
            string sQry = string.Empty;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
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

                sQry = "select frozenFor  from OCRD where CardCode = '" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim() + "'";
                oRecordSet.DoQuery(sQry);
                if (oRecordSet.Fields.Item(0).Value.ToString().Trim() == "Y")
                {
                    errMessage = "거래처코드 비활성화 상태 입니다. 확인하세요. 코드 :" + oForm.Items.Item("CardCode").Specific.Value.ToString().Trim();
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "수주는 필수입니다.";
                        oMat01.Columns.Item("OrderNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    
                    if (Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value) <= 0)
                    {
                        errMessage = "중량(수량)은 필수입니다.";
                        oMat01.Columns.Item("Weight").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    
                    for (j = i + 1; j <= oMat01.VisualRowCount - 1; j++)
                    {
                        if (oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrderNum").Cells.Item(j).Specific.Value)
                        {
                            errMessage = "동일한 수주가 존재합니다.";
                            oMat01.Columns.Item("OrderNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                }

                if (PS_SD030_Validate("검사") == false)
                {
                    errMessage = " ";
                    throw new Exception();
                }

                oDS_PS_SD030L.RemoveRecord(oDS_PS_SD030L.Size - 1);
                oMat01.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_SD030_SetDocEntry();
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_SD030_Validate(string ValidateType)
        {
            bool returnValue = false;
            int i;
            int j;
            string query;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_SD030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할 수 없습니다.";
                    throw new Exception();
                }

                bool Exist = false;
                if (ValidateType == "검사")
                {
                    for (i = 1; i <= oMat01.VisualRowCount - 1; i++) //입력된 행에 대해
                    {
                        if (Convert.ToDouble(dataHelpClass.GetValue("SELECT COUNT(*) FROM [ORDR] ORDR LEFT JOIN [RDR1] RDR1 ON ORDR.DocEntry = RDR1.DocEntry WHERE CONVERT(NVARCHAR,ORDR.DocEntry) + '-' + CONVERT(NVARCHAR,RDR1.LineNum) = '" + oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                        {
                            errMessage = "판매오더문서가 존재하지 않습니다.";
                            throw new Exception();
                        }
                    }
                    //삭제된 행을 찾아서 삭제가능성 검사
                    Exist = false;
                    query = "SELECT DocEntry,LineId FROM [@PS_SD030L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    RecordSet01.DoQuery(query);
                    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        Exist = false;
                        for (j = 1; j <= oMat01.RowCount - 1; j++)
                        {
                            //라인번호가 같고, 품목코드가 같으면 존재하는 행, LineNum에 값이 존재하는지 확인 필요(행삭제된행인경우 LineNum이 존재하지않음)
                            string lineID = oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value;
                            if (Convert.ToInt32(RecordSet01.Fields.Item(1).Value) == Convert.ToInt32(lineID == "" ? "0" : lineID) && !string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                            {
                                Exist = true;
                            }
                        }
                        
                        if (Exist == false) //삭제된 행 중에서
                        {
                            if (oForm.Items.Item("DocType").Specific.Value == "1") //출하요청
                            {
                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + RecordSet01.Fields.Item(0).Value + "' AND PS_SD040L.U_SD030L = '" + RecordSet01.Fields.Item(1).Value + "'", 0, 1)) > 0)
                                {
                                    errMessage = "삭제된 행이 다른 사용자에 의해 납품되었습니다. 적용할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                            else if (oForm.Items.Item("DocType").Specific.Value == "2") //선출요청
                            {
                                if (dataHelpClass.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "3")
                                {
                                    errMessage = "이미 납품된 문서입니다. 삭제할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                        }
                        RecordSet01.MoveNext();
                    }
                    //수량가능성검사
                    for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value)) //새로추가된 행인경우, 검사 불필요
                        {
                        }
                        else
                        {
                            //매트릭스에 입력된 수량과 DB상에 존재하는 수량의 값비교
                            if (Convert.ToDouble(oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value) < Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(U_Weight) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_SD040L.U_SD030L = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)))
                            {
                                errMessage = "출하요청,선출요청 수량보다 작습니다.";
                                throw new Exception();
                            }
                            //납품된 행이 있으면 수정 불가
                            if (Convert.ToDouble(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_SD040L.U_SD030L = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0)
                            {
                                query = "  SELECT   U_OrderNum, ";
                                query += "          U_ItemCode, ";
                                query += "          U_WhsCode ";
                                query += " FROM     [@PS_SD030L] PS_SD030L";
                                query += " WHERE    PS_SD030L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                                query += "          AND PS_SD030L.LineId = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'";
                                RecordSet01.DoQuery(query);

                                if (RecordSet01.Fields.Item(0).Value == oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(1).Value == oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(2).Value == oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value)
                                {
                                }
                                else
                                {
                                    errMessage = "이미 납품된 행입니다. 수정할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                        }
                    }
                }
                else if (ValidateType == "수정")
                {
                }
                else if (ValidateType == "행삭제")
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oLastColRow01).Specific.Value))
                    {
                        //새로추가된 행인경우, 삭제 가능
                    }
                    else
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) //추가,수정모드일때행삭제가능검사
                        {

                            if (oForm.Items.Item("DocType").Specific.Value == "1") //출하요청
                            {
                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_SD040L.U_SD030L = '" + oMat01.Columns.Item("LineId").Cells.Item(oLastColRow01).Specific.Value + "'", 0, 1)) > 0)
                                {
                                    errMessage = "납품된 행입니다. 삭제할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                            else if (oForm.Items.Item("DocType").Specific.Value == "2") //선출요청
                            {
                                if (dataHelpClass.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "3")
                                {
                                    errMessage = "납품된 행입니다. 삭제할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                        }
                    }
                }
                else if (ValidateType == "취소")
                {
                    query = "SELECT DocEntry,LineId,U_ItemCode FROM [@PS_SD030L] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    RecordSet01.DoQuery(query);
                    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        //출하요청
                        if (oForm.Items.Item("DocType").Specific.Value == "1")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_SD040H] PS_SD040H LEFT JOIN [@PS_SD040L] PS_SD040L ON PS_SD040H.DocEntry = PS_SD040L.DocEntry WHERE PS_SD040H.Canceled = 'N' AND PS_SD040L.U_SD030H = '" + RecordSet01.Fields.Item(0).Value + "' AND PS_SD040L.U_SD030L = '" + RecordSet01.Fields.Item(1).Value + "'", 0, 1)) > 0)
                            {
                                errMessage = "납품된 문서입니다. 삭제할 수 없습니다.";
                                throw new Exception();
                            }
                        }
                        else if (oForm.Items.Item("DocType").Specific.Value == "2") //선출요청
                        {
                            if (dataHelpClass.GetValue("SELECT U_ProgStat FROM [@PS_SD030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "3")
                            {
                                errMessage = "납품된 문서입니다. 삭제할 수 없습니다.";
                                throw new Exception();
                            }
                        }
                        RecordSet01.MoveNext();
                    }
                }

                returnValue = true;
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 거래명세서 출력
        /// </summary>
        [STAThread]
        private void PS_SD030_Print_Report01()
        {
            string WinTitle;
            string ReportName;
            int i;
            string errMessage = string.Empty;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();

            try
            {
                if (oForm.Items.Item("ProgStat").Specific.Selected.Value != "3" && oDocType01 != "선출요청")
                {
                    errMessage = "문서상태가 납품이 아닙니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("DocType").Specific.Selected.Value != "2")
                {
                    errMessage = "선출요청이 아닙니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("BPLId").Specific.Selected.Value != "2")
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

                WinTitle = "[PS_PP540_10] 레포트";
                ReportName = "PS_PP540_10.rpt";
                //쿼리 : PS_PP540_10

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass>
                {
                    new PSH_DataPackClass("@DocType", "선출"),
                    new PSH_DataPackClass("@DocEntry", oForm.Items.Item("DocEntry").Specific.Value),
                    new PSH_DataPackClass("@UserSign", PSH_Globals.oCompany.UserSignature)
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
        /// 출하요청서
        /// </summary>
        [STAThread]
        private void PS_SD030_Print_Report02()
        {
            string WinTitle;
            string ReportName;
            string sQry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                WinTitle = "[PS_SD030_10] 레포트";
                ReportName = "PS_SD030_10.rpt";
                //쿼리 : PS_SD030_10

                sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
                oRecordSet01.DoQuery(sQry);

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass> //Parameter List
                {
                    new PSH_DataPackClass("@SD030HNo", oForm.Items.Item("DocEntry").Specific.Value)
                };

                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass> //Formula List
                {
                    new PSH_DataPackClass("@BPLName", oRecordSet01.Fields.Item(0).Value)
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }
        }

        /// <summary>
        /// 출하요청서(임시) 출력
        /// </summary>
        [STAThread]
        private void PS_SD030_Print_Report03()
        {
            string WinTitle;
            string ReportName;
            string sQry;
            PSH_FormHelpClass formHelpClass = new PSH_FormHelpClass();
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                WinTitle = "[PS_SD030_20] 레포트";
                ReportName = "PS_SD030_20.rpt";
                //쿼리 : PS_SD030_10

                sQry = "SELECT BPLName FROM [OBPL] WHERE BPLId = '" + oForm.Items.Item("BPLId").Specific.Value + "'";
                oRecordSet01.DoQuery(sQry);

                List<PSH_DataPackClass> dataPackParameter = new List<PSH_DataPackClass> //Parameter List
                {
                    new PSH_DataPackClass("@SD030HNo", oForm.Items.Item("DocEntry").Specific.Value)
                };

                List<PSH_DataPackClass> dataPackFormula = new List<PSH_DataPackClass> //Formula List
                {
                    new PSH_DataPackClass("@BPLName", oRecordSet01.Fields.Item(0).Value)
                };

                formHelpClass.OpenCrystalReport(WinTitle, ReportName, dataPackParameter, dataPackFormula);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
                    //Raise_EVENT_MATRIX_LINK_PRESSED(FormUID, ref pVal, ref BubbleEvent);
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
                            if (PS_SD030_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_SD030_CheckDataValid() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    //else if (pVal.ItemUID == "Button01") //납품전기(선출요청[PS_SD031] 에서 사용)
                    //{
                    //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    //    {
                    //    }
                    //    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    //    {
                    //    }
                    //    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    //    {
                    //        //작지완료여부검사후 DI '//납품전기가능상황에서 납품(실적이 등록된것)
                    //        if (PS_SD030_ValidateDelivery() == false)
                    //        {
                    //            BubbleEvent = false;
                    //            return;
                    //        }
                    //        if (PS_SD030_DI_API_01() == false)
                    //        {
                    //            BubbleEvent = false;
                    //            return;
                    //        }
                    //        PSH_Globals.SBO_Application.ActivateMenuItem("1281");
                    //    }
                    //}
                    //else if (pVal.ItemUID == "Button04") //반품전기(선출요청[PS_SD031] 에서 사용)
                    //{
                    //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    //    {
                    //    }
                    //    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                    //    {
                    //    }
                    //    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    //    {
                    //        if (oForm.Items.Item("ProgStat").Specific.Selected.Value == "3") //납품상태이면
                    //        {
                    //            for (int i = 1; i <= oMat01.VisualRowCount - 1; i++) //AR송장처리된 문서존재유무검사
                    //            {
                    //                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [INV1] WHERE BaseType = '15' AND BaseEntry = '" + oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value + "' AND BaseLine = '" + oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0)
                    //                {
                    //                    PSH_Globals.SBO_Application.StatusBar.SetText("AR송장처리된 문서가 존재합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    //                    BubbleEvent = false;
                    //                    return;
                    //                }
                    //            }
                    //            //반품처리된 문서존재유무검사
                    //            for (int i = 1; i <= oMat01.VisualRowCount - 1; i++)
                    //            {
                    //                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [RDN1] WHERE BaseType = '15' AND BaseEntry = '" + oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value + "' AND BaseLine = '" + oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0)
                    //                {
                    //                    PSH_Globals.SBO_Application.StatusBar.SetText("반품처리된 문서가 존재합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                    //                    BubbleEvent = false;
                    //                    return;
                    //                }
                    //            }
                    //            if (PS_SD030_DI_API_02() == false)
                    //            {
                    //                BubbleEvent = false;
                    //                return;
                    //            }
                    //            PSH_Globals.SBO_Application.ActivateMenuItem("1281");
                    //        }
                    //    }
                    //}
                    else if (pVal.ItemUID == "Button02") //거래명세서
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_SD030_Print_Report01);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                    else if (pVal.ItemUID == "Button03") //출하요청서
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_SD030_Print_Report02);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
                        }
                    }
                    else if (pVal.ItemUID == "Button05") //출하요청서(임시)
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            System.Threading.Thread thread = new System.Threading.Thread(PS_SD030_Print_Report03);
                            thread.SetApartmentState(System.Threading.ApartmentState.STA);
                            thread.Start();
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
                                PS_SD030_FormItemEnabled();
                                PS_SD030_AddMatrixRow(oMat01.RowCount, true);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                PS_SD030_FormItemEnabled();
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
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrderNum");
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
        private void  Raise_EVENT_VALIDATE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            int i;
            string Query01;
            string ItemCode01;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (PS_SD030_Validate("수정") == false)
                            {
                                oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_SD030L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "OrderNum")
                                {
                                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrderNum").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        errMessage = " ";
                                        throw new Exception();
                                    }
                                    for (i = 1; i <= oMat01.RowCount; i++)
                                    {
                                        if (pVal.Row != i) //현재 선택되어있는 행이 아니면
                                        {
                                            if (oMat01.Columns.Item("OrderNum").Cells.Item(pVal.Row).Specific.Value == oMat01.Columns.Item("OrderNum").Cells.Item(i).Specific.Value)
                                            {
                                                errMessage = "동일한 수주가 존재합니다.";
                                                oMat01.Columns.Item("OrderNum").Cells.Item(pVal.Row).Specific.Value = "";
                                                throw new Exception();
                                            }
                                        }
                                    }
                                    RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    Query01 = "EXEC PS_SD030_01 '" + oMat01.Columns.Item("OrderNum").Cells.Item(pVal.Row).Specific.Value + "','" + oForm.Items.Item("DocType").Specific.Selected.Value + "'";
                                    RecordSet01.DoQuery(Query01);
                                    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                                    {
                                        oDS_PS_SD030L.SetValue("U_OrderNum", pVal.Row - 1, RecordSet01.Fields.Item("OrderNum").Value);
                                        oDS_PS_SD030L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                        oDS_PS_SD030L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
                                        oDS_PS_SD030L.SetValue("U_ItemGpCd", pVal.Row - 1, RecordSet01.Fields.Item("ItemGpCd").Value);
                                        oDS_PS_SD030L.SetValue("U_ItmBsort", pVal.Row - 1, RecordSet01.Fields.Item("ItmBsort").Value);
                                        oDS_PS_SD030L.SetValue("U_ItmMsort", pVal.Row - 1, RecordSet01.Fields.Item("ItmMsort").Value);
                                        oDS_PS_SD030L.SetValue("U_Unit1", pVal.Row - 1, RecordSet01.Fields.Item("Unit1").Value);
                                        oDS_PS_SD030L.SetValue("U_Size", pVal.Row - 1, RecordSet01.Fields.Item("Size").Value);
                                        oDS_PS_SD030L.SetValue("U_ItemType", pVal.Row - 1, RecordSet01.Fields.Item("ItemType").Value);
                                        oDS_PS_SD030L.SetValue("U_Quality", pVal.Row - 1, RecordSet01.Fields.Item("Quality").Value);
                                        oDS_PS_SD030L.SetValue("U_Mark", pVal.Row - 1, RecordSet01.Fields.Item("Mark").Value);
                                        oDS_PS_SD030L.SetValue("U_SbasUnit", pVal.Row - 1, RecordSet01.Fields.Item("SbasUnit").Value);
                                        oDS_PS_SD030L.SetValue("U_SjQty", pVal.Row - 1, RecordSet01.Fields.Item("SjQty").Value);
                                        oDS_PS_SD030L.SetValue("U_SjWeight", pVal.Row - 1, RecordSet01.Fields.Item("SjWeight").Value);
                                        oDS_PS_SD030L.SetValue("U_Qty", pVal.Row - 1, RecordSet01.Fields.Item("Qty").Value);
                                        oDS_PS_SD030L.SetValue("U_UnWeight", pVal.Row - 1, RecordSet01.Fields.Item("UnWeight").Value);
                                        oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, RecordSet01.Fields.Item("Weight").Value);
                                        oDS_PS_SD030L.SetValue("U_Currency", pVal.Row - 1, RecordSet01.Fields.Item("Currency").Value);
                                        oDS_PS_SD030L.SetValue("U_Price", pVal.Row - 1, RecordSet01.Fields.Item("Price").Value);
                                        oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, RecordSet01.Fields.Item("LinTotal").Value);
                                        oDS_PS_SD030L.SetValue("U_WhsCode", pVal.Row - 1, RecordSet01.Fields.Item("WhsCode").Value);
                                        oDS_PS_SD030L.SetValue("U_WhsName", pVal.Row - 1, RecordSet01.Fields.Item("WhsName").Value);
                                        oDS_PS_SD030L.SetValue("U_Comments", pVal.Row - 1, RecordSet01.Fields.Item("Comments").Value);
                                        oDS_PS_SD030L.SetValue("U_TrType", pVal.Row - 1, RecordSet01.Fields.Item("TrType").Value);
                                        oDS_PS_SD030L.SetValue("U_ORDRNum", pVal.Row - 1, RecordSet01.Fields.Item("ORDRNum").Value);
                                        oDS_PS_SD030L.SetValue("U_RDR1Num", pVal.Row - 1, RecordSet01.Fields.Item("RDR1Num").Value);
                                        oDS_PS_SD030L.SetValue("U_Status", pVal.Row - 1, RecordSet01.Fields.Item("Status").Value);
                                        oDS_PS_SD030L.SetValue("U_LineId", pVal.Row - 1, RecordSet01.Fields.Item("LineId").Value);
                                        RecordSet01.MoveNext();
                                    }
                                    if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_SD030L.GetValue("U_OrderNum", pVal.Row - 1).ToString().Trim()))
                                    {
                                        PS_SD030_AddMatrixRow(pVal.Row, false);
                                    }
                                    oMat01.LoadFromDataSource();
                                    oMat01.AutoResizeColumns();

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
                                        oForm.Items.Item("TrType").Enabled = true;
                                    }
                                    oForm.Update();
                                }
                                else if (pVal.ColUID == "Qty")
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                        oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, "0");
                                        oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, "0");
                                    }
                                    else
                                    {
                                        ItemCode01 = oDS_PS_SD030L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
                                        
                                        if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "101") //EA자체품
                                        {
                                            oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value);
                                        }
                                        else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "102") //EAUOM
                                        {
                                            oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_Unit1(ItemCode01))));
                                        }
                                        else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "201") //KGSPEC
                                        {
                                            oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(dataHelpClass.GetItem_Spec1(ItemCode01)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(ItemCode01)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "202") //KG단중
                                        {
                                            oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
                                        }
                                        else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "203") //KG선택
                                        {
                                        }
                                        oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oDS_PS_SD030L.GetValue("U_Weight", pVal.Row - 1).ToString().Trim()) * Convert.ToDouble(oDS_PS_SD030L.GetValue("U_Price", pVal.Row - 1).ToString().Trim())));
                                        oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                }
                                else if (pVal.ColUID == "Weight")
                                {
                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        oDS_PS_SD030L.SetValue("U_Qty", pVal.Row - 1, "0");
                                        oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                        oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, "0");
                                    }
                                    else
                                    {
                                        ItemCode01 = oDS_PS_SD030L.GetValue("U_ItemCode", pVal.Row - 1).ToString().Trim();
                                        if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "101")
                                        {
                                            oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        }
                                        else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "102") //EAUOM
                                        {
                                            oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        }
                                        else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "201") //KGSPEC
                                        {
                                            oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(dataHelpClass.GetItem_Spec1(ItemCode01)) - Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01))) * Convert.ToDouble(dataHelpClass.GetItem_Spec2(ItemCode01)) * 0.02808 * (Convert.ToDouble(dataHelpClass.GetItem_Spec3(ItemCode01)) / 1000) * Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "202") //KG단중
                                        {
                                            oDS_PS_SD030L.SetValue("U_Weight", pVal.Row - 1, Convert.ToString(System.Math.Round(Convert.ToDouble(oMat01.Columns.Item("Qty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(dataHelpClass.GetItem_UnWeight(ItemCode01)) / 1000, 0)));
                                            
                                        }
                                        else if (dataHelpClass.GetItem_SbasUnit(ItemCode01) == "203") //KG선택
                                        {
                                            oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        }
                                        oDS_PS_SD030L.SetValue("U_LinTotal", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oDS_PS_SD030L.GetValue("U_Weight", pVal.Row - 1).ToString().Trim()) * Convert.ToDouble(oDS_PS_SD030L.GetValue("U_Price", pVal.Row - 1).ToString().Trim())));
                                    }
                                }
                                else
                                {
                                    oDS_PS_SD030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
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
                                oDS_PS_SD030H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_SD030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_SD030H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_SD030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_SD030_FormItemEnabled();
                    PS_SD030_AddMatrixRow(oMat01.VisualRowCount, false);
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD030H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_SD030L);
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
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_SD030H", "U_CardCode,U_CardName", "", 0, "", "", "");
                        oDS_PS_SD030H.SetValue("U_Managed", 0, dataHelpClass.GetValue("SELECT isnull(QryGroup15,'N') FROM OCRD WHERE CardCode ='" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                        oDS_PS_SD030H.SetValue("U_CheckYN", 0, "Y");
                    }
                    else if (pVal.ItemUID == "DCardCod" || pVal.ItemUID == "DCardNam")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, (pVal.FormUID), "@PS_SD030H", "U_DCardCod,U_DCardNam", "", 0, "", "", "");
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "WhsCode")
                        {
                            if(((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects != null)
                            {
                                SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당
                                oDS_PS_SD030L.SetValue("U_WhsCode", pVal.Row - 1, oDataTable01.Columns.Item("WhsCode").Cells.Item(0).Value);
                                oDS_PS_SD030L.SetValue("U_WhsName", pVal.Row - 1, oDataTable01.Columns.Item("WhsName").Cells.Item(0).Value);
                                oDataTable01 = null;
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
            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (PS_SD030_Validate("행삭제") == false)
                        {
                            BubbleEvent = false;
                            return;
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        for (int i = 1; i <= oMat01.VisualRowCount; i++)
                        {
                            oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat01.FlushToDataSource();
                        oDS_PS_SD030L.RemoveRecord(oDS_PS_SD030L.Size - 1);
                        oMat01.LoadFromDataSource();
                        if (oMat01.RowCount == 0)
                        {
                            PS_SD030_AddMatrixRow(0, false);
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(oDS_PS_SD030L.GetValue("U_OrderNum", oMat01.RowCount - 1).ToString().Trim()))
                            {
                                PS_SD030_AddMatrixRow(oMat01.RowCount, false);
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
                                if (PS_SD030_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("현재 모드에서는 취소할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
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
                            PS_SD030_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_SD030_FormItemEnabled();
                            PS_SD030_AddMatrixRow(0, true);
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

        #region 선출요청[PS_SD031]에서 사용(구현 필요할 경우를 대비하여 주석 처리후 유지 중(2021.04.03) // 2011.06.25 이후 선출요청으로 등록된 자료 없음, 선출요청 클래스는 삭제하여도 무방할 것으로 보임)
        #region PS_SD030_DI_API_01(납품전기, 선출요청에서 사용)
        //private bool PS_SD030_DI_API_01()
        //{
        //    bool returnValue = false;

        //    int i = 0;
        //    int j = 0;
        //    SAPbobsCOM.Documents oDIObject = null;
        //    int RetVal = 0;
        //    int LineNumCount = 0;
        //    int ResultDocNum = 0;

        //    if (SubMain.Sbo_Company.InTransaction == true)
        //    {
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //    }
        //    SubMain.Sbo_Company.StartTransaction();

        //    ItemInformation = new ItemInformations[1];
        //    ItemInformationCount = 0;
        //    for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
        //    {
        //        Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //        ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
        //        ItemInformation[ItemInformationCount].Qty = oMat01.Columns.Item("Qty").Cells.Item(i).Specific.Value;
        //        ItemInformation[ItemInformationCount].Weight = oMat01.Columns.Item("Weight").Cells.Item(i).Specific.Value;
        //        ItemInformation[ItemInformationCount].Currency = oMat01.Columns.Item("Currency").Cells.Item(i).Specific.Value;
        //        ItemInformation[ItemInformationCount].Price = oMat01.Columns.Item("Price").Cells.Item(i).Specific.Value;
        //        ItemInformation[ItemInformationCount].LineTotal = oMat01.Columns.Item("LinTotal").Cells.Item(i).Specific.Value;
        //        ItemInformation[ItemInformationCount].WhsCode = oMat01.Columns.Item("WhsCode").Cells.Item(i).Specific.Value;
        //        ItemInformation[ItemInformationCount].ORDRNum = Conversion.Val(oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value);
        //        ItemInformation[ItemInformationCount].RDR1Num = Conversion.Val(oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value);
        //        ItemInformation[ItemInformationCount].SD030HNum = Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value);
        //        ItemInformation[ItemInformationCount].SD030LNum = Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value);
        //        ItemInformationCount = ItemInformationCount + 1;
        //    }

        //    LineNumCount = 0;
        //    oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
        //    oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm.Items.Item("BPLId").Specific.Selected.Value));
        //    oDIObject.CardCode = Strings.Trim(oForm.Items.Item("CardCode").Specific.Value);
        //    oDIObject.UserFields.Fields.Item("U_DCardCod").Value = Strings.Trim(oForm.Items.Item("DCardCod").Specific.Value);
        //    oDIObject.UserFields.Fields.Item("U_DCardNam").Value = Strings.Trim(oForm.Items.Item("DCardNam").Specific.Value);
        //    oDIObject.UserFields.Fields.Item("U_TradeType").Value = Strings.Trim(oForm.Items.Item("TrType").Specific.Selected.Value);
        //    if (!string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
        //    {
        //        oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DocDate").Specific.Value, "&&&&-&&-&&"));
        //    }
        //    if (!string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value))
        //    {
        //        oDIObject.DocDueDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm.Items.Item("DueDate").Specific.Value, "&&&&-&&-&&"));
        //    }

        //    for (i = 0; i <= ItemInformationCount - 1; i++)
        //    {
        //        if (i != 0)
        //        {
        //            oDIObject.Lines.Add();
        //        }

        //        oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //        oDIObject.Lines.WarehouseCode = ItemInformation[i].WhsCode;
        //        oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD030";
        //        oDIObject.Lines.BaseType = 17;
        //        oDIObject.Lines.BaseEntry = ItemInformation[i].ORDRNum;
        //        oDIObject.Lines.BaseLine = ItemInformation[i].RDR1Num;

        //        oDIObject.Lines.UserFields.Fields.Item("U_Qty").Value = ItemInformation[i].Qty;
        //        oDIObject.Lines.Quantity = ItemInformation[i].Weight;
        //        oDIObject.Lines.Currency = ItemInformation[i].Currency;
        //        oDIObject.Lines.Price = ItemInformation[i].Price;
        //        oDIObject.Lines.LineTotal = ItemInformation[i].LineTotal;
        //        ItemInformation[i].DLN1Num = LineNumCount;
        //        LineNumCount = LineNumCount + 1;
        //    }
        //    RetVal = oDIObject.Add();
        //    if (RetVal == 0)
        //    {
        //        ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //        //문서상태 납품으로 변경
        //        MDC_PS_Common.DoQuery(("UPDATE [@PS_SD030H] SET U_ProgStat = '3' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'"));
        //        //납품,반품문서번호 업데이트
        //        for (i = 0; i <= ItemInformationCount - 1; i++)
        //        {
        //            MDC_PS_Common.DoQuery(("UPDATE [@PS_SD030L] SET U_ODLNNum = '" + ResultDocNum + "', U_DLN1Num = '" + ItemInformation[i].DLN1Num + "', U_ORDNNum = '', U_RDN1Num = '' WHERE DocEntry = '" + ItemInformation[i].SD030HNum + "' AND LineId = '" + ItemInformation[i].SD030LNum + "'"));
        //        }
        //    }
        //    else
        //    {
        //        goto PS_SD030_DI_API_01_DI_Error;
        //    }

        //    if (SubMain.Sbo_Company.InTransaction == true)
        //    {
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //    }
        //    oMat01.LoadFromDataSource();
        //    oMat01.AutoResizeColumns();
        //    oForm.Update();
        //    oDIObject = null;
        //    return returnValue;
        //    PS_SD030_DI_API_01_DI_Error:
        //    if (SubMain.Sbo_Company.InTransaction == true)
        //    {
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //    }
        //    SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //    returnValue = false;
        //    oDIObject = null;
        //    return returnValue;
        //    PS_SD030_DI_API_01_Error:
        //    if (SubMain.Sbo_Company.InTransaction == true)
        //    {
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //    }
        //    SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_DI_API_01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //    returnValue = false;
        //    oDIObject = null;
        //    return returnValue;
        //}
        #endregion

        #region PS_SD030_DI_API_02(반품전기, 선출요청에서 사용)
        //private bool PS_SD030_DI_API_02()
        //{
        //	bool returnValue = false;
        //	////반품
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	returnValue = true;
        //	object i = null;
        //	int j = 0;
        //	SAPbobsCOM.Documents oDIObject = null;
        //	int RetVal = 0;
        //	int LineNumCount = 0;
        //	int ResultDocNum = 0;

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Company.StartTransaction();

        //	ItemInformation = new ItemInformations[1];
        //	ItemInformationCount = 0;
        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
        //		Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
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
        //		ItemInformation[ItemInformationCount].ORDRNum = Conversion.Val(oMat01.Columns.Item("ORDRNum").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].RDR1Num = Conversion.Val(oMat01.Columns.Item("RDR1Num").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD030HNum = Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].SD030LNum = Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value);
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].ODLNNum = Conversion.Val(oMat01.Columns.Item("ODLNNum").Cells.Item(i).Specific.Value);
        //		////납품
        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		ItemInformation[ItemInformationCount].DLN1Num = Conversion.Val(oMat01.Columns.Item("DLN1Num").Cells.Item(i).Specific.Value);
        //		////납품
        //		////ItemInformation(ItemInformationCount).Check = False
        //		ItemInformationCount = ItemInformationCount + 1;
        //	}

        //	LineNumCount = 0;
        //	oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oReturns);
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
        //		oDIObject.Lines.UserFields.Fields.Item("U_BaseType").Value = "PS_SD030";
        //		////별로 의미가 없을듯..
        //		oDIObject.Lines.BaseType = 15;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseEntry = ItemInformation[i].ODLNNum;
        //		//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		oDIObject.Lines.BaseLine = ItemInformation[i].DLN1Num;

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
        //		ItemInformation[i].RDN1Num = LineNumCount;
        //		LineNumCount = LineNumCount + 1;
        //	}
        //	RetVal = oDIObject.Add();
        //	if (RetVal == 0) {
        //		ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //		////문서상태 반품으로 변경
        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //		MDC_PS_Common.DoQuery(("UPDATE [@PS_SD030H] SET U_ProgStat = '4' WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'"));
        //		////납품,반품문서번호 업데이트
        //		for (i = 0; i <= ItemInformationCount - 1; i++) {
        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			MDC_PS_Common.DoQuery(("UPDATE [@PS_SD030L] SET U_ORDNNum = '" + ResultDocNum + "', U_RDN1Num = '" + ItemInformation[i].RDN1Num + "' WHERE DocEntry = '" + ItemInformation[i].SD030HNum + "' AND LineId = '" + ItemInformation[i].SD030LNum + "'"));
        //		}
        //	} else {
        //		goto PS_SD030_DI_API_02_DI_Error;
        //	}

        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //	}
        //	oMat01.LoadFromDataSource();
        //	oMat01.AutoResizeColumns();
        //	oForm.Update();
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return returnValue;
        //	PS_SD030_DI_API_02_DI_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	returnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return returnValue;
        //	PS_SD030_DI_API_02_Error:
        //	if (SubMain.Sbo_Company.InTransaction == true) {
        //		SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //	}
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_DI_API_02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	returnValue = false;
        //	//UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	oDIObject = null;
        //	return returnValue;
        //}
        #endregion

        #region PS_SD030_FindValidateDocument
        //public bool PS_SD030_FindValidateDocument(string ObjectType)
        //{
        //	bool returnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	returnValue = true;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string Query02 = null;
        //	SAPbobsCOM.Recordset RecordSet02 = null;

        //	int i = 0;
        //	string DocEntry = null;
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
        //	////원본문서

        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	Query01 = " SELECT DocEntry";
        //	Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry = ";
        //	Query01 = Query01 + DocEntry;
        //	if ((oDocType01 == "출하요청")) {
        //		Query01 = Query01 + " AND U_DocType = '1'";
        //	} else if ((oDocType01 == "선출요청")) {
        //		Query01 = Query01 + " AND U_DocType = '2'";
        //	}
        //	RecordSet01.DoQuery(Query01);
        //	if ((RecordSet01.RecordCount == 0)) {
        //		if ((oDocType01 == "출하요청")) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("선출요청문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		} else if ((oDocType01 == "선출요청")) {
        //			SubMain.Sbo_Application.SetStatusBarMessage("출하요청문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //		}
        //		returnValue = false;
        //		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		RecordSet01 = null;
        //		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		RecordSet02 = null;
        //		return returnValue;
        //	}

        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return returnValue;
        //	PS_SD030_FindValidateDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	returnValue = false;
        //	return returnValue;
        //}
        #endregion

        #region PS_SD030_DirectionValidateDocument
        //public bool PS_SD030_DirectionValidateDocument(string DocEntry, string DocEntryNext, string Direction, string ObjectType)
        //{
        //	bool returnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string Query02 = null;
        //	SAPbobsCOM.Recordset RecordSet02 = null;

        //	int i = 0;
        //	string MaxDocEntry = null;
        //	string MinDocEntry = null;
        //	bool DoNext = false;
        //	bool IsFirst = false;
        //	////시작유무
        //	DoNext = true;
        //	IsFirst = true;

        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //	RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //	while ((DoNext == true)) {
        //		if ((IsFirst != true)) {
        //			////문서전체를 경유하고도 유효값을 찾지못했다면
        //			if ((DocEntry == DocEntryNext)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("유효한문서가 존재하지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				returnValue = false;
        //				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				RecordSet01 = null;
        //				//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				RecordSet02 = null;
        //				return returnValue;
        //			}
        //		}
        //		if ((Direction == "Next")) {
        //			Query01 = " SELECT TOP 1 DocEntry";
        //			Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry > ";
        //			Query01 = Query01 + DocEntryNext;
        //			if ((oDocType01 == "출하요청")) {
        //				Query01 = Query01 + " AND U_DocType = '1'";
        //			} else if ((oDocType01 == "선출요청")) {
        //				Query01 = Query01 + " AND U_DocType = '2'";
        //			}
        //			Query01 = Query01 + " ORDER BY DocEntry ASC";
        //		} else if ((Direction == "Prev")) {
        //			Query01 = " SELECT TOP 1 DocEntry";
        //			Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
        //			Query01 = Query01 + DocEntryNext;
        //			if ((oDocType01 == "출하요청")) {
        //				Query01 = Query01 + " AND U_DocType = '1'";
        //			} else if ((oDocType01 == "선출요청")) {
        //				Query01 = Query01 + " AND U_DocType = '2'";
        //			}
        //			Query01 = Query01 + " ORDER BY DocEntry DESC";
        //		}
        //		RecordSet01.DoQuery(Query01);
        //		////해당문서가 마지막문서라면
        //		if ((RecordSet01.Fields.Item(0).Value == 0)) {
        //			if ((Direction == "Next")) {
        //				Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
        //				if ((oDocType01 == "출하요청")) {
        //					Query02 = Query02 + " WHERE U_DocType = '1'";
        //				} else if ((oDocType01 == "선출요청")) {
        //					Query02 = Query02 + " WHERE U_DocType = '2'";
        //				}
        //				Query02 = Query02 + " ORDER BY DocEntry ASC";
        //			} else if ((Direction == "Prev")) {
        //				Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
        //				if ((oDocType01 == "출하요청")) {
        //					Query02 = Query02 + " WHERE U_DocType = '1'";
        //				} else if ((oDocType01 == "선출요청")) {
        //					Query02 = Query02 + " WHERE U_DocType = '2'";
        //				}
        //				Query02 = Query02 + " ORDER BY DocEntry DESC";
        //			}
        //			RecordSet02.DoQuery(Query02);
        //			////문서가 아예 존재하지 않는다면
        //			if ((RecordSet02.RecordCount == 0)) {
        //				SubMain.Sbo_Application.SetStatusBarMessage("유효한문서가 존재하지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				RecordSet01 = null;
        //				//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //				RecordSet02 = null;
        //				returnValue = false;
        //				return returnValue;
        //			} else {
        //				if ((Direction == "Next")) {
        //					DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) - 1);
        //					Query01 = " SELECT TOP 1 DocEntry";
        //					Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry > ";
        //					Query01 = Query01 + DocEntryNext;
        //					if ((oDocType01 == "출하요청")) {
        //						Query01 = Query01 + " AND U_DocType = '1'";
        //					} else if ((oDocType01 == "선출요청")) {
        //						Query01 = Query01 + " AND U_DocType = '2'";
        //					}
        //					Query01 = Query01 + " ORDER BY DocEntry ASC";
        //					RecordSet01.DoQuery(Query01);
        //				} else if ((Direction == "Prev")) {
        //					DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) + 1);
        //					Query01 = " SELECT TOP 1 DocNum";
        //					Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
        //					Query01 = Query01 + DocEntryNext;
        //					if ((oDocType01 == "출하요청")) {
        //						Query01 = Query01 + " AND U_DocType = '1'";
        //					} else if ((oDocType01 == "선출요청")) {
        //						Query01 = Query01 + " AND U_DocType = '2'";
        //					}
        //					Query01 = Query01 + " ORDER BY DocEntry DESC";
        //					RecordSet01.DoQuery(Query01);
        //				}
        //			}
        //		}
        //		if ((oDocType01 == "출하요청")) {
        //			DoNext = false;
        //			if ((Direction == "Next")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
        //			} else if ((Direction == "Prev")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
        //			}
        //		} else if ((oDocType01 == "선출요청")) {
        //			DoNext = false;
        //			if ((Direction == "Next")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
        //			} else if ((Direction == "Prev")) {
        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
        //			}
        //		}
        //		IsFirst = false;
        //	}
        //	////다음문서가 유효하다면 그냥 넘어가고
        //	if ((DocEntry == DocEntryNext)) {
        //		PS_SD030_FormItemEnabled();
        //		//
        //	////다음문서가 유효하지 않다면
        //	} else {
        //		oForm.Freeze(true);
        //		oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //		PS_SD030_FormItemEnabled();
        //		//
        //		////문서번호 필드가 입력이 가능하다면
        //		if (oForm.Items.Item("DocEntry").Enabled == true) {
        //			if ((Direction == "Next")) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.Value = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) + 1));
        //			} else if ((Direction == "Prev")) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				oForm.Items.Item("DocEntry").Specific.Value = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) - 1));
        //			}
        //			oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //		}
        //		oForm.Freeze(false);
        //		returnValue = false;
        //		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		RecordSet01 = null;
        //		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //		RecordSet02 = null;
        //		return returnValue;
        //	}
        //	returnValue = true;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return returnValue;
        //	PS_SD030_DirectionValidateDocument_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	returnValue = false;
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet02 = null;
        //	return returnValue;
        //}
        #endregion

        #region PS_SD030_ValidateDelivery
        //private bool PS_SD030_ValidateDelivery()
        //{
        //	bool returnValue = false;
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	returnValue = true;
        //	if (PS_SD030_ValidateCreditLine() == false) {
        //		returnValue = false;
        //		return returnValue;
        //	}
        //	return returnValue;
        //	PS_SD030_ValidateDelivery_Error:

        //	////생산완료 되었는지 확인
        //	//    Dim i As Long
        //	//
        //	//    For i = 1 To oMat01.VisualRowCount - 1
        //	//        mdc_ps_common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H
        //	//    Next
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_ValidateDelivery_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return returnValue;
        //}
        #endregion

        #region PS_SD030_ValidateCreditLine
        //private bool PS_SD030_ValidateCreditLine()
        //{
        //	bool returnValue = false;
        //	////여신한도체크
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	returnValue = true;
        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;

        //	int i = 0;
        //	decimal OCRDCreditLine = default(decimal);
        //	////고객여신한도
        //	decimal SD080CreditLine = default(decimal);
        //	////추가여신한도
        //	decimal OCRDBalance = default(decimal);
        //	////계정잔액
        //	decimal OCRDDNotesBal = default(decimal);
        //	////납품액
        //	decimal CurrentLineSum = default(decimal);
        //	////현재문서총계
        //	decimal OutPreP = default(decimal);
        //	//출고예정금액
        //	////If oMat01.Columns("WhsCode").Cells(1).Specific.Value = "104" And oMat01.Columns("ItmBsort").Cells(1).Specific.Value = "101" Then
        //	////휘팅 서울출고
        //	//UPGRADE_WARNING: oMat01.Columns(ItmBsort).Cells(1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	//UPGRADE_WARNING: oMat01.Columns(WhsCode).Cells(1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	if (oMat01.Columns.Item("WhsCode").Cells.Item(1).Specific.Value == "101" & oMat01.Columns.Item("ItmBsort").Cells.Item(1).Specific.Value == "111") {
        //		////창원 분말
        //		if (oDS_PS_SD030H.GetValue("U_PrtYn", 0) != "Y") {
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			//UPGRADE_WARNING: oForm.Items(CardCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //			Query01 = "EXEC PS_SD081_hando '" + Strings.Trim(oDS_PS_SD030H.GetValue("U_BPLId", 0)) + "', '" + oForm.Items.Item("CardCode").Specific.Value + "','" + oForm.Items.Item("DocDate").Specific.Value + "'";
        //			RecordSet01.DoQuery(Query01);

        //			if (RecordSet01.RecordCount > 0) {
        //				////여신한도금액
        //				if (RecordSet01.Fields.Item("OverAmt").Value > 0) {
        //					MDC_Com.MDC_GF_Message(ref "여신한도로 초과했습니다.", ref "W");
        //					returnValue = false;
        //					return returnValue;
        //				}
        //			}
        //		}


        //		//EXEC [PS_SD081_hando] '4','12494','20110316'

        //		//        OCRDCreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT CreditLine FROM [OCRD] WHERE CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//        SD080CreditLine = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Sum(PS_SD080L.U_RequestP) FROM [@PS_SD080H] PS_SD080H LEFT JOIN [@PS_SD080L] PS_SD080L ON PS_SD080H.DocEntry = PS_SD080L.DocEntry WHERE PS_SD080H.Canceled = 'N' And PS_SD080H.U_OkYN = 'Y' AND PS_SD080H.U_DocDate = '" & oForm.Items("DocDate").Specific.Value & "' AND PS_SD080L.U_CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//'        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT Balance FROM [OCRD] WHERE CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//        OCRDBalance = MDC_PS_Common.GetValue("SELECT ISNULL((Select Sum(Debit - Credit) from JDT1 where ShortName = '" & oForm.Items("CardCode").Specific.Value & "' And Account = '11104010'),0)", 0, 1)
        //		//        OCRDDNotesBal = MDC_PS_Common.GetValue("SELECT ISNULL((SELECT DNotesBal FROM [OCRD] WHERE CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//'        OutPreP = MDC_PS_Common.GetValue("Select IsNull((Select Sum(LineTotal) From  [ORDR] a Inner Join [RDR1] b On a.DocEntry = b.DocEntry Where  b.LineStatus = 'O' And a.DocStatus = 'O' And a.CardCode = '" & oForm.Items("CardCode").Specific.Value & "'),0)", 0, 1)
        //		//        CurrentLineSum = 0
        //		//        For i = 1 To oMat01.VisualRowCount - 1
        //		//            CurrentLineSum = CurrentLineSum + Val(oMat01.Columns("LinTotal").Cells(i).Specific.Value)
        //		//        Next
        //		//'        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal) < CurrentLineSum) Then
        //		//        If ((OCRDCreditLine + SD080CreditLine) - (OCRDBalance + OCRDDNotesBal + OutPreP) < CurrentLineSum) Then
        //		//            Call MDC_Com.MDC_GF_Message("여신한도가 부족합니다.", "W")
        //		//            PS_SD030_ValidateCreditLine = False
        //		//            Exit Function
        //		//        End If
        //	}
        //	return returnValue;
        //	PS_SD030_ValidateCreditLine_Error:
        //	returnValue = false;
        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_SD030_ValidateCreditLine_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //	return returnValue;
        //}
        #endregion

        #region Raise_EVENT_RECORD_MOVE
        //private void Raise_EVENT_RECORD_MOVE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
        //{
        //	 // ERROR: Not supported in C#: OnErrorStatement

        //	string Query01 = null;
        //	SAPbobsCOM.Recordset RecordSet01 = null;
        //	string DocEntry = null;
        //	string DocEntryNext = null;
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
        //	////원본문서
        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //	DocEntryNext = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
        //	////다음문서

        //	////다음
        //	if (pVal.MenuUID == "1288") {
        //		if (pVal.BeforeAction == true) {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				SubMain.Sbo_Application.ActivateMenuItem(("1290"));
        //				BubbleEvent = false;
        //				return;
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if ((string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))) {
        //					SubMain.Sbo_Application.ActivateMenuItem(("1290"));
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //			if (PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_SD030H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////이전
        //	} else if (pVal.MenuUID == "1289") {
        //		if (pVal.BeforeAction == true) {
        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
        //				SubMain.Sbo_Application.ActivateMenuItem(("1291"));
        //				BubbleEvent = false;
        //				return;
        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //				if ((string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))) {
        //					SubMain.Sbo_Application.ActivateMenuItem(("1291"));
        //					BubbleEvent = false;
        //					return;
        //				}
        //			}
        //			if (PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_SD030H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////첫번째레코드로이동
        //	} else if (pVal.MenuUID == "1290") {
        //		if (pVal.BeforeAction == true) {
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			Query01 = " SELECT TOP 1 DocEntry FROM [@PS_SD030H] ORDER BY DocEntry DESC";
        //			////가장마지막행을 부여
        //			RecordSet01.DoQuery(Query01);
        //			DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////원본문서
        //			DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////다음문서
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			if (PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_SD030H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	////마지막문서로이동
        //	} else if (pVal.MenuUID == "1291") {
        //		if (pVal.BeforeAction == true) {
        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //			Query01 = " SELECT TOP 1 DocEntry FROM [@PS_SD030H] ORDER BY DocEntry ASC";
        //			////가장 첫행을 부여
        //			RecordSet01.DoQuery(Query01);
        //			DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////원본문서
        //			DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //			////다음문서
        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //			RecordSet01 = null;
        //			if (PS_SD030_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_SD030H") == false) {
        //				BubbleEvent = false;
        //				return;
        //			}
        //		}
        //	}
        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //	RecordSet01 = null;
        //	return;
        //	Raise_EVENT_RECORD_MOVE_Error:
        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RECORD_MOVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion
        #endregion
    }
}
