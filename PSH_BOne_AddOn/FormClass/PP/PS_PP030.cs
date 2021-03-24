using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
	/// <summary>
	/// 작업지시등록
	/// </summary>
	internal class PS_PP030 : PSH_BaseClass
	{
		private string oFormUniqueID;
		private SAPbouiCOM.Matrix oMat01;
		private SAPbouiCOM.Matrix oMat02;
		private SAPbouiCOM.Matrix oMat03;
		private SAPbouiCOM.DBDataSource oDS_PS_USERDS01; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP030H; //등록헤더
		private SAPbouiCOM.DBDataSource oDS_PS_PP030L; //등록라인
		private SAPbouiCOM.DBDataSource oDS_PS_PP030M; //등록라인
		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
		private int oMat01Row01;
		private int oMat02Row02;
		private int oMat03Row03;
        private string oDocEntry01;
        private string oSCardCod01;
        private SAPbouiCOM.BoFormMode oFormMode01;
        private bool oHasMatrix01;

        /// <summary>
        /// Form 호출
        /// </summary>
        /// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
		{
			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

			try
			{
				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP030.srf");
				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

				//매트릭스의 타이틀높이와 셀높이를 고정
				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
				{
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
				}

				oFormUniqueID = "PS_PP030_" + SubMain.Get_TotalFormsCount();
				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP030");

				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

				oForm.SupportedModes = -1;
				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
				oForm.DataBrowser.BrowseBy = "DocEntry";

				oForm.Freeze(true);
                PS_PP030_CreateItems();
                PS_PP030_ComboBox_Setting();
                PS_PP030_CF_ChooseFromList();
                PS_PP030_EnableMenus();
                PS_PP030_SetDocument(oFormDocEntry);
                PS_PP030_FormResize();

                oForm.EnableMenu("1283", false); //삭제
				oForm.EnableMenu("1287", false); //복제
				oForm.EnableMenu("1286", true); //닫기
				oForm.EnableMenu("1284", true); //취소
				oForm.EnableMenu("1293", true); //행삭제
				oForm.EnableMenu("1299", false); //행닫기
            }
			catch(Exception ex)
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
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP030_CreateItems()
        {
            try
            {
                oDS_PS_USERDS01 = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
                oDS_PS_PP030H = oForm.DataSources.DBDataSources.Item("@PS_PP030H");
                oDS_PS_PP030L = oForm.DataSources.DBDataSources.Item("@PS_PP030L");
                oDS_PS_PP030M = oForm.DataSources.DBDataSources.Item("@PS_PP030M");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat03 = oForm.Items.Item("Mat03").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();

                oForm.DataSources.UserDataSources.Add("SBPLId", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("SBPLId").Specific.DataBind.SetBound(true, "", "SBPLId");

                oForm.DataSources.UserDataSources.Add("ItmBsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmBsort").Specific.DataBind.SetBound(true, "", "ItmBsort");

                oForm.DataSources.UserDataSources.Add("ItmMsort", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ItmMsort").Specific.DataBind.SetBound(true, "", "ItmMsort");

                oForm.DataSources.UserDataSources.Add("ReqType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("ReqType").Specific.DataBind.SetBound(true, "", "ReqType");

                oForm.DataSources.UserDataSources.Add("SItemCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SItemCod").Specific.DataBind.SetBound(true, "", "SItemCod");

                oForm.DataSources.UserDataSources.Add("SItemNam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SItemNam").Specific.DataBind.SetBound(true, "", "SItemNam");

                oForm.DataSources.UserDataSources.Add("SCardCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("SCardCod").Specific.DataBind.SetBound(true, "", "SCardCod");

                oForm.DataSources.UserDataSources.Add("SCardNam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("SCardNam").Specific.DataBind.SetBound(true, "", "SCardNam");

                oForm.DataSources.UserDataSources.Add("ReqCod", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                oForm.Items.Item("ReqCod").Specific.DataBind.SetBound(true, "", "ReqCod");

                oForm.DataSources.UserDataSources.Add("ReqNam", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.Items.Item("ReqNam").Specific.DataBind.SetBound(true, "", "ReqNam");

                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");

                oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");

                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");

                oForm.DataSources.UserDataSources.Add("Total", SAPbouiCOM.BoDataType.dt_SUM);
                oForm.Items.Item("Total").Specific.DataBind.SetBound(true, "", "Total");
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP030_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "ReqType", "", "10", "계획생산요청");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "ReqType", "", "20", "수주생산요청");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("ReqType").Specific, "PS_PP030", "ReqType", true);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ReqType", "10", "계획생산요청");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat01", "ReqType", "20", "수주생산요청");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat01.Columns.Item("ReqType"), "PS_PP030", "Mat01", "ReqType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "BasicGub", "", "10", "통합");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "BasicGub", "", "20", "비통합");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("BasicGub").Specific, "PS_PP030", "BasicGub", false);
                oForm.Items.Item("BasicGub").Specific.Select("비통합", SAPbouiCOM.BoSearchKey.psk_ByValue);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "MulGbn1", "", "10", "탈지");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "MulGbn1", "", "20", "비탈지");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("MulGbn1").Specific, "PS_PP030", "MulGbn1", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "MulGbn2", "", "10", "시계");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "MulGbn2", "", "20", "반시계");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("MulGbn2").Specific, "PS_PP030", "MulGbn2", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "MulGbn3", "", "10", "배면");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "MulGbn3", "", "20", "상면");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("MulGbn3").Specific, "PS_PP030", "MulGbn3", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "10", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "20", "원재료");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn", "30", "스크랩");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn2", "20", "원재료");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat02", "InputGbn2", "30", "스크랩");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat02.Columns.Item("InputGbn"), "PS_PP030", "Mat02", "InputGbn", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "10", "청구");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "20", "잔재");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat02", "ProcType", "30", "취소");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat02.Columns.Item("ProcType"), "PS_PP030", "Mat02", "ProcType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "10", "자가");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "20", "정밀");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "WorkGbn", "30", "외주");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("WorkGbn"), "PS_PP030", "Mat03", "WorkGbn", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ResultYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ResultYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("ResultYN"), "PS_PP030", "Mat03", "ResultYN", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReWorkYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReWorkYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("ReWorkYN"), "PS_PP030", "Mat03", "ReWorkYN", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReportYN", "Y", "예");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP030", "Mat03", "ReportYN", "N", "아니오");
                dataHelpClass.Combo_ValidValues_SetValueColumn(oMat03.Columns.Item("ReportYN"), "PS_PP030", "Mat03", "ReportYN", false);

                dataHelpClass.Set_ComboList(oForm.Items.Item("SBPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, true);
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmBsort").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", false, true);
                dataHelpClass.Set_ComboList(oForm.Items.Item("ItmMsort").Specific, "SELECT PSH_ITMMSORT.U_Code, PSH_ITMMSORT.U_CodeName FROM [@PSH_ITMMSORT] PSH_ITMMSORT LEFT JOIN [@PSH_ITMBSORT] PSH_ITMBSORT ON PSH_ITMBSORT.Code = PSH_ITMMSORT.U_rCode WHERE PSH_ITMBSORT.U_PudYN = 'Y'", "", false, true);
                dataHelpClass.Set_ComboList(oForm.Items.Item("Mark").Specific, "SELECT Code, Name FROM [@PSH_MARK] order by Code", "", false, true);
                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", false, false);
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("ItmBsort"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat02.Columns.Item("ItemGpCd"), "SELECT ItmsGrpCod, ItmsGrpNam FROM [OITB]", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat03.Columns.Item("Unit"), "SELECT Code, Name FROM [@PSH_CPUOM]", "", "");

                //재청구사유(라인.Mat02)
                sQry = "  SELECT    U_Minor,";
                sQry += "           U_CdName";
                sQry += " FROM      [@PS_SY001L]";
                sQry += " WHERE     Code = 'P203'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += "           AND U_Minor <> 'A'";
                sQry += " ORDER BY  U_Seq";
                dataHelpClass.GP_MatrixSetMatComboList(oMat02.Columns.Item("RCode"), sQry, "", "");

                oForm.Items.Item("SBPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue); //조회영역 사업장:사용자 소속사업장으로 세팅
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_PP030_CF_ChooseFromList()
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
            SAPbouiCOM.ChooseFromList oCFL = null;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
            SAPbouiCOM.EditText oEdit = null;
            
            try
            {
                //거래처코드
                oEdit = oForm.Items.Item("SCardCod").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = "2";
                oCFLCreationParams.UniqueID = "CFLSCARDCOD";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oEdit.ChooseFromListUID = "CFLSCARDCOD";
                oEdit.ChooseFromListAlias = "CardCode";

                //품목코드
                oEdit = oForm.Items.Item("SItemCod").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLSITEMCOD";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oEdit.ChooseFromListUID = "CFLSITEMCOD";
                oEdit.ChooseFromListAlias = "ItemCode";
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP030_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, true, false, false, false, false, false);
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_PP030_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP030_FormItemEnabled();
                    PS_PP030_AddMatrixRow01(0, true);
                    PS_PP030_AddMatrixRow02(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_PP030_FormItemEnabled();
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
        /// 각모드에따른 아이템설정
        /// </summary>        
        private void PS_PP030_FormItemEnabled()
        {
            string sQry01;
            string sQry02;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                oForm.Freeze(true);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("BasicGub").Enabled = true;
                    oForm.Items.Item("MulGbn1").Enabled = false;
                    oForm.Items.Item("MulGbn2").Enabled = false;
                    oForm.Items.Item("MulGbn3").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("OrdMgNum").Enabled = true;
                    oForm.Items.Item("ReqWt").Enabled = false;
                    oForm.Items.Item("SelWt").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Mat02").Enabled = true;
                    oForm.Items.Item("Mat03").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;

                    oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
                    oForm.Items.Item("DueDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
                    oForm.Items.Item("SItemCod").Specific.Value = "";
                    oForm.Items.Item("SCardCod").Specific.Value = "";
                    oForm.Items.Item("OrdMgNum").Specific.Value = "";
                    oForm.Items.Item("OrdNum").Specific.Value = "";
                    oForm.Items.Item("OrdSub1").Specific.Value = "";
                    oForm.Items.Item("OrdSub2").Specific.Value = "";
                    oForm.Items.Item("BasicGub").Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("OrdMgNum").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(DateAndTime.Today, "YYYYMMDD");
                    oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                    oForm.Items.Item("Total").Specific.Value = "0";
                    PS_PP030_FormClear();
                    
                    oMat02.Columns.Item("BatchNum").Editable = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("OrdGbn").Enabled = true;
                    oForm.Items.Item("BPLId").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("DueDate").Enabled = true;
                    oForm.Items.Item("ItemCode").Enabled = true;
                    oForm.Items.Item("CntcCode").Enabled = true;
                    oForm.Items.Item("BasicGub").Enabled = true;
                    oForm.Items.Item("MulGbn1").Enabled = false;
                    oForm.Items.Item("MulGbn2").Enabled = false;
                    oForm.Items.Item("MulGbn3").Enabled = false;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("OrdMgNum").Enabled = true;
                    oForm.Items.Item("ReqWt").Enabled = false;
                    oForm.Items.Item("SelWt").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.Items.Item("Mat02").Enabled = false;
                    oForm.Items.Item("Mat03").Enabled = false;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;
                    oMat02.Columns.Item("BatchNum").Editable = false;
                    oForm.Items.Item("Total").Specific.Value = "0";
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oMat02.Columns.Item("BatchNum").Editable = false;
                    
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" + oDS_PS_PP030H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "Y")
                    {
                        oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("OrdGbn").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("DueDate").Enabled = false;
                        oForm.Items.Item("ItemCode").Enabled = false;
                        oForm.Items.Item("CntcCode").Enabled = false;
                        oForm.Items.Item("MulGbn1").Enabled = false;
                        oForm.Items.Item("MulGbn2").Enabled = false;
                        oForm.Items.Item("MulGbn3").Enabled = false;
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("OrdMgNum").Enabled = false;
                        oForm.Items.Item("ReqWt").Enabled = false;
                        oForm.Items.Item("SelWt").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("Mat02").Enabled = false;
                        oForm.Items.Item("Mat03").Enabled = false;
                        oForm.Items.Item("Button01").Enabled = false;
                        oForm.Items.Item("1").Enabled = false;
                    }
                    else
                    {
                        oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("Mat01").Enabled = true;
                        oForm.Items.Item("Mat02").Enabled = true;
                        oForm.Items.Item("Mat03").Enabled = true;
                        oForm.Items.Item("Button01").Enabled = true;
                        oForm.Items.Item("1").Enabled = true;

                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("OrdGbn").Enabled = false;
                        oForm.Items.Item("BPLId").Enabled = false;
                        oForm.Items.Item("ItemCode").Enabled = false;
                        oForm.Items.Item("OrdMgNum").Enabled = false;

                        //실적(작업일보)문서가 없고 원가 상의 재공에 투입된 자료가 아니라면 아래 필드의 데이터는 수정(2017.02.21 송명규)
                        //실적 자료 조회용 쿼리
                        sQry01 = "  SELECT  COUNT(*)";
                        sQry01 += " FROM    [@PS_PP040H] AS T0";
                        sQry01 += "         INNER JOIN";
                        sQry01 += "         [@PS_PP040L] AS T1";
                        sQry01 += "             ON T0.DocEntry = T1.DocEntry";
                        sQry01 += " WHERE   T0.Canceled = 'N'";
                        sQry01 += "         AND T1.U_PP030HNo = " + oDS_PS_PP030H.GetValue("DocEntry", 0).ToString().Trim();

                        //원가 자료 조회용 쿼리
                        sQry02 = "  SELECT  COUNT(*)";
                        sQry02 += " FROM    [@PS_CO130L] AS T0";
                        sQry02 += " WHERE   T0.U_POEntry = " + oDS_PS_PP030H.GetValue("DocEntry", 0).ToString().Trim();

                        if (dataHelpClass.GetValue(sQry01, 0, 1) == "0" || dataHelpClass.GetValue(sQry02, 0, 1) == "0") //실적(작업일보) 및 원가계산된 자료가 없으면 아래 자료는 수정가능
                        {
                            oForm.Items.Item("DocDate").Enabled = true;
                            oForm.Items.Item("DueDate").Enabled = true;
                            oForm.Items.Item("CntcCode").Enabled = true;
                            oForm.Items.Item("SelWt").Enabled = true;
                            oForm.Items.Item("ReqWt").Enabled = true;
                            
                            if (oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "104") //작업구분 멀티
                            {
                                oForm.Items.Item("BasicGub").Enabled = true;
                                oForm.Items.Item("MulGbn1").Enabled = true;
                                oForm.Items.Item("MulGbn2").Enabled = true;
                                oForm.Items.Item("MulGbn3").Enabled = true;
                            }
                            else
                            {
                                oForm.Items.Item("BasicGub").Enabled = false;
                                oForm.Items.Item("MulGbn1").Enabled = false;
                                oForm.Items.Item("MulGbn2").Enabled = false;
                                oForm.Items.Item("MulGbn3").Enabled = false;
                            }
                        }
                        else //실적 등록 및 원가 계산된 자료가 있으면
                        {   
                            if (codeHelpClass.Left(oDS_PS_PP030H.GetValue("U_OrdNum", 0).ToString().Trim(), 1) == "E") //멀티 작번인 경우
                            {
                                oForm.Items.Item("DocDate").Enabled = false;
                                oForm.Items.Item("DueDate").Enabled = false;
                                oForm.Items.Item("CntcCode").Enabled = false;
                                oForm.Items.Item("ReqWt").Enabled = true;
                                oForm.Items.Item("SelWt").Enabled = true;
                                oForm.Items.Item("MulGbn1").Enabled = true;
                                oForm.Items.Item("MulGbn2").Enabled = false;
                                oForm.Items.Item("MulGbn3").Enabled = false;
                            }
                            else
                            {
                                oForm.Items.Item("DocDate").Enabled = false;
                                oForm.Items.Item("DueDate").Enabled = false;
                                oForm.Items.Item("CntcCode").Enabled = false;
                                oForm.Items.Item("ReqWt").Enabled = false;
                                oForm.Items.Item("SelWt").Enabled = false;
                                oForm.Items.Item("MulGbn1").Enabled = true;
                                oForm.Items.Item("MulGbn2").Enabled = false;
                                oForm.Items.Item("MulGbn3").Enabled = false;
                            }
                        }

                        if (oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "107")
                        {
                            oMat02.Columns.Item("InputGbn").Editable = true;
                        }
                        else
                        {
                            oMat02.Columns.Item("InputGbn").Editable = false;
                        }

                        if (oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "105" || oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "106")
                        {
                            oMat02.Columns.Item("Weight").Editable = true;
                        }
                        else
                        {
                            oMat02.Columns.Item("Weight").Editable = false;
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
        private void PS_PP030_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP030'", "");

                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
                }
                else
                {
                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
                }

                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value != "105" && oForm.Items.Item("OrdGbn").Specific.Selected.Value != "106" && oForm.Items.Item("OrdGbn").Specific.Selected.Value != "선택")
                {
                    if (!string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.Value))
                    {
                        oForm.Items.Item("OrdNum").Specific.Value = oForm.Items.Item("OrdMgNum").Specific.Value + dataHelpClass.GetValue("EXEC PS_PP030_01 '" + oForm.Items.Item("OrdMgNum").Specific.Value + "'", 0, 1);
                        oForm.Items.Item("OrdSub1").Specific.Value = "00";
                        oForm.Items.Item("OrdSub2").Specific.Value = "000";
                    }
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 Row추가(Mat02)
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param> 
        private void PS_PP030_AddMatrixRow01(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false)
                {
                    oDS_PS_PP030L.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_PP030L.Offset = oRow;
                oDS_PS_PP030L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                
                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim() == "107") //엔드베어링은 투입구분,원재료,스크랩
                {
                    oDS_PS_PP030L.SetValue("U_InputGbn", oRow, "20");
                }
                else //나머지경우는 일반으로 선택
                {
                    oDS_PS_PP030L.SetValue("U_InputGbn", oRow, "10");
                }
                oDS_PS_PP030L.SetValue("U_ProcType", oRow, "20");
                oMat02.LoadFromDataSource();
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
        /// 메트릭스 Row추가(Mat03)
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param> 
        private void PS_PP030_AddMatrixRow02(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);
                
                if (RowIserted == false)
                {
                    oDS_PS_PP030M.InsertRecord(oRow);
                }
                oMat03.AddRow();
                oDS_PS_PP030M.Offset = oRow;
                oDS_PS_PP030M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PS_PP030M.SetValue("U_Sequence", oRow, Convert.ToString(oRow + 1));
                oDS_PS_PP030M.SetValue("U_WorkGbn", oRow, "10");
                oDS_PS_PP030M.SetValue("U_ReWorkYN", oRow, "N");
                oDS_PS_PP030M.SetValue("U_ResultYN", oRow, "N");
                oDS_PS_PP030M.SetValue("U_ReportYN", oRow, "Y");
                oMat03.LoadFromDataSource();
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
        /// 필수입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP030_DataValidCheck()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            int i = 0;
            short Lot104Exsits;
            string query01;
            double baseItemWeight = 0;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClass = new PSH_CodeHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP030_FormClear();
                }

                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                {
                    errCode = "1";
                    throw new Exception();
                }
                else if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택")
                {
                    errCode = "2";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value) || string.IsNullOrEmpty(oForm.Items.Item("OrdSub1").Specific.Value) || string.IsNullOrEmpty(oForm.Items.Item("OrdSub2").Specific.Value))
                {
                    errCode = "3";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                {
                    errCode = "4";
                    throw new Exception();
                }
                else if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                {
                    errCode = "5";
                    throw new Exception();
                }
                else if (Convert.ToDouble(oForm.Items.Item("SelWt").Specific.Value) <= 0)
                {
                    errCode = "6";
                    throw new Exception();
                }

                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104")
                {
                    query01 = "EXEC [PS_PP030_09] '" + oForm.Items.Item("ItemCode").Specific.Value + "','";
                    query01 += codeHelpClass.Right(codeHelpClass.Left(oForm.Items.Item("DocDate").Specific.Value, 6), 4) + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() + "'";
                    RecordSet01.DoQuery(query01);
                    if (RecordSet01.Fields.Item(0).Value == 1)
                    {
                        errCode = "7";
                        baseItemWeight = RecordSet01.Fields.Item(1).Value;
                        throw new Exception();
                    }

                    query01 = "SELECT Count(*) FROM Z_DSMDFRY Where lotno = '" + oForm.Items.Item("OrdNum").Specific.Value + "'";
                    RecordSet01.DoQuery(query01);
                    Lot104Exsits = Convert.ToInt32(RecordSet01.Fields.Item(0).Value);

                    if (Lot104Exsits == 0)
                    {
                        if (oMat02.VisualRowCount <= 1)
                        {
                            errCode = "8";
                            throw new Exception();
                        }
                    }
                }
                else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105")
                {

                    if (PS_PP030_CheckDate() == false)
                    {
                        errCode = "9";
                        throw new Exception();
                    }

                }
                else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101") //휘팅일경우
                {
                }

                if (oMat03.VisualRowCount <= 1)
                {
                    errCode = "10";
                    throw new Exception();
                }

                for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                {
                    if (oMat02.Columns.Item("InputGbn").Cells.Item(i).Specific.Selected == null)
                    {
                        errCode = "11";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value))
                    {
                        errCode = "12";
                        throw new Exception();
                    }
                    else if (oMat02.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Selected == null)
                    {
                        errCode = "13";
                        throw new Exception();
                    }
                    else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value != "104"
                          && oForm.Items.Item("OrdGbn").Specific.Selected.Value != "105"
                          && oForm.Items.Item("OrdGbn").Specific.Selected.Value != "106"
                          && oForm.Items.Item("OrdGbn").Specific.Selected.Value != "107") //휘팅,부품,엔드베어링일경우
                    {
                        if (oMat02.VisualRowCount > 2)
                        {
                            errCode = "14";
                            throw new Exception();
                        }
                    }
                    else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106") //기계공구,몰드인경우
                    {
                        if (oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected == null)
                        {
                            errCode = "15";
                            throw new Exception();
                        }
                        else if (Convert.ToDouble(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.Value) <= 0)
                        {
                            errCode = "16";
                            throw new Exception();
                        }

                        if (PS_PP030_Check_DupReq(oForm.Items.Item("DocEntry").Specific.Value, oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value, oMat02.Columns.Item("LineId").Cells.Item(i).Specific.Value) == true) //원재료 중복 청구 시(2018.09.17 송명규, 김석태 과장 요청)
                        {
                            if (oMat02.Columns.Item("RCode").Cells.Item(i).Specific.Selected == null)
                            {
                                errCode = "17";
                                throw new Exception();
                            }
                        }
                    }
                    else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107") //멀티,엔드베어링인경우
                    {
                        if (Convert.ToDouble(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.Value) <= 0)
                        {
                            errCode = "18";
                            throw new Exception();
                        }
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("BatchNum").Cells.Item(i).Specific.Value))
                        {
                            errCode = "19";
                            throw new Exception();
                        }
                    }
                }

                for (i = 1; i <= oMat03.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.Value))
                    {
                        errCode = "20";
                        throw new Exception();
                    }
                    else if (string.IsNullOrEmpty(oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value))
                    {
                        errCode = "21";
                        throw new Exception();
                    }
                }

                if (PS_PP030_Validate("검사01") == false)
                {
                    errCode = "22";
                    throw new Exception();
                }

                if (PS_PP030_Validate("검사02") == false)
                {
                    errCode = "23";
                    throw new Exception();
                }

                if (PS_PP030_Validate("검사03") == false)
                {
                    errCode = "24";
                    throw new Exception();
                }

                oDS_PS_PP030L.RemoveRecord(oDS_PS_PP030L.Size - 1);
                oDS_PS_PP030M.RemoveRecord(oDS_PS_PP030M.Size - 1);
                oMat02.LoadFromDataSource();
                oMat03.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작지구분은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작지번호는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("OrdMgNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지시일자는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("DocDate").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("품목코드는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("ItemCode").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "6")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("지시수,중량이 올바르지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("SelWt").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "7")
                {
                    PSH_Globals.SBO_Application.MessageBox("원재료 사용량을 초과하였습니다. 담당자에게 문의하세요. (" + baseItemWeight + " kg)");
                }
                else if (errCode == "8")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("투입자재라인이 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "9")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업지시등록일은 작번등록일과 같거나 늦어야합니다. 확인하십시오.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "10")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("공정리스트 라인이 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "11")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("투입구분은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat02.Columns.Item("InputGbn").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "12")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("품목은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat02.Columns.Item("ItemCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "13")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("품목그룹은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "14")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당작지는 투입자재 한품목만 입력가능합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "15")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("조달방식은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat02.Columns.Item("ProcType").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "16")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("수,중량은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat02.Columns.Item("Weight").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "17")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(i + "행의 원재료 청구가 중복되어 재청구사유를 필수로 입력하여야 합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat02.Columns.Item("RCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "18")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("수,중량은 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "19")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("배치번호는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "20")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("공정대분류는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat03.Columns.Item("CpBCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "21")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("공정중분류는 필수입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat03.Columns.Item("CpCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "22")
                {
                    //PS_PP030_Validate에서 에러 출력
                }
                else if (errCode == "23")
                {
                    //PS_PP030_Validate에서 에러 출력
                }
                else if (errCode == "24")
                {
                    //PS_PP030_Validate에서 에러 출력
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {

            }

            return returnValue;
        }

        /// <summary>
        /// 선행프로세스와 일자 비교
        /// </summary>
        /// <returns>true:선행프로세스보다 일자가 같거나 느릴 경우, false:선행프로세스보다 일자가 빠를 경우</returns>
        private bool PS_PP030_CheckDate()
        {
            bool returnValue = false;
            string query01;
            string baseEntry;
            string baseLine;
            string docType;
            string CurDocDate;

            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                baseEntry = oForm.Items.Item("BaseNum").Specific.Value.ToString().Trim();
                baseLine = "";
                docType = "PS_PP030";
                CurDocDate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();

                query01 = "EXEC PS_Z_CHECK_DATE '";
                query01 += baseEntry + "','";
                query01 += baseLine + "','";
                query01 += docType + "','";
                query01 += CurDocDate + "'";

                oRecordSet01.DoQuery(query01);

                if (oRecordSet01.Fields.Item("ReturnValue").Value != "False")
                { 
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 중복청구 여부 조회
        /// </summary>
        /// <param name="pDocEntry">문서번호</param>
        /// <param name="pItemCode">원재료품목코드</param>
        /// <param name="pLineID">라인번호</param>
        /// <returns></returns>
        private bool PS_PP030_Check_DupReq(string pDocEntry, string pItemCode, string pLineID)
        {
            bool returnValue = false;
            string query01;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                query01 = "EXEC PS_Z_Check_DupReq '";
                query01 += pDocEntry + "','";
                query01 += pItemCode + "','";
                query01 += pLineID + "'";

                oRecordSet01.DoQuery(query01);

                if (oRecordSet01.Fields.Item("ReturnValue").Value != "FALSE")
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP030_Validate(string ValidateType)
        {
            bool returnValue = false;
            int i;
            int j;
            string query01;
            bool Exist;
            string errCode = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                {
                    errCode = "취소";
                    throw new Exception();
                }
                
                if (ValidateType == "검사01")
                {   
                    //투입자재 매트릭스에 대한 검사
                }
                else if (ValidateType == "검사02")
                {
                    //삭제된 행을 찾아서 삭제가능성 검사
                    query01 = "SELECT PS_PP030L.DocEntry,PS_PP030L.LineId,PS_PP030L.U_ProcType FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030L] PS_PP030L ON PS_PP030H.DocEntry = PS_PP030L.DocEntry WHERE PS_PP030H.Canceled = 'N' AND PS_PP030L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    RecordSet01.DoQuery(query01);
                    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        Exist = false;
                        for (j = 1; j <= oMat02.RowCount - 1; j++)
                        {
                            if (string.IsNullOrEmpty(oMat02.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                            {
                                //새로추가된 행인경우 검사 불필요
                            }
                            else
                            {
                                //라인번호가 같고, 문서번호가 같으면 존재하는행
                                if (Convert.ToInt32(RecordSet01.Fields.Item(0).Value) == Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value) && Convert.ToInt32(RecordSet01.Fields.Item(1).Value) == Convert.ToInt32(oMat02.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                                {
                                    Exist = true;
                                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106") //몰드,기계공구
                                    {
                                        //DB상에는 청구이고 매트릭스의 조달방법이 잔재로 변경된경우 수정할수 없다.
                                        if (RecordSet01.Fields.Item(2).Value == "10" && oMat02.Columns.Item("ProcType").Cells.Item(j).Specific.Selected.Value != "10")
                                        {
                                            errCode = "구매요청";
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                        }
                        
                        if (Exist == false) //삭제된 행중 구매요청에 아직 존재하면 수정불가
                        {
                            
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106") //몰드,기계공구
                            {
                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + RecordSet01.Fields.Item(0).Value + "' AND U_PP030LNo = '" + RecordSet01.Fields.Item(1).Value + "'", 0, 1)) > 0)
                                {
                                    errCode = "구매요청";
                                    throw new Exception();
                                }
                            }
                            
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107") //삭제된 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정불가
                            {
                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + RecordSet01.Fields.Item(0).Value + "'", 0, 1)) > 0)
                                {
                                    errCode = "작업일보";
                                    throw new Exception();
                                }
                            }
                            //휘팅,부품은 삭제 체크 불필요
                        }
                        RecordSet01.MoveNext();
                    }

                    for (i = 1; i <= oMat02.RowCount - 1; i++)
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("LineId").Cells.Item(i).Specific.Value))
                        {
                            //새로추가된 행인경우 검사불필요
                        }
                        else
                        {
                            //기존에 있던 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정불가
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107") 
                            {
                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                                {
                                    query01 = "  SELECT     PS_PP030L.U_ItemCode,";
                                    query01 += "            PS_PP030L.U_ItemName,";
                                    query01 += "            PS_PP030L.U_ItemGpCd,";
                                    query01 += "            PS_PP030L.U_Weight,";
                                    query01 += "            PS_PP030H.U_BPLId,";
                                    query01 += "            CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),";
                                    query01 += "            PS_PP030L.U_CntcCode,";
                                    query01 += "            PS_PP030L.U_CntcName,";
                                    query01 += "            PS_PP030L.U_ProcType,";
                                    query01 += "            PS_PP030L.U_Comments";
                                    query01 += " FROM       [@PS_PP030H] PS_PP030H";
                                    query01 += "            LEFT JOIN";
                                    query01 += "            [@PS_PP030L] PS_PP030L";
                                    query01 += "                ON PS_PP030H.DocEntry = PS_PP030L.DocEntry";
                                    query01 += " WHERE      PS_PP030H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                                    query01 += "            AND PS_PP030L.LineId = '" + oMat02.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                    query01 += "            AND PS_PP030H.Canceled = 'N'";

                                    RecordSet01.DoQuery(query01);

                                    if (RecordSet01.Fields.Item(0).Value == oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value 
                                     && RecordSet01.Fields.Item(1).Value == oMat02.Columns.Item("ItemName").Cells.Item(i).Specific.Value 
                                     && RecordSet01.Fields.Item(2).Value == oMat02.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Selected.Value 
                                     && Convert.ToDouble(RecordSet01.Fields.Item(3).Value) == Convert.ToDouble(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.Value) 
                                     && RecordSet01.Fields.Item(4).Value == oForm.Items.Item("BPLId").Specific.Selected.Value 
                                     && RecordSet01.Fields.Item(5).Value == oMat02.Columns.Item("DueDate").Cells.Item(i).Specific.Value 
                                     && RecordSet01.Fields.Item(6).Value == oMat02.Columns.Item("CntcCode").Cells.Item(i).Specific.Value 
                                     && RecordSet01.Fields.Item(7).Value == oMat02.Columns.Item("CntcName").Cells.Item(i).Specific.Value 
                                     && RecordSet01.Fields.Item(8).Value == oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.Value 
                                     && RecordSet01.Fields.Item(9).Value == oMat02.Columns.Item("Comments").Cells.Item(i).Specific.Value)
                                    {
                                        //값이 변경된 행의경우
                                    }
                                    else
                                    {
                                        errCode = "작업일보";
                                        throw new Exception();
                                    }
                                }
                            }
                            
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106") //몰드,기계공구
                            {
                                if (oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.Value == "20") //잔재인 행은 제외
                                {
                                    //취소인 행은 제외
                                }
                                else if (oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.Value == "30")
                                {
                                    //청구인행에 대해
                                }
                                else
                                {
                                    if (dataHelpClass.GetValue("SELECT U_OKYN FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_PP030LNo = '" + oMat02.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString().Trim() + "'", 0, 1) == "Y")
                                    {
                                        //결재가 완료된 값중
                                        query01 = "  SELECT     PS_PP030L.U_ItemCode,";
                                        query01 += "            PS_PP030L.U_ItemName,";
                                        query01 += "            PS_PP030L.U_ItemGpCd,";
                                        query01 += "            Round(PS_PP030L.U_Weight,2),";
                                        query01 += "            PS_PP030H.U_BPLId,";
                                        query01 += "            CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),";
                                        query01 += "            PS_PP030L.U_CntcCode,";
                                        query01 += "            PS_PP030L.U_CntcName,";
                                        query01 += "            PS_PP030L.U_ProcType,";
                                        query01 += "            PS_PP030L.U_Comments";
                                        query01 += " FROM       [@PS_PP030H] PS_PP030H";
                                        query01 += "            LEFT JOIN";
                                        query01 += "            [@PS_PP030L] PS_PP030L";
                                        query01 += "                ON PS_PP030H.DocEntry = PS_PP030L.DocEntry";
                                        query01 += " WHERE      PS_PP030H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                                        query01 += "            AND PS_PP030L.LineId = '" + oMat02.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                        query01 += "            AND PS_PP030H.Canceled = 'N'";

                                        RecordSet01.DoQuery(query01);

                                        if (RecordSet01.Fields.Item(0).Value == oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(1).Value == oMat02.Columns.Item("ItemName").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(2).Value == oMat02.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Selected.Value 
                                         && Convert.ToDouble(RecordSet01.Fields.Item(3).Value) == Convert.ToDouble(oMat02.Columns.Item("Weight").Cells.Item(i).Specific.Value) 
                                         && RecordSet01.Fields.Item(4).Value == oForm.Items.Item("BPLId").Specific.Selected.Value 
                                         && RecordSet01.Fields.Item(5).Value == oMat02.Columns.Item("DueDate").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(6).Value == oMat02.Columns.Item("CntcCode").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(7).Value == oMat02.Columns.Item("CntcName").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(8).Value == oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.Value 
                                         && RecordSet01.Fields.Item(9).Value == oMat02.Columns.Item("Comments").Cells.Item(i).Specific.Value)
                                        {
                                            //값이 변경된 행의경우
                                        }
                                        else
                                        {
                                            errCode = "구매요청";
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ValidateType == "검사03") //공정 매트릭스에 대한 검사
                {
                    //삭제된 행을 찾아서 삭제가능성 검사
                    query01 = "  SELECT     PS_PP030M.DocEntry,";
                    query01 += "            PS_PP030M.LineId,";
                    query01 += "            PS_PP030M.U_Sequence,";
                    query01 += "            PS_PP030M.U_WorkGbn";
                    query01 += " FROM       [@PS_PP030H] PS_PP030H";
                    query01 += "            LEFT JOIN";
                    query01 += "            [@PS_PP030M] PS_PP030M";
                    query01 += "                ON PS_PP030H.DocEntry = PS_PP030M.DocEntry";
                    query01 += " WHERE      PS_PP030M.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                    RecordSet01.DoQuery(query01);
                    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        Exist = false;
                        for (j = 1; j <= oMat03.RowCount - 1; j++)
                        {
                            if (string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                            {
                                //새로추가된 행인경우 검사불필요
                            }
                            else
                            {
                                //라인번호가 같고, 문서번호가 같으면 존재하는행,시퀀스도 같아야 한다. 행을 삭제할경우 시퀀스가 변경될수 있기때문에.
                                if (Convert.ToInt32(RecordSet01.Fields.Item(0).Value) == Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value) 
                                 && Convert.ToInt32(RecordSet01.Fields.Item(1).Value) == Convert.ToInt32(oMat03.Columns.Item("LineId").Cells.Item(j).Specific.Value) 
                                 && Convert.ToInt32(RecordSet01.Fields.Item(2).Value) == Convert.ToInt32(oMat03.Columns.Item("Sequence").Cells.Item(j).Specific.Value))
                                {
                                    Exist = true;
                                }
                            }
                        }
                        
                        if (Exist == false) //삭제된 행중 작업일보에 등록된행
                        {
                            //삭제된행중에 외주반출등록된행
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101")
                            {
                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + RecordSet01.Fields.Item(0).Value + "' AND PS_MM130L.U_PP030MNo = '" + RecordSet01.Fields.Item(1).Value + "'", 0, 1)) > 0)
                                {
                                    errCode = "외주반출";
                                    throw new Exception();
                                }
                            }

                            //삭제된행중에 외주등록된행
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105") //기계공구
                            {
                                if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE PS_MM005H.U_OrdType in ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + RecordSet01.Fields.Item(0).Value + "-" + RecordSet01.Fields.Item(1).Value + "'", 0, 1)) > 0)
                                {
                                    errCode = "외주청구";
                                    throw new Exception();
                                }
                            }
                        }
                        RecordSet01.MoveNext();
                    }

                    for (i = 1; i <= oMat03.RowCount - 1; i++)
                    {
                        
                        if (string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(i).Specific.Value))
                        { 
                            //새로추가된 행인경우 검사불필요
                        }
                        else
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND PS_PP040L.U_PP030MNo = '" + oMat03.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                //작업일보등록된문서중에 수정이 된문서를 구함
                                query01 = "  SELECT     PS_PP030M.U_CpBCode,";
                                query01 += "            PS_PP030M.U_CpCode,";
                                query01 += "            PS_PP030M.U_ResultYN,";
                                query01 += "            PS_PP030M.U_ReportYN";
                                query01 += " FROM       [@PS_PP030H] PS_PP030H";
                                query01 += "            LEFT JOIN";
                                query01 += "            [@PS_PP030M] PS_PP030M";
                                query01 += "                ON PS_PP030H.DocEntry = PS_PP030M.DocEntry";
                                query01 += " WHERE      PS_PP030H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                                query01 += "            AND PS_PP030M.LineId = '" + oMat03.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                query01 += "            AND PS_PP030H.Canceled = 'N'";
                                RecordSet01.DoQuery(query01);
                                
                                if (RecordSet01.Fields.Item(1).Value == "CP40101" || RecordSet01.Fields.Item(1).Value == "CP40102") //CP40101,2 공정코드는 일보,실적 수정가능 배병관대리 요청 20200603
                                {
                                }
                                else
                                {
                                    if (RecordSet01.Fields.Item(0).Value == oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.Value 
                                     && RecordSet01.Fields.Item(1).Value == oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value 
                                     && RecordSet01.Fields.Item(2).Value == oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Selected.Value 
                                     && RecordSet01.Fields.Item(3).Value == oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Selected.Value)
                                    {
                                        //값이 변경된 행의경우
                                    }
                                    else
                                    {
                                        oMat03.SelectRow(i, true, false);
                                        errCode = "작업일보";
                                        throw new Exception();
                                    }
                                }
                            }

                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101")
                            {
                                
                                if (oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.Value == "10") //자가인 행은 제외
                                {
                                }
                                else if (oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.Value == "20") //정밀인 행은 제외
                                {
                                }
                                else //외주
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND PS_MM130L.U_PP030MNo = '" + oMat03.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString() + "'", 0, 1)) > 0)
                                    {
                                        //외주반출등록된문서중에 수정이 된문서를 구함
                                        query01 = "  SELECT     PS_PP030M.U_CpBCode,";
                                        query01 += "            PS_PP030M.U_CpCode,";
                                        query01 += "            PS_PP030M.U_ResultYN,";
                                        query01 += "            PS_PP030M.U_ReportYN,";
                                        query01 += "            PS_PP030M.U_WorkGbn";
                                        query01 += " FROM       [@PS_PP030H] PS_PP030H";
                                        query01 += "            LEFT JOIN";
                                        query01 += "            [@PS_PP030M] PS_PP030M";
                                        query01 += "                ON PS_PP030H.DocEntry = PS_PP030M.DocEntry";
                                        query01 += " WHERE      PS_PP030H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                                        query01 += "            AND PS_PP030M.LineId = '" + oMat03.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString().Trim() + "'";
                                        query01 += " AND PS_PP030H.Canceled = 'N'";
                                        RecordSet01.DoQuery(query01);

                                        if (RecordSet01.Fields.Item(0).Value == oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(1).Value == oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(2).Value == oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Selected.Value 
                                         && RecordSet01.Fields.Item(3).Value == oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Selected.Value 
                                         && RecordSet01.Fields.Item(4).Value == oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.Value)
                                        {
                                            //값이 변경된 행의경우
                                        }
                                        else
                                        {
                                            oMat03.SelectRow(i, true, false);
                                            errCode = "외주반출";
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                            
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105") //기계공구
                            {
                                if (oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.Value == "10")
                                {
                                    //자가인 행은 제외
                                }
                                else if (oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.Value == "20")
                                {
                                    //정밀인 행은 제외
                                }
                                else //외주
                                {
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "-" + oMat03.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                                    {
                                        //외주청구등록된문서중에 수정이 된문서를 구함
                                        query01 = "  SELECT     PS_PP030M.U_CpBCode,";
                                        query01 += "            PS_PP030M.U_CpCode,";
                                        query01 += "            PS_PP030M.U_ResultYN,";
                                        query01 += "            PS_PP030M.U_ReportYN,";
                                        query01 += "            PS_PP030M.U_WorkGbn";
                                        query01 += " FROM       [@PS_PP030H] PS_PP030H";
                                        query01 += "            LEFT JOIN";
                                        query01 += "            [@PS_PP030M] PS_PP030M";
                                        query01 += "                ON PS_PP030H.DocEntry = PS_PP030M.DocEntry";
                                        query01 += " WHERE      PS_PP030H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                                        query01 += " AND PS_PP030M.LineId = '" + oMat03.Columns.Item("LineId").Cells.Item(i).Specific.Value.ToString.Trim() + "'";
                                        query01 += " AND PS_PP030H.Canceled = 'N'";
                                        RecordSet01.DoQuery(query01);
                                        if (RecordSet01.Fields.Item(0).Value == oMat03.Columns.Item("CpBCode").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(1).Value == oMat03.Columns.Item("CpCode").Cells.Item(i).Specific.Value 
                                         && RecordSet01.Fields.Item(2).Value == oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Selected.Value 
                                         && RecordSet01.Fields.Item(3).Value == oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Selected.Value 
                                         && RecordSet01.Fields.Item(4).Value == oMat03.Columns.Item("WorkGbn").Cells.Item(i).Specific.Selected.Value)
                                        {
                                            //값이 변경된 행의경우
                                        }
                                        else
                                        {
                                            oMat03.SelectRow(i, true, false);
                                            errCode = "외주청구";
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ValidateType == "수정02") //모든값의 변경에 대해 수정가능검사
                {
                    if (string.IsNullOrEmpty(oMat02.Columns.Item("LineId").Cells.Item(oMat02Row02).Specific.Value))
                    {
                        //새로추가된 행인경우 수정가능
                    }
                    else
                    {
                        //삭제된 행중에 멀티,엔드베어링중 작업일보에 등록된 행이면 수정불가
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107") //MG
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "작업일보";
                                throw new Exception();
                            }
                        }
                        else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106") //몰드,기계공구
                        {
                            if (dataHelpClass.GetValue("SELECT U_OKYN FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_PP030LNo = '" + oMat02.Columns.Item("LineId").Cells.Item(oMat02Row02).Specific.Value.ToString().Trim() + "'", 0, 1) == "Y")
                            {
                                errCode = "구매요청";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "행삭제02") //행삭제전 행삭제가능여부검사
                {   
                    if (string.IsNullOrEmpty(oMat02.Columns.Item("LineId").Cells.Item(oMat02Row02).Specific.Value))
                    {
                        //새로추가된 행인경우 삭제가능
                    }
                    else
                    {
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "작업일보";
                                throw new Exception();
                            }
                        }
                        else if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106") //몰드,기계공구
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND U_PP030LNo = '" + oMat02.Columns.Item("LineId").Cells.Item(oMat02Row02).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "구매요청";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "수정03") //모든값의 변경에 대해 수정가능검사
                {
                    if (string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value))
                    {
                        //새로추가된 행인경우 수정 가능
                    }
                    else
                    {
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value != "102")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND PS_PP040L.U_PP030MNo = '" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "작업일보";
                                throw new Exception();
                            }
                        }

                        //삭제된행중에 외주반출등록된행                        
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND PS_MM130L.U_PP030MNo = '" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "외주반출";
                                throw new Exception();
                            }
                        }
                        
                        //삭제된행중에 외주청구등록된행
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105") //기계공구일때
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "-" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "외주청구";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "행삭제03") //행삭제전 행삭제가능여부검사
                {
                    if (string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim()))
                    {
                        //새로추가된 행인경우, 삭제가능
                    }
                    else
                    {
                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND PS_PP040L.U_PP030MNo = '" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                        {
                            errCode = "작업일보";
                            throw new Exception();
                        }

                        //삭제된행중에 외주반출등록된행
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND PS_MM130L.U_PP030MNo = '" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "외주반출";
                                throw new Exception();
                            }
                        }

                        //삭제된행중에 외주청구등록된행
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105") 
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "-" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "외주청구";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "행추가03") //행추가전 행추가가능여부검사
                {   
                    if (string.IsNullOrEmpty(oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim()))
                    {
                        //새로추가된 행인경우 삭제 가능
                    }
                    else
                    {
                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND PS_PP040L.U_PP030MNo = '" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                        {
                            errCode = "작업일보";
                            throw new Exception();
                        }

                        //삭제된행중에 외주반출등록된행
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101")
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "' AND PS_MM130L.U_PP030MNo = '" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "외주반출";
                                throw new Exception();
                            }
                        }

                        //삭제된행중에 외주청구등록된행
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105") //기계공구일때
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40) AND PS_MM005H.Canceled = 'N' AND PS_MM005H.U_PP030DL = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "-" + oMat03.Columns.Item("LineId").Cells.Item(oMat03Row03).Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                            {
                                errCode = "외주청구";
                                throw new Exception();
                            }
                        }
                    }
                }
                else if (ValidateType == "취소") //취소가능유무검사
                {
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                    {
                        errCode = "취소";
                        throw new Exception();
                    }
                    
                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106") //몰드,기계공구
                    {
                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] WHERE U_OrdType = '10' AND Canceled = 'N' AND U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND U_OKYN = 'Y'", 0, 1)) > 0)
                        {
                            errCode = "구매요청";
                            throw new Exception();
                        }
                    }

                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040H.Canceled = 'N' AND PS_PP040L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1)) > 0)
                    {
                        errCode = "";
                        throw new Exception();
                    }

                    //삭제된행중에 외주반출등록된행
                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101")
                    {
                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM130H] PS_MM130H LEFT JOIN [@PS_MM130L] PS_MM130L ON PS_MM130H.DocEntry = PS_MM130L.DocEntry WHERE PS_MM130H.Canceled = 'N' AND PS_MM130L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                        {
                            errCode = "외주반출";
                            throw new Exception();
                        }
                    }

                    //삭제된행중에 외주청구등록된행
                    //기계공구일때
                    if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105")
                    {
                        if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_MM005H] PS_MM005H WHERE U_OrdType IN ('30','40') AND PS_MM005H.Canceled = 'N' AND U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'", 0, 1)) > 0)
                        {
                            errCode = "외주청구";
                            throw new Exception();
                        }
                    }
                }
                else if (ValidateType == "닫기")
                {
                    //닫기가능유무검사
                    if (dataHelpClass.GetValue("SELECT Status FROM [@PS_PP030H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "C")
                    {
                        errCode = "닫기";
                        throw new Exception();
                    }

                    //재고가 존재하면 닫기(종료) 불가 기능 추가(2012.01.11 송명규 추가)
                    query01 = "  SELECT     ISNULL(SUM(A.InQty) - SUM(A.OutQty), 0) AS [StockQty]";
                    query01 += " FROM       OINM AS A";
                    query01 += "            INNER JOIN";
                    query01 += "            OITM As B";
                    query01 += "                ON A.ItemCode = B.ItemCode";
                    query01 += " WHERE      B.U_ItmBsort IN ('105','106')";
                    query01 += "            AND A.ItemCode = '" + oForm.Items.Item("ItemCode").Specific.Value + "'";
                    query01 += " GROUP BY   A.ItemCode";

                    if (Convert.ToDouble(dataHelpClass.GetValue(query01, 0, 1)) > 0)
                    {
                        errCode = "재고";
                        throw new Exception();
                    }
                }

                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errCode == "구매요청")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("구매요청 등록된 행입니다. 처리할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "외주반출")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("외주반출 등록된 행입니다. 처리할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "외주청구")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("외주청구 등록된 행입니다. 처리할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "작업일보")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업일보 등록된 행입니다. 처리할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "취소")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이미 취소된 문서입니다. 처리할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "닫기")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("이미 닫기(종료)된 문서입니다. 처리할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "재고")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("재고가 존재하는 문서입니다. 처리할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
            }

            return returnValue;
        }

        /// <summary>
        /// FormResize
        /// </summary>
        private void PS_PP030_FormResize()
        {
            try
            {
                //생산요청, 작지리스트(Option)
                oForm.Items.Item("Opt01").Left = 10;

                //생산요청, 작지리스트(Matrix)
                oForm.Items.Item("Mat01").Top = 58;
                oForm.Items.Item("Mat01").Height = oForm.Height / 2 - 120;
                oForm.Items.Item("Mat01").Left = oForm.Items.Item("Opt01").Left;
                oForm.Items.Item("Mat01").Width = oForm.Width - 30;

                //작업구분(Label)
                oForm.Items.Item("9").Top = oForm.Items.Item("Mat01").Height + oForm.Items.Item("Mat01").Top + 5;
                oForm.Items.Item("9").Left = oForm.Items.Item("Opt01").Left;
                //작업구분(TextBox)
                oForm.Items.Item("OrdGbn").Top = oForm.Items.Item("9").Top;
                oForm.Items.Item("OrdGbn").Left = oForm.Items.Item("9").Left + oForm.Items.Item("9").Width;
                oForm.Items.Item("BasicGub").Top = oForm.Items.Item("9").Top;
                oForm.Items.Item("BasicGub").Left = oForm.Items.Item("9").Left + oForm.Items.Item("9").Width + oForm.Items.Item("9").Width;
                //제품코드(Label)
                oForm.Items.Item("17").Top = oForm.Items.Item("9").Top + oForm.Items.Item("9").Height + 1;
                oForm.Items.Item("17").Left = oForm.Items.Item("9").Left;
                //제품코드(Link)
                oForm.Items.Item("1000001").Top = oForm.Items.Item("17").Top + 1;
                oForm.Items.Item("1000001").Left = oForm.Items.Item("17").Left + oForm.Items.Item("17").Width - 15;

                //제품코드(TextBox)
                oForm.Items.Item("ItemCode").Top = oForm.Items.Item("17").Top;
                oForm.Items.Item("ItemCode").Left = oForm.Items.Item("OrdGbn").Left;
                //제품명(TextBox)
                oForm.Items.Item("ItemName").Top = oForm.Items.Item("ItemCode").Top;
                oForm.Items.Item("ItemName").Left = oForm.Items.Item("ItemCode").Left + oForm.Items.Item("ItemCode").Width;

                //기준일자(Label)
                oForm.Items.Item("14").Top = oForm.Items.Item("17").Top + oForm.Items.Item("17").Height + 1;
                oForm.Items.Item("14").Left = oForm.Items.Item("17").Left;
                //기준일자(TextBox)
                oForm.Items.Item("OrdMgNum").Top = oForm.Items.Item("14").Top;
                oForm.Items.Item("OrdMgNum").Left = oForm.Items.Item("ItemCode").Left;

                //작업지시번호(Label)
                oForm.Items.Item("67").Top = oForm.Items.Item("14").Top + oForm.Items.Item("14").Height + 1;
                oForm.Items.Item("67").Left = oForm.Items.Item("14").Left;
                //작업지시번호(TextBox)
                oForm.Items.Item("OrdNum").Top = oForm.Items.Item("67").Top;
                oForm.Items.Item("OrdNum").Left = oForm.Items.Item("OrdMgNum").Left;
                //작업지시번호(Sub)(TextBox)
                oForm.Items.Item("OrdSub1").Top = oForm.Items.Item("67").Top;
                oForm.Items.Item("OrdSub1").Left = oForm.Items.Item("OrdNum").Left + oForm.Items.Item("OrdNum").Width;
                oForm.Items.Item("OrdSub2").Top = oForm.Items.Item("67").Top;
                oForm.Items.Item("OrdSub2").Left = oForm.Items.Item("OrdSub1").Left + oForm.Items.Item("OrdSub1").Width;

                //지시,완료일자(Label)
                oForm.Items.Item("18").Top = oForm.Items.Item("67").Top + oForm.Items.Item("67").Height + 1;
                oForm.Items.Item("18").Left = oForm.Items.Item("67").Left;
                //지시일자(TextBox)
                oForm.Items.Item("DocDate").Top = oForm.Items.Item("18").Top;
                oForm.Items.Item("DocDate").Left = oForm.Items.Item("OrdNum").Left;
                //완료일자(TextBox)
                oForm.Items.Item("DueDate").Top = oForm.Items.Item("18").Top;
                oForm.Items.Item("DueDate").Left = oForm.Items.Item("DocDate").Left + oForm.Items.Item("DocDate").Width;

                //담당자(Label)
                oForm.Items.Item("15").Top = oForm.Items.Item("18").Top + oForm.Items.Item("18").Height + 1;
                oForm.Items.Item("15").Left = oForm.Items.Item("18").Left;
                //담당자(TextBox)
                oForm.Items.Item("CntcCode").Top = oForm.Items.Item("15").Top;
                oForm.Items.Item("CntcCode").Left = oForm.Items.Item("DocDate").Left;
                //담당자명(TextBox)
                oForm.Items.Item("CntcName").Top = oForm.Items.Item("15").Top;
                oForm.Items.Item("CntcName").Left = oForm.Items.Item("CntcCode").Left + oForm.Items.Item("CntcCode").Width;

                //수주번호(Label)
                oForm.Items.Item("13").Top = oForm.Items.Item("15").Top + oForm.Items.Item("15").Height + 1;
                oForm.Items.Item("13").Left = oForm.Items.Item("15").Left;
                //수주번호(TextBox)
                oForm.Items.Item("SjNum").Top = oForm.Items.Item("13").Top;
                oForm.Items.Item("SjNum").Left = oForm.Items.Item("CntcCode").Left;
                //수주라인(TextBox)
                oForm.Items.Item("SjLine").Top = oForm.Items.Item("13").Top;
                oForm.Items.Item("SjLine").Left = oForm.Items.Item("SjNum").Left + oForm.Items.Item("SjNum").Width;

                //수주LOT번호(Label)
                oForm.Items.Item("39").Top = oForm.Items.Item("13").Top + oForm.Items.Item("13").Height + 1;
                oForm.Items.Item("39").Left = oForm.Items.Item("13").Left;
                //수주LOT번호(TextBox)
                oForm.Items.Item("LotNo").Top = oForm.Items.Item("39").Top;
                oForm.Items.Item("LotNo").Left = oForm.Items.Item("SjNum").Left;

                //멀티작업구분(Label)
                oForm.Items.Item("1000005").Top = oForm.Items.Item("39").Top + oForm.Items.Item("39").Height + 1;
                oForm.Items.Item("1000005").Left = oForm.Items.Item("39").Left;
                //멀티작업구분1(TextBox)
                oForm.Items.Item("MulGbn1").Top = oForm.Items.Item("1000005").Top;
                oForm.Items.Item("MulGbn1").Left = oForm.Items.Item("LotNo").Left;
                //멀티작업구분2(TextBox)
                oForm.Items.Item("MulGbn2").Top = oForm.Items.Item("1000005").Top;
                oForm.Items.Item("MulGbn2").Left = oForm.Items.Item("MulGbn1").Left + oForm.Items.Item("MulGbn1").Width;
                //멀티작업구분3(TextBox)
                oForm.Items.Item("MulGbn3").Top = oForm.Items.Item("1000005").Top;
                oForm.Items.Item("MulGbn3").Left = oForm.Items.Item("MulGbn2").Left + oForm.Items.Item("MulGbn2").Width;

                //기준문서구분(Label)
                oForm.Items.Item("63").Top = oForm.Items.Item("1000005").Top + oForm.Items.Item("1000005").Height + 1;
                oForm.Items.Item("63").Left = oForm.Items.Item("1000005").Left;
                //기준문서구분(TextBox)
                oForm.Items.Item("BaseType").Top = oForm.Items.Item("63").Top;
                oForm.Items.Item("BaseType").Left = oForm.Items.Item("MulGbn1").Left;

                //기준문서번호(Label)
                oForm.Items.Item("65").Top = oForm.Items.Item("63").Top;
                oForm.Items.Item("65").Left = oForm.Items.Item("BaseType").Left + oForm.Items.Item("BaseType").Width;
                //기준문서번호(TextBox)
                oForm.Items.Item("BaseNum").Top = oForm.Items.Item("65").Top;
                oForm.Items.Item("BaseNum").Left = oForm.Items.Item("65").Left + oForm.Items.Item("65").Width;

                //투입자재(Option)
                oForm.Items.Item("Opt02").Top = oForm.Items.Item("63").Top + oForm.Items.Item("63").Height + 15;
                oForm.Items.Item("Opt02").Left = oForm.Items.Item("63").Left;

                //투입자재(Matrix)
                oForm.Items.Item("Mat02").Top = oForm.Items.Item("Opt02").Top + oForm.Items.Item("Opt02").Height + 1;
                oForm.Items.Item("Mat02").Left = oForm.Items.Item("63").Left;
                oForm.Items.Item("Mat02").Width = oForm.Width / 2 - 25;
                oForm.Items.Item("Mat02").Height = oForm.Height - oForm.Items.Item("Mat02").Top - 60;

                //문서번호(Label)
                oForm.Items.Item("11").Top = oForm.Items.Item("9").Top;
                oForm.Items.Item("11").Left = 320;
                //문서번호(TextBox)
                oForm.Items.Item("DocEntry").Top = oForm.Items.Item("9").Top;
                oForm.Items.Item("DocEntry").Left = oForm.Items.Item("11").Left + oForm.Items.Item("11").Width;

                //사업장(Label)
                oForm.Items.Item("1000002").Top = oForm.Items.Item("14").Top;
                oForm.Items.Item("1000002").Left = 255;
                //사업장(TextBox)
                oForm.Items.Item("BPLId").Top = oForm.Items.Item("14").Top;
                oForm.Items.Item("BPLId").Left = 335;

                //작번이름(Label)
                oForm.Items.Item("70").Top = oForm.Items.Item("1000002").Top + oForm.Items.Item("1000002").Height + 1;
                oForm.Items.Item("70").Left = oForm.Items.Item("11").Left;
                //작번이름(TextBox)
                oForm.Items.Item("JakMyung").Top = oForm.Items.Item("70").Top;
                oForm.Items.Item("JakMyung").Left = oForm.Items.Item("70").Left + oForm.Items.Item("70").Width;

                //작번규격,단위(Label)
                oForm.Items.Item("72").Top = oForm.Items.Item("70").Top + oForm.Items.Item("70").Height + 1;
                oForm.Items.Item("72").Left = oForm.Items.Item("70").Left;
                //작번규격(TextBox)
                oForm.Items.Item("JakSize").Top = oForm.Items.Item("72").Top;
                oForm.Items.Item("JakSize").Left = oForm.Items.Item("72").Left + oForm.Items.Item("72").Width;
                //작번단위(TextBox)
                oForm.Items.Item("JakUnit").Top = oForm.Items.Item("72").Top;
                oForm.Items.Item("JakUnit").Left = oForm.Items.Item("JakSize").Left + oForm.Items.Item("JakSize").Width;

                //요청수,중량(Label)
                oForm.Items.Item("42").Top = oForm.Items.Item("72").Top + oForm.Items.Item("72").Height + 1;
                oForm.Items.Item("42").Left = oForm.Items.Item("72").Left;
                //요청수, 중량
                oForm.Items.Item("ReqWt").Top = oForm.Items.Item("42").Top;
                oForm.Items.Item("ReqWt").Left = oForm.Items.Item("42").Left + oForm.Items.Item("42").Width;

                //지시수,중량(Label)
                oForm.Items.Item("40").Top = oForm.Items.Item("42").Top + oForm.Items.Item("42").Height + 1;
                oForm.Items.Item("40").Left = oForm.Items.Item("42").Left;
                //지시수,중량
                oForm.Items.Item("SelWt").Top = oForm.Items.Item("40").Top;
                oForm.Items.Item("SelWt").Left = oForm.Items.Item("40").Left + oForm.Items.Item("40").Width;

                //수주금액(Label)
                oForm.Items.Item("38").Top = oForm.Items.Item("40").Top + oForm.Items.Item("40").Height + 1;
                oForm.Items.Item("38").Left = oForm.Items.Item("40").Left;
                //수주금액
                oForm.Items.Item("SjPrice").Top = oForm.Items.Item("38").Top;
                oForm.Items.Item("SjPrice").Left = oForm.Items.Item("38").Left + oForm.Items.Item("38").Width;

                //문서상태(Label)
                oForm.Items.Item("79").Top = oForm.Items.Item("38").Top + oForm.Items.Item("38").Height + 1;
                oForm.Items.Item("79").Left = oForm.Items.Item("38").Left + 65;
                //문서상태(TextBox)
                oForm.Items.Item("Status").Top = oForm.Items.Item("79").Top;
                oForm.Items.Item("Status").Left = oForm.Items.Item("79").Left + oForm.Items.Item("79").Width;

                //취소여부(Label)
                oForm.Items.Item("71").Top = oForm.Items.Item("79").Top + oForm.Items.Item("79").Height + 1;
                oForm.Items.Item("71").Left = oForm.Items.Item("79").Left;
                //취소여부(TextBox)
                oForm.Items.Item("Canceled").Top = oForm.Items.Item("71").Top;
                oForm.Items.Item("Canceled").Left = oForm.Items.Item("71").Left + oForm.Items.Item("71").Width;

                //공정리스트(Option)
                oForm.Items.Item("Opt03").Top = oForm.Items.Item("9").Top;
                oForm.Items.Item("Opt03").Left = oForm.Width / 2;

                //표준공수조회(BUTTON)
                oForm.Items.Item("btnWkSrch").Top = oForm.Items.Item("9").Top - 2;
                oForm.Items.Item("btnWkSrch").Left = oForm.Items.Item("Opt03").Left + oForm.Items.Item("Opt03").Width + 3;

                //품목별공수조회(BUTTON)
                oForm.Items.Item("btnItmSrch").Top = oForm.Items.Item("btnWkSrch").Top;
                oForm.Items.Item("btnItmSrch").Left = oForm.Items.Item("btnWkSrch").Left + oForm.Items.Item("btnWkSrch").Width + 3;

                //공정금액합계(Label)
                oForm.Items.Item("77").Top = oForm.Items.Item("9").Top;
                oForm.Items.Item("77").Left = oForm.Items.Item("btnItmSrch").Left + oForm.Items.Item("btnItmSrch").Width + 5;

                //공정금액합계(TextBox)
                oForm.Items.Item("Total").Top = oForm.Items.Item("9").Top;
                oForm.Items.Item("Total").Left = oForm.Items.Item("77").Left + oForm.Items.Item("77").Width;

                //공정리스트(Matrix)
                oForm.Items.Item("Mat03").Left = oForm.Items.Item("Opt03").Left;
                oForm.Items.Item("Mat03").Top = oForm.Items.Item("9").Top + 18;
                oForm.Items.Item("Mat03").Height = oForm.Height - oForm.Items.Item("Mat03").Top - 60;
                oForm.Items.Item("Mat03").Width = oForm.Width - oForm.Items.Item("Mat03").Left - 20;
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 메트릭스 데이터 로드
        /// </summary>
        private void PS_PP030_MTX01()
        {
            int i;
            string query01;
            string Param01;
            string Param02;
            string Param03;
            string Param04;
            string Param05;
            string Param06;
            string Param07;
            string Param08;
            string errMessage = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbouiCOM.ProgressBar ProgressBar01 = null;

            try
            {
                oForm.Freeze(true);

                ProgressBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                Param01 = oForm.Items.Item("SBPLId").Specific.Selected.Value.ToString().Trim();
                Param02 = oForm.Items.Item("ItmBsort").Specific.Selected.Value.ToString().Trim();
                Param03 = oForm.Items.Item("ItmMsort").Specific.Selected.Value.ToString().Trim();
                Param04 = oForm.Items.Item("ReqType").Specific.Selected.Value.ToString().Trim();
                Param05 = oForm.Items.Item("SItemCod").Specific.Value.ToString().Trim();
                Param06 = oForm.Items.Item("SCardCod").Specific.Value.ToString().Trim();
                Param07 = oForm.Items.Item("Mark").Specific.Value.ToString().Trim();
                Param08 = oForm.Items.Item("ReqCod").Specific.Value.ToString().Trim();
                
                query01 = "EXEC PS_PP030_02 '";
                query01 += Param01 + "','";
                query01 += Param02 + "','";
                query01 += Param03 + "','";
                query01 += Param04 + "','";
                query01 += Param05 + "','";
                query01 += Param06 + "','";
                query01 += Param07 + "','";
                query01 += Param08 + "'";
                RecordSet01.DoQuery(query01);

                oMat01.Clear();
                oMat01.FlushToDataSource();
                oMat01.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_USERDS01.InsertRecord(i);
                    }
                    oDS_PS_USERDS01.Offset = i;
                    oDS_PS_USERDS01.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_USERDS01.SetValue("U_ColReg01", i, RecordSet01.Fields.Item(0).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg02", i, RecordSet01.Fields.Item(1).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg03", i, RecordSet01.Fields.Item(2).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg04", i, RecordSet01.Fields.Item(3).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg05", i, RecordSet01.Fields.Item(4).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg06", i, RecordSet01.Fields.Item(5).Value);
                    oDS_PS_USERDS01.SetValue("U_ColDt01", i, RecordSet01.Fields.Item(6).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg07", i, RecordSet01.Fields.Item(7).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg08", i, RecordSet01.Fields.Item(8).Value);
                    oDS_PS_USERDS01.SetValue("U_ColQty01", i, RecordSet01.Fields.Item(9).Value);
                    oDS_PS_USERDS01.SetValue("U_ColQty02", i, RecordSet01.Fields.Item(10).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg09", i, RecordSet01.Fields.Item(11).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg10", i, RecordSet01.Fields.Item(12).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg11", i, RecordSet01.Fields.Item(13).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg12", i, RecordSet01.Fields.Item(14).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg13", i, RecordSet01.Fields.Item(15).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg14", i, RecordSet01.Fields.Item(16).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg15", i, RecordSet01.Fields.Item(17).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg16", i, RecordSet01.Fields.Item(18).Value);
                    oDS_PS_USERDS01.SetValue("U_ColReg17", i, RecordSet01.Fields.Item(19).Value);
                    RecordSet01.MoveNext();
                    ProgressBar01.Value += 1;
                    ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
                }
                oMat01.LoadFromDataSource();
                oMat01.AutoResizeColumns();
                oForm.Update();
            }
            catch(Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                oForm.Freeze(false);
                if (ProgressBar01 != null)
                {
                    ProgressBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgressBar01);
                }
            }
        }

        /// <summary>
        /// 메트릭스 데이터 로드
        /// </summary>
        private void PS_PP030_MTX02()
        {
            int i;
            string query01;
            string errMessage = string.Empty;

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);
                
                query01 = "  SELECT     PS_PP005H.U_ItemCod2,";
                query01 += "            PS_PP005H.U_ItemNam2,";
                query01 += "            OITM.ItmsGrpCod";
                query01 += " FROM       [@PS_PP005H] PS_PP005H";
                query01 += "            LEFT JOIN";
                query01 += "            [OITM] OITM";
                query01 += "                ON PS_PP005H.U_ItemCod2 = OITM.ItemCode";
                query01 += " WHERE      U_ItemCod1 = '" + oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() + "'";
                RecordSet01.DoQuery(query01);

                oMat02.Clear();
                oMat02.FlushToDataSource();
                oMat02.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    PS_PP030_AddMatrixRow01(0, true);
                    errMessage = "결과가 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                {
                    if (i != 0)
                    {
                        oDS_PS_PP030L.InsertRecord(i);
                    }
                    oDS_PS_PP030L.Offset = i;
                    oDS_PS_PP030L.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                    oDS_PS_PP030L.SetValue("U_InputGbn", i, "10"); //투입구분 : 휘팅,부품의경우만 실행되므로 항상 10이다
                    oDS_PS_PP030L.SetValue("U_ItemCode", i, RecordSet01.Fields.Item(0).Value); //품목코드
                    oDS_PS_PP030L.SetValue("U_ItemName", i, RecordSet01.Fields.Item(1).Value); //품목이름
                    oDS_PS_PP030L.SetValue("U_ItemGpCd", i, RecordSet01.Fields.Item(2).Value); //품목그룹
                    oDS_PS_PP030L.SetValue("U_BatchNum", i, ""); //배치번호
                    oDS_PS_PP030L.SetValue("U_Weight", i, "0"); //중량
                    oDS_PS_PP030L.SetValue("U_DueDate", i, "");
                    oDS_PS_PP030L.SetValue("U_CntcCode", i, "");
                    oDS_PS_PP030L.SetValue("U_CntcName", i, "");
                    oDS_PS_PP030L.SetValue("U_ProcType", i, "20");
                    oDS_PS_PP030L.SetValue("U_Comments", i, "");
                    oDS_PS_PP030L.SetValue("U_LineId", i, "");
                    if (i == RecordSet01.RecordCount - 1)
                    {
                        PS_PP030_AddMatrixRow01(i + 1, false); //마지막행에 한줄추가
                    }
                    RecordSet01.MoveNext();
                }
                oMat02.LoadFromDataSource();
                oMat02.AutoResizeColumns();
                oForm.Update();
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(errMessage, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 메트릭스 데이터 로드
        /// </summary>
        private void PS_PP030_MTX03()
        {
            int i;
            string query01;
            string itemCode;
            string BasicGub;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                oForm.Freeze(true);

                itemCode = oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim();
                BasicGub = oForm.Items.Item("BasicGub").Specific.Value.ToString().Trim();

                query01 = "EXEC [PS_PP030_07] '";
                query01 += itemCode + "','";
                query01 += BasicGub + "'";

                RecordSet01.DoQuery(query01);

                oMat03.Clear();
                oMat03.FlushToDataSource();
                oMat03.LoadFromDataSource();

                if (RecordSet01.RecordCount == 0)
                {
                    if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "105" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "106") 
                    {
                        oForm.Items.Item("Mat03").Enabled = true;
                    } 
                    else
                    {
                        oForm.Items.Item("Mat03").Enabled = false; //휘팅,부품,멀티,엔베는 표준공정이 등록되지 않으면 진행불가능
                    }
                    PS_PP030_AddMatrixRow02(0, true);
                }
                else
                {
                    oForm.Items.Item("Mat03").Enabled = true;
                }

                if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() != "105") 
                {
                    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                    {
                        if (i != 0)
                        {
                            oDS_PS_PP030M.InsertRecord(i);
                        }
                        oDS_PS_PP030M.Offset = i;
                        oDS_PS_PP030M.SetValue("U_LineNum", i, Convert.ToString(i + 1));
                        oDS_PS_PP030M.SetValue("U_Sequence", i, Convert.ToString(i + 1));
                        oDS_PS_PP030M.SetValue("U_CpBCode", i, RecordSet01.Fields.Item(0).Value);
                        oDS_PS_PP030M.SetValue("U_CpBName", i, RecordSet01.Fields.Item(1).Value);
                        oDS_PS_PP030M.SetValue("U_CpCode", i, RecordSet01.Fields.Item(2).Value);
                        oDS_PS_PP030M.SetValue("U_CpName", i, RecordSet01.Fields.Item(3).Value);
                        oDS_PS_PP030M.SetValue("U_Unit", i, RecordSet01.Fields.Item(4).Value);
                        oDS_PS_PP030M.SetValue("U_ReWorkYN", i, "N");
                        oDS_PS_PP030M.SetValue("U_ResultYN", i, RecordSet01.Fields.Item(5).Value);
                        oDS_PS_PP030M.SetValue("U_ReportYN", i, RecordSet01.Fields.Item(6).Value);
                        oDS_PS_PP030M.SetValue("U_WorkGbn", i, "10");
                        if (i == RecordSet01.RecordCount - 1)
                        {
                            PS_PP030_AddMatrixRow02(i + 1, false); //마지막행 추가

                        }
                        RecordSet01.MoveNext();
                    }
                }
                else //기계공구는 검사공정 기본 등록
                {
                    oDS_PS_PP030M.Offset = 0;
                    oDS_PS_PP030M.SetValue("U_LineNum", 0, "1");
                    oDS_PS_PP030M.SetValue("U_Sequence", 0, "1");
                    oDS_PS_PP030M.SetValue("U_CpBCode", 0, "CP204");
                    oDS_PS_PP030M.SetValue("U_CpBName", 0, "검사");
                    oDS_PS_PP030M.SetValue("U_CpCode", 0, "CP20402");
                    oDS_PS_PP030M.SetValue("U_CpName", 0, "최종검사");
                    oDS_PS_PP030M.SetValue("U_Unit", 0, "");
                    oDS_PS_PP030M.SetValue("U_ReWorkYN", 0, "N");
                    oDS_PS_PP030M.SetValue("U_ResultYN", 0, "N");
                    oDS_PS_PP030M.SetValue("U_ReportYN", 0, "N");
                    oDS_PS_PP030M.SetValue("U_WorkGbn", 0, "10");
                    PS_PP030_AddMatrixRow02(1, false); //마지막행 추가
                }
                oMat03.LoadFromDataSource();
                oMat03.AutoResizeColumns();
                oForm.Update();
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// 구매요청 INSERT
        /// (조달방식이 청구이면 [@PS_MM005H]에 추가, 구매요청의 결재(OKYN) 값이 Y로 변경된 경우 수정불가, 작지에서는 청구행에 대해 행삭제불가)
        /// </summary>
        /// <param name="oDocEntry02"></param>
        /// <param name="oLineId02"></param>
        private void PS_PP030_PurchaseRequest(int oDocEntry02, int oLineId02)
        {
            string query01;
            string query02;
            string itemName;
            string DocEntry = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset RecordSet02 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                query01 = "   SELECT '" +   DocEntry + "',";
                query01 += "'" +           DocEntry + "',";
                query01 += "               PS_PP030L.U_ItemCode,";
                query01 += "               PS_PP030L.U_ItemName,";
                query01 += "               PS_PP030L.U_Weight,";
                query01 += "               PS_PP030L.U_Weight,";
                query01 += "               0,";
                query01 += "               0,";
                query01 += "               PS_PP030H.U_BPLId,";
                query01 += "'" +           DocEntry + "',";
                query01 += "               CONVERT(NVARCHAR,GETDATE(),112),";
                query01 += "               CONVERT(NVARCHAR,PS_PP030L.U_DueDate,112),";
                query01 += "               PS_PP030L.U_CntcCode,";
                query01 += "               PS_PP030L.U_CntcName,";
                query01 += "               (SELECT dept FROM [OHEM] WHERE empID = PS_PP030L.U_CntcCode),";
                query01 += "               '',";
                query01 += "               'N',";
                query01 += "               'Y',";
                query01 += "               '10',";
                query01 += "               PS_PP030L.U_Comments,";
                query01 += "               'N',";
                query01 += "               '',";
                query01 += "               '10',";
                query01 += "               '',";
                query01 += "               'O',";
                query01 += "               PS_PP030H.DocEntry,";
                query01 += "               PS_PP030L.LineId,";
                query01 += "               CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030L.LineId),";
                query01 += "               CONVERT(NVARCHAR,PS_PP030L.U_CGDate,112) As CGDate,";
                query01 += "               PS_PP030H.U_OrdNum + '-' + PS_PP030H.U_OrdSub1 + '-' + PS_PP030H.U_OrdSub2 As OrdNum,";
                query01 += "               PS_PP030L.U_ImportYN As ImportYN,"; //수입품여부
                query01 += "               PS_PP030L.U_EmergYN As EmergYN,"; //긴급여부
                query01 += "               PS_PP030L.U_RCode As RCode,"; //재작업사유
                query01 += "               PS_PP030L.U_RName As RName,"; //재작업사유내용
                query01 += "               PS_PP030L.U_PartNo As PartNo"; //PartNo 추가(2020.04.16 송명규, 송채린(생산팀) 요청)
                query01 += " FROM          [@PS_PP030H] PS_PP030H";
                query01 += "               LEFT JOIN";
                query01 += "               [@PS_PP030L] PS_PP030L";
                query01 += "                   ON PS_PP030H.DocEntry = PS_PP030L.DocEntry";
                query01 += " WHERE         PS_PP030H.DocEntry = '" + oDocEntry02 + "'";
                query01 += "               AND PS_PP030L.LineId = '" + oLineId02 + "'";
                query01 += "               AND PS_PP030H.Canceled = 'N'";
                RecordSet01.DoQuery(query01);

                string temp = RecordSet01.Fields.Item(3).Value.ToString().Trim();

                itemName = dataHelpClass.Make_ItemName(RecordSet01.Fields.Item(3).Value.ToString().Trim());

                DocEntry = dataHelpClass.GetValue("SELECT CASE WHEN ISNULL(MAX(CONVERT(INT,DocEntry)),0) = 0 THEN LEFT(CONVERT(NVARCHAR,'" + RecordSet01.Fields.Item("CGDate").Value + "',112),6) + '0001' ELSE ISNULL(MAX(CONVERT(INT,DocEntry)),0)+1 END FROM [@PS_MM005H] WHERE LEFT(CONVERT(NVARCHAR,'" + RecordSet01.Fields.Item("CGDate").Value + "',112),6) = LEFT(DocEntry,6)", 0, 1);

                //구매요청이 취소되면 안되고 삭제되어야 한다. 삭제하면서 작업지시등록의 구매요청행도 삭제, 단 작업지시등록에 행이 1개만 존재한다면 삭제할수 없다.
                query02 = "  SELECT COUNT(*)";
                query02 += " FROM   [@PS_MM005H]";
                query02 += " WHERE  U_OrdType = '10'";
                query02 += "        AND U_PP030HNo = '" + oDocEntry02 + "'";
                query02 += "        AND U_PP030LNo = '" + oLineId02 + "'";
                RecordSet02.DoQuery(query02);

                if (Convert.ToInt32(RecordSet02.Fields.Item(0).Value) == 0)
                {
                    query01 = "INSERT INTO [@PS_MM005H]";
                    query01 += " (";
                    query01 += " DocEntry,";
                    query01 += " DocNum,";
                    query01 += " U_ItemCode,";
                    query01 += " U_ItemName,";
                    query01 += " U_Qty,";
                    query01 += " U_Weight,";
                    query01 += " U_BPLId,";
                    query01 += " U_CgNum,";
                    query01 += " U_DocDate,";
                    query01 += " U_DueDate,";
                    query01 += " U_CntcCode,";
                    query01 += " U_CntcName,";
                    query01 += " U_DeptCode,";
                    query01 += " U_UseDept,";
                    query01 += " U_Auto,";
                    query01 += " U_QCYN,";
                    query01 += " U_OKYN,";
                    query01 += " U_OKDate,";
                    query01 += " U_OrdType,";
                    query01 += " U_ProcCode,";
                    query01 += " U_Status,";
                    query01 += " U_Comments,";
                    query01 += " U_OrdNum,";
                    query01 += " U_PP030HNo,";
                    query01 += " U_PP030LNo,";
                    query01 += " U_PP030DL,";
                    query01 += " U_ImportYN,"; //수입품여부(2018.09.12 송명규, 김석태 과장 요청)
                    query01 += " U_EmergYN,"; //긴급여부(2018.09.12 송명규, 김석태 과장 요청)
                    query01 += " U_RCode,"; //재청구사유(2018.09.17 송명규, 김석태 과장 요청)
                    query01 += " U_RName,"; //재청구사유내용(2018.09.17 송명규, 김석태 과장 요청)
                    query01 += " U_PartNo,"; //PartNo 추가(2020.04.16 송명규, 송채린(생산팀) 요청)
                    query01 += " UserSign,"; //UserSign 추가(2020.04.16 송명규)
                    query01 += " CreateDate"; //생성일 추가(2014.02.24 송명규)
                    query01 += " ) ";
                    query01 += "VALUES(";
                    query01 += "'" + DocEntry + "',";
                    query01 += "'" + DocEntry + "',";
                    query01 += "'" + RecordSet01.Fields.Item(2).Value + "',"; //ItemCode
                    query01 += "'" + itemName + "',"; //ItemName
                    query01 += "" + RecordSet01.Fields.Item(4).Value + ","; //Qty
                    query01 += "" + RecordSet01.Fields.Item(5).Value + ","; //Weight
                    query01 += "'" + RecordSet01.Fields.Item(8).Value + "',"; //BPLId
                    query01 += "'" + DocEntry + "',";
                    query01 += "'" + RecordSet01.Fields.Item("CGDate").Value + "',"; //DocDate
                    query01 += "'" + RecordSet01.Fields.Item(11).Value + "',"; //DueDate
                    query01 += "'" + RecordSet01.Fields.Item(12).Value + "',"; //CntcCode
                    query01 += "'" + RecordSet01.Fields.Item(13).Value + "',"; //CntcName
                    query01 += "'" + RecordSet01.Fields.Item(14).Value + "',"; //DeptCode
                    query01 += "'" + RecordSet01.Fields.Item(15).Value + "',"; //UseDept
                    query01 += "'" + RecordSet01.Fields.Item(16).Value + "',"; //Auto
                    query01 += "'" + RecordSet01.Fields.Item(17).Value + "',"; //QCYN
                    query01 += "'" + RecordSet01.Fields.Item(20).Value + "',"; //OKYN
                    query01 += "'" + RecordSet01.Fields.Item(21).Value + "',"; //OKDate
                    query01 += "'" + RecordSet01.Fields.Item(22).Value + "',"; //OrdType
                    query01 += "'" + RecordSet01.Fields.Item(23).Value + "',"; //ProcCode
                    query01 += "'" + RecordSet01.Fields.Item(24).Value + "',"; //Status
                    query01 += "'" + RecordSet01.Fields.Item(19).Value + "',"; //U_Comments
                    query01 += "'" + RecordSet01.Fields.Item("OrdNum").Value + "',"; //U_OrdNum
                    query01 += "'" + RecordSet01.Fields.Item(25).Value + "',"; //PP030HNo
                    query01 += "'" + RecordSet01.Fields.Item(26).Value + "',"; //PP030LNo
                    query01 += "'" + RecordSet01.Fields.Item(27).Value + "',"; //PP030DL
                    query01 += "'" + RecordSet01.Fields.Item("ImportYN").Value + "',"; //수입품여부(2018.09.12 송명규, 김석태 과장 요청)
                    query01 += "'" + RecordSet01.Fields.Item("EmergYN").Value + "',"; //긴급여부(2018.09.12 송명규, 김석태 과장 요청)
                    query01 += "'" + RecordSet01.Fields.Item("RCode").Value + "',"; //재청구사유(2018.09.17 송명규, 김석태 과장 요청)
                    query01 += "'" + RecordSet01.Fields.Item("RName").Value + "',"; //재청구사유내용(2018.09.17 송명규, 김석태 과장 요청)
                    query01 += "'" + RecordSet01.Fields.Item("PartNo").Value + "',"; //PartNo 추가(2020.04.16 송명규, 송채린(생산팀) 요청)
                    query01 += "'" + PSH_Globals.oCompany.UserSignature + "',"; //UserSign 추가(2020.04.16 송명규)
                    query01 += " GETDATE()"; //생성일 추가(2014.02.24 송명규)
                    query01 += ")";
                    RecordSet01.DoQuery(query01);
                }
                else
                {
                    query01 = "UPDATE [@PS_MM005H] SET";
                    query01 += " U_ItemCode = '" + RecordSet01.Fields.Item(2).Value + "',";
                    query01 += " U_ItemName = '" + itemName + "',";
                    query01 += " U_Qty = " + RecordSet01.Fields.Item(4).Value + ",";
                    query01 += " U_Weight = " + RecordSet01.Fields.Item(5).Value + ",";
                    query01 += " U_BPLId = '" + RecordSet01.Fields.Item(8).Value + "',";
                    query01 += " U_DocDate = '" + RecordSet01.Fields.Item("CGDate").Value + "',";
                    query01 += " U_DueDate = '" + RecordSet01.Fields.Item(11).Value + "',";
                    query01 += " U_CntcCode = '" + RecordSet01.Fields.Item(12).Value + "',";
                    query01 += " U_CntcName = '" + RecordSet01.Fields.Item(13).Value + "',";
                    query01 += " U_DeptCode = '" + RecordSet01.Fields.Item(14).Value + "',";
                    query01 += " U_UseDept = '" + RecordSet01.Fields.Item(15).Value + "',";
                    query01 += " U_Auto = '" + RecordSet01.Fields.Item(16).Value + "',";
                    query01 += " U_QCYN = '" + RecordSet01.Fields.Item(17).Value + "',";
                    query01 += " U_Note = '" + RecordSet01.Fields.Item(19).Value + "',";
                    query01 += " U_OrdType = '" + RecordSet01.Fields.Item(22).Value + "',";
                    query01 += " U_ProcCode = '" + RecordSet01.Fields.Item(23).Value + "',";
                    query01 += " U_Comments = '" + RecordSet01.Fields.Item(19).Value + "',";
                    query01 += " U_ImportYN = '" + RecordSet01.Fields.Item("ImportYN").Value + "',"; //수입품여부(2018.09.12 송명규, 김석태 과장 요청)
                    query01 += " U_EmergYN = '" + RecordSet01.Fields.Item("EmergYN").Value + "',"; //긴급여부(2018.09.12 송명규, 김석태 과장 요청)
                    query01 += " U_RCode = '" + RecordSet01.Fields.Item("RCode").Value + "',"; //재작업사유(2018.09.17 송명규, 김석태 과장 요청)
                    query01 += " U_RName = '" + RecordSet01.Fields.Item("RName").Value + "',"; //재작업사유내용(2018.09.17 송명규, 김석태 과장 요청)
                    query01 += " U_PartNo = '" + RecordSet01.Fields.Item("PartNo").Value + "',"; //PartNo 추가(2020.04.16 송명규, 송채린(생산팀) 요청)
                    query01 += " UserSign = '" + PSH_Globals.oCompany.UserSignature + "',"; //UserSign(2020.04.16 송명규)
                    query01 += " UpdateDate = GETDATE()"; //수정일 추가(2014.02.24 송명규)
                    query01 += " WHERE U_OrdType = '10' And U_PP030HNo = '" + RecordSet01.Fields.Item(25).Value + "'";
                    query01 += " AND U_PP030LNo = '" + RecordSet01.Fields.Item(26).Value + "'";
                    RecordSet01.DoQuery(query01);
                }
            }
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet02);
            }
        }

        /// <summary>
        /// MG 작업지시 일괄 반복 등록
        /// </summary>
        /// <returns></returns>
        private bool PS_PP030_AutoCreateMultiGage()
        {
        	bool returnValue = false;
            int i;
            int j;
        	string query01;
            int CurrentDocEntry;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();
                
                for (i = 1; i <= oMat02.VisualRowCount; i++) //투입자재의 수량만큼
                {
                    query01 = "SELECT AutoKey FROM [ONNM] WHERE ObjectCode = 'PS_PP030'";
                    RecordSet01.DoQuery(query01);
                    CurrentDocEntry = Convert.ToInt32(RecordSet01.Fields.Item(0).Value);
                    query01 = "UPDATE [ONNM] SET AutoKey = AutoKey + 1 WHERE ObjectCode = 'PS_PP030'";
                    RecordSet01.DoQuery(query01);
                    
                    query01 = "INSERT INTO [@PS_PP030H] (";
                    query01 += "DocEntry,";
                    query01 += "DocNum,";
                    query01 += "Period,";
                    query01 += "Instance,";
                    query01 += "Series,";
                    query01 += "Handwrtten,";
                    query01 += "Canceled,";
                    query01 += "Object,";
                    query01 += "LogInst,";
                    query01 += "UserSign,";
                    query01 += "Transfered,";
                    query01 += "Status,";
                    query01 += "CreateDate,";
                    query01 += "CreateTime,";
                    query01 += "UpdateDate,";
                    query01 += "UpdateTime,";
                    query01 += "DataSource,";
                    query01 += "U_BaseType,";
                    query01 += "U_BaseNum,";
                    query01 += "U_OrdGbn,";
                    query01 += "U_DocDate,";
                    query01 += "U_DueDate,";
                    query01 += "U_ItemCode,";
                    query01 += "U_ItemName,";
                    query01 += "U_CntcCode,";
                    query01 += "U_CntcName,";
                    query01 += "U_SjNum,";
                    query01 += "U_SjLine,";
                    query01 += "U_OrdMgNum,";
                    query01 += "U_OrdNum,";
                    query01 += "U_OrdSub1,";
                    query01 += "U_OrdSub2,";
                    query01 += "U_JakMyung,";
                    query01 += "U_ReqWt,";
                    query01 += "U_SelWt,";
                    query01 += "U_LotNo,";
                    query01 += "U_SjPrice,";
                    query01 += "U_MulGbn1,";
                    query01 += "U_MulGbn2,";
                    query01 += "U_MulGbn3,";
                    query01 += "U_Comments,";
                    query01 += "U_BPLId,";
                    query01 += "U_BasicGub";
                    query01 += ")";
                    query01 += " VALUES";
                    query01 += "(";
                    query01 += "'" + CurrentDocEntry + "',";
                    query01 += "'" + CurrentDocEntry + "',";
                    query01 += "'11'" + ",";
                    query01 += "'0'" + ",";
                    query01 += "'-1'" + ",";
                    query01 += "'N'" + ",";
                    query01 += "'N'" + ",";
                    query01 += "'PS_PP030'" + ",";
                    query01 += "NULL" + ",";
                    query01 += "'" + PSH_Globals.oCompany.UserSignature + "',";
                    query01 += "'N'" + ",";
                    query01 += "'O'" + ","; //Status
                    query01 += "CONVERT(NVARCHAR,GETDATE(),112),";
                    query01 += "SUBSTRING(CONVERT(NVARCHAR,GETDATE(),108),1,2) + SUBSTRING(CONVERT(NVARCHAR,GETDATE(),108),4,2),";
                    query01 += "NULL,"; //UpdateDate
                    query01 += "NULL,"; //UpdateTime
                    query01 += "'I',"; //DataSource
                    query01 += "NULL,"; //BaseType
                    query01 += "NULL,"; //BaseNum
                    query01 += "'" + oForm.Items.Item("OrdGbn").Specific.Selected.Value + "',";
                    if (string.IsNullOrEmpty(oForm.Items.Item("DocDate").Specific.Value))
                    {
                        query01 += "NULL,";
                    }
                    else
                    {
                        query01 += "'" + oForm.Items.Item("DocDate").Specific.Value + "',";
                    }
                    if (string.IsNullOrEmpty(oForm.Items.Item("DueDate").Specific.Value))
                    {
                        query01 += "NULL,";
                    }
                    else
                    {
                        query01 += "'" + oForm.Items.Item("DueDate").Specific.Value + "',";
                    }
                    query01 += "'" + oForm.Items.Item("ItemCode").Specific.Value + "',";
                    query01 += "'" + oForm.Items.Item("ItemName").Specific.Value + "',";
                    query01 += "'" + oForm.Items.Item("CntcCode").Specific.Value + "',";
                    query01 += "'" + oForm.Items.Item("CntcName").Specific.Value + "',";
                    query01 += "NULL,"; //SjNum
                    query01 += "NULL,"; //SjLine
                    query01 += "'" + oForm.Items.Item("OrdMgNum").Specific.Value + "',"; //신규작업지시번호를 조회
                    query01 += "'" + oForm.Items.Item("OrdMgNum").Specific.Value + dataHelpClass.GetValue("EXEC PS_PP030_01 '" + oForm.Items.Item("OrdMgNum").Specific.Value + "'", 0, 1) + "',";
                    query01 += "'" + "00',";
                    query01 += "'" + "000',";
                    query01 += "'" + oForm.Items.Item("JakMyung").Specific.Value + "',";
                    query01 += "'" + oForm.Items.Item("ReqWt").Specific.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("Weight").Cells.Item(i).Specific.Value + "',"; //투입자재의 중량으로 입력되어야함
                    query01 += "NULL,"; //LotNo
                    query01 += "'" + oForm.Items.Item("SjPrice").Specific.Value + "',";
                    if (oForm.Items.Item("MulGbn1").Specific.Selected == null)
                    {
                        query01 += "'',";
                    }
                    else
                    {
                        query01 += "'" + oForm.Items.Item("MulGbn1").Specific.Selected.Value.ToString().Trim() + "',";
                    }

                    if (oForm.Items.Item("MulGbn2").Specific.Selected == null)
                    {
                        query01 += "'',";
                    }
                    else
                    {
                        query01 += "'" + oForm.Items.Item("MulGbn2").Specific.Selected.Value.ToString().Trim() + "',";
                    }

                    if (oForm.Items.Item("MulGbn3").Specific.Selected == null)
                    {
                        query01 += "'',";
                    }
                    else
                    {
                        query01 += "'" + oForm.Items.Item("MulGbn3").Specific.Selected.Value.ToString().Trim() + "',";
                    }
                    query01 += "'" + oForm.Items.Item("Comments").Specific.Value.ToString().Trim() + "',";
                    query01 += "'" + oForm.Items.Item("BPLId").Specific.Value.ToString().Trim() + "',";
                    query01 += "'" + oForm.Items.Item("BasicGub").Specific.Selected.Value + "'";
                    query01 += ")";
                    RecordSet01.DoQuery(query01);

                    query01 = "INSERT INTO [@PS_PP030L] (";
                    query01 += "DocEntry,";
                    query01 += "LineId,";
                    query01 += "VisOrder,";
                    query01 += "Object,";
                    query01 += "LogInst,";
                    query01 += "U_LineNum,";
                    query01 += "U_InputGbn,";
                    query01 += "U_ItemCode,";
                    query01 += "U_ItemName,";
                    query01 += "U_ItemGpCd,";
                    query01 += "U_Weight,";
                    query01 += "U_DueDate,";
                    query01 += "U_CntcCode,";
                    query01 += "U_CntcName,";
                    query01 += "U_ProcType,";
                    query01 += "U_Comments,";
                    query01 += "U_BatchNum,";
                    query01 += "U_LineId";
                    query01 += ")";
                    query01 += " VALUES(";
                    query01 += "'" + CurrentDocEntry + "',";
                    query01 += "'" + i + "',";
                    query01 += "'" + (i - 1) + "',";
                    query01 += "'PS_PP030',";
                    query01 += "NULL,";
                    query01 += "'" + i + "',";
                    query01 += "'" + oMat02.Columns.Item("InputGbn").Cells.Item(i).Specific.Selected.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("ItemCode").Cells.Item(i).Specific.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("ItemName").Cells.Item(i).Specific.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("ItemGpCd").Cells.Item(i).Specific.Selected.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("Weight").Cells.Item(i).Specific.Value + "',";
                    if (string.IsNullOrEmpty(oMat02.Columns.Item("DueDate").Cells.Item(i).Specific.Value))
                    {
                        query01 += "NULL,";
                    }
                    else
                    {
                        query01 += "'" + oMat02.Columns.Item("DueDate").Cells.Item(i).Specific.Value + "',";
                    }
                    query01 += "'" + oMat02.Columns.Item("CntcCode").Cells.Item(i).Specific.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("CntcName").Cells.Item(i).Specific.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("ProcType").Cells.Item(i).Specific.Selected.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("Comments").Cells.Item(i).Specific.Value + "',";
                    query01 += "'" + oMat02.Columns.Item("BatchNum").Cells.Item(i).Specific.Value + "',";
                    query01 += "'" + i + "'";
                    query01 += ")";
                    RecordSet01.DoQuery(query01);

                    for (j = 1; j <= oMat03.VisualRowCount; j++)
                    {
                        query01 = "INSERT INTO [@PS_PP030M] (";
                        query01 += "DocEntry,";
                        query01 += "LineId,";
                        query01 += "VisOrder,";
                        query01 += "Object,";
                        query01 += "LogInst,";
                        query01 += "U_LineNum,";
                        query01 += "U_Sequence,";
                        query01 += "U_CpBCode,";
                        query01 += "U_CpBName,";
                        query01 += "U_CpCode,";
                        query01 += "U_CpName,";
                        query01 += "U_StdHour,";
                        query01 += "U_Unit,";
                        query01 += "U_ReDate,";
                        query01 += "U_WorkGbn,";
                        query01 += "U_ReWorkYN,";
                        query01 += "U_ResultYN,";
                        query01 += "U_ReportYN,";
                        query01 += "U_LineId";
                        query01 += ")";
                        query01 += " VALUES";
                        query01 += "(";
                        query01 += "'" + CurrentDocEntry + "',";
                        query01 += "'" + j + "',";
                        query01 += "'" + (j - 1) + "',";
                        query01 += "'PS_PP030',";
                        query01 += "NULL,";
                        query01 += "'" + j + "',";
                        query01 += "'" + j + "',";
                        query01 += "'" + oMat03.Columns.Item("CpBCode").Cells.Item(j).Specific.Value + "',";
                        query01 += "'" + oMat03.Columns.Item("CpBName").Cells.Item(j).Specific.Value + "',";
                        query01 += "'" + oMat03.Columns.Item("CpCode").Cells.Item(j).Specific.Value + "',";
                        query01 += "'" + oMat03.Columns.Item("CpName").Cells.Item(j).Specific.Value + "',";
                        query01 += "'" + oMat03.Columns.Item("StdHour").Cells.Item(j).Specific.Value + "',";
                        query01 += "'" + oMat03.Columns.Item("Unit").Cells.Item(j).Specific.Value + "',";
                        if (string.IsNullOrEmpty(oMat03.Columns.Item("ReDate").Cells.Item(j).Specific.Value))
                        {
                            query01 += "NULL,";
                        }
                        else
                        {
                            query01 += "'" + oMat03.Columns.Item("ReDate").Cells.Item(j).Specific.Value + "',";
                        }
                        query01 += "'" + oMat03.Columns.Item("WorkGbn").Cells.Item(j).Specific.Selected.Value + "',";
                        query01 += "'" + oMat03.Columns.Item("ReWorkYN").Cells.Item(j).Specific.Selected.Value + "',";
                        query01 += "'" + oMat03.Columns.Item("ResultYN").Cells.Item(j).Specific.Selected.Value + "',";
                        query01 += "'" + oMat03.Columns.Item("ReportYN").Cells.Item(j).Specific.Selected.Value + "',";
                        query01 += "'" + j + "'";
                        query01 += ")";
                        RecordSet01.DoQuery(query01);
                    }
                }

                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                returnValue = true;
            }
            catch(Exception ex)
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
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
                case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS: //4
                    //Raise_EVENT_LOST_FOCUS(FormUID, ref pVal, ref BubbleEvent);
                    break;
                case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT: //5
                    Raise_EVENT_COMBO_SELECT(FormUID, ref pVal, ref BubbleEvent);
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
                    Raise_EVENT_FORM_RESIZE(FormUID, ref pVal, ref BubbleEvent);
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
            string errMessage = string.Empty;
            int i;
            string query01;
            string oOrdGbn01;
            short li_Cnt;
            short li_LineId;
            //object lChildForm = null; //팝업창 호출 용 변수(2012.04.12 송명규)

            SAPbobsCOM.Recordset RecordSet01 = null;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_PP030_MTX01(); //조회
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "1")
                    {
                        ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP030_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oSCardCod01 = oForm.Items.Item("SCardCod").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;

                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107")
                            {
                                if (PS_PP030_AutoCreateMultiGage() == false)
                                {
                                    PS_PP030_AddMatrixRow01(oMat02.VisualRowCount, false);
                                    PS_PP030_AddMatrixRow02(oMat03.VisualRowCount, false);
                                    BubbleEvent = false;
                                    return;
                                }
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                PS_PP030_FormItemEnabled();
                                PSH_Globals.SBO_Application.ActivateMenuItem("1282");
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                //멀티게이지를 제외한 나머지 경우는 자동으로 입력
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP030_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                            if (oMat02.VisualRowCount == 0)
                            {
                                oMat02.Clear();
                                oMat02.AddRow();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }
                    else if (pVal.ItemUID == "btnWkSrch") //표준공수조회 버튼 클릭
                    {
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                        {
                            errMessage = "작업구분을 선택하십시오.";
                            throw new Exception();
                        }
                        else
                        {
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106")
                            {
                                PS_PP033 oTempClass = new PS_PP033();
                                oTempClass.LoadForm(this);
                            }
                            else
                            {
                                errMessage = "작업구분이 [제품_기계공구] 또는 [제품_몰드] 일 경우에만 사용이 가능합니다.";
                                throw new Exception();
                            }
                        }
                    }
                    else if (pVal.ItemUID == "btnItmSrch") //품목별공수조회 버튼 클릭
                    {
                        if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
                        {
                            errMessage = "작업구분을 선택하십시오.";
                            throw new Exception();
                        }
                        else
                        {
                            if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "105" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "106")
                            {
                                PS_PP031 oTempClass = new PS_PP031();
                                oTempClass.LoadForm("");
                            }
                            else
                            {
                                errMessage = "작업구분이 [제품_기계공구] 또는 [제품_몰드] 일 경우에만 사용이 가능합니다.";
                                throw new Exception();
                            }
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
                                RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                oOrdGbn01 = dataHelpClass.GetValue("SELECT U_OrdGbn FROM [@PS_PP030H] WHERE DocEntry = '" + oDocEntry01 + "'", 0, 1);
                                if (oOrdGbn01 == "105" || oOrdGbn01 == "106") //기계공구, 몰드
                                {
                                    query01 = "SELECT U_ProcType, DocEntry, LineId FROM [@PS_PP030L] WHERE DocEntry = '" + oDocEntry01 + "'";
                                    RecordSet01.DoQuery(query01);
                                    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                                    {
                                        if (RecordSet01.Fields.Item(0).Value == "10")
                                        {
                                            PS_PP030_PurchaseRequest(RecordSet01.Fields.Item(1).Value, RecordSet01.Fields.Item(2).Value);
                                        }
                                        RecordSet01.MoveNext();
                                    }
                                }
                                
                                PS_PP030_FormItemEnabled();
                                PS_PP030_AddMatrixRow01(0, true);
                                PS_PP030_AddMatrixRow02(0, true);
                                oForm.Items.Item("SCardCod").Specific.Value = oSCardCod01;
                                oForm.Items.Item("Total").Specific.Value = 0; //공정금액 합계 초기화
                                oForm.Items.Item("Button01").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                oForm.Items.Item("Total").Specific.Value = 0; //공정금액 합계 초기화
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    oOrdGbn01 = dataHelpClass.GetValue("SELECT U_OrdGbn FROM [@PS_PP030H] WHERE DocEntry = '" + oDocEntry01 + "'", 0, 1);
                                    
                                    if (oOrdGbn01 == "105" || oOrdGbn01 == "106") //기계공구, 몰드
                                    {
                                        query01 = "SELECT U_ProcType, DocEntry, LineId FROM [@PS_PP030L] WHERE DocEntry = '" + oDocEntry01 + "'";
                                        RecordSet01.DoQuery(query01);
                                        for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                                        {
                                            if (RecordSet01.Fields.Item(0).Value == "10")
                                            {
                                                PS_PP030_PurchaseRequest(RecordSet01.Fields.Item(1).Value, RecordSet01.Fields.Item(2).Value);
                                            }
                                            RecordSet01.MoveNext();
                                        }
                                    }
                                    else if (oOrdGbn01 == "104")
                                    {
                                        query01 = "Update [@PS_PP030M] set VisOrder = U_Sequence - 1, LineId = U_Sequence, U_LineId = U_Sequence WHERE LineId <> U_Sequence And DocEntry = '" + oDocEntry01 + "'";
                                        RecordSet01.DoQuery(query01);

                                        query01 = "SELECT Count(*), Min(LineId) FROM [@PS_PP030M] WHERE DocEntry = '" + oDocEntry01 + "' and U_CpCode = 'CP50107'";
                                        RecordSet01.DoQuery(query01);

                                        li_Cnt = RecordSet01.Fields.Item(0).Value;
                                        li_LineId = RecordSet01.Fields.Item(1).Value;

                                        if (li_Cnt > 1)
                                        {
                                            query01 = "Update [@PS_PP030M] set U_ResultYN = 'N' WHERE DocEntry = '" + oDocEntry01 + "' and LineId = '" + li_LineId + "'";
                                            RecordSet01.DoQuery(query01);
                                        }
                                        else
                                        {
                                            query01 = "Update [@PS_PP030M] set U_ResultYN = 'Y' WHERE DocEntry = '" + oDocEntry01 + "' and LineId = '" + li_LineId + "'";
                                            RecordSet01.DoQuery(query01);
                                        }
                                    }

                                    oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    PS_PP030_FormItemEnabled();
                                    oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                PS_PP030_FormItemEnabled();
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
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);
                }
            }
            finally
            {
                if (RecordSet01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RecordSet01);
                }

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }
                oForm.Freeze(false);
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
            string ordGbn;
            string InputGbn;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "CntcCode", "");
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                    {
                        dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
                    }
                    else
                    {
                        dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "ItemCode", "");
                    }

                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "CntcCode");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat03", "CpBCode");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat03", "CpCode");

                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.ColUID == "ItemCode")
                        {
                            if (oMat02.Columns.Item("InputGbn").Cells.Item(pVal.Row).Specific.Selected == null)
                            {
                                dataHelpClass.MDC_GF_Message("투입구분을 선택하세요", "W");
                                oMat02.Columns.Item("InputGbn").Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                if (PS_PP030_Validate("수정02") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() == "107010002" || oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() == "107010004") 
                                {
                                    return;
                                }
                                
                                ordGbn = oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim();
                                InputGbn = oMat02.Columns.Item("InputGbn").Cells.Item(pVal.Row).Specific.Selected.Value.ToString().Trim();

                                PS_SM021 oTempClass = new PS_SM021(); //품목선택 화면 호출
                                oTempClass.LoadForm(oForm, pVal.ItemUID, pVal.ColUID, pVal.Row, ordGbn, InputGbn, oDS_PS_PP030H.GetValue("U_BPLId", 0).ToString().Trim());
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.ColUID == "FailCode")
                        {
                            if (string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(pVal.Row).Specific.Value))
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("7425");
                                BubbleEvent = false;
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
                    if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02" || pVal.ItemUID == "Mat03")
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
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01Row01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02Row02 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat03Row03 = pVal.Row;
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
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemChanged == true) 
                    {
                        if (pVal.ItemUID == "Mat02")
                        {
                            if (PS_PP030_Validate("수정02") == false)
                            {
                                oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP030L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "특정컬럼")
                                {
                                }
                                else
                                {
                                    oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (PS_PP030_Validate("수정03") == false)
                            {
                                oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP030M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "WorkGbn")
                                {
                                    if (oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value == "10")
                                    {
                                        oDS_PS_PP030M.SetValue("U_CpPrice", pVal.Row - 1, Convert.ToString(dataHelpClass.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1) * oMat03.Columns.Item("StdHour").Cells.Item(pVal.Row).Specific.Value));
                                    }
                                    else if (oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value == "20")
                                    {
                                        oDS_PS_PP030M.SetValue("U_CpPrice", pVal.Row - 1, Convert.ToString(dataHelpClass.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1) * oMat03.Columns.Item("StdHour").Cells.Item(pVal.Row).Specific.Value));
                                    }

                                    oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                }
                                else if (pVal.ColUID == "특정컬럼")
                                {
                                    oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    if (oMat03.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP030M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                    {
                                        PS_PP030_AddMatrixRow02(pVal.Row, false);
                                    }
                                }
                                else
                                {
                                    oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "OrdGbn")
                            {
                                oDS_PS_PP030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
                                if (oHasMatrix01 == true)
                                {
                                    if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "104") //작업구분 멀티
                                    {
                                        oForm.Items.Item("BasicGub").Enabled = true;
                                        oForm.Items.Item("MulGbn1").Enabled = true;
                                        oForm.Items.Item("MulGbn2").Enabled = true;
                                        oForm.Items.Item("MulGbn3").Enabled = true;
                                    }
                                    else
                                    {
                                        oForm.Items.Item("BasicGub").Enabled = false;
                                        oForm.Items.Item("MulGbn1").Enabled = false;
                                        oForm.Items.Item("MulGbn2").Enabled = false;
                                        oForm.Items.Item("MulGbn3").Enabled = false;
                                    }
                                    
                                    if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "107") //엔드베어링
                                    {
                                        oMat02.Columns.Item("InputGbn").Editable = true;
                                    }
                                    else
                                    {
                                        oMat02.Columns.Item("InputGbn").Editable = false;
                                    }

                                    if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "105" || oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "106")
                                    {
                                        oMat02.Columns.Item("Weight").Editable = true;
                                    }
                                    else
                                    {
                                        oMat02.Columns.Item("Weight").Editable = false;
                                    }

                                    oMat02.Clear();
                                    oMat02.FlushToDataSource();
                                    oMat02.LoadFromDataSource();
                                    PS_PP030_AddMatrixRow01(0, true);
                                    oMat03.Clear();
                                    oMat03.FlushToDataSource();
                                    oMat03.LoadFromDataSource();
                                    PS_PP030_AddMatrixRow02(0, true);
                                    oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                    oForm.Items.Item("DocDate").Specific.Value = "";
                                    oForm.Items.Item("DueDate").Specific.Value = "";
                                    oForm.Items.Item("ItemCode").Specific.Value = "";
                                    oForm.Items.Item("ItemName").Specific.Value = "";
                                    oForm.Items.Item("JakMyung").Specific.Value = "";
                                    oForm.Items.Item("JakSize").Specific.Value = "";
                                    oForm.Items.Item("JakUnit").Specific.Value = "";
                                    oForm.Items.Item("CntcCode").Specific.Value = "";
                                    oForm.Items.Item("CntcName").Specific.Value = "";
                                    oForm.Items.Item("OrdMgNum").Specific.Value = "";
                                    oForm.Items.Item("ReqWt").Specific.Value = 0;
                                    oForm.Items.Item("SelWt").Specific.Value = 0;
                                }
                                else
                                {
                                    
                                    if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "104" || oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "107") //작업구분 멀티
                                    {
                                        
                                        if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "104") //작업구분 멀티
                                        {
                                            oForm.Items.Item("BasicGub").Enabled = true;
                                            oForm.Items.Item("MulGbn1").Enabled = true;
                                            oForm.Items.Item("MulGbn2").Enabled = true;
                                            oForm.Items.Item("MulGbn3").Enabled = true;
                                        }
                                        else
                                        {
                                            oForm.Items.Item("BasicGub").Enabled = false;
                                            oForm.Items.Item("MulGbn1").Enabled = false;
                                            oForm.Items.Item("MulGbn2").Enabled = false;
                                            oForm.Items.Item("MulGbn3").Enabled = false;
                                        }
                                        
                                        if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "107") //엔드베어링
                                        {
                                            oMat02.Columns.Item("InputGbn").Editable = true;
                                        }
                                        else
                                        {
                                            oMat02.Columns.Item("InputGbn").Editable = false;
                                        }

                                        if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "105" || oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "106")
                                        {
                                            oMat02.Columns.Item("Weight").Editable = true;
                                        }
                                        else
                                        {
                                            oMat02.Columns.Item("Weight").Editable = false;
                                        }
                                        
                                        oMat02.Clear();
                                        oMat02.FlushToDataSource();
                                        oMat02.LoadFromDataSource();
                                        PS_PP030_AddMatrixRow01(0, true);
                                        oMat03.Clear();
                                        oMat03.FlushToDataSource();
                                        oMat03.LoadFromDataSource();
                                        PS_PP030_AddMatrixRow02(0, true);
                                        oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                                        oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                                        oForm.Items.Item("DueDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                                        oForm.Items.Item("ItemCode").Specific.Value = "";
                                        oForm.Items.Item("ItemName").Specific.Value = "";
                                        oForm.Items.Item("JakMyung").Specific.Value = "";
                                        oForm.Items.Item("JakSize").Specific.Value = "";
                                        oForm.Items.Item("JakUnit").Specific.Value = "";
                                        oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                                        oForm.Items.Item("OrdMgNum").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                                        oForm.Items.Item("ReqWt").Specific.Value = 0;
                                        oForm.Items.Item("SelWt").Specific.Value = 0;
                                    }
                                    else if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "선택")
                                    {
                                        //멀티랑 엔드베어링일때
                                    }
                                    else
                                    {
                                        dataHelpClass.MDC_GF_Message("멀티,엔드베어링작업만 선택할수 있습니다.", "W");
                                        oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                    }
                                }
                            }
                        }

                        oMat02.LoadFromDataSource();
                        oMat03.LoadFromDataSource();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                        oForm.Update();

                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else
                        {
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
        /// CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            bool True_False;

            try
            {
                oForm.Freeze(true);

                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Opt01")
                    {
                        True_False = oMat02.Columns.Item("Weight").Editable;
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat02.Columns.Item("Weight").Editable = True_False;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                    }
                    if (pVal.ItemUID == "Opt02")
                    {
                        True_False = oMat02.Columns.Item("Weight").Editable;
                        oForm.Settings.MatrixUID = "Mat02";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat02.Columns.Item("Weight").Editable = True_False;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                    }
                    if (pVal.ItemUID == "Opt03")
                    {
                        True_False = oMat02.Columns.Item("Weight").Editable;
                        oForm.Settings.MatrixUID = "Mat03";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat02.Columns.Item("Weight").Editable = True_False;
                        oMat01.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oMat03.AutoResizeColumns();
                    }
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                            oMat01Row01 = pVal.Row;
                        }
                    }
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02.SelectRow(pVal.Row, true, false);
                            oMat02Row02 = pVal.Row;
                        }
                    }
                    if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat03.SelectRow(pVal.Row, true, false);
                            oMat03Row03 = pVal.Row;
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
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row == 0)
                        {
                            oMat01.Columns.Item(pVal.ColUID).TitleObject.Sortable = true;
                            oMat01.FlushToDataSource();
                        }
                        else if (pVal.Row > 0)
                        {
                            oHasMatrix01 = true;
                            oForm.Freeze(true);
                            oMat02.Clear();
                            oMat02.FlushToDataSource();
                            oMat02.LoadFromDataSource();
                            PS_PP030_AddMatrixRow01(0, true);
                            oForm.Items.Item("OrdGbn").Enabled = true;
                            oForm.Items.Item("BPLId").Enabled = true;
                            oForm.Items.Item("ItemCode").Enabled = true;
                            oForm.Items.Item("OrdMgNum").Enabled = true;
                            if (oMat01.Columns.Item("ItmBsort").Cells.Item(pVal.Row).Specific.Selected == null)
                            {
                                oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("BPLId").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                                oForm.Items.Item("DocDate").Specific.Value = "";
                                oForm.Items.Item("DueDate").Specific.Value = "";
                                oForm.Items.Item("BaseType").Specific.Value = "";
                                oForm.Items.Item("BaseNum").Specific.Value = "";
                                oForm.Items.Item("OrdMgNum").Specific.Value = "";
                                oForm.Items.Item("OrdNum").Specific.Value = "";
                                oForm.Items.Item("OrdSub1").Specific.Value = "";
                                oForm.Items.Item("OrdSub2").Specific.Value = "";
                                oForm.Items.Item("ItemCode").Specific.Value = "";
                                oForm.Items.Item("JakMyung").Specific.Value = "";
                                oForm.Items.Item("JakSize").Specific.Value = "";
                                oForm.Items.Item("JakUnit").Specific.Value = "";
                                oForm.Items.Item("ReqWt").Specific.Value = 0;
                                oForm.Items.Item("SelWt").Specific.Value = 0;
                                oForm.Items.Item("SjNum").Specific.Value = "";
                                oForm.Items.Item("SjLine").Specific.Value = "";
                                oForm.Items.Item("LotNo").Specific.Value = "";
                                oForm.Items.Item("SjPrice").Specific.Value = 0;
                            }
                            else
                            {
                                oForm.Items.Item("OrdGbn").Specific.Select(oMat01.Columns.Item("ItmBsort").Cells.Item(pVal.Row).Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("BPLId").Specific.Select(oMat01.Columns.Item("BPLId").Cells.Item(pVal.Row).Specific.Selected.Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                oForm.Items.Item("CntcCode").Specific.Value = dataHelpClass.User_MSTCOD();
                                oForm.Items.Item("ItemCode").Specific.Value = oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value;
                                oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                                oForm.Items.Item("DueDate").Specific.Value = oMat01.Columns.Item("ReqDate").Cells.Item(pVal.Row).Specific.Value;
                                oForm.Items.Item("BaseType").Specific.Value = oMat01.Columns.Item("BaseType").Cells.Item(pVal.Row).Specific.Value;
                                oForm.Items.Item("BaseNum").Specific.Value = oMat01.Columns.Item("BaseNum").Cells.Item(pVal.Row).Specific.Value;
                                oForm.Items.Item("OrdMgNum").Specific.Value = "";
                                if (oMat01.Columns.Item("BaseType").Cells.Item(pVal.Row).Specific.Value == "작번요청")
                                {
                                    oDS_PS_PP030H.SetValue("U_OrdNum", 0, oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP030H.SetValue("U_OrdSub1", 0, oMat01.Columns.Item("OrdSub1").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP030H.SetValue("U_OrdSub2", 0, oMat01.Columns.Item("OrdSub2").Cells.Item(pVal.Row).Specific.Value);

                                    if (oMat01.Columns.Item("OrdSub1").Cells.Item(pVal.Row).Specific.Value == "00")
                                    {
                                        oDS_PS_PP030H.SetValue("U_JakMyung", 0, dataHelpClass.GetValue("SELECT FrgnName FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                        oDS_PS_PP030H.SetValue("U_JakSize", 0, dataHelpClass.GetValue("SELECT U_Size FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                        oDS_PS_PP030H.SetValue("U_JakUnit", 0, dataHelpClass.GetValue("SELECT salUnitMsr FROM [OITM] WHERE ItemCode = '" + oMat01.Columns.Item("OrdNum").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                    }
                                    else
                                    {
                                        oForm.Items.Item("JakMyung").Specific.Value = oMat01.Columns.Item("JakMyung").Cells.Item(pVal.Row).Specific.Value;
                                        oForm.Items.Item("JakSize").Specific.Value = oMat01.Columns.Item("JakSize").Cells.Item(pVal.Row).Specific.Value;
                                        oForm.Items.Item("JakUnit").Specific.Value = oMat01.Columns.Item("JakUnit").Cells.Item(pVal.Row).Specific.Value;
                                    }
                                }
                                else if (oMat01.Columns.Item("BaseType").Cells.Item(pVal.Row).Specific.Value == "생산요청")
                                {
                                    oForm.Items.Item("OrdMgNum").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                                    oForm.Items.Item("OrdNum").Specific.Value = DateTime.Now.ToString("yyyyMMdd") + dataHelpClass.GetValue("EXEC PS_PP030_01 '" + oForm.Items.Item("OrdNum").Specific.Value + "'", 0, 1);
                                    oForm.Items.Item("OrdSub1").Specific.Value = "00";
                                    oForm.Items.Item("OrdSub2").Specific.Value = "000";

                                    oForm.Items.Item("JakMyung").Specific.Value = oMat01.Columns.Item("JakMyung").Cells.Item(pVal.Row).Specific.Value;
                                    oForm.Items.Item("JakSize").Specific.Value = oMat01.Columns.Item("JakSize").Cells.Item(pVal.Row).Specific.Value;
                                    oForm.Items.Item("JakUnit").Specific.Value = oMat01.Columns.Item("JakUnit").Cells.Item(pVal.Row).Specific.Value;
                                }

                                oForm.Items.Item("ReqWt").Specific.Value = oMat01.Columns.Item("RemainWt").Cells.Item(pVal.Row).Specific.Value;
                                oForm.Items.Item("SelWt").Specific.Value = oMat01.Columns.Item("RemainWt").Cells.Item(pVal.Row).Specific.Value;
                                oForm.Items.Item("SjNum").Specific.Value = oMat01.Columns.Item("ORDRNum").Cells.Item(pVal.Row).Specific.Value;
                                oForm.Items.Item("SjLine").Specific.Value = oMat01.Columns.Item("RDR1Num").Cells.Item(pVal.Row).Specific.Value;
                                oForm.Items.Item("LotNo").Specific.Value = dataHelpClass.GetValue("SELECT U_LotNo FROM [ORDR] WHERE DocEntry = '" + oMat01.Columns.Item("ORDRNum").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1);
                                oForm.Items.Item("SjPrice").Specific.Value = dataHelpClass.GetValue("SELECT LineTotal FROM [RDR1] WHERE DocEntry = '" + oMat01.Columns.Item("ORDRNum").Cells.Item(pVal.Row).Specific.Value + "' AND LineNum = '" + oMat01.Columns.Item("RDR1Num").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1);
                            }
                            oForm.Items.Item("OrdGbn").Enabled = false;
                            oForm.Items.Item("BPLId").Enabled = false;
                            oForm.Items.Item("ItemCode").Enabled = false;
                            if (oMat01.Columns.Item("BaseType").Cells.Item(pVal.Row).Specific.Value == "작번요청")
                            {
                                oForm.Items.Item("OrdMgNum").Enabled = false;
                            }
                            else if (oMat01.Columns.Item("BaseType").Cells.Item(pVal.Row).Specific.Value == "생산요청")
                            {
                                oForm.Items.Item("OrdMgNum").Enabled = true;
                            }
                            oForm.Freeze(false);
                            oHasMatrix01 = false;
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
            string sQry;
            double TotalAmt = 0;
            double TotalQty = 0;
            double useMkg;
            SAPbobsCOM.Recordset oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();            

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat02")
                        {
                            if (PS_PP030_Validate("수정02") == false)
                            {
                                oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP030L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "ItemCode")
                                {
                                    oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP030L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                    {
                                        PS_PP030_AddMatrixRow01(pVal.Row, false);
                                    }
                                    if (oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() == "107010002" || oForm.Items.Item("ItemCode").Specific.Value.ToString().Trim() == "107010004") 
                                    {
                                        oMat02.Columns.Item("BatchNum").Editable = true;
                                    }
                                }
                                else if (pVal.ColUID == "BatchNum")
                                {   
                                    oMat02.FlushToDataSource();

                                    sQry = "EXEC [PS_PP030_06] '" + oMat02.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "', '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'";
                                    oRecordSet01.DoQuery(sQry);

                                    oDS_PS_PP030L.SetValue("U_ItemName", pVal.Row - 1, oRecordSet01.Fields.Item("ItemName").Value);
                                    oDS_PS_PP030L.SetValue("U_ItemGpCd", pVal.Row - 1, oRecordSet01.Fields.Item("ItmsGrpCod").Value);
                                    oDS_PS_PP030L.SetValue("U_Unit", pVal.Row - 1, oRecordSet01.Fields.Item("InvntryUom").Value);
                                    oDS_PS_PP030L.SetValue("U_Weight", pVal.Row - 1, oRecordSet01.Fields.Item("Quantity").Value);
                                    oMat02.SetLineData(pVal.Row);
                                }
                                else if (pVal.ColUID == "Weight")
                                {
                                    if (Convert.ToDouble(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) < 0)
                                    {
                                        oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                    }
                                    else
                                    {
                                        oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                }
                                else if (pVal.ColUID == "CntcCode")
                                {
                                    oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP030L.SetValue("U_CntcName", pVal.Row - 1, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else if (pVal.ColUID == "Comments")
                                {
                                    if (oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "104" || oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "107") 
                                    {
                                        if (oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "107") 
                                        {
                                            sQry = "Select IsNull(U_useMkg, 0) From [OITM] Where ItemCode = '" + oDS_PS_PP030H.GetValue("U_ItemCode", 0).ToString().Trim() + "'";
                                            oRecordSet01.DoQuery(sQry);
                                            useMkg = Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) / 1000;

                                            for (i = 1; i <= oDS_PS_PP030L.Size - 1; i++)
                                            {
                                                TotalQty += Convert.ToDouble(oDS_PS_PP030L.GetValue("U_Weight", i - 1));
                                            }
                                            if (useMkg == 0)
                                            {
                                                oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(System.Math.Round(TotalQty, 0)));
                                                oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            }
                                            else
                                            {
                                                oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(System.Math.Round(TotalQty / useMkg, 0)));
                                                oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            }
                                        } 
                                        else
                                        {
                                            for (i = 1; i <= oDS_PS_PP030L.Size - 1; i++)
                                            {
                                                TotalQty += Convert.ToDouble(oDS_PS_PP030L.GetValue("U_Weight", i - 1));
                                            }
                                            oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(TotalQty));
                                            oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        }
                                    }
                                    oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                                else
                                {
                                    oDS_PS_PP030L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "StdHour" || pVal.ColUID == "ReDate")
                            {
                                oMat03.FlushToDataSource();
                                
                                if (pVal.ColUID == "StdHour") //표준공수 등록 시
                                {
                                    //공정단가 계산_S
                                    if (oMat03.Columns.Item("WorkGbn").Cells.Item(pVal.Row).Specific.Value == "10")
                                    {
                                        oDS_PS_PP030M.SetValue("U_CpPrice", pVal.Row - 1, Convert.ToString(Convert.ToDouble(dataHelpClass.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) * Convert.ToDouble(oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                    else if (oMat03.Columns.Item("WorkGbn").Cells.Item(pVal.Row).Specific.Value == "20")
                                    {
                                        oDS_PS_PP030M.SetValue("U_CpPrice", pVal.Row - 1, Convert.ToString(Convert.ToDouble(dataHelpClass.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) * Convert.ToDouble(oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    }
                                    oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    //공정단가 계산_E

                                    //합계 계산_S
                                    for (i = 0; i <= oMat03.VisualRowCount - 1; i++)
                                    {
                                        TotalAmt += Convert.ToDouble(oDS_PS_PP030M.GetValue("U_CpPrice", i));
                                    }

                                    oForm.Items.Item("Total").Specific.Value = TotalAmt;
                                    //합계 계산_E
                                }
                            }
                            
                            if (pVal.ColUID == "CpBCode" || pVal.ColUID == "CpCode") //작업일보가 등록된 작지 중에서 공정대분류와 공정중분류는 수정 불가
                            {
                                if (PS_PP030_Validate("수정03") == false)
                                {
                                    oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP030M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                                }
                                else
                                {
                                    if (pVal.ColUID == "CpBCode")
                                    {
                                        oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP030M.SetValue("U_CpBName", pVal.Row - 1, dataHelpClass.GetValue("SELECT Name FROM [@PS_PP001H] WHERE Code = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                        if (oMat03.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP030M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                        {
                                            PS_PP030_AddMatrixRow02(pVal.Row, false);
                                        }
                                    }
                                    else if (pVal.ColUID == "CpCode")
                                    {
                                        oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP030M.SetValue("U_CpName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE Code = '" + oMat03.Columns.Item("CpBCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                        oDS_PS_PP030M.SetValue("U_StdHour", pVal.Row - 1, "0");
                                        oDS_PS_PP030M.SetValue("U_CpPrice", pVal.Row - 1, "0");
                                        oDS_PS_PP030M.SetValue("U_ResultYN", pVal.Row - 1, "Y");
                                        if (oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value == "CP50103" || oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value == "CP50106")
                                        {
                                            oDS_PS_PP030M.SetValue("U_ReWorkYN", pVal.Row - 1, "Y");
                                        }
                                        else
                                        {
                                            oDS_PS_PP030M.SetValue("U_ReWorkYN", pVal.Row - 1, "N");
                                        }
                                    }
                                    else if (pVal.ColUID == "StdHour")
                                    {
                                        if (oMat03.Columns.Item("WorkGbn").Cells.Item(pVal.Row).Specific.Value == "10")
                                        {
                                            oDS_PS_PP030M.SetValue("U_CpPrice", pVal.Row - 1, Convert.ToString(dataHelpClass.GetValue("Select U_Price From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1) * oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value));
                                        }
                                        else if (oMat03.Columns.Item("WorkGbn").Cells.Item(pVal.Row).Specific.Value == "20")
                                        {
                                            oDS_PS_PP030M.SetValue("U_CpPrice", pVal.Row - 1, Convert.ToString(dataHelpClass.GetValue("Select U_PsmtP From [@PS_PP001L] Where U_CpCode = '" + oMat03.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1) * oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value));
                                        }
                                        oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                }
                            }
                            else if (pVal.ColUID == "FailCode")
                            {
                                if (oMat03.Columns.Item("ReWorkYN").Cells.Item(pVal.Row).Specific.Value == "Y")
                                {
                                    oDS_PS_PP030M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP030M.SetValue("U_FailName", pVal.Row - 1, dataHelpClass.GetValue("Select U_SmalName From [@PS_PP003L] Where U_SmalCode = '" + oMat03.Columns.Item("FailCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                            }
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP030H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "OrdMgNum")
                            {
                                if (oForm.Items.Item("BaseType").Specific.Value == "작번요청")
                                {
                                    oDS_PS_PP030H.SetValue("U_" + pVal.ItemUID, 0, oDS_PS_PP030H.GetValue("U_" + pVal.ItemUID, 0));
                                }
                                else
                                {
                                    if (string.IsNullOrEmpty(oForm.Items.Item(pVal.ItemUID).Specific.Value))
                                    {
                                        oDS_PS_PP030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                        oDS_PS_PP030H.SetValue("U_OrdNum", 0, "");
                                        oDS_PS_PP030H.SetValue("U_OrdSub1", 0, "");
                                        oDS_PS_PP030H.SetValue("U_OrdSub2", 0, "");
                                    }
                                    else
                                    {
                                        oDS_PS_PP030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                        oDS_PS_PP030H.SetValue("U_OrdNum", 0, oForm.Items.Item(pVal.ItemUID).Specific.Value + dataHelpClass.GetValue("EXEC PS_PP030_01 '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                                        oDS_PS_PP030H.SetValue("U_OrdSub1", 0, "00");
                                        oDS_PS_PP030H.SetValue("U_OrdSub2", 0, "000");
                                    }
                                }
                            }
                            else if (pVal.ItemUID == "ItemCode")
                            {
                                oDS_PS_PP030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_PP030H.SetValue("U_ItemName", 0, dataHelpClass.Get_ReData("ItemName", "ItemCode", "[OITM]", "'" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", ""));

                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104")
                                {
                                    oDS_PS_PP030H.SetValue("U_MulGbn1", 0, "");
                                    oDS_PS_PP030H.SetValue("U_MulGbn2", 0, "");
                                    oDS_PS_PP030H.SetValue("U_MulGbn3", 0, "");
                                    oDS_PS_PP030H.SetValue("U_MulGbn1", 0, dataHelpClass.GetValue("SELECT U_Jakup1 FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                                    oDS_PS_PP030H.SetValue("U_MulGbn2", 0, dataHelpClass.GetValue("SELECT U_Jakup2 FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                                    oDS_PS_PP030H.SetValue("U_MulGbn3", 0, dataHelpClass.GetValue("SELECT U_Jakup3 FROM [OITM] WHERE ItemCode = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                                }
                                else
                                {
                                    oDS_PS_PP030H.SetValue("U_MulGbn1", 0, "");
                                    oDS_PS_PP030H.SetValue("U_MulGbn2", 0, "");
                                    oDS_PS_PP030H.SetValue("U_MulGbn3", 0, "");
                                }

                                PS_PP030_MTX03();

                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "101"
                                 || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "102"
                                 || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "111"
                                 || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "601"
                                 || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "602")
                                {
                                    PS_PP030_MTX02(); //투입자재 처리
                                }
                                
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104" || oForm.Items.Item("OrdGbn").Specific.Selected.Value == "107")
                                {
                                    oDS_PS_PP030H.SetValue("U_JakMyung", 0, oDS_PS_PP030H.GetValue("U_ItemName", 0));
                                }
                            }
                            else if (pVal.ItemUID == "SelWt")
                            {
                                oDS_PS_PP030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                if (Convert.ToDouble(oForm.Items.Item(pVal.ItemUID).Specific.Value) < 0)
                                {
                                    oDS_PS_PP030H.SetValue("U_SelWt", 0, "0");
                                    dataHelpClass.MDC_GF_Message("수,중량이 올바르지 않습니다.", "W");
                                }

                                if (!string.IsNullOrEmpty(oForm.Items.Item("BaseType").Specific.Value))
                                {
                                    if (Convert.ToDouble(oForm.Items.Item(pVal.ItemUID).Specific.Value) > Convert.ToDouble(oForm.Items.Item("ReqWt").Specific.Value))
                                    {
                                        oDS_PS_PP030H.SetValue("U_SelWt", 0, oForm.Items.Item("ReqWt").Specific.Value);
                                        dataHelpClass.MDC_GF_Message("수,중량이 올바르지 않습니다.", "W");
                                    }
                                }
                            }
                            else if (pVal.ItemUID == "ReqCod") // 요청자 추가 20180726 황영수
                            {
                                sQry = "SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item("ReqCod").Specific.Value.ToString().Trim() + "'";
                                oRecordSet01.DoQuery(sQry);
                                oForm.Items.Item("ReqNam").Specific.Value = oRecordSet01.Fields.Item(0).Value;
                            }
                            else if (pVal.ItemUID == "CntcCode")
                            {
                                oDS_PS_PP030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oDS_PS_PP030H.SetValue("U_CntcName", 0, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oForm.Items.Item(pVal.ItemUID).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                if (pVal.ItemUID == "SItemCod" || pVal.ItemUID == "SCardCod")
                                {
                                }
                                else
                                {
                                    oDS_PS_PP030H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                }
                            }
                        }
                        oMat02.LoadFromDataSource();
                        oMat03.LoadFromDataSource();
                        oMat02.AutoResizeColumns();
                        oMat02.AutoResizeColumns();
                        oForm.Update();
                        if (pVal.ItemUID == "Mat01")
                        {
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
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
                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
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
            

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_PP030_FormItemEnabled();
                    if (pVal.ItemUID == "Mat01")
                    {
                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        for (i = 1; i <= oMat02.VisualRowCount; i++)
                        {
                            if (i <= oMat02.VisualRowCount)
                            {
                                if (string.IsNullOrEmpty(oMat02.Columns.Item("InputGbn").Cells.Item(i).Specific.Value))
                                {
                                    oMat02.DeleteRow(i);
                                    i -= 1;
                                }
                            }
                        }
                        for (i = 1; i <= oMat02.VisualRowCount; i++)
                        {
                            oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                        }
                        oMat02.FlushToDataSource();
                        if (oMat02.VisualRowCount == 0)
                        {
                            PS_PP030_AddMatrixRow01(oMat02.VisualRowCount, true);
                        }
                        else
                        {
                            PS_PP030_AddMatrixRow01(oMat02.VisualRowCount, false);
                        }
                        oMat02.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        PS_PP030_AddMatrixRow02(oMat03.VisualRowCount, false);

                        if (oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "104") 
                        {
                            oMat03.Columns.Item("Sequence").TitleObject.Sortable = true;
                            oMat03.Columns.Item("Sequence").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                            oMat03.Columns.Item("Sequence").TitleObject.Sortable = false;
                            oMat03.FlushToDataSource();
                        }
                        oMat03.AutoResizeColumns();
                    }

                    //공정금액 합계_S
                    double total = 0;
                    for (i = 0; i <= oMat03.VisualRowCount - 1; i++)
                    {
                        total += Convert.ToDouble(oMat03.Columns.Item("CpPrice").Cells.Item(i + 1).Specific.Value);
                    }
                    oForm.Items.Item("Total").Specific.Value = total;
                    //공정금액 합계_E
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
                    SubMain.Remove_Forms(oFormUniqueID);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oForm);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_USERDS01);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP030H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP030L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP030M);
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
        /// FORM_RESIZE 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_FORM_RESIZE(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
            	    PS_PP030_FormResize();
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
            SAPbouiCOM.DataTable oDataTable01 = ((SAPbouiCOM.ChooseFromListEvent)pVal).SelectedObjects; //ItemEvent를 ChooseFromListEvent로 명시적 형변환 후 SelectedObjects 할당

            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    if (oDataTable01 != null) //SelectedObjects 가 null이 아닐때만 실행(ChooseFromList 팝업창을 취소했을 때 미실행)
                    {
                        if (pVal.ItemUID == "SItemCod")
                        {
                            oForm.DataSources.UserDataSources.Item("SItemCod").Value = oDataTable01.Columns.Item("ItemCode").Cells.Item(0).Value;
                            oForm.DataSources.UserDataSources.Item("SItemNam").Value = oDataTable01.Columns.Item("ItemName").Cells.Item(0).Value;
                        }
                        else if (pVal.ItemUID == "SCardCod")
                        {
                            oForm.DataSources.UserDataSources.Item("SCardCod").Value = oDataTable01.Columns.Item("CardCode").Cells.Item(0).Value;
                            oForm.DataSources.UserDataSources.Item("SCardNam").Value = oDataTable01.Columns.Item("CardName").Cells.Item(0).Value;
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
                if (oDataTable01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDataTable01);
                }
            }
        }

        /// <summary>
        /// 행삭제 체크 메소드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_ROW_DELETE(string FormUID, SAPbouiCOM.MenuEvent pVal, bool BubbleEvent)
        {
            int i;
            double TotalQty = 0;
            double useMkg;
            string sQry;
            SAPbobsCOM.Recordset oRecordSet01 = null;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (oLastItemUID01 == "Mat02")
                        {
                            if (PS_PP030_Validate("행삭제02") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                        else if (oLastItemUID01 == "Mat03")
                        {
                            if (PS_PP030_Validate("행삭제03") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oLastItemUID01 == "Mat02")
                        {
                            for (i = 1; i <= oMat02.VisualRowCount; i++)
                            {
                                oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                                oMat02.Columns.Item("InputGbn").Cells.Item(i).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                            oMat02.FlushToDataSource();
                            oDS_PS_PP030L.RemoveRecord(oDS_PS_PP030L.Size - 1);
                            oMat02.LoadFromDataSource();
                            if (oMat02.RowCount == 0)
                            {
                                PS_PP030_AddMatrixRow01(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP030L.GetValue("U_ItemCode", oMat02.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP030_AddMatrixRow01(oMat02.RowCount, false);
                                }
                            }

                            if (oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "104" || oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "107") 
                            {
                                oRecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                if (oDS_PS_PP030H.GetValue("U_OrdGbn", 0).ToString().Trim() == "107") 
                                {
                                    sQry = "Select IsNull(U_useMkg, 0) From [OITM] Where ItemCode = '" + oDS_PS_PP030H.GetValue("U_ItemCode", 0).ToString().Trim() + "'";
                                    oRecordSet01.DoQuery(sQry);
                                    useMkg = Convert.ToDouble(oRecordSet01.Fields.Item(0).Value) / 1000;

                                    for (i = 1; i <= oDS_PS_PP030L.Size - 1; i++)
                                    {
                                        TotalQty += Convert.ToDouble(oDS_PS_PP030L.GetValue("U_Weight", i - 1));
                                    }
                                    if (useMkg == 0)
                                    {
                                        oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(TotalQty));
                                    }
                                    else
                                    {
                                        oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(TotalQty / useMkg));
                                    }
                                } 
                                else
                                {
                                    for (i = 1; i <= oDS_PS_PP030L.Size - 1; i++)
                                    {
                                        TotalQty += Convert.ToDouble(oDS_PS_PP030L.GetValue("U_Weight", i - 1));
                                    }
                                    oDS_PS_PP030H.SetValue("U_SelWt", 0, Convert.ToString(TotalQty));
                                }

                                oMat02.LoadFromDataSource();
                                oForm.Update();
                            }
                        }
                        else if (oLastItemUID01 == "Mat03")
                        {
                            double totalCpPrice = 0;

                            for (i = 1; i <= oMat03.VisualRowCount; i++)
                            {
                                oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                                oMat03.Columns.Item("Sequence").Cells.Item(i).Specific.Value = i;
                                totalCpPrice += Convert.ToDouble(oMat03.Columns.Item("CpPrice").Cells.Item(i).Specific.Value);
                            }
                            oMat03.FlushToDataSource();
                            oDS_PS_PP030M.RemoveRecord(oDS_PS_PP030M.Size - 1);
                            oMat03.LoadFromDataSource();
                            if (oMat03.RowCount == 0)
                            {
                                PS_PP030_AddMatrixRow02(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP030M.GetValue("U_CpBCode", oMat03.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP030_AddMatrixRow02(oMat03.RowCount, false);
                                }
                            }

                            oForm.Items.Item("Total").Specific.Value = totalCpPrice; //공정금액 합계
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
                if (oRecordSet01 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet01);
                }
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
                                if (PS_PP030_Validate("취소") == false)
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
                                dataHelpClass.MDC_GF_Message("현재 모드에서는 취소할수 없습니다.", "W");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1286": //닫기
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (PS_PP030_Validate("닫기") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PSH_Globals.SBO_Application.MessageBox("문서를 닫기(종료) 처리하겠습니까?", 1, "예", "아니오") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                dataHelpClass.MDC_GF_Message("현재 모드에서는 닫기(종료) 처리할 수 없습니다.", "W");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1292": //행추가
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
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
                        case "1292": //행추가
                            if (oLastItemUID01 == "Mat03")
                            {
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "104")
                                {
                                    if (PS_PP030_Validate("행추가03") == false) //행추가 가능검사
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                    oMat03.AddRow(1, oMat03Row03 - 1);
                                    for (i = 1; i <= oMat03.VisualRowCount; i++)
                                    {
                                        oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                                        oMat03.Columns.Item("Sequence").Cells.Item(i).Specific.Value = i;

                                        //새로추가된 행의 값 설정
                                        if (oMat03Row03 == i)
                                        {
                                            oMat03.Columns.Item("ReWorkYN").Cells.Item(i).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //PK/탈지일때 재작업여부 예
                                            oMat03.Columns.Item("ResultYN").Cells.Item(i).Specific.Select(1, SAPbouiCOM.BoSearchKey.psk_Index); //실적여부 아니오
                                            oMat03.Columns.Item("ReportYN").Cells.Item(i).Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index); //일보여부 예
                                        }
                                    }
                                    oMat03.FlushToDataSource();
                                    oMat03.LoadFromDataSource();
                                }
                                else
                                {
                                    dataHelpClass.MDC_GF_Message("멀티인 경우만 행추가 가능합니다.", "W");
                                }
                            }
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, pVal, BubbleEvent);
                            break;
                        case "1281": //찾기
                            PS_PP030_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP030_FormItemEnabled();
                            PS_PP030_AddMatrixRow01(0, true);
                            PS_PP030_AddMatrixRow02(0, true);
                            break;
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291": //레코드이동버튼
                            PS_PP030_FormItemEnabled();
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

                if (pVal.ItemUID == "Mat01" || pVal.ItemUID == "Mat02" || pVal.ItemUID == "Mat03")
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
                if (pVal.ItemUID == "Mat01")
                {
                    if (pVal.Row > 0)
                    {
                        oMat01Row01 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat02")
                {
                    if (pVal.Row > 0)
                    {
                        oMat02Row02 = pVal.Row;
                    }
                }
                else if (pVal.ItemUID == "Mat03")
                {
                    if (pVal.Row > 0)
                    {
                        oMat03Row03 = pVal.Row;
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
