//using System;
//using SAPbouiCOM;
//using PSH_BOne_AddOn.Data;
//using PSH_BOne_AddOn.Code;

//namespace PSH_BOne_AddOn
//{
//	/// <summary>
//	/// 작업일보등록(방산부품)
//	/// </summary>
//	internal class PS_PP044 : PSH_BaseClass
//	{
//		private string oFormUniqueID;
//		private SAPbouiCOM.Matrix oMat01;
//		private SAPbouiCOM.Matrix oMat02;
//		private SAPbouiCOM.Matrix oMat03;
//		private SAPbouiCOM.Matrix oMat04;
//		private SAPbouiCOM.Matrix oMat05;
//		private SAPbouiCOM.DBDataSource oDS_PS_PP044H; //등록헤더
//		private SAPbouiCOM.DBDataSource oDS_PS_PP044L; //등록라인
//		private SAPbouiCOM.DBDataSource oDS_PS_PP044M; //등록라인
//		private SAPbouiCOM.DBDataSource oDS_PS_PP044N; //등록라인
//		private SAPbouiCOM.DBDataSource oDS_PS_PP044T; //사원조회라인
//		private SAPbouiCOM.DBDataSource oDS_PS_PP044U; //사원조회라인
//		private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
//		private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
//		private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
//		private int oMat01Row01;
//		private int oMat02Row02;
//		private int oMat03Row03;
//        private string oDocType01;
//        private string oDocEntry01;
//        private string oOrdGbn;
//        private string oSequence;
//        private string oDocdate;
//        private string oGubun;
//        private SAPbouiCOM.BoFormMode oFormMode01;

//        /// <summary>
//        /// Form 호출
//        /// </summary>
//        /// <param name="oFormDocEntry"></param>
//        public override void LoadForm(string oFormDocEntry)
//		{
//			MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

//			try
//			{
//				oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP044.srf");
//				oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
//				oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
//				oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

//				//매트릭스의 타이틀높이와 셀높이를 고정
//				for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
//				{
//					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
//					oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
//				}

//				oFormUniqueID = "PS_PP044_" + SubMain.Get_TotalFormsCount();
//				SubMain.Add_Forms(this, oFormUniqueID, "PS_PP044");

//				PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
//				oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

//				oForm.SupportedModes = -1;
//				oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
//				oForm.DataBrowser.BrowseBy = "DocEntry";
				
//				oForm.Freeze(true);

//                PS_PP044_CreateItems();
//                PS_PP044_ComboBox_Setting();
//                PS_PP044_CF_ChooseFromList();
//                PS_PP044_EnableMenus();
//                PS_PP044_SetDocument(oFormDocEntry);
//                //PS_PP044_FormResize();
//            }
//			catch (Exception ex)
//			{
//				PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//			}
//			finally
//			{
//				oForm.Update();
//				oForm.Freeze(false);
//				oForm.Visible = true;
//				System.Runtime.InteropServices.Marshal.ReleaseComObject(oXmlDoc);
//			}
//		}

//        /// <summary>
//        /// 화면 Item 생성
//        /// </summary>
//        private void PS_PP044_CreateItems()
//        {
//            try
//            {
//                oDS_PS_PP044H = oForm.DataSources.DBDataSources.Item("@PS_PP040H");
//                oDS_PS_PP044L = oForm.DataSources.DBDataSources.Item("@PS_PP040L");
//                oDS_PS_PP044M = oForm.DataSources.DBDataSources.Item("@PS_PP040M");
//                oDS_PS_PP044N = oForm.DataSources.DBDataSources.Item("@PS_PP040N");
//                oDS_PS_PP044T = oForm.DataSources.DBDataSources.Item("@PS_USERDS01");
//                oDS_PS_PP044U = oForm.DataSources.DBDataSources.Item("@PS_USERDS02");

//                oMat01 = oForm.Items.Item("Mat01").Specific;
//                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//                oMat01.AutoResizeColumns();

//                oMat02 = oForm.Items.Item("Mat02").Specific;
//                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//                oMat02.AutoResizeColumns();

//                oMat03 = oForm.Items.Item("Mat03").Specific;
//                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//                oMat03.AutoResizeColumns();

//                oMat04 = oForm.Items.Item("Mat04").Specific;
//                oMat04.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//                oMat04.AutoResizeColumns();

//                oMat05 = oForm.Items.Item("Mat05").Specific;
//                oMat05.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
//                oMat05.AutoResizeColumns();

//                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");

//                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");

//                oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");

//                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
//                oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");

//                oForm.DataSources.UserDataSources.Add("Gubun", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("Gubun").Specific.DataBind.SetBound(true, "", "Gubun");

//                oForm.DataSources.UserDataSources.Add("EmpChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
//                oForm.Items.Item("EmpChk").Specific.DataBind.SetBound(true, "", "EmpChk");

//                oDocType01 = "작업일보등록(작지)";

//                if (oDocType01 == "작업일보등록(작지)")
//                {
//                    oForm.Items.Item("DocType").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
//                }
//                else if (oDocType01 == "작업일보등록(공정)")
//                {
//                    oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
//                }
//            }
//            catch (Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// Combobox 설정
//        /// </summary>
//        private void PS_PP044_ComboBox_Setting()
//        {
//            string sQry;
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
//                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "10", "일반");
//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "20", "PSMT지원");
//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "30", "외주");
//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "40", "실적");
//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "50", "일반조정");
//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "60", "외주조정");
//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "OrdType", "", "70", "설계시간");
//                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OrdType").Specific, "PS_PP044", "OrdType", false);

//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "DocType", "", "10", "작지기준");
//                dataHelpClass.Combo_ValidValues_Insert("PS_PP044", "DocType", "", "20", "공정기준");
//                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PP044", "DocType", false);

//                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
//                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND CODE IN('102','602') order by Code", "", false, false);

//                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
//                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");

//                oForm.Items.Item("Gubun").Specific.ValidValues.Add("선택", "선택");
//                oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                //작업구분코드(2014.04.15 송명규 수정)
//                sQry = "  SELECT    U_Minor,";
//                sQry += "           U_CdName";
//                sQry += " FROM      [@PS_SY001L]";
//                sQry += " WHERE     Code = 'P203'";
//                sQry += "           AND U_UseYN = 'Y'";
//                sQry += " ORDER BY  U_Seq";

//                if (oMat01.Columns.Item("WorkCls").ValidValues.Count > 0)
//                {
//                    for (int loopCount = 0; loopCount <= oMat01.Columns.Item("WorkCls").ValidValues.Count - 1; loopCount++)
//                    {
//                        oMat01.Columns.Item("WorkCls").ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
//                    }

//                    dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkCls"), sQry, "", "");
//                }
//                else
//                {
//                    dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkCls"), sQry, "", "");
//                }
//            }
//            catch(Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// ChooseFromList 설정
//        /// </summary>
//        private void PS_PP044_CF_ChooseFromList()
//        {
//            SAPbouiCOM.ChooseFromListCollection oCFLs = null;
//            SAPbouiCOM.Conditions oCons = null;
//            SAPbouiCOM.Condition oCon = null;
//            SAPbouiCOM.ChooseFromList oCFL = null;
//            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
//            SAPbouiCOM.EditText oEdit = null;

//            try
//            {
//                oEdit = oForm.Items.Item("ItemCode").Specific;
//                oCFLs = oForm.ChooseFromLists;
//                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

//                oCFLCreationParams.ObjectType = "4";
//                oCFLCreationParams.UniqueID = "CFLITEMCODE";
//                oCFLCreationParams.MultiSelection = false;
//                oCFL = oCFLs.Add(oCFLCreationParams);

//                oCons = oCFL.GetConditions();
//                oCon = oCons.Add();
//                oCon.Alias = "ItmsGrpCod";
//                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
//                oCon.CondVal = "102";
//                oCFL.SetConditions(oCons);

//                oEdit.ChooseFromListUID = "CFLITEMCODE";
//                oEdit.ChooseFromListAlias = "ItemCode";
//            }
//            catch(Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//            finally
//            {
//                if (oCFLs != null)
//                {
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLs);
//                }

//                if (oCons != null)
//                {
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons);
//                }

//                if (oCon != null)
//                {
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon);
//                }

//                if (oCFL != null)
//                {
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
//                }

//                if (oCFLCreationParams != null)
//                {
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
//                }

//                if (oEdit != null)
//                {
//                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
//                }
//            }
//        }

//        /// <summary>
//        /// 메뉴설정
//        /// </summary>
//        private void PS_PP044_EnableMenus()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false);
//            }
//            catch(Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// SetDocument
//        /// </summary>
//        /// <param name="oFormDocEntry">DocEntry</param>
//        private void PS_PP044_SetDocument(string oFormDocEntry)
//        {
//            try
//            {
//                if (string.IsNullOrEmpty(oFormDocEntry))
//                {
//                    PS_PP044_FormItemEnabled();
//                    PS_PP044_AddMatrixRow01(0, true);
//                    PS_PP044_AddMatrixRow02(0, true);
//                }
//                else
//                {
//                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//                    PS_PP044_FormItemEnabled();
//                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
//                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//                }
//            }
//            catch(Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// 각모드에따른 아이템설정
//        /// </summary>
//        private void PS_PP044_FormItemEnabled()
//        {
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                oForm.Freeze(true);
//                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
//                {
//                    oForm.EnableMenu("1281", true); //찾기
//                    oForm.EnableMenu("1282", false); //추가
//                    oForm.Items.Item("DocEntry").Enabled = false;
//                    oForm.Items.Item("OrdType").Enabled = true;
//                    oForm.Items.Item("OrdMgNum").Enabled = true;
//                    oForm.Items.Item("DocDate").Enabled = true;
//                    oForm.Items.Item("Button01").Enabled = true;
//                    oForm.Items.Item("1").Enabled = true;
//                    oForm.Items.Item("Mat01").Enabled = true;
//                    oForm.Items.Item("Mat02").Enabled = true;
//                    oForm.Items.Item("Mat03").Enabled = true;
//                    oMat02.Columns.Item("NTime").Editable = true; //비가동시간만 사용

//                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_Index);
//                    oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

//                    if (string.IsNullOrEmpty(oOrdGbn))
//                    {
//                        oForm.Items.Item("OrdGbn").Specific.Select("102", SAPbouiCOM.BoSearchKey.psk_ByValue);
//                    }
//                    else
//                    {
//                        oForm.Items.Item("OrdGbn").Specific.Select(oOrdGbn, SAPbouiCOM.BoSearchKey.psk_ByValue);
//                    }

//                    if (oGubun == "선택" || string.IsNullOrEmpty(oGubun))
//                    {
//                        oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//                    }
//                    else
//                    {
//                    oForm.Items.Item("Gubun").Specific.Select(oGubun, SAPbouiCOM.BoSearchKey.psk_ByValue);
//                    }

//                    PS_PP044_FormClear();

//                    if (oDocType01 == "작업일보등록(작지)")
//                    {
//                        oDS_PS_PP044H.SetValue("U_DocType", 0, "10");
//                    }
//                    else if (oDocType01 == "작업일보등록(공정)")
//                    {
//                        oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
//                    }
//                    if (string.IsNullOrEmpty(oDocdate))
//                    {
//                        oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
//                    }
//                    else
//                    {
//                        oForm.Items.Item("DocDate").Specific.Value = oDocdate;
//                    }
//                }
//                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
//                {
//                    oForm.EnableMenu("1281", false); //찾기
//                    oForm.EnableMenu("1282", true); //추가
//                    oForm.Items.Item("DocEntry").Enabled = true;
//                    oForm.Items.Item("OrdType").Enabled = true;
//                    oForm.Items.Item("OrdMgNum").Enabled = true;
//                    oForm.Items.Item("DocDate").Enabled = true;
//                    oForm.Items.Item("Button01").Enabled = true;
//                    oForm.Items.Item("1").Enabled = true;
//                    oForm.Items.Item("Mat01").Enabled = false;
//                    oForm.Items.Item("Mat02").Enabled = false;
//                    oForm.Items.Item("Mat03").Enabled = false;
//                }
//                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
//                {
//                    oForm.EnableMenu("1281", true); //찾기
//                    oForm.EnableMenu("1282", true); //추가

//                    if (oGubun == "선택" || string.IsNullOrEmpty(oGubun))
//                    {
//                        oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//                    }
//                    else
//                    {
//                        oForm.Items.Item("Gubun").Specific.Select(oGubun, SAPbouiCOM.BoSearchKey.psk_ByValue);
//                    }

//                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oDS_PS_PP044H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "Y")
//                    {
//                        oForm.Items.Item("DocEntry").Enabled = false;
//                        oForm.Items.Item("OrdType").Enabled = false;
//                        oForm.Items.Item("OrdMgNum").Enabled = false;
//                        oForm.Items.Item("DocDate").Enabled = false;
//                        oForm.Items.Item("Button01").Enabled = false;
//                        oForm.Items.Item("1").Enabled = false;
//                        oForm.Items.Item("Mat01").Enabled = false;
//                        oForm.Items.Item("Mat02").Enabled = false;
//                        oForm.Items.Item("Mat03").Enabled = false;
//                    }
//                    else
//                    {
//                        if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "10" 
//                         || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "50" 
//                         || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "60" 
//                         || oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "70") //조정, 설계
//                        {
//                            oForm.Items.Item("DocEntry").Enabled = false;
//                            oForm.Items.Item("OrdType").Enabled = false;
//                            oForm.Items.Item("OrdMgNum").Enabled = true;
//                            oForm.Items.Item("DocDate").Enabled = true;
//                            oForm.Items.Item("Button01").Enabled = true;
//                            oForm.Items.Item("1").Enabled = true;
//                            oForm.Items.Item("Mat01").Enabled = true;
//                            oForm.Items.Item("Mat02").Enabled = true;
//                            oForm.Items.Item("Mat03").Enabled = true;
//                        }
//                        else if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "20") //PSMT
//                        {
//                            oForm.Items.Item("DocEntry").Enabled = false;
//                            oForm.Items.Item("OrdType").Enabled = false;
//                            oForm.Items.Item("OrdMgNum").Enabled = true;
//                            oForm.Items.Item("DocDate").Enabled = true;
//                            oForm.Items.Item("Button01").Enabled = true;
//                            oForm.Items.Item("1").Enabled = true;
//                            oForm.Items.Item("Mat01").Enabled = true;
//                            oForm.Items.Item("Mat02").Enabled = true;
//                            oForm.Items.Item("Mat03").Enabled = true;
//                        }
//                        else if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "30") //외주
//                        {
//                            oForm.Items.Item("DocEntry").Enabled = false;
//                            oForm.Items.Item("OrdType").Enabled = false;
//                            oForm.Items.Item("OrdMgNum").Enabled = false;
//                            oForm.Items.Item("DocDate").Enabled = false;
//                            oForm.Items.Item("Button01").Enabled = false;
//                            oForm.Items.Item("1").Enabled = false;
//                            oForm.Items.Item("Mat01").Enabled = false;
//                            oForm.Items.Item("Mat02").Enabled = false;
//                            oForm.Items.Item("Mat03").Enabled = false;
//                        }
//                        else if (oDS_PS_PP044H.GetValue("U_OrdType", 0).ToString().Trim() == "40") //실적
//                        {
//                            oForm.Items.Item("DocEntry").Enabled = false;
//                            oForm.Items.Item("OrdType").Enabled = false;
//                            oForm.Items.Item("OrdMgNum").Enabled = false;
//                            oForm.Items.Item("DocDate").Enabled = false;
//                            oForm.Items.Item("Button01").Enabled = false;
//                            oForm.Items.Item("1").Enabled = false;
//                            oForm.Items.Item("Mat01").Enabled = false;
//                            oForm.Items.Item("Mat02").Enabled = false;
//                            oForm.Items.Item("Mat03").Enabled = false;
//                        }
//                    }
//                }
//            }
//            catch(Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }

//        /// <summary>
//        /// DocEntry 초기화
//        /// </summary>
//        private void PS_PP044_FormClear()
//        {
//            string DocEntry;
//            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

//            try
//            {
//                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP040'", "");

//                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
//                {
//                    oForm.Items.Item("DocEntry").Specific.Value = "1";
//                }
//                else
//                {
//                    oForm.Items.Item("DocEntry").Specific.Value = DocEntry;
//                }
//            }
//            catch(Exception ex)
//            {
//                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
//            }
//        }



//        #region Raise_ItemEvent
//        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	switch (pVal.EventType) {
//        //		case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
//        //			////1
//        //			Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
//        //			////2
//        //			Raise_EVENT_KEY_DOWN(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
//        //			////5
//        //			Raise_EVENT_COMBO_SELECT(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_CLICK:
//        //			////6
//        //			Raise_EVENT_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
//        //			////7
//        //			Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
//        //			////8
//        //			Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_VALIDATE:
//        //			////10
//        //			Raise_EVENT_VALIDATE(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
//        //			////11
//        //			Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pVal, ref BubbleEvent);

//        //			// 작업시간 합계 추가 S
//        //			//            Dim i&
//        //			//            Dim Total As Currency
//        //			//
//        //			//
//        //			//                For i = 0 To oMat01.VisualRowCount - 1
//        //			//
//        //			//                    Total = Total + Val(oMat01.Columns("WorkTime").Cells(i + 1).Specific.Value)
//        //			//'                 oMat01.Columns("Total").Cells.Specific.Value = Total
//        //			//                Next i
//        //			//                oForm.Items("Total").Specific.Value = Total
//        //			PS_PP044_SumWorkTime();
//        //			break;
//        //		// 작업시간 합계 추가 E

//        //		//            Call Raise_EVENT_MATRIX_LOAD(FormUID, pVal, BubbleEvent)

//        //		case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
//        //			////18
//        //			break;
//        //		////et_FORM_ACTIVATE
//        //		case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
//        //			////19
//        //			break;
//        //		////et_FORM_DEACTIVATE
//        //		case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
//        //			////20
//        //			Raise_EVENT_RESIZE(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
//        //			////27
//        //			Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
//        //			////3
//        //			Raise_EVENT_GOT_FOCUS(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //		case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
//        //			////4
//        //			break;
//        //		////et_LOST_FOCUS
//        //		case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
//        //			////17
//        //			Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pVal, ref BubbleEvent);
//        //			break;
//        //	}
//        //	return;
//        //	Raise_ItemEvent_Error:
//        //	///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_MenuEvent
//        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	////BeforeAction = True
//        //	if ((pVal.BeforeAction == true)) {
//        //		switch (pVal.MenuUID) {
//        //			case "1284":
//        //				//취소
//        //				if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//        //					if ((PS_PP044_Validate("취소") == false)) {
//        //						BubbleEvent = false;
//        //						return;
//        //					}
//        //					if (SubMain.Sbo_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1")) {
//        //						BubbleEvent = false;
//        //						return;
//        //					}
//        //				} else {
//        //					MDC_Com.MDC_GF_Message(ref "현재 모드에서는 취소할수 없습니다.", ref "W");
//        //					BubbleEvent = false;
//        //					return;
//        //				}
//        //				break;
//        //			case "1286":
//        //				//닫기
//        //				break;
//        //			case "1293":
//        //				//행삭제
//        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
//        //				break;
//        //			case "1281":
//        //				//찾기
//        //				break;
//        //			case "1282":
//        //				//추가
//        //				break;
//        //			case "1288":
//        //			case "1289":
//        //			case "1290":
//        //			case "1291":
//        //				//레코드이동버튼
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("Gubun").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
//        //				Raise_EVENT_RECORD_MOVE(ref FormUID, ref pVal, ref BubbleEvent);
//        //				break;
//        //		}
//        //	////BeforeAction = False
//        //	} else if ((pVal.BeforeAction == false)) {
//        //		switch (pVal.MenuUID) {
//        //			case "1284":
//        //				//취소
//        //				break;
//        //			case "1286":
//        //				//닫기
//        //				break;
//        //			case "1293":
//        //				//행삭제
//        //				Raise_EVENT_ROW_DELETE(ref FormUID, ref pVal, ref BubbleEvent);
//        //				break;
//        //			case "1281":
//        //				//찾기
//        //				PS_PP044_FormItemEnabled();
//        //				////UDO방식
//        //				oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //				break;
//        //			case "1282":
//        //				//추가
//        //				PS_PP044_FormItemEnabled();
//        //				////UDO방식
//        //				PS_PP044_AddMatrixRow01(0, ref true);
//        //				////UDO방식
//        //				PS_PP044_AddMatrixRow02(0, ref true);
//        //				////UDO방식
//        //				break;
//        //			case "1288":
//        //			case "1289":
//        //			case "1290":
//        //			case "1291":
//        //				//레코드이동버튼

//        //				Raise_EVENT_RECORD_MOVE(ref FormUID, ref pVal, ref BubbleEvent);
//        //				break;
//        //		}
//        //	}
//        //	return;
//        //	Raise_MenuEvent_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_FormDataEvent
//        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	////BeforeAction = True
//        //	if ((BusinessObjectInfo.BeforeAction == true)) {
//        //		switch (BusinessObjectInfo.EventType) {
//        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//        //				////33
//        //				break;
//        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//        //				////34
//        //				break;
//        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//        //				////35
//        //				break;
//        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//        //				////36
//        //				break;
//        //		}
//        //	////BeforeAction = False
//        //	} else if ((BusinessObjectInfo.BeforeAction == false)) {
//        //		switch (BusinessObjectInfo.EventType) {
//        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
//        //				////33
//        //				if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//        //					if ((PS_PP044_FindValidateDocument("@PS_PP040H") == false)) {
//        //						////찾기메뉴 활성화일때 수행
//        //						if (SubMain.Sbo_Application.Menus.Item("1281").Enabled == true) {
//        //							SubMain.Sbo_Application.ActivateMenuItem(("1281"));
//        //						} else {
//        //							SubMain.Sbo_Application.SetStatusBarMessage("관리자에게 문의바랍니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //						}
//        //						BubbleEvent = false;
//        //						return;
//        //					}
//        //				}
//        //				break;
//        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
//        //				////34
//        //				break;
//        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
//        //				////35
//        //				break;
//        //			case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
//        //				////36
//        //				break;
//        //		}
//        //	}
//        //	return;
//        //	Raise_FormDataEvent_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_RightClickEvent
//        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	if (pVal.BeforeAction == true) {
//        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
//        //		//            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
//        //		//            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
//        //		//            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
//        //		//            MenuCreationParams01.uniqueID = "MenuUID"
//        //		//            MenuCreationParams01.String = "메뉴명"
//        //		//            MenuCreationParams01.Enabled = True
//        //		//            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
//        //		//        End If
//        //	} else if (pVal.BeforeAction == false) {
//        //		//        If pVal.ItemUID = "Mat01" And pVal.Row > 0 And pVal.Row <= oMat01.RowCount Then
//        //		//                Call Sbo_Application.Menus.RemoveEx("MenuUID")
//        //		//        End If
//        //	}
//        //	if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02" | pVal.ItemUID == "Mat03") {
//        //		if (pVal.Row > 0) {
//        //			oLastItemUID01 = pVal.ItemUID;
//        //			oLastColUID01 = pVal.ColUID;
//        //			oLastColRow01 = pVal.Row;
//        //		}
//        //	} else {
//        //		oLastItemUID01 = pVal.ItemUID;
//        //		oLastColUID01 = "";
//        //		oLastColRow01 = 0;
//        //	}
//        //	if (pVal.ItemUID == "Mat01") {
//        //		if (pVal.Row > 0) {
//        //			oMat01Row01 = pVal.Row;
//        //		}
//        //	} else if (pVal.ItemUID == "Mat02") {
//        //		if (pVal.Row > 0) {
//        //			oMat02Row02 = pVal.Row;
//        //		}
//        //	} else if (pVal.ItemUID == "Mat03") {
//        //		if (pVal.Row > 0) {
//        //			oMat03Row03 = pVal.Row;
//        //		}
//        //	}
//        //	return;
//        //	Raise_RightClickEvent_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_ITEM_PRESSED
//        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	string vReturnValue = null;
//        //	int i = 0;
//        //	int j = 0;
//        //	string CntcCode = null;
//        //	string IsYN = null;
//        //	double YTime = 0;

//        //	string DocEntry = null;
//        //	string LineNum = null;
//        //	int ErrNum = 0;
//        //	string DocNum = null;
//        //	string WinTitle = null;
//        //	string ReportName = null;
//        //	string[] oText = new string[2];
//        //	string sQry = null;
//        //	string sQryS = null;
//        //	string sQry1 = null;
//        //	string WorkName = null;
//        //	SAPbobsCOM.Recordset oRecordSet01 = null;
//        //	if (pVal.BeforeAction == true) {
//        //		if (pVal.ItemUID == "PS_PP044") {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//        //			}
//        //		}
//        //		if (pVal.ItemUID == "1") {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//        //				if (PS_PP044_DataValidCheck() == false) {
//        //					BubbleEvent = false;
//        //					return;
//        //				}
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oDocEntry01 = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oOrdGbn = Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value);
//        //				////작업구분
//        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oSequence = Strings.Trim(oMat01.Columns.Item("Sequence").Cells.Item(1).Specific.Value);
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oDocdate = Strings.Trim(oForm.Items.Item("DocDate").Specific.Value);
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oGubun = Strings.Trim(oForm.Items.Item("Gubun").Specific.Selected.Value);
//        //				////선택한 반을 다시 선택하도록
//        //				oFormMode01 = oForm.Mode;
//        //				////해야할일 작업
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//        //				if (PS_PP044_DataValidCheck() == false) {
//        //					BubbleEvent = false;
//        //					return;
//        //				}
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oDocEntry01 = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
//        //				oFormMode01 = oForm.Mode;
//        //				////해야할일 작업
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//        //			}
//        //		}

//        //		////취소버튼 누를시 저장할 자료가 있으면 메시지 표시
//        //		if (pVal.ItemUID == "2") {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//        //				if (oMat01.VisualRowCount > 1) {
//        //					vReturnValue = Convert.ToString(SubMain.Sbo_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소"));
//        //					switch (vReturnValue) {
//        //						case Convert.ToString(1):
//        //							break;
//        //						case Convert.ToString(2):
//        //							BubbleEvent = false;
//        //							return;

//        //							break;
//        //					}
//        //				}
//        //			}
//        //		}

//        //		if (pVal.ItemUID == "Button01") {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//        //				PS_PP044_OrderInfoLoad();
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//        //				PS_PP044_OrderInfoLoad();
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//        //			}
//        //		}
//        //		if (pVal.ItemUID == "Button02") {


//        //			oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        //			MDC_PS_Common.ConnectODBC();

//        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
//        //			for (i = 0; i <= oMat01.VisualRowCount - 1; i++) {
//        //				if (oMat01.IsRowSelected(i + 1) == true) {
//        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					LineNum = oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value;
//        //				}
//        //			}

//        //			WinTitle = " 공정카드 [PS_PP044]";
//        //			ReportName = "PS_PP044_01.rpt";

//        //			sQry1 = "Select U_WorkName From [@PS_PP040M] Where DocEntry = '" + DocEntry + "' And IsNull(U_WorkName, '') <> ''";
//        //			oRecordSet01.DoQuery(sQry1);

//        //			while (!(oRecordSet01.EoF)) {
//        //				WorkName = WorkName + "     " + oRecordSet01.Fields.Item(0).Value;
//        //				oRecordSet01.MoveNext();
//        //			}
//        //			MDC_Globals.gRpt_Formula = new string[2];
//        //			MDC_Globals.gRpt_Formula_Value = new string[2];

//        //			////Formula 수식필드

//        //			oText[1] = WorkName;

//        //			for (i = 1; i <= 1; i++) {
//        //				if (Strings.Len("" + i + "") == 1) {
//        //					MDC_Globals.gRpt_Formula[i] = "F0" + i + "";
//        //				} else {
//        //					MDC_Globals.gRpt_Formula[i] = "F" + i + "";
//        //				}
//        //				MDC_Globals.gRpt_Formula_Value[i] = oText[i];
//        //			}
//        //			MDC_Globals.gRpt_SRptSqry = new string[2];
//        //			MDC_Globals.gRpt_SRptName = new string[2];
//        //			MDC_Globals.gRpt_SFormula = new string[2, 2];
//        //			MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

//        //			////SubReport

//        //			MDC_Globals.gRpt_SFormula[1, 1] = "";
//        //			MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

//        //			sQryS = "EXEC [PS_PP040_06] '" + DocEntry + "', '" + LineNum + "', 'S'";

//        //			MDC_Globals.gRpt_SRptSqry[1] = sQryS;
//        //			MDC_Globals.gRpt_SRptName[1] = "PS_PP044_S1";

//        //			////조회조건문
//        //			sQry = "EXEC [PS_PP040_06] '" + DocEntry + "', '" + LineNum + "', 'M'";
//        //			oRecordSet01.DoQuery(sQry);
//        //			if (oRecordSet01.RecordCount == 0) {
//        //				MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다.확인해 주세요.", ref "E");
//        //				return;
//        //			}

//        //			////CR Action
//        //			if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "N", "V") == false) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //			}
//        //			//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //			oRecordSet01 = null;
//        //		}
//        //		if (pVal.ItemUID == "Button03") {
//        //			IsYN = "N";
//        //			if (oMat04.VisualRowCount == 0) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("작업자정보 라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //			} else {

//        //				oForm.Freeze(true);

//        //				for (i = 1; i <= oMat04.VisualRowCount; i++) {
//        //					//UPGRADE_WARNING: oMat04.Columns(Check).Cells(i).Specific.Checked 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if ((oMat04.Columns.Item("Check").Cells.Item(i).Specific.Checked == true)) {
//        //						//UPGRADE_WARNING: oMat04.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						CntcCode = oMat04.Columns.Item("CntcCode").Cells.Item(i).Specific.Value;
//        //						//UPGRADE_WARNING: oMat04.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						YTime = Conversion.Val(oMat04.Columns.Item("Base").Cells.Item(i).Specific.Value) + Conversion.Val(oMat04.Columns.Item("Extend").Cells.Item(i).Specific.Value);
//        //						for (j = 1; j <= oMat02.VisualRowCount - 1; j++) {
//        //							//UPGRADE_WARNING: oMat02.Columns(WorkCode).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							if (CntcCode == oMat02.Columns.Item("WorkCode").Cells.Item(j).Specific.Value) {
//        //								IsYN = "Y";
//        //							}
//        //						}
//        //						if (IsYN == "N") {
//        //							oDS_PS_PP044M.SetValue("U_TTime", oMat02.VisualRowCount - 1, Strings.Trim(Conversion.Str(YTime)));
//        //							//oMat02.Columns("TTime").Cells(oMat02.VisualRowCount).Specific.Value = Str(YTime)
//        //							//UPGRADE_WARNING: oMat02.Columns(YTime).Cells(oMat02.VisualRowCount).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							oMat02.Columns.Item("YTime").Cells.Item(oMat02.VisualRowCount).Specific.Value = Conversion.Str(YTime);
//        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							oMat02.Columns.Item("WorkCode").Cells.Item(oMat02.VisualRowCount).Specific.Value = CntcCode;
//        //						}
//        //						IsYN = "N";
//        //					}
//        //				}
//        //				oForm.Freeze(false);

//        //			}
//        //		}

//        //	} else if (pVal.BeforeAction == false) {
//        //		if (pVal.ItemUID == "PS_PP044") {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//        //			}
//        //		}
//        //		if (pVal.ItemUID == "1") {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//        //				if (pVal.ActionSuccess == true) {
//        //					if (oOrdGbn == "101" & oSequence == "1") {
//        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//        //						PS_PP044_FormItemEnabled();
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
//        //						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //					} else {
//        //						PS_PP044_FormItemEnabled();
//        //						PS_PP044_AddMatrixRow01(0, ref true);
//        //						////UDO방식일때
//        //						PS_PP044_AddMatrixRow02(0, ref true);
//        //						////UDO방식일때
//        //					}
//        //					//
//        //				}
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//        //				if (pVal.ActionSuccess == true) {
//        //					if ((oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)) {
//        //						oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
//        //						oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//        //						PS_PP044_FormItemEnabled();
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
//        //						oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //					}
//        //					PS_PP044_FormItemEnabled();
//        //				}
//        //			}
//        //		}
//        //		if (pVal.ItemUID == "Button01") {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) {
//        //			}
//        //		}
//        //	}
//        //	return;
//        //	Raise_EVENT_ITEM_PRESSED_Error:
//        //	oForm.Freeze(false);
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_KEY_DOWN
//        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	if (pVal.BeforeAction == true) {
//        //		if (pVal.ItemUID == "OrdMgNum") {
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			////작업타입이 일반,조정일때
//        //			if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm.Items.Item("OrdType").Specific.Selected.Value == "60") {
//        //				MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "OrdMgNum", "");
//        //				////사용자값활성
//        //			}
//        //		}
//        //		if (pVal.ItemUID == "Mat01") {
//        //			if (pVal.ColUID == "OrdMgNum") {
//        //				//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////일반,조정, 설계
//        //				if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm.Items.Item("OrdType").Specific.Selected.Value == "60" | oForm.Items.Item("OrdType").Specific.Selected.Value == "70") {
//        //					//UPGRADE_WARNING: oForm.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택") {
//        //						MDC_Com.MDC_GF_Message(ref "작업구분이 선택되지 않았습니다.", ref "W");
//        //						BubbleEvent = false;
//        //						return;
//        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택") {
//        //						MDC_Com.MDC_GF_Message(ref "사업장이 선택되지 않았습니다.", ref "W");
//        //						BubbleEvent = false;
//        //						return;
//        //						//UPGRADE_WARNING: oForm.Items(ItemCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value)) {
//        //						MDC_Com.MDC_GF_Message(ref "품목코드가 선택되지 않았습니다.", ref "W");
//        //						BubbleEvent = false;
//        //						return;
//        //						//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value)) {
//        //						MDC_Com.MDC_GF_Message(ref "작지번호가 선택되지 않았습니다.", ref "W");
//        //						BubbleEvent = false;
//        //						return;
//        //						//UPGRADE_WARNING: oForm.Items(PP030HNo).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (string.IsNullOrEmpty(oForm.Items.Item("PP030HNo").Specific.Value)) {
//        //						MDC_Com.MDC_GF_Message(ref "작지문서번호가 선택되지 않았습니다.", ref "W");
//        //						BubbleEvent = false;
//        //						return;
//        //					} else {
//        //						MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdMgNum");
//        //						////사용자값활성
//        //					}
//        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////지원
//        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {
//        //					//UPGRADE_WARNING: oForm.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (oForm.Items.Item("OrdGbn").Specific.Selected.Value == "선택") {
//        //						MDC_Com.MDC_GF_Message(ref "작업구분이 선택되지 않았습니다.", ref "W");
//        //						oForm.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //						BubbleEvent = false;
//        //						return;
//        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (oForm.Items.Item("BPLId").Specific.Selected.Value == "선택") {
//        //						MDC_Com.MDC_GF_Message(ref "사업장이 선택되지 않았습니다.", ref "W");
//        //						oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //						BubbleEvent = false;
//        //						return;
//        //						//                    ElseIf oForm.Items("ItemCode").Specific.Value = "" Then
//        //						//                        Call MDC_Com.MDC_GF_Message("품목코드가 선택되지 않았습니다.", "W")
//        //						//                        oForm.Items("ItemCode").Click ct_Regular
//        //						//                        BubbleEvent = False
//        //						//                        Exit Sub
//        //						//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value)) {
//        //						MDC_Com.MDC_GF_Message(ref "작지번호가 선택되지 않았습니다.", ref "W");
//        //						BubbleEvent = false;
//        //						return;
//        //					} else {
//        //						MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdMgNum");
//        //						////사용자값활성
//        //					}
//        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////외주
//        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {

//        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////실적
//        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {

//        //				}

//        //			}
//        //		}
//        //		if (pVal.ItemUID == "Mat02") {
//        //			if (pVal.ColUID == "WorkCode") {
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				if (Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value) == 0) {
//        //					MDC_Com.MDC_GF_Message(ref "기준시간을 입력하지 않았습니다.", ref "W");
//        //					oForm.Items.Item("BaseTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //					BubbleEvent = false;
//        //					return;
//        //				}
//        //			}
//        //		}
//        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "WorkCode");
//        //		////사용자값활성
//        //		MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "NCode");
//        //		////사용자값활성
//        //		MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat03", "FailCode");
//        //		////사용자값활성

//        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MachCode");
//        //		////설비코드 사용자값활성
//        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "Mat01", "SubLot") '//sub작지번호 사용자값활성
//        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CItemCod");
//        //		////원재료코드 사용자값활성
//        //		MDC_PS_Common.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "UseMCode", "");
//        //		////작업장비 사용자값활성
//        //		//        Call MDC_PS_Common.ActiveUserDefineValue(oForm, pVal, BubbleEvent, "Mat01", "ItemCode") '//사용자값활성
//        //	} else if (pVal.BeforeAction == false) {
//        //		//// 화살표 이동 강제 코딩 - 류영조
//        //		if (pVal.ItemUID == "Mat01") {
//        //			////위쪽 화살표
//        //			if (pVal.CharPressed == 38) {
//        //				if (pVal.Row > 1 & pVal.Row <= oMat01.VisualRowCount) {
//        //					oForm.Freeze(true);
//        //					oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //					oForm.Freeze(false);
//        //				}
//        //			////아래 화살표
//        //			} else if (pVal.CharPressed == 40) {
//        //				if (pVal.Row > 0 & pVal.Row < oMat01.VisualRowCount) {
//        //					oForm.Freeze(true);
//        //					oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //					oForm.Freeze(false);
//        //				}
//        //			}

//        //			//작업시간 입력 시마다 합계 계산(2011.09.26 송명규 추가)
//        //			if (pVal.ColUID == "WorkTime" & pVal.Row != Convert.ToDouble("0")) {

//        //				PS_PP044_SumWorkTime();

//        //			}

//        //		} else if (pVal.ItemUID == "BaseTime") {

//        //			//탭 키 Press
//        //			if (pVal.CharPressed == 9) {

//        //				oMat02.Columns.Item("WorkCode").Cells.Item(1).Click();

//        //			}

//        //		}
//        //	}
//        //	return;
//        //	Raise_EVENT_KEY_DOWN_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_COMBO_SELECT
//        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement


//        //	string OrdGbn = null;
//        //	string BPLID = null;
//        //	string DocDate = null;
//        //	string Gubun = null;

//        //	int i = 0;
//        //	string sQry = null;

//        //	int sCount = 0;
//        //	int sSeq = 0;
//        //	SAPbobsCOM.Recordset oRecordSet01 = null;

//        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        //	oForm.Freeze(true);
//        //	if (pVal.BeforeAction == true) {

//        //	} else if (pVal.BeforeAction == false) {
//        //		if (pVal.ItemChanged == true) {
//        //			oForm.Freeze(true);
//        //			if ((pVal.ItemUID == "Mat01")) {
//        //				if ((pVal.ColUID == "특정컬럼")) {
//        //					////기타작업
//        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
//        //					if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)))) {
//        //						//PS_PP044_AddMatrixRow (pVal.Row)
//        //					}
//        //				} else {
//        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
//        //				}
//        //			} else if ((pVal.ItemUID == "Mat02")) {
//        //				if ((pVal.ColUID == "특정컬럼")) {
//        //					////기타작업
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
//        //					if (oMat02.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP044M.GetValue("U_" + pVal.ColUID, pVal.Row - 1)))) {
//        //						//PS_PP044_AddMatrixRow (pVal.Row)
//        //					}
//        //				} else {
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
//        //				}
//        //			} else if ((pVal.ItemUID == "Mat03")) {
//        //				if ((pVal.ColUID == "특정컬럼")) {
//        //				} else {
//        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value);
//        //				}
//        //			} else {
//        //				if ((pVal.ItemUID == "OrdType")) {
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
//        //					//UPGRADE_WARNING: oForm.Items(pVal.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					////일반,조정,설계
//        //					if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "10" | oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "50" | oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "60" | oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "70") {
//        //						//창원은 품목구분 선택하도록 수정 '2015/04/09
//        //						//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if (oForm.Items.Item("BPLId").Specific.Value == "1") {
//        //							oForm.Items.Item("OrdGbn").Enabled = true;
//        //						} else {
//        //							oForm.Items.Item("OrdGbn").Enabled = false;
//        //						}
//        //						oForm.Items.Item("BPLId").Enabled = false;
//        //						oForm.Items.Item("ItemCode").Enabled = false;
//        //						//UPGRADE_WARNING: oForm.Items(pVal.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "20") {
//        //						oForm.Items.Item("OrdGbn").Enabled = true;
//        //						oForm.Items.Item("BPLId").Enabled = true;
//        //						oForm.Items.Item("ItemCode").Enabled = true;
//        //						//UPGRADE_WARNING: oForm.Items(pVal.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "30") {
//        //						oForm.Items.Item("OrdGbn").Enabled = false;
//        //						oForm.Items.Item("BPLId").Enabled = false;
//        //						oForm.Items.Item("ItemCode").Enabled = false;
//        //						//UPGRADE_WARNING: oForm.Items(pVal.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					} else if (oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value == "40") {
//        //						oForm.Items.Item("OrdGbn").Enabled = false;
//        //						oForm.Items.Item("BPLId").Enabled = false;
//        //						oForm.Items.Item("ItemCode").Enabled = false;
//        //					}

//        //					//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdGbn").Specific.Select("102", SAPbouiCOM.BoSearchKey.psk_ByValue);
//        //					//                    Call oForm.Items("OrdGbn").Specific.Select(0, psk_Index)
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdMgNum").Specific.Value = "";
//        //					//Call oForm.Items("BPLId").Specific.Select(0, psk_Index)
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("ItemCode").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("ItemName").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdNum").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdSub1").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdSub2").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("PP030HNo").Specific.Value = "";
//        //					oMat01.Clear();
//        //					oMat01.FlushToDataSource();
//        //					oMat01.LoadFromDataSource();
//        //					PS_PP044_AddMatrixRow01(0, ref true);
//        //					oMat02.Clear();
//        //					oMat02.FlushToDataSource();
//        //					oMat02.LoadFromDataSource();
//        //					PS_PP044_AddMatrixRow02(0, ref true);
//        //					oMat03.Clear();
//        //					oMat03.FlushToDataSource();
//        //					oMat03.LoadFromDataSource();
//        //				} else if ((pVal.ItemUID == "OrdGbn")) {
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);

//        //					//부품생산은 제품과 재공을 같이 작업할 수 있기 때문에 처리 2015/11/28 노근용 '배병관기사 요청
//        //					//                    oMat01.Clear
//        //					//                    oMat01.FlushToDataSource
//        //					//                    oMat01.LoadFromDataSource
//        //					//                    Call PS_PP044_AddMatrixRow01(0, True)
//        //					//                    oMat02.Clear
//        //					//                    oMat02.FlushToDataSource
//        //					//                    oMat02.LoadFromDataSource
//        //					//                    Call PS_PP044_AddMatrixRow02(0, True)
//        //					//                    oMat03.Clear
//        //					//                    oMat03.FlushToDataSource
//        //					//                    oMat03.LoadFromDataSource
//        //					//                    oMat04.Clear
//        //					//                    oMat04.FlushToDataSource
//        //					//                    oMat04.LoadFromDataSource
//        //					//                    oMat05.Clear
//        //					//                    oMat05.FlushToDataSource
//        //					//                    oMat05.LoadFromDataSource

//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					OrdGbn = Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value);

//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdMgNum").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdNum").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("ItemCode").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("PP030HNo").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("ItemName").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdSub1").Specific.Value = "";
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("OrdSub2").Specific.Value = "";


//        //					//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					sCount = oForm.Items.Item("Gubun").Specific.ValidValues.Count;
//        //					sSeq = sCount;
//        //					for (i = 1; i <= sCount; i++) {
//        //						//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oForm.Items.Item("Gubun").Specific.ValidValues.Remove(sSeq - 1, SAPbouiCOM.BoSearchKey.psk_Index);
//        //						sSeq = sSeq - 1;
//        //					}

//        //					sQry = "SELECT U_Minor, U_CdName From [@PS_SY001L] Where Code = 'P208' and U_RelCd = '" + OrdGbn + "' Order by U_Minor";
//        //					oRecordSet01.DoQuery(sQry);
//        //					//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("Gubun").Specific.ValidValues.Add("선택", "선택");
//        //					while (!(oRecordSet01.EoF)) {

//        //						//UPGRADE_WARNING: oForm.Items().Specific.ValidValues 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oForm.Items.Item("Gubun").Specific.ValidValues.Add(Strings.Trim(oRecordSet01.Fields.Item(0).Value), Strings.Trim(oRecordSet01.Fields.Item(1).Value));
//        //						oRecordSet01.MoveNext();
//        //					}
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("Gubun").Specific.Select("선택", SAPbouiCOM.BoSearchKey.psk_ByValue);
//        //				} else if ((pVal.ItemUID == "BPLId")) {
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
//        //					oMat01.Clear();
//        //					oMat01.FlushToDataSource();
//        //					oMat01.LoadFromDataSource();
//        //					PS_PP044_AddMatrixRow01(0, ref true);
//        //					oMat02.Clear();
//        //					oMat02.FlushToDataSource();
//        //					oMat02.LoadFromDataSource();
//        //					PS_PP044_AddMatrixRow02(0, ref true);
//        //					oMat03.Clear();
//        //					oMat03.FlushToDataSource();
//        //					oMat03.LoadFromDataSource();
//        //					oMat04.Clear();
//        //					oMat04.FlushToDataSource();
//        //					oMat04.LoadFromDataSource();

//        //					oMat05.Clear();
//        //					oMat05.FlushToDataSource();
//        //					oMat05.LoadFromDataSource();
//        //				} else if ((pVal.ItemUID == "Gubun")) {

//        //					PS_PP044_MTX01();
//        //					PS_PP044_MTX02();

//        //				} else {
//        //					//구분이 아닐 경우만 실행(2012.02.02 송명규 추가)
//        //					if (pVal.ItemUID != "Gubun") {
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value);
//        //					}
//        //				}
//        //			}
//        //			oMat01.LoadFromDataSource();
//        //			oMat01.AutoResizeColumns();
//        //			oMat02.LoadFromDataSource();
//        //			oMat02.AutoResizeColumns();
//        //			oMat03.LoadFromDataSource();
//        //			oMat03.AutoResizeColumns();
//        //			oMat04.LoadFromDataSource();
//        //			oMat04.AutoResizeColumns();
//        //			oMat05.LoadFromDataSource();
//        //			oMat05.AutoResizeColumns();
//        //			oForm.Update();
//        //			if (pVal.ItemUID == "Mat01") {
//        //				oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
//        //			} else if (pVal.ItemUID == "Mat02") {
//        //				oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
//        //			} else if (pVal.ItemUID == "Mat03") {
//        //				oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
//        //			} else if (pVal.ItemUID == "Mat04") {
//        //				oMat04.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
//        //			} else {

//        //			}
//        //			oForm.Freeze(false);
//        //		}
//        //	}
//        //	oForm.Freeze(false);
//        //	return;
//        //	Raise_EVENT_COMBO_SELECT_Error:
//        //	oForm.Freeze(false);
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_CLICK
//        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	object TempForm01 = null;

//        //	if (pVal.BeforeAction == true) {
//        //		if (pVal.ItemUID == "Opt01") {
//        //			oForm.Freeze(true);
//        //			oForm.Settings.MatrixUID = "Mat02";
//        //			oForm.Settings.EnableRowFormat = true;
//        //			oForm.Settings.Enabled = true;
//        //			oMat01.AutoResizeColumns();
//        //			oMat02.AutoResizeColumns();
//        //			oMat03.AutoResizeColumns();
//        //			oForm.Freeze(false);
//        //		}
//        //		if (pVal.ItemUID == "Opt02") {
//        //			oForm.Freeze(true);
//        //			oForm.Settings.MatrixUID = "Mat03";
//        //			oForm.Settings.EnableRowFormat = true;
//        //			oForm.Settings.Enabled = true;
//        //			oMat01.AutoResizeColumns();
//        //			oMat02.AutoResizeColumns();
//        //			oMat03.AutoResizeColumns();
//        //			oForm.Freeze(false);
//        //		}
//        //		if (pVal.ItemUID == "Opt03") {
//        //			oForm.Freeze(true);
//        //			oForm.Settings.MatrixUID = "Mat01";
//        //			oForm.Settings.EnableRowFormat = true;
//        //			oForm.Settings.Enabled = true;
//        //			oMat01.AutoResizeColumns();
//        //			oMat02.AutoResizeColumns();
//        //			oMat03.AutoResizeColumns();
//        //			oForm.Freeze(false);
//        //		}
//        //		//        If pVal.ItemUID = "Mat01" Then
//        //		//            If pVal.Row > 0 Then
//        //		//                Call oMat01.SelectRow(pVal.Row, True, False)
//        //		//            End If
//        //		//        End If
//        //		if (pVal.ItemUID == "Mat01") {
//        //			if (pVal.Row > 0) {
//        //				oMat01.SelectRow(pVal.Row, true, false);
//        //				oMat01Row01 = pVal.Row;
//        //			}
//        //		}
//        //		if (pVal.ItemUID == "Mat02") {
//        //			if (pVal.Row > 0) {
//        //				oMat02.SelectRow(pVal.Row, true, false);
//        //				oMat02Row02 = pVal.Row;
//        //			}
//        //		}
//        //		if (pVal.ItemUID == "Mat03") {
//        //			if (pVal.Row > 0) {
//        //				oMat03.SelectRow(pVal.Row, true, false);
//        //				oMat03Row03 = pVal.Row;
//        //			}
//        //		}
//        //	} else if (pVal.BeforeAction == false) {
//        //		//// 작업지시번호 링크 번튼 - 류영조
//        //		if (pVal.ItemUID == "LBtn01") {
//        //			TempForm01 = new PS_PP030();
//        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: TempForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			TempForm01.LoadForm(oForm.Items.Item("PP030HNo").Specific.Value);
//        //			//UPGRADE_NOTE: TempForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //			TempForm01 = null;
//        //		}
//        //	}
//        //	return;
//        //	Raise_EVENT_CLICK_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_DOUBLE_CLICK
//        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	if (pVal.BeforeAction == true) {
//        //		if (pVal.ItemUID == "Mat01") {
//        //			if (pVal.Row > 0) {
//        //				//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////작업타입이 일반,조정인경우
//        //				if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm.Items.Item("OrdType").Specific.Selected.Value == "60") {
//        //					//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value)) {

//        //					} else {
//        //						if (oMat03.VisualRowCount == 0) {
//        //							PS_PP044_AddMatrixRow03(0, ref true);
//        //						} else {
//        //							PS_PP044_AddMatrixRow03(oMat03.VisualRowCount);
//        //						}
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value);
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value);
//        //						oDS_PS_PP044N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));
//        //						oMat03.LoadFromDataSource();
//        //						oMat03.AutoResizeColumns();
//        //						//                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
//        //						//                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
//        //						oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
//        //						oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
//        //						oMat03.FlushToDataSource();
//        //					}
//        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////작업타입이 PSMT지원인경우
//        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {
//        //					//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value)) {

//        //					} else {
//        //						if (oMat03.VisualRowCount == 0) {
//        //							PS_PP044_AddMatrixRow03(0, ref true);
//        //						} else {
//        //							PS_PP044_AddMatrixRow03(oMat03.VisualRowCount);
//        //						}
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value);
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value);
//        //						oDS_PS_PP044N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));
//        //						oMat03.LoadFromDataSource();
//        //						oMat03.AutoResizeColumns();
//        //						//                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
//        //						//                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
//        //						oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
//        //						oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
//        //						oMat03.FlushToDataSource();
//        //					}
//        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////작업타입이 외주인경우
//        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {
//        //					//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////작업타입이 실적인경우
//        //				} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {
//        //				}
//        //			}
//        //		}
//        //	} else if (pVal.BeforeAction == false) {

//        //	}
//        //	return;
//        //	Raise_EVENT_DOUBLE_CLICK_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_MATRIX_LINK_PRESSED
//        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	object oTempClass = null;
//        //	if (pVal.BeforeAction == true) {
//        //		if (pVal.ItemUID == "Mat01") {
//        //			if (pVal.ColUID == "OrdMgNum") {
//        //				oTempClass = new PS_PP030();
//        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oTempClass.LoadForm(Strings.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, "-") - 1));
//        //			}
//        //			if (pVal.ColUID == "PP030HNo") {
//        //				oTempClass = new PS_PP030();
//        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oTempClass.LoadForm(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //			}
//        //		}
//        //		if (pVal.ItemUID == "Mat03") {
//        //			if (pVal.ColUID == "OrdMgNum") {
//        //				oTempClass = new PS_PP030();
//        //				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oTempClass.LoadForm(Strings.Mid(oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 1, Strings.InStr(oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, "-") - 1));
//        //			}
//        //		}
//        //	} else if (pVal.BeforeAction == false) {

//        //	}
//        //	return;
//        //	Raise_EVENT_MATRIX_LINK_PRESSED_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_VALIDATE
//        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	int i = 0;
//        //	string Query01 = null;
//        //	SAPbobsCOM.Recordset RecordSet01 = null;
//        //	double Weight = 0;

//        //	double Time = 0;
//        //	//UPGRADE_NOTE: Hour이(가) Hour_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//        //	int Hour_Renamed = 0;
//        //	//UPGRADE_NOTE: Minute이(가) Minute_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
//        //	int Minute_Renamed = 0;
//        //	string BPLID = null;
//        //	string OrdGbn = null;
//        //	string DocDate = null;
//        //	string Gubun = null;

//        //	oForm.Freeze(true);
//        //	SAPbouiCOM.ProgressBar ProgBar01 = null;
//        //	string WkCmDt = null;
//        //	if (pVal.BeforeAction == true) {
//        //		if (pVal.ItemChanged == true) {
//        //			if ((pVal.ItemUID == "Mat01")) {
//        //				if ((PS_PP044_Validate("수정01") == false)) {
//        //					oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Strings.Trim(oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)));
//        //				} else {
//        //					if ((pVal.ColUID == "OrdMgNum")) {
//        //						RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        //						ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("실행 중...", 100, false);

//        //						//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						////작지번호에 값이 없으면 작업지시가 불러오기전
//        //						if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value)) {
//        //							oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
//        //						////작업지시가 선택된상태
//        //						} else {
//        //							//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							////작업타입이 일반,조정, 설계
//        //							if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm.Items.Item("OrdType").Specific.Selected.Value == "60" | oForm.Items.Item("OrdType").Specific.Selected.Value == "70") {
//        //								////작지문서헤더번호가 일치하지 않으면
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								//UPGRADE_WARNING: oForm.Items(PP030HNo).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								if (oForm.Items.Item("PP030HNo").Specific.Value != Strings.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, "-") - 1)) {
//        //									oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
//        //								////작지문서번호가 일치하면
//        //								} else {
//        //									//UPGRADE_WARNING: oForm.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1") {
//        //										////신동사업부를 제외한 사업부만 체크
//        //										for (i = 1; i <= oMat01.RowCount; i++) {
//        //											////현재 입력한 값이 이미 입력되어 있는경우
//        //											//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //											//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //											if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value & i != pVal.Row) {
//        //												MDC_Com.MDC_GF_Message(ref "이미 입력한 공정입니다.", ref "W");
//        //												oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
//        //												goto Continue_Renamed;
//        //											}
//        //											//                                        '//공정라인의 공정순서가 앞공정보다 높으면
//        //											//                                        If Val(oMat01.Columns("Sequence").Cells(i).Specific.Value) >= MDC_PS_Common.GetValue("SELECT PS_PP030M.U_Sequence FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE CONVERT(NVARCHAR,PS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" & oMat01.Columns("OrdMgNum").Cells(pVal.Row).Specific.Value & "'") Then
//        //											//                                            Call MDC_Com.MDC_GF_Message("공정순서가 올바르지 않습니다.", "W")
//        //											//                                            Call oDS_PS_PP044L.setValue("U_" & pVal.ColUID, pVal.Row - 1, "")
//        //											//                                            GoTo Continue
//        //											//                                        End If
//        //										}

//        //										//생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_S
//        //										//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //										Query01 = "EXEC PS_PP040_90 '" + Strings.Mid(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, "-") - 1) + "'";
//        //										//oMat01.Columns("OrdMgNum").Cells(pVal.Row).Specific.Value & "'"
//        //										RecordSet01.DoQuery(Query01);
//        //										//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //										WkCmDt = RecordSet01.Fields.Item("WkCmDt").Value;

//        //										//생산완료수량이 작업지시수량만큼 모두 등록이 되었다면
//        //										if (RecordSet01.Fields.Item("Return").Value == "1") {
//        //											if (SubMain.Sbo_Application.MessageBox("생산완료가 모두 등록된 작번(완료일자:" + WkCmDt + ")입니다. 계속 진행하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1")) {
//        //												//계속 진행시에는 해당 작업지시문서번호 등록
//        //											} else {
//        //												//                                                Call MDC_Com.MDC_GF_Message("등록이 취소되었습니다.", "W")
//        //												oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
//        //												goto Continue_Renamed;
//        //											}
//        //										}
//        //										//생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_E

//        //									}

//        //									//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									Query01 = "EXEC PS_PP040_01 '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "', '" + oForm.Items.Item("OrdType").Specific.Selected.Value + "'";
//        //									RecordSet01.DoQuery(Query01);
//        //									if (RecordSet01.RecordCount == 0) {
//        //										oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
//        //									} else {
//        //										oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
//        //										oDS_PS_PP044L.SetValue("U_Sequence", pVal.Row - 1, RecordSet01.Fields.Item("Sequence").Value);
//        //										oDS_PS_PP044L.SetValue("U_CpCode", pVal.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
//        //										oDS_PS_PP044L.SetValue("U_CpName", pVal.Row - 1, RecordSet01.Fields.Item("CpName").Value);
//        //										oDS_PS_PP044L.SetValue("U_OrdGbn", pVal.Row - 1, RecordSet01.Fields.Item("OrdGbn").Value);
//        //										oDS_PS_PP044L.SetValue("U_BPLId", pVal.Row - 1, RecordSet01.Fields.Item("BPLId").Value);
//        //										oDS_PS_PP044L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
//        //										oDS_PS_PP044L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
//        //										oDS_PS_PP044L.SetValue("U_OrdNum", pVal.Row - 1, RecordSet01.Fields.Item("OrdNum").Value);
//        //										oDS_PS_PP044L.SetValue("U_OrdSub1", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub1").Value);
//        //										oDS_PS_PP044L.SetValue("U_OrdSub2", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub2").Value);
//        //										oDS_PS_PP044L.SetValue("U_PP030HNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030HNo").Value);
//        //										oDS_PS_PP044L.SetValue("U_PP030MNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030MNo").Value);
//        //										oDS_PS_PP044L.SetValue("U_SelWt", pVal.Row - 1, RecordSet01.Fields.Item("SelWt").Value);
//        //										oDS_PS_PP044L.SetValue("U_PSum", pVal.Row - 1, RecordSet01.Fields.Item("PSum").Value);
//        //										oDS_PS_PP044L.SetValue("U_BQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
//        //										oDS_PS_PP044L.SetValue("U_PQty", pVal.Row - 1, Convert.ToString(0));
//        //										oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(0));
//        //										oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(0));
//        //										oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(0));
//        //										oDS_PS_PP044L.SetValue("U_NQty", pVal.Row - 1, Convert.ToString(0));
//        //										oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(0));
//        //										oDS_PS_PP044L.SetValue("U_ScrapWt", pVal.Row - 1, Convert.ToString(0));
//        //										oDS_PS_PP044L.SetValue("U_WorkTime", pVal.Row - 1, Convert.ToString(0));
//        //										oDS_PS_PP044L.SetValue("U_LineId", pVal.Row - 1, "");

//        //										////설비코드,명 Reset
//        //										oDS_PS_PP044L.SetValue("U_MachCode", pVal.Row - 1, "");
//        //										oDS_PS_PP044L.SetValue("U_MachName", pVal.Row - 1, "");
//        //										////불량코드테이블
//        //										if (oMat03.VisualRowCount == 0) {
//        //											PS_PP044_AddMatrixRow03(0, ref true);
//        //										} else {
//        //											PS_PP044_AddMatrixRow03(oMat03.VisualRowCount);
//        //										}

//        //										oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
//        //										oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpCode").Value);
//        //										oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpName").Value);
//        //										oDS_PS_PP044N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));



//        //										//// 류영조
//        //										//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //										if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm.Items.Item("OrdType").Specific.Selected.Value == "60") {
//        //											oDS_PS_PP044H.SetValue("U_BaseTime", 0, "1");
//        //											//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //											oMat02.Columns.Item("WorkCode").Cells.Item(1).Specific.Value = "9999999";
//        //											//                                            oMat02.Columns("WorkName").Cells(1).Specific.Value = "조정"
//        //											//                                            Call oDS_PS_PP044M.setValue("U_WorkCode", 0, "9999999")
//        //											oDS_PS_PP044M.SetValue("U_WorkName", 0, "조정");
//        //											oMat02.LoadFromDataSource();
//        //										} else {
//        //											//                                            Call oDS_PS_PP044H.setValue("U_BaseTime", 0, "")
//        //											//                                            oMat02.Columns("WorkCode").Cells(1).Specific.Value = ""
//        //											//                                            oMat02.Columns("WorkName").Cells(1).Specific.Value = ""
//        //											//                        Call oDS_PS_PP044M.setValue("U_WorkCode", 0, "")
//        //											//                        Call oDS_PS_PP044M.setValue("U_WorkName", 0, "")
//        //										}
//        //									}
//        //								}
//        //								//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							////작업타입이 PSMT지원
//        //							} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {
//        //								////올바른 공정코드인지 검사
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [PS_PP001L] WHERE U_CpCode = ' & oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'") == 0) {
//        //									oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
//        //								} else {
//        //									for (i = 1; i <= oMat01.RowCount; i++) {
//        //										////현재 입력한 값이 이미 입력되어 있는경우
//        //										//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //										//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //										if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value & i != pVal.Row) {
//        //											MDC_Com.MDC_GF_Message(ref "이미 입력한 공정입니다.", ref "W");
//        //											oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
//        //											goto Continue_Renamed;
//        //										}
//        //									}
//        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044L.SetValue("U_CpCode", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044L.SetValue("U_CpName", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
//        //									//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044L.SetValue("U_OrdGbn", pVal.Row - 1, oForm.Items.Item("OrdGbn").Specific.Selected.Value);
//        //									//UPGRADE_WARNING: oForm.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044L.SetValue("U_BPLId", pVal.Row - 1, oForm.Items.Item("BPLId").Specific.Selected.Value);
//        //									oDS_PS_PP044L.SetValue("U_ItemCode", pVal.Row - 1, "");
//        //									oDS_PS_PP044L.SetValue("U_ItemName", pVal.Row - 1, "");
//        //									////PSMT지원은 품목코드 필요없음
//        //									//                                    Call oDS_PS_PP044L.setValue("U_ItemCode", pVal.Row - 1, oForm.Items("ItemCode").Specific.Value)
//        //									//                                    Call oDS_PS_PP044L.setValue("U_ItemName", pVal.Row - 1, oForm.Items("ItemName").Specific.Value)
//        //									//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044L.SetValue("U_OrdNum", pVal.Row - 1, oForm.Items.Item("OrdNum").Specific.Value);
//        //									//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044L.SetValue("U_OrdSub1", pVal.Row - 1, oForm.Items.Item("OrdSub1").Specific.Value);
//        //									//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044L.SetValue("U_OrdSub2", pVal.Row - 1, oForm.Items.Item("OrdSub2").Specific.Value);
//        //									oDS_PS_PP044L.SetValue("U_PP030HNo", pVal.Row - 1, "");
//        //									oDS_PS_PP044L.SetValue("U_PP030MNo", pVal.Row - 1, "");
//        //									oDS_PS_PP044L.SetValue("U_PSum", pVal.Row - 1, Convert.ToString(0));
//        //									oDS_PS_PP044L.SetValue("U_PQty", pVal.Row - 1, Convert.ToString(0));
//        //									oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(0));
//        //									oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(0));
//        //									oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(0));
//        //									oDS_PS_PP044L.SetValue("U_NQty", pVal.Row - 1, Convert.ToString(0));
//        //									oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(0));
//        //									oDS_PS_PP044L.SetValue("U_ScrapWt", pVal.Row - 1, Convert.ToString(0));
//        //									////불량코드테이블
//        //									if (oMat03.VisualRowCount == 0) {
//        //										PS_PP044_AddMatrixRow03(0, ref true);
//        //									} else {
//        //										if (oDS_PS_PP044L.GetValue("U_OrdMgNum", pVal.Row - 1) == oDS_PS_PP044N.GetValue("U_OrdMgNum", oMat03.VisualRowCount - 1)) {
//        //										} else {
//        //											PS_PP044_AddMatrixRow03(oMat03.VisualRowCount);
//        //										}
//        //									}
//        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
//        //								}
//        //								//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							////작업타입이 외주
//        //							} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {

//        //								//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							////작업타입이 실적
//        //							} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {

//        //							}
//        //							Continue_Renamed:
//        //							if (oMat01.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1)))) {
//        //								PS_PP044_AddMatrixRow01(pVal.Row);
//        //							}
//        //						}

//        //						ProgBar01.Value = 100;
//        //						ProgBar01.Stop();
//        //						//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //						ProgBar01 = null;

//        //						//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //						RecordSet01 = null;
//        //					} else if (pVal.ColUID == "PQty") {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if (Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0) {
//        //							if (Strings.Trim(oDS_PS_PP044H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP044H.GetValue("U_OrdType", 0)) == "60") {
//        //								goto Skip_PQty;
//        //							} else {
//        //								oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
//        //							}
//        //						} else {
//        //							Skip_PQty:
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //							//UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							Weight = Conversion.Val(MDC_PS_Common.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
//        //							if (Weight == 0) {
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //							} else {
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //							}
//        //							oDS_PS_PP044L.SetValue("U_NQty", pVal.Row - 1, Convert.ToString(0));
//        //							oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(0));
//        //						}
//        //					} else if (pVal.ColUID == "NQty") {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if (Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0) {
//        //							if (Strings.Trim(oDS_PS_PP044H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP044H.GetValue("U_OrdType", 0)) == "60") {
//        //								goto skip_Nqty;
//        //							} else {
//        //								oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
//        //							}
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						} else if (Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value)) {
//        //							if (Strings.Trim(oDS_PS_PP044H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP044H.GetValue("U_OrdType", 0)) == "60") {
//        //								goto skip_Nqty;
//        //							} else {
//        //								oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP044L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
//        //							}
//        //						} else {
//        //							skip_Nqty:
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							oDS_PS_PP044L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //							//UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							Weight = Conversion.Val(MDC_PS_Common.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) / 1000;
//        //							if (Weight == 0) {
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //							} else {
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								oDS_PS_PP044L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Weight * (Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
//        //							}
//        //						}
//        //					} else if (pVal.ColUID == "WorkTime") {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //					////기존도면매수
//        //					} else if (pVal.ColUID == "BdwQty") {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_AdwQTy", pVal.Row - 1, Convert.ToString((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
//        //					////도면 적용율
//        //					} else if (pVal.ColUID == "DwRate") {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_AdwQTy", pVal.Row - 1, Convert.ToString((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
//        //					////신규도면매수
//        //					} else if (pVal.ColUID == "NdwQTy") {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Conversion.Val(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //					} else if (pVal.ColUID == "MachCode") {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_MachName", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
//        //						//Call oDS_PS_PP044L.setValue("U_MachGrCd", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT U_MacdGrCd FROM [@PS_PP130H] WHERE U_MachCode = '" & oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value & "'", 0, 1))
//        //					} else if (pVal.ColUID == "CItemCod") {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //						//UPGRADE_WARNING: oMat01.Columns(pVal.ColUID).Cells(pVal.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_CItemNam", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT U_ItemNam2 FROM [@PS_PP005H] WHERE U_ItemCod1 = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' and U_ItemCod2 = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
//        //					} else {
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //					}
//        //				}
//        //			} else if ((pVal.ItemUID == "Mat02")) {
//        //				if ((pVal.ColUID == "WorkCode")) {
//        //					////기타작업
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_WorkName", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
//        //					if (oMat02.RowCount == pVal.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP044M.GetValue("U_" + pVal.ColUID, pVal.Row - 1)))) {
//        //						PS_PP044_AddMatrixRow02(pVal.Row);
//        //					}
//        //				} else if (pVal.ColUID == "NStart") {
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 | Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0) {
//        //						oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(0));
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value)));
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value)));
//        //					} else {
//        //						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value)) {
//        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							Time = Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
//        //						} else {
//        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							Time = (2400 - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
//        //						}
//        //						Hour_Renamed = Conversion.Fix(Time / 100);
//        //						//UPGRADE_WARNING: Mod에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
//        //						Minute_Renamed = Time % 100;
//        //						Time = Hour_Renamed;
//        //						if (Minute_Renamed > 0) {
//        //							Time = Time + 0.5;
//        //						}
//        //						oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(Time));
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value) - Time));
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value) - Time));
//        //					}
//        //				} else if (pVal.ColUID == "NEnd") {
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 | Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0) {
//        //						oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(0));
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value)));
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value)));
//        //					} else {
//        //						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value)) {
//        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							Time = Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
//        //						} else {
//        //							//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							Time = (2400 - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
//        //						}
//        //						Hour_Renamed = Conversion.Fix(Time / 100);
//        //						//UPGRADE_WARNING: Mod에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
//        //						Minute_Renamed = Time % 100;
//        //						Time = Hour_Renamed;
//        //						if (Minute_Renamed > 0) {
//        //							Time = Time + 0.5;
//        //						}
//        //						oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(Time));
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value) - Time));
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oForm.Items.Item("BaseTime").Specific.Value) - Time));
//        //					}
//        //				} else if (pVal.ColUID == "YTime") {
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (Conversion.Val(oMat02.Columns.Item("TTime").Cells.Item(pVal.Row).Specific.Value) > 0) {
//        //						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item("TTime").Cells.Item(pVal.Row).Specific.Value) - Conversion.Val(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //					}
//        //				} else if (pVal.ColUID == "NTime") {
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (Conversion.Val(oMat02.Columns.Item("TTime").Cells.Item(pVal.Row).Specific.Value) > 0) {
//        //						//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item("TTime").Cells.Item(pVal.Row).Specific.Value) - Conversion.Val(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
//        //					}
//        //				} else {
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //				}
//        //			} else if ((pVal.ItemUID == "Mat03")) {
//        //				if ((pVal.ColUID == "FailCode")) {
//        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044N.SetValue("U_FailName", pVal.Row - 1, MDC_PS_Common.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
//        //				} else {
//        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
//        //				}
//        //			} else {
//        //				if ((pVal.ItemUID == "DocEntry")) {
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
//        //				} else if ((pVal.ItemUID == "BaseTime")) {
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, Convert.ToString(Conversion.Val(oForm.Items.Item(pVal.ItemUID).Specific.Value)));
//        //				} else if ((pVal.ItemUID == "OrdMgNum")) {
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
//        //					if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) {
//        //						PS_PP044_OrderInfoLoad();
//        //					}
//        //				} else if ((pVal.ItemUID == "ItemCode")) {
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
//        //					oMat01.Clear();
//        //					oMat01.FlushToDataSource();
//        //					oMat01.LoadFromDataSource();
//        //					PS_PP044_AddMatrixRow01(0, ref true);
//        //					oMat02.Clear();
//        //					oMat02.FlushToDataSource();
//        //					oMat02.LoadFromDataSource();
//        //					PS_PP044_AddMatrixRow02(0, ref true);
//        //					oMat03.Clear();
//        //					oMat03.FlushToDataSource();
//        //					oMat03.LoadFromDataSource();

//        //				} else if ((pVal.ItemUID == "UseMCode")) {

//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					Query01 = "EXEC PS_PP040_98 '" + oForm.Items.Item("UseMCode").Specific.Value;

//        //					RecordSet01.DoQuery(Query01);

//        //					//UPGRADE_WARNING: oForm.Items(UseMName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oForm.Items.Item("UseMName").Specific.Value = Strings.Trim(RecordSet01.Fields.Item(0).Value);
//        //				} else if ((pVal.ItemUID == "DocDate")) {
//        //					PS_PP044_MTX01();
//        //					PS_PP044_MTX02();
//        //				} else {
//        //					//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oDS_PS_PP044H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
//        //				}
//        //			}
//        //			oMat01.LoadFromDataSource();
//        //			oMat01.AutoResizeColumns();
//        //			oMat02.LoadFromDataSource();
//        //			oMat02.AutoResizeColumns();
//        //			oMat03.LoadFromDataSource();
//        //			oMat03.AutoResizeColumns();
//        //			oForm.Update();
//        //			if (pVal.ItemUID == "Mat01") {
//        //				oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //			} else if (pVal.ItemUID == "Mat02") {
//        //				oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //			} else if (pVal.ItemUID == "Mat03") {
//        //				oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //			} else {
//        //				oForm.Items.Item(pVal.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //			}
//        //		}
//        //	} else if (pVal.BeforeAction == false) {

//        //	}
//        //	oForm.Freeze(false);
//        //	return;
//        //	Raise_EVENT_VALIDATE_Error:
//        //	oForm.Freeze(false);
//        //	ProgBar01.Value = 100;
//        //	ProgBar01.Stop();
//        //	//UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	ProgBar01 = null;
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_MATRIX_LOAD
//        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	if (pVal.BeforeAction == true) {

//        //	} else if (pVal.BeforeAction == false) {
//        //		PS_PP044_FormItemEnabled();
//        //		if (pVal.ItemUID == "Mat01") {
//        //			PS_PP044_AddMatrixRow01(oMat01.VisualRowCount);
//        //			////UDO방식
//        //		} else if (pVal.ItemUID == "Mat02") {
//        //			PS_PP044_AddMatrixRow02(oMat02.VisualRowCount);
//        //			////UDO방식
//        //		}
//        //	}
//        //	return;
//        //	Raise_EVENT_MATRIX_LOAD_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_RESIZE
//        //private void Raise_EVENT_RESIZE(ref object FormUID = null, ref SAPbouiCOM.ItemEvent pVal = null, ref bool BubbleEvent = false)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	if (pVal.BeforeAction == true) {

//        //	} else if (pVal.BeforeAction == false) {
//        //		PS_PP044_FormResize();
//        //	}
//        //	return;
//        //	Raise_EVENT_RESIZE_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RESIZE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_CHOOSE_FROM_LIST
//        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	SAPbouiCOM.DataTable oDataTable01 = null;
//        //	if (pVal.BeforeAction == true) {

//        //	} else if (pVal.BeforeAction == false) {
//        //		//        If (pVal.ItemUID = "ItemCode") Then
//        //		//            Dim oDataTable01 As SAPbouiCOM.DataTable
//        //		//            Set oDataTable01 = pVal.SelectedObjects
//        //		//            oForm.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
//        //		//            Set oDataTable01 = Nothing
//        //		//        End If
//        //		//        If (pVal.ItemUID = "CardCode" Or pVal.ItemUID = "CardName") Then
//        //		//            Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP040H", "U_CardCode,U_CardName")
//        //		//        End If
//        //		if ((pVal.ItemUID == "ItemCode")) {
//        //			//UPGRADE_WARNING: pVal.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if (pVal.SelectedObjects == null) {
//        //			} else {
//        //				MDC_Com.MDC_GP_CF_DBDatasourceReturn(pVal, (pVal.FormUID), "@PS_PP040H", "U_ItemCode,U_ItemName");
//        //				oMat01.Clear();
//        //				oMat01.FlushToDataSource();
//        //				oMat01.LoadFromDataSource();
//        //				PS_PP044_AddMatrixRow01(0, ref true);
//        //				oMat02.Clear();
//        //				oMat02.FlushToDataSource();
//        //				oMat02.LoadFromDataSource();
//        //				PS_PP044_AddMatrixRow02(0, ref true);
//        //				oMat03.Clear();
//        //				oMat03.FlushToDataSource();
//        //				oMat03.LoadFromDataSource();
//        //			}
//        //		}
//        //		//        If (pVal.ItemUID = "Mat02") Then
//        //		//            If (pVal.ColUID = "WorkCode") Then
//        //		//                If pVal.SelectedObjects Is Nothing Then
//        //		//                Else
//        //		//                    Set oDataTable01 = pVal.SelectedObjects
//        //		//                    Call oDS_PS_PP044M.setValue("U_WorkCode", pVal.Row - 1, oDataTable01.Columns("empID").Cells(0).Value)
//        //		//                    Call oDS_PS_PP044M.setValue("U_WorkName", pVal.Row - 1, oDataTable01.Columns("firstName").Cells(0).Value & oDataTable01.Columns("lastName").Cells(0).Value)
//        //		//                    If oMat02.RowCount = pVal.Row And Trim(oDS_PS_PP044M.GetValue("U_" & pVal.ColUID, pVal.Row - 1)) <> "" Then
//        //		//                        Call PS_PP044_AddMatrixRow02(pVal.Row)
//        //		//                    End If
//        //		//                    Set oDataTable01 = Nothing
//        //		//                    'Call MDC_GP_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
//        //		//                    oMat02.LoadFromDataSource
//        //		//                    oMat02.Columns(pVal.ColUID).Cells(pVal.Row).Click ct_Regular
//        //		//                End If
//        //		//            End If
//        //		//        End If
//        //	}
//        //	return;
//        //	Raise_EVENT_CHOOSE_FROM_LIST_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_GOT_FOCUS
//        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	if (pVal.ItemUID == "Mat01" | pVal.ItemUID == "Mat02" | pVal.ItemUID == "Mat03") {
//        //		if (pVal.Row > 0) {
//        //			oLastItemUID01 = pVal.ItemUID;
//        //			oLastColUID01 = pVal.ColUID;
//        //			oLastColRow01 = pVal.Row;
//        //		}
//        //	} else {
//        //		oLastItemUID01 = pVal.ItemUID;
//        //		oLastColUID01 = "";
//        //		oLastColRow01 = 0;
//        //	}
//        //	if (pVal.ItemUID == "Mat01") {
//        //		if (pVal.Row > 0) {
//        //			oMat01Row01 = pVal.Row;
//        //		}
//        //	} else if (pVal.ItemUID == "Mat02") {
//        //		if (pVal.Row > 0) {
//        //			oMat02Row02 = pVal.Row;
//        //		}
//        //	} else if (pVal.ItemUID == "Mat03") {
//        //		if (pVal.Row > 0) {
//        //			oMat03Row03 = pVal.Row;
//        //		}
//        //	}
//        //	return;
//        //	Raise_EVENT_GOT_FOCUS_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_FORM_UNLOAD
//        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	if (pVal.BeforeAction == true) {
//        //	} else if (pVal.BeforeAction == false) {
//        //		SubMain.RemoveForms(oFormUniqueID);
//        //		//UPGRADE_NOTE: oForm 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //		oForm = null;
//        //		//UPGRADE_NOTE: oMat01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //		oMat01 = null;
//        //	}
//        //	return;
//        //	Raise_EVENT_FORM_UNLOAD_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_ROW_DELETE
//        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	object i = null;
//        //	int j = 0;
//        //	bool Exist = false;
//        //	if ((oLastColRow01 > 0)) {
//        //		if (pVal.BeforeAction == true) {
//        //			if (oLastItemUID01 == "Mat01") {
//        //				if ((PS_PP044_Validate("행삭제01") == false)) {
//        //					BubbleEvent = false;
//        //					return;
//        //				}
//        //				Continue_Renamed:
//        //				for (i = 1; i <= oMat03.RowCount; i++) {
//        //					//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: oMat01.Columns(LineNum).Cells(oLastColRow01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: oMat03.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(oLastColRow01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (oMat01.Columns.Item("OrdMgNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value & oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value) {
//        //						////If oMat01.Columns("OrdMgNum").Cells(oLastColRow01).Specific.Value = oMat03.Columns("OrdMgNum").Cells(i).Specific.Value Then
//        //						//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.RemoveRecord((i - 1));
//        //						oMat03.DeleteRow((i));
//        //						oMat03.FlushToDataSource();
//        //						goto Continue_Renamed;
//        //					}
//        //				}
//        //			}
//        //			////행삭제전 행삭제가능여부검사
//        //		} else if (pVal.BeforeAction == false) {
//        //			if (oLastItemUID01 == "Mat01") {
//        //				for (i = 1; i <= oMat01.VisualRowCount; i++) {
//        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
//        //				}

//        //				for (i = 1; i <= oMat03.VisualRowCount; i++) {
//        //					//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value != 1) {
//        //						//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value = oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value - 1;
//        //						////i
//        //					}
//        //				}

//        //				oMat01.FlushToDataSource();
//        //				oDS_PS_PP044L.RemoveRecord(oDS_PS_PP044L.Size - 1);
//        //				oMat01.LoadFromDataSource();
//        //				if (oMat01.RowCount == 0) {
//        //					PS_PP044_AddMatrixRow01(0);
//        //				} else {
//        //					if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP044L.GetValue("U_OrdMgNum", oMat01.RowCount - 1)))) {
//        //						PS_PP044_AddMatrixRow01(oMat01.RowCount);
//        //					}
//        //				}
//        //			} else if (oLastItemUID01 == "Mat02") {
//        //				for (i = 1; i <= oMat02.VisualRowCount; i++) {
//        //					//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
//        //				}
//        //				oMat02.FlushToDataSource();
//        //				oDS_PS_PP044M.RemoveRecord(oDS_PS_PP044M.Size - 1);
//        //				oMat02.LoadFromDataSource();
//        //				if (oMat02.RowCount == 0) {
//        //					PS_PP044_AddMatrixRow02(0);
//        //				} else {
//        //					if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP044M.GetValue("U_WorkCode", oMat02.RowCount - 1)))) {
//        //						PS_PP044_AddMatrixRow02(oMat02.RowCount);
//        //					}
//        //				}
//        //			} else if (oLastItemUID01 == "Mat03") {
//        //				for (i = 1; i <= oMat03.VisualRowCount; i++) {
//        //					//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
//        //				}
//        //				oMat03.FlushToDataSource();
//        //				////사이즈가 0일때는 행을 빼주면 oMat03.VisualRowCount 가 0 으로 변경되어서 문제가 생김
//        //				if (oDS_PS_PP044N.Size == 1) {
//        //				} else {
//        //					oDS_PS_PP044N.RemoveRecord(oDS_PS_PP044N.Size - 1);
//        //				}
//        //				oMat03.LoadFromDataSource();

//        //				////공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
//        //				for (i = 1; i <= oMat01.RowCount - 1; i++) {
//        //					Exist = false;
//        //					for (j = 1; j <= oMat03.RowCount; j++) {
//        //						//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oMat01.Columns(LineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oMat03.Columns(OrdMgNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value & oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.Value) {
//        //							////If oMat01.Columns("OrdMgNum").Cells(i).Specific.Value = oMat03.Columns("OrdMgNum").Cells(j).Specific.Value Then
//        //							Exist = true;
//        //						}
//        //					}
//        //					////불량코드테이블에 값이 존재하지 않으면
//        //					if (Exist == false) {
//        //						if (oMat03.VisualRowCount == 0) {
//        //							PS_PP044_AddMatrixRow03(0, ref true);
//        //						} else {
//        //							PS_PP044_AddMatrixRow03(oMat03.VisualRowCount);
//        //						}
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value);
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value);
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(i).Specific.Value);
//        //						//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						oDS_PS_PP044N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, i);
//        //						oMat03.LoadFromDataSource();
//        //						oMat03.AutoResizeColumns();
//        //						oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
//        //						oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
//        //						//                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
//        //						//                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
//        //						oMat03.FlushToDataSource();
//        //					}
//        //				}
//        //			}
//        //		}
//        //	}
//        //	return;
//        //	Raise_EVENT_ROW_DELETE_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region Raise_EVENT_RECORD_MOVE
//        //private void Raise_EVENT_RECORD_MOVE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pVal, ref bool BubbleEvent)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	string Query01 = null;
//        //	SAPbobsCOM.Recordset RecordSet01 = null;
//        //	string DocEntry = null;
//        //	string DocEntryNext = null;
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
//        //	////원본문서
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	DocEntryNext = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
//        //	////다음문서

//        //	////다음
//        //	if (pVal.MenuUID == "1288") {
//        //		if (pVal.BeforeAction == true) {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {

//        //				SubMain.Sbo_Application.ActivateMenuItem(("1290"));
//        //				BubbleEvent = false;
//        //				return;
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				if ((string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))) {
//        //					SubMain.Sbo_Application.ActivateMenuItem(("1290"));
//        //					BubbleEvent = false;
//        //					return;
//        //				}
//        //			}
//        //			if (PS_PP044_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") == false) {
//        //				BubbleEvent = false;
//        //				return;
//        //			}
//        //		}
//        //	////이전
//        //	} else if (pVal.MenuUID == "1289") {
//        //		if (pVal.BeforeAction == true) {
//        //			if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE) {
//        //				SubMain.Sbo_Application.ActivateMenuItem(("1291"));
//        //				BubbleEvent = false;
//        //				return;
//        //			} else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) {
//        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				if ((string.IsNullOrEmpty(oForm.Items.Item("DocEntry").Specific.Value))) {
//        //					SubMain.Sbo_Application.ActivateMenuItem(("1291"));
//        //					BubbleEvent = false;
//        //					return;
//        //				}
//        //			}
//        //			if (PS_PP044_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") == false) {
//        //				BubbleEvent = false;
//        //				return;
//        //			}
//        //		}
//        //	////첫번째레코드로이동
//        //	} else if (pVal.MenuUID == "1290") {
//        //		if (pVal.BeforeAction == true) {
//        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//        //			Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry DESC";
//        //			////가장마지막행을 부여
//        //			RecordSet01.DoQuery(Query01);
//        //			DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
//        //			////원본문서
//        //			DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
//        //			////다음문서
//        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //			RecordSet01 = null;
//        //			if (PS_PP044_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") == false) {
//        //				BubbleEvent = false;
//        //				return;
//        //			}
//        //		}
//        //	////마지막문서로이동
//        //	} else if (pVal.MenuUID == "1291") {
//        //		if (pVal.BeforeAction == true) {
//        //			RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//        //			Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry ASC";
//        //			////가장 첫행을 부여
//        //			RecordSet01.DoQuery(Query01);
//        //			DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
//        //			////원본문서
//        //			DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
//        //			////다음문서
//        //			//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //			RecordSet01 = null;
//        //			if (PS_PP044_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") == false) {
//        //				BubbleEvent = false;
//        //				return;
//        //			}
//        //		}
//        //	}
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	return;
//        //	Raise_EVENT_RECORD_MOVE_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RECORD_MOVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion








//        #region PS_PP044_AddMatrixRow01
//        //public void PS_PP044_AddMatrixRow01(int oRow, ref bool RowIserted = false)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	oForm.Freeze(true);
//        //	////행추가여부
//        //	if (RowIserted == false) {
//        //		oDS_PS_PP044L.InsertRecord((oRow));
//        //	}
//        //	oMat01.AddRow();
//        //	oDS_PS_PP044L.Offset = oRow;
//        //	oDS_PS_PP044L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//        //	oDS_PS_PP044L.SetValue("U_WorkCls", oRow, "A");
//        //	//작업구분을 기본으로 선택(2014.04.15 송명규 추가)
//        //	oMat01.LoadFromDataSource();
//        //	oForm.Freeze(false);
//        //	return;
//        //	PS_PP044_AddMatrixRow01_Error:
//        //	oForm.Freeze(false);
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_AddMatrixRow01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region PS_PP044_AddMatrixRow02
//        //public void PS_PP044_AddMatrixRow02(int oRow, ref bool RowIserted = false)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	oForm.Freeze(true);
//        //	////행추가여부
//        //	if (RowIserted == false) {
//        //		oDS_PS_PP044M.InsertRecord((oRow));
//        //	}
//        //	oMat02.AddRow();
//        //	oDS_PS_PP044M.Offset = oRow;
//        //	oDS_PS_PP044M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//        //	oMat02.LoadFromDataSource();
//        //	oForm.Freeze(false);
//        //	return;
//        //	PS_PP044_AddMatrixRow02_Error:
//        //	oForm.Freeze(false);
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_AddMatrixRow02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region PS_PP044_AddMatrixRow03
//        //public void PS_PP044_AddMatrixRow03(int oRow, ref bool RowIserted = false)
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	oForm.Freeze(true);
//        //	////행추가여부
//        //	if (RowIserted == false) {
//        //		oDS_PS_PP044N.InsertRecord((oRow));
//        //	}
//        //	oMat03.AddRow();
//        //	oDS_PS_PP044N.Offset = oRow;
//        //	oDS_PS_PP044N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
//        //	oMat03.LoadFromDataSource();
//        //	oForm.Freeze(false);
//        //	return;
//        //	PS_PP044_AddMatrixRow03_Error:
//        //	oForm.Freeze(false);
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_AddMatrixRow03_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion



//        #region PS_PP044_DataValidCheck
//        //public bool PS_PP044_DataValidCheck()
//        //{
//        //	bool functionReturnValue = false;
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	functionReturnValue = false;
//        //	object i = null;
//        //	int j = 0;
//        //	double FailQty = 0;
//        //	double sYTime = 0;
//        //	double sNTime = 0;
//        //	double sTTime = 0;

//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	//UPGRADE_WARNING: MDC_PS_Common.GetValue(select Count(*) from OFPR Where ' & oForm.Items(DocDate).Specific.Value & ' between F_RefDate and T_RefDate And PeriodStat = 'Y') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	if (MDC_PS_Common.GetValue("select Count(*) from OFPR Where '" + oForm.Items.Item("DocDate").Specific.Value + "' between F_RefDate and T_RefDate And PeriodStat = 'Y'") > 0) {
//        //		SubMain.Sbo_Application.SetStatusBarMessage("해당일자는 전기기간이 잠겼습니다. 일자를 확인바랍니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		functionReturnValue = false;
//        //		return functionReturnValue;
//        //	}
//        //	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//        //		PS_PP044_FormClear();
//        //	}
//        //	//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	if (oForm.Items.Item("OrdType").Specific.Selected.Value != "10" & oForm.Items.Item("OrdType").Specific.Selected.Value != "20" & oForm.Items.Item("OrdType").Specific.Selected.Value != "50" & oForm.Items.Item("OrdType").Specific.Selected.Value != "60" & oForm.Items.Item("OrdType").Specific.Selected.Value != "70") {
//        //		SubMain.Sbo_Application.SetStatusBarMessage("작업타입이 일반, PSMT지원, 조정, 설계가 아닙니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		functionReturnValue = false;
//        //		return functionReturnValue;
//        //	}

//        //	//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value)) {
//        //		SubMain.Sbo_Application.SetStatusBarMessage("작지번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		oForm.Items.Item("OrdNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //		functionReturnValue = false;
//        //		return functionReturnValue;
//        //	}

//        //	if (oMat01.VisualRowCount == 1) {
//        //		SubMain.Sbo_Application.SetStatusBarMessage("공정정보 라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		functionReturnValue = false;
//        //		return functionReturnValue;
//        //	}
//        //	if (oMat02.VisualRowCount == 1) {
//        //		SubMain.Sbo_Application.SetStatusBarMessage("작업자정보 라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		functionReturnValue = false;
//        //		return functionReturnValue;
//        //	}

//        //	// 작업자 1명이상 가능토록 수정 (이병각)
//        //	//    If oMat02.VisualRowCount > 2 Then '//한명이상 입력했을경우
//        //	//        If oForm.Items("OrdGbn").Specific.Selected.Value = "106" Then '//몰드
//        //	//            Sbo_Application.SetStatusBarMessage "작업자정보 한명만 입력할수 있습니다.", bmt_Short, True
//        //	//            PS_PP044_DataValidCheck = False
//        //	//            Exit Function
//        //	//        Else
//        //	//            '//휘팅,부품은 여러명 입력할수 있다.
//        //	//        End If
//        //	//    End If

//        //	for (j = 1; j <= oMat02.VisualRowCount; j++) {
//        //		//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		sYTime = sYTime + Conversion.Val(oMat02.Columns.Item("YTime").Cells.Item(j).Specific.Value);
//        //		//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		sNTime = sNTime + Conversion.Val(oMat02.Columns.Item("NTime").Cells.Item(j).Specific.Value);
//        //		//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		sTTime = sNTime + Conversion.Val(oMat02.Columns.Item("TTime").Cells.Item(j).Specific.Value);
//        //	}

//        //	if (oMat03.VisualRowCount == 0) {
//        //		SubMain.Sbo_Application.SetStatusBarMessage("불량정보 라인이 존재하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		functionReturnValue = false;
//        //		return functionReturnValue;
//        //	}

//        //	//마감상태 체크_S(2017.11.23 송명규 추가)
//        //	//UPGRADE_WARNING: oForm.Items(DocDate).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	if (MDC_PS_Common.Check_Finish_Status(Strings.Trim(oForm.Items.Item("BPLId").Specific.Value), oForm.Items.Item("DocDate").Specific.Value, oForm.TypeEx) == false) {
//        //		SubMain.Sbo_Application.SetStatusBarMessage("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작업일보일자를 확인하고, 회계부서로 문의하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		functionReturnValue = false;
//        //		return functionReturnValue;
//        //	}
//        //	//마감상태 체크_E(2017.11.23 송명규 추가)

//        //	for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
//        //		//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if ((string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value))) {
//        //			SubMain.Sbo_Application.SetStatusBarMessage("작지문서번호는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //			oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //			functionReturnValue = false;
//        //			return functionReturnValue;
//        //		}

//        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (Strings.Trim(oForm.Items.Item("OrdType").Specific.Value) != "50" & Strings.Trim(oForm.Items.Item("OrdType").Specific.Value) != "60") {
//        //			//작업시간이 0 보다 클때
//        //			if (sYTime > 0) {
//        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				if ((Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) <= 0)) {
//        //					SubMain.Sbo_Application.SetStatusBarMessage("생산수량은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //					oMat01.Columns.Item("PQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //					functionReturnValue = false;
//        //					return functionReturnValue;
//        //				}
//        //			}
//        //		}

//        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (Strings.Trim(oForm.Items.Item("OrdType").Specific.Value) != "50" & Strings.Trim(oForm.Items.Item("OrdType").Specific.Value) != "60" & Strings.Trim(oForm.Items.Item("OrdType").Specific.Value) != "70") {
//        //			//작업시간이 0 보다 클때
//        //			if (sYTime > 0) {
//        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				if ((Conversion.Val(oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value) <= 0)) {
//        //					SubMain.Sbo_Application.SetStatusBarMessage("실동시간은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //					oMat01.Columns.Item("WorkTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //					functionReturnValue = false;
//        //					return functionReturnValue;
//        //				}
//        //			}
//        //		}

//        //		//작업완료여부(2012.02.02. 송명규 추가)
//        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		//기계공구, 몰드일 경우만 작업완료여부 필수 체크
//        //		if (Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value) == "105" | Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value) == "106") {

//        //			//UPGRADE_WARNING: oMat01.Columns(CompltYN).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if ((oMat01.Columns.Item("CompltYN").Cells.Item(i).Specific.Value == "%")) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("작업구분이 기계공구, 몰드일경우는 작업완료여부가 필수입니다. 확인하십시오.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				oMat01.Columns.Item("CompltYN").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //				functionReturnValue = false;
//        //				return functionReturnValue;
//        //			}

//        //		}

//        //		////불량수량 검사
//        //		FailQty = 0;
//        //		for (j = 1; j <= oMat03.VisualRowCount; j++) {
//        //			////불량코드를 입력했는지 check
//        //			//UPGRADE_WARNING: oMat03.Columns(FailCode).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if (Conversion.Val(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) != 0 & string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value)) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("불량수량이 입력되었을 때는 불량코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //				functionReturnValue = false;
//        //				return functionReturnValue;
//        //			}

//        //			//UPGRADE_WARNING: oMat03.Columns(FailCode).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if (Conversion.Val(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) == 0 & !string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value)) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("불량코드를 확인하세요.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //				functionReturnValue = false;
//        //				return functionReturnValue;
//        //			}

//        //			//UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: oMat01.Columns(LineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: oMat03.Columns(OrdMgNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value) & (oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.Value)) {
//        //				//UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				FailQty = FailQty + Conversion.Val(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value);
//        //			}
//        //		}
//        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (Strings.Trim(oForm.Items.Item("OrdType").Specific.Value) != "50" & Strings.Trim(oForm.Items.Item("OrdType").Specific.Value) != "60") {
//        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if (Conversion.Val(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value) != FailQty) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("공정리스트의 불량수량과 불량정보의 불량수량이 일치하지 않습니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				functionReturnValue = false;
//        //				return functionReturnValue;
//        //			}
//        //		}

//        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value) == "601" | Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value) == "111") {
//        //			//If oMat01.Columns("CpCode").Cells(i).Specific.Value = "CP80101" And Trim(oMat01.Columns("CItemCod").Cells(i).Specific.Value) = "" Then
//        //			//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: oMat01.Columns(Sequence).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if (oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value == 1 & string.IsNullOrEmpty(Strings.Trim(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value))) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("공정 사용 원재료코드가 없습니다. 사용 원재료를 선택해 주세요", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				functionReturnValue = false;
//        //				return functionReturnValue;
//        //			}
//        //		}
//        //	}

//        //	//비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_S
//        //	for (i = 1; i <= oMat02.VisualRowCount - 1; i++) {

//        //		//UPGRADE_WARNING: oMat02.Columns(NCode).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if ((!string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))) {

//        //			//UPGRADE_WARNING: oMat02.Columns(NTime).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if (string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value)) {

//        //				SubMain.Sbo_Application.SetStatusBarMessage("비가동코드가 입력되었을 때는 비가동시간은 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				oMat02.Columns.Item("NTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //				functionReturnValue = false;
//        //				return functionReturnValue;

//        //			}

//        //		}

//        //		//UPGRADE_WARNING: oMat02.Columns(NTime).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if ((!string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value) & oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value != "0")) {

//        //			//UPGRADE_WARNING: oMat02.Columns(NCode).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			if (string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value)) {

//        //				SubMain.Sbo_Application.SetStatusBarMessage("비가동시간이 입력되었을 때는 비가동코드는 필수입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				oMat02.Columns.Item("NCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //				functionReturnValue = false;
//        //				return functionReturnValue;

//        //			}

//        //		}
//        //		//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (Conversion.Val(oMat02.Columns.Item("TTime").Cells.Item(i).Specific.Value) == 0) {
//        //			//UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			oDS_PS_PP044M.SetValue("U_TTime", i - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item("YTime").Cells.Item(i).Specific.Value) + Conversion.Val(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value)));
//        //		}


//        //	}
//        //	//비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_E

//        //	if ((PS_PP044_Validate("검사01") == false)) {
//        //		functionReturnValue = false;
//        //		return functionReturnValue;
//        //	}

//        //	oDS_PS_PP044L.RemoveRecord(oDS_PS_PP044L.Size - 1);
//        //	oMat01.LoadFromDataSource();
//        //	oDS_PS_PP044M.RemoveRecord(oDS_PS_PP044M.Size - 1);
//        //	oMat02.LoadFromDataSource();

//        //	if ((oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//        //		PS_PP044_FormClear();
//        //	}
//        //	functionReturnValue = true;
//        //	return functionReturnValue;
//        //	PS_PP044_DataValidCheck_Error:
//        //	functionReturnValue = false;
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_DataValidCheck_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //	return functionReturnValue;
//        //}
//        #endregion

//        #region PS_PP044_MTX01
//        //private void PS_PP044_MTX01()
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	string OrdGbn = null;
//        //	string BPLID = null;
//        //	string DocDate = null;
//        //	string Gubun = null;

//        //	int i = 0;
//        //	string sQry = null;

//        //	int sCount = 0;
//        //	int sSeq = 0;
//        //	SAPbobsCOM.Recordset oRecordSet01 = null;

//        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//        //	////메트릭스에 데이터 로드
//        //	oForm.Freeze(true);

//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	BPLID = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	DocDate = Strings.Trim(oForm.Items.Item("DocDate").Specific.Value);
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	OrdGbn = Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value);
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	Gubun = Strings.Trim(oForm.Items.Item("Gubun").Specific.Value);

//        //	sQry = "EXEC PS_PP044_01 '" + BPLID + "','" + DocDate + "', '" + OrdGbn + "', '" + Gubun + "'";
//        //	oRecordSet01.DoQuery(sQry);

//        //	oMat04.Clear();
//        //	oDS_PS_PP044T.Clear();
//        //	oMat04.FlushToDataSource();
//        //	oMat04.LoadFromDataSource();



//        //	for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
//        //		if (i + 1 > oDS_PS_PP044T.Size) {
//        //			oDS_PS_PP044T.InsertRecord((i));
//        //		}

//        //		oMat04.AddRow();
//        //		oDS_PS_PP044T.Offset = i;

//        //		oDS_PS_PP044T.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//        //		oDS_PS_PP044T.SetValue("U_ColReg01", i, "N");
//        //		oDS_PS_PP044T.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("CntcCode").Value));
//        //		oDS_PS_PP044T.SetValue("U_ColReg03", i, Strings.Trim(oRecordSet01.Fields.Item("FullName").Value));
//        //		oDS_PS_PP044T.SetValue("U_ColQty01", i, Strings.Trim(oRecordSet01.Fields.Item("Base").Value));
//        //		oDS_PS_PP044T.SetValue("U_ColQty02", i, Strings.Trim(oRecordSet01.Fields.Item("Extend").Value));
//        //		oDS_PS_PP044T.SetValue("U_ColQty03", i, Strings.Trim(oRecordSet01.Fields.Item("YTime").Value));
//        //		oDS_PS_PP044T.SetValue("U_ColQty04", i, Strings.Trim(oRecordSet01.Fields.Item("NTime").Value));

//        //		oRecordSet01.MoveNext();

//        //	}

//        //	oMat04.LoadFromDataSource();
//        //	oMat04.AutoResizeColumns();

//        //	oForm.Update();


//        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	oRecordSet01 = null;
//        //	oForm.Freeze(false);
//        //	return;
//        //	PS_PP044_MTX01_Exit:
//        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	oRecordSet01 = null;
//        //	oForm.Freeze(false);

//        //	return;
//        //	PS_PP044_MTX01_Error:

//        //	oForm.Freeze(false);
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region PS_PP044_MTX02
//        //private void PS_PP044_MTX02()
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	string OrdGbn = null;
//        //	string BPLID = null;
//        //	string DocDate = null;
//        //	string Gubun = null;

//        //	int i = 0;
//        //	string sQry = null;

//        //	int sCount = 0;
//        //	int sSeq = 0;
//        //	SAPbobsCOM.Recordset oRecordSet01 = null;

//        //	oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


//        //	////메트릭스에 데이터 로드
//        //	oForm.Freeze(true);

//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	BPLID = Strings.Trim(oForm.Items.Item("BPLId").Specific.Value);
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	DocDate = Strings.Trim(oForm.Items.Item("DocDate").Specific.Value);
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	OrdGbn = Strings.Trim(oForm.Items.Item("OrdGbn").Specific.Value);
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	Gubun = Strings.Trim(oForm.Items.Item("Gubun").Specific.Value);

//        //	sQry = "EXEC PS_PP044_02 '" + BPLID + "','" + DocDate + "', '" + OrdGbn + "', '" + Gubun + "'";
//        //	oRecordSet01.DoQuery(sQry);

//        //	oMat05.Clear();
//        //	oDS_PS_PP044U.Clear();
//        //	oMat05.FlushToDataSource();
//        //	oMat05.LoadFromDataSource();



//        //	for (i = 0; i <= oRecordSet01.RecordCount - 1; i++) {
//        //		if (i + 1 > oDS_PS_PP044U.Size) {
//        //			oDS_PS_PP044U.InsertRecord((i));
//        //		}

//        //		oMat05.AddRow();
//        //		oDS_PS_PP044U.Offset = i;

//        //		oDS_PS_PP044U.SetValue("U_LineNum", i, Convert.ToString(i + 1));
//        //		oDS_PS_PP044U.SetValue("U_ColReg01", i, Strings.Trim(oRecordSet01.Fields.Item("CntcCode").Value));
//        //		oDS_PS_PP044U.SetValue("U_ColReg02", i, Strings.Trim(oRecordSet01.Fields.Item("FullName").Value));
//        //		oDS_PS_PP044U.SetValue("U_ColQty01", i, Strings.Trim(oRecordSet01.Fields.Item("Base").Value));
//        //		oDS_PS_PP044U.SetValue("U_ColQty02", i, Strings.Trim(oRecordSet01.Fields.Item("Extend").Value));
//        //		oDS_PS_PP044U.SetValue("U_ColQty03", i, Strings.Trim(oRecordSet01.Fields.Item("YTime").Value));
//        //		oDS_PS_PP044U.SetValue("U_ColQty04", i, Strings.Trim(oRecordSet01.Fields.Item("NTime").Value));

//        //		oRecordSet01.MoveNext();

//        //	}

//        //	oMat05.LoadFromDataSource();
//        //	oMat05.AutoResizeColumns();

//        //	oForm.Update();


//        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	oRecordSet01 = null;
//        //	oForm.Freeze(false);
//        //	return;
//        //	PS_PP044_MTX02_Exit:
//        //	//UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	oRecordSet01 = null;
//        //	oForm.Freeze(false);

//        //	return;
//        //	PS_PP044_MTX02_Error:

//        //	oForm.Freeze(false);
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_MTX02_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region PS_PP044_SumWorkTime
//        //private void PS_PP044_SumWorkTime()
//        //{
//        //	//******************************************************************************
//        //	//Function ID    : PS_PP044_SumWorkTime()
//        //	//해 당 모 듈    : 생산관리
//        //	//기        능    : 근무시간의 총합을 구함
//        //	//인        수    : 없음
//        //	//반   환   값   : 없음
//        //	//특 이 사 항    : 없음
//        //	//******************************************************************************
//        //	 // ERROR: Not supported in C#: OnErrorStatement


//        //	short loopCount = 0;
//        //	double Total = 0;

//        //	for (loopCount = 0; loopCount <= oMat01.RowCount - 2; loopCount++) {
//        //		//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		Total = Total + Convert.ToDouble((string.IsNullOrEmpty(Strings.Trim(oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.Value)) ? 0 : Strings.Trim(oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.Value)));
//        //	}

//        //	//UPGRADE_WARNING: oForm.Items(Total).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	oForm.Items.Item("Total").Specific.Value = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(Total, "##0.#0");

//        //	return;
//        //	PS_PP044_SumWorkTime_Error:

//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_SumWorkTime_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}

//        //private void PS_PP044_FormResize()
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	oForm.Items.Item("Mat02").Top = (oForm.Height / 5) * 2;
//        //	//200 '185 '170
//        //	oForm.Items.Item("Mat02").Left = 7;
//        //	oForm.Items.Item("Mat02").Height = ((oForm.Height - 170) / 3 * 1) - 20;
//        //	oForm.Items.Item("Mat02").Width = oForm.Width / 2 - 14;

//        //	oForm.Items.Item("Mat03").Top = (oForm.Height / 5) * 2;
//        //	//200 '185 '170
//        //	oForm.Items.Item("Mat03").Left = oForm.Width / 2;
//        //	oForm.Items.Item("Mat03").Height = ((oForm.Height - 170) / 3 * 1) - 20;
//        //	oForm.Items.Item("Mat03").Width = oForm.Width / 2 - 14;

//        //	oForm.Items.Item("Mat04").Top = 10;
//        //	oForm.Items.Item("Mat04").Left = oForm.Width / 2;
//        //	oForm.Items.Item("Mat04").Height = (oForm.Height / 4) - 10;
//        //	//((oForm.Height - 170) / 4 * 1) - 20
//        //	oForm.Items.Item("Mat04").Width = oForm.Width / 3 - 25;
//        //	//14
//        //	oMat04.AutoResizeColumns();

//        //	oForm.Items.Item("Mat05").Top = oForm.Items.Item("Mat04").Top + oForm.Items.Item("Mat04").Height;
//        //	oForm.Items.Item("Mat05").Left = oForm.Width / 2;
//        //	oForm.Items.Item("Mat05").Height = oForm.Items.Item("Mat04").Height / 2;
//        //	//((oForm.Height - 170) / 4 * 1) - 20
//        //	oForm.Items.Item("Mat05").Width = oForm.Width / 3 - 25;
//        //	//14
//        //	oMat04.AutoResizeColumns();

//        //	oForm.Items.Item("Mat01").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 20;
//        //	oForm.Items.Item("Mat01").Left = 7;
//        //	oForm.Items.Item("Mat01").Height = oForm.Height / 4;
//        //	//((oForm.Height - 170) / 3 * 2) - 120 '80
//        //	oForm.Items.Item("Mat01").Width = oForm.Width - 21;

//        //	oForm.Items.Item("Opt01").Top = oForm.Items.Item("Mat02").Top - 20;
//        //	oForm.Items.Item("Opt02").Top = oForm.Items.Item("Mat03").Top - 20;
//        //	oForm.Items.Item("EmpChk").Top = oForm.Items.Item("Mat02").Top - 20;
//        //	oForm.Items.Item("Button03").Top = oForm.Items.Item("Mat02").Top - 20;

//        //	oForm.Items.Item("Opt01").Left = 10;
//        //	oForm.Items.Item("Opt02").Left = oForm.Width / 2;
//        //	oForm.Items.Item("Opt03").Left = 10;
//        //	oForm.Items.Item("Opt03").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 5;
//        //	//20

//        //	oForm.Items.Item("Button03").Left = oForm.Items.Item("Mat02").Width - oForm.Items.Item("Button03").Width;

//        //	return;
//        //	PS_PP044_FormResize_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_FormResize_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region PS_PP044_Validate
//        //public bool PS_PP044_Validate(string ValidateType)
//        //{
//        //	bool functionReturnValue = false;
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	functionReturnValue = true;
//        //	object i = null;
//        //	int j = 0;
//        //	string Query01 = null;
//        //	SAPbobsCOM.Recordset RecordSet01 = null;
//        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        //	int PrevDBCpQty = 0;
//        //	int PrevMATRIXCpQty = 0;
//        //	int CurrentDBCpQty = 0;
//        //	int CurrentMATRIXCpQty = 0;
//        //	int NextDBCpQty = 0;
//        //	int NextMATRIXCpQty = 0;
//        //	string PrevCpInfo = null;
//        //	string CurrentCpInfo = null;
//        //	string NextCpInfo = null;

//        //	string OrdMgNum = null;
//        //	bool Exist = false;

//        //	if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_PP040H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y") {
//        //			MDC_Com.MDC_GF_Message(ref "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할수 없습니다.", ref "W");
//        //			functionReturnValue = false;
//        //			goto PS_PP044_Validate_Exit;
//        //		}
//        //	}

//        //	//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	////작업타입이 일반,조정인경우
//        //	if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm.Items.Item("OrdType").Specific.Selected.Value == "60") {
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	////작업타입이 PSMT지원인경우
//        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	////작업타입이 외주인경우
//        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {
//        //		MDC_Com.MDC_GF_Message(ref "해당작업타입은 변경이 불가능합니다.", ref "W");
//        //		functionReturnValue = false;
//        //		goto PS_PP044_Validate_Exit;
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	////작업타입이 실적인경우
//        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {
//        //		MDC_Com.MDC_GF_Message(ref "해당작업타입은 변경이 불가능합니다.", ref "W");
//        //		functionReturnValue = false;
//        //		goto PS_PP044_Validate_Exit;
//        //	}

//        //	string QueryString = null;
//        //	if (ValidateType == "검사01") {
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 일반인경우
//        //		if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") {
//        //			////입력된 행에 대해
//        //			for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
//        //				//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [PS_PP030H] PS_PP030H LEFT JOIN [PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = ' & oMat01.Columns(OrdMgNum).Cells(i).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1) <= 0) {
//        //					MDC_Com.MDC_GF_Message(ref "작업지시문서가 존재하지 않습니다.", ref "W");
//        //					functionReturnValue = false;
//        //					goto PS_PP044_Validate_Exit;
//        //				}
//        //			}

//        //			if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//        //				////삭제된 행에 대한처리
//        //				Query01 = "SELECT ";
//        //				Query01 = Query01 + " PS_PP044H.DocEntry,";
//        //				Query01 = Query01 + " PS_PP044L.LineId,";
//        //				Query01 = Query01 + " CONVERT(NVARCHAR,PS_PP044H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP044L.LineId) AS DocInfo,";
//        //				Query01 = Query01 + " PS_PP044L.U_OrdGbn AS OrdGbn,";
//        //				Query01 = Query01 + " PS_PP044L.U_PP030HNo AS PP030HNo,";
//        //				Query01 = Query01 + " PS_PP044L.U_PP030MNo AS PP030MNo,";
//        //				Query01 = Query01 + " PS_PP044L.U_OrdMgNum AS OrdMgNum ";
//        //				Query01 = Query01 + " FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry ";
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				Query01 = Query01 + " WHERE PS_PP044L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
//        //				RecordSet01.DoQuery(Query01);
//        //				for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
//        //					Exist = false;
//        //					////기존에 있는 행에대한처리
//        //					for (j = 1; j <= oMat01.VisualRowCount - 1; j++) {
//        //						//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						////새로추가된 행인경우, 검사할필요없다
//        //						if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value))) {
//        //						} else {
//        //							////라인번호가 같고, 문서번호가 같으면 존재하는행
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							if (Conversion.Val(RecordSet01.Fields.Item(0).Value) == Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) & Conversion.Val(RecordSet01.Fields.Item(1).Value) == Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value)) {
//        //								Exist = true;
//        //							}
//        //						}
//        //					}
//        //					////삭제된 행중 수량관계를 알아본다.
//        //					if (Exist == false) {
//        //						////휘팅이면서
//        //						if (RecordSet01.Fields.Item("OrdGbn").Value == "101") {
//        //							////현재 공정이 실적공정이면..
//        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP040_05 ' & RecordSet01.Fields(OrdMgNum).Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							if (MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y") {
//        //								////휘팅벌크포장
//        //								//                            PP040_CurrentPQty = 0
//        //								//                            PP040_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044H.Canceled = 'N' AND PS_PP044L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP044L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)
//        //								//                            PP070_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP070L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)
//        //								//                            PP080_DBPQty = MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" & RecordSet01.Fields("PP030HNo").Value & "' AND PS_PP070L.U_PP030MNo = '" & RecordSet01.Fields("PP030MNo").Value & "'", 0, 1)

//        //								if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP070L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) {
//        //									MDC_Com.MDC_GF_Message(ref "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
//        //									functionReturnValue = false;
//        //									goto PS_PP044_Validate_Exit;
//        //								}
//        //								////휘팅실적
//        //								if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) {
//        //									MDC_Com.MDC_GF_Message(ref "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
//        //									functionReturnValue = false;
//        //									goto PS_PP044_Validate_Exit;
//        //								}
//        //							}
//        //						}

//        //						////기계공구,몰드
//        //						if (RecordSet01.Fields.Item("OrdGbn").Value == "105" | RecordSet01.Fields.Item("OrdGbn").Value == "106") {
//        //							////그냥 입력가능
//        //						////휘팅,부품
//        //						} else if (RecordSet01.Fields.Item("OrdGbn").Value == "101" | RecordSet01.Fields.Item("OrdGbn").Value == "102") {
//        //							////삭제된 행에 대한 검사..
//        //							//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							OrdMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
//        //							//// DocEntry + '-' + LineId
//        //							CurrentCpInfo = OrdMgNum;

//        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" + OrdMgNum + "'");
//        //							if (string.IsNullOrEmpty(PrevCpInfo)) {
//        //								////해당공정이 첫공정이면 입력되어도 상관없다.
//        //							} else {
//        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP044H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP044H.Canceled = 'N'");
//        //								////재공이동 수량
//        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");
//        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");

//        //								PrevMATRIXCpQty = 0;
//        //								for (j = 1; j <= oMat01.VisualRowCount - 1; j++) {
//        //									//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == PrevCpInfo)) {
//        //										//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //										PrevMATRIXCpQty = PrevMATRIXCpQty + Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
//        //									}
//        //								}
//        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP044L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP044H.Canceled = 'N'");
//        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");
//        //								//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");

//        //								CurrentMATRIXCpQty = 0;
//        //								for (j = 1; j <= oMat01.VisualRowCount - 1; j++) {
//        //									//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == CurrentCpInfo)) {
//        //										//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //										CurrentMATRIXCpQty = CurrentMATRIXCpQty + Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
//        //									}
//        //								}
//        //								if (((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))) {
//        //									SubMain.Sbo_Application.SetStatusBarMessage("삭제된 공정의 선행공정의 생산수량이 삭제된 공정의 생산수량을 미달합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //									functionReturnValue = false;
//        //									goto PS_PP044_Validate_Exit;
//        //								}
//        //							}
//        //							//                        If oForm.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
//        //							//                            NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & OrdMgNum & "'")
//        //							//                            If NextCpInfo = "" Then
//        //							//                                '//해당공정이 마지막공정이면 삭제되어도 상관없다.
//        //							//                            Else
//        //							//                                NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP044H.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP044H.Canceled = 'N'")
//        //							//                                NextMATRIXCpQty = 0
//        //							//                                For j = 1 To oMat01.VisualRowCount - 1
//        //							//                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = NextCpInfo) Then
//        //							//                                        NextMATRIXCpQty = NextMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
//        //							//                                    End If
//        //							//                                Next
//        //							//                                CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP044L.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP044H.Canceled = 'N'")
//        //							//                                CurrentMATRIXCpQty = 0
//        //							//                                For j = 1 To oMat01.VisualRowCount - 1 '//현재공정은 삭제되었으므로.. 매트릭스에 존재하지 않는다.
//        //							//                                    If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
//        //							//                                        CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
//        //							//                                    End If
//        //							//                                Next
//        //							//                                If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
//        //							//                                    Sbo_Application.SetStatusBarMessage "삭제된 공정의 후행공정의 생산수량이 삭제된 공정의 생산수량을 초과합니다.", bmt_Short, True
//        //							//                                    PS_PP044_Validate = False
//        //							//                                    GoTo PS_PP044_Validate_Exit
//        //							//                                End If
//        //							//                            End If
//        //							//                        End If
//        //						}
//        //					}
//        //					RecordSet01.MoveNext();
//        //				}
//        //			}

//        //			if ((oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)) {
//        //				for (i = 1; i <= oMat01.VisualRowCount - 1; i++) {
//        //					//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					////새로추가된 행인경우, 검사할필요없다
//        //					if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value))) {
//        //					} else {
//        //						//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						////휘팅이면서
//        //						if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "101") {
//        //							////현재공정이 실적공정이면
//        //							//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP040_05 ' & Val(oForm.Items(DocEntry).Specific.Value) & - & Val(oMat01.Columns(LineId).Cells(i).Specific.Value) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //							////현재 공정이 바렐 앞공정이면..
//        //							if (MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "-" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value) + "'", 0, 1) == "Y") {
//        //								//                            '//휘팅벌크포장,휘팅실적
//        //								//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //								if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "' AND PS_PP070L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value) + "'", 0, 1)) > 0 | (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "' AND PS_PP080L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value) + "'", 0, 1)) > 0) {
//        //									////작업일보등록된문서중에 수정이 된문서를 구함
//        //									Query01 = "SELECT ";
//        //									Query01 = Query01 + " PS_PP044L.U_OrdMgNum,";
//        //									Query01 = Query01 + " PS_PP044L.U_Sequence,";
//        //									Query01 = Query01 + " PS_PP044L.U_CpCode,";
//        //									Query01 = Query01 + " PS_PP044L.U_ItemCode,";
//        //									Query01 = Query01 + " PS_PP044L.U_PP030HNo,";
//        //									Query01 = Query01 + " PS_PP044L.U_PP030MNo,";
//        //									Query01 = Query01 + " PS_PP044L.U_PQty,";
//        //									Query01 = Query01 + " PS_PP044L.U_NQty,";
//        //									Query01 = Query01 + " PS_PP044L.U_ScrapWt,";
//        //									Query01 = Query01 + " PS_PP044L.U_WorkTime";
//        //									Query01 = Query01 + " FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry";
//        //									//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									Query01 = Query01 + " WHERE PS_PP044H.DocEntry = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "'";
//        //									//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									Query01 = Query01 + " AND PS_PP044L.LineId = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value) + "'";
//        //									Query01 = Query01 + " AND PS_PP044H.Canceled = 'N'";
//        //									RecordSet01.DoQuery(Query01);
//        //									//UPGRADE_WARNING: oMat01.Columns(WorkTime).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(ScrapWt).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(NQty).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(PQty).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(PP030MNo).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(PP030HNo).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(ItemCode).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(Sequence).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									//UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //									if ((RecordSet01.Fields.Item(0).Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(1).Value == oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(2).Value == oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(3).Value == oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(4).Value == oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(5).Value == oMat01.Columns.Item("PP030MNo").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(6).Value == oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(7).Value == oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(8).Value == oMat01.Columns.Item("ScrapWt").Cells.Item(i).Specific.Value & RecordSet01.Fields.Item(9).Value == oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value)) {
//        //									////값이 변경된 행의경우
//        //									} else {
//        //										MDC_Com.MDC_GF_Message(ref "생산실적이 등록된 행은 수정할수 없습니다.", ref "W");
//        //										functionReturnValue = false;
//        //										goto PS_PP044_Validate_Exit;
//        //									}
//        //								}
//        //							}
//        //						}
//        //					}
//        //				}
//        //			}

//        //			////저장 속도가 너무 느려 임시로 막음

//        //			//            For i = 1 To oMat01.VisualRowCount - 1 '//입력된 모든행에 대해 입력가능성 검사
//        //			//                If oMat01.Columns("OrdGbn").Cells(i).Specific.Value = "105" Or oMat01.Columns("OrdGbn").Cells(i).Specific.Value = "106" Then '//기계공구,몰드
//        //			//                    '//그냥 입력가능
//        //			//                ElseIf oMat01.Columns("OrdGbn").Cells(i).Specific.Value = "101" Or oMat01.Columns("OrdGbn").Cells(i).Specific.Value = "102" Then '//휘팅,부품
//        //			//                    OrdMgNum = oMat01.Columns("OrdMgNum").Cells(i).Specific.Value
//        //			//                    CurrentCpInfo = OrdMgNum
//        //			//
//        //			//                    PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" & OrdMgNum & "'")
//        //			//                    If PrevCpInfo = "" Then
//        //			//                        '//해당공정이 첫공정이면 입력되어도 상관없다.
//        //			//                    Else
//        //			//
//        //			//                        PrevDBCpQty = MDC_PS_Common.GetValue("EXEC PS_PP040_07 '" & PrevCpInfo & "', '" & oForm.Items("DocEntry").Specific.Value & "'")
//        //			//                        '//재공 이동수량 반영
//        //			//                        PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" & PrevCpInfo & "' AND a.Canceled = 'N'")
//        //			//                        PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" & PrevCpInfo & "' AND a.Canceled = 'N'")
//        //			//
//        //			//                        PrevMATRIXCpQty = 0
//        //			//                        For j = 1 To oMat01.VisualRowCount - 1
//        //			//                            If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = PrevCpInfo) Then
//        //			//                                PrevMATRIXCpQty = PrevMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
//        //			//                            End If
//        //			//                        Next
//        //			//
//        //			//                        CurrentDBCpQty = MDC_PS_Common.GetValue("EXEC PS_PP040_07 '" & CurrentCpInfo & "', '" & oForm.Items("DocEntry").Specific.Value & "'")
//        //			//                        CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" & CurrentCpInfo & "' AND a.Canceled = 'N'")
//        //			//                        CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" & CurrentCpInfo & "' AND a.Canceled = 'N'")
//        //			//
//        //			//                        CurrentMATRIXCpQty = 0
//        //			//                        For j = 1 To oMat01.VisualRowCount - 1
//        //			//                            If (oMat01.Columns("OrdMgNum").Cells(j).Specific.Value = CurrentCpInfo) Then
//        //			//                                CurrentMATRIXCpQty = CurrentMATRIXCpQty + Val(oMat01.Columns("PQty").Cells(j).Specific.Value)
//        //			//                            End If
//        //			//                        Next
//        //			//                        '// 노대리님 요청 주석
//        //			//                        If ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
//        //			//                            Sbo_Application.SetStatusBarMessage "선행공정의 생산수량이 현공정의 생산수량에 미달 합니다.", bmt_Short, True
//        //			//                            Call oMat01.SelectRow(i, True, False)
//        //			//                            PS_PP044_Validate = False
//        //			//                            GoTo PS_PP044_Validate_Exit
//        //			//                        End If
//        //			//
//        //			//                    End If
//        //			//                End If
//        //			//            Next

//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 PSMT지원인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {
//        //			////현재는 특별한 조건이 필요치 않음
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 외주인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {
//        //			////현재는 특별한 조건이 필요치 않음
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 실적인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {
//        //			////현재는 특별한 조건이 필요치 않음
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 조정인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") {
//        //			////현재는 특별한 조건이 필요치 않음
//        //		}
//        //	} else if (ValidateType == "행삭제01") {
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 일반인경우
//        //		if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") {
//        //			////행삭제전 행삭제가능여부검사
//        //			//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(oMat01Row01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			////새로추가된 행인경우, 삭제하여도 무방하다
//        //			if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value))) {
//        //			} else {
//        //				//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(oMat01Row01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////휘팅이면서
//        //				if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "101") {
//        //					////현재공정이 실적공정이면
//        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP040_05 ' & Val(oMat01.Columns(PP030HNo).Cells(oMat01Row01).Specific.Value) & - & Val(oMat01.Columns(PP030MNo).Cells(oMat01Row01).Specific.Value) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					////현재 공정이 바렐 앞공정이면..
//        //					if (MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" + Conversion.Val(oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value) + "-" + Conversion.Val(oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value) + "'", 0, 1) == "Y") {
//        //						////휘팅벌크포장
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "' AND PS_PP070L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value) + "'", 0, 1)) > 0) {
//        //							MDC_Com.MDC_GF_Message(ref "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
//        //							functionReturnValue = false;
//        //							goto PS_PP044_Validate_Exit;
//        //						}
//        //						////휘팅실적
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "' AND PS_PP080L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value) + "'", 0, 1)) > 0) {
//        //							MDC_Com.MDC_GF_Message(ref "삭제된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
//        //							functionReturnValue = false;
//        //							goto PS_PP044_Validate_Exit;
//        //						}
//        //					}

//        //					//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(oMat01Row01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////기계공구,몰드
//        //				} else if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "105" | oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "106") {

//        //					//재고가 존재하면 행삭제 불가 기능 추가(2011.12.15 송명규 추가)

//        //					QueryString = "                     SELECT      SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]";
//        //					QueryString = QueryString + "  FROM       OINM AS A";
//        //					QueryString = QueryString + "                 INNER JOIN";
//        //					QueryString = QueryString + "                 OITM As B";
//        //					QueryString = QueryString + "                     ON A.ItemCode = B.ItemCode";
//        //					QueryString = QueryString + "  WHERE      B.U_ItmBsort IN ('105','106')";
//        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					QueryString = QueryString + "                 AND A.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value + "'";
//        //					QueryString = QueryString + "  GROUP BY  A.ItemCode";

//        //					if ((string.IsNullOrEmpty((MDC_PS_Common.GetValue(QueryString, 0, 1))) ? 0 : (MDC_PS_Common.GetValue(QueryString, 0, 1))) > 0) {

//        //						MDC_Com.MDC_GF_Message(ref "재고가 존재하는 작번입니다. 삭제할 수 없습니다.", ref "W");
//        //						functionReturnValue = false;
//        //						goto PS_PP044_Validate_Exit;

//        //					}

//        //				}

//        //			}
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 PSMT인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 외주인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 실적인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 조정인경
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") {

//        //		}
//        //	} else if (ValidateType == "수정01") {
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 일반인경우
//        //		if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") {
//        //			////수정전 수정가능여부검사
//        //			//UPGRADE_WARNING: oMat01.Columns(LineId).Cells(oMat01Row01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			////새로추가된 행인경우, 수정하여도 무방하다
//        //			if ((string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value))) {
//        //			} else {
//        //				//UPGRADE_WARNING: oMat01.Columns(OrdGbn).Cells(oMat01Row01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				////휘팅이면서
//        //				if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "101") {
//        //					////현재공정이 실적공정이면
//        //					//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP040_05 ' & Val(oMat01.Columns(PP030HNo).Cells(oMat01Row01).Specific.Value) & - & Val(oMat01.Columns(PP030MNo).Cells(oMat01Row01).Specific.Value) & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					////현재 공정이 바렐 앞공정이면..
//        //					if (MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" + Conversion.Val(oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value) + "-" + Conversion.Val(oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value) + "'", 0, 1) == "Y") {
//        //						////휘팅벌크포장
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "' AND PS_PP070L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value) + "'", 0, 1)) > 0) {
//        //							MDC_Com.MDC_GF_Message(ref "수정된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
//        //							functionReturnValue = false;
//        //							goto PS_PP044_Validate_Exit;
//        //						}
//        //						////휘팅실적
//        //						//UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + Conversion.Val(oForm.Items.Item("DocEntry").Specific.Value) + "' AND PS_PP080L.U_PP030MNo = '" + Conversion.Val(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value) + "'", 0, 1)) > 0) {
//        //							MDC_Com.MDC_GF_Message(ref "수정된행이 생산실적 등록된 행입니다. 적용할수 없습니다.", ref "W");
//        //							functionReturnValue = false;
//        //							goto PS_PP044_Validate_Exit;
//        //						}
//        //					}
//        //				}
//        //			}
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 PSMT인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 외주인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 실적인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 조정인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") {
//        //		}
//        //	} else if (ValidateType == "취소") {
//        //		//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		//UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT Canceled FROM [PS_PP040H] WHERE DocEntry = ' & oForm.Items(DocEntry).Specific.Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (MDC_PS_Common.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y") {
//        //			MDC_Com.MDC_GF_Message(ref "이미취소된 문서 입니다. 취소할수 없습니다.", ref "W");
//        //			functionReturnValue = false;
//        //			goto PS_PP044_Validate_Exit;
//        //		}
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 일반인경우
//        //		if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10") {
//        //			////삭제된 행에 대한처리
//        //			Query01 = "SELECT ";
//        //			Query01 = Query01 + " PS_PP044H.DocEntry,";
//        //			Query01 = Query01 + " PS_PP044L.LineId,";
//        //			Query01 = Query01 + " CONVERT(NVARCHAR,PS_PP044H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP044L.LineId) AS DocInfo,";
//        //			Query01 = Query01 + " PS_PP044L.U_OrdGbn AS OrdGbn,";
//        //			Query01 = Query01 + " PS_PP044L.U_PP030HNo AS PP030HNo,";
//        //			Query01 = Query01 + " PS_PP044L.U_PP030MNo AS PP030MNo,";
//        //			Query01 = Query01 + " PS_PP044L.U_OrdMgNum AS OrdMgNum ";
//        //			Query01 = Query01 + " FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry ";
//        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			Query01 = Query01 + " WHERE PS_PP044L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
//        //			RecordSet01.DoQuery(Query01);
//        //			for (i = 0; i <= RecordSet01.RecordCount - 1; i++) {
//        //				////휘팅이면서
//        //				if (RecordSet01.Fields.Item("OrdGbn").Value == "101") {
//        //					////현재공정이 실적포인트이면
//        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue(EXEC PS_PP040_05 ' & RecordSet01.Fields(OrdMgNum).Value & ', 0, 1) 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					if (MDC_PS_Common.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y") {
//        //						////휘팅벌크포장
//        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP070L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) {
//        //							MDC_Com.MDC_GF_Message(ref "생산실적 등록된 문서입니다. 적용할수 없습니다.", ref "W");
//        //							functionReturnValue = false;
//        //							goto PS_PP044_Validate_Exit;
//        //						}
//        //						////휘팅실적
//        //						if ((MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0) {
//        //							MDC_Com.MDC_GF_Message(ref "생산실적 등록된 문서입니다. 적용할수 없습니다.", ref "W");
//        //							functionReturnValue = false;
//        //							goto PS_PP044_Validate_Exit;
//        //						}
//        //					}
//        //				}

//        //				////기계공구,몰드
//        //				if (RecordSet01.Fields.Item("OrdGbn").Value == "105" | RecordSet01.Fields.Item("OrdGbn").Value == "106") {
//        //					////그냥 입력가능
//        //				////휘팅,부품
//        //				} else if (RecordSet01.Fields.Item("OrdGbn").Value == "101" | RecordSet01.Fields.Item("OrdGbn").Value == "102") {
//        //					////삭제된 행에 대한 검사..
//        //					//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					OrdMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
//        //					//// DocEntry + '-' + LineId
//        //					CurrentCpInfo = OrdMgNum;

//        //					//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //					PrevCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_02 '" + OrdMgNum + "'");
//        //					if (string.IsNullOrEmpty(PrevCpInfo)) {
//        //						////해당공정이 첫공정이면 입력되어도 상관없다.
//        //					} else {
//        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						PrevDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP044H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP044H.Canceled = 'N'");
//        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");
//        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						PrevDBCpQty = PrevDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'");

//        //						PrevMATRIXCpQty = 0;
//        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP044L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP044H.Canceled = 'N'");
//        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");
//        //						//UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //						CurrentDBCpQty = CurrentDBCpQty + MDC_PS_Common.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'");
//        //						CurrentMATRIXCpQty = 0;
//        //						if (((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))) {
//        //							SubMain.Sbo_Application.SetStatusBarMessage("취소문서의 선행공정의 생산수량이 취소문서의 생산수량을 미달합니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //							functionReturnValue = false;
//        //							goto PS_PP044_Validate_Exit;
//        //						}
//        //					}

//        //					//                    If oForm.Mode = fm_UPDATE_MODE Then '//후행공정은 수정모드에서만 수정함
//        //					//                        NextCpInfo = MDC_PS_Common.GetValue("EXEC PS_PP040_03 '" & OrdMgNum & "'")
//        //					//                        If NextCpInfo = "" Then
//        //					//                            '//해당공정이 마지막공정이면 삭제되어도 상관없다.
//        //					//                        Else
//        //					//                            NextDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" & NextCpInfo & "' AND PS_PP044H.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP044H.Canceled = 'N'")
//        //					//                            NextMATRIXCpQty = 0
//        //					//                            CurrentDBCpQty = MDC_PS_Common.GetValue("SELECT SUM(PS_PP044L.U_PQty) FROM [@PS_PP040H] PS_PP044H LEFT JOIN [@PS_PP040L] PS_PP044L ON PS_PP044H.DocEntry = PS_PP044L.DocEntry WHERE PS_PP044L.U_OrdMgNum = '" & CurrentCpInfo & "' AND PS_PP044L.DocEntry <> '" & RecordSet01.Fields(0).Value & "' AND PS_PP044H.Canceled = 'N'")
//        //					//                            CurrentMATRIXCpQty = 0
//        //					//                            If ((NextDBCpQty + NextMATRIXCpQty) > (CurrentDBCpQty + CurrentMATRIXCpQty)) Then
//        //					//                                Sbo_Application.SetStatusBarMessage "취소문서의 후행공정의 생산수량이 취소문서의 생산수량을 초과합니다.", bmt_Short, True
//        //					//                                PS_PP044_Validate = False
//        //					//                                GoTo PS_PP044_Validate_Exit
//        //					//                            End If
//        //					//                        End If
//        //					//                    End If
//        //				}
//        //				RecordSet01.MoveNext();
//        //			}
//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 PSMT인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {

//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 외주인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {

//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 실적인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {

//        //			//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		////작업타입이 조정인경우
//        //		} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "50") {

//        //		}
//        //	}
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	return functionReturnValue;
//        //	PS_PP044_Validate_Exit:
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	return functionReturnValue;
//        //	PS_PP044_Validate_Error:
//        //	functionReturnValue = false;
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_Validate_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //	return functionReturnValue;
//        //}
//        #endregion

//        #region PS_PP044_OrderInfoLoad
//        //private void PS_PP044_OrderInfoLoad()
//        //{
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	string Query01 = null;
//        //	SAPbobsCOM.Recordset RecordSet01 = null;
//        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//        //	//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	////일반,조정, 설계
//        //	if (oForm.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm.Items.Item("OrdType").Specific.Selected.Value == "60" | oForm.Items.Item("OrdType").Specific.Selected.Value == "70") {
//        //		//UPGRADE_WARNING: oForm.Items(OrdMgNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.Value)) {
//        //			MDC_Com.MDC_GF_Message(ref "작업지시 관리번호를 입력하지 않습니다.", ref "W");
//        //			goto PS_PP044_OrderInfoLoad_Exit;
//        //		} else {
//        //			Query01 = "SELECT ";
//        //			Query01 = Query01 + "U_OrdGbn,";
//        //			Query01 = Query01 + "U_BPLId,";
//        //			Query01 = Query01 + "U_ItemCode,";
//        //			Query01 = Query01 + "U_ItemName,";
//        //			Query01 = Query01 + "U_OrdNum,";
//        //			Query01 = Query01 + "U_OrdSub1,";
//        //			Query01 = Query01 + "U_OrdSub2,";
//        //			Query01 = Query01 + "DocEntry";
//        //			Query01 = Query01 + " FROM [@PS_PP030H]";
//        //			Query01 = Query01 + " WHERE ";
//        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			Query01 = Query01 + " U_OrdNum + U_OrdSub1 + U_OrdSub2 = '" + oForm.Items.Item("OrdMgNum").Specific.Value + "'";
//        //			Query01 = Query01 + " AND U_OrdGbn NOT IN('104','107') ";
//        //			Query01 = Query01 + " AND Canceled = 'N'";
//        //			RecordSet01.DoQuery(Query01);
//        //			if (RecordSet01.RecordCount == 0) {
//        //				MDC_Com.MDC_GF_Message(ref "작업지시 정보가 존재하지 않습니다.", ref "W");
//        //				goto PS_PP044_OrderInfoLoad_Exit;
//        //			} else {
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("OrdGbn").Specific.Select(RecordSet01.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//        //				//UPGRADE_WARNING: oForm.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("BPLId").Specific.Select(RecordSet01.Fields.Item(1).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
//        //				//UPGRADE_WARNING: oForm.Items(ItemCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("ItemCode").Specific.Value = RecordSet01.Fields.Item(2).Value;
//        //				//UPGRADE_WARNING: oForm.Items(ItemName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("ItemName").Specific.Value = RecordSet01.Fields.Item(3).Value;
//        //				//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("OrdNum").Specific.Value = RecordSet01.Fields.Item(4).Value;
//        //				//UPGRADE_WARNING: oForm.Items(OrdSub1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("OrdSub1").Specific.Value = RecordSet01.Fields.Item(5).Value;
//        //				//UPGRADE_WARNING: oForm.Items(OrdSub2).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("OrdSub2").Specific.Value = RecordSet01.Fields.Item(6).Value;
//        //				//UPGRADE_WARNING: oForm.Items(PP030HNo).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				//UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("PP030HNo").Specific.Value = RecordSet01.Fields.Item(7).Value;
//        //				//                '//매트릭스삭제
//        //				//                oMat01.Clear
//        //				//                oMat01.FlushToDataSource
//        //				//                oMat01.LoadFromDataSource
//        //				//                Call PS_PP044_AddMatrixRow01(0, True)
//        //				//                oMat02.Clear
//        //				//                oMat02.FlushToDataSource
//        //				//                oMat02.LoadFromDataSource
//        //				//                Call PS_PP044_AddMatrixRow02(0, True)
//        //				//                oMat03.Clear
//        //				//                oMat03.FlushToDataSource
//        //				//                oMat03.LoadFromDataSource
//        //				oForm.Update();
//        //			}
//        //		}
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	////PSMT
//        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "20") {
//        //		//UPGRADE_WARNING: oForm.Items(OrdMgNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //		if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.Value)) {
//        //			MDC_Com.MDC_GF_Message(ref "작업지시 관리번호를 입력하지 않습니다.", ref "W");
//        //			goto PS_PP044_OrderInfoLoad_Exit;
//        //		} else {
//        //			//UPGRADE_WARNING: oForm.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			oForm.Items.Item("OrdNum").Specific.Value = oForm.Items.Item("OrdMgNum").Specific.Value;
//        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			oForm.Items.Item("OrdSub1").Specific.Value = "000";
//        //			//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //			oForm.Items.Item("OrdSub2").Specific.Value = "00";
//        //			////매트릭스삭제
//        //			oMat01.Clear();
//        //			oMat01.FlushToDataSource();
//        //			oMat01.LoadFromDataSource();
//        //			PS_PP044_AddMatrixRow01(0, ref true);
//        //			oMat02.Clear();
//        //			oMat02.FlushToDataSource();
//        //			oMat02.LoadFromDataSource();
//        //			PS_PP044_AddMatrixRow02(0, ref true);
//        //			oMat03.Clear();
//        //			oMat03.FlushToDataSource();
//        //			oMat03.LoadFromDataSource();
//        //			oForm.Update();
//        //		}
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "30") {
//        //		MDC_Com.MDC_GF_Message(ref "외주은 입력할수 없습니다.", ref "W");
//        //		goto PS_PP044_OrderInfoLoad_Exit;
//        //		//UPGRADE_WARNING: oForm.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	} else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40") {
//        //		MDC_Com.MDC_GF_Message(ref "실적은 입력할수 없습니다.", ref "W");
//        //		goto PS_PP044_OrderInfoLoad_Exit;
//        //	}
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	return;
//        //	PS_PP044_OrderInfoLoad_Exit:
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	return;
//        //	PS_PP044_OrderInfoLoad_Error:
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	SubMain.Sbo_Application.SetStatusBarMessage("PS_PP044_OrderInfoLoad_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //}
//        #endregion

//        #region PS_PP044_FindValidateDocument
//        //public bool PS_PP044_FindValidateDocument(string ObjectType)
//        //{
//        //	bool functionReturnValue = false;
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	functionReturnValue = true;
//        //	string Query01 = null;
//        //	SAPbobsCOM.Recordset RecordSet01 = null;
//        //	string Query02 = null;
//        //	SAPbobsCOM.Recordset RecordSet02 = null;

//        //	int i = 0;
//        //	string DocEntry = null;
//        //	//UPGRADE_WARNING: oForm.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //	DocEntry = Strings.Trim(oForm.Items.Item("DocEntry").Specific.Value);
//        //	////원본문서

//        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//        //	RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        //	Query01 = " SELECT DocEntry";
//        //	Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry = ";
//        //	Query01 = Query01 + DocEntry;
//        //	if ((oDocType01 == "작업일보등록(작지)")) {
//        //		Query01 = Query01 + " AND U_DocType = '10'";
//        //	} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //		Query01 = Query01 + " AND U_DocType = '20'";
//        //	}
//        //	RecordSet01.DoQuery(Query01);
//        //	if ((RecordSet01.RecordCount == 0)) {
//        //		if ((oDocType01 == "작업일보등록(작지)")) {
//        //			SubMain.Sbo_Application.SetStatusBarMessage("작업일보등록(공정)문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //			SubMain.Sbo_Application.SetStatusBarMessage("작업일보등록(작지)문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //		}
//        //		functionReturnValue = false;
//        //		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //		RecordSet01 = null;
//        //		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //		RecordSet02 = null;
//        //		return functionReturnValue;
//        //	}

//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet02 = null;
//        //	return functionReturnValue;
//        //	PS_PP044_FindValidateDocument_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet02 = null;
//        //	functionReturnValue = false;
//        //	return functionReturnValue;
//        //}
//        #endregion

//        #region PS_PP044_DirectionValidateDocument
//        //public bool PS_PP044_DirectionValidateDocument(string DocEntry, string DocEntryNext, string Direction, string ObjectType)
//        //{
//        //	bool functionReturnValue = false;
//        //	 // ERROR: Not supported in C#: OnErrorStatement

//        //	string Query01 = null;
//        //	SAPbobsCOM.Recordset RecordSet01 = null;
//        //	string Query02 = null;
//        //	SAPbobsCOM.Recordset RecordSet02 = null;

//        //	int i = 0;
//        //	string MaxDocEntry = null;
//        //	string MinDocEntry = null;
//        //	bool DoNext = false;
//        //	bool IsFirst = false;
//        //	////시작유무
//        //	DoNext = true;
//        //	IsFirst = true;

//        //	RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
//        //	RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

//        //	while ((DoNext == true)) {
//        //		if ((IsFirst != true)) {
//        //			////문서전체를 경유하고도 유효값을 찾지못했다면
//        //			if ((DocEntry == DocEntryNext)) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("유효한문서가 존재하지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				functionReturnValue = false;
//        //				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //				RecordSet01 = null;
//        //				//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //				RecordSet02 = null;
//        //				return functionReturnValue;
//        //			}
//        //		}
//        //		if ((Direction == "Next")) {
//        //			Query01 = " SELECT TOP 1 DocEntry";
//        //			Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry > ";
//        //			Query01 = Query01 + DocEntryNext;
//        //			if ((oDocType01 == "작업일보등록(작지)")) {
//        //				Query01 = Query01 + " AND U_DocType = '10'";
//        //			} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //				Query01 = Query01 + " AND U_DocType = '20'";
//        //			}
//        //			Query01 = Query01 + " ORDER BY DocEntry ASC";
//        //		} else if ((Direction == "Prev")) {
//        //			Query01 = " SELECT TOP 1 DocEntry";
//        //			Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
//        //			Query01 = Query01 + DocEntryNext;
//        //			if ((oDocType01 == "작업일보등록(작지)")) {
//        //				Query01 = Query01 + " AND U_DocType = '10'";
//        //			} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //				Query01 = Query01 + " AND U_DocType = '20'";
//        //			}
//        //			Query01 = Query01 + " ORDER BY DocEntry DESC";
//        //		}
//        //		RecordSet01.DoQuery(Query01);
//        //		////해당문서가 마지막문서라면
//        //		if ((RecordSet01.Fields.Item(0).Value == 0)) {
//        //			if ((Direction == "Next")) {
//        //				Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
//        //				if ((oDocType01 == "작업일보등록(작지)")) {
//        //					Query02 = Query02 + " WHERE U_DocType = '10'";
//        //				} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //					Query02 = Query02 + " WHERE U_DocType = '20'";
//        //				}
//        //				Query02 = Query02 + " ORDER BY DocEntry ASC";
//        //			} else if ((Direction == "Prev")) {
//        //				Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
//        //				if ((oDocType01 == "작업일보등록(작지)")) {
//        //					Query02 = Query02 + " WHERE U_DocType = '10'";
//        //				} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //					Query02 = Query02 + " WHERE U_DocType = '20'";
//        //				}
//        //				Query02 = Query02 + " ORDER BY DocEntry DESC";
//        //			}
//        //			RecordSet02.DoQuery(Query02);
//        //			////문서가 아예 존재하지 않는다면
//        //			if ((RecordSet02.RecordCount == 0)) {
//        //				SubMain.Sbo_Application.SetStatusBarMessage("유효한문서가 존재하지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //				//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //				RecordSet01 = null;
//        //				//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //				RecordSet02 = null;
//        //				functionReturnValue = false;
//        //				return functionReturnValue;
//        //			} else {
//        //				if ((Direction == "Next")) {
//        //					DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) - 1);
//        //					Query01 = " SELECT TOP 1 DocEntry";
//        //					Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry > ";
//        //					Query01 = Query01 + DocEntryNext;
//        //					if ((oDocType01 == "작업일보등록(작지)")) {
//        //						Query01 = Query01 + " AND U_DocType = '10'";
//        //					} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //						Query01 = Query01 + " AND U_DocType = '20'";
//        //					}
//        //					Query01 = Query01 + " ORDER BY DocEntry ASC";
//        //					RecordSet01.DoQuery(Query01);
//        //				} else if ((Direction == "Prev")) {
//        //					DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) + 1);
//        //					Query01 = " SELECT TOP 1 DocNum";
//        //					Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
//        //					Query01 = Query01 + DocEntryNext;
//        //					if ((oDocType01 == "작업일보등록(작지)")) {
//        //						Query01 = Query01 + " AND U_DocType = '10'";
//        //					} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //						Query01 = Query01 + " AND U_DocType = '20'";
//        //					}
//        //					Query01 = Query01 + " ORDER BY DocEntry DESC";
//        //					RecordSet01.DoQuery(Query01);
//        //				}
//        //			}
//        //		}
//        //		if ((oDocType01 == "작업일보등록(작지)")) {
//        //			DoNext = false;
//        //			if ((Direction == "Next")) {
//        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
//        //			} else if ((Direction == "Prev")) {
//        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
//        //			}
//        //		} else if ((oDocType01 == "작업일보등록(공정)")) {
//        //			DoNext = false;
//        //			if ((Direction == "Next")) {
//        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
//        //			} else if ((Direction == "Prev")) {
//        //				DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
//        //			}
//        //		}
//        //		IsFirst = false;
//        //	}
//        //	////다음문서가 유효하다면 그냥 넘어가고
//        //	if ((DocEntry == DocEntryNext)) {
//        //		PS_PP044_FormItemEnabled();
//        //		////UDO방식
//        //	////다음문서가 유효하지 않다면
//        //	} else {
//        //		oForm.Freeze(true);
//        //		oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
//        //		PS_PP044_FormItemEnabled();
//        //		////UDO방식
//        //		////문서번호 필드가 입력이 가능하다면
//        //		if (oForm.Items.Item("DocEntry").Enabled == true) {
//        //			if ((Direction == "Next")) {
//        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("DocEntry").Specific.Value = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) + 1));
//        //			} else if ((Direction == "Prev")) {
//        //				//UPGRADE_WARNING: oForm.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
//        //				oForm.Items.Item("DocEntry").Specific.Value = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) - 1));
//        //			}
//        //			oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
//        //		}
//        //		oForm.Freeze(false);
//        //		functionReturnValue = false;
//        //		//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //		RecordSet01 = null;
//        //		//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //		RecordSet02 = null;
//        //		return functionReturnValue;
//        //	}
//        //	functionReturnValue = true;
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet02 = null;
//        //	return functionReturnValue;
//        //	PS_PP044_DirectionValidateDocument_Error:
//        //	SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
//        //	functionReturnValue = false;
//        //	//UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet01 = null;
//        //	//UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
//        //	RecordSet02 = null;
//        //	return functionReturnValue;
//        //}
//        #endregion
//    }
//}
