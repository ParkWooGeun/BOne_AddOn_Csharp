using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;
using PSH_BOne_AddOn.Code;

namespace PSH_BOne_AddOn
{
    /// <summary>
    /// 작업일보등록(작지)
    /// </summary>
    internal class PS_PP040 : PSH_BaseClass
    {
        private string oFormUniqueID;
        private SAPbouiCOM.Matrix oMat01;
        private SAPbouiCOM.Matrix oMat02;
        private SAPbouiCOM.Matrix oMat03;
        private SAPbouiCOM.DBDataSource oDS_PS_PP040H; //등록헤더
        private SAPbouiCOM.DBDataSource oDS_PS_PP040L; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP040M; //등록라인
        private SAPbouiCOM.DBDataSource oDS_PS_PP040N; //등록라인
        private string oLastItemUID01; //클래스에서 선택한 마지막 아이템 Uid값
        private string oLastColUID01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Col의 Uid값
        private int oLastColRow01; //마지막아이템이 메트릭스일경우에 마지막 선택된 Row값
        private int oMat01Row01;
        private int oMat02Row02;
        private int oMat03Row03;
        private string oDocType01;
        private string oDocEntry01;
        private string oOrdGbn;
        private string oSequence;
        private string oDocdate;
        private SAPbouiCOM.BoFormMode oFormMode01;

        /// <summary>
		/// Form 호출
		/// </summary>
		/// <param name="oFormDocEntry"></param>
        public override void LoadForm(string oFormDocEntry)
        {
            MSXML2.DOMDocument oXmlDoc = new MSXML2.DOMDocument();

            try
            {
                PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();
                string mainJob = dataHelpClass.User_MainJob();

                //생산팀서무는 작업일보(작지)의 공정정보 매트릭스 컬럼 세팅을 FIX 시킴(전용화면 사용) (2016.03.16 송명규, 강주란 요청)
                if (mainJob == "생산팀서무")
                {
                    oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP040_01.srf");
                }
                else
                {
                    oXmlDoc.load(PSH_Globals.SP_Path + "\\" + PSH_Globals.Screen + "\\PS_PP040.srf");
                }

                oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue = oXmlDoc.selectSingleNode("Application/forms/action/form/@uid").nodeValue + "_" + (SubMain.Get_TotalFormsCount());
                oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@top").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);
                oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue = Convert.ToInt32(oXmlDoc.selectSingleNode("Application/forms/action/form/@left").nodeValue.ToString()) + (SubMain.Get_CurrentFormsCount() * 10);

                //매트릭스의 타이틀높이와 셀높이를 고정
                for (int i = 1; i <= (oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight").length); i++)
                {
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@titleHeight")[i - 1].nodeValue = 20;
                    oXmlDoc.selectNodes("Application/forms/action/form/items/action/item/specific/@cellHeight")[i - 1].nodeValue = 16;
                }

                oFormUniqueID = "PS_PP040_" + SubMain.Get_TotalFormsCount();
                SubMain.Add_Forms(this, oFormUniqueID, "PS_PP040");

                PSH_Globals.SBO_Application.LoadBatchActions(oXmlDoc.xml.ToString());
                oForm = PSH_Globals.SBO_Application.Forms.Item(oFormUniqueID);

                oForm.SupportedModes = -1;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                oForm.DataBrowser.BrowseBy = "DocEntry";

                oForm.Freeze(true);
                PS_PP040_CreateItems();
                PS_PP040_ComboBox_Setting();
                PS_PP040_CF_ChooseFromList();
                PS_PP040_EnableMenus();
                PS_PP040_SetDocument(oFormDocEntry);
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
        /// 화면 Item 생성
        /// </summary>
        private void PS_PP040_CreateItems()
        {
            try
            {
                oDS_PS_PP040H = oForm.DataSources.DBDataSources.Item("@PS_PP040H");
                oDS_PS_PP040L = oForm.DataSources.DBDataSources.Item("@PS_PP040L");
                oDS_PS_PP040M = oForm.DataSources.DBDataSources.Item("@PS_PP040M");
                oDS_PS_PP040N = oForm.DataSources.DBDataSources.Item("@PS_PP040N");

                oMat01 = oForm.Items.Item("Mat01").Specific;
                oMat01.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat01.AutoResizeColumns();

                oMat02 = oForm.Items.Item("Mat02").Specific;
                oMat02.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat02.AutoResizeColumns();

                oMat03 = oForm.Items.Item("Mat03").Specific;
                oMat03.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_NotSupported;
                oMat03.AutoResizeColumns();

                //기본매트릭스 선택용 라디오버튼
                oForm.DataSources.UserDataSources.Add("Opt01", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Add("Opt02", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Add("Opt03", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("Opt01").Specific.DataBind.SetBound(true, "", "Opt01");
                oForm.Items.Item("Opt02").Specific.DataBind.SetBound(true, "", "Opt02");
                oForm.Items.Item("Opt03").Specific.DataBind.SetBound(true, "", "Opt03");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt02");
                oForm.Items.Item("Opt01").Specific.GroupWith("Opt03");

                //거래처구분 콤보박스
                oForm.DataSources.UserDataSources.Add("CardType", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("CardType").Specific.DataBind.SetBound(true, "", "CardType");

                //전체사원보기 체크박스
                oForm.DataSources.UserDataSources.Add("EmpChk", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.Items.Item("EmpChk").Specific.DataBind.SetBound(true, "", "EmpChk");

                oDocType01 = "작업일보등록(작지)";
                if (oDocType01 == "작업일보등록(작지)")
                {
                    oForm.Items.Item("DocType").Specific.Select("10", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else if (oDocType01 == "작업일보등록(공정)")
                {
                    oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Combobox 설정
        /// </summary>
        private void PS_PP040_ComboBox_Setting()
        {
            string sQry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Items.Item("BPLId").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("BPLId").Specific, "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", false, false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "10", "일반");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "20", "PSMT지원");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "30", "외주");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "40", "실적");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "50", "일반조정");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "60", "외주조정");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "70", "설계시간");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "OrdType", "", "80", "외주제작지원");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OrdType").Specific, "PS_PP040", "OrdType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "DocType", "", "10", "작지기준");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "DocType", "", "20", "공정기준");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PP040", "DocType", false);

                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND CODE NOT IN('104','107','102','602') order by Code", "", false, false);
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("BPLId"), "SELECT BPLId, BPLName FROM OBPL order by BPLId", "", "");
                dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("OrdGbn"), "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' order by Code", "", "");

                //거래처구분 콤보(2012.02.02 송명규 추가)
                oForm.Items.Item("CardType").Specific.ValidValues.Add("%", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("CardType").Specific, "SELECT U_Minor, U_CdName FROM [@PS_SY001L] WHERE Code = 'C100' ORDER BY Code", "", false, false);
                oForm.Items.Item("CardType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                //작업구분코드(2014.04.15 송명규 수정)
                sQry = "  SELECT    U_Minor,";
                sQry += "           U_CdName";
                sQry += " FROM      [@PS_SY001L]";
                sQry += " WHERE     Code = 'P203'";
                sQry += "           AND U_UseYN = 'Y'";
                sQry += " ORDER BY  U_Seq";

                int workClsComboCount = oMat01.Columns.Item("WorkCls").ValidValues.Count - 1;

                if (workClsComboCount > 0)
                {
                    for (int loopCount = 0; loopCount <= workClsComboCount; loopCount++)
                    {
                        oMat01.Columns.Item("WorkCls").ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }

                    dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkCls"), sQry, "", "");
                }
                else
                {
                    dataHelpClass.GP_MatrixSetMatComboList(oMat01.Columns.Item("WorkCls"), sQry, "", "");
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// ChooseFromList 설정
        /// </summary>
        private void PS_PP040_CF_ChooseFromList()
        {
            ChooseFromListCollection oCFLs = null;
            Conditions oCons = null;
            Condition oCon = null;
            ChooseFromList oCFL = null;
            ChooseFromListCreationParams oCFLCreationParams = null;
            EditText oEdit = null;

            try
            {
                oEdit = oForm.Items.Item("ItemCode").Specific;
                oCFLs = oForm.ChooseFromLists;
                oCFLCreationParams = PSH_Globals.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLITEMCODE";
                oCFLCreationParams.MultiSelection = false;
                oCFL = oCFLs.Add(oCFLCreationParams);

                oCons = oCFL.GetConditions();
                oCon = oCons.Add();
                oCon.Alias = "ItmsGrpCod";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "102";
                oCFL.SetConditions(oCons);

                oEdit.ChooseFromListUID = "CFLITEMCODE";
                oEdit.ChooseFromListAlias = "ItemCode";
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
                if (oCons != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCons);
                }
                if (oCon != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCon);
                }
                if (oCFL != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFL);
                }
                if (oCFLCreationParams != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCFLCreationParams);
                }
                if (oEdit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEdit);
                }
            }
        }

        /// <summary>
        /// EnableMenus
        /// </summary>
        private void PS_PP040_EnableMenus()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                dataHelpClass.SetEnableMenus(oForm, false, false, true, true, false, true, true, true, true, true, false, false, false, false, false, false);
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// SetDocument
        /// </summary>
        /// <param name="oFormDocEntry">DocEntry</param>
        private void PS_PP040_SetDocument(string oFormDocEntry)
        {
            try
            {
                if (string.IsNullOrEmpty(oFormDocEntry))
                {
                    PS_PP040_FormItemEnabled();
                    PS_PP040_AddMatrixRow01(0, true);
                    PS_PP040_AddMatrixRow02(0, true);
                }
                else
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    PS_PP040_FormItemEnabled();
                    oForm.Items.Item("DocEntry").Specific.Value = oFormDocEntry;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각 모드에 따른 아이템설정
        /// </summary>
        private void PS_PP040_FormItemEnabled()
        {
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", false); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                    oForm.Items.Item("OrdType").Enabled = true;
                    oForm.Items.Item("OrdMgNum").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = true;
                    oForm.Items.Item("Mat02").Enabled = true;
                    oForm.Items.Item("Mat03").Enabled = true;
                    oMat02.Columns.Item("NTime").Editable = true;
                    oForm.Items.Item("BPLId").Specific.Select(dataHelpClass.User_BPLID(), SAPbouiCOM.BoSearchKey.psk_Index);
                    oForm.Items.Item("OrdType").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                    if (string.IsNullOrEmpty(oOrdGbn)) //oOrdGbn 변수 데이터할당 타이밍 재 확인 필요
                    {
                        oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    else
                    {
                        oForm.Items.Item("OrdGbn").Specific.Select(oOrdGbn, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }

                    PS_PP040_FormClear();

                    if (oDocType01 == "작업일보등록(작지)")
                    {
                        oDS_PS_PP040H.SetValue("U_DocType", 0, "10");
                    }
                    else if (oDocType01 == "작업일보등록(공정)")
                    {
                        oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    if (string.IsNullOrEmpty(oDocdate)) //oDocdate 변수 데이터할당 타이밍 재확인 필요
                    {
                        oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd");
                    }
                    else
                    {
                        oForm.Items.Item("DocDate").Specific.Value = oDocdate;
                    }
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                {
                    oForm.EnableMenu("1281", false); //찾기
                    oForm.EnableMenu("1282", true); //추가
                    oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = true;
                    oForm.Items.Item("OrdType").Enabled = true;
                    oForm.Items.Item("OrdMgNum").Enabled = true;
                    oForm.Items.Item("DocDate").Enabled = true;
                    oForm.Items.Item("Button01").Enabled = true;
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("Mat01").Enabled = false;
                    oForm.Items.Item("Mat02").Enabled = false;
                    oForm.Items.Item("Mat03").Enabled = false;
                }
                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.EnableMenu("1281", true); //찾기
                    oForm.EnableMenu("1282", true); //추가

                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oDS_PS_PP040H.GetValue("DocEntry", 0).ToString().Trim() + "'", 0, 1) == "Y")
                    {
                        oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("DocEntry").Enabled = false;
                        oForm.Items.Item("OrdType").Enabled = false;
                        oForm.Items.Item("OrdMgNum").Enabled = false;
                        oForm.Items.Item("DocDate").Enabled = false;
                        oForm.Items.Item("Button01").Enabled = false;
                        oForm.Items.Item("1").Enabled = false;
                        oForm.Items.Item("Mat01").Enabled = false;
                        oForm.Items.Item("Mat02").Enabled = false;
                        oForm.Items.Item("Mat03").Enabled = false;
                    }
                    else
                    {
                        if (oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "10" || oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "50" || oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "60" || oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "70")
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("OrdType").Enabled = false;
                            oForm.Items.Item("OrdMgNum").Enabled = true;
                            oForm.Items.Item("DocDate").Enabled = true;
                            oForm.Items.Item("Button01").Enabled = true;
                            oForm.Items.Item("1").Enabled = true;
                            oForm.Items.Item("Mat01").Enabled = true;
                            oForm.Items.Item("Mat02").Enabled = true;
                            oForm.Items.Item("Mat03").Enabled = true;
                        }
                        else if (oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "20")
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("OrdType").Enabled = false;
                            oForm.Items.Item("OrdMgNum").Enabled = true;
                            oForm.Items.Item("DocDate").Enabled = true;
                            oForm.Items.Item("Button01").Enabled = true;
                            oForm.Items.Item("1").Enabled = true;
                            oForm.Items.Item("Mat01").Enabled = true;
                            oForm.Items.Item("Mat02").Enabled = true;
                            oForm.Items.Item("Mat03").Enabled = true;
                        }
                        else if (oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "30")
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("OrdType").Enabled = false;
                            oForm.Items.Item("OrdMgNum").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("1").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("Mat02").Enabled = false;
                            oForm.Items.Item("Mat03").Enabled = false;
                        }
                        else if (oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "40")
                        {
                            oForm.Items.Item("Focus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oForm.Items.Item("DocEntry").Enabled = false;
                            oForm.Items.Item("OrdType").Enabled = false;
                            oForm.Items.Item("OrdMgNum").Enabled = false;
                            oForm.Items.Item("DocDate").Enabled = false;
                            oForm.Items.Item("Button01").Enabled = false;
                            oForm.Items.Item("1").Enabled = false;
                            oForm.Items.Item("Mat01").Enabled = false;
                            oForm.Items.Item("Mat02").Enabled = false;
                            oForm.Items.Item("Mat03").Enabled = false;
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
        /// DocEntry 초기화
        /// </summary>
        private void PS_PP040_FormClear()
        {
            string DocEntry;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                DocEntry = dataHelpClass.Get_ReData("AutoKey", "ObjectCode", "ONNM", "'PS_PP040'", "");

                if (string.IsNullOrEmpty(DocEntry) || DocEntry == "0")
                {
                    oForm.Items.Item("DocEntry").Specific.Value = 1;
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
        /// 메트릭스 Row추가(Mat01)
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param> 
        private void PS_PP040_AddMatrixRow01(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false) //행추가여부
                {
                    oDS_PS_PP040L.InsertRecord(oRow);
                }
                oMat01.AddRow();
                oDS_PS_PP040L.Offset = oRow;
                oDS_PS_PP040L.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oDS_PS_PP040L.SetValue("U_WorkCls", oRow, "A"); //작업구분을 기본으로 선택(2014.04.15 송명규 추가)
                oMat01.LoadFromDataSource();
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
        /// 메트릭스 Row추가(Mat02)
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param> 
        private void PS_PP040_AddMatrixRow02(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false) //행추가여부
                {
                    oDS_PS_PP040M.InsertRecord(oRow);
                }
                oMat02.AddRow();
                oDS_PS_PP040M.Offset = oRow;
                oDS_PS_PP040M.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat02.LoadFromDataSource();
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
        /// 메트릭스 Row추가(Mat03)
        /// </summary>
        /// <param name="oRow"></param>
        /// <param name="RowIserted"></param> 
        private void PS_PP040_AddMatrixRow03(int oRow, bool RowIserted)
        {
            try
            {
                oForm.Freeze(true);

                if (RowIserted == false) //행추가여부
                {
                    oDS_PS_PP040N.InsertRecord(oRow);
                }
                oMat03.AddRow();
                oDS_PS_PP040N.Offset = oRow;
                oDS_PS_PP040N.SetValue("U_LineNum", oRow, Convert.ToString(oRow + 1));
                oMat03.LoadFromDataSource();
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
        /// FormResize
        /// </summary>
        private void PS_PP040_FormResize()
        {
            try
            {
                oForm.Items.Item("Mat02").Top = 170;
                oForm.Items.Item("Mat02").Left = 7;
                oForm.Items.Item("Mat02").Height = ((oForm.Height - 170) / 3 * 1) - 20;
                oForm.Items.Item("Mat02").Width = oForm.Width / 2 - 14;

                oForm.Items.Item("Mat03").Top = 170;
                oForm.Items.Item("Mat03").Left = oForm.Width / 2;
                oForm.Items.Item("Mat03").Height = ((oForm.Height - 170) / 3 * 1) - 20;
                oForm.Items.Item("Mat03").Width = oForm.Width / 2 - 14;

                oForm.Items.Item("Mat01").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 40;
                oForm.Items.Item("Mat01").Left = 7;
                oForm.Items.Item("Mat01").Height = ((oForm.Height - 170) / 3 * 2) - 80;
                oForm.Items.Item("Mat01").Width = oForm.Width - 21;

                oForm.Items.Item("Opt01").Left = 10;
                oForm.Items.Item("Opt02").Left = oForm.Width / 2;
                oForm.Items.Item("Opt03").Left = 10;
                oForm.Items.Item("Opt03").Top = oForm.Items.Item("Mat03").Top + oForm.Items.Item("Mat03").Height + 20;

                oMat01.AutoResizeColumns();
                oMat02.AutoResizeColumns();
                oMat03.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 필수입력사항 체크
        /// </summary>
        /// <returns></returns>
        private bool PS_PP040_DataValidCheck()
        {
            bool returnValue = false;
            int i;
            int j;
            int failQty = 0;
            string sQry;
            string errMessage = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP040_FormClear();
                }

                if (Convert.ToInt32(dataHelpClass.GetValue("select Count(*) from OFPR Where '" + oForm.Items.Item("DocDate").Specific.Value + "' between F_RefDate and T_RefDate And PeriodStat = 'Y'", 0, 1)) > 0)
                {
                    errMessage = "해당일자는 전기기간이 잠겼습니다. 일자를 확인바랍니다.";
                    throw new Exception();
                }
                //마감상태 체크(원가)
                if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0,6)) == false)
                {
                    errMessage = "마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작업일보일자를 확인하고, 회계부서로 문의하세요.";
                    throw new Exception();
                }

                if (oForm.Items.Item("OrdType").Specific.Selected.Value != "10" 
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "20" 
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "50" 
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "60" 
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "70"
                 && oForm.Items.Item("OrdType").Specific.Selected.Value != "80")
                {
                    errMessage = "작업타입이 일반, PSMT지원, 조정, 설계, 외주제작지원이 아닙니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("OrdType").Specific.Selected.Value != "80") //외주제작지원은 헤더 작지번호 불필요
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                    {
                        errMessage = "작지번호는 필수입니다.";
                        oForm.Items.Item("OrdNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                }

                if (oMat01.VisualRowCount == 1)
                {
                    errMessage = "공정정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                if (oForm.Items.Item("OrdType").Specific.Selected.Value != "80") //외주제작지원은 작업자 정보 불필요
                {
                    if (oMat02.VisualRowCount == 1)
                    {
                        errMessage = "작업자정보 라인이 존재하지 않습니다.";
                        throw new Exception();
                    }
                }

                if (oMat03.VisualRowCount == 0)
                {
                    errMessage = "불량정보 라인이 존재하지 않습니다.";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "작지문서번호는 필수입니다.";
                        oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50" 
                          && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60"
                          && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "80")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) <= 0)
                        {
                            errMessage = "생산수량은 필수입니다.";
                            oMat01.Columns.Item("PQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50" 
                          && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60" 
                          && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "70"
                          && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "80")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value) <= 0)
                        {
                            errMessage = "실동시간은 필수입니다.";
                            oMat01.Columns.Item("WorkTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                    else if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "105" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "106") //작업완료여부(2012.02.02. 송명규 추가)(기계공구, 몰드일 경우만 작업완료여부 필수 체크)
                    {
                        if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "80")
                        {
                            if (oMat01.Columns.Item("CompltYN").Cells.Item(i).Specific.Value.ToString().Trim() == "%")
                            {
                                errMessage = "작업구분이 기계공구, 몰드일경우는 작업완료여부가 필수입니다. 확인하십시오.";
                                oMat01.Columns.Item("CompltYN").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                throw new Exception();
                            }
                        }
                    }

                    //불량수량 검사
                    for (j = 1; j <= oMat03.VisualRowCount; j++)
                    {
                        //불량코드 입력 여부 check
                        if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) != 0 && string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value.ToString().Trim()))
                        {
                            errMessage = "불량수량이 입력되었을 때는 불량코드는 필수입니다.";
                            oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                        else if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) == 0 && !string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value.ToString().Trim()))
                        {
                            errMessage = "불량코드를 확인하세요.";
                            oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }

                        if ((oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value) && (oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.Value))
                        {
                            failQty += Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value);
                        }

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (oMat01.Columns.Item("CpCode").Cells.Item(j).Specific.Value.ToString() == "CP10105" || oMat01.Columns.Item("CpCode").Cells.Item(j).Specific.Value.ToString().Trim() == "CP20402")
                            {
                                sQry = "  SELECT    U_TeamCode ";
                                sQry += " FROM      [@PH_PY001A] ";
                                sQry += " WHERE     CODE IN (";
                                sQry += "                       SELECT  U_MSTCOD ";
                                sQry += "                       FROM    OHEM ";
                                sQry += "                       WHERE   userId IN (";
                                sQry += "                                           SELECT USERID";
                                sQry += "                                           FROM OUSR ";
                                sQry += "                                           WHERE USER_CODE ='" + PSH_Globals.oCompany.UserName + "'";
                                sQry += "                                          )";
                                sQry += "                   )";

                                if (dataHelpClass.GetValue(sQry, 0, 1) != "2600")
                                {
                                    errMessage = "기계사업부 품질팀만 등록 및 수정이 가능합니다.";
                                    throw new Exception();
                                }
                            }
                        }
                    }

                    if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50" && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value) != failQty)
                        {
                            errMessage = "공정리스트의 불량수량과 불량정보의 불량수량이 일치하지 않습니다.";
                            throw new Exception();
                        }
                    }

                    if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "601" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "111")
                    {
                        if (Convert.ToInt32(oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value) == 1 && string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            errMessage = "공정 사용 원재료코드가 없습니다. 사용 원재료를 선택해 주세요.";
                            throw new Exception();
                        }
                    }
                }

                //작업자 테이블 필수 정보 체크
                for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat02.Columns.Item("WorkCode").Cells.Item(i).Specific.Value))
                    {
                        errMessage = "작업자 사번이 없습니다.";
                        oMat02.Columns.Item("WorkCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        throw new Exception();
                    }

                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "비가동코드가 입력되었을 때는 비가동시간은 필수입니다.";
                            oMat02.Columns.Item("NTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }

                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value))
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                        {
                            errMessage = "비가동시간이 입력되었을 때는 비가동코드는 필수입니다.";
                            oMat02.Columns.Item("NCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            throw new Exception();
                        }
                    }
                }

                if (PS_PP040_Validate("검사01") == false)
                {
                    errMessage = "";
                    throw new Exception();
                }

                oDS_PS_PP040L.RemoveRecord(oDS_PS_PP040L.Size - 1);
                oMat01.LoadFromDataSource();
                oDS_PS_PP040M.RemoveRecord(oDS_PS_PP040M.Size - 1);
                oMat02.LoadFromDataSource();

                returnValue = true;
            }
            catch (Exception ex)
            {
                if (errMessage != string.Empty)
                {
                    PSH_Globals.SBO_Application.MessageBox(errMessage);
                }
                else if (errMessage == "")
                {
                    //처리 없음
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + (char)13 + ex.Message);
                }
            }
            finally
            {
            }

            return returnValue;
        }

        /// <summary>
        /// 처리가능한 Action인지 검사
        /// </summary>
        /// <param name="ValidateType"></param>
        /// <returns></returns>
        private bool PS_PP040_Validate(string ValidateType)
        {
            bool returnValue = false;
            int i;
            int j;
            string Query01;
            double PrevDBCpQty;
            double PrevMATRIXCpQty;
            double CurrentDBCpQty;
            double CurrentMATRIXCpQty;
            string PrevCpInfo;
            string CurrentCpInfo;
            string OrdMgNum;
            bool Exist;
            string LineNum;
            string DocEntry;
            string errMessage = string.Empty;

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                    {
                        errMessage = "해당문서는 다른사용자에 의해 취소되었습니다. 작업을 진행할 수 없습니다.";
                        throw new Exception();
                    }
                }

                //작업타입이 일반,조정인경우
                if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "10" || oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "50" || oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "60")
                {
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "20") //작업타입이 PSMT지원인경우
                {
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "30" || oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "40") //작업타입이 외주, 실적인경우
                {
                    errMessage = "해당작업타입은 변경이 불가능합니다.";
                    throw new Exception();
                }

                if (ValidateType == "검사01")
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "10") //작업타입이 일반인경우
                    {
                        for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1)) <= 0)
                            {
                                errMessage = "작업지시문서가 존재하지 않습니다.";
                                throw new Exception();
                            }
                        }

                        if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            //삭제된 행에 대한처리
                            Query01 = "  SELECT     PS_PP040H.DocEntry,";
                            Query01 += "            PS_PP040L.LineId,";
                            Query01 += "            CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,";
                            Query01 += "            PS_PP040L.U_OrdGbn AS OrdGbn,";
                            Query01 += "            PS_PP040L.U_PP030HNo AS PP030HNo,";
                            Query01 += "            PS_PP040L.U_PP030MNo AS PP030MNo,";
                            Query01 += "            PS_PP040L.U_OrdMgNum AS OrdMgNum ";
                            Query01 += " FROM       [@PS_PP040H] PS_PP040H";
                            Query01 += "            LEFT JOIN";
                            Query01 += "            [@PS_PP040L] PS_PP040L";
                            Query01 += "                ON PS_PP040H.DocEntry = PS_PP040L.DocEntry ";
                            Query01 += " WHERE      PS_PP040L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                            RecordSet01.DoQuery(Query01);
                            for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                            {
                                Exist = false;
                                for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                {
                                    if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value)) //새로추가된 행인경우, 검사미필요
                                    {
                                    }
                                    else
                                    {
                                        //라인번호가 같고, 문서번호가 같으면 존재하는행
                                        if (Convert.ToInt32(RecordSet01.Fields.Item(0).Value) == Convert.ToInt32(oForm.Items.Item("DocEntry").Specific.Value) && Convert.ToInt32(RecordSet01.Fields.Item(1).Value) == Convert.ToInt32(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value))
                                        {
                                            Exist = true;
                                        }
                                    }
                                }

                                if (Exist == false) //삭제된 행중 수량관계를 알아본다.
                                {
                                    //휘팅이면서
                                    if (RecordSet01.Fields.Item("OrdGbn").Value == "101")
                                    {
                                        //현재 공정이 실적공정이면..
                                        if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y")
                                        {
                                            //휘팅벌크포장
                                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP070L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                            {
                                                errMessage = "삭제된행이 생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                                throw new Exception();
                                            }
                                            //휘팅실적
                                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                            {
                                                errMessage = "삭제된행이 생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                                throw new Exception();
                                            }
                                        }
                                    }

                                    if (RecordSet01.Fields.Item("OrdGbn").Value == "105" || RecordSet01.Fields.Item("OrdGbn").Value == "106") //기계공구,몰드
                                    {
                                        //입력가능
                                    }
                                    else if (RecordSet01.Fields.Item("OrdGbn").Value == "101" || RecordSet01.Fields.Item("OrdGbn").Value == "102") //휘팅,부품
                                    {
                                        //삭제된 행에 대한 검사
                                        OrdMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
                                        CurrentCpInfo = OrdMgNum;
                                        PrevCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_02 '" + OrdMgNum + "'", 0, 1);

                                        if (string.IsNullOrEmpty(PrevCpInfo))
                                        {
                                            //해당공정이 첫공정이면 입력가능
                                        }
                                        else
                                        {
                                            PrevDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP040H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                            //재공이동 수량
                                            PrevDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                            PrevDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                            PrevMATRIXCpQty = 0;
                                            for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                            {
                                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == PrevCpInfo)
                                                {
                                                    PrevMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                                }
                                            }

                                            CurrentDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP040L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                            CurrentDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                            CurrentDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                            CurrentMATRIXCpQty = 0;
                                            for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                            {
                                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == CurrentCpInfo)
                                                {
                                                    CurrentMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                                }
                                            }
                                            if ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))
                                            {
                                                errMessage = "삭제된 공정의 선행공정의 생산수량이 삭제된 공정의 생산수량을 미달합니다.";
                                                throw new Exception();
                                            }
                                        }
                                    }
                                }
                                RecordSet01.MoveNext();
                            }
                        }

                        if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value)) //새로추가된 행인경우
                                {
                                    //검사 불필요
                                }
                                else
                                {
                                    if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "101") //휘팅이면서
                                    {
                                        if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oForm.Items.Item("DocEntry").Specific.Value + "-" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1) == "Y")
                                        {
                                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0
                                             || Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'", 0, 1)) > 0)
                                            {
                                                //작업일보등록된문서중에 수정이 된문서를 구함
                                                Query01 = "  SELECT     PS_PP040L.U_OrdMgNum,";
                                                Query01 += "            PS_PP040L.U_Sequence,";
                                                Query01 += "            PS_PP040L.U_CpCode,";
                                                Query01 += "            PS_PP040L.U_ItemCode,";
                                                Query01 += "            PS_PP040L.U_PP030HNo,";
                                                Query01 += "            PS_PP040L.U_PP030MNo,";
                                                Query01 += "            PS_PP040L.U_PQty,";
                                                Query01 += "            PS_PP040L.U_NQty,";
                                                Query01 += "            PS_PP040L.U_ScrapWt,";
                                                Query01 += "            PS_PP040L.U_WorkTime";
                                                Query01 += " FROM       [@PS_PP040H] PS_PP040H";
                                                Query01 += "            LEFT JOIN";
                                                Query01 += "            [@PS_PP040L] PS_PP040L";
                                                Query01 += "                ON PS_PP040H.DocEntry = PS_PP040L.DocEntry";
                                                Query01 += " WHERE      PS_PP040H.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                                                Query01 += "            AND PS_PP040L.LineId = '" + oMat01.Columns.Item("LineId").Cells.Item(i).Specific.Value + "'";
                                                Query01 += "            AND PS_PP040H.Canceled = 'N'";
                                                RecordSet01.DoQuery(Query01);

                                                if (RecordSet01.Fields.Item(0).Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(1).Value == oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(2).Value == oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(3).Value == oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(4).Value == oMat01.Columns.Item("PP030HNo").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(5).Value == oMat01.Columns.Item("PP030MNo").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(6).Value == oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(7).Value == oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(8).Value == oMat01.Columns.Item("ScrapWt").Cells.Item(i).Specific.Value
                                                 && RecordSet01.Fields.Item(9).Value == oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value) //값이 변경되지 않은 경우
                                                {
                                                    //수정가능
                                                }
                                                else //값이 변경된 행의 경우
                                                {
                                                    errMessage = "생산실적이 등록된 행은 수정할 수 없습니다.";
                                                    throw new Exception();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        for (i = 1; i <= oMat01.VisualRowCount - 1; i++) //입력된 모든행에 대해 입력가능성 검사
                        {
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "105" || oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "106") //기계공구,몰드
                            {
                                //입력 가능
                            }
                            else if (oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "101" || oMat01.Columns.Item("OrdGbn").Cells.Item(i).Specific.Value == "102") //휘팅,부품
                            {
                                OrdMgNum = oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value;
                                CurrentCpInfo = OrdMgNum;
                                PrevCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_02 '" + OrdMgNum + "'", 0, 1);

                                if (string.IsNullOrEmpty(PrevCpInfo))
                                {
                                    //해당공정이 첫공정이면 입력 가능
                                }
                                else
                                {
                                    PrevDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("EXEC PS_PP040_07 '" + PrevCpInfo + "', '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1));
                                    //재공 이동수량 반영
                                    PrevDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    PrevDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                    PrevMATRIXCpQty = 0;
                                    for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == PrevCpInfo)
                                        {
                                            PrevMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                        }
                                    }
                                    CurrentDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("EXEC PS_PP040_07 '" + CurrentCpInfo + "', '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1));
                                    CurrentDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    CurrentDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                    CurrentMATRIXCpQty = 0;
                                    for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == CurrentCpInfo)
                                        {
                                            CurrentMATRIXCpQty += Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                        }
                                    }

                                    if ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))
                                    {
                                        oMat01.SelectRow(i, true, false);
                                        errMessage = "선행공정의 생산수량이 현공정의 생산수량에 미달 합니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "20") //작업타입이 PSMT지원인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "30") //작업타입이 외주인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "40") //작업타입이 실적인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "50") //작업타입이 조정인경우
                    {
                        //별도 조치 불필요
                    }
                }
                else if (ValidateType == "행삭제01") //행삭제전 행삭제가능여부검사
                {
                    //작업타입이 일반인경우
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "10")
                    {
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value)) //새로추가된 행인경우
                        {
                            //삭제 가능
                        }
                        else
                        {
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "101") //휘팅
                            {
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "-" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1) == "Y")
                                {
                                    //휘팅벌크포장
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "삭제된행이 생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }

                                    //휘팅실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "삭제된행이 생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                            else if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "105" || oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "106") //기계공구,몰드
                            {
                                //재고가 존재하면 행삭제 불가 기능 추가(2011.12.15 송명규 추가)
                                Query01 = "  SELECT     SUM(A.InQty) - SUM(A.OutQty) AS [StockQty]";
                                Query01 += " FROM       OINM AS A";
                                Query01 += "            INNER JOIN";
                                Query01 += "            OITM As B";
                                Query01 += "                ON A.ItemCode = B.ItemCode";
                                Query01 += " WHERE      B.U_ItmBsort IN ('105','106')";
                                Query01 += "            AND A.ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(oMat01Row01).Specific.Value + "'";
                                Query01 += " GROUP BY  A.ItemCode";

                                string stockQty = string.IsNullOrEmpty(dataHelpClass.GetValue(Query01, 0, 1)) ? "0" : dataHelpClass.GetValue(Query01, 0, 1);

                                if (Convert.ToDouble(stockQty) > 0)
                                {
                                    errMessage = "재고가 존재하는 작번입니다. 삭제할 수 없습니다.";
                                    throw new Exception();
                                }
                            }
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "20") //작업타입이 PSMT인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "30") //작업타입이 외주인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "40") //작업타입이 실적인경우
                    {
                        //별도 조치 불필요   
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "50") //작업타입이 조정인경우
                    {
                        //별도 조치 불필요
                    }
                }
                else if (ValidateType == "수정01") //수정전 수정가능여부검사
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "10") //작업타입이 일반인경우
                    {
                        //oMat01.VisualRowCount가 1인 경우는 최초 행추가이므로 빈문자열("") 반환, 그게 아닐경우만 Matrix의 LineID 반환(matrix index 오류 처리)
                        string tempLineID = oMat01.VisualRowCount == 1 ? "" : oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value;

                        if (string.IsNullOrEmpty(tempLineID)) //새로 추가된 행인경우
                        {
                            //수정 가능
                        }
                        else
                        {
                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "111" || oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "601") //분말
                            {
                                if (oMat01.Columns.Item("CpCode").Cells.Item(oMat01Row01).Specific.Value == "CP80111" || oMat01.Columns.Item("CpCode").Cells.Item(oMat01Row01).Specific.Value == "CP80101")
                                {
                                    DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                                    LineNum = oMat01.Columns.Item("LineNum").Cells.Item(oMat01Row01).Specific.Value.ToString().Trim();

                                    if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(oMat01Row01).Specific.Value.ToString().Trim()) != Convert.ToDouble(dataHelpClass.GetValue("select U_pqty from [@PS_PP040L] where DocEntry ='" + DocEntry + "' and u_linenum ='" + LineNum + "'", 0, 1)))
                                    {
                                        errMessage = "원자재 불출이 진행된 행은 생산수량을 수정할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                            }

                            if (oMat01.Columns.Item("OrdGbn").Cells.Item(oMat01Row01).Specific.Value == "101") //휘팅
                            {
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + oMat01.Columns.Item("PP030HNo").Cells.Item(oMat01Row01).Specific.Value + "-" + oMat01.Columns.Item("PP030MNo").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1) == "Y")
                                {
                                    //휘팅벌크포장
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP070L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "수정된행이 생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }

                                    //휘팅실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "수정된행이 생산실적 등록된 행입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "20") //작업타입이 PSMT인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "30") //작업타입이 외주인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "40") //작업타입이 실적인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "50") //작업타입이 조정인경우
                    {
                        //별도 조치 불필요
                    }
                }
                else if (ValidateType == "취소")
                {
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                    {
                        errMessage = "이미취소된 문서 입니다. 취소할 수 없습니다.";
                        throw new Exception();
                    }

                    if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "10") //작업타입이 일반인경우
                    {
                        //삭제된 행에 대한처리
                        Query01 = "  SELECT     PS_PP040H.DocEntry,";
                        Query01 += "            PS_PP040L.LineId,";
                        Query01 += "            CONVERT(NVARCHAR,PS_PP040H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP040L.LineId) AS DocInfo,";
                        Query01 += "            PS_PP040L.U_OrdGbn AS OrdGbn,";
                        Query01 += "            PS_PP040L.U_PP030HNo AS PP030HNo,";
                        Query01 += "            PS_PP040L.U_PP030MNo AS PP030MNo,";
                        Query01 += "            PS_PP040L.U_OrdMgNum AS OrdMgNum ";
                        Query01 += " FROM       [@PS_PP040H] PS_PP040H";
                        Query01 += "            LEFT JOIN";
                        Query01 += "            [@PS_PP040L] PS_PP040L";
                        Query01 += "                ON PS_PP040H.DocEntry = PS_PP040L.DocEntry ";
                        Query01 += " WHERE      PS_PP040L.DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'";
                        RecordSet01.DoQuery(Query01);

                        for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
                        {
                            if (RecordSet01.Fields.Item("OrdGbn").Value == "101") //휘팅
                            {
                                if (dataHelpClass.GetValue("EXEC PS_PP040_05 '" + RecordSet01.Fields.Item("OrdMgNum").Value + "'", 0, 1) == "Y") //현재공정이 실적포인트이면
                                {
                                    //휘팅벌크포장
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP070H] PS_PP070H LEFT JOIN [@PS_PP070L] PS_PP070L ON PS_PP070H.DocEntry = PS_PP070L.DocEntry WHERE PS_PP070H.Canceled = 'N' AND PS_PP070L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP070L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "생산실적 등록된 문서입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }

                                    //휘팅실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                    {
                                        errMessage = "생산실적 등록된 문서입니다. 적용할 수 없습니다.";
                                        throw new Exception();
                                    }
                                }
                            }

                            if (RecordSet01.Fields.Item("OrdGbn").Value == "105" || RecordSet01.Fields.Item("OrdGbn").Value == "106") //기계공구,몰드
                            {
                                //입력가능
                            }
                            else if (RecordSet01.Fields.Item("OrdGbn").Value == "101" || RecordSet01.Fields.Item("OrdGbn").Value == "102") //휘팅,부품
                            {
                                //삭제된 행에 대한 검사
                                OrdMgNum = RecordSet01.Fields.Item("OrdMgNum").Value;
                                CurrentCpInfo = OrdMgNum;
                                PrevCpInfo = dataHelpClass.GetValue("EXEC PS_PP040_02 '" + OrdMgNum + "'", 0, 1);

                                if (string.IsNullOrEmpty(PrevCpInfo))
                                {
                                    //해당공정이 첫공정이면 입력 가능
                                }
                                else
                                {
                                    PrevDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP040H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                    PrevDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    PrevDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    PrevMATRIXCpQty = 0;

                                    CurrentDBCpQty = Convert.ToDouble(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP040L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                    CurrentDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    CurrentDBCpQty += Convert.ToDouble(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    CurrentMATRIXCpQty = 0;

                                    if ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))
                                    {
                                        errMessage = "취소문서의 선행공정의 생산수량이 취소문서의 생산수량을 미달합니다.";
                                        throw new Exception();
                                    }
                                }
                            }
                            RecordSet01.MoveNext();
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "20") //작업타입이 PSMT인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "30") //작업타입이 외주인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "40") //작업타입이 실적인경우
                    {
                        //별도 조치 불필요
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "50") //작업타입이 조정인경우
                    {
                        //별도 조치 불필요
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
        /// 근무시간 총합계산
        /// </summary>
        private void PS_PP040_SumWorkTime()
        {
            short loopCount;
            double total = 0;

            try
            {
                for (loopCount = 0; loopCount <= oMat01.RowCount - 2; loopCount++)
                {
                    total += Convert.ToDouble(string.IsNullOrEmpty(oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.Value.ToString().Trim()) ? 0 : oMat01.Columns.Item("WorkTime").Cells.Item(loopCount + 1).Specific.Value.ToString().Trim());
                }

                oForm.Items.Item("Total").Specific.Value = total.ToString("#,##0.##");
            }
            catch (Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// OrderInfoLoad
        /// </summary>
        private void PS_PP040_OrderInfoLoad()
        {
            string Query01;
            string errCode = string.Empty;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "10"
                 || oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "50"
                 || oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "60"
                 || oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "70") //일반,조정, 설계
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.Value))
                    {
                        errCode = "1";
                        throw new Exception();
                    }
                    else
                    {
                        Query01 = " SELECT      U_OrdGbn,";
                        Query01 += "             U_BPLId,";
                        Query01 += "            U_ItemCode,";
                        Query01 += "            U_ItemName,";
                        Query01 += "            U_OrdNum,";
                        Query01 += "            U_OrdSub1,";
                        Query01 += "            U_OrdSub2,";
                        Query01 += "            DocEntry";
                        Query01 += " FROM       [@PS_PP030H]";
                        Query01 += " WHERE      U_OrdNum + U_OrdSub1 + U_OrdSub2 = '" + oForm.Items.Item("OrdMgNum").Specific.Value + "'";
                        Query01 += "            AND U_OrdGbn NOT IN('104','107') ";
                        Query01 += "            AND Canceled = 'N'";
                        RecordSet01.DoQuery(Query01);

                        if (RecordSet01.RecordCount == 0)
                        {
                            errCode = "2";
                            throw new Exception();
                        }
                        else
                        {
                            oForm.Items.Item("OrdGbn").Specific.Select(RecordSet01.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("BPLId").Specific.Select(RecordSet01.Fields.Item(1).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            oForm.Items.Item("ItemCode").Specific.Value = RecordSet01.Fields.Item(2).Value;
                            oForm.Items.Item("ItemName").Specific.Value = RecordSet01.Fields.Item(3).Value;
                            oForm.Items.Item("OrdNum").Specific.Value = RecordSet01.Fields.Item(4).Value;
                            oForm.Items.Item("OrdSub1").Specific.Value = RecordSet01.Fields.Item(5).Value;
                            oForm.Items.Item("OrdSub2").Specific.Value = RecordSet01.Fields.Item(6).Value;
                            oForm.Items.Item("PP030HNo").Specific.Value = RecordSet01.Fields.Item(7).Value;
                            oForm.Update();
                        }
                    }
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "20") //PSMT
                {
                    if (string.IsNullOrEmpty(oForm.Items.Item("OrdMgNum").Specific.Value))
                    {
                        errCode = "1";
                        throw new Exception();
                    }
                    else
                    {
                        oForm.Items.Item("OrdNum").Specific.Value = oForm.Items.Item("OrdMgNum").Specific.Value;
                        oForm.Items.Item("OrdSub1").Specific.Value = "000";
                        oForm.Items.Item("OrdSub2").Specific.Value = "00";
                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();
                        PS_PP040_AddMatrixRow01(0, true);
                        oMat02.Clear();
                        oMat02.FlushToDataSource();
                        oMat02.LoadFromDataSource();
                        PS_PP040_AddMatrixRow02(0, true);
                        oMat03.Clear();
                        oMat03.FlushToDataSource();
                        oMat03.LoadFromDataSource();
                        oForm.Update();
                    }
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "30")
                {
                    errCode = "4";
                    throw new Exception();
                }
                else if (oForm.Items.Item("OrdType").Specific.Selected.Value == "40")
                {
                    errCode = "5";
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업지시 관리번호를 입력하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업지시 정보가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("외주는 입력할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("실적은 입력할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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
        }

        /// <summary>
        /// FindValidateDocument : 작업일보등록(작지) 문서인지 조회
        /// </summary>
        /// <param name="ObjectType"></param>
        /// <returns></returns>
        private bool PS_PP040_FindValidateDocument(string ObjectType)
        {
            bool returnValue = false;
            string Query01;
            string errMessage = string.Empty;
            string DocEntry;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            try
            {
                DocEntry = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();

                Query01 = "  SELECT     DocEntry";
                Query01 += " FROM       [" + ObjectType + "]";
                Query01 += " WHERE      DocEntry = " + DocEntry;
                if (oDocType01 == "작업일보등록(작지)")
                {
                    Query01 += " AND U_DocType = '10'";
                }
                else if (oDocType01 == "작업일보등록(공정)")
                {
                    Query01 += " AND U_DocType = '20'";
                }
                RecordSet01.DoQuery(Query01);

                if (RecordSet01.RecordCount == 0)
                {
                    if (oDocType01 == "작업일보등록(작지)")
                    {
                        errMessage = "작업일보등록(공정)문서 이거나 존재하지 않는 문서입니다";
                    }
                    else if (oDocType01 == "작업일보등록(공정)")
                    {
                        errMessage = "작업일보등록(작지)문서 이거나 존재하지 않는 문서입니다.";
                    }

                    throw new Exception();
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
        /// 출고 DI API(원재료 출고용)
        /// </summary>
        /// <returns></returns>
        private bool PS_PP040_AddoInventoryGenExit()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int j;
            int i;
            int Cnt = 0;
            int RetVal;
            string CpCode;
            string DocNum;
            string DocDate;
            string CItemCod;
            string WhsCode;
            double IssueQty;
            double IssueWt;
            string SDocEntry;
            string sQry;
            double Price;

            SAPbobsCOM.Documents DI_oInventoryGenExit = null; //재고출고 DI객체
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClss = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                oMat01.FlushToDataSource();
                DocDate = dataHelpClass.ConvertDateType(oDS_PS_PP040H.GetValue("U_DocDate", 0), "-");
                DocNum = oDS_PS_PP040H.GetValue("DocEntry", 0).ToString().Trim();

                if (string.IsNullOrEmpty(oMat01.Columns.Item("OutDoc").Cells.Item(1).Specific.Value))
                {
                    DI_oInventoryGenExit = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                    DI_oInventoryGenExit.DocDate = Convert.ToDateTime(DocDate);
                    DI_oInventoryGenExit.TaxDate = Convert.ToDateTime(DocDate);
                    DI_oInventoryGenExit.Comments = "원재료 불출 등록(" + DocNum + ") 출고_PS_PP040";

                    j = 0;
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        sQry = "  SELECT    PRICE";
                        sQry += " FROM      OIVL a";
                        sQry += "           INNER JOIN";
                        sQry += "           OIGN b";
                        sQry += "               ON a.BASE_REF = b.DocEntry";
                        sQry += "               and b.U_Comments ='Convert Meterial'";
                        sQry += " WHERE     a.ITEMCODE = '" + oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value + "'";
                        sQry += "           AND CONVERT(CHAR(6),a.DocDate,112) ='" + codeHelpClss.Left(oDS_PS_PP040H.GetValue("U_DocDate", 0), 6) + "'";

                        oRecordSet.DoQuery(sQry);

                        CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                        IssueQty = Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value);
                        IssueWt = Convert.ToDouble(oMat01.Columns.Item("PWeight").Cells.Item(i + 1).Specific.Value);
                        CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                        Price = Convert.ToString(oRecordSet.Fields.Item(0).Value) == "" ? 0 : Convert.ToDouble(oRecordSet.Fields.Item(0).Value);

                        WhsCode = "101";

                        if ((CpCode == "CP80101" || CpCode == "CP80111") && !string.IsNullOrEmpty(CItemCod) && IssueQty >= 0 && IssueWt != 0 && !string.IsNullOrEmpty(WhsCode))
                        {
                            if (j > 0)
                            {
                                DI_oInventoryGenExit.Lines.Add();
                            }
                            DI_oInventoryGenExit.Lines.SetCurrentLine(j);
                            DI_oInventoryGenExit.Lines.ItemCode = CItemCod;
                            DI_oInventoryGenExit.Lines.WarehouseCode = WhsCode;
                            DI_oInventoryGenExit.Lines.Quantity = IssueWt;
                            DI_oInventoryGenExit.Lines.UserFields.Fields.Item("U_Qty").Value = IssueQty;

                            if (oRecordSet.EoF) //제품원재료 변환 품목은 단가를 계산 후 입력
                            {
                            }
                            else
                            {
                                DI_oInventoryGenExit.Lines.Price = Price;
                                DI_oInventoryGenExit.Lines.UnitPrice = Price;
                                DI_oInventoryGenExit.Lines.LineTotal = Price * IssueWt;
                            }

                            Cnt += 1;
                            j += 1;
                        }
                    }

                    //완료
                    if (Cnt > 0)
                    {
                        RetVal = DI_oInventoryGenExit.Add();
                        if (0 != RetVal)
                        {
                            PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                            errCode = "1";
                            throw new Exception();
                        }
                        else
                        {
                            PSH_Globals.oCompany.GetNewObjectCode(out SDocEntry);
                            Cnt = 1;
                            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                            {
                                CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.Value;
                                if (CpCode == "CP80101" || CpCode == "CP80111")
                                {
                                    oDS_PS_PP040L.SetValue("U_OutDoc", i, SDocEntry);
                                    oDS_PS_PP040L.SetValue("U_OutLin", i, Convert.ToString(Cnt));
                                    Cnt += 1;
                                }
                            }
                        }
                    }

                    oMat01.LoadFromDataSource();
                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                if (DI_oInventoryGenExit != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenExit);
                }
            }

            return returnValue;
        }

        /// <summary>
        /// 입고 DI API(원재료 입고용(출고 취소))
        /// </summary>
        /// <returns></returns>
        private bool PS_PP040_AddoInventoryGenEntry()
        {
            bool returnValue = false;
            string errCode = string.Empty;
            string errDIMsg = string.Empty;
            int errDICode = 0;
            int j;
            int i;
            int Cnt = 0;
            int RetVal;
            string CpCode;
            string DocNum;
            string DocDate;
            string CItemCod;
            string WhsCode;
            double IssueQty;
            double IssueWt;
            string SDocEntry = string.Empty;
            string sQry;
            string OIGEDoc;
            double Price;

            SAPbobsCOM.Documents DI_oInventoryGenEntry = null;
            SAPbobsCOM.Recordset oRecordSet = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_CodeHelpClass codeHelpClss = new PSH_CodeHelpClass();
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (PSH_Globals.oCompany.InTransaction == true)
                {
                    PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                PSH_Globals.oCompany.StartTransaction();

                //현재월의 전기기간 체크 후 잠겨있으면 DI API 미실행
                if (dataHelpClass.Get_ReData("PeriodStat", "[NAME]", "OFPR", "'" + DateTime.Now.ToString("yyyy") + "-" + DateTime.Now.ToString("MM") + "'", "") == "Y")
                {
                    errCode = "2";
                    throw new Exception();
                }

                oMat01.FlushToDataSource();

                DocDate = dataHelpClass.ConvertDateType(oDS_PS_PP040H.GetValue("U_DocDate", 0), "-");
                DocNum = oDS_PS_PP040H.GetValue("DocEntry", 0).ToString().Trim();
                OIGEDoc = oMat01.Columns.Item("OutDoc").Cells.Item(1).Specific.Value.ToString().Trim();

                if (string.IsNullOrEmpty(oMat01.Columns.Item("OutDocC").Cells.Item(1).Specific.Value))
                {
                    DI_oInventoryGenEntry = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

                    DI_oInventoryGenEntry.DocDate = Convert.ToDateTime(DocDate);
                    DI_oInventoryGenEntry.TaxDate = Convert.ToDateTime(DocDate);
                    DI_oInventoryGenEntry.Comments = "원재료 불출 등록 출고 취소 (" + DocNum + ") 입고_PS_PP040";
                    DI_oInventoryGenEntry.UserFields.Fields.Item("U_CancDoc").Value = OIGEDoc; //입고취소 문서번호

                    j = 0;
                    for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
                    {
                        sQry = "  SELECT    PRICE";
                        sQry += " FROM      OIVL a";
                        sQry += "           INNER JOIN";
                        sQry += "           OIGN b";
                        sQry += "               ON a.BASE_REF = b.DocEntry";
                        sQry += "               AND b.U_Comments = 'Convert Meterial'";
                        sQry += " WHERE     a.ITEMCODE = '" + oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value + "'";
                        sQry += "           AND CONVERT(CHAR(6), a.DocDate,112) = '" + codeHelpClss.Left(oDS_PS_PP040H.GetValue("U_DocDate", 0), 6) + "'";

                        oRecordSet.DoQuery(sQry);

                        CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                        IssueQty = Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value);
                        IssueWt = Convert.ToDouble(oMat01.Columns.Item("PWeight").Cells.Item(i + 1).Specific.Value);
                        CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.Value.ToString().Trim();
                        Price = Convert.ToString(oRecordSet.Fields.Item(0).Value) == "" ? 0 : Convert.ToDouble(oRecordSet.Fields.Item(0).Value);

                        WhsCode = "101";

                        if ((CpCode == "CP80101" || CpCode == "CP80111") && !string.IsNullOrEmpty(CItemCod) && IssueQty >= 0 && IssueWt != 0 && !string.IsNullOrEmpty(WhsCode))
                        {
                            if (j > 0)
                            {
                                DI_oInventoryGenEntry.Lines.Add();
                            }
                            DI_oInventoryGenEntry.Lines.SetCurrentLine(j);
                            DI_oInventoryGenEntry.Lines.ItemCode = CItemCod;
                            DI_oInventoryGenEntry.Lines.WarehouseCode = WhsCode;
                            DI_oInventoryGenEntry.Lines.Quantity = IssueWt;
                            DI_oInventoryGenEntry.Lines.UserFields.Fields.Item("U_Qty").Value = IssueQty;

                            if (oRecordSet.EoF) //제품원재료 변환 품목은 단가를 계산 후 입력
                            {
                            }
                            else
                            {
                                DI_oInventoryGenEntry.Lines.Price = Price;
                                DI_oInventoryGenEntry.Lines.UnitPrice = Price;
                                DI_oInventoryGenEntry.Lines.LineTotal = Price * IssueWt;
                            }

                            Cnt += 1;
                            j += 1;
                        }
                    }

                    //완료
                    if (Cnt > 0)
                    {
                        RetVal = DI_oInventoryGenEntry.Add();
                        if (0 != RetVal)
                        {
                            PSH_Globals.oCompany.GetLastError(out errDICode, out errDIMsg);
                            errCode = "1";
                            throw new Exception();
                        }
                        else
                        {
                            PSH_Globals.oCompany.GetNewObjectCode(out SDocEntry);

                            sQry = "  UPDATE    [@PS_PP040L]";
                            sQry += " SET       U_OutDocC = '" + SDocEntry + "',";
                            sQry += "           U_OutLinC = U_OutLin";
                            sQry += " FROM      [@PS_PP040L]";
                            sQry += " WHERE     U_CpCode in ('CP80101','CP80111')";
                            sQry += "           AND DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim() + "'";
                            oRecordSet.DoQuery(sQry);
                        }
                    }

                    oMat01.LoadFromDataSource();
                    if (PSH_Globals.oCompany.InTransaction == true)
                    {
                        PSH_Globals.oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                }
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                if (DI_oInventoryGenEntry != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(DI_oInventoryGenEntry);
                }
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
            SAPbouiCOM.ProgressBar ProgBar01 = null;

            try
            {
                oForm.Freeze(true);
                if (pVal.BeforeAction == true)
                {
                    if (pVal.ItemUID == "1")
                    {
                        ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (PS_PP040_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "111" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "601") // 분말 첫번째 공정 투입시 원자재 불출로직 추가(황영수 20181101)
                            {
                                if (PS_PP040_AddoInventoryGenExit() == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }

                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oOrdGbn = oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim();
                            oSequence = oMat01.Columns.Item("Sequence").Cells.Item(1).Specific.Value;
                            oDocdate = oForm.Items.Item("DocDate").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (PS_PP040_DataValidCheck() == false)
                            {
                                BubbleEvent = false;
                                return;
                            }
                            oDocEntry01 = oForm.Items.Item("DocEntry").Specific.Value.ToString().Trim();
                            oFormMode01 = oForm.Mode;
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                        }
                    }

                    if (pVal.ItemUID == "2") //취소버튼 클릭 : 저장할 자료가 있으면 메시지 표시
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            if (oMat01.VisualRowCount > 1)
                            {
                                if (PSH_Globals.SBO_Application.MessageBox("저장하지 않은 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소") == 2)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                    }

                    if (pVal.ItemUID == "Button01")
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            PS_PP040_OrderInfoLoad();
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                            PS_PP040_OrderInfoLoad();
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
                                if (oOrdGbn == "101" && oSequence == "1")
                                {
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    PS_PP040_FormItemEnabled();
                                    oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                else
                                {
                                    PS_PP040_FormItemEnabled();
                                    PS_PP040_AddMatrixRow01(0, true);
                                    PS_PP040_AddMatrixRow02(0, true);
                                }
                            }
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                        {
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.ActionSuccess == true)
                            {
                                if (oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    PS_PP040_FormItemEnabled();
                                    oForm.Items.Item("DocEntry").Specific.Value = oDocEntry01;
                                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                PS_PP040_FormItemEnabled();
                            }
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
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "OrdMgNum")
                    {
                        string ordType = oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim();

                        if (ordType == "10" || ordType == "50" || ordType == "60") //작업타입이 일반,조정
                        {
                            dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "OrdMgNum", "");
                        }
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.ColUID == "OrdMgNum")
                        {
                            string ordType = oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim();

                            if (ordType == "10" || ordType == "50" || ordType == "60" || ordType == "70") //일반,조정,설계
                            {
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim() == "선택")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("작업구분이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim() == "선택")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("ItemCode").Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("품목코드가 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("작지번호가 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("PP030HNo").Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("작지문서번호가 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdMgNum");
                                }
                            }
                            else if (ordType == "20") //지원
                            {
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim() == "선택")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("작업구분이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    oForm.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim() == "선택")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("작지번호가 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdMgNum");
                                }
                            }
                            else if (ordType == "30") //외주
                            {
                            }
                            else if (ordType == "40") //실적
                            {
                            }
                            else if (ordType == "80") //외주제작지원
                            {
                                if (oForm.Items.Item("OrdGbn").Specific.Selected.Value.ToString().Trim() == "선택")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("작업구분이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    oForm.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    BubbleEvent = false;
                                    return;
                                }
                                else if (oForm.Items.Item("BPLId").Specific.Selected.Value.ToString().Trim() == "선택")
                                {
                                    PSH_Globals.SBO_Application.StatusBar.SetText("사업장이 선택되지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                    oForm.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "OrdMgNum");
                                }
                            }
                        }
                    }
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.ColUID == "WorkCode")
                        {
                            if (string.IsNullOrEmpty(oForm.Items.Item("BaseTime").Specific.Value) || Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) == 0)
                            {
                                PSH_Globals.SBO_Application.StatusBar.SetText("기준시간을 입력하지 않았습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                oForm.Items.Item("BaseTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                BubbleEvent = false;
                                return;
                            }
                        }

                    }
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "WorkCode");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat02", "NCode");
                    dataHelpClass.ActiveUserDefineValueAlways(ref oForm, ref pVal, ref BubbleEvent, "Mat03", "FailCode");
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "MachCode"); //설비코드
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "CItemCod"); //원재료코드
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "Mat01", "SCpCode"); //지원공정(2018.05.30 송명규)
                    dataHelpClass.ActiveUserDefineValue(ref oForm, ref pVal, ref BubbleEvent, "UseMCode", ""); //작업장비
                }
                else if (pVal.Before_Action == false)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.CharPressed == 38) //방향키(↑)
                        {
                            if (pVal.Row > 1 && pVal.Row <= oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }
                        else if (pVal.CharPressed == 40) //방향키(↓)
                        {
                            if (pVal.Row > 0 && pVal.Row < oMat01.VisualRowCount)
                            {
                                oForm.Freeze(true);
                                oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                oForm.Freeze(false);
                            }
                        }

                        //작업시간 입력 시마다 합계 계산(2011.09.26 송명규 추가)
                        if (pVal.ColUID == "WorkTime" && pVal.Row != 0)
                        {
                            PS_PP040_SumWorkTime();
                        }
                    }
                    if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.ColUID == "YTime" || pVal.ColUID == "NTime")
                        {
                            if (string.IsNullOrEmpty(oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))
                            {
                                oForm.Freeze(true);
                                oMat02.FlushToDataSource();
                                oDS_PS_PP040M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "0");
                                oMat02.LoadFromDataSource();
                                oForm.Freeze(false);
                            }
                        }
                    }
                    else if (pVal.ItemUID == "BaseTime")
                    {
                        if (pVal.CharPressed == 9) //탭
                        {
                            oMat02.Columns.Item("WorkCode").Cells.Item(1).Click();
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
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value.ToString().Trim()); //기타작업
                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP040L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    //PS_PP040_AddMatrixRow(pVal.Row)
                                }
                            }
                            else
                            {
                                oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value.ToString().Trim());
                            }
                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                                oDS_PS_PP040M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value.ToString().Trim()); //기타작업
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP040M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    //PS_PP040_AddMatrixRow(pVal.Row)
                                }
                            }
                            else
                            {
                                oDS_PS_PP040M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value.ToString().Trim());
                            }
                            oMat02.LoadFromDataSource();
                            oMat02.AutoResizeColumns();
                            oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "특정컬럼")
                            {
                            }
                            else
                            {
                                oDS_PS_PP040N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Selected.Value.ToString().Trim());
                            }
                            oMat03.LoadFromDataSource();
                            oMat03.AutoResizeColumns();
                            oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
                        }
                        else
                        {
                            if (pVal.ItemUID == "OrdType")
                            {
                                oDS_PS_PP040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value.ToString().Trim());

                                string itemUID = oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value.ToString().Trim();

                                if (itemUID == "10" || itemUID == "50" || itemUID == "60" || itemUID == "70") //일반,조정,설계
                                {
                                    if (oForm.Items.Item("BPLId").Specific.Value == "1") //창원은 품목구분 선택하도록 수정 '2015.04.09
                                    {
                                        oForm.Items.Item("OrdGbn").Enabled = true;
                                    }
                                    else
                                    {
                                        oForm.Items.Item("OrdGbn").Enabled = false;
                                    }
                                    oForm.Items.Item("BPLId").Enabled = false;
                                    oForm.Items.Item("ItemCode").Enabled = false;
                                }
                                else if (itemUID == "20")
                                {
                                    oForm.Items.Item("OrdGbn").Enabled = true;
                                    oForm.Items.Item("BPLId").Enabled = true;
                                    oForm.Items.Item("ItemCode").Enabled = true;
                                }
                                else if (itemUID == "30")
                                {
                                    oForm.Items.Item("OrdGbn").Enabled = false;
                                    oForm.Items.Item("BPLId").Enabled = false;
                                    oForm.Items.Item("ItemCode").Enabled = false;
                                }
                                else if (itemUID == "40")
                                {
                                    oForm.Items.Item("OrdGbn").Enabled = false;
                                    oForm.Items.Item("BPLId").Enabled = false;
                                    oForm.Items.Item("ItemCode").Enabled = false;
                                }
                                else if (itemUID == "80") //외주제작지원
                                {
                                    oForm.Items.Item("OrdGbn").Enabled = true;
                                }

                                oForm.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                oForm.Items.Item("OrdMgNum").Specific.Value = "";
                                oForm.Items.Item("ItemCode").Specific.Value = "";
                                oForm.Items.Item("ItemName").Specific.Value = "";
                                oForm.Items.Item("OrdNum").Specific.Value = "";
                                oForm.Items.Item("OrdSub1").Specific.Value = "";
                                oForm.Items.Item("OrdSub2").Specific.Value = "";
                                oForm.Items.Item("PP030HNo").Specific.Value = "";

                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP040_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP040_AddMatrixRow02(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                            }
                            else if (pVal.ItemUID == "OrdGbn")
                            {
                                oDS_PS_PP040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value.ToString().Trim());
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP040_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP040_AddMatrixRow02(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                            }
                            else if (pVal.ItemUID == "BPLId")
                            {
                                oDS_PS_PP040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value.ToString().Trim());
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP040_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP040_AddMatrixRow02(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                            }
                            else
                            {
                                if (pVal.ItemUID != "CardType") //거래처구분이 아닐 경우만 실행(2012.02.02 송명규 추가)
                                {
                                    oDS_PS_PP040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Selected.Value.ToString().Trim());
                                }
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
                    if (pVal.ItemUID == "Opt01")
                    {
                        oForm.Settings.MatrixUID = "Mat02";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat02.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Opt02")
                    {
                        oForm.Settings.MatrixUID = "Mat03";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat03.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Opt03")
                    {
                        oForm.Settings.MatrixUID = "Mat01";
                        oForm.Settings.EnableRowFormat = true;
                        oForm.Settings.Enabled = true;
                        oMat01.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat01.SelectRow(pVal.Row, true, false);
                            oMat01Row01 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        if (pVal.Row > 0)
                        {
                            oMat02.SelectRow(pVal.Row, true, false);
                            oMat02Row02 = pVal.Row;
                        }
                    }
                    else if (pVal.ItemUID == "Mat03")
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
                    if (pVal.ItemUID == "LBtn01")
                    {
                        PS_PP030 oTempClass = new PS_PP030();
                        oTempClass.LoadForm(Convert.ToString(oForm.Items.Item("PP030HNo").Specific.Value));
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
        /// DOUBLE_CLICK 이벤트
        /// </summary>
        /// <param name="FormUID">Form UID</param>
        /// <param name="pVal">ItemEvent 객체</param>
        /// <param name="BubbleEvent">BubbleEvnet(true, false)</param>
        private void Raise_EVENT_DOUBLE_CLICK(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemUID == "Mat01")
                    {
                        if (pVal.Row > 0)
                        {
                            string ordType = oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim();

                            if (ordType == "10" || ordType == "50" || ordType == "60") //작업타입이 일반,조정인경우
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value))
                                {
                                }
                                else
                                {
                                    if (oMat03.VisualRowCount == 0)
                                    {
                                        PS_PP040_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP040_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }
                                    oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));
                                    oMat03.LoadFromDataSource();
                                    oMat03.AutoResizeColumns();
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                                    oMat03.FlushToDataSource();
                                }
                            }
                            else if (ordType == "20") //작업타입이 PSMT지원인경우
                            {
                                if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value))
                                {
                                }
                                else
                                {
                                    if (oMat03.VisualRowCount == 0)
                                    {
                                        PS_PP040_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP040_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }
                                    oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));
                                    oMat03.LoadFromDataSource();
                                    oMat03.AutoResizeColumns();
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                                    oMat03.FlushToDataSource();
                                }
                            }
                            else if (ordType == "30") //작업타입이 외주인경우
                            {
                            }
                            else if (ordType == "40") //작업타입이 실적인경우
                            {
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
                        if (pVal.ColUID == "OrdMgNum" || pVal.ColUID == "PP030HNo")
                        {
                            PS_PP030 oTempClass = new PS_PP030();
                            oTempClass.LoadForm(oMat01.Columns.Item("PP030HNo").Cells.Item(pVal.Row).Specific.Value);
                        }
                    }
                    if (pVal.ItemUID == "Mat03")
                    {
                        if (pVal.ColUID == "OrdMgNum")
                        {
                            PS_PP030 oTempClass = new PS_PP030();
                            oTempClass.LoadForm(codeHelpClass.Mid(oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value, 0, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().IndexOf("-")));
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
            string query01;
            string errCode = string.Empty;
            SAPbouiCOM.ProgressBar ProgBar01 = null;
            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                oForm.Freeze(true);
                if (pVal.Before_Action == true)
                {
                    if (pVal.ItemChanged == true)
                    {
                        if (pVal.ItemUID == "Mat01")
                        {
                            if (PS_PP040_Validate("수정01") == false)
                            {
                                oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP040L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim());
                            }
                            else
                            {
                                if (pVal.ColUID == "OrdMgNum")
                                {
                                    ProgBar01 = PSH_Globals.SBO_Application.StatusBar.CreateProgressBar("", 0, false);

                                    string ordType = oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim();

                                    if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value)) //작지번호(헤더)에 값이 없으면
                                    {
                                        if (ordType == "80")
                                        {
                                            query01 = "EXEC PS_PP040_01 '";
                                            query01 += oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "','";
                                            query01 += oForm.Items.Item("OrdType").Specific.Selected.Value + "'";
                                            RecordSet01.DoQuery(query01);
                                            if (RecordSet01.RecordCount == 0)
                                            {
                                                oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                            }
                                            else
                                            {
                                                oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
                                                oDS_PS_PP040L.SetValue("U_Sequence", pVal.Row - 1, RecordSet01.Fields.Item("Sequence").Value);
                                                oDS_PS_PP040L.SetValue("U_CpCode", pVal.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
                                                oDS_PS_PP040L.SetValue("U_CpName", pVal.Row - 1, RecordSet01.Fields.Item("CpName").Value);
                                                oDS_PS_PP040L.SetValue("U_OrdGbn", pVal.Row - 1, RecordSet01.Fields.Item("OrdGbn").Value);
                                                oDS_PS_PP040L.SetValue("U_BPLId", pVal.Row - 1, RecordSet01.Fields.Item("BPLId").Value);
                                                oDS_PS_PP040L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                                oDS_PS_PP040L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
                                                oDS_PS_PP040L.SetValue("U_OrdNum", pVal.Row - 1, RecordSet01.Fields.Item("OrdNum").Value);
                                                oDS_PS_PP040L.SetValue("U_OrdSub1", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub1").Value);
                                                oDS_PS_PP040L.SetValue("U_OrdSub2", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub2").Value);
                                                oDS_PS_PP040L.SetValue("U_PP030HNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030HNo").Value);
                                                oDS_PS_PP040L.SetValue("U_PP030MNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030MNo").Value);
                                                oDS_PS_PP040L.SetValue("U_SelWt", pVal.Row - 1, RecordSet01.Fields.Item("SelWt").Value);
                                                oDS_PS_PP040L.SetValue("U_PSum", pVal.Row - 1, RecordSet01.Fields.Item("PSum").Value);
                                                oDS_PS_PP040L.SetValue("U_BQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                oDS_PS_PP040L.SetValue("U_PQty", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_YQty", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_NQty", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_WorkTime", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_LineId", pVal.Row - 1, "");

                                                //설비코드,명 Reset
                                                oDS_PS_PP040L.SetValue("U_MachCode", pVal.Row - 1, "");
                                                oDS_PS_PP040L.SetValue("U_MachName", pVal.Row - 1, "");

                                                if (oMat03.VisualRowCount == 0) //불량코드테이블
                                                {
                                                    PS_PP040_AddMatrixRow03(0, true);
                                                }
                                                else
                                                {
                                                    PS_PP040_AddMatrixRow03(oMat03.VisualRowCount, false);
                                                }

                                                oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
                                                oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpCode").Value);
                                                oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpName").Value);
                                                oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));

                                                if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP040L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                                {
                                                    PS_PP040_AddMatrixRow01(pVal.Row, false);
                                                }
                                            }
                                        }
                                        else
                                        {
                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                        }
                                    }
                                    else //작업지시가 선택된상태
                                    {
                                        if (ordType == "10" || ordType == "50" || ordType == "60" || ordType == "70") //작업타입이 일반,조정,설계
                                        {
                                            string ordDocEntry = oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value.ToString().Split('-')[0]; //공정정보 매트릭스의 현재 행 작지문서번호

                                            if (oForm.Items.Item("PP030HNo").Specific.Value != ordDocEntry) //작지문서헤더번호가 일치하지 않으면
                                            {
                                                oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                            }
                                            else //작지문서번호가 일치하면
                                            {
                                                if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1")
                                                {
                                                    for (i = 1; i <= oMat01.RowCount; i++) //신동사업부를 제외한 사업부만 체크
                                                    {
                                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value && i != pVal.Row) //현재 입력한 값이 이미 입력되어 있는경우
                                                        {
                                                            PSH_Globals.SBO_Application.MessageBox("이미 입력한 공정입니다.");
                                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                            oMat01.LoadFromDataSourceEx();
                                                            errCode = "1";
                                                            throw new Exception();
                                                        }
                                                    }

                                                    //생산완료등록이 완료된 작번인지 체크(수량으로 비교, 2012.08.27 송명규 추가)_S
                                                    query01 = "EXEC PS_PP040_90 '";
                                                    query01 += ordDocEntry + "'";
                                                    RecordSet01.DoQuery(query01);
                                                    string WkCmDt = RecordSet01.Fields.Item("WkCmDt").Value;

                                                    if (RecordSet01.Fields.Item("Return").Value == "1") //생산완료수량이 작업지시수량만큼 모두 등록이 되었다면
                                                    {
                                                        if (PSH_Globals.SBO_Application.MessageBox("생산완료가 모두 등록된 작번(완료일자:" + WkCmDt + ")입니다." + (char)13 + "계속 진행하시겠습니까?", 1, "예", "아니오") == 1)
                                                        {
                                                            //계속 진행시에는 해당 작업지시문서번호 등록
                                                        }
                                                        else
                                                        {
                                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                            oMat01.LoadFromDataSourceEx();
                                                            errCode = "1";
                                                            throw new Exception();
                                                        }
                                                    }
                                                    //생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_E

                                                    //판매완료등록 체크_S(2015.07.14 송명규 추가)
                                                    query01 = "EXEC PS_PP040_91 '";
                                                    query01 += ordDocEntry + "','";
                                                    query01 += oDS_PS_PP040H.GetValue("U_DocDate", 0) + "'";
                                                    RecordSet01.DoQuery(query01);
                                                    string OINV_Dt = RecordSet01.Fields.Item("OINV_Dt").Value;

                                                    if (RecordSet01.Fields.Item("Return").Value == "1") //판매확정수량이 판매오더수량만큼 모두 등록이 되었다면
                                                    {
                                                        PSH_Globals.SBO_Application.MessageBox("판매완료(최종일자:" + OINV_Dt + ")된 작번입니다." + (char)13 + "등록이 불가능합니다.", 1, "확인");
                                                        oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                        oMat01.LoadFromDataSourceEx();
                                                        errCode = "1";
                                                        throw new Exception();
                                                    }
                                                    //판매완료등록 체크_E(2015.07.14 송명규 추가)
                                                }

                                                query01 = "EXEC PS_PP040_01 '";
                                                query01 += oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "','";
                                                query01 += oForm.Items.Item("OrdType").Specific.Selected.Value + "'";
                                                RecordSet01.DoQuery(query01);
                                                if (RecordSet01.RecordCount == 0)
                                                {
                                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                }
                                                else
                                                {
                                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
                                                    oDS_PS_PP040L.SetValue("U_Sequence", pVal.Row - 1, RecordSet01.Fields.Item("Sequence").Value);
                                                    oDS_PS_PP040L.SetValue("U_CpCode", pVal.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
                                                    oDS_PS_PP040L.SetValue("U_CpName", pVal.Row - 1, RecordSet01.Fields.Item("CpName").Value);
                                                    oDS_PS_PP040L.SetValue("U_OrdGbn", pVal.Row - 1, RecordSet01.Fields.Item("OrdGbn").Value);
                                                    oDS_PS_PP040L.SetValue("U_BPLId", pVal.Row - 1, RecordSet01.Fields.Item("BPLId").Value);
                                                    oDS_PS_PP040L.SetValue("U_ItemCode", pVal.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
                                                    oDS_PS_PP040L.SetValue("U_ItemName", pVal.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
                                                    oDS_PS_PP040L.SetValue("U_OrdNum", pVal.Row - 1, RecordSet01.Fields.Item("OrdNum").Value);
                                                    oDS_PS_PP040L.SetValue("U_OrdSub1", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub1").Value);
                                                    oDS_PS_PP040L.SetValue("U_OrdSub2", pVal.Row - 1, RecordSet01.Fields.Item("OrdSub2").Value);
                                                    oDS_PS_PP040L.SetValue("U_PP030HNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030HNo").Value);
                                                    oDS_PS_PP040L.SetValue("U_PP030MNo", pVal.Row - 1, RecordSet01.Fields.Item("PP030MNo").Value);
                                                    oDS_PS_PP040L.SetValue("U_SelWt", pVal.Row - 1, RecordSet01.Fields.Item("SelWt").Value);
                                                    oDS_PS_PP040L.SetValue("U_PSum", pVal.Row - 1, RecordSet01.Fields.Item("PSum").Value);
                                                    oDS_PS_PP040L.SetValue("U_BQty", pVal.Row - 1, RecordSet01.Fields.Item("BQty").Value);
                                                    oDS_PS_PP040L.SetValue("U_PQty", pVal.Row - 1, "0");
                                                    oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, "0");
                                                    oDS_PS_PP040L.SetValue("U_YQty", pVal.Row - 1, "0");
                                                    oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, "0");
                                                    oDS_PS_PP040L.SetValue("U_NQty", pVal.Row - 1, "0");
                                                    oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                                    oDS_PS_PP040L.SetValue("U_ScrapWt", pVal.Row - 1, "0");
                                                    oDS_PS_PP040L.SetValue("U_WorkTime", pVal.Row - 1, "0");
                                                    oDS_PS_PP040L.SetValue("U_LineId", pVal.Row - 1, "");

                                                    //설비코드,명 Reset
                                                    oDS_PS_PP040L.SetValue("U_MachCode", pVal.Row - 1, "");
                                                    oDS_PS_PP040L.SetValue("U_MachName", pVal.Row - 1, "");

                                                    if (oMat03.VisualRowCount == 0) //불량코드테이블
                                                    {
                                                        PS_PP040_AddMatrixRow03(0, true);
                                                    }
                                                    else
                                                    {
                                                        PS_PP040_AddMatrixRow03(oMat03.VisualRowCount, false);
                                                    }

                                                    oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
                                                    oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpCode").Value);
                                                    oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpName").Value);
                                                    oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pVal.Row));

                                                    if (ordType == "50" || ordType == "60")
                                                    {
                                                        oDS_PS_PP040H.SetValue("U_BaseTime", 0, "1");
                                                        oMat02.Columns.Item("WorkCode").Cells.Item(1).Specific.Value = "9999999";
                                                        oDS_PS_PP040M.SetValue("U_WorkName", 0, "조정");
                                                        oMat02.LoadFromDataSource();
                                                    }
                                                    else
                                                    {
                                                    }
                                                }
                                            }
                                        }
                                        else if (ordType == "20") //작업타입이 PSMT지원
                                        {
                                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1)) == 0) //올바른 공정코드인지 검사
                                            {
                                                oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                            }
                                            else
                                            {
                                                for (i = 1; i <= oMat01.RowCount; i++)
                                                {

                                                    if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pVal.Row).Specific.Value && i != pVal.Row) //현재 입력한 값이 이미 입력되어 있는경우
                                                    {
                                                        PSH_Globals.SBO_Application.StatusBar.SetText("이미 입력한 공정입니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                                                        oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, "");
                                                        errCode = "1";
                                                        throw new Exception();
                                                    }
                                                }
                                                oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP040L.SetValue("U_CpCode", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP040L.SetValue("U_CpName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                                oDS_PS_PP040L.SetValue("U_OrdGbn", pVal.Row - 1, oForm.Items.Item("OrdGbn").Specific.Selected.Value);
                                                oDS_PS_PP040L.SetValue("U_BPLId", pVal.Row - 1, oForm.Items.Item("BPLId").Specific.Selected.Value);
                                                oDS_PS_PP040L.SetValue("U_ItemCode", pVal.Row - 1, "");
                                                oDS_PS_PP040L.SetValue("U_ItemName", pVal.Row - 1, "");
                                                oDS_PS_PP040L.SetValue("U_OrdNum", pVal.Row - 1, oForm.Items.Item("OrdNum").Specific.Value);
                                                oDS_PS_PP040L.SetValue("U_OrdSub1", pVal.Row - 1, oForm.Items.Item("OrdSub1").Specific.Value);
                                                oDS_PS_PP040L.SetValue("U_OrdSub2", pVal.Row - 1, oForm.Items.Item("OrdSub2").Specific.Value);
                                                oDS_PS_PP040L.SetValue("U_PP030HNo", pVal.Row - 1, "");
                                                oDS_PS_PP040L.SetValue("U_PP030MNo", pVal.Row - 1, "");
                                                oDS_PS_PP040L.SetValue("U_PSum", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_PQty", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_YQty", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_NQty", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                                oDS_PS_PP040L.SetValue("U_ScrapWt", pVal.Row - 1, "0");

                                                if (oMat03.VisualRowCount == 0) //불량코드테이블
                                                {
                                                    PS_PP040_AddMatrixRow03(0, true);
                                                }
                                                else
                                                {
                                                    if (oDS_PS_PP040L.GetValue("U_OrdMgNum", pVal.Row - 1) == oDS_PS_PP040N.GetValue("U_OrdMgNum", oMat03.VisualRowCount - 1))
                                                    {
                                                    }
                                                    else
                                                    {
                                                        PS_PP040_AddMatrixRow03(oMat03.VisualRowCount, false);
                                                    }
                                                }
                                                oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                            }
                                        }
                                        else if (ordType == "30") //작업타입이 외주
                                        {
                                        }
                                        else if (ordType == "40") //작업타입이 실적
                                        {
                                        }

                                        if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP040L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                        {
                                            PS_PP040_AddMatrixRow01(pVal.Row, false);
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "PQty") //공정정보(Mat01).생산수량(PQty)
                                {
                                    string query = "SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'";
                                    string returnValue = dataHelpClass.GetValue(query, 0, 1);
                                    double weight = Convert.ToDouble(returnValue == "" ? "0" : returnValue) / 1000;

                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0)
                                    {
                                        if (oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "50" || oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "60")
                                        {
                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP040L.SetValue("U_YQty", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            if (weight == 0)
                                            {
                                                oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            }
                                            else
                                            {
                                                oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                            oDS_PS_PP040L.SetValue("U_NQty", pVal.Row - 1, "0");
                                            oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                        }
                                        else
                                        {
                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP040L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                    }
                                    else
                                    {
                                        oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP040L.SetValue("U_YQty", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);

                                        if (weight == 0)
                                        {
                                            oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        }
                                        else
                                        {
                                            oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        oDS_PS_PP040L.SetValue("U_NQty", pVal.Row - 1, "0");
                                        oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, "0");
                                    }
                                }
                                else if (pVal.ColUID == "NQty") //공정정보(Mat01).불량수량(NQty)
                                {
                                    string query = "SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pVal.Row).Specific.Value + "'";
                                    string returnValue = dataHelpClass.GetValue(query, 0, 1);
                                    double weight = Convert.ToDouble(returnValue == "" ? "0" : returnValue) / 1000;

                                    if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) <= 0) //불량 수량이 0보다 작거나 같으면
                                    {
                                        if (oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "50" || oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "60")
                                        {
                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP040L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            if (weight == 0)
                                            {
                                                oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                            else
                                            {
                                                oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(weight * (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
                                            }
                                        }
                                        else
                                        {
                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP040L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                    }
                                    else if (Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) > Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value)) //불량수량이 생산수량보다 크면
                                    {
                                        if (oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "50" || oDS_PS_PP040H.GetValue("U_OrdType", 0).ToString().Trim() == "60")
                                        {
                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP040L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            if (weight == 0)
                                            {
                                                oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                                oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            }
                                            else
                                            {
                                                oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                                oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(weight * (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
                                            }
                                        }
                                        else
                                        {
                                            oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oDS_PS_PP040L.GetValue("U_" + pVal.ColUID, pVal.Row - 1));
                                        }
                                    }
                                    else
                                    {
                                        oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                        oDS_PS_PP040L.SetValue("U_YQty", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        if (weight == 0)
                                        {
                                            oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                            oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                        }
                                        else
                                        {
                                            oDS_PS_PP040L.SetValue("U_NWeight", pVal.Row - 1, Convert.ToString(weight * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                            oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(weight * (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value))));
                                        }
                                    }
                                }
                                else if (pVal.ColUID == "WorkTime") //작업시간(공수)을 입력할 때
                                {
                                    if (oForm.Items.Item("BPLId").Specific.Selected.Value != "1")
                                    {
                                        oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    }
                                }
                                else if (pVal.ColUID == "BdwQty") //기존도면매수
                                {
                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040L.SetValue("U_AdwQTy", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100));
                                    oDS_PS_PP040L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("DwRate").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                }
                                else if (pVal.ColUID == "DwRate") //도면 적용률
                                {
                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040L.SetValue("U_AdwQTy", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100));
                                    oDS_PS_PP040L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString((Convert.ToDouble(oMat01.Columns.Item("BdwQty").Cells.Item(pVal.Row).Specific.Value) * Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value) / 100) + Convert.ToDouble(oMat01.Columns.Item("NdwQTy").Cells.Item(pVal.Row).Specific.Value)));
                                }
                                else if (pVal.ColUID == "NdwQTy") //신규도면매수
                                {
                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040L.SetValue("U_PQTy", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_PWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_YQTy", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                    oDS_PS_PP040L.SetValue("U_YWeight", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oMat01.Columns.Item("AdwQty").Cells.Item(pVal.Row).Specific.Value) + Convert.ToDouble(oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value)));
                                }
                                else if (pVal.ColUID == "MachCode")
                                {
                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040L.SetValue("U_MachName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else if (pVal.ColUID == "CItemCod")
                                {
                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040L.SetValue("U_CItemNam", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_ItemNam2 FROM [@PS_PP005H] WHERE U_ItemCod1 = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pVal.Row).Specific.Value + "' and U_ItemCod2 = '" + oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else if (pVal.ColUID == "SCpCode") //지원공정코드
                                {
                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                    oDS_PS_PP040L.SetValue("U_SCpName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item("SCpCode").Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                }
                                else
                                {
                                    oDS_PS_PP040L.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                }
                            }
                            oMat01.LoadFromDataSource();
                            oMat01.AutoResizeColumns();
                            oMat01.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            oMat03.LoadFromDataSource();
                            oMat03.AutoResizeColumns();
                        }
                        else if (pVal.ItemUID == "Mat02")
                        {
                            if (pVal.ColUID == "WorkCode")
                            {
                                oDS_PS_PP040M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP040M.SetValue("U_WorkName", pVal.Row - 1, dataHelpClass.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                                if (oMat02.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP040M.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                                {
                                    PS_PP040_AddMatrixRow02(pVal.Row, false);
                                }
                            }
                            else if (pVal.ColUID == "NStart") //필드 없음(사용안되는듯, C#마이그레이션 구현, 2021.03.15 송명규)
                            {
                                oDS_PS_PP040M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP040M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP040M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    oDS_PS_PP040M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                }
                                else
                                {
                                    double time = 0;
                                    if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        time = Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        time = (2400 - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    double hour = time / 100;
                                    double minute = time % 100;
                                    time = hour;
                                    if (minute > 0)
                                    {
                                        time += 0.5;
                                    }
                                    oDS_PS_PP040M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(time));
                                    oDS_PS_PP040M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                    oDS_PS_PP040M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                }
                            }
                            else if (pVal.ColUID == "NEnd") //필드 없음(사용안되는듯, C#마이그레이션 구현, 2021.03.15 송명규)
                            {
                                oDS_PS_PP040M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) == 0 || Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) == 0)
                                {
                                    oDS_PS_PP040M.SetValue("U_NTime", pVal.Row - 1, "0");
                                    oDS_PS_PP040M.SetValue("U_YTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                    oDS_PS_PP040M.SetValue("U_TTime", pVal.Row - 1, oForm.Items.Item("BaseTime").Specific.Value);
                                }
                                else
                                {
                                    double time = 0;
                                    if (Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value) <= Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value))
                                    {
                                        time = Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value) - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    else
                                    {
                                        time = (2400 - Convert.ToDouble(oMat02.Columns.Item("NStart").Cells.Item(pVal.Row).Specific.Value)) + Convert.ToDouble(oMat02.Columns.Item("NEnd").Cells.Item(pVal.Row).Specific.Value);
                                    }
                                    double hour = time / 100;
                                    double minute = time % 100;
                                    time = hour;
                                    if (minute > 0)
                                    {
                                        time += 0.5;
                                    }
                                    oDS_PS_PP040M.SetValue("U_NTime", pVal.Row - 1, Convert.ToString(time));
                                    oDS_PS_PP040M.SetValue("U_YTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                    oDS_PS_PP040M.SetValue("U_TTime", pVal.Row - 1, Convert.ToString(Convert.ToDouble(oForm.Items.Item("BaseTime").Specific.Value) - time));
                                }
                            }
                            else if (pVal.ColUID == "YTime")
                            {
                                oDS_PS_PP040M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP040M.SetValue("U_TTime", pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                            else
                            {
                                oDS_PS_PP040M.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                            oMat02.LoadFromDataSource();
                            oMat02.AutoResizeColumns();
                            oMat02.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else if (pVal.ItemUID == "Mat03")
                        {
                            if (pVal.ColUID == "FailCode")
                            {
                                oDS_PS_PP040N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                                oDS_PS_PP040N.SetValue("U_FailName", pVal.Row - 1, dataHelpClass.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value + "'", 0, 1));
                            }
                            else
                            {
                                oDS_PS_PP040N.SetValue("U_" + pVal.ColUID, pVal.Row - 1, oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific.Value);
                            }
                            oMat03.LoadFromDataSource();
                            oMat03.AutoResizeColumns();
                            oMat03.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }
                        else
                        {
                            if (pVal.ItemUID == "DocEntry")
                            {
                                oDS_PS_PP040H.SetValue(pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "BaseTime")
                            {
                                oDS_PS_PP040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
                            else if (pVal.ItemUID == "OrdMgNum")
                            {
                                oDS_PS_PP040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    PS_PP040_OrderInfoLoad();
                                }
                            }
                            else if (pVal.ItemUID == "ItemCode")
                            {
                                oDS_PS_PP040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                                oMat01.Clear();
                                oMat01.FlushToDataSource();
                                oMat01.LoadFromDataSource();
                                PS_PP040_AddMatrixRow01(0, true);
                                oMat02.Clear();
                                oMat02.FlushToDataSource();
                                oMat02.LoadFromDataSource();
                                PS_PP040_AddMatrixRow02(0, true);
                                oMat03.Clear();
                                oMat03.FlushToDataSource();
                                oMat03.LoadFromDataSource();
                            }
                            else if (pVal.ItemUID == "UseMCode")
                            {
                                query01 = "EXEC PS_PP040_98 '" + oForm.Items.Item("UseMCode").Specific.Value;
                                RecordSet01.DoQuery(query01);
                                oForm.Items.Item("UseMName").Specific.Value = RecordSet01.Fields.Item(0).Value.ToString().Trim();
                            }
                            else
                            {
                                oDS_PS_PP040H.SetValue("U_" + pVal.ItemUID, 0, oForm.Items.Item(pVal.ItemUID).Specific.Value);
                            }
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
                if (errCode == "1")
                {
                    if (oMat01.RowCount == pVal.Row && !string.IsNullOrEmpty(oDS_PS_PP040L.GetValue("U_" + pVal.ColUID, pVal.Row - 1).ToString().Trim()))
                    {
                        PS_PP040_AddMatrixRow01(pVal.Row, false);
                    }
                }
                else
                {
                    PSH_Globals.SBO_Application.MessageBox(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message);                    
                }

                BubbleEvent = false;
            }
            finally
            {
                oForm.Freeze(false);

                if (ProgBar01 != null)
                {
                    ProgBar01.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ProgBar01);
                }

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
                PS_PP040_SumWorkTime();

                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    PS_PP040_FormItemEnabled();
                    if (pVal.ItemUID == "Mat01")
                    {
                        PS_PP040_AddMatrixRow01(oMat01.VisualRowCount, false);
                        oMat01.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Mat02")
                    {
                        PS_PP040_AddMatrixRow02(oMat02.VisualRowCount, false);
                        oMat02.AutoResizeColumns();
                    }
                    else if (pVal.ItemUID == "Mat03")
                    {
                        oMat03.AutoResizeColumns();
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat02);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMat03);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP040H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP040L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP040M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP040N);
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
                    PS_PP040_FormResize();
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
                    if (pVal.ItemUID == "ItemCode")
                    {
                        dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PS_PP040H", "U_ItemCode,U_ItemName", pVal.ItemUID, (short)pVal.Row, "", "", "");
                        oMat01.Clear();
                        oMat01.FlushToDataSource();
                        oMat01.LoadFromDataSource();
                        PS_PP040_AddMatrixRow01(0, true);
                        oMat02.Clear();
                        oMat02.FlushToDataSource();
                        oMat02.LoadFromDataSource();
                        PS_PP040_AddMatrixRow02(0, true);
                        oMat03.Clear();
                        oMat03.FlushToDataSource();
                        oMat03.LoadFromDataSource();
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
            int j;
            bool exist;

            try
            {
                if (oLastColRow01 > 0)
                {
                    if (pVal.BeforeAction == true)
                    {
                        if (oForm.Items.Item("OrdGbn").Specific.Value == "111"
                            && (oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.Value == "CP80111"
                               || oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.Value == "CP80101")
                           ) //분말 첫번째 공정
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("첫공정은 행삭제 할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }
                        else if (oForm.Items.Item("OrdGbn").Specific.Value == "601"
                                 && (oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.Value == "CP80111"
                                    || oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.Value) == "CP80101"
                                ) //분말 첫번째 공정
                        {
                            PSH_Globals.SBO_Application.StatusBar.SetText("첫공정은 행삭제 할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }

                        if (oLastItemUID01 == "Mat01")
                        {
                            if (PS_PP040_Validate("행삭제01") == false)
                            {
                                BubbleEvent = false;
                                return;
                            }

                            for (i = 1; i <= oMat03.RowCount; i++)
                            {
                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value
                                 && oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value)
                                {
                                    oDS_PS_PP040N.RemoveRecord(i - 1);
                                    oMat03.DeleteRow(i);
                                    oMat03.FlushToDataSource();
                                }
                            }
                        }
                    }
                    else if (pVal.BeforeAction == false)
                    {
                        if (oLastItemUID01 == "Mat01")
                        {
                            for (i = 1; i <= oMat01.VisualRowCount; i++)
                            {
                                oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                            }

                            for (i = 1; i <= oMat03.VisualRowCount; i++)
                            {
                                if (Convert.ToInt32(oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value) != 1)
                                {
                                    oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value = Convert.ToInt32(oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value) - 1;
                                }
                            }
                            oMat01.FlushToDataSource();
                            oDS_PS_PP040L.RemoveRecord(oDS_PS_PP040L.Size - 1);
                            oMat01.LoadFromDataSource();
                            if (oMat01.RowCount == 0)
                            {
                                PS_PP040_AddMatrixRow01(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP040L.GetValue("U_OrdMgNum", oMat01.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP040_AddMatrixRow01(oMat01.RowCount, false);
                                }
                            }
                            PS_PP040_SumWorkTime();
                        }
                        else if (oLastItemUID01 == "Mat02")
                        {
                            for (i = 1; i <= oMat02.VisualRowCount; i++)
                            {
                                oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                            }
                            oMat02.FlushToDataSource();
                            oDS_PS_PP040M.RemoveRecord(oDS_PS_PP040M.Size - 1);
                            oMat02.LoadFromDataSource();
                            if (oMat02.RowCount == 0)
                            {
                                PS_PP040_AddMatrixRow02(0, false);
                            }
                            else
                            {
                                if (!string.IsNullOrEmpty(oDS_PS_PP040M.GetValue("U_WorkCode", oMat02.RowCount - 1).ToString().Trim()))
                                {
                                    PS_PP040_AddMatrixRow02(oMat02.RowCount, false);
                                }
                            }
                        }
                        else if (oLastItemUID01 == "Mat03")
                        {
                            for (i = 1; i <= oMat03.VisualRowCount; i++)
                            {
                                oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
                            }
                            oMat03.FlushToDataSource();

                            if (oDS_PS_PP040N.Size == 1) //사이즈가 0일때는 행을 빼주면 oMat03.VisualRowCount 가 0 으로 변경되어서 문제가 생김
                            {
                            }
                            else
                            {
                                oDS_PS_PP040N.RemoveRecord(oDS_PS_PP040N.Size - 1);
                            }
                            oMat03.LoadFromDataSource();

                            for (i = 1; i <= oMat01.RowCount - 1; i++) //공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
                            {
                                exist = false;
                                for (j = 1; j <= oMat03.RowCount; j++)
                                {
                                    if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value
                                     && oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.Value)
                                    {
                                        exist = true;
                                    }
                                }

                                if (exist == false) //불량코드테이블에 값이 존재하지 않으면
                                {
                                    if (oMat03.VisualRowCount == 0)
                                    {
                                        PS_PP040_AddMatrixRow03(0, true);
                                    }
                                    else
                                    {
                                        PS_PP040_AddMatrixRow03(oMat03.VisualRowCount, false);
                                    }
                                    oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(i).Specific.Value);
                                    oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(i));
                                    oMat03.LoadFromDataSource();
                                    oMat03.AutoResizeColumns();
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
                                    oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
                                    oMat03.FlushToDataSource();
                                }
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
        /// 네비게이션 메소드(Raise_FormMenuEvent 에서 사용)
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="pVal"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_RECORD_MOVE(string FormUID, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            string query01;
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
                        query01 = "  SELECT		ISNULL";
                        query01 += "            (";
                        query01 += "                MIN(DocEntry),";
                        query01 += "                (SELECT MIN(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '10' AND U_OrdGbn IN ('101','103','105','106','108','109','110','111','601'))";
                        query01 += "            )";
                        query01 += " FROM       [@PS_PP040H]";
                        query01 += " WHERE      U_DocType = '10'";
                        query01 += "            AND U_OrdGbn IN ('101','103','105','106','108','109','110','111','601')";
                        query01 += "            AND DocEntry > " + docEntry;

                        oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(query01, 0, 1);
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
                        query01 = "  SELECT		ISNULL";
                        query01 += "            (";
                        query01 += "                MAX(DocEntry),";
                        query01 += "                (SELECT MAX(DocEntry) FROM [@PS_PP040H] WHERE U_DocType = '10' AND U_OrdGbn IN ('101','103','105','106','108','109','110','111','601'))";
                        query01 += "            )";
                        query01 += " FROM       [@PS_PP040H]";
                        query01 += " WHERE      U_DocType = '10'";
                        query01 += "            AND U_OrdGbn IN ('101','103','105','106','108','109','110','111','601')";
                        query01 += "            AND DocEntry < " + docEntry;

                        oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(query01, 0, 1);
                        oForm.Items.Item("1").Enabled = true;
                        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        oForm.Items.Item("DocEntry").Enabled = false;
                    }
                }
                else if (pVal.MenuUID == "1290") //최초
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    query01 = "  SELECT     MIN(DocEntry)";
                    query01 += " FROM       [@PS_PP040H]";
                    query01 += " WHERE      U_DocType = '10'";
                    query01 += "            AND U_OrdGbn IN ('101','103','105','106','108','109','110','111','601')";

                    oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(query01, 0, 1);
                    oForm.Items.Item("1").Enabled = true;
                    oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    oForm.Items.Item("DocEntry").Enabled = false;
                }
                else if (pVal.MenuUID == "1291") //최종
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                    oForm.Items.Item("DocEntry").Enabled = true;
                    query01 = "  SELECT     MAX(DocEntry)";
                    query01 += " FROM       [@PS_PP040H]";
                    query01 += " WHERE      U_DocType = '10'";
                    query01 += "            AND U_OrdGbn IN ('101','103','105','106','108','109','110','111','601')";

                    oForm.Items.Item("DocEntry").Specific.Value = dataHelpClass.GetValue(query01, 0, 1);
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
                            if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 취소할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
                                BubbleEvent = false;
                                return;
                            }
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (PS_PP040_Validate("취소") == false)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (PSH_Globals.SBO_Application.MessageBox("정말로 취소하시겠습니까?", 1, "예", "아니오") != 1)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "111" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "601")
                                {
                                    if (PS_PP040_AddoInventoryGenEntry() == false)
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
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
                            if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value.ToString().Trim().Substring(0, 6)) == false)
                            {
                                PSH_Globals.SBO_Application.MessageBox("마감상태가 잠금입니다. 해당 일자로 닫기할 수 없습니다. 작성일자를 확인하고, 회계부서로 문의하세요.");
                                BubbleEvent = false;
                                return;
                            }
                            break;
                        case "1293": //행삭제
                            Raise_EVENT_ROW_DELETE(FormUID, ref pVal, ref BubbleEvent);
                            break;
                        case "1281": //찾기
                            break;
                        case "1282": //추가
                            break;
                        case "1288": //레코드이동버튼(다음)
                        case "1289": //레코드이동버튼(이전)
                        case "1290": //레코드이동버튼(최초)
                        case "1291": //레코드이동버튼(최종)
                            Raise_EVENT_RECORD_MOVE(FormUID, ref pVal, ref BubbleEvent);
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
                            PS_PP040_FormItemEnabled();
                            oForm.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            break;
                        case "1282": //추가
                            PS_PP040_FormItemEnabled();
                            PS_PP040_AddMatrixRow01(0, true);
                            PS_PP040_AddMatrixRow02(0, true);
                            break;
                        case "1288": //레코드이동버튼(다음)
                        case "1289": //레코드이동버튼(이전)
                        case "1290": //레코드이동버튼(최초)
                        case "1291": //레코드이동버튼(최종)
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
                        Raise_EVENT_FORM_DATA_LOAD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD: //34
                        //Raise_EVENT_FORM_DATA_ADDD(FormUID, ref BusinessObjectInfo, ref BubbleEvent);
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
        /// FORM_DATA_LOAD 이벤트
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        private void Raise_EVENT_FORM_DATA_LOAD(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
                {
                }
                else if (BusinessObjectInfo.BeforeAction == false)
                {
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        if (PS_PP040_FindValidateDocument("@PS_PP040H") == false)
                        {
                            if (PSH_Globals.SBO_Application.Menus.Item("1281").Enabled == true) //찾기메뉴 활성화일때 수행
                            {
                                PSH_Globals.SBO_Application.ActivateMenuItem("1281");
                            }
                            else
                            {
                                PSH_Globals.SBO_Application.SetStatusBarMessage("관리자에게 문의바랍니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            }
                            BubbleEvent = false;
                            return;
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
