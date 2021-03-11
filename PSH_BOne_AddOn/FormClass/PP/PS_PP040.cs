using System;
using SAPbouiCOM;
using PSH_BOne_AddOn.Data;

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

        //사용자구조체
        private struct ItemInformations
        {
            public string ItemCode;
            public string BatchNum;
            public int Quantity;
            public int OPORNo;
            public int POR1No;
            public bool Check;
            public int OPDNNo;
            public int PDN1No;
        }
        private ItemInformations[] ItemInformation;
        private int ItemInformationCount;

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
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("OrdType").Specific, "PS_PP040", "OrdType", false);

                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "DocType", "", "10", "작지기준");
                dataHelpClass.Combo_ValidValues_Insert("PS_PP040", "DocType", "", "20", "공정기준");
                dataHelpClass.Combo_ValidValues_SetValueItem(oForm.Items.Item("DocType").Specific, "PS_PP040", "DocType", false);

                oForm.Items.Item("OrdGbn").Specific.ValidValues.Add("선택", "선택");
                dataHelpClass.Set_ComboList(oForm.Items.Item("OrdGbn").Specific, "SELECT Code, Name FROM [@PSH_ITMBSORT] WHERE U_PudYN = 'Y' AND CODE NOT IN('104','107') order by Code", "", false, false);
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
            catch(Exception ex)
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
            catch(Exception ex)
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
            catch(Exception ex)
            {
                PSH_Globals.SBO_Application.StatusBar.SetText(System.Reflection.MethodBase.GetCurrentMethod().Name + "_Error : " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// 각모드에따른 아이템설정
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

                    if ((oDocType01 == "작업일보등록(작지)"))
                    {
                        oDS_PS_PP040H.SetValue("U_DocType", 0, "10");
                    }
                    else if ((oDocType01 == "작업일보등록(공정)"))
                    {
                        oForm.Items.Item("DocType").Specific.Select("20", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    }
                    if (string.IsNullOrEmpty(oDocdate)) //oDocdate 변수 데이터할당 타이밍 재확인 필요
                    {
                        oForm.Items.Item("DocDate").Specific.Value = DateTime.Now.ToString("yyyyMMdd"); //Microsoft.VisualBasic.Compatibility.VB6.Support.Format(System.Date.FromOADate(DateAndTime.Now.ToOADate() - 1), "YYYYMMDD");
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
            catch(Exception ex)
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
            }
            catch(Exception ex)
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
            int i = 0;
            int j;
            int failQty = 0;
            string sQry;
            string errCode = string.Empty;
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    PS_PP040_FormClear();
                }

                if (dataHelpClass.GetValue("select Count(*) from OFPR Where '" + oForm.Items.Item("DocDate").Specific.Value + "' between F_RefDate and T_RefDate And PeriodStat = 'Y'", 0, 1) > 0)
                {
                    errCode = "1";
                    throw new Exception();
                }

                if (oForm.Items.Item("OrdType").Specific.Selected.Value != "10" && oForm.Items.Item("OrdType").Specific.Selected.Value != "20" && oForm.Items.Item("OrdType").Specific.Selected.Value != "50" && oForm.Items.Item("OrdType").Specific.Selected.Value != "60" && oForm.Items.Item("OrdType").Specific.Selected.Value != "70")
                {
                    errCode = "2";
                    throw new Exception();
                }

                if (string.IsNullOrEmpty(oForm.Items.Item("OrdNum").Specific.Value))
                {
                    errCode = "3";
                    throw new Exception();
                }

                if (oMat01.VisualRowCount == 1)
                {
                    errCode = "4";
                    throw new Exception();
                }

                if (oMat02.VisualRowCount == 1)
                {
                    errCode = "5";
                    throw new Exception();
                }

                //마감상태 체크_S(2017.11.23 송명규 추가)
                if (dataHelpClass.Check_Finish_Status(oForm.Items.Item("BPLId").Specific.Value.ToString().Trim(), oForm.Items.Item("DocDate").Specific.Value, oForm.TypeEx) == false)
                {
                    errCode = "6";
                    throw new Exception();
                }
                //마감상태 체크_E(2017.11.23 송명규 추가)

                if (oMat03.VisualRowCount == 0)
                {
                    errCode = "7";
                    throw new Exception();
                }

                for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                {
                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value))
                    {
                        errCode = "8";
                        throw new Exception();
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50" && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("PQty").Cells.Item(i).Specific.Value) <= 0)
                        {
                            errCode = "9";
                            throw new Exception();
                        }
                    }
                    else if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50" && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60" && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim()() != "70")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("WorkTime").Cells.Item(i).Specific.Value) <= 0)
                        {
                            errCode = "10";
                            throw new Exception();
                        }
                    }
                    else if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "105" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "106") //작업완료여부(2012.02.02. 송명규 추가)(기계공구, 몰드일 경우만 작업완료여부 필수 체크)
                    {
                        if (oMat01.Columns.Item("CompltYN").Cells.Item(i).Specific.Value.ToString().Trim() == "%")
                        {
                            errCode = "11";
                            throw new Exception();
                        }
                    }

                    //불량수량 검사
                    for (j = 1; j <= oMat03.VisualRowCount; j++)
                    {
                        //불량코드 입력 여부 check
                        if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) != 0 && string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value.ToString().Trim()))
                        {
                            errCode = "12";
                            throw new Exception();
                        }
                        else if (Convert.ToDouble(oMat03.Columns.Item("FailQty").Cells.Item(j).Specific.Value) == 0 && !string.IsNullOrEmpty(oMat03.Columns.Item("FailCode").Cells.Item(j).Specific.Value.ToString().Trim()))
                        {
                            errCode = "13";
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
                                    errCode = "14";
                                    throw new Exception();
                                }
                            }
                        }
                    }

                    if (oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "50" && oForm.Items.Item("OrdType").Specific.Value.ToString().Trim() != "60")
                    {
                        if (Convert.ToDouble(oMat01.Columns.Item("NQty").Cells.Item(i).Specific.Value) != failQty)
                        {
                            errCode = "15";
                            throw new Exception();
                        }
                    }

                    if (oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "601" || oForm.Items.Item("OrdGbn").Specific.Value.ToString().Trim() == "111")
                    {
                        if (oMat01.Columns.Item("Sequence").Cells.Item(i).Specific.Value == 1 && string.IsNullOrEmpty(oMat01.Columns.Item("CItemCod").Cells.Item(i).Specific.Value.ToString().Trim()))
                        {
                            errCode = "16";
                            throw new Exception();
                        }
                    }
                }

                //비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_S
                for (i = 1; i <= oMat02.VisualRowCount - 1; i++)
                {
                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value))
                        {
                            errCode = "17";
                            throw new Exception();
                        }
                    }

                    if (!string.IsNullOrEmpty(oMat02.Columns.Item("NTime").Cells.Item(i).Specific.Value))
                    {
                        if (string.IsNullOrEmpty(oMat02.Columns.Item("NCode").Cells.Item(i).Specific.Value))
                        {
                            errCode = "18";
                            throw new Exception();
                        }
                    }
                }
                //비가동코드와 비가동시간 체크(2012.06.14 송명규 추가)_E

                if (PS_PP040_Validate("검사01") == false)
                {
                    errCode = "19";
                    throw new Exception();
                }

                oDS_PS_PP040L.RemoveRecord(oDS_PS_PP040L.Size - 1);
                oMat01.LoadFromDataSource();
                oDS_PS_PP040M.RemoveRecord(oDS_PS_PP040M.Size - 1);
                oMat02.LoadFromDataSource();
                
                returnValue = true;
            }
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당일자는 전기기간이 잠겼습니다. 일자를 확인바랍니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업타입이 일반, PSMT지원, 조정, 설계가 아닙니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "3")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작지번호는 필수입니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oForm.Items.Item("OrdNum").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("공정정보 라인이 존재하지 않습니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업자정보 라인이 존재하지 않습니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "6")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("마감상태가 잠금입니다. 해당 일자로 등록할 수 없습니다. 작업일보일자를 확인하고, 회계부서로 문의하세요." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "7")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("불량정보 라인이 존재하지 않습니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "8")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작지문서번호는 필수입니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "9")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("생산수량은 필수입니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("PQty").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "10")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("실동시간은 필수입니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("WorkTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "11")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업구분이 기계공구, 몰드일경우는 작업완료여부가 필수입니다. 확인하십시오." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat01.Columns.Item("CompltYN").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "12")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("불량수량이 입력되었을 때는 불량코드는 필수입니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "13")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("불량코드를 확인하세요." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat03.Columns.Item("FailCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "14")
                {
                    PSH_Globals.SBO_Application.MessageBox("기계사업부 품질팀만 등록 및 수정이 가능합니다.");
                }
                else if (errCode == "15")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("공정리스트의 불량수량과 불량정보의 불량수량이 일치하지 않습니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "16")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("공정 사용 원재료코드가 없습니다. 사용 원재료를 선택해 주세요." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "17")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("비가동코드가 입력되었을 때는 비가동시간은 필수입니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat02.Columns.Item("NTime").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "18")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("비가동시간이 입력되었을 때는 비가동코드는 필수입니다." + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    oMat02.Columns.Item("NCode").Cells.Item(i).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else if (errCode == "19") //PS_PP040_Validate("검사01") == false 인 경우 : PS_PP040_Validate 메소드에서 에러 메시지 출력
                {
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
            int PrevDBCpQty;
            int PrevMATRIXCpQty;
            int CurrentDBCpQty;
            int CurrentMATRIXCpQty;
            string PrevCpInfo;
            string CurrentCpInfo;
            string OrdMgNum;
            bool Exist;
            string LineNum;
            string DocEntry;
            string errCode = string.Empty;

            SAPbobsCOM.Recordset RecordSet01 = PSH_Globals.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            PSH_DataHelpClass dataHelpClass = new PSH_DataHelpClass();

            try
            {
                if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (dataHelpClass.GetValue("SELECT Canceled FROM [@PS_PP040H] WHERE DocEntry = '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1) == "Y")
                    {
                        errCode = "1";
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
                    errCode = "2";
                    throw new Exception();
                }
                
                if (ValidateType == "검사01")
                {
                    if (oForm.Items.Item("OrdType").Specific.Selected.Value.ToString().Trim() == "10") //작업타입이 일반인경우
                    {
                        for (i = 1; i <= oMat01.VisualRowCount - 1; i++)
                        {
                            if (dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE PS_PP030H.Canceled = 'N' AND CONVERT(NVARCHAR,PS_PP030H.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" + oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value + "'", 0, 1) <= 0)
                            {
                                errCode = "4";
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
                                        if (Convert.ToInt64(RecordSet01.Fields.Item(0).Value) == Convert.ToInt64(oForm.Items.Item("DocEntry").Specific.Value) && Convert.ToInt64(RecordSet01.Fields.Item(1).Value) == Convert.ToInt64(oMat01.Columns.Item("LineId").Cells.Item(j).Specific.Value))
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
                                                errCode = "5";
                                                throw new Exception();
                                            }
                                            //휘팅실적
                                            if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                            {
                                                errCode = "4";
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
                                            PrevDBCpQty = Convert.ToInt32(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP040H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                            //재공이동 수량
                                            PrevDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                            PrevDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                            PrevMATRIXCpQty = 0;
                                            for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                            {
                                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == PrevCpInfo)
                                                {
                                                    PrevMATRIXCpQty += Convert.ToInt32(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                                }
                                            }

                                            CurrentDBCpQty = Convert.ToInt32(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP040L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                            CurrentDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                            CurrentDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                            CurrentMATRIXCpQty = 0;
                                            for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                            {
                                                if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == CurrentCpInfo)
                                                {
                                                    CurrentMATRIXCpQty += Convert.ToInt32(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                                }
                                            }
                                            if ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))
                                            {
                                                errCode = "7";
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
                                                else //값이 변경된 행의경우
                                                {
                                                    errCode = "4";
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
                                    PrevDBCpQty = Convert.ToInt32(dataHelpClass.GetValue("EXEC PS_PP040_07 '" + PrevCpInfo + "', '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1));
                                    //재공 이동수량 반영
                                    PrevDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    PrevDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                    PrevMATRIXCpQty = 0;
                                    for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == PrevCpInfo)
                                        {
                                            PrevMATRIXCpQty += Convert.ToInt32(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                        }
                                    }
                                    CurrentDBCpQty = Convert.ToInt32(dataHelpClass.GetValue("EXEC PS_PP040_07 '" + CurrentCpInfo + "', '" + oForm.Items.Item("DocEntry").Specific.Value + "'", 0, 1));
                                    CurrentDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    CurrentDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                    CurrentMATRIXCpQty = 0;
                                    for (j = 1; j <= oMat01.VisualRowCount - 1; j++)
                                    {
                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value == CurrentCpInfo)
                                        {
                                            CurrentMATRIXCpQty += Convert.ToInt32(oMat01.Columns.Item("PQty").Cells.Item(j).Specific.Value);
                                        }
                                    }
                                    
                                    if ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))
                                    {
                                        oMat01.SelectRow(i, true, false);
                                        errCode = "7";
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
                                        errCode = "4";
                                        throw new Exception();
                                    }

                                    //휘팅실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errCode = "4";
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

                                if (Convert.ToInt32(stockQty) > 0)
                                {
                                    errCode = "12";
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
                        if (string.IsNullOrEmpty(oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value)) //새로추가된 행인경우
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

                                    if (oMat01.Columns.Item("PQty").Cells.Item(oMat01Row01).Specific.Value.ToString().Trim() != dataHelpClass.GetValue("select U_pqty from [@PS_PP040L] where DocEntry ='" + DocEntry + "' and u_linenum ='" + LineNum + "'", 0, 1))
                                    {
                                        errCode = "13";
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
                                        errCode = "4";
                                        throw new Exception();
                                    }

                                    //휘팅실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + oForm.Items.Item("DocEntry").Specific.Value + "' AND PS_PP080L.U_PP030MNo = '" + oMat01.Columns.Item("LineId").Cells.Item(oMat01Row01).Specific.Value + "'", 0, 1)) > 0)
                                    {
                                        errCode = "4";
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
                        errCode = "1";
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
                                        errCode = "17";
                                        throw new Exception();
                                    }

                                    //휘팅실적
                                    if (Convert.ToInt32(dataHelpClass.GetValue("SELECT COUNT(*) FROM [@PS_PP080H] PS_PP080H LEFT JOIN [@PS_PP080L] PS_PP080L ON PS_PP080H.DocEntry = PS_PP080L.DocEntry WHERE PS_PP080H.Canceled = 'N' AND PS_PP080L.U_PP030HNo = '" + RecordSet01.Fields.Item("PP030HNo").Value + "' AND PS_PP080L.U_PP030MNo = '" + RecordSet01.Fields.Item("PP030MNo").Value + "'", 0, 1)) > 0)
                                    {
                                        errCode = "17";
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
                                    PrevDBCpQty = Convert.ToInt32(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + PrevCpInfo + "' AND PS_PP040H.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                    PrevDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    PrevDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + PrevCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                    PrevMATRIXCpQty = 0;

                                    CurrentDBCpQty = Convert.ToInt32(dataHelpClass.GetValue("SELECT SUM(PS_PP040L.U_PQty) FROM [@PS_PP040H] PS_PP040H LEFT JOIN [@PS_PP040L] PS_PP040L ON PS_PP040H.DocEntry = PS_PP040L.DocEntry WHERE PS_PP040L.U_OrdMgNum = '" + CurrentCpInfo + "' AND PS_PP040L.DocEntry <> '" + RecordSet01.Fields.Item(0).Value + "' AND PS_PP040H.Canceled = 'N'", 0, 1));
                                    CurrentDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_MPO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));
                                    CurrentDBCpQty += Convert.ToInt32(dataHelpClass.GetValue("SELECT Isnull(SUM(b.U_PQty),0) * -1 FROM [@PS_CO160H] a Inner JOIN [@PS_CO160L] b ON a.DocEntry = b.DocEntry WHERE b.U_PO = '" + CurrentCpInfo + "' AND a.Canceled = 'N'", 0, 1));

                                    CurrentMATRIXCpQty = 0;

                                    if ((PrevDBCpQty + PrevMATRIXCpQty) < (CurrentDBCpQty + CurrentMATRIXCpQty))
                                    {
                                        errCode = "7";
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
            catch(Exception ex)
            {
                if (errCode == "1")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당문서는 다른사용자에 의해 취소되었습니다. 적용할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "2")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("해당작업타입은 변경이 불가능합니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "4")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("생산실적 등록된 행입니다. 적용할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "5")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("작업지시문서가 존재하지 않습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "7")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("선행공정의 생산수량이 현공정의 생산수량에 미달 합니다. 적용할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "12")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("재고가 존재하는 작번입니다. 적용할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "13")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("원자재 불출이 진행된 행은 생산수량을 수정할 수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                else if (errCode == "17")
                {
                    PSH_Globals.SBO_Application.StatusBar.SetText("생산실적 등록된 문서입니다. 적용할수 없습니다.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
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

                oForm.Items.Item("Total").Specific.Value = total.ToString("#,###.##");
            }
            catch(Exception ex)
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
            try
            {
                if (pVal.BeforeAction == true)
                {
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
            try
            {
                if (pVal.Before_Action == true)
                {
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
            try
            {
                if (pVal.Before_Action == true)
                {
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
            try
            {
                if (pVal.Before_Action == true)
                {
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
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP040H);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP040L);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP040M);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDS_PS_PP040N);
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
            try
            {
                if (pVal.Before_Action == true)
                {
                }
                else if (pVal.Before_Action == false)
                {
                    //원본 소스(VB6.0 주석처리되어 있음)
                    //if(pVal.ItemUID == "Code")
                    //{
                    //    dataHelpClass.PSH_CF_DBDatasourceReturn(pVal, pVal.FormUID, "@PH_PY001A", "Code", "", 0, "", "", "");
                    //}
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
        /// FormDataEvent
        /// </summary>
        /// <param name="FormUID"></param>
        /// <param name="BusinessObjectInfo"></param>
        /// <param name="BubbleEvent"></param>
        public override void Raise_FormDataEvent(string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                if (BusinessObjectInfo.BeforeAction == true)
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
                else if (BusinessObjectInfo.BeforeAction == false)
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


        








        #region Raise_ItemEvent
        //public void Raise_ItemEvent(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    switch (pval.EventType)
        //    {
        //        case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
        //            ////1
        //            Raise_EVENT_ITEM_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
        //            ////2
        //            Raise_EVENT_KEY_DOWN(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
        //            ////5
        //            Raise_EVENT_COMBO_SELECT(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_CLICK:
        //            ////6
        //            Raise_EVENT_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
        //            ////7
        //            Raise_EVENT_DOUBLE_CLICK(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
        //            ////8
        //            Raise_EVENT_MATRIX_LINK_PRESSED(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_VALIDATE:
        //            ////10
        //            Raise_EVENT_VALIDATE(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_MATRIX_LOAD:
        //            ////11
        //            Raise_EVENT_MATRIX_LOAD(ref FormUID, ref pval, ref BubbleEvent);

        //            // 작업시간 합계 추가 S
        //            //            Dim i&
        //            //            Dim Total As Currency
        //            //
        //            //
        //            //                For i = 0 To oMat01.VisualRowCount - 1
        //            //
        //            //                    Total = Total + Val(oMat01.Columns("WorkTime").Cells(i + 1).Specific.Value)
        //            //'                 oMat01.Columns("Total").Cells.Specific.Value = Total
        //            //                Next i
        //            //                oForm01.Items("Total").Specific.Value = Total
        //            PS_PP040_SumWorkTime();
        //            break;
        //        // 작업시간 합계 추가 E

        //        //            Call Raise_EVENT_MATRIX_LOAD(FormUID, pval, BubbleEvent)

        //        case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
        //            ////18
        //            break;
        //        ////et_FORM_ACTIVATE
        //        case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
        //            ////19
        //            break;
        //        ////et_FORM_DEACTIVATE
        //        case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE:
        //            ////20
        //            Raise_EVENT_RESIZE(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
        //            ////27
        //            Raise_EVENT_CHOOSE_FROM_LIST(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
        //            ////3
        //            Raise_EVENT_GOT_FOCUS(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //        case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
        //            ////4
        //            break;
        //        ////et_LOST_FOCUS
        //        case SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD:
        //            ////17
        //            Raise_EVENT_FORM_UNLOAD(ref FormUID, ref pval, ref BubbleEvent);
        //            break;
        //    }
        //    return;
        //Raise_ItemEvent_Error:
        //    ///''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_ItemEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_MenuEvent
        //public void Raise_MenuEvent(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    ////BeforeAction = True
        //    if ((pval.BeforeAction == true))
        //    {
        //        switch (pval.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //                {
        //                    if ((PS_PP040_Validate("취소") == false))
        //                    {
        //                        BubbleEvent = false;
        //                        return;
        //                    }
        //                    if (SubMain.Sbo_Application.MessageBox("정말로 취소하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") != Convert.ToDouble("1"))
        //                    {
        //                        BubbleEvent = false;
        //                        return;
        //                    }
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    // 분말 첫번째 공정 투입시 원자재 불출로직 추가(황영수 20181101)
        //                    if (Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.Value) == "111" | Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.Value) == "601")
        //                    {
        //                        if (Add_oInventoryGenEntry(ref 2) == false)
        //                        {
        //                            BubbleEvent = false;
        //                            return;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    MDC_Com.MDC_GF_Message(ref "현재 모드에서는 취소할수 없습니다.", ref "W");
        //                    BubbleEvent = false;
        //                    return;
        //                }
        //                break;
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1293":
        //                //행삭제
        //                Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
        //                break;
        //            case "1281":
        //                //찾기
        //                break;
        //            case "1282":
        //                //추가
        //                break;
        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                Raise_EVENT_RECORD_MOVE(ref FormUID, ref pval, ref BubbleEvent);
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((pval.BeforeAction == false))
        //    {
        //        switch (pval.MenuUID)
        //        {
        //            case "1284":
        //                //취소
        //                break;
        //            case "1286":
        //                //닫기
        //                break;
        //            case "1293":
        //                //행삭제
        //                Raise_EVENT_ROW_DELETE(ref FormUID, ref pval, ref BubbleEvent);
        //                break;
        //            case "1281":
        //                //찾기
        //                PS_PP040_FormItemEnabled();
        //                ////UDO방식
        //                oForm01.Items.Item("DocEntry").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                break;
        //            case "1282":
        //                //추가
        //                PS_PP040_FormItemEnabled();
        //                ////UDO방식
        //                PS_PP040_AddMatrixRow01(0, ref true);
        //                ////UDO방식
        //                PS_PP040_AddMatrixRow02(0, ref true);
        //                ////UDO방식
        //                break;
        //            case "1288":
        //            case "1289":
        //            case "1290":
        //            case "1291":
        //                //레코드이동버튼
        //                Raise_EVENT_RECORD_MOVE(ref FormUID, ref pval, ref BubbleEvent);
        //                break;
        //        }
        //    }
        //    return;
        //Raise_MenuEvent_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_MenuEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_FormDataEvent
        //public void Raise_FormDataEvent(ref string FormUID, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    ////BeforeAction = True
        //    if ((BusinessObjectInfo.BeforeAction == true))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //        ////BeforeAction = False
        //    }
        //    else if ((BusinessObjectInfo.BeforeAction == false))
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                ////33
        //                if ((oForm01.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE))
        //                {
        //                    if ((PS_PP040_FindValidateDocument("@PS_PP040H") == false))
        //                    {
        //                        ////찾기메뉴 활성화일때 수행
        //                        if (SubMain.Sbo_Application.Menus.Item("1281").Enabled == true)
        //                        {
        //                            SubMain.Sbo_Application.ActivateMenuItem(("1281"));
        //                        }
        //                        else
        //                        {
        //                            SubMain.Sbo_Application.SetStatusBarMessage("관리자에게 문의바랍니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //                        }
        //                        BubbleEvent = false;
        //                        return;
        //                    }
        //                }
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                ////34
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
        //                ////35
        //                break;
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE:
        //                ////36
        //                break;
        //        }
        //    }
        //    return;
        //Raise_FormDataEvent_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_FormDataEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_RightClickEvent
        //public void Raise_RightClickEvent(ref string FormUID, ref SAPbouiCOM.ContextMenuInfo pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if (pval.BeforeAction == true)
        //    {
        //        //        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //        //            Dim MenuCreationParams01 As SAPbouiCOM.MenuCreationParams
        //        //            Set MenuCreationParams01 = Sbo_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        //        //            MenuCreationParams01.Type = SAPbouiCOM.BoMenuType.mt_STRING
        //        //            MenuCreationParams01.uniqueID = "MenuUID"
        //        //            MenuCreationParams01.String = "메뉴명"
        //        //            MenuCreationParams01.Enabled = True
        //        //            Call Sbo_Application.Menus.Item("1280").SubMenus.AddEx(MenuCreationParams01)
        //        //        End If
        //    }
        //    else if (pval.BeforeAction == false)
        //    {
        //        //        If pval.ItemUID = "Mat01" And pval.Row > 0 And pval.Row <= oMat01.RowCount Then
        //        //                Call Sbo_Application.Menus.RemoveEx("MenuUID")
        //        //        End If
        //    }
        //    if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02" | pval.ItemUID == "Mat03")
        //    {
        //        if (pval.Row > 0)
        //        {
        //            oLastItemUID01 = pval.ItemUID;
        //            oLastColUID01 = pval.ColUID;
        //            oLastColRow01 = pval.Row;
        //        }
        //    }
        //    else
        //    {
        //        oLastItemUID01 = pval.ItemUID;
        //        oLastColUID01 = "";
        //        oLastColRow01 = 0;
        //    }
        //    if (pval.ItemUID == "Mat01")
        //    {
        //        if (pval.Row > 0)
        //        {
        //            oMat01Row01 = pval.Row;
        //        }
        //    }
        //    else if (pval.ItemUID == "Mat02")
        //    {
        //        if (pval.Row > 0)
        //        {
        //            oMat02Row02 = pval.Row;
        //        }
        //    }
        //    else if (pval.ItemUID == "Mat03")
        //    {
        //        if (pval.Row > 0)
        //        {
        //            oMat03Row03 = pval.Row;
        //        }
        //    }
        //    return;
        //Raise_RightClickEvent_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_RightClickEvent_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ITEM_PRESSED
        //private void Raise_EVENT_ITEM_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    string vReturnValue = null;

        //    SAPbouiCOM.ProgressBar ProgBar01 = null;
        //    ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("처리 중..", 100, false);

        //    string DocEntry = null;
        //    string LineNum = null;
        //    int i = 0;
        //    int ErrNum = 0;
        //    string DocNum = null;
        //    string WinTitle = null;
        //    string ReportName = null;
        //    string[] oText = new string[2];
        //    string sQry = null;
        //    string sQryS = null;
        //    string sQry1 = null;
        //    string WorkName = null;
        //    SAPbobsCOM.Recordset oRecordSet01 = null;
        //    if (pval.BeforeAction == true)
        //    {
        //        if (pval.ItemUID == "PS_PP040")
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //            {
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //            {
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //            {
        //            }
        //        }
        //        if (pval.ItemUID == "1")
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //            {
        //                if (PS_PP040_DataValidCheck() == false)
        //                {
        //                    BubbleEvent = false;
        //                    return;
        //                }

        //                //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                // 분말 첫번째 공정 투입시 원자재 불출로직 추가(황영수 20181101)
        //                if (Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.Value) == "111" | Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.Value) == "601")
        //                {
        //                    if (Add_oInventoryGenExit(ref 2) == false)
        //                    {
        //                        BubbleEvent = false;
        //                        return;
        //                    }
        //                    else
        //                    {
        //                    }
        //                    // End If
        //                }


        //                //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oDocEntry01 = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.Value);
        //                //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oOrdGbn = Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.Value);
        //                ////작업구분
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oSequence = oMat01.Columns.Item("Sequence").Cells.Item(1).Specific.Value;
        //                //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oDocdate = Strings.Trim(oForm01.Items.Item("DocDate").Specific.Value);
        //                oFormMode01 = oForm01.Mode;
        //                ////해야할일 작업
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //            {
        //                if (PS_PP040_DataValidCheck() == false)
        //                {
        //                    BubbleEvent = false;
        //                    return;
        //                }
        //                //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oDocEntry01 = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.Value);
        //                oFormMode01 = oForm01.Mode;
        //                ////해야할일 작업
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //            {
        //            }
        //        }

        //        ////취소버튼 누를시 저장할 자료가 있으면 메시지 표시
        //        if (pval.ItemUID == "2")
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //            {
        //                if (oMat01.VisualRowCount > 1)
        //                {
        //                    vReturnValue = Convert.ToString(SubMain.Sbo_Application.MessageBox("저장하지 않는 자료가 있습니다. 취소하시겠습니까?", 2, "&확인", "&취소"));
        //                    switch (vReturnValue)
        //                    {
        //                        case Convert.ToString(1):
        //                            break;
        //                        case Convert.ToString(2):
        //                            BubbleEvent = false;
        //                            return;

        //                            break;
        //                    }
        //                }
        //            }
        //        }

        //        if (pval.ItemUID == "Button01")
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //            {
        //                PS_PP040_OrderInfoLoad();
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //            {
        //                PS_PP040_OrderInfoLoad();
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //            {
        //            }
        //        }
        //        if (pval.ItemUID == "Button02")
        //        {

        //            oRecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //            MDC_PS_Common.ConnectODBC();

        //            //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            DocEntry = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.Value);
        //            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //            {
        //                if (oMat01.IsRowSelected(i + 1) == true)
        //                {
        //                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    LineNum = oMat01.Columns.Item("LineNum").Cells.Item(i + 1).Specific.Value;
        //                }
        //            }

        //            WinTitle = " 공정카드 [PS_PP040]";
        //            ReportName = "PS_PP040_01.rpt";

        //            sQry1 = "Select U_WorkName From [@PS_PP040M] Where DocEntry = '" + DocEntry + "' And IsNull(U_WorkName, '') <> ''";
        //            oRecordSet01.DoQuery(sQry1);

        //            while (!(oRecordSet01.EoF))
        //            {
        //                WorkName = WorkName + "     " + oRecordSet01.Fields.Item(0).Value;
        //                oRecordSet01.MoveNext();
        //            }
        //            MDC_Globals.gRpt_Formula = new string[2];
        //            MDC_Globals.gRpt_Formula_Value = new string[2];

        //            ////Formula 수식필드

        //            oText[1] = WorkName;

        //            for (i = 1; i <= 1; i++)
        //            {
        //                if (Strings.Len("" + i + "") == 1)
        //                {
        //                    MDC_Globals.gRpt_Formula[i] = "F0" + i + "";
        //                }
        //                else
        //                {
        //                    MDC_Globals.gRpt_Formula[i] = "F" + i + "";
        //                }
        //                MDC_Globals.gRpt_Formula_Value[i] = oText[i];
        //            }
        //            MDC_Globals.gRpt_SRptSqry = new string[2];
        //            MDC_Globals.gRpt_SRptName = new string[2];
        //            MDC_Globals.gRpt_SFormula = new string[2, 2];
        //            MDC_Globals.gRpt_SFormula_Value = new string[2, 2];

        //            ////SubReport

        //            MDC_Globals.gRpt_SFormula[1, 1] = "";
        //            MDC_Globals.gRpt_SFormula_Value[1, 1] = "";

        //            sQryS = "EXEC [PS_PP040_06] '" + DocEntry + "', '" + LineNum + "', 'S'";

        //            MDC_Globals.gRpt_SRptSqry[1] = sQryS;
        //            MDC_Globals.gRpt_SRptName[1] = "PS_PP040_S1";

        //            ////조회조건문
        //            sQry = "EXEC [PS_PP040_06] '" + DocEntry + "', '" + LineNum + "', 'M'";
        //            oRecordSet01.DoQuery(sQry);
        //            if (oRecordSet01.RecordCount == 0)
        //            {
        //                MDC_Com.MDC_GF_Message(ref "출력할 데이터가 없습니다.확인해 주세요.", ref "E");
        //                return;
        //            }

        //            ////CR Action
        //            if (MDC_SetMod.gCryReport_Action(WinTitle, ReportName, "Y", sQry, "1", "N", "V") == false)
        //            {
        //                SubMain.Sbo_Application.SetStatusBarMessage("gCryReport_Action : 실패!", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //            }
        //            //UPGRADE_NOTE: oRecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //            oRecordSet01 = null;
        //        }
        //    }
        //    else if (pval.BeforeAction == false)
        //    {
        //        if (pval.ItemUID == "PS_PP040")
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //            {
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //            {
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //            {
        //            }
        //        }
        //        if (pval.ItemUID == "1")
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //            {
        //                if (pval.ActionSuccess == true)
        //                {
        //                    if (oOrdGbn == "101" & oSequence == "1")
        //                    {
        //                        oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //                        PS_PP040_FormItemEnabled();
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oForm01.Items.Item("DocEntry").Specific.Value = oDocEntry01;
        //                        oForm01.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                    }
        //                    else
        //                    {
        //                        PS_PP040_FormItemEnabled();
        //                        PS_PP040_AddMatrixRow01(0, ref true);
        //                        ////UDO방식일때
        //                        PS_PP040_AddMatrixRow02(0, ref true);
        //                        ////UDO방식일때
        //                    }
        //                    //
        //                }
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //            {
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //            {
        //                if (pval.ActionSuccess == true)
        //                {
        //                    if ((oFormMode01 == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
        //                    {
        //                        oFormMode01 = SAPbouiCOM.BoFormMode.fm_OK_MODE;
        //                        oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //                        PS_PP040_FormItemEnabled();
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oForm01.Items.Item("DocEntry").Specific.Value = oDocEntry01;
        //                        oForm01.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                    }
        //                    PS_PP040_FormItemEnabled();
        //                }
        //            }
        //        }
        //        if (pval.ItemUID == "Button01")
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //            {
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //            {
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //            {
        //            }
        //        }
        //    }

        //    ProgBar01.Value = 100;
        //    ProgBar01.Stop();
        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;

        //    return;
        //Raise_EVENT_ITEM_PRESSED_Error:


        //    ProgBar01.Value = 100;
        //    ProgBar01.Stop();
        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;

        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ITEM_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_KEY_DOWN
        //private void Raise_EVENT_KEY_DOWN(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if (pval.BeforeAction == true)
        //    {
        //        if (pval.ItemUID == "OrdMgNum")
        //        {
        //            //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            ////작업타입이 일반,조정일때
        //            if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "60")
        //            {
        //                MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "OrdMgNum", "");
        //                ////사용자값활성
        //            }
        //        }





        //        if (pval.ItemUID == "Mat01")
        //        {
        //            if (pval.ColUID == "OrdMgNum")
        //            {
        //                //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                ////일반,조정, 설계
        //                if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "60" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "70")
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (oForm01.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
        //                    {
        //                        MDC_Com.MDC_GF_Message(ref "작업구분이 선택되지 않았습니다.", ref "W");
        //                        BubbleEvent = false;
        //                        return;
        //                        //UPGRADE_WARNING: oForm01.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (oForm01.Items.Item("BPLId").Specific.Selected.Value == "선택")
        //                    {
        //                        MDC_Com.MDC_GF_Message(ref "사업장이 선택되지 않았습니다.", ref "W");
        //                        BubbleEvent = false;
        //                        return;
        //                        //UPGRADE_WARNING: oForm01.Items(ItemCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (string.IsNullOrEmpty(oForm01.Items.Item("ItemCode").Specific.Value))
        //                    {
        //                        MDC_Com.MDC_GF_Message(ref "품목코드가 선택되지 않았습니다.", ref "W");
        //                        BubbleEvent = false;
        //                        return;
        //                        //UPGRADE_WARNING: oForm01.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (string.IsNullOrEmpty(oForm01.Items.Item("OrdNum").Specific.Value))
        //                    {
        //                        MDC_Com.MDC_GF_Message(ref "작지번호가 선택되지 않았습니다.", ref "W");
        //                        BubbleEvent = false;
        //                        return;
        //                        //UPGRADE_WARNING: oForm01.Items(PP030HNo).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (string.IsNullOrEmpty(oForm01.Items.Item("PP030HNo").Specific.Value))
        //                    {
        //                        MDC_Com.MDC_GF_Message(ref "작지문서번호가 선택되지 않았습니다.", ref "W");
        //                        BubbleEvent = false;
        //                        return;
        //                    }
        //                    else
        //                    {
        //                        MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "Mat01", "OrdMgNum");
        //                        ////사용자값활성
        //                    }
        //                    //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    ////지원
        //                }
        //                else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "20")
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items(OrdGbn).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (oForm01.Items.Item("OrdGbn").Specific.Selected.Value == "선택")
        //                    {
        //                        MDC_Com.MDC_GF_Message(ref "작업구분이 선택되지 않았습니다.", ref "W");
        //                        oForm01.Items.Item("OrdGbn").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                        BubbleEvent = false;
        //                        return;
        //                        //UPGRADE_WARNING: oForm01.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (oForm01.Items.Item("BPLId").Specific.Selected.Value == "선택")
        //                    {
        //                        MDC_Com.MDC_GF_Message(ref "사업장이 선택되지 않았습니다.", ref "W");
        //                        oForm01.Items.Item("BPLId").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                        BubbleEvent = false;
        //                        return;
        //                        //                    ElseIf oForm01.Items("ItemCode").Specific.Value = "" Then
        //                        //                        Call MDC_Com.MDC_GF_Message("품목코드가 선택되지 않았습니다.", "W")
        //                        //                        oForm01.Items("ItemCode").Click ct_Regular
        //                        //                        BubbleEvent = False
        //                        //                        Exit Sub
        //                        //UPGRADE_WARNING: oForm01.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (string.IsNullOrEmpty(oForm01.Items.Item("OrdNum").Specific.Value))
        //                    {
        //                        MDC_Com.MDC_GF_Message(ref "작지번호가 선택되지 않았습니다.", ref "W");
        //                        BubbleEvent = false;
        //                        return;
        //                    }
        //                    else
        //                    {
        //                        MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "Mat01", "OrdMgNum");
        //                        ////사용자값활성
        //                    }
        //                    //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    ////외주
        //                }
        //                else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "30")
        //                {

        //                    //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    ////실적
        //                }
        //                else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "40")
        //                {

        //                }

        //            }
        //        }
        //        if (pval.ItemUID == "Mat02")
        //        {
        //            if (pval.ColUID == "WorkCode")
        //            {
        //                //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                if (Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value) == 0)
        //                {
        //                    MDC_Com.MDC_GF_Message(ref "기준시간을 입력하지 않았습니다.", ref "W");
        //                    oForm01.Items.Item("BaseTime").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                    BubbleEvent = false;
        //                    return;
        //                }
        //            }
        //        }
        //        MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "Mat02", "WorkCode");
        //        //사용자값활성
        //        MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "Mat02", "NCode");
        //        //사용자값활성
        //        MDC_PS_Common.ActiveUserDefineValueAlways(ref oForm01, ref pval, ref BubbleEvent, "Mat03", "FailCode");
        //        //사용자값활성

        //        MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "Mat01", "MachCode");
        //        //설비코드 사용자값활성
        //        //        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "SubLot") 'sub작지번호 사용자값활성
        //        MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "Mat01", "CItemCod");
        //        //원재료코드 사용자값활성
        //        MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "Mat01", "SCpCode");
        //        //지원공정추가(2018.05.30 송명규)
        //        MDC_PS_Common.ActiveUserDefineValue(ref oForm01, ref pval, ref BubbleEvent, "UseMCode", "");
        //        //작업장비 사용자값활성
        //        //        Call MDC_PS_Common.ActiveUserDefineValue(oForm01, pval, BubbleEvent, "Mat01", "ItemCode") '사용자값활성
        //    }
        //    else if (pval.BeforeAction == false)
        //    {
        //        //// 화살표 이동 강제 코딩 - 류영조
        //        if (pval.ItemUID == "Mat01")
        //        {
        //            ////위쪽 화살표
        //            if (pval.CharPressed == 38)
        //            {
        //                if (pval.Row > 1 & pval.Row <= oMat01.VisualRowCount)
        //                {
        //                    oForm01.Freeze(true);
        //                    oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row - 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                    oForm01.Freeze(false);
        //                }
        //                ////아래 화살표
        //            }
        //            else if (pval.CharPressed == 40)
        //            {
        //                if (pval.Row > 0 & pval.Row < oMat01.VisualRowCount)
        //                {
        //                    oForm01.Freeze(true);
        //                    oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //                    oForm01.Freeze(false);
        //                }
        //            }

        //            //작업시간 입력 시마다 합계 계산(2011.09.26 송명규 추가)
        //            if (pval.ColUID == "WorkTime" & pval.Row != Convert.ToDouble("0"))
        //            {

        //                PS_PP040_SumWorkTime();

        //            }

        //        }
        //        else if (pval.ItemUID == "BaseTime")
        //        {

        //            //탭 키 Press
        //            if (pval.CharPressed == 9)
        //            {

        //                oMat02.Columns.Item("WorkCode").Cells.Item(1).Click();

        //            }

        //        }
        //    }
        //    return;
        //Raise_EVENT_KEY_DOWN_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_KEY_DOWN_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_COMBO_SELECT
        //private void Raise_EVENT_COMBO_SELECT(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    oForm01.Freeze(true);
        //    if (pval.BeforeAction == true)
        //    {

        //    }
        //    else if (pval.BeforeAction == false)
        //    {
        //        if (pval.ItemChanged == true)
        //        {
        //            oForm01.Freeze(true);
        //            if ((pval.ItemUID == "Mat01"))
        //            {
        //                if ((pval.ColUID == "특정컬럼"))
        //                {
        //                    ////기타작업
        //                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.Value);
        //                    if (oMat01.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP040L.GetValue("U_" + pval.ColUID, pval.Row - 1))))
        //                    {
        //                        //PS_PP040_AddMatrixRow (pval.Row)
        //                    }
        //                }
        //                else
        //                {
        //                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.Value);
        //                }
        //            }
        //            else if ((pval.ItemUID == "Mat02"))
        //            {
        //                if ((pval.ColUID == "특정컬럼"))
        //                {
        //                    ////기타작업
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.Value);
        //                    if (oMat02.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP040M.GetValue("U_" + pval.ColUID, pval.Row - 1))))
        //                    {
        //                        //PS_PP040_AddMatrixRow (pval.Row)
        //                    }
        //                }
        //                else
        //                {
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.Value);
        //                }
        //            }
        //            else if ((pval.ItemUID == "Mat03"))
        //            {
        //                if ((pval.ColUID == "특정컬럼"))
        //                {
        //                }
        //                else
        //                {
        //                    //UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040N.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Selected.Value);
        //                }
        //            }
        //            else
        //            {
        //                if ((pval.ItemUID == "OrdType"))
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value);
        //                    //UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    ////일반,조정,설계
        //                    if (oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value == "10" | oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value == "50" | oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value == "60" | oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value == "70")
        //                    {
        //                        //창원은 품목구분 선택하도록 수정 '2015/04/09
        //                        //UPGRADE_WARNING: oForm01.Items(BPLId).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (oForm01.Items.Item("BPLId").Specific.Value == "1")
        //                        {
        //                            oForm01.Items.Item("OrdGbn").Enabled = true;
        //                        }
        //                        else
        //                        {
        //                            oForm01.Items.Item("OrdGbn").Enabled = false;
        //                        }
        //                        oForm01.Items.Item("BPLId").Enabled = false;
        //                        oForm01.Items.Item("ItemCode").Enabled = false;
        //                        //UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value == "20")
        //                    {
        //                        oForm01.Items.Item("OrdGbn").Enabled = true;
        //                        oForm01.Items.Item("BPLId").Enabled = true;
        //                        oForm01.Items.Item("ItemCode").Enabled = true;
        //                        //UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value == "30")
        //                    {
        //                        oForm01.Items.Item("OrdGbn").Enabled = false;
        //                        oForm01.Items.Item("BPLId").Enabled = false;
        //                        oForm01.Items.Item("ItemCode").Enabled = false;
        //                        //UPGRADE_WARNING: oForm01.Items(pval.ItemUID).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    }
        //                    else if (oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value == "40")
        //                    {
        //                        oForm01.Items.Item("OrdGbn").Enabled = false;
        //                        oForm01.Items.Item("BPLId").Enabled = false;
        //                        oForm01.Items.Item("ItemCode").Enabled = false;
        //                    }

        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("OrdGbn").Specific.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("OrdMgNum").Specific.Value = "";
        //                    //Call oForm01.Items("BPLId").Specific.Select(0, psk_Index)
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("ItemCode").Specific.Value = "";
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("ItemName").Specific.Value = "";
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("OrdNum").Specific.Value = "";
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("OrdSub1").Specific.Value = "";
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("OrdSub2").Specific.Value = "";
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("PP030HNo").Specific.Value = "";
        //                    oMat01.Clear();
        //                    oMat01.FlushToDataSource();
        //                    oMat01.LoadFromDataSource();
        //                    PS_PP040_AddMatrixRow01(0, ref true);
        //                    oMat02.Clear();
        //                    oMat02.FlushToDataSource();
        //                    oMat02.LoadFromDataSource();
        //                    PS_PP040_AddMatrixRow02(0, ref true);
        //                    oMat03.Clear();
        //                    oMat03.FlushToDataSource();
        //                    oMat03.LoadFromDataSource();
        //                }
        //                else if ((pval.ItemUID == "OrdGbn"))
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value);
        //                    oMat01.Clear();
        //                    oMat01.FlushToDataSource();
        //                    oMat01.LoadFromDataSource();
        //                    PS_PP040_AddMatrixRow01(0, ref true);
        //                    oMat02.Clear();
        //                    oMat02.FlushToDataSource();
        //                    oMat02.LoadFromDataSource();
        //                    PS_PP040_AddMatrixRow02(0, ref true);
        //                    oMat03.Clear();
        //                    oMat03.FlushToDataSource();
        //                    oMat03.LoadFromDataSource();
        //                }
        //                else if ((pval.ItemUID == "BPLId"))
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value);
        //                    oMat01.Clear();
        //                    oMat01.FlushToDataSource();
        //                    oMat01.LoadFromDataSource();
        //                    PS_PP040_AddMatrixRow01(0, ref true);
        //                    oMat02.Clear();
        //                    oMat02.FlushToDataSource();
        //                    oMat02.LoadFromDataSource();
        //                    PS_PP040_AddMatrixRow02(0, ref true);
        //                    oMat03.Clear();
        //                    oMat03.FlushToDataSource();
        //                    oMat03.LoadFromDataSource();
        //                }
        //                else
        //                {
        //                    //거래처구분이 아닐 경우만 실행(2012.02.02 송명규 추가)
        //                    if (pval.ItemUID != "CardType")
        //                    {
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Selected.Value);
        //                    }
        //                }
        //            }
        //            oMat01.LoadFromDataSource();
        //            oMat01.AutoResizeColumns();
        //            oMat02.LoadFromDataSource();
        //            oMat02.AutoResizeColumns();
        //            oMat03.LoadFromDataSource();
        //            oMat03.AutoResizeColumns();
        //            oForm01.Update();
        //            if (pval.ItemUID == "Mat01")
        //            {
        //                oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //            }
        //            else if (pval.ItemUID == "Mat02")
        //            {
        //                oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //            }
        //            else if (pval.ItemUID == "Mat03")
        //            {
        //                oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Collapsed);
        //            }
        //            else
        //            {

        //            }
        //            oForm01.Freeze(false);
        //        }
        //    }
        //    oForm01.Freeze(false);
        //    return;
        //Raise_EVENT_COMBO_SELECT_Error:
        //    oForm01.Freeze(false);
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_COMBO_SELECT_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CLICK
        //private void Raise_EVENT_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    object TempForm01 = null;

        //    if (pval.BeforeAction == true)
        //    {
        //        if (pval.ItemUID == "Opt01")
        //        {
        //            oForm01.Freeze(true);
        //            oForm01.Settings.MatrixUID = "Mat02";
        //            oForm01.Settings.EnableRowFormat = true;
        //            oForm01.Settings.Enabled = true;
        //            oMat01.AutoResizeColumns();
        //            oMat02.AutoResizeColumns();
        //            oMat03.AutoResizeColumns();
        //            oForm01.Freeze(false);
        //        }
        //        if (pval.ItemUID == "Opt02")
        //        {
        //            oForm01.Freeze(true);
        //            oForm01.Settings.MatrixUID = "Mat03";
        //            oForm01.Settings.EnableRowFormat = true;
        //            oForm01.Settings.Enabled = true;
        //            oMat01.AutoResizeColumns();
        //            oMat02.AutoResizeColumns();
        //            oMat03.AutoResizeColumns();
        //            oForm01.Freeze(false);
        //        }
        //        if (pval.ItemUID == "Opt03")
        //        {
        //            oForm01.Freeze(true);
        //            oForm01.Settings.MatrixUID = "Mat01";
        //            oForm01.Settings.EnableRowFormat = true;
        //            oForm01.Settings.Enabled = true;
        //            oMat01.AutoResizeColumns();
        //            oMat02.AutoResizeColumns();
        //            oMat03.AutoResizeColumns();
        //            oForm01.Freeze(false);
        //        }
        //        //        If pval.ItemUID = "Mat01" Then
        //        //            If pval.Row > 0 Then
        //        //                Call oMat01.SelectRow(pval.Row, True, False)
        //        //            End If
        //        //        End If
        //        if (pval.ItemUID == "Mat01")
        //        {
        //            if (pval.Row > 0)
        //            {
        //                oMat01.SelectRow(pval.Row, true, false);
        //                oMat01Row01 = pval.Row;
        //            }
        //        }
        //        if (pval.ItemUID == "Mat02")
        //        {
        //            if (pval.Row > 0)
        //            {
        //                oMat02.SelectRow(pval.Row, true, false);
        //                oMat02Row02 = pval.Row;
        //            }
        //        }
        //        if (pval.ItemUID == "Mat03")
        //        {
        //            if (pval.Row > 0)
        //            {
        //                oMat03.SelectRow(pval.Row, true, false);
        //                oMat03Row03 = pval.Row;
        //            }
        //        }
        //    }
        //    else if (pval.BeforeAction == false)
        //    {
        //        //// 작업지시번호 링크 번튼 - 류영조
        //        if (pval.ItemUID == "LBtn01")
        //        {
        //            TempForm01 = new PS_PP030();
        //            //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            //UPGRADE_WARNING: TempForm01.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            TempForm01.LoadForm(oForm01.Items.Item("PP030HNo").Specific.Value);
        //            //UPGRADE_NOTE: TempForm01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //            TempForm01 = null;
        //        }
        //    }
        //    return;
        //Raise_EVENT_CLICK_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_DOUBLE_CLICK
        //private void Raise_EVENT_DOUBLE_CLICK(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if (pval.BeforeAction == true)
        //    {
        //        if (pval.ItemUID == "Mat01")
        //        {
        //            if (pval.Row > 0)
        //            {
        //                //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                ////작업타입이 일반,조정인경우
        //                if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "60")
        //                {
        //                    //UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pval.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.Value))
        //                    {

        //                    }
        //                    else
        //                    {
        //                        if (oMat03.VisualRowCount == 0)
        //                        {
        //                            PS_PP040_AddMatrixRow03(0, ref true);
        //                        }
        //                        else
        //                        {
        //                            PS_PP040_AddMatrixRow03(oMat03.VisualRowCount);
        //                        }
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pval.Row).Specific.Value);
        //                        oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pval.Row));
        //                        oMat03.LoadFromDataSource();
        //                        oMat03.AutoResizeColumns();
        //                        //                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
        //                        //                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
        //                        oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
        //                        oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
        //                        oMat03.FlushToDataSource();
        //                    }
        //                    //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    ////작업타입이 PSMT지원인경우
        //                }
        //                else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "20")
        //                {
        //                    //UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pval.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (string.IsNullOrEmpty(oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.Value))
        //                    {

        //                    }
        //                    else
        //                    {
        //                        if (oMat03.VisualRowCount == 0)
        //                        {
        //                            PS_PP040_AddMatrixRow03(0, ref true);
        //                        }
        //                        else
        //                        {
        //                            PS_PP040_AddMatrixRow03(oMat03.VisualRowCount);
        //                        }
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(pval.Row).Specific.Value);
        //                        oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pval.Row));
        //                        oMat03.LoadFromDataSource();
        //                        oMat03.AutoResizeColumns();
        //                        //                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
        //                        //                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
        //                        oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
        //                        oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
        //                        oMat03.FlushToDataSource();
        //                    }
        //                    //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    ////작업타입이 외주인경우
        //                }
        //                else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "30")
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    ////작업타입이 실적인경우
        //                }
        //                else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "40")
        //                {
        //                }
        //            }
        //        }
        //    }
        //    else if (pval.BeforeAction == false)
        //    {

        //    }
        //    return;
        //Raise_EVENT_DOUBLE_CLICK_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_DOUBLE_CLICK_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LINK_PRESSED
        //private void Raise_EVENT_MATRIX_LINK_PRESSED(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    object oTempClass = null;
        //    if (pval.BeforeAction == true)
        //    {
        //        if (pval.ItemUID == "Mat01")
        //        {
        //            if (pval.ColUID == "OrdMgNum")
        //            {
        //                oTempClass = new PS_PP030();
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oTempClass.LoadForm(Strings.Mid(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, "-") - 1));
        //            }
        //            if (pval.ColUID == "PP030HNo")
        //            {
        //                oTempClass = new PS_PP030();
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oTempClass.LoadForm(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //            }
        //        }
        //        if (pval.ItemUID == "Mat03")
        //        {
        //            if (pval.ColUID == "OrdMgNum")
        //            {
        //                oTempClass = new PS_PP030();
        //                //UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: oTempClass.LoadForm 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oTempClass.LoadForm(Strings.Mid(oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, 1, Strings.InStr(oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, "-") - 1));
        //            }
        //        }
        //    }
        //    else if (pval.BeforeAction == false)
        //    {

        //    }
        //    return;
        //Raise_EVENT_MATRIX_LINK_PRESSED_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LINK_PRESSED_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_VALIDATE
        //private void Raise_EVENT_VALIDATE(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    int i = 0;
        //    string Query01 = null;
        //    SAPbobsCOM.Recordset RecordSet01 = null;
        //    double Weight = 0;

        //    double Time = 0;
        //    //UPGRADE_NOTE: Hour이(가) Hour_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //    int Hour_Renamed = 0;
        //    //UPGRADE_NOTE: Minute이(가) Minute_Renamed(으)로 업그레이드되었습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
        //    int Minute_Renamed = 0;

        //    oForm01.Freeze(true);
        //    SAPbouiCOM.ProgressBar ProgBar01 = null;
        //    string WkCmDt = null;
        //    string OINV_Dt = null;
        //    string ReturnValue = null;
        //    if (pval.BeforeAction == true)
        //    {
        //        if (pval.ItemChanged == true)
        //        {
        //            if ((pval.ItemUID == "Mat01"))
        //            {
        //                if ((PS_PP040_Validate("수정01") == false))
        //                {
        //                    oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, Strings.Trim(oDS_PS_PP040L.GetValue("U_" + pval.ColUID, pval.Row - 1)));
        //                }
        //                else
        //                {
        //                    if ((pval.ColUID == "OrdMgNum"))
        //                    {
        //                        RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //                        ProgBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("실행 중...", 100, false);

        //                        //UPGRADE_WARNING: oForm01.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        ////작지번호에 값이 없으면 작업지시가 불러오기전
        //                        if (string.IsNullOrEmpty(oForm01.Items.Item("OrdNum").Specific.Value))
        //                        {
        //                            oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //                            ////작업지시가 선택된상태
        //                        }
        //                        else
        //                        {
        //                            //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            ////작업타입이 일반,조정, 설계
        //                            if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "60" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "70")
        //                            {
        //                                ////작지문서헤더번호가 일치하지 않으면
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                //UPGRADE_WARNING: oForm01.Items(PP030HNo).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                if (oForm01.Items.Item("PP030HNo").Specific.Value != Strings.Mid(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, "-") - 1))
        //                                {
        //                                    oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //                                    ////작지문서번호가 일치하면
        //                                }
        //                                else
        //                                {
        //                                    //UPGRADE_WARNING: oForm01.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    if (oForm01.Items.Item("BPLId").Specific.Selected.Value != "1")
        //                                    {
        //                                        ////신동사업부를 제외한 사업부만 체크
        //                                        for (i = 1; i <= oMat01.RowCount; i++)
        //                                        {
        //                                            ////현재 입력한 값이 이미 입력되어 있는경우
        //                                            //UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pval.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                            //UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                            if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.Value & i != pval.Row)
        //                                            {
        //                                                MDC_Com.MDC_GF_Message(ref "이미 입력한 공정입니다.", ref "W");
        //                                                oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //                                                goto Continue_Renamed;
        //                                            }
        //                                            //                                        '//공정라인의 공정순서가 앞공정보다 높으면
        //                                            //                                        If Val(oMat01.Columns("Sequence").Cells(i).Specific.Value) >= MDC_PS_Common.GetValue("SELECT PS_PP030M.U_Sequence FROM [@PS_PP030H] PS_PP030H LEFT JOIN [@PS_PP030M] PS_PP030M ON PS_PP030H.DocEntry = PS_PP030M.DocEntry WHERE CONVERT(NVARCHAR,PS_PP030M.DocEntry) + '-' + CONVERT(NVARCHAR,PS_PP030M.LineId) = '" & oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.Value & "'") Then
        //                                            //                                            Call MDC_Com.MDC_GF_Message("공정순서가 올바르지 않습니다.", "W")
        //                                            //                                            Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //                                            //                                            GoTo Continue
        //                                            //                                        End If
        //                                        }

        //                                        //생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_S
        //                                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                        Query01 = "EXEC PS_PP040_90 '" + Strings.Mid(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, "-") - 1) + "'";
        //                                        //oMat01.Columns("OrdMgNum").Cells(pval.Row).Specific.Value & "'"
        //                                        RecordSet01.DoQuery(Query01);
        //                                        //UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                        WkCmDt = RecordSet01.Fields.Item("WkCmDt").Value;

        //                                        //생산완료수량이 작업지시수량만큼 모두 등록이 되었다면
        //                                        if (RecordSet01.Fields.Item("Return").Value == "1")
        //                                        {
        //                                            if (SubMain.Sbo_Application.MessageBox("생산완료가 모두 등록된 작번(완료일자:" + WkCmDt + ")입니다. 계속 진행하시겠습니까?", Convert.ToInt32("1"), "예", "아니오") == Convert.ToDouble("1"))
        //                                            {
        //                                                //계속 진행시에는 해당 작업지시문서번호 등록
        //                                            }
        //                                            else
        //                                            {
        //                                                //                                                Call MDC_Com.MDC_GF_Message("등록이 취소되었습니다.", "W")
        //                                                oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //                                                goto Continue_Renamed;
        //                                            }
        //                                        }
        //                                        //생산완료등록이 완료된 작번인지 체크_수량으로 비교(2012.08.27 송명규 추가)_E

        //                                        //판매완료등록 체크_S(2015.07.14 송명규 추가)
        //                                        Query01 = "EXEC PS_PP040_91 '";
        //                                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                        Query01 = Query01 + Strings.Mid(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, 1, Strings.InStr(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value, "-") - 1) + "','";
        //                                        Query01 = Query01 + oDS_PS_PP040H.GetValue("U_DocDate", 0) + "'";
        //                                        RecordSet01.DoQuery(Query01);
        //                                        //UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                        OINV_Dt = RecordSet01.Fields.Item("OINV_Dt").Value;

        //                                        //판매확정수량이 판매오더수량만큼 모두 등록이 되었다면
        //                                        if (RecordSet01.Fields.Item("Return").Value == "1")
        //                                        {
        //                                            SubMain.Sbo_Application.MessageBox("판매완료(최종일자:" + OINV_Dt + ")된 작번입니다. 등록이 불가능합니다.", 1, "확인");
        //                                            oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //                                            goto Continue_Renamed;
        //                                        }
        //                                        //판매완료등록 체크_E(2015.07.14 송명규 추가)


        //                                    }

        //                                    //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    Query01 = "EXEC PS_PP040_01 '" + oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value + "', '" + oForm01.Items.Item("OrdType").Specific.Selected.Value + "'";
        //                                    RecordSet01.DoQuery(Query01);
        //                                    if (RecordSet01.RecordCount == 0)
        //                                    {
        //                                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //                                    }
        //                                    else
        //                                    {
        //                                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
        //                                        oDS_PS_PP040L.SetValue("U_Sequence", pval.Row - 1, RecordSet01.Fields.Item("Sequence").Value);
        //                                        oDS_PS_PP040L.SetValue("U_CpCode", pval.Row - 1, RecordSet01.Fields.Item("CpCode").Value);
        //                                        oDS_PS_PP040L.SetValue("U_CpName", pval.Row - 1, RecordSet01.Fields.Item("CpName").Value);
        //                                        oDS_PS_PP040L.SetValue("U_OrdGbn", pval.Row - 1, RecordSet01.Fields.Item("OrdGbn").Value);
        //                                        oDS_PS_PP040L.SetValue("U_BPLId", pval.Row - 1, RecordSet01.Fields.Item("BPLId").Value);
        //                                        oDS_PS_PP040L.SetValue("U_ItemCode", pval.Row - 1, RecordSet01.Fields.Item("ItemCode").Value);
        //                                        oDS_PS_PP040L.SetValue("U_ItemName", pval.Row - 1, RecordSet01.Fields.Item("ItemName").Value);
        //                                        oDS_PS_PP040L.SetValue("U_OrdNum", pval.Row - 1, RecordSet01.Fields.Item("OrdNum").Value);
        //                                        oDS_PS_PP040L.SetValue("U_OrdSub1", pval.Row - 1, RecordSet01.Fields.Item("OrdSub1").Value);
        //                                        oDS_PS_PP040L.SetValue("U_OrdSub2", pval.Row - 1, RecordSet01.Fields.Item("OrdSub2").Value);
        //                                        oDS_PS_PP040L.SetValue("U_PP030HNo", pval.Row - 1, RecordSet01.Fields.Item("PP030HNo").Value);
        //                                        oDS_PS_PP040L.SetValue("U_PP030MNo", pval.Row - 1, RecordSet01.Fields.Item("PP030MNo").Value);
        //                                        oDS_PS_PP040L.SetValue("U_SelWt", pval.Row - 1, RecordSet01.Fields.Item("SelWt").Value);
        //                                        oDS_PS_PP040L.SetValue("U_PSum", pval.Row - 1, RecordSet01.Fields.Item("PSum").Value);
        //                                        oDS_PS_PP040L.SetValue("U_BQty", pval.Row - 1, RecordSet01.Fields.Item("BQty").Value);
        //                                        oDS_PS_PP040L.SetValue("U_PQty", pval.Row - 1, Convert.ToString(0));
        //                                        oDS_PS_PP040L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(0));
        //                                        oDS_PS_PP040L.SetValue("U_YQty", pval.Row - 1, Convert.ToString(0));
        //                                        oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(0));
        //                                        oDS_PS_PP040L.SetValue("U_NQty", pval.Row - 1, Convert.ToString(0));
        //                                        oDS_PS_PP040L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(0));
        //                                        oDS_PS_PP040L.SetValue("U_ScrapWt", pval.Row - 1, Convert.ToString(0));
        //                                        oDS_PS_PP040L.SetValue("U_WorkTime", pval.Row - 1, Convert.ToString(0));
        //                                        oDS_PS_PP040L.SetValue("U_LineId", pval.Row - 1, "");

        //                                        ////설비코드,명 Reset
        //                                        oDS_PS_PP040L.SetValue("U_MachCode", pval.Row - 1, "");
        //                                        oDS_PS_PP040L.SetValue("U_MachName", pval.Row - 1, "");
        //                                        ////불량코드테이블
        //                                        if (oMat03.VisualRowCount == 0)
        //                                        {
        //                                            PS_PP040_AddMatrixRow03(0, ref true);
        //                                        }
        //                                        else
        //                                        {
        //                                            PS_PP040_AddMatrixRow03(oMat03.VisualRowCount);
        //                                        }

        //                                        oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("OrdMgNum").Value);
        //                                        oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpCode").Value);
        //                                        oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, RecordSet01.Fields.Item("CpName").Value);
        //                                        oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, Convert.ToString(pval.Row));



        //                                        //// 류영조
        //                                        //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                        if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "60")
        //                                        {
        //                                            oDS_PS_PP040H.SetValue("U_BaseTime", 0, "1");
        //                                            //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                            oMat02.Columns.Item("WorkCode").Cells.Item(1).Specific.Value = "9999999";
        //                                            //                                            oMat02.Columns("WorkName").Cells(1).Specific.Value = "조정"
        //                                            //                                            Call oDS_PS_PP040M.setValue("U_WorkCode", 0, "9999999")
        //                                            oDS_PS_PP040M.SetValue("U_WorkName", 0, "조정");
        //                                            oMat02.LoadFromDataSource();
        //                                        }
        //                                        else
        //                                        {
        //                                            //                                            Call oDS_PS_PP040H.setValue("U_BaseTime", 0, "")
        //                                            //                                            oMat02.Columns("WorkCode").Cells(1).Specific.Value = ""
        //                                            //                                            oMat02.Columns("WorkName").Cells(1).Specific.Value = ""
        //                                            //                        Call oDS_PS_PP040M.setValue("U_WorkCode", 0, "")
        //                                            //                        Call oDS_PS_PP040M.setValue("U_WorkName", 0, "")
        //                                        }
        //                                    }
        //                                }
        //                                //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                ////작업타입이 PSMT지원
        //                            }
        //                            else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "20")
        //                            {
        //                                ////올바른 공정코드인지 검사
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                //UPGRADE_WARNING: MDC_PS_Common.GetValue(SELECT COUNT(*) FROM [PS_PP001L] WHERE U_CpCode = ' & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value & ') 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                if (MDC_PS_Common.GetValue("SELECT COUNT(*) FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value + "'") == 0)
        //                                {
        //                                    oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //                                }
        //                                else
        //                                {
        //                                    for (i = 1; i <= oMat01.RowCount; i++)
        //                                    {
        //                                        ////현재 입력한 값이 이미 입력되어 있는경우
        //                                        //UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(pval.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                        //UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat01.Columns.Item("OrdMgNum").Cells.Item(pval.Row).Specific.Value & i != pval.Row)
        //                                        {
        //                                            MDC_Com.MDC_GF_Message(ref "이미 입력한 공정입니다.", ref "W");
        //                                            oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, "");
        //                                            goto Continue_Renamed;
        //                                        }
        //                                    }
        //                                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040L.SetValue("U_CpCode", pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040L.SetValue("U_CpName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value + "'", 0, 1));
        //                                    //UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040L.SetValue("U_OrdGbn", pval.Row - 1, oForm01.Items.Item("OrdGbn").Specific.Selected.Value);
        //                                    //UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040L.SetValue("U_BPLId", pval.Row - 1, oForm01.Items.Item("BPLId").Specific.Selected.Value);
        //                                    oDS_PS_PP040L.SetValue("U_ItemCode", pval.Row - 1, "");
        //                                    oDS_PS_PP040L.SetValue("U_ItemName", pval.Row - 1, "");
        //                                    ////PSMT지원은 품목코드 필요없음
        //                                    //                                    Call oDS_PS_PP040L.setValue("U_ItemCode", pval.Row - 1, oForm01.Items("ItemCode").Specific.Value)
        //                                    //                                    Call oDS_PS_PP040L.setValue("U_ItemName", pval.Row - 1, oForm01.Items("ItemName").Specific.Value)
        //                                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040L.SetValue("U_OrdNum", pval.Row - 1, oForm01.Items.Item("OrdNum").Specific.Value);
        //                                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040L.SetValue("U_OrdSub1", pval.Row - 1, oForm01.Items.Item("OrdSub1").Specific.Value);
        //                                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040L.SetValue("U_OrdSub2", pval.Row - 1, oForm01.Items.Item("OrdSub2").Specific.Value);
        //                                    oDS_PS_PP040L.SetValue("U_PP030HNo", pval.Row - 1, "");
        //                                    oDS_PS_PP040L.SetValue("U_PP030MNo", pval.Row - 1, "");
        //                                    oDS_PS_PP040L.SetValue("U_PSum", pval.Row - 1, Convert.ToString(0));
        //                                    oDS_PS_PP040L.SetValue("U_PQty", pval.Row - 1, Convert.ToString(0));
        //                                    oDS_PS_PP040L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(0));
        //                                    oDS_PS_PP040L.SetValue("U_YQty", pval.Row - 1, Convert.ToString(0));
        //                                    oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(0));
        //                                    oDS_PS_PP040L.SetValue("U_NQty", pval.Row - 1, Convert.ToString(0));
        //                                    oDS_PS_PP040L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(0));
        //                                    oDS_PS_PP040L.SetValue("U_ScrapWt", pval.Row - 1, Convert.ToString(0));
        //                                    ////불량코드테이블
        //                                    if (oMat03.VisualRowCount == 0)
        //                                    {
        //                                        PS_PP040_AddMatrixRow03(0, ref true);
        //                                    }
        //                                    else
        //                                    {
        //                                        if (oDS_PS_PP040L.GetValue("U_OrdMgNum", pval.Row - 1) == oDS_PS_PP040N.GetValue("U_OrdMgNum", oMat03.VisualRowCount - 1))
        //                                        {
        //                                        }
        //                                        else
        //                                        {
        //                                            PS_PP040_AddMatrixRow03(oMat03.VisualRowCount);
        //                                        }
        //                                    }
        //                                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                    oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value + "'", 0, 1));
        //                                }
        //                                //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                ////작업타입이 외주
        //                            }
        //                            else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "30")
        //                            {

        //                                //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                ////작업타입이 실적
        //                            }
        //                            else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "40")
        //                            {

        //                            }
        //                        Continue_Renamed:
        //                            if (oMat01.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP040L.GetValue("U_" + pval.ColUID, pval.Row - 1))))
        //                            {
        //                                PS_PP040_AddMatrixRow01(pval.Row);
        //                            }
        //                        }

        //                        ProgBar01.Value = 100;
        //                        ProgBar01.Stop();
        //                        //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                        ProgBar01 = null;

        //                        //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                        RecordSet01 = null;
        //                    }
        //                    else if (pval.ColUID == "PQty")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value) <= 0)
        //                        {
        //                            if (Strings.Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) == "60")
        //                            {
        //                                goto Skip_PQty;
        //                            }
        //                            else
        //                            {
        //                                oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oDS_PS_PP040L.GetValue("U_" + pval.ColUID, pval.Row - 1));
        //                            }
        //                        }
        //                        else
        //                        {
        //                        Skip_PQty:
        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            oDS_PS_PP040L.SetValue("U_YQty", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                            //UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(pval.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            Weight = Conversion.Val(MDC_PS_Common.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.Value + "'", 0, 1)) / 1000;
        //                            if (Weight == 0)
        //                            {
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                oDS_PS_PP040L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                            }
        //                            else
        //                            {
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                oDS_PS_PP040L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                            }
        //                            oDS_PS_PP040L.SetValue("U_NQty", pval.Row - 1, Convert.ToString(0));
        //                            oDS_PS_PP040L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(0));
        //                        }
        //                    }
        //                    else if (pval.ColUID == "NQty")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value) <= 0)
        //                        {
        //                            if (Strings.Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) == "60")
        //                            {
        //                                goto skip_Nqty;
        //                            }
        //                            else
        //                            {
        //                                oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oDS_PS_PP040L.GetValue("U_" + pval.ColUID, pval.Row - 1));
        //                            }
        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        }
        //                        else if (Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value) > Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pval.Row).Specific.Value))
        //                        {
        //                            if (Strings.Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) == "50" | Strings.Trim(oDS_PS_PP040H.GetValue("U_OrdType", 0)) == "60")
        //                            {
        //                                goto skip_Nqty;
        //                            }
        //                            else
        //                            {
        //                                oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oDS_PS_PP040L.GetValue("U_" + pval.ColUID, pval.Row - 1));
        //                            }
        //                        }
        //                        else
        //                        {
        //                        skip_Nqty:
        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            oDS_PS_PP040L.SetValue("U_YQty", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pval.Row).Specific.Value) - Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                            //UPGRADE_WARNING: oMat01.Columns(CpCode).Cells(pval.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            Weight = Conversion.Val(MDC_PS_Common.GetValue("SELECT U_CpUnWt  FROM [@PS_PP004H] WHERE U_ItemCode = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.Value + "' AND U_CpCode = '" + oMat01.Columns.Item("CpCode").Cells.Item(pval.Row).Specific.Value + "'", 0, 1)) / 1000;
        //                            if (Weight == 0)
        //                            {
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                oDS_PS_PP040L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pval.Row).Specific.Value) - Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                            }
        //                            else
        //                            {
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                oDS_PS_PP040L.SetValue("U_NWeight", pval.Row - 1, Convert.ToString(Weight * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                                oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Weight * (Conversion.Val(oMat01.Columns.Item("PQty").Cells.Item(pval.Row).Specific.Value) - Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value))));
        //                            }
        //                        }

        //                        //작업시간(공수)을 입력할 때
        //                    }
        //                    else if (pval.ColUID == "WorkTime")
        //                    {

        //                        RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //                        //UPGRADE_WARNING: oForm01.Items(BPLId).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (oForm01.Items.Item("BPLId").Specific.Selected.Value != "1")
        //                        {

        //                            //적자 여부 확인 체크(2016.05.20 송명규 추가)_S
        //                            //                            Query01 = "EXEC PS_PP040_92 '"
        //                            //                            Query01 = Query01 & oMat01.Columns("PP030HNo").Cells(pval.Row).Specific.Value & "','"
        //                            //                            Query01 = Query01 & oMat01.Columns("PP030MNo").Cells(pval.Row).Specific.Value & "'"
        //                            //
        //                            //                            Call RecordSet01.DoQuery(Query01)
        //                            //
        //                            //                            ReturnValue = RecordSet01.Fields("ReturnValue").Value
        //                            //
        //                            //                            If ReturnValue <> "" Then '적자가 예상되는 작번은 메시지 출력
        //                            //                                If Sbo_Application.MessageBox(ReturnValue, "1", "예", "아니오") = "1" Then
        //                            //                                    '계속 진행시에는 해당 공수 등록
        //                            //                                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, Val(oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value))
        //                            //                                Else
        //                            //                                    Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, 0)
        //                            //    '                                Call oDS_PS_PP040L.setValue("U_" & pval.ColUID, pval.Row - 1, "")
        //                            //    '                                GoTo Continue
        //                            //                                End If
        //                            //
        //                            //                            Else

        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));

        //                            //                            End If
        //                            //적자 여부 확인 체크(2016.05.20 송명규 추가)_S

        //                        }
        //                        else
        //                        {

        //                            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));

        //                        }

        //                        //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                        RecordSet01 = null;

        //                        ////기존도면매수
        //                    }
        //                    else if (pval.ColUID == "BdwQty")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_AdwQTy", pval.Row - 1, Convert.ToString((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_PQTy", pval.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_YQTy", pval.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("DwRate").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pval.Row).Specific.Value)));
        //                        ////도면 적용율
        //                    }
        //                    else if (pval.ColUID == "DwRate")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_AdwQTy", pval.Row - 1, Convert.ToString((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_PQTy", pval.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_YQTy", pval.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(((Conversion.Val(oMat01.Columns.Item("BdwQty").Cells.Item(pval.Row).Specific.Value) * Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)) / 100) + Conversion.Val(oMat01.Columns.Item("NdwQTy").Cells.Item(pval.Row).Specific.Value)));
        //                        ////신규도면매수
        //                    }
        //                    else if (pval.ColUID == "NdwQTy")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_PQTy", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("AdwQty").Cells.Item(pval.Row).Specific.Value) + Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_PWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("AdwQty").Cells.Item(pval.Row).Specific.Value) + Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_YQTy", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("AdwQty").Cells.Item(pval.Row).Specific.Value) + Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_YWeight", pval.Row - 1, Convert.ToString(Conversion.Val(oMat01.Columns.Item("AdwQty").Cells.Item(pval.Row).Specific.Value) + Conversion.Val(oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                    }
        //                    else if (pval.ColUID == "MachCode")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_MachName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_MachName FROM [@PS_PP130H] WHERE U_MachCode = '" + oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value + "'", 0, 1));
        //                        //Call oDS_PS_PP040L.setValue("U_MachGrCd", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_MacdGrCd FROM [@PS_PP130H] WHERE U_MachCode = '" & oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value & "'", 0, 1))
        //                    }
        //                    else if (pval.ColUID == "CItemCod")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns(pval.ColUID).Cells(pval.Row).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_CItemNam", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_ItemNam2 FROM [@PS_PP005H] WHERE U_ItemCod1 = '" + oMat01.Columns.Item("ItemCode").Cells.Item(pval.Row).Specific.Value + "' and U_ItemCod2 = '" + oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value + "'", 0, 1));
        //                        //지원공정코드
        //                    }
        //                    else if (pval.ColUID == "SCpCode")
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_SCpName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_CpName FROM [@PS_PP001L] WHERE U_CpCode = '" + oMat01.Columns.Item("SCpCode").Cells.Item(pval.Row).Specific.Value + "'", 0, 1));
        //                    }
        //                    else
        //                    {
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040L.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                    }
        //                }
        //            }
        //            else if ((pval.ItemUID == "Mat02"))
        //            {
        //                if ((pval.ColUID == "WorkCode"))
        //                {
        //                    ////기타작업
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_WorkName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT LastName + FirstName FROM [OHEM] WHERE U_MSTCOD = '" + oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value + "'", 0, 1));
        //                    if (oMat02.RowCount == pval.Row & !string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP040M.GetValue("U_" + pval.ColUID, pval.Row - 1))))
        //                    {
        //                        PS_PP040_AddMatrixRow02(pval.Row);
        //                    }
        //                }
        //                else if (pval.ColUID == "NStart")
        //                {
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.Value) == 0 | Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.Value) == 0)
        //                    {
        //                        oDS_PS_PP040M.SetValue("U_NTime", pval.Row - 1, Convert.ToString(0));
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040M.SetValue("U_YTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value)));
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value)));
        //                    }
        //                    else
        //                    {
        //                        //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.Value) <= Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.Value))
        //                        {
        //                            //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            Time = Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.Value) - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.Value);
        //                        }
        //                        else
        //                        {
        //                            //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            Time = (2400 - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.Value)) + Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.Value);
        //                        }
        //                        Hour_Renamed = Conversion.Fix(Time / 100);
        //                        //UPGRADE_WARNING: Mod에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        //                        Minute_Renamed = Time % 100;
        //                        Time = Hour_Renamed;
        //                        if (Minute_Renamed > 0)
        //                        {
        //                            Time = Time + 0.5;
        //                        }
        //                        oDS_PS_PP040M.SetValue("U_NTime", pval.Row - 1, Convert.ToString(Time));
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040M.SetValue("U_YTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value) - Time));
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value) - Time));
        //                    }
        //                }
        //                else if (pval.ColUID == "NEnd")
        //                {
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.Value) == 0 | Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.Value) == 0)
        //                    {
        //                        oDS_PS_PP040M.SetValue("U_NTime", pval.Row - 1, Convert.ToString(0));
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040M.SetValue("U_YTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value)));
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value)));
        //                    }
        //                    else
        //                    {
        //                        //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.Value) <= Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.Value))
        //                        {
        //                            //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            Time = Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.Value) - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.Value);
        //                        }
        //                        else
        //                        {
        //                            //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                            Time = (2400 - Conversion.Val(oMat02.Columns.Item("NStart").Cells.Item(pval.Row).Specific.Value)) + Conversion.Val(oMat02.Columns.Item("NEnd").Cells.Item(pval.Row).Specific.Value);
        //                        }
        //                        Hour_Renamed = Conversion.Fix(Time / 100);
        //                        //UPGRADE_WARNING: Mod에 새 동작이 있습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        //                        Minute_Renamed = Time % 100;
        //                        Time = Hour_Renamed;
        //                        if (Minute_Renamed > 0)
        //                        {
        //                            Time = Time + 0.5;
        //                        }
        //                        oDS_PS_PP040M.SetValue("U_NTime", pval.Row - 1, Convert.ToString(Time));
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040M.SetValue("U_YTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value) - Time));
        //                        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oForm01.Items.Item("BaseTime").Specific.Value) - Time));
        //                    }
        //                }
        //                else if (pval.ColUID == "YTime")
        //                {
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_" + pval.ColUID, pval.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_TTime", pval.Row - 1, Convert.ToString(Conversion.Val(oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value)));
        //                }
        //                else
        //                {
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040M.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                }
        //            }
        //            else if ((pval.ItemUID == "Mat03"))
        //            {
        //                if ((pval.ColUID == "FailCode"))
        //                {
        //                    //UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040N.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                    //UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    //UPGRADE_WARNING: MDC_PS_Common.GetValue() 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040N.SetValue("U_FailName", pval.Row - 1, MDC_PS_Common.GetValue("SELECT U_SmalName FROM [@PS_PP003L] WHERE U_SmalCode = '" + oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value + "'", 0, 1));
        //                }
        //                else
        //                {
        //                    //UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040N.SetValue("U_" + pval.ColUID, pval.Row - 1, oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Specific.Value);
        //                }
        //            }
        //            else
        //            {
        //                if ((pval.ItemUID == "DocEntry"))
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040H.SetValue(pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Value);
        //                }
        //                else if ((pval.ItemUID == "BaseTime"))
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040H.SetValue("U_" + pval.ItemUID, 0, Convert.ToString(Conversion.Val(oForm01.Items.Item(pval.ItemUID).Specific.Value)));
        //                }
        //                else if ((pval.ItemUID == "OrdMgNum"))
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Value);
        //                    if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE | oForm01.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
        //                    {
        //                        PS_PP040_OrderInfoLoad();
        //                    }
        //                }
        //                else if ((pval.ItemUID == "ItemCode"))
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Value);
        //                    oMat01.Clear();
        //                    oMat01.FlushToDataSource();
        //                    oMat01.LoadFromDataSource();
        //                    PS_PP040_AddMatrixRow01(0, ref true);
        //                    oMat02.Clear();
        //                    oMat02.FlushToDataSource();
        //                    oMat02.LoadFromDataSource();
        //                    PS_PP040_AddMatrixRow02(0, ref true);
        //                    oMat03.Clear();
        //                    oMat03.FlushToDataSource();
        //                    oMat03.LoadFromDataSource();

        //                }
        //                else if ((pval.ItemUID == "UseMCode"))
        //                {

        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    Query01 = "EXEC PS_PP040_98 '" + oForm01.Items.Item("UseMCode").Specific.Value;

        //                    RecordSet01.DoQuery(Query01);

        //                    //UPGRADE_WARNING: oForm01.Items(UseMName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oForm01.Items.Item("UseMName").Specific.Value = Strings.Trim(RecordSet01.Fields.Item(0).Value);

        //                }
        //                else
        //                {
        //                    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oDS_PS_PP040H.SetValue("U_" + pval.ItemUID, 0, oForm01.Items.Item(pval.ItemUID).Specific.Value);
        //                }
        //            }
        //            oMat01.LoadFromDataSource();
        //            oMat01.AutoResizeColumns();
        //            oMat02.LoadFromDataSource();
        //            oMat02.AutoResizeColumns();
        //            oMat03.LoadFromDataSource();
        //            oMat03.AutoResizeColumns();
        //            oForm01.Update();
        //            if (pval.ItemUID == "Mat01")
        //            {
        //                oMat01.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //            }
        //            else if (pval.ItemUID == "Mat02")
        //            {
        //                oMat02.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //            }
        //            else if (pval.ItemUID == "Mat03")
        //            {
        //                oMat03.Columns.Item(pval.ColUID).Cells.Item(pval.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //            }
        //            else
        //            {
        //                oForm01.Items.Item(pval.ItemUID).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //            }
        //        }
        //    }
        //    else if (pval.BeforeAction == false)
        //    {

        //    }
        //    oForm01.Freeze(false);
        //    return;
        //Raise_EVENT_VALIDATE_Error:
        //    oForm01.Freeze(false);
        //    ProgBar01.Value = 100;
        //    ProgBar01.Stop();
        //    //UPGRADE_NOTE: ProgBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgBar01 = null;
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_VALIDATE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_MATRIX_LOAD
        //private void Raise_EVENT_MATRIX_LOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if (pval.BeforeAction == true)
        //    {

        //    }
        //    else if (pval.BeforeAction == false)
        //    {
        //        PS_PP040_FormItemEnabled();
        //        if (pval.ItemUID == "Mat01")
        //        {
        //            PS_PP040_AddMatrixRow01(oMat01.VisualRowCount);
        //            ////UDO방식
        //        }
        //        else if (pval.ItemUID == "Mat02")
        //        {
        //            PS_PP040_AddMatrixRow02(oMat02.VisualRowCount);
        //            ////UDO방식
        //        }
        //    }
        //    return;
        //Raise_EVENT_MATRIX_LOAD_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_MATRIX_LOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_CHOOSE_FROM_LIST
        //private void Raise_EVENT_CHOOSE_FROM_LIST(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    SAPbouiCOM.DataTable oDataTable01 = null;
        //    if (pval.BeforeAction == true)
        //    {

        //    }
        //    else if (pval.BeforeAction == false)
        //    {
        //        //        If (pval.ItemUID = "ItemCode") Then
        //        //            Dim oDataTable01 As SAPbouiCOM.DataTable
        //        //            Set oDataTable01 = pval.SelectedObjects
        //        //            oForm01.DataSources.UserDataSources("ItemCode").Value = oDataTable01.Columns(0).Cells(0).Value
        //        //            Set oDataTable01 = Nothing
        //        //        End If
        //        //        If (pval.ItemUID = "CardCode" Or pval.ItemUID = "CardName") Then
        //        //            Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP040H", "U_CardCode,U_CardName")
        //        //        End If
        //        if ((pval.ItemUID == "ItemCode"))
        //        {
        //            //UPGRADE_WARNING: pval.SelectedObjects 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            if (pval.SelectedObjects == null)
        //            {
        //            }
        //            else
        //            {
        //                MDC_Com.MDC_GP_CF_DBDatasourceReturn(pval, (pval.FormUID), "@PS_PP040H", "U_ItemCode,U_ItemName");
        //                oMat01.Clear();
        //                oMat01.FlushToDataSource();
        //                oMat01.LoadFromDataSource();
        //                PS_PP040_AddMatrixRow01(0, ref true);
        //                oMat02.Clear();
        //                oMat02.FlushToDataSource();
        //                oMat02.LoadFromDataSource();
        //                PS_PP040_AddMatrixRow02(0, ref true);
        //                oMat03.Clear();
        //                oMat03.FlushToDataSource();
        //                oMat03.LoadFromDataSource();
        //            }
        //        }
        //        //        If (pval.ItemUID = "Mat02") Then
        //        //            If (pval.ColUID = "WorkCode") Then
        //        //                If pval.SelectedObjects Is Nothing Then
        //        //                Else
        //        //                    Set oDataTable01 = pval.SelectedObjects
        //        //                    Call oDS_PS_PP040M.setValue("U_WorkCode", pval.Row - 1, oDataTable01.Columns("empID").Cells(0).Value)
        //        //                    Call oDS_PS_PP040M.setValue("U_WorkName", pval.Row - 1, oDataTable01.Columns("firstName").Cells(0).Value & oDataTable01.Columns("lastName").Cells(0).Value)
        //        //                    If oMat02.RowCount = pval.Row And Trim(oDS_PS_PP040M.GetValue("U_" & pval.ColUID, pval.Row - 1)) <> "" Then
        //        //                        Call PS_PP040_AddMatrixRow02(pval.Row)
        //        //                    End If
        //        //                    Set oDataTable01 = Nothing
        //        //                    'Call MDC_GP_CF_DBDatasourceReturn(pval, pval.FormUID, "@PS_PP030L", "U_CntcCode,U_CntcName")
        //        //                    oMat02.LoadFromDataSource
        //        //                    oMat02.Columns(pval.ColUID).Cells(pval.Row).Click ct_Regular
        //        //                End If
        //        //            End If
        //        //        End If
        //    }
        //    return;
        //Raise_EVENT_CHOOSE_FROM_LIST_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_CHOOSE_FROM_LIST_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_GOT_FOCUS
        //private void Raise_EVENT_GOT_FOCUS(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if (pval.ItemUID == "Mat01" | pval.ItemUID == "Mat02" | pval.ItemUID == "Mat03")
        //    {
        //        if (pval.Row > 0)
        //        {
        //            oLastItemUID01 = pval.ItemUID;
        //            oLastColUID01 = pval.ColUID;
        //            oLastColRow01 = pval.Row;
        //        }
        //    }
        //    else
        //    {
        //        oLastItemUID01 = pval.ItemUID;
        //        oLastColUID01 = "";
        //        oLastColRow01 = 0;
        //    }
        //    if (pval.ItemUID == "Mat01")
        //    {
        //        if (pval.Row > 0)
        //        {
        //            oMat01Row01 = pval.Row;
        //        }
        //    }
        //    else if (pval.ItemUID == "Mat02")
        //    {
        //        if (pval.Row > 0)
        //        {
        //            oMat02Row02 = pval.Row;
        //        }
        //    }
        //    else if (pval.ItemUID == "Mat03")
        //    {
        //        if (pval.Row > 0)
        //        {
        //            oMat03Row03 = pval.Row;
        //        }
        //    }
        //    return;
        //Raise_EVENT_GOT_FOCUS_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_GOT_FOCUS_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_FORM_UNLOAD
        //private void Raise_EVENT_FORM_UNLOAD(ref object FormUID, ref SAPbouiCOM.ItemEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    if (pval.BeforeAction == true)
        //    {
        //    }
        //    else if (pval.BeforeAction == false)
        //    {

        //    }
        //    return;
        //Raise_EVENT_FORM_UNLOAD_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_FORM_UNLOAD_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_ROW_DELETE
        //private void Raise_EVENT_ROW_DELETE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    object i = null;
        //    int j = 0;
        //    bool Exist = false;
        //    if ((oLastColRow01 > 0))
        //    {
        //        if (pval.BeforeAction == true)
        //        {
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            //분말 첫번째 공정일 경우 오류
        //            if (Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.Value) == "111" & (Strings.Trim(oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.Value) == "CP80111" | Strings.Trim(oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.Value) == "CP80101"))
        //            {
        //                MDC_Com.MDC_GF_Message(ref "첫공정은 행삭제 할수 없습니다.", ref "E");
        //                BubbleEvent = false;
        //                return;
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //분말 첫번째 공정일 경우 오류
        //            }
        //            else if (Strings.Trim(oForm01.Items.Item("OrdGbn").Specific.Value) == "601" & (Strings.Trim(oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.Value) == "CP80111" | Strings.Trim(oMat01.Columns.Item("CpCode").Cells.Item(oLastColRow01).Specific.Value) == "CP80101"))
        //            {
        //                MDC_Com.MDC_GF_Message(ref "첫공정은 행삭제 할수 없습니다.", ref "E");
        //                BubbleEvent = false;
        //                return;
        //            }
        //            //추가 End
        //            if (oLastItemUID01 == "Mat01")
        //            {
        //                if ((PS_PP040_Validate("행삭제01") == false))
        //                {
        //                    BubbleEvent = false;
        //                    return;
        //                }
        //            Continue_Renamed:
        //                for (i = 1; i <= oMat03.RowCount; i++)
        //                {
        //                    //UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    //UPGRADE_WARNING: oMat01.Columns(LineNum).Cells(oLastColRow01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    //UPGRADE_WARNING: oMat03.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    //UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(oLastColRow01).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (oMat01.Columns.Item("OrdMgNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value & oMat01.Columns.Item("LineNum").Cells.Item(oLastColRow01).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value)
        //                    {
        //                        ////If oMat01.Columns("OrdMgNum").Cells(oLastColRow01).Specific.Value = oMat03.Columns("OrdMgNum").Cells(i).Specific.Value Then
        //                        //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.RemoveRecord((i - 1));
        //                        oMat03.DeleteRow((i));
        //                        oMat03.FlushToDataSource();
        //                        goto Continue_Renamed;
        //                    }
        //                }
        //            }
        //            ////행삭제전 행삭제가능여부검사
        //        }
        //        else if (pval.BeforeAction == false)
        //        {
        //            if (oLastItemUID01 == "Mat01")
        //            {
        //                for (i = 1; i <= oMat01.VisualRowCount; i++)
        //                {
        //                    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //                }

        //                for (i = 1; i <= oMat03.VisualRowCount; i++)
        //                {
        //                    //UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    if (oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value != 1)
        //                    {
        //                        //UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        //UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value = oMat03.Columns.Item("OLineNum").Cells.Item(i).Specific.Value - 1;
        //                        ////i
        //                    }
        //                }

        //                oMat01.FlushToDataSource();
        //                oDS_PS_PP040L.RemoveRecord(oDS_PS_PP040L.Size - 1);
        //                oMat01.LoadFromDataSource();
        //                if (oMat01.RowCount == 0)
        //                {
        //                    PS_PP040_AddMatrixRow01(0);
        //                }
        //                else
        //                {
        //                    if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP040L.GetValue("U_OrdMgNum", oMat01.RowCount - 1))))
        //                    {
        //                        PS_PP040_AddMatrixRow01(oMat01.RowCount);
        //                    }
        //                }
        //            }
        //            else if (oLastItemUID01 == "Mat02")
        //            {
        //                for (i = 1; i <= oMat02.VisualRowCount; i++)
        //                {
        //                    //UPGRADE_WARNING: oMat02.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oMat02.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //                }
        //                oMat02.FlushToDataSource();
        //                oDS_PS_PP040M.RemoveRecord(oDS_PS_PP040M.Size - 1);
        //                oMat02.LoadFromDataSource();
        //                if (oMat02.RowCount == 0)
        //                {
        //                    PS_PP040_AddMatrixRow02(0);
        //                }
        //                else
        //                {
        //                    if (!string.IsNullOrEmpty(Strings.Trim(oDS_PS_PP040M.GetValue("U_WorkCode", oMat02.RowCount - 1))))
        //                    {
        //                        PS_PP040_AddMatrixRow02(oMat02.RowCount);
        //                    }
        //                }
        //            }
        //            else if (oLastItemUID01 == "Mat03")
        //            {
        //                for (i = 1; i <= oMat03.VisualRowCount; i++)
        //                {
        //                    //UPGRADE_WARNING: oMat03.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                    oMat03.Columns.Item("LineNum").Cells.Item(i).Specific.Value = i;
        //                }
        //                oMat03.FlushToDataSource();
        //                ////사이즈가 0일때는 행을 빼주면 oMat03.VisualRowCount 가 0 으로 변경되어서 문제가 생김
        //                if (oDS_PS_PP040N.Size == 1)
        //                {
        //                }
        //                else
        //                {
        //                    oDS_PS_PP040N.RemoveRecord(oDS_PS_PP040N.Size - 1);
        //                }
        //                oMat03.LoadFromDataSource();

        //                ////공정 테이블에는 있는데 불량 테이블에 존재하지 않는값이 있는경우 불량테이블에 값을 추가함
        //                for (i = 1; i <= oMat01.RowCount - 1; i++)
        //                {
        //                    Exist = false;
        //                    for (j = 1; j <= oMat03.RowCount; j++)
        //                    {
        //                        //UPGRADE_WARNING: oMat03.Columns(OLineNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        //UPGRADE_WARNING: oMat01.Columns(LineNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        //UPGRADE_WARNING: oMat03.Columns(OrdMgNum).Cells(j).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        //UPGRADE_WARNING: oMat01.Columns(OrdMgNum).Cells(i).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        if (oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OrdMgNum").Cells.Item(j).Specific.Value & oMat01.Columns.Item("LineNum").Cells.Item(i).Specific.Value == oMat03.Columns.Item("OLineNum").Cells.Item(j).Specific.Value)
        //                        {
        //                            ////If oMat01.Columns("OrdMgNum").Cells(i).Specific.Value = oMat03.Columns("OrdMgNum").Cells(j).Specific.Value Then
        //                            Exist = true;
        //                        }
        //                    }
        //                    ////불량코드테이블에 값이 존재하지 않으면
        //                    if (Exist == false)
        //                    {
        //                        if (oMat03.VisualRowCount == 0)
        //                        {
        //                            PS_PP040_AddMatrixRow03(0, ref true);
        //                        }
        //                        else
        //                        {
        //                            PS_PP040_AddMatrixRow03(oMat03.VisualRowCount);
        //                        }
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_OrdMgNum", oMat03.VisualRowCount - 1, oMat01.Columns.Item("OrdMgNum").Cells.Item(i).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_CpCode", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpCode").Cells.Item(i).Specific.Value);
        //                        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_CpName", oMat03.VisualRowCount - 1, oMat01.Columns.Item("CpName").Cells.Item(i).Specific.Value);
        //                        //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                        oDS_PS_PP040N.SetValue("U_OLineNum", oMat03.VisualRowCount - 1, i);
        //                        oMat03.LoadFromDataSource();
        //                        oMat03.AutoResizeColumns();
        //                        oMat03.Columns.Item("OLineNum").TitleObject.Sortable = true;
        //                        oMat03.Columns.Item("OLineNum").TitleObject.Sort(SAPbouiCOM.BoGridSortType.gst_Ascending);
        //                        //                        oMat03.Columns("OrdMgNum").TitleObject.Sortable = True
        //                        //                        Call oMat03.Columns("OrdMgNum").TitleObject.Sort(gst_Ascending)
        //                        oMat03.FlushToDataSource();
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    return;
        //Raise_EVENT_ROW_DELETE_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_ROW_DELETE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region Raise_EVENT_RECORD_MOVE
        //private void Raise_EVENT_RECORD_MOVE(ref string FormUID, ref SAPbouiCOM.IMenuEvent pval, ref bool BubbleEvent)
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    string Query01 = null;
        //    SAPbobsCOM.Recordset RecordSet01 = null;
        //    string DocEntry = null;
        //    string DocEntryNext = null;
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    DocEntry = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.Value);
        //    ////원본문서
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    DocEntryNext = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.Value);
        //    ////다음문서

        //    ////다음
        //    if (pval.MenuUID == "1288")
        //    {
        //        if (pval.BeforeAction == true)
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //            {
        //                SubMain.Sbo_Application.ActivateMenuItem(("1290"));
        //                BubbleEvent = false;
        //                return;
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
        //            {
        //                //UPGRADE_WARNING: oForm01.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                if ((string.IsNullOrEmpty(oForm01.Items.Item("DocEntry").Specific.Value)))
        //                {
        //                    SubMain.Sbo_Application.ActivateMenuItem(("1290"));
        //                    BubbleEvent = false;
        //                    return;
        //                }
        //            }
        //            if (PS_PP040_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") == false)
        //            {
        //                BubbleEvent = false;
        //                return;
        //            }
        //        }
        //        ////이전
        //    }
        //    else if (pval.MenuUID == "1289")
        //    {
        //        if (pval.BeforeAction == true)
        //        {
        //            if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
        //            {
        //                SubMain.Sbo_Application.ActivateMenuItem(("1291"));
        //                BubbleEvent = false;
        //                return;
        //            }
        //            else if (oForm01.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
        //            {
        //                //UPGRADE_WARNING: oForm01.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                if ((string.IsNullOrEmpty(oForm01.Items.Item("DocEntry").Specific.Value)))
        //                {
        //                    SubMain.Sbo_Application.ActivateMenuItem(("1291"));
        //                    BubbleEvent = false;
        //                    return;
        //                }
        //            }
        //            if (PS_PP040_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") == false)
        //            {
        //                BubbleEvent = false;
        //                return;
        //            }
        //        }
        //        ////첫번째레코드로이동
        //    }
        //    else if (pval.MenuUID == "1290")
        //    {
        //        if (pval.BeforeAction == true)
        //        {
        //            RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //            Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry DESC";
        //            ////가장마지막행을 부여
        //            RecordSet01.DoQuery(Query01);
        //            DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //            ////원본문서
        //            DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //            ////다음문서
        //            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //            RecordSet01 = null;
        //            if (PS_PP040_DirectionValidateDocument(DocEntry, DocEntryNext, "Next", "@PS_PP040H") == false)
        //            {
        //                BubbleEvent = false;
        //                return;
        //            }
        //        }
        //        ////마지막문서로이동
        //    }
        //    else if (pval.MenuUID == "1291")
        //    {
        //        if (pval.BeforeAction == true)
        //        {
        //            RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //            Query01 = " SELECT TOP 1 DocEntry FROM [@PS_PP040H] ORDER BY DocEntry ASC";
        //            ////가장 첫행을 부여
        //            RecordSet01.DoQuery(Query01);
        //            DocEntry = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //            ////원본문서
        //            DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value));
        //            ////다음문서
        //            //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //            RecordSet01 = null;
        //            if (PS_PP040_DirectionValidateDocument(DocEntry, DocEntryNext, "Prev", "@PS_PP040H") == false)
        //            {
        //                BubbleEvent = false;
        //                return;
        //            }
        //        }
        //    }
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    return;
        //Raise_EVENT_RECORD_MOVE_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("Raise_EVENT_RECORD_MOVE_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion
















        #region PS_PP040_DI_API
        //private bool PS_PP040_DI_API()
        //{
        //    bool functionReturnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    functionReturnValue = true;
        //    object i = null;
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
        //    for (i = 1; i <= oMat01.VisualRowCount; i++)
        //    {
        //        Array.Resize(ref ItemInformation, ItemInformationCount + 1);
        //        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        ItemInformation[ItemInformationCount].ItemCode = oMat01.Columns.Item("ItemCode").Cells.Item(i).Specific.Value;
        //        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        ItemInformation[ItemInformationCount].BatchNum = oMat01.Columns.Item("BatchNum").Cells.Item(i).Specific.Value;
        //        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        ItemInformation[ItemInformationCount].Quantity = oMat01.Columns.Item("Quantity").Cells.Item(i).Specific.Value;
        //        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        ItemInformation[ItemInformationCount].OPORNo = oMat01.Columns.Item("OPORNo").Cells.Item(i).Specific.Value;
        //        //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        ItemInformation[ItemInformationCount].POR1No = oMat01.Columns.Item("POR1No").Cells.Item(i).Specific.Value;
        //        ItemInformation[ItemInformationCount].Check = false;
        //        ItemInformationCount = ItemInformationCount + 1;
        //    }

        //    LineNumCount = 0;
        //    oDIObject = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes);
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    oDIObject.BPL_IDAssignedToInvoice = Convert.ToInt32(Strings.Trim(oForm01.Items.Item("BPLId").Specific.Selected.Value));
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    oDIObject.CardCode = Strings.Trim(oForm01.Items.Item("CardCode").Specific.Value);
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    oDIObject.DocDate = Convert.ToDateTime(Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oForm01.Items.Item("InDate").Specific.Value, "&&&&-&&-&&"));
        //    for (i = 0; i <= ItemInformationCount - 1; i++)
        //    {
        //        //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        if (ItemInformation[i].Check == true)
        //        {
        //            goto Continue_First;
        //        }
        //        //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        if (i != 0)
        //        {
        //            oDIObject.Lines.Add();
        //        }
        //        //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oDIObject.Lines.ItemCode = ItemInformation[i].ItemCode;
        //        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oDIObject.Lines.WarehouseCode = Strings.Trim(oForm01.Items.Item("WhsCode").Specific.Value);
        //        oDIObject.Lines.BaseType = Convert.ToInt32("22");
        //        //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oDIObject.Lines.BaseEntry = ItemInformation[i].OPORNo;
        //        //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        oDIObject.Lines.BaseLine = ItemInformation[i].POR1No;
        //        //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        for (j = i; j <= Information.UBound(ItemInformation); j++)
        //        {
        //            if (ItemInformation[j].Check == true)
        //            {
        //                goto Continue_Second;
        //            }
        //            //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            if ((ItemInformation[i].ItemCode != ItemInformation[j].ItemCode | ItemInformation[i].OPORNo != ItemInformation[j].OPORNo | ItemInformation[i].POR1No != ItemInformation[j].POR1No))
        //            {
        //                goto Continue_Second;
        //            }
        //            ////같은것
        //            oDIObject.Lines.Quantity = oDIObject.Lines.Quantity + ItemInformation[j].Quantity;
        //            oDIObject.Lines.BatchNumbers.BatchNumber = ItemInformation[j].BatchNum;
        //            oDIObject.Lines.BatchNumbers.Quantity = ItemInformation[j].Quantity;
        //            oDIObject.Lines.BatchNumbers.Add();
        //            ItemInformation[j].PDN1No = LineNumCount;
        //            ItemInformation[j].Check = true;
        //        Continue_Second:

        //        }
        //        LineNumCount = LineNumCount + 1;
        //    Continue_First:

        //    }
        //    RetVal = oDIObject.Add();
        //    if (RetVal == 0)
        //    {
        //        ResultDocNum = Convert.ToInt32(SubMain.Sbo_Company.GetNewObjectKey());
        //        for (i = 0; i <= Information.UBound(ItemInformation); i++)
        //        {
        //            //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            oDS_PS_PP040L.SetValue("U_OPDNNo", i, Convert.ToString(ResultDocNum));
        //            //UPGRADE_WARNING: i 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            oDS_PS_PP040L.SetValue("U_PDN1No", i, Convert.ToString(ItemInformation[i].PDN1No));
        //        }
        //    }
        //    else
        //    {
        //        goto PS_PP040_DI_API_Error;
        //    }

        //    if (SubMain.Sbo_Company.InTransaction == true)
        //    {
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //    }
        //    oMat01.LoadFromDataSource();
        //    oMat01.AutoResizeColumns();

        //    //UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oDIObject = null;
        //    return functionReturnValue;
        //PS_PP040_DI_API_DI_Error:
        //    if (SubMain.Sbo_Company.InTransaction == true)
        //    {
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //    }
        //    SubMain.Sbo_Application.SetStatusBarMessage(SubMain.Sbo_Company.GetLastErrorCode() + " - " + SubMain.Sbo_Company.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //    functionReturnValue = false;
        //    //UPGRADE_NOTE: oDIObject 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oDIObject = null;
        //    return functionReturnValue;
        //PS_PP040_DI_API_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage("PS_PP040_DI_API_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //    functionReturnValue = false;
        //    return functionReturnValue;
        //}
        #endregion


        #region PS_PP040_OrderInfoLoad
        //private void PS_PP040_OrderInfoLoad()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    string Query01 = null;
        //    SAPbobsCOM.Recordset RecordSet01 = null;
        //    RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    ////일반,조정, 설계
        //    if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "10" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "50" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "60" | oForm01.Items.Item("OrdType").Specific.Selected.Value == "70")
        //    {
        //        //UPGRADE_WARNING: oForm01.Items(OrdMgNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        if (string.IsNullOrEmpty(oForm01.Items.Item("OrdMgNum").Specific.Value))
        //        {
        //            MDC_Com.MDC_GF_Message(ref "작업지시 관리번호를 입력하지 않습니다.", ref "W");
        //            goto PS_PP040_OrderInfoLoad_Exit;
        //        }
        //        else
        //        {
        //            Query01 = "SELECT ";
        //            Query01 = Query01 + "U_OrdGbn,";
        //            Query01 = Query01 + "U_BPLId,";
        //            Query01 = Query01 + "U_ItemCode,";
        //            Query01 = Query01 + "U_ItemName,";
        //            Query01 = Query01 + "U_OrdNum,";
        //            Query01 = Query01 + "U_OrdSub1,";
        //            Query01 = Query01 + "U_OrdSub2,";
        //            Query01 = Query01 + "DocEntry";
        //            Query01 = Query01 + " FROM [@PS_PP030H]";
        //            Query01 = Query01 + " WHERE ";
        //            //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            Query01 = Query01 + " U_OrdNum + U_OrdSub1 + U_OrdSub2 = '" + oForm01.Items.Item("OrdMgNum").Specific.Value + "'";
        //            Query01 = Query01 + " AND U_OrdGbn NOT IN('104','107') ";
        //            Query01 = Query01 + " AND Canceled = 'N'";
        //            RecordSet01.DoQuery(Query01);
        //            if (RecordSet01.RecordCount == 0)
        //            {
        //                MDC_Com.MDC_GF_Message(ref "작업지시 정보가 존재하지 않습니다.", ref "W");
        //                goto PS_PP040_OrderInfoLoad_Exit;
        //            }
        //            else
        //            {
        //                //UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("OrdGbn").Specific.Select(RecordSet01.Fields.Item(0).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
        //                //UPGRADE_WARNING: oForm01.Items().Specific.Select 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("BPLId").Specific.Select(RecordSet01.Fields.Item(1).Value, SAPbouiCOM.BoSearchKey.psk_ByValue);
        //                //UPGRADE_WARNING: oForm01.Items(ItemCode).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("ItemCode").Specific.Value = RecordSet01.Fields.Item(2).Value;
        //                //UPGRADE_WARNING: oForm01.Items(ItemName).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("ItemName").Specific.Value = RecordSet01.Fields.Item(3).Value;
        //                //UPGRADE_WARNING: oForm01.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("OrdNum").Specific.Value = RecordSet01.Fields.Item(4).Value;
        //                //UPGRADE_WARNING: oForm01.Items(OrdSub1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("OrdSub1").Specific.Value = RecordSet01.Fields.Item(5).Value;
        //                //UPGRADE_WARNING: oForm01.Items(OrdSub2).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("OrdSub2").Specific.Value = RecordSet01.Fields.Item(6).Value;
        //                //UPGRADE_WARNING: oForm01.Items(PP030HNo).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                //UPGRADE_WARNING: RecordSet01.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("PP030HNo").Specific.Value = RecordSet01.Fields.Item(7).Value;
        //                //                '//매트릭스삭제
        //                //                oMat01.Clear
        //                //                oMat01.FlushToDataSource
        //                //                oMat01.LoadFromDataSource
        //                //                Call PS_PP040_AddMatrixRow01(0, True)
        //                //                oMat02.Clear
        //                //                oMat02.FlushToDataSource
        //                //                oMat02.LoadFromDataSource
        //                //                Call PS_PP040_AddMatrixRow02(0, True)
        //                //                oMat03.Clear
        //                //                oMat03.FlushToDataSource
        //                //                oMat03.LoadFromDataSource
        //                oForm01.Update();
        //            }
        //        }
        //        //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        ////PSMT
        //    }
        //    else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "20")
        //    {
        //        //UPGRADE_WARNING: oForm01.Items(OrdMgNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        if (string.IsNullOrEmpty(oForm01.Items.Item("OrdMgNum").Specific.Value))
        //        {
        //            MDC_Com.MDC_GF_Message(ref "작업지시 관리번호를 입력하지 않습니다.", ref "W");
        //            goto PS_PP040_OrderInfoLoad_Exit;
        //        }
        //        else
        //        {
        //            //UPGRADE_WARNING: oForm01.Items(OrdNum).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            oForm01.Items.Item("OrdNum").Specific.Value = oForm01.Items.Item("OrdMgNum").Specific.Value;
        //            //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            oForm01.Items.Item("OrdSub1").Specific.Value = "000";
        //            //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            oForm01.Items.Item("OrdSub2").Specific.Value = "00";
        //            ////매트릭스삭제
        //            oMat01.Clear();
        //            oMat01.FlushToDataSource();
        //            oMat01.LoadFromDataSource();
        //            PS_PP040_AddMatrixRow01(0, ref true);
        //            oMat02.Clear();
        //            oMat02.FlushToDataSource();
        //            oMat02.LoadFromDataSource();
        //            PS_PP040_AddMatrixRow02(0, ref true);
        //            oMat03.Clear();
        //            oMat03.FlushToDataSource();
        //            oMat03.LoadFromDataSource();
        //            oForm01.Update();
        //        }
        //        //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    }
        //    else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "30")
        //    {
        //        MDC_Com.MDC_GF_Message(ref "외주은 입력할수 없습니다.", ref "W");
        //        goto PS_PP040_OrderInfoLoad_Exit;
        //        //UPGRADE_WARNING: oForm01.Items(OrdType).Specific.Selected 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    }
        //    else if (oForm01.Items.Item("OrdType").Specific.Selected.Value == "40")
        //    {
        //        MDC_Com.MDC_GF_Message(ref "실적은 입력할수 없습니다.", ref "W");
        //        goto PS_PP040_OrderInfoLoad_Exit;
        //    }
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    return;
        //PS_PP040_OrderInfoLoad_Exit:
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    return;
        //PS_PP040_OrderInfoLoad_Error:
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    SubMain.Sbo_Application.SetStatusBarMessage("PS_PP040_OrderInfoLoad_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

        #region PS_PP040_FindValidateDocument
        //public bool PS_PP040_FindValidateDocument(string ObjectType)
        //{
        //    bool functionReturnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    functionReturnValue = true;
        //    string Query01 = null;
        //    SAPbobsCOM.Recordset RecordSet01 = null;
        //    string Query02 = null;
        //    SAPbobsCOM.Recordset RecordSet02 = null;

        //    int i = 0;
        //    string DocEntry = null;
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    DocEntry = Strings.Trim(oForm01.Items.Item("DocEntry").Specific.Value);
        //    ////원본문서

        //    RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    Query01 = " SELECT DocEntry";
        //    Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry = ";
        //    Query01 = Query01 + DocEntry;
        //    if ((oDocType01 == "작업일보등록(작지)"))
        //    {
        //        Query01 = Query01 + " AND U_DocType = '10'";
        //    }
        //    else if ((oDocType01 == "작업일보등록(공정)"))
        //    {
        //        Query01 = Query01 + " AND U_DocType = '20'";
        //    }
        //    RecordSet01.DoQuery(Query01);
        //    if ((RecordSet01.RecordCount == 0))
        //    {
        //        if ((oDocType01 == "작업일보등록(작지)"))
        //        {
        //            SubMain.Sbo_Application.SetStatusBarMessage("작업일보등록(공정)문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //        }
        //        else if ((oDocType01 == "작업일보등록(공정)"))
        //        {
        //            SubMain.Sbo_Application.SetStatusBarMessage("작업일보등록(작지)문서 이거나 존재하지 않는 문서입니다.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //        }
        //        functionReturnValue = false;
        //        //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        RecordSet01 = null;
        //        //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        RecordSet02 = null;
        //        return functionReturnValue;
        //    }

        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet02 = null;
        //    return functionReturnValue;
        //PS_PP040_FindValidateDocument_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet02 = null;
        //    functionReturnValue = false;
        //    return functionReturnValue;
        //}
        #endregion

        #region PS_PP040_DirectionValidateDocument
        //public bool PS_PP040_DirectionValidateDocument(string DocEntry, string DocEntryNext, string Direction, string ObjectType)
        //{
        //    bool functionReturnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    string Query01 = null;
        //    SAPbobsCOM.Recordset RecordSet01 = null;
        //    string Query02 = null;
        //    SAPbobsCOM.Recordset RecordSet02 = null;

        //    int i = 0;
        //    string MaxDocEntry = null;
        //    string MinDocEntry = null;
        //    bool DoNext = false;
        //    bool IsFirst = false;
        //    ////시작유무
        //    DoNext = true;
        //    IsFirst = true;

        //    RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    RecordSet02 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    while ((DoNext == true))
        //    {
        //        if ((IsFirst != true))
        //        {
        //            ////문서전체를 경유하고도 유효값을 찾지못했다면
        //            if ((DocEntry == DocEntryNext))
        //            {
        //                SubMain.Sbo_Application.SetStatusBarMessage("유효한문서가 존재하지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //                functionReturnValue = false;
        //                //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                RecordSet01 = null;
        //                //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                RecordSet02 = null;
        //                return functionReturnValue;
        //            }
        //        }
        //        if ((Direction == "Next"))
        //        {
        //            Query01 = " SELECT TOP 1 DocEntry";
        //            Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry > ";
        //            Query01 = Query01 + DocEntryNext;
        //            if ((oDocType01 == "작업일보등록(작지)"))
        //            {
        //                Query01 = Query01 + " AND U_DocType = '10'";
        //            }
        //            else if ((oDocType01 == "작업일보등록(공정)"))
        //            {
        //                Query01 = Query01 + " AND U_DocType = '20'";
        //            }
        //            Query01 = Query01 + " ORDER BY DocEntry ASC";
        //        }
        //        else if ((Direction == "Prev"))
        //        {
        //            Query01 = " SELECT TOP 1 DocEntry";
        //            Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
        //            Query01 = Query01 + DocEntryNext;
        //            if ((oDocType01 == "작업일보등록(작지)"))
        //            {
        //                Query01 = Query01 + " AND U_DocType = '10'";
        //            }
        //            else if ((oDocType01 == "작업일보등록(공정)"))
        //            {
        //                Query01 = Query01 + " AND U_DocType = '20'";
        //            }
        //            Query01 = Query01 + " ORDER BY DocEntry DESC";
        //        }
        //        RecordSet01.DoQuery(Query01);
        //        ////해당문서가 마지막문서라면
        //        if ((RecordSet01.Fields.Item(0).Value == 0))
        //        {
        //            if ((Direction == "Next"))
        //            {
        //                Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
        //                if ((oDocType01 == "작업일보등록(작지)"))
        //                {
        //                    Query02 = Query02 + " WHERE U_DocType = '10'";
        //                }
        //                else if ((oDocType01 == "작업일보등록(공정)"))
        //                {
        //                    Query02 = Query02 + " WHERE U_DocType = '20'";
        //                }
        //                Query02 = Query02 + " ORDER BY DocEntry ASC";
        //            }
        //            else if ((Direction == "Prev"))
        //            {
        //                Query02 = " SELECT TOP 1 DocEntry FROM [" + ObjectType + "]";
        //                if ((oDocType01 == "작업일보등록(작지)"))
        //                {
        //                    Query02 = Query02 + " WHERE U_DocType = '10'";
        //                }
        //                else if ((oDocType01 == "작업일보등록(공정)"))
        //                {
        //                    Query02 = Query02 + " WHERE U_DocType = '20'";
        //                }
        //                Query02 = Query02 + " ORDER BY DocEntry DESC";
        //            }
        //            RecordSet02.DoQuery(Query02);
        //            ////문서가 아예 존재하지 않는다면
        //            if ((RecordSet02.RecordCount == 0))
        //            {
        //                SubMain.Sbo_Application.SetStatusBarMessage("유효한문서가 존재하지 않습니다", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //                //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                RecordSet01 = null;
        //                //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //                RecordSet02 = null;
        //                functionReturnValue = false;
        //                return functionReturnValue;
        //            }
        //            else
        //            {
        //                if ((Direction == "Next"))
        //                {
        //                    DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) - 1);
        //                    Query01 = " SELECT TOP 1 DocEntry";
        //                    Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry > ";
        //                    Query01 = Query01 + DocEntryNext;
        //                    if ((oDocType01 == "작업일보등록(작지)"))
        //                    {
        //                        Query01 = Query01 + " AND U_DocType = '10'";
        //                    }
        //                    else if ((oDocType01 == "작업일보등록(공정)"))
        //                    {
        //                        Query01 = Query01 + " AND U_DocType = '20'";
        //                    }
        //                    Query01 = Query01 + " ORDER BY DocEntry ASC";
        //                    RecordSet01.DoQuery(Query01);
        //                }
        //                else if ((Direction == "Prev"))
        //                {
        //                    DocEntryNext = Convert.ToString(Conversion.Val(RecordSet02.Fields.Item(0).Value) + 1);
        //                    Query01 = " SELECT TOP 1 DocNum";
        //                    Query01 = Query01 + " FROM [" + ObjectType + "] Where DocEntry < ";
        //                    Query01 = Query01 + DocEntryNext;
        //                    if ((oDocType01 == "작업일보등록(작지)"))
        //                    {
        //                        Query01 = Query01 + " AND U_DocType = '10'";
        //                    }
        //                    else if ((oDocType01 == "작업일보등록(공정)"))
        //                    {
        //                        Query01 = Query01 + " AND U_DocType = '20'";
        //                    }
        //                    Query01 = Query01 + " ORDER BY DocEntry DESC";
        //                    RecordSet01.DoQuery(Query01);
        //                }
        //            }
        //        }
        //        if ((oDocType01 == "작업일보등록(작지)"))
        //        {
        //            DoNext = false;
        //            if ((Direction == "Next"))
        //            {
        //                DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
        //            }
        //            else if ((Direction == "Prev"))
        //            {
        //                DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
        //            }
        //        }
        //        else if ((oDocType01 == "작업일보등록(공정)"))
        //        {
        //            DoNext = false;
        //            if ((Direction == "Next"))
        //            {
        //                DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) - 1);
        //            }
        //            else if ((Direction == "Prev"))
        //            {
        //                DocEntryNext = Convert.ToString(Conversion.Val(RecordSet01.Fields.Item(0).Value) + 1);
        //            }
        //        }
        //        IsFirst = false;
        //    }
        //    ////다음문서가 유효하다면 그냥 넘어가고
        //    if ((DocEntry == DocEntryNext))
        //    {
        //        PS_PP040_FormItemEnabled();
        //        ////UDO방식
        //        ////다음문서가 유효하지 않다면
        //    }
        //    else
        //    {
        //        oForm01.Freeze(true);
        //        oForm01.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
        //        PS_PP040_FormItemEnabled();
        //        ////UDO방식
        //        ////문서번호 필드가 입력이 가능하다면
        //        if (oForm01.Items.Item("DocEntry").Enabled == true)
        //        {
        //            if ((Direction == "Next"))
        //            {
        //                //UPGRADE_WARNING: oForm01.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("DocEntry").Specific.Value = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) + 1));
        //            }
        //            else if ((Direction == "Prev"))
        //            {
        //                //UPGRADE_WARNING: oForm01.Items(DocEntry).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                oForm01.Items.Item("DocEntry").Specific.Value = Conversion.Val(Convert.ToString(Convert.ToDouble(DocEntryNext) - 1));
        //            }
        //            oForm01.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
        //        }
        //        oForm01.Freeze(false);
        //        functionReturnValue = false;
        //        //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        RecordSet01 = null;
        //        //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        RecordSet02 = null;
        //        return functionReturnValue;
        //    }
        //    functionReturnValue = true;
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet02 = null;
        //    return functionReturnValue;
        //PS_PP040_DirectionValidateDocument_Error:
        //    SubMain.Sbo_Application.SetStatusBarMessage(Err().Number + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //    functionReturnValue = false;
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    //UPGRADE_NOTE: RecordSet02 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet02 = null;
        //    return functionReturnValue;
        //}
        #endregion

        #region Add_oInventoryGenExit
        //private bool Add_oInventoryGenExit(ref short ChkType)
        //{
        //    bool functionReturnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    SAPbobsCOM.Documents DI_oInventoryGenExit = null;
        //    ////재고출고 문서 객체

        //    int j = 0;
        //    int i = 0;
        //    int Cnt = 0;
        //    short ErrNum = 0;
        //    int errCode = 0;
        //    string ErrMsg = null;
        //    int RetVal = 0;

        //    string CpCode = null;
        //    string DocNum = null;
        //    string DocDate = null;
        //    string CItemCod = null;
        //    string WhsCode = null;
        //    int IssueQty = 0;
        //    decimal IssueWt = default(decimal);
        //    string SDocEntry = null;
        //    SAPbobsCOM.Recordset oRecordSet = null;
        //    string sQry = null;
        //    short Price = 0;

        //    Cnt = 0;

        //    oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    oMat01.FlushToDataSource();

        //    DocDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDS_PS_PP040H.GetValue("U_DocDate", 0), "0000-00-00");
        //    //[If ChkType = 2 Then Call FormClear
        //    DocNum = Strings.Trim(oDS_PS_PP040H.GetValue("DocEntry", 0));


        //    //UPGRADE_WARNING: oMat01.Columns(OutDoc).Cells(1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    if (string.IsNullOrEmpty(oMat01.Columns.Item("OutDoc").Cells.Item(1).Specific.Value))
        //    {
        //        SubMain.Sbo_Company.StartTransaction();
        //        //UPGRADE_NOTE: DI_oInventoryGenExit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        DI_oInventoryGenExit = null;
        //        DI_oInventoryGenExit = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

        //        var _with1 = DI_oInventoryGenExit;
        //        _with1.DocDate = Convert.ToDateTime(DocDate);
        //        _with1.TaxDate = Convert.ToDateTime(DocDate);
        //        _with1.Comments = "원재료 불출 등록(" + DocNum + ") 출고";

        //        j = 0;
        //        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //        {

        //            sQry = " SELECT PRICE";
        //            sQry = sQry + "  FROM OIVL a inner join OIGN b on a.BASE_REF = b.DocEntry and b.U_Comments ='Convert Meterial'";
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            sQry = Convert.ToString(Convert.ToDouble(sQry + "  WHERE a.ITEMCODE ='") + oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value + Convert.ToDouble("'"));
        //            sQry = sQry + "  and convert(char(6),a.DocDate,112) ='" + Strings.Left(oDS_PS_PP040H.GetValue("U_DocDate", 0), 6) + "'";

        //            oRecordSet.DoQuery((sQry));

        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value;
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            IssueQty = oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value;
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            IssueWt = oMat01.Columns.Item("PWeight").Cells.Item(i + 1).Specific.Value;
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.Value;
        //            //UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            Price = oRecordSet.Fields.Item(0).Value;

        //            WhsCode = "101";

        //            if ((CpCode == "CP80101" | CpCode == "CP80111") & !string.IsNullOrEmpty(CItemCod) & IssueQty >= 0 & IssueWt != 0 & !string.IsNullOrEmpty(WhsCode))
        //            {
        //                //If (CpCode = "CP80101" Or CpCode = "CP80111" Or CpCode = "CP80104" Or CpCode = "CP80105") And CItemCod <> "" And IssueQty >= 0 And IssueWt <> 0 And WhsCode <> "" Then
        //                if (j > 0)
        //                    _with1.Lines.Add();
        //                _with1.Lines.SetCurrentLine(j);
        //                _with1.Lines.ItemCode = CItemCod;
        //                _with1.Lines.WarehouseCode = WhsCode;
        //                _with1.Lines.Quantity = IssueWt;
        //                _with1.Lines.UserFields.Fields.Item("U_Qty").Value = IssueQty;
        //                //제품원재료 변환 품목은 단가를 계산 후 입력
        //                if ((oRecordSet.EoF))
        //                {
        //                }
        //                else
        //                {
        //                    _with1.Lines.Price = Price;
        //                    _with1.Lines.UnitPrice = Price;
        //                    _with1.Lines.LineTotal = Price * IssueWt;

        //                }

        //                Cnt = Cnt + 1;
        //                j = j + 1;
        //            }
        //        }

        //        //// 완료
        //        if (Cnt > 0)
        //        {
        //            RetVal = DI_oInventoryGenExit.Add();
        //            if ((0 != RetVal))
        //            {
        //                SubMain.Sbo_Company.GetLastError(out errCode, out ErrMsg);
        //                ErrNum = 1;
        //                goto Add_oInventoryGenExit_Error;
        //            }
        //        }

        //        if (ChkType == 1)
        //        {
        //            SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }
        //        else if (ChkType == 2)
        //        {
        //            SubMain.Sbo_Company.GetNewObjectCode(out SDocEntry);
        //            Cnt = 1;
        //            for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //            {
        //                //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //                CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.Value;
        //                if (CpCode == "CP80101" | CpCode == "CP80111")
        //                {
        //                    //If CpCode = "CP80101" Or CpCode = "CP80111" Or CpCode = "CP80104" Or CpCode = "CP80105" Then
        //                    oDS_PS_PP040L.SetValue("U_OutDoc", i, SDocEntry);
        //                    oDS_PS_PP040L.SetValue("U_OutLin", i, Convert.ToString(Cnt));
        //                    Cnt = Cnt + 1;
        //                }
        //            }
        //            oMat01.LoadFromDataSource();
        //            SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //        }
        //    }

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    //UPGRADE_NOTE: DI_oInventoryGenExit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    DI_oInventoryGenExit = null;
        //    functionReturnValue = true;
        //    return functionReturnValue;
        //Add_oInventoryGenExit_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    //UPGRADE_NOTE: DI_oInventoryGenExit 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    DI_oInventoryGenExit = null;
        //    if (SubMain.Sbo_Company.InTransaction)
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //    functionReturnValue = false;
        //    if (ErrNum == 1)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "Add_oInventoryGenExit_Error:" + errCode + " - " + ErrMsg, ref "E");
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "Add_oInventoryGenExit_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //    return functionReturnValue;
        //}
        #endregion

        #region Add_oInventoryGenEntry
        //private bool Add_oInventoryGenEntry(ref short ChkType)
        //{
        //    bool functionReturnValue = false;
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    SAPbobsCOM.Documents DI_oInventoryGenEntry = null;
        //    ////재고출고 문서 객체

        //    int j = 0;
        //    int i = 0;
        //    int Cnt = 0;
        //    short ErrNum = 0;
        //    int errCode = 0;
        //    string ErrMsg = null;
        //    int RetVal = 0;

        //    string CpCode = null;
        //    string DocNum = null;
        //    string DocDate = null;
        //    string CItemCod = null;
        //    string WhsCode = null;
        //    int IssueQty = 0;
        //    decimal IssueWt = default(decimal);
        //    string SDocEntry = null;
        //    SAPbobsCOM.Recordset oRecordSet = null;
        //    string sQry = null;
        //    string OIGEDoc = null;
        //    short Price = 0;

        //    Cnt = 0;

        //    oRecordSet = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    oMat01.FlushToDataSource();

        //    DocDate = Microsoft.VisualBasic.Compatibility.VB6.Support.Format(oDS_PS_PP040H.GetValue("U_DocDate", 0), "0000-00-00");
        //    DocNum = Strings.Trim(oDS_PS_PP040H.GetValue("DocEntry", 0));
        //    //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    OIGEDoc = Strings.Trim(oMat01.Columns.Item("OutDoc").Cells.Item(1).Specific.Value);

        //    //UPGRADE_WARNING: oMat01.Columns(OutDocC).Cells(1).Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    if (string.IsNullOrEmpty(oMat01.Columns.Item("OutDocC").Cells.Item(1).Specific.Value))
        //    {
        //        SubMain.Sbo_Company.StartTransaction();
        //        //UPGRADE_NOTE: DI_oInventoryGenEntry 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //        DI_oInventoryGenEntry = null;
        //        DI_oInventoryGenEntry = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

        //        var _with2 = DI_oInventoryGenEntry;
        //        _with2.DocDate = Convert.ToDateTime(DocDate);
        //        _with2.TaxDate = Convert.ToDateTime(DocDate);
        //        _with2.Comments = "원재료 불출 등록 출고 취소 (" + DocNum + ") 입고";

        //        _with2.UserFields.Fields.Item("U_CancDoc").Value = OIGEDoc;
        //        ////입고취소 문서번호

        //        j = 0;
        //        for (i = 0; i <= oMat01.VisualRowCount - 1; i++)
        //        {

        //            sQry = " SELECT PRICE";
        //            sQry = sQry + "  FROM OIVL a inner join OIGN b on a.BASE_REF = b.DocEntry and b.U_Comments ='Convert Meterial'";
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            sQry = Convert.ToString(Convert.ToDouble(sQry + "  WHERE a.ITEMCODE ='") + oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value + Convert.ToDouble("'"));
        //            sQry = sQry + "  and convert(char(6),a.DocDate,112) ='" + Strings.Left(oDS_PS_PP040H.GetValue("U_DocDate", 0), 6) + "'";

        //            oRecordSet.DoQuery((sQry));

        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            CItemCod = oMat01.Columns.Item("CItemCod").Cells.Item(i + 1).Specific.Value;
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            IssueQty = oMat01.Columns.Item("PQty").Cells.Item(i + 1).Specific.Value;
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            IssueWt = oMat01.Columns.Item("PWeight").Cells.Item(i + 1).Specific.Value;
        //            //UPGRADE_WARNING: oMat01.Columns().Cells().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            CpCode = oMat01.Columns.Item("CpCode").Cells.Item(i + 1).Specific.Value;
        //            WhsCode = "101";
        //            //UPGRADE_WARNING: oRecordSet.Fields().Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //            Price = oRecordSet.Fields.Item(0).Value;

        //            if ((CpCode == "CP80101" | CpCode == "CP80111") & !string.IsNullOrEmpty(CItemCod) & IssueQty >= 0 & IssueWt != 0 & !string.IsNullOrEmpty(WhsCode))
        //            {
        //                //If (CpCode = "CP80101" Or CpCode = "CP80111" Or CpCode = "CP80104" Or CpCode = "CP80105") And CItemCod <> "" And IssueQty >= 0 And IssueWt <> 0 And WhsCode <> "" Then
        //                if (j > 0)
        //                    _with2.Lines.Add();
        //                _with2.Lines.SetCurrentLine(j);
        //                _with2.Lines.ItemCode = CItemCod;
        //                _with2.Lines.WarehouseCode = WhsCode;
        //                //            .Lines.AccountCode = Trim(sAccount)
        //                _with2.Lines.Quantity = IssueWt;
        //                _with2.Lines.UserFields.Fields.Item("U_Qty").Value = IssueQty;
        //                //제품원재료 변환 품목은 단가를 계산 후 입력
        //                if ((oRecordSet.EoF))
        //                {
        //                }
        //                else
        //                {
        //                    _with2.Lines.Price = Price;
        //                    _with2.Lines.UnitPrice = Price;
        //                    _with2.Lines.LineTotal = Price * IssueWt;
        //                }
        //                Cnt = Cnt + 1;

        //                j = j + 1;
        //            }
        //        }

        //        //// 완료
        //        if (Cnt > 0)
        //        {
        //            RetVal = DI_oInventoryGenEntry.Add();
        //            if ((0 != RetVal))
        //            {
        //                SubMain.Sbo_Company.GetLastError(out errCode, out ErrMsg);
        //                ErrNum = 1;
        //                goto Add_oInventoryGenEntry_Error;
        //            }
        //        }


        //        if (ChkType == 1)
        //        {
        //            SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        }
        //        else if (ChkType == 2)
        //        {
        //            SubMain.Sbo_Company.GetNewObjectCode(out SDocEntry);
        //            //cnt = 1
        //            //For i = 0 To oMat01.VisualRowCount - 1
        //            //     CpCode = oMat01.Columns("CpCode").Cells(i + 1).Specific.Value
        //            //     If CpCode = "CP80101" Or CpCode = "CP80111" Then
        //            //         oDS_PS_PP040L.setValue "U_OutDocC", i, sDocEntry
        //            //         oDS_PS_PP040L.setValue "U_OutLinC", i, cnt
        //            //         cnt = cnt + 1
        //        }
        //        // Next i
        //        oMat01.LoadFromDataSource();
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

        //        sQry = "Update [@PS_PP040L] set U_OutDocC = '" + SDocEntry + "', U_OutLinC = U_OutLin";
        //        //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //        sQry = sQry + " From [@PS_PP040L] where 1=1 and u_cpcode in ('CP80101','CP80111') and docentry = '" + Strings.Trim(oForm01.Items.Item("DocEntry").Specific.Value) + "' ";
        //        oRecordSet.DoQuery(sQry);
        //        //End If
        //    }

        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    //UPGRADE_NOTE: DI_oInventoryGenEntry 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    DI_oInventoryGenEntry = null;
        //    functionReturnValue = true;
        //    return functionReturnValue;
        //Add_oInventoryGenEntry_Error:
        //    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //    //UPGRADE_NOTE: oRecordSet 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    oRecordSet = null;
        //    //UPGRADE_NOTE: DI_oInventoryGenEntry 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    DI_oInventoryGenEntry = null;
        //    if (SubMain.Sbo_Company.InTransaction)
        //        SubMain.Sbo_Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //    functionReturnValue = false;
        //    if (ErrNum == 1)
        //    {
        //        MDC_Com.MDC_GF_Message(ref "Add_oInventoryGenEntry_Error:" + errCode + " - " + ErrMsg, ref "E");
        //    }
        //    else
        //    {
        //        MDC_Com.MDC_GF_Message(ref "Add_oInventoryGenEntry_Error:" + Err().Number + " - " + Err().Description, ref "E");
        //    }
        //    return functionReturnValue;
        //}
        #endregion





        #region PS_PP040_MTX01(미사용 추정, 본 클래스 마이그레이션 완료 후 재확인하여 삭제 예정)
        //private void PS_PP040_MTX01()
        //{
        //    // ERROR: Not supported in C#: OnErrorStatement

        //    ////메트릭스에 데이터 로드
        //    oForm01.Freeze(true);
        //    int i = 0;
        //    string Query01 = null;
        //    SAPbobsCOM.Recordset RecordSet01 = null;
        //    RecordSet01 = SubMain.Sbo_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //    string Param01 = null;
        //    string Param02 = null;
        //    string Param03 = null;
        //    string Param04 = null;
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    Param01 = Strings.Trim(oForm01.Items.Item("Param01").Specific.Value);
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    Param02 = Strings.Trim(oForm01.Items.Item("Param01").Specific.Value);
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    Param03 = Strings.Trim(oForm01.Items.Item("Param01").Specific.Value);
        //    //UPGRADE_WARNING: oForm01.Items().Specific.Value 개체의 기본 속성을 확인할 수 없습니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        //    Param04 = Strings.Trim(oForm01.Items.Item("Param01").Specific.Value);

        //    Query01 = "SELECT 10";
        //    RecordSet01.DoQuery(Query01);

        //    oMat01.Clear();
        //    oMat01.FlushToDataSource();
        //    oMat01.LoadFromDataSource();

        //    if ((RecordSet01.RecordCount == 0))
        //    {
        //        MDC_Com.MDC_GF_Message(ref "결과가 존재하지 않습니다.", ref "E");
        //        goto PS_PP040_MTX01_Exit;
        //    }

        //    SAPbouiCOM.ProgressBar ProgressBar01 = null;
        //    ProgressBar01 = SubMain.Sbo_Application.StatusBar.CreateProgressBar("조회시작!", RecordSet01.RecordCount, false);

        //    for (i = 0; i <= RecordSet01.RecordCount - 1; i++)
        //    {
        //        if (i != 0)
        //        {
        //            oDS_PS_PP040L.InsertRecord((i));
        //        }
        //        oDS_PS_PP040L.Offset = i;
        //        oDS_PS_PP040L.SetValue("U_COL01", i, RecordSet01.Fields.Item(0).Value);
        //        oDS_PS_PP040L.SetValue("U_COL02", i, RecordSet01.Fields.Item(1).Value);
        //        RecordSet01.MoveNext();
        //        ProgressBar01.Value = ProgressBar01.Value + 1;
        //        ProgressBar01.Text = ProgressBar01.Value + "/" + RecordSet01.RecordCount + "건 조회중...!";
        //    }
        //    oMat01.LoadFromDataSource();
        //    oMat01.AutoResizeColumns();
        //    oForm01.Update();

        //    ProgressBar01.Stop();
        //    //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgressBar01 = null;
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    oForm01.Freeze(false);
        //    return;
        //PS_PP040_MTX01_Exit:
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    oForm01.Freeze(false);
        //    if ((ProgressBar01 != null))
        //    {
        //        ProgressBar01.Stop();
        //    }
        //    return;
        //PS_PP040_MTX01_Error:
        //    ProgressBar01.Stop();
        //    //UPGRADE_NOTE: ProgressBar01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    ProgressBar01 = null;
        //    //UPGRADE_NOTE: RecordSet01 개체는 가비지가 수집되어야 소멸됩니다. 자세한 내용은 다음을 참조하십시오. 'ms-help://MS.VSExpressCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        //    RecordSet01 = null;
        //    oForm01.Freeze(false);
        //    SubMain.Sbo_Application.SetStatusBarMessage("PS_PP040_MTX01_Error: " + Err().Number + " - " + Err().Description, SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //}
        #endregion

    }
}
